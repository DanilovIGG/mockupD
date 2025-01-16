<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use Bitrix\Main\Loader;
use \Bitrix\Iblock\Elements\ElementClothesTable;
use CIBlockSection;
use CIBlockElement;

require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");

function processExcelFile($fileTmpPath)
{
  try {
    if (!Loader::includeModule("iblock")) {
      throw new Exception("Модуль инфоблоков не загружен");
    }

    if (!Loader::includeModule("catalog")) {
      throw new Exception("Модуль catalog не загружен");
    }

    if (!file_exists($fileTmpPath) || pathinfo($fileTmpPath, PATHINFO_EXTENSION) !== 'xlsx') {
      throw new Exception("Неверный формат файла. Ожидается .xlsx");
    }

    $spreadsheet = IOFactory::load($fileTmpPath);
    $sheet = $spreadsheet->getActiveSheet();

    $currentSectionId = null; // ID текущего раздела
    $parentSectionId = null;

    foreach ($sheet->getRowIterator() as $row) {
      $cellIterator = $row->getCellIterator();
      $cellIterator->setIterateOnlyExistingCells(false);

      $cells = iterator_to_array($cellIterator);
      $firstCellValue = trim($cells[0]->getValue());

      if (!empty($firstCellValue) && preg_match('/^[\p{L}\s.,-]+$/u', $firstCellValue)) {
        $sectionId = issetSection($firstCellValue);
        $currentSectionId = $sectionId ?: createSection($firstCellValue, $parentSectionId);
        $parentSectionId = $currentSectionId;
      } else {
        $article = trim($cells[0]->getValue());
        $name = trim($cells[3]->getValue());
        $price = floatval($cells[5]->getValue());
        $quantity = intval($cells[6]->getValue());

        if ($article && $name) {
          $elementId = issetElement($article);
          if ($elementId) {
            if ($_POST['update'] === 'y'){
              updateElementInCatalog($elementId, $quantity);
              updatePriceToCatalogElement($elementId, $price);
            }
          } else {
            createElement($currentSectionId, $article, $name, $price, $quantity);
          }
        }
      }
    }

    return ['success' => true, 'message' => 'Файл обработан успешно'];
  } catch (Exception $e) {
    return ['success' => false, 'message' => "Ошибка: " . $e->getMessage()];
  }
}

function createSection($sectionName, $parentSectionId = null)
{
  $section = new CIBlockSection;
  $arParams = ["replace_space" => "-", "replace_other" => "-"];
  $sectionCode = Cutil::translit($sectionName, "ru", $arParams);

  $fields = [
    "IBLOCK_ID"         => 2,
    "IBLOCK_SECTION_ID" => $parentSectionId,
    "CODE"              => $sectionCode,
    "NAME"              => $sectionName,
  ];

  $sectionId = $section->Add($fields);

  if (!$sectionId) {
    throw new Exception("Ошибка создания раздела: " . $section->LAST_ERROR);
  }

  return $sectionId;
}

function issetSection($sectionName)
{
  $sectionId = null;
  $dbSections = CIBlockSection::GetList(
    ['SORT' => 'ASC'],
    ['IBLOCK_ID' => 2, 'NAME' => $sectionName],
    false,
    ['ID']
  );

  if ($section = $dbSections->Fetch()) {
    $sectionId = $section['ID'];
  }

  return $sectionId;
}

function addElementToCatalog($productId, $quantity)
{
  $arFields = [
    "ID"       => intval($productId),
    "QUANTITY" => intval($quantity),
  ];

  $result = \Bitrix\Catalog\Model\Product::add($arFields);
  if (!$result->isSuccess()) {
    throw new Exception("Ошибка добавления параметров товара: " . implode(", ", $result->getErrorMessages()));
  }

  return $result->getId();
}

function addPriceToCatalogElement($productId, $price)
{
  $arFields = [
    "PRODUCT_ID"       => $productId,
    "CATALOG_GROUP_ID" => 1,
    "PRICE"            => $price,
    "CURRENCY"         => "RUB",
  ];

  $result = \Bitrix\Catalog\Model\Price::add($arFields);
  if (!$result->isSuccess()) {
    throw new Exception("Ошибка добавления цены: " . implode(", ", $result->getErrorMessages()));
  }
}

function updateElementInCatalog($productId, $quantity)
{
  $arFields = ["QUANTITY" => intval($quantity)];

  $result = \Bitrix\Catalog\Model\Product::update($productId, $arFields);
  if (!$result->isSuccess()) {
    throw new Exception("Ошибка обновления кол-ва товара: " . implode(", ", $result->getErrorMessages()));
  }

  return $result->getId();
}

function updatePriceToCatalogElement($productId, $price)
{
  $dbPrice = \Bitrix\Catalog\Model\Price::getList([
    "filter" => [
      "PRODUCT_ID"       => $productId,
      "CATALOG_GROUP_ID" => 1,
    ],
  ]);

  $arFields = [
    "PRODUCT_ID"       => $productId,
    "CATALOG_GROUP_ID" => 1,
    "PRICE"            => $price,
    "CURRENCY"         => "RUB",
  ];

  if ($arPrice = $dbPrice->fetch()) {
    $result = \Bitrix\Catalog\Model\Price::update($arPrice["ID"], $arFields);
  } else {
    $result = \Bitrix\Catalog\Model\Price::add($arFields);
  }

  if (!$result->isSuccess()) {
    throw new Exception("Ошибка обновления цены: " . implode(", ", $result->getErrorMessages()));
  }
}

function createElement($sectionId, $article, $name, $price, $quantity)
{
  $element = new CIBlockElement;

  $fields = [
    "IBLOCK_ID"         => 2,
    "IBLOCK_SECTION_ID" => $sectionId,
    "NAME"              => $name,
    "ACTIVE"            => "Y",
    "PROPERTY_VALUES"   => [
      "ARTNUMBER" => $article,
    ],
  ];

  $elementId = $element->Add($fields);

  if (!$elementId) {
    throw new Exception("Ошибка создания элемента: " . $element->LAST_ERROR);
  }

  addElementToCatalog($elementId, $quantity);
  addPriceToCatalogElement($elementId, $price);

  return $elementId;
}

function issetElement($article)
{
  $elementId = null;
  $elements = \Bitrix\Iblock\Elements\ElementCatalogTable::getList([
    'select' => ['ID'],
    'filter' => ['=ACTIVE' => 'Y', '=ARTNUMBER_VALUE' => $article],
  ])->fetch();

  if ($elements) {
    $elementId = $elements['ID'];
  }

  return $elementId;
}

if ($_SERVER["REQUEST_METHOD"] == "POST" && isset($_FILES["excelFile"])) {
  $fileTmpPath = $_FILES["excelFile"]["tmp_name"];
  $response = processExcelFile($fileTmpPath);

  header('Content-Type: application/json');
  echo json_encode($response);
  die();
}

http_response_code(400);
echo json_encode(['success' => false, 'message' => 'Неверный запрос']);
die();
