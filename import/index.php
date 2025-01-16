<?php
require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/header.php");
$APPLICATION->SetTitle("Импорт каталога");

$APPLICATION->IncludeFile(SITE_DIR . '/_inc.files/forms/import-form.php', [], ['SHOW_BORDER' => false]);

require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/footer.php"); ?>
