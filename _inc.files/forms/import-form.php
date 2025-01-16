<style>
  h2 {
    color: #333;
    text-align: center;
  }
  .form-container {
    background-color: #fff;
    max-width: 400px;
    margin: 20px auto;
    padding: 20px;
    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
    border-radius: 8px;
  }
  label {
    display: block;
    margin-bottom: 10px;
    font-weight: bold;
    color: #333;
  }
  input[type="file"] {
    width: 100%;
    padding: 8px;
    margin-bottom: 20px;
    border: 1px solid #ccc;
    border-radius: 4px;
  }
  input[type="submit"] {
    background-color: #24a844;
    color: #fff;
    border: none;
    padding: 10px 15px;
    cursor: pointer;
    border-radius: 4px;
    width: 100%;
    font-size: 16px;
  }
  input[type="submit"]:hover {
    background-color: #1f913b;
  }
  .success {
    color: #24a844;
    font-weight: bold;
    text-align: center;
    margin-top: 20px;
  }
  .error {
    color: #ff4d4d;
    font-weight: bold;
    text-align: center;
    margin-top: 20px;
  }
  .result {
    background-color: #fff;
    padding: 10px;
    margin-top: 20px;
    border-radius: 4px;
    font-family: monospace;
    font-size: 14px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
  }
  table, th, td {
    border: 1px solid #ccc;
  }
  th, td {
    padding: 8px;
    text-align: center;
  }
  th {
    background-color: #24a844;
  }
</style>
<div class="max-width">
  <h2>Загрузите Excel файл</h2>
  <div class="form-container">
    <form id="uploadForm" enctype="multipart/form-data">
      <label for="excelFile">Выберите Excel файл:</label>
      <input type="file" name="excelFile" id="excelFile" required>
      <input type="submit" value="Загрузить и обработать">
    </form>
    <div id="responseMessage"></div>
  </div>

  <div id="resultContainer" style="display: none;">
    <h2>Результаты обработки</h2>
    <div id="result" class="result"></div>
  </div>
</div>
<script>
  $(document).ready(function() {
    $('#uploadForm').on('submit', function(event) {
      event.preventDefault();

      var formData = new FormData(this)

      $.ajax({
        url: '/ajax/parse_excel.php',
        type: 'POST',
        data: formData,
        contentType: false,
        processData: false,
        beforeSend: function() {
          $('#responseMessage').text('Загрузка...').removeClass('success error')
          $('#resultContainer').hide()
        },
        success: function(response) {
          try {
            var data = response
            if (data.success) {
              $('#responseMessage').text('Файл успешно обработан.').addClass('success')
              var tableHtml = '<table><thead><tr><th>ID</th><th>Код</th><th>Цена</th><th>Статус обновления</th><th>Ошибка</th><th>Детальная страница</th><th>В админке</th></tr></thead><tbody>'
              data.result.forEach(function(item) {
                tableHtml += '<tr><td>' + (item.ID || '—') + '</td><td>' + item.code + '</td><td>' + item.price + '</td><td>' + item.status + '</td><td>' + (item.error || '—') +
                    '</td><td><a href="' + item.detail_page_url + '" target="_blank">Открыть</a></td><td><a href="' + item.admin_url +
                    '" target="_blank">Редактировать</a></td></tr>'
              })
              tableHtml += '</tbody></table>'

              $('#result').html(tableHtml)
              $('#resultContainer').show()
            } else {
              $('#responseMessage').text(data.message).addClass('error')
            }
          } catch (error) {
            $('#responseMessage').text('Ошибка при получении ответа сервера.').addClass('error')
          }
        },
        error: function(xhr, status, error) {
          $('#responseMessage').text('Произошла ошибка при загрузке файла: ' + error).addClass('error')
        },
      })
    })
  })
</script>
