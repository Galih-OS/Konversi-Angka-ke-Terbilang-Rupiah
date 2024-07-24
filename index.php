<!DOCTYPE html>
<html>
<head>
    <title>Konversi Angka ke Terbilang Rupiah</title>
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    <script>
        $(document).ready(function(){
            $('#convertForm').on('submit', function(e){
                e.preventDefault();
                var number = $('#number').val().replace(/,/g, '.').replace(/\./g, '');
                $.ajax({
                    type: 'POST',
                    url: 'convert.php',
                    data: {number: number},
                    success: function(response) {
                        $('#result').html(response);
                        $('#copyButton').show();
                    }
                });
            });

            $('#copyButton').on('click', function(){
                var resultText = $('#result').text();
                var tempInput = $('<input>');
                $('body').append(tempInput);
                tempInput.val(resultText).select();
                document.execCommand('copy');
                tempInput.remove();
                showCopySuccess();
            });

            $('#resetButton').on('click', function(){
                $('#convertForm')[0].reset();
                $('#result').html('');
                $('#copyButton').hide();
            });

            function showCopySuccess() {
                var alertBox = $('<div class="alert alert-success" role="alert">Tersalin, segera Paste (CTRL + V).</div>');
                $('.container').prepend(alertBox);
                setTimeout(function() {
                    alertBox.fadeOut('slow', function() {
                        $(this).remove();
                    });
                }, 3000); // 3 seconds
            }
        });
    </script>
    <style>
        body {
            background-color: #f8f9fa;
        }
        .container {
            max-width: 600px;
            margin-top: 50px;
            padding: 20px;
            background: #ffffff;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }
        h1 {
            margin-bottom: 20px;
        }
        #copyButton {
            display: none;
        }
    </style>
    <link rel="shortcut icon" href="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEgWo_QOl2BvBYszDGJkAQJqUT8TEL7c4FpB553Jj-iQHLVvx9IC7J3xBqcu5sR4qmzvf3JCc3g5WrGOYbWPo4H7TdWO7jgfyUeIvPxzj2rPFzMLd2GPvbUwz5p5bfZWPTg/w68-h68-p-k-no-nu/gos.png">
</head>
<body>
    <div class="container">

        <center><h3>Konversi Angka ke Terbilang Rupiah</h3></center>
        <center><h3>Pembatasan Terbilang Hanya s/d Triliun</h3></center>
        <br>
        <form id="convertForm" method="post">
            <div class="form-group">
                <label for="number">Masukkan Angka, bisa menggunakan (koma atau titik):</label>
                <input type="text" id="number" name="number" class="form-control" required>
            </div>
            <button type="submit" class="btn btn-primary">Konversi</button>
            <button type="button" id="copyButton" class="btn btn-secondary ml-2">Copy</button>
            <button type="button" id="resetButton" class="btn btn-danger ml-2">Reset</button>
        </form>
        <div id="result" class="mt-4"></div>
        <hr>
        <p class="mb-0 text-center text-muted">©
            <script>document.write(new Date().getFullYear())</script>
            <i class="mdi mdi-heart text-danger"></i> by <a href="https://galih-os.github.io/" target="_blank" class="text-muted">Galih-OS.</a> | Call Me, Telegram : 
			<svg xmlns="http://www.w3.org/2000/svg" width="1em" height="1em" viewBox="0 0 24 24"><a href="https://t.me/galihos" target="_blank"><path d="M9.78 18.65l.28-4.23l7.68-6.92c.34-.31-.07-.46-.52-.19L7.74 13.3L3.64 12c-.88-.25-.89-.86.2-1.3l15.97-6.16c.73-.33 1.43.18 1.15 1.3l-2.72 12.81c-.19.91-.74 1.13-1.5.71L12.6 16.3l-1.99 1.93c-.23.23-.42.42-.83.42z" fill="currentColor"/></svg>
    
			</br>
			</br>
			<i class="mdi mdi-heart text-danger" id="google_translate_element">
				<script type="text/javascript">
				function googleTranslateElementInit() {
				new google.translate.TranslateElement({pageLanguage: "en", layout: google.translate.TranslateElement.InlineLayout.SIMPLE, multilanguagePage: true}, "google_translate_element");
				}
				</script> 
				<script type="text/javascript" src="//translate.google.com/translate_a/element.js?cb=googleTranslateElementInit"></script>
			
			</i>
        </p>
    </div>
</body>
</html>
