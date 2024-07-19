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
                var number = $('#number').val().replace(/\./g, '').replace(',', '.');
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

            function showCopySuccess() {
                var alertBox = $('<div class="alert alert-success" role="alert">Tersalin, Segera CTRL + V (PASTE).</div>');
                $('.container').prepend(alertBox);
                setTimeout(function() {
                    alertBox.fadeOut('slow', function() {
                        $(this).remove();
                    });
                }, 3000); // waktu tampil notifikasi berhasil
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
		</br>
        <form id="convertForm" method="post">
            <div class="form-group">
                <label for="number">Masukkan Angka:</label>
                <input type="text" id="number" name="number" class="form-control" required>
            </div>
            <button type="submit" class="btn btn-primary">Konversi</button>
            <button type="button" id="copyButton" class="btn btn-secondary ml-2">Copy</button>
        </form>
        <div id="result" class="mt-4"></div>
		<hr>
		<p class="mb-0 text-center text-muted">©
			<script>document.write(new Date().getFullYear())</script>
			<i class="mdi mdi-heart text-danger"></i> by <a href="https://galih-os.github.io/" target="_blank" class="text-muted">Galih-OS.</a>
        </p>
    </div>
</body>
</html>
