<?php
require 'C:/xampp/phpMyAdmin/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

if ($_SERVER["REQUEST_METHOD"] == "POST" && isset($_FILES["excelFile"])) {
    $file = $_FILES["excelFile"]["tmp_name"];
    $spreadsheet = IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();

    $highestRow = $sheet->getHighestRow();
    for ($row = 1; $row <= $highestRow; $row++) {
        $number = $sheet->getCellByColumnAndRow(1, $row)->getValue();
        if (is_numeric($number)) {
            $terbilang = terbilang($number);
            $sheet->setCellValueByColumnAndRow(2, $row, $terbilang . ' Rupiah');
        }
    }

    $writer = new Xlsx($spreadsheet);
    $outputFile = 'output.xlsx';
    $writer->save($outputFile);

    echo '<a href="'.$outputFile.'" class="btn btn-success">Download Hasil</a>';
}

function terbilang($number) {
    $number = (float)$number;
    $units = [
        '', 'Ribu', 'Juta', 'Miliar', 'Triliun'
    ];
    
    if ($number < 12) {
        return satuan($number);
    } elseif ($number < 20) {
        return satuan($number - 10) . ' Belas';
    } elseif ($number < 100) {
        return satuan((int)($number / 10)) . ' Puluh ' . satuan($number % 10);
    } elseif ($number < 200) {
        return 'Seratus ' . terbilang($number - 100);
    } elseif ($number < 1000) {
        return satuan((int)($number / 100)) . ' Ratus ' . terbilang($number % 100);
    } else {
        for ($i = 1; $i < count($units); $i++) {
            $unitValue = pow(1000, $i);
            if ($number < pow(1000, $i + 1)) {
                return terbilang((int)($number / $unitValue)) . ' ' . $units[$i] . ' ' . terbilang($number % $unitValue);
            }
        }
    }
}

function satuan($number) {
    switch ($number) {
        case 1:
            return 'Satu';
        case 2:
            return 'Dua';
        case 3:
            return 'Tiga';
        case 4:
            return 'Empat';
        case 5:
            return 'Lima';
        case 6:
            return 'Enam';
        case 7:
            return 'Tujuh';
        case 8:
            return 'Delapan';
        case 9:
            return 'Sembilan';
        case 10:
            return 'Sepuluh';        
        case 11:
            return 'Sebelas';
        default:
            return '';
    }
}
?>
