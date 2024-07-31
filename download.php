<?php
require 'C:/xampp/phpMyAdmin/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Jumlah Angka');
$sheet->setCellValue('B1', 'Terbilang');

$writer = new Xlsx($spreadsheet);

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="template.xlsx"');
header('Cache-Control: max-age=0');

$writer->save('php://output');
?>
