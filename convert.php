<?php
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

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $number = $_POST['number'];
    $number = str_replace(',', '.', str_replace('.', '', $number)); // Mengubah format angka
    if (is_numeric($number) && $number >= 0 && $number <= pow(10, 33)) { // Batas hingga Puluhan Triliun
        echo terbilang($number) . " Rupiah";
    } else {
        echo "Masukkan angka yang valid antara 0 dan 10^33.";
    }
}
?>
