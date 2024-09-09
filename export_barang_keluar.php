<?php
require 'vendor/autoload.php'; // Pastikan composer autoload di-include

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

// Informasi koneksi database
$host = 'localhost';
$db = 'inventorywise'; // Ganti dengan nama database Anda
$user = 'root';             // Ganti dengan username database Anda
$pass = '';                 // Ganti dengan password database Anda

$conn = new mysqli($host, $user, $pass, $db);

// Periksa koneksi
if ($conn->connect_error) {
    die("Koneksi gagal: " . $conn->connect_error);
}

// Query untuk mengambil data dari tb_barang_keluar
$sql = "SELECT * FROM tb_barang_keluar";
$result = $conn->query($sql);

if ($result->num_rows > 0) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Set judul laporan di baris pertama (misalnya, baris 1, kolom A sampai H)
    $sheet->mergeCells('A1:H1'); // Menggabungkan sel A1 sampai H1 untuk judul
    $sheet->setCellValue('A1', 'Laporan Barang Keluar');
    $sheet->getStyle('A1')->getFont()->setBold(true); // Menjadikan teks judul tebal
    $sheet->getStyle('A1')->getFont()->setSize(16); // Mengatur ukuran font judul
    $sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER); // Menyelaraskan teks ke tengah

    // Mengisi judul kolom di baris kedua
    $sheet->setCellValue('A2', 'ID Transaksi');
    $sheet->setCellValue('B2', 'Tanggal Masuk');
    $sheet->setCellValue('C2', 'Tanggal Keluar');
    $sheet->setCellValue('D2', 'Nama Pengguna');
    $sheet->setCellValue('E2', 'Kode Barang');
    $sheet->setCellValue('F2', 'Nama Barang');
    $sheet->setCellValue('G2', 'Satuan Barang');
    $sheet->setCellValue('H2', 'Jumlah Barang');

    // Menambahkan border pada judul kolom
    $sheet->getStyle('A2:H2')->applyFromArray([
        'font' => ['bold' => true],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => ['argb' => 'FFFF99']
        ]
    ]);

    $rowNumber = 3; // Mulai dari baris ketiga karena baris pertama untuk judul dan baris kedua untuk judul kolom

    while ($row = $result->fetch_assoc()) {
        $sheet->setCellValue('A' . $rowNumber, $row['id_transaksi']);
        $sheet->setCellValue('B' . $rowNumber, $row['tanggal_masuk']);
        $sheet->setCellValue('C' . $rowNumber, $row['tanggal_keluar']);
        $sheet->setCellValue('D' . $rowNumber, $row['nama_pengguna']);
        $sheet->setCellValue('E' . $rowNumber, $row['kode_barang']);
        $sheet->setCellValue('F' . $rowNumber, $row['nama_barang']);
        $sheet->setCellValue('G' . $rowNumber, $row['satuan']);
        $sheet->setCellValue('H' . $rowNumber, $row['jumlah']);
        $rowNumber++;
    }

    // Menambahkan border pada seluruh data
    $sheet->getStyle('A2:H' . ($rowNumber - 1))->applyFromArray([
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ]);

    // Menyesuaikan lebar kolom secara otomatis
    foreach (range('A', 'H') as $columnID) {
        $sheet->getColumnDimension($columnID)->setAutoSize(true);
    }

    // Menambahkan filter pada baris kedua
    $sheet->setAutoFilter('A2:H' . ($rowNumber - 1));

    $writer = new Xlsx($spreadsheet);
    $filename = 'barang_keluar.xlsx';

    // Mengatur header untuk download file Excel
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $filename . '"');
    header('Cache-Control: max-age=0');

    $writer->save('php://output');
} else {
    echo "Tidak ada data untuk diekspor.";
}

$conn->close();
?>