<?php
/*
-- Source Code from My Notes Code (www.mynotescode.com)
-- 
-- Follow Us on Social Media
-- Facebook : http://facebook.com/mynotescode/
-- Twitter  : http://twitter.com/code_notes
-- Google+  : http://plus.google.com/118319575543333993544
--
-- Terimakasih telah mengunjungi blog kami.
-- Jangan lupa untuk Like dan Share catatan-catatan yang ada di blog kami.
*/

// Load file koneksi.php
include "koneksi.php";

if(isset($_POST['import'])){ // Jika user mengklik tombol Import
	$nama_file_baru = 'data.xlsx';
	
	// Load librari PHPExcel nya
	require_once 'PHPExcel/PHPExcel.php';
	
	$excelreader = new PHPExcel_Reader_Excel2007();
	$loadexcel = $excelreader->load('tmp/'.$nama_file_baru); // Load file excel yang tadi diupload ke folder tmp
	$sheet = $loadexcel->getActiveSheet()->toArray(null, true, true ,true);
	
	// Buat query Insert
	$sql = $pdo->prepare("INSERT INTO eauction VALUES(:no_pr,:nama_barang,:qty,:uom,:tanggal,:nilai_total_pr,:harga_akhir,:nilai_total,:penghematan,:suplier_pemenang,:bagian)");
	
	$numrow = 1;
	foreach($sheet as $row){
		// Ambil data pada excel sesuai Kolom
		
		$no_pr = $row['B']; // Ambil data NIS
		$nama_barang = $row['C']; // Ambil data nama
		$qty = $row['D']; // Ambil data jenis kelamin
		$uom = $row['E']; // Ambil data telepon
		$tanggal = $row['F']; // Ambil data alamat
		$nilai_total_pr = $row['G']; // Ambil data alamat
		$harga_akhir = $row['H']; // Ambil data alamat
		$nilai_total = $row['I']; // Ambil data alamat
		$penghematan = $row['J']; // Ambil data alamat
		$suplier_pemenang = $row['K']; // Ambil data alamat
		$bagian = $row['L']; // Ambil data alamat
		// Cek jika semua data tidak diisi
		if(empty($no_pr) && empty($nama_barang) && empty($qty) && empty($uom) && empty($tanggal) && empty($nilai_total_pr) && empty($harga_akhir) && empty($nilai_total) && empty($penghematan) && empty($suplier_pemenang) && empty($bagian)){
		continue; }// Lewat data pada baris ini (masuk ke looping selanjutnya / baris selanjutnya)
		
		// Cek $numrow apakah lebih dari 1
		// Artinya karena baris pertama adalah nama-nama kolom
		// Jadi dilewat saja, tidak usah diimport
		if($numrow > 1){
			// Proses simpan ke Database
			$sql->bindParam(':no_pr', $no_pr);
			$sql->bindParam(':nama_barang', $nama_barang);
			$sql->bindParam(':qty', $qty);
			$sql->bindParam(':uom', $uom);
			$sql->bindParam(':tanggal', $tanggal);
			$sql->bindParam(':nilai_total_pr', $nilai_total_pr);
			$sql->bindParam(':harga_akhir', $harga_akhir);
			$sql->bindParam(':nilai_total', $nilai_total);
			$sql->bindParam(':penghematan', $penghematan);
			$sql->bindParam(':suplier_pemenang', $suplier_pemenang);
			$sql->bindParam(':bagian', $bagian);
			$sql->execute(); // Eksekusi query insert
		}
		
		$numrow++; // Tambah 1 setiap kali looping
	}
}

header('location: index.php'); // Redirect ke halaman awal
?>
