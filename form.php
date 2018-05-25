<!--
-- Source Code from My Notes Code (www.mynotescode.com)
-- 
-- Follow Us on Social Media
-- Facebook : http://facebook.com/mynotescode/
-- Twitter  : http://twitter.com/code_notes
-- Google+  : http://plus.google.com/118319575543333993544
--
-- Terimakasih telah mengunjungi blog kami.
-- Jangan lupa untuk Like dan Share catatan-catatan yang ada di blog kami.
-->

<!DOCTYPE html>
<html lang="en">
	<head>
		<meta charset="utf-8">
		<meta http-equiv="X-UA-Compatible" content="IE=edge">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Import Data Excel dengan PHP</title>

		<!-- Load File bootstrap.min.css yang ada difolder css -->
		<link href="css/bootstrap.min.css" rel="stylesheet">

		<!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
		<!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
		<!--[if lt IE 9]>
		  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
		  <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
		<![endif]-->
		
		<!-- Style untuk Loading -->
		<style>
        #loading{
			background: whitesmoke;
			position: absolute;
			top: 140px;
			left: 82px;
			padding: 5px 10px;
			border: 1px solid #ccc;
		}
		</style>
		
		<!-- Load File jquery.min.js yang ada difolder js -->
		<script src="js/jquery.min.js"></script>
		
		<script>
		$(document).ready(function(){
			// Sembunyikan alert validasi kosong
			$("#kosong").hide();
		});
		</script>
	</head>
	<body>
		<!-- Membuat Menu Header / Navbar -->
		<nav class="navbar navbar-inverse" role="navigation">
			<div class="container-fluid">
				<div class="navbar-header">
					<a class="navbar-brand" href="#" style="color: white;"><b>Import Data Excel dengan PHP</b></a>
				</div>
				<p class="navbar-text navbar-right hidden-xs" style="color: white;padding-right: 10px;">
					FOLLOW US ON &nbsp;
					<a target="_blank" style="background: #3b5998; padding: 0 5px; border-radius: 4px; color: #f7f7f7; text-decoration: none;" href="https://www.facebook.com/mynotescode">Facebook</a> 
					<a target="_blank" style="background: #00aced; padding: 0 5px; border-radius: 4px; color: #ffffff; text-decoration: none;" href="https://twitter.com/code_notes">Twitter</a> 
					<a target="_blank" style="background: #d34836; padding: 0 5px; border-radius: 4px; color: #ffffff; text-decoration: none;" href="https://plus.google.com/118319575543333993544">Google+</a>
				</p>
			</div>
		</nav>
		
		<!-- Content -->
		<div style="padding: 0 15px;">
			<!-- Buat sebuah tombol Cancel untuk kemabli ke halaman awal / view data -->
			<a href="index.php" class="btn btn-danger pull-right">
				<span class="glyphicon glyphicon-remove"></span> Cancel
			</a>
			
			<h3>Form Import Data</h3>
			<hr>
			
			<!-- Buat sebuah tag form dan arahkan action nya ke file ini lagi -->
			<form method="post" action="" enctype="multipart/form-data">
				<a href="Format.xlsx" class="btn btn-default">
					<span class="glyphicon glyphicon-download"></span>
					Download Format
				</a><br><br>
				
				<!-- 
				-- Buat sebuah input type file
				-- class pull-left berfungsi agar file input berada di sebelah kiri
				-->
				<input type="file" name="file" class="pull-left">
				
				<button type="submit" name="preview" class="btn btn-success btn-sm">
					<span class="glyphicon glyphicon-eye-open"></span> Preview
				</button>
			</form>
			
			<hr>
			
			<!-- Buat Preview Data -->
			<?php
			// Jika user telah mengklik tombol Preview
			if(isset($_POST['preview'])){
				//$ip = ; // Ambil IP Address dari User
				$nama_file_baru = 'data.xlsx';
				
				// Cek apakah terdapat file data.xlsx pada folder tmp
				if(is_file('tmp/'.$nama_file_baru)) // Jika file tersebut ada
					unlink('tmp/'.$nama_file_baru); // Hapus file tersebut
				
				$tipe_file = $_FILES['file']['type']; // Ambil tipe file yang akan diupload
				$tmp_file = $_FILES['file']['tmp_name'];
				
				// Cek apakah file yang diupload adalah file Excel 2007 (.xlsx)
				if($tipe_file == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
					// Upload file yang dipilih ke folder tmp
					// dan rename file tersebut menjadi data{ip_address}.xlsx
					// {ip_address} diganti jadi ip address user yang ada di variabel $ip
					// Contoh nama file setelah di rename : data127.0.0.1.xlsx
					move_uploaded_file($tmp_file, 'tmp/'.$nama_file_baru);
					
					// Load librari PHPExcel nya
					require_once 'PHPExcel/PHPExcel.php';
					
					$excelreader = new PHPExcel_Reader_Excel2007();
					$loadexcel = $excelreader->load('tmp/'.$nama_file_baru); // Load file yang tadi diupload ke folder tmp
					$sheet = $loadexcel->getActiveSheet()->toArray(null, true, true ,true);
					
					// Buat sebuah tag form untuk proses import data ke database
					echo "<form method='post' action='import.php'>";
					
					// Buat sebuah div untuk alert validasi kosong
					echo "<div class='alert alert-danger' id='kosong'>
					Semua data belum diisi, Ada <span id='jumlah_kosong'></span> data yang belum diisi.
					</div>";
					
					echo "<table class='table table-bordered'>
					<tr>
						<th colspan='11' class='text-center'>Preview Data</th>
					</tr>
					<tr> 
						
						<th>Nomor PR</th>
						<th>Nama Barang</th>
						<th>Qty</th>
						<th>UoM</th>
						<th>Tanggal</th>
						<th>Nilai Total PR</th>
						<th>Harga Akhir</th>
						<th>Nilai Total</th>
						<th>Penghematan</th>
						<th>Suplier Pemenang</th>
						<th>Bagian</th>
					</tr>";
					
					$numrow = 1;
					//$kosong = 0;
					foreach($sheet as $row){ // Lakukan perulangan dari data yang ada di excel
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
						continue;} // Lewat data pada baris ini (masuk ke looping selanjutnya / baris selanjutnya)
						
						// Cek $numrow apakah lebih dari 1
						// Artinya karena baris pertama adalah nama-nama kolom
						// Jadi dilewat saja, tidak usah diimport
						if($numrow > 1){
							// Validasi apakah semua data telah diisi
							$no_pr_td = ( ! empty($no_pr))? "" : " style='background: #E07171;'"; // Jika NIS kosong, beri warna merah
							$nama_barang_td = ( ! empty($nama_barang))? "" : " style='background: #E07171;'"; // Jika Nama kosong, beri warna merah
							$qty_td = ( ! empty($qty))? "" : " style='background: #E07171;'"; // Jika Jenis Kelamin kosong, beri warna merah
							$uom_td = ( ! empty($uom))? "" : " style='background: #E07171;'"; // Jika Telepon kosong, beri warna merah
							$tanggal_td = ( ! empty($tanggal))? "" : " style='background: #E07171;'"; // Jika Alamat kosong, beri warna merah
							$nilai_total_pr_td = ( ! empty($nilai_total_pr))? "" : " style='background: #E07171;'";
							$harga_akhir_td = ( ! empty($harga_akhir))? "" : " style='background: #E07171;'";
							$nilai_total_td = ( ! empty($nilai_total))? "" : " style='background: #E07171;'";
							$penghematan_td = ( ($penghematan))? "" : " style='background: #E07171;'";
							$suplier_pemenang_td = ( ! empty($suplier_pemenang))? "" : " style='background: #E07171;'";
							$bagian_td = ( ! empty($bagian))? "" : " style='background: #E07171;'";
							// Jika salah satu data ada yang kosong
							if(empty($no_pr) or empty($nama_barang) or empty($qty) or empty($uom) or empty($tanggal) or empty($nilai_total_pr) or empty($harga_akhir) or empty($nilai_total) or empty($penghematan) or empty($suplier_pemenang) or empty($bagian)){
								//$kosong++; // Tambah 1 variabel $kosong
							}
							
							echo "<tr>";
							echo "<td".$no_pr_td.">".$no_pr."</td>";
							echo "<td".$nama_barang_td.">".$nama_barang."</td>";
							echo "<td".$qty_td.">".$qty."</td>";
							echo "<td".$uom_td.">".$uom."</td>";
							echo "<td".$tanggal_td.">".$tanggal."</td>";
							echo "<td".$nilai_total_pr_td.">".$nilai_total_pr."</td>";
							echo "<td".$harga_akhir_td.">".$harga_akhir."</td>";
							echo "<td".$nilai_total_td.">".$nilai_total."</td>";
							echo "<td".$penghematan_td.">".$penghematan."</td>";
							echo "<td".$suplier_pemenang_td.">".$suplier_pemenang."</td>";
							echo "<td".$bagian_td.">".$bagian."</td>";
							echo "</tr>";
						}
						
						$numrow++; // Tambah 1 setiap kali looping
					}
					
					echo "</table>";
					
					// Cek apakah variabel kosong lebih dari 1
					// Jika lebih dari 1, berarti ada data yang masih kosong
				
						echo "<hr>";
						
						// Buat sebuah tombol untuk mengimport data ke database
						echo "<button type='submit' name='import' class='btn btn-primary'><span class='glyphicon glyphicon-upload'></span> Import</button>";
					}
					
					echo "</form>";
				}else{ // Jika file yang diupload bukan File Excel 2007 (.xlsx)
					// Munculkan pesan validasi
					echo "<div class='alert alert-danger'>
					Hanya File Excel 2007 (.xlsx) yang diperbolehkan
					</div>";
				}
			
			?>
		</div>
	</body>
</html>

