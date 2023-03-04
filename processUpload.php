<?php 
// var_dump($_FILES);
$src = $_FILES["excel"]["tmp_name"];  // đi từ excel vào trong thư mục tạm của xamp
$filename = $_FILES["excel"]["name"]; // đổi tên
$dest = "upload/$filename";
move_uploaded_file($src, $dest);// duy chuyển từ nơi này đến nơi kia
require 'vendor/autoload.php'; //chỉ định file cụ thể

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$inputFileName = 'upload/DSNV.xlsx';

// đọc files
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
$sheet = $spreadsheet->getActiveSheet(); // làm từng cái sheet getActiveSheet() là sheet hiện tại
// lấy giá trị ra
// $cell = $sheet-> getcell("B3");
// echo $cell->getvalue();

// lấy dữ liệu ra
$start = 2;
$end   = $sheet->getHighestDataRow(); 
// echo $end; // lấy dòng cuối
// echo "<br>";
?>
<!-- cắt từ đầu đến <tbody>(1) -->
<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>DANH SÁCH NHÂN VIÊN</title>
	<link rel="stylesheet" href="public/vendor/bootstrap-4.5.3-dist/css/bootstrap.min.css">
	
</head>
<body>

<div class="container-fluid">
	<div class="table-responsive">
		<table class="table table-striped">
		  
		  <thead class="thead-dark">
		    <tr>
		      <th scope="col">Số Thứ Tự</th>
		      <th scope="col">Mã Nhân Viên</th>
		      <th scope="col">Họ Và Tên</th>
		      <th scope="col">Số Điện Thoại</th>
		      <th scope="col">Email</th>
		      <th scope="col">giới Tính</th>
		      <th scope="col">Lương</th>
		      <th scope="col">Bộ Phận</th>
		      <th scope="col">Ngày Đi Làm</th>
		    </tr>
		  </thead>
		  <tbody>

			<?php 
			$order = 0;
			$salaryTotal = 0;
			$salaries = [];
			for ($row = $start ; $row <= $end ; $row++) {  // duyệt từng dòng
				# code...
				$staffcode 	= $sheet->getCell("A$row")->getValue();
				$name 		= $sheet->getCell("B$row")->getValue();
				$mobile 	= $sheet->getCell("C$row")->getValue();
				$emaile 	= $sheet->getCell("D$row")->getValue();
				$gender 	= $sheet->getCell("E$row")->getValue();
				$salary 	= $sheet->getCell("F$row")->getValue();
				$department = $sheet->getCell("G$row")->getValue();
				$date 		= $sheet->getCell("H$row")->getValue();
				// echo $staffcode;
				// echo "<br>";
				$order++;
				$salaryTotal += $salary; 
				$salaries[] = $salary;  // thêm 1 phần tử lương vào biến $sallaries[] chạy 3 lần thì biến $sallaries[] chứa 3 phần tử

				?>
				<!-- cắt từ <tr> đến </tr>(3) -->
				<tr>
			      <td><?=$order?></td>
			      <td><?=$staffcode; ?></td>
			      <td><?=$name ?></td>
			      <td><?=$mobile ?></td>
			      <th><?=$emaile ?></th>
			      <td><?=$gender ?></td>			    
			      <td><?=number_format($salary) ?>đ</td>
			      <td><?=$department ?></td>
			      <td><?=$date ?></td>			     
			    </tr>
				<?php
			}
			 ?>
				 <tr>
			      <td> Số Lượng:</td>
			      <td><?=$order?></td>
			      <td><?=$end - 1?></td>
			      <td></td>
			      <th></th>
			      <td></td>
			      <td></td>
			      <td></td>		     
			    </tr>
				  <tr>
			      <td> Tổng Lương:</td>
			      <td></td>
			      <td></td>
			      <td></td>
			      <th></th>
			      <td></td>
			      <td><?= number_format($salaryTotal)?>đ</td>
			      <td></td>
			      <td></td>
			     	     
			    </tr>
			     <tr>
			     	<?php 
			     	$highestSalary = max($salaries);
			     	?>
			      <td> Lương cao nhất:</td>
			      <td></td>
			      <td></td>
			      <td></td>
			      <th></th>
			      <td></td>
			      <td><?=number_format($highestSalary)?>đ</td>
			      <td></td>
			      <td></td>
			    </tr>
			    <tr>
			    	<?php 
			     	$lowestSalary = min($salaries);
			     	?>
			      <td> Lương Thấp Nhất:</td>
			      <td></td>
			      <td></td>
			      <td></td>
			      <th></th>
			      <td></td>
			      <td><?=number_format($lowestSalary)?>đ</td>
			      <td></td>
			      <td></td>		     
			    </tr>
			    <tr>
			    	<?php 
			    	// round là làm tròn
			    	$averageSalary = round($salaryTotal / $order /1000) * 1000;
			    	?>
			      <td> Lương Trung Bình:</td>
			     <td></td>
			      <td></td>
			      <td></td>
			      <th></th>
			      <td></td>
			      <td><?=number_format($averageSalary)?>đ</td>
			      <td></td>
			      <td></td>		     
			    </tr>
			    <!-- cắt từ  </tbody> đến hết(2) -->
 		  </tbody>
		</table>
	</div>
</div>
	
	<script type="text/javascript" src="public/vendor/jquery-3.5.1.min.js"></script>
	<script type="text/javascript" src="public/vendor/bootstrap-4.5.3-dist/js/bootstrap.min.js"></script>
	<script src="public/js/script.js"></script>
</body>
</html>