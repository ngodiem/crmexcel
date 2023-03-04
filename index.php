<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>Upload Excel</title>
	<link rel="stylesheet" href="public/vendor/bootstrap-4.5.3-dist/css/bootstrap.min.css">
	
</head>
<body>

	<form action="processUpload.php" class="container-fluid" style="max-width: 500px" enctype="multipart/form-data" method="POST">
		<h1>Upload file excel</h1>
		<div class="input-group mb-3">
			<div class="input-group-prepend">
				<span class="input-group-text">Upload</span>
			</div>
			<div class="custom-file">
				<input type="file" class="custom-file-input" id="excel" name="excel">
				<label class="custom-file-label" for="excel">Choose file</label>
			</div>
		</div>
		<div class="form-group" id="filename">
			
		</div>
		<div class="form-group text-right">
			<button type="submit" class="btn btn-primary">Submit</button>
		</div>
	</form>
	
	<script type="text/javascript" src="public/vendor/jquery-3.5.1.min.js"></script>
	<script type="text/javascript" src="public/vendor/bootstrap-4.5.3-dist/js/bootstrap.min.js"></script>
	<script src="public/js/script.js"></script>
</body>
</html>