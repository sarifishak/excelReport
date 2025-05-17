<?php

if(!empty($_FILES["fileToUpload"]["name"])){
    $target_dir = "uploads/";
    //$target_file = $target_dir . basename($_FILES["fileToUpload"]["name"]);
    $target_file = $target_dir . 'latestReport.xlsx';
    $uploadOk = 1;
    $imageFileType = strtolower(pathinfo($target_file,PATHINFO_EXTENSION));
    
    // Check if file already exists
    if (file_exists($target_file)) {
      //echo "Sorry, file already exists.";
      //$uploadOk = 0;
      unlink($target_file);
    }
    
    if($_FILES["fileToUpload"])
    
    // Check file size
    if ($_FILES["fileToUpload"]["size"] > 500000) {
      echo "Sorry, your file is too large.";
      $uploadOk = 0;
    }
    
    // Allow certain file formats
    if($imageFileType != "xlsx" && $imageFileType != "xlsx") {
      echo "Sorry, only excel files are allowed.";
      $uploadOk = 0;
    }
    
    // Check if $uploadOk is set to 0 by an error
    if ($uploadOk == 0) {
      echo "Sorry, your file was not uploaded.";
    // if everything is ok, try to upload file
    } else {
      if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $target_file)) {
        //echo "The file ". htmlspecialchars( basename( $_FILES["fileToUpload"]["name"])). " has been uploaded.";
        header("Location: check_valid_latest_report.php");
        die();
      } else {
        echo "Sorry, there was an error uploading your file.";
      }
    }    
} else {
?>
<!DOCTYPE html>
<html>
<body>

<form action="uploadLatestReport.php" method="post" enctype="multipart/form-data">
  Select latest report to upload:
  <input type="file" name="fileToUpload" id="fileToUpload">
  <input type="submit" value="Upload Now" name="submit">
</form>

</body>
</html>

<?php
}
?>


