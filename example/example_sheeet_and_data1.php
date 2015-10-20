<?php
	set_time_limit(120);
	$sDpath=dirname(__DIR__);
	require_once $sDpath.'/classes/class.ExcelWork.php';
	$objWork=new ExcelWork();

	$objWork->createSheet(); // creating a sheet
	$objWork->setNameOfSheetCreater("Admin-PLus91"); // giving sheet creater name to file
	$objWork->setColumnWidth(20); // width for columns
	$objWork->setHeaderTextFirst("This is a Example sheet with dummy data");
	
	$aHeaders=array('Sr No','Name','State','City','Pin Code','Mobile No'); 
	
	
	$aExportData=array(array('1','JOHN','UK','London','411028','1234567890'),
				array('2','mark','US','NY','411028','123456789'),
				array('3','Petert','Brasil','Parogue','493661','0123456789'),
				array('4','Michel','','123456','493661',''),
				array('5','Abdul','Pakistan','karachi','411028','0123456789'),
			);
	
	$objWork->setDataAndColumnToExcel($aHeaders,$aExportData);


	$objWork->setRowsInsideSheet(2);
	
	// give color to column , colum & row by pass agrument
	
	$objWork->setColorByColum(1,5,'C01AA1'); // colore direct columns by passing like this
	
	// give file name to your excel
	$objWork->setExcelFilename("MyFirstExcel");
	$objWork->setExcelFileExtension('xlsx'); // by default xls format
	
	$objWork->setHeaderTextAtEnd("This is a Our footer of Excel sheet..");
	$objWork->exportFile();
?>