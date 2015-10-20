<?php
	set_time_limit(120);
 	$sDpath=dirname(__DIR__);
	include_once $sDpath.'/classes/class.ExcelWork.php';

	$objWork=new ExcelWork();

	$iSheet=5;
	$objWork->setTotalSheet($iSheet); // creating a sheet
	

	for($iKey=0;$iKey < $iSheet;$iKey++){
		$objWork->setActiveSheetOfExcel($iKey);
		// Use only for multiple sheets..
		$objWork->setInitialRowValue();
		$objWork->setColumnWidth(20);
		
		$objWork->setHeaderTextFirst("This is a Example sheet with dummy data");

		$aHeaders=array('Sr No','Name','State','City','Pin Code','Mobile No'); 
		$objWork->setHeaderNames($aHeaders);
	
	$aExportData=array(array('1','JOHN','UK','London','411028','1234567890'),
				array('2','mark','US','NY','411028','123456789'),
				array('3','Petert','Brasil','Parogue','493661','0123456789'),
				array('4','Michel','','123456','493661',''),
				array('5','Abdul','Pakistan','karachi','411028','0123456789'),
			);
	
	$objWork->setDataToExcel($aExportData);

	}

	
	
	
	

	//$objWork->setRowsInsideSheet(2);
	
	
	
	
	
	// give file name to your excel
	$objWork->setExcelFilename("MyFirstExcel");
	$objWork->setExcelFileExtension('xlsx'); // by default xls format
	$objWork->setActiveSheetOfExcel(0);
	$objWork->setHeaderTextAtEnd("This is a Our footer of Excel sheet..");
	$objWork->exportFile();
?>