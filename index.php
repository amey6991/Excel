<?php
	set_time_limit(120);
	require_once 'classes/class.ExcelWork.php';
	$con=mysqli_connect('localhost','root','','amey_medixcel_basic');
	$aExportData=array();
	if($con){
		$q='SELECT * FROM mxcel_item_rates';
		$sQuery=mysqli_query($con,$q);
		while($row=mysqli_fetch_row($sQuery)){
			$aExportData[]=$row;
		}
		//echo "<pre>";print_r($aExportData);
	}
	
	//exit();
	$objWork=new ExcelWork();

	$objWork->createSheet(); // creating a sheet
	//$objWork->setTotalSheet(5); // creating multiple sheet by pass no of sheets
	
	//$objWork->getActiveSheetIndexValue();
	
	$objWork->setNameOfSheetCreater("Amey Damle"); // giving sheet creater name to file
	//$objWork->setActiveSheetOfExcel(3); // sheet 3 is active now , bydefault is 0
	//$objWork->setTitleForSheet('TitleSheet');
	//$objWork->setTitleBySheet(1,'TitleSheet-1');
	//$objWork->setTitleBySheet(2,'TitleSheet-2');
	//$objWork->setNoOfColumn(6); // setting number of column gonna be
	$objWork->setColumnWidth(20); // width for columns
	//$objWork->setColumnWidthColumn(6,15); // set width by colum & size given in parameters
	
	$objWork->setHeaderTextFirst("This is a header between this merged column");
	$aHeaders=array('item_rate_id','item_id','item_rate_type_id','parent_item_id','rate','contrast_charge','procedure_charge','added_by','added_on','status','is_valid');
	//$aHeaders=array('Sr No','Name','State','City','Pin Code','Mobile No'); // hardcoded colNames for file
	//$objWork->setHeaderNames($aHeaders); // setted names to the column
	
	// $aExportData=array(array('1','amey','Maharashtra','Pune','411028','9926134428'),
	// 	array('2','AMIT','Maharashtra','Raipur','411028','9926134428'),
	// 	array('3','Subhojit','GOA','BILASPUR','493661','8412052144'),
	// 	array('4','Subhojit','','BILASPUR','493661',''),
	// 	array('5','Sumit','BIHAR','Patna','411028','9926134428'),
	// 	);
	
	$objWork->setDataAndColumnToExcel($aHeaders,$aExportData); // according to data will set the column but data should be constant in array
	//$objWork->setDataToExcel($aExportData);


	//$objWork->setRowsInsideSheet(2);
	
	// give color to column , colum & row by pass agrument
	
	//$objWork->setColorByColum(1,5,'C01AA1'); // colore direct columns by passing like this
	//$objWork->setColorByColumRow(1,1,3,2,'C01AA1'); // color colum & row by passign argument
	//$objWork->setColorByColumn(2,'C01AA1');
	
	//$objWork->setRowsInsideSheet(2); // it will increment the row counter 2 times

	// give file name to your excel
	$objWork->setExcelFilename("MyFirstExcel");
	$objWork->setExcelFileExtension('xlsx'); // by default xls format
	//$objWork->getNoOfCol($aExportData);
	// text style in excel sheet
	//$objWork->setTextBold(); // make all text as bold , to according you pass argument
	//$objWork->setTextBoldAccor(1,1,26,1);
	//$objWork->setTextSingleBold(4,1);  // make your colum & row bold
	$objWork->setTextColumnBold(3);    // make any single column bold by passign column index
	$objWork->setTextBoldAccor(1,1,26,1); // argument as f_col,f_row to l_col,l_row
	$objWork->setTextFontSize(12); // giving font size to all column & row
	//$objWork->excelmergeColumnsByRow(1,4,7); // pass the first starting column & last column to be merge

	//$objWork->setHeaderText(0,7,8,"This is a header between this merged column");
	$objWork->setHeaderTextAtEnd("This is a header between this merged column");
	
	// force to download the file with above activity , if you enter wrong path it will create a download
	// folder inside the directory and place file there..
	$aFileDetails=$objWork->exportFileWithPath('/var/www/Test/Excel/download');
	print_r($aFileDetails);
	//$objWork->exportFile();
?>