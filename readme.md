

**PHPExcel Library**
---------------
____________________________________________________________

Excel Library is a library for creating excel sheet. using library different functions we can create an excel file.

**FEATURES**
_______________________________________________________________

 - Create one or more then excel sheets in a file.
 - Set the data fetched from database in to excel sheet.
 - Merge the columns & rows.
 - Download the file directly to the browser or in a path. 

**INSTALLATION**
_____________________________________________________________________
Download the library and include 'class.ExcelWork.php' in your code.

**USAGE**
_____________________________________________________________________
After include the class file in your code you have your Header & Data array , structure for you excel how you want to display. now just start using it's function. create object of ExcelWork class and use function.

**Methods & Example**
_____________________________________________________________________

 1.  *createSheet()* : 
		This function is create a sheet in excel file which is by default index as ‘0’.
		example:

    $objWork=new ExcelWork(); // creating object of included class
    
	$objWork->createSheet();  // it will create sheet

 2. *setTotalSheet() :*
	 If you want to create multiple sheets just pass the number is argument as ‘setTotalSheet(4)’. it will create 5 sheets inside the excel file.
	 

    $objWork->setTotalSheet(5);  //  

 3. *setActiveSheetOfExcel() :* 
	 If you are doing multiple task in multiple sheets , then you need to use this function for individual task in the current sheet. just pass the number of sheet and then after all other function will work on this sheet only , to change again the sheet you can use this.like this ‘setActiveSheetOfExcel(2)’ . indexing starts from 0 to n.
	 

    $objWork->setActiveSheetOfExcel(3); // if you have multiple sheet

 4. *setNameOfSheetCreater(‘any name’) :* 
	 By this function you can give name to sheet who created that this sheet. pass the name is argument.
	 

    $objWork->setNameOfSheetCreater("Peter Parker");

 5. *setTitleForSheet(‘Title’) :*
	 By this function we can set the title for our current working sheet.
	 

    $objWork->setTitleForSheet('TitleSheet');

 6. *setNoOfColumn(5) :* 
	If you have fixed column known then just pass the number of columns in this function.	
	

    $objWork->setNoOfColumn(6);

 7. *setColumnWidthColumn() :*
 By this function we can set the width for particular column. like i written in argument (6,15) the ‘6’ is column index which start from 1 and the ‘15’ is the size for width.
 

    $objWork->setColumnWidthColumn(6,15);

 8. *setColumnWidth(20) :* 
We can use this function to set default width size for columns , by default width is ‘12’ but as i shown we are setting it ‘20’.

    $objWork->setColumnWidth(20);

 9. *excelMergeColumnsByRow() :* 
We can use this function to merge the columns , as i shown just pass these arguments like (1,6,2) here ‘1’ is first column index , ‘6’ is last column index till we want to merge and ‘2’ is the row number. it will merge the column on this particular row. Must to pass the row number.

    $objWork->excelmergeColumnsByRow(1,6,2);

 10. *setHeaderNames($array) :* 
Header names means the column names for the sheet , so if you want to set the value in just use this function and pass the array for header names.

     $objWork->setHeaderNames($aHeaders);

 11. *setHeaderTextFirst(‘Header text’) :* 
In our sheet we want a header text means the starting caption or little description in the first row , use this function and pass the String in the argument. it will set it in first row and merge the first row for proper output display.

    $objWork->setHeaderTextFirst("This is a header between this merged column");

 12. *set HeaderText(1,8,3,’Header Text’) :*
 By using this function we can set the header text as per our need like i done in argument (1,8,3,’Header Text’) here ‘1’ is the first column , ‘8’ is last column till text will display , ‘3’ is the row number & ‘text’ the text to set. it will merge also for the given column & row then set this text.
 

    $objWork->setHeaderText(0,7,8,"This is a header between this merged column");

 13. *setHeaderTextAtEnd(‘HeadetText’) :*
This function will set a header Text as FOOTER at the last row of the sheet. Need to pass only $string value. 

    $objWork->setHeaderTextAtEnd("This is a header between this merged column");


 14. *getNoOfCol($DataArray) :* 
By using this function we can get the number of column , passing the data array which we want to insert. it will return the number of column should have in the sheet.

 15. *setRowsInsideSheet(4) :*
 By this function we can give blank rows in sheet for proper output display. need to pass no of row you want  to skip or left blank.
 

    $objWork->setRowsInsideSheet(2);

 20. *setDataToExcel($DataArray) :* 
Here comes the important one by this function we can set values to the columns by passing the array of data.

    $objWork->setDataToExcel($aExportData);

 21. *setDataAndColumnToExce($arrayHeader,$arrayData):*
This function will take two arguments of array one of for header & other is for Data values. by this excel sheet will filled by values.

    $objWork->setDataAndColumnToExcel($aHeaders,$aExportData);

 22. *setExcelFilename(‘filename’) :*
 we can set the file name for our excel file by passing suitable name for it.
 

    $objWork->setExcelFilename("MyFirstExcel");

 24. *setTextColumnBold(68) :*
By this funtion we can set the text as bold for this column only , here it will set on ‘D’ row.

    $objWork->setTextColumnBold(3);

 25. *setTextSingleBold(4,6) :*
By this function you can set the text as bold by giving specific column number & row number. here ‘4’ is column number & ‘6’ is row number. 

    $objWork->setTextSingleBold(4,1); 

 27. *setTextFontSize(15) :* 
By using this function we can set the text font size for the sheet , ‘15’ is size which we are setting.

    $objWork->setTextFontSize(12);

 28. *setExcelFileExtension(‘xlsx’) :*
 By Default the extension for excel file is ‘.xls’ but we can pass the ‘xlsx’ and set these extension to file.
 

    $objWork->setExcelFileExtension('xlsx');

 29. *exportFile() :* 
By using this we will forcefully download the created excel files. as we gives data & file name & extension.

    $objWork->exportFile();

 30. *exportFileWithPath(‘PATH’) :* 
By using this function we can get the excel file in the given path. it return an array of path,full path and file name. if you entered a wrong path then it will create ‘download’ folder inside the current directory and place the file there.

    $aFileDetails=$objWork->exportFileWithPath('/var/www/Test/Excel/download');

**Example**

 1. creating sheet & download it :

        <?php
	    require_once 'classes/class.ExcelWork.php';
	    $objWork=new ExcelWork();
	    $objWork->createSheet();
	    $objWork->setExcelFilename("MyFirstExcel");
		$objWork->setExcelFileExtension('xlsx');
		$objWork->exportFile();
	   ?>

 2. creating sheet & download it in given path :
 

      <?php
	    require_once 'classes/class.ExcelWork.php';
	    $objWork=new ExcelWork();
	    $objWork->createSheet();
	    $objWork->setExcelFilename("MyFirstExcel");
		$objWork->setExcelFileExtension('xlsx');
		$aFileDetails= $objWork->exportFileWithPath('/var/www/Test/Excel/download');
		// $aFileDetails will have path , fullpath & filename in array
	   ?>

 3. Give Header Data to set in sheet :
 

       <?php
	    require_once 'classes/class.ExcelWork.php';
	    $objWork=new ExcelWork();
	    $objWork->createSheet();
		$aHeaders=array('Sr No','Name','State','City','Pin Code','Mobile No'); 
		$objWork->setHeaderNames($aHeaders);
		$objWork->setExcelFilename("MyFirstExcel");
		$objWork->setExcelFileExtension('xlsx');
		$objWork->exportFile();
		?>

 4. Give Header & Row data to set in sheet :
 

	    <?php
		    require_once 'classes/class.ExcelWork.php';
		    $objWork=new ExcelWork();
		    $objWork->createSheet();
			$aHeaders=array('Sr No','Name','State','City','Pin Code','Mobile No'); 
			$aExportData=array(array('1','JOHN','UK','London','411028','1234567890'),
				array('2','mark','US','NY','411028','123456789'),
				array('3','Petert','Brasil','Parogue','493661','0123456789'),
				array('4','Michel','','123456','493661',''),
				array('5','Abdul','Pakistan','karachi','411028','0123456789'),
			);
			$objWork->setDataAndColumnToExcel($aHeaders,$aExportData);
			$objWork->setExcelFilename("MyFirstExcel");
			$objWork->setExcelFileExtension('xlsx');
		$objWork->exportFile();
		?>

 5. Merge Cell , Color in column , Bold Text :
 

       <?php
		   require_once 'classes/class.ExcelWork.php';
		    $objWork=new ExcelWork();
		    $objWork->createSheet();
			$aHeaders=array('Sr No','Name','State','City','Pin Code','Mobile No'); 
			$aExportData=array(array('1','JOHN','UK','London','411028','1234567890'),
				array('2','mark','US','NY','411028','123456789'),
				array('3','Petert','Brasil','Parogue','493661','0123456789'),
				array('4','Michel','','123456','493661',''),
				array('5','Abdul','Pakistan','karachi','411028','0123456789'),
			);
			$objWork->setDataAndColumnToExcel($aHeaders,$aExportData);
			$objWork->setColorByColum(1,3,'C01AA1');
			$objWork->excelmergeColumnsByRow(1,4,8);
			$objWork->setHeaderTextAtEnd("This is a header between this merged column");
			$objWork->setExcelFilename("MyFirstExcel");
			$objWork->setExcelFileExtension('xlsx');
			$objWork->exportFile();
		?>