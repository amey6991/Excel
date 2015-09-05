<?php
	include_once 'lib/PHPExcel/PHPExcel.php';
	
	class ExcelWork extends PHPExcel{
		public $iNoOfCol=0;
		public $sFileName='defaultExcelFile';
		public $iRow=1;
		public $iColumn=0;
		protected $sFileExtension='xls';

		// constructor which called whenever the object is created of ExcelWork class
		function __construct() {
			$this->excel = new PHPExcel(); // object created for PHPExcel
		}
		// create a sheet inside the excel worksheet ..
		function createSheet(){
				$this->excel->createSheet(); // a sheet created in excel file
				$this->excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(17);
		}
		// if you have miltiple sheets to create just pass the number of sheet ,
		// it will create that number of sheets.
		function setTotalSheet($iNumber){
				for ($iLoop=1; $iLoop < $iNumber; $iLoop++) { 
					$this->excel->createSheet();	// creating sheets..
				}
		}
		function getActiveSheetIndexValue(){
			$sheet=$this->excel->getActiveSheet();
				// echo "<pre>";
				// print_r($sheet);
				// exit();
		}
		// it will set the sheet as current sheet in excel file , by just passign the index of sheet
		function setActiveSheetOfExcel($iNumber){
				$this->excel->setActiveSheetIndex($iNumber); // seted the active sheet by passing the index	
		}
		// it will set the sheet create name
		function setNameOfSheetCreater($sName){
			$this->excel->getProperties()->setCreator($sName);
		}
		// it will set the title for the current sheet
		function setTitleForSheet($sTitle){
			$this->excel->getActiveSheet()->setTitle($sTitle);
		}
		// if you have multple sheets in excel file then 
		// it will set the title for the sheet index you given ,
		// just pass the index of sheet & title for it.
		function setTitleBySheet($iSheetIndex,$sTitle){
			$this->excel->setActiveSheetIndex($iSheetIndex);
			$this->excel->getActiveSheet()->setTitle($sTitle);
		}
		// column functions
		// it will set the column number and also create a fixed size array 
		function setNoOfColumn($iColumnNumber){
			$this->iNoOfCol=$iColumnNumber;
			$aColarray=new SplFixedArray($iColumnNumber);
		}
		// it will set the width of the given column , just pass the index of the column and the size
		function setColumnWidthColumn($iColumnValue,$iSize){
			$sColumnChar=$this->getColumnValueChar($iColumnValue);
			$this->excel->getActiveSheet()->getColumnDimension("$sColumnChar")->setWidth($iSize);
		}
		// it will set the header names to the sheet 
		function setHeaderNames($aColNames){
			foreach ($aColNames as $iCol => $sValue) {
				$this->excel->getActiveSheet()->setCellValueByColumnAndRow($iCol,$this->iRow,$sValue);
			}
			$this->iRow++;
		}
		// seted width for all column
		function setColumnWidth($iWid){
			$this->excel->getActiveSheet()->getDefaultColumnDimension('A:Z')->setWidth($iWid);
		}
		// it will merge the columns by row in sheet , pass the start , end column & row number
		function excelMergeColumnsByRow($iFirstColumn,$iLastColumn,$iRoww){
			$sFirstColumnChar=$this->getColumnValueChar($iFirstColumn);
			$sFirstColumnChar=$sFirstColumnChar.$iRoww;
			$sLastColumnChar=$this->getColumnValueChar($iLastColumn);
			$sLastColumnChar=$sLastColumnChar.$iRoww;
			$this->excel->getActiveSheet()->mergeCells("$sFirstColumnChar:$sLastColumnChar");
		}
		// merging the column by passing start & end column
		function excelMergeColumns($iFirstColumn,$iLastColumn){
			$sFirstColumnChar=$this->getColumnValueChar($iFirstColumn);
			$sFirstColumnChar=$sFirstColumnChar.$this->iRow;
			$sLastColumnChar=$this->getColumnValueChar($iLastColumn);
			$sLastColumnChar=$sLastColumnChar.$this->iRow;
			$this->excel->getActiveSheet()->mergeCells("$sFirstColumnChar:$sLastColumnChar");
		}
		// it will return the highest column index in the sheet
		function getHighestColumnNumber(){
			return $this->excel->setActiveSheetIndex(0)->getHighestColumn();
		}
		// get no of columns from the data array you passed.
		function getNoOfCol($aExportData){
			$iCountt=0;
		foreach ($aExportData as $iKey => $sValue) {
			foreach ($sValue as $iKey1 => $sValue1) {
					$iCountt++;
				}
				return $iCountt;
				break;
			}
		}
		// it returns the column index , argument shoul be in A..Z , it return the
		// index value of that column 
		function getColumnValueNumber($columnstring){
					$iVal=ord($columnstring);
					$iDiff=$iVal-64;
					return $iDiff;
		}
		// return column name in characters from A..Z 
		function getColumnValueChar($iIndex){
			$sChars=$iIndex+64;
			return chr($sChars);
		}
		// it will set color by given columns , argument should be index of start & end column 
		function setColorByColum($iFirstColumn,$iLastColumn,$colorcode){
			$sFirstColumn=$this->getColumnValueChar($iFirstColumn);
			$sLastColumn=$this->getColumnValueChar($iLastColumn);
			$this->excel->getActiveSheet()->getStyle("$sFirstColumn:$sLastColumn")->applyFromArray(
		        array(
		            'fill' => array(
		                'type' => PHPExcel_Style_Fill::FILL_SOLID,
		                'color' => array('rgb' => "$colorcode")
		           )));
		}
		// it will set color in given colum & row , just pass the first column , first row , last column
		// last row & color code.. 
		function setColorByColumRow($iFirstColumn,$iFirstRow,$iLastColumn,$iLastRow,$colorcode){
			$iFirstColumn=$this->getColumnValueChar($iFirstColumn);$iFirstColumn=$iFirstColumn.$iFirstRow;
			$iLastColumn=$this->getColumnValueChar($iLastColumn);$iLastColumn=$iLastColumn.$iLastRow;
			$this->excel->getActiveSheet()->getStyle("$iFirstColumn:$iLastColumn")->applyFromArray(
		        array(
		            'fill' => array(
		                'type' => PHPExcel_Style_Fill::FILL_SOLID,
		                'color' => array('rgb' => "$colorcode")
		           )));
		}
		// it will set color in column you set , just pass the column index & color code
		function setColorByColumn($iColumn,$sColorCode){
			$sColumnValue=$this->getColumnValueChar($iColumn);
			$this->excel->getActiveSheet()->getStyle("$sColumnValue")->applyFromArray(
		        array(
		            'fill' => array(
		                'type' => PHPExcel_Style_Fill::FILL_SOLID,
		                'color' => array('rgb' => "$sColorCode")
		           )));
		}

		// set Header text in the first row
		function setHeaderTextFirst($sString){
			$this->excelMergeColumnsByRow(1,12,$this->iRow);
			//$this->excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$this->excel->getActiveSheet()->setCellValueByColumnAndRow(0,$this->iRow,$sString);
			$this->iRow++;
		}
		// set the header Text by passing the column & row number and the text
		function setHeaderText($iFirstColumn,$iLastColumn,$iRoww,$sString){
			$this->excel->getActiveSheet()->setCellValueByColumnAndRow($iFirstColumn,$iRoww,$sString);
			$this->iRow=$iRoww;
			$this->iRow++;
		}
		// it will set Header Text at the end of the sheet..
		function setHeaderTextAtEnd($sString){
				$this->excelMergeColumnsByRow(1,15,$this->iRow);
				$sRow='A'.$this->iRow;
				$this->excel->getActiveSheet()->getStyle("$sRow")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$this->excel->getActiveSheet()->setCellValueByColumnAndRow(0,$this->iRow,$sString);
				$this->iRow++;
		}
		// it skips row or create blank rows in sheet 
		function setRowsInsideSheet($iNumber){
			for ($iLoop=0; $iLoop <$iNumber ; $iLoop++) { 
				$this->iRow++;
			}
		}
		// fill data into excel file by passing the array of header & data ..
		function setDataAndColumnToExcel($aHeader,$aData){
			foreach ($aHeader as $iCol => $sValue) {
				$this->excel->getActiveSheet()->setCellValueByColumnAndRow($iCol,$this->iRow,$sValue);
			}
			// $this->excel->getActiveSheet()->getStyle('A1:Z1')->getFont()->setBold(true);
			$this->iRow++;
			foreach ($aData as $iKey => $aValue) {
				foreach ($aValue as $iKey1 => $sValue1) {
					$this->excel->getActiveSheet()->setCellValueByColumnAndRow($iKey1,$this->iRow,$sValue1);	
				}
				$this->iRow++;
			}
		}
		// if there is no Header or already seted then it will set your excel data only.
		function setDataToExcel($aData){
			foreach ($aData as $iKey => $aValue) {
				foreach ($aValue as $iKey1 => $sValue1) {
					$this->excel->getActiveSheet()->setCellValueByColumnAndRow($iKey1,$this->iRow,$sValue1);	
				}
				$this->iRow++;
			}
		}

		// giving filename to excel file
		function setExcelFilename($sFilenamee){
				$this->sFileName=$sFilenamee;
		}
		// it return the filename for excel
		function getExcelFileName(){
			return $this->sFileName;
		}

		// text style functions..

		// pass the column index which text want to set as bold
		function setTextColumnBold($iColumn){
			$sColumn=$this->getColumnValueChar($iColumn);
			$this->excel->getActiveSheet()->getStyle("$sColumn")->getFont()->setBold(true);
		}
		// pass the index of first column & row number to highlight the value as bold text
		function setTextSingleBold($iFirstColumn,$iFirstRow){
			$sFirstColumn=$this->getColumnValueChar($iFirstColumn);
			$sFirstColumn=$sFirstColumn.$iFirstRow;
			$this->excel->getActiveSheet()->getStyle("$sFirstColumn")->getFont()->setBold(true);
		}
		
		// setting text as bold for all column
		function setTextBold(){
			$this->excel->getActiveSheet()->getStyle('A:Z')->getFont()->setBold(true);
		}

		// set the specific column & row for text bold in our data sheet 
		function setTextBoldAccor($iFirstColumn,$iFirstRow,$iLastColumn,$iLastRow){
			$sFirstColumn=$this->getColumnValueChar($iFirstColumn);
			$sFirstColumn=$sFirstColumn.$iFirstRow;
			$sLastColumn=$this->getColumnValueChar($iLastColumn);
			$sLastColumn=$sLastColumn.$iLastRow;
			$this->excel->getActiveSheet()->getStyle("$sFirstColumn:$sLastColumn")->getFont()->setBold(true);
		}
		// setting font size for all the column
		function setTextFontSize($iSize){
			$this->excel->getActiveSheet()->getStyle('A:Z')->getFont()->setSize($iSize);
			$this->excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight($iSize+3);

		}
		// it set the file extension you specify
		function setExcelFileExtension($sFileex){
			$this->sFileExtension=$sFileex; 
		}

		// it download the file in the given path if wrong path given it create a download folder
		// inside the current directory and place the file there..
		function exportFileWithPath($sPath){
			$sPath=realpath($sPath);
			if(empty($sPath)){
					$sDIRPath=dirname(__DIR__);
					if(is_dir($sDIRPath.'/downloads')){
						$sFullPath=$sDIRPath.'/downloads/';
						$this->sFileName=$this->getExcelFileName().'.'.$this->sFileExtension;
						$objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
						$objWriter->save($sFullPath.'/'.$this->sFileName);
						$aFileDetail=array('filename'=>$this->sFileName,'path'=>$sFullPath,'fullpath'=>$sFullPath.'/'.$this->sFileName);
						return $aFileDetail;
						exit;
					}else{
						mkdir($sDIRPath.'/downloads/',0777);
						$sFullPath=$sDIRPath.'/downloads/';
						$this->sFileName=$this->getExcelFileName().'.'.$this->sFileExtension;
						$objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
						$objWriter->save($sFullPath.'/'.$this->sFileName);
						$aFileDetail=array('filename'=>$this->sFileName,'path'=>$sFullPath,'fullpath'=>$sFullPath.'/'.$this->sFileName);
						return $aFileDetail;
						exit;
					}
				
			}else{
				if($this->sFileExtension=='xls'){
						$this->sFileName=$this->getExcelFileName().'.'.$this->sFileExtension;
						$objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
						$objWriter->save($sPath.'/'.$this->sFileName);
						$aFileDetail=array('filename'=>$this->sFileName,'path'=>$sPath,'fullpath'=>$sPath.'/'.$this->sFileName);
						return $aFileDetail;
						exit;
				}elseif ($this->sFileExtension=='xlsx') {
						$this->sFileName=$this->getExcelFileName().'.'.$this->sFileExtension;
						$objWriter = new PHPExcel_Writer_Excel2007($this->excel);
						$objWriter->save($sPath.'/'.$this->sFileName);
						$aFileDetail=array('filename'=>$this->sFileName,'path'=>$sPath,'fullpath'=>$sPath.'/'.$this->sFileName);
						return $aFileDetail;
						exit;
				}
			}
		}
		// it download the file in default download folder & force to browser to download file 
		function exportFile(){
			if($this->sFileExtension=='xls'){
				header('Content-type: application/vnd.ms-excel');
				$this->sFileName=$this->getExcelFileName().'.'.$this->sFileExtension;
				header('Content-Disposition: attachment; filename='.$this->sFileName);
				$objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
				$objWriter->save('php://output');
				exit;

			}elseif ($this->sFileExtension=='xlsx') {
				header('Content-type: application/vnd.ms-excel');
				$this->sFileName=$this->getExcelFileName().'.'.$this->sFileExtension;
				header('Content-Disposition: attachment; filename='.$this->sFileName);
				$objWriter = new PHPExcel_Writer_Excel2007($this->excel);
				$objWriter->save('php://output');
				exit;
			}
		}
	}
		
?>