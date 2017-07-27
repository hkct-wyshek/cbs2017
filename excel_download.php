<?php
session_start();
date_default_timezone_set('GMT+0');


require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";

require_once 'PHPExcel.php';
PHPExcel_Settings::setZipClass(PHPExcel_Settings::PCLZIP);


if(!getAuthorised(2))
{
	showAccessDenied($day, $month, $year, $area);
	exit();
}


$objPHPExcel = new PHPExcel;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");

$currencyFormat = '#,#0.## \€;[Red]-#,#0.## \€';

$numberFormat = '#,#0.##;[Red]-#,#0.##';
$objSheet = $objPHPExcel->getActiveSheet();

// rename the sheet
$objSheet->setTitle('Booking List');
$objSheet->getStyle('A1:I1')->getFont()->setBold(true);

$col   = array('A','B','C','D','E','F','G','H','I');
$title = array("開始日期\rYYYY-MM-DD",
		       "結束日期\rYYYY-MM-DD",
			   "開始時間\rhh:mm",
			   "結束時間\rhh:mm",
			   "重覆的星期\r(e.g逢星期一, 四重複 = 14)\rp.s.星期日為0",
			   "校舍",
			   "使用單位",
			   "班別編號\r(長度不可以超過80字母)",
			   "備註");

$content1 = array("2016-07-06", 
				  "", 
				  "14:00", 
				  "17:00", 
				  "", 
				  "佐敦培訓中心 11樓課室-1", 
				  "毅進文憑 (兼讀制)", 
				  "", 
				  "");
$content2 = array("2016-07-04", 
				  "2016-07-05", 
				  "09:30", 
				  "13:30", 
				  "1", 
				  "佐敦培訓中心 11樓課室-1", 
				  "其他", 
				  "中學到訪活動 0930-1330", 
				  "");
$content3 = array("2016-07-13", 
				  "2016-07-31", 
				  "14:00", 
				  "17:00", 
				  "4", 
				  "佐敦培訓中心 11樓課室-1", 
				  "應用學習", 
				  "", 
				  "no.");

for($i = 0; $i < count($col); $i++){
	$objSheet->getCell($col[$i]."1")->setValue($title[$i]);
	$objPHPExcel->getActiveSheet()->getStyle($col[$i]."1")->getAlignment()->setWrapText(true);
	$objSheet->getCell($col[$i]."2")->setValue($content1[$i]);
	$objSheet->getCell($col[$i]."3")->setValue($content2[$i]);
	$objSheet->getCell($col[$i]."4")->setValue($content3[$i]);
}

$objSheet->getColumnDimension('A')->setWidth(20);
$objSheet->getColumnDimension('B')->setWidth(20);
$objSheet->getColumnDimension('C')->setWidth(15);
$objSheet->getColumnDimension('D')->setWidth(15);
$objSheet->getColumnDimension('E')->setWidth(30);
$objSheet->getColumnDimension('F')->setWidth(30);
$objSheet->getColumnDimension('G')->setWidth(30);
$objSheet->getColumnDimension('H')->setWidth(30);
$objSheet->getColumnDimension('I')->setWidth(30);

$objSheet->getStyle('A1:A1000')->getNumberFormat()->setFormatCode('yyyy-mm-dd');
$objSheet->getStyle('B1:B1000')->getNumberFormat()->setFormatCode('yyyy-mm-dd');
$objSheet->getStyle('C1:C1000')->getNumberFormat()->setFormatCode('h:mm');
$objSheet->getStyle('D1:D1000')->getNumberFormat()->setFormatCode('h:mm');

$objWorkSheet = $objPHPExcel->createSheet();  
$objPHPExcel->setActiveSheetIndex(1); 
$objSheet = $objPHPExcel->getActiveSheet();
$objSheet->setTitle('Campus');
$row = 1; 
if (!empty($_SESSION['campus'])){
	$campus = $_SESSION['campus'];
	unset($_SESSION['campus']);
	for($i = 0; $i < count($campus); $i++){
		$objSheet->getCell('A'.($i+1))->setValue($campus[$i]);
		$row +=1 ;
	}
}
			
$objPHPExcel->addNamedRange( 
	new PHPExcel_NamedRange(
		'Campus', 
		$objPHPExcel->setActiveSheetIndex(1), 
		'A1:A'.$row
	) 
);


$objWorkSheet = $objPHPExcel->createSheet();  
$objPHPExcel->setActiveSheetIndex(2); 
$objSheet = $objPHPExcel->getActiveSheet();
$objSheet->setTitle('Course');
$row2 = 1;
if (!empty($_SESSION['course'])){
	$course = $_SESSION['course'];
	for ($i = 0; $i < count($course); $i++){
		$objSheet->SetCellValue('A'.($i+1), $course[$i]);
		$row2 += 1;
	}
}

$objPHPExcel->addNamedRange( 
	new PHPExcel_NamedRange(
		'Course', 
		$objPHPExcel->setActiveSheetIndex(2), 
		'A1:A'.$row2
	) 
);

$objPHPExcel->setActiveSheetIndex(0);
			
			$sheetname = array(
							   array('type'=>'Campus','colname'=>'F'),
							   array('type'=>'Course','colname'=>'G')  
							);
			foreach ($sheetname as $key => $val){
				for ($i = 2; $i <= 1000; $i++){
					$objValidation = $objPHPExcel->getSheet(0)->getCell($val['colname']."$i")->getDataValidation();
					$objValidation->setType( PHPExcel_Cell_DataValidation::TYPE_LIST );
					$objValidation->setErrorStyle( PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
					$objValidation->setAllowBlank(false);
					$objValidation->setShowInputMessage(true);
					$objValidation->setShowErrorMessage(true);
					$objValidation->setShowDropDown(true);
					$objValidation->setErrorTitle('Input error');
					$objValidation->setError('Value is not in list.');
					$objValidation->setFormula1("=".$val['type']);
				}
			}


//Setting the header type
ob_end_clean();
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="upload_booking_template.xlsx"');
header('Cache-Control: max-age=0');

$objWriter->save('php://output');


function get_campus_data(){
	
}
?>
