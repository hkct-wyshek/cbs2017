<?php 

require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";

require_once 'PHPExcel.php';
ini_set('max_execution_time', 1000);

if (empty($area))
{
    $area = get_default_area();
}

if(!getAuthorised(1)){
	showAccessDenied($day, $month, $year, $area);
	exit;
}

if (!isCourseAdmin()){
	showAccessDenied($day, $month, $year, $area);
	exit;
}

if (isset($submit)){
	if ($submit == 0){
		
		$start_date = mktime(0, 0, 0,$month, $day, $year);
		$end_date = mktime(0, 0, 0, $end_month, $end_day, $end_year);
		$today = mktime(0, 0, 0, date('m'), date('d'), date('Y'));
		
		
			
		if(($end_date < $start_date) ){
			
			print_header($day, $month, $year, $area);
			echo "<p>".get_vocab("wrong_date")."</p><br>";
			echo "<a href=\"$HTTP_REFERER\">" . get_vocab("returnprev") . "</a>";
			include "trailer.inc";
			exit;
		}elseif (($start_date < $today) || ($end_date < $today)){
			print_header($day, $month, $year, $area);
			echo "<p>".get_vocab("wrong_date2")."</p><br>";
			echo "<a href=\"$HTTP_REFERER\">" . get_vocab("returnprev") . "</a>";
			include "trailer.inc";
			exit;
		}else{
			
			$objPHPExcel = new PHPExcel;
			
			$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
			
			$objSheet = $objPHPExcel->getActiveSheet();
			$objSheet->setTitle('Course List');
			$objSheet->getStyle('A1:E1')->getFont()->setBold(true);
			
			$styleArray = array(
								'font'  => array(
								'color' => array('rgb' => 'FF0000'),
								));
								
			$objSheet->getStyle('D1:E1')->applyFromArray($styleArray);
			
			$sql = "SELECT start_time, end_time, name, e.description, area_name, room_name ". 
				   "FROM $tbl_entry e, $tbl_area a, $tbl_room r " . 
				   "WHERE e.room_id = r.id AND a.id = r.area_id AND type='".getUserName()."' and e.start_time >= '$start_date' and e.end_time <= '$end_date' AND a.cpname IN('".implode("','", $cp_name)."')";

			$result = mysql_query($sql) or die(mysql_error());
			
			$objSheet->getCell("A1")->setValue(get_vocab("course_xls1"));
			$objSheet->getCell("B1")->setValue(get_vocab("course_xls2"));
			$objSheet->getCell("C1")->setValue(get_vocab("course_xls3"));
			$objSheet->getCell("D1")->setValue(get_vocab("course_xls4"));
			$objSheet->getCell("E1")->setValue(get_vocab("course_xls5"));
			
			$objSheet->getColumnDimension('A')->setWidth(20);
			$objSheet->getColumnDimension('B')->setWidth(20);
			$objSheet->getColumnDimension('C')->setWidth(30);
			$objSheet->getColumnDimension('D')->setWidth(30);
			$objSheet->getColumnDimension('E')->setWidth(30);

			$rowCount = 2; 
			while($row = mysql_fetch_array($result)){ 	
				$objSheet->getCell("A".$rowCount)->setValue(date('Y-m-d H:i', $row['start_time']));
				$objSheet->getCell("B".$rowCount)->setValue(date('Y-m-d H:i', $row['end_time']));
				$objSheet->getCell("C".$rowCount)->setValue($row['area_name']."-".$row['room_name']);
				$objSheet->getCell("D".$rowCount)->setValue($row['name']);
				$objSheet->getCell("E".$rowCount)->setValue($row['description']);
				$rowCount++; 
			}
			
			
			$objSheet->getStyle("A2:A$rowCount")->getNumberFormat()->setFormatCode('yyyy-mm-dd hh:mm');
			$objSheet->getStyle("B2:B$rowCount")->getNumberFormat()->setFormatCode('yyyy-mm-dd hh:mm');
			$objSheet->getStyle("A2:A$rowCount")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objSheet->getStyle("B2:B$rowCount")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

			ob_end_clean();
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header('Content-Disposition: attachment;filename="course_info.xlsx"');
			header('Cache-Control: max-age=0');
			
			$objWriter->save('php://output');

		}	
	}
}


?>