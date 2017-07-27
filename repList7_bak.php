<?php
# $Id: report.php,v 1.22 2004/04/17 15:28:37 thierry_bo Exp $

require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
include "mrbs_sql.inc";

//error_reporting(E_ALL);
//ini_set('display_errors', TRUE);

ini_set('max_execution_time', 300);
require_once 'PHPExcel.php';

#If we dont know the right date then make it up
if(!isset($day) or !isset($month) or !isset($year))
{
	$day   = date("d");
	$month = date("m");
	$year  = date("Y");
}
if(empty($area))
	$area = get_default_area();

	$user = getUserName();
	if (getUserLevel($user) < 1.9)
	{
		showAccessDenied($day, $month, $year, $area);
		exit();
	}


	# print the page header
	print_header($day, $month, $year, $area);

	if ($mode=="day") { $modeName = "日間"; $modeTitle = "日間";} /* mon - fri 0900 - 1759 */
	if ($mode=="night") { $modeName = "晚間"; $modeTitle = "晚間";} /* mon - fri 1800 - 2200, sat - sun 0900 - 2200 */
	if ($mode=="all") { $modeName = "全日"; $modeTitle = "全日";}

	$from = $_POST["dfrom"];
	$to   = $_POST["dto"];
	$mode = $_POST["mode"];


	$isValidDate = true;
	if (check_date_format($_POST["dfrom"]) == false || check_date_format($_POST["dto"]) == false){
		$isValidDate = false;
		$msg = "請輸入正確的日期格式(YYYY-MM-DD)";
	}

	if (strtotime($_POST['dfrom']) > strtotime(date('Y-m-d')) || strtotime($_POST['dto']) > strtotime(date('Y-m-d'))){
		$isValidDate = false;
		$msg = "所輸入的日期不能大於今天";
	}

	if (!$isValidDate){
		echo "<H2>" . get_vocab("wrong_date_head") . "</H2>";
		echo $msg;

		echo "<p><a href=\"report7.php\">返回上一頁</a></p>";
		include "trailer.inc";
		exit;
	}

	$xcunit = $cunit;
	$campusName = "";
	$location_code = "";
	$xls_campus = "";
	if (isset($_POST["c_code"])){
		$location_code = $_POST["c_code"];
		foreach($cpname as $key => $val){
			if ($_POST["c_code"] == $val['code']){
				$campusName = $campusName."[".$val['name']."] ";
				$xls_campus = $val['name'];
			}
		}
	}


	$objPHPExcel = new PHPExcel;
	$objPHPExcel2 = new PHPExcel;
		
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
	$objWriter2 = PHPExcel_IOFactory::createWriter($objPHPExcel2, "Excel2007");

	$objSheet = $objPHPExcel->getActiveSheet();
	$objSheet2 = $objPHPExcel2->getActiveSheet();


	$objSheet->setTitle($xls_campus);
	$objSheet2->setTitle($xls_campus);

	$objSheet->getCell('A1')->setValue($xls_campus.$modeTitle."課室使用資料");
	$objSheet->mergeCells('A1:E1');

	$objSheet2->getCell('A1')->setValue($xls_campus.$modeTitle."課室使用資料");
	$objSheet2->mergeCells('A1:E1');

	?>

<h2>課室使用率報告 - 搜尋結果</h2>
<h4><font color=green>數據計算需時，可能要稍等4~5分鐘...</font></h4>
<h4><?php echo "[".$modeName."] ".$from."至".$to." ".$campusName ?></h4>

<hr>


<?php


$from_timestamp = strtotime($from);
$to_timestamp   = strtotime($to);

function getNameFromNumber($num) {
    $numeric = ($num - 1) % 26;
    $letter = chr(65 + $numeric);
    $num2 = intval(($num - 1) / 26);
    if ($num2 > 0) {
        return getNameFromNumber($num2) . $letter;
    } else {
        return $letter;
    }
}


function create_rich_text($tmp_xls){
	$objRichText = new PHPExcel_RichText();
	$total_tmp = count($tmp_xls);
	
	foreach ($tmp_xls as $num =>$rowData){
		if ($rowData['isUsedRoom'] == 1){
			$run1 = $objRichText->createTextRun($rowData['val']);
			$run1->getFont()->setColor( new PHPExcel_Style_Color( PHPExcel_Style_Color::COLOR_RED ) );
			if ((--$total_tmp)){
				$objRichText->createText("\n&\n");
			}	
		}else{
			$objRichText->createText("\n".$rowData['val']);
			if ((--$total_tmp)){
				$objRichText->createText("\n&\n");
			}
		}	
	}
	return $objRichText;
}

function cellColor($cells){
    global $objPHPExcel;

    $objPHPExcel->getActiveSheet()->getStyle($cells)->getFill()->applyFromArray(array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => array(
             'rgb' => 'D8E4BC'
        )
    ));
}

$date = $from;
$end_date = $to;
$sql = "SELECT id, room_name FROM $tbl_room WHERE location = '$location_code' ORDER BY sequence";

$res = sql_query($sql);

if (! $res) fatal_error(0, sql_error());

	if (sql_count($res) == 0) {
		echo get_vocab("norooms");
	} else {
		$data_ary = array();
		for ($i = 0; ($row = sql_row($res, $i)); $i++) {
			$tmp = array();
	 		$tmp['id'] = $row[0];
	 		$tmp['room_name'] = $row[1]; 
 			array_push($data_ary, $tmp);
 			$aid[$i] =  $row[0];
		}
}



$isfirst = true;
$week_str = "<td class='title'></td>";
$day_str  = "<td class='title'></td>";

$i = 0;

$curr_year = "";
$curr_month = "";

echo "<table class=\"tbl_report7\">";

$row1 = 5;
$col1 = 2;



foreach($aid as $key =>$val){
	
	$time_slot1 = "<td class='title' style=' font-weight:bold;' ><p>".$data_ary[$key]['room_name']."</p></td>";
	$time_slot2 = "<td class='title'><p></p></td>";
	$time_slot3 = "<td class='title'><p></p></td>";
	
	$objSheet->getCell("A".$row1)->setValue($data_ary[$key]['room_name']);
	$objSheet->getColumnDimension("A")->setAutoSize(true);
	$objSheet2->getColumnDimension("A")->setAutoSize(true);

	$date = $from;
	while (strtotime($date) <= strtotime($end_date)) {
		
		if ($mode=="day") { $start_time = "09:00"; $end_time = "17:59"; } 
		if ($mode=="night") { $start_time = "18:00"; $end_time = "22:00"; } 
		if ($mode=="all") { $start_time = "09:00"; $end_time = "22:00"; }
		
		if ($isfirst){
			$count += 1;
			if (date('m', strtotime($date))!= $curr_month){
				$curr_year = date('Y', strtotime($date));
				$curr_month = date('m', strtotime($date));
				$month_str .= "<td colspan='2' style=' font-weight:bold;'>" . $curr_year. "年" . $curr_month . "月</td>";
				$objSheet->getCell(getNameFromNumber(($col1 == 2 ? $col1-1 : $col1))."2")->setValue($curr_year. "年" . $curr_month . "月");
				$objSheet2->getCell(getNameFromNumber(($col1 == 2 ? $col1-1 : $col1))."2")->setValue($curr_year. "年" . $curr_month . "月");
			}else{
				$month_str .= "<td></td>";
			}
			$week_str .= "<td style=' font-weight:bold;'>" . date('D', strtotime($date)) . "</td>";
			$day_str  .= "<td style=' font-weight:bold;'>" . date('d', strtotime($date)) . "</td>";
			
			$objSheet->getCell(getNameFromNumber($col1)."3")->setValue(date('D', strtotime($date)));
			$objSheet->getCell(getNameFromNumber($col1)."4")->setValue(date('d', strtotime($date)));
			$objSheet->getStyle(getNameFromNumber($col1)."4")->getAlignment()->applyFromArray(
			    array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
			);
			
			$objSheet2->getCell(getNameFromNumber($col1)."3")->setValue(date('D', strtotime($date)));
			$objSheet2->getCell(getNameFromNumber($col1)."4")->setValue(date('d', strtotime($date)));
			$objSheet2->getStyle(getNameFromNumber($col1)."4")->getAlignment()->applyFromArray(
			    array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
			);
		}
		
		
		
		if ($mode=="night" && (date('w', strtotime($date)) == 0 || date('w', strtotime($date)) == 6)){
			$start_time = "09:00";
			$end_time   = "22:00";
		}
	
		$tmp_start = strtotime("$date $start_time");
		$tmp_end   = strtotime("$date $end_time");
		
		$morning_end = strtotime("$date 09:00");
		$afternoon_end = strtotime("$date 12:00");
		$evening_end = strtotime("$date 18:00");
		
		if (($mode == "day" && date('w', strtotime($date)) == 0) || ($mode == "day" && date('w', strtotime($date)) == 6)){
			$time_slot1 .= "<td><p></p></td>";
			$time_slot2 .= "<td><p></p></td>";
			$time_slot3 .= "<td><p></p></td>";
			
		}else {
			$sql = "SELECT type, isUsedRoom, start_time, end_time FROM $tbl_entry WHERE room_id=$val AND start_time >= $tmp_start AND start_time <= $tmp_end ORDER BY start_time";
			$res = sql_query($sql);
			
			if (! $res) fatal_error(0, sql_error());
		
			if (sql_count($res) == 0) {
				$time_slot1 .= "<td><p></p></td>";
				$time_slot2 .= "<td><p></p></td>";
				$time_slot3 .= "<td><p></p></td>";
			} else {
				
				$tmp_ary1 = array();
				$tmp_ary2 = array();
				$tmp_ary3 = array();
				
				$tmp = array();
				$tmp_xls_val1 = array();
				$tmp_xls_val2 = array();
				$tmp_xls_val3 = array();
				
				for ($i = 0; ($row = sql_row($res, $i)); $i++) {
					$type_str = "<p style='".($row[1] == 1 ? "color: #ED1C24; background: #FFF200;"  : "")."'>".$row[0]."<br>(".date('H:i', $row[2])."-".date('H:i', $row[3]).")<p>";
					$tmp['isUsedRoom'] = $row[1];
					$tmp['val'] 	   = $row[0] . "(".date('H:i', $row[2])."-".date('H:i', $row[3]).")";
					
					if ($row[2] >= $evening_end ){
						array_push($tmp_ary3, $type_str);
						array_push($tmp_xls_val3, $tmp);
						
					}elseif($row[2] >= $afternoon_end){
						array_push($tmp_ary2, $type_str);
						array_push($tmp_xls_val2, $tmp);
						
					}elseif($row[2] >= $morning_end){
						array_push($tmp_ary1, $type_str);
						array_push($tmp_xls_val1, $tmp);
					}
				}
				
				$time_slot1 .= (count($tmp_ary1) > 0 ? "<td>".implode(" & ", $tmp_ary1)."</td>" : "<td><p></p></td>");
				$time_slot2 .= (count($tmp_ary2) > 0 ? "<td>".implode(" & ", $tmp_ary2)."</td>" : "<td><p></p></td>");
				$time_slot3 .= (count($tmp_ary3) > 0 ? "<td>".implode(" & ", $tmp_ary3)."</td>" : "<td><p></p></td>");
				
				$objSheet->setCellValue(getNameFromNumber($col1).$row1, create_rich_text($tmp_xls_val1));
				$objSheet->getColumnDimension(getNameFromNumber($col1))->setWidth(25);
				$objSheet->getStyle(getNameFromNumber($col1).$row1)->getAlignment()->setWrapText(true);
				$objSheet->getStyle(getNameFromNumber($col1).$row1)->getAlignment()->applyFromArray(
				    array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
				);
				
				$objSheet->setCellValue(getNameFromNumber($col1).($row1+1), create_rich_text($tmp_xls_val2));
				$objSheet->getStyle(getNameFromNumber($col1).($row1+1))->getAlignment()->setWrapText(true);
				$objSheet->getStyle(getNameFromNumber($col1).($row1+1))->getAlignment()->applyFromArray(
				    array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
				);
				
				$objSheet->getStyle(getNameFromNumber($col1).($row1+2))->getAlignment()->setWrapText(true);
				$objSheet->getStyle(getNameFromNumber($col1).($row1+2))->getAlignment()->applyFromArray(
				    array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
				);
			}
		}
		$date = date ("Y-m-d", strtotime("+1 day", strtotime($date)));
		$col1++;
	}
	
	$row1 += 3;
	$col1 = 2;
	
	if ($isfirst){
		echo "<tr>$month_str</tr>";
		echo "<tr>$week_str</tr>";
		echo "<tr>$day_str</tr>";
		$count += 1;
		$isfirst = false;
	}
	
	echo "<tr>$time_slot1</tr>";
	echo "<tr>$time_slot2</tr>";
	echo "<tr>$time_slot3</tr>";		
}

echo "<tr><td colspan='$count' style='text-align:left;'>(B) booking request; (A) Acutual Utilization</td></tr>";

$objSheet->getCell("A".$row1)->setValue("(B) booking request; (A) Acutual Utilization");
$objSheet->mergeCells("A".$row1.":E".$row1++);

$col1 = 1;
$row2 = 5;

foreach ($cunit as $key => $val){
	$b = "<td>$val (B)</td>";
	$a = "<td>$val (A)</td>";
	$rate = "<td>$val (%)</td>";
	$date = $from;

	$objSheet->getCell(getNameFromNumber($col1).$row1)->setValue($val."(B)");
	$objSheet->getCell(getNameFromNumber($col1).($row1+1))->setValue($val."(A)");
	$objSheet->getCell(getNameFromNumber($col1).($row1+2))->setValue($val."(%)");
	cellColor(getNameFromNumber($col1).($row1+2));
	
	$objSheet2->getCell(getNameFromNumber($col1++).($row2))->setValue($val."(%)");
	
	while (strtotime($date) <= strtotime($end_date)) {
		
		
		if ($mode=="day") { $start_time = "09:00"; $end_time = "17:59"; } 
		if ($mode=="night") { $start_time = "18:00"; $end_time = "22:00"; } 
		if ($mode=="all") { $start_time = "09:00"; $end_time = "22:00"; }
		
		if ($mode=="night" && (date('w', strtotime($date)) == 0 || date('w', strtotime($date)) == 6)){
			$start_time = "09:00";
			$end_time   = "22:00";
		}
	
		$tmp_start = strtotime("$date $start_time");
		$tmp_end   = strtotime("$date $end_time");
		
		if (($mode == "day" && date('w', strtotime($date)) == 0) || ($mode == "day" && date('w', strtotime($date)) == 6)){
			$b    .= "<td><p></p></td>";
			$a    .= "<td><p></p></td>";
			$rate .= "<td><p></p></td>";
		}else{
			
			$sql = "SELECT count(id) FROM $tbl_entry ".
			   "WHERE room_id IN (" . implode(",", $aid) . ") AND type = '$val' ".
			   "AND start_time >= $tmp_start AND start_time <= $tmp_end";

			$total_booked = sql_query1($sql);
			
			$sql = "SELECT count(id) FROM $tbl_entry ".
					"WHERE room_id IN (" . implode(",", $aid) . ") AND type = '$val' ".
					"AND start_time >= $tmp_start AND start_time <= $tmp_end AND isUsedRoom = 1";
				
			$total_used = sql_query1($sql);
			
			if ($total_booked == 0 && $total_used == 0){
				$total_booked = "";
				$total_used   = "";
			}
			$b    .= "<td>$total_booked</td>";
			$a    .= "<td>$total_used</td>";
			$persentage = round($total_used / $total_booked * 100);
			
			if ($total_used!="" && $total_booked != ""){
				$rate .= "<td>$persentage</td>";
				$objSheet->getCell(getNameFromNumber($col1).$row1)->setValue($total_booked);
				$objSheet->getCell(getNameFromNumber($col1).($row1+1))->setValue($total_used);
				$objSheet->getCell(getNameFromNumber($col1).($row1+2))->setValue($persentage);
				
				$objSheet2->getCell(getNameFromNumber($col1).($row2))->setValue($persentage);
			}else{
				$rate .= "<td><p></p></td>";
			}
		}	
		cellColor(getNameFromNumber($col1).($row1+2));
		$date = date ("Y-m-d", strtotime("+1 day", strtotime($date)));
		$col1++;
	}
	
	$objSheet->getStyle("A1:".getNameFromNumber($col1-1).($row1+2))->applyFromArray(
	    array(
	        'borders' => array(
	            'allborders' => array(
	                'style' => PHPExcel_Style_Border::BORDER_THIN,
	                'color' => array('rgb' => '000000')
	            )
	        )
	    )
	);
	
	$objSheet2->getStyle("A1:".getNameFromNumber($col1-1).($row2))->applyFromArray(
	    array(
	        'borders' => array(
	            'allborders' => array(
	                'style' => PHPExcel_Style_Border::BORDER_THIN,
	                'color' => array('rgb' => '000000')
	            )
	        )
	    )
	);
	
	$row1 += 3;
	$row2++;
	$col1 = 1;	
	echo "<tr>$b</tr>";
	echo "<tr>$a</tr>";
	echo "<tr class='rate_col'>$rate</tr>";
}

echo "</table>";	



$filename = "download/ClassRoomUsageInfoWithProgramme".date('YmdHis').".xlsx";
$filename2 = "download/ClassRoomUsageInfo".date('YmdHis').".xlsx";
$objWriter->save($filename);
$objWriter2->save($filename2);

?>

<br>

<p>如要下載檔案，請右擊以下連結以另存目標。</p>
<A href="<?php echo $filename2 ?>">下載課室使用率報告</A><br>
<A href="<?php echo $filename ?>">下載課室使用率報告(包含課程名稱)</A><br>
<br>
<A href="javascript:history.go(-1)">&nbsp;返回上一頁</A>


<?php     



?>