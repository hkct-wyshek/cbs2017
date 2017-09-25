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

if ($mode=="day") { $modeName = "全日制"; $modeTitle = "全日制";} /* mon - fri 0900 - 1759 */
if ($mode=="night") { $modeName = "兼讀制"; $modeTitle = "兼讀制";} /* mon - fri 1800 - 2200, sat - sun 0900 - 2200 */
if ($mode=="all") { $modeName = "全日制 和兼讀制"; $modeTitle = "全日制 和兼讀制";}

$from = $_POST["dfrom"]."-01";
$mode = $_POST["mode"];

$isValidDate = true;
if (check_date_format($from) == false){
	$isValidDate = false;
	$msg = "請輸入正確的日期格式(YYYY-MM)";
}else{
	$to = $_POST["dfrom"]."-".date('t',strtotime($from));
}

if (!$isValidDate){
	echo "<H2>" . get_vocab("wrong_date_head") . "</H2>";
	echo $msg;
	
	echo "<p><a href=\"report8.php\">返回上一頁</a></p>";
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
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");

$objSheet = $objPHPExcel->getActiveSheet();
$objSheet->getDefaultColumnDimension()->setWidth(15);

function create_rich_text($tmp_xls){
	$objRichText = new PHPExcel_RichText();
	$run1 = $objRichText->createTextRun($tmp_xls);
	$run1->getFont()->setColor( new PHPExcel_Style_Color( PHPExcel_Style_Color::COLOR_RED ) );
	$objRichText->createText();
	
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

function cellBorder($fromCell, $toCell, $sheet){
	global $objPHPExcel;
	$styleArray = array(
			'borders' => array(
					'allborders' => array(
							'style' => PHPExcel_Style_Border::BORDER_THIN
					)
			)
	);
	
	$sheet->getStyle("$fromCell:$toCell")->applyFromArray($styleArray);
}

function cellValue($sheet, $str, $from, $to, $isMerge, $isCenter, $isBold, $col, $cellWidth, $isWrap, $isRichText, $fontSize){

	/*
	 * 1. $sheet    = excelWorkbook
	 * 2. $str      = cell value
	 * 3. $from     = cell index e.g A1
	 * 4. $to       = cell index for merge cell use
	 * 5. $isCenter = cell align center
	 * 6. $isBold   = cell font weight
	 * 6. $col      = cell column e.g. A
	 * 7. $cellWidth= cell column width
	 * 8. $isWrap   = text line break
	 * 9. $isRichText = Rich Text
	 */

	if (!$isRichText){
		$sheet->getCell($from)->setValue($str);
	}


	if ($isWrap){
		$sheet->getStyle($from)->getAlignment()->setWrapText(true);
	}

	if ($isMerge){
		$sheet->mergeCells("$from:$to");
	}

	if ($isCenter){
		$sheet->getStyle($from)->getAlignment()->applyFromArray(array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER));
	}

	if ($isBold){
		$sheet->getStyle($from)->getFont()->setBold(true);
	}

	if (!empty($cellWidth)){
		$sheet->getColumnDimension($col)->setWidth($cellWidth);
	}
	
	if (!empty($fontSize)){
		$sheet->getStyle($from)->getFont()->setSize($fontSize);
	}
}

function cellValueForRichText($sheet, $str, $from, $to, $isMerge, $isCenter, $isBold, $col,  $cellWidth, $isWrap){
	$sheet->setCellValue($from, create_rich_text($str));
	cellValue($sheet, "", $from, $to, $isMerge, $isCenter, $isBold, $col, $cellWidth, $isWrap, true);
}
function get_section_code($val){
	
	$dd = date('Y-m-d', $val);
	$dd_arr = array(strtotime("$dd 18:00"), strtotime("$dd 16:00"), strtotime("$dd 13:00"), strtotime("$dd 11:00"), strtotime("$dd 09:00"));
	//print_array($dd_arr);
	foreach ($dd_arr as $key => $item){
		
		if ($val >= $item){
			return abs($key - 4);
		}
	}
}

function get_loc_shift_val($loc_id, $section, $dd){
	global $tbl_entry;
	
	switch ($section){
		case 0:
			$starttime = "09:00";
			$endtime = "10:59";
			break;
			//0900-1100
		case 1:
			$starttime = "11:00";
			$endtime = "12:59";
			break;
			//1101-1300
		case 2:
			$starttime = "13:00";
			$endtime = "15:59";
			break;
			//1400-1600
		case 3:
			$starttime = "16:00";
			$endtime = "17:59";
			break;
			//1601-1800
		case 4:
			$starttime = "18:00";
			$endtime = "22:00";
			break;
			//1801-2200
	}
	
	$starttime = strtotime("$dd $starttime");
	$endtime = strtotime("$dd $endtime");
	
	$res = sql_query("SELECT name, start_time, end_time, isUsedRoom, type FROM $tbl_entry WHERE start_time >= '$starttime' and start_time <= '$endtime'  and room_id = $loc_id limit 1");
	if (! $res) fatal_error(0, sql_error());
	if (sql_count($res) == 0) {
		$tmp = array();
		$tmp['name'] = "";
		$tmp['starttime'] = "";
		$tmp['endtime'] = "";
		$tmp['duration'] = 1;
		$tmp['isUsedRoom'] = 0;
		$tmp['type'] = "";
	} else {
		$data_ary = array();
		for ($i = 0; ($row = sql_row($res, $i)); $i++) {
			$tmp = array();
			
			$tmp['starttime'] = date('H:i',$row[1]);
			$tmp['endtime'] = date('H:i',$row[2]);
			$tmp['duration'] = date('H',$row[2]) - date('H',$row[1]);
			$tmp['name'] = ($row[4] == $row[0] ? $row[0]." (".$tmp['starttime']."-".$tmp['endtime'].")": "[" . $row[4] . "] " . $row[0]." (".$tmp['starttime']."-".$tmp['endtime'].")");
			$tmp['isUsedRoom'] = $row[3];
			$tmp['type'] = $row[4];
		}
	}
	return $tmp;
}

function addTBLCell($loc_data, $isSatorSun, $NightClass, $course_colspan){
	
	if ($loc_data["isUsedRoom"] == 1){
		$bgcolor = "color: #ED1C24; background: #FFF200;";
	}else{
		$bgcolor = "";
	}
	
	if ($loc_data["duration"] > 2 && $isSatorSun && !$NightClass){
		$course_colspan = $course_colspan;
	}
	
	//if ($course_colspan > 4) $course_colspan = 4;	
	return "<td colspan=\"$course_colspan\" style=\"$bgcolor\">".$loc_data["name"]. "</td>";
}


function calUsage($course_booked, $course_used){
	//echo "course_booked: $course_booked -  course_used: $course_used<BR>";
	if ($course_booked == '') return '';
	if ($course_booked == 0 || $course_used == 0) return 0;
	if ($course_booked == 0 && $course_used == '' ) return 0;
	
	return round(($course_used / $course_booked) * 100);	
}

$objSheet->setTitle($xls_campus);

$title = $xls_campus.$modeTitle."每月課室使用狀況";

cellValue($objSheet, $title, "A1", "E1", true, false, true, "", "", false, false, 16);
cellValue($objSheet, $from."至".$to, "A2", "E2", true, false, true, "", "", false, false, 14);

?>

<h2><?php echo $title; ?> - 搜尋結果</h2>
<h4><font color=green>數據計算需時，可能要稍等4~5分鐘...</font></h4>
<h4><?php echo "[".$modeName."] ".$from."至".$to." ".$campusName ?></h4>

<hr>


<?php

$from_timestamp = strtotime($from);
$to_timestamp   = strtotime($to);
cellValue($objSheet, date('Y', $from_timestamp).'年'.date('m', $from_timestamp).'月', "A3");

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

function calColspan($current_num, $duration, $ddTime){
	$dd = date('Y-m-d', $ddTime);
	switch ($current_num) {
		case 0:
			if ($duration >= 4) 
				return 4; 
			else 
				return $duration;
			break;
		case 1:
			if ($duration >= 3) 
				return 3; 
			else 
				return $duration;
			break;
		case 2:
			if ($duration >= 2) 
				return 2; 
			else 
				return $duration;
			break;
		case 3 :
			if ($duration >= 1) 
				return 1; 
			else 
				return $duration;
			break;
		case 4 :
			if ($duration >= 1) 
				return 1; 
			else 
				return $duration;
			break;
	}
}

$date = $from;
$end_date = $to;
$sql = "SELECT r.id, r.room_name, a.area_desc FROM $tbl_room r, $tbl_area a WHERE r.location = '$location_code' and r.area_id = a.id ORDER BY r.area_id, r.room_name                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         ";

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
	 		$tmp['area_desc'] = $row[2]; 
 			array_push($data_ary, $tmp);
 			$aid[$i] =  $row[0];
		}
	}

	//$date = $from;
	$loc_val = array();
	
	$course_daliy_total_usage= array();
	$course_total_booked = array();
	$course_total_used = array();
	$total_num = "<tr>";
	
	
	//print_array($course_total_usage);
	$row1 = 3;
	$col1 = 1;
	
	foreach($data_ary as $key => $loc){
		
		cellValue($objSheet, $loc['room_name']." - (".$loc['area_desc'].")", getNameFromNumber($col1).($row1+2), "", false, false, false);		
		$loc_val[$key] = "<tr><td>".$loc['room_name']."<BR>(".$loc['area_desc'].")"."</td>";	
		$row1++;
	}
	
	$row1 += 3;
	
	while (strtotime($date) <= strtotime($to)) {
		$strotimeDate =  strtotime($date);
		$weekday = date('w', $strotimeDate);
		if ((($mode == 'day') && ($weekday != 0 && $weekday != 6))){
			cellValue($objSheet, date('d', $strotimeDate)."(".date('D', $strotimeDate).")", getNameFromNumber(++$col1).($row1), "", false, true, false);
		}			
		if ($mode == 'night' || $mode == 'all'){
			cellValue($objSheet, date('d', $strotimeDate)."(".date('D', $strotimeDate).")", getNameFromNumber(++$col1).($row1), "", false, true, false);			
		}
		$date = date ("Y-m-d", strtotime("+1 day", $strotimeDate));
	}
	
	$tbl2StartRow = $row1 + 1;
	
	$row1++;
	$col1 = 1;
	$date = $from;
	
	foreach ($cunit as $key => $unit){
		cellValue($objSheet, $unit."(%)", getNameFromNumber($col1).($row1), "", false, false, false);
		$course_daliy_total_usage[$unit] = "<tr><td>$unit(%)</td>";
		$row1++;
	}
	
	$row1 = 3;
	$col1++;
	$fullTime_str = array('早', '早', '午', '午');
	$partTime_str = array('早', '早', '午', '午', '晚');
	if ($mode == 'night' || $mode == 'all'){$shift_str = $partTime_str; }
	if ($mode == 'day'){$shift_str = $fullTime_str; }
	//$shift_str = array('早', '早', '午', '午', '晚');
	$total_column = 1;
	$unit_col = 2;
	
	while (strtotime($date) <= strtotime($to)) {
		
		$strotimeDate =  strtotime($date);
		$weekday = date('w', $strotimeDate);
		
		if(($mode == 'day') && ($weekday == 0 || $weekday == 6)){
			$date = date ("Y-m-d", strtotime("+1 day", $strotimeDate));			
		}else{
			$colspan = 0;
			$course_booked = array();
			$course_used = array();
			
			$start_col = $col1;
			
			$DateIsMerge = false;
			
			if ((($weekday == 0 || $weekday == 6) && $mode == 'night') ||
					($mode == 'day' && ($weekday != 0 && $weekday != 6)) || ($mode == 'all')){
				
				$toDate = getNameFromNumber($start_col + count($shift_str)-1).($row1);
				$DateIsMerge = true;
				$colspan = count($shift_str);
				foreach ($shift_str as $key => $sStr){
					$tmp_shift .= "<td>$sStr</td>";
				}
				//$tmp_shift .= "<td>早</td><td>早</td><td>午</td><td>午</td><td>晚</td>";
				$total_column += count($shift_str);
				
			}else{
				if($mode == 'night'){
					$tmp_shift .= "<td>晚</td>";
					$total_column++;
				}
				
			}
			
			cellValue($objSheet, date('d', $strotimeDate)."(".date('D', $strotimeDate).")", getNameFromNumber($start_col).($row1), $toDate, $DateIsMerge, true, false);
			$row1++;
			
			if ($DateIsMerge==true){
				foreach ($shift_str as $key => $sStr){
					cellValue($objSheet, $sStr, getNameFromNumber($start_col).($row1), "", false, true, false);
					$start_col++;
				}
			}else{
				cellValue($objSheet, '晚', getNameFromNumber($start_col).($row1), "", false, true, false);
			}
			$start_col = $col1;
			
			foreach($data_ary as $key => $loc){
				$row1++;
				if ((($weekday == 0 || $weekday == 6) && $mode == 'night') ||
						(($mode == 'day') && ($weekday != 0 && $weekday != 6)) || ($mode == 'all')){
							
							$i = 0;
							$bgcolor = "";
							
							while ($i <= count($shift_str) - 1){
								
								$currDate = date('Y-m-d', $strotimeDate);
								$loc_data = get_loc_shift_val($loc['id'], $i, $currDate);
								$start_section = get_section_code(strtotime("$currDate ".$loc_data["starttime"]));
								$new_end_time = explode(":", $loc_data["endtime"]);
								$new_hour = $new_end_time[0];
								
								if ($new_end_time[1] == "00"){
									$new_min= 59;
									$new_hour -= 1;
								}else{
									$new_min = $new_end_time[1] - 1;
								}
								
								$new_end_time2 = "$new_hour:$new_min";

								$end_section = get_section_code(strtotime("$currDate $new_end_time2"));
								if ($mode == 'day'){
									if($end_section >= 3){
										$end_section = 3;
									}
								}

								$addColNum = $end_section - $start_section;
								$tblAddColNum = $addColNum + 1;
								
								$course_booked[$loc_data['type']] += 1;
								
								if ($loc_data['isUsedRoom'] == 1){
									$course_used[$loc_data['type']] += 1;
								}	
								if($tblAddColNum > 1){
									$i += $tblAddColNum;

									$check_data = get_loc_shift_val($loc['id'], ($i - 1), $currDate);
									
									if ($check_data["name"] != ''){
										$i--;
										$tblAddColNum--;
										$addColNum--;
									}
									
									$isMerge = true;
								}else{
									$i++;
									$isMerge = false;
								};
					
								$loc_val[$key] .= addTBLCell($loc_data, true ,($i <=  count($shift_str) - 1? false: true), $tblAddColNum);
								
								if ($addColNum <= 0 ){
									$addColNum = 0;
								}
								if ($isMerge){
									$toCol = getNameFromNumber($start_col + $addColNum).($row1);
								}else{
									$toCol = "";
								}
								if ($loc_data['isUsedRoom'] == 1){
									//echo date('Y-m-d', $strotimeDate) . " - " . $loc_data["name"] . "<BR>";
									//cellValueForRichText($objSheet, $loc_data['name']." - " . ($isMerge? "isMerge":"NotMerge"). " - $i - $aa", getNameFromNumber($start_col).($row1), $toCol,  $isMerge, true, false, "","", true);
									cellValueForRichText($objSheet, $loc_data['name'], getNameFromNumber($start_col).($row1), $toCol,  $isMerge, true, false, "","", true);
								}else{
									cellValue($objSheet, $loc_data['name'], getNameFromNumber($start_col).($row1), $toCol, $isMerge, true, false, "", "", true);
									//echo date('Y-m-d', $strotimeDate) . " - " . $loc_data["name"] . "<BR>";
									//cellValue($objSheet, $loc_data['name']." - " . ($isMerge? "isMerge":"NotMerge"). " - $i - $aa", getNameFromNumber($start_col).($row1), $toCol, $isMerge, true, false, "", "", true);
								}
								
								if ($isMerge){	
									$start_col += $tblAddColNum;
								}else{
									$start_col++;
								}
							}
							
							$start_col = $col1;
							
				}else{
					if ($mode=='night'){
						$loc_data = get_loc_shift_val($loc['id'], 4, date('Y-m-d', $strotimeDate));
						$loc_val[$key] .= addTBLCell($loc_data, false, true, 0);
						$course_booked[$loc_data['type']] += 1;
						if ($loc_data['isUsedRoom'] == 1){
							$course_used[$loc_data['type']] += 1;
						}
						
						if ($loc_data['isUsedRoom'] == 1){
							cellValueForRichText($objSheet, $loc_data['name'], getNameFromNumber($start_col).($row1), "",  false, true, false, "","", true);
						}else{
							cellValue($objSheet, $loc_data['name'], getNameFromNumber($start_col).($row1), "", false, true, false, "", "", true);
						}
					}
					
					
				}
			}
			
			
			$unit_row = $tbl2StartRow;
			
			foreach ($cunit as $key => $unit){
				
				$daily_persentage = calUsage($course_booked[$unit], $course_used[$unit]);
				cellValue($objSheet, $daily_persentage, getNameFromNumber($unit_col).($unit_row++), "", false, true, false, "", "", true);
				
				$course_daliy_total_usage[$unit] .= "<td>".calUsage($course_booked[$unit], $course_used[$unit]). "</td>";
				$course_total_booked[$unit] += $course_booked[$unit];
				$course_total_used[$unit] += $course_used[$unit];
			}
			
			$end_unit_row = $unit_row;
			$unit_col++;
			
			$total_row = $row1;
			$tmp_date .= "<td colspan=\"$colspan\">".date('d', $strotimeDate)."(".date('D', $strotimeDate).")</td>";
			$tmp_date2 .= "<td>".date('d', $strotimeDate)."(".date('D', $strotimeDate).")</td>";
			$total_num .= "<td></td>";
			$date = date ("Y-m-d", strtotime("+1 day", $strotimeDate));
			$row1 = 3;
			
			if ($colspan == count($shift_str)){
				$col1 += count($shift_str);
			}else {
				$col1++;
			}
		}		
	}
																													
	foreach($data_ary as $key => $loc){
		$loc_val[$key] .= "</tr>";
		$td_data .= $loc_val[$key];
	}
	
	$td_date = "<tr><td>日期</td>$tmp_date";
	$td_date2 = "<tr><td>日期</td>$tmp_date2";
	$td_shift = "<tr><td>課室</td>$tmp_shift</tr>";

	$table = "<table class=\"tbl_report7\">$td_date</tr>$td_shift.$td_data</table><BR><BR>";
	$ptable = "<table class=\"tbl_report7\">$td_date2<td></td><td></td><tr>";
	
	$unit_row = $tbl2StartRow;
	
	foreach ($cunit as $key => $unit){
		$total_booked += $course_total_booked[$unit];
		$total_used += $course_total_used[$unit];
		$course_total_persentage = calUsage($course_total_booked[$unit], $course_total_used[$unit]);
		
		cellValue($objSheet, $course_total_persentage, getNameFromNumber($unit_col).($unit_row), "", false, true, false, "", "", true);
		cellValue($objSheet, $unit."(%)", getNameFromNumber($unit_col+1).($unit_row++), "", false, true, false, "", "", true);
		$course_total_num = "<td style=\"background: #FFF200;\">$course_total_persentage</td><td style='min-width:125px;'>$unit(%)</td>";
		$ptable .= $course_daliy_total_usage[$unit]."$course_total_num<tr>";
	}	
	
	$total_Usage_persentage = calUsage($total_booked, $total_used);
	$total_num .= "<td></td><td style=\"background: #FFF200;\">$total_Usage_persentage</td><td>全校使用率(%)</td></tr>";
	
	cellValue($objSheet, "全校使用率(%)", getNameFromNumber($unit_col+1).($unit_row), "", false, true, false, "", "", true);
	cellValue($objSheet, $total_Usage_persentage, getNameFromNumber($unit_col).($unit_row), "", false, true, false, "", "", true);

	echo $table;
	echo "$ptable$total_num</table>";
	
	cellBorder("A3", getNameFromNumber($total_column).$total_row, $objSheet);
	cellBorder("A".($tbl2StartRow-1), getNameFromNumber($unit_col+1).$end_unit_row, $objSheet);
	
	$filename = "download/ClassRoomActualUsage".date('YmdHis').".xlsx";
	
	$objWriter->save($filename);

?>

<br>

<p>如要下載檔案，請右擊以下連結以另存目標。</p>
<A href="<?php echo $filename ?>">下載每月課室使用狀況</A><br>
<br>
<A href="javascript:history.go(-1)">&nbsp;返回上一頁</A>
