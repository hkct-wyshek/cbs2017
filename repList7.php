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
$objSheet->getDefaultColumnDimension()->setWidth(25);

$objSheet2 = $objPHPExcel2->getActiveSheet();
$objSheet2->getDefaultColumnDimension()->setWidth(15);


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

$objSheet->setTitle($xls_campus);
$objSheet2->setTitle($xls_campus);

cellValue($objSheet, $xls_campus.$modeTitle."課室使用資料", "A1", "E1", true, false, true, "", "", false, false, 16);
cellValue($objSheet2, $xls_campus.$modeTitle."課室使用資料", "A1", "E1", true, false, true, "", "", false, false, 16);

cellValue($objSheet, $from."至".$to, "A2", "E2", true, false, true, "", "", false, false, 14);
cellValue($objSheet2, $from."至".$to, "A2", "E2", true, false, true, "", "", false, false, 14);

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

$row1 = 6;
$col1 = 2;
$second_table = array();
$total_row = 1;


foreach($aid as $key =>$val){
	
	$time_slot1 = "<td class='title' style=' font-weight:bold;' ><p>".$data_ary[$key]['room_name']."</p></td>";
	$time_slot2 = "<td class='title'><p></p></td>";
	$time_slot3 = "<td class='title'><p></p></td>";

	cellValue($objSheet, $data_ary[$key]['room_name'], "A".$row1, "", false, false, true);
	
	foreach ($cunit as $num => $locName){
		if (!empty($locName)){
			$loc_num[strtoupper($locName) . '_' . $data_ary[$key]['room_name']] = 0;
			$loc_num[strtoupper($locName).'_'. $data_ary[$key]['room_name'] .'_ACTUAL'] = 0;
			$loc_num[strtoupper($locName).'_'. $data_ary[$key]['room_name'] .'_ARRANGED'] = 0;
		}	
	}
		
	$loc_num[strtoupper($row[0]) . '_' . $data_ary[$key]['room_name']] += 1;
	$loc_num[strtoupper($row[0]).'_'. $data_ary[$key]['room_name'] .'_ACTUAL'] += 1;
	$loc_num[strtoupper($row[0]).'_'. $data_ary[$key]['room_name'] .'_ARRANGED'] += 1;

	$date = $from;
	while (strtotime($date) <= strtotime($end_date)) {
		
		if ($mode=="day") { $start_time = "09:00"; $end_time = "18:59"; } 
		if ($mode=="night") { $start_time = "19:00"; $end_time = "22:00"; } 
		if ($mode=="all") { $start_time = "09:00"; $end_time = "22:00"; }
		
		if ($isfirst){
			$count += 1;
			if (!empty($date)){

				if (date('m', strtotime($date))!= $curr_month ){
					$curr_year = date('Y', strtotime($date));
					$curr_month = date('m', strtotime($date));
					if ($count == 1){
						$month_str.="<td></td>";
					}
					
					$m_str = $curr_year. "年" . $curr_month . "月";
					$month_str .= "<td style=' font-weight:bold;'>$m_str</td>";
					
					cellValue($objSheet, $m_str, getNameFromNumber(($col1 == 2 ? $col1-1 : $col1))."3", "", false, true, true);

				}else{
					$month_str .= "<td></td>";
				}
			}
			
			$w_str = date('D', strtotime($date));
			$d_str = date('d', strtotime($date));
			
			$week_str .= "<td style=' font-weight:bold;'>$w_str</td>";
			$day_str  .= "<td style=' font-weight:bold;'>$d_str</td>";
			
			cellValue($objSheet, $w_str, getNameFromNumber($col1)."4", "", false, true, true);
			cellValue($objSheet, $d_str, getNameFromNumber($col1)."5", "", false, true, true);
			
		}
		
		
		if ($mode=="night" && (date('w', strtotime($date)) == 0 || date('w', strtotime($date)) == 6)){
			$start_time = "09:00";
			$end_time   = "22:00";
		}
	
		$tmp_start = strtotime("$date $start_time");
		$tmp_end   = strtotime("$date $end_time");
		
		$morning_end = strtotime("$date 09:00");
		$afternoon_end = strtotime("$date 12:00");
		$evening_end = strtotime("$date 19:00");
		
		if (($mode == "day" && date('w', strtotime($date)) == 0) || ($mode == "day" && date('w', strtotime($date)) == 6)){
			$time_slot1 .= "<td><p></p></td>";
			$time_slot2 .= "<td><p></p></td>";
			$time_slot3 .= "<td><p></p></td>";
			
		}else {
			$sql = "SELECT type, isUsedRoom, start_time, end_time, name FROM $tbl_entry WHERE room_id=$val AND start_time >= $tmp_start AND start_time <= $tmp_end ORDER BY start_time";
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
					$tmp['name']       = $row[4];
					$tmp['type']       = $row[0];
					
					if ($row[0] != ''){
						$loc_num[strtoupper($row[0]) . '_' . $data_ary[$key]['room_name']] += 1;
						if ($tmp['isUsedRoom'] == 1){
							$loc_num[strtoupper($row[0]).'_'. $data_ary[$key]['room_name'] .'_ACTUAL'] += 1;
							$loc_num[strtoupper($row[0]).'_'. $data_ary[$key]['room_name'] .'_ARRANGED'] += 1;
						}
						elseif(strtoupper($tmp['type']) != $tmp['name'] && $tmp['name'] != '' && $tmp['isUsedRoom'] == 0){
							$loc_num[strtoupper($row[0]).'_'. $data_ary[$key]['room_name'] .'_ARRANGED'] += 1;
						}
					}
					
						
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
				
				
				cellValueForRichText($objSheet, $tmp_xls_val1, getNameFromNumber($col1).$row1, "",  false, true, false, getNameFromNumber($col1), 25, true);
				cellValueForRichText($objSheet, $tmp_xls_val2, getNameFromNumber($col1).($row1+1), "",  false, true, false, "", "", true);
				cellValueForRichText($objSheet, $tmp_xls_val3, getNameFromNumber($col1).($row1+2), "",  false, true, false, "", "", true);
			}
			
			
		}
		$date = date ("Y-m-d", strtotime("+1 day", strtotime($date)));
		$col1++;
	}
	
	/*Total*/
	$tmp_table .= "<tr>";
	$tmp_table .= "<td>" . $data_ary[$key]['room_name'] . "</td>";
	$tmp_array = array();
	array_push($tmp_array, $data_ary[$key]['room_name']);
	
	foreach ($cunit as $key2 => $unitName){
		
		$b = $loc_num[strtoupper($unitName) . '_' . $data_ary[$key]['room_name']];
		$a = $loc_num[strtoupper($unitName) . '_' . $data_ary[$key]['room_name'] . '_ACTUAL'];
		$c = $loc_num[strtoupper($unitName) . '_' . $data_ary[$key]['room_name'] . '_ARRANGED'];
		
		$loc_num[strtoupper($unitName)] += $b; 
		$loc_num[strtoupper($unitName) . '_ACTUAL'] += $a;
		$loc_num[strtoupper($unitName) . '_ARRANGED'] += $c;
		
		$total_b += $b;
		$total_a += $a;
		$total_c += $c;
		
		$ab = round($a / $b * 100);
		$cb = round($c / $b * 100);
		
		$total_ab = round($total_a / $total_b * 100);
		$total_cb = round($total_c / $total_b * 100);
		
		$tmp_table .= "<td>" . $b . "</td>";
		$tmp_table .= "<td>" . $c . "</td>";
		$tmp_table .= "<td>" . $cb . "%</td>";
		$tmp_table .= "<td>" . $a . "</td>";
		$tmp_table .= "<td>" . $ab . "%</td>";
		
		array_push($tmp_array, $b, $c, $cb."%", $a, $ab."%");
	}
	
	$tmp_table .= "<td>" . $total_b . "</td>";
	$tmp_table .= "<td>" . $total_c . "</td>";
	$tmp_table .= "<td>" . $total_cb . "%</td>";
	$tmp_table .= "<td>" . $total_a . "</td>";
	$tmp_table .= "<td>" . $total_ab . "%</td>";
	
	array_push($tmp_array, $total_b, $total_c, $total_cb."%", $total_a, $total_ab."%");	
	array_push($second_table, $tmp_array);
	
	$total_row ++;
	
	$total_b = 0;
	$total_a = 0;
	$total_c = 0;
	
	$tmp_table ."</tr>";	
	
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

cellBorder("A3", getNameFromNumber($count).($row1-1), $objSheet);

$row1 += 3;
$row2 =  3;

$second_tbl_first_row = $row1;
$thead .= "<th></th>";
$thead2 .= "<th></th>";

$sub_thead   = array("預約總數", "預計使用", "預計使用率(%)", "實際使用", "實際使用率(%)", "總數");

$thead2_html = "<th>".$sub_thead[0]."</th><th>".$sub_thead[1]."</th><th>".$sub_thead[2]."</th><th>".$sub_thead[3]."</th><th>".$sub_thead[4]."</th>";

foreach ($cunit as $key2 => $unit_name){
	
	$thead .= "<th colspan='5' width='100'>$unit_name</th>";
	$thead2 .= $thead2_html;
	if ($key2 == 0){
		$colNum = 2;
		$toColNum = 6;
	}else{
		$colNum = $toColNum + 1;
		$toColNum = $colNum + 4;			
	}
	
	$fromCol = getNameFromNumber($colNum).$row1;
	$toCol = getNameFromNumber($toColNum).$row1;
	
	
	
	cellValue($objSheet, $unit_name, $fromCol, $toCol, true, true, true);
	cellValue($objSheet, $sub_thead[0], getNameFromNumber($colNum).($row1+1), "", false, true, true);
	cellValue($objSheet, $sub_thead[1], getNameFromNumber($colNum+1).($row1+1), "", false, true, true);
	cellValue($objSheet, $sub_thead[2], getNameFromNumber($colNum+2).($row1+1), "", false, true, true);
	cellValue($objSheet, $sub_thead[3], getNameFromNumber($colNum+3).($row1+1), "", false, true, true);
	cellValue($objSheet, $sub_thead[4], getNameFromNumber($colNum+4).($row1+1), "", false, true, true);
	
	cellValue($objSheet2, $unit_name, getNameFromNumber($colNum).$row2, getNameFromNumber($toColNum).$row2, true, true, true);
	cellValue($objSheet2, $sub_thead[0], getNameFromNumber($colNum).($row2+1), "", false, true, true);
	cellValue($objSheet2, $sub_thead[1], getNameFromNumber($colNum+1).($row2+1), "", false, true, true);
	cellValue($objSheet2, $sub_thead[2], getNameFromNumber($colNum+2).($row2+1), "", false, true, true);
	cellValue($objSheet2, $sub_thead[3], getNameFromNumber($colNum+3).($row2+1), "", false, true, true);
	cellValue($objSheet2, $sub_thead[4], getNameFromNumber($colNum+4).($row2+1), "", false, true, true);
		
		
}

$thead  .= "<th colspan='5' width='100'>總數</th>";
$thead2 .= $thead2_html;

$colNum = $toColNum + 1;
$toColNum = $colNum + 4;

$fromCol = getNameFromNumber($colNum).$row1;
$toCol = getNameFromNumber($toColNum).$row1;

$objSheet->mergeCells("$fromCol:$toCol");
cellValue($objSheet, $sub_thead[5], $fromCol, $toCol, true, true, true);
cellValue($objSheet, $sub_thead[0], getNameFromNumber($colNum).($row1+1), "", false, true, true);
cellValue($objSheet, $sub_thead[1], getNameFromNumber($colNum+1).($row1+1), "", false, true, true);
cellValue($objSheet, $sub_thead[2], getNameFromNumber($colNum+2).($row1+1), "", false, true, true);
cellValue($objSheet, $sub_thead[3], getNameFromNumber($colNum+3).($row1+1), "", false, true, true);
cellValue($objSheet, $sub_thead[4], getNameFromNumber($colNum+4).($row1+1), "", false, true, true);

cellValue($objSheet2, $sub_thead[5], getNameFromNumber($colNum).$row2, getNameFromNumber($toColNum).$row2, true, true, true);
cellValue($objSheet2, $sub_thead[0], getNameFromNumber($colNum).($row2+1), "", false, true, true);
cellValue($objSheet2, $sub_thead[1], getNameFromNumber($colNum+1).($row2+1), "", false, true, true);
cellValue($objSheet2, $sub_thead[2], getNameFromNumber($colNum+2).($row2+1), "", false, true, true);
cellValue($objSheet2, $sub_thead[3], getNameFromNumber($colNum+3).($row2+1), "", false, true, true);
cellValue($objSheet2, $sub_thead[4], getNameFromNumber($colNum+4).($row2+1), "", false, true, true);

$row1+=2;
$row2+=2;

$isfirst = true;
echo "<table class=\"tbl_report7\" style='margin-top:20px;'>";


foreach($aid as $key =>$val){

	$time_slot1 = "<td class='title' style=' font-weight:bold;' ><p>".$data_ary[$key]['room_name']."</p></td>";
	$time_slot2 = "<td class='title'><p></p></td>";
	$time_slot3 = "<td class='title'><p></p></td>";

	if ($isfirst){
		echo "<tr>$thead</tr>";
		echo "<tr>$thead2</tr>";
		
		$isfirst = false;
	}	
}

foreach ($second_table as $key => $val){
	foreach ($val as $i => $str){
		cellValue($objSheet, $str, (getNameFromNumber($i+1).$row1), "", false, ($i==0? false: true), ($i==0? true : false));
		cellValue($objSheet2, $str, (getNameFromNumber($i+1).$row2), "", false, ($i==0? false: true), ($i==0? true : false));
	}
	$row1++;
	$row2++;
}



echo $tmp_table;
echo "<tr><td>".$sub_thead[5]."</td>";

cellValue($objSheet, $sub_thead[5], getNameFromNumber(1).$row1, "", false, false, true);
cellValue($objSheet2, $sub_thead[5], getNameFromNumber(1).$row2, "", false, false, true);
$col1 = 2;

foreach ($cunit as $key2 => $unitName){
	
	$b = $loc_num[strtoupper($unitName)];
	$a = $loc_num[strtoupper($unitName) . '_ACTUAL'];
	$c = $loc_num[strtoupper($unitName) . '_ARRANGED'];
		
	$total_b += $b;
	$total_a += $a;
	$total_c += $c;
	
	$ab = round($a / $b * 100);
	$cb = round($c / $b * 100);
	
	$total_ab = round($total_a / $total_b * 100);
	$total_cb = round($total_c / $total_b * 100);
	
	echo "<td>" . $b. "</td>";
	echo "<td>" . $c. "</td>";
	echo "<td>" . $cb. "%</td>";
	echo "<td>" . $a. "</td>";
	echo "<td>" . $ab . "%</td>";	
	
	cellValue($objSheet, $b, getNameFromNumber(($col1)).$row1, "", false, true, false);
	cellValue($objSheet, $c, getNameFromNumber(($col1+1)).$row1, "", false, true, false);
    cellValue($objSheet, "$cb%", getNameFromNumber(($col1+2)).$row1, "", false, true, false);
    cellValue($objSheet, $a, getNameFromNumber(($col1+3)).$row1, "", false, true, false);
    cellValue($objSheet, "$ab%", getNameFromNumber(($col1+4)).$row1, "", false, true, false);
    
    cellValue($objSheet2, $b, getNameFromNumber(($col1++)).$row2, "", false, true, false);
    cellValue($objSheet2, $c, getNameFromNumber(($col1++)).$row2, "", false, true, false);
    cellValue($objSheet2, "$cb%", getNameFromNumber(($col1++)).$row2, "", false, true, false);
    cellValue($objSheet2, $a, getNameFromNumber(($col1++)).$row2, "", false, true, false);
    cellValue($objSheet2, "$ab%", getNameFromNumber(($col1++)).$row2, "", false, true, false);
}


echo  "<td>" . $total_b . "</td>";
echo  "<td>" . $total_c . "</td>";
echo  "<td>" . $total_cb . "%</td>";
echo  "<td>" . $total_a . "</td>";
echo  "<td>" . $total_ab . "%</td>";

cellValue($objSheet, $total_b, getNameFromNumber($col1).$row1, "", false, true, false);
cellValue($objSheet, $total_c, getNameFromNumber($col1+1).$row1, "", false, true, false);
cellValue($objSheet, "$total_cb%", getNameFromNumber($col1+2).$row1, "", false, true, false);
cellValue($objSheet, $total_a, getNameFromNumber($col1+3).$row1, "", false, true, false);
cellValue($objSheet, "$total_ab%", getNameFromNumber($col1+4).$row1, "", false, true, false);

cellValue($objSheet2, $total_b, getNameFromNumber($col1++).$row2, "", false, true, false);
cellValue($objSheet2, $total_c, getNameFromNumber($col1++).$row2, "", false, true, false);
cellValue($objSheet2, "$total_cb%", getNameFromNumber($col1++).$row2, "", false, true, false);
cellValue($objSheet2, $total_a, getNameFromNumber($col1++).$row2, "", false, true, false);
cellValue($objSheet2, "$total_ab%", getNameFromNumber($col1++).$row2, "", false, true, false);


echo "</tr>";
echo "</table>";


cellBorder("A".$second_tbl_first_row, getNameFromNumber(--$col1) . $row1, $objSheet);
cellBorder("A3", getNameFromNumber($col1) . $row2, $objSheet2);

$objSheet->getColumnDimension("A")->setAutoSize(true);
$objSheet2->getColumnDimension("A")->setAutoSize(true);

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