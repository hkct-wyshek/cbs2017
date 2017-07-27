<?php 

require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";

require 'PHPExcel.php';
require_once 'PHPExcel/Writer/Excel5.php';
require_once 'PHPExcel/Writer/Excel2007.php';
require_once 'PHPExcel/IOFactory.php';

ini_set('max_execution_time', 1000);

if(!getAuthorised(1)){
	showAccessDenied($day, $month, $year, $area);
	exit;
}


if (!isCourseAdmin()){
	showAccessDenied($day, $month, $year, $area);
	exit;
}
print_header($day, $month, $year, $area);

$start_day   = date("d");
$start_month = date("m");
$start_year  = date("Y");

function check_isValid_data($start, $end, $room){
	global $tbl_entry;
	$query = "SELECT id FROM $tbl_entry e ". 
			 "WHERE start_time = '$start' AND end_time = '$end' ".
			 "AND create_by = '". getUserName()."' AND room_id = $room";
	return sql_query1($query);
}

function check_repeat_data($tmpData, $dataArr){
	foreach($dataArr as $key => $val){
		if ($val['starttime'] == $tmpData['starttime'] && $val['endtime'] == $tmpData['endtime'] && $val['course'] == $tmpData['course']){
			return $val['xls_row'];
		}
	}
	return -1;
}


if (isset($submit)){
	if ($submit == 1){
	
		$msg = "";
	
	$file_type = array('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/vnd.ms-excel' ) ;

	if (!in_array($_FILES['booking_detail']['type'], $file_type)){
		$msg =  get_vocab('upload_err2');
			
	}else if ($_FILES['booking_detail']['size'] == 0){
		$msg = get_vocab('upload_err1');
		
	}else{
		
		$temp = explode(".", $_FILES["booking_detail"]["name"]);	
		$path = "upload/" .date('YmdHi').".".end($temp);
		
		if (!move_uploaded_file($_FILES['booking_detail']['tmp_name'], $path)) {
	    	$msg =  get_vocab('upload_err3');
		    break;  
		}
					
	    if(!file_exists($path)) return false;
	   
	    $pathinfo = pathinfo($path);
	    $file_extension= isset($pathinfo['extension'])?$pathinfo['extension']:'';
	
	    $file_extension = strtolower($file_extension);
	
	    $excel_reader_type = '';
	
	    if($file_extension == 'xls')
	      $excel_reader_type = 'Excel5';
	    else if($file_extension == 'xlsx')
	      $excel_reader_type = 'Excel2007';
	   

	    if(!$excel_reader_type) return false;
	  
		try {
			
		    $objReader = PHPExcel_IOFactory::createReader($excel_reader_type);
		   
		    $objPHPExcel = $objReader->load($path);
		    
		     
		    
			if ($objReader->canRead($path)) {
	
				$dataArr = array();
				$tmpArr = array();	 
				
				$export_title = array(get_vocab("xls_row"),
									  get_vocab("xls_starttime"),
									  get_vocab("xls_endtime"),
									  get_vocab("xls_class_room"),
									  get_vocab("xls_class_no"),
									  get_vocab("xls_remarks"),
									  get_vocab("xls_system_msg"));
				
				$export_name = array(
									"starttime",
									"endtime",
									"course",
									"class_no",
									"remarks",
									"notValid");
				
				$CurrentWorkSheetIndex = 0; 
				$total_wrong = 0;	
				
				foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
					
					if ($CurrentWorkSheetIndex++ == 0){
						 
						$worksheetTitle     = $worksheet->getTitle();
					    $highestRow         = $worksheet->getHighestRow(); // e.g. 10
					    $highestColumn      = $worksheet->getHighestColumn(); // e.g 'F'
					    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);
					    $row_data = $worksheet->rangeToArray('A1:' . $highestColumn . $highestRow);
					    $row_data = array_map('array_filter', $row_data);
		       			$row_data = array_filter($row_data);
		       			
		       			if (count($row_data)<=1){
		       				$msg = get_vocab('upload_err4');
		       				break;
		       			}

		       			$str.= "<form method=\"get\" action=\"update_course_detail.php\" >";
					    $str.= "<table width=100% border=1 style='margin:10px 0;'>";
					    $str.= "<tr>";
					    
					    foreach ($export_title as $title){
					    	$str.= "<th>$title</th>";  	  	
					    }
					    
						$str.= "</tr>";
						$num = 0;
						$msg = "";
					    for ($row = 2; $row <= $highestRow; ++ $row) {
					    	
					    	$row_wrong = 0;
					    	if (! empty($row_data[$row-1])){
					    		
					    		$str.= "<tr><td>$row</td>";	
						    	
						    	for ($col = 0; $col < 5; ++ $col) {
						    		
						        	$isWrong = false;
						        	$isSingle = false;
						            $cell = $worksheet->getCellByColumnAndRow($col, $row);
						            $val = $cell->getValue();
						          
						        	switch ($export_name[$col]) {

						            	case "starttime":
						            	case "endtime":
						            		
						            		$cell_value = PHPExcel_Style_NumberFormat::toFormattedString($cell->getCalculatedValue(), 'yyyy-mm-dd hh:mm');
						            		if (PHPExcel_Shared_Date::isDateTime($worksheet->getCellByColumnAndRow($col, $row))){
						            			$cell_value = $cell->getFormattedValue();
						            			$val = $cell_value;
						            		}

						        			if ( preg_match("/^[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1]) (2[0-4]|[01][1-9]|10):([0-5][0-9])$/", (string)$cell_value)){
						        				$tmp_date = explode(" ", $cell_value);
						        				
						        				$tmp_dd = explode("-", $tmp_date[0]);
						        				$tmp_tt = explode(":", $tmp_date[1]);
						        				
						        				$data = mktime($tmp_tt[0], $tmp_tt[1], 0, $tmp_dd[1], $tmp_dd[2], $tmp_dd[0]);
						        				$today = mktime(0, 0, 0, date('m'), date('d'), date('Y'));
						        				
												if ($data > $today){
													if ($col == 0){
							        					$tmpstarttime = $data;
							        				}else{
							        					$tmpendtime = $data;
							        				}
												}else{
													$isWrong = TRUE;
												}
						        				
						        			}else{
						        				$isWrong = TRUE;
						        			}
											break;
						            		
						            	case "course":
				
						            		if (!empty($val)){
						            			$arr = explode("-", $val);
						            			if (count($arr) == 2){
						            				$id = sql_query1("SELECT id FROM $tbl_area Where area_name ='" . mysql_real_escape_string(trim((string)$arr[0]," ")) . "' LIMIT 1");			  
							            			if ($id != -1){
							            				$id = sql_query1("SELECT id FROM $tbl_room Where room_name ='" . mysql_real_escape_string(trim((string)$arr[1]," ")) . "' AND area_id = $id LIMIT 1");
							            				if ($id != -1){
								            				$tmproom = $id;
								            				$data = $id;						            				
								            				break;
								            			}
							            			}
						            			}	
						            		}
						            		$data = "";
						            		$isWrong = TRUE;
						            		break;
						            		
						            	case "class_no":
						            		if (strlen($val) > 80){
						            			$isWrong=true;
						            		}	
						            		$data = $val;			  
						            		break;   
						            		       		
						            	default:
						            		$data = $val;
						            		break;	
						            } //End Switch
						            
						             if ($isWrong == false){  
						             	
						             	 $tmp_data[$export_name[$col]] = (string)$data;	
						             	 $tmp[$export_name[$col]] = (string)$data;
										 $str.= "<td>".htmlspecialchars($val)."</td>";
										 
						             	 if ($export_name[$col] == "remarks"){
						             	 	
						             	 	$tmp["xls_row"] = $row;
						             	 	$tmp_data["xls_row"] = $row;
						             	 		
						             	 	if (isset($tmpstarttime) && isset($tmpendtime) && isset($tmproom)){
						             	 		
						             	 		$id = check_isValid_data($tmpstarttime, $tmpendtime, $tmproom);
						             	 		
							             	 	if ($id == -1){
							             	 		$row_wrong ++;
									             	$total_wrong++;
									             	$str.= "<td style=\"background-color:#ffe6e6\">".get_vocab("xls_err1")."</td>";
							             	 	}else{
							             	 		$id = check_repeat_data($tmp_data, $dataArr);
							             	 		if ($id > 0){
							             	 			$str.= "<td>該記錄將會覆蓋行數:".$id."的記錄</td>";
							             	 		}else{
							             	 			$str.= "<td></td>";
							             	 		}
							             	 	}
							             	 	
						             	 	}else{
						             	 		$row_wrong ++;
									            $total_wrong++;
						             	 		$str.= "<td style=\"background-color:#ffe6e6\">".get_vocab("xls_err1")."</td>";
						             	 	}
						             	 }
						             	
						             }else {
						             	$row_wrong ++;
						             	$total_wrong++;
						             	$str.= "<td style=\"background-color:#ffe6e6\">".htmlspecialchars($val)."</td>";
						             }      
						              $input_str.= "<input type=\"hidden\" value=\"$data\" name=\"data[".$num."][".$export_name[$col]."]\"  />";
						        }//End for loop $col
						        
						        if ($row_wrong == 0){
						        	array_push($dataArr, $tmp);
						       		array_push($tmpArr, $tmp_data);
						        }
						        
						        $str.= "</tr>";
					    	}
					    	
						    $tmpstarttime = '';
						    $tmpendtime = '';
						    $tmproom = '';
					    }//End for loop $row
					    
					    $str.= "</table>";

					    if ($total_wrong == 0){
					    	$str.= "<div style='text-align:center'>";
					    	$str.= "<button style=\"margin-right:10px;\" value=\"0\" name=\"confirm\" type=\"submit\" onclick=\"validation()\">".get_vocab("confirm")."</button>" . 
					    		   "<button value=\"1\" name=\"confirm\">".get_vocab("cancel")."</button>";
					    	$str.= "</div>";
							$_SESSION['C_List'] = $dataArr;
					    	$msg = "";
					    }else{
					    	$msg= get_vocab("excel_data_err2");
					    	unset($_SESSION['input_data']);
					    }
					    $str .= "</form>";
					}  
				};//End Loop excel worksheet
				
			}else{
				$msg = get_vocab("cannot_read_file");
			}
		} catch(PHPExcel_Exception $e) {
		    $msg = get_vocab("cannot_read_file");
		}// End try
	}
	unlink($path); // delete file
	}
}

if (isset($_GET['msg']) && $_GET['msg'] == 0){
	$msg = get_vocab("success_upload");
}

?>

<h2><?php echo get_vocab("course_page_heading"); ?> <span style="font-size:13px;">(<?php echo get_vocab("wrong_date2"); ?>)</span></h2>

<h3><?php echo get_vocab("course_page_sub1"); ?></h3>

<form method="get" action="dlCourseInfo.php">
	<TABLE BORDER=0>
	<TR>
		<TD CLASS=CR><B><?php echo get_vocab("start_date_from");?></B></TD>
		<TD CLASS=CL><?php genDateSelector("", $start_day, $start_month, $start_year); ?></TD>
	</TR>
	
	<TR>
		<TD CLASS=CR><B><?php echo get_vocab("start_date_to");?></B></TD>
		<TD CLASS=CL><?php genDateSelector("end_", $start_day, $start_month, $start_year+1); ?></TD>
	</TR>
	
	<tr>
		<td colspan="2" ><button name="submit" value="0"><?php echo get_vocab("download_xls_file");?></button></td>
	</tr>
	
	</TABLE>
</form>


<h3><?php echo get_vocab("course_page_sub2"); ?></h3>

<form action="" enctype="multipart/form-data" method="post" name="form1" id="form1">
	<input type="file" name="booking_detail" accept=".xlsx, .xls">
	<button name="submit" type="submit" value="1"><?php echo get_vocab("upload"); ?></button>
</form>

<?php 
	if (!empty($str)){
		echo $str;
	}
?>
<p id="message" style="color:#ff0000;"><?php echo $msg; ?></p>

<?php

if(isLoggedIn())
{
	include "trailer.inc";
} 
else
{
	include "trailerNotYetLogin.inc";
}

?>