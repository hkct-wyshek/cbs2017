<?php
session_start();
# $Id: admin.php,v 1.16.2.1 2005/03/29 13:26:15 jberanek Exp $
require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
include "mrbs_sql.inc";



require 'PHPExcel.php';
require_once 'PHPExcel/Writer/Excel5.php';
require_once 'PHPExcel/Writer/Excel2007.php';
require_once 'PHPExcel/IOFactory.php';

ini_set('max_execution_time', 1000);

//ini_set('display_errors', 'On');
//error_reporting (E_ALL ^ E_NOTICE);
#If we dont know the right date then make it up 
if(!isset($day) or !isset($month) or !isset($year))
{
	$day   = date("d");
	$month = date("m");
	$year  = date("Y");
}

if (empty($area))
{
    $area = get_default_area();
}

if(!getAuthorised(2))
{
	showAccessDenied($day, $month, $year, $area);
	exit();
}


print_header($day, $month, $year, isset($area) ? $area : "");

// If area is set but area name is not known, get the name.
if (isset($area))
{
	if (empty($area_name))
	{
		$res = sql_query("select area_name from $tbl_area where id=$area");
    	if (! $res) fatal_error(0, sql_error());
		if (sql_count($res) == 1)
		{
			$row = sql_row($res, 0);
			$area_name = $row[0];
		}
		sql_free($res);
	} else {
		$area_name = unslashes($area_name);
	}
}

?>

<?php 

function get_data_start_end_date($start, $end, $weekly, $area){
	if ($weekly != -1 && !empty($weekly)){
		if (date('Y-m-d', $start) != date('Y-m-d', $end)){
			$new_date = Get_Divide_Date($start, $end, $weekly);
			return $new_date;
		}
	}
		$new_date = array();
		$new_tt = date('H:i', $end);
		$new_dd = date('Y-m-d', $start);
		$tmp['start'] = $start;
		$tmp['end'] = strtotime("$new_dd $new_tt");
		array_push($new_date, $tmp);
	
	return $new_date;
}

function check_repeat_excel_data($data, $tmp){
	
	$newStarttime = $tmp['starttime'];
	$newEndtime = $tmp['endtime'];
	$newweekly = $tmp['weekly'];
	$newArea = $tmp['area'];
	
	foreach ($data as $key => $d){
				
		$o_starttime = $d['starttime'];
		$o_endtime = $d['endtime'];
		$o_weekly = $d['weekly'];
		$o_area = $d['area'];
		
		$n_data = get_data_start_end_date($newStarttime, $newEndtime, $newweekly, $newArea);
		$o_data = get_data_start_end_date($o_starttime, $o_endtime, $o_weekly, $o_area);
		
		foreach ($n_data as $key => $n_val){
			foreach ($o_data as $num => $o_val){
				
				if (date('Y-m-d', $n_val['start']) == date('Y-m-d', $o_val['start'])){
					
					if (($n_val['start'] >= $o_val['start'] && $n_val['start'] < $o_val['end']) ||
					    ($n_val['end'] > $o_val['start'] && $n_val['end'] <= $o_val['end'])){
						if ($newArea == $o_area){
							return $d['xls_row'];
						}
					}
				}
			}
		}
	}// End $d foreach
	return "";
}//End function

function check_valid_date($dd){
	
	if (!preg_match("/^[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])$/",$dd)){
		return -1;
		//$tmpstartdate = "";
		//$isWrong = true;
		//break;
	}
	$dd = strtotime($dd);
	if((date('Y',$dd) <= date('Y')+5 && date('Y',$dd) >= date('Y')-5)){
		return $dd;
	}
	return -1;
		/*
		$val = date('Y-m-d',$dd);
		$data = $val;
								            				
		if ($export_name[$col] == "startdate"){							            					
			$tmpstartdate = $val;
		}else{
			if (!empty($tmpstartdate)){
				if(strtotime($tmpstartdate)<= strtotime($val)) 
					$tmpenddate = $val;
			}
		}
	}	*/
}

if (isset($submit) == "1"){
	
	$msg = "";
	
	$file_type = array('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/vnd.ms-excel' ) ;
	
	echo $_FILES["booking_detail"]["name"];
	
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
									  get_vocab("xls_startdate"),
									  get_vocab("xls_enddate"),
									  get_vocab("xls_starttime"),
									  get_vocab("xls_endtime"),
									  get_vocab("xls_repeat_week"),
									  get_vocab("areas"),
									  get_vocab("xls_type"),
									  get_vocab("xls_class_no"),
									  get_vocab("xls_remarks"),
									  get_vocab("xls_system_msg"));
				
				$export_name = array(
									"startdate",
									"enddate",
									"starttime",
									"endtime",
									"weekly",
									"area",
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

		       			$str.= "<form method=\"get\" action=\"import_booking_detail.php\" >";
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
						    	
						    	for ($col = 0; $col < 9; ++ $col) {
						    		
						        	$isWrong = false;
						        	$isSingle = false;
						            $cell = $worksheet->getCellByColumnAndRow($col, $row);
						            $val = $cell->getValue();
						            
						        	switch ($export_name[$col]) {
						        		
						            	case "startdate":
						            	case "enddate":
						            		
						        			if($export_name[$col]=="enddate" && empty($val) && !empty($tmpstartdate))	{
						            			$tmpenddate = $tmpstartdate;	
						            			break;
						            		}elseif($export_name[$col]=="enddate" && empty($val) && empty($tmpstartdate)){
						            			$isWrong = FALSE;
						            			break;
						            		}
						            		
						            		if (PHPExcel_Shared_Date::isDateTime($worksheet->getCellByColumnAndRow($col, $row))){
						            			$cell_value = $cell->getFormattedValue();
						            			$val = $cell_value;
						            		}
						            			
						            		if (preg_match("/^[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])$/", (string)$val)){
						            			$tmp_date = check_valid_date($cell->getFormattedValue());
						            			if ($tmp_date != -1){
						            				$val = $cell->getFormattedValue();
						            				
						            				if ($export_name[$col] == "startdate"){							            					
							            				$tmpstartdate = $val;
							            			}else{
							            				if (!empty($tmpstartdate)){
							            					if(strtotime($tmpstartdate)< strtotime($val)) {
							            						$tmpenddate = $val;
							            					}else{
							            						$isWrong = true;
							            					}
							            				}
							            			}
							            			break;
						            			}
						            		}
	
						            		$tmpstartdate = "";
						            		$isWrong = TRUE;
						            		break;
						            		
						            	case "starttime":
						            	case "endtime":
						            		
						            		$cell_value = PHPExcel_Style_NumberFormat::toFormattedString($cell->getCalculatedValue(), 'hh:mm');
						            		
						        			if (PHPExcel_Shared_Date::isDateTime($worksheet->getCellByColumnAndRow($col, $row)) && preg_match("/(2[0-4]|[01][1-9]|10):([0-5][0-9])/", (string)$cell_value)){
						        				
						        				$val = $cell_value;
						        				
						            			if (!empty($tmpstartdate)){
						            				$tmpdate = $export_name[$col] == "starttime"?$tmpstartdate:$tmpenddate;	
							            			$utime = strtotime($tmpdate . " " . date('H:i',strtotime($val)));
								            		$hh = date("H", $utime);
								            		$ii = date("i", $utime);
										            if ($export_name[$col] == "starttime"){						            				
										            	$tmptime = "$hh:00";
										            }else{
										            	$ii = date("i", $utime);
										            	if ($ii>0){
										            		$tmptime = ($hh+1).":00";
										            	}else 
										            		$tmptime = date('H:i',$utime);
										            }
										            
										            $data = strtotime($tmpdate . " " . $tmptime);		
										            							
													if (($data >= strtotime($tmpdate . " 09:00")) && ($data <= strtotime($tmpdate . " 22:00"))){
														if ($export_name[$col] == "starttime")
															$tmpstarttime = strtotime("$tmpdate $hh:$ii");
														else{
															if (isset($tmpstarttime)){
																if (strtotime(date('Y-m-d', $tmpstarttime) . " " . $tmptime) > $tmpstarttime )
																	$tmpendtime = strtotime("$tmpdate $hh:$ii");
																else{
																	$tmpendtime = "";
																	$isWrong=TRUE;
																}	
															}else $tmpendtime = "";
														}
							            			}else{
							            				if ($export_name[$col] == "starttime")
															$tmpstarttime = "";
														else
															$tmpendtime = "";
							            				$isWrong = TRUE;
							            			}			
							            		}
							            		break;
						        			}
						            		$isWrong = TRUE;
						            		break;
						            		
						            	case "weekly":
						            		$data="";
						            		$tmpweekly=-1;
						            		
						            		if ((!empty($tmpstarttime) && !empty($tmpendtime)) && ($tmpstartdate != $tmpenddate)){
						            			//echo "$val<br>";
						            			//var_dump(str_split($val));
							            		//if(!empty($val)){
							            			$arr = str_split($val);
							            			if (count($arr)>0){
							            				for ($i = 0; $i < count($arr); $i ++){
								            				if(!is_numeric($arr[$i]) || ($arr[$i] < 0 || $arr[$i] > 6)){
								            					$isWrong = TRUE;
								            				}
							            				}
							            			}else{
							            				$isWrong = TRUE;
							            			}
							            			
							            			$data="";
							            			if (!$isWrong){
							            				for ($i = 0; $i < 7; $i++){ 
								            				 $data .= (in_array($i, $arr) ? "1" : "0");
								            				 $tmpweekly=$data;	
							            				}	
							            			}
							            			
							            			break;
							            		//}else{
							            			//$isWrong = TRUE;
							            			//break;
							            		//}
						            		}
						            		$isWrong = false;
							            	break;		
						            			
						            	case "area":
						            		
						            		if (! empty($val)){
						            			$arr = explode("-", $val);
						            			if (count($arr) == 2){
						            				$id = sql_query1("SELECT id FROM $tbl_area Where area_name ='" . mysql_real_escape_string(trim((string)$arr[0]," ")) . "' LIMIT 1");
						            									  
							            			if ($id == -1){
							            				$isWrong = TRUE;
							            				break;
							            			}
							            			$data = $id;
							            			if (isset($id)){
							            				$id = sql_query1("SELECT id FROM $tbl_room Where room_name ='" . mysql_real_escape_string(trim((string)$arr[1]," ")) . "' AND area_id = $id LIMIT 1");
							            				if ($id == -1){
								            				$isWrong = TRUE;
								            				break;	
								            			}
								            			$tmproom = $id;
								            			$data .= "-".$id;
							            				break;
							            			}
						            			}	
						            		}
						            		$data='';
						            		$isWrong = TRUE;
						            		break;
		
						            	case "course":
						            		if (in_array(trim((string)$val," "), $cunitname)){
						            			$key = array_search(trim((string)$val," "), $cunitname);
						            			$data = $cunit[$key];
						            			break;
						            		}
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
							             if ($export_name[$col]=='starttime' || $export_name[$col]=='endtime'){
											$data = $utime;
										}
						             	 $tmp[$export_name[$col]] = (string)$data;	
						             	 if ($export_name[$col] == "remarks"){
						             	 	$tmp["xls_row"] = $row;
						             	 	$tmp_data["xls_row"] = $row;
						             	 }
						             	$str.= "<td>".htmlspecialchars($val)."</td>";
						             }else {
						             	$row_wrong ++;
						             	$total_wrong++;
						             	$str.= "<td style=\"background-color:#ffe6e6\">".htmlspecialchars($val)."</td>";
						             }      
						              $input_str.= "<input type=\"hidden\" value=\"$data\" name=\"data[".$num."][".$export_name[$col]."]\"  />";
						        }//End for loop $col
						        
						        if ((date('Y-m-d', $tmpstarttime) != date('Y-m-d', $tmpendtime)) && $tmpweekly == -1 ){
						        		$msg = "<ul style=\"margin:0px;font-weight:bold;\"><li>".get_vocab("missing_rep_week")."</li></ul>";
						        	}else{
						        		$msg="";
						        	}
						        	

						        if ((!empty($tmpstarttime)) && (!empty($tmpendtime)) && (!empty($tmproom)) && (!empty($tmpweekly)) && $row_wrong == 0){
						        	
									if ($tmpweekly!=-1 && $tmpweekly!=''){
										$w = date('w', $tmpstarttime);
										$tweek = str_split($tmpweekly);
										if($tweek[$w]!=1){
											 $tmpdd = strtotime("+1 day", $tmpstarttime) ;
											
											 while ($tmpdd <= $tmpendtime) {
											 	
											 	$ww = date('w', $tmpdd);
											 	if ($tweek[$ww]!=1){
											 		$tmpdd = strtotime("+1 day", $tmpdd);
											 	}else{
											 		$tmp['startdate'] = date('Y-m-d', $tmpdd);
											 		$tmp_data['startdate'] = date('Y-m-d', $tmpdd);
											 		$tmp['starttime'] = $tmpdd;
											 		$tmp_data['starttime'] = $tmpdd;
											 		break;
											 	} 
 											}	
										}
									}
							        $str1 = check_exsiting_record($tmpstarttime, $tmpendtime, $tmproom, $tmpweekly); 	
							        if ($num > 0){
	
							        	$isRepeatRow = check_repeat_excel_data($tmpArr, $tmp_data);
							        }					
							         if( $str1 ==''){
							         
							         	if (! empty($isRepeatRow)){	
							        		$str .= "<td style=\"background-color:#ffe6e6\">".get_vocab('excel_data_err').$isRepeatRow.get_vocab('bookingsfor').$msg."</td>";
							        		$total_wrong++;
							         	}else{
							         		$str .= "<td>$msg</td>";
							         	}
							        }else{			        	
							        	if ((!empty($tmpstarttime)) && (!empty($tmpendtime)) && (!empty($tmproom))){
							        		$str .= "<td style=\"background-color:#ffe6e6\"><ul style=\"margin:0px;\">$str1</ul>$msg</td>";
							        		$total_wrong++;	
							        	}else {
							        		$str .= "<td>$msg</td>";
							        	}			        			
							        }	
						        }else{
						        	$str .= "<td></td>";
						        }
						        if ($row_wrong == 0){
						        	array_push($dataArr, $tmp);
						       		array_push($tmpArr, $tmp_data);
						        }
						        
						        $str.= "</tr>";
						        $num++;

					    	}
					    	
					    	$tmpstartdate = '';
						    $tmpenddate = '';
						    $tmpstarttime = '';
						    $tmpendtime = '';
						    $tmproom = '';
							
					    	
					    }//End for loop $row
					    
					    $str.= "</table>";
					    
					    $str.= get_vocab("xls_system_remind1");
						//echo "<pre>";
						//var_dump($tmpArr);
						//echo "</pre>";
					    if ($total_wrong == 0){
					    	$str.= "<div style='text-align:center'>";
					    	$str.= "<button style=\"margin-right:10px;\" value=\"0\" name=\"confirm\" type=\"submit\" onclick=\"validation()\">".get_vocab("confirm")."</button>" . 
					    		   "<button value=\"1\" name=\"confirm\">".get_vocab("cancel")."</button>";
					    	$str.= "</div>";
					    	$_SESSION['input_data'] = $dataArr;
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

}//End $submit

if (isset($msg) && $_SERVER['REQUEST_METHOD'] == 'GET'){	
	if ($msg == "0"){
		$msg =  get_vocab("success_upload");
	}
}
?>
<script type="text/javascript" src="js/jquery-2.1.1.js"></script>
<script type="text/javascript">
$(document).ready(function (e) {
	$("#form1").submit(function (e) { 
		$("#message").hide();
	    $("#readdata").show();
	    $("#form1").submit();
	    set_message();
	});
});
</script>
<?php // End Batch Upload?>




<h2><?php echo get_vocab("administration") ?></h2>





<table border=1>
<tr>
<th><center><b><?php echo get_vocab("areas") ?></b></center></th>
<th><center><b><?php echo get_vocab("rooms") ?> <?php if(isset($area_name)) { echo get_vocab("in") . " " .
  htmlspecialchars($area_name); }?></b></center></th>
</tr>

<tr>
<td>
<?php 
# This cell has the areas
$res = sql_query("select id, area_name from $tbl_area order by area_name");
if (! $res) fatal_error(0, sql_error());

if (sql_count($res) == 0) {
	echo get_vocab("noareas");
} else {
	echo "<ul>";
	for ($i = 0; ($row = sql_row($res, $i)); $i++) {
		$area_name_q = urlencode($row[1]);
		echo "<li><a href=\"admin.php?area=$row[0]&area_name=$area_name_q\">"
			. htmlspecialchars($row[1]) . "</a> (<a href=\"edit_area_room.php?area=$row[0]\">" . get_vocab("edit") . "</a>) (<a href=\"del.php?type=area&area=$row[0]\">" .  get_vocab("delete") . "</a>)\n";
	}
	echo "</ul>";
}
?>
</td>
<td>
<?php
# This one has the rooms
if(isset($area)) {
	$res = sql_query("select id, room_name, description, capacity from $tbl_room where area_id=$area order by room_name");
	if (! $res) fatal_error(0, sql_error());
	if (sql_count($res) == 0) {
		echo get_vocab("norooms");
	} else {
		echo "<ul>";
		for ($i = 0; ($row = sql_row($res, $i)); $i++) {
			echo "<li>" . htmlspecialchars($row[1]) . "(" . htmlspecialchars($row[2])
			. ", $row[3]) (<a href=\"edit_area_room.php?room=$row[0]\">" . get_vocab("edit") . "</a>) (<a href=\"del.php?type=room&room=$row[0]\">" . get_vocab("delete") . "</a>)\n";
		}
		echo "</ul>";
	}
} else {
	echo get_vocab("noarea");
}
?>

</tr>
<tr>
<td>
<h3 ALIGN=CENTER><?php echo get_vocab("addarea") ?></h3>
<form action=add.php method=post>

<input type=hidden name=type value=area>

<TABLE>
<TR><TD><?php echo get_vocab("name") ?>:       </TD><TD><input type=text name=name></TD></TR>
</TABLE>
<input type=submit value="<?php echo get_vocab("addarea") ?>">
</form>
</td>

<td>
<?php if (0 != $area) { ?>
<h3 ALIGN=CENTER><?php echo get_vocab("addroom") ?></h3>
<form action=add.php method=post>

<input type=hidden name=type value=room>
<input type=hidden name=area value=<?php echo $area; ?>>

<TABLE>
<TR><TD><?php echo get_vocab("name") ?>:       </TD><TD><input type=text name=name></TD></TR>
<TR><TD><?php echo get_vocab("description") ?></TD><TD><input type=text name=description></TD></TR>
<TR><TD><?php echo get_vocab("capacity") ?>:   </TD><TD><input type=text name=capacity></TD></TR>
</TABLE>
<input type=submit value="<?php echo get_vocab("addroom") ?>">
</form>

<?php } else { echo "&nbsp;"; }?>
</td>
</tr>
</table>
<?php 
	/* 2016-07-04 Add batch upload function */
?>
<form action="" enctype="multipart/form-data" method="post" name="form1" id="form1">
	<table border="1" style="margin-top:10px;">
		<tr>
			<th><?php echo get_vocab("batch_upload") ; ?></th>
			<th width="100"></th>
		</tr>
		<tr>
			<td><input type="file" name="booking_detail" accept=".xlsx, .xls">
			<button name="submit" type="submit" value="1"><?php echo get_vocab("upload"); ?></button></td>
			<td align="center" style="vertical-align:middle"><a target="_blank" href="excel_download.php"><?php echo get_vocab("dl_template");?></a></td>
		</tr>	
	</table>
	
</form>

<script type="text/javascript" src="js/jquery-2.1.1.js"></script>
<script type="text/javascript">
$('.dl_xls').on('click', function(e) {
	e.preventDefault();
	<?php 
			$query = "SELECT a.area_name, r.room_name FROM $tbl_area a, $tbl_room r WHERE a.id = r.area_id ORDER BY a.area_name, r.room_name"; 
			$result = mysql_query($query) or die(mysql_error());
			$rowCount = 1; 
			$campus = array();
			while($row = mysql_fetch_array($result)){ 	
				$campus[$rowCount-1] = $row['area_name']."-".$row['room_name'];
				$rowCount++; 
			}
			$_SESSION['campus'] = $campus;
			$_SESSION['course'] = $cunitname;
	?>
    window.open(e.currentTarget.href);
});
</script>

<p id="readdata" style="display:none"><?php echo get_vocab("loading_data")?></p>
<?php 
	if (!empty($str)){
		echo $str;
	}
?>
<p id="message" style="color:#ff0000;"><?php echo $msg; ?></p>
<?php //End Batch upload function ?>

<br>
<?php echo get_vocab("browserlang") . " " . $HTTP_ACCEPT_LANGUAGE . " " . get_vocab("postbrowserlang") ; ?>

<?php include "trailer.inc" ?>


