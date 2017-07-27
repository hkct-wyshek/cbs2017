<?php
require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
include "mrbs_sql.inc";

ini_set('max_execution_time', 1000);
if(!isset($day) or !isset($month) or !isset($year))
{
    $day   = date("d");
    $month = date("m");
    $year  = date("Y");
}

if(empty($area))
    $area = get_default_area();

if(!getAuthorised(2))
{
    showAccessDenied($day, $month, $year, $area);
    exit;
}

if(!getWritable($create_by, getUserName()))
{
    showAccessDenied($day, $month, $year, $area);
    exit;
}

function process_data($data, $action){
	
	

                                                                                                    
	foreach ($data as $key => $value){
		
		$starttime = $value['starttime'];
		$endtime = $value['endtime'];
		$entry_type = 0;
		$repeat_id = 0;
		$r = explode("-", $value['area']);
		$room_id = $r[1];
		$owner = $value['course'];
		if ($value['class_no'] == ''){
			$value['class_no'] = $value['course'];
		}
		$rep_opt = $value['weekly'];
		$name = $value['class_no'];
		$type = $value['course'];
		$description = $value['remarks'];
		$isUsedRoom = 1;
		if ($type == $name || $type == '' || empty($type)){
			$isUsedRoom = 0;
		}
		
		//echo "starttime: $starttime, endtime: $endtime, room: $room_id, owner: $owner, Repeat Opt: $rep_opt, Class no: $name, Course: $type, Remarks: $description<br>";
		if ($action == 0){
			$err.= check_exsiting_record($starttime, $endtime, $room_id, $rep_opt);	
		}else{
			$time = date('H:i', $endtime);
			$end = strtotime(date('Y-m-d', $starttime)." $time");
			
			if (!empty($rep_opt) && $rep_opt!=-1 && date('Y-m-d', $starttime)!=date('Y-m-d', $endtime)){
				mrbsCreateRepeatingEntrys($starttime,$end , 2, $endtime, $rep_opt,
                                      				$room_id, $owner, $name, $type, $description, "", $isUsedRoom);  
			}else {
				mrbsCreateSingleEntry($starttime, $end, 0, 0, $room_id, $owner, $name, $type, $description, $isUsedRoom);
			}                       		
		}	
	}
	return $err;
}

if (isset($confirm) && !empty($_SESSION['input_data']) && $confirm == 0){
	
	$data = $_SESSION['input_data'];
	unset($_SESSION['input_data']);
	
	$err = process_data($data, 0);
	
	if (empty($err)){
		$err = process_data($data, 1);
		header("Location: admin.php?msg=0");
		exit;	
	}else{
		$hide_title  = 1;
		if(strlen($err)){
		    print_header($day, $month, $year, $area);
		    
		    echo "<H2>" . get_vocab("sched_conflict") . "</H2>";
		    echo get_vocab("conflict");
		    echo "<UL>";		    
		    echo $err;
		    echo "</UL>";
		}
		
		echo "<a href=\"admin.php\">".get_vocab("backadmin")."</a><p>";
		include "trailer.inc"; 
	}		
}else{
	header("Location: admin.php");	
}

?>

