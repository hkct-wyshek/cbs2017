<?php
require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
include "mrbs_sql.inc";


if(!isset($day) or !isset($month) or !isset($year))
{
    $day   = date("d");
    $month = date("m");
    $year  = date("Y");
}

if(!getAuthorised(1)){
	showAccessDenied($day, $month, $year, $area);
	exit;
}

if (!isCourseAdmin()){
	showAccessDenied($day, $month, $year, $area);
	exit;
}

function process_data($data, $action){

	global $tbl_entry;
	foreach ($data as $key => $value){
		
		$starttime = $value['starttime'];
		$endtime = $value['endtime'];
		$class = $value['class_no'];
		$course = $value['course'];
		$description = $value['remarks'];
		if ($action == 0){
			
			$query = sql_query1("SELECT id FROM $tbl_entry e ". 
							    "WHERE start_time = '$starttime' AND end_time = '$endtime' ".
							    "AND create_by = '". getUserName()."' AND room_id = $course");
	
			if ($query == -1){
				return -1;
			}
		}elseif($action == 1){
			if ($class != getUserName() && $class != '') {
				$isUsedRoom = 1;
			}else{
				$class = strtoupper(getUserName());
				$isUsedRoom = 0;
			}
			$sql = "UPDATE $tbl_entry SET name='" . mysql_real_escape_string($class) . "', isUsedRoom = $isUsedRoom, ".
				   "description = '" . mysql_real_escape_string($description) . "' ".
				   "WHERE start_time = '$starttime' AND end_time = '$endtime' ".
				   "AND create_by = '". getUserName()."' AND room_id = $course";
			
			sql_command($sql);
		}
		
	}
	return 0;
}

if (isset($confirm) && !empty($_SESSION['C_List']) && $confirm == 0){
	
	$data = $_SESSION['C_List'];
	unset($_SESSION['C_List']);
	$err = process_data($data, 0);
	
	if ($err == 0){
		$err = process_data($data, 1);
		
		header("Location: maintainCourse.php?msg=0");
		exit;	
	}else{
		$hide_title  = 1;
		if(strlen($err)){
		    print_header($day, $month, $year, $area);
		    
		    echo "<H2>".get_vocab("excel_data_err2")."</H2>";
		}
		
		echo "<a href=\"maintainCourse.php\">".get_vocab("back_to_course")."</a><p>";
		include "trailer.inc"; 
	}	
	unset($confirm);	
}else{
	header("Location: maintainCourse.php");	
}

?>

