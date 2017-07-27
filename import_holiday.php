<?php
session_start();
# $Id: admin.php,v 1.16.2.1 2005/03/29 13:26:15 jberanek Exp $

require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
include "mrbs_sql.inc";
include 'class.iCalReader.php';


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


if ($submit == 0){
	foreach ($data as $key => $val){
		$dd= $val["date"];
		$desc = $val["desc"];
		$sql = "INSERT INTO `mrbs_holiday`(`hday`, `description`) VALUES ('$dd','$desc')";
		sql_query($sql);
	}
}

$ical = new ical('holiday.ics');
$array= $ical->events();

// The ical date
$str = "<form method=\"post\">";
$num = 0;
foreach($array as $key => $val){
	$start = $array[$key]['DTSTART'];
	$desc = $array[$key]['SUMMARY'];
	$dd = sql_query1("SELECT hday FROM mrbs_holiday WHERE hday='$start'");


	if ($dd==-1){
		echo $array[$key]['DTSTART']."=>".$array[$key]['SUMMARY']."<br>";
		$str.= "<input type=\"hidden\" name=\"data[$num][date]\" value=\"$start\">";
		$str.= "<input type=\"hidden\" name=\"data[$num][desc]\" value=\"$desc\">";
		$num++;
	}
}
$str.= "<button type=\"submit\" value=\"0\" name=\"submit\">Upload</button>";
$str.= "</form>";

if ($num > 0){
	echo $str;
}else {
	echo "此文件沒有新假期可匯入";
}
?>