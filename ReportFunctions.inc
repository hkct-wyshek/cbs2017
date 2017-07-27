<?php
# $Id: report.php,v 1.22 2004/04/17 15:28:37 thierry_bo Exp $

function GetRoomName($roomid)
{
	global $db_host;
	global $db_login;
	global $db_password;
	global $db_database;
	$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
	
	// utf8 fixes
	//
	mysql_query("SET character_set_client=utf8", $fcn);
	mysql_query("SET character_set_connection=utf8", $fcn);
	mysql_query("SET character_set_results=utf8", $fcn);
	
	mysql_select_db($db_database,$fcn) or die ("Database not found!");
	$fsql = "select room_name from mrbs_room where id=\"".$roomid."\"";
	$frs = mysql_query($fsql, $fcn);
	mysql_data_seek($frs,0);
	$frow = mysql_fetch_row($frs);
	mysql_free_result($frs);
	mysql_close($fcn);
	return $frow[0];
}

function GetAreaName($areaid)
{
	global $db_host;
	global $db_login;
	global $db_password;
	global $db_database;
	$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
	
	// utf8 fixes
	//
	mysql_query("SET character_set_client=utf8", $fcn);
	mysql_query("SET character_set_connection=utf8", $fcn);
	mysql_query("SET character_set_results=utf8", $fcn);
	
	mysql_select_db($db_database,$fcn) or die ("Database not found!");
	$fsql = "select area_name from mrbs_area where id=\"".$areaid."\"";
	$frs = mysql_query($fsql, $fcn);
	mysql_data_seek($frs,0);
	$frow = mysql_fetch_row($frs);
	mysql_free_result($frs);
	mysql_close($fcn);
	return $frow[0];
}

function CalcSupply($from,$to,$mode,$special)
{
	$session = 0; 
	$sTime = mktime(0,0,0,substr($from,5,2),substr($from,8,2),substr($from,0,4));
	$eTime = mktime(23,59,59,substr($to,5,2),substr($to,8,2),substr($to,0,4));

	global $db_host;
	global $db_login;
	global $db_password;
	global $db_database;
	$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
	
	// utf8 fixes
	//
	mysql_query("SET character_set_client=utf8", $fcn);
	mysql_query("SET character_set_connection=utf8", $fcn);
	mysql_query("SET character_set_results=utf8", $fcn);
	
	mysql_select_db($db_database,$fcn) or die ("Database not found!");

	while ($sTime < $eTime) 
	{
		$cDate = date("Ymd", $sTime);
		$dow = getdate($sTime);
		$fsql = "select hday from mrbs_holiday where hday=\"".$cDate."\"";
		$frs = mysql_query($fsql, $fcn);
		$fnumRow = mysql_num_rows($frs);

		if ($fnumRow == 0) 
		{/*
			mysql_data_seek($frs,0);
			$frow = mysql_fetch_row($frs);
			if ($frow[0]==0.5)
			{
				if ( ($mode=="day") && ($special==false) )
				{ 
					if ($dow[wday]!=6) 
						{ $session = $session + 2; } 
					else 
						{ $session = $session + 0; }
				}
				if ($mode=="night") { $session = $session + 0; }
				
				if ( ($mode=="all") && ($special==false) ) 
				{ 
					if ($dow[wday]!=6) 
						{ $session = $session + 2; } 
					else 
						{ $session = $session + 0; }
				}
			}*/
		
		//else 
		//{
			if ( ($mode=="day") && ($special==false) )
			{ 
				if ($dow[wday]!=6 && $dow[wday]!=0) 
					{ $session = $session + 2; } 
				else 
					{ $session = $session + 0; }
			}
			if (($mode=="night")&& ($special==false))
			{ 
				if ($dow[wday]!=6 && $dow[wday]!=0) 
					{ $session = $session + 1; } 
				else if($dow[wday]!=0)
					{ $session = $session + 3; }
			}
			if (($mode=="night")&& ($special==true))
			{
				if ($dow[wday]!=6 && $dow[wday]!=0)
				{ $session = $session + 1; }
				else
				{ $session = $session + 3; }
			}
			if ( ($mode=="all") && ($special==false) )
			{ 
				if ($dow[wday]!=0) 
					{ $session = $session + 3; } 
				//else 
					//{ $session = $session + 3; }
			}
			if ( ($mode=="all") && ($special==true) )
			{
				
				$session = $session + 3;
				/*if ($dow[wday]!=0) 
					{ $session = $session + 1; } 
				else 
					{ $session = $session + 2; }
					*/
			}
			
			
			if ((($mode == "night") || ($mode == "all"))  && (($cDate == "20160915") || ($cDate == "20161221"))){
				$session = $session - 1;
			}
			
			
		//}
		}
		$sTime = $sTime + 86400;
	}		
	mysql_free_result($frs);
	mysql_close($fcn);

	return $session;
}

function CalcSlot($starttime, $endtime, $actual){
	if ((($starttime>="09:00:00") && ($starttime<"11:00:00")) && (($endtime>"09:00:00") && ($endtime<="11:00:00")))
	{ $actual = $actual + 0.5; }
	if ((($starttime>="09:00:00") && ($starttime<"11:00:00")) && (($endtime>"11:00:00") && ($endtime<="14:00:00")))
	{ $actual = $actual + 1; }
	if ((($starttime>="09:00:00") && ($starttime<"11:00:00")) && (($endtime>"14:00:00") && ($endtime<="16:00:00")))
	{ $actual = $actual + 1.5; }
	if ((($starttime>="09:00:00") && ($starttime<"11:00:00")) && (($endtime>"16:00:00") && ($endtime<="19:00:00")))
	{ $actual = $actual + 2; }
	if ((($starttime>="11:00:00") && ($starttime<"13:00:00")) && (($endtime>"11:00:00") && ($endtime<="14:00:00")))
	{ $actual = $actual + 0.5; }
	if ((($starttime>="11:00:00") && ($starttime<"13:00:00")) && (($endtime>"14:00:00") && ($endtime<="16:00:00")))
	{ $actual = $actual + 1; }
	if ((($starttime>="11:00:00") && ($starttime<"13:00:00")) && (($endtime>"16:00:00") && ($endtime<="19:00:00")))
	{ $actual = $actual + 1.5; }
	if ((($starttime>="13:00:00") && ($starttime<"16:00:00")) && (($endtime>"14:00:00") && ($endtime<="16:00:00")))
	{ $actual = $actual + 0.5; }
	if ((($starttime>="13:00:00") && ($starttime<"16:00:00")) && (($endtime>"16:00:00") && ($endtime<="19:00:00")))
	{ $actual = $actual + 1; }
	if ((($starttime>="16:00:00") && ($starttime<"18:00:00")) && (($endtime>"16:00:00") && ($endtime<="19:00:00")))
	{ $actual = $actual + 0.5; }
	
	return $actual;
}

function CalcActualUsage($cunit,$from,$to,$areaid,$roomid,$mode,$special) 
{

	global $db_host;
	global $db_login;
	global $db_password;
	global $db_database;
	$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
	
	// utf8 fixes
	//
	mysql_query("SET character_set_client=utf8", $fcn);
	mysql_query("SET character_set_connection=utf8", $fcn);
	mysql_query("SET character_set_results=utf8", $fcn);
	
	mysql_select_db($db_database,$fcn) or die ("Database not found!");
	$actual=0;
	$sql=     "select type, room_id, start_time, end_time,\n";
	$sql=$sql."date(from_unixtime(start_time)), time(from_unixtime(start_time)),\n";
	$sql=$sql."date(from_unixtime(end_time)), time(from_unixtime(end_time))\n";
	$sql=$sql."from mrbs_entry\n";
	$sql=$sql."where type=\"".$cunit."\" and room_id=".$roomid."\n";
	$sql=$sql."      and isUsedRoom=1 \n";
	$sql=$sql."      and date(from_unixtime(start_time))>=\"".$from."\"\n";
	$sql=$sql."      and date(from_unixtime(end_time))<=\"".$to."\"\n";
	$sql=$sql."order by type, date(from_unixtime(start_time)), time(from_unixtime(start_time))\n";

	//echo $sql."<BR><BR>";
	
	$rs = mysql_query($sql, $fcn);

	if ($rs == false) 
	{
		echo $sql."<br>\n";
		echo "Query failed!\n";
		exit();
	}
	$numRow = mysql_num_rows($rs);
	
	for ($i=0; $i<$numRow; $i++) 
	{
		mysql_data_seek($rs,$i);
		$row = mysql_fetch_row($rs);
		$dow = getdate($row[2]);
		$sql2 = "select duration from mrbs_holiday where hday=\"".str_replace("-", "", $row[4])."\"";
		$rs2 = mysql_query($sql2, $fcn);
		$numRow2 = mysql_num_rows($rs2);	
		if ( ($numRow2==0)) {
				
			if ( (($mode=="day") || ($mode=="all")) && ($dow[wday]!=6) && ($dow[wday]!=0))
			{
				$actual = CalcSlot($row[5], $row[7], $actual);
			}
			if (($mode=="night") || ($mode=="all"))
			{
				if (($dow[wday] >=1) && ($dow[wday] <=5))
				{
					if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
					{ $actual = $actual + 1; }
				}
				if ($dow[wday] == 6)
				{
					$actual = CalcSlot($row[5], $row[7], $actual);
					if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
					{ $actual = $actual + 1; }
				}
				$specialRoom = array(1, 2, 3, 5, 6, 7, 8, 9, 43, 44, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 113, 114, 115, 116, 122, 139, 142, 143);
		
				if (in_array($roomid, $specialRoom) && $dow[wday] == 0){
					$actual = CalcSlot($row[5], $row[7], $actual);
					if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
					{ $actual = $actual + 1; }
				}
				
				if ((($mode == "night") || ($mode == "all"))  && ((str_replace("-", "", $row[4]) == "20160915") || (str_replace("-", "", $row[4]) == "20161221")) && (($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00"))){
					$actual = $actual - 1;
				}
			}
		}
		//echo "$roomid - $cunit". $row[4]. $row[5]."-". $row[7]. " - $actual<BR>";
	}	
	
	
	mysql_free_result($rs);
	mysql_free_result($rs2);
	mysql_close($fcn);
	return $actual;
}

function CalcReserve($cunit,$from,$to,$areaid,$roomid,$mode,$special) 
{
	global $db_host;
	global $db_login;
	global $db_password;
	global $db_database;
	$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
	
	// utf8 fixes
	//
	mysql_query("SET character_set_client=utf8", $fcn);
	mysql_query("SET character_set_connection=utf8", $fcn);
	mysql_query("SET character_set_results=utf8", $fcn);
	
	mysql_select_db($db_database,$fcn) or die ("Database not found!");
	$reserve=0;
	$sql=     "select type, room_id, start_time, end_time,\n";
	$sql=$sql."date(from_unixtime(start_time)), time(from_unixtime(start_time)),\n";
	$sql=$sql."date(from_unixtime(end_time)), time(from_unixtime(end_time))\n";
	$sql=$sql."from mrbs_entry\n";
	$sql=$sql."where type=\"".$cunit."\" and room_id=".$roomid."\n";
	$sql=$sql."      and date(from_unixtime(start_time))>=\"".$from."\"\n";
	$sql=$sql."      and date(from_unixtime(end_time))<=\"".$to."\"\n";
	$sql=$sql."order by type, date(from_unixtime(start_time)), time(from_unixtime(start_time))\n";
	
	//echo $sql."<BR><BR>";

	$rs = mysql_query($sql, $fcn);
	if ($rs == false) 
	{
		echo $sql."<br>\n";
		echo "Query failed!\n";
		exit();
	}
	$numRow = mysql_num_rows($rs);
	for ($i=0; $i<$numRow; $i++) 
	{
		mysql_data_seek($rs,$i);
		$row = mysql_fetch_row($rs);
		$dow = getdate($row[2]);
		$sql2 = "select duration from mrbs_holiday where hday=\"".str_replace("-", "", $row[4])."\"";
		$rs2 = mysql_query($sql2, $fcn);
		$numRow2 = mysql_num_rows($rs2);
	//	echo $numRow2 . " ";
		if ( ($numRow2==0) ) {
			
			if ( (($mode=="day") || ($mode=="all")) && ($dow[wday]!=6) && ($dow[wday]!=0))
			{
				$reserve = CalcSlot($row[5], $row[7], $reserve);
			}
			if (($mode=="night") || ($mode=="all")) 
			{
				if (($dow[wday] >=1) && ($dow[wday] <=5))
				{
					if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
						{ $reserve = $reserve + 1; }
				}
				if ($dow[wday] == 6)
				{
					$reserve = CalcSlot($row[5], $row[7], $reserve);
					if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
					{ $reserve = $reserve + 1; }
				}
				$specialRoom = array(1, 2, 3, 5, 6, 7, 8, 9, 43, 44, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 113, 114, 115, 116, 122, 139, 142, 143);
				
				if (in_array($roomid, $specialRoom)){
					if ($dow[wday] == 0)
					{
						$reserve = CalcSlot($row[5], $row[7], $reserve);
						if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
						{ $reserve = $reserve + 1; }
					}
				}
				
				
				if ((($mode == "night") || ($mode == "all"))  && ((str_replace("-", "", $row[4]) == "20160915") || (str_replace("-", "", $row[4]) == "20161221")) && (($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00"))){
					$reserve = $reserve - 1;
				}
			}
		}
		
		//echo $roomid . " " . $cunit .' '. $row[4]. $row[5]."-". $row[7]. $reserve . "<BR>";
	}
	
	mysql_free_result($rs);
	mysql_free_result($rs2);
	mysql_close($fcn);
	return $reserve;
}

function CalcExtraUsage($cunit,$from,$to,$areaid,$roomid,$mode,$special) 
{
	global $db_host;
	global $db_login;
	global $db_password;
	global $db_database;
	$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
	
	// utf8 fixes
	//
	mysql_query("SET character_set_client=utf8", $fcn);
	mysql_query("SET character_set_connection=utf8", $fcn);
	mysql_query("SET character_set_results=utf8", $fcn);
	
	mysql_select_db($db_database,$fcn) or die ("Database not found!");
	$extra=0;
	$sql=     "select type, room_id, start_time, end_time,\n";
	$sql=$sql."date(from_unixtime(start_time)), time(from_unixtime(start_time)),\n";
	$sql=$sql."date(from_unixtime(end_time)), time(from_unixtime(end_time))\n";
	$sql=$sql."from mrbs_entry\n";
	$sql=$sql."where type=\"OTHER\" and room_id=".$roomid."\n";
	$sql=$sql."      and isUsedRoom=1 \n";
	//$sql=$sql."      or name like \"%[*".$cunit."%\" )\n";
	$sql=$sql."      and date(from_unixtime(start_time))>=\"".$from."\"\n";
	$sql=$sql."      and date(from_unixtime(end_time))<=\"".$to."\"\n";
	$sql=$sql."order by type, date(from_unixtime(start_time)), time(from_unixtime(start_time))\n";
	$rs = mysql_query($sql, $fcn);
	if ($rs == false) 
	{
		echo $sql."<br>\n";
		echo "Query failed!\n";
		exit();
	}
	$numRow = mysql_num_rows($rs);
	for ($i=0; $i<$numRow; $i++) 
	{
		mysql_data_seek($rs,$i);
		$row = mysql_fetch_row($rs);
		$dow = getdate($row[2]);
		$sql2 = "select duration from mrbs_holiday where hday=\"".$row[4]."\"";
		$rs2 = mysql_query($sql2, $fcn);
		$numRow2 = mysql_num_rows($rs2);
		if ($numRow2 >0 ) {
			mysql_data_seek($rs2,0);
			$row2 = mysql_fetch_row($rs2);
			if ($row2[0]==0.5) { $NoEveningOnly=true; } else { $NoEveningOnly=false; }
		} else { $NoEveningOnly=false; }
		if ( ($numRow2==0) || ($NoEveningOnly==true) ) {
			if ( (($mode=="day") || ($mode=="all")) && ($dow[wday]!=6) )
			{
				if ((($row[5]>="09:00:00") && ($row[5]<"11:00:00")) && (($row[7]>"09:00:00") && ($row[7]<="11:00:00")))
					{ $extra = $extra + 0.5; }
				if ((($row[5]>="09:00:00") && ($row[5]<"11:00:00")) && (($row[7]>"11:00:00") && ($row[7]<="14:00:00")))
					{ $extra = $extra + 1; }
				if ((($row[5]>="09:00:00") && ($row[5]<"11:00:00")) && (($row[7]>"14:00:00") && ($row[7]<="16:00:00")))
					{ $extra = $extra + 1.5; }
				if ((($row[5]>="09:00:00") && ($row[5]<"11:00:00")) && (($row[7]>"16:00:00") && ($row[7]<="19:00:00")))
					{ $extra = $extra + 2; }
				if ((($row[5]>="11:00:00") && ($row[5]<"13:00:00")) && (($row[7]>"11:00:00") && ($row[7]<="14:00:00")))
					{ $extra = $extra + 0.5; }
				if ((($row[5]>="11:00:00") && ($row[5]<"13:00:00")) && (($row[7]>"14:00:00") && ($row[7]<="16:00:00")))
					{ $extra = $extra + 1; }
				if ((($row[5]>="11:00:00") && ($row[5]<"13:00:00")) && (($row[7]>"16:00:00") && ($row[7]<="19:00:00")))
					{ $extra = $extra + 1.5; }
				if ((($row[5]>="13:00:00") && ($row[5]<"16:00:00")) && (($row[7]>"14:00:00") && ($row[7]<="16:00:00")))
					{ $extra = $extra + 0.5; }
				if ((($row[5]>="13:00:00") && ($row[5]<"16:00:00")) && (($row[7]>"16:00:00") && ($row[7]<="19:00:00")))
					{ $extra = $extra + 1; }
				if ((($row[5]>="16:00:00") && ($row[5]<"18:00:00")) && (($row[7]>"16:00:00") && ($row[7]<="19:00:00")))
					{ $extra = $extra + 0.5; }
			}
		}
		if ($numRow2==0) {
			if (($mode=="night") || ($mode=="all")) 
			{
				if (($dow[wday] >=1) && ($dow[wday] <=5))
				{
					if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
						{ $extra = $extra + 1; }
				}
				if ($dow[wday] == 6)
				{
					if ((($row[5]>="13:00:00") && ($row[5]<"18:00:00")) && (($row[7]>"14:00:00") && ($row[7]<="19:00:00")))
						{ $extra = $extra + 1; }
					if ((($row[5]>="13:00:00") && ($row[5]<"18:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
						{ $extra = $extra + 2; }
					if ((($row[5]>="18:00:00") && ($row[5]<"22:00:00")) && (($row[7]>"19:00:00") && ($row[7]<="22:00:00")))
						{ $extra = $extra + 1; }
				}	
			}
		}
	}	
	mysql_free_result($rs);
	mysql_free_result($rs2);
	mysql_close($fcn);
	return 0;
}

