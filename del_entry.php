<?php
# $Id: del_entry.php,v 1.5.2.1 2005/03/29 13:26:16 jberanek Exp $

require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
include "mrbs_sql.inc";

#2016-07-15 Find Repeat ID For "add_remove_day" Function Use
function find_repeat_id($id, $starttime){
	global $tbl_entry;

	$repeat_id = sql_query1("SELECT repeat_id FROM $tbl_entry WHERE start_time = $starttime AND id=$id LIMIT 1");
	return $repeat_id;
}

#2016-07-15 
function add_remove_day($id, $repeat_id, $starttime){
	global $tbl_remove;	
	if ($repeat_id > 0){
		sql_query("INSERT INTO $tbl_remove(entry_id, repeat_id, start_time) VALUES($id , $repeat_id, $starttime)");
	}
}

if(getAuthorised(1) && ($info = mrbsGetEntryInfo($id)))
{

	$day   = strftime("%d", $info["start_time"]);
	$month = strftime("%m", $info["start_time"]);
	$year  = strftime("%Y", $info["start_time"]);
	$area  = mrbsGetRoomArea($info["room_id"]);

    if (MAIL_ADMIN_ON_DELETE)
    {
        include_once "functions_mail.inc";
        // Gather all fields values for use in emails.
        $mail_previous = getPreviousEntryData($id, $series);
    }
    
    if ($series==0)
    	$repeat_id = find_repeat_id($id, $info["start_time"]);

    sql_begin();
	$result = mrbsDelEntry(getUserName(), $id, $series, 1);
	sql_commit();
	if ($result)
	{

		if ($series==0 && $repeat_id > 0)
			add_remove_day($id, $repeat_id, $info["start_time"]);
			
        // Send a mail to the Administrator
        (MAIL_ADMIN_ON_DELETE) ? $result = notifyAdminOnDelete($mail_previous) : '';
        Header("Location: day.php?day=$day&month=$month&year=$year&area=$area");
		exit();
	}
}
// If you got this far then we got an access denied.
showAccessDenied($day, $month, $year, $area);
?>