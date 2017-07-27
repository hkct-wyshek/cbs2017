<?php
# $Id: report.php,v 1.22 2004/04/17 15:28:37 thierry_bo Exp $

require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";


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
if (getUserLevel($user) < 1.5)
{
	showAccessDenied($day, $month, $year, $area);
	exit();
}

# print the page header
print_header($day, $month, $year, $area);
?>

<h2>&nbsp;報表目錄</h2>
<ul>

<?php
if (getUserLevel($user) >= 1.9) {
?>
  <li><font size=4><A href="report1.php">實際及預算校舍使用率</A></font></li>
  <li><font size=4><A href="report7.php">課室使用率報告</A></font></li><br><br>
  <li><font size=4><A href="report3.php">每日課室表</A></font></li>
  <li><font size=4><A href="report5.php">日間課程每週課室表</A></font></li>
  <li><font size=4><A href="report6.php">兼讀/短期課程每週課室表</A></font></li>


<?php } ?>

<?php
if (getUserLevel($user) == 1.7) {
?>
  <li><font size=4><A href="report2.php">學院使用率</A></font></li>
<?php } ?>

</ul>

<br>
<A href="javascript:history.go(-1)">&nbsp;返回上一頁</A>

<?php

?>