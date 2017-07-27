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
if (getUserLevel($user) != 1.7)
{
	showAccessDenied($day, $month, $year, $area);
	exit();
}

# print the page header
print_header($day, $month, $year, $area);
?>

<h2>&nbsp;學院使用率</h2>

<form name="frmR2" method="post" action="repList2.php">
  
  <strong>請選擇校舍 (可選多個) :-</strong>  
  <table width="400" border="0" cellpadding="5">
    
     <!-- 2016-2017  --> 
    <?php 
    	
    	foreach($cpname as $key => $val){
    		$num = $key + 1;
    		$code = $val['code'];
    		$name = $val['name'];
    		echo "<tr>";
    		echo "<td><input type=\"checkbox\" name=\"c$num\" value=\"$code\" checked></td>";
    		echo "<td>$name</td>";
    		echo "</tr>";
    	}
    ?>
  </table>
  <br>

  <strong>請輸入查詢範圍 :-</strong>
  <table width="400" border="0" cellpadding="5">
    <tr>
      <td width="83">查詢日期由</td>
      <td width="301"><input type="text" name="dfrom" value="2015-09-01">&nbsp;(YYYY-MM-DD)</td>
    </tr>
    <tr>
      <td>查詢日期至</td>
      <td><input type="text" name="dto" value="2016-08-31">&nbsp;(YYYY-MM-DD)</td>
    </tr>
  </table>
  <br>

  <strong>請點選時段模式 :-</strong>
  <table width="400" border="0" cellpadding="5">
    <tr>
      <td width="30"><input name="mode" type="radio" value="day" checked></td>
      <td width="354">全日制時段 (星期一至五9AM-6PM) </td>
    </tr>
    <tr>
      <td><input name="mode" type="radio" value="night"></td>
      <td>兼讀制時段 (星期一至五6PM-10PM及星期六2PM-10PM) </td>
    </tr>
    <tr>
      <td><input name="mode" type="radio" value="all"></td>
      <td>全日制 + 兼讀制時段 </td>
    </tr>
  </table>

  <br>

  <p>
    <input type="submit" name="submit" value="送出">
    <input type="reset" name="reset" value="重設">
  </p>


</form>

<br>
<A href="javascript:history.go(-1)">&nbsp;返回上一頁</A>


<?php

?>