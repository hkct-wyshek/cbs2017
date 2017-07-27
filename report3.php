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

<h2>每日課室表</h2>

<form name="frmR1" method="post" action="RmBookListByDay.php">
  
  <strong>請選擇校舍 :-</strong>  
  <table width="400" border="0" cellpadding="5">  
    <!-- 2016-2017 --> 
    <?php 
    	
    	foreach($cpname as $key => $val){
    		$num = $key + 1;
    		$code = $val['code'];
    		$name = $val['name'];
    		echo "<tr>";
    		echo "<td><input type=\"checkbox\" name=\"c$num\" value=\"$code\"></td>";
    		echo "<td>$name</td>";
    		echo "</tr>";
    	}
    ?>
  </table>
  <br>

  <strong>請輸入課室使用日期 :-</strong>
  <table width="400" border="0" cellpadding="5">
    <tr>
      <td width="83">查詢日期</td>
      <td width="301"><input type="text" name="useDate" value="<?php echo date("Y")."-".date("m")."-".date("d"); ?>">&nbsp;(YYYY-MM-DD)</td>
    </tr>
  </table>
  <br>

  <strong>請選擇顯示時段 :-</strong>
  <table width="400" border="0" cellpadding="5">
    <tr>
      <td width="30"><input type="radio" name="showPeriod" value="DAY"></td>
      <td width="354">日間 (7:00PM或之前開課)</td>
    </tr>
    <tr>
      <td width="30"><input type="radio" name="showPeriod" value="EVENING"></td>
      <td width="354">晚間 (6:00PM或之後開課)</td>
    </tr>
    <tr>
      <td width="30"><input type="radio" name="showPeriod" value="ALL" checked></td>
      <td width="354">全天</td>
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