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
if (getUserLevel($user) < 1.9)
{
	showAccessDenied($day, $month, $year, $area);
	exit();
}

# print the page header
print_header($day, $month, $year, $area);

?>

<script type="text/javascript" src="js/jquery-1.11.1.js"></script>
<script type="text/javascript" src="js/jquery-2.1.1.js"></script>
<script type="text/javascript" src="js/jquery-ui.js"></script>
<script type="text/javascript" src="js/jquery-ui.multidatespicker.js"></script>
<link rel="stylesheet" href="css/mdp.css" />
<script type="text/javascript">
$( document ).ready(function() {


	$('.datepicker1').datepicker({dateFormat: "yy-mm-dd", maxDate:	0});
	$('.datepicker2').datepicker({
		dateFormat: "yy-mm-dd",
		maxDate:	0
	});
	
});

$( function() {
    var dateFormat = "yy-mm-dd",
      from = $( ".datepicker1" )
        .datepicker({
          defaultDate: "+1w",
          changeMonth: true
        })
        .on( "change", function() {
          to.datepicker( "option", "minDate", getDate( this ) );
          to.datepicker( "option", "maxDate", 0 );
        }),
      to = $( ".datepicker2" ).datepicker({
        defaultDate: "+1w",
        changeMonth: true,
        maxDate:  0
      })
      .on( "change", function() {
        from.datepicker( "option", "maxDate", 0);
      });
 
    function getDate( element ) {
      var date;
      try {
        date = $.datepicker.parseDate( dateFormat, element.value );
        console.log(element.value);
      } catch( error ) {
        date = null;
      }
 
      return date;
    }
  } );
</script>

<h2>課室使用率報告</h2>

<?php
# <form name="frmR1" method="post" action="repList1.php">
?>
<form name="frmR1" method="post" action="repList7.php">
  
  <strong>請選擇校舍 :-</strong>  
  <table width="400" border="0" cellpadding="5"> 
    <!-- 2016-2017 --> 
    <?php 
    	
    	foreach($cpname as $key => $val){
    		$num = $key + 1;
    		$code = $val['code'];
    		$name = $val['name'];
    		echo "<tr>";
    		echo "<td><input type=\"radio\" name=\"c_code\" value=\"$code\" " . ($key == 0 ? "checked" : "") . "></td>";
    		echo "<td>$name</td>";
    		echo "</tr>";
    	}
    ?>
  </table>
  <br>
	
  <strong>請輸入查詢範圍:</strong>
  <table width="400" border="0" cellpadding="5">
    <tr>
      <td width="83">查詢日期由</td>
      <td width="301"><input class="datepicker1" type="text" name="dfrom" value="<?php echo date('Y-m-d', strtotime('-1 week')); ?>" >&nbsp;(YYYY-MM-DD)</td>
    </tr>
    <tr>
      <td>查詢日期至</td>
      <td><input class="datepicker2" type="text" name="dto" value="<?php echo date('Y-m-d');?>">&nbsp;(YYYY-MM-DD)</td>
    </tr>
  </table>
  <br>

  <strong>請點選時段 :-</strong>
  <table width="400" border="0" cellpadding="5">
    <tr>
      <td width="30"><input name="mode" type="radio" value="day" checked></td>
      <td width="354">日間(星期一至五9AM-7PM) </td>
    </tr>
    <tr>
      <td><input name="mode" type="radio" value="night"></td>
      <td>晚間 (星期一至五7PM-10PM及星期六至日9AM-10PM) </td>
    </tr>
    <tr>
      <td><input name="mode" type="radio" value="all"></td>
      <td>全日 </td>
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