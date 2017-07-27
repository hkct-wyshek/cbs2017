<?php
# $Id: edit_entry.php,v 1.30.2.2 2005/03/29 13:26:27 jberanek Exp $10/9/2006

require_once('grab_globals.inc.php');
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
//include 'class.iCalReader.php';

global $twentyfourhour_format;

#If we dont know the right date then make it up
if(!isset($day) or !isset($month) or !isset($year))
{
	$day   = date("d");
	$month = date("m");
	$year  = date("Y");
}
if(empty($area))
	$area = get_default_area();
if(!isset($edit_type))
	$edit_type = "";

if (!isset($id)) {  //no book id, new booing require admin right
	if(!getAuthorised(2))
	{
		showAccessDenied($day, $month, $year, $area);
		exit;
	}
} else {
	if(!getAuthorised(1))
	{
		showAccessDenied($day, $month, $year, $area);
		exit;
	}
}

# This page will either add or modify a booking

# We need to know:
#  Name of booker
#  Description of meeting
#  Date (option select box for day, month, year)
#  Time
#  Duration
#  Internal/External

# Firstly we need to know if this is a new booking or modifying an old one
# and if it's a modification we need to get all the old data from the db.
# If we had $id passed in then it's a modification.
if (isset($id))
{
	$sql = "select name, create_by, description, start_time, end_time,
	        type, room_id, entry_type, repeat_id, isUsedRoom from $tbl_entry where id=$id";
	
	$res = sql_query($sql);
	if (! $res) fatal_error(1, sql_error());
	if (sql_count($res) != 1) fatal_error(1, get_vocab("entryid") . $id . get_vocab("not_found"));
	
	$row = sql_row($res, 0);
	sql_free($res);
# Note: Removed stripslashes() calls from name and description. Previous
# versions of MRBS mistakenly had the backslash-escapes in the actual database
# records because of an extra addslashes going on. Fix your database and
# leave this code alone, please.
	$name        = $row[0];
//	$create_by   = $row[1];
	$create_by   = getUserName();
	$description = $row[2];
	$start_day   = strftime('%d', $row[3]);
	$start_month = strftime('%m', $row[3]);
	$start_year  = strftime('%Y', $row[3]);
	$start_hour  = strftime('%H', $row[3]);
	$start_min   = strftime('%M', $row[3]);
	$duration    = $row[4] - $row[3] - cross_dst($row[3], $row[4]);
	$type        = $row[5];
	$room_id     = $row[6];
	$entry_type  = $row[7];
	$rep_id      = $row[8];
	$is_used	 = ($row[9] == 1 ? "Checked" : "");
	
	if($entry_type >= 1)
	{
		$sql = "SELECT rep_type, start_time, end_date, rep_opt, rep_num_weeks
		        FROM $tbl_repeat WHERE id=$rep_id";
		
		$res = sql_query($sql);
		if (! $res) fatal_error(1, sql_error());
		if (sql_count($res) != 1) fatal_error(1, get_vocab("repeat_id") . $rep_id . get_vocab("not_found"));
		
		$row = sql_row($res, 0);
		sql_free($res);
		
		$rep_type = $row[0];

		if($edit_type == "series")
		{
			$start_day   = (int)strftime('%d', $row[1]);
			$start_month = (int)strftime('%m', $row[1]);
			$start_year  = (int)strftime('%Y', $row[1]);
			
			$rep_end_day   = (int)strftime('%d', $row[2]);
			$rep_end_month = (int)strftime('%m', $row[2]);
			$rep_end_year  = (int)strftime('%Y', $row[2]);
			
			switch($rep_type)
			{
				case 2:
				case 6:
					$rep_day[0] = $row[3][0] != "0";
					$rep_day[1] = $row[3][1] != "0";
					$rep_day[2] = $row[3][2] != "0";
					$rep_day[3] = $row[3][3] != "0";
					$rep_day[4] = $row[3][4] != "0";
					$rep_day[5] = $row[3][5] != "0";
					$rep_day[6] = $row[3][6] != "0";

					if ($rep_type == 6)
					{
						$rep_num_weeks = $row[4];
					}
					
					break;
				
				default:
					$rep_day = array(0, 0, 0, 0, 0, 0, 0);
			}
		}
		else
		{
			$rep_type     = $row[0];
			$rep_end_date = utf8_strftime('%A %d %B %Y',$row[2]);
			$rep_opt      = $row[3];
		}
	}
}
else
{
	# It is a new booking. The data comes from whichever button the user clicked
	$edit_type   = "series";
	$name        = "";
	$create_by   = getUserName();
	$description = "";
	$start_day   = $day;
	$start_month = $month;
	$start_year  = $year;
    // Avoid notices for $hour and $minute if periods is enabled
    (isset($hour)) ? $start_hour = $hour : '';
	(isset($minute)) ? $start_min = $minute : '';
	$duration    = ($enable_periods ? 60 : 60 * 60);
	$type        = "";
	$room_id     = $room;
	$is_used	 = "";
    unset($id);

	$rep_id        = 0;
	$rep_type      = 0;
	$rep_end_day   = $day;
	$rep_end_month = $month;
	$rep_end_year  = $year;
	$rep_day       = array(0, 0, 0, 0, 0, 0, 0);
}

# These next 4 if statements handle the situation where
# this page has been accessed directly and no arguments have
# been passed to it.
# If we have not been provided with a room_id
if( empty( $room_id ) )
{
	$sql = "select id from $tbl_room limit 1";
	$res = sql_query($sql);
	$row = sql_row($res, 0);
	$room_id = $row[0];

}

# If we have not been provided with starting time
if( empty( $start_hour ) && $morningstarts < 10 )
	$start_hour = "0$morningstarts";

if( empty( $start_hour ) )
	$start_hour = "$morningstarts";

if( empty( $start_min ) )
	$start_min = "00";

// Remove "Undefined variable" notice
if (!isset($rep_num_weeks))
{
    $rep_num_weeks = "";
}

$enable_periods ? toPeriodString($start_min, $duration, $dur_units) : toTimeString($duration, $dur_units);

#now that we know all the data to fill the form with we start drawing it

if(!getWritable($create_by, getUserName()))
{
	showAccessDenied($day, $month, $year, $area);
	exit;
}

if (isCourseAdmin()){
	$edit_date = mktime(0, 0, 0, $start_month, $start_day, $start_year);
	$today = mktime(0, 0, 0, date('m'), date('d'), date('Y'));
	if($today > $edit_date){
		showAccessDenied($day, $month, $year, $area);
		exit;
	}
}


print_header($day, $month, $year, $area);

?>

<SCRIPT LANGUAGE="JavaScript">
// do a little form verifying
function validate_and_submit ()
{

  if(document.forms["main"].name.value.length > 80)
  {
    alert ( "<?php echo get_vocab("exceed_max")?>");
    return false;
  }

  // null strings and spaces only strings not allowed
//  if(/(^$)|(^\s+$)/.test(document.forms["main"].name.value))
//  {
//    alert ( "<?php echo get_vocab("you_have_not_entered") . '\n' . get_vocab("brief_description") ?>");
//    return false;
//  }
  <?php if( ! $enable_periods ) { ?>

  h = parseInt(document.forms["main"].hour.value);
  m = parseInt(document.forms["main"].minute.value);

  if(h > 23 || m > 59)
  {
    alert ("<?php echo get_vocab("you_have_not_entered") . '\n' . get_vocab("valid_time_of_day") ?>");
    return false;
  }
  <?php } ?>
0
  // check form element exist before trying to access it
  if( document.forms["main"].id )
    i1 = parseInt(document.forms["main"].id.value);
  else
    i1 = 0;

  i2 = parseInt(document.forms["main"].rep_id.value);
  if ( document.forms["main"].rep_num_weeks)
  {
  	n = parseInt(document.forms["main"].rep_num_weeks.value);
  }
  if ((!i1 || (i1 && i2)) && document.forms["main"].rep_type && document.forms["main"].rep_type[6].checked && (!n || n < 2))
  {
    alert("<?php echo get_vocab("you_have_not_entered") . '\n' . get_vocab("useful_n-weekly_value") ?>");
    return false;
  }

  // check that a room(s) has been selected
  // this is needed as edit_entry_handler does not check that a room(s)
  // has been chosen
  if( document.forms["main"].elements['rooms'].selectedIndex == -1 )
  {
    alert("<?php echo get_vocab("you_have_not_selected") . '\n' . get_vocab("valid_room") ?>");
    return false;
  }

<?php if (isAdmin()) { ?>

  if (document.forms["main"].ampm[0].checked) { x = document.forms["main"].ampm[0].value; }
  if (document.forms["main"].ampm[1].checked) { x = document.forms["main"].ampm[1].value; }

  if ( ((parseInt(document.main.hour.value) < 9) && (x == "am") && (document.main.hour.value != "09")) || ((parseInt(document.main.hour.value) ==12 ) && (x == "am")) )
  {
    alert("Invalid lesson start time!");
    return false;
  }

  if ( (parseInt(document.main.hour.value) >= 10) && (x == "pm") && (document.main.hour.value != "12") )
  {
    alert("Invalid lesson start time!");
    return false;
  }

<?php } ?>

  // Form submit can take some times, especially if mails are enabled and
  // there are more than one recipient. To avoid users doing weird things
  // like clicking more than one time on submit button, we hide it as soon
  // it is clicked.

<?php
  if ( (!isset($id)) || (strtoupper(getUserName())=="DFM") ) {
?>
  msg = "You have chosen the [" + document.main.type.value + "] unit in this booking.\n";
  msg = msg + "If continue to save, press OK.";
<?php
  } else {
?>
  msg = "Class = [" + document.main.name.value + "].\n";
  msg = msg + "If continue to save, press OK.";  
<?php
  }
?>

  goAhead = window.confirm(msg);
  if (goAhead) {
	  document.forms["main"].save_button.disabled="true";
	  if (document.forms["main"].create_by.value.toUpperCase() == "DFM")
		    document.forms["main"].create_by.value = document.forms["main"].type.value;

	  if (document.forms["main"].name.value.length == 0) 
			document.forms["main"].name.value = document.forms["main"].type.value;

<?php
	if ( (isset($id)) && (strtoupper(getUserName())=="DFM") ) {
?>
			if (document.forms["main"].name.value == document.forms["main"].create_by.value) {
				document.forms["main"].name.value = document.forms["main"].type.value;
			}	
		    document.forms["main"].create_by.value = document.forms["main"].type.value;
<?php
	}
	
	/** 2016-07-06 Get Datepicker Value */
?>

	var dates = $('#simpliest-usage').multiDatesPicker('value');
	document.forms["main"].remove_date.value = dates;

//  document.forms["main"].save_button.disabled="true";
  
//  if (document.forms["main"].create_by.value == "DFM") 
//    document.forms["main"].create_by.value = document.forms["main"].type.value;

//  alert(document.forms["main"].type.value);


  // would be nice to also check date to not allow Feb 31, etc...
  	

    document.forms["main"].submit();
    return true;
  
  } else {

	return false;

  }


}
function OnAllDayClick(allday) // Executed when the user clicks on the all_day checkbox.
{
  form = document.forms["main"];
  if (allday.checked) // If checking the box...
  {
    <?php if( ! $enable_periods ) { ?>
      form.hour.value = "00";
      form.minute.value = "00";
    <?php } ?>
    if (form.dur_units.value!="days") // Don't change it if the user already did.
    {
      form.duration.value = "1";
      form.dur_units.value = "days";
    }
  }
}

//
</SCRIPT>
<?php 

function get_holiday(){
$sql = "SELECT hday FROM mrbs_holiday";
$res = sql_query($sql);
if (! $res) fatal_error(0, sql_error());
for ($i = 0; ($row = sql_row($res, $i)); $i++) {
	$holiday_arr[$i] = $row[0];
}
return $holiday_arr;
}



/* 2016-07-07 Multiple Date Picker For Remove Specific Date. */
?>
<script type="text/javascript" src="js/jquery-1.11.1.js"></script>
<script type="text/javascript" src="js/jquery-2.1.1.js"></script>
<script type="text/javascript" src="js/jquery-ui-1.11.1.js"></script>
<script type="text/javascript" src="js/jquery-ui.multidatespicker.js"></script>
<link rel="stylesheet" href="css/mdp.css" />

<script type="text/javascript">

function weekday(){
	var x = document.getElementsByName("rm_week");
	if (x[7].checked){
		x[7].checked = false;
	}
}
function everyday(){
	var x = document.getElementsByName("rm_week");
	if (x[7].checked){
		for (i = 0; i < 7; i++) {
			x[i].checked = false;
		}
	}
}
function get_date(type){
	return $('form[name="main"] select[name="' + type + '"]').val();
}

function del_not_exist_date(curr_val, addDD, dates, data){
	for(i=0; i<curr_val.length; i++){
		var DD_num = addDD.indexOf(curr_val[i]);
		var dates_num = dates.indexOf(curr_val[i]);
		var data_num = data.indexOf(curr_val[i]);
		if (DD_num < 0 && dates_num < 0 && data_num < 0 ){
			$('#simpliest-usage').multiDatesPicker('removeDates', curr_val[i]);
		}
	}
}

function clear_all_selected_date(){
	var $cal = $('#simpliest-usage');
	var start_date = get_date("day") + "/" + get_date("month") + "/" + get_date("year"); 
	var end_date = get_date("rep_end_day") + "/" + get_date("rep_end_month")  + "/" + get_date("rep_end_year");

	$cal.datepicker('destroy');
	
	var d1 = $.datepicker.parseDate('dd/mm/yy', start_date);
	var d2 = $.datepicker.parseDate('dd/mm/yy', end_date);

	var data = calcBusinessDays(d1,d2);

	if (data.length>0){
		$cal.multiDatesPicker( {
		    minDate: get_date("month") + "/" + get_date("day") + "/" + get_date("year"),
		    maxDate: get_date("rep_end_month") + "/" + get_date("rep_end_day")  + "/" + get_date("rep_end_year"),
		    addDates: data,
		    addDisabledDates: data
		});	

	}else{
		$cal.multiDatesPicker( {
			minDate: get_date("month") + "/" + get_date("day") + "/" + get_date("year"),
		    maxDate: get_date("rep_end_month") + "/" + get_date("rep_end_day")  + "/" + get_date("rep_end_year")
		});
	}

	var curr_val = $cal.multiDatesPicker('getDates');
	del_not_exist_date(curr_val, "", "", data);
}

function clear_range_input(){
	var x = document.getElementsByName("rm_week");
	$('#from_date').val('');
	$('#to_date').val('');
	for (i = 0; i < x.length; i++) {
		x[i].checked = false;
	}

	var data = calcBusinessDays(arr[3],arr[4]);
	$cal.multiDatesPicker('addDates', dates);
}

function add_range_date(){

	var arr = check_date_range();
	
	if (arr != false){
		var $cal = $('#simpliest-usage');
		var addDD = $cal.multiDatesPicker('getDates');
		$cal.datepicker('destroy');
		var data = calcBusinessDays(arr[3],arr[4]);
		var dates = new Array();
		var date  = arr[0];
		dates.push($.datepicker.formatDate('mm/dd/yy', new Date(date)));

		if (data.length>0){
			$cal.multiDatesPicker( {
			    minDate: get_date("month") + "/" + get_date("day") + "/" + get_date("year"),
			    maxDate: get_date("rep_end_month") + "/" + get_date("rep_end_day")  + "/" + get_date("rep_end_year"),
			    addDates: data,
			    addDisabledDates: data
			});	
		}else{
			$cal.multiDatesPicker( {
				minDate: get_date("month") + "/" + get_date("day") + "/" + get_date("year"),
			    maxDate: get_date("rep_end_month") + "/" + get_date("rep_end_day")  + "/" + get_date("rep_end_year")
			});
		}

		if (dates.length > 0){
			$cal.multiDatesPicker('addDates', dates);
		}	

		var curr_val = $cal.multiDatesPicker('getDates');
		del_not_exist_date(curr_val, addDD, dates, data);
		clear_range_input()	;
	}			
}

function check_date_range(){
	
	var fdate = $('#from_date').val();
	var fy = fdate.substr(0,4), fm = fdate.substr(4,2) - 1, fd = fdate.substr(6,2);
	if (fd.length == 2){
		var f_D = new Date(fy,fm,fd);
		var start_date = get_date("day") + "/" + get_date("month") + "/" + get_date("year"); 
		var end_date = get_date("rep_end_day") + "/" + get_date("rep_end_month")  + "/" + get_date("rep_end_year");
		var d1 = $.datepicker.parseDate('dd/mm/yy', start_date); //start date
		var d2 = $.datepicker.parseDate('dd/mm/yy', end_date); // end date
		var minDD = d1.getFullYear() + "" + ('0' + (d1.getMonth()+1)).slice(-2) + "" + ('0' + d1.getDate()).slice(-2);
		var maxDD = d2.getFullYear() + "" + ('0' + (d2.getMonth()+1)).slice(-2) + "" + ('0' + d2.getDate()).slice(-2);
		var arr = new Array(f_D, 
							f_D, 
							"", 
							d1, 
							d2,
							minDD,
							maxDD);
		return arr;
	} 
    alert("請輸入正確的日期"); 
	return false;
}


function remove_range_date(){
	
	var arr = check_date_range();
	if (arr != false){
		var $cal = $('#simpliest-usage');
		var addDD = $cal.multiDatesPicker('getDates');
		var data = calcBusinessDays(arr[3], arr[4]);
		
		var date  = arr[0]; //from date
		$('#simpliest-usage').multiDatesPicker('removeDates', $.datepicker.formatDate('mm/dd/yy', new Date(date)));
		if (data.length>0){
			$cal.multiDatesPicker( {
			    minDate: get_date("month") + "/" + get_date("day") + "/" + get_date("year"),
			    maxDate: get_date("rep_end_month") + "/" + get_date("rep_end_day")  + "/" + get_date("rep_end_year"),
			    addDates: data,
			    addDisabledDates: data
			});	
	
		}
		var curr_val = $cal.multiDatesPicker('getDates');
		del_not_exist_date(curr_val, addDD, "", "");
		clear_range_input();
	}			
}

function calcBusinessDays(dDate1, dDate2) {
    if (dDate1 > dDate2) return false;
    var date  = dDate1;
    var dates = [];
    var holiday = [<?php echo '"'.implode('","',  get_holiday() ).'"' ?>];
    
    while (date <= dDate2) {
        if( $.inArray($.datepicker.formatDate('yymmdd', new Date(date)), holiday) > -1){
        	dates.push($.datepicker.formatDate('mm/dd/yy', new Date(date)));
        }
        
        date.setDate( date.getDate() + 1 );
    }  
 
    return dates;
}


$( document ).ready(function() {

	var d1 = $.datepicker.parseDate('dd/mm/yy', <?php echo "'".date("d/m/Y", strtotime("$start_day-$start_month-$start_year"))."'"?>);
	var d2 = $.datepicker.parseDate('dd/mm/yy', <?php echo "'".date("d/m/Y", strtotime("$rep_end_day-$rep_end_month-$rep_end_year"))."'"?>);
	var data = calcBusinessDays(d1,d2);
	if (data != ''){
		$('#simpliest-usage').multiDatesPicker('addDates', data);
	}else{
		data[0] = '';
	}
	
	$('#simpliest-usage').multiDatesPicker({
		
		<?php 
				echo "minDate: '$start_month/$start_day/$start_year',";
			    echo "maxDate: '$rep_end_month/$rep_end_day/$rep_end_year',";
				$sql = "SELECT distinct(start_time) FROM $tbl_remove WHERE repeat_id = $rep_id";
				$res = sql_query($sql);
				if (! $res) fatal_error(0, sql_error());
				
				$dd = array();
				for ($i = 0; ($row = sql_row($res, $i)); $i++) {
					$dd[$i] = "'".date('m/d/Y',$row[0])."'";
				}		
				if (!empty($dd)){
					echo "addDates: [".implode(",", $dd)."]";
				}	
		?>		
	});	
	
	if (data.length > 0){
		$('#simpliest-usage').multiDatesPicker({
			addDisabledDates: data
		});
	}
	
	
	var curr_val = $('#simpliest-usage').multiDatesPicker('getDates');
	var addDD = "<?php echo implode(",", $dd); ?>";
		
	del_not_exist_date(curr_val, addDD, "", data);	


	
	
});

</script>
<style>

#simpliest-usage .ui-state-disabled {
    BACKGROUND-IMAGE: none; FILTER: none; opacity: .99;
}
.ui-datepicker .ui-datepicker-calendar .ui-state-highlight a {
    background: #743620 none;
    color: white;
}  

</style>
<?php 
/*   End Multiple Date Picker */
?>

<h2><?php echo isset($id) ? ($edit_type == "series" ? get_vocab("editseries") : get_vocab("editentry")) : get_vocab("addentry"); ?></H2>

<FORM NAME="main" ACTION="edit_entry_handler.php" METHOD="GET">

<TABLE BORDER=0>

<TR><TD CLASS=CR><B><?php echo get_vocab("namebooker")?></B></TD>
  <TD CLASS=CL><INPUT NAME="name" SIZE=50 VALUE="<?php echo htmlspecialchars($name,ENT_NOQUOTES) ?>"><?php if (isAdmin() && isset($id)) { ?><input name="isUsedRoom" type="checkbox" <?php echo $is_used?> value="1">已使用該課室<?php } ?></TD></TR>

<TR><TD CLASS=TR><B><?php echo get_vocab("fulldescription")?></B></TD>
  <TD CLASS=TL><TEXTAREA NAME="description" ROWS=8 COLS=40 WRAP="virtual"><?php echo
htmlspecialchars ( $description ); ?></TEXTAREA></TD></TR>

<?php if (isAdmin()) { ?>
<TR><TD CLASS=CR><B><?php echo get_vocab("date"); ?></B></TD>
	<TD CLASS=CL>
	<?php genDateSelector("", $start_day, $start_month, $start_year) ?>
	</TD>
</TR>
<?php } else { ?>
<TR><TD CLASS=CR><B><?php echo get_vocab("date")?></B></TD>
	<TD CLASS=CL>
	<?php echo $start_year."年 ".$start_month."月 ".$start_day."日<br>" ?>
	</TD>
	<input name="day" type="hidden" value="<?php echo $start_day ?>">
	<input name="month" type="hidden" value="<?php echo $start_month ?>">
	<input name="year" type="hidden" value="<?php echo $start_year ?>">
</TR>
<?php } ?>

<?php if(! $enable_periods ) { ?>

<?php if (isAdmin()) { ?>

<TR><TD CLASS=CR><B><?php echo get_vocab("time");?></B></TD>
  <TD CLASS=CL><INPUT NAME="hour" SIZE=2 VALUE="<?php if (!$twentyfourhour_format && ($start_hour > 12)){ echo ($start_hour - 12);} else { echo $start_hour;} ?>" MAXLENGTH=2>:<INPUT NAME="minute" SIZE=2 VALUE="<?php echo $start_min;?>" MAXLENGTH=2>
<?php
if (!$twentyfourhour_format)
{
  $checked = ($start_hour < 12) ? "checked" : "";
  echo "<INPUT NAME=\"ampm\" type=\"radio\" value=\"am\" $checked>".utf8_date("a",mktime(1,0,0,1,1,2000));
  $checked = ($start_hour >= 12) ? "checked" : "";
  echo "<INPUT NAME=\"ampm\" type=\"radio\" value=\"pm\" $checked>".utf8_date("a",mktime(13,0,0,1,1,2000));
}
?>
</TD></TR>

<?php } else { ?>

<TR><TD CLASS=CR><B><?php echo get_vocab("time")?></B></TD>
  <TD CLASS=CL><INPUT NAME="hour" TYPE="hidden" SIZE=2 VALUE="<?php if (!$twentyfourhour_format && ($start_hour > 12)){ echo ($start_hour - 12);} else { echo $start_hour;} ?>" MAXLENGTH=2><?php if (!$twentyfourhour_format && ($start_hour > 12)){ echo ($start_hour - 12);} else { echo $start_hour;} ?>:<INPUT NAME="minute" TYPE="hidden" SIZE=2 VALUE="<?php echo $start_min;?>" MAXLENGTH=2><?php echo $start_min;?>&nbsp;
<?php
if (!$twentyfourhour_format)
{
  if ($start_hour < 12) 
	  echo "<INPUT NAME=\"ampm\" TYPE=\"hidden\" VALUE=\"am\">".AM;
  else
	  echo "<INPUT NAME=\"ampm\" TYPE=\"hidden\" VALUE=\"pm\">".PM;
}
?>
</TD></TR>

<?php } ?>

<?php } else { ?> 

<TR><TD CLASS=CR><B><?php echo get_vocab("period")?></B></TD>
    <TD CLASS=CL>
    <SELECT NAME="period">
<?php
foreach ($periods as $p_num => $p_val)
{
	echo "<OPTION VALUE=$p_num";
	if( ( isset( $period ) && $period == $p_num ) || $p_num == $start_min)
        	echo " SELECTED";
	echo ">$p_val";
}
?>
	    </SELECT>
		</TD></TR>

<?php } ?>

<?php if (isAdmin()) { ?>

<TR><TD CLASS=CR><B><?php echo get_vocab("duration");?></B></TD>
  <TD CLASS=CL><INPUT NAME="duration" SIZE=7 VALUE="<?php echo $duration;?>">
    <SELECT NAME="dur_units">
<?php
if( $enable_periods )
	$units = array("periods", "days");
else
	$units = array("minutes", "hours", "days", "weeks");

while (list(,$unit) = each($units))
{
	echo "<OPTION VALUE=$unit";
	if ($dur_units == get_vocab($unit)) echo " SELECTED";
	echo ">".get_vocab($unit);
}
?>
    </SELECT>
</TD></TR>

<?php } else { ?>

<TR><TD CLASS=CR><B><?php echo get_vocab("duration");?></B></TD>
  <TD CLASS=CL><INPUT NAME="duration" TYPE="hidden" SIZE=7 VALUE="<?php echo $duration;?>"><?php echo $duration;?>
<?php
if( $enable_periods )
	$units = array("periods", "days");
else
	$units = array("minutes", "hours", "days", "weeks");

$savedUnit="";
while (list(,$unit) = each($units))
{
	if ($dur_units == get_vocab($unit)) $savedUnit=$unit;
}
?>
	<INPUT NAME="dur_units" TYPE="hidden" VALUE="<?php echo $savedUnit ?>"><?php echo get_vocab($savedUnit);?>

</TD></TR>

<?php } ?>

<?php
      # Determine the area id of the room in question first
      $sql = "select area_id from $tbl_room where id=$room_id";
      $res = sql_query($sql);
      $row = sql_row($res, 0);
      $area_id = $row[0];
      # determine if there is more than one area
      $sql = "select id from $tbl_area";
      $res = sql_query($sql);
      $num_areas = sql_count($res);
      # if there is more than one area then give the option
      # to choose areas.

	  if (isAdmin()) { 

	  if( $num_areas > 1 ) {

?>

<script language="JavaScript">
<!--
function changeRooms( formObj )
{
    areasObj = eval( "formObj.areas" );

    area = areasObj[areasObj.selectedIndex].value
//    roomsObj = eval( "formObj.elements['rooms']" )
    roomsObj = eval( "formObj.rooms" )
    // remove all entries
//    for (i=0; i < (roomsObj.length); i++)
//    {
//      roomsObj.options[i] = null;
//    }
	roomsObj.length = 0;
    // add entries based on area selected
    switch (area){
<?php
        # get the area id for case statement
	$sql = "select id, area_name from $tbl_area order by area_name";
        $res = sql_query($sql);
	if ($res) for ($i = 0; ($row = sql_row($res, $i)); $i++)
	{

                print "      case \"".$row[0]."\":\n";
        	# get rooms for this area
		$sql2 = "select id, room_name from $tbl_room where area_id='".$row[0]."' order by room_name";
        	$res2 = sql_query($sql2);
		if ($res2) for ($j = 0; ($row2 = sql_row($res2, $j)); $j++)
		{
                	print "        roomsObj.options[$j] = new Option(\"".str_replace('"','\\"',$row2[1])."\",".$row2[0] .")\n";
                }
		# select the first entry by default to ensure
		# that one room is selected to begin with
		print "        roomsObj.options[0].selected = true\n";
		print "        break\n";
	}
?>
    } //switch
}

// create area selector if javascript is enabled as this is required
// if the room selector is to be updated.
this.document.writeln("<tr><td class=CR><b><?php echo get_vocab("areas") ?>:</b></td><td class=CL valign=top>");
this.document.writeln("          <select name=\"areas\" onChange=\"changeRooms(this.form)\">");
<?php
# get list of areas
$sql = "select id, area_name from $tbl_area order by area_name";
$res = sql_query($sql);
if ($res) for ($i = 0; ($row = sql_row($res, $i)); $i++)
{
	$selected = "";
	if ($row[0] == $area_id) {
		$selected = "SELECTED";
	}
	print "this.document.writeln(\"            <option $selected value=\\\"".$row[0]."\\\">".$row[1]."\")\n";
}
?>
this.document.writeln("          </select>");
this.document.writeln("</td></tr>");
// -->
</script>

<?php
} # if $num_areas
?>

<tr><td class=CR><b><?php echo get_vocab("rooms") ?>:</b></td>
  <td class=CL valign=top><table><tr><td><select name="rooms">
  <?php
        # select the rooms in the area determined above
	$sql = "select id, room_name from $tbl_room where area_id=$area_id order by room_name";
   	$res = sql_query($sql);


   	if ($res) for ($i = 0; ($row = sql_row($res, $i)); $i++)
   	{
		$selected = "";
		if ($row[0] == $room_id) {
			$selected = "SELECTED";
		}
		echo "<option $selected value=\"".$row[0]."\">".$row[1];
        // store room names for emails
        $room_names[$i] = $row[1];
   	}
  ?>
  </select></td><td><?php echo "" ?></td></tr></table>
    </td></tr>

<?php } else { ?> 

	<tr><td class=CR><b><?php echo get_vocab("areas") ?>:</b></td>
<?php
# get list of areas
$sql = "select id, area_name from $tbl_area order by area_name";
$res = sql_query($sql);
$areaName = "";
if ($res) for ($i = 0; ($row = sql_row($res, $i)); $i++)
{
	$selected = "";
	if ($row[0] == $area_id) {
		$areaName = $row[1];
	}
}
?>
	<td class=CL valign=top>
	<input name="areas" type="hidden" value="<?php echo $area_id ?>"><?php echo $areaName ?>
	</td></tr>


<tr><td class=CR><b><?php echo get_vocab("rooms") ?>:</b></td>
  <?php
        # select the rooms in the area determined above
	$sql = "select id, room_name from $tbl_room where area_id=$area_id order by room_name";
   	$res = sql_query($sql);

	$roomName = "";
   	if ($res) for ($i = 0; ($row = sql_row($res, $i)); $i++)
   	{
		$selected = "";
		if ($row[0] == $room_id) {
			$roomName = $row[1];
		}
        // store room names for emails
//        $room_names[$i] = $row[1];
   	}
  ?>
  <td class=CL valign=top>
	<input name="rooms" type="hidden" value="<?php echo $room_id ?>"><?php echo $roomName ?>
  </td>
<?php } ?> 


<TR><TD CLASS=CR><B><?php echo get_vocab("type")?></B></TD>
  <TD CLASS=CL>
  
<?php 

if (isAdmin()) {
  echo "<SELECT NAME=\"type\">";

  for ($c = 0; $c <= 20; $c++) {
	if (!empty($cunitname[$c]))
		echo "<OPTION VALUE=$cunit[$c]" . ($type == $cunit[$c] ? " SELECTED" : "") . ">$cunitname[$c]\n";
  }
  echo "</SELECT>";

} else {

  $located=0;
  for ($c = 0; $c <= 20; $c++) {
	if ($type==$cunit[$c])
		$located=$c;
	}
  echo "<INPUT NAME=\"type\" type=\"hidden\" value=\"$cunit[$located]\">";
  echo $cunitname[$located]."<br>";
}
?>

</TD></TR>


<?php if (isAdmin()) { ?>

<?php if($edit_type == "series") { ?>

<TR>
 <TD CLASS=CR><B><?php echo get_vocab("rep_type")?></B></TD>
 <TD CLASS=CL>
<?php

for($i = 0; isset($vocab["rep_type_$i"]); $i++)
{
	echo "<INPUT NAME=\"rep_type\" TYPE=\"RADIO\" VALUE=\"" . $i . "\"";

	if($i == $rep_type)
		echo " CHECKED";

	echo ">" . get_vocab("rep_type_$i") . "\n";
}

?>
 </TD>
</TR>

<TR>
 <TD CLASS=CR><B><?php echo get_vocab("rep_end_date")?></B></TD>
 <TD CLASS=CL><?php genDateSelector("rep_end_", $rep_end_day, $rep_end_month, $rep_end_year) ?></TD>
</TR>

<?php /** 2016-07-06 Add multiple datepicker */ ?>
<TR>
 <TD CLASS=CR><B><?php echo get_vocab("remove_repeat_day")?></B></TD>
 <TD CLASS=CL>
 	<div style='margin-top:10px; float:left;' id="simpliest-usage" class="box"></div>
 	<input style="margin-top:125px; margin-left:10px;  float:left;" type="button"  style="margin-top:5px;"  onclick="clear_all_selected_date()" value="清除所有已選取的日期">
 </TD>
</TR>

<tr>
	<td CLASS=CR><b>需要除去的日期:</b></td>
	<td CLASS=CL>
		<div style="margin-top:10px; float:left;">
		 	<input type="text" name="from_date" id="from_date">(YYYYMMDD)
		 	<?php /*<label>至</label><input type="text" name="from_to" id="to_date">*/?>
	 	</div>
	 	<div style="margin-left:10px;margin-top:2px; float:left;">
		 	<?php
			# Display day name checkboxes according to language and preferred weekday start.
			/*
			for ($i = 0; $i < 7; $i++)
			{
				$wday = ($i + $weekstarts) % 7;
				echo "<INPUT NAME=\"rm_week\" TYPE=CHECKBOX onclick=\"weekday()\">" . day_name($wday) . "\n";
			}
			echo "<INPUT NAME=\"rm_week\" TYPE=CHECKBOX onclick=\"everyday()\">每天";
			*/
			?>
		
		<input style="margin-top:5px;" type="button"  style="margin-top:5px;"  onclick="add_range_date()" value="加入至日曆">
		<input style="margin-top:5px;" type="button"  style="margin-top:5px;"  onclick="remove_range_date()" value="從日曆移除">
		</div>
 	</td>
</tr>
<?php /** End multiple datepicker*/ ?>

<TR>
 <TD CLASS=CR><B><?php echo get_vocab("rep_rep_day")?></B> <?php echo get_vocab("rep_for_weekly")?></TD>
 <TD CLASS=CL>
<?php
# Display day name checkboxes according to language and preferred weekday start.
for ($i = 0; $i < 7; $i++)
{
	$wday = ($i + $weekstarts) % 7;
	echo "<INPUT NAME=\"rep_day[$wday]\" TYPE=CHECKBOX";
	if ($rep_day[$wday]) echo " CHECKED";
	echo ">" . day_name($wday) . "\n";
}
?>
 </TD>
</TR>

<?php
}
else
{
	$key = "rep_type_" . (isset($rep_type) ? $rep_type : "0");

	echo "<tr><td class=\"CR\"><b>".get_vocab("rep_type")."</b></td><td class=\"CL\">".get_vocab($key)."</td></tr>\n";

	if(isset($rep_type) && ($rep_type != 0))
	{
		$opt = "";
		if ($rep_type == 2)
		{
			# Display day names according to language and preferred weekday start.
			for ($i = 0; $i < 7; $i++)
			{
				$wday = ($i + $weekstarts) % 7;
				if ($rep_opt[$wday]) $opt .= day_name($wday) . " ";
			}
		}
		if($opt)
			echo "<tr><td class=\"CR\"><b>".get_vocab("rep_rep_day")."</b></td><td class=\"CL\">$opt</td></tr>\n";

		echo "<tr><td class=\"CR\"><b>".get_vocab("rep_end_date")."</b></td><td class=\"CL\">$rep_end_date</td></tr>\n";
	}
}
/* We display the rep_num_weeks box only if:
   - this is a new entry ($id is not set)
   Xor
   - we are editing an existing repeating entry ($rep_type is set and
     $rep_type != 0 and $edit_type == "series" )
*/
if ( ( !isset( $id ) ) Xor ( isset( $rep_type ) && ( $rep_type != 0 ) && ( "series" == $edit_type ) ) )
{
?>

<TR>
 <TD CLASS=CR><B><?php echo get_vocab("rep_num_weeks")?></B> <?php echo get_vocab("rep_for_nweekly")?></TD>
 <TD CLASS=CL><INPUT TYPE=TEXT NAME="rep_num_weeks" VALUE="<?php echo $rep_num_weeks?>">
</TR>
<?php } ?>

<?php } ?>


<TR>
 <TD colspan=2 align=center>
  <SCRIPT LANGUAGE="JavaScript">
   document.writeln ( '<INPUT TYPE="button" NAME="save_button" VALUE="<?php echo get_vocab("save")?>" ONCLICK="validate_and_submit()">' );
  </SCRIPT>
  <NOSCRIPT>
   <INPUT TYPE="submit" VALUE="<?php echo get_vocab("save")?>">
  </NOSCRIPT>
 </TD></TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="returl"    VALUE="<?php echo $HTTP_REFERER?>">
<!--INPUT TYPE=HIDDEN NAME="room_id"   VALUE="<?php echo $room_id?>"-->
<INPUT TYPE=HIDDEN NAME="create_by" VALUE="<?php echo $create_by?>">
<INPUT TYPE=HIDDEN NAME="rep_id"    VALUE="<?php echo $rep_id?>">
<INPUT TYPE=HIDDEN NAME="edit_type" VALUE="<?php echo $edit_type?>">
<?php /** Add remove_date For Remove Specific Date Function */?>
<INPUT TYPE=HIDDEN NAME="remove_date" VALUE="">

<?php if(isset($id)) echo "<INPUT TYPE=HIDDEN NAME=\"id\"        VALUE=\"$id\">\n";
?>

</FORM>



<script>
$('form[name="main"] .datestr').change(function() { 
	var $cal = $('#simpliest-usage');
	var start_day = get_date("day"); 
	var start_month = get_date("month");
	var start_year = get_date("year");

	var end_day = get_date("rep_end_day");
	var end_month = get_date("rep_end_month");
	var end_year = get_date("rep_end_year");

	var addDD = $cal.multiDatesPicker('getDates');
	$cal.datepicker('destroy');
	
	var d1 = $.datepicker.parseDate('dd/mm/yy', start_day+"/"+start_month+"/"+start_year);
	var d2 = $.datepicker.parseDate('dd/mm/yy', end_day+"/"+end_month+"/"+end_year);
	var data = calcBusinessDays(d1,d2);
	mindd = start_year +""+ start_month +""+ start_day;
	if (data!=''){
		$cal.multiDatesPicker( {
		    minDate: start_month+"/"+start_day+"/"+start_year,
		    maxDate: end_month+"/"+end_day+"/"+end_year,
		    addDates: data,
		    addDisabledDates: data
		});
	}else{
		$cal.multiDatesPicker( {
		    minDate: start_month+"/"+start_day+"/"+start_year,
		    maxDate: end_month+"/"+end_day+"/"+end_year
		});
	}

	var curr_val = $cal.multiDatesPicker('getDates');
	del_not_exist_date(curr_val, addDD, "", data);	
});
</script>

<?php include "trailer.inc" ?>