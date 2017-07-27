<?php
# $Id: trailer.inc,v 1.10 2004/01/15 21:01:10 jflarvoire Exp $

if ( $pview != 1 ) {

echo "<P><HR><B>".get_vocab("viewday").":</B>\n";

$currentDay   = date("d");
$currentMonth = date("m");
$currentYear  = date("Y");

if (empty($area))
	$params = "";
else
	$params = "&area=$area";

for($i = 0; $i <= 13; $i++)
{
	$ctime = mktime(0, 0, 0, $currentMonth, $currentDay + $i, $currentYear);

	$str = utf8_strftime(empty($dateformat)? "%b %d" : "%d %b", $ctime);

	$cyear  = date("Y", $ctime);
	$cmonth = date("m", $ctime);
	$cday   = date("d", $ctime);
	if ($i != -6) echo " | ";
	echo "<a href=\"day.php?year=$cyear&month=$cmonth&day=$cday$params\">$str</a>\n";
}


}
?>

<HR>
</BODY>
</HTML>