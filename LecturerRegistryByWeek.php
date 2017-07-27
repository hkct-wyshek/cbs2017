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


$fid = md5(rand());
define("CSVFILE","CampusUsage/"."LecturerRegistryByWeek_".$fid.".xls");
header("Content-Type: text/html; charset=utf-8");

$fp = fopen(CSVFILE,"w") or die("未能開啟檔案!\n");
flock($fp, LOCK_EX);

# print the page header
print_header($day, $month, $year, $area);

$campus = $_POST["campus"];
$monDate = $_POST["monDate"];

$starttime = mktime(9,0,0,substr($monDate,5,2),substr($monDate,8,2),substr($monDate,0,4));
#$endtime = mktime(22,0,0,substr($monDate,5,2),substr($monDate,8,2),substr($monDate,0,4));
$endtime = $starttime + 6*24*60*60 +13*60*60;

$sunDate = date("Y",$endtime)."-".date("m",$endtime)."-".date("d",$endtime);

$fcn = mysql_connect("localhost","dfm","niloDFM2007") or die ("DB Connection Error.");

// utf8 fixes
//
mysql_query("SET character_set_client=utf8", $fcn);
mysql_query("SET character_set_connection=utf8", $fcn);
mysql_query("SET character_set_results=utf8", $fcn);

mysql_select_db("classrm2011",$fcn) or die ("Database not found!");

# $fsql = "select id,room_name from mrbs_room where location like \"".$campus."%\" and sequence <> 0 order by sequence asc";

$fsql = "select id,room_name from mrbs_room where location like \"".$campus."%\" and sequence >= 0 order by location,sequence,area_id,room_name asc";

$frs = mysql_query($fsql, $fcn);
$fnumRow = mysql_num_rows($frs);

?>

<HTML>
  <HEAD>
    <META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
  </HEAD>
  <BODY>

<h2>Lecturer's Attendance Registry By Week</h2>
<h4><font color=green>數據計算需時，可能要稍等2~3分鐘...</font></h4>

<hr>

<?php

echo "<h4><FONT COLOR=RED><STRONG>完成製作 - ".$campus." (".$monDate." MON - ".$sunDate." SUN)"."</STRONG></FONT></h4>";

fwrite($fp, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
fwrite($fp, "<?mso-application progid=\"Excel.Sheet\"?>\n");
fwrite($fp, "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
fwrite($fp, " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
fwrite($fp, " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
fwrite($fp, " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
fwrite($fp, " xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

fwrite($fp, " <Styles>\n");

fwrite($fp, "  <Style ss:ID=\"Default\">\n");
fwrite($fp, "   <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"12\"/>\n");
fwrite($fp, "  </Style>\n");

fwrite($fp, "  <Style ss:ID=\"HeaderTitle\">\n");
fwrite($fp, "  <Alignment ss:Horizontal=\"CenterAcrossSelection\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"22\" ss:Bold=\"1\"/>\n");
fwrite($fp, "  </Style>\n");

fwrite($fp, "  <Style ss:ID=\"HeaderTitle_RED\">\n");
fwrite($fp, "  <Alignment ss:Horizontal=\"CenterAcrossSelection\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"16\" ss:Color=\"#FF0000\" ss:Bold=\"1\"/>\n");
fwrite($fp, "  </Style>\n");

fwrite($fp, "  <Style ss:ID=\"CenterBold\">\n");
fwrite($fp, "  <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"12\" ss:Bold=\"1\"/>\n");
fwrite($fp, "  </Style>\n");

fwrite($fp, "  <Style ss:ID=\"SignBox\">\n");
fwrite($fp, "  <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "  <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "  </Borders>\n");
fwrite($fp, "  <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"10\" ss:Color=\"#FF0000\"/>\n");
fwrite($fp, "  <Interior/>\n");
fwrite($fp, "  </Style>\n");

fwrite($fp, "  <Style ss:ID=\"RoomBox\">\n");
fwrite($fp, "  <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "  <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "  </Borders>\n");
fwrite($fp, "  <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"10\" ss:Color=\"#000000\" ss:Bold=\"1\"/>\n");
fwrite($fp, "  <Interior/>\n");
fwrite($fp, "  </Style>\n");

fwrite($fp, "    <Style ss:ID=\"ClassBox\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Color=\"#000000\"/>\n");
fwrite($fp, "     <Interior/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, "    <Style ss:ID=\"CenterBoldFontWithBorderLine\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"12\" ss:Color=\"#000000\" ss:Bold=\"1\"/>\n");
fwrite($fp, "     <Interior/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, "    <Style ss:ID=\"CenterBoldFont_14_WithBorderLine\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"14\" ss:Color=\"#000000\" ss:Bold=\"1\"/>\n");
fwrite($fp, "     <Interior/>\n");
fwrite($fp, "    </Style>\n");



fwrite($fp, "  <Style ss:ID=\"Center\">\n");
fwrite($fp, "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, "  <Style ss:ID=\"Right\">\n");
fwrite($fp, "   <Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, "  <Style ss:ID=\"Left\">\n");
fwrite($fp, "   <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  </Style>\n");

fwrite($fp, "    <Style ss:ID=\"BoldFontWithBorderLine\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"12\" ss:Color=\"#000000\" ss:Bold=\"1\"/>\n");
fwrite($fp, "     <Interior ss:Color=\"#D8D8D8\" ss:Pattern=\"Solid\"/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, "    <Style ss:ID=\"BorderLine\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"12\" ss:Color=\"#000000\"/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, "    <Style ss:ID=\"CenterFontWithBorderLine\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Arial Unicode MS\" x:CharSet=\"136\" x:Family=\"Swiss\" ss:Size=\"12\" ss:Color=\"#000000\"/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, " </Styles>\n");

fwrite($fp, " <Worksheet ss:Name=\"".$campus."(".$monDate." to ".$sunDate.")"."\">\n");

fwrite($fp, "  <Names>\n");
fwrite($fp, "   <NamedRange ss:Name=\"Print_Titles\" ss:RefersTo=\"='".$campus."(".$monDate." to ".$sunDate.")"."'!R1:R6\"/>\n");
fwrite($fp, "  </Names>\n");
fwrite($fp, "  <Table ss:StyleID=\"Default\" ss:DefaultColumnWidth=\"54\" ss:DefaultRowHeight=\"17.25\">\n");

fwrite($fp, "  <Column ss:Index=\"2\" ss:StyleID=\"Default\" ss:AutoFitWidth=\"0\" ss:Width=\"63.75\" ss:Span=\"31\"/>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"39\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><Data ss:Type=\"String\">Lecturer's Attendance Register</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"27.75\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle_RED\"><Data ss:Type=\"String\">For classes conducted at ".$campus."</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"21.75\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle_RED\"><Data ss:Type=\"String\">2007-2008 "."(".$monDate." to ".$sunDate.")"."</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"15.75\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle_RED\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"HeaderTitle\"><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"20.25\" ss:StyleID=\"CenterBold\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">Date</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
$x = $starttime;
$strDate = date("d",$x)."-".date("M",$x)."-".date("Y",$x)." (".date("D",$x).")";
fwrite($fp, "  <Cell ss:MergeAcross=\"4\" ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">".$strDate."</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
$x = $x + 24*60*60;
$strDate = date("d",$x)."-".date("M",$x)."-".date("Y",$x)." (".date("D",$x).")";
fwrite($fp, "  <Cell ss:MergeAcross=\"4\" ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">".$strDate."</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
$x = $x + 24*60*60;
$strDate = date("d",$x)."-".date("M",$x)."-".date("Y",$x)." (".date("D",$x).")";
fwrite($fp, "  <Cell ss:MergeAcross=\"4\" ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">".$strDate."</Data><NamedCell
ss:Name=\"Print_Titles\"/></Cell>\n");
$x = $x + 24*60*60;
$strDate = date("d",$x)."-".date("M",$x)."-".date("Y",$x)." (".date("D",$x).")";
fwrite($fp, "  <Cell ss:MergeAcross=\"4\" ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">".$strDate."</Data><NamedCell
ss:Name=\"Print_Titles\"/></Cell>\n");
$x = $x + 24*60*60;
$strDate = date("d",$x)."-".date("M",$x)."-".date("Y",$x)." (".date("D",$x).")";
fwrite($fp, "  <Cell ss:MergeAcross=\"4\" ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">".$strDate."</Data><NamedCell
ss:Name=\"Print_Titles\"/></Cell>\n");
$x = $x + 24*60*60;
$strDate = date("d",$x)."-".date("M",$x)."-".date("Y",$x)." (".date("D",$x).")";
fwrite($fp, "  <Cell ss:MergeAcross=\"2\" ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">".$strDate."</Data><NamedCell
ss:Name=\"Print_Titles\"/></Cell>\n");
$x = $x + 24*60*60;
$strDate = date("d",$x)."-".date("M",$x)."-".date("Y",$x)." (".date("D",$x).")";
fwrite($fp, "  <Cell ss:MergeAcross=\"3\" ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">".$strDate."</Data><NamedCell
ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"22.5\" ss:StyleID=\"CenterBold\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">課室</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">晚</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">晚</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">晚</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">晚</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">晚</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">晚</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">早</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"CenterBoldFont_14_WithBorderLine\"><Data ss:Type=\"String\">午</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "  </Row>\n");

$nowtime = time()+8*60*60;
$strTimeStamp = date("Y",$nowtime)."-".date("m",$nowtime)."-".date("d",$nowtime)."T".date("H",$nowtime).":".date("i",$nowtime).":".date("s",$nowtime);

while ($frow = mysql_fetch_array($frs)) { 
	$roomid = $frow["id"]; 
	$roomname = $frow["room_name"];
#	echo "<strong>".$roomname."</strong><br>";

	fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"69.9375\">\n");
	fwrite($fp, "  <Cell ss:StyleID=\"RoomBox\"><Data ss:Type=\"String\">".$roomname."</Data></Cell>\n");

	$sql = "select start_time, end_time, room_id, name from mrbs_entry where room_id=".$roomid;
	$sql = $sql." and name<>create_by";
	$sql = $sql." and not upper(name) like \"%CANCEL%\" and not upper(name) like \"%VACAN%\" and not name like \"%取消%\"";
	$sql = $sql." and start_time>=".$starttime." and start_time<=".$endtime." order by start_time";
	$rs = mysql_query($sql, $fcn);

	$r=0;
	while ($row = mysql_fetch_array($rs)) { 
		$lessonstart = $row["start_time"]; 
		$lessonend = $row["end_time"]; 
		$lessonname = $row["name"];
#		echo strftime("%m-%d [%H:%M]",$lessonstart)."-".strftime("[%H:%M]",$lessonend)." ".$lessonname."<br>";
		$rmbook[$r]["lessonstart"] = $lessonstart;
		$rmbook[$r]["lessonend"] = $lessonend;
		$rmbook[$r]["lessonname"] = $lessonname;
		$r++;
	}

	for ($c=0;$c<5;$c++) {

		$t1 = $starttime + $c*24*60*60;

		$numSuit=0;
		$loc = array();
		for ($i=0;$i<$r;$i++) {
			if ( $rmbook[$i]["lessonstart"]>=($t1) and $rmbook[$i]["lessonend"]<=($t1+4*60*60) ) {
				$numSuit++;
				$loc[$numSuit] = $i;
			}
		}
		if ($numSuit==2) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[2]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==1) {
			if ( ($rmbook[$loc[1]]["lessonend"]-$rmbook[$loc[1]]["lessonstart"]) > 2*60*60 ) {
				fwrite($fp, "  <Cell ss:MergeAcross=\"1\" ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			} else {
				if ( $rmbook[$loc[1]]["lessonstart"]>=($t1) and $rmbook[$loc[1]]["lessonend"]<=($t1+2*60*60) ) {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
				} else {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
				}
			}
		}
		elseif ($numSuit==0) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
		}

		$numSuit=0;
		$loc = array();
		for ($i=0;$i<$r;$i++) {
			if ( $rmbook[$i]["lessonstart"]>=($t1+5*60*60) and $rmbook[$i]["lessonend"]<=($t1+9*60*60) ) {
				$numSuit++;
				$loc[$numSuit] = $i;
			}
		}
		if ($numSuit==2) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[2]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==1) {
			if ( ($rmbook[$loc[1]]["lessonend"]-$rmbook[$loc[1]]["lessonstart"])> 2*60*60 ) {
				fwrite($fp, "  <Cell ss:MergeAcross=\"1\" ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			} else {
				if ( $rmbook[$loc[1]]["lessonstart"]>=($t1+5*60*60) and $rmbook[$loc[1]]["lessonend"]<=($t1+7*60*60) ) {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
				} else {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
				}
			}
		}
		elseif ($numSuit==0) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
		}

		$numSuit=0;
		$loc = array();
		for ($i=0;$i<$r;$i++) {
			if ( $rmbook[$i]["lessonstart"]>=($t1+10*60*60) and $rmbook[$i]["lessonend"]<=($t1+13*60*60) ) {
				$numSuit++;
				$loc[$numSuit] = $i;
			}
		}
		if ($numSuit==2) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]." +	".$rmbook[$loc[2]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==1) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==0) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
		}

	}

	$t1 = $starttime + 5*24*60*60;

	$numSuit=0;
	$loc = array();
	$longLesson = false;
	for ($i=0;$i<$r;$i++) {
		if ( $rmbook[$i]["lessonstart"]>=($t1+5*60*60) and $rmbook[$i]["lessonend"]<=($t1+13*60*60) ) {
			$numSuit++;
			$loc[$numSuit] = $i;
		}
	}
	if ($numSuit==1) {
		if ( ($rmbook[$loc[1]]["lessonend"]-$rmbook[$loc[1]]["lessonstart"]) > 6*60*60 ) {
			fwrite($fp, "  <Cell ss:MergeAcross=\"2\" ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			$longLesson = TRUE;
		}
	}

	if (! $longLesson) {

		$numSuit=0;
		$loc = array();
		for ($i=0;$i<$r;$i++) {
			if ( $rmbook[$i]["lessonstart"]>=($t1+5*60*60) and $rmbook[$i]["lessonend"]<=($t1+9*60*60) ) {
				$numSuit++;
				$loc[$numSuit] = $i;
			}
		}
		if ($numSuit==2) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[2]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==1) {
			if ( ($rmbook[$loc[1]]["lessonend"]-$rmbook[$loc[1]]["lessonstart"])> 2*60*60 ) {
				fwrite($fp, "  <Cell ss:MergeAcross=\"1\" ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			} else {
				if ( $rmbook[$loc[1]]["lessonstart"]>=($t1+5*60*60) and $rmbook[$loc[1]]["lessonend"]<=($t1+7*60*60) ) {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
				} else {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
				}
			}
		}
		elseif ($numSuit==0) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
		}

		$numSuit=0;
		$loc = array();
		for ($i=0;$i<$r;$i++) {
			if ( $rmbook[$i]["lessonstart"]>=($t1+9*60*60) and $rmbook[$i]["lessonend"]<=($t1+13*60*60) ) {
				$numSuit++;
				$loc[$numSuit] = $i;
			}
		}
		if ($numSuit==2) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]." +	".$rmbook[$loc[2]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==1) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==0) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
		}

	}

	$t1 = $starttime + 6*24*60*60;

	$numSuit=0;
	$loc = array();
	$longLesson = false;
	for ($i=0;$i<$r;$i++) {
		if ( $rmbook[$i]["lessonstart"]>=($t1) and $rmbook[$i]["lessonend"]<=($t1+9*60*60) ) {
			$numSuit++;
			$loc[$numSuit] = $i;
		}
	}
	if ($numSuit==1) {
		if ( ($rmbook[$loc[1]]["lessonend"]-$rmbook[$loc[1]]["lessonstart"]) > 6*60*60 ) {
			fwrite($fp, "  <Cell ss:MergeAcross=\"3\" ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			$longLesson = TRUE;
		}
	}

	if (! $longLesson) {

		$numSuit=0;
		$loc = array();
		for ($i=0;$i<$r;$i++) {
			if ( $rmbook[$i]["lessonstart"]>=($t1) and $rmbook[$i]["lessonend"]<=($t1+4*60*60) ) {
				$numSuit++;
				$loc[$numSuit] = $i;
			}
		}
		if ($numSuit==2) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[2]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==1) {
			if ( ($rmbook[$loc[1]]["lessonend"]-$rmbook[$loc[1]]["lessonstart"]) > 2*60*60 ) {
				fwrite($fp, "  <Cell ss:MergeAcross=\"1\" ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			} else {
				if ( $rmbook[$loc[1]]["lessonstart"]>=($t1) and $rmbook[$loc[1]]["lessonend"]<=($t1+2*60*60) ) {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
				} else {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
				}
			}
		}
		elseif ($numSuit==0) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
		}

		$numSuit=0;
		$loc = array();
		for ($i=0;$i<$r;$i++) {
			if ( $rmbook[$i]["lessonstart"]>=($t1+5*60*60) and $rmbook[$i]["lessonend"]<=($t1+9*60*60) ) {
				$numSuit++;
				$loc[$numSuit] = $i;
			}
		}
		if ($numSuit==2) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[2]]["lessonname"]."</Data></Cell>\n");
		}
		elseif ($numSuit==1) {
			if ( ($rmbook[$loc[1]]["lessonend"]-$rmbook[$loc[1]]["lessonstart"])> 2*60*60 ) {
				fwrite($fp, "  <Cell ss:MergeAcross=\"1\" ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
			} else {
				if ( $rmbook[$loc[1]]["lessonstart"]>=($t1+5*60*60) and $rmbook[$loc[1]]["lessonend"]<=($t1+7*60*60) ) {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
				} else {
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
					fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\">".$rmbook[$loc[1]]["lessonname"]."</Data></Cell>\n");
				}
			}
		}
		elseif ($numSuit==0) {
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
			fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"><Data ss:Type=\"String\"></Data></Cell>\n");
		}

	}

	fwrite($fp, "  </Row>\n");




	fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"69.9375\">\n");
	fwrite($fp, "  <Cell ss:StyleID=\"SignBox\"><Data ss:Type=\"String\">Lecturers' Sign-in</Data></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
	fwrite($fp, "  </Row>\n");

}

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"69.9375\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"RoomBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"69.9375\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"SignBox\"><Data ss:Type=\"String\">Lecturers' Sign-in</Data></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"69.9375\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"RoomBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  <Row ss:AutoFitHeight=\"0\" ss:Height=\"69.9375\">\n");
fwrite($fp, "  <Cell ss:StyleID=\"SignBox\"><Data ss:Type=\"String\">Lecturers' Sign-in</Data></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  <Cell ss:StyleID=\"ClassBox\"></Cell>\n");
fwrite($fp, "  </Row>\n");

fwrite($fp, "  </Table>\n");

fwrite($fp, "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n");
fwrite($fp, "  <PageSetup>\n");
fwrite($fp, "  <Layout x:Orientation=\"Landscape\" x:CenterHorizontal=\"1\" x:CenterVertical=\"0\"/>\n");
fwrite($fp, "  <Header x:Margin=\"0.31496062992125984\"/>\n");
fwrite($fp, "  <Footer x:Margin=\"0.31496062992125984\" x:Data=\"&amp;LLast Updated: ".$strTimeStamp."&amp;R&amp;P/&amp;N\"/>\n");
fwrite($fp, "  <PageMargins x:Bottom=\"0.51181102362204722\" x:Left=\"0.23622047244094491\" x:Right=\"0\" x:Top=\"0.62992125984251968\"/>\n");
fwrite($fp, "  </PageSetup>\n");
fwrite($fp, "  <Unsynced/>\n");
fwrite($fp, "  <FitToPage/>\n");
fwrite($fp, "  <Print>\n");
fwrite($fp, "  <FitHeight>99</FitHeight>\n");
fwrite($fp, "  <ValidPrinterInfo/>\n");
fwrite($fp, "  <PaperSizeIndex>9</PaperSizeIndex>\n");
fwrite($fp, "  <Scale>38</Scale>\n");
fwrite($fp, "  <VerticalResolution>0</VerticalResolution>\n");
fwrite($fp, "  </Print>\n");
fwrite($fp, "  <Zoom>80</Zoom>\n");
fwrite($fp, "  <Selected/>\n");
fwrite($fp, "  <FreezePanes/>\n");
fwrite($fp, "  <FrozenNoSplit/>\n");
fwrite($fp, "  <SplitHorizontal>6</SplitHorizontal>\n");
fwrite($fp, "  <TopRowBottomPane>6</TopRowBottomPane>\n");
fwrite($fp, "  <SplitVertical>1</SplitVertical>\n");
fwrite($fp, "  <LeftColumnRightPane>1</LeftColumnRightPane>\n");
fwrite($fp, "  <ActivePane>0</ActivePane>\n");
fwrite($fp, "  <Panes>\n");
fwrite($fp, "  <Pane>\n");
fwrite($fp, "  <Number>3</Number>\n");
fwrite($fp, "  </Pane>\n");
fwrite($fp, "  <Pane>\n");
fwrite($fp, "  <Number>1</Number>\n");
fwrite($fp, "  </Pane>\n");
fwrite($fp, "  <Pane>\n");
fwrite($fp, "  <Number>2</Number>\n");
fwrite($fp, "  </Pane>\n");
fwrite($fp, "  <Pane>\n");
fwrite($fp, "  <Number>0</Number>\n");
fwrite($fp, "  <ActiveRow>0</ActiveRow>\n");
fwrite($fp, "  <ActiveCol>0</ActiveCol>\n");
fwrite($fp, "  </Pane>\n");
fwrite($fp, "  </Panes>\n");
fwrite($fp, "  <ProtectObjects>False</ProtectObjects>\n");
fwrite($fp, "  <ProtectScenarios>False</ProtectScenarios>\n");
fwrite($fp, "  </WorksheetOptions>\n");

fwrite($fp, " </Worksheet>\n");
fwrite($fp, "</Workbook>\n");

flock($fp, LOCK_UN);
fclose($fp);

$url = "CampusUsage/LecturerRegistryByWeek_".$fid.".xls";

mysql_free_result($frs);
mysql_close($fcn);

?>

  <br>
  <A href="<?php echo $url ?>">如要下載檔案，請在此處用滑鼠右擊並另存目標，下載時可以把檔案名稱更改。</A><br>
  <br>
  <A href="javascript:history.go(-1)">&nbsp;返回上一頁</A>

  </BODY>
</HTML>