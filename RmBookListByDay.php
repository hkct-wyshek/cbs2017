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
define("CSVFILE","CampusUsage/"."ClassRmListingByDay_".$fid.".xls");
header("Content-Type: text/html; charset=utf-8");

$fp = fopen(CSVFILE,"w") or die("未能開啟檔案!\n");
flock($fp, LOCK_EX);

# print the page header
print_header($day, $month, $year, $area);

$campus = $_POST["campus"];
$useDate = $_POST["useDate"];
$showPeriod = $_POST["showPeriod"];

foreach($cpname as $key => $val){
	if (isset($_POST["c".($key+1)])){
		if ($_POST["c".($key+1)] == $val['code']){
			$campus = $val['code'];
			$campusName = $val['name'];
		}
	}	
}
/*
if ($_POST["c1"]=="HMTSB") {
	$campus = $_POST["c1"];  }
if ($_POST["c2"]=="MOS") {
	$campus = $_POST["c2"];  }
if ($_POST["c3"]=="MUC") {
	$campus = $_POST["c3"]; }
if ($_POST["c4"]=="JDN") {
	$campus = $_POST["c4"];  }
if ($_POST["c5"]=="OPC") {
	$campus = $_POST["c5"];  }
if ($_POST["c6"]=="PLS") {
	$campus = $_POST["c6"];  }
if ($_POST["c7"]=="TKO") {
	$campus = $_POST["c7"];  }
if ($_POST["c8"]=="HMT") {
	$campus = $_POST["c8"];  }
if ($_POST["c9"]=="YL") {
	$campus = $_POST["c9"];  }
if ($_POST["c10"]=="CSW") {
	$campus = $_POST["c10"];  }
if ($_POST["c11"]=="AUS") {
	$campus = $_POST["c11"];  }
*/	
/*
switch ($campus) {
	// 2012 - 2013
	/*
	case 'HMTSB' : $campusName="何文田南座"; break;
	case 'MFR' : $campusName="文福道校舍"; break;
	case 'MOS' : $campusName="馬鞍山校舍"; break;
	case 'CWB' : $campusName="銅鑼灣校舍"; break;
	*/
	
	// 2015 - 2016 
	/*
	case 'HMTSB' : $campusName = "何文田校舍"; break;
	case 'HMT' : $campusName = "何文田(勞校)"; break;
	case 'MOS' : $campusName = "馬鞍山校舍(MOS)"; break;
	case 'MUC' : $campusName = "馬鞍山本科校園(MUC)"; break;
	case 'JDN' : $campusName = "佐敦培訓中心"; break;
	case 'OPC' : $campusName = "開源道培訓中心"; break;
	case 'PLS' : $campusName = "砵蘭街培訓中心"; break;
	case 'TKO' : $campusName = "將軍澳培訓中心"; break;
	case 'YL' : $campusName = "元朗教學中心"; break;
	case 'CSW' : $campusName = "長沙灣培訓中心"; break;
	case 'AUS' : $campusName = "柯士甸道教學中心"; break;
	
}*/

switch($showPeriod) {
case 'DAY':   
	$starttime = mktime(0,0,0,substr($useDate,5,2),substr($useDate,8,2),substr($useDate,0,4));
	$endtime = mktime(18,59,59,substr($useDate,5,2),substr($useDate,8,2),substr($useDate,0,4));
	break;
case 'EVENING': 
	$starttime = mktime(18,0,0,substr($useDate,5,2),substr($useDate,8,2),substr($useDate,0,4));
	$endtime = mktime(23,59,59,substr($useDate,5,2),substr($useDate,8,2),substr($useDate,0,4));
	break; 
default:      
	$starttime = mktime(0,0,0,substr($useDate,5,2),substr($useDate,8,2),substr($useDate,0,4));
	$endtime = mktime(23,59,59,substr($useDate,5,2),substr($useDate,8,2),substr($useDate,0,4));
	break;
}



?>

<HTML>
  <HEAD>
    <META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
  </HEAD>
  <BODY>

<h2>每日課室表</h2>
<h4><font color=green>數據計算需時，可能要稍等2~3分鐘...</font></h4>

<hr>

<?php

echo "<FONT COLOR=RED><STRONG>".$campusName." (".$useDate.") - ".$showPeriod."</STRONG></FONT><BR><BR>";

$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
	
	// utf8 fixes
	//
mysql_query("SET character_set_client=utf8", $fcn);
mysql_query("SET character_set_connection=utf8", $fcn);
mysql_query("SET character_set_results=utf8", $fcn);
	
mysql_select_db($db_database,$fcn) or die ("Database not found!");

$sql = "select e.start_time, e.end_time, e.room_id, r.room_name, e.name from mrbs_entry e left join mrbs_room r on e.room_id=r.id ";
$sql = $sql."where r.location = \"".$campus."\" ";
$sql = $sql."and ( e.name <> e.create_by ";
$sql = $sql."and not upper(e.name) like \"%CANCEL%\" and not upper(e.name) like \"%VANCAN%\" and not e.name like \"%取消%\" ";
$sql = $sql."and not upper(e.name) like \"%VACA%\" ) ";
$sql = $sql."and e.start_time>=".$starttime." and e.start_time<=".$endtime." order by e.start_time, r.room_name";

$rs = mysql_query($sql, $fcn);
$numRow = mysql_num_rows($rs);

?>


<?php
fwrite($fp, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
fwrite($fp, "<?mso-application progid=\"Excel.Sheet\"?>\n");
fwrite($fp, "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
fwrite($fp, " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
fwrite($fp, " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
fwrite($fp, " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
fwrite($fp, " xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

fwrite($fp, " <Styles>\n");
fwrite($fp, "  <Style ss:ID=\"Default\" ss:Name=\"Normal\">\n");
fwrite($fp, "   <Font ss:FontName=\"Arial\" ss:Size=\"12\"/>\n");
fwrite($fp, "  </Style>\n");
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
fwrite($fp, "     <Font ss:FontName=\"Times New Roman\" x:CharSet=\"136\" x:Family=\"Roman\" ss:Size=\"20\" ss:Color=\"#000000\" ss:Bold=\"1\"/>\n");
fwrite($fp, "     <Interior ss:Color=\"#D8D8D8\" ss:Pattern=\"Solid\"/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, "    <Style ss:ID=\"CenterBoldFontWithBorderLine\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Times New Roman\" x:CharSet=\"136\" x:Family=\"Roman\" ss:Size=\"20\" ss:Color=\"#000000\" ss:Bold=\"1\"/>\n");
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
fwrite($fp, "     <Font ss:FontName=\"Times New Roman\" x:CharSet=\"136\" x:Family=\"Roman\" ss:Size=\"20\" ss:Color=\"#000000\"/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, "    <Style ss:ID=\"CenterFontWithBorderLine\">\n");
fwrite($fp, "     <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\n");
fwrite($fp, "     <Borders>\n");
fwrite($fp, "      <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "      <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n");
fwrite($fp, "     </Borders>\n");
fwrite($fp, "     <Font ss:FontName=\"Times New Roman\" x:CharSet=\"136\" x:Family=\"Roman\" ss:Size=\"20\" ss:Color=\"#000000\"/>\n");
fwrite($fp, "    </Style>\n");

fwrite($fp, " </Styles>\n");

fwrite($fp, " <Worksheet ss:Name=\"".$campusName." (".$useDate.")"."\">\n");
fwrite($fp, "    <Names>\n");
fwrite($fp, "    <NamedRange ss:Name=\"Print_Titles\" ss:RefersTo=\"='".$campusName." (".$useDate.")"."'!R1\"/>\n");
fwrite($fp, "    </Names>\n");

fwrite($fp, "  <Table ss:DefaultRowHeight=\"40\">\n");

fwrite($fp, "   <Column ss:StyleID=\"Left\" ss:AutoFitWidth=\"0\" ss:Width=\"450\"/>\n");
fwrite($fp, "   <Column ss:StyleID=\"Center\" ss:AutoFitWidth=\"0\" ss:Width=\"200\"/>\n");
fwrite($fp, "   <Column ss:StyleID=\"Center\" ss:AutoFitWidth=\"0\" ss:Width=\"150\"/>\n");

fwrite($fp, "   <Row>\n");
fwrite($fp, "    <Cell ss:StyleID=\"BoldFontWithBorderLine\"><Data ss:Type=\"String\">課程名稱</Data><NamedCell
      ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "    <Cell ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">上課時間</Data><NamedCell
      ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "    <Cell ss:StyleID=\"CenterBoldFontWithBorderLine\"><Data ss:Type=\"String\">上課地點</Data><NamedCell
      ss:Name=\"Print_Titles\"/></Cell>\n");
fwrite($fp, "   </Row>\n");

$nowtime = time()+8*60*60;
$strTimeStamp = date("Y",$nowtime)."-".date("m",$nowtime)."-".date("d",$nowtime)."T".date("H",$nowtime).":".date("i",$nowtime).":".date("s",$nowtime);

while ($row = mysql_fetch_array($rs)) { 
	$roomname = $row["room_name"];
	$lessonstart = $row["start_time"]; 
	$lessonend = $row["end_time"]; 
	$lessonname = $row["name"];
    $ampm1 = utf8_date("a",$lessonstart);
    $ampm2 = utf8_date("a",$lessonend);

	echo "[".$lessonname."] [".strftime("%I:%M $ampm1",$lessonstart)." - ".strftime("%I:%M $ampm2",$lessonend)."] [".$roomname."]<br>";

	fwrite($fp, "   <Row>\n");
	fwrite($fp, "    <Cell ss:StyleID=\"BorderLine\"><Data ss:Type=\"String\">".$lessonname."</Data></Cell>\n");
	fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"><Data ss:Type=\"String\">".strftime("%I:%M $ampm1",$lessonstart)." - ".strftime("%I:%M $ampm2",$lessonend)."</Data></Cell>\n");
	fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"><Data ss:Type=\"String\">".$roomname."</Data></Cell>\n");
	fwrite($fp, "   </Row>\n");
	
}

if (mysql_num_rows($rs)<12) {
	for ($i=1;$i<=(12-mysql_num_rows($rs));$i++)	{
		fwrite($fp, "   <Row>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"BorderLine\"></Cell>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
		fwrite($fp, "   </Row>\n");
	}
} else {
	if ((mysql_num_rows($rs)>12) and (mysql_num_rows($rs)<24)) {
		for ($i=1;$i<=(24-mysql_num_rows($rs));$i++)	{
			fwrite($fp, "   <Row>\n");
			fwrite($fp, "    <Cell ss:StyleID=\"BorderLine\"></Cell>\n");
			fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
			fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
			fwrite($fp, "   </Row>\n");
		}
	} else {
		if ((mysql_num_rows($rs)>24) and (mysql_num_rows($rs)<36)) {
			for ($i=1;$i<=(36-mysql_num_rows($rs));$i++)	{
				fwrite($fp, "   <Row>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"BorderLine\"></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
				fwrite($fp, "   </Row>\n");
			}
		} else {
			if ((mysql_num_rows($rs)>36) and (mysql_num_rows($rs)<48)) {
				for ($i=1;$i<=(48-mysql_num_rows($rs));$i++)	{
					fwrite($fp, "   <Row>\n");
					fwrite($fp, "    <Cell ss:StyleID=\"BorderLine\"></Cell>\n");
					fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
					fwrite($fp, "    <Cell ss:StyleID=\"CenterFontWithBorderLine\"></Cell>\n");
					fwrite($fp, "   </Row>\n");
				}		
			}
		}

	}

}


fwrite($fp, "  </Table>\n");
fwrite($fp, "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n");
fwrite($fp, "   <PageSetup>\n");
fwrite($fp, "    <Layout x:Orientation=\"Landscape\"/>\n");
fwrite($fp, "    <Header x:Margin=\"0.19685039370078741\"/>\n");
fwrite($fp, "    <Footer x:Margin=\"0.39370078740157483\" x:Data=\"&amp;LLast Updated: ".$strTimeStamp."&amp;R&amp;A\"/>\n");
fwrite($fp, "    <PageMargins x:Bottom=\"0.39370078740157483\" x:Left=\"0.39370078740157483\" x:Right=\"0.39370078740157483\" x:Top=\"0.39370078740157483\"/>\n");
fwrite($fp, "   </PageSetup>\n");
fwrite($fp, "   <Unsynced/>\n");
fwrite($fp, "   <FitToPage/>\n");
fwrite($fp, "   <Print>\n");
fwrite($fp, "    <FitHeight>99</FitHeight>\n");
fwrite($fp, "    <ValidPrinterInfo/>\n");
fwrite($fp, "    <PaperSizeIndex>9</PaperSizeIndex>\n");
fwrite($fp, "    <VerticalResolution>0</VerticalResolution>\n");
fwrite($fp, "   </Print>\n");

fwrite($fp, "   <Zoom>100</Zoom>\n");
fwrite($fp, "  </WorksheetOptions>\n");
fwrite($fp, " </Worksheet>\n");
fwrite($fp, "</Workbook>\n");

flock($fp, LOCK_UN);
fclose($fp);

$url = "CampusUsage/ClassRmListingByDay_".$fid.".xls";

# marco mark @2013-7-30
#mysql_free_result($frs);
mysql_free_result($rs);

mysql_close($fcn);

?>

  <br>
  <A href="<?php echo $url ?>">如要下載檔案，請在此處用滑鼠右擊並另存目標，下載時可以把檔案名稱更改。</A><br>
  <br>
  <A href="javascript:history.go(-1)">&nbsp;返回上一頁</A>

  </BODY>
</HTML>