<?php
# $Id: report.php,v 1.22 2004/04/17 15:28:37 thierry_bo Exp $
 
require_once "grab_globals.inc.php";
include "config.inc.php";
include "functions.inc";
include "$dbsys.inc";
include "mrbs_auth.inc";
include "ReportFunctions.inc";

ini_set('max_execution_time', 0);

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

$dept = substr(strtoupper($user),0,strlen($user)-4);
if ($dept == "AEC-YJDP"){
	$dept = "AEC-YJDPT";
}elseif($dept == "DTH-THP"){
	$dept = "DTH";
}elseif($dept == "DDTBT-TP"){
	$dept = "DDTB-TP";
}elseif($dept == "DTH-THP"){
	$dept = "DTH";
}
$fid = md5(rand());
define("CSVFILE","CampusUsage/"."DeptUsage(".$dept.")_".$fid.".xls");
header("Content-Type: text/html; charset=utf-8");

$fp = fopen(CSVFILE,"w") or die("未能開啟檔案!\n");
flock($fp, LOCK_EX);

# print the page header
print_header($day, $month, $year, $area);
?>

<?php
$from = $_POST["dfrom"];
$to = $_POST["dto"];
$mode = $_POST["mode"];

if ($mode=="day") { $modeName = "全日制時段"; }
if ($mode=="night") { $modeName = "兼讀制時段"; }
if ($mode=="all") { $modeName = "全日制 + 兼讀制時段"; }

$xcunit = $cunit;
$campusName = "";

foreach($cpname as $key => $val){
	if (isset($_POST["c".($key+1)])){
		if ($_POST["c".($key+1)] == $val['code']){
			$campusName = $campusName."[".$val['name']."] ";
		}
	}	
}
	

	
?>

<h2>實際及預算校舍使用率 - 搜尋結果(<?php echo $dept ?>)</h2>
<h4><font color=green>數據計算需時，可能要稍等2~3分鐘...</font></h4>
<h4><?php echo "[".$modeName."] ".$from."至".$to." ".$campusName ?></h4>

〔全日制每個課室每週節數供應: 10節 (星期一至五09:00-18:00, 4小時計一節, 2小時計半節)〕<br>
〔兼讀制每個課室每週節數供應: 7節 (星期一至五每晚一節, 星期六下午一節, 星期六晚上一節)〕<br>
〔全年供應計算已減去的日子 : 星期日 ， 公眾假期 ， 校慶/退修 ， 中秋/冬至晚上〕<br>
〔何文田勞校課室只提供晚上節數給本校課程使用〕<br>
〔113中心只提供晚上節數給非SVT課程使用〕<br>

<hr>

<table>

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
fwrite($fp, "  <Style ss:ID=\"SubTotal\">\n");
fwrite($fp, "   <Interior ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, "  <Style ss:ID=\"Percent\">\n");
fwrite($fp, "   <NumberFormat ss:Format=\"0%\"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, "  <Style ss:ID=\"2dp\">\n");
fwrite($fp, "   <NumberFormat ss:Format=\"0.00_ \"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, "  <Style ss:ID=\"Centre\">\n");
fwrite($fp, "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, "  <Style ss:ID=\"Right\">\n");
fwrite($fp, "   <Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, "  <Style ss:ID=\"Left\">\n");
fwrite($fp, "   <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Center\"/>\n");
fwrite($fp, "  </Style>\n");
fwrite($fp, " </Styles>\n");

fwrite($fp, " <Worksheet ss:Name=\"".$modeName."\">\n");
fwrite($fp, "  <Table ss:DefaultRowHeight=\"16.5\">\n");
fwrite($fp, "   <Column ss:Width=\"150\"/>\n");
fwrite($fp, "   <Column ss:Width=\"50\"/>\n");
fwrite($fp, "   <Column ss:Width=\"100\"/>\n");
fwrite($fp, "   <Column ss:Width=\"100\"/>\n");
fwrite($fp, "   <Column ss:Width=\"100\"/>\n");

fwrite($fp, "   <Row>\n");
fwrite($fp, "    <Cell><Data ss:Type=\"String\">實際及預算校舍使用率(".$dept.")</Data></Cell>\n");
fwrite($fp, "   </Row>\n");

fwrite($fp, "   <Row/>\n");

fwrite($fp, "   <Row>\n");
fwrite($fp, "    <Cell><Data ss:Type=\"String\">檢視期間 :</Data></Cell>\n");
fwrite($fp, "    <Cell><Data ss:Type=\"String\">".$from." 至 ".$to."</Data></Cell>\n");
fwrite($fp, "   </Row>\n");

fwrite($fp, "   <Row>\n");
fwrite($fp, "    <Cell><Data ss:Type=\"String\">檢視模式 :</Data></Cell>\n");
fwrite($fp, "    <Cell><Data ss:Type=\"String\">".$modeName."</Data></Cell>\n");
fwrite($fp, "   </Row>\n");

fwrite($fp, "   <Row>\n");
fwrite($fp, "    <Cell><Data ss:Type=\"String\">地點 :</Data></Cell>\n");
fwrite($fp, "    <Cell><Data ss:Type=\"String\">".$campusName."</Data></Cell>\n");
fwrite($fp, "   </Row>\n");

fwrite($fp, "   <Row/>\n");


$special = false;
$supply = CalcSupply($from,$to,$mode,$special);

function set_conn(){
	
	global $db_host;
	global $db_login;
	global $db_password;
	global $db_database;
	global $tbl_room;
	
	$fcn = mysql_connect($db_host,$db_login,$db_password) or die ("DB Connection Error.");
		
	mysql_query("SET character_set_client=utf8", $fcn);
	mysql_query("SET character_set_connection=utf8", $fcn);
	mysql_query("SET character_set_results=utf8", $fcn);
		
	mysql_select_db($db_database,$fcn) or die ("Database not found!");
	
	return $fcn;
}

function get_area_list($id){
	global $tbl_room;
	
	$fcn = set_conn();			
	$sql = "SELECT id from $tbl_room WHERE area_id=$id";
	$res = mysql_query($sql, $fcn);
	
	$tmp = array();
	$tmp[0] = $id;
	$i = 1;
	while($r = mysql_fetch_array($res)){ 
		$tmp[$i] = $r['id'];
		$i++;
	}
	
	return $tmp;
}

foreach ($cpname as $key => $val){
	$num = $key + 1;
	if ($_POST["c$num"] == $val['code']){
		
		$fcn = set_conn();
		$query = "select id from $tbl_area where cpname='".$val['name']."'"; 
		$result = mysql_query($query, $fcn);

		$i = 1; 
		while($row = mysql_fetch_array($result)){ 
			$campus[$num][$i][] = get_area_list($row['id']);	
			$fp = fwriteData2($fp, $campus[$num][$i], $supply, $dept, $from, $to, $mode, $special);
			$i++; 
		}
	}
}


/*
if ($_POST["c1"]=="HMTSB") {
	$campus[1][1][] = array(1, 1,2,3,4,5,6,7);
	$campus[1][2][] = array(2, 8);
	$campus[1][3][] = array(3, 9);
	$campus[1][4][] = array(30, 122);

	$special = true;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[1][1], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[1][2], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[1][3], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[1][4], $supply, $dept, $from, $to, $mode, $special);
	
}

if ($_POST["c2"]=="MOS") {
	$campus[2][1][] = array(4, 10,11,12,13,14,15,16,17,18,19,20,21,22);
	$campus[2][2][] = array(5, 23,24);
	$campus[2][3][] = array(6, 25);
	$campus[2][4][] = array(23, 82, 83, 84, 85, 86, 87, 123, 124);
	
	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[2][1], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[2][2], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[2][3], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[2][4], $supply, $dept, $from, $to, $mode, $special);

}

if ($_POST["c3"]=="MUC") {
	$campus[3][1][] = array(7, 26,27,28,29,30,31,32,33,34,35,36,37,38);
	$campus[3][2][] = array(8, 39,40);
	$campus[3][3][] = array(9, 41,42);
	$campus[3][4][] = array(24, 88,89,90,91,92,93,94,95,96,97,98,99,100,101,102);
	$campus[3][5][] = array(25, 104,105);
	
	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[3][1], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[3][2], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[3][3], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[3][4], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[3][5], $supply, $dept, $from, $to, $mode, $special);
	
}

if ($_POST["c4"]=="JDN") {
	$campus[4][1][] = array(10, 43,44,45);
	$campus[4][2][] = array(11, 48,49);
	$campus[4][3][] = array(12, 50,51,52,53,54);
	$campus[4][4][] = array(13, 55,56);
	$campus[4][5][] = array(27, 113,114,115,116);
	
	// updated @2013-10-21
	$campus[4][3][] = array(24, 93,94,95,96,97,98,99);
	$campus[4][4][] = array(25, 100,101);
	
	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[4][1], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[4][2], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[4][3], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[4][4], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[4][5], $supply, $dept, $from, $to, $mode, $special);
	
}

if ($_POST["c5"]=="OPC") {
	$campus[5][1][] = array(20, 72);
	$campus[5][2][] = array(21, 73,74,75,76,77,78);
	$campus[5][3][] = array(22, 79,80,81);
	
	$special = true;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[5][1], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[5][2], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[5][3], $supply, $dept, $from, $to, $mode, $special);
	
}

if ($_POST["c6"]=="PLS") {
	$campus[6][1][] = array(16, 61,62,63);
	
	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[6][1], $supply, $dept, $from, $to, $mode, $special);
}

if ($_POST["c7"]=="TKO") {
	$campus[7][1][] = array(17, 64);

	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);
	$fp = fwriteData2($fp, $campus[7][1], $supply, $dept, $from, $to, $mode, $special);
}

if ($_POST["c8"]=="HMT") {
	$campus[8][1][] = array(26, 106,107,108,109);
	$campus[8][2][] = array(33, 136);
	

	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[8][1], $supply, $dept, $from, $to, $mode, $special);
	$fp = fwriteData2($fp, $campus[8][2], $supply, $dept, $from, $to, $mode, $special);

}

if ($_POST["c9"]=="YL") {
	$campus[9][1][] = array(28, 117);
	

	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[9][1], $supply, $dept, $from, $to, $mode, $special);

}

if ($_POST["c10"]=="CSW") {
	$campus[10][1][] = array(31, 130,131,132,133);
	

	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[10][1], $supply, $dept, $from, $to, $mode, $special);

}

if ($_POST["c11"]=="AUS") {
	$campus[11][1][] = array(32, 134,135);
	

	$special = false;
	$supply = CalcSupply($from,$to,$mode,$special);

	$fp = fwriteData2($fp, $campus[11][1], $supply, $dept, $from, $to, $mode, $special);

}
*/


fwrite($fp, "  </Table>\n");
fwrite($fp, "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n");
fwrite($fp, "   <PageSetup>\n");
fwrite($fp, "    <Layout x:Orientation=\"Landscape\"/>\n");
fwrite($fp, "   </PageSetup>\n");
fwrite($fp, "   <Unsynced/>\n");
fwrite($fp, "   <FitToPage/>\n");
fwrite($fp, "   <Print>\n");
fwrite($fp, "    <FitHeight>99</FitHeight>\n");
fwrite($fp, "    <ValidPrinterInfo/>\n");
fwrite($fp, "    <PaperSizeIndex>9</PaperSizeIndex>\n");
fwrite($fp, "    <VerticalResolution>0</VerticalResolution>\n");
fwrite($fp, "   </Print>\n");

fwrite($fp, "   <Zoom>75</Zoom>\n");
fwrite($fp, "  </WorksheetOptions>\n");
fwrite($fp, " </Worksheet>\n");
fwrite($fp, "</Workbook>\n");

flock($fp, LOCK_UN);
fclose($fp);

$url = "CampusUsage/DeptUsage(".$dept.")_".$fid.".xls";

?>
</table>

<br>

<A href="<?php echo $url ?>">如要下載檔案，請在此處用滑鼠右擊並另存目標，下載時可以把檔案名稱更改。</A><br>
<br>
<A href="javascript:history.go(-1)">&nbsp;返回上一頁</A>


<?php     

// @ 2013-8-27
// unify a fwrite function
//
function fwriteData2($fwriteData, $campusArray, $supply, $dept, $from, $to, $mode, $special)
{
	
	$fp = $fwriteData;
	
	$totalActual_V = 0;
	$totalReserve_V = 0;
	$totalExtra_V = 0;
	$totalSupply=0;
	foreach ($campusArray as $rm)
	{
		$areaName = GetAreaName($rm[0]);
		fwrite($fp, "   <Row>\n");
		fwrite($fp, "    <Cell><Data ss:Type=\"String\">".$areaName."</Data></Cell>\n");
		fwrite($fp, "    <Cell/>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">".$dept."</Data></Cell>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">".$dept."</Data></Cell>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">".$dept."</Data></Cell>\n");
		fwrite($fp, "   </Row>\n");
	
		echo "<tr><td>".$areaName."</td><td>&nbsp;</td><td>".$dept."&nbsp;</td><td>".$dept."&nbsp;</td>";
		echo "<td>".$dept."&nbsp;</td></tr>\n";
	
		fwrite($fp, "   <Row>\n");
		fwrite($fp, "    <Cell/>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">供應</Data></Cell>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">總預算</Data></Cell>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">實際(連續)</Data></Cell>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">實際(非連續)</Data></Cell>\n");
		fwrite($fp, "   </Row>\n");
	
		echo "<tr><td>&nbsp;</td><td>供應</td><td>總預算</td><td>實際(連續)</td><td>實際(非連續)</td></tr>\n";
	
		for ($i=1; $i<sizeof($rm); $i++)
		{
			$roomName = GetRoomName($rm[$i]);
	
			fwrite($fp, "   <Row>\n");
			fwrite($fp, "    <Cell ss:StyleID=\"Left\"><Data ss:Type=\"String\">".$roomName."</Data></Cell>\n");
			fwrite($fp, "    <Cell><Data ss:Type=\"Number\">".$supply."</Data></Cell>\n");
			
			echo "<tr>\n";
			echo "<td>".$roomName."</td>";
			echo "<td>".$supply."</td>";
		
			$actual=CalcActualUsage($dept,$from,$to,$rm[0],$rm[$i],$mode,$special);
			$reserve=CalcReserve($dept,$from,$to,$rm[0],$rm[$i],$mode,$special);
			$extra=CalcExtraUsage($dept,$from,$to,$rm[0],$rm[$i],$mode,$special);
			$reserve=$reserve+$extra;

			if ($supply==0)
			{
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">0</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">0</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">0</Data></Cell>\n");
				echo "<td>0%</td>";
				echo "<td>0%</td>";
				echo "<td>0%</td>";
			} else
			{
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">".($reserve/$supply)."</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">".($actual/$supply)."</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">".($extra/$supply)."</Data></Cell>\n");
				echo "<td>".round($reserve/$supply*100)."%</td>";
				echo "<td>".round($actual/$supply*100)."%</td>";
				echo "<td>".round($extra/$supply*100)."%</td></tr>\n";
			}
			$totalActual_V = $totalActual_V + $actual;
			$totalReserve_V = $totalReserve_V + $reserve;
			$totalExtra_V = $totalExtra_V + $extra;
			$totalSupply = $totalSupply + $supply;
			fwrite($fp, "   </Row>\n");
		}
	
		fwrite($fp, "   <Row>\n");
		fwrite($fp, "    <Cell ss:StyleID=\"Right\"><Data ss:Type=\"String\">TOTAL</Data></Cell>\n");
		fwrite($fp, "    <Cell><Data ss:Type=\"Number\">".$totalSupply."</Data></Cell>\n");
		
		echo "<tr>\n";
		echo "<td>TOTAL</td>";
		echo "<td>".$totalSupply."</td>";
	
		if ($totalSupply==0)
		{
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">0</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">0</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">0</Data></Cell>\n");
				echo "<td>0%</td>";
				echo "<td>0%</td>";
				echo "<td>0%</td>";
		} else
		{
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">".($totalReserve_V/$totalSupply)."</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">".($totalActual_V/$totalSupply)."</Data></Cell>\n");
				fwrite($fp, "    <Cell ss:StyleID=\"Percent\"><Data ss:Type=\"Number\">".($totalExtra_V/$totalSupply)."</Data></Cell>\n");
				echo "<td>".round($totalReserve_V/$totalSupply*100)."%</td>";
				echo "<td>".round($totalActual_V/$totalSupply*100)."%</td>";
				echo "<td>".round($totalExtra_V/$totalSupply*100)."%</td></tr>\n";
		}
	
		fwrite($fp, "   </Row>\n");
	
		fwrite($fp, "   <Row/>\n");
	
		echo "<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>\n";
	}
	
	return $fp;
	
}
?>



