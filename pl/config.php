<?php
require_once dirname(__FILE__) . '/../phpexcel/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();
include dirname(__FILE__) . "/../phpexcel/common.php";

//表格属性
$objPHPExcel->getProperties()->setCreator("吾幼内控管理系统")->setTitle("PL表")->setSubject("损益表");

$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(30);
$objPHPExcel->getActiveSheet()->getRowDimension('9')->setRowHeight(40);

//标题
$objPHPExcel->getActiveSheet()->setCellValue('A1', date('Y') . '中心损益表');
$objPHPExcel->getActiveSheet()->mergeCells('A1:P1');
//中心公司名称
$objPHPExcel->getActiveSheet()->setCellValue('A2', '中心公司名称:');
$objPHPExcel->getActiveSheet()->mergeCells('A2:B2');
$objPHPExcel->getActiveSheet()->mergeCells('C2:P2');
//月份
$objPHPExcel->getActiveSheet()->setCellValue('A3', '月份:');
$objPHPExcel->getActiveSheet()->setCellValue('C3', '期初数');
$objPHPExcel->getActiveSheet()->setCellValue('D3', '1月');
$objPHPExcel->getActiveSheet()->setCellValue('E3', '2月');
$objPHPExcel->getActiveSheet()->setCellValue('F3', '3月');
$objPHPExcel->getActiveSheet()->setCellValue('G3', '4月');
$objPHPExcel->getActiveSheet()->setCellValue('H3', '5月');
$objPHPExcel->getActiveSheet()->setCellValue('I3', '6月');
$objPHPExcel->getActiveSheet()->setCellValue('J3', '7月');
$objPHPExcel->getActiveSheet()->setCellValue('K3', '8月');
$objPHPExcel->getActiveSheet()->setCellValue('L3', '9月');
$objPHPExcel->getActiveSheet()->setCellValue('M3', '10月');
$objPHPExcel->getActiveSheet()->setCellValue('N3', '11月');
$objPHPExcel->getActiveSheet()->setCellValue('O3', '12月');
$objPHPExcel->getActiveSheet()->setCellValue('P3', '合计');
$objPHPExcel->getActiveSheet()->mergeCells('A3:B3');
//收入：
$objPHPExcel->getActiveSheet()->setCellValue('A4', '收入:');
$objPHPExcel->getActiveSheet()->mergeCells('A4:B4');
//签单收入（按签单数）
$objPHPExcel->getActiveSheet()->setCellValue('A5', '签单收入（按签单数）');
$objPHPExcel->getActiveSheet()->mergeCells('A5:B5');
//签单收入（按收入表）
$objPHPExcel->getActiveSheet()->setCellValue('A6', '签单收入（按收入表）');
$objPHPExcel->getActiveSheet()->mergeCells('A6:B6');
//前台销售收入
$objPHPExcel->getActiveSheet()->setCellValue('A7', '前台销售收入');
$objPHPExcel->getActiveSheet()->mergeCells('A7:B7');

////运营成本：=SUM(C9:C30))
$objPHPExcel->getActiveSheet()->setCellValue('A8', '运营成本：');
$objPHPExcel->getActiveSheet()->mergeCells('A8:B8');

//固定支出: 房租，物业、水电费 ,工资 ,社保 ,住房公积金 ,绿植租金 ,打印机租金 ,主营业务税金及附加 ,工会经费
$objPHPExcel->getActiveSheet()->setCellValue('A9',  '固定支出');
$objPHPExcel->getActiveSheet()->setCellValue('B9',  '房租');
$objPHPExcel->getActiveSheet()->setCellValue('B10', '物业、水电费');
$objPHPExcel->getActiveSheet()->setCellValue('B11', '工资');
$objPHPExcel->getActiveSheet()->setCellValue('B12', '社保');
$objPHPExcel->getActiveSheet()->setCellValue('B13', '住房公积金');
$objPHPExcel->getActiveSheet()->setCellValue('B14', '绿植租金');
$objPHPExcel->getActiveSheet()->setCellValue('B15', '打印机租金');
$objPHPExcel->getActiveSheet()->setCellValue('B16', '主营业务税金及附加');
$objPHPExcel->getActiveSheet()->setCellValue('B17', '工会经费');
$objPHPExcel->getActiveSheet()->mergeCells('A9:A17');

//变动支出： 耗材&商品成本 ,低值易耗品 ,办公费 ,运费 税费 、快递费 ,通讯费 ,福利费 ,交际费 ,市内交通费 ,差旅费 ,广告及市场推广费 ,其他 ,税费 ,财务费用及刷卡手续费
$objPHPExcel->getActiveSheet()->setCellValue('A18', '变动支出');
$objPHPExcel->getActiveSheet()->setCellValue('B18', '耗材&商品成本');
$objPHPExcel->getActiveSheet()->setCellValue('B19', '低值易耗品');
$objPHPExcel->getActiveSheet()->setCellValue('B20', '办公费');
$objPHPExcel->getActiveSheet()->setCellValue('B21', '运费、快递费');
$objPHPExcel->getActiveSheet()->setCellValue('B22', '通讯费');
$objPHPExcel->getActiveSheet()->setCellValue('B23', '福利费');
$objPHPExcel->getActiveSheet()->setCellValue('B24', '交际费');
$objPHPExcel->getActiveSheet()->setCellValue('B25', '市内交通费');
$objPHPExcel->getActiveSheet()->setCellValue('B26', '差旅费');
$objPHPExcel->getActiveSheet()->setCellValue('B27', '广告及市场推广费');
$objPHPExcel->getActiveSheet()->setCellValue('B28', '其他');
$objPHPExcel->getActiveSheet()->setCellValue('B29', '税费');
$objPHPExcel->getActiveSheet()->setCellValue('B30', '财务费用及刷卡手续费');
$objPHPExcel->getActiveSheet()->mergeCells('A18:A30');

//运营利润（不扣筹建费按签单单数确认收入）=C5+C7-C8
$objPHPExcel->getActiveSheet()->setCellValue('A31', '运营利润（不扣筹建费按签单单数确认收入）');
$objPHPExcel->getActiveSheet()->mergeCells('A31:B31');

//运营利润（扣筹建费按收入表确认收入）=C6+C7-C8
$objPHPExcel->getActiveSheet()->setCellValue('A32', '运营利润（扣筹建费按收入表确认收入）');
$objPHPExcel->getActiveSheet()->mergeCells('A32:B32');

//筹建费用：装修费 ,设计费 ,开办费 ,验资费 ,固定资产 ,在建工程 ,低值易耗品 ,购入办公用品等 ,开业前期房租及押金 ,开业前期物业水电费 ,开业前期广告投放费 ,开业前期人事费用 ,其他成本支出 ,其他
$objPHPExcel->getActiveSheet()->setCellValue('A33', '筹建费用');
$objPHPExcel->getActiveSheet()->mergeCells('A33:B33');

//筹建费用：装修费 ,设计费 ,开办费 ,验资费 ,固定资产 ,在建工程 ,低值易耗品 ,购入办公用品等 ,开业前期房租及押金 ,开业前期物业水电费 ,开业前期广告投放费 ,开业前期人事费用 ,其他成本支出 ,其他
$objPHPExcel->getActiveSheet()->setCellValue('A34', '装修费');
$objPHPExcel->getActiveSheet()->setCellValue('A35', '设计费');
$objPHPExcel->getActiveSheet()->setCellValue('A36', '开办费');
$objPHPExcel->getActiveSheet()->setCellValue('A37', '验资费');
$objPHPExcel->getActiveSheet()->setCellValue('A38', '固定资产');
$objPHPExcel->getActiveSheet()->setCellValue('A39', '在建工程');
$objPHPExcel->getActiveSheet()->setCellValue('A40', '低值易耗品');
$objPHPExcel->getActiveSheet()->setCellValue('A41', '购入办公用品等');
$objPHPExcel->getActiveSheet()->setCellValue('A42', '开业前期房租及押金');
$objPHPExcel->getActiveSheet()->setCellValue('A43', '开业前期物业水电费');
$objPHPExcel->getActiveSheet()->setCellValue('A44', '开业前期广告投放费');
$objPHPExcel->getActiveSheet()->setCellValue('A45', '开业前期人事费用');
$objPHPExcel->getActiveSheet()->setCellValue('A46', '其他成本支出');
$objPHPExcel->getActiveSheet()->setCellValue('A47', '其他');
$objPHPExcel->getActiveSheet()->mergeCells('A34:B34');
$objPHPExcel->getActiveSheet()->mergeCells('A35:B35');
$objPHPExcel->getActiveSheet()->mergeCells('A36:B36');
$objPHPExcel->getActiveSheet()->mergeCells('A37:B37');
$objPHPExcel->getActiveSheet()->mergeCells('A38:B38');
$objPHPExcel->getActiveSheet()->mergeCells('A39:B39');
$objPHPExcel->getActiveSheet()->mergeCells('A40:B40');
$objPHPExcel->getActiveSheet()->mergeCells('A41:B41');
$objPHPExcel->getActiveSheet()->mergeCells('A42:B42');
$objPHPExcel->getActiveSheet()->mergeCells('A43:B43');
$objPHPExcel->getActiveSheet()->mergeCells('A44:B44');
$objPHPExcel->getActiveSheet()->mergeCells('A45:B45');
$objPHPExcel->getActiveSheet()->mergeCells('A46:B46');
$objPHPExcel->getActiveSheet()->mergeCells('A47:B47');

//其他费用：退课费, 折旧费
$objPHPExcel->getActiveSheet()->setCellValue('A48', '其他费用');
$objPHPExcel->getActiveSheet()->setCellValue('A49', '退课费');
$objPHPExcel->getActiveSheet()->setCellValue('A50', '折旧费');
$objPHPExcel->getActiveSheet()->mergeCells('A48:B48');
$objPHPExcel->getActiveSheet()->mergeCells('A49:B49');
$objPHPExcel->getActiveSheet()->mergeCells('A50:B50');
//总利润（扣筹建费按签单单数确认收入)
$objPHPExcel->getActiveSheet()->setCellValue('A51', '总利润（扣筹建费按签单单数确认收入)');
$objPHPExcel->getActiveSheet()->mergeCells('A51:B51');
//总利润（扣筹建费按收入表确认收入）
$objPHPExcel->getActiveSheet()->setCellValue('A52', '总利润（扣筹建费按收入表确认收入）');
$objPHPExcel->getActiveSheet()->mergeCells('A52:B52');

//填充真实数据
//$_POST["data"] = '{"params":{"filename":"阿罗海中心(20140701-20140831)", "centerName":"阿罗海中心"},"640101":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":32000,"8":0,"9":0,"10":0,"11":0,"12":0},"640102":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":2077.1,"8":0,"9":0,"10":0,"11":0,"12":0},"640105":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":10307.630000000001,"8":0,"9":0,"10":0,"11":0,"12":0},"640106":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":741.12,"8":0,"9":0,"10":0,"11":0,"12":0},"640107":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":33229.4,"8":0,"9":0,"10":0,"11":0,"12":0},"640202":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":49,"8":0,"9":0,"10":0,"11":0,"12":0},"640204":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":278,"8":0,"9":0,"10":0,"11":0,"12":0},"640205":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":3019.8,"8":0,"9":0,"10":0,"11":0,"12":0},"640206":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":5880,"8":0,"9":0,"10":0,"11":0,"12":0},"660204":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":140,"8":0,"9":0,"10":0,"11":0,"12":0},"660207":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":4595,"8":0,"9":0,"10":0,"11":0,"12":0},"660210":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":1520.6,"8":0,"9":0,"10":0,"11":0,"12":0},"660211":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":1610,"8":0,"9":0,"10":0,"11":0,"12":0},"660212":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":12637.4,"8":0,"9":0,"10":0,"11":0,"12":0},"66020102":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":2668.5,"8":0,"9":0,"10":0,"11":0,"12":0},"66020201":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":312,"8":0,"9":0,"10":0,"11":0,"12":0},"66020204":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":1318,"8":0,"9":0,"10":0,"11":0,"12":0},"66020301":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":1000,"7":8692.8,"8":0,"9":0,"10":0,"11":0,"12":0},"66020302":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":463.6,"8":0,"9":0,"10":0,"11":0,"12":0},"66020303":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":1290,"8":0,"9":0,"10":0,"11":0,"12":0},"66020304":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":1049.93,"8":0,"9":0,"10":0,"11":0,"12":0},"66020305":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":225.8,"8":0,"9":0,"10":0,"11":0,"12":0},"66020306":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":1600,"8":0,"9":0,"10":0,"11":0,"12":0},"66020307":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":2509.7000000000003,"8":0,"9":0,"10":0,"11":0,"12":0},"66020502":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":341,"8":0,"9":0,"10":0,"11":0,"12":0},"66020503":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":717.5,"8":0,"9":0,"10":0,"11":0,"12":0},"qiandanshu":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":175110,"8":0,"9":0,"10":0,"11":0,"12":0},"shourushu":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":170300,"8":0,"9":0,"10":0,"11":0,"12":0},"qiantaixiaoshou":{"1":0,"2":0,"3":0,"4":0,"5":0,"6":0,"7":12136,"8":0,"9":0,"10":0,"11":0,"12":0}}';
$_POST["data"] = str_replace("\\", "", $_POST["data"]);
$data = json_decode($_POST["data"]);

$params = (array)$data->params;
unset($data->params);
$objPHPExcel->getActiveSheet()->setCellValue('C2', $params["centerName"]);

$tableMap = array();
$tableMap["B5"]  = array("code" => "qiandanshu", "name" => "签单收入（按签单数）", "months" => "D5:O5");
$tableMap["B6"]  = array("code" => "shourushu", "name" => "签单收入（按签单数）", "months" => "D6:O6");
$tableMap["B7"]  = array("code" => "qiantaixiaoshou", "name" => "前台销售收入", "months" => "D7:O7");
$tableMap["B9"]  = array("code" => "640101", "name" => "房租", "months" => "D9:O9");
$tableMap["B10"] = array("code" => "640102&66020301", "name" => "物业、水电费", "months" => "D10:O10");
$tableMap["B11"] = array("code" => "66020101&66020102", "name" => "工资", "months" => "D11:O11");
$tableMap["B14"] = array("code" => "66020306", "name" => "绿植租金", "months" => "D14:O14");
$tableMap["B17"] = array("code" => "66020902", "name" => "工会经费", "months" => "D17:O17");
$tableMap["B18"] = array("code" => "640103&640104&640105&640106&640107", "name" => "耗材&商品成本", "months" => "D18:O18");
$tableMap["B19"] = array("code" => "660204", "name" => "低值易耗品", "months" => "D19:O19");
$tableMap["B20"] = array("code" => "66020302", "name" => "办公费", "months" => "D20:O20");
$tableMap["B21"] = array("code" => "66020303", "name" => "运费", "months" => "D21:O21");
$tableMap["B22"] = array("code" => "66020304", "name" => "通讯费", "months" => "D22:O22");
$tableMap["B23"] = array("code" => "66020201&66020202&66020203&66020204", "name" => "福利费", "months" => "D23:O23");
$tableMap["B24"] = array("code" => "660206", "name" => "交际费", "months" => "D24:O24");
$tableMap["B25"] = array("code" => "66020305", "name" => "市内交通费", "months" => "D25:O25");
$tableMap["B26"] = array("code" => "66020501&66020502&66020503", "name" => "差旅费", "months" => "D26:O26");
$tableMap["B27"] = array("code" => "660207", "name" => "广告及市场推广费", "months" => "D27:O27");
$tableMap["B28"] = array("code" => "660212", "name" => "其他", "months" => "D28:O28");
$tableMap["B29"] = array("code" => "660209", "name" => "税费", "months" => "D29:O29");
$tableMap["A34"] = array("code" => "660211", "name" => "装修费", "months" => "D34:O34");
$tableMap["A36"] = array("code" => "660210", "name" => "开办费", "months" => "D36:O36");

foreach($tableMap as $cell => $item)
{
    $codes = explode("&", $item["code"]);
    $newMonthValues = array("1"=>0, "2"=>0, "3"=>0, "4"=>0, "5"=>0, "6"=>0, "7"=>0, "8"=>0, "9"=>0, "10"=>0, "11"=>0, "12"=>0);
    foreach($codes as $code)
    {
        if(!isset($data->{$code})) continue;
        $monthValues = $data->{$code};
        for($i=1; $i<=12; $i++) $newMonthValues[$i] += $monthValues->{$i};
    }
    $monthCellArray = set($item["months"]);
    foreach($monthCellArray as $index => $monthCell)
    {
        $objPHPExcel->getActiveSheet()->setCellValue($monthCell, $newMonthValues[$index + 1]);
    }
}

//设置公式
set("C8:O8",   "setCellValue", "=SUM(%s9:%s30)");
set("C31:O31", "setCellValue", "=%s5+%s7-%s8");
set("C32:O32", "setCellValue", "=%s6+%s7-%s8");
set("C33:O33", "setCellValue", "=SUM(%s34:%s47)");
set("C48:O48", "setCellValue", "=SUM(%s49:%s50)");
set("C51:O51", "setCellValue", "=%s31-%s33-%s48");
set("C52:O52", "setCellValue", "=%s32-%s33-%s48");
set("P5:P52",  "setCellValue", "=SUM(C%d:O%d)");

//样式设计
//标题
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(16);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);

set("A8:P8",   "setBgColor", "FFB7DEE8");
set("A33:P33", "setBgColor", "FFB7DEE8");
set("A48:P48", "setBgColor", "FFB7DEE8");
set("A31:P31", "setBgColor", "FFFCD5B4");
set("A32:P32", "setBgColor", "FFFCD5B4");
set("A51:P51", "setBgColor", "FFFCD5B4");
set("A52:P52", "setBgColor", "FFFCD5B4");
set("P3:P52",  "setBgColor", "FFB7DEE8");
set("A2:P52",  "setBorder");

$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(40);
for($i = 2; $i <= 52; $i++)
{
    $objPHPExcel->getActiveSheet()->getRowDimension($i)->setRowHeight(20);
}
set("A2:P52", "setVAlign");
