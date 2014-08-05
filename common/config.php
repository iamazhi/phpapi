<?php
require_once '../phpexcel/Classes/PHPExcel.php';
$objPHPExcel = new PHPExcel();
include "../phpexcel/common.php";

// $_POST["data"] = <<<EOT
// {
// "basic":{"title":"福州台江美术"},
// "style":{"A1":"font-size:16;font-bold:true;align-center:true","A2:G2":"font-bold:true","B":"width:20","C":"width:30"},
// "A1:G1":{"title":"福州台江美术日报表"},
// "A2:G2":{"id":"序号", "date ":"日期",       "desc ":"摘要",             "income ":"收入", "payout ":"支出", "remain ":"余额", "remark ":"备注"},
// "A3:G3":{"id":1,      "date ":"2014-07-22", "desc ":"课堂用品",         "income ":"",     "payout ":147.2,  "remain ":"",     "remark ":""},
// "A4:G4":{"id":2,      "date ":"2014-07-22", "desc ":"办公费",           "income ":"",     "payout ":50,     "remain ":"",     "remark ":""},
// "A5:G5":{"id":3,      "date ":"2014-07-22", "desc ":"市内交通费",       "income ":"",     "payout ":109,    "remain ":"",     "remark ":""},
// "A6:G6":{"id":4,      "date ":"2014-07-23", "desc ":"课堂用品",         "income ":"",     "payout ":888,    "remain ":"",     "remark ":""},
// "A7:G7":{"id":5,      "date ":"2014-07-23", "desc ":"其他（前台销售）", "income ":"",     "payout ":234,    "remain ":"",     "remark ":""}
// }
// EOT;

//$_POST["data"] = str_replace("\\", "", $_POST["pldata"]);
$data = json_decode($_POST["data"], true);

//表格属性
$objPHPExcel->getProperties()->setCreator("吾幼内控管理系统")->setTitle($data["basic"]["title"])->setSubject("日记账");
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(16);

//样式设计
foreach($data["style"] as $scope => $style){
    if(strpos($scope, ":") !== false){
        $cells = set($scope);
        foreach($cells as $cell) setStyle($cell, $style);
    }else{
        setStyle($scope, $style);
    }
}

unset($data["basic"]);
unset($data["style"]);

//填充真实数据
foreach($data as $scope => $values){
    $values = array_values($values);

    $cells = set($scope);
    set($scope, "setBorder");
    set($scope, "setVAlign");
    if(count($values) ==  1) $objPHPExcel->getActiveSheet()->mergeCells($scope);
    foreach($cells as $index => $cell) if(isset($values[$index])) $objPHPExcel->getActiveSheet()->setCellValue($cell, $values[$index]);
}
