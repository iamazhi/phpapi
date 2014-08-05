<?php
function setBgColor($cell, $bgColor)
{
    global $objPHPExcel;
    $objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB($bgColor);
}

function setBorder($cell, $borderColor = "FF000000")
{
    global $objPHPExcel;
    $style = array(
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => $borderColor),
            ),
        ),
    );
    $objPHPExcel->getActiveSheet()->getStyle("$cell:$cell")->applyFromArray($style);
}

function setVAlign($cell, $params = "")
{
    global $objPHPExcel;
    $objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
}

//setcellvalue("C8", "=SUM(%s9:%s30)")->setCellValue("C8", "=SUM(C9:C30)");
//setcellvalue("P5", "=SUM(C%d:O%d)")->setcellvalue("P5", "=SUM(C5:O5)");
function setCellValue($cell, $value)
{
    // $cell=C8, $value=SUM(%s9:%s30)
    global $objPHPExcel;
    $letter = preg_replace("|[0-9]+|", "", $cell); //C
    $number = str_replace($letter, "", $cell); //8
    $value  = str_replace("%s", $letter, $value); //SUM(C9:C30)
    $value  = str_replace("%d", $number, $value);
    $objPHPExcel->getActiveSheet()->setCellValue($cell, $value);
}

// $css=font-size:16;font-bold:true;align-center:true
function setStyle($cell, $css){
    global $objPHPExcel;
    $params  = explode(";", $css);
    $matches = preg_grep ('/^font-size|^font-bold|^align-center|^width|^height/i', $params);
    if(!$matches) continue;
    foreach($matches as $match){
        $cssItem = explode(":", $match);
        $key     = $cssItem[0];
        $value   = $cssItem[1];
        switch($key){
        case "font-size":    $objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->setSize($value);break;
        case "font-bold":    $objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->setBold(true);break;
        case "align-center": $objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);break;
        case "width":        $objPHPExcel->getActiveSheet()->getColumnDimension($cell)->setWidth($value);break;
        default:break;
        }
    }
}

//multiSetting通用方法，范围内修改
function set($scope, $func = "", $params = "")
{
    $abc         = array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z");
    $cells       = array();

    // $scope=C8:O8, $func=setcellvalue, $params="=SUM(%s9:%s30)"
    // $start=C8, $end=O8, $startLetter=C, $endLetter=O, $startNumber=8, $endNumber=8, $startPos=2,$endPos=14
    $scopeArray  = explode(":", $scope);
    $start       = $scopeArray[0];
    $end         = $scopeArray[1];
    $startLetter = preg_replace("|[0-9]+|", "", $start);
    $endLetter   = preg_replace("|[0-9]+|", "", $end);
    $startNumber = str_replace($startLetter, "", $start);
    $endNumber   = str_replace($endLetter, "", $end);
    $startPos    = array_search($startLetter, $abc);
    $endPos      = array_search($endLetter, $abc);

    // $cells：取到scope范围内所有的cell
    for($i = $startPos; $i <= $endPos; $i++) //2-14
    {
        for($j = $startNumber; $j <= $endNumber; $j++)//8-8
        {
            $currentCell = $abc[$i] . "$j";//C8->D8....->O8
            $cells[] = $currentCell;
            if($func) $func($currentCell, $params);//setcellvalue(C8,"=SUM(%s9:%s30)")

        }
    }
    return $cells;
}
?>
