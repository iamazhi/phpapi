<?php
echo header("Access-Control-Allow-Origin: *");
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
date_default_timezone_set('Europe/London');

$type = $_GET['type'];
if(strpos("pl|common", $type) === false) return print "你是我的小苹果！";

include "$type/config.php";

/** Include PHPExcel_IOFactory */
require_once './phpexcel/Classes/PHPExcel/IOFactory.php';

$filename = $params["filename"] . "(" . date("YmdHis") . ")";

// Save Excel 2007 file
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("data/$filename.xlsx");

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("data/$filename.xls");

echo "http://" . $_SERVER["SERVER_NAME"] . "/data/$filename.xlsx";
?>
