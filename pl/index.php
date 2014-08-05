<?php
echo header("Access-Control-Allow-Origin: http://wms.wooyou.com.cn");
echo "http://" . $_SERVER["SERVER_NAME"] . "/pl/output.xlsx";

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

include "config.php";

/** Include PHPExcel_IOFactory */
require_once '../phpexcel/Classes/PHPExcel/IOFactory.php';

// Save Excel 2007 file
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("output.xlsx");

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("output.xls");
?>
