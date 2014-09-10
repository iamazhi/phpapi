<?php
error_reporting(E_ALL);
date_default_timezone_set("Asia/Shanghai");

if($_GET)  extract($_GET);
if($_POST) extract($_POST);

function falseReturn($info){
    $return            = array();
    $return["success"] = false;
    $return["datas"]   = $info;
    // print_r($return);
    $return            = json_encode($return);
    return $return;
}

function trueReturn($info){
    $return            = array();
    $return["success"] = true;
    $return["datas"]   = $info;
    // print_r($return);
    $return            = json_encode($return);
    return $return;
}

/* 判断参数是否正确 */
if(!isset($actor) || !isset($action) || !isset($target) || !isset($param) || !isset($token)) return falseReturn("wrong query");
if($token != "cptbtptp") return falseReturn("wrong token");
if(!isset($extra))   $extra = "";
if(!isset($comment)) $comment = "";

/* 判断方法是存在 */
$function = $action . ucfirst($target);
if(!function_exists($function)) return falseReturn("function:$function() not exists");

/* 链接数据库 */
$con = mysql_connect("192.168.1.16", "root", "root");
if(!$con) return falseReturn(mysql_error());
mysql_query("SET NAMES UTF8");
if(!mysql_select_db("wooyou_test", $con)) return falseReturn("failed to access db:wooyou_test");

/* 调用对应方法 */
record();
return print $function();

function changeScore(){
    global $con, $param, $extra;
    $condition = "pid='$param'";
    $sql = "update parent set score=score+$extra where $condition";
    $result = mysql_query($sql, $con);
    if(!$result) return falseReturn("failed to exec sql:$sql");
    return selectMemberByID();
}

function selectMemberByID(){
    global $con, $param;
    $condition = "parent.pid=$param";
    $sql = "select parent.*, child.center_id from parent left join child on child.pid=parent.pid where $condition";
    $result = mysql_query($sql, $con);
    if(!$result) return falseReturn("failed to exec sql:$sql");
    while($row = mysql_fetch_array($result, MYSQL_ASSOC)) return trueReturn($row);
    return falseReturn("no matched item");
}

function selectMemberByMobile(){
    global $con, $param;
    $condition = "telephone=$param";
    $sql = "select * from parent where $condition";
    $result = mysql_query($sql, $con);
    if(!$result) return falseReturn("failed to exec sql:$sql");
    while($row = mysql_fetch_array($result, MYSQL_ASSOC)) return trueReturn($row);
    return falseReturn("no matched item");
}

function selectCenters(){
    global $con;
    $sql = "select cid,name from center";
    $result = mysql_query($sql, $con);
    if(!$result) return falseReturn("failed to exec sql:$sql");

    $centerList = array();
    while($row = mysql_fetch_array($result, MYSQL_ASSOC)){
        $centerList[] = array("desc" => $row["name"], "value" => $row["cid"]);
    }
    return $centerList ? trueReturn($centerList) : falseReturn("no matched item");
}

function selectCenterName(){
    global $con, $param;
    $condition = "cid='$param'";
    $sql = "select cid,name from center where $condition";
    $result = mysql_query($sql, $con);
    if(!$result) return falseReturn("failed to exec sql:$sql");
    while($row = mysql_fetch_array($result, MYSQL_ASSOC)) return trueReturn($row);
    return falseReturn("no matched item");
}


function updatePassword(){
    global $con, $param, $extra;
    $condition = "pid='$param'";
    $sql = "update parent set password='$extra' where $condition";
    $result = mysql_query($sql, $con);
    if(!$result) return falseReturn("failed to exec sql:$sql");
    return selectMemberByID();
}

function record(){
    global $con, $actor, $action, $target, $param, $extra, $comment;
    $now = date('YmdHis');
    $sql = "insert record (actor, action, target, param, extra, comment, time) values('$actor', '$action', '$target', '$param', '$extra', '$comment', '$now');";
    return mysql_query($sql, $con);
}

function selectRecordScore(){
    global $con, $param;
    $condition = "target='score' and param=$param";
    $sql = "select * from record where $condition";
    $result = mysql_query($sql, $con);
    if(!$result) return falseReturn("failed to exec sql:$sql");

    $scoreList = array();
    while($row = mysql_fetch_array($result, MYSQL_ASSOC)) $scoreList[] = $row;

    return $scoreList ? trueReturn($scoreList) : falseReturn("no matched item");
}

?>
