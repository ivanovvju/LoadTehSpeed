<?php
require_once('config.php');

class Database {
 
 private static $connect;

 public static function connect() {
   global $DB, $DBPASS, $DBUSER;
   self::$connect = odbc_connect($DB, $DBUSER, $DBPASS);
   if (!self::$connect) {
    echo "<p><b>К сожалению, не удалось подключиться к базе данных</b></p>";
    exit();
    return false;

    //throw new Exception("Невозможно соединиться с базой данных", E_USER_ERROR);
   }
    
 }
 
 public static function disconnect() {
   if (self::$connect) {
      odbc_close(self::$connect);
   }
 }
 
 //select из базы------------------------
 public static function select($sql) {
  $result=odbc_exec(self::$connect,$sql);  
  if (!is_resource($result)) {
    //вывод ошибки
    $err = odbc_errormsg(self::$connect);
    throw new Exception($err);
  } 
  
  $arReturn = array();
  
  while ($row=@odbc_fetch_array($result)) {
    $arReturn[] = $row;  
  }
  
  //$result->close();
  
  return $arReturn; 
 }
 
 //получаем ассоциативный массив, для пар id - значение
 public static function select1($sql, $field1, $field2) {
    
  $result=odbc_exec(self::$connect,$sql);  
  if (!is_resource($result)) {
    //вывод ошибки
    $err = odbc_errormsg(self::$connect);
    throw new Exception($err);
  } 
 
  $arReturn = array();
  
  while ($row=@odbc_fetch_array($result)) {
    $val = self::Encoding($row[$field2]);
    $arReturn[$row[$field1]] = $val;  
  }
  
  return $arReturn; 
 }
 
 //$tab - имя таблицы
 //$arCond - ассоциативный массив пар "имя поля - значение" для where
 //$arfields - массив полей для выборки
 public static function selectW($tab, $arfields, $arCond) {
  $arWhere = array();
  foreach ($arCond as $field => $val) {
    //if (!is_numeric($val)) {
    //    $val = "'".$val."'";
    //}
    
    $arWhere[] = $field." = ".$val; 
  }
  
  $sql = "SELECT ".join(", ", $arfields)." FROM ".$tab;
  $sql = $sql." WHERE ". join(" AND ", $arWhere); 
  //echo $sql."<br><br>";
  
  $result=odbc_exec(self::$connect,$sql);  
  if (!is_resource($result)) {
    //вывод ошибки
    $err = odbc_errormsg(self::$connect);
    throw new Exception($err);
  } 
  
  $arReturn = array();
  
  while ($row=@odbc_fetch_array($result)) {
    $arReturn[] = $row;    
  }
  
  return $arReturn; 
   
 }
 
 public static function upd_ins($sql) {
  
   //odbc_exec(self::$connect,$sql) or die("<p>".odbc_errormsg());
  $bool = false;   
  $bool = odbc_exec(self::$connect,$sql);
  return $bool;

     
 } 
 
 //удаление из базы-----------------------
 //$table - таблица из которой удаляем
 //$arCond - ассоциативный массив пар "имя поля - значение" для where
 public static function delete($table, $arCond) {
  $arWhere = array();
  foreach ($arCond as $field => $val) {
    if (!is_numeric($val)) {
        $val = "'".$val."'";
    }
    
    $arWhere[] = $field." = ".$val; 
  }
 
  $sql = "DELETE FROM ".$table." WHERE " . join(' AND ', $arWhere);
  $result=odbc_exec(self::$connect,$sql);  
  
  if (!is_resource($result)) {
    //вывод ошибки
    $err = odbc_errormsg(self::$connect);
    throw new Exception($err);
  }
     
 }
 
 static function Encoding($per) {
    return iconv('windows-1251', 'UTF-8', $per);  
  }
}
?>