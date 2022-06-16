<?php

/**
 * @author admin
 * @copyright 2016
 */

class log {

    /**
     * @var string Файл, в который аписывать логи.
     */
    protected static $file = "log/log_%date%.txt";

    /**
     * Запись информационного сообщения.
     * @param $message - Сообщение для записи в лог.
     */
    public static function Info($message)
    {
        $dateObj = new DateTime();
        $dateObjPr = $dateObj->modify('0 day');
        $date = $dateObjPr->format('Y-m');
        $dateTime = $dateObjPr->format('Y-m-d H:i:s');

        $file = str_replace("%date%", $date, self::$file);

        $message = "$dateTime - INFO: $message\n";

        file_put_contents(  $file, $message, FILE_APPEND);
    }

    /**
     * Запись ошибки.
     * @param $message - Сообщение для записи в лог.
     */
    public static function Error($message)
    {
        $dateObj = new DateTime();
        $dateObjPr = $dateObj->modify('0 day');
        $date = $dateObjPr->format('Y-m');
        $dateTime = $dateObjPr->format('Y-m-d H:i:s');

        $file = str_replace("%date%", $date, self::$file);

        $message = "$dateTime - ERROR: $message\n";

        file_put_contents(  $file, $message, FILE_APPEND);
    }

    /**
     * Запись предупредительного сообщения.
     * @param $message - Сообщение для записи в лог.
     */
    public static function Warn($message)
    {
        $dateObj = new DateTime();
        $dateObjPr = $dateObj->modify('0 day');
        $date = $dateObjPr->format('Y-m');
        $dateTime = $dateObjPr->format('Y-m-d H:i:s');

        $file = str_replace("%date%", $date, self::$file);

        $message = "$dateTime - WARN: $message\n";

        file_put_contents(  $file, $message, FILE_APPEND);
    }
}

?>