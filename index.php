<?php

include 'D:/wwwnew/Classes/PHPExcel.php';
include 'D:/wwwnew/libPHP/Database.php';
include 'D:/wwwnew/libPHP/log.php';

$object = new ParseExcel();

class ParseExcel
{
    /**
     * @var array ������ Excel-��������� ��� ��������.
     */
    private $listDoc = array();
    /**
     * @var string ���� �� Excel-����������.
     */
    private $pathToExcel;
    /**
     * @var array ������ �������� � ���� [������������_������� => ���_�������].
     */
    private $nameRegions;
    /**
     * @var array ������ �������� � ���� [���_������� => ������������_�������].
     */
    private $codeRegions;
    /**
     * @var string ����.
     */
    private $date;
    /**
     * @var string ���� �������� ����.
     */
    private $lastYearDate;
    /**
     * @var int ���� �������� ����. j
     */
    private $day;
    /**
     * @var string ������������� ������� - [���_������� => ������������_�������].
     */
    private $dispUch;

    public function __construct()
    {
        log::Info("--- START PROGRAM ---");

        $this->date = date("Y-m-d");
//        $this->date = date("2021-09-14");

        $raznDay = (isset($_SERVER['argv'][1])) ? -1 : 0;

        log::Warn($raznDay);
        log::Warn($_SERVER['argv'][1]);

        $this->setDate($this->date, $raznDay);

        log::Info("�������� � ����� {$this->date}");
        echo "�������� � ����� {$this->date}<br>";

        // ������������ � �� DOCLAD
        Database::connect();

        try {
            $this->getNameReg();
            $this->getDispUch();

            $this->pathToExcel = "E:\Diskor_new\\IHLP\\tech_speed\\";

            $this->iniListDoc();

            foreach ($this->listDoc as $id => $nameFile) {
                if ($this->day != 1 && ($id == 3 || $id == 6)) {
                    continue;
                }
                $pathFile = $this->pathToExcel . $nameFile;
                $parseResult = $this->parseDocument($pathFile, $id);

                $this->loadDataToDb($parseResult, $id);
            }
        } catch (Exception $ex) {
            log::Error("��������� ������. {$ex->getMessage()}");
            echo "��������� ������. {$ex->getMessage()}";
            return;
        }

        // ���������� �� �� DOCLAD
        Database::disconnect();

        log::Info("--- END PROGRAM ---");
        echo "��� ����� ����������!";

    }

    /**
     * ���������� ����������� ���������� � ������
     */
    private function setDate($date, $raznDay)
    {
        $dateObj = new DateTime($date);
        $dateObjPr = $dateObj->modify("$raznDay day");
        $this->date = $dateObjPr->format('Y-m-d');

        $dateObj = new DateTime($this->date);
        $dateObjPr = $dateObj->modify('-1 year');
        $this->lastYearDate = $dateObjPr->format('Y-m-d');

        $dateObj = new DateTime($this->date);
        $dateObjPr = $dateObj->modify('0 day');
        $this->day = $dateObjPr->format('j');
    }

    /**
     * �������������� ������������ Excel-������.
     */
    private function iniListDoc()
    {
        $this->listDoc[1] = "file1_{$this->date}.xlsx";
        $this->listDoc[2] = "file2_{$this->date}.xlsx";
        $this->listDoc[3] = "file3_{$this->date}.xlsx";
        $this->listDoc[4] = "file4_{$this->date}.xlsx";
        $this->listDoc[5] = "file5_{$this->date}.xlsx";
        $this->listDoc[6] = "file6_{$this->date}.xlsx";
    }

    /**
     * ������� Excel-������.
     * @param $nameFile - ������������ �����.
     * @param $typeData - ��� ������ (1,2,3,4,5,6 - id ������).
     * @return array - ������ � �������.
     * @throws Exception - ���� ��� ������.
     */
    private function parseDocument($nameFile, $typeData)
    {
        $dataList = array();
        $dataValue = array();
        $codeRegion = 0;

        try {
            $xls = PHPExcel_IOFactory::load($nameFile);
            $xls->setActiveSheetIndex(0);
            $sheet = $xls->getActiveSheet();

            $dataList = $sheet->toArray();
        } catch (Exception $ex) {
            log::Error("�������� ������ �� ����� ��������� ������� ������ �� Excel. ������: {$ex->getMessage()}");
            throw new Exception("�������� ������ �� ����� ��������� ������� ������ �� Excel. ������: {$ex->getMessage()}");
        }

        switch ($typeData) {
            // Excel ����������� ����������� �������� �� ����� �� ��������. - file1.xlsx
            case 1:
                $finishRow = count($dataList);
                for ($row = 4; $row < $finishRow; $row++) {
                    switch (iconv("UTF-8", "cp1251", $dataList[$row][0])) {
                        case "����������":
                            $codeRegion = 1;
                            break;

                        case "�����-������":
                            $codeRegion = 2;
                            break;

                        case "���������":
                            $codeRegion = 3;
                            break;

                        case "����������":
                            $codeRegion = 4;
                            break;

                        case "�����":
                            $codeRegion = 0;
                            break;

                        default:
                            $codeRegion = 0;
                            break;
                    }

                    $dataValue[$codeRegion][5] = (isset($dataList[$row][1])) ? $dataList[$row][1] : 0;
                    $dataValue[$codeRegion][1] = (isset($dataList[$row][2])) ? $dataList[$row][2] : 0;
                    $dataValue[$codeRegion][6] = (isset($dataList[$row][3])) ? $dataList[$row][3] : 0;
                }
                break;

            // Excel ����������� ����������� �������� ����������� ������ � ������ ������ �� ��������. - file2.xlsx
            case 2:
                $finishRow = count($dataList);
                for ($row = 4; $row < $finishRow; $row++) {
                    switch (iconv("UTF-8", "cp1251", $dataList[$row][0])) {
                        case "����������":
                            $codeRegion = 1;
                            break;

                        case "�����-������":
                            $codeRegion = 2;
                            break;

                        case "���������":
                            $codeRegion = 3;
                            break;

                        case "����������":
                            $codeRegion = 4;
                            break;

                        case "�����":
                            $codeRegion = 0;
                            break;

                        default:
                            throw new Exception("����� parseDocument. �� ����� ������ ������. ����: " . iconv("UTF-8", "cp1251", $dataList[$row][0]) . "typeData = $typeData");
                    }

                    $dataValue[$codeRegion][5] = (isset($dataList[$row][1])) ? $dataList[$row][1] : 0;
                    $dataValue[$codeRegion][1] = (isset($dataList[$row][2])) ? $dataList[$row][2] : 0;
                    $dataValue[$codeRegion][6] = (isset($dataList[$row][3])) ? $dataList[$row][3] : 0;
                }
                break;

            // Excel ����������� �������� �� ����� �� ������� ���������� �� ���������� ��� �� ��������. - file3.xlsx
            case 3:
                $finishRow = count($dataList);
                for ($row = 4; $row < $finishRow; $row++) {
                    switch (iconv("UTF-8", "cp1251", $dataList[$row][0])) {
                        case "����������":
                            $codeRegion = 1;
                            break;

                        case "�����-������":
                            $codeRegion = 2;
                            break;

                        case "���������":
                            $codeRegion = 3;
                            break;

                        case "����������":
                            $codeRegion = 4;
                            break;

                        case "�����":
                            $codeRegion = 0;
                            break;

                        default:
                            throw new Exception("����� parseDocument. �� ����� ������ ������. ����: " . iconv("UTF-8", "cp1251", $dataList[$row][0]) . "typeData = $typeData");
                            return;
                    }

                    $dataValue[$codeRegion][1] = (isset($dataList[$row][2])) ? $dataList[$row][2] : 0;
                    $dataValue[$codeRegion][5] = (isset($dataList[$row][1])) ? $dataList[$row][1] : 0;
                    $dataValue[$codeRegion][6] = (isset($dataList[$row][3])) ? $dataList[$row][3] : 0;
                }
                break;

            // Excel ����������� �������� �� ����� �� ����������� ���������� �� ������� ��� �� ������������� ��������. - file4.xlsx
            case 4:
            // Excel ����������� �������� ����������� ������ �� ����������� ���������� �� ������� ��� �� ������������� ��������.
            case 5:
            // Excel ����������� �������� �� ����� �� ������� ���������� �� ���������� ��� �� ������������� ��������.
            case 6:
                $finishRow = count($dataList);
                for ($row = 4; $row < $finishRow; $row++) {
                    $nameUch = iconv("UTF-8", "cp1251", $dataList[$row][0]);
                    $codeUch = substr(mb_eregi_replace("[^0-9]", '', $dataList[$row][0]), 0, 3);

                    if (isset($this->dispUch[$codeUch])) {
                        $dataValue[$codeUch][5] = (isset($dataList[$row][1])) ? $dataList[$row][1] : 0;
                        $dataValue[$codeUch][1] = (isset($dataList[$row][2])) ? $dataList[$row][2] : 0;
                        $dataValue[$codeUch][6] = (isset($dataList[$row][3])) ? $dataList[$row][3] : 0;
                    } elseif($nameUch == "�����") {
                        $dataValue[999][5] = (isset($dataList[$row][1])) ? $dataList[$row][1] : 0;
                        $dataValue[999][1] = (isset($dataList[$row][2])) ? $dataList[$row][2] : 0;
                        $dataValue[999][6] = (isset($dataList[$row][3])) ? $dataList[$row][3] : 0;
                    } else {
                        log::Warn("������� '$nameUch' �� ������ � ���.");
                    }
                }
                break;

            default:
                log::Warn("������� �������� ID ��������� ��� ��������.");
                for ($row = 0; $row < 5; $row++) {
                    $dataValue[$row][1] = 0;
                    $dataValue[$row][5] = 0;
                    $dataValue[$row][6] = 0;
                }
                break;
        }

        return $dataValue;

    }

    /**
     * �������� ������ � ��.
     * @param $data - ������ ��� ������
     * @param $typeData - id ���������, ������� ��������
     * @throws Exception
     */
    private function loadDataToDb($data, $typeData)
    {
        $date = "";
        $table = "DISKOR.FACT_TEH_SPEED_REGION";
        $nameColumns = "(CODE_REG, CODE_TYPE_POKAZ, TYPE_DATA, VAL, REPORT_DATE)";
        $listDataToDb = array();
        $fileName = $this->listDoc[$typeData];

        if ($typeData >= 4) {
            $table = "DISKOR.FACT_TEH_SPEED_DISP_UCH";
            $nameColumns = "(CODE_DISP_UCH, CODE_TYPE_POKAZ, TYPE_DATA, VAL, REPORT_DATE)";
        }

        switch ($typeData) {
            // ������ �� �����.
            case 1:
            case 4:
                $typeData = 'S';
                $date = $this->date;
                break;
            // ������ �� ����� �����������.
            case 2:
            case 5:
                $typeData = 'M';
                $date = $this->date;
                break;
            // ������ �� ����� ����������� - ������� ���.
            case 3:
            case 6:
                $typeData = 'L_M';
                $date = $this->lastYearDate;
                break;
            // ���� �������� ���-�� �� ��.
            default:
                throw new Exception("����� loadDataToDb. � �������� typeData �������� ������������ ������: $typeData");
        }
        if ($typeData == 'S' || $typeData == 'M' || $typeData == 'L_M') {

            foreach ($data as $codeReg => $dataReg) {
                foreach ($dataReg as $key => $val) {
                    $val = str_replace(",", "", $val);
                    $val = str_replace(" ", "", $val);
                    $val = ($val == "") ? 0.00 : round($val, 2);
                    $listDataToDb[] = "($codeReg, $key, '$typeData', $val, '$date')";
                }
            }

            $values = implode("," , $listDataToDb);

            if ($values == "") {
                throw new Exception("������ ��� ������ ����! typeData = $typeData");
            }

            $sql = "delete from " . $table . " where REPORT_DATE='$date' and CODE_TYPE_POKAZ in (1,5,6) and TYPE_DATA  = '$typeData'";
            if (!Database::upd_ins($sql)) {
                throw new Exception("��������� ������ �� ����� �������� ������ �� ��. SQL: $sql");
            }

            $sql = "insert into $table $nameColumns values $values";
            Database::upd_ins($sql);

            log::Info("������� �������� ������ �� ����� $fileName!");
            echo "������� �������� ������ �� ����� $fileName!<br>";
        }
    }

    /**
     * ��������� ������������ ��������.
     */
    private function getNameReg() {
        $sql = "select ID, NAME from NSI_API.REGION";

        $result = Database::select($sql);

        foreach ($result as $item) {
            $codeReg = $item['ID'];
            $nameReg = $item['NAME'];

            $this->nameRegions[$nameReg] = $codeReg;
            $this->codeRegions[$codeReg] = $nameReg;
        }
    }

    /**
     * ��������� ����� ������������� ��������.
     */
    private function getDispUch() {
        $sql = "select ID, NAME_UCH, CODE_UCH from NSI_API.DISP_UCH";

        $result = Database::select($sql);

        foreach ($result as $item) {
            $codeUch = $item['CODE_UCH'];
            $nameUch = $item['NAME_UCH'];

            $this->dispUch[$codeUch] = $nameUch;
        }
    }

}