<?php
/**
 * ExcelからDBにデータを投入、sqlite3ファイルを作成
 *
 * @category tool
 * @package  tool
 * @author   Nobuo Tsuchiya <develop@m.tsuchi99.net>
 * @license  http://www.opensource.org/licenses/mit-license.html  MIT License
 * @link     https://github.com/Tsuchiy/MasterManageTool
 **/

error_reporting(E_ALL);
ini_set('display_errors', 1);
set_time_limit(0);
mb_internal_encoding("utf-8");

date_default_timezone_set('Asia/Tokyo');

define("EXCEL_PATH", '../excel/');
define("SQLITE_PATH", '../sqlite/');

/** -------- edit configure -------- **/
/** Excel **/
$inputFileExtend = 'xlsx';
//  $inputFileType = 'Excel5';
    $inputFileType = 'Excel2007';
//  $inputFileType = 'Excel2003XML';
//  $inputFileType = 'OOCalc';
//  $inputFileType = 'Gnumeric';

/** Mysql **/
$dbHost = '127.0.0.1';
$dbPort = '3306';
$dbName = 'master_data';
$dbUser = 'masterUser';
$dbPassword = 'masterPassword';

/** Sqlite **/
$sqliteFileName = 'itemMaster.sqlite3';
/** -------- edit configure -------- **/

$dbDsn = 'mysql:host=' . $dbHost . ';port=' . $dbPort . ';dbname=' . $dbName . ';charset=utf8';
$sqliteDsn = 'sqlite:' . SQLITE_PATH . $sqliteFileName;

// seek Excel files
$excelFileList = array();

$dh = opendir(EXCEL_PATH);
while (($file = readdir($dh)) !== false) {
    if ($file == '.' || $file == '..') {
        continue;
    }
    if (is_file(EXCEL_PATH . $file) && preg_match('/^([^~].+)\.' . $inputFileExtend . '$/s', $file, $matches)) {
        $excelFileList[] = $file;
    }
}
closedir($dh);


/**  prepare PHPExcel **/

// add include path
set_include_path(dirname(realpath(__FILE__)) . '/PHPExcel/Classes/');

// PHPExcel_IOFactory
include 'PHPExcel/IOFactory.php';

$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objReader->setReadDataOnly(true);

/** prepare Mysql **/
$dbAttr = array(
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
    );
$pdo = new PDO($dbDsn, $dbUser, $dbPassword, $dbAttr);

/** prepare sqlite **/
$sqlitePdo = new PDO($sqliteDsn, null, null, $dbAttr);


foreach ($excelFileList as $excelFile) {

    // read file
    $inputFileName = EXCEL_PATH . DIRECTORY_SEPARATOR . $excelFile;
    $objPHPExcel = $objReader->load($inputFileName);

    // controlシートから読むシートを制御
    $controlSheet = $objPHPExcel->getSheetByName('control');
    $rownum = 1;
    $tableNames = array();
    $sheetNames = array();
    while ($sheetName = $controlSheet->getCellByColumnAndRow(1, ++$rownum)->getValue()) {
        $tableNames[$sheetName] = $controlSheet->getCellByColumnAndRow(3, $rownum)->getValue();
    }

    // controlシートの内容から各テーブル毎のシートを読みに行く
    foreach ($tableNames as $sheetName => $tableName) {
        $sheet = $objPHPExcel->getSheetByName($sheetName);
        $colnum = 0;
        $columns = array();
        $topColumnIndex = null;
        while ($sheet->getCellByColumnAndRow($colnum, 1)->getValue()) {
            if ($columnName = $sheet->getCellByColumnAndRow($colnum, 2)->getValue()) {
                $topColumnIndex = is_null($topColumnIndex) ? $colnum : $topColumnIndex;
                $columns[$colnum] = $columnName;
            }
            ++$colnum;
        }
        if (!$columns) {
            throw new Exception;
        }

        echo $tableName . "\n";
        
        // mysql用の処理はココらへん
        $truncateSql = 'TRUNCATE TABLE ' . $tableName . ' ';
        $pdo->exec($truncateSql);


        $insertSql  = 'INSERT INTO ' . $tableName . ' ';
        $insertSql .= '(' . implode(',', $columns) . ') ';
        $insertSql .= 'VALUES ';
        $insertSql .= '(' . implode(',', array_fill(0, count($columns), '?')) . ') ';
        // echo $insertSql . "\n";

        $rownum = 3;
        $rows = array();
        while ($sheet->getCellByColumnAndRow($topColumnIndex, $rownum)->getValue()) {
            $row = array();
            foreach(array_keys($columns) as $colIndex) {
                $tmp = mb_convert_encoding($sheet->getCellByColumnAndRow($colIndex, $rownum)->getValue(), "UTF8");
                $row[] = $tmp === "" ? null : $tmp;
            }
            $rows[] = $row;
            $rownum++;
        }
        // print_r($params);
        // echo "\n\n";

        $stmt = $pdo->prepare($insertSql);
        foreach ($rows as $row) {
            $stmt->execute($row);
        }

        $descSql = 'DESC ' . $tableName . ' ';
        $stmt = $pdo->prepare($descSql);
        $stmt->execute();
        $descResult = $stmt->fetchAll();

        // mysqlのテーブルスキーマからsqliteのテーブル作るCreate文つくるよ
        $primaryKeys = array();
        $createSql  = 'CREATE TABLE ' . $tableName . ' ( ';
        foreach ($descResult as $colummData) {
            $createSql .= '"' . $colummData['Field'] . '" ';
            if (preg_match('/int/', $colummData['Type']) || preg_match('/bit/', $colummData['Type'])) {
                $createSql .= 'INTEGER ';
            } elseif (preg_match('/float/', $colummData['Type']) || preg_match('/double/', $colummData['Type']) || preg_match('/decimal/', $colummData['Type'])) {
                $createSql .= 'REAL ';
            } elseif (preg_match('/blob/', $colummData['Type']) || preg_match('/binary/', $colummData['Type'])) {
                $createSql .= 'BLOB ';
            } else {
                $createSql .= 'TEXT ';
            }
            if ($colummData['Null'] == 'NO') {
                $createSql .= 'NOT NULL ';
            }
            $createSql .= ',';

            if (preg_match('/PRI/', $colummData['Key'])) {
                $primaryKeys[] = $colummData['Field'];
            }
        }
        $createSql .= ' PRIMARY KEY ("' . implode('","',$primaryKeys) . '") ';
        $createSql .= ' ) ';

        try {
            $sqlitePdo->exec('DROP TABLE ' . $tableName);
        } catch (Exception $ex) {
        }
        $sqlitePdo->exec($createSql);

        $stmt = $sqlitePdo->prepare($insertSql);
        foreach ($rows as $row) {
            $stmt->execute($row);
        }
        // echo $createSql;
        // echo "\n\n";
    }
}

