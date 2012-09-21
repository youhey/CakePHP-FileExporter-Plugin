<?php
/**
 * ファイル出力
 * 
 * PHP versions >= 5.2
 *
 * CakePHP(tm) : Rapid Development Framework (http://cakephp.org)
 * Copyright 2005-2011, Cake Software Foundation, Inc. (http://cakefoundation.org)
 *
 * Licensed under The MIT License
 * Redistributions of files must retain the above copyright notice.
 * 
 * @since   FileExporter 1.0
 * @license MIT License (http://www.opensource.org/licenses/mit-license.php)
 */

App::import('Helper', 'FileExporter.AbstractFileExporter');
if (!class_exists('PHPExcel')) {
    App::import('Vendor', 'PHPExcel', array('file' => 'PHPExcel/PHPExcel.php'));
}

/**
 * Excel出力の抽象化クラス
 * 
 * <p>UserモデルのデータからExcelデータを作成するサンプルコードです。</p>
 * <code>
 * class TestUserExcelExporterHelper extends AbstractExcelExporterHelper
 * {
 * 
 *     protected
 *         $fileTitle   = 'Excelサンプルデータ', 
 *         $fileSubject = 'Excelサンプルデータ';
 * 
 *     protected $procedureForEveryColumns = array(
 *             'A' => array('process' => 'extractName', 'type' => PHPExcel_Cell_DataType::TYPE_STRING), 
 *             'B' => array('process' => 'extractEmail', 'type' => PHPExcel_Cell_DataType::TYPE_STRING), 
 *             'C' => array('process' => 'extractTelephone', 'type' => PHPExcel_Cell_DataType::TYPE_STRING), 
 *         );
 * 
 *     protected function getTableHeader()
 *     {
 *         $header = array('A1' => '氏名', 'B1' => 'メールアドレス', 'C1' => '電話番号');
 * 
 *         return $header;
 *     }
 * 
 *     protected function extractName($data)
 *     {
 *         $name = null;
 *         if (isset($data['User']['name'])) {
 *             $name = $data['User']['name'];
 *         }
 * 
 *         return $name;
 *     }
 * 
 *     protected function extractEmail($data)
 *     {
 *         $email = null;
 *         if (isset($data['User']['email'])) {
 *             $email = $data['User']['email'];
 *         }
 * 
 *         return $email;
 *     }
 * 
 *     protected function extractTelephone($data)
 *     {
 *         $telephone = null;
 *         if (isset($data['User']['telephone'])) {
 *             $telephone = $data['User']['telephone'];
 *         }
 * 
 *         return $telephone;
 *     }
 * }
 * </code>
 *
 * @author IKEDA Youhei <youhey.ikeda@gmail.com>
 */
abstract class AbstractExcelExporterHelper extends AbstractFileExporterHelper {

    /** 作業シートのINDEX */
    const SHEET_INDEX = 0;

    /** データの流し込みを開始する行 */
    const START_ROW = 2;

    /** 書き込みオブジェクトのタイプ */
    const WRITER_TYPE = 'Excel5';

    /** フォント */
    const 
        BODY_FONT_NAME  = 'ＭＳ Ｐゴシック', 
        BODY_FONT_SIZE  = 9, 
        BODY_FONT_BOLD  = false;

    /** セルのサイズ */
    const 
        COLUMN_DIMENSION  = '12', 
        ROW_DIMENSION     = '12';

    /**
     * ファイルの作成者
     * 
     * @var string
     */
    protected $fileCreator = '';

    /**
     * ファイルのタイトル
     * 
     * @var string
     */
    protected $fileTitle = '';

    /**
     * ファイルの件名
     * 
     * @var string
     */
    protected $fileSubject = '';

    /**
     * ファイルの説明
     * 
     * @var string
     */
    protected $fileDescription = '';

    /**
     * ファイルのキーワード
     * 
     * @var string
     */
    protected $fileKeywords = '';

    /**
     * ファイルのカテゴリ
     * 
     * @var string
     */
    protected $fileCategory = '';

    /**
     *データセルごとの処理定義
     * 
     * @var array
     */
    protected $procedureForEveryColumns = array();

    /**
     * タイトル行データを返却
     * 
     * @return array 安タイトル行
     */
    abstract protected function getTableHeader();

    /**
     * Excelバイナリデータを作成
     * 
     * <p>作成したExcelのバイナリデータはメモリ上に保持するので注意</p>
     * 
     * @param array  $data データ
     * @return string Excelバイナリデータ
     */
    public function compile($data) {
        $binary = null;

        try {
            if (!is_array($data)) {
                $message = 'argument data must be array';
                throw new InvalidArgumentException($message);
            }

            $excel = $this->createPHPExcel();
            $sheet = $excel->setActiveSheetIndex(self::SHEET_INDEX);

            foreach ($this->getTableHeader() as $position => $label) {
                $sheet->setCellValue($position, $label);
            }

            $index = 0;
            foreach ($data as $row) {
                $number = (self::START_ROW + $index++);
                foreach ($this->procedureForEveryColumns as $p => $c) {
                    $address = ($p.$number);
                    $this->setExcelCellValue($sheet, $address, $c, $row);
                }
            }
            $binary = $this->writeExcelBinary($excel);
        } catch (Exception $e) {
            $message = 'Unable to compile the Excel binary: '.$e->getMessage();
            $this->log($message, LOG_ERROR);
        }

        return 
            $this->output($binary);
    }

    /**
     * PHPExcelのインスタンスを作成
     * 
     * @return PHPExcel PHPExcelオブジェクト
     */
    private function createPHPExcel() {
        $excel = new PHPExcel();

        $this->bindExcelProperties($excel);
        $this->bindExcelStyleFont($excel);
        $this->bindExcelStyleDimension($excel);

        return $excel;
    }

    /**
     * Excelファイルのプロパティを設定
     * 
     * @param  PHPExcel $excel PHPExcelオブジェクト
     * @return void
     */
    private function bindExcelProperties(PHPExcel $excel) {
        $excel->getProperties()->setCreator($this->fileCreator)
                               ->setLastModifiedBy($this->fileCreator)
                               ->setTitle($this->fileTitle)
                               ->setSubject($this->fileSubject)
                               ->setDescription($this->fileDescription)
                               ->setKeywords($this->fileKeywords)
                               ->setCategory($this->fileCategory);
    }

    /**
     * Excelファイルのフォントを設定
     * 
     * @param  PHPExcel $excel PHPExcelオブジェクト
     * @return void
     */
    private function bindExcelStyleFont(PHPExcel $excel) {
        $font = $excel->getDefaultStyle()->getFont();
        $font->setName(self::BODY_FONT_NAME);
        $font->setSize(self::BODY_FONT_SIZE);
        $font->setBold(self::BODY_FONT_BOLD);
    }

    /**
     * Excelファイルのセルサイズを設定
     * 
     * @param  PHPExcel $excel PHPExcelオブジェクト
     * @return void
     */
    private function bindExcelStyleDimension(PHPExcel $excel) {
        $sheet = $excel->setActiveSheetIndex(self::SHEET_INDEX);
        $sheet->getDefaultColumnDimension()->setWidth(self::COLUMN_DIMENSION);
        $sheet->getDefaultRowDimension()->setRowHeight(self::ROW_DIMENSION);
    }

    /**
     * Excelバイナリデータを作成
     * 
     * @param  PHPExcel $excel PHPExcelオブジェクト
     * @return string Excelバイナリデータ
     */
    private function writeExcelBinary(PHPExcel $excel) {
        $writer = PHPExcel_IOFactory::createWriter($excel, self::WRITER_TYPE);

        ob_start();
        $writer->save('php://output');
        $buffer = ob_get_clean();

        return $buffer;
    }

    /**
     * Excelのセルにデータを書き込む
     * 
     * @param  PHPExcel_Worksheet $sheet 作業中のExcelワークシート
     * @param  string $address データを書きこむワークシートのアドレス
     * @param  array $procedureForColumn セルのデータ処理手順
     * @param  array $data 書きこむデータ
     * @return string Excelバイナリデータ
     * @throws InvalidArgumentException 書きこむアドレスが空だったとき
     * @throws InvalidArgumentException 処理手順のデータ型が不正だったとき
     * @throws InvalidArgumentException 書きこむデータのデータ型が不正だったとき
     * @throws DomainException 処理手順のデータ定義が不正だったとき
     * @throws BadMethodCallException データを処理するメソッドが未定義だったとき
     
     */
    private function setExcelCellValue($sheet, $address, $procedureForColumn, $data) {
        if (empty($address)) {
            $message = 'Excel cell address mut not be empty';
            throw new InvalidArgumentException($message);
        }
        if (!is_array($procedureForColumn)) {
            $message = "procedure for cell must be array in {$address}";
            throw new InvalidArgumentException($message);
        }
        if (!is_array($data)) {
            $message = "row data must be array in {$address}";
            throw new InvalidArgumentException($message);
        }

        if (!isset($procedureForColumn['process'])) {
            $message = "process method name does not exist in {$address}";
            throw new DomainException($message);
        }
        $method = $procedureForColumn['process'];
        if (!method_exists($this,$method)) {
            $message = "Method '{$method}' does not exist in {$address}";
            throw new BadMethodCallException ($message);
        }
        $value = $this->{$method}($data);

        if (!isset($procedureForColumn['type'])) {
            $message = "Excel cell data type does not exist in {$address}";
            throw new DomainException($message);
        }
        $dataType = $procedureForColumn['type'];

        $sheet->setCellValueExplicit($address, $value, $dataType);
    }
}
