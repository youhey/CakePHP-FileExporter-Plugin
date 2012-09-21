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

/**
 * CSVファイルを出力するヘルパ
 * 
 * @author IKEDA Youhei <youhey.ikeda@gmail.com>
 * @see <a href="http://blog.asial.co.jp/721">PHPでのCSV出力について</a>
 * @see <a href="http://bakery.cakephp.org/articles/ifunk/2007/09/10/csv-helper-php5">CSV Helper</a>
 */
class CsvExporterHelper extends AbstractFileExporterHelper {

    /** サーバのlocale */
    const LOCALE = 'ja_JP.utf8';

    /** CSVの設定 */
    const DELIMITER = ',', ENCLOSURE = '"'; 

    /** 改行コード */
    const LF = "\n", CR = "\r", LFCR = "\n\r", CRLF = "\r\n";

    /** シフトJISエンコーディング */
    const SJIS_ENCODING = 'SJIS-win';

    /**
     * １次元の連想配列から、CSVデータの行を生成
     * 
     * @param array $data 災害安否データ
     * @return string Excelバイナリデータ
     */
    public function row($data) {
        $csv = '';
        try {
            $csv = $this->writeCsvRow($data);
        } catch (Exception $e) {
            $this->log($e->getMessage(), LOG_ERROR);
        }

        return 
            $this->output($csv);
    }

    /**
     * CSVフォーマットの文字列を、Windows系のソフトウェアに最適化
     * 
     * <p>改行コードを「CRLF」に変換する</p>
     * <p>指定があれば文字コードを変換する（SJISに変換するための対応）</p>
     * 
     * @param string $value CSVフォーマットの文字列
     * @param string $encoding 変換する文字コード、NULLなら変換なし
     * @return string Windows向きなCSVフォーマットに変換した文字列
     */
    public function windows($value, $encoding = self::SJIS_ENCODING) {
        try {
            $windows = $this->toCrLf($value);
            if ($encoding !== null) {
                $windows = $this->toEncoding($windows, $encoding);
            }
        } catch (Exception $e) {
            $this->log($e->getMessage(), LOG_ERROR);
        }

        return 
            $this->output($windows);
    }

    /**
     * CSVデータの行を作成
     * 
     * <p>バイナリデータをメモリ上に保持するので注意</p>
     * 
     * @param array $row データ
     * @return string CSVデータの行
     * @throws InvalidArgumentException 行データの型が配列ではないとき
     * @throws RuntimeException 書き込むストリームをオープンできないとき
     * @throws RuntimeException ロケール情報の設定に失敗したとき
     * @throws RuntimeException CSV形式のフォーマットに失敗したとき
     * @throws RuntimeException ストリームリソースの操作に失敗したとき
     * @see <a href="http://jp.php.net/manual/ja/function.fputcsv.php">fputcsv</a>
     */
    private function writeCsvRow($row) {
        if (!is_array($row)) {
            throw new InvalidArgumentException('CSV record row must be array');
        }

        $buffer = fopen('php://memory', 'r+');
        if ($buffer === false) {
            $message = 'failed to open stream: cannot access memory';
            throw new RuntimeException($message);
        }

        try {
            $result = setlocale(LC_ALL, self::LOCALE);
            if ($result === false) {
                $message = 'failed to set locale: '.self::LOCALE;
                throw new RuntimeException($message);
            }

            $result = fputcsv($buffer, $row, self::DELIMITER, self::ENCLOSURE);
            if ($result === false) {
                $message = 'failed to put csv: cannot write CSV record row';
                throw new RuntimeException($message);
            }

            $result = rewind($buffer);
            if ($result === false) {
                $message = 'failed to rewind stream: cannot access buffering';
                throw new RuntimeException($message);
            }
            $csv = stream_get_contents($buffer);
        } catch (Exception $e) {
            fclose($buffer);
            throw $e;
        }
        fclose($buffer);

        return $csv;
    }

    /**
     * 文字列の改行コード「LF」を「CRLF」に変換
     * 
     * @param string $value 改行コードを変換する文字列
     * @return string 改行コードを「CRLF」に変換した文字列
     */
    private function toCrLf($value) {
        $buf = str_replace($this->getNewlineCharacters(), self::LF, $value);

        return 
            str_replace(self::LF, self::CRLF, $buf);
    }

    /**
     * 文字コードを変換
     * 
     * @param string $value 文字コードを変換する文字列
     * @param string $encoding 変換する文字コード
     * @return string 文字コードを変換した文字列
     * @throws LogicException 文字コードの変換に失敗したとき
     */
    private function toEncoding($value, $encoding) {
        $internal = Configure::read('App.encoding');
        $result   = mb_convert_variables($encoding, $internal, $value);
        if ($result === false) {
            $message = "failed to convert encoding: invalid encoding: " 
                     . "from={$internal} to={$encoding}";
            throw new LogicException($message);
        }

        return $value;
    }

    /**
     * 改行コードの配列を返却
     * 
     * @return array 改行コードの配列
     */
    private function getNewlineCharacters() {
        $newlineCharacters = array(
                self::LFCR, 
                self::CRLF, 
                self::LF, 
                self::CR, 
            );

        return $newlineCharacters;
    }
}
