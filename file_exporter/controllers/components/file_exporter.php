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

/**
 * ファイルダウンロードのためにリクエストをコントロール
 * 
 * @author IKEDA Youhei <youhey.ikeda@gmail.com>
 */
class FileExporterComponent extends Object {

    /** ファイル拡張子 */
    const 
        EXTENSION_CSV   = 'csv', 
        EXTENSION_EXCEL = 'xls';

    /** コンテントタイプのマッピング */
    const 
        CONTENTTYPE_MAPPINGS_PLAINTEXT = 'text/plain', 
        CONTENTTYPE_MAPPINGS_MSEXCEL   = 'application/vnd.ms-excel';

    /** コンテンツのキャッシュ有効時間 */
    const TIMELIMIT_EXPIRES = 1800; // 60sec * 30min

    /** 日付フォーマット */
    const 
        HTTP_DATE_FORMAT = 'D, d M Y H:i:s \G\M\T';

    /** HTTPレスポンスヘッダ */
    const 
        HTTP_LAST_MODIFIED = 'Last-Modified: %1$s', 
        HTTP_EXPIRES       = 'Expires: %1$s', 
        HTTP_CACHE_CONTROL = 'Cache-Control: public', 
        HTTP_PRAGMA        = 'Pragma: public';

    /**
     * 使用するコンポーネント
     *
     * @var array
     */
    public $components = array('RequestHandler');

    /**
     * コンポーネントを初期化
     *
     * @param  Controller $controller コントローラ
     * @return void
     * @link   http://book.cakephp.org/ja/view/64/Creating-Components
     */
    public function initialize(Controller $controller) {
        if ($this->isExcelExtension()) {
            $mapping = array(
                    self::CONTENTTYPE_MAPPINGS_MSEXCEL, 
                    self::CONTENTTYPE_MAPPINGS_PLAINTEXT, 
                );
            $this->RequestHandler->setContent(self::EXTENSION_EXCEL, $mapping);
        }
    }

    /**
     * リクエストの拡張子がExcelファイル（*.xls）かをチェック
     * 
     * @return boolean 拡張子がExcelファイルであればTRUE
     */
    public function isExcelExtension() {
        return 
            ($this->RequestHandler->ext === self::EXTENSION_EXCEL);
    }

    /**
     * リクエストの拡張子がCSVファイル（*.csv）かをチェック
     * 
     * @return boolean 拡張子がCSVファイルであればTRUE
     */
    public function isCsvExtension() {
        return 
            ($this->RequestHandler->ext === self::EXTENSION_CSV);
    }

    /**
     * ExcelファイルをダウンロードするためのHTTPレスポンスを出力
     * 
     * @param string $filename ダウンロードするファイル名
     * @return void
     */
    public function respondAsExcel($filename) {
        $this->enableCacheControl(self::TIMELIMIT_EXPIRES);

        $options = array('attachment' => $filename);
        $this->RequestHandler->respondAs(self::EXTENSION_EXCEL, $options);
    }

    /**
     * CSVファイルをダウンロードするためのHTTPレスポンスを出力
     * 
     * @param string $filename ダウンロードするファイル名
     * @return void
     */
    public function respondAsCsv($filename) {
        $this->enableCacheControl(self::TIMELIMIT_EXPIRES);

        $options = array('attachment' => $filename);
        $this->RequestHandler->respondAs(self::EXTENSION_CSV, $options);
    }

    /**
     * IE対応のためにキャッシュを有効にするHTTPレスポンスヘッダ
     * 
     * <p>HTTPS（SSL暗号化通信）経由のファイルダウンロードにおいて、
     * Internet Explorer ではキャッシュコントローラが機能しない。<br />
     * HTTPヘッダでキャッシュの無効を指定しているレスポンスでは、
     * IE+SSLの環境で一部のファイルがダウンロードできない障害です。<br />
     * エラーメッセージ「ダウンロードすることができません。」が発生<br />
     * 上記の対応に対応するため、キャッシュ制御の指定を上書きします。</p>
     * 
     * @param  integer $ttl キャッシュ生存時間
     * @return void
     * @see <a href="http://support.microsoft.com/kb/323308/ja">Microsoftサポート</a>
     * @see <a href="http://support.microsoft.com/kb/316431/ja">この動作は仕様です。</a>
     * @see <a href="http://support.microsoft.com/kb/436605/ja">マイクロソフトの問題認識</a>
     */
    private function enableCacheControl($ttl = 0) {
        $current = null;
        if (isset($_SERVER['REQUEST_TIME'])) {
            $current = $_SERVER['REQUEST_TIME'];
        }

        $timelimit = ($current + $ttl);
        $datetime  = gmdate(self::HTTP_DATE_FORMAT, $timelimit);
        $response  = sprintf(self::HTTP_EXPIRES, $datetime);
        header($response);

        $datetime = gmdate(self::HTTP_DATE_FORMAT, $current);
        $response = sprintf(self::HTTP_LAST_MODIFIED, $datetime);
        header($response);

        header(self::HTTP_CACHE_CONTROL);
        header(self::HTTP_PRAGMA);
    }
}
