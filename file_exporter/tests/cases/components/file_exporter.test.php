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

App::import('Component', 'RequestHandler');
App::import('Component', 'FileExporter.FileExporter');

/**
 * コンポーネントのテストケース
 * 
 * @author IKEDA Youhei <youhey.ikeda@gmail.com>
 */
class FileExporterComponentTest extends CakeTestCase
{

    public function startTest() {
        $this->FileExporter = new FileExporterComponent;
        $this->FileExporter->RequestHandler = new RequestHandlerComponent;
    }
    public function endTest() {
        $this->FileExporter = null;
        ClassRegistry::flush();
    }

    public function test：リクエストURLの拡張子がExcelかをテスト() {
        $this->FileExporter->RequestHandler->ext = 'html';
        $this->assertFalse($this->FileExporter->isExcelExtension(), "拡張子HTMLは偽");

        $this->FileExporter->RequestHandler->ext = 'txt';
        $this->assertFalse($this->FileExporter->isExcelExtension(), "拡張子TXTは偽");

        $this->FileExporter->RequestHandler->ext = 'csv';
        $this->assertFalse($this->FileExporter->isExcelExtension(), "拡張子CSVは偽");

        $this->FileExporter->RequestHandler->ext = 'xls';
        $this->assertTrue($this->FileExporter->isExcelExtension(), "拡張子XLSは真");

        $this->FileExporter->RequestHandler->ext = 'xlsx';
        $this->assertFalse($this->FileExporter->isExcelExtension(), "拡張子XLSXは偽");
    }

    public function test：リクエストURLの拡張子がCSVかをテスト() {
        $this->FileExporter->RequestHandler->ext = 'html';
        $this->assertFalse($this->FileExporter->isCsvExtension(), "拡張子HTMLは偽");

        $this->FileExporter->RequestHandler->ext = 'txt';
        $this->assertFalse($this->FileExporter->isCsvExtension(), "拡張子TXTは偽");

        $this->FileExporter->RequestHandler->ext = 'tsv';
        $this->assertFalse($this->FileExporter->isCsvExtension(), "拡張子TSVは偽");

        $this->FileExporter->RequestHandler->ext = 'xls';
        $this->assertFalse($this->FileExporter->isCsvExtension(), "拡張子XLSは偽");

        $this->FileExporter->RequestHandler->ext = 'csv';
        $this->assertTrue($this->FileExporter->isCsvExtension(), "拡張子CSVは真");
    }
}
