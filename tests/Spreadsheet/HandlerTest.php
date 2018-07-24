<?php
namespace Test\Spreadsheet;

use EasySpreadsheets\Handler;
use PHPUnit\Framework\TestCase;

class HandlerTest extends TestCase
{
    protected $file;
    protected $spreadsheet;

    protected function setUp()
    {
        parent::setUp();
        $this->file = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'test.xlsx';
        $this->spreadsheet = new Handler();       
    }
    
    /** 
     * Test the spreadsheet open
     */
    public function testOpen()
    {
        $this->spreadsheet->load($this->file);
        $this->assertTrue(true);
    }
    
    public function testGetHead()
    {
        $this->spreadsheet->load($this->file);
        $head = [
            'Column A',
            'Column B',
            'Column C',
            'Column D',
            'Column E',
        ];
        $this->assertEquals($head, $this->spreadsheet->getHeader());
    }
    
    public function testGetRows()
    {
        $this->spreadsheet->load($this->file);
        $rows = [
            2 => [
                'Column A' => 'Line 2A',
                'Column B' => 'Line 2B',
                'Column C' => 2,
                'Column D' => '26/10/89',
                'Column E' => 'Line 2E'],

            3 => [
                'Column A' => 'Line 3A',
                'Column B' => 'Line 3B',
                'Column C' => 3,
                'Column D' => '26/10/89',
                'Column E' => 'Line 2E'],

            4 => [
                'Column A' => 'Line 4A',
                'Column B' => 'Line 4B',
                'Column C' => 4,
                'Column D' => '26/10/89',
                'Column E' => 'Line 2E'],

        ];
        $this->assertEquals($rows, $this->spreadsheet->getRows());
    }
}