<?php
namespace Test\Spreadsheet;

use EasySpreadsheets\Handler;
use PHPUnit\Framework\TestCase;

class HandlerTest extends TestCase
{
    protected $file;
    protected $spreadsheet;
    protected $header = [
        'Column A',
        'Column B',
        'Column C',
        'Column D',
        'Column E',
    ];
    protected $rows =  [
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
        $this->assertEquals($this->header, $this->spreadsheet->getHeader());
    }
    
    public function testGetRows()
    {
        $this->spreadsheet->load($this->file);
        $this->assertEquals($this->rows, $this->spreadsheet->getRows());
    }

    public function testNavigateBetweenRows()
    {
        $this->spreadsheet->load($this->file);
        $this->assertEquals($this->rows[2], $this->spreadsheet->getRow());
        $this->assertTrue($this->spreadsheet->hasNext());
        $this->assertEquals($this->rows[3], $this->spreadsheet->getRow());
        $this->assertTrue($this->spreadsheet->hasNext());
        $this->assertEquals($this->rows[4], $this->spreadsheet->getRow());
        $this->assertFalse($this->spreadsheet->hasNext());
        $this->assertEquals($this->rows[2], $this->spreadsheet->getRow(2));
    }

    public function testGetColors()
    {
        $this->spreadsheet->load($this->file);
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('white', $row['Column E']['colors']['font']);
        $this->assertEquals('red', $row['Column E']['colors']['fill']);
    }
}