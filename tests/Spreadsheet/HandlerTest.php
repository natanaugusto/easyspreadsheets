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
        $this->assertEquals('FFFFFFFF', $row['Column E']['colors']['font']);
        $this->assertEquals(
            ['start' => "FFED1C24", 'end' => "FF993300"],
            $row['Column E']['colors']['fill']
        );

        $row = $this->spreadsheet->getRowFullInfo(3);
        $this->assertEquals('FFFFFFFF', $row['Column E']['colors']['font']);
        $this->assertEquals(
            ['start' => "FF00A65D", 'end' => "FF008080"],
            $row['Column E']['colors']['fill']
        );

        $row = $this->spreadsheet->getRowFullInfo(4);
        $this->assertEquals('FFFFFFFF', $row['Column E']['colors']['font']);
        $this->assertEquals(
            ['start' => "FF0066B3", 'end' => "FF008080"],
            $row['Column E']['colors']['fill']
        );
    }

    public function testCellWrite()
    {
        $this->spreadsheet->load($this->file);
        $this->spreadsheet->writeCell('A2', 'Writed');
        $this->spreadsheet->save();

        $this->spreadsheet->load($this->file, true);
        $row = $this->spreadsheet->getRow(2);
        $this->assertEquals('Writed', $row['Column A']);

        $this->spreadsheet->writeCell('A2', 'Line 2A');
        $this->spreadsheet->save();
        
        $this->spreadsheet->load($this->file, true);
        $row = $this->spreadsheet->getRow(2);
        $this->assertEquals('Line 2A', $row['Column A']);
    }

    public function testPaintTexts()
    {
        $this->spreadsheet->load($this->file, true);

        $this->setColorsSaveLoad('font', 'A2', 'FF00A65D');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('FF00A65D', $row['Column A']['colors']['font']);

        $this->setColorsSaveLoad('font', 'A2', 'FF000000');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('FF000000', $row['Column A']['colors']['font']);

        $this->setColorsSaveLoad('font', 'A2:B3', 'FF00A65D');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('FF00A65D', $row['Column A']['colors']['font']);
        $this->assertEquals('FF00A65D', $row['Column B']['colors']['font']);
        $row = $this->spreadsheet->getRowFullInfo(3);
        $this->assertEquals('FF00A65D', $row['Column A']['colors']['font']);
        $this->assertEquals('FF00A65D', $row['Column B']['colors']['font']);

        $this->setColorsSaveLoad('font', 'A2:B3', 'FF000000');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('FF000000', $row['Column A']['colors']['font']);
        $this->assertEquals('FF000000', $row['Column B']['colors']['font']);
        $row = $this->spreadsheet->getRowFullInfo(3);
        $this->assertEquals('FF000000', $row['Column A']['colors']['font']);
        $this->assertEquals('FF000000', $row['Column B']['colors']['font']);
        
    }

    public function testPaintFills()
    {
        $this->spreadsheet->load($this->file, true);

        $this->setColorsSaveLoad('fill', 'A2', 'FF00A65D');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals(
            ['start' => 'FF00A65D', 'end' => 'FF00A65D'],
            $row['Column A']['colors']['fill']
        );

        $this->setColorsSaveLoad('fill', 'A2', ['end' => 'FF0066B3']);
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('FF0066B3', $row['Column A']['colors']['fill']['end']);

        $this->setColorsSaveLoad('fill', 'A2', 'FF000000');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals(
            ['start' => 'FF000000', 'end' => 'FF000000'],
            $row['Column A']['colors']['fill']
        );

        $this->setColorsSaveLoad('fill', 'A2:B3', 'FF00A65D');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals(
            ['start' => 'FF00A65D', 'end' => 'FF00A65D'],
            $row['Column A']['colors']['fill']
        );
        $this->assertEquals(
            ['start' => 'FF00A65D', 'end' => 'FF00A65D'],
            $row['Column B']['colors']['fill']
        );
        $row = $this->spreadsheet->getRowFullInfo(3);
        $this->assertEquals(
            ['start' => 'FF00A65D', 'end' => 'FF00A65D'],
            $row['Column A']['colors']['fill']
        );
        $this->assertEquals(
            ['start' => 'FF00A65D', 'end' => 'FF00A65D'],
            $row['Column B']['colors']['fill']
        );

        $this->setColorsSaveLoad('fill', 'A2:B3', 'FF000000');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals(
            ['start' => 'FF000000', 'end' => 'FF000000'],
            $row['Column A']['colors']['fill']
        );
        $this->assertEquals(
            ['start' => 'FF000000', 'end' => 'FF000000'],
            $row['Column B']['colors']['fill']
        );
        $row = $this->spreadsheet->getRowFullInfo(3);
        $this->assertEquals(
            ['start' => 'FF000000', 'end' => 'FF000000'],
            $row['Column A']['colors']['fill']
        );
        $this->assertEquals(
            ['start' => 'FF000000', 'end' => 'FF000000'],
            $row['Column B']['colors']['fill']
        );
    }

    /**
     * Set the new color on fill/font, save, load and try the assertion.
     *
     * @param string $type   Use font/fill
     * @param string $ranges The ranges to be changed
     * @param array  $values Format exemples:
     * 'FF00A65D'
     * ['start' => 'FF00A65D']
     * ['end' => 'FF00A65D']
     * ['start' => 'FF00A65D', 'end' => 'FF00A65D']
     * @return void
     */
    protected function setColorsSaveLoad($type, $ranges, $values)
    {
        // Verify if the ranges and values has the same elements amount
        if(is_array($ranges) && is_array($values)) {
            if(count($ranges) !== count($values)) {
                throw new Exception("\$ranges and \$values must have the same numbers of elements.", 1);
            }
        }
        // Validate the type passed
        switch ($type) {
            case 'fill':
                $method = 'setFillColor';
                break;
            case 'font':
                $method = 'setTextColor';
                break;
            default:
                throw new Exception("Error: the type must be 'font' or 'fill'. {$type} is passed", 2);
                break;
        }
        // Set the values
        switch(gettype($ranges)) {
            case 'array':
                foreach($ranges as $key => $range) {
                    $this->spreadsheet->{$method}($range, $values[$key]);
                }
                break;
            case 'string':
                $this->spreadsheet->{$method}($ranges, $values);                
                break;
        }
        // Save
        $this->spreadsheet->save();
        $this->spreadsheet->load($this->file, true);
    }
}