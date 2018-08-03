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
            'Line 2A',
            'Line 2B',
            2,
            '26/10/89',
            'Line 2E'],

        3 => [
            'Line 3A',
            'Line 3B',
             3,
            '26/10/89',
            'Line 2E'],

        4 => [
            'Line 4A',
            'Line 4B',
            4,
            '26/10/89',
            'Line 2E'],

    ];

    protected function setUp()
    {
        parent::setUp();
        $this->file = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'test.xlsx';
        $this->spreadsheet = new Handler();       
    }
    
    public function testLoad()
    {
        $this->spreadsheet->load($this->file);
        $this->assertEquals(Handler::class, get_class($this->spreadsheet));
    }
    
    public function testLoadWithoutHeader()
    {
        $this->spreadsheet->load($this->file, false);
        $this->assertEquals(Handler::class, get_class($this->spreadsheet));

        $this->assertEquals($this->getRows(true), $this->spreadsheet->getRows());
    }
    
    public function testGetHead()
    {
        $this->spreadsheet->load($this->file);
        $this->assertEquals($this->header, $this->spreadsheet->getHeader());
    }
    
    public function testGetRows()
    {
        $this->spreadsheet->load($this->file);
        $this->assertEquals($this->getRows(), $this->spreadsheet->getRows());
    }

    public function testNavigateBetweenRows()
    {
        $this->spreadsheet->load($this->file);
        $this->assertEquals($this->getRows()[2], $this->spreadsheet->getRow());
        $this->assertTrue($this->spreadsheet->hasNext());
        $this->assertEquals($this->getRows()[3], $this->spreadsheet->getRow());
        $this->assertTrue($this->spreadsheet->hasNext());
        $this->assertEquals($this->getRows()[4], $this->spreadsheet->getRow());
        $this->assertFalse($this->spreadsheet->hasNext());
        $this->assertEquals($this->getRows()[2], $this->spreadsheet->getRow(2));
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
    }

    public function testPaintTexts()
    {
        $this->spreadsheet->load($this->file, true);

        $this->setColorsSaveLoad('font', 'A2', 'FF00A65D');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('FF00A65D', $row['Column A']['colors']['font']);
    
        $this->setColorsSaveLoad('font', 'A2:B3', 'FF00A65D');
        $row = $this->spreadsheet->getRowFullInfo(2);
        $this->assertEquals('FF00A65D', $row['Column A']['colors']['font']);
        $this->assertEquals('FF00A65D', $row['Column B']['colors']['font']);
        $row = $this->spreadsheet->getRowFullInfo(3);
        $this->assertEquals('FF00A65D', $row['Column A']['colors']['font']);
        $this->assertEquals('FF00A65D', $row['Column B']['colors']['font']);
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
    }

    public function testLoadException()
    {
        $this->expectException(\Exception::class);
        $this->spreadsheet->load('file/not/found');
    }

    public function testGetRowsException()
    {
        $this->spreadsheet->load($this->file);
        $this->expectException(\Exception::class);
        $this->spreadsheet->getRow('b');
    }

    public function tearDown()
    {
        $this->spreadsheet->load($this->file);
        $this->spreadsheet->writeCell('A2', 'Line 2A');
        $this->setColorsSaveLoad('font', 'A2:B3', 'FF000000');
        $this->setColorsSaveLoad('fill', 'A2:B3', 'FF000000');
    }

    /**
     * Return the preconfigured rows
     *
     * @param boolean $noheader
     * @return array
     */
    protected function getRows($noheader = false)
    {
        if($noheader) {
            return $this->rows;
        }
        $rows = $this->rows;
        array_walk($rows, function (&$item) {
            $item = array_combine($this->header, $item);
        });
        return $rows;
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