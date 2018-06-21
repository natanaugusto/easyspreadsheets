<?php
namespace Test\Spreadsheet;

use EasySpreadsheet\Handler;
use PHPUnit\Framework\TestCase;

class HandlerTest extends TestCase
{
    protected $spreadsheet;

    protected function setUp()
    {
        parent::setUp();
        $this->spreadsheet = new Handler();       
    }
    
    /** 
     * Test the spreadsheet open
     */
    public function testOpenSpreadsheet()
    {
        $file = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'test.xlsx';
        $this->spreadsheet->load($file);
        $this->assertTrue(true);
    }
        
}