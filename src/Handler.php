<?php
namespace EasySpreadsheets;
use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
class Handler
{
    /**
     * Maximum lines that must be loaded at a time.
     *
     * @var integer $MAX_LINES_LOAD
     */
    protected static $MAX_LINES_LOAD = 5000;
    /**
     * Maximum number of lines to be loaded at one time. * First line to be read.
     * (Line one, usually the one in the header)
     *
     * @var integer
     */
    protected static $OFFSET_LINE = 1;
    /**
     * Object PHPSpreadsheet
     *
     * @var PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected $resource;
    /**
     * Object PHPSpreadsheet
     *
     * @var PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    protected $activesheet;
    /**
     * Path to spreadsheet
     *
     * @va
     */
    protected $path;
    /**
     * Object Coordinate
     *
     * @var PhpOffice\PhpSpreadsheet\Cell\Coordinate
     */
    protected $coordinate;
    /**
     * Last spreadsheet column used
     *
     * @var string
     */
    protected $highestColumn;
    /**
     * Last spreadsheet line used
     *
     * @var integer
     */
    protected $highestRow;
    /**
     * Lines readed number
     *
     * @var integer
     */
    protected $linesRead = 0;
    /**
     * Current row on memory
     *
     * @var integer
     */
    protected $currentRow;
    /**
     * Primeira linha da planilha Header
     *
     * @var array
     */
    protected $header = [];
    /**
     * Rows array
     *
     * @var array
     */
    protected $rows = [];
    /**
     * Return the current row number @var $currentRow
     * (@var $currentRow init with the @var $OFFSET_LINE value)
     * @return integer
     */
    public function getCurrentRow()
    {
        if(is_null($this->currentRow)) {
            $this->currentRow += self::$OFFSET_LINE;
        }
        return $this->currentRow;
    }
    /**
     * Return the highest row @var $highestRow
     *
     * @return integer
     */
    public function getHighestRow()
    {
        return $this->highestRow;
    }
    /**
     * Return the highest column @var $highestColumn
     *
     * @return string
     */
    public function getHighestColumn()
    {
        return $this->highestColumn;
    }
    /**
     * Return the laoded row
     *
     * @return integer
     */
    public function getRows()
    {
        return $this->rows;
    }
    /**
     * Retorna a linha atual ou a linha passada por parametro
     * (Some mistakes was made down there)
     * @param integer $line
     * @throws Exception
     * @return array
     */
    public function getRow($line = null)
    {
        $line = is_null($line) ? $this->getCurrentRow() : $line;
        if(!is_integer($line)) {
            throw new Exception("The line is not a integer", 3);
        }
        $row = null;
        do {
            if(empty($this->rows[$line])) {
                $this->loadRows();
                continue;
            }
            $row = $this->rows[$line];
            break;
        } while($this->linesRead < $this->highestRow);
        
        if(is_null($row)) {
            return [];
        }        
        $this->currentRow++;
        return $row;
    }
    /**
     * Return row with colors info
     *
     * @param integer $line
     * @return array
     */
    public function getRowFullInfo($line = null)
    {
        $row = $this->getRow($line);
        $column = 1;
        foreach($row as $index => $value) {
            $row[$index] = array_merge(
                ['value' => $value], 
                $this->getColors($line, $column)
            );
            $column++;
        }
        return $row;
    }
    /**
     * Return the path to the current file on memory @var $path
     *
     * @return string
     */
    public function getPath()
    {
        return $this->path;
    }
    /**
     * Return the header @var $header
     *
     * @return void
     */
    public function getHeader()
    {
        return $this->header;
    }
    /**
     * Load the spreadshee resource
     *
     * @param string  $path
     * @param boolean $header
     * @param boolean $force
     * @throws Exception
     * @return Handler
     */
    public function load($path, $header = true, $force = false)
    {
        if(!file_exists($path)) {
            throw new Exception("The file {$path} not found.", 1);
        }
        $this->path = $path;
        $this->resource = IOFactory::load($path);
        $this->activesheet = $this->resource->getActiveSheet();
        $this->highestRow = (int)$this->activesheet->getHighestRow();
        $this->highestColumn = $this->activesheet->getHighestColumn();
        if($force) {
            $this->header = [];
            $this->rows = [];
            $this->currentRow = self::$OFFSET_LINE;
            $this->linesRead = 0;
        }
        if($header === true) {
            self::$OFFSET_LINE = 2;
            $this->loadHeader();
        } else {
            self::$OFFSET_LINE = 1;
        }
        $this->loadRows();
        return $this;
    }
    /**
     * Save the spreadsheet setted on @var $resource on path @var $path
     *
     * @return Handler
     */
    public function save()
    {
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx(
            $this->resource
        );
        $writer->save($this->path);
        return $this;
    }
    /**
     *Verify if exists more rows to read
     *
     * @return boolean
     */
    public function hasNext()
    {
        return $this->currentRow <= $this->highestRow ? true : false;
    }
    /**
     * Recover the first line of the spreadsheet assumed that's the header of spreadsheet
     * 
     * @throws Exception
     * @return Handler
     */
    public function loadHeader()
    {
        $header = $this->activesheet->rangeToArray(
            "A1:{$this->highestColumn}1"
        )[0];
        if(empty($header)) {
            throw new Exception('The first spreadsheet first line is empty', 2);
        }
        
        foreach($header as $k => $attr) {
            if(is_null($attr)) {
                $this->highestColumn = Coordinate::stringFromColumnIndex(
                    count($this->header)
                );
                return;
            }
            $attr = trim($attr);
            $this->header[] = $attr;
        }

        return $this;
    }
    /**
     * Load the @var $resource rows to @var $rows
     * 
     * @return void
     */
    public function loadRows()
    {
        $noHeader = empty($this->header);
        if($this->linesRead === 0) {
            $this->linesRead += self::$OFFSET_LINE;
        }
        $limit = $this->highestRow;
        if($limit > self::$MAX_LINES_LOAD) {
            $linesToRead = $this->highestRow - $this->linesRead;
            if($linesToRead > self::$MAX_LINES_LOAD) {
                $limit = $this->linesRead + self::$MAX_LINES_LOAD;
            }
        }
        $rangeToRead = "A{$this->linesRead}:{$this->highestColumn}{$limit}";
        $result = $this->activesheet->rangeToArray($rangeToRead);
        $rows = [];
        foreach($result as $key => $row) {
            $row = $noHeader ?
                $this->removeNullCells($row) : array_combine($this->header, $row);
                    
            array_walk($row, function (&$val) {
                $val = trim($val);
            });
            
            $rows[$this->linesRead + $key] = $row;
        }
        $this->linesRead += count($rows);
        $this->rows = $rows;
    }
    /**
     * Paint the text of a cells range
     *
     * @param string $rang
     * @param string $color
     * @return Handler
     */
    public function setTextColor($range, $color)
    {
        $this->activesheet
            ->getStyle($range)
            ->getFont()
            ->getColor()
            ->setARGB($color);
        return $this;
    }
    /**
     * Set a background color on a fill
     *
     * @param string $range
     * @param mixed  $color
     * @return Handler
     */
    public function setFillColor($range, $color)
    {
        if(is_array($color)) {
            if(is_assoc($color)) {
                if(!empty($color['start'])) {
                    $colorStart = $color['start'];
                }
                if(!empty($color['end'])) {
                    $colorEnd = $color['end'];
                }
            }
        } else {
            $colorStart = $colorEnd = $color;
        }
        if(!empty($colorStart)) {
            $this->activesheet
                ->getStyle($range)
                ->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getStartColor()
                ->setARGB($colorStart);
        }        
        if(!empty($colorEnd)) {
            $this->activesheet
                ->getStyle($range)
                ->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getEndColor()
                ->setARGB($colorEnd);
        }
        return $this;       
    }
    /**
     * Write a text on a cell
     *
     * @param string $cell
     * @param string $text
     * @return Handler
     */
    public function writeCell($cell, $text)
    {
        $this->activesheet->getCell($cell)->setValue($text);
        return $this;
    }
    /**
     * Recover background and font colors from a worksheet cell
     * 
     * @param integer $line  The line that must be retrieved. Per
     * default, the line must be incremented by 2 to bypass the
     * the first line question is the header and array @var $ rows
     * start at 0.
     * @param string  $index The index that the color should be retrieved
     * 
     * @return array [colors=>[fill => 'cor da fill', font => 'cor da font']]
     */
    public function getColors($line, $index)
    {
        $position = $this->convertPosition($index, $line);
        return [
            'colors' => [
                'font' =>
                    $this->activesheet
                        ->getStyle($position)
                        ->getFont()
                        ->getColor()
                        ->getARGB(),
                'fill' =>
                    [
                        'start' => $this->activesheet
                                    ->getStyle($position)
                                    ->getFill()
                                    ->getStartColor()
                                    ->getARGB(),
                        'end' => $this->activesheet
                                    ->getStyle($position)
                                    ->getFill()
                                    ->getEndColor()
                                    ->getARGB(),
                    ],
            ]
        ];
    }
    /**
     * Converts the numerical position passed by column and line to a
     * position in spreadsheet pattern A1, A2, B5, etc.
     *
     * @param mixed $column String or Integer representing the column in question
     * @param integer $line Line to be returned
     * @return void
     */
    public function convertPosition($column, $line)
    {
        return Coordinate::stringFromColumnIndex($column) . $line;
    }
    /**
     * Remove all empty cells.
     * That's a hack. And is not beautiful
     * I'm sorry
     *
     * @param array $row
     * @return array
     */
    protected function removeNullCells(array $row)
    {
        $column = count($row) - 1;
        while(is_null($row[$column])) {
            unset($row[$column]);
            --$column;
        }
        return $row;
    }
}