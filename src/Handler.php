<?php
namespace EasySpreadsheets;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
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
    protected static $OFFSET_LINE = 2;
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
     * Constructor
     */
    public function __construct()
    {
        $this->currentRow += self::$OFFSET_LINE;
    }
    /**
     * Return the current row number @var $currentRow
     * (@var $currentRow init with the @var $OFFSET_LINE value)
     * @return integer
     */
    public function getCurrentRow()
    {
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
     * @return void
     */
    public function getRows()
    {
        return $this->rows;
    }
    /**
     * Retorna a linha atual ou a linha passada por parametro
     *
     * @param integer $line
     * 
     * @return array
     */
    public function getRow($line = null)
    {
        $line = is_null($line) ? $this->currentRow : $line;
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
        $column = 1;
        foreach($row as $index => $value) {
            $row[$index] = array_merge(
                ['value' => $value], 
                $this->getColors($line, $column)
            );
            $column++;
        }
        $this->currentRow++;
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
     * @param string $path
     * @return void
     */
    public function load($path)
    {
        $this->path = $path;
        $this->resource = IOFactory::load($path);
        $this->activesheet = $this->resource->getActiveSheet();
        $this->highestRow = (int)$this->activesheet->getHighestRow();
        $this->highestColumn = $this->activesheet->getHighestColumn();
        $this->loadHeader();
        $this->loadRows();
    }
    /**
     * Save the spreadsheet setted on @var $resource on path @var $path
     *
     * @return void
     */
    public function save()
    {
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx(
            $this->resource
        );
        return $writer->save($this->path);
    }
    /**
     * Verifica se existem linhas a serem processadas
     *
     * @return boolean
     */
    public function hasNext()
    {
        return $this->currentRow < $this->highestRow ? true : false;
    }
    /**
     * Recover the first line of the spreadsheet assumed that's the header of spreadsheet
     * 
     * @return void
     */
    public function loadHeader()
    {
        $header = $this->activesheet->rangeToArray(
            "A1:{$this->highestColumn}1"
        )[0];
        if(empty($header)) {
            throw new \Exception('The first spreadsheet first line is empty');
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
    }
    /**
     * Load the @var $resource rows to @var $rows
     * 
     * @return void
     */
    public function loadRows()
    {
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
            $rows[$this->linesRead + $key] = array_combine($this->header, $row);
        }
        $this->linesRead += count($rows);
        $this->rows = $rows;
    }
    /**
     * Paint a cells range
     *
     * @param string $begin
     * @param string $range
     * @param string $type Types fill or font
     * @return void
     */
    public function paintRange($range, $color, $type = 'fill')
    {
        switch($type) {
            case 'fill':
                $this->activesheet
                    ->getStyle($range)
                    ->getFill()
                    ->getStartColor()
                    ->setARGB($color);
                $this->activesheet
                    ->getStyle($range)
                    ->getFill()
                    ->getEndColor()
                    ->setARGB($color);
                break;
            case 'font':
                $this->activesheet
                    ->getStyle($range)
                    ->getFont()
                    ->getColor()
                    ->setARGB($color);
                break;
        }
    }
    /**
     * Write a text on a cell
     *
     * @param string $cell
     * @param string $text
     * @return void
     */
    public function writeColumn($cell, $text)
    {
        return $this->activesheet->getCell($cell)->setValue($text);
    }
    /**
     * Recupera as cores do backgroud e da fonte de uma celula da 
     * planilha
     * 
     * @param integer $line  A linha que deve ser recuperada. Por 
     * padrão, a linha deve ser incrementada em 2 p/ contornar a 
     * questão da primeira linha ser a header e o array @var $rows
     * começar em 0.
     * @param string  $index O indice que a cor deve ser recuperada
     * 
     * @return array [colors=>[fill => 'cor da fill', font => 'cor da font']]
     */
    public function getColors($line, $index)
    {
        $position = $this->convertPosition($index, $line);
        return [
            'colors' => [
                'font' => $this->getColorName(
                    $this->activesheet
                        ->getStyle($position)
                        ->getFont()
                        ->getColor()
                        ->getARGB()
                    ),
                'fill' => $this->getColorName(
                    $this->activesheet
                        ->getStyle($position)
                        ->getFill()
                        ->getEndColor()
                        ->getARGB()
                    ),
            ]
        ];
    }
    /**
     * Corverte uma cor ARGB em uma nomenclatura variante entre red, green, blue, gray, black, white
     * (Os numeros chumbados foram escolhidos arbitráriamente)
     * 
     * @param string $argb Valor ARGB da cor
     * @return string
     */
    public function getColorName($argb)
    {
        if (strlen($argb) == 8) { //ARGB
            $hex = array($argb[2].$argb[3], $argb[4].$argb[5], $argb[6].$argb[7]);
            $rgb = array_map('hexdec', $hex);
        }
        // Validando as cores de maneira Pepeada
        // Branco, preto, cinza
        if($rgb[0] === $rgb[1] && $rgb[1] === $rgb[2]) {
            $rgb = array_sum($rgb);
            if($rgb > 750) {
                return 'white';
            }
            if($rgb < 180) {
                return 'black';
            }
            return 'gray';
        }
        $pepe = 80;
        // Vermelho
        if($rgb[0] > 150 && $rgb[0] > $rgb[1] && $rgb[0] > $rgb[2]) {
            if(($rgb[0] - $rgb[1]) > $pepe && ($rgb[0] - $rgb[2]) > $pepe) {
                return 'red';
            }
        }
        // Verde
        if($rgb[1] > 150 && $rgb[1] > $rgb[0] && $rgb[1] > $rgb[2]) {
            if(($rgb[1] - $rgb[0]) > $pepe && ($rgb[1] - $rgb[2]) > $pepe) {
                return 'green';
            }
        }
        // Azul
        if($rgb[2] > 150 && $rgb[2] > $rgb[0] && $rgb[2] > $rgb[1]) {
            if(($rgb[2] - $rgb[0]) > $pepe && ($rgb[2] - $rgb[1]) > $pepe) {
                return 'green';
            }
        }
        return 'unkown';
    }
    /**
     * Converte a posição numerica passada por columna e linha para uma 
     * posição no padrão de planilhas A1, A2, B5, etc
     *
     * @param mixed $column String ou Inteiro representando a coluna em questão
     * @param integer $line Linha que deve ser retornada
     * @return void
     */
    public function convertPosition($column, $line)
    {
        return Coordinate::stringFromColumnIndex($column) . $line;
    }
}