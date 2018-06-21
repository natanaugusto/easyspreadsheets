<?php
namespace EasySpreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
class Handler
{
    /**
     * Máximo de linhas que devem ser carregadas por vez.
     *
     * @var integer $MAX_LINES_LOAD
     */
    protected static $MAX_LINES_LOAD = 5000;
    /**
     * Primeira linha a ser lida.
     * (A linha um, normalmente é a do cabeçalho)
     *
     * @var integer
     */
    protected static $OFFSET_LINE = 2;
    /**
     * Objeto PHPSpreadsheet
     *
     * @var PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
     */
    protected $spreadsheet;
    /**
     * Objeto PHPSpreadsheet
     *
     * @var PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $activesheet
     */
    protected $activesheet;
    /**
     * Path da planilha
     *
     * @var string
     */
    protected $spreadsheetPath;
    /**
     * Obejeto Coordinate
     *
     * @var PhpOffice\PhpSpreadsheet\Cell\Coordinate $coordinate
     */
    protected $coordinate;
    /**
     * Ultima coluna usada na planilha
     *
     * @var string $highestColumn
     */
    protected $highestColumn;
    /**
     * Ultima linha usada
     *
     * @var integer $highestRow
     */
    protected $highestRow;
    /**
     * Quantidade de linhas que já foram lidas e armazenadas na
     * memoria.
     *
     * @var integer $linesRead
     */
    protected $linesRead = 0;
    /**
     * A linha atual que está sendo analisada
     *
     * @var integer $currentRow
     */
    protected $currentRow;
    /**
     * Primeira linha da planilha Header
     *
     * @var array $headerRow
     */
    protected $headerRow = [];
    /**
     * Array com a estrutura que facilite o processamento e
     * envio para a API da PLuggTO
     *
     * @var array $rows
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
     * Retorna a linha atual
     *
     * @return integer
     */
    public function getCurrentLine()
    {
        return $this->currentRow;
    }
    /**
     * Recupera a ultima linha utilizada
     *
     * @return integer
     */
    public function getHighestRow()
    {
        return $this->highestRow;
    }
    /**
     * Recupera a ultima coluna utilizada
     *
     * @return string
     */
    public function getHighestColumn()
    {
        return $this->highestColumn;
    }
    /**
     * Retorna o path da planilha
     *
     * @return string
     */
    public function getSpreadsheetPath()
    {
        return $this->spreadsheetPath;
    }
    /**
     * Carrega a planilha p/ a classe
     *
     * @param string $spreadsheetPath
     * @return void
     */
    public function loadSpreadsheet($spreadsheetPath)
    {
        $this->spreadsheetPath = $spreadsheetPath;
        $this->spreadsheet = IOFactory::load($spreadsheetPath);
        $this->activesheet = $this->spreadsheet->getActiveSheet();
        $this->highestRow = (int)$this->activesheet->getHighestRow() + self::$OFFSET_LINE;
        $this->highestColumn = $this->activesheet->getHighestColumn();
        $this->loadSpreadsheetHeader();
        $this->loadSpreadsheetRows();
    }
    /**
     * Salva a planilha
     *
     * @return void
     */
    public function saveSpreadsheet()
    {
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx(
            $this->spreadsheet
        );
        return $writer->save($this->spreadsheetPath);
    }
    /**
     * Retorna a linha atual ou a linha passada por parametro
     *
     * @param integer $line
     * 
     * @return array
     */
    public function getSpreadsheetRow($line = null)
    {
        $line = is_null($line) ? $this->currentRow : $line;
        $row = null;
        do {
            if(empty($this->rows[$line])) {
                $this->loadSpreadsheetRows();
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
     * Verifica se existem linhas a serem processadas
     *
     * @return boolean
     */
    public function hasNext()
    {
        return $this->currentRow < $this->highestRow ? true : false;
    }
    /**
     * Recupera as colunas que fazem parte da Header do documento. Os nomes 
     * das colunas, são as propriedades do objeto json que será enviado a
     * API.
     * 
     * @return void
     */
    public function loadSpreadsheetHeader()
    {
        $headerRow = $this->activesheet->rangeToArray(
            "A1:{$this->highestColumn}1"
        )[0];
        if(empty($headerRow)) {
            throw new \Exception('The first spreadsheet first line is empty');
        }
        
        foreach($headerRow as $k => $attr) {
            if(is_null($attr)) {
                $this->highestColumn = Coordinate::stringFromColumnIndex(
                    count($this->headerRow)
                );
                return;
            }
            $attr = strtolower(
                preg_replace(
                    ['/ \(\*\)/', '/ /'],
                    ['', '_'],
                    trim($attr)
                )
            );
            $this->headerRow[] = $attr;
        }
    }
    /**
     * Carrega as linhas da planilha p/ um array
     * 
     * @return void
     */
    public function loadSpreadsheetRows()
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
            $rows[$this->linesRead + $key] = array_combine($this->headerRow, $row);
        }
        $this->linesRead += count($rows);
        $this->rows = $rows;
    }
    /**
     * Pinta um range de celulas
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
     * Escreve um texto numa celula
     *
     * @param [type] $cell
     * @param [type] $text
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