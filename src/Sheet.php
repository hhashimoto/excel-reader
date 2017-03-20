<?php
namespace hhashimoto\excel;

/**
 * xl/workbook.xml
 *   - workbook
 *     - sheets
 *       - sheet
 */
class Sheet {
    private static $book;

    private $name;
    private $sheetId;
    private $refId;

    // Cell
    private $cells = null;

    private $left   = null;
    private $top    = null;
    private $right  = null;
    private $bottom = null;

    static function belongsTo(Book $book) {
        self::$book = $book;
    }

    /**
     * Construct
     *
     * @param (string | SimpleXMLElement) $name
     * @param (string | SimpleXMLElement) $sheetId
     * @param (string | SimpleXMLElement) $refId
     */
    function __construct($name, $sheetId, $refId) {
        // '(string)$var' convert SimpleXMLElement to string
        $this->name    = (string)$name;
        $this->sheetId = (string)$sheetId;
        $this->refId   = (string)$refId;
    }

    /**
     * Return this name
     * @return string
     */
    public function name() {
        return $this->name;
    }

    /**
     * Return this id
     * @return string
     */
    public function id() {
        return $this->sheetId;
    }

    /**
     * Return this referenceId
     * @return string
     */
    public function refId() {
        return $this->refId;
    }

    /**
     * Unload all cell
     */
    public function unload() {
        $this->cells = null;
    }

    private function load() {
        $zip = new \ZipArchive;
        if (! $zip->open(self::$book->name())) {
            throw new \Exception("'{$name}' could not open!");
        }

        try {
            $sheetNum = substr($this->refId, strlen('rId'));
            $sheet = $zip->getFromName('xl/worksheets/sheet' . $sheetNum . '.xml');
            $zip->close();
            $zip = null;
    
            $xml = new \SimpleXMLElement($sheet);
            $sheet = null;
    
            if (! $this->left) {
                // ex) ref="A6:K42"
                preg_match('/^([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)$/', $xml->dimension['ref'], $matches);
                list($_, $this->left, $this->top, $this->right, $this->bottom) = $matches;
            }
    
            $cells = [];
            foreach ($xml->sheetData->row as $row) {
                $cols = [];
                foreach ($row->c as $c) {
                    $val = '';
                    if ($c->v) {
                        if ($c['t'] && $c['t'] == 's') {
                            $val = self::$book->getSharedString($c->v);
                        } else {
                            $val = $c->v;
                        }
                    }
                    $cols[] = new Cell($val);
                }
                $cells[] = $cols;
            }
            $this->cells = $cells;
        } catch (\Exception $e) {
            echo __FILE__ . ' : ' . __LINE__ . ' ' . $e . PHP_EOL;

            $this->cells = null;
        } finally {
            if ($zip) {
                $zip->close();
                $zip = null;
            }
        }
    }

    private function cells() {
        if (! $this->cells) {
            $this->load();
        }
        return $this->cells;
    }

    /**
     * @param string $pos
     * @return Cell
     */
    public function getCell($pos) {
        $cells = $this->cells();

        preg_match('/^(\w+)(\d+)$/', $pos, $matches);
        list($_, $col, $row) = $matches;

        $x = ord($col) - ord($this->left);
        $y = $row - $this->top;

        if (array_key_exists($y, $cells) &&
            array_key_exists($x, $cells[$y])) {
            return $cells[$y][$x];
        }
        return new Cell;
    }
}
