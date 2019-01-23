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

    private $targetColumnsToLoad = null;

    // Cell
    private $cells = null;

    private $minColumn;
    private $minRow;
    private $maxColumn;
    private $maxRow;

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

    /**
     * @param array $targetColumns
     */
    public function setTargetColumnsToLoad(array $targetColumns) {
        if (!empty($targetColumns)) {
            foreach ($targetColumns as $col) {
                assert(preg_match('/^[a-zA-Z]+$/', $col), "'{$col}' is not valid for column name.");
            }
            $this->targetColumnsToLoad = $targetColumns;
        }
    }

    private function loadDimension($xml) {
        list($min, $max) = explode(':', $xml->dimension['ref'][0]);
        if (preg_match('/^([a-zA-Z]+)(\d+)$/', $min, $matches)) {
            list($_, $this->minColumn, $r) = $matches;
            $this->minRow = (int)$r;
        }
        if (preg_match('/^([a-zA-Z]+)(\d+)$/', $max, $matches)) {
            list($_, $this->maxColumn, $r) = $matches;
            $this->maxRow = (int)$r;
        }
    }

    private function loadAllCells($xml) {
        $cells = [];
        foreach ($xml->sheetData->row as $row) {
            $cols = [];
            $y = (string)$row['r'];
            foreach ($row->c as $c) {
                preg_match('/^(\w+?)\d+$/', $c['r'], $matches);
                list($_, $x) = $matches;
                $val = '';
                if ($c->v) {
                    if ($c['t'] && $c['t'] == 's') {
                        $val = self::$book->getSharedString($c->v);
                    } else {
                        $val = $c->v;
                    }
                }
                $cols[$x] = new Cell($val);
            }
            $cells[$y] = $cols;
        }
        return $cells;
    }

    private function loadOnlyTargetColumns($xml) {
        // 指定カラムだけを連想配列として読み出す
        $columns = implode('|', $this->targetColumnsToLoad);

        $cells = [];
        foreach ($xml->sheetData->row as $row) {
            $cols = [];
            $y = (string)$row['r'];
            foreach ($row->c as $c) {
                if (!preg_match('/^(' . $columns . ')\d+$/', $c['r'], $matches)) {
                    continue;
                }
                list($_, $x) = $matches;
                $val = '';
                if ($c->v) {
                    if ($c['t'] && $c['t'] == 's') {
                        $val = self::$book->getSharedString($c->v);
                    } else {
                        $val = $c->v;
                    }
                }
                $cols[$x] = new Cell($val);
            }
            $cells[$y] = $cols;
        }
        return $cells;
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

            $this->loadDimension($xml);

            if ($this->targetColumnsToLoad) {
                $this->cells = $this->loadOnlyTargetColumns($xml);
            } else {
                $this->cells = $this->loadAllCells($xml);
            }
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
        $this->loadCellsIfNeeded();
        return $this->cells;
    }

    private function loadCellsIfNeeded() {
        if (! $this->cells) {
            $this->load();
        }
    }

    /**
     * @param string $pos
     * @return Cell
     */
    public function getCell($pos) {
        $cells = $this->cells();

        preg_match('/^(\w+?)(\d+)$/', $pos, $matches);
        list($_, $col, $row) = $matches;

        if (array_key_exists($row, $cells) &&
            array_key_exists($col, $cells[$row])) {
            return $cells[$row][$col];
        }
        return new Cell;
    }

    /**
     * @return int
     */
    public function maxRow() {
        $this->loadCellsIfNeeded();
        return $this->maxRow;
    }

    /**
     * @return string
     */
    public function maxColumn() {
        $this->loadCellsIfNeeded();
        return $this->maxColumn;
    }
}
