<?php
namespace hhashimoto\excel;

/**
 * xl/workbook.xml
 *   - workbook
 *     - sheets
 */
class Sheets implements \Countable {
    private $sheets;

    public function __construct() {
        $this->sheets = [];
    }

    /**
     * Add sheet
     *
     * NOTE: $sheet must have a unique name in the book
     * @param Sheet $sheet
     */
    public function add(Sheet $sheet) {
        $this->sheets[] = $sheet;
    }

    /**
     * Return the sheet same name with $sheetName
     *
     * If the sheet not found, return null
     * @param string $sheetName
     * @return Sheet | null
     */
    public function find($sheetName) {
        $l = mb_strlen($sheetName);
        $found = null;
        foreach ($this->sheets as $sheet) {
            if (($l !== mb_strlen($sheet->name())) ||
                (mb_strpos($sheetName, $sheet->name()) !== 0)) {
                $sheet->unload();
                continue;
            }

            $found = $sheet;
        }
        return $found;
    }

    /**
     * Return number of sheets
     *
     * @return int
     */
    public function count() {
        return count($this->sheets);
    }

    /**
     * @return array
     */
    public function sheetNames() {
        $names = [];
        foreach ($this->sheets as $sheet) {
            $names[] = $sheet->name();
        }
        return $names;
    }
}
