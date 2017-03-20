<?php

/**
 * xl/workbook.xml
 *   - workbook
 *     - sheets
 */
class Sheets implements Countable {
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
     * If the sheet not founded, return null
     * @param string $sheetName
     * @return Sheet | null
     */
    public function find($sheetName) {
        $l = mb_strlen($sheetName);
        foreach ($this->sheets as $sheet) {
            if ($l !== mb_strlen($sheet->name())) continue;

            if (mb_strpos($sheetName, $sheet->name()) === 0) return $sheet;
        }
        return null;
    }

    /**
     * Return number of sheets
     *
     * @return int
     */
    public function count() {
        return count($this->sheets);
    }
}
