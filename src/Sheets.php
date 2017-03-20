<?php

/**
 * xl/workbook.xml
 *   - workbook
 *     - sheets
 */
class Sheets {
    private $sheets;

    public function __construct() {
        $this->sheets = [];
    }

    /**
     * Add sheet with unique sheet name
     * @param Sheet $sheet
     */
    public function add(Sheet $sheet) {
        $this->sheets[$sheet->name()] = $sheet;
    }

    /**
     * Return the sheet same name with $sheetName
     *
     * If the sheet not founded, return null
     * @param string $sheetName
     * @return Sheet | null
     */
    public function find($sheetName) {
        if (array_key_exists($sheetName, $this->sheets)) {
            return $this->sheets[$sheetName];
        } else {
            return null;
        }
    }
}
