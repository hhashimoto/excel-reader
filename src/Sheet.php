<?php

/**
 * xl/workbook.xml
 *   - workbook
 *     - sheets
 *       - sheet
 */
class Sheet {
    private $name;
    private $sheetId;
    private $refId;

    public function __construct($name, $sheetId, $refId) {
        $this->name    = $name;
        $this->sheetId = $sheetId;
        $this->refId   = $refId;
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
}
