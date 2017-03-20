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

    /**
     * Construct
     *
     * @param (string | SimpleXMLElement) $name
     * @param (string | SimpleXMLElement) $sheetId
     * @param (string | SimpleXMLElement) $refId
     */
    public function __construct($name, $sheetId, $refId) {
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
}
