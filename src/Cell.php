<?php
namespace hhashimoto\excel;

class Cell {
    private $value;

    /**
     * @param (string | SimpleXMLElement) $value
     */
    public function __construct($value = '') {
        $this->value = (string)$value;
    }

    public function getValue() {
        return $this->value;
    }
}
