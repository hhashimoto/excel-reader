<?php
namespace hhashimoto\excel;

class SharedStrings {
    private $strings;

    public function __construct() {
        $this->strings = [];
    }

    /**
     * @param (string | SimpleXMLElement) $str
     */
    public function add($str) {
        $this->strings[] = (string)$str;
    }

    public function get($index) {
        return $this->strings[$index];
    }
}
