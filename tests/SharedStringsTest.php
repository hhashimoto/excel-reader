<?php
namespace hhashimoto\excel\tests;

use hhashimoto\excel\Book;

class SharedStringsTest extends \PHPUnit_Framework_TestCase {
    /**
     * @test
     */
    public function get() {
        $book = new Book();
        $book->load('tests/fixtures/test1.xlsx');

        $this->assertEquals('sheetName', $book->getSharedString(0));
        $this->assertEquals('Sheet3', $book->getSharedString(1));
    }
}
