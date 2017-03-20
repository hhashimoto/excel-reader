<?php

class SheetTest extends \PHPUnit_Framework_TestCase {
    /**
     * @test
     */
    public function getAccessorsReturnsSameAsConstructorArgs() {
        $sut = new Sheet('testSheet', '1', 'rId1');

        $this->assertSame('testSheet', $sut->name());
        $this->assertSame('1',         $sut->id());
        $this->assertSame('rId1',      $sut->refId());
    }

    /**
     * @test
     */
    public function getCellWhenImmediateValue() {
        $book = new Book;
        $book->load('tests/fixtures/test1.xlsx');

        $sut = $book->getSheetByName('Sheet1');

        $expected = new Cell('10');
        $actual = $sut->getCell('A1');
        $this->assertEquals($expected, $actual);

        $expected = new Cell('20');
        $actual = $sut->getCell('B1');
        $this->assertEquals($expected, $actual);
    }

    /**
     * @test
     */
    public function getCellWhenSharedStrings() {
        $book = new Book;
        $book->load('tests/fixtures/test1.xlsx');

        $sut = $book->getSheetByName('Sheet2');

        $expected = new Cell('sheetName');
        $actual = $sut->getCell('B2');
        $this->assertEquals($expected, $actual);

        $expected = new Cell('Sheet2');
        $actual = $sut->getCell('C2');
        $this->assertEquals($expected, $actual);
    }
}
