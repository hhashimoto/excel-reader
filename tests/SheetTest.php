<?php
namespace hhashimoto\excel\tests;

use hhashimoto\excel\Book;
use hhashimoto\excel\Sheet;
use hhashimoto\excel\Cell;

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
        $actual = $sut->getCell('C1');
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

    /**
     * @test
     */
    public function cannotGetCellBelongToNotTargetColumns() {
        $book = new Book;
        $book->load('tests/fixtures/test1.xlsx');

        $sut = $book->getSheetByName('Sheet2');
        $sut->setTargetColumnsToLoad(['B']);

        $expected = new Cell('sheetName');
        $actual = $sut->getCell('B2');
        $this->assertEquals($expected, $actual);

        $expected = new Cell('');
        $actual = $sut->getCell('C2');
        $this->assertEquals($expected, $actual);
    }

    /**
     * @test
     */
    public function maxColumnReturnsRightColumnName() {
        $book = new Book;
        $book->load('tests/fixtures/test1.xlsx');

        $sut = $book->getSheetByName('Sheet1');
        $this->assertSame('E', $sut->maxColumn());

        $sut = $book->getSheetByName('Sheet2');
        $this->assertSame('C', $sut->maxColumn());
    }

    /**
     * @test
     */
    public function maxRowReturnsBottomRowNumber() {
        $book = new Book;
        $book->load('tests/fixtures/test1.xlsx');

        $sut = $book->getSheetByName('Sheet1');
        $this->assertSame(1, $sut->maxRow());

        $sut = $book->getSheetByName('Sheet2');
        $this->assertSame(2, $sut->maxRow());
    }
}
