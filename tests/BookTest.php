<?php
namespace hhashimoto\excel\tests;

use hhashimoto\excel\Book;
use hhashimoto\excel\Sheet;

class BookTest extends \PHPUnit_Framework_TestCase {
    /**
     * @test
     * @expectedException Exception
     */
    public function numberOfSheetsThrowExceptionWhenBookNotLoaded() {
        $sut = new Book;
        $sut->numberOfSheets();
    }

    /**
     * @test
     */
    public function numberOfSheetsReturnsNumberOfSheetsInTheBook() {
        $sut = new Book;

        $sut->load('tests/fixtures/test1.xlsx');
        $this->assertSame(3, $sut->numberOfSheets());

        $sut->load('tests/fixtures/test2.xlsx');
        $this->assertSame(5, $sut->numberOfSheets());
    }

    /**
     * @test
     * @expectedException Exception
     */
    public function getSheetByNameThrowExceptionWhenBookNotLoaded() {
        $sut = new Book;
        $sut->getSheetByName('sheet1');
    }

    /**
     * @test
     */
    public function getSheetByNameReturnsNullWhenTargetSheetNotExists() {
        $sut = new Book;
        $sut->load('tests/fixtures/test1.xlsx');
        $this->assertNull($sut->getSheetByName('nothing!'));
    }

    /**
     * @test
     */
    public function getSheetByName() {
        $sut = new Book;
        $sut->load('tests/fixtures/test1.xlsx');

        // Sheet1 | Sheet3 | Sheet2
        $actual = $sut->getSheetByName('Sheet3');
        $expected = new Sheet('Sheet3', '3', 'rId2');
        $this->assertEquals($expected, $actual);
    }

    /**
     * @test
     */
    public function name() {
        $sut = new Book;
        $sut->load('tests/fixtures/test1.xlsx');

        $this->assertSame('tests/fixtures/test1.xlsx', $sut->name());
    }

    /**
     * @test
     */
    public function getSheetNames() {
        $sut = new Book;
        $sut->load('tests/fixtures/test1.xlsx');

        $actual = $sut->getSheetNames();
        $expected = ['Sheet1', 'Sheet3', 'Sheet2'];

        $this->assertEquals($expected, $actual);
    }
}
