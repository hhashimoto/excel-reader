<?php

class SheetsTest extends \PHPUnit_Framework_TestCase {
    /**
     * @test
     */
    public function findReturnsSheetWithSameName() {
        $sheet1 = new Sheet('test1', 1, 'rId1');
        $sheet2 = new Sheet('test2', 2, 'rId2');
        $sheet3 = new Sheet('test3', 3, 'rId3');

        $sut = new Sheets;
        $sut->add($sheet1);
        $sut->add($sheet2);
        $sut->add($sheet3);

        $this->assertEquals($sheet1, $sut->find('test1'));
        $this->assertEquals($sheet2, $sut->find('test2'));
        $this->assertEquals($sheet3, $sut->find('test3'));

        $this->assertEquals(null, $sut->find('test１'));
        $this->assertEquals(null, $sut->find(''));
    }
}
