<?php
namespace hhashimoto\excel\tests;

use hhashimoto\excel\Cell;

class CellTest extends \PHPUnit_Framework_TestCase {
    /**
     * @test
     */
    public function getValue() {
        $sut = new Cell;
        $this->assertEmpty($sut->getValue());

        $sut = new Cell('test');
        $this->assertSame('test', $sut->getValue());
    }
}
