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
}
