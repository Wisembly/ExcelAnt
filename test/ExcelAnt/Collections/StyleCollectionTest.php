<?php

namespace ExcelAnt\Collections;

use ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Style\Format,
    ExcelAnt\Style\Alignment;

class StyleCollectionTest extends \PHPUnit_Framework_TestCase
{
    private $styleCollection;

    public function setUp()
    {
        $this->styleCollection = new StyleCollection([new Fill(), new Font()]);
    }

    public function testAdd()
    {
        $style = new Alignment();
        $this->styleCollection->add($style);
        $this->assertInstanceOf('ExcelAnt\Style\Alignment', $this->styleCollection[2]);
    }

    public function testAddAnExistingStyle()
    {
        $fill = new Fill();
        $fill->setColor('000000');

        $this->assertEquals('ffffff', $this->styleCollection[0]->getColor());
        $this->styleCollection->add($fill);
        $this->assertEquals('000000', $this->styleCollection[0]->getColor());
        $this->assertCount(2, $this->styleCollection);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetWithExistingStyleAndDifferentIndex()
    {
        $this->styleCollection->set(1, new Fill());
    }

    public function testSetWithExistingStyleWithSameIndex()
    {
        $fill = new Fill();
        $fill->setColor('000000');

        $this->assertEquals('ffffff', $this->styleCollection[0]->getColor());
        $this->styleCollection->set(0, $fill);
        $this->assertEquals('000000', $this->styleCollection[0]->getColor());
        $this->assertCount(2, $this->styleCollection);
    }

    public function testSet()
    {
        $this->styleCollection->set(0, new Alignment());

        $this->assertInstanceOf('ExcelAnt\Style\Alignment', $this->styleCollection[0]);
        $this->assertCount(2, $this->styleCollection);
    }

    public function testContains()
    {
        $this->assertFalse($this->styleCollection->contains(new Alignment()));
        $this->assertEquals(0, $this->styleCollection->contains(new Fill()));
    }

    public function testClear()
    {
        $this->assertCount(2, $this->styleCollection);
        $this->styleCollection->clear();
        $this->assertCount(0, $this->styleCollection);
    }

    public function testIsEmpty()
    {
        $this->styleCollection->clear();
        $this->assertTrue($this->styleCollection->isEmpty());
    }

    public function testRemove()
    {
        $style = $this->styleCollection->remove(0);

        $this->assertInstanceOf('ExcelAnt\Style\Fill', $style);
        $this->assertFalse($this->styleCollection->contains($style));
        $this->assertNull($this->styleCollection->remove(0));
    }

    public function testRemoveElement()
    {
        $style = new Fill();

        $this->assertTrue($this->styleCollection->removeElement($style));
        $this->assertFalse($this->styleCollection->contains($style));
        $this->assertFalse($this->styleCollection->removeElement($style));
    }

    public function testContainsKey()
    {
        $this->assertTrue($this->styleCollection->containsKey(0));
        $this->assertFalse($this->styleCollection->containsKey(2));
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testGetElementWithNonExistentElement()
    {
        $this->styleCollection->getElement(new Format());
    }

    public function testGetElement()
    {
        $this->styleCollection->add((new Format())->setFormat(Format::TYPE_NUMERIC));
        $element = $this->styleCollection->getElement(new Format());

        $this->assertInstanceOf('ExcelAnt\Style\Format', $element);
        $this->assertEquals(Format::TYPE_NUMERIC, $element->getFormat());
    }

    public function testGet()
    {
        $this->assertInstanceOf('ExcelAnt\Style\Fill', $this->styleCollection->get(0));
    }

    public function testGetKeys()
    {
        $this->assertEquals([0, 1], $this->styleCollection->getKeys());
    }

    public function testToArray()
    {
        $styles = $this->styleCollection->toArray();

        $this->assertInstanceOf('ExcelAnt\Style\Fill', $styles[0]);
        $this->assertInstanceOf('ExcelAnt\Style\Font', $styles[1]);
    }

    public function testCount()
    {
        $this->assertEquals(2, $this->styleCollection->count());
    }
}