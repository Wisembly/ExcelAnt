<?php

namespace ExcelAnt\Traits;

use ExcelAnt\Traits\StyleCollectionRequirements;

class StyleCollectionRequirementsTest extends \PHPUnit_Framework_TestCase
{
    private $styleCollectionRequirements;

    public function setUp()
    {
        $this->styleCollectionRequirements = $this->getObjectForTrait('ExcelAnt\Traits\StyleCollectionRequirements');
    }

    /**
     * @expectedException \InvalidArgumentException
     * @dataProvider      getStyles
     */
    public function testCheckStyleCollectionRequirementsWithWrongParameters($styles)
    {
        $this->styleCollectionRequirements->checkStyleCollectionRequirements($styles);
    }

    public function getStyles()
    {
        return [
            ['foo'],
            [$this->getStyleInterfaceMock(), 'foo'],
        ];
    }

    public function testCheckStyleCollectionRequirements()
    {
        $styles = [$this->getStyleInterfaceMock(), $this->getStyleInterfaceMock()];
        $this->styleCollectionRequirements->checkStyleCollectionRequirements($styles);
    }

    /**
     * Mock StyleInterface
     *
     * @return Mock_StyleInterface
     */
    private function getStyleInterfaceMock()
    {
        return $this->getMockBuilder('ExcelAnt\Style\StyleInterface')->disableOriginalConstructor()->getMock();
    }
}