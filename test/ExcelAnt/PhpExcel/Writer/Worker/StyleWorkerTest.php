<?php

namespace ExcelAnt\PhpExcel\Writer\Worker;

use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Collections\StyleCollection;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Style\Format;
use ExcelAnt\Style\Borders;
use ExcelAnt\Style\Border;
use ExcelAnt\Style\Alignment;

class StyleWorkerTest extends \PHPUnit_Framework_TestCase
{
    public function testApplyStyleOnCell()
    {
        $localStyleStorage = [];

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelStyle = $this->getPhpExcelStyleMock();

        $phpExcelStyle->expects($this->any())
            ->method('applyFromArray')
            ->will($this->returnCallback(function($styles) use (&$localStyleStorage) {
                $localStyleStorage = $styles;
            }));

        $phpExcelWorksheet->expects($this->any())
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));

        $styleCollection = new StyleCollection([
            (new Fill())->setColor('000000'),
            (new Font())->setName('Verdana'),
            (new Format())->setFormat(Format::TYPE_BOOL),
            (new Borders())->setTop((new Border())->setColor('ff0000')),
            (new Alignment())->setVertical(Alignment::VERTICAL_CENTER),
        ]);

        $expected = [
            'fill' => ['color' => ['rgb' => '000000']],
            'font' => [
                'name'      => 'Verdana',
                'size'      => 11,
                'bold'      => false,
                'italic'    => false,
                'color'     => ['rgb' => '000000'],
                'underline' => Font::UNDERLINE_NONE,
            ],
            'alignment'     => [
                'horizontal' => 'general',
                'vertical'   => 'center',
            ],
            'borders' => [
                'top' => [
                    'color' => 'ff0000',
                    'type'  => Border::BORDER_NONE,
                ],
            ],
        ];

        $styleWorker = (new StyleWorker())->applyStyles($phpExcelWorksheet, new Coordinate(1, 1), $styleCollection);

        $this->assertEquals($expected, $localStyleStorage);
    }

    /**
     * Mock PHPExcel_Worksheet
     * @return Mock
     */
    private function getPhpExcelWorksheetMock()
    {
        return $this->getMockBuilder('PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock PHPExcel_Style
     * @return Mock
     */
    private function getPhpExcelStyleMock()
    {
        return $this->getMockBuilder('PHPExcel_Style')->disableOriginalConstructor()->getMock();
    }
}