<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Style\Format,
    ExcelAnt\Style\Borders,
    ExcelAnt\Style\Border,
    ExcelAnt\Style\Alignment;

class StyleWorkerTest extends \PHPUnit_Framework_TestCase
{
    public function testApplyStyleOnCell()
    {
        $styleCollection = new StyleCollection([
            (new Fill())->setColor('000000'),
            (new Font())->setName('Verdana'),
            (new Format())->setFormat(Format::TYPE_BOOL),
            (new Borders())->setTop((new Border())->setColor('ff0000')),
            (new Alignment())->setVertical(Alignment::VERTICAL_CENTER),
        ]);

        $expected = [
            'fill' => [
                'type'  => 'solid',
                'color' => ['rgb' => '000000']
            ],
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
                    'color'  => ['rgb' => 'ff0000'],
                    'style'  => Border::BORDER_THIN,
                ],
            ],
        ];

        $styles = (new StyleWorker())->convertStyles($styleCollection);

        $this->assertEquals($expected, $styles);
    }
}