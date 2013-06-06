<?php

namespace ExcelAnt\PhpExcel\Writer\Worker;

use PHPExcel_Worksheet;

use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Collections\StyleCollection;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Style\Borders;
use ExcelAnt\Style\Alignment;

class StyleWorker
{
    /**
     * Convert StyleCollection to PHPExcel_Worksheet data
     *
     * @param  PHPExcel_Worksheet $phpExcelWorksheet The current worksheet
     * @param  Coordinate         $coordinate        The coordinate where the style must to be applied
     * @param  StyleCollection    $styleCollection   The StyleCollection
     *
     * @return PHPExcel_Worksheet
     */
    public function applyStyles(PHPExcel_Worksheet $phpExcelWorksheet, Coordinate $coordinate, StyleCollection $styleCollection)
    {
        $styles = [];

        foreach ($styleCollection as $style) {

            switch (get_class($style)) {

                case 'ExcelAnt\Style\Fill':
                    $styles['fill'] = $this->fillManager($style);
                    break;

                case 'ExcelAnt\Style\Font':
                    $styles['font'] = $this->fontManager($style);
                    break;

                case 'ExcelAnt\Style\Borders':
                    $styles['borders'] = $this->bordersManager($style);
                    break;

                case 'ExcelAnt\Style\Alignment':
                    $styles['alignment'] = $this->alignmentManager($style);
                    break;
            }
        }
        error_log(var_export($styles, true));
        $phpExcelWorksheet
            ->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())
            ->applyFromArray($styles);
    }

    /**
     * Fill management
     *
     * @param Fill $fill
     *
     * @return array
     */
    private function fillManager(Fill $fill)
    {
        return [
            'color' => ['rgb' => $fill->getColor()],
        ];
    }

    /**
     * Font management
     *
     * @param Font $font
     *
     * @return array
     */
    private function fontManager(Font $font)
    {
        return [
            'name'      => $font->getName(),
            'size'      => $font->getSize(),
            'bold'      => $font->isBold(),
            'italic'    => $font->isItalic(),
            'color'     => ['rgb' => $font->getColor()],
            'underline' => $font->getUnderline(),
        ];
    }

    /**
     * Alignment management
     *
     * @param Alignment $alignment
     *
     * @return array
     */
    private function alignmentManager(Alignment $alignemnt)
    {
        return [
            'horizontal' => $alignemnt->getHorizontal(),
            'vertical'   => $alignemnt->getVertical(),
        ];
    }

    /**
     * Boders management
     *
     * @param Borders $borders
     *
     * @return array
     */
    private function bordersManager(Borders $borders)
    {
        $bordersToReturn = [];

        if (null !== $borders->getTop()) {
            $bordersToReturn['top'] = [
                'color' => ['rgb' => $borders->getTop()->getColor()],
                'style'  => $borders->getTop()->getType(),
            ];
        }

        if (null !== $borders->getBottom()) {
            $bordersToReturn['bottom'] = [
                'color' => ['rgb' => $borders->getBottom()->getColor()],
                'style'  => $borders->getBottom()->getType(),
            ];
        }

        if (null !== $borders->getLeft()) {
            $bordersToReturn['left'] = [
                'color' => ['rgb' => $borders->getLeft()->getColor()],
                'style'  => $borders->getLeft()->getType(),
            ];
        }

        if (null !== $borders->getRight()) {
            $bordersToReturn['right'] = [
                'color' => ['rgb' => $borders->getRight()->getColor()],
                'style'  => $borders->getRight()->getType(),
            ];
        }

        return $bordersToReturn;
    }
}