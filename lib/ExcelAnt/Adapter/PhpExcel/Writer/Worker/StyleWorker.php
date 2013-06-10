<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Style\Borders,
    ExcelAnt\Style\Alignment;

class StyleWorker
{
    /**
     * Convert StyleCollection to PHPExcel_Worksheet data
     *
     * @param  StyleCollection    $styleCollection   The StyleCollection
     *
     * @return PHPExcel_Worksheet
     */
    public function convertStyles(StyleCollection $styleCollection)
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

        return $styles;
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
            'type'  => $fill->getType(),
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