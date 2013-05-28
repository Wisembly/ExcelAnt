<?php

namespace ExcelAnt\Traits;

use ExcelAnt\Style\StyleInterface;

trait StyleCollectionRequirements
{
    public function checkStyleCollectionRequirements($styles)
    {
        if (isset($styles)) {
            if (!is_array($styles)) {
                throw new \InvalidArgumentException("styles must be an array of StyleInterface");
            }

            foreach ($styles as $style) {
                if (!($style instanceof StyleInterface)) {
                    throw new \InvalidArgumentException("style must be an instance of StyleInterface");
                }
            }
        }
    }
}