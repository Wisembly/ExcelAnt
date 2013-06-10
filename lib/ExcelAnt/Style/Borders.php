<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface,
    ExcelAnt\Style\Border;

class Borders implements StyleInterface
{
    private $top;
    private $bottom;
    private $left;
    private $right;

    public function setTop(Border $border)
    {
        $this->top = $border;

        return $this;
    }

    public function getTop()
    {
        return $this->top;
    }

    public function setBottom(Border $border)
    {
        $this->bottom = $border;

        return $this;
    }

    public function getBottom()
    {
        return $this->bottom;
    }

    public function setLeft(Border $border)
    {
        $this->left = $border;

        return $this;
    }

    public function getLeft()
    {
        return $this->left;
    }

    public function setRight(Border $border)
    {
        $this->right = $border;

        return $this;
    }

    public function getRight()
    {
        return $this->right;
    }

    public function getBorders()
    {
        return [
            'top' => $this->top,
            'bottom' => $this->bottom,
            'left' => $this->left,
            'right' => $this->right,
        ];
    }
}