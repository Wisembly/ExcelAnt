<?php

namespace ExcelAnt\Style;

abstract class StyleTestCase extends \PHPUnit_Framework_TestCase
{
    public function getWrongParameters()
    {
        return [
            ['foo'],
            [''],
            [null],
            ['@&'],
        ];
    }
}