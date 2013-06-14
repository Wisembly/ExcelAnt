<?php

if (!($loader = include __DIR__ . '/../vendor/autoload.php')) {
    throw new RuntimeException('Install dependencies to run test suite. "php composer.phar install --dev"');
}

$loader->add('ExcelAnt\Tests', __DIR__);