           |     |
            \   /
             \_/
        __   /^\   __
       '  `. \_/ ,'  `
            \/ \/
       _,--./| |\.--._
    _,'   _.-\_/-._   `._
         |   / \   |
         |  /   \  |
        /   |   |   \
      -'    \___/    `-

#ExcelAnt

[![Build Status](https://travis-ci.org/Wisembly/ExcelAnt.png?branch=master)](https://travis-ci.org/Wisembly/ExcelAnt)

ExcelAnt is an Excel manipulation library for PHP 5.4. It currently works on top of [PHPExcel](https://github.com/PHPOffice/PHPExcel).
If you want to add / use another library, feel free to fork and contribute !

#Version

1.0.0

#Installation

1. Install composer : `curl -s http://getcomposer.org/installer | php`
(more info at getcomposer.org)
2. Create a `composer.json` file in your project root :
(or add only the excelant line in your existing composer file)

```yml
  {
    "require": {
      "wisembly/excelant": "*",
    }
  }
```

3. Install via composer : `php composer.phar install`

#Use ExcelAnt

Create a simple Table :

```php
use ExcelAnt\Adapter\PhpExcel\Workbook\Workbook,
    ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Adapter\PhpExcel\Writer\Writer,
    ExcelAnt\Table\Table,
    ExcelAnt\Coordinate\Coordinate;

Class Export
{
    public function createExport(array $users)
    {
        $workbook = new Workbook();
        $sheet = new Sheet($workbook);
        $table = new Table();

        foreach ($users as $user) {
            $table->setRow([
                $user->getName(),
                $user->getEmail(),
            ]);
        }

        $sheet->addTable($table, new Coordinate(1, 1));
        $workbook->addSheet($sheet);
    }
}
```

Now, to export your Workbook, you need to create a Writer :

```php
use ExcelAnt\Adapter\PhpExcel\Writer\WriterFactory,
    ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\Excel5;

$writer = (new WriterFactory())->createWriter(new Excel5('/path/to/myExport.xls'));
```

Convert your Worbook to create a PHPExcel object and export it :

```php
$phpExcel = $writer->convert($workbook);
$writer->write($phpExcel);
```


![Simple table](https://raw.github.com/Wisembly/ExcelAnt/master/docs/simple-table.png)

#Documentation

Coming soon...

#Contributing

ExcelAnt is an open source project. If you would like to contribute, fork the repository and submit a pull request.

#Running ExcelAnt Tests
