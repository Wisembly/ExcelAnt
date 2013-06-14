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

ExcelAnt is an Excel manipulation library for PHP 5.4. Actually it works with [PHPExcel](https://github.com/PHPOffice/PHPExcel).
If you want to add another library, fork and contribute to improve it.

#Installation

1. Install composer in your project: `curl -s http://getcomposer.org/installer | php`
2. Create a `composer.json` file in your project root:

```yml
  {
    "require": {
      "wisembly/excelant": "dev-master",
    }
  }
```

3. Install via composer : `php composer.phar install`

#Usage with PHPExcel

```php
use ExcelAnt\Adapter\PhpExcel\Workbook\Workbook,
    ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Table\Table;

Class Export
{
    public function createExport(array $users)
    {
        $workbook = new Workbook();
        $sheet = new Sheet();
        $table = new Table();

        foreach($user as $user) {
            $table->setRow([
                $user->getName(),
                $user->getEmail(),
            ]);
        }

        $sheet->addTable($table);
        $workbook->addSheet($sheet);
    }
}
```

Now to export your Workbook, you need to create a Writer :

```php
use ExcelAnt\Adapter\PhpExcel\Writer\Worker\StyleWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\LabelWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\TableWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\Excel5;

$styleWorker = new StyleWorker();
$cellWorker = new CellWorker($styleWorker);
$labelWorker = new LabelWorker($cellWorker);
$tableWorker = new TableWorker($cellWorker, $labelWorker);

$writer = new Writer(new Excel5('myExport.xls'), $tableWorker, $cellWorker, $styleWorker);
```

Convert your Worbook to create a PHPExcel object and export it :

```php
$phpExcel = $writer->convert($workbook);
$writer->write($phpExcel);
```

#Contributing
ExcelAnt is an open source project. If you would like to contribute, fork the repository and submit a pull request.

#Running ExcelAnt Tests