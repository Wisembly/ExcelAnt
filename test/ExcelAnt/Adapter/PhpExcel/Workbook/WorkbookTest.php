<?php

namespace ExcelAnt\Adapter\PhpExcel;

use \PHPExcel;
use \PHPExcel_DocumentProperties;

use ExcelAnt\Adapter\PhpExcel\Workbook\Workbook,
    ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font;

class WorkbookTest extends \PHPUnit_Framework_TestCase
{
    public function testRawClass()
    {
        $workbook = $this->createWorkbook();
        $this->assertInstanceOf("PHPExcel", $workbook->getRawClass());
    }

    public function testCreateSheet()
    {
        $workbook = $this->createWorkbook();
        $sheet = $workbook->createSheet();

        $this->assertInstanceOf("ExcelAnt\Adapter\PhpExcel\Sheet\Sheet", $sheet);
        $this->assertCount(1, $workbook->getAllSheets());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetSheetWithInvalidArgument()
    {
        $workbook = $this->createWorkbook()->getSheet("foo");
    }

    /**
     * @expectedException \RuntimeException
     */
    public function testGetSheetWithNonExistentIndex()
    {
        $workbook = $this->createWorkbook();
        $workbook->getSheet(count($workbook->getAllSheets()) + 1);
    }

    public function testGetSheet()
    {
        $workbook = $this->createWorkbook();
        $workbook->createSheet();
        $sheet = $workbook->getSheet(0);

        $this->assertInstanceOf("ExcelAnt\Adapter\PhpExcel\Sheet\Sheet", $sheet);
    }

    public function testCountSheets()
    {
        $workbook = $this->createWorkbook();
        $workbook->createSheet();
        $workbook->createSheet();
        $workbook->createSheet();

        $this->assertEquals(3, $workbook->countSheets());
    }

    public function testAddSheetWithoutIndex()
    {
        $workbook = $this->createWorkbook();
        $sheet = $this->getSheetMock('Foo');
        $workbook->addSheet($sheet);

        $this->assertCount(1, $workbook->getAllSheets());
        $this->assertEquals('Foo', $workbook->getSheet($workbook->countSheets() - 1)->getTitle());
    }

    public function testAddSheetWithIndex()
    {
        $workbook = $this->createWorkbook();
        $sheet = $this->getSheetMock('Foo');
        $workbook->addSheet($sheet);

        $sheet = $this->getSheetMock('Bar');
        $workbook->addSheet($sheet, 0);

        $this->assertEquals('Bar', $workbook->getSheet(0)->getTitle());
        $this->assertCount(1, $workbook->getAllSheets());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testAddSheetWithAWrongParameter()
    {
        $workbook = $this->createWorkbook();
        $sheet = $this->getSheetMock('Foo');
        $workbook->addSheet($sheet, 'foo');
    }

    /**
     * @dataProvider getDataToInsert
     */
    public function testInsertSheet($sheetCollection, $sheetToInsert, $index)
    {
        $workbook = $this->createWorkbook();

        // Add sheets
        foreach ($sheetCollection as $sheet) {
            $workbook->addSheet($sheet);
        }

        // Insert
        $workbook->addSheet($sheetToInsert, $index, true);

        // Get new data
        $newSheetCollection = $workbook->getAllSheets();
        $countNewSheetCollection = $workbook->countSheets();

        // Asserts
        $this->assertCount(count($sheetCollection) + 1, $newSheetCollection);

        for ($i=0; $i < $countNewSheetCollection; $i++) {
            if ($i < $index) {
                $this->assertEquals($sheetCollection[$i]->getTitle(), $newSheetCollection[$i]->getTitle());

                continue;
            }

            if ($i === $index) {
                $this->assertEquals($sheetToInsert->getTitle(), $newSheetCollection[$i]->getTitle());

                continue;
            }

            $this->assertEquals($sheetCollection[$i - 1]->getTitle(), $newSheetCollection[$i]->getTitle());
        }
    }

    public function getDataToInsert()
    {
        $workbook = $this->createWorkbook();

        return [
            [[$this->getSheetMock('Foo'), $this->getSheetMock('Bar'), $this->getSheetMock('Baz')], $this->getSheetMock('Insert'), 0],
            [[$this->getSheetMock('Foo'), $this->getSheetMock('Bar'), $this->getSheetMock('Baz')], $this->getSheetMock('Insert'), 1],
            [[$this->getSheetMock('Foo'), $this->getSheetMock('Bar'), $this->getSheetMock('Baz')], $this->getSheetMock('Insert'), 2],
        ];
    }

    /**
     * @expectedException \RuntimeException
     */
    public function testRemoveSheetWithNonExistentIndex()
    {
        $workbook = $this->createWorkbook();
        $workbook->createSheet();
        $workbook->createSheet();
        $workbook->removeSheet(2);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testRemoveSheetWithAWrongParameter()
    {
        $workbook = $this->createWorkbook();
        $workbook->createSheet();
        $workbook->removeSheet('foo');
    }

    public function testRemoveSheet()
    {
        $workbook = $this->createWorkbook()
            ->addSheet($this->getSheetMock('Foo'))
            ->addSheet($this->getSheetMock('Bar'))
            ->removeSheet(0);
        $sheeCollection = $workbook->getAllSheets();

        $this->assertCount(1, $sheeCollection);
        $this->assertTrue(array_key_exists(0, $sheeCollection));
        $this->assertEquals('Bar', $sheeCollection[0]->getTitle());
    }

    public function testSetTitle()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('setTitle')
            ->will($this->returnValue($phpExcelDocument));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel)
            ->setTitle('Foo');
    }

    public function testGetTitle()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('getTitle')
            ->will($this->returnValue('Foo'));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel);

        $this->assertEquals('Foo', $workbook->getTitle());
    }

    public function testSetCreator()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('setCreator')
            ->will($this->returnValue($phpExcelDocument));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel)
            ->setCreator('Foo');
    }

    public function testGetCreator()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('getCreator')
            ->will($this->returnValue('Foo'));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel);

        $this->assertEquals('Foo', $workbook->getCreator());
    }

    public function testSetDescription()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('setDescription')
            ->will($this->returnValue($phpExcelDocument));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel)
            ->setDescription('Foo');
    }

    public function testGetDescription()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('getDescription')
            ->will($this->returnValue('Foo'));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel);

        $this->assertEquals('Foo', $workbook->getDescription());
    }

    public function testSetCompany()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('setCompany')
            ->will($this->returnValue($phpExcelDocument));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel)
            ->setCompany('Foo');
    }

    public function testGetCompany()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('getCompany')
            ->will($this->returnValue('Foo'));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel);

        $this->assertEquals('Foo', $workbook->getCompany());
    }

    public function testSetSubject()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('setSubject')
            ->will($this->returnValue($phpExcelDocument));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel)
            ->setSubject('Foo');
    }

    public function testGetSubject()
    {
        $phpExcelDocument = $this->getPhpExcelDocumentPropertiesMock();
        $phpExcelDocument->expects($this->once())
            ->method('getSubject')
            ->will($this->returnValue('Foo'));

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getProperties')
            ->will($this->returnValue($phpExcelDocument));

        $workbook = $this->createWorkbook($phpExcel);

        $this->assertEquals('Foo', $workbook->getSubject());
    }

    public function testAddStyleAndGetStyleCollection()
    {
        $workbook = $this->createWorkbook()
            ->addStyles(new StyleCollection([new Fill(), new Font()]));

        $this->assertCount(2, $workbook->getStyles());
        $this->assertInstanceOf('ExcelAnt\Collections\StyleCollection', $workbook->getStyles());
    }

    /**
     * Create a Workbook
     *
     * @param  MockPhpExcel $phpExcel
     *
     * @return Workbook
     */
    public function createWorkbook($phpExcel = null)
    {
        $phpExcel = $phpExcel ?: $this->getPhpExcelMock();

        return new Workbook($phpExcel);
    }

    /**
     * Mock PHPExcel
     *
     * @return Mock_PHPExcel
     */
    private function getPhpExcelMock()
    {
        return $this->getMockBuilder('PHPExcel')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock ExcelAnt\Adapter\PhpExcel\Sheet\Sheet
     *
     * @return Mock_Sheet
     */
    private function getSheetMock($title = null)
    {
        $sheet = $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Sheet\Sheet')->disableOriginalConstructor()->getMock();

        if (null !== $title) {
            $sheet->expects($this->any())
             ->method('getTitle')
             ->will($this->returnValue($title));
        }

        return $sheet;
    }

    /**
     * Mock PHPExcel_DocumentProperties
     *
     * @return Mock_PHPExcel_DocumentProperties
     */
    private function getPhpExcelDocumentPropertiesMock()
    {
        return $this->getMockBuilder('PHPExcel_DocumentProperties')->disableOriginalConstructor()->getMock();
    }
}