<?php

namespace ExcelAnt\PhpExcel;

use PHPExcel;
use PHPExcel_DocumentProperties;

use ExcelAnt\PhpExcel\Worksheet;
use ExcelAnt\PhpExcel\Sheet;
use ExcelAnt\Collections\StyleCollection;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;

class WorksheetTest extends \PHPUnit_Framework_TestCase
{
    public function testRawClass()
    {
        $worksheet = $this->createWorksheet();
        $this->assertInstanceOf("PHPExcel", $worksheet->getRawClass());
    }

    public function testCreateSheet()
    {
        $worksheet = $this->createWorksheet();
        $sheet = $worksheet->createSheet();

        $this->assertInstanceOf("ExcelAnt\PhpExcel\Sheet", $sheet);
        $this->assertCount(1, $worksheet->getAllSheets());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetSheetWithInvalidArgument()
    {
        $worksheet = $this->createWorksheet();
        $worksheet->getSheet("foo");
    }

    /**
     * @expectedException \RuntimeException
     */
    public function testGetSheetWithNonExistentIndex()
    {
        $worksheet = $this->createWorksheet();
        $worksheet->getSheet(count($worksheet->getAllSheets()) + 1);
    }

    public function testGetSheet()
    {
        $worksheet = $this->createWorksheet();
        $worksheet->createSheet();
        $sheet = $worksheet->getSheet(0);

        $this->assertInstanceOf("ExcelAnt\PhpExcel\Sheet", $sheet);
    }

    public function testCountSheets()
    {
        $worksheet = $this->createWorksheet();
        $worksheet->createSheet();
        $worksheet->createSheet();
        $worksheet->createSheet();

        $this->assertEquals(3, $worksheet->countSheets());
    }

    public function testAddSheetWithoutIndex()
    {
        $worksheet = $this->createWorksheet();
        $sheet = $this->getSheetMock('Foo');
        $worksheet->addSheet($sheet);

        $this->assertCount(1, $worksheet->getAllSheets());
        $this->assertEquals('Foo', $worksheet->getSheet($worksheet->countSheets() - 1)->getTitle());
    }

    public function testAddSheetWithIndex()
    {
        $worksheet = $this->createWorksheet();
        $sheet = $this->getSheetMock('Foo');
        $worksheet->addSheet($sheet);

        $sheet = $this->getSheetMock('Bar');
        $worksheet->addSheet($sheet, 0);

        $this->assertEquals('Bar', $worksheet->getSheet(0)->getTitle());
        $this->assertCount(1, $worksheet->getAllSheets());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testAddSheetWithAWrongParameter()
    {
        $worksheet = $this->createWorksheet();
        $sheet = $this->getSheetMock('Foo');
        $worksheet->addSheet($sheet, 'foo');
    }

    /**
     * @dataProvider getDataToInsert
     */
    public function testInsertSheet($sheetCollection, $sheetToInsert, $index)
    {
        $worksheet = $this->createWorksheet();

        // Add sheets
        foreach ($sheetCollection as $sheet) {
            $worksheet->addSheet($sheet);
        }

        // Insert
        $worksheet->addSheet($sheetToInsert, $index, true);

        // Get new data
        $newSheetCollection = $worksheet->getAllSheets();
        $countNewSheetCollection = $worksheet->countSheets();

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
        $worksheet = $this->createWorksheet();

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
        $worksheet = $this->createWorksheet();
        $worksheet->createSheet();
        $worksheet->createSheet();

        $worksheet->removeSheet(2);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testRemoveSheetWithAWrongParameter()
    {
        $worksheet = $this->createWorksheet();
        $worksheet->createSheet();

        $worksheet->removeSheet('foo');
    }

    public function testRemoveSheet()
    {
        $worksheet = $this->createWorksheet();
        $worksheet->addSheet($this->getSheetMock('Foo'));
        $worksheet->addSheet($this->getSheetMock('Bar'));

        $worksheet->removeSheet(0);
        $sheeCollection = $worksheet->getAllSheets();

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

        $worksheet = $this->createWorksheet($phpExcel);
        $worksheet->setTitle('Foo');
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

        $worksheet = $this->createWorksheet($phpExcel);

        $this->assertEquals('Foo', $worksheet->getTitle());
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

        $worksheet = $this->createWorksheet($phpExcel);
        $worksheet->setCreator('Foo');
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

        $worksheet = $this->createWorksheet($phpExcel);

        $this->assertEquals('Foo', $worksheet->getCreator());
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

        $worksheet = $this->createWorksheet($phpExcel);
        $worksheet->setDescription('Foo');
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

        $worksheet = $this->createWorksheet($phpExcel);

        $this->assertEquals('Foo', $worksheet->getDescription());
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

        $worksheet = $this->createWorksheet($phpExcel);
        $worksheet->setCompany('Foo');
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

        $worksheet = $this->createWorksheet($phpExcel);

        $this->assertEquals('Foo', $worksheet->getCompany());
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

        $worksheet = $this->createWorksheet($phpExcel);
        $worksheet->setSubject('Foo');
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

        $worksheet = $this->createWorksheet($phpExcel);

        $this->assertEquals('Foo', $worksheet->getSubject());
    }

    public function testAddStyleAndGetStyleCollection()
    {
        $worksheet = $this->createWorksheet();
        $worksheet->addStyles(new StyleCollection([new Fill(), new Font()]));

        $this->assertCount(2, $worksheet->getStyles());
        $this->assertInstanceOf('ExcelAnt\Collections\StyleCollection', $worksheet->getStyles());
    }

    /**
     * Create a Worksheet
     *
     * @param  MockPhpExcel $phpExcel
     *
     * @return Worksheet
     */
    public function createWorksheet($phpExcel = null)
    {
        $phpExcel = $phpExcel ?: $this->getPhpExcelMock();

        return new Worksheet($phpExcel);
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
     * Mock ExcelAnt\PhpExcel\Sheet
     *
     * @return Mock_Sheet
     */
    private function getSheetMock($title = null)
    {
        $sheet = $this->getMockBuilder('ExcelAnt\PhpExcel\Sheet')->disableOriginalConstructor()->getMock();

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