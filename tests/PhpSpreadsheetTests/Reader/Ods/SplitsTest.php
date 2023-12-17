<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Reader\Ods;

use PhpOffice\PhpSpreadsheet\Reader\Ods;
use PhpOffice\PhpSpreadsheetTests\Functional\AbstractFunctional;

class SplitsTest extends AbstractFunctional
{
    private string $testbook;

    protected function setUp(): void
    {
        $this->testbook = dirname(__DIR__, 3) . '/data/Reader/Ods/splits.ods';
    }

    public function testSplits(): void
    {
        $reader = new Ods();
        $spreadsheet = $reader->load($this->testbook);

        $sheet = $spreadsheet->getSheetByNameOrThrow('Freeze');
        self::assertSame('E7', $sheet->getFreezePane());
        self::assertSame('frozen', $sheet->getPaneState());
        self::assertSame('L7', $sheet->getPaneTopLeftCell());
        self::assertSame('L7', $sheet->getSelectedCells());

        $sheet = $spreadsheet->getSheetByNameOrThrow('SplitVertical');
        self::assertNull($sheet->getFreezePane());
        self::assertNotEquals(0, $sheet->getXSplit());
        self::assertEquals(0, $sheet->getYSplit());
        self::assertSame('G1', $sheet->getTopLeftCell());
        self::assertSame('E1', $sheet->getPaneTopLeftCell());
        self::assertSame('E1', $sheet->getSelectedCells());

        $sheet = $spreadsheet->getSheetByNameOrThrow('SplitHorizontal');
        self::assertNull($sheet->getFreezePane());
        self::assertEquals(0, $sheet->getXSplit());
        self::assertNotEquals(0, $sheet->getYSplit());
        self::assertSame('A3', $sheet->getTopLeftCell());
        self::assertSame('A6', $sheet->getPaneTopLeftCell());
        self::assertSame('A7', $sheet->getSelectedCells());

        $sheet = $spreadsheet->getSheetByNameOrThrow('SplitBoth');
        self::assertNull($sheet->getFreezePane());
        self::assertNotEquals(0, $sheet->getXSplit());
        self::assertNotEquals(0, $sheet->getYSplit());
        self::assertSame('H3', $sheet->getTopLeftCell());
        self::assertSame('E19', $sheet->getPaneTopLeftCell());
        self::assertSame('E20', $sheet->getSelectedCells());

        $sheet = $spreadsheet->getSheetByNameOrThrow('NoFreezeNorSplit');
        self::assertNull($sheet->getFreezePane());
        self::assertSame('D3', $sheet->getTopLeftCell());
        self::assertSame('D5', $sheet->getSelectedCells());

        $sheet = $spreadsheet->getSheetByNameOrThrow('FrozenSplit');
        self::assertSame('B4', $sheet->getFreezePane());
        self::assertSame('frozen', $sheet->getPaneState());
        self::assertSame('B4', $sheet->getPaneTopLeftCell());
        self::assertSame('B4', $sheet->getTopLeftCell());
        self::assertSame('B4', $sheet->getSelectedCells());

        $reloadedSpreadsheet = $this->writeAndReload($spreadsheet, 'Ods');
        $spreadsheet->disconnectWorksheets();

        $sheet = $reloadedSpreadsheet->getSheetByNameOrThrow('Freeze');
        self::assertSame('E7', $sheet->getFreezePane());
        self::assertSame('frozen', $sheet->getPaneState());
        self::assertSame('L7', $sheet->getPaneTopLeftCell());
        self::assertSame('L7', $sheet->getSelectedCells());

        $sheet = $reloadedSpreadsheet->getSheetByNameOrThrow('SplitVertical');
        self::assertNull($sheet->getFreezePane());
        self::assertNotEquals(0, $sheet->getXSplit());
        self::assertEquals(0, $sheet->getYSplit());
        self::assertSame('E1', $sheet->getPaneTopLeftCell());
        self::assertSame('E1', $sheet->getSelectedCells());

        $sheet = $reloadedSpreadsheet->getSheetByNameOrThrow('SplitHorizontal');
        self::assertNull($sheet->getFreezePane());
        self::assertEquals(0, $sheet->getXSplit());
        self::assertNotEquals(0, $sheet->getYSplit());
        self::assertSame('A6', $sheet->getPaneTopLeftCell());
        self::assertSame('A7', $sheet->getSelectedCells());

        $sheet = $reloadedSpreadsheet->getSheetByNameOrThrow('SplitBoth');
        self::assertNull($sheet->getFreezePane());
        self::assertNotEquals(0, $sheet->getXSplit());
        self::assertNotEquals(0, $sheet->getYSplit());
        self::assertSame('H3', $sheet->getTopLeftCell());
        self::assertSame('E19', $sheet->getPaneTopLeftCell());
        self::assertSame('E20', $sheet->getSelectedCells());

        $sheet = $reloadedSpreadsheet->getSheetByNameOrThrow('NoFreezeNorSplit');
        self::assertNull($sheet->getFreezePane());
        self::assertSame('D5', $sheet->getSelectedCells());

        $sheet = $reloadedSpreadsheet->getSheetByNameOrThrow('FrozenSplit');
        self::assertSame('B4', $sheet->getFreezePane());
        self::assertSame('frozen', $sheet->getPaneState());
        self::assertSame('B4', $sheet->getPaneTopLeftCell());
        self::assertSame('B4', $sheet->getTopLeftCell());
        self::assertSame('B4', $sheet->getSelectedCells());

        $reloadedSpreadsheet->disconnectWorksheets();
    }
}
