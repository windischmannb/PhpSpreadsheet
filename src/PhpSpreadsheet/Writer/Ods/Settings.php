<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Ods;

use PhpOffice\PhpSpreadsheet\Cell\CellAddress;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Shared\XMLWriter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Settings extends WriterPart
{
    /**
     * Write settings.xml to XML format.
     *
     * @return string XML Output
     */
    public function write(): string
    {
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // Settings
        $objWriter->startElement('office:document-settings');
        $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $objWriter->writeAttribute('xmlns:config', 'urn:oasis:names:tc:opendocument:xmlns:config:1.0');
        $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $objWriter->writeAttribute('office:version', '1.2');

        $objWriter->startElement('office:settings');
        $objWriter->startElement('config:config-item-set');
        $objWriter->writeAttribute('config:name', 'ooo:view-settings');
        $objWriter->startElement('config:config-item-map-indexed');
        $objWriter->writeAttribute('config:name', 'Views');
        $objWriter->startElement('config:config-item-map-entry');
        $spreadsheet = $this->getParentWriter()->getSpreadsheet();

        $objWriter->startElement('config:config-item');
        $objWriter->writeAttribute('config:name', 'ViewId');
        $objWriter->writeAttribute('config:type', 'string');
        $objWriter->text('view1');
        $objWriter->endElement(); // ViewId
        $objWriter->startElement('config:config-item-map-named');

        $this->writeAllWorksheetSettings($objWriter, $spreadsheet);

        $wstitle = $spreadsheet->getActiveSheet()->getTitle();
        $objWriter->startElement('config:config-item');
        $objWriter->writeAttribute('config:name', 'ActiveTable');
        $objWriter->writeAttribute('config:type', 'string');
        $objWriter->text($wstitle);
        $objWriter->endElement(); // config:config-item ActiveTable

        $objWriter->endElement(); // config:config-item-map-entry
        $objWriter->endElement(); // config:config-item-map-indexed Views
        $objWriter->endElement(); // config:config-item-set ooo:view-settings
        $objWriter->startElement('config:config-item-set');
        $objWriter->writeAttribute('config:name', 'ooo:configuration-settings');
        $objWriter->endElement(); // config:config-item-set ooo:configuration-settings
        $objWriter->endElement(); // office:settings
        $objWriter->endElement(); // office:document-settings

        return $objWriter->getData();
    }

    private function writeAllWorksheetSettings(XMLWriter $objWriter, Spreadsheet $spreadsheet): void
    {
        $objWriter->writeAttribute('config:name', 'Tables');

        foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
            $this->writeWorksheetSettings($objWriter, $worksheet);
        }

        $objWriter->endElement(); // config:config-item-map-entry Tables
    }

    private function writeWorksheetSettings(XMLWriter $objWriter, Worksheet $worksheet): void
    {
        $objWriter->startElement('config:config-item-map-entry');
        $objWriter->writeAttribute('config:name', $worksheet->getTitle());

        $this->writeSelectedCells($objWriter, $worksheet);
        $this->writeFreezePane($objWriter, $worksheet);

        $objWriter->endElement(); // config:config-item-map-entry Worksheet
    }

    private function writeSelectedCells(XMLWriter $objWriter, Worksheet $worksheet): void
    {
        $selected = $worksheet->getSelectedCells();
        if (preg_match('/^([a-z]+)([0-9]+)/i', $selected, $matches) === 1) {
            $colSel = Coordinate::columnIndexFromString($matches[1]) - 1;
            $rowSel = (int) $matches[2] - 1;
            $objWriter->startElement('config:config-item');
            $objWriter->writeAttribute('config:name', 'CursorPositionX');
            $objWriter->writeAttribute('config:type', 'int');
            $objWriter->text((string) $colSel);
            $objWriter->endElement();
            $objWriter->startElement('config:config-item');
            $objWriter->writeAttribute('config:name', 'CursorPositionY');
            $objWriter->writeAttribute('config:type', 'int');
            $objWriter->text((string) $rowSel);
            $objWriter->endElement();
        }
    }

    private function writeSplitValue(XMLWriter $objWriter, string $splitMode, string $type, string $value): void
    {
        $objWriter->startElement('config:config-item');
        $objWriter->writeAttribute('config:name', $splitMode);
        $objWriter->writeAttribute('config:type', $type);
        $objWriter->text($value);
        $objWriter->endElement();
    }

    private function writeFreezePane(XMLWriter $objWriter, Worksheet $worksheet): void
    {
        $freezePane = CellAddress::fromCellAddress($worksheet->getFreezePane() ?: 'A1');
        $xSplit = $worksheet->getXSplit();
        $ySplit = $worksheet->getYSplit();

        $columnId = $freezePane->columnId();
        $columnName = $freezePane->columnName();
        $row = $freezePane->rowId();
        $paneTopLeft = CellAddress::fromCellAddress($worksheet->getPaneTopLeftCell() ?: 'A1');
        $topLeftCell = CellAddress::fromCellAddress($worksheet->getTopLeftCell() ?: 'A1');
        $positionLeft = (string) ($topLeftCell->columnId() - 1);
        $positionRight = (string) ($paneTopLeft->columnId() - 1);
        $positionTop = (string) ($topLeftCell->rowId() - 1);
        $positionBottom = (string) ($paneTopLeft->rowId() - 1);
        if ($columnId > 1 || $xSplit === 0) {
            $horizontalSlitMode = $columnId === 1 ? '0' : '2';
            $horizontalSlitPosition = (string) ($columnId - 1);
            $positionTop = '0';
        } else {
            $horizontalSlitMode = '1';
            $horizontalSlitPosition = (string) $xSplit;
        }
        if ($row > 1 || $ySplit === 0) {
            $verticalSplitMode = $row === 1 ? '0' : '2';
            $verticalSplitPosition = (string) ($row - 1);
            $positionLeft = '0';
        } else {
            $verticalSplitMode = '1';
            $verticalSplitPosition = (string) $ySplit;
        }

        $this->writeSplitValue($objWriter, 'HorizontalSplitMode', 'short', $horizontalSlitMode);
        $this->writeSplitValue($objWriter, 'HorizontalSplitPosition', 'int', $horizontalSlitPosition);
        $this->writeSplitValue($objWriter, 'PositionLeft', 'short', $positionLeft);
        $this->writeSplitValue($objWriter, 'PositionRight', 'short', $positionRight);

        for ($column = 'A'; $column !== $columnName; ++$column) {
            $worksheet->getColumnDimension($column)->setAutoSize(true);
        }

        $this->writeSplitValue($objWriter, 'VerticalSplitMode', 'short', $verticalSplitMode);
        $this->writeSplitValue($objWriter, 'VerticalSplitPosition', 'int', $verticalSplitPosition);
        $this->writeSplitValue($objWriter, 'PositionTop', 'short', $positionTop);
        $this->writeSplitValue($objWriter, 'PositionBottom', 'short', $positionBottom);

        $this->writeSplitValue($objWriter, 'ActiveSplitRange', 'short', '3');
    }
}
