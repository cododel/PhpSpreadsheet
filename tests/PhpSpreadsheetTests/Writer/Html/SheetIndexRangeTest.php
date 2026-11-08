<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Writer\Html;

use DOMDocument;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Shared\File;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Writer\Pdf;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Dompdf;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Tcpdf;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;

class SheetIndexRangeTest extends TestCase
{
    private function threeSheets(): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $first = $spreadsheet->getActiveSheet();
        $first->setTitle('First');
        $first->setCellValue('A1', 'alpha');
        $second = $spreadsheet->createSheet();
        $second->setTitle('Second');
        $second->setCellValue('A1', 'beta');
        $second->mergeCells('A1:B1');
        $third = $spreadsheet->createSheet();
        $third->setTitle('Third');
        $third->setCellValue('A1', 'gamma');
        $third->mergeCells('A1:C1');

        return $spreadsheet;
    }

    public function testSelectedSheetsAreWrittenInOrder(): void
    {
        $spreadsheet = $this->threeSheets();
        $writer = new Html($spreadsheet);
        $writer->setSheetIndex([2, 0]);
        self::assertSame([2, 0], $writer->getSheetIndex());

        $html = $writer->generateHTMLAll();
        self::assertStringContainsString('gamma', $html);
        self::assertStringContainsString('alpha', $html);
        self::assertStringNotContainsString('beta', $html);
        self::assertStringContainsString('colspan="3"', $html);
        self::assertStringNotContainsString('colspan="2"', $html);
        self::assertStringContainsString('table.sheet2', $html);
        self::assertStringContainsString('table.sheet0', $html);
        self::assertStringNotContainsString('table.sheet1', $html);

        $eol = $writer->getLineEnding();
        $expectedNav = '<ul class="navigation">'
            . $eol
            . '  <li class="sheet2"><a href="#sheet2">Third</a></li>'
            . $eol
            . '  <li class="sheet0"><a href="#sheet0">First</a></li>'
            . $eol
            . '</ul>';
        self::assertStringContainsString($expectedNav, $html);

        $dom = new DOMDocument();
        $dom->loadHTML($html);
        $tables = $dom->getElementsByTagName('table');
        self::assertCount(2, $tables);
        self::assertSame('sheet2', $tables->item(0)?->getAttribute('id'));
        self::assertSame('sheet0', $tables->item(1)?->getAttribute('id'));
        self::assertSame(0, $spreadsheet->getActiveSheetIndex());
        $spreadsheet->disconnectWorksheets();
    }

    public function testDuplicateIndexesAreWrittenAgain(): void
    {
        $spreadsheet = $this->threeSheets();
        $writer = new Html($spreadsheet);
        $writer->setSheetIndex([2, 0, 2]);
        $html = $writer->generateHTMLAll();
        self::assertSame(2, substr_count($html, "id='sheet2'"));
        self::assertSame(1, substr_count($html, "id='sheet0'"));
        self::assertSame(0, substr_count($html, "id='sheet1'"));
        self::assertLessThan(
            strpos($html, "id='sheet0'"),
            strpos($html, "id='sheet2'")
        );
        self::assertSame(2, substr_count($html, 'gamma'));
        self::assertStringNotContainsString('beta', $html);
        $spreadsheet->disconnectWorksheets();
    }

    public function testIntegerSelectionIsUnchanged(): void
    {
        $spreadsheet = $this->threeSheets();
        $sheet = $spreadsheet->getSheet(2);
        $sheet->fromArray([
            [1, 2, 3, 4, 5],
            [6, 7, 8, 9, 10],
            [11, 12, 13, 14, 15],
            [16, 17, 18, 19, 20],
        ]);
        $sheet->getPageSetup()->setPrintArea('B2:D3');

        $single = new Html($spreadsheet);
        $single->setSheetIndex(2);
        $singleHtml = $single->generateHTMLAll();
        self::assertStringContainsString('@page page0 {', $singleHtml);
        self::assertStringContainsString('table.sheet0 tr.row0 td { display:none }', $singleHtml);
        self::assertStringContainsString("id='sheet2'", $singleHtml);
        self::assertStringContainsString("style='page: page2'", $singleHtml);
        self::assertStringNotContainsString('<ul class="navigation">', $singleHtml);

        $selected = new Html($spreadsheet);
        $selected->setSheetIndex([2]);
        $selectedHtml = $selected->generateHTMLAll();
        self::assertStringContainsString('@page page2 {', $selectedHtml);
        self::assertStringNotContainsString('@page page0 {', $selectedHtml);
        self::assertStringContainsString('table.sheet2 tr.row0 td { display:none }', $selectedHtml);
        self::assertStringNotContainsString('table.sheet0', $selectedHtml);
        self::assertStringContainsString("id='sheet2'", $selectedHtml);
        self::assertStringContainsString("style='page: page2'", $selectedHtml);
        self::assertStringNotContainsString('beta', $selectedHtml);
        self::assertStringNotContainsString('alpha', $selectedHtml);
        $spreadsheet->disconnectWorksheets();
    }

    public function testSelectedSheetsHonorRightToLeft(): void
    {
        $spreadsheet = $this->threeSheets();
        $spreadsheet->getSheet(0)->setRightToLeft(true);
        $spreadsheet->getSheet(2)->setRightToLeft(false);

        $writer = new Html($spreadsheet);
        $writer->setSheetIndex([0, 2]);
        $html = $writer->generateHTMLAll();
        self::assertSame(1, substr_count($html, "dir='rtl'"));
        self::assertSame(1, substr_count($html, "dir='ltr'"));
        self::assertStringNotContainsString('beta', $html);

        $onlyRtl = new Html($spreadsheet);
        $onlyRtl->setSheetIndex([0]);
        $rtlHtml = $onlyRtl->generateHTMLAll();
        self::assertSame(2, substr_count($rtlHtml, "dir='rtl'"));
        self::assertSame(0, substr_count($rtlHtml, "dir='ltr'"));
        $spreadsheet->disconnectWorksheets();
    }

    public function testTcpdfPageBreakBetweenSelectedSheets(): void
    {
        $spreadsheet = $this->threeSheets();
        $writer = new Tcpdf($spreadsheet);
        $writer->setSheetIndex([0, 2]);
        $html = $writer->generateHTMLAll();
        self::assertSame(1, substr_count($html, 'page-break-before:always'));
        self::assertStringContainsString('alpha', $html);
        self::assertStringContainsString('gamma', $html);
        self::assertStringNotContainsString('beta', $html);

        $all = new Tcpdf($spreadsheet);
        $all->writeAllSheets();
        $allHtml = $all->generateHTMLAll();
        self::assertSame(2, substr_count($allHtml, 'page-break-before:always'));
        $spreadsheet->disconnectWorksheets();
    }

    public function testTcpdfSingleSheetHasNoExtraPageBreak(): void
    {
        $spreadsheet = $this->threeSheets();
        $byIndex = new Tcpdf($spreadsheet);
        $byIndex->setSheetIndex(1);
        self::assertSame(0, substr_count($byIndex->generateHTMLAll(), 'page-break-before:always'));

        $byList = new Tcpdf($spreadsheet);
        $byList->setSheetIndex([1]);
        $html = $byList->generateHTMLAll();
        self::assertSame(0, substr_count($html, 'page-break-before:always'));
        self::assertStringContainsString('beta', $html);
        self::assertStringNotContainsString('alpha', $html);
        $spreadsheet->disconnectWorksheets();
    }

    public function testPageSetupFollowsFirstSelectedSheet(): void
    {
        $spreadsheet = $this->threeSheets();
        $spreadsheet->getSheet(0)->getPageSetup()->setOrientation(PageSetup::ORIENTATION_PORTRAIT);
        $spreadsheet->getSheet(2)->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);

        $probe = new class ($spreadsheet) extends Html {
            public function pageSetupSheetIndex(): int
            {
                return $this->getPageSetupSheetIndex();
            }
        };
        self::assertSame(0, $probe->pageSetupSheetIndex());
        $probe->setSheetIndex([2, 0]);
        self::assertSame(2, $probe->pageSetupSheetIndex());
        $probe->setSheetIndex([]);
        self::assertSame(0, $probe->pageSetupSheetIndex());
        $probe->writeAllSheets();
        self::assertNull($probe->getSheetIndex());
        self::assertSame(0, $probe->pageSetupSheetIndex());
        $probe->setSheetIndex(1);
        self::assertSame(1, $probe->pageSetupSheetIndex());

        $writer = new Mpdf($spreadsheet);
        $writer->setSheetIndex([2, 0]);
        $html = $writer->generateHTMLAll();
        self::assertStringContainsString('@page page2 {', $html);
        self::assertStringContainsString('@page page0 {', $html);
        self::assertStringNotContainsString('@page page1 {', $html);
        self::assertSame(1, substr_count($html, 'size: landscape;'));
        self::assertSame(1, substr_count($html, 'size: portrait;'));
        $spreadsheet->disconnectWorksheets();
    }

    public function testSheetIndexApi(): void
    {
        $spreadsheet = $this->threeSheets();
        $writer = new Html($spreadsheet);
        self::assertSame(0, $writer->getSheetIndex());
        $chained = $writer->setSheetIndex(['later' => 2, 'earlier' => 0]);
        self::assertSame($writer, $chained);
        self::assertSame([2, 0], $writer->getSheetIndex());
        $writer->writeAllSheets();
        self::assertNull($writer->getSheetIndex());
        $writer->setSheetIndex(1);
        self::assertSame(1, $writer->getSheetIndex());
        $spreadsheet->disconnectWorksheets();
    }

    public function testOutOfRangeIndexMatchesSingleSheet(): void
    {
        $spreadsheet = $this->threeSheets();
        $single = new Html($spreadsheet);
        $single->setSheetIndex(5);

        try {
            $single->generateHTMLAll();
            self::fail('An out-of-range sheet index must be rejected');
        } catch (PhpSpreadsheetException $exception) {
            self::assertSame(
                'Your requested sheet index: 5 is out of bounds. The actual number of sheets is 3.',
                $exception->getMessage()
            );
        }

        $selected = new Html($spreadsheet);
        $selected->setSheetIndex([0, 5]);
        $this->expectException(PhpSpreadsheetException::class);
        $this->expectExceptionMessage('Your requested sheet index: 5 is out of bounds. The actual number of sheets is 3.');
        $selected->generateHTMLAll();
    }

    public function testNegativeIndexIsOutOfBounds(): void
    {
        $spreadsheet = $this->threeSheets();
        $writer = new Html($spreadsheet);
        $writer->setSheetIndex([0, -1]);
        $this->expectException(PhpSpreadsheetException::class);
        $this->expectExceptionMessage('Your requested sheet index: -1 is out of bounds. The actual number of sheets is 3.');
        $writer->generateHTMLAll();
    }

    public function testNonIntegerSheetIndexIsRejected(): void
    {
        $spreadsheet = $this->threeSheets();
        $writer = new Html($spreadsheet);
        $this->expectException(WriterException::class);
        $this->expectExceptionMessage('Sheet index must be an integer, string given');
        $writer->setSheetIndex([0, '1']); // @phpstan-ignore argument.type (deliberate testing runtime)
    }

    public function testEmptySelectionWritesNoSheets(): void
    {
        $spreadsheet = $this->threeSheets();
        $writer = new Html($spreadsheet);
        $writer->setSheetIndex([]);
        self::assertSame([], $writer->getSheetIndex());
        $html = $writer->generateHTMLAll();
        self::assertStringNotContainsString("id='sheet", $html);
        self::assertStringNotContainsString('<ul class="navigation">', $html);
        self::assertStringNotContainsString('alpha', $html);
        self::assertStringNotContainsString('beta', $html);
        self::assertStringNotContainsString('gamma', $html);
        $spreadsheet->disconnectWorksheets();
    }

    /** @return array<string, array{0: class-string<Pdf>}> */
    public static function pdfWriterProvider(): array
    {
        return [
            'mpdf' => [Mpdf::class],
            'dompdf' => [Dompdf::class],
            'tcpdf' => [Tcpdf::class],
        ];
    }

    /** @param class-string<Pdf> $writerClass */
    #[DataProvider('pdfWriterProvider')]
    public function testPdfWriterSavesSelectedSheets(string $writerClass): void
    {
        $spreadsheet = $this->threeSheets();
        $spreadsheet->getSheet(2)->getPageSetup()
            ->setOrientation(PageSetup::ORIENTATION_LANDSCAPE)
            ->setPaperSize(PageSetup::PAPERSIZE_A4);
        $writer = new $writerClass($spreadsheet);
        $writer->setSheetIndex([2, 0]);
        $path = File::temporaryFilename();

        try {
            $writer->save($path);
            $contents = file_get_contents($path);
            self::assertIsString($contents);
            self::assertStringStartsWith('%PDF', $contents);
        } finally {
            if (is_file($path)) {
                unlink($path);
            }
        }
        $spreadsheet->disconnectWorksheets();
    }
}
