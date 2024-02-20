<?php
namespace TestSheets\Tests\ExcelWriter;

use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPUnit\Framework\TestStatus\Success;

class ExcelWriter {
    public static $row = 9;
    public $merges = [
        ['A', 'C'],
        ['D', 'I'],
        ['J', 'V'],
        ['W', 'AK'],
        ['AL', 'AZ'],
        ['BA', 'BD'],
        ['BE', 'BF'],
    ];
    public static $writer;
    public static $spreadsheet;
    public static $sheet;
    public static $page_name;
    public static $project;
    public static $filename;
    public function __construct()
    {
        \PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder( new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder() );
        if (!self::$spreadsheet) !self::$spreadsheet = new Spreadsheet();
    }

    public static function setUp($sheet): void
    {
        self::$sheet = $sheet;

        self::$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(4);
        self::$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
        self::$spreadsheet->getDefaultStyle()->getFont()->setName("ＭＳ Ｐゴシック");
        self::$spreadsheet->getDefaultStyle()->getFont()->setSize(9);
        self::$sheet->mergeCells("M2:AP3");
        self::$sheet->getStyle("M2:AP3")->getFont()->setSize(16);
        self::$sheet->getStyle("M2:AP3")->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        self::$sheet->getStyle("M2:AP3")->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        self::setColors();
        Borders::applyBorders(self::$sheet);
        StaticValues::writeCells(self::$sheet, self::$project, self::$page_name);
    }

    public static function setColors(): void
    {
        self::setFill('M1:AP4', 'FFFF99');
        self::setFill('A6:F6', 'CCFFCC');
        self::setFill('A8:BF8', 'CCFFCC');
    }

    public static function setFill($range, $color)
    {
        self::$sheet->getStyle($range)->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()->setRGB($color);
    }

    public function mergeCells(): static
    {
        foreach ($this->merges as $cells) {
            $row = self::$row;
            self::$sheet->mergeCells("{$cells[0]}{$row}:{$cells[1]}{$row}");
        }
        return $this;
    }

    public function setRowBorders(): static
    {
        Borders::setRowBorders(self::$sheet, self::$row - 1);
        return $this;
    }

    public function setRowNum(): static
    {
        $row = self::$row;
        self::$sheet->setCellValue("A{$row}", self::$row - 8);
        return $this;
    }

    public function incrementRow(): static
    {
        self::$row++;
        return $this;
    }

    public function writeDate(): static
    {
        self::$sheet->getStyle("BA" . self::$row)
        ->getNumberFormat()
        ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
        $date = Carbon::now()->format('Y/m/d');
        $this->write('BA', $date);
        return $this;
    }

    public function writeResult($is_success): static
    {
        $result = $is_success ? 'OK' : 'NG';
        $this->write('BE', $result);
        return $this;
    }

    public function write($col, $value): static
    {
        $row = self::$spreadsheet->getActiveSheet()->getRowDimension(self::$row);
        $new_height = self::calcHeight($value);
        if ($row->getRowHeight() < $new_height) {
            $row->setRowHeight(self::calcHeight($value));
        }
        self::$sheet->getStyle($col . self::$row)->getAlignment()->setWrapText(true);
        self::$sheet->setCellValue($col . self::$row, $value);
        return $this;
    }

    public static function calcHeight($value)
    {
        $multiplier = 0;
        $sub_strings = explode("\n", $value);
        foreach ($sub_strings as $line) {
            $multiplier += ceil(strlen($line) / 102);
        }

        return $multiplier * 12.75;
    }

    public function save(): static
    {
        $writer = new Xlsx(self::$spreadsheet);
        $writer->save("tests/".self::$filename.".xlsx");
        return $this;
    }
}