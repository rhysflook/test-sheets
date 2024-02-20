<?php
namespace TestSheets\Tests\ExcelWriter;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
class Borders {
    const BORDERS = [
        [
            'effect' => 'right',
            'cells' => "E1:E4,AS1:AS4,BA1:BA4",
            'style' => BORDER::BORDER_DOTTED
        ],
        [
            'effect' => 'bottom',
            'cells' => "AQ1:BF1,AQ3:BF3",
            'style' => BORDER::BORDER_DOTTED
        ],
        [
            'effect' => 'bottom',
            'cells' => "A1:L1,A2:L2,A3:L3,AQ2:BF2",
            'style' => BORDER::BORDER_THIN
        ],
        [
            'effect' => 'bottom',
            'cells' => "A5:AL5,A6:AL6",
            'style' => BORDER::BORDER_THIN
        ],
        [
            'effect' => 'right',
            'cells' => "L1:L4,AP1:AP4,AX1:AX4",
            'style' => BORDER::BORDER_THIN
        ],
        [
            'effect' => 'right',
            'cells' => "F6,AL6",
            'style' => BORDER::BORDER_THIN
        ],
        [
            'effect' => 'bottom',
            'cells' => "A4:BF4",
            'style' => BORDER::BORDER_MEDIUM
        ],
        [
            'effect' => 'right',
            'cells' => "BF1:BF4",
            'style' => BORDER::BORDER_MEDIUM
        ],
    ];

    public static function applyBorders($sheet)
    {
        foreach (self::BORDERS as $border) {
            self::apply($sheet, $border['cells'], $border['effect'], $border['style']);
        }
        self::setRowBorders($sheet, 7);
    }

    public static function apply($sheet, $cells, $effect, $style)
    {
        $cells = self::getCellList($cells);
        foreach ($cells as $cell) {
            $sheet->getStyle($cell)->applyFromArray(['borders' => [
                $effect => [
                    'borderStyle' => $style,
                ],
            ]]);
        }
    }

    public static function getCellList($cells): array
    {

        $ranges = explode(',', $cells);
        $cells = [];
        foreach ($ranges as $range) {
            $cells = array_merge($cells, Coordinate::extractAllCellReferencesInRange($range));
        }
        return $cells;
    }

    public static function setRowBorders($sheet, $row)
    {
        $bottom = $row + 1;
        $borders = [[
            'effect' => 'bottom',
            'cells' => "A$row:BF$row,A$bottom:BF$bottom",
            'style' => BORDER::BORDER_THIN
        ],
        [
            'effect' => 'right',
            'cells' => "C$bottom,I$bottom,V$bottom,AK$bottom,AZ$bottom,BD$bottom,BF$bottom",
            'style' => BORDER::BORDER_THIN
        ]];
        foreach ($borders as $border) {
            self::apply($sheet, $border['cells'], $border['effect'], $border['style']);
        }
    }
}