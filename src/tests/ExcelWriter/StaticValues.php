<?php
namespace TestSheets\Tests\ExcelWriter;

class StaticValues {
    const VALUES = [
        ['A1', 'プロジェクト名：'],
        ['A2', '開発工程：'],
        ['A3', 'サブシステム名：'],
        ['A4', '処理機能名：'],
        ['M1', 'タイトル：'],
        ['M2', '手動単体テスト仕様書'],
        ['AQ1', '作成日'],
        ['AQ2', '承認日'],
        ['AQ3', '改訂日'],
        ['AQ4', '承認日'],
        ['AY1', '作成者'],
        ['AY2', '承認者'],
        ['AY3', '改訂者'],
        ['AY4', '承認者'],
        ['A6', '画面名'],
        ['A8', 'No.'],
        ['D8', 'イベント'],
        ['J8', '項目'],
        ['W8', '入力条件'],
        ['AL8', '確認内容'],
        ['BA8', 'テスト完了日'],
        ['BE8', '確認'],
    ];

    public static function writeCells($sheet, $project, $page_name)
    {
        foreach (self::VALUES as [$cell, $value]) {
            $sheet->setCellValue($cell, $value);
        }
        $sheet->setCellValue('F1', $project);
        $sheet->setCellValue('G6', $page_name);
    }
}