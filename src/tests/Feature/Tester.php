<?php
namespace TestSheets\Tests\Feature;

use TestSheets\Tests\ExcelWriter\ExcelWriter;
use TestSheets\Tests\Feature\Traits\ValidatesForm;
use Tests\TestCase;

class Tester extends TestCase
{

    use ValidatesForm;

    public static $writer;
    public $implicit = true;
    public static function setUpBeforeClass(): void
    {
        parent::setUpBeforeClass();
        self::$writer = resolve(ExcelWriter::class);
        self::$writer::$row = 9;
    }

    public static function setUpSheet($name, $index)
    {
        $sheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet(self::$writer::$spreadsheet, $name);
        self::$writer::$project = config('sheets.project');
        self::$writer::$spreadsheet->addSheet($sheet, $index);
        self::$writer::$spreadsheet->setActiveSheetIndex($index);
        self::$writer::setUp($sheet);
    }

    public function setUp(): void
    {
        parent::setUp();
    }

    public function tearDown(): void
    {
        if ($this->implicit) {
            $this->writeRow($this->status()->isSuccess());
        }
        parent::tearDown();
    }

    public function writeRow($is_success): static
    {
        self::$writer
            ->setRowBorders()
            ->setRowNum()
            ->mergeCells()
            ->writeDate()
            ->writeResult($is_success)
            ->incrementRow()
            ->save();
        return $this;
    }

    public function event($event): static
    {
        self::$writer->write("D", $event);
        return $this;
    }

    public function item($item): static
    {
        self::$writer->write("J", $item);
        return $this;
    }

    public function conditions($conditions): static
    {
        if (is_array($conditions)) {
            $conditions = implode("\n", $conditions);
        }
        self::$writer->write("W", $conditions);
        return $this;
    }

    public function checks($checks): static
    {
        if (is_array($checks)) {
            $checks = implode("\n", $checks);
        }
        self::$writer->write("AL", $checks);
        return $this;
    }

    public function assignAll($event, $item, $conditions, $checks)
    {
        return $this->event($event)
            ->item($item)
            ->conditions($conditions)
            ->checks($checks);
    }

    public function validate($key, $rule, $conditions, $value="なし", $is_failure=true, $args=[])
    {
        $error = __("validation.$rule", array_merge(['attribute' => $key], $args));
        $check = $is_failure ? "表示される" : "表示されない";
        $this
            ->item('バリデーションエラーの確認')
            ->conditions("$conditions: $value")
            ->checks("「{$error}」というエラーが{$check}");
    }

    public function display(): static
    {
        $this->event("初期画面表示")
            ->item("表示内容の確認")
            ->checks(["レイアウト通り表示"]);
        return $this;
    }

    public function filter($col, $search_val): static
    {
        return $this->event("{$col}のフィルター")
            ->item("一覧テーブルのフィルター処理の確認")
            ->conditions("検索値：{$search_val}")
            ->checks("{$search_val}が入っている{$col}のみ表示される");
    }

    public function ordering($col, $click_count): static
    {
        $key = [1 => "昇順", 2 => "降順", 3 => "初期"];
        $this->conditions("{$col}: {$click_count}目のクリック")
            ->checks("{$col}の{$key[$click_count]}に並べる");
        return $this;
    }

    public function rescue($callback): static
    {
        try {
            $callback();
            $this->writeRow(true);
        } catch (\Throwable $th) {
            $this->writeRow(false);
            throw $th;
        }

        return $this;
    }

}