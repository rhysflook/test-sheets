<?php
namespace TestSheets\Console\Commands;

use Illuminate\Console\Command;

class FeatureTest extends Command
{
    /**
     * コンソールコマンドの名前とシグネチャ
     *
     * @var string
     */
    protected $signature = 'test-sheets:feature-test {class} {--file=} {--sheet=}';

    /**
     * コンソールコマンドの説明
     *
     * @var string
     */
    protected $description = 'Setup required classes for feature test with excel generation.';

    /**
     * コンソールコマンドの実行
     */
    public function handle()
    {
        $class = $this->argument('class');
        $file = $this->option('file') ?? $class;
        $sheet = $this->option('sheet') ?? $class;
        $file_str = $this->getFileString($class, $file, $sheet);

        // write to file in laravel test directory
        $path = base_path("tests/Feature/$class.php");
        file_put_contents($path, $file_str);
    }

    public function getFileString($className, $fileName, $sheetName)
    {
        $file = "<?php\n";
        $file .= "namespace Tests\Feature;\n";
        $file .= "use TestSheets\Tests\Feature\Tester;\n";
        $file .= "use Illuminate\Foundation\Testing\RefreshDatabase;\n";
        $file .= "\n";
        $file .= "class $className extends Tester\n";
        $file .= "{\n";
        $file .= "    use RefreshDatabase;\n";
        $file .= "    public static function setUpBeforeClass(): void\n";
        $file .= "    {\n";
        $file .= "        parent::setUpBeforeClass();\n";
        $file .= "        self::\$writer::\$filename = \"$fileName\";\n";
        $file .= "        self::\$writer::\$page_name = \"$sheetName\";\n";
        $file .= "    }\n";
        $file .= "    public function setUp(): void\n";
        $file .= "    {\n";
        $file .= "        parent::setUp();\n";
        $file .= "        self::setUpSheet('一覧', 1);\n";
        $file .= "    }\n";
        $file .= "}\n";
        $file .= "\n";
        return $file;
    }

}