<?php
namespace TestSheets\Tests\Feature\Traits;
trait ValidatesForm {
    public function validate($key, $rule, $conditions, $value="なし", $is_failure=true, $args=[])
    {
        $error = __("validation.$rule", array_merge(['attribute' => $key], $args));
        $check = $is_failure ? "表示される" : "表示されない";
        $this
            ->item('バリデーションエラーの確認')
            ->conditions("$conditions: $value")
            ->checks("「{$error}」というエラーが{$check}");
    }

    public function required($keys, $component, $action="check") {
        if (is_array($keys)) {
            foreach ($keys as $key) {
                $this->rescue(function () use ($key, $component, $action) {
                    $this->validate($key, 'required', $key);
                    $component->call($action)
                        ->assertSee("{$key}は必須です。");
                });
            }
            return;
        }
        $this->validate($keys, 'required', $keys);
        $component->call($action)
            ->assertSee("{$keys}は必須です。");
    }
}