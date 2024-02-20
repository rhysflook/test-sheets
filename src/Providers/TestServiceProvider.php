<?php
namespace Src\Providers;
use Illuminate\Support\ServiceProvider;
use TestSheets\Console\Commands\FeatureTest;

class TestServiceProvider extends ServiceProvider {
    /**
     * Bootstrap any package services.
     */
    public function boot(): void
    {
        if ($this->app->runningInConsole()) {
            $this->commands([
                FeatureTest::class,
            ]);
        }
    }
}