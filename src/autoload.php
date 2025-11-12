<?php

spl_autoload_register(static function ($class) {
    $namespace = 'avadim\\FastExcelWriter\\';
    if (0 === strpos($class, $namespace)) {
        $file = str_replace('\\', '/', __DIR__ . '/FastExcelWriter/' . str_replace($namespace, '', $class) . '.php');
        if (file_exists($file)) {
            include $file;
        }
    }
});

// EOF
