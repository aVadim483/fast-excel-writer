<?php

spl_autoload_register(static function ($class) {
    $namespace = 'avadim\\FastExcelWriter\\';
    if (0 === strpos($class, $namespace)) {
        $file = str_replace('\\', DIRECTORY_SEPARATOR, str_replace($namespace, '', $class) . '.php');
        include_once __DIR__ . DIRECTORY_SEPARATOR . 'FastExcelWriter' . DIRECTORY_SEPARATOR . $file;
    }
});

// EOF
