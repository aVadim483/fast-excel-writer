<?php

namespace avadim\FastExcelWriter\Exceptions;

/**
 * Class Exception
 *
 * @package avadim\FastExcelWriter
 */
class Exception extends \RuntimeException
{
    public const ERROR_ADDRESS = 101;
    public const ERROR_FILE = 201;
    public const ERROR_RUNTIME = 901;

    protected static int $defaultCode = self::ERROR_RUNTIME;

    public static function throwNew($message, ...$args)
    {
        $parts = explode('\\', __NAMESPACE__);
        $namespace = $parts[0] . '\\' . $parts[1] . '\\';
        $stack = debug_backtrace();
        foreach ($stack as $point) {
            if (!isset($point['class']) || strpos($point['class'], $namespace) !== 0) {
                break;
            }
        }
        if (isset($point['file'], $point['line'])) {
            $message .= ' (called in ' . $point['file'] . ':' . $point['line'] . ')';
        }
        $class = get_called_class();
        throw new $class(sprintf($message, ...$args), self::$defaultCode);
    }

}

// EOF