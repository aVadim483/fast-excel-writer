<?php

namespace avadim\FastExcelWriter\Writer;

/** @internal A date conversion shared by width measurement and one cell write. */
final class PreparedDateTime
{
    /** @var mixed */
    public $sourceValue;
    /** @var float|int|false */
    public $serial;
    public string $timezone;

    public function __construct($value)
    {
        $this->sourceValue = $value;
        $this->timezone = date_default_timezone_get();
        $this->serial = Writer::convertDateTime($value);
    }
}
