<?php

namespace avadim\FastExcelWriter;

/**
 * Class Options
 *
 * @package avadim\FastExcelWriter
 */
class Options implements \ArrayAccess, \IteratorAggregate
{
    protected array $options = [];

    /**
     * Options constructor
     *
     * @param array $options
     */
    public function __construct(array $options = [])
    {
        $this->options = $options;
    }

    /**
     * Create Options instance
     *
     * @param array $options
     *
     * @return Options
     */
    public static function create(array $options = []): Options
    {
        return new self($options);
    }

    /**
     * Set temporary directory
     *
     * @param string $tempDir
     *
     * @return $this
     */
    public function tempDir(string $tempDir): Options
    {
        $this->options['temp_dir'] = $tempDir;

        return $this;
    }

    /**
     * Set prefix for temporary files
     *
     * @param string $tempPrefix
     *
     * @return $this
     */
    public function tempPrefix(string $tempPrefix): Options
    {
        $this->options['temp_prefix'] = $tempPrefix;

        return $this;
    }

    /**
     * Set auto conversion for numbers
     *
     * @param bool $autoConvertNumber
     *
     * @return $this
     */
    public function autoConvertNumber(bool $autoConvertNumber = true): Options
    {
        $this->options['auto_convert_number'] = $autoConvertNumber;

        return $this;
    }

    /**
     * Use shared strings
     *
     * @param bool $sharedString
     *
     * @return $this
     */
    public function sharedString(bool $sharedString = true): Options
    {
        $this->options['shared_string'] = $sharedString;

        return $this;
    }

    /**
     * Set locale
     *
     * @param string $locale
     *
     * @return $this
     */
    public function locale(string $locale): Options
    {
        $this->options['locale'] = $locale;

        return $this;
    }

    /**
     * Set default font
     *
     * @param array $fontOptions
     *
     * @return $this
     */
    public function defaultFont(array $fontOptions): Options
    {
        $this->options['default_font'] = $fontOptions;

        return $this;
    }

    /**
     * Set writer class
     *
     * @param string $writerClass
     *
     * @return $this
     */
    public function writerClass(string $writerClass): Options
    {
        $this->options['writer_class'] = $writerClass;

        return $this;
    }

    /**
     * Set style manager class
     *
     * @param string $styleManagerClass
     *
     * @return $this
     */
    public function styleManagerClass(string $styleManagerClass): Options
    {
        $this->options['style_manager'] = $styleManagerClass;

        return $this;
    }

    /**
     * Return options as array
     *
     * @return array
     */
    public function toArray(): array
    {
        return $this->options;
    }

    // ArrayAccess implementation

    /**
     * Whether an offset exists
     *
     * @param mixed $offset
     *
     * @return bool
     */
    public function offsetExists($offset): bool
    {
        return isset($this->options[$offset]);
    }

    /**
     * Offset to retrieve
     *
     * @param mixed $offset
     *
     * @return mixed
     */
    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        return $this->options[$offset] ?? null;
    }

    /**
     * Offset to set
     *
     * @param mixed $offset
     * @param mixed $value
     */
    public function offsetSet($offset, $value): void
    {
        if (is_null($offset)) {
            $this->options[] = $value;
        } else {
            $this->options[$offset] = $value;
        }
    }

    /**
     * Offset to unset
     *
     * @param mixed $offset
     */
    public function offsetUnset($offset): void
    {
        unset($this->options[$offset]);
    }

    // IteratorAggregate implementation

    /**
     * Retrieve an external iterator
     *
     * @return \Traversable
     */
    public function getIterator(): \Traversable
    {
        return new \ArrayIterator($this->options);
    }
}
