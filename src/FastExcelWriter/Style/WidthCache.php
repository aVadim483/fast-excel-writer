<?php

namespace avadim\FastExcelWriter\Style;

/** Bounded FIFO cache used by automatic width calculations. */
class WidthCache
{
    protected array $values = [];
    protected array $queue = [];
    protected int $head = 0;
    protected int $tail = 0;
    protected int $bytes = 0;
    protected int $capacity;
    protected int $byteLimit;
    protected int $keyLimit;

    public function __construct(int $capacity = 4096, int $byteLimit = 1048576, int $keyLimit = 1024)
    {
        if ($capacity < 1 || $byteLimit < 1 || $keyLimit < 1) {
            throw new \InvalidArgumentException('Cache limits must be positive');
        }
        $this->capacity = $capacity;
        $this->byteLimit = $byteLimit;
        $this->keyLimit = $keyLimit;
    }

    public function get(string $key)
    {
        return $this->values[$key] ?? null;
    }

    public function put(string $key, $value): void
    {
        $length = strlen($key);
        if ($length > $this->keyLimit || $length > $this->byteLimit) {
            return;
        }
        if (array_key_exists($key, $this->values)) {
            $this->values[$key] = $value;
            return;
        }
        while (count($this->values) >= $this->capacity || $this->bytes + $length > $this->byteLimit) {
            $old = $this->queue[$this->head];
            unset($this->values[$old], $this->queue[$this->head]);
            $this->head = ($this->head + 1) % $this->capacity;
            $this->bytes -= strlen($old);
        }
        $this->values[$key] = $value;
        $this->queue[$this->tail] = $key;
        $this->tail = ($this->tail + 1) % $this->capacity;
        $this->bytes += $length;
    }
}
