<?php

namespace avadim\FastExcelWriter\Style;

/** @internal Prepared geometry and lazily loaded metrics for one font. */
final class FontContext
{
    public array $font = [];
    public string $name;
    public float $size;
    public int $face;
    public string $prefix;
    public ?array $metrics = null;

    public function __construct(array $font)
    {
        $this->font = $font;
        // Keep PHP's cheap array identity check unless geometry contains mutable references.
        foreach (['name', 'size', 'style-bold', 'style-italic'] as $property) {
            if (array_key_exists($property, $font) && \ReflectionReference::fromArrayElement($font, $property) !== null) {
                $this->font = [];
                foreach ($font as $key => $value) {
                    $this->font[$key] = $value;
                }
                break;
            }
        }
        $this->name = $font['name'] ?? Font::DEFAULT_FONT_NAME;
        $this->size = (float)($font['size'] ?? Font::DEFAULT_FONT_SIZE);
        $this->face = (!empty($font['style-bold']) ? 1 : 0) + (!empty($font['style-italic']) ? 2 : 0);
        $this->prefix = 'T' . strlen($this->name) . ':' . $this->name . ':' . $this->size . ':' . $this->face . ':';
    }
}
