<?php

namespace avadim\FastExcelWriter\Charts;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class GridLines extends Properties
{
    private bool $objectState = false;

    private array $lineProperties = [
        'color' => [
            'type' => self::EXCEL_COLOR_TYPE_STANDARD,
            'value' => null,
            'alpha' => 0
        ],
        'style' => [
            'width' => '9525',
            'compound' => self::LINE_STYLE_COMPOUND_SIMPLE,
            'dash' => self::LINE_STYLE_DASH_SOLID,
            'cap' => self::LINE_STYLE_CAP_FLAT,
            'join' => self::LINE_STYLE_JOIN_BEVEL,
            'arrow' => [
                'head' => [
                    'type' => self::LINE_STYLE_ARROW_TYPE_NOARROW,
                    'size' => self::LINE_STYLE_ARROW_SIZE_5
                ],
                'end' => [
                    'type' => self::LINE_STYLE_ARROW_TYPE_NOARROW,
                    'size' => self::LINE_STYLE_ARROW_SIZE_8
                ],
            ]
        ]
    ];

    private array $shadowProperties = [
        'presets' => self::SHADOW_PRESETS_NOSHADOW,
        'effect' => null,
        'color' => [
            'type' => self::EXCEL_COLOR_TYPE_STANDARD,
            'value' => 'black',
            'alpha' => 85,
        ],
        'size' => [
            'sx' => null,
            'sy' => null,
            'kx' => null
        ],
        'blur' => null,
        'direction' => null,
        'distance' => null,
        'algn' => null,
        'rotWithShape' => null
    ];

    private array $glowProperties = [
        'size' => null,
        'color' => [
            'type' => self::EXCEL_COLOR_TYPE_STANDARD,
            'value' => 'black',
            'alpha' => 40
        ]
    ];

    private array $softEdges = [
        'size' => null
    ];

    /**
     * Get Object State
     *
     * @return bool
     */
    public function getObjectState(): bool
    {
        return $this->objectState;
    }

    /**
     * Change Object State to True
     */
    private function activateObject()
    {
        $this->objectState = true;
    }

    /**
     * Set Line Color Properties
     *
     * @param string $value
     * @param int $alpha
     * @param string $type
     */
    public function setLineColorProperties(string $value, int $alpha = 0, string $type = self::EXCEL_COLOR_TYPE_STANDARD)
    {
        $this->activateObject();
        $this->lineProperties['color'] = $this->setColorProperties($value, $alpha, $type);
    }

    /**
     * Set Line Color Properties
     *
     * @param float|null $lineWidth
     * @param string|null $compoundType
     * @param string|null $dashType
     * @param string|null $capType
     * @param string|null $joinType
     * @param string|null $headArrowType
     * @param string|null $headArrowSize
     * @param string|null $endArrowType
     * @param string|null $endArrowSize
     */
    public function setLineStyleProperties(?float $lineWidth = null, ?string $compoundType = null, ?string $dashType = null,
                                           ?string $capType = null, ?string $joinType = null, ?string $headArrowType = null,
                                           ?string $headArrowSize = null, ?string $endArrowType = null, ?string $endArrowSize = null)
    {
        $this->activateObject();
        $this->lineProperties['style']['width'] = ($lineWidth !== null) ? Properties::excelPointsWidth($lineWidth) : null;
        $this->lineProperties['style']['compound'] = $compoundType;
        $this->lineProperties['style']['dash'] = $dashType ;
        $this->lineProperties['style']['cap'] = $capType;
        $this->lineProperties['style']['join'] = $joinType;
        $this->lineProperties['style']['arrow']['head']['type'] = $headArrowType;
        $this->lineProperties['style']['arrow']['head']['size'] = $headArrowSize;
        $this->lineProperties['style']['arrow']['end']['type'] = $endArrowType;
        $this->lineProperties['style']['arrow']['end']['size'] = $endArrowSize;

    }

    /**
     * Get Line Color Property
     *
     * @param string $parameter
     *
     * @return string|null
     */
    public function getLineColorProperty(string $parameter): ?string
    {
        return $this->lineProperties['color'][$parameter] ?? null;
    }

    /**
     * Get Line Style Property
     *
     * @param    array|string $elements
     *
     * @return string|null
     */
    public function getLineStyleProperty($elements): ?string
    {
        return $this->getArrayElementsValue($this->lineProperties['style'], $elements) ?? null;
    }

    /**
     * Set Glow Properties
     *
     * @param float $size
     * @param string|null $colorValue
     * @param int|null $colorAlpha
     * @param string|null $colorType
     */
    public function setGlowProperties(float $size, ?string $colorValue = null, ?int $colorAlpha = null, ?string $colorType = null)
    {
        $this->activateObject();
        $this->setGlowSize($size);
        $this->setGlowColor($colorValue, $colorAlpha, $colorType);
    }

    /**
     * Get Glow Color Property
     *
     * @param string $property
     *
     * @return string|null
     */
    public function getGlowColor(string $property): ?string
    {
        return $this->glowProperties['color'][$property] ?? null;
    }

    /**
     * Get Glow Size
     *
     * @return string|null
     */
    public function getGlowSize(): ?string
    {
        return $this->glowProperties['size'] ?? null;
    }

    /**
     * Set Glow Size
     *
     * @param float $size
     */
    private function setGlowSize(float $size)
    {
        $this->glowProperties['size'] = Properties::excelPointsWidth((float) $size);
    }

    /**
     * Set Glow Color
     *
     * @param string|null $color
     * @param int|null $alpha
     * @param string|null $type
     */
    private function setGlowColor(?string $color, ?int $alpha, ?string $type)
    {
        if ($color !== null) {
            $this->glowProperties['color']['value'] = (string) $color;
        }
        if ($alpha !== null) {
            $this->glowProperties['color']['alpha'] = $this->trueAlpha((int) $alpha);
        }
        if ($type !== null) {
            $this->glowProperties['color']['type'] = (string) $type;
        }
    }

    /**
     * Get Line Style Arrow Parameters
     *
     * @param string $arrowSelector
     * @param string $propertySelector
     *
     * @return string
     */
    public function getLineStyleArrowParameters(string $arrowSelector, string $propertySelector): string
    {
        return $this->getLineStyleArrowSize($this->lineProperties['style']['arrow'][$arrowSelector]['size'], $propertySelector);
    }

    /**
     * Set Shadow Properties
     *
     * @param int $shPresets
     * @param string|null $shColorValue
     * @param string|null $shColorType
     * @param int|null $shColorAlpha
     * @param float|null $shBlur
     * @param int|null $shAngle
     * @param float|null $shDistance
     *
     * @return $this
     */
    public function setShadowProperties(int  $shPresets, ?string $shColorValue = null, ?string $shColorType = null,
                                        ?int $shColorAlpha = null, float $shBlur = null, ?int $shAngle = null,
                                        ?float $shDistance = null): GridLines
    {
        $this->activateObject();
        $this->setShadowPresetsProperties($shPresets);
        $this->setShadowColor(
                ($shColorValue === null) ? $this->shadowProperties['color']['value'] : $shColorValue,
                ($shColorAlpha === null) ? (int) $this->shadowProperties['color']['alpha'] : $this->trueAlpha($shColorAlpha),
                ($shColorType === null) ? $this->shadowProperties['color']['type'] : $shColorType
            );
        $this->setShadowBlur($shBlur);
        $this->setShadowAngle($shAngle);
        $this->setShadowDistance($shDistance);

        return $this;
    }

    /**
     * Set Shadow Presets Properties
     *
     * @param int $shadowPresets
     */
    private function setShadowPresetsProperties(int $shadowPresets)
    {
        $this->shadowProperties['presets'] = $shadowPresets;
        $this->setShadowPropertiesMapValues($this->getShadowPresetsMap($shadowPresets));
    }


    /**
     * Set Shadow Color
     *
     * @param string|null $color
     * @param int|null $alpha
     * @param string|null $type
     */
    private function setShadowColor(?string $color, ?int $alpha, ?string $type)
    {
        if ($color !== null) {
            $this->shadowProperties['color']['value'] = (string) $color;
        }
        if ($alpha !== null) {
            $this->shadowProperties['color']['alpha'] = $this->trueAlpha((int) $alpha);
        }
        if ($type !== null) {
            $this->shadowProperties['color']['type'] = (string) $type;
        }
    }

    /**
     * Set Shadow Blur
     *
     * @param float|null $blur
     */
    private function setShadowBlur(?float $blur)
    {
        if ($blur !== null) {
            $this->shadowProperties['blur'] = (string) Properties::excelPointsWidth($blur);
        }
    }

    /**
     * Set Shadow Angle
     *
     * @param int|null $angle
     */
    private function setShadowAngle(?int $angle)
    {
        if ($angle !== null) {
            $this->shadowProperties['direction'] = (string) $this->excelPointsAngle($angle);
        }
    }

    /**
     * Set Shadow Distance
     *
     * @param float|null $distance
     */
    private function setShadowDistance(?float $distance)
    {
        if ($distance !== null) {
            $this->shadowProperties['distance'] = (string) Properties::excelPointsWidth($distance);
        }
    }

    /**
     * Get Shadow Property
     *
     * @param string|array $elements
     *
     * @return string|null
     */
    public function getShadowProperty($elements): ?string
    {
        return $this->getArrayElementsValue($this->shadowProperties, $elements) ?? null;
    }

    /**
     * Set Soft Edges Size
     *
     * @param float|null $size
     */
    public function setSoftEdgesSize(?float $size)
    {
        if (!is_null($size)) {
            $this->activateObject();
            $this->softEdges['size'] = (string) Properties::excelPointsWidth($size);
        }
    }

    /**
     * Get Soft Edges Size
     *
     * @return string|null
     */
    public function getSoftEdgesSize(): ?string
    {
        return $this->softEdges['size'] ?? null;
    }

}