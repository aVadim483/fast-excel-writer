<?php

namespace avadim\FastExcelWriter\Charts;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class Axis extends Properties
{
    /**
     * Axis Number
     *
     * @var  array of mixed
     */
    private array $axisNumber = [
        'format' => self::FORMAT_CODE_GENERAL,
        'source_linked' => 1
    ];

    /**
     * Axis Options
     *
     * @var  array of mixed
     */
    private array $axisOptions = [
        'minimum' => null,
        'maximum' => null,
        'major_unit' => null,
        'minor_unit' => null,
        'orientation' => self::ORIENTATION_NORMAL,
        'minor_tick_mark' => self::TICK_MARK_NONE,
        'major_tick_mark' => self::TICK_MARK_NONE,
        'axis_labels' => self::AXIS_LABELS_NEXT_TO,
        'horizontal_crosses' => self::HORIZONTAL_CROSSES_AUTOZERO,
        'horizontal_crosses_value' => null
    ];

    /**
     * Fill Properties
     *
     * @var  array of mixed
     */
    private array $fillProperties = [
        'type' => self::EXCEL_COLOR_TYPE_ARGB,
        'value' => null,
        'alpha' => 0
    ];

    /**
     * Line Properties
     *
     * @var  array of mixed
     */
    private array $lineProperties = [
        'type' => self::EXCEL_COLOR_TYPE_ARGB,
        'value' => null,
        'alpha' => 0
    ];

    /**
     * Line Style Properties
     *
     * @var  array of mixed
     */
    private array $lineStyleProperties = [
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
    ];

    /**
     * Shadow Properties
     *
     * @var  array of mixed
     */
    private array $shadowProperties = [
        'presets' => self::SHADOW_PRESETS_NOSHADOW,
        'effect' => null,
        'color' => [
            'type' => self::EXCEL_COLOR_TYPE_STANDARD,
            'value' => 'black',
            'alpha' => 40,
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

    /**
     * Glow Properties
     *
     * @var  array of mixed
     */
    private array $glowProperties = [
        'size' => null,
        'color' => [
            'type' => self::EXCEL_COLOR_TYPE_STANDARD,
            'value' => 'black',
            'alpha' => 40
        ]
    ];

    /**
     * Soft Edge Properties
     *
     * @var  array of mixed
     */
    private array $softEdges = [
        'size' => null
    ];

    private string $orientation;


    /**
     * Axis constructor
     */
    public function __construct()
    {
    }

    /**
     * Set Axis Number Properties
     *
     * @param string $formatCode
     *
     * @return $this
     */
    public function setAxisNumberProperties(string $formatCode): Axis
    {
        $this->axisNumber['format'] = $formatCode;
        $this->axisNumber['source_linked'] = 0;
        
        return $this;
    }

    /**
     * Get Axis Number Format Data Type
     *
     * @return string
     */
    public function getAxisNumberFormat(): string
    {
        return (string)$this->axisNumber['format'];
    }

    /**
     * Get Axis Number Source Linked
     *
     * @return string
     */
    public function getAxisNumberSourceLinked(): string
    {
        return (string)$this->axisNumber['source_linked'];
    }

    /**
     * Set Axis Options Properties
     *
     * @param string $axisLabels
     * @param string|null $horizontalCrossesValue
     * @param string|null $horizontalCrosses
     * @param string|null $axisOrientation
     * @param string|null $majorTmt
     * @param string|null $minorTmt
     * @param string|null $minimum
     * @param string|null $maximum
     * @param string|null $majorUnit
     * @param string|null $minorUnit
     * 
     * @return $this
     */
    public function setAxisOptionsProperties(string $axisLabels, ?string $horizontalCrossesValue = null,
                                             ?string $horizontalCrosses = null, ?string $axisOrientation = null,
                                             ?string $majorTmt = null, ?string $minorTmt = null, ?string $minimum = null,
                                             ?string $maximum = null, ?string $majorUnit = null, ?string $minorUnit = null): Axis
    {
        $this->axisOptions['axis_labels'] = $axisLabels ?: null;
        $this->axisOptions['horizontal_crosses_value'] = $horizontalCrossesValue ?: null;
        $this->axisOptions['horizontal_crosses'] = $horizontalCrosses ?: null;
        $this->axisOptions['orientation'] = $axisOrientation ?: null;
        $this->axisOptions['major_tick_mark'] = $majorTmt ?: null;
        $this->axisOptions['minor_tick_mark'] = $minorTmt ?: null;
        $this->axisOptions['minimum'] = $minimum ?: null;
        $this->axisOptions['maximum'] = $maximum ?: null;
        $this->axisOptions['major_unit'] = $majorUnit ?: null;
        $this->axisOptions['minor_unit'] = $minorUnit ?: null;

        return $this;
    }

    /**
     * Get Axis Options Property
     *
     * @param string $property
     *
     * @return string|null
     */
    public function getAxisOptionsProperty(string $property): ?string
    {
        return $this->axisOptions[$property] ?? null;
    }

    /**
     * Set Axis Orientation Property
     *
     * @param string $orientation
     *
     * @return $this
     */
    public function setAxisOrientation(string $orientation): Axis
    {
        $this->orientation = (string)$orientation;

        return $this;
    }

    /**
     * Set Fill Property
     *
     * @param string $color
     * @param int $alpha
     * @param string $type
     *
     * @return $this
     */
    public function setFillParameters(string $color, int $alpha = 0, string $type = self::EXCEL_COLOR_TYPE_ARGB): Axis
    {
        $this->fillProperties = $this->setColorProperties($color, $alpha, $type);

        return $this;
    }

    /**
     * Set Line Property
     *
     * @param string $color
     * @param int $alpha
     * @param string $type
     *
     * @return $this
     */
    public function setLineParameters(string $color, int $alpha = 0, string $type = self::EXCEL_COLOR_TYPE_ARGB): Axis
    {
        $this->lineProperties = $this->setColorProperties($color, $alpha, $type);

        return $this;
    }

    /**
     * Get Fill Property
     *
     * @param string $property
     *
     * @return string|null
     */
    public function getFillProperty(string $property): ?string
    {
        return $this->fillProperties[$property] ?? null;
    }

    /**
     * Get Line Property
     *
     * @param string $property
     *
     * @return string
     */
    public function getLineProperty(string $property): ?string
    {
        return $this->lineProperties[$property] ?? null;
    }

    /**
     * Set Line Style Properties
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
     *
     * @return $this
     */
    public function setLineStyleProperties(?float $lineWidth = null, ?string $compoundType = null, ?string $dashType = null,
                                           ?string $capType = null, ?string $joinType = null, ?string $headArrowType = null,
                                           ?string $headArrowSize = null, ?string $endArrowType = null, ?string $endArrowSize = null): Axis
    {
        $this->lineStyleProperties['width'] = $lineWidth ? Properties::excelPointsWidth((float)$lineWidth) : null;
        $this->lineStyleProperties['compound'] = $compoundType ? (string)$compoundType : null;
        $this->lineStyleProperties['dash'] = $dashType ? (string)$dashType : null;
        $this->lineStyleProperties['cap'] = $capType ? (string)$capType : null;
        $this->lineStyleProperties['join'] = $joinType ? (string)$joinType : null;
        $this->lineStyleProperties['arrow']['head']['type'] = $headArrowType ? (string)$headArrowType : null;
        $this->lineStyleProperties['arrow']['head']['size'] = $headArrowSize ? (string)$headArrowSize : null;
        $this->lineStyleProperties['arrow']['end']['type'] = $endArrowType ? (string)$endArrowType : null;
        $this->lineStyleProperties['arrow']['end']['size'] = $endArrowSize ? (string)$endArrowSize : null;

        return $this;
    }

    /**
     * Get Line Style Property
     *
     * @param array|string $elements
     *
     * @return string|null
     */
    public function getLineStyleProperty($elements): ?string
    {
        return $this->getArrayElementsValue($this->lineStyleProperties, $elements) ?? null;
    }

    /**
     * Get Line Style Arrow Excel Width
     *
     * @param string $arrow
     *
     * @return string|null
     */
    public function getLineStyleArrowWidth(string $arrow): ?string
    {
        return $this->getLineStyleArrowSize($this->lineStyleProperties['arrow'][$arrow]['size'], 'w') ?? null;
    }

    /**
     * Get Line Style Arrow Excel Length
     *
     * @param string $arrow
     *
     * @return string|null
     */
    public function getLineStyleArrowLength(string $arrow): ?string
    {
        return $this->getLineStyleArrowSize($this->lineStyleProperties['arrow'][$arrow]['size'], 'len') ?? null;
    }

    /**
     * Set Shadow Properties
     *
     * @param int $shPresets
     * @param string|null $shColorValue
     * @param string|null $shColorType
     * @param string|null $shColorAlpha
     * @param float|null $shBlur
     * @param int|null $shAngle
     * @param float|null $shDistance
     *
     * @return $this
     */
    public function setShadowProperties(int    $shPresets, ?string $shColorValue = null, ?string $shColorType = null,
                                        ?string $shColorAlpha = null, ?float $shBlur = null, ?int $shAngle = null,
                                        ?float $shDistance = null): Axis
    {
        $this->setShadowPresetsProperties($shPresets);
        $this->setShadowColor(
                ($shColorValue === null) ? $this->shadowProperties['color']['value'] : $shColorValue,
                ($shColorAlpha === null) ? (int) $this->shadowProperties['color']['alpha'] : $shColorAlpha,
                ($shColorType === null) ? $this->shadowProperties['color']['type'] : $shColorType
            );
        $this->setShadowBlur($shBlur);
        $this->setShadowAngle($shAngle);
        $this->setShadowDistance($shDistance);

        return $this;
    }

    /**
     * Set Shadow Color
     *
     * @param int $shadowPresets
     *
     * @return void
     */
    private function setShadowPresetsProperties(int $shadowPresets)
    {
        $this->shadowProperties['presets'] = $shadowPresets;
        $this->setShadowPropertiesMapValues($this->getShadowPresetsMap($shadowPresets));
    }


    /**
     * Set Shadow Color
     *
     * @param string $color
     * @param int $alpha
     * @param string $type
     *
     * @return void
     */
    private function setShadowColor(string $color, int $alpha, string $type)
    {
        $this->shadowProperties['color'] = $this->setColorProperties($color, $alpha, $type);
    }

    /**
     * Set Shadow Blur
     *
     * @param float|null $blur
     *
     * @return void
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
     *
     * @return void
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
     *
     * @return void
     */
    private function setShadowDistance(?float $distance)
    {
        if ($distance !== null) {
            $this->shadowProperties['distance'] = (string)Properties::excelPointsWidth($distance);
        }
    }

    /**
     * Get Shadow Property
     *
     * @param array|string $elements
     *
     * @return mixed|null
     */
    public function getShadowProperty($elements)
    {
        return $this->getArrayElementsValue($this->shadowProperties, $elements) ?? null;
    }

    /**
     * Set Glow Properties
     *
     * @param float $size
     * @param string|null $colorValue
     * @param int|null $colorAlpha
     * @param string|null $colorType
     *
     * @return $this
     */
    public function setGlowProperties(float $size, ?string $colorValue = null, ?int $colorAlpha = null, ?string $colorType = null): Axis
    {
        $this->setGlowSize($size);
        $this->setGlowColor(
                ($colorValue === null) ? $this->glowProperties['color']['value'] : $colorValue,
                ($colorAlpha === null) ? (int) $this->glowProperties['color']['alpha'] : $colorAlpha,
                ($colorType === null) ? $this->glowProperties['color']['type'] : $colorType
            );

        return $this;
    }

    /**
     * Get Glow Property
     *
     * @param array|string $property
     *
     * @return string|null
     */
    public function getGlowProperty($property): ?string
    {
        return $this->getArrayElementsValue($this->glowProperties, $property) ?? null;
    }

    /**
     * Set Glow Size
     *
     * @param float $size
     *
     * @return void
     */
    private function setGlowSize(float $size)
    {
        $this->glowProperties['size'] = Properties::excelPointsWidth($size);
    }

    /**
     * Set Glow Color
     *
     * @param string $color
     * @param int $alpha
     * @param string $type
     *
     * @return void
     */
    private function setGlowColor(string $color, int $alpha, string $type)
    {
        $this->glowProperties['color'] = $this->setColorProperties($color, $alpha, $type);
    }

    /**
     * Set Soft Edges Size
     *
     * @param float $size
     *
     * @return $this
     */
    public function setSoftEdges(float $size): Axis
    {
        $this->softEdges['size'] = (string) Properties::excelPointsWidth((float) $size);

        return $this;
    }

    /**
     * Get Soft Edges Size
     *
     * @return string|null
     */
    public function getSoftEdgesSize(): ?string
    {
        return $this->softEdges['size'];
    }

}