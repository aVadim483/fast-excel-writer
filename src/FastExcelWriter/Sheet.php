<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;
use avadim\FastExcelWriter\Exception\ExceptionAddress;
use avadim\FastExcelWriter\Exception\ExceptionFile;
use avadim\FastExcelWriter\Exception\ExceptionRangeName;

/**
 * Class Sheet
 *
 * @package avadim\FastExcelWriter
 */
class Sheet
{
    // constants for auo width
    protected const WIDTH_LOWER_CASE_LETTER = 1.05;
    protected const WIDTH_UPPER_CASE_LETTER = 1.25;
    protected const WIDTH_WIDE_LETTER = 1.70;
    protected const WIDTH_DOTS_SYMBOLS = 0.50;
    protected const WIDTH_PADDING = 5;

    // constants for notes
    protected const NOTE_LEFT_OFFSET = 1.5;
    protected const NOTE_LEFT_INC = 48.65;
    protected const NOTE_TOP_OFFSET = -4.2;
    protected const NOTE_TOP_INC = 14.4;
    protected const NOTE_DEFAULT_TOP = '1.5pt';
    protected const NOTE_DEFAULT_WIDTH = '96pt';
    protected const NOTE_DEFAULT_HEIGHT = '55.5pt';
    protected const NOTE_DEFAULT_COLOR = '#FFFFE1';

    /** @var Excel */
    public Excel $excel;

    /** @var int Index of the sheet */
    public int $index;

    /** @var string Key of the sheet */
    public string $key;

    /** @var string $relId Id of the relationship */
    public string $relId;

    public bool $active = false;

    /** @var string Temporary file */
    public string $fileTempName = '';

    /** @var string Real sheet name */
    public string $sheetName = '';

    public string $sanitizedSheetName = '';
    public string $xmlName = '';

    public string $fileRels = '';
    public string $xmlRels = '';

    public bool $open = false;
    public bool $close = false;

    // written rows of sheet
    public int $rowCountWritten = 0;

    // written cols of row
    public int $colCountWritten = 0;

    public ?WriterBuffer $fileWriter = null;

    public array $defaultStyle = [];

    protected array $sheetStylesSummary = [];

    // ZERO based
    public int $freezeRows = 0;
    public int $freezeColumns = 0;

    public ?string $autoFilter = null;
    public string $absoluteAutoFilter = '';

    // ZERO based
    public array $colWidths = [];
    public array $colFormulas = [];
    public array $colStyles = [];

    protected array $colStylesSummary = [];

    // special styles by field names
    protected array $fieldStyles = [];

    // ZERO based
    public array $rowHeights = [];
    public array $rowStyles = [];

    // ZERO based
    protected array $cells = [];

    // Current row index
    protected int $currentRow = 0;

    // Current column index
    protected int $currentCol = 0;

    protected int $offsetCol = 0;

    protected array $mergeCells = [];
    protected array $totalArea = [];
    protected array $areas = [];

    protected array $pageOptions = [];


    protected int $relationshipId = 0;

    protected array $relationships = [];

    //protected array $externalLinks = [];

    protected array $lastTouch = [];
    protected int $minRow = 0;
    protected int $minCol = 0;
    protected int $maxRow = 0;
    protected int $maxCol = 0;

    protected array $namedRanges = [];


    protected array $notes = [];

    protected array $media = [];

    /**
     * Sheet constructor
     *
     * @param string $sheetName
     */
    public function __construct(string $sheetName)
    {
        $this->setName($sheetName);
        $this->pageOptions['orientation'] = 'portrait';
        $this->cells = [
            'values' => [],
            'styles' => [],
        ];
        $this->_setCellData('A1', '', null, false);
        $this->lastTouch = [
            'cell' => [
                'row_idx' => 0,
                'col_idx' => 0,
            ],
            'row' => [],
            'area' => [
                'row_idx1' => 0,
                'row_idx2' => 0,
                'col_idx1' => 0,
                'col_idx2' => 0,
            ],
            'ref' => 'cell',
        ];
        $this->_touchEnd(0, 0, 'cell');
    }

    /**
     * Compatibility with previous versions
     *
     * @param $name
     *
     * @return mixed|null
     */
    public function __get($name)
    {
        if ($name === 'book') {
            return $this->excel;
        }
        $trace = debug_backtrace();

        trigger_error(
            'Undefined property: ' . get_class() . '::$' . $name .
            ' (call in file ' . $trace[0]['file'] .
            ' on line ' . $trace[0]['line'] . ') ',
            E_USER_NOTICE);

        return null;
    }

    /**
     * @param int $row
     * @param int $col
     *
     * @return void
     */
    protected function _setDimension(int $row, int $col)
    {
        if (!$this->minRow || $this->minRow > $row) {
            $this->minRow = $row;
        }
        if (!$this->maxRow || $this->maxRow < $row) {
            $this->maxRow = $row;
        }
        if (!$this->minCol || $this->minCol > $col) {
            $this->minCol = $col;
        }
        if (!$this->maxCol || $this->maxCol < $col) {
            $this->maxCol = $col;
        }
    }

    /**
     * @param WriterBuffer $fileWriter
     *
     * @return $this
     */
    public function setFileWriter(WriterBuffer $fileWriter): Sheet
    {
        if (!$this->fileWriter) {
            $this->fileWriter = $fileWriter;
            $this->fileTempName = $fileWriter->getFileName();
            $this->fileRels = $this->fileTempName . '.rels';
        }

        return $this;
    }

    /**
     * @param WriterBuffer $fileWriter
     *
     * @return $this
     */
    public function resetFileWriter(WriterBuffer $fileWriter): Sheet
    {
        $this->fileWriter = $fileWriter;
        $this->fileTempName = $fileWriter->getFileName();
        $this->fileRels = $this->fileTempName . '.rels';

        return $this;
    }

    /**
     * @return array
     */
    public function getExternalLinks(): array
    {
        return $this->relationships['links'] ?? [];
    }

    /**
     * @return string|null
     */
    public function getXmlRels(): ?string
    {
        if ($this->relationships) {
            $result = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            $result .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
            foreach ($this->relationships as $rels) {
                foreach ($rels as $rId => $data) {
                    $result .= '<Relationship Id="rId' . $rId . '" Type="' . $data['type'] . '" Target="' . $data['link'] . '" ' . $data['extra'] . '/>';
                }
            }
            $result .= '</Relationships>';

            return $result;
        }

        return null;
    }

    /**
     * Set sheet name
     *
     * @param string $sheetName
     *
     * @return $this
     */
    public function setName(string $sheetName): Sheet
    {
        $this->sheetName = $sheetName;
        $this->sanitizedSheetName = Writer::sanitizeSheetName($sheetName);

        return $this;
    }

    /**
     * Get sheet name
     *
     * @return string
     */
    public function getName(): string
    {
        return $this->sheetName;
    }

    /**
     * @param string $option
     * @param mixed $value
     *
     * @return $this
     */
    public function setPageOptions(string $option, $value): Sheet
    {
        if ($this->rowCountWritten) {
            throw new Exception('Cannot set page settings after rows writing');
        }
        $this->pageOptions[$option] = $value;

        return $this;
    }

    /**
     * @param $option
     *
     * @return mixed|null
     */
    public function getPageOptions($option)
    {
        return $this->pageOptions[$option] ?? null;
    }

    /**
     * @return $this
     */
    public function pageOrientationPortrait(): Sheet
    {
        return $this->setPageOptions('orientation', 'portrait');
    }

    /**
     * Set page orientation as Portrait
     *
     * @return $this
     */
    public function pagePortrait(): Sheet
    {
        return $this->setPageOptions('orientation', 'portrait');
    }

    /**
     * @return $this
     */
    public function pageOrientationLandscape(): Sheet
    {
        return $this->setPageOptions('orientation', 'landscape');
    }

    /**
     * Set page orientation as Landscape
     *
     * @return $this
     */
    public function pageLandscape(): Sheet
    {
        return $this->setPageOptions('orientation', 'landscape');
    }

    /**
     * @param int|string|null $numPage
     *
     * @return $this
     */
    public function pageFitToWidth($numPage = 1): Sheet
    {
        if (strtolower($numPage) === 'auto') {
            $numPage = 0;
        }
        if ($numPage >=0) {
            $this->setPageOptions('fit_width', (int)$numPage);
        }
        return $this;
    }

    /**
     * @param int|string|null $numPage
     *
     * @return $this
     */
    public function pageFitToHeight($numPage = 1): Sheet
    {
        if (strtolower($numPage) === 'auto') {
            $numPage = 0;
        }
        if ($numPage >=0) {
            $this->setPageOptions('fit_height', (int)$numPage);
        }
        return $this;
    }

    /**
     * @return string
     */
    public function getPageOrientation(): string
    {
        return $this->pageOptions['orientation'] ?? 'portrait';
    }

    /**
     * @return int
     */
    public function getPageFitToWidth(): int
    {
        return (int)($this->pageOptions['fit_width'] ?? 0);
    }

    /**
     * @return int
     */
    public function getPageFitToHeight(): int
    {
        return (int)($this->pageOptions['fit_height'] ?? 0);
    }

    /**
     * @return bool
     */
    public function getPageFit(): bool
    {
        return $this->getPageFitToWidth() || $this->getPageFitToHeight();
    }

    /**
     * setFreeze(3, 3) - number rows and columns to freeze
     * setFreeze('C3') - left top cell of free area
     *
     * @param mixed $freezeRows
     * @param mixed $freezeColumns
     *
     * @return $this
     */
    public function setFreeze($freezeRows, $freezeColumns = null): Sheet
    {
        if (!is_numeric($freezeRows) && null === $freezeColumns) {
            $dimension = Excel::rangeDimension($freezeRows);
            if ($dimension) {
                $this->setFreezeRows($dimension['row'] - 1);
                $this->setFreezeColumns($dimension['col'] - 1);
            }
        }
        else {
            $this->setFreezeRows((int)$freezeRows);
            $this->setFreezeColumns((int)$freezeColumns);
        }

        return $this;
    }

    /**
     * @param int $freezeRows Number rows to freeze
     *
     * @return $this
     */
    public function setFreezeRows(int $freezeRows): Sheet
    {
        $this->freezeRows = ($freezeRows > 0) ? $freezeRows : 0;

        return $this;
    }

    /**
     * @param int $freezeColumns Number columns to freeze
     *
     * @return $this
     */
    public function setFreezeColumns(int $freezeColumns): Sheet
    {
        $this->freezeColumns = ($freezeColumns > 0) ? $freezeColumns : 0;

        return $this;
    }

    /**
     * @param int|null $row
     * @param int|null $col
     *
     * @return $this
     */
    public function setAutofilter(?int $row = 1, ?int $col = 1): Sheet
    {
        if ($row >= 0) {
            if (empty($row)) {
                $this->autoFilter = null;
            }
            else {
                $this->autoFilter = Excel::cellAddress($row, $col);
            }
        }
        $this->absoluteAutoFilter = Excel::cellAddress($row, $col, true);

        return $this;
    }

    /**
     * @param string|array $cellAddress
     *
     * @return $this
     */
    public function setTopLeftCell($cellAddress): Sheet
    {
        $address = $this->_moveTo($cellAddress);
        $this->_touchStart($address['rowIndex'], $address['colIndex'], 'cell');
        $this->_touchEnd($address['rowIndex'], $address['colIndex'], 'cell');

        $this->currentRow = $address['rowIndex'];
        $this->currentCol = $this->offsetCol = $address['colIndex'];

        return $this;
    }

    /**
     * @param array $columns
     *
     * @return array
     */
    protected function normalizeColKeys(array $columns): array
    {
        $keys = array_keys($columns);
        if (reset($keys) === 0) {
            foreach ($keys as $n => $key) {
                if (is_int($key)) {
                    $keys[$n] = $key + 1;
                }
            }
            $result = array_combine($keys, array_values($columns));
        }
        else {
            $result = $columns;
        }

        return $result;
    }

    /**
     * Set options of columns (widths, styles, formats, etc)
     *
     * Call examples:
     *  setColOptions('B', ['width' = 20]) - options for column 'B'
     *  setColOptions('B:D', ['width' = 'auto']) - options for range of columns
     *  setColOptions(['B' => ['width' = 20], 'C' => ['font-color' = '#f00']]) - options for several columns 'B' and 'C'
     *
     * @param mixed $arg1
     * @param array|null $arg2
     *
     * @return $this
     */
    public function setColOptions($arg1, array $arg2 = null): Sheet
    {
        if ($this->currentCol) {
            $this->_writeCurrentRow();
        }

        if ($arg2 === null) {
            $options = array_combine(Excel::colLetterRange(array_keys($arg1)), array_values($arg1));
            foreach ($options as $col => $colOptions) {
                if ($colOptions) {
                    $options[$col] = Style::normalize($colOptions);
                }
                else {
                    $options[$col] = null;
                }
            }
        }
        else {
            $options = [];
            $colNumbers = Excel::colNumberRange($arg1);
            if ($colNumbers) {
                $colOptions = Style::normalize($arg2);
                foreach ($colNumbers as $col) {
                    $options[$col] = $colOptions;
                }
            }
        }
        $options = array_filter($options);
        if ($options) {
            $options = $this->normalizeColKeys($options);
            foreach($options as $col => $colOptions) {
                $style = [];
                if (isset($colOptions['options']['width'])) {
                    $this->setColWidth($col, $colOptions['options']['width']);
                }
                foreach($colOptions as $optionName => $optionValues) {
                    if (is_array($optionValues)) {
                        foreach ($optionValues as $key => $val) {
                            if ($key === 'width') {
                                $this->setColWidth($col, $val);
                            }
                            elseif ($key === 'formula') {
                                $this->setColFormula($col, $val);
                            }
                            else {
                                $style[$optionName][$key] = $val;
                            }
                        }
                    }
                    else {
                        $style[$optionName] = $optionValues;
                    }
                }
                if ($style) {
                    $this->setColStyle($col, $style);
                }
            }
        }
        $this->clearSummary();

        return $this;
    }

    /**
     * Alias for setColOptions()
     *
     * @param $arg1
     * @param array|null $arg2
     *
     * @return $this
     */
    public function setColStyles($arg1, array $arg2 = null): Sheet
    {
        return $this->setColOptions($arg1, $arg2);
    }

    /**
     * Set width of single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param int|float|string $width
     *
     * @return $this
     */
    public function setColWidth($col, $width): Sheet
    {
        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                if (strtolower($width) === 'auto') {
                    $this->colStyles[$colIdx]['options']['width-auto'] = true;
                }
                else {
                    $width = Style::numFloat($width);
                    if (is_numeric($width)) {
                        $this->colWidths[$colIdx] = $width;
                    }
                }
            }
        }
        $this->clearSummary();

        return $this;
    }

    /**
     * @param array $widths
     *
     * @return $this
     */
    public function setColWidths(array $widths): Sheet
    {
        if ($widths) {
            $widths = Excel::colKeysToLetters($widths);
            foreach ($widths as $col => $width) {
                $this->setColWidth($col, $width);
            }
        }

        return $this;
    }

    /**
     * Set width of single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     *
     * @return $this
     */
    public function setColWidthAuto($col): Sheet
    {
        return $this->setColWidth($col, 'auto');
    }

    /**
     * Set style of single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param mixed $style
     *
     * @return $this
     */
    public function setColStyle($col, $style): Sheet
    {
        if ($this->currentCol) {
            $this->_writeCurrentRow();
        }

        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                $style = Style::normalize($style);
                if (!empty($this->colStyles[$colIdx])) {
                    $this->colStyles[$colIdx] = array_replace_recursive($this->colStyles[$colIdx], $style);
                }
                else {
                    $this->colStyles[$colIdx] = $style;
                }
            }
        }
        $this->clearSummary();

        return $this;
    }

    /**
     * Set formula for single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param string $formula
     *
     * @return $this
     */
    public function setColFormula($col, string $formula): Sheet
    {
        if ($this->currentCol) {
            $this->_writeCurrentRow();
        }

        $colIndexes = Excel::colIndexRange($col);
        if ($formula) {
            if ($formula[0] !== '=') {
                $formula = '=' . $formula;
            }
        }
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                $this->colFormulas[$colIdx] = $formula;
            }
        }

        return $this;
    }

    /**
     * Set format of single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param mixed $format
     *
     * @return $this
     */
    public function setColFormat($col, $format): Sheet
    {
        $this->setColStyle($col, ['format' => $format]);
        $this->clearSummary();

        return $this;
    }

    /**
     * Set formats of columns
     *
     * @param array $formats
     *
     * @return $this
     */
    public function setColFormats(array $formats): Sheet
    {
        if ($formats) {
            $formats = array_combine(Excel::colLetterRange(array_keys($formats)), array_values($formats));
            foreach ($formats as $col => $format) {
                $this->setColFormat($col, $format);
            }
        }

        return $this;
    }

    /**
     * @param array $styles
     *
     * @return $this
     */
    public function setFieldStyles(array $styles): Sheet
    {
        if ($styles) {
            foreach ($styles as $field => $style) {
                $style = Style::normalize($style);
                if (empty($this->fieldStyles[$field])) {
                    $this->fieldStyles[$field] = $style;
                }
                else {
                    $this->fieldStyles[$field] = array_merge_recursive($this->fieldStyles[$field], $style);
                }
            }
        }

        return $this;
    }

    /**
     * @param array $formats
     *
     * @return $this
     */
    public function setFieldFormats(array $formats): Sheet
    {
        if ($formats) {
            $styles = [];
            foreach ($formats as $field => $format) {
                $styles[$field] = Style::normalize(['format' => $format]);
            }
            $this->setFieldStyles($styles);
        }

        return $this;
    }

    /**
     * Height of a specific row
     *
     * @param $rowNum
     * @param $height
     *
     * @return $this
     */
    public function setRowHeight($rowNum, $height): Sheet
    {
        if (is_array($rowNum)) {
            foreach ($rowNum as $row) {
                $this->setRowHeight($row, $height);
            }
        }
        elseif (is_numeric($rowNum)) {
            $this->rowHeights[(int)$rowNum - 1] = str_replace(',', '.', (float)$height);
        }
        return $this;
    }

    /**
     * Multiple rows height
     *
     * @param array $heights
     *
     * @return $this
     */
    public function setRowHeights(array $heights): Sheet
    {
        foreach ($heights as $rowNum => $rowHeight) {
            $this->setRowHeight($rowNum, $rowHeight);
        }
        return $this;
    }

    /**
     * setRowOptions(3, ['height' = 20]) - options for row number 3
     * setRowOptions([3 => ['height' = 20], 4 => ['font-color' = '#f00']]) - options for several rows
     * setRowOptions('2:5', ['font-color' = '#f00']) - options for range of rows
     *
     * @param mixed $arg1
     * @param array|null $arg2
     *
     * @return $this
     */
    public function setRowOptions($arg1, array $arg2 = null): Sheet
    {
        if ($arg2 === null) {
            $options = $arg1;
        }
        else {
            if (is_string($arg1) && preg_match('/^(\d+):(\d+)$/', $arg1, $m)) {
                $options = [];
                for ($row = $m[1]; $row <= $m[2]; $row++) {
                    $options[$row] = $arg2;
                }
            }
            elseif (is_numeric($arg1)) {
                $options[(int)$arg1] = $arg2;
            }
            else {
                $options = [];
            }
        }
        foreach ($options as $rowNum => $rowOptions) {
            $rowIdx = (int)$rowNum - 1;
            if ($rowOptions) {
                $rowOptions = Style::normalize($rowOptions);
                if (isset($rowOptions['height'])) {
                    $this->setRowHeight($rowNum, $rowOptions['height']);
                    unset($rowOptions['height']);
                }
                if (isset($this->rowStyles[$rowIdx])) {
                    $this->rowStyles[$rowIdx] = array_replace_recursive($this->rowStyles[$rowIdx], $rowOptions);
                }
                else {
                    $this->rowStyles[$rowIdx] = $rowOptions;
                }
            }
        }
        return $this;
    }

    public function setRowStyles($arg1, array $arg2 = null): Sheet
    {
        return $this->setRowOptions($arg1, $arg2);
    }

    /**
     * @param string $address
     * @param string $link
     */
    protected function _addExternalLink(string $address, string $link)
    {
        $this->relationships['links'][++$this->relationshipId] = [
            'cell' => $address,
            'link' => $link,
            'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            'extra' => 'TargetMode="External"',
        ];
    }

    /**
     * @param Writer|null $writer
     * @param array|null $row Values of all cells of row (incl. empty)
     * @param array|null $rowOptions Specified style for the row
     * @param array|null $cellsOptions Styles of all cells of row (incl. empty)
     */
    protected function _writeRow(?Writer $writer, array $row = [], array $rowOptions = [], array $cellsOptions = [])
    {
        static $_styleCache = [];

        if ($this->rowCountWritten === 0) {
            $_styleCache = [];
        }

        $rowAttr = '';
        if (!empty($rowOptions['height'])) {
            $height = $rowOptions['height'];
        }
        elseif (isset($this->rowHeights[$this->rowCountWritten])) {
            $height = $this->rowHeights[$this->rowCountWritten];
        }
        else {
            $height = null;
        }
        if ($height !== null) {
            $rowAttr .= ' customHeight="1" ht="' . (float)$height . '" ';
        }
        if (!empty($rowOptions['hidden'])) {
            $rowAttr .= ' hidden="1" ';
        }
        if (!empty($rowOptions['collapsed'])) {
            $rowAttr .= ' collapsed="1" ';
        }

        // add auto formulas of columns
        if ($this->colFormulas && $row) {
            foreach($this->colFormulas as $colIdx => $formula) {
                if (!isset($row[$colIdx])) {
                    $row[$colIdx] = $formula;
                }
            }
            ksort($row);
        }

        if ($row) {
            if (empty($this->sheetStylesSummary)) {
                if ($this->defaultStyle) {
                    $this->sheetStylesSummary = [
                        'general_style' => Style::mergeStyles([$this->excel->style->defaultStyle, $this->defaultStyle]),
                        'hyperlink_style' => Style::mergeStyles([$this->excel->style->hyperlinkStyle, $this->defaultStyle]),
                    ];
                }
                else {
                    $this->sheetStylesSummary = [
                        'general_style' => $this->excel->style->defaultStyle,
                        'hyperlink_style' => $this->excel->style->hyperlinkStyle,
                    ];
                }
            }
            $this->fileWriter->write('<row r="' . ($this->rowCountWritten + 1) . '" ' . $rowAttr . '>');
            $rowIdx = $this->rowCountWritten;
            foreach ($row as $colIdx => $cellValue) {
                if (!isset($this->colStylesSummary[$colIdx])) {
                    if (!isset($this->colStyles[$colIdx])) {
                        $this->colStylesSummary[$colIdx] = $this->sheetStylesSummary;
                    }
                    else {
                        $this->colStylesSummary[$colIdx] = [
                            'general_style' => Style::mergeStyles([
                                $this->sheetStylesSummary['general_style'],
                                $this->colStyles[$colIdx],
                            ]),
                            'hyperlink_style' => Style::mergeStyles([
                                $this->sheetStylesSummary['hyperlink_style'],
                                $this->colStyles[$colIdx],
                            ]),
                        ];
                    }
                }

                $styleStack = [];
                $baseStyle = !empty($cellsOptions[$colIdx]['hyperlink']) ? $this->colStylesSummary[$colIdx]['hyperlink_style'] : $this->colStylesSummary[$colIdx]['general_style'];
                if ($baseStyle) {
                    $styleStack = [$baseStyle];
                }

                if (!empty($this->rowStyles[$rowIdx])) {
                    $styleStack[] = $this->rowStyles[$rowIdx];
                }
                if (!empty($this->cells['styles'][$rowIdx][$colIdx])) {
                    $styleStack[] = $this->cells['styles'][$rowIdx][$colIdx];
                }
                if ($rowOptions) {
                    $styleStack[] = $rowOptions;
                }
                if (!empty($cellsOptions[$colIdx])) {
                    $styleStack[] = $cellsOptions[$colIdx];
                }
                if (count($styleStack) > 1) {
                    $cellStyle = Style::mergeStyles($styleStack);
                }
                else {
                    $cellStyle = $styleStack ? $styleStack[0] : [];
                }
                if (!empty($cellStyle['format']['format-pattern']) && !empty($this->excel->style->defaultFormatStyles[$cellStyle['format']['format-pattern']])) {
                    $cellStyle = Style::mergeStyles([$this->excel->style->defaultFormatStyles[$cellStyle['format']['format-pattern']], $cellStyle]);
                }

                if (isset($cellStyle['hyperlink'])) {
                    if (!empty($cellStyle['hyperlink'])) {
                        if (is_string($cellStyle['hyperlink'])) {
                            $link = $cellStyle['hyperlink'];
                        }
                        else {
                            $link = $cellValue;
                        }
                        $cellValue = [
                            'shared_value' => $cellValue,
                            'shared_index' => $this->excel->addSharedString($cellValue),
                        ];
                        $this->_addExternalLink(Excel::cellAddress($rowIdx + 1, $colIdx + 1), $link);
                        if (!empty($this->excel->style->hyperlinkStyle)) {
                            $cellStyle = Style::mergeStyles([$this->excel->style->hyperlinkStyle, $cellStyle]);
                        }
                    }
                    unset($cellStyle['hyperlink']);
                }

                $styleHash = json_encode($cellStyle);
                if (!isset($_styleCache[$styleHash])) {
                    $cellStyleIdx = $this->excel->style->addStyle($cellStyle, $resultStyle);
                    $_styleCache[$styleHash] = ['cell_style' => $cellStyle, 'result_style' => $resultStyle, 'style_idx' => $cellStyleIdx];
                }
                else {
                    $resultStyle = $_styleCache[$styleHash]['result_style'];
                    $cellStyleIdx = $_styleCache[$styleHash]['style_idx'];
                }

                $numberFormat = $resultStyle['number_format'];
                $numberFormatType = $resultStyle['number_format_type'];

                if (!empty($cellStyle['options']['width-auto'])) {
                    $this->_columnWidth($colIdx, $cellValue, $numberFormat, $resultStyle ?? []);
                }

                if (!$writer) {
                    $writer = $this->excel->getWriter();
                }
                $writer->_writeCell($this->fileWriter, $rowIdx + 1, $colIdx + 1, $cellValue, $numberFormatType, $cellStyleIdx);
                $this->_setDimension($rowIdx + 1, $colIdx + 1);
                $colIdx++;
                if ($colIdx > $this->colCountWritten) {
                    $this->colCountWritten = $colIdx;
                }
            }
            $this->fileWriter->write('</row>');
        }
        else {
            $this->fileWriter->write('<row r="' . ($this->rowCountWritten + 1) . '" ' . $rowAttr . '/>');
        }
        $this->rowCountWritten++;
    }

    /**
     * @param string $str
     * @param int|float $fontSize
     * @param bool|null $numFormat
     *
     * @return float
     */
    protected function _calcWidth(string $str, $fontSize, ?bool $numFormat = false): float
    {
        if ($numFormat && strpos($str, ';')) {
            $lenArray = [];
            foreach (explode(';', $str) as $part) {
                $lenArray[] = $this->_calcWidth($part, $fontSize);
            }

            return max(...$lenArray);
        }

        $len = mb_strlen($str);
        $wideCount = 0;
        $upperCount = 0;
        $dotsCount = 0;
        if (preg_match_all("/[@%&WM]/u", $str, $matches)) {
            $wideCount = count($matches[0]);
            $str = preg_replace("/[@%&WM]/u", '', $str);
        }
        if (preg_match_all("/[,\.\-\:\';`Iil\[\]]/u", $str, $matches)) {
            $dotsCount = count($matches[0]);
            $str = preg_replace("/[,\.\-\:\';`Iil\[\]]/u", '', $str);
        }
        if (preg_match_all("/[[:upper:]#@w]/u", $str, $matches)) {
            $upperCount = count($matches[0]);
        }

        // width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
        $k = $fontSize * 0.66;

        $n = ($len - $wideCount - $upperCount - $dotsCount) * self::WIDTH_LOWER_CASE_LETTER * $k +
            $wideCount * self::WIDTH_WIDE_LETTER * $k +
            $upperCount * self::WIDTH_UPPER_CASE_LETTER * $k +
            $dotsCount * self::WIDTH_DOTS_SYMBOLS * $k + self::WIDTH_PADDING;

        return round($n / $k, 8);
    }

    /**
     * @param mixed $value
     * @param string $format
     *
     * @return string
     */
    protected function _formatValue($value, string $format): string
    {
        if (is_numeric($value)) {
            if (strpos($format, ';')) {
                $formats = explode(';', $format);
                if ($value > 0 && !empty($formats[0])) {
                    return $this->_formatValue($value, $formats[0]);
                }
                if ($value < 0 && !empty($formats[1])) {
                    return $this->_formatValue($value, $formats[1]);
                }
                if ((int)$value === 0 && !empty($formats[2])) {
                    return $this->_formatValue($value, $formats[2]);
                }
                return $this->_formatValue($value, '0');
            }
            else {
                if (preg_match('/[#0](\.0+)/', $format, $m)) {
                    $value = number_format($value, strlen($m[1]) - 1);
                }
                else {
                    $value = number_format($value, 0);
                }
                $cnt = substr_count($format, '\\');
                if ($cnt) {
                    $value .= str_repeat('-', $cnt);
                }
                if (preg_match('/\[\$.+\]/U', $format, $m)) {
                    $value .= str_replace(['[$', ']'], '', $m[0]);
                }

                return $value;
            }
        }
        elseif (strpos($format, ';')) {
            // value is not numeric but format for number
            return $value;
        }

        return $format;
    }

    /**
     * @param $colNum
     * @param $cellValue
     * @param $numberFormat
     * @param $style
     */
    protected function _columnWidth($colNum, $cellValue, $numberFormat, $style)
    {
        static $cache = [];

        if ($cellValue) {
            $fontSize = $style['font']['val']['size'] ?? 10;
            $value = (isset($cellValue['shared_value'])) ? $cellValue['shared_value'] : $cellValue;
            $key = '[[[' . $fontSize . ']]][[[' . $numberFormat . ']]][[[' . $value . ']]]';
            if (isset($cache[$key])) {
                $len = $cache[$key];
            }
            else {
                $len = $this->_calcWidth($value, $fontSize);
                if ($numberFormat !== 'GENERAL' && $numberFormat !== '0' && $numberFormat != '@') {
                    $numberFormat = $this->_formatValue($value, $numberFormat);
                    $len = max($len, $this->_calcWidth(str_replace('\\', '', $numberFormat), $fontSize, true));
                }
                $cache[$key] = $len;
            }
            if (empty($this->colWidths[$colNum]) || $this->colWidths[$colNum] < $len) {
                $this->colWidths[$colNum] = $len;
            }
        }
    }

    protected function clearSummary()
    {
        if ($this->sheetStylesSummary) {
            $this->sheetStylesSummary = [];
        }
        if ($this->colStylesSummary) {
            $this->colStylesSummary = [];
        }
    }

    /**
     * Sets default style
     *
     * @param array $style
     *
     * @return $this
     */
    public function setDefaultStyle(array $style): Sheet
    {
        $this->clearSummary();
        $this->defaultStyle = Style::normalize($style);

        return $this;
    }

    /**
     * Returns default style
     *
     * @return array
     */
    public function getDefaultStyle(): array
    {
        return $this->defaultStyle;
    }

    // ++++++++++++++++++++ //
    // +++ DEFAULT FONT +++ //

    /**
     * @param string|array $font
     *
     * @return $this
     */
    public function setDefaultFont($font): Sheet
    {
        $normStyle = Style::normalizeFont($font);
        if (isset($normStyle['font'])) {
            if (isset($this->defaultStyle['font'])) {
                $this->defaultStyle['font'] = array_replace($this->defaultStyle['font'], $normStyle['font']);
            }
            else {
                $this->defaultStyle['font'] = $normStyle['font'];
            }
        }

        return $this;
    }

    /**
     * @param string $fontName
     *
     * @return $this
     */
    public function setDefaultFontName(string $fontName): Sheet
    {
        $this->defaultStyle['font']['font-name'] = $fontName;

        return $this;
    }

    /**
     * @param int $fontSize
     *
     * @return $this
     */
    public function setDefaultFontSize(int $fontSize): Sheet
    {
        $this->defaultStyle['font']['font-size'] = $fontSize;

        return $this;
    }

    /**
     * @param string $fontStyle
     *
     * @return $this
     */
    public function setDefaultFontStyle(string $fontStyle): Sheet
    {
        $key = 'font-style-' . strtolower($fontStyle);
        $this->defaultStyle['font'][$key] = 1;

        return $this;
    }

    public function setDefaultFontStyleBold(): Sheet
    {
        return $this->applyFontStyle('bold');
    }

    /**
     * @return $this
     */
    public function setDefaultFontStyleItalic(): Sheet
    {
        return $this->applyFontStyle('italic');
    }

    /**
     * @param bool|null $double
     *
     * @return $this
     */
    public function setDefaultFontStyleUnderline(?bool $double = false): Sheet
    {
        $this->defaultStyle['font']['font-style-underline'] = ['font-style-underline' => $double ? 2 : 1];

        return $this;
    }

    /**
     * @return $this
     */
    public function setDefaultFontStyleStrikethrough(): Sheet
    {
        return $this->applyFontStyle('strikethrough');
    }

    /**
     * @param string $fontColor
     *
     * @return $this
     */
    public function setDefaultFontColor(string $fontColor): Sheet
    {
        $this->defaultStyle['font']['font-style-underline'] = ['font-color' => $fontColor];

        return $this;
    }

    // --- DEFAULT FONT --- //
    // ---------------------//


    /**
     * Write value to the current cell and move pointer to the next cell in the row
     *
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     *
     * @throws \Exception
     */
    public function writeCell($value, array $styles = null): Sheet
    {
        if ($this->lastTouch['ref'] === 'row') {
            $this->_writeCurrentRow();
        }
        ///-- $styles = $styles ? Style::normalize($styles) : [];
        if ($this->currentRow < $this->rowCountWritten) {
            $this->currentRow = $this->rowCountWritten;
        }
        //$this->withLastCell();
        $cellAddress = [
            'row' => 1 + $this->currentRow,
            'col' => 1 + $this->currentCol,
        ];
        $this->_setCellData($cellAddress, $value, $styles, false);
        $this->_touchEnd($this->currentRow, $this->currentCol, 'cell');
        ++$this->currentCol;

        return $this;
    }

    /**
     * @return $this
     *
     * @throws \Exception
     */
    public function nextCell(): Sheet
    {
        $this->writeCell(null);

        return $this;
    }

    /**
     * writeHeader(['title1', 'title2', 'title3']) - texts for cells of header
     * writeHeader(['title1' => '@text', 'title2' => 'YYYY-MM-DD', 'title3' => ['format' => ..., 'font' => ...]]) - texts and formats of columns
     * writeHeader([<cell values>], [<row styles>], [<col styles>]) - texts and formats of columns and options of row
     *
     * @param array $header
     * @param array|null $rowStyle
     * @param array|null $colStyles
     *
     * @return $this
     */
    public function writeHeader(array $header, array $rowStyle = null, ?array $colStyles = []): Sheet
    {
        $rowValues = [];
        $colNum = 0;
        foreach($header as $key => $val) {
            if (!is_int($key)) {
                $rowValues[$colNum] = $key;
                if (is_scalar($val)) {
                    $colStyles[$colNum + $this->offsetCol]['format'] = $val;
                }
                else {
                    $colStyles[$colNum + $this->offsetCol] = isset($colStyles[$colNum + $this->currentCol]) ? array_replace_recursive($colStyles[$colNum + $this->currentCol], $val) : $val;
                }
            }
            else {
                $rowValues[$colNum] = $val;
                $colStyles[$colNum + $this->offsetCol] = null;
            }
            $colNum++;
        }
        $this->writeRow($rowValues, $rowStyle);
        if ($colStyles) {
            // column styles for next rows
            $this->colStyles[-1] = $colStyles;
        }

        return $this;
    }

    /**
     * @param string|array $cellAddress
     *
     * @return array
     */
    protected function _moveTo(&$cellAddress): ?array
    {
        $address = $this->_parseAddress($cellAddress);
        if (!isset($address['row'], $address['col'])) {
            ExceptionAddress::throwNew('Wrong cell address %s', print_r($address, 1));
        }
        elseif ($address['row'] <= $this->rowCountWritten) {
            ExceptionAddress::throwNew('Row number must be greater then written rows');
        }
        else {
            $cellAddress = $address;
        }

        if (isset($address['colIndex'], $address['rowIndex'])) {
            $this->currentCol = $address['colIndex'];
            $this->currentRow = $address['rowIndex'];
        }
        else {
            while ($this->currentRow < $cellAddress['row'] - 1) {
                $this->nextRow();
            }
        }

        return $address;
    }

    /**
     * Write value to the specified cell and move pointer to the next cell in the row
     *
     * $cellAddress formats:
     *  'B5'
     *  'B5:C7'
     *  ['row' => 6, 'col' => 7]
     *  [6, 7]
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     */
    public function writeTo($cellAddress, $value, ?array $styles = []): Sheet
    {
        $address = $this->_moveTo($cellAddress);
        $this->_touchStart($address['rowIndex'], $address['colIndex'], 'cell');
        //$this->withLastCell();

        ///-- $styles = $styles ? Style::normalize($styles) : null;
        $this->_setCellData($cellAddress, $value, $styles, true, true);
        if (isset($address['width'], $address['range']) && $address['width'] > 1) {
            $this->mergeCells($address['range']);
            $this->currentCol += $address['width'];
        }
        else {
            $this->currentCol++;
        }
        $this->_touchEnd($address['rowIndex'], $address['colIndex'], 'cell');

        return $this;
    }

    /**
     * Merge cells
     *
     * mergeCells('A1:C3')
     * mergeCells(['A1:B2', 'C1:D2'])
     *
     * @param array|string|int $rangeSet
     *
     * @return $this
     */
    public function mergeCells($rangeSet): Sheet
    {
        foreach((array)$rangeSet as $range) {
            if (isset($this->mergeCells[$range]) || empty($range)) {
                // cells are already merged
                continue;
            }
            $dimension = Excel::rangeDimension($range, true);
            // check intersection with saved merged cells
            foreach ($this->mergeCells as $savedRange => $savedDimension) {
                if (
                    ((($dimension['rowNum1'] >= $savedDimension['rowNum1']) && ($dimension['rowNum1'] <= $savedDimension['rowNum2']))
                        || (($dimension['rowNum2'] >= $savedDimension['rowNum1']) && ($dimension['rowNum2'] <= $savedDimension['rowNum2'])))
                    && ((($dimension['colNum1'] >= $savedDimension['colNum1']) && ($dimension['colNum1'] <= $savedDimension['colNum2']))
                        || (($dimension['colNum2'] >= $savedDimension['colNum1']) && ($dimension['colNum2'] <= $savedDimension['colNum2'])))
                ) {
                    throw new Exception("Cannot merge cells $range because they are intersecting with $savedRange");
                }
            }
            $this->mergeCells[$dimension['range']] = $dimension;
        }

        return $this;
    }

    /**
     * Merge relative cells
     *
     * mergeCells(3) -> 3 columns of current row -> mergeCells('A5:C5') // if current row is 5
     * mergeCells(['RC3:RC5', 'RC6:RC7']) -> mergeCells(['C7:E7', 'F7:G7']) // if current row is 7
     *
     * @param array|string|int $rangeSet
     *
     * @return $this
     */
    public function mergeRelCells($rangeSet): Sheet
    {
        if (is_int($rangeSet)) {
            $rangeSet = 'A' . $this->rowCountWritten . ':' . Excel::colLetter($rangeSet)  . $this->rowCountWritten;
        }
        foreach((array)$rangeSet as $range) {
            if (isset($this->mergeCells[$range]) || empty($range)) {
                // cells are already merged
                continue;
            }
            $dimension = $this->_rangeDimension($range, 1, 0);
            // check intersection with saved merged cells
            foreach ($this->mergeCells as $savedRange => $savedDimension) {
                if (
                    ((($dimension['rowNum1'] >= $savedDimension['rowNum1']) && ($dimension['rowNum1'] <= $savedDimension['rowNum2']))
                        || (($dimension['rowNum2'] >= $savedDimension['rowNum1']) && ($dimension['rowNum2'] <= $savedDimension['rowNum2'])))
                    && ((($dimension['colNum1'] >= $savedDimension['colNum1']) && ($dimension['colNum1'] <= $savedDimension['colNum2']))
                        || (($dimension['colNum2'] >= $savedDimension['colNum1']) && ($dimension['colNum2'] <= $savedDimension['colNum2'])))
                ) {
                    if ($range !== $dimension['range']) {
                        $range .= ' (' . $dimension['range'] . ')';
                    }
                    throw new Exception("Cannot merge cells $range because they are intersecting with $savedRange");
                }
            }
            $this->mergeCells[$dimension['range']] = $dimension;
        }

        return $this;
    }

    /**
     * Returns merged cells
     *
     * @return array
     */
    public function getMergedCells(): array
    {
        return array_keys($this->mergeCells);
    }

    /**
     * @return void
     */
    protected function _writeCurrentRow()
    {
        if (!empty($this->cells['values']) || !empty($this->cells['styles'])) {
            $writer = $this->excel->getWriter();
            if (!$this->open) {
                $writer->writeSheetDataBegin($this);
            }
            $maxRowIdx = max($this->cells['values'] ? max(array_keys($this->cells['values'])) : -1,
                $this->cells['styles'] ? max(array_keys($this->cells['styles'])) : -1);
            if ($maxRowIdx >= 0) {
                // has values or styles
                if ($maxRowIdx < $this->currentRow) {
                    $maxRowIdx = $this->currentRow;
                }

                for ($rowIdx = $this->rowCountWritten; $rowIdx <= $maxRowIdx; $rowIdx++) {
                    if (isset($this->cells['values'][$rowIdx])) {
                        $values = $this->cells['values'][$rowIdx];
                        unset($this->cells['values'][$rowIdx]);
                    }
                    else {
                        $values = [];
                    }
                    if (isset($this->cells['styles'][$rowIdx])) {
                        $styles = $this->cells['styles'][$rowIdx];
                        unset($this->cells['styles'][$rowIdx]);
                    }
                    else {
                        $styles = [];
                    }
                    if ($values || $styles) {
                        $this->_writeRow($writer, $values, [], $styles);
                    }
                    else {
                        //$this->rowCount++;
                        $this->_writeRow($writer, [], [], []);
                    }
                    if (isset($this->rowStyles[$rowIdx])) {
                        unset($this->rowStyles[$rowIdx]);
                    }
                }
                $this->currentRow++;
            }

            $this->currentCol = $this->offsetCol;

            if (isset($this->colStyles[-1])) {
                $this->setColOptions($this->colStyles[-1]);
                unset($this->colStyles[-1]);
            }
            $this->withLastCell();
        }
    }

    /**
     * Write values to the current row
     *
     * @param array $rowValues Values of cells
     * @param array|null $rowStyle Style applied to the entire row
     * @param array|null $cellStyles Styles of specified cells in the row
     *
     * @return $this
     */
    public function writeRow(array $rowValues = [], array $rowStyle = null, array $cellStyles = null): Sheet
    {
        if (($this->currentCol > $this->offsetCol) || $this->areas) {
            $this->_writeCurrentRow();
        }

        if (!is_array($rowValues)) {
            $rowFieldNames = [0];
            $rowValues = [$rowValues];
        }
        else {
            $rowFieldNames = array_keys($rowValues);
            $rowValues = array_values($rowValues);
        }
        if (is_array($cellStyles)) {
            $key = array_key_first($cellStyles);
            if (!is_int($key)) {
                $cellStyles = Excel::colKeysToIndexes($cellStyles);
            }
        }
        else {
            $cellStyles = [];
        }

        if ($rowStyle) {
            $rowStyle = Style::normalize($rowStyle);
            $this->rowStyles[$this->currentRow] = $rowStyle;
            if (isset($rowStyle['options']['height'])) {
                $this->setRowHeight($this->currentRow + 1, $rowStyle['options']['height']);
            }
        }

        $this->lastTouch['area']['col_idx1'] = $this->lastTouch['area']['col_idx2'] = -1;
        $maxColIdx = max($cellStyles ? max(array_keys($cellStyles)) : 0, $rowValues ? max(array_keys($rowValues)) : 0);
        $this->_touchStart($this->currentRow, $this->offsetCol, 'row');
        for ($colIdx = 0; $colIdx <= $maxColIdx; $colIdx++) {
            if (isset($rowValues[$colIdx]) || isset($cellStyles[$colIdx])) {
                if ($this->lastTouch['area']['col_idx1'] === -1) {
                    $this->lastTouch['area']['col_idx1'] = $colIdx;
                }
                $this->lastTouch['area']['col_idx2'] = $colIdx;
                if (isset($this->fieldStyles[$rowFieldNames[$colIdx]]) && isset($cellStyles[$colIdx])) {
                    $cellComboStyle = array_merge_recursive($this->fieldStyles[$rowFieldNames[$colIdx]], $cellStyles[$colIdx]);
                }
                elseif (isset($this->fieldStyles[$rowFieldNames[$colIdx]])) {
                    $cellComboStyle = $this->fieldStyles[$rowFieldNames[$colIdx]];
                }
                elseif (isset($cellStyles[$colIdx])) {
                    $cellComboStyle = $cellStyles[$colIdx];
                }
                else {
                    $cellComboStyle = null;
                }
                //$this->_setCellData(null, $rowValues[$colIdx] ?? null, $cellComboStyle);
                $this->_setCellData(['col_idx' => $this->offsetCol + $colIdx, 'row_idx' => $this->currentRow], $rowValues[$colIdx] ?? null, $cellComboStyle);
            }
            $this->lastTouch['cell']['col_idx'] = ++$this->currentCol;
        }
        $this->_touchEnd($this->currentRow, $maxColIdx, 'row');

        //$this->withLastRow();

        return $this;
    }

    /**
     * Move to the next row
     *
     * @param array|null $options
     *
     * @return $this
     */
    public function nextRow(?array $options = []): Sheet
    {
        if (!empty($options)) {
            $this->rowStyles[$this->currentRow] = $options;
        }
        $this->_writeCurrentRow();

        return $this;
    }

    /**
     * Skip rows
     *
     * @param int|null $rowCount
     *
     * @return $this
     */
    public function skipRow(?int $rowCount = 1): Sheet
    {
        for ($i = 0; $i < $rowCount; $i++) {
            $this->nextRow();
        }

        return $this;
    }

    /**
     * Make area for writing
     *
     * @param string $range A1:Z9 or R1C1:R9C28
     *
     * @return Area
     */
    public function makeArea(string $range): Area
    {
        $area = new Area($this, $range);

        $this->areas[] = $area->setIndex(count($this->areas));
        $coord = $area->getCoord();
        if (empty($this->totalArea['coord'])) {
            $this->totalArea['coord'] = $coord;
        }
        else {
            if ($this->totalArea['coord'][0]['row'] > $coord[0]['row']) {
                $this->totalArea['coord'][0]['row'] = $coord[0]['row'];
            }
            if ($this->totalArea['coord'][0]['col'] > $coord[0]['col']) {
                $this->totalArea['coord'][0]['col'] = $coord[0]['col'];
            }
            if ($this->totalArea['coord'][1]['row'] > $coord[1]['row']) {
                $this->totalArea['coord'][1]['row'] = $coord[1]['row'];
            }
            if ($this->totalArea['coord'][1]['col'] > $coord[1]['col']) {
                $this->totalArea['coord'][1]['col'] = $coord[1]['col'];
            }
        }

        return $area;
    }

    /**
     * Begin a new area
     *
     * @param string|null $cellAddress Upper left cell of area
     *
     * @return Area
     */
    public function beginArea(string $cellAddress = null): Area
    {
        if (null === $cellAddress) {
            $cellAddress = 'A' . ($this->rowCountWritten + 1);
        }
        $dimension = Excel::rangeDimension($cellAddress, true);
        if ($dimension['rowNum1'] <= $this->rowCountWritten) {
            throw new Exception("Cannot make area from $cellAddress (row number must be greater then written rows)");
        }
        $maxCell = Excel::cellAddress(Excel::MAX_ROW, Excel::MAX_COL);

        return $this->makeArea($cellAddress . ':' . $maxCell);
    }

    /**
     * @return $this
     */
    public function endAreas(): Sheet
    {
        $this->_writeCurrentRow();

        return $this;
    }

    /**
     * @param $cellAddress
     *
     * @return array
     */
    protected function _parseAddress($cellAddress): ?array
    {
        if (is_array($cellAddress) && isset($cellAddress['row'], $cellAddress['col'])) {
            return $cellAddress;
        }

        $result = ['row' => null, 'col' => null];
        if (is_string($cellAddress)) {
            $result = $this->_rangeDimension($cellAddress);
        }
        elseif (is_array($cellAddress)) {
            if (isset($cellAddress['row'], $cellAddress['col'])) {
                $result = $cellAddress;
            }
            else {
                [$row, $col] = array_values($cellAddress);
                $result = ['row' => $row, 'col' => $col];
            }
        }

        return $result;
    }

    /**
     * @param string $cellAddress
     * @param int|null $colOffset
     * @param int|null $rowOffset
     *
     * @return array|null
     */
    protected function _rangeDimension(string $cellAddress, ?int $colOffset = 1, ?int $rowOffset = 1): ?array
    {
        if (preg_match('/^R\[?(-?\d+)?\]?C/', $cellAddress)) {
            // relative address
            $relAddress = $cellAddress;
            $cellAddress = Excel::colLetter($colOffset) . ($this->rowCountWritten + $rowOffset);
            $dimension = Excel::rangeDimensionRelative($cellAddress, $relAddress, true);
        }
        else {
            // absolute address
            $dimension = Excel::rangeDimension($cellAddress, true);
        }

        return $dimension;
    }

    /**
     * @param string|array|null $cellAddress
     * @param mixed $value
     * @param mixed|null $styles
     * @param bool|null $merge
     * @param bool|null $changeCurrent
     *
     * @return array
     */
    protected function _setCellData($cellAddress, $value, $styles = null, ?bool $merge = false, ?bool $changeCurrent = false)
    {
        $dimension = [];
        if (null === $cellAddress) {
            $rowIdx = $this->lastTouch['cell']['row_idx'];
            $colIdx = $this->lastTouch['cell']['col_idx'];
        }
        else {
            if (isset($cellAddress['row_idx'], $cellAddress['col_idx'])) {
                $rowIdx = $cellAddress['row_idx'];
                $colIdx = $cellAddress['col_idx'];
                $this->lastTouch['cell'] = ['row_idx' => $rowIdx, 'col_idx' => $colIdx];
            }
            else {
                if (isset($cellAddress['row'], $cellAddress['col'])) {
                    $row = $cellAddress['row'];
                    $col = $cellAddress['col'];
                    $dimension = $cellAddress;
                }
                else {
                    $dimension = $this->_parseAddress($cellAddress);
                    $row = $dimension['row'];
                    $col = $dimension['col'];
                    if ($merge && isset($dimension['width'], $dimension['height']) && ($dimension['width'] > 1 || $dimension['height'] > 1)) {
                        $this->mergeCells($dimension['range']);
                    }
                }

                if ($row === null || $col === null) {
                    ExceptionAddress::throwNew('Wrong cell address %s', print_r($cellAddress, 1));
                }
                if ($row < $this->currentRow) {
                    ExceptionAddress::throwNew('Row number must be greater then written rows');
                }
                $rowIdx = $row - 1;
                $colIdx = $col - 1;

                $rowCnt = isset($dimension['rowNum1'], $dimension['rowNum2']) ? $dimension['rowNum2'] - $dimension['colNum1'] : 0;
                $colCnt = isset($dimension['colNum1'], $dimension['colNum2']) ? $dimension['colNum2'] - $dimension['colNum1'] : 0;

                $this->lastTouch['cell'] = ['row_idx' => $rowIdx, 'col_idx' => $colIdx, 'dimension' => $dimension];
                $this->lastTouch['area']['col_idx1'] = $colIdx;
                $this->lastTouch['area']['row_idx2'] = $rowIdx + $rowCnt;
                $this->lastTouch['area']['col_idx2'] = $colIdx + $colCnt;
            }
        }

        if ($value !== null) {
            if (!is_scalar($value)) {
                $addr = Excel::cellAddress($colIdx + 1, $rowIdx + 1);
                Exception::throwNew('Value for cell %s must be scalar', $addr);
            }
            $this->cells['values'][$rowIdx][$colIdx] = $value;
            if ($changeCurrent) {
                $this->currentRow = $rowIdx;
                $this->currentCol = $colIdx;
            }
        }
        if ($styles) {
            $this->cells['styles'][$rowIdx][$colIdx] = Style::normalize($styles);
        }

        return $dimension;
    }

    /**
     * Set a value to the single cell or to the cell range
     *
     * $cellAddress formats:
     *      'B5'
     *      'B5:C7'
     *      ['row' => 6, 'col' => 7]
     *      [6, 7]
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     */
    public function setValue($cellAddress, $value, ?array $styles = null): Sheet
    {
        $dimension = $this->_setCellData($cellAddress, $value, $styles, true);
        if (isset($dimension['rowNum1'], $dimension['rowNum2'], $dimension['colNum1'], $dimension['colNum2'])) {
            $rowCnt = $dimension['rowNum2'] - $dimension['colNum1'];
            $colCnt = $dimension['colNum2'] - $dimension['colNum1'];

            if ($rowCnt === 0 && $colCnt === 0) {
                $this->lastTouch['row'] = 'cell';
            }
            elseif ($rowCnt === 0) {
                $this->lastTouch['row'] = 'row';
            }
            else {
                $this->lastTouch['row'] = 'area';
            }
        }
        //$this->lastTouch['row'] = 'area';

        return $this;
    }

    /**
     * Set a formula to the single cell or to the cell range
     *
     * $cellAddress formats:
     *      'B5'
     *      'B5:C7'
     *      ['row' => 6, 'col' => 7]
     *      [6, 7]
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     */
    public function setFormula($cellAddress, $value, array $styles = null): Sheet
    {
        if (empty($value)) {
            $value = null;
        }
        elseif (strpos($value, '=') !== 0) {
            $value = '=' . $value;
        }

        ///-- $styles = $styles ? Style::normalize($styles) : null;
        $this->_setCellData($cellAddress, $value, $styles, true);

        return $this;
    }

    /**
     * @param string $cellAddress
     * @param mixed $style
     * @param bool|null $mergeStyles
     *
     * @return $this
     */
    public function setStyle(string $cellAddress, $style, ?bool $mergeStyles = false): Sheet
    {
        $dimension = $this->_rangeDimension($cellAddress);
        if ($dimension['rowNum1'] <= $this->rowCountWritten) {
            throw new Exception('Row number must be greater then written rows');
        }
        $oldStyle = $style;
        $style = Style::normalize($style);
        for ($row = $dimension['rowNum1'] - 1; $row < $dimension['rowNum2']; $row++) {
            for ($col = $dimension['colNum1'] - 1; $col < $dimension['colNum2']; $col++) {
                if ($mergeStyles && isset($this->cells['styles'][$row][$col])) {
                    $this->cells['styles'][$row][$col] = array_merge($this->cells['styles'][$row][$col], $style);
                }
                else {
                    $this->cells['styles'][$row][$col] = $style;
                }
                if (!isset($this->cells['values'][$row][$col])) {
                    $this->cells['values'][$row][$col] = '';
                }
            }
        }

        return $this;
    }

    /**
     * Alias for setStyle()
     *
     * @param string $cellAddress
     * @param $style
     * @param bool|null $mergeStyles
     *
     * @return $this
     */
    public function setCellStyle(string $cellAddress, $style, ?bool $mergeStyles = false): Sheet
    {
        return $this->setStyle($cellAddress, $style, $mergeStyles);
    }

    /**
     * @param string $cellAddr
     * @param array $style
     *
     * @return $this
     */
    public function addStyle(string $cellAddr, array $style): Sheet
    {
        return $this->setStyle($cellAddr, $style, true);
    }

    /**
     * @param string $cellAddr
     * @param string $color
     *
     * @return $this
     */
    public function setBgColor(string $cellAddr, string $color): Sheet
    {
        return $this->setStyle($cellAddr, ['fill-color' => $color], true);
    }

    /**
     * @param string $cellAddr
     * @param string $format
     *
     * @return $this
     */
    public function setFormat(string $cellAddr, string $format): Sheet
    {
        return $this->setStyle($cellAddr, ['format' => $format], true);
    }

    /**
     * @param string $range
     * @param string|array $style
     *
     * @return $this
     */
    public function setOuterBorder(string $range, $style): Sheet
    {
        $borderStyle = Style::borderOptions($style);
        $this->withRange($range)->applyOuterBorder($borderStyle['border-left-style'], $borderStyle['border-left-color']);


        return $this;
    }

    /**
     * @return $this
     */
    public function writeAreas(): Sheet
    {
        $writer = $this->excel->getWriter();
        if ($this->open) {
            $this->writeAreasRows($writer);
        }
        else {
            $writer->writeSheetDataBegin($this);
        }
        return $this;
    }

    /**
     * @return $this
     */
    public function clearAreas(): Sheet
    {
        $this->cells = ['values' => [], 'styles' => []];
        $this->areas = [];
        $this->totalArea = [];

        return $this;
    }

    /**
     * Write all active areas
     *
     * @return $this
     */
    public function writeAreasRows($writer): Sheet
    {
        if (!$this->areas) {
            return $this;
        }

        if (!empty($this->cells['values']) || !empty($this->cells['styles'])) {
            $maxRow = max(array_keys($this->cells['values']) + array_keys($this->cells['styles']));
            // writes row by row
            for ($numRow = $this->rowCountWritten; $numRow <= $maxRow; $numRow++) {
                if (isset($this->cells['values'][$numRow]) || isset($this->cells['styles'][$numRow])) {
                    $colMax = 0;
                    $rowValues = $this->cells['values'][$numRow] ?? [];
                    if ($rowValues && ($keyMax = max(array_keys($rowValues))) > $colMax) {
                        $colMax = $keyMax;
                    }
                    $cellStyles = $this->cells['styles'][$numRow] ?? [];
                    if ($cellStyles && ($keyMax = max(array_keys($cellStyles))) > $colMax) {
                        $colMax = $keyMax;
                    }

                    for ($numCol = Excel::MIN_COL; $numCol <= $colMax; $numCol++) {
                        if (!isset($rowValues[$numCol])) {
                            $rowValues[$numCol] = null;
                        }
                        if (!isset($cellStyles[$numCol])) {
                            $cellStyles[$numCol] = [];
                        }
                    }
                    // array of all values
                    ksort($rowValues);
                    // array of all styles
                    ksort($cellStyles);
                    $this->_writeRow($writer, array_values($rowValues), [], $cellStyles ? array_values($cellStyles) : []);
                }
                else {
                    $this->_writeRow($writer, [null]);
                }
            }
            $this->clearAreas();
        }
        $this->currentRow = $this->rowCountWritten;
        $this->currentCol = 0;
        $this->_touchEnd($this->currentRow, $this->currentCol, 'cell');
        $this->withLastCell();

        return $this;
    }

    /**
     * @param $writer
     *
     * @return void
     */
    public function writeDataBegin($writer)
    {
        // if already initialized
        if ($this->open) {
            return;
        }

        $sheetFileName = $writer->tempFilename('xl/worksheets/' . $this->xmlName);
        $this->setFileWriter($writer::makeWriteBuffer($sheetFileName));

        $this->fileWriter->write('<sheetData>');

        $this->open = true;
        if ($this->areas) {
            $this->writeAreasRows($writer);
        }
    }

    /**
     * @return void
     */
    public function writeDataEnd()
    {
        if ($this->close) {
            return;
        }
        if ($this->areas) {
            $this->writeAreas();
        }

        if ($this->currentCol) {
            $this->_writeCurrentRow();
        }

        $this->fileWriter->flush(true);
        $this->fileWriter->write('</sheetData>');
    }

    /**
     * @return string
     */
    public function minCell(): string
    {
        return Excel::cellAddress(($this->minRow === 0) ? 1 : $this->minRow, ($this->minCol === 0) ? 1 : $this->minCol);
    }

    /**
     * @return string
     */
    public function maxCell(): string
    {
        return Excel::cellAddress(($this->maxRow === 0) ? 1 : $this->maxRow, ($this->maxCol === 0) ? 1 : $this->maxCol);
    }

    /**
     * @return bool
     */
    public function isRightToLeft(): bool
    {
        return $this->excel->isRightToLeft();
    }

    /**
     * @param array $addr
     * @param string $key
     * @param array $options
     * @param bool|null $replace
     *
     * @return void
     */
    protected function _setStyleOptions(array $addr, string $key, array $options, ?bool $replace = false)
    {
        if ($addr) {
            $rowIdx = $addr['row_idx'];
            $colIdx = $addr['col_idx'];
            $ref = 'cell';
        }
        else {
            $rowIdx = $this->lastTouch['cell']['row_idx'];
            $colIdx = $this->lastTouch['cell']['col_idx'];
            $ref = $this->lastTouch['ref'];
        }

        if ($ref === 'cell') {
            if ($replace || !isset($this->cells['styles'][$rowIdx][$colIdx][$key])) {
                $this->cells['styles'][$rowIdx][$colIdx][$key] = $options;
            }
            else {
                $this->cells['styles'][$rowIdx][$colIdx][$key] = array_replace_recursive($this->cells['styles'][$rowIdx][$colIdx][$key], $options);
            }
            if (!isset($this->cells['values'][$rowIdx][$colIdx])) {
                $this->cells['values'][$rowIdx][$colIdx] = '';
            }
        }
        elseif ($ref === 'area') {
            for ($rowIdx = $this->lastTouch['area']['row_idx1']; $rowIdx <= $this->lastTouch['area']['row_idx2']; $rowIdx++) {
                for ($colIdx = $this->lastTouch['area']['col_idx1']; $colIdx <= $this->lastTouch['area']['col_idx2']; $colIdx++) {
                    if ($replace || !isset($this->cells['styles'][$rowIdx][$colIdx][$key])) {
                        $this->cells['styles'][$rowIdx][$colIdx][$key] = $options;
                    }
                    else {
                        $this->cells['styles'][$rowIdx][$colIdx][$key] = array_replace_recursive($this->cells['styles'][$rowIdx][$colIdx][$key], $options);
                    }
                    if (!isset($this->cells['values'][$rowIdx][$colIdx])) {
                        $this->cells['values'][$rowIdx][$colIdx] = '';
                    }
                }
            }
        }
        else {
            $rowIdx = $this->lastTouch['row']['row_idx'];
            if ($replace || !isset($this->rowStyles[$rowIdx][$key])) {
                $this->rowStyles[$rowIdx][$key] = $options;
            }
            else {
                $this->rowStyles[$rowIdx][$key] = array_replace_recursive($this->rowStyles[$rowIdx][$key], $options);
            }
        }
    }

    /**
     * @param int $row
     * @param int $col
     * @param string|null $ref
     *
     * @return void
     */
    protected function _touchStart(int $row, int $col, ?string $ref = null)
    {
        $this->lastTouch['row'] = [
            'row_idx' => $row,
        ];
        $this->lastTouch['cell'] = [
            'row_idx' => $row,
            'col_idx' => $col,
        ];
        $this->lastTouch['area'] = [
            'row_idx1' => $row,
            'row_idx2' => $row,
            'col_idx1' => $col,
            'col_idx2' => $col,
        ];
        if ($ref) {
            $this->lastTouch['ref'] = $ref;
        }
    }

    /**
     * @param int $row
     * @param int $col
     * @param string|null $ref
     *
     * @return void
     */
    protected function _touchEnd(int $row, int $col, ?string $ref = null)
    {
        $this->lastTouch['row'] = [
            'row_idx' => $row,
        ];
        $this->lastTouch['cell'] = [
            'row_idx' => $row,
            'col_idx' => $col,
        ];
        $this->lastTouch['area']['row_idx2'] = $row;
        $this->lastTouch['area']['col_idx2'] = $col;

        if ($ref) {
            $this->lastTouch['ref'] = $ref;
            if ($ref === 'cell') {
                $this->lastTouch['area']['row_idx1'] = $row;
                $this->lastTouch['area']['col_idx1'] = $col;
            }
        }
    }

    /**
     * Select last written cell for applying
     *
     * @return $this
     */
    public function withLastCell(): Sheet
    {
        /*
        $this->lastTouch['row'] = [
            'row_idx' => $this->currentRow,
        ];
        if (isset($this->cells['values'][$this->currentRow])) {
            $maxCol = max(array_keys($this->cells['values'][$this->currentRow]));
        }
        else {
            $maxCol = 0;
        }
        $this->lastTouch['cell'] = [
            'row_idx' => $this->currentRow,
            'col_idx' => $maxCol,
        ];
        $this->lastTouch['area'] = [
            'row_idx1' => $this->currentRow,
            'row_idx2' => $this->currentRow,
            'col_idx1' => $maxCol,
            'col_idx2' => $maxCol,
        ];
        */
        $this->lastTouch['ref'] = 'cell';

        return $this;
    }

    /**
     * Select last written row for applying
     *
     * @return $this
     */
    public function withLastRow(): Sheet
    {
        /*
        $this->lastTouch['cell'] = [
            'row_idx' => $this->currentRow,
            'col_idx' => $this->currentCol,
        ];
        $this->lastTouch['row'] = ['row_idx' => $this->currentRow];
        $this->lastTouch['area'] = [
            'row_idx1' => $this->currentRow,
            'row_idx2' => $this->currentRow,
            'col_idx1' => $this->currentCol,
            'col_idx2' => $this->currentCol,
        ];
        */
        $this->lastTouch['ref'] = 'row';

        return $this;
    }

    /**
     * Select custom range for applying
     *
     * @param array|string $range
     *
     * @return $this
     */
    public function withRange($range): Sheet
    {
        $dimension = self::_rangeDimension($range);
        if ($dimension['rowNum1'] <= $this->rowCountWritten) {
            throw new Exception('Row number must be greater then written rows');
        }

        $this->lastTouch['area'] = [
            'row_idx1' => $dimension['rowNum1'] - 1,
            'row_idx2' => $dimension['rowNum2'] - 1,
            'col_idx1' => $dimension['colNum1'] - 1,
            'col_idx2' => $dimension['colNum2'] - 1,
        ];
        $this->lastTouch['ref'] = 'area';

        return $this;
    }

    /**
     * Define named range
     *
     * @param string $range
     * @param string $name
     *
     * @return $this
     */
    public function addNamedRange(string $range, string $name): Sheet
    {
        $dimension = self::_rangeDimension($range);
        if (isset($dimension['absAddress'])) {
            if (!preg_match('/^\w+$/u', $name)) {
                ExceptionRangeName::throwNew('Wrong name for range');
            }
            if (mb_strlen($name) > 255) {
                ExceptionRangeName::throwNew('Name for range cannot be more then 255');
            }
            foreach ($this->excel->getSheets() as $sheet) {
                foreach ($sheet->getNamedRanges() as $range) {
                    if (mb_strtolower($range['name']) === mb_strtolower($name)) {
                        ExceptionRangeName::throwNew('Named range "' . $name . '" already exists on sheet "' . $sheet->sheetName . '"');
                    }
                }
            }
            $this->namedRanges[] = ['range' => $dimension['absAddress'], 'name' => $name];
            $this->_setDimension($dimension['rowNum1'], $dimension['colNum1']);
            $this->_setDimension($dimension['rowNum2'], $dimension['colNum2']);
        }

        return $this;
    }

    /**
     * Returns named ranges with full addresses
     *
     * @return array
     */
    public function getNamedRanges(): array
    {
        return $this->namedRanges;
    }

    /**
     * Add note to the sheet
     * $sheet->addNote('A1', $noteText, $noteStyle)
     * $sheet->writeCell($cellValue)->addNote($noteText, $noteStyle)
     *
     * @param string $cell
     * @param string|array|null $comment
     * @param array $noteStyle
     *
     * @return $this
     */
    public function addNote(string $cell, $comment = null, array $noteStyle = []): Sheet
    {
        if (func_num_args() === 1 || (func_num_args() === 2 && is_array($comment)) ) {
            if ( func_num_args() === 2) {
                $noteStyle = $comment;
            }
            $comment = $cell;
            $rowIdx = $this->lastTouch['cell']['row_idx'];
            $colIdx = $this->lastTouch['cell']['col_idx'];
            $cell = Excel::cellAddress($rowIdx + 1, $colIdx + 1);
        }
        else {
            $dimension = self::_rangeDimension($cell);
            $cell = $dimension['cell1'];
            $rowIdx = $dimension['rowIndex'];
            $colIdx = $dimension['colIndex'];
        }
        if ($cell) {
            $marginLeft = number_format(self::NOTE_LEFT_OFFSET + self::NOTE_LEFT_INC * ($colIdx + 1), 2, '.', '') . 'pt';
            if ($rowIdx === 0) {
                $marginTop = self::NOTE_DEFAULT_TOP;
            }
            else {
                $marginTop = number_format(self::NOTE_TOP_OFFSET + self::NOTE_TOP_INC * $rowIdx, 2, '.', '') . 'pt';
            }
            if (!empty($noteStyle['fill_color'])) {
                $noteStyle['fill_color'] = '#' . substr(Style::normalizeColor($noteStyle['fill_color']), 2);
            }
            elseif (!empty($noteStyle['bg_color'])) {
                $noteStyle['fill_color'] = '#' . substr(Style::normalizeColor($noteStyle['bg_color']), 2);
            }
            if (!empty($noteStyle['width']) && (is_int($noteStyle['width']) || is_float($noteStyle['width']))) {
                $noteStyle['width'] = number_format($noteStyle['width'], 2, '.', '') . 'pt';
            }
            if (!empty($noteStyle['height']) && (is_int($noteStyle['height']) || is_float($noteStyle['height']))) {
                $noteStyle['height'] = number_format($noteStyle['height'], 2, '.', '') . 'pt';
            }
            $this->notes[$cell] = [
                'cell' => $cell,
                'row_index' => $rowIdx,
                'col_index' => $colIdx,
                'text' => htmlspecialchars($comment),
                'style' => array_merge( [
                    'width' => self::NOTE_DEFAULT_WIDTH,
                    'height' => self::NOTE_DEFAULT_HEIGHT,
                    'margin_left' => $marginLeft,
                    'margin_top' => $marginTop,
                    'fill_color' => self::NOTE_DEFAULT_COLOR,
                ], $noteStyle ),
            ];
            $this->_setDimension($rowIdx + 1, $colIdx + 1);
            if (!isset($this->relationships['legacyDrawing'])) {
                $file = 'vmlDrawing' . $this->index . '.vml';
                $this->relationships['legacyDrawing'][++$this->relationshipId] = [
                    'file' => $file,
                    'link' => '../drawings/' . $file,
                    'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
                    'extra' => '',
                ];
            }
            if (!isset($this->relationships['comments'])) {
                $file = 'comments' . $this->index . '.xml';
                $this->relationships['comments'][++$this->relationshipId] = [
                    'file' => $file,
                    'link' => '../' . $file,
                    'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
                    'extra' => '',
                ];
            }
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getNotes(): array
    {
        return $this->notes;
    }

    public function getLegacyDrawingId()
    {
        if (isset($this->relationships['legacyDrawing'])) {
            return array_key_first($this->relationships['legacyDrawing']);
        }

        return 0;
    }

    public function getDrawingId()
    {
        if (isset($this->relationships['drawing'])) {
            return array_key_first($this->relationships['drawing']);
        }

        return 0;
    }

    // === IMAGES === //

    /**
     * Add image to the sheet
     * $sheet->addImage('A1', 'path/to/file')
     * $sheet->addImage('A1', 'path/to/file', ['width => 100])
     *
     * @param string $cell
     * @param string $imageFile
     * @param array|null $imageStyle
     *
     * @return $this
     */
    public function addImage(string $cell, string $imageFile, ?array $imageStyle = []): Sheet
    {
        if (func_num_args() === 1) {
            $imageFile = $cell;
            $rowIdx = $this->lastTouch['cell']['row_idx'];
            $colIdx = $this->lastTouch['cell']['col_idx'];
            $cell = Excel::cellAddress($rowIdx + 1, $colIdx + 1);
        }
        else {
            $dimension = self::_rangeDimension($cell);
            $cell = $dimension['cell1'];
            $rowIdx = $dimension['rowIndex'];
            $colIdx = $dimension['colIndex'];
        }
        if (!is_file($imageFile)) {
            ExceptionFile::throwNew('Image file "%s" does not exist', $imageFile);
        }

        if ($cell) {
            $imageData = $this->excel->loadImageFile($imageFile);
            if ($imageData) {
                $imageData['cell'] = $cell;
                $imageData['row_index'] = $rowIdx;
                $imageData['col_index'] = $colIdx;
                if (!empty($imageStyle['width']) || !empty($imageStyle['height'])) {
                    if (!empty($imageStyle['width']) && empty($imageStyle['height'])) {
                        $ratio = $imageStyle['width'] / $imageData['width'];
                        $imageData['width'] = $imageStyle['width'];
                        $imageData['height'] = (int)round($imageData['height'] * $ratio);
                    }
                    elseif (empty($imageStyle['width']) && !empty($imageStyle['height'])) {
                        $ratio = $imageStyle['height'] / $imageData['height'];
                        $imageData['width'] = (int)round($imageData['width'] * $ratio);
                        $imageData['height'] = $imageStyle['height'];
                    }
                    else {
                        $imageData['width'] = $imageStyle['width'];
                        $imageData['height'] = $imageStyle['height'];
                    }
                }
                $this->media['images'][] = $imageData;
                $this->_setDimension($rowIdx + 1, $colIdx + 1);

                if (!isset($this->relationships['drawing'])) {
                    $file = 'drawing' . $this->index . '.xml';
                    $this->relationships['drawing'][++$this->relationshipId] = [
                        'file' => $file,
                        'link' => '../drawings/' . $file,
                        'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                        'extra' => '',
                    ];
                }
            }
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getImages()
    {

        return $this->media['images'] ?? [];
    }

    // === DESIGN STYLES === //

    /**
     * Sets height to the current row
     *
     * @param float $height
     *
     * @return $this
     */
    public function applyRowHeight(float $height): Sheet
    {
        $this->setRowHeight($this->currentRow + 1, $height);

        return $this;
    }

    /**
     * Sets all borders style
     *
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function applyBorder(string $style, ?string $color = '#000000'): Sheet
    {
        $options = [
            'border-left-style' => $style,
            'border-left-color' => $color,
            'border-right-style' => $style,
            'border-right-color' => $color,
            'border-top-style' => $style,
            'border-top-color' => $color,
            'border-bottom-style' => $style,
            'border-bottom-color' => $color,
            'border-diagonal-up' => 0,
            'border-diagonal-down' => 0,
        ];

        $this->_setStyleOptions([], Style::BORDER, $options, true);

        return $this;
    }

    /**
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function applyBorderLeft(string $style, ?string $color = '#000000'): Sheet
    {
        $options = [
            'border-left-style' => $style,
            'border-left-color' => $color,
        ];
        for ($rowIdx = $this->lastTouch['area']['row_idx1']; $rowIdx <= $this->lastTouch['area']['row_idx2']; $rowIdx++) {
            $addr = [
                'row_idx' => $rowIdx,
                'col_idx' => $this->lastTouch['area']['col_idx1'],
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);
        }

        return $this;
    }

    /**
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function applyBorderRight(string $style, ?string $color = '#000000'): Sheet
    {
        $options = [
            'border-right-style' => $style,
            'border-right-color' => $color,
        ];
        for ($rowIdx = $this->lastTouch['area']['row_idx1']; $rowIdx <= $this->lastTouch['area']['row_idx2']; $rowIdx++) {
            $addr = [
                'row_idx' => $rowIdx,
                'col_idx' => $this->lastTouch['area']['col_idx2'],
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);
        }

        return $this;
    }

    /**
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function applyBorderTop(string $style, ?string $color = '#000000'): Sheet
    {
        $options = [
            'border-top-style' => $style,
            'border-top-color' => $color,
        ];
        for ($colIdx = $this->lastTouch['area']['col_idx1']; $colIdx <= $this->lastTouch['area']['col_idx2']; $colIdx++) {
            $addr = [
                'row_idx' => $this->lastTouch['area']['row_idx1'],
                'col_idx' => $colIdx,
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);
        }

        return $this;
    }

    /**
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function applyBorderBottom(string $style, ?string $color = '#000000'): Sheet
    {
        $options = [
            'border-bottom-style' => $style,
            'border-bottom-color' => $color,
        ];
        for ($colIdx = $this->lastTouch['area']['col_idx1']; $colIdx <= $this->lastTouch['area']['col_idx2']; $colIdx++) {
            $addr = [
                'row_idx' => $this->lastTouch['area']['row_idx2'],
                'col_idx' => $colIdx,
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);
        }

        return $this;
    }

    /**
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function applyOuterBorder(string $style, ?string $color = '#000000'): Sheet
    {
        if ($this->lastTouch['area']['row_idx1'] === $this->lastTouch['area']['row_idx2']
            && $this->lastTouch['area']['col_idx1'] === $this->lastTouch['area']['col_idx2']) {
            $options = [
                'border-left-style' => $style,
                'border-left-color' => $color,
                'border-right-style' => $style,
                'border-right-color' => $color,
                'border-top-style' => $style,
                'border-top-color' => $color,
                'border-bottom-style' => $style,
                'border-bottom-color' => $color,
            ];
            $addr = [
                'row_idx' => $this->lastTouch['area']['row_idx1'],
                'col_idx' => $this->lastTouch['area']['col_idx1'],
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);
        }
        else {
            $rowMin = $this->lastTouch['area']['row_idx1'];
            $rowMax = $this->lastTouch['area']['row_idx2'];
            $colMin = $this->lastTouch['area']['col_idx1'];
            $colMax = $this->lastTouch['area']['col_idx2'];

            $options = [
                'border-left-style' => $style,
                'border-left-color' => $color,
                'border-top-style' => $style,
                'border-top-color' => $color,
            ];
            $addr = [
                'row_idx' => $rowMin,
                'col_idx' => $colMin,
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);

            $options = [
                'border-top-style' => $style,
                'border-top-color' => $color,
                'border-right-style' => $style,
                'border-right-color' => $color,
            ];
            $addr = [
                'row_idx' => $rowMin,
                'col_idx' => $colMax,
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);

            $options = [
                'border-right-style' => $style,
                'border-right-color' => $color,
                'border-bottom-style' => $style,
                'border-bottom-color' => $color,
            ];
            $addr = [
                'row_idx' => $rowMax,
                'col_idx' => $colMax,
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);

            $options = [
                'border-bottom-style' => $style,
                'border-bottom-color' => $color,
                'border-left-style' => $style,
                'border-left-color' => $color,
            ];
            $addr = [
                'row_idx' => $rowMax,
                'col_idx' => $colMin,
            ];
            $this->_setStyleOptions($addr, Style::BORDER, $options);

            if ($colMax - $colMin > 1) {
                $options1 = [
                    'border-top-style' => $style,
                    'border-top-color' => $color,
                ];
                $options2 = [
                    'border-bottom-style' => $style,
                    'border-bottom-color' => $color,
                ];
                for ($colIdx = $colMin + 1; $colIdx < $colMax; $colIdx++) {
                    $addr['col_idx'] = $colIdx;
                    $addr['row_idx'] = $rowMin;
                    $this->_setStyleOptions($addr, Style::BORDER, $options1);

                    $addr['row_idx'] = $rowMax;
                    $this->_setStyleOptions($addr, Style::BORDER, $options2);
                }
            }

            if ($rowMax - $rowMin > 1) {
                $options1 = [
                    'border-left-style' => $style,
                    'border-left-color' => $color,
                ];
                $options2 = [
                    'border-right-style' => $style,
                    'border-right-color' => $color,
                ];
                for ($rowIdx = $rowMin + 1; $rowIdx < $rowMax; $rowIdx++) {
                    $addr['row_idx'] = $rowIdx;
                    $addr['col_idx'] = $colMin;
                    $this->_setStyleOptions($addr, Style::BORDER, $options1);
                    $addr['col_idx'] = $colMax;
                    $this->_setStyleOptions($addr, Style::BORDER, $options2);
                }
            }

        }

        return $this;
    }

    /**
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function applyInnerBorder(string $style, ?string $color = '#000000'): Sheet
    {
        if ($this->lastTouch['area']['row_idx1'] < $this->lastTouch['area']['row_idx2']
            || $this->lastTouch['area']['col_idx1'] < $this->lastTouch['area']['col_idx2']) {

            $rowMin = $this->lastTouch['area']['row_idx1'];
            $rowMax = $this->lastTouch['area']['row_idx2'];
            $colMin = $this->lastTouch['area']['col_idx1'];
            $colMax = $this->lastTouch['area']['col_idx2'];

            $options = [
                'border-right-style' => $style,
                'border-right-color' => $color,
            ];
            for ($colIdx = $colMin; $colIdx < $colMax; $colIdx++) {
                $addr = [
                    'row_idx' => $rowMax,
                    'col_idx' => $colIdx,
                ];
                $this->_setStyleOptions($addr, 'border', $options);
            }

            $options1 = [
                'border-right-style' => $style,
                'border-right-color' => $color,
                'border-bottom-style' => $style,
                'border-bottom-color' => $color,
            ];
            $options2 = [
                'border-bottom-style' => $style,
                'border-bottom-color' => $color,
            ];
            for ($rowIdx = $rowMin; $rowIdx < $rowMax; $rowIdx++) {
                for ($colIdx = $colMin; $colIdx < $colMax; $colIdx++) {
                    $addr = [
                        'row_idx' => $rowIdx,
                        'col_idx' => $colIdx,
                    ];
                    $this->_setStyleOptions($addr, 'border', $options1);
                }
                $addr = [
                    'row_idx' => $rowIdx,
                    'col_idx' => $this->lastTouch['area']['col_idx2'],
                ];
                $this->_setStyleOptions($addr, 'border', $options2);
            }
        }

        return $this;
    }

    /**
     * @param string $fontName
     * @param int|null $fontSize
     * @param string|null $fontStyle
     * @param string|null $fontColor
     *
     * @return $this
     */
    public function applyFont(string $fontName, ?int $fontSize = null, ?string $fontStyle = null, ?string $fontColor = null): Sheet
    {
        $font = ['font-name' => $fontName];
        if ($fontSize) {
            $font = ['font-size' => $fontSize];
        }
        if ($fontStyle) {
            $font = ['font-style' => $fontStyle];
        }
        if ($fontColor) {
            $font = ['font-color' => $fontColor];
        }

        $this->_setStyleOptions([], 'font', $font);

        return $this;
    }


    /**
     * @param string $fontName
     *
     * @return $this
     */
    public function applyFontName(string $fontName): Sheet
    {
        $this->_setStyleOptions([], 'font', ['font-name' => $fontName]);

        return $this;
    }

    /**
     * @param float $fontSize
     *
     * @return $this
     */
    public function applyFontSize(float $fontSize): Sheet
    {
        $this->_setStyleOptions([], 'font', ['font-size' => $fontSize]);

        return $this;
    }

    /**
     * @param string $fontStyle
     *
     * @return $this
     */
    public function applyFontStyle(string $fontStyle): Sheet
    {
        $this->_setStyleOptions([], 'font', ['font-style-' . strtolower($fontStyle) => 1]);

        return $this;
    }

	public function applyTextRotation(int $degrees): Sheet
	{
		$this->_setStyleOptions([], 'format', [ 'format-text-rotation' => $degrees ] );

		return $this;
	}

    public function applyFontStyleBold(): Sheet
    {
        return $this->applyFontStyle('bold');
    }

    /**
     * @return $this
     */
    public function applyFontStyleItalic(): Sheet
    {
        return $this->applyFontStyle('italic');
    }

    /**
     * @param bool|null $double
     *
     * @return $this
     */
    public function applyFontStyleUnderline(?bool $double = false): Sheet
    {
        $this->_setStyleOptions([], 'font', ['font-style-underline' => $double ? 2 : 1]);

        return $this;
    }

    /**
     * @return $this
     */
    public function applyFontStyleStrikethrough(): Sheet
    {
        return $this->applyFontStyle('strikethrough');
    }

    /**
     * @param string $fontColor
     *
     * @return $this
     */
    public function applyFontColor(string $fontColor): Sheet
    {
        $this->_setStyleOptions([], 'font', ['font-color' => $fontColor]);

        return $this;
    }

    /**
     * Alias of 'setFontColor()'
     *
     * @param string $color
     *
     * @return $this
     */
    public function applyColor(string $color): Sheet
    {
        return $this->applyFontColor($color);
    }

    /**
     * Alias of 'backgroundColor()'
     *
     * @param string $color
     *
     * @return $this
     */
    public function applyFillColor(string $color): Sheet
    {
        $this->_setStyleOptions([], 'fill', ['fill-color' => $color]);

        return $this;
    }

    /**
     * @param string $color
     *
     * @return $this
     */
    public function applyBgColor(string $color): Sheet
    {
        return $this->applyFillColor($color);
    }

    /**
     * @param string $textAlign
     * @param string|null $verticalAlign
     *
     * @return $this
     */
    public function applyTextAlign(string $textAlign, ?string $verticalAlign = null): Sheet
    {
        $options = ['format-align-horizontal' => $textAlign];
        if ($verticalAlign !== null) {
            $options['format-align-vertical'] = $verticalAlign;
        }
        $this->_setStyleOptions([], 'format', $options);

        return $this;
    }

    /**
     * @param string $verticalAlign
     *
     * @return $this
     */
    public function applyVerticalAlign(string $verticalAlign): Sheet
    {
        $this->_setStyleOptions([], 'format', ['vertical-align' => $verticalAlign]);

        return $this;
    }

    /**
     * @return $this
     */
    public function applyTextCenter(): Sheet
    {
        $this->_setStyleOptions([], 'format', ['format-align-horizontal' => 'center', 'format-align-vertical' => 'center']);

        return $this;
    }

    /**
     * @param bool $textWrap
     *
     * @return $this
     */
    public function applyTextWrap(bool $textWrap): Sheet
    {
        $this->_setStyleOptions([], 'format', ['format-text-wrap' => 1]);

        return $this;
    }

    /**
     * @param string $color
     *
     * @return $this
     */
    public function applyTextColor(string $color): Sheet
    {
        $this->_setStyleOptions([], 'font', ['font-color' => $color]);

        return $this;
    }


}

// EOF