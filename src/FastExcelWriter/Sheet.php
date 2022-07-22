<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;
use avadim\FastExcelWriter\Exception\ExceptionAddress;

/**
 * Class Sheet
 *
 * @package avadim\FastExcelWriter
 */
class Sheet
{
    // constants for auo width
    protected const WIDTH_LOWER_CASE_LETTER = 1.00;
    protected const WIDTH_UPPER_CASE_LETTER = 1.26;
    protected const WIDTH_DOTS_SYMBOLS = 0.20;
    protected const WIDTH_PADDING = 0.56;

    /** @var Excel */
    public Excel $excel;

    /** @var int Index of the sheet */
    public int $index;

    /** @var string Key of the sheet */
    public string $key;

    public $active = false;
    public $fileName = '';
    public $sheetName = '';
    public $xmlName = '';

    public $rowCount = 0;
    public $colCount = 0;

    /** @var WriterBuffer */
    public $fileWriter = null;

    // ZERO based
    public $freezeRows = 0;
    public $freezeColumns = 0;
    public $autoFilter = 0;
    public $absoluteAutoFilter = '';

    // ZERO based
    public $colWidths = [];
    public $colFormulas = [];
    public $colStyles = [];

    // ZERO based
    public $rowHeights = [];
    public $rowStyles = [];

    protected $currentRow = Excel::MIN_ROW;
    protected $currentCol = Excel::MIN_COL;

    // ZERO based
    protected $cells = [];

    public $open = false;
    public $close = false;

    protected $mergeCells = [];
    protected $totalArea = [];
    protected $areas = [];
    protected $defaultStyle = [];

    protected $pageOptions = [];

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
     * @param WriterBuffer $fileWriter
     *
     * @return $this
     */
    public function setFileWriter($fileWriter)
    {
        if (!$this->fileWriter) {
            $this->fileWriter = $fileWriter;
            $this->fileName = $fileWriter->getFileName();
        }

        return $this;
    }

    /**
     * @param WriterBuffer $fileWriter
     *
     * @return $this
     */
    public function resetFileWriter($fileWriter)
    {
        $this->fileWriter = $fileWriter;
        $this->fileName = $fileWriter->getFileName();

        return $this;
    }

    /**
     * Set sheet name
     *
     * @param string $sheetName
     *
     * @return $this
     */
    public function setName(string $sheetName)
    {
        $this->sheetName = $sheetName;

        return $this;
    }

    /**
     * Get sheet name
     *
     * @return string
     */
    public function getName()
    {
        return $this->sheetName;
    }

    /**
     * @param string $option
     * @param mixed $value
     *
     * @return $this
     */
    public function setPageOptions(string $option, $value)
    {
        if ($this->rowCount) {
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
    public function pageOrientationPortrait()
    {
        return $this->setPageOptions('orientation', 'portrait');
    }

    /**
     * @return $this
     */
    public function pageOrientationLandscape()
    {
        return $this->setPageOptions('orientation', 'landscape');
    }

    /**
     * @param int|string|null $numPage
     *
     * @return $this
     */
    public function pageFitToWidth($numPage = 1)
    {
        if ($numPage === 'auto') {
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
    public function pageFitToHeight($numPage = 1)
    {
        if ($numPage === 'auto') {
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
    public function getPageOrientation()
    {
        return $this->pageOptions['orientation'] ?? 'portrait';
    }

    /**
     * @return int
     */
    public function getPageFitToWidth()
    {
        return (int)($this->pageOptions['fit_width'] ?? 0);
    }

    /**
     * @return int
     */
    public function getPageFitToHeight()
    {
        return (int)($this->pageOptions['fit_height'] ?? 0);
    }

    /**
     * @return bool
     */
    public function getPageFit()
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
    public function setFreeze($freezeRows, $freezeColumns = null)
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
    public function setFreezeRows(int $freezeRows)
    {
        $this->freezeRows = ($freezeRows > 0) ? $freezeRows : 0;

        return $this;
    }

    /**
     * @param int $freezeColumns Number columns to freeze
     *
     * @return $this
     */
    public function setFreezeColumns(int $freezeColumns)
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
    public function setAutofilter(?int $row = 1, ?int $col = 1)
    {
        if ($row >= 0) {
            if (empty($row)) {
                $this->autoFilter = false;
            } else {
                $this->autoFilter = Excel::cellAddress($row, $col);
            }
        }
        $this->absoluteAutoFilter = Excel::cellAddress($row, $col, true);

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
     * Set width of single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param int|float|string $width
     *
     * @return $this
     */
    public function setColWidth($col, $width)
    {
        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                if ($width === 'auto') {
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

        return $this;
    }

    /**
     * Set width of single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     *
     * @return $this
     */
    public function setColWidthAuto($col)
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
    public function setColStyle($col, $style)
    {
        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                if (!empty($this->colStyles[$colIdx])) {
                    $this->colStyles[$colIdx] = array_replace_recursive($this->colStyles[$colIdx], $style);
                }
                else {
                    $this->colStyles[$colIdx] = $style;
                }
            }
        }

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
    public function setColFormula($col, string $formula)
    {
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
    public function setColFormat($col, $format)
    {
        $this->setColStyle($col, ['format' => $format]);

        return $this;
    }

    /**
     * Set options of columns (widths, styles, formats, etc)
     *
     * Call examples:
     *  setColOptions('B', ['width' = 20]) - options for column 'B'
     *  setColOptions('B:D', ['width' = 'auto']) - options for range of columns
     *  setColOptions(['B' => ['width' = 20], 'C' => ['color' = '#f00']]) - options for several columns 'B' and 'C'
     *
     * @param mixed $arg1
     * @param array|null $arg2
     *
     * @return $this
     */
    public function setColOptions($arg1, array $arg2 = null)
    {
        if ($arg2 === null) {
            $options = $arg1;
            foreach ($options as $col => $colOptions) {
                $options[$col] = Style::normalize($colOptions);
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
        if ($options) {
            $options = $this->normalizeColKeys($options);
            foreach($options as $col => $colOptions) {
                $style = [];
                foreach($colOptions as $optionName => $optionValue) {
                    if ($optionName === 'width') {
                        $this->setColWidth($col, $optionValue);
                    }
                    elseif ($optionName === 'formula') {
                        $this->setColFormula($col, $optionValue);
                    }
                    else {
                        $style[$optionName] = $optionValue;
                    }
                }
                if ($style) {
                    $this->setColStyle($col, $style);
                }
            }
        }

        return $this;
    }

    /**
     * Height of a specific line
     *
     * @param $rowNum
     * @param $height
     *
     * @return $this
     */
    public function setRowHeight($rowNum, $height)
    {
        if (is_numeric($rowNum)) {
            $this->rowHeights[(int)$rowNum - 1] = str_replace(',', '.', (float)$height);
        }
        return $this;
    }

    /**
     * Multiple line height
     *
     * @param array $heights
     *
     * @return $this
     */
    public function setRowHeights(array $heights)
    {
        foreach ($heights as $rowNum => $rowHeight) {
            $this->setRowHeight($rowNum, $rowHeight);
        }
        return $this;
    }

    /**
     * setRowOptions(3, ['height' = 20]) - options for row number 3
     * setRowOptions([3 => ['height' = 20], 4 => ['color' = '#f00']]) - options for several rows
     * setRowOptions('2:5', ['color' = '#f00']) - options for range of rows
     *
     * @param mixed $arg1
     * @param array|null $arg2
     *
     * @return $this
     */
    public function setRowOptions($arg1, array $arg2 = null)
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
            if (isset($rowOptions['height'])) {
                $this->setRowHeight($rowNum, $rowOptions['height']);
                unset($rowOptions['height']);
            }
            if ($rowOptions) {
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

    /**
     * @param $style
     *
     * @return $this
     */
    public function setDefaultStyle($style)
    {
        if ($style) {
            $this->defaultStyle = Style::normalize($style);
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getDefaultStyle()
    {
        return $this->defaultStyle;
    }

    /**
     * @param Writer $writer
     * @param array|null $row Values of all cells of row (incl. empty)
     * @param array|null $rowOptions Specified style for the row
     * @param array|null $cellsOptions Styles of all cells of row (incl. empty)
     */
    protected function _writeRow(Writer $writer, array $row = [], array $rowOptions = null, array $cellsOptions = null)
    {
        $rowAttr = '';
        if (!empty($rowOptions['height'])) {
            $height = $rowOptions['height'];
        }
        elseif (isset($this->rowHeights[$this->rowCount])) {
            $height = $this->rowHeights[$this->rowCount];
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
        $this->fileWriter->write('<row r="' . ($this->rowCount + 1) . '" outlineLevel="0" ' . $rowAttr . '>');

        // add auto formulas of columns
        if ($this->colFormulas && $row) {
            foreach($this->colFormulas as $colIdx => $formula) {
                if (!isset($row[$colIdx])) {
                    $row[$colIdx] = $formula;
                }
            }
            ksort($row);
        }

        $rowIdx = $this->rowCount;
        foreach ($row as $colIdx => $cellValue) {
            $styleStack = [
                $this->defaultStyle ?? null, // default style of sheet
                $this->colStyles[$colIdx] ?? null, // defined style of column
                $this->rowStyles[$rowIdx] ?? null, // defined style of row
                $this->cells['styles'][$rowIdx][$colIdx] ?? null, // defined style of cell
                $rowOptions ?? null, // runtime style of tje row
                $cellsOptions[$rowIdx] ?? null, // runtime style of the cell
            ];
            $cellStyle = Style::mergeStyles($styleStack);
            $cellStyleIdx = $this->excel->style->addStyle($cellStyle, $resultStyle);

            $numberFormat = $resultStyle['number_format'];
            $numberFormatType = $resultStyle['number_format_type'];

            if (!empty($cellStyle['options']['width-auto'])) {
                $this->_columnWidth($colIdx, $cellValue, $numberFormat, $resultStyle ?? []);
            }
            $writer->_writeCell($this->fileWriter, $rowIdx + 1, $colIdx + 1, $cellValue, $numberFormatType, $cellStyleIdx);
            $colIdx++;
            if ($colIdx > $this->colCount) {
                $this->colCount = $colIdx;
            }
        }
        $this->fileWriter->write('</row>');
        $this->rowCount++;
    }

    /**
     * @param $str
     * @param $fontSize
     *
     * @return float
     */
    protected function _calcWidth($str, $fontSize)
    {
        $len = mb_strlen($str);
        $upperCount = 0;
        $dotsCount = 0;
        if (preg_match_all("/[[:upper:]#@02-9]/u", $str, $matches)) {
            $upperCount = count($matches[0]);
        }
        if (preg_match_all("/[,\.\-\+]/u", $str, $matches)) {
            $dotsCount = count($matches[0]);
        }
        $k = $fontSize / 10;

        return ($len - $upperCount - $dotsCount) * self::WIDTH_LOWER_CASE_LETTER * $k +
            $upperCount * self::WIDTH_UPPER_CASE_LETTER * $k +
            $dotsCount * self::WIDTH_DOTS_SYMBOLS * $k + self::WIDTH_PADDING;
    }

    /**
     * @param $colNum
     * @param $cellValue
     * @param $numberFormat
     * @param $style
     */
    protected function _columnWidth($colNum, $cellValue, $numberFormat, $style)
    {
        if ($cellValue) {
            $fontSize = $style['font']['size'] ?? 10;
            $len = $this->_calcWidth($cellValue, $fontSize);
            if ($numberFormat !== 'GENERAL') {
                $len = max($len, $this->_calcWidth(str_replace('\\', '', $numberFormat), $fontSize));
            }
            if (empty($this->colWidths[$colNum]) || $this->colWidths[$colNum] < $len) {
                $this->colWidths[$colNum] = $len;
            }
        }
    }

    /**
     * @param int $colNumber
     * @param array $options
     * @param array|null $defaultOptions
     * @param int|null $keyOffset
     *
     * @return array
     */
    protected function _parseColumnOptions(int $colNumber, array $options, array $defaultOptions = null, ?int $keyOffset = 1)
    {
        $colOptions = [];
        foreach($options as $key => $val) {
            if (is_int($key)) {
                $colIndex = $key + $keyOffset;
            }
            else {
                $colIndex = Excel::colNumber($key) + $keyOffset - 1;
            }
            if ($colIndex > 0) {
                if ($defaultOptions) {
                    $colOptions[$colIndex] = array_merge($defaultOptions, $val);
                }
                else {
                    $colOptions[$colIndex] = $val;
                }
            } else {
                $colOptions = [];
                break;
            }
        }
        if (!$colOptions) {
            $result = array_fill(0, $colNumber, $options);
        } else {
            $result = [];
            for($colNum = 0; $colNum < $colNumber; $colNum++) {
                $result[$colNum] = $colOptions[$colNum + 1] ?? $defaultOptions;
            }
        }
        return $result;
    }

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
    public function writeCell($value, array $styles = null)
    {
        $styles = $styles ? Style::normalize($styles) : null;
        if ($this->currentRow < $this->rowCount) {
            $this->currentRow = $this->rowCount;
        }
        $cellAddress = [
            'row' => 1 + $this->currentRow,
            'col' => 1 + $this->currentCol++,
        ];
        $this->_setCellData($cellAddress, $value, $styles, false);

        return $this;
    }

    /**
     * @return $this
     *
     * @throws \Exception
     */
    public function nextCell()
    {
        $this->writeCell(null);

        return $this;
    }

    /**
     * writeHeader(['title1', 'title2', 'title3']) - texts for cells of header
     * writeHeader(['title1' => 'text', 'title2' => 'YYYY-MM-DD', 'title3' => ['format' => ..., 'font' => ...]]) - texts and formats of columns
     * writeHeader([...], [...]) - texts and formats of columns and options of row
     *
     * @param array $header
     * @param array|null $rowStyle
     *
     * @return $this
     */
    public function writeHeader(array $header, array $rowStyle = null)
    {
        $rowValues = [];
        $colStyles = [];
        $colNum = 0;
        foreach($header as $key => $val) {
            if (!is_int($key)) {
                $rowValues[$colNum] = $key;
                //$colStyles[$colNum]['format'] = $val;
                if (is_scalar($val)) {
                    $colStyles[$colNum]['format'] = $val;
                }
                else {
                    $colStyles[$colNum] = $val;
                }
            }
            else {
                $rowValues[$colNum] = $val;
                $colStyles[$colNum] = null;
            }
            $colNum++;
        }
        $this->writeRow($rowValues, $rowStyle);
        $this->setColOptions($colStyles);

        return $this;
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
    public function writeTo($cellAddress, $value, ?array $styles = [])
    {
        $address = $this->_parseAddress($cellAddress);
        if (!isset($address['row'], $address['col'])) {
            ExceptionAddress::throwNew('Wrong cell address %s', print_r($address));
        }
        else {
            $cellAddress = $address;
        }

        while ($this->currentRow < $cellAddress['row'] - 1) {
            $this->nextRow();
        }

        $styles = $styles ? Style::normalize($styles) : null;
        $this->_setCellData($cellAddress, $value, $styles, false, true);
        if (isset($address['width'], $address['range']) && $address['width'] > 1) {
            $this->mergeCells($address['range']);
            $this->currentCol += $address['width'];
        }
        else {
            $this->currentCol++;
        }

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
    public function mergeCells($rangeSet)
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
    public function mergeRelCells($rangeSet)
    {
        if (is_int($rangeSet)) {
            $rangeSet = 'A' . $this->rowCount . ':' . Excel::colLetter($rangeSet)  . $this->rowCount;
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
     * @return array
     */
    public function getMergedCells()
    {
        return array_keys($this->mergeCells);
    }

    /**
     * @param array|null $options
     *
     * @return $this
     */
    public function nextRow(?array $options = [])
    {
        $styles = $this->cells['styles'][$this->currentRow] ?? [];
        if (empty($options)) {
            $rowStyles = $styles;
        }
        elseif (empty($styles)) {
            $rowStyles = $options;
        }
        else {
            $rowStyles = array_replace_recursive($styles, $options);
        }
        $this->writeRow($this->cells['values'][$this->currentRow] ?? [], $rowStyles);

        return $this;
    }

    /**
     * @param array|mixed $row Values of cells
     * @param array|null $rowStyle
     *
     * @return $this
     */
    public function writeRow($row = [], array $rowStyle = null)
    {
        $writer = $this->excel->getWriter();
        $writer->writeSheetDataBegin($this);

        if (!is_array($row)) {
            $row = [$row];
        }

        $this->_writeRow($writer, $row, $rowStyle);
        $this->currentCol = Excel::MIN_COL;

        $this->currentRow++;

        return $this;
    }

    /**
     * @param int|null $rowCount
     *
     * @return $this
     */
    public function skipRow(?int $rowCount = 1)
    {
        for($i = 0; $i < $rowCount; $i++) {
            $this->writeRow([null]);
        }

        return $this;
    }

    /**
     * @param string $range A1:Z9 or R1C1:R9C28
     *
     * @return Area
     */
    public function makeArea(string $range)
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
     * Begin area
     *
     * @param string|null $cellAddress Upper left cell of area
     *
     * @return Area
     */
    public function beginArea(string $cellAddress = null)
    {
        if (null === $cellAddress) {
            $cellAddress = 'A' . ($this->rowCount + 1);
        }
        $dimension = Excel::rangeDimension($cellAddress, true);
        if ($dimension['rowNum1'] <= $this->rowCount) {
            throw new Exception("Cannot make area from $cellAddress (row number must be greater then written rows)");
        }
        $maxCell = Excel::cellAddress(Excel::MAX_ROW, Excel::MAX_COL);

        return $this->makeArea($cellAddress . ':' . $maxCell);
    }

    /**
     * @param $cellAddress
     *
     * @return array|bool
     */
    protected function _parseAddress($cellAddress)
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
     * @return array|bool
     */
    protected function _rangeDimension(string $cellAddress, ?int $colOffset = 1, ?int $rowOffset = 1)
    {
        if (preg_match('/^R\[?(-?\d+)?\]?C/', $cellAddress)) {
            // relative address
            $relAddress = $cellAddress;
            $cellAddress = Excel::colLetter($colOffset) . ($this->rowCount + $rowOffset);
            $dimension = Excel::rangeDimensionRelative($cellAddress, $relAddress, true);
        }
        else {
            // absolute address
            $dimension = Excel::rangeDimension($cellAddress, true);
        }

        return $dimension;
    }

    /**
     * @param string|array $cellAddress
     * @param mixed $values
     * @param mixed $styles
     * @param bool|null $merge
     * @param bool|null $changeCurrent
     *
     * @throws Exception
     */
    protected function _setCellData($cellAddress, $values, $styles, ?bool $merge, ?bool $changeCurrent = false)
    {
        $dimension = $this->_parseAddress($cellAddress);
        $row = $dimension['row'];
        $col = $dimension['col'];
        if ($merge && isset($dimension['width'], $dimension['height']) && ($dimension['width'] > 1 || $dimension['height'] > 1)) {
            $this->mergeCells($dimension['range']);
        }

        if ($row === null || $col === null) {
            ExceptionAddress::throwNew('Wrong cell address %s', print_r($cellAddress));
        }
        if ($row < $this->currentRow) {
            ExceptionAddress::throwNew('Row number must be greater then written rows');
        }

        $rowIdx = $row - 1;
        $colIdx = $col - 1;
        if ($values !== null) {
            $this->cells['values'][$rowIdx][$colIdx] = $values;
            if ($changeCurrent) {
                $this->currentRow = $rowIdx;
                $this->currentCol = $colIdx;
            }
        }
        if ($styles !== null) {
            $this->cells['styles'][$rowIdx][$colIdx] = $styles;
        }
    }

    /**
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
    public function setValue($cellAddress, $value, ?array $styles = null)
    {
        $styles = $styles ? Style::normalize($styles) : null;
        $this->_setCellData($cellAddress, $value, $styles, true);

        return $this;
    }

    /**
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
    public function setFormula($cellAddress, $value, array $styles = null)
    {
        if (empty($value)) {
            $value = null;
        }
        elseif (strpos($value, '=') !== 0) {
            $value = '=' . $value;
        }

        $styles = $styles ? Style::normalize($styles) : null;
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
    public function applayStyle(string $cellAddress, $style, ?bool $mergeStyles = false)
    {
        $dimension = $this->_rangeDimension($cellAddress);
        if ($dimension['rowNum1'] <= $this->rowCount) {
            throw new Exception('Row number must be greater then written rows');
        }
        $style = Style::normalize($style);
        for ($row = $dimension['rowNum1'] - 1; $row < $dimension['rowNum2']; $row++) {
            for ($col = $dimension['colNum1'] - 1; $col < $dimension['colNum2']; $col++) {
                if ($mergeStyles && isset($this->cells['styles'][$row][$col])) {
                    $this->cells['styles'][$row][$col] = array_merge($this->cells['styles'][$row][$col], $style);
                }
                else {
                    $this->cells['styles'][$row][$col] = $style;
                }
            }
        }

        return $this;
    }

    /**
     * @param string $cellAddr
     * @param $style
     *
     * @return $this
     */
    public function setStyle($cellAddr, $style)
    {
        return $this->applayStyle($cellAddr, $style, false);
    }

    /**
     * @param string $cellAddr
     * @param $style
     *
     * @return $this
     */
    public function addStyle($cellAddr, $style)
    {
        return $this->applayStyle($cellAddr, $style, true);
    }

    /**
     * @param string $cellAddr
     * @param $format
     *
     * @return $this
     */
    public function setFormat($cellAddr, $format)
    {
        return $this->applayStyle($cellAddr, ['format' => $format], true);
    }

    /**
     * @param $range
     * @param $style
     *
     * @return $this
     */
    public function setOuterBorder($range, $style)
    {
        $dimension = $this->_rangeDimension($range);
        if ($dimension['rowNum1'] <= $this->rowCount) {
            throw new Exception('Row number must be greater then written rows');
        }
        $border = Style::normalizeBorder($style);
        // top
        if (!empty($border['top'])) {
            $rowIdx = $dimension['row'] - 1;
            for ($colIdx = $dimension['colNum1'] - 1; $colIdx < $dimension['colNum2']; $colIdx++) {
                if (!empty($this->cells['styles'][$rowIdx][$colIdx]['border']['top'])) {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['top'] = array_merge($this->cells['styles'][$rowIdx][$colIdx]['border']['top'], $border['top']);
                } else {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['top'] = $border['top'];
                }
            }
        }

        // bottom
        if (!empty($border['bottom'])) {
            $rowIdx = $dimension['rowNum2'] - 1;
            for ($colIdx = $dimension['colNum1'] - 1; $colIdx < $dimension['colNum2']; $colIdx++) {
                if (!empty($this->cells['styles'][$rowIdx][$colIdx]['border']['bottom'])) {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['bottom'] = array_merge($this->cells['styles'][$rowIdx][$colIdx]['border']['top'], $border['bottom']);
                } else {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['bottom'] = $border['bottom'];
                }
            }
        }

        // left
        if (!empty($border['left'])) {
            $colIdx = $dimension['colNum1'] - 1;
            for ($rowIdx = $dimension['rowNum1'] - 1; $rowIdx < $dimension['rowNum2']; $rowIdx++) {
                if (!empty($this->cells['styles'][$rowIdx][$colIdx]['border']['left'])) {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['left'] = array_merge($this->cells['styles'][$rowIdx][$colIdx]['border']['left'], $border['left']);
                } else {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['left'] = $border['left'];
                }
            }
        }

        // right
        if (!empty($border['right'])) {
            $colIdx = $dimension['colNum2'] - 1;
            for ($rowIdx = $dimension['rowNum1'] - 1; $rowIdx < $dimension['rowNum2']; $rowIdx++) {
                if (!empty($this->cells['styles'][$rowIdx][$colIdx]['border']['right'])) {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['right'] = array_merge($this->cells['styles'][$rowIdx][$colIdx]['border']['right'], $border['right']);
                } else {
                    $this->cells['styles'][$rowIdx][$colIdx]['border']['right'] = $border['right'];
                }
            }
        }

        return $this;
    }

    public function writeAreas()
    {
        $writer = $this->excel->getWriter();
        if ($this->open) {
            $this->writeAreasRows($writer);
        } else {
            $writer->writeSheetDataBegin($this);
        }
        return $this;
    }

    /**
     * @return $this
     */
    public function clearAreas()
    {
        $this->cells = [];
        $this->areas = [];
        $this->totalArea = [];

        return $this;
    }

    /**
     * Write all active areas
     *
     * @return $this
     */
    public function writeAreasRows($writer)
    {
        if (!$this->areas) {
            return $this;
        }

        if (!empty($this->cells['values']) || !empty($this->cells['styles'])) {
            $maxRow = max(array_keys($this->cells['values']) + array_keys($this->cells['styles']));
            // writes row by row
            for ($numRow = $this->rowCount; $numRow <= $maxRow; $numRow++) {
                if (isset($this->cells['values'][$numRow]) || isset($this->cells['styles'][$numRow])) {
                    $colMax = 0;
                    $rowValues = $this->cells['values'][$numRow] ?? [];
                    if ($rowValues && ($keyMax = max(array_keys($rowValues))) > $colMax) {
                        $colMax = $keyMax;
                    }
                    $rowStyles = $this->cells['styles'][$numRow] ?? [];
                    if ($rowStyles && ($keyMax = max(array_keys($rowStyles))) > $colMax) {
                        $colMax = $keyMax;
                    }

                    for ($numCol = Excel::MIN_COL; $numCol <= $colMax; $numCol++) {
                        if (!isset($rowValues[$numCol])) {
                            $rowValues[$numCol] = null;
                        }
                        if (!isset($rowStyles[$numCol])) {
                            $rowStyles[$numCol] = [];
                        }
                    }
                    // array of all values
                    ksort($rowValues);
                    // array of all styles
                    ksort($rowStyles);
                    $this->_writeRow($writer, array_values($rowValues), null, array_values($rowStyles));
                }
                else {
                    $this->_writeRow($writer, [null]);
                }
            }
            $this->clearAreas();
        }

        return $this;
    }


    public function writeDataBegin($writer)
    {
        //if already initialized
        if ($this->open) {
            return;
        }

        $sheetFileName = $writer->tempFilename();
        $this->setFileWriter($writer::makeWriteBuffer($sheetFileName));

        $this->fileWriter->write('<sheetData>');

        $this->open = true;
        if ($this->areas) {
            $this->writeAreasRows($writer);
        }
    }


    public function writeDataEnd()
    {
        if ($this->close) {
            return;
        }
        if ($this->areas) {
            $this->writeAreas();
        }
        $this->nextRow();
        $this->fileWriter->flush(true);
        $this->fileWriter->write('</sheetData>');
    }

    /**
     * @return string
     */
    public function maxCell()
    {
        return Excel::cellAddress($this->rowCount, $this->colCount);
    }

    /**
     * @return bool
     */
    public function isRightToLeft()
    {
        return $this->excel->isRightToLeft();
    }
}

// EOF