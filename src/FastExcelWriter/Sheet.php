<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;

/**
 * Class Sheet
 *
 * @package avadim\FastExcelWriter
 */
class Sheet
{
    /** @var Excel */
    public $book;

    /** @var string */
    public $key;

    /** @var int */
    public $index;

    public $active = false;
    public $fileName  = '';
    public $sheetName = '';
    public $xmlName   = '';
    public $rowCount  = 0;
    public $colCount  = 0;

    /**
     * @var WriterBuffer
     */
    public $fileWriter      = null;

    public $columns         = [];
    public $freezeRows      = 0;
    public $freezeColumns   = 0;
    public $autoFilter      = 0;
    public $absoluteAutoFilter = '';

    // zero based
    public $colWidths       = [];
    public $colFormats      = [];
    public $colStyles       = [];

    public $rowHeights = [];

    public $open = false;
    public $close = false;

    protected $mergeCells = [];
    protected $cells = [];
    protected $totalArea = [];
    protected $areas = [];
    protected $defaultStyle = [];

    protected $currentRow = 0;
    protected $currentCol = 0;

    protected $pageSetup = [];

    /**
     * Sheet constructor
     *
     * @param string $sheetName
     */
    public function __construct($sheetName)
    {
        $this->setName($sheetName);
        $this->pageSetup['orientation'] = 'portrait';
        $this->cells = [
            'values' => [],
            'styles' => [],
        ];
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
    public function setName($sheetName)
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
     * @param $option
     * @param $value
     *
     * @return $this
     */
    public function setPageSetup($option, $value)
    {
        if ($this->rowCount) {
            throw new Exception('Cannot set page settings after rows writing');
        }
        $this->pageSetup[$option] = $value;

        return $this;
    }

    /**
     * @param $option
     *
     * @return mixed|null
     */
    public function getPageSetup($option)
    {
        return $this->pageSetup[$option] ?? null;
    }

    /**
     * @return $this
     */
    public function pageOrientationPortrait()
    {
        return $this->setPageSetup('orientation', 'portrait');
    }

    /**
     * @return $this
     */
    public function pageOrientationLandscape()
    {
        return $this->setPageSetup('orientation', 'landscape');
    }

    /**
     * @param int $numPage
     *
     * @return $this
     */
    public function pageFitToWidth($numPage = 1)
    {
        if ($numPage === 'auto') {
            $numPage = 0;
        }
        if ($numPage >=0) {
            $this->setPageSetup('fit_width', (int)$numPage);
        }
        return $this;
    }

    /**
     * @param int $numPage
     *
     * @return $this
     */
    public function pageFitToHeight($numPage = 1)
    {
        if ($numPage === 'auto') {
            $numPage = 0;
        }
        if ($numPage >=0) {
            $this->setPageSetup('fit_height', (int)$numPage);
        }
        return $this;
    }

    /**
     * @return string
     */
    public function getPageOrientation()
    {
        return $this->pageSetup['orientation'] ?? 'portrait';
    }

    /**
     * @return int
     */
    public function getPageFitToWidth()
    {
        return (int)($this->pageSetup['fit_width'] ?? 0);
    }

    /**
     * @return int
     */
    public function getPageFitToHeight()
    {
        return (int)($this->pageSetup['fit_height'] ?? 0);
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
     * @param $freezeRows
     * @param $freezeColumns
     *
     * @return $this
     */
    public function setFreeze($freezeRows, $freezeColumns = null)
    {
        if (!is_numeric($freezeRows) && null === $freezeColumns) {
            $dimension = Excel::rangeDimension($freezeRows);
            if ($dimension) {
                $this->setFreezeRows($dimension['rowIndex1'] - 1);
                $this->setFreezeColumns($dimension['colIndex1'] - 1);
            }
        } else {
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
    public function setFreezeRows($freezeRows)
    {
        $this->freezeRows = ($freezeRows > 0) ? $freezeRows : 0;

        return $this;
    }

    /**
     * @param int $freezeColumns Number columns to freeze
     *
     * @return $this
     */
    public function setFreezeColumns($freezeColumns)
    {
        $this->freezeColumns = ($freezeColumns > 0) ? $freezeColumns : 0;

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     *
     * @return $this
     */
    public function setAutofilter($row = 1, $col = 1)
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
     * @param int|string $col Column index or column letter
     * @param string $optionName
     * @param mixed $optionValue
     */
    protected function _setColOption($col, $optionName, $optionValue)
    {
        if ($optionValue !== null) {
            if (is_numeric($col)) {
                $colIndex = (int)$col;
            } else {
                $colIndex = Excel::colIndex($col);
            }
            if ($optionName === 'width' && is_numeric($optionValue)) {
                $this->colWidths[$colIndex - 1] = str_replace(',', '.', (float)$optionValue);
            } elseif ($optionName === 'format') {
                $this->colFormats[$colIndex - 1] = $optionValue;
                if (isset($this->columns[$colIndex])) {
                    unset($this->columns[$colIndex]);
                }
            } elseif ($optionName === 'style') {
                if (!empty($this->defaultStyle)) {
                    $optionValue = $optionValue ? array_merge($optionValue, $this->defaultStyle) : $this->defaultStyle;
                }
                $this->colStyles[$colIndex - 1] = $optionValue ? Style::normalize($optionValue) : $optionValue;
                if (isset($this->columns[$colIndex])) {
                    unset($this->columns[$colIndex]);
                }
            }
        }
    }

    /**
     * @param string $optionName
     * @param array $colOptions
     */
    protected function _setColOptions($optionName, $colOptions)
    {
        if (is_array($colOptions)) {
            foreach($colOptions as $col => $value) {
                $this->_setColOption(is_int($col) ? ++$col : $col, $optionName, $value);
            }
        }
    }

    /**
     * Set width of single column
     *
     * @param int|string|array $col Column index or column letter
     * @param int|float $width
     *
     * @return $this
     */
    public function setColWidth($col, $width)
    {
        foreach((array)$col as $colName) {
            $this->_setColOption($colName, 'width', $width);
        }

        return $this;
    }

    /**
     * Set format of single column
     *
     * @param int|string|array $col Column index or column letter
     * @param string $format
     *
     * @return $this
     */
    public function setColFormat($col, $format)
    {
        foreach((array)$col as $colName) {
            $this->_setColOption($colName, 'format', $format);
        }

        return $this;
    }

    /**
     * Set style of single column
     *
     * @param int|string $col Column index or column letter
     * @param mixed $style
     *
     * @return $this
     */
    public function setColStyle($col, $style)
    {
        foreach((array)$col as $colName) {
            $this->_setColOption($colName, 'style', $style);
        }

        return $this;
    }

    /**
     * Set widths of columns
     *
     * @param array $colWidths
     *
     * @return $this
     */
    public function setColWidths(array $colWidths)
    {
        $this->_setColOptions('width', $colWidths);

        return $this;
    }

    /**
     * Set formats of columns
     *
     * @param array $colFormats
     *
     * @return $this
     */
    public function setColFormats(array $colFormats)
    {
        $this->_setColOptions('format', $colFormats);

        return $this;
    }

    /**
     * Set styles of columns
     *
     * @param array $colFormats
     *
     * @return $this
     */
    public function setColStyles(array $colFormats)
    {
        $this->_setColOptions('style', $colFormats);

        return $this;
    }

    /**
     * Set options of columns (widths and/or formats)
     *  $sheet->setColumns([
     *      ['format' => 'integer'],
     *      ['format' => 'date', 'width' => 16],
     *      ['format' => 'text', 'width' => 40, 'text-wrap' => true],
     *  ]);
     *
     * @param array $options
     *
     * @return $this
     */
    public function setColumns($options)
    {
        if (is_array($options)) {
            if (key($options) === 0) {
                $colIndex = 1;
            } else {
                $colIndex = -1;
            }
            foreach($options as $colName => $colOptions) {
                $col = ($colIndex > 0) ? $colIndex++ : $colName;
                $style = [];
                foreach($colOptions as $optionName => $optionValue) {
                    if ($optionName === 'width') {
                        $this->setColWidth($col, $colOptions['width']);
                    } elseif ($optionName === 'format') {
                        $this->setColFormat($col, $colOptions['format']);
                    } elseif ($optionName === 'style') {
                        $style = $colOptions['style'];
                    } else {
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

    public function setRowHeights($heights)
    {
        foreach ($heights as $rowNum => $rowHeight) {
            $this->setRowHeight($rowNum, $rowHeight);
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
            if ($this->colStyles) {
                foreach($this->colStyles as $colIndex => $colStyle) {
                    $colStyle = !empty($colStyle) ? array_merge($colStyle, $style) : $style;
                    $this->colStyles[$colIndex] = $colStyle ? Style::normalize($colStyle) : $colStyle;
                    if (isset($this->columns[$colIndex])) {
                        unset($this->columns[$colIndex]);
                    }
                }
            }
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
     * @param array $row
     * @param array|null $rowOptions
     * @param array|null $cellsOptions
     */
    protected function _writeRow($writer, array $row = [], $rowOptions = null, $cellsOptions = null)
    {
        if (count($this->colFormats) > count($this->columns)) {
            foreach($this->colFormats as $colNum => $format) {
                $colIndex = $colNum + 1;
                if (!isset($this->columns[$colIndex])) {
                    $this->columns[$colIndex] = $writer->defineFormatType($format);
                }
            }
        }

        $defaultFormatType = $writer->defaultFormatType();

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

        $rowFormat = $rowStyle = $rowCellStyles = null;
        // styles for each cell
        $rowCellStyles = $cellsOptions;

        $colNum = 0;
        foreach ($row as $cellValue) {
            $formatType = $this->columns[$colNum + 1] ?? $defaultFormatType;
            $numberFormat = $formatType['number_format'];
            $numberFormatType = $formatType['number_format_type'];
            if (empty($rowCellStyles) && empty($rowStyle)) {
                $cellStyleIdx = $formatType['default_style_idx'];
            } else {
                if (isset($rowCellStyles[$colNum])) {
                    $cellStyle = $rowCellStyles[$colNum];
                    if (isset($cellStyle['format'])) {
                        $cellFormat = $writer->defineFormatType($cellStyle['format']);
                        $numberFormat = $cellFormat['number_format'];
                        $numberFormatType = $cellFormat['number_format_type'];
                    }
                } else {
                    $cellStyle = $rowStyle;
                    if ($rowFormat) {
                        $numberFormat = $rowFormat['number_format'];
                        $numberFormatType = $rowFormat['number_format_type'];
                    }
                }
                $cellStyleIdx = $writer->addCellStyle($numberFormat, $cellStyle);
            }
            if (!empty($cellStyle['autofit'])) {
                $len = max(mb_strlen((string)$cellValue) ?: 1, mb_strlen(str_replace('\\', '', $numberFormat)));
                if (empty($this->colWidths[$colNum]) || $this->colWidths[$colNum] < $len) {
                    $this->colWidths[$colNum] = $len * 1.2;
                }
            }
            $writer->writeCell($this->fileWriter, $this->rowCount + 1, $colNum + 1, $cellValue, $numberFormatType, $cellStyleIdx);
            $colNum++;
            if ($colNum > $this->colCount) {
                $this->colCount = $colNum;
            }
        }
        $this->fileWriter->write('</row>');
        $this->rowCount++;
    }

    /**
     * @param int $colNumber
     * @param array $options
     * @param array|null $defaultOptions
     * @param int $keyOffset
     *
     * @return array
     */
    protected function _parseColumnOptions($colNumber, $options, $defaultOptions = null, $keyOffset = 1)
    {
        $colOptions = [];
        foreach($options as $key => $val) {
            if (is_int($key)) {
                $colIndex = $key + $keyOffset;
            } else {
                $colIndex = Excel::colIndex($key) + $keyOffset - 1;
            }
            if ($colIndex > 0) {
                if ($defaultOptions) {
                    $colOptions[$colIndex] = array_merge($defaultOptions, $val);
                } else {
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
     * @param mixed $value
     * @param array|null $options
     *
     * @return $this
     *
     * @throws \Exception
     */
    public function writeCell($value, $options = null)
    {
        $data = [
            'values' => $value,
            'styles' => Style::normalize($options),
        ];
        if ($this->currentRow < $this->rowCount) {
            $this->currentRow = $this->rowCount;
        }
        $cellAddr = [
            'row' => $this->currentRow,
            'col' => $this->currentCol++,
        ];
        $this->_setCellData($cellAddr, $data, false, true);

        return $this;
    }

    /**
     * @return $this
     */
    public function nextCell()
    {
        $this->writeCell(null);

        return $this;
    }

    /**
     * @param array $options
     *
     * @return $this
     */
    public function nextRow($options = [])
    {
        $styles = $this->cells['styles'][$this->currentRow] ?? [];
        $styles = array_merge_recursive($styles, $options);
        $this->writeRow($this->cells['values'][$this->currentRow] ?? [], $styles);
        $this->currentCol = 0;

        return $this;
    }

    /**
     * @param array|mixed $row
     * @param array|null $options
     *
     * @return $this
     */
    public function writeRow($row = [], array $options = null)
    {
        $writer = $this->book->getWriter();
        $writer->writeSheetDataBegin($this);

        if (!is_array($row)) {
            $row = [$row];
        }

        // ++++++++++++++++++++++++++++
        // splits row and cells options
        $rowOptions = $rowCellsOptions = [];
        if ($options) {
            if (isset($options['height'])) {
                $rowOptions['height'] = (float)$options['height'];
                unset($options['height']);
            }
            if (isset($options['hidden'])) {
                $rowOptions['hidden'] = (bool)$options['hidden'];
                unset($options['height']);
            }
            if (isset($options['collapsed'])) {
                $rowOptions['collapsed'] = (bool)$options['collapsed'];
                unset($options['collapsed']);
            }
            $rowCellsOptions = $options;
        }
        // ----------------------------

        $cellsCount = count($row);

        // combine default options and column options to $colOptions
        $defaultStyle = $this->getDefaultStyle() ?: null;
        $colOptions = [];

        if ($defaultStyle) {
            if (!empty($this->colStyles)) {
                for($colNum = 0; $colNum < $cellsCount; $colNum++) {
                    if (!isset($this->colStyles[$colNum])) {
                        $colOptions[$colNum] = $defaultStyle;
                    } else {
                        $colOptions[$colNum] = $this->colStyles[$colNum];
                    }
                }
            } else {
                $colOptions = array_fill(0, count($row), $defaultStyle);
            }
        } else {
            $colOptions = $this->colStyles;
        }

        // combine column options and cell options for current row
        if (!empty($rowCellsOptions) && is_array($rowCellsOptions)) {
            $firstKey = key($rowCellsOptions);
            if (is_int($firstKey) || ($firstKey && is_string($firstKey) && $firstKey[0] >= 'A' && $firstKey[0] <= 'Z')) {
                $rowCellsOptions = $this->_parseColumnOptions(count($row), $rowCellsOptions);
            } elseif ($rowCellsOptions) {
                $rowCellsOptions = array_fill(0, count($row), $rowCellsOptions);
            }
            if ($colOptions) {
                foreach ($rowCellsOptions as $colNum => $cellOptions) {
                    if (!empty($colOptions[$colNum])) {
                        if (!empty($cellOptions)) {
                            $rowCellsOptions[$colNum] = array_merge($cellOptions, $colOptions[$colNum]);
                        } else {
                            $rowCellsOptions[$colNum] = $colOptions[$colNum];
                        }
                    }
                }
            }
        } else {
            $rowCellsOptions = $colOptions;
        }

        $this->_writeRow($writer, $row, $rowOptions, $rowCellsOptions);

        return $this;
    }

    /**
     * @param int $rowCount
     *
     * @return $this
     */
    public function skipRow($rowCount = 1)
    {
        for($i = 0; $i < $rowCount; $i++) {
            $this->writeRow([null]);
        }

        return $this;
    }

    /**
     * @param array $header
     * @param array|null $options
     *
     * @return $this
     */
    public function writeHeader($header, $options = null)
    {
        $rowValues = [];
        $colFormats = [];
        foreach($header as $key => $val) {
            if (!is_int($key)) {
                $rowValues[] = $key;
                $colFormats[] = $val;
            } else {
                $rowValues[] = $val;
                $colFormats[] = null;
            }
        }
        $this->writeRow($rowValues, $options);
        $this->setColFormats($colFormats);

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
                    ((($dimension['rowIndex1'] >= $savedDimension['rowIndex1']) && ($dimension['rowIndex1'] <= $savedDimension['rowIndex2']))
                        || (($dimension['rowIndex2'] >= $savedDimension['rowIndex1']) && ($dimension['rowIndex2'] <= $savedDimension['rowIndex2'])))
                    && ((($dimension['colIndex1'] >= $savedDimension['colIndex1']) && ($dimension['colIndex1'] <= $savedDimension['colIndex2']))
                        || (($dimension['colIndex2'] >= $savedDimension['colIndex1']) && ($dimension['colIndex2'] <= $savedDimension['colIndex2'])))
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
                    ((($dimension['rowIndex1'] >= $savedDimension['rowIndex1']) && ($dimension['rowIndex1'] <= $savedDimension['rowIndex2']))
                        || (($dimension['rowIndex2'] >= $savedDimension['rowIndex1']) && ($dimension['rowIndex2'] <= $savedDimension['rowIndex2'])))
                    && ((($dimension['colIndex1'] >= $savedDimension['colIndex1']) && ($dimension['colIndex1'] <= $savedDimension['colIndex2']))
                        || (($dimension['colIndex2'] >= $savedDimension['colIndex1']) && ($dimension['colIndex2'] <= $savedDimension['colIndex2'])))
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
     * @param string $range A1:Z9 or R1C1:R9C28
     *
     * @return Area
     */
    public function makeArea($range)
    {
        $area = new Area($this, $range);

        $this->areas[] = $area->setIndex(count($this->areas));
        $coord = $area->getCoord();
        if (empty($this->totalArea['coord'])) {
            $this->totalArea['coord'] = $coord;
        } else {
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
     * @param string $cellAddr Upper left cell of area
     *
     * @return Area
     */
    public function beginArea($cellAddr = null)
    {
        if (null === $cellAddr) {
            $cellAddr = 'A' . ($this->rowCount + 1);
        }
        $dimension = Excel::rangeDimension($cellAddr, true);
        if ($dimension['rowIndex1'] <= $this->rowCount) {
            throw new Exception("Cannot make area from $cellAddr (row number must be greater then written rows)");
        }
        $maxCell = Excel::cellAddress(Excel::EXCEL_2007_MAX_ROW, Excel::EXCEL_2007_MAX_COL);

        return $this->makeArea($cellAddr . ':' . $maxCell);
    }

    /**
     * @param $cellAddr
     * @param $colOffset
     * @param $rowOffset
     *
     * @return array|bool
     */
    protected function _rangeDimension($cellAddr, $colOffset = 1, $rowOffset = 1)
    {
        if (preg_match('/^R\[?(-?\d+)?\]?C/', $cellAddr)) {
            $relAddr = $cellAddr;
            $cellAddr = Excel::colLetter($colOffset) . ($this->rowCount + $rowOffset);
            $dimension = Excel::rangeDimensionRelative($cellAddr, $relAddr, true);
        } else {
            $dimension = Excel::rangeDimension($cellAddr, true);
        }
        return $dimension;
    }

    /**
     * @param string|array $cellAddr
     * @param $data
     * @param $merge
     * @param $currentRow
     *
     * @throws \Exception
     */
    protected function _setCellData($cellAddr, $data, $merge, $currentRow = false)
    {
        $row = $col = null;
        if (is_string($cellAddr)) {
            $dimension = $this->_rangeDimension($cellAddr);
            $row = $dimension['rowIndex1'];
            $col = $dimension['colIndex1'];
            if ($merge && ($dimension['width'] > 1 || $dimension['height'] > 1)) {
                $this->mergeCells($dimension['range']);
            }
        } elseif (is_array($cellAddr)) {
            if (isset($cellAddr['row'], $cellAddr['col'])) {
                $row = $cellAddr['row'];
                $col = $cellAddr['col'];
            } else {
                [$row, $col] = $cellAddr;
            }
        }
        if ($row === null || $col === null) {
            throw new Exception('Wrong cell address ' . print_r($cellAddr));
        }
        if ($row < $this->rowCount + ($currentRow ? 0 : 1)) {
            throw new \Exception('Row number must be greater then written rows');
        }

        foreach ($data as $key => $val) {
            $this->cells[$key][$row][$col] = $val;
        }
    }

    /**
     * @param string|array $cellAddr
     * @param $value
     * @param $style
     *
     * @return $this
     */
    public function setValue($cellAddr, $value, $style = [])
    {
        $data = [
            'values' => $value,
            'styles' => Style::normalize($style),
        ];
        $this->_setCellData($cellAddr, $data, true);

        return $this;
    }

    /**
     * @param $cellAddr
     * @param $value
     * @param $style
     *
     * @return $this
     */
    public function setFormula($cellAddr, $value, $style)
    {
        if (empty($value)) {
            $value = null;
        } elseif (strpos($value, '=') !== 0) {
            $value = '=' . $value;
        }
        $data = [
            'values' => $value,
            'styles' => Style::normalize($style),
        ];
        $this->_setCellData($cellAddr, $data, true);

        return $this;
    }

    /**
     * @param string $cellAddr
     * @param $style
     * @param $mergeStyles
     *
     * @return $this
     */
    public function applayStyle($cellAddr, $style, $mergeStyles = false)
    {
        $dimension = $this->_rangeDimension($cellAddr);
        if ($dimension['rowIndex1'] <= $this->rowCount) {
            throw new Exception('Row number must be greater then written rows');
        }
        $style = Style::normalize($style);
        for($row = $dimension['rowIndex1']; $row <= $dimension['rowIndex2']; $row++) {
            for ($col = $dimension['colIndex1']; $col <= $dimension['colIndex2']; $col++) {
                if ($mergeStyles && isset($this->cells['styles'][$row][$col])) {
                    $this->cells['styles'][$row][$col] = array_merge($this->cells['styles'][$row][$col], $style);
                } else {
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
        if ($dimension['rowIndex1'] <= $this->rowCount) {
            throw new Exception('Row number must be greater then written rows');
        }
        $border = Style::normalizeBorder($style);
        // top
        if (!empty($border['top'])) {
            $row = $dimension['rowIndex1'];
            for ($col = $dimension['colIndex1']; $col <= $dimension['colIndex2']; $col++) {
                if (!empty($this->cells['styles'][$row][$col]['border']['top'])) {
                    $this->cells['styles'][$row][$col]['border']['top'] = array_merge($this->cells['styles'][$row][$col]['border']['top'], $border['top']);
                } else {
                    $this->cells['styles'][$row][$col]['border']['top'] = $border['top'];
                }
            }
        }
        // bottom
        if (!empty($border['bottom'])) {
            $row = $dimension['rowIndex2'];
            for ($col = $dimension['colIndex1']; $col <= $dimension['colIndex2']; $col++) {
                if (!empty($this->cells['styles'][$row][$col]['border']['bottom'])) {
                    $this->cells['styles'][$row][$col]['border']['bottom'] = array_merge($this->cells['styles'][$row][$col]['border']['top'], $border['bottom']);
                } else {
                    $this->cells['styles'][$row][$col]['border']['bottom'] = $border['bottom'];
                }
            }
        }
        // left
        if (!empty($border['left'])) {
            $col = $dimension['colIndex1'];
            for ($row = $dimension['rowIndex1']; $row <= $dimension['rowIndex2']; $row++) {
                if (!empty($this->cells['styles'][$row][$col]['border']['left'])) {
                    $this->cells['styles'][$row][$col]['border']['left'] = array_merge($this->cells['styles'][$row][$col]['border']['left'], $border['left']);
                } else {
                    $this->cells['styles'][$row][$col]['border']['left'] = $border['left'];
                }
            }
        }
        // right
        if (!empty($border['right'])) {
            $col = $dimension['colIndex2'];
            for ($row = $dimension['rowIndex1']; $row <= $dimension['rowIndex2']; $row++) {
                if (!empty($this->cells['styles'][$row][$col]['border']['right'])) {
                    $this->cells['styles'][$row][$col]['border']['right'] = array_merge($this->cells['styles'][$row][$col]['border']['right'], $border['right']);
                } else {
                    $this->cells['styles'][$row][$col]['border']['right'] = $border['right'];
                }
            }
        }

        return $this;
    }

    public function writeAreas()
    {
        $writer = $this->book->getWriter();
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
        if (!empty($this->cells['values']) || !empty($this->cells['styles'])) {
            $maxRow = max(array_keys($this->cells['values']) + array_keys($this->cells['styles']));
            for ($numRow = $this->rowCount + 1; $numRow <= $maxRow; $numRow++) {
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

                    for ($numCol = 1; $numCol <= $colMax; $numCol++) {
                        if (!isset($rowValues[$numCol])) {
                            $rowValues[$numCol] = null;
                        }
                        if (!isset($rowStyles[$numCol])) {
                            $rowStyles[$numCol] = [];
                        }
                    }
                    ksort($rowValues);
                    ksort($rowStyles);
                    $this->_writeRow($writer, array_values($rowValues), null, array_values($rowStyles));
                } else {
                    $this->_writeRow($writer, [null]);
                }
            }
            $this->clearAreas();
        }

        return $this;
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
        return $this->book->isRightToLeft();
    }
}

// EOF