<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelWriter\Charts\Chart;
use avadim\FastExcelWriter\Conditional\Conditional;
use avadim\FastExcelWriter\DataValidation\DataValidation;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Exceptions\ExceptionAddress;
use avadim\FastExcelWriter\Exceptions\ExceptionRangeName;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;
use avadim\FastExcelWriter\Writer\FileWriter;
use avadim\FastExcelWriter\Writer\Writer;

/**
 * Class Sheet
 *
 * @package avadim\FastExcelWriter
 */
class Sheet implements InterfaceSheetWriter
{
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

    /** @var int Local ID of the sheet */
    public int $localSheetId;

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

    public ?FileWriter $fileWriter = null;

    public array $defaultStyle = [];

    protected array $sheetStylesSummary = [];

    // ZERO based
    public int $freezeRows = 0;
    public int $freezeColumns = 0;

    public ?string $autoFilter = null;
    public string $absoluteAutoFilter = '';

    // ++ ZERO based
    public array $colFormulas = [];
    public array $colStyles = [];
    protected array $colAttributes = [];
    // --

    // minimal with of columns
    protected array $colMinWidths = [];
    protected array $colStylesSummary = [];

    // ++ ZERO based
    protected array $rowSettings = [];
    public array $rowStyles = [];
    protected array $rowAttributes = [];
    // --

    // ZERO based
    protected array $cells = [];

    // Current row index
    protected int $currentRowIdx = 0;

    // Current column index
    protected int $currentColIdx = 0;

    protected int $rowOutlineLevel = 0;
    protected array $rowOutlineCollapsed = [];

    protected int $offsetCol = 0;

    protected array $mergeCells = [];
    protected array $mergedCellsArray = ['rowNum1' => Excel::MAX_ROW, 'colNum1' => Excel::MAX_COL, 'rowNum2' => 0, 'colNum2' => 0];

    protected array $totalArea = [];
    protected array $areas = [];

    protected int $relationshipId = 0;

    protected array $relationships = [];

    protected array $lastTouch = [];
    protected int $minRow = 0;
    protected int $minCol = 0;
    protected int $maxRow = 0;
    protected int $maxCol = 0;

    protected array $namedRanges = [];

    protected array $notes = [];

    protected array $media = [];

    protected array $charts = [];

    protected int $drawingRelsId = 0;

    // Data validations
    protected array $validations = [];

    // Conditional formatting
    protected array $conditionals = [];

    protected array $protection = [];

    protected ?string $activeCell = null;
    protected ?string $activeRef = null;

    protected array $sheetViews = [];

    protected array $sheetFormatPr = [];

    protected array $sheetProperties = [];

    // bottom sheet nodes
    protected array $bottomNodesOptions = [];

    protected array $printAreas = [];
    protected string $printTopRows = '';
    protected string $printLeftColumns = '';


    /**
     * Sheet constructor
     *
     * @param string $sheetName
     */
    public function __construct(string $sheetName)
    {
        $this->setName($sheetName);
        $this->bottomNodesOptions = [
            'printOptions' => [],
            'pageMargins' => [
                'left' => '0.5',
                'right' => '0.5',
                'top' => '1.0',
                'bottom' => '1.0',
                'header' => '0.5',
                'footer' => '0.5',
            ],
            'pageSetup' => [
                //'paperSize' => '1',
                'useFirstPageNumber' => '1',
                'horizontalDpi' => '0',
                'verticalDpi' => '0',
                'orientation' => 'portrait',
            ],
            /*
            'headerFooter' => [
                'differentFirst' => false,
                'differentOddEven' => false,
                '__kids' => [
                    [
                        '__name' => 'oddHeader',
                        '__attr' => [],
                    ],
                    [
                        '__name' => 'oddFooter',
                        '__attr' => [],
                    ],
                ],
            ],
            */
        ];

        $this->sheetViews = [
            [
                'workbookViewId' => '0',
                'view' => 'normal',
                'topLeftCell' => 'A1',
            ]
        ];

        $this->cells = [
            'values' => [],
            'styles' => [],
        ];
        $this->_setCellData('A1', null, null, false);
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
     * @return object|null
     */
    public function __get($name)
    {
        if ($name === 'book') {
            return $this->excel;
        }
        $trace = debug_backtrace();

        trigger_error(
            'Undefined property: ' . __CLASS__ . '::$' . $name .
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
     * @param FileWriter $fileWriter
     *
     * @return $this
     */
    public function setFileWriter(FileWriter $fileWriter): Sheet
    {
        if (!$this->fileWriter) {
            $this->fileWriter = $fileWriter;
            $this->fileTempName = $fileWriter->getFileName();
            $this->fileRels = $this->fileTempName . '.rels';
        }

        return $this;
    }

    /**
     * @param FileWriter $fileWriter
     *
     * @return $this
     */
    public function resetFileWriter(FileWriter $fileWriter): Sheet
    {
        $this->fileWriter = $fileWriter;
        $this->fileTempName = $fileWriter->getFileName();
        $this->fileRels = $this->fileTempName . '.rels';

        return $this;
    }

    /**
    * @return int
    */
    public function getCurrentRowId(): int
    {
        return $this->currentRowIdx;
    }

    /**
    * @return int
    */
    public function getCurrentColId(): int
    {
        return $this->currentColIdx;
    }

    /**
     * Returns current row number
     *
     * @return int
     */
    public function getCurrentRow(): int
    {
        return $this->currentRowIdx + 1;
    }

    /**
     * Returns current column letter
     * *
     * @return string
     */
    public function getCurrentCol(): string
    {
        return Helper::colLetter($this->currentColIdx + 1);
    }

    /**
     * Returns address of the current cell
     *
     * @return string
     */
    public function getCurrentCell(): string
    {
        return $this->getCurrentCol() . $this->getCurrentRow();
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
     * @return string
     */
    public function incDrawingRelsId(): string
    {
        return 'rId' . (++$this->drawingRelsId);
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
     * Case-insensitive name checking
     *
     * @param string $name
     *
     * @return bool
     */
    public function isName(string $name): bool
    {
        return strcasecmp($this->sheetName, $name) === 0;
    }

    /**
     * @param string $node
     * @param string $key
     * @param mixed $value
     *
     * @return $this
     */
    public function setBottomNodeOption(string $node, string $key, $value): Sheet
    {
        $this->bottomNodesOptions[$node][$key] = $value;

        return $this;
    }

    /**
     * @param string $node
     * @param array $options
     *
     * @return $this
     */
    public function setBottomNodesOptions(string $node, array $options): Sheet
    {
        $this->bottomNodesOptions[$node] = $options;

        return $this;
    }

    /**
     * @deprecated
     *
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
        $this->bottomNodesOptions['pageSetup'][$option] = $value;

        return $this;
    }

    /**
     * Set page orientation as Portrait, alias of pagePortrait()
     *
     * @return $this
     */
    public function pageOrientationPortrait(): Sheet
    {
        $this->bottomNodesOptions['pageSetup']['orientation'] = 'portrait';

        return $this;
    }

    /**
     * Set page orientation as Portrait
     *
     * @return $this
     */
    public function pagePortrait(): Sheet
    {
        return $this->pageOrientationPortrait();
    }

    /**
     * Set page orientation as Landscape, alias of pageLandscape()
     * *
     * @return $this
     */
    public function pageOrientationLandscape(): Sheet
    {
        $this->bottomNodesOptions['pageSetup']['orientation'] = 'landscape';

        return $this;
    }

    /**
     * Set page orientation as Landscape
     *
     * @return $this
     */
    public function pageLandscape(): Sheet
    {
        return $this->pageOrientationLandscape();
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
            $this->bottomNodesOptions['pageSetup']['fitToWidth'] = (int)$numPage;
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
            $this->bottomNodesOptions['pageSetup']['fitToHeight'] = (int)$numPage;
        }
        return $this;
    }

    /**
     * @param int $scale
     *
     * @return $this
     */
    public function pageScale(int $scale): Sheet
    {
        $this->bottomNodesOptions['pageSetup']['scale'] = $scale;

        return $this;
    }

    /**
     * @return string
     */
    public function getPageOrientation(): string
    {
        return $this->bottomNodesOptions['pageSetup']['orientation'] ?? 'portrait';
    }

    /**
     * @return int
     */
    public function getPageFitToWidth(): int
    {
        return (int)($this->bottomNodesOptions['pageSetup']['fitToWidth'] ?? 0);
    }

    /**
     * @return int
     */
    public function getPageFitToHeight(): int
    {
        return (int)($this->bottomNodesOptions['pageSetup']['fitToHeight'] ?? 0);
    }

    /**
     * @return bool
     */
    public function getPageFit(): bool
    {
        return !empty($this->bottomNodesOptions['pageSetup']['fitToWidth']) || !empty($this->bottomNodesOptions['pageSetup']['fitToHeight']);
    }

    /**
     * @return array
     */
    public function getSheetProperties(): array
    {
        if ($this->getPageFit()) {
            $this->sheetProperties['pageSetUpPr'] = [
                '_tag' => 'pageSetUpPr',
                '_attr' => ['fitToPage' => '1'],
            ];
        }

        return $this->sheetProperties;
    }

    /**
     * @return array
     */
    public function getPageSetup(): array
    {
        return $this->bottomNodesOptions['pageSetup'];
    }

    /**
     * @param array $options
     *
     * @return $this
     */
    public function setPageSetup(array $options): Sheet
    {
        $this->bottomNodesOptions['pageSetup'] = $options;

        return $this;
    }

    /**
     * Freeze rows/columns
     *
     * @example
     * $sheet->setFreeze(3, 3); // number rows and columns to freeze
     * $sheet->setFreeze('C3'); // left top cell of free area
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
     * Freeze rows
     *
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
     * Freeze columns
     *
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
     * Set active cell
     *
     * @param $cellAddress
     *
     * @return $this
     */
    public function setActiveCell($cellAddress): Sheet
    {
        $address = $this->_parseAddress($cellAddress);
        if (!isset($address['row'], $address['col'])) {
            ExceptionAddress::throwNew('Wrong cell address %s', print_r($address, 1));
        }
        if ($address['cell1'] === $address['cell2']) {
            $this->activeRef = $this->activeCell = $address['cell1'];
        }
        else {
            $this->activeCell = $address['cell1'];
            $this->activeRef = $address['cell1'] . ':' . $address['cell2'];
        }

        return $this;
    }


    /**
     * Set auto filter
     *
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
     * Set top left cell for writing
     *
     * @param string|array $cellAddress
     *
     * @return $this
     *
     * @example
     * $sheet->setTopLeftCell('C3');
     * $sheet->writeRow([11, 22, 33]); // Will be written in cells C3, D3, E3
     * $sheet->setTopLeftCell('G7');
     * $sheet->writeRow([44, 55]); // Will be written in cells G7, H7
     */
    public function setTopLeftCell($cellAddress): Sheet
    {
        $address = $this->_moveTo($cellAddress);
        $this->_touch($address['rowIndex'], $address['colIndex'], $address['rowIndex'], $address['colIndex'], 'cell');

        $this->currentRowIdx = $address['rowIndex'];
        $this->currentColIdx = $this->offsetCol = $address['colIndex'];

        return $this;
    }

    /**
     * Set color for the sheet tab
     *
     * @param string|null $color
     *
     * @return $this
     */
    public function setTabColor(?string $color): Sheet
    {
        if (!$color) {
            if (isset($this->sheetProperties['tabColor'])) {
                unset($this->sheetProperties['tabColor']);
            }
        }
        else {
            $color = StyleManager::normalizeColor($color);
            $this->sheetProperties['tabColor'] = [
                '_tag' => 'tabColor',
                '_attr' => ['rgb' => $color],
            ];
        }

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
     * Set style of single or multiple column(s)
     *
     * Styles are applied to the entire sheet column(s) (even if it is empty)
     *
     * @param int|string|array $colRange Column number or column letter (or array of these)
     * @param mixed $style
     *
     * @return $this
     *
     * @example
     * $sheet->setColStyle('B', $style);
     * $sheet->setColStyle(2, $style); // 'B' is number 2 column
     * $sheet->setColStyle('C:F', $style);
     * $sheet->setColStyle(['A', 'B', 'C'], $style);
     */
    public function setColStyle($colRange, $style): Sheet
    {
        if ($this->currentColIdx) {
            $this->_writeCurrentRow();
        }
        $this->_setColCellStyle($colRange, $style, true);
        $this->clearSummary();

        return $this;
    }

    /**
     * Set style of single or multiple column(s)
     *
     * Styles are applied to the entire sheet column(s) (even if it is empty)
     *
     * @param array $colStyles
     *
     * @return $this
     *
     * @example
     * $sheet->setColStyleArray(['B' => ['width' = 20], 'C' => ['font-color' = '#f00']]);
     */
    public function setColStyleArray(array $colStyles): Sheet
    {
        foreach ($colStyles as $col => $style) {
            $this->setColStyle($col, $style);
        }

        return $this;
    }

    /**
     * Set style of column cells (colors, formats, etc.)
     *
     * Styles are applied only to non-empty cells in a column and only take effect starting with the current row
     *
     * @param int|string|array $colRange
     * @param array $colStyle
     *
     * @return $this
     *
     * @example
     * $sheet->setColDataStyle('B', ['width' = 20]); // style for cells of column 'B'
     * $sheet->setColDataStyle(2, ['width' = 20]); // 'B' is number 2 column
     * $sheet->setColDataStyle('B:D', ['width' = 'auto']); // options for range of columns
     * $sheet->setColDataStyle(['A', 'B', 'C'], $style); // options for several columns 'A', 'B' and 'C'
     */
    public function setColDataStyle($colRange, array $colStyle): Sheet
    {
        if ($this->currentColIdx) {
            $this->_writeCurrentRow();
        }

        $options = [];
        $colNumbers = Excel::colNumberRange($colRange);
        if ($colNumbers) {
            $colOptions = StyleManager::normalize($colStyle);
            foreach ($colNumbers as $col) {
                $options[$col] = $colOptions;
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
                    $this->_setColCellStyle($col, $style, false);
                }
            }
        }
        $this->clearSummary();

        return $this;
    }

    /**
     * Set style of column cells (colors, formats, etc.)
     *
     * Styles are applied only to non-empty cells in a column and only take effect starting with the current row
     *
     * @param array $colStyles
     *
     * @return $this
     *
     * @example
     * $sheet->setColDataStyleArray(['B' => $style1, 'C' => $style2]); // options for columns 'B' and 'C'
     */
    public function setColDataStyleArray(array $colStyles): Sheet
    {
        $styles = array_combine(Excel::colLetterRange(array_keys($colStyles)), array_values($colStyles));
        foreach ($styles as $col => $style) {
            $this->setColDataStyle($col, $style);
        }

        return $this;
    }

    /**
     * @param $arg1
     * @param array|null $arg2
     *
     * @return $this
     *@deprecated since v.6.1
     *
     */
    public function setColStyles($arg1, ?array $arg2 = null): Sheet
    {
        return $this->setColStyle($arg1, $arg2);
    }

    /**
     * Use 'setColDataStyle()' or 'setColDataStyleArray()' instead
     * @deprecated since v.6.1
     */
    public function setColOptions($arg1, ?array $arg2 = null): Sheet
    {
        if ($arg2 === null) {
            return $this->setColDataStyleArray($arg1);
        }

        return $this->setColDataStyle($arg1, $arg2);
    }

    /**
     * Show/hide a column
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param bool $val
     *
     * @return $this
     */
    public function setColVisible($col, bool $val): Sheet
    {
        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                if ($val) {
                    $this->_delColAttributes($colIdx, ['hidden']);
                }
                else {
                    $this->_setColAttributes($colIdx, ['hidden' => 1]);
                }
            }
        }

        return $this;
    }

    /**
     * @param int|string|array $col Column number or column letter (or array of these)
     *
     * @return $this
     */
    public function setColHidden($col): Sheet
    {

        return $this->setColVisible($col, false);
    }

    /**
     * Set width of single or multiple column(s)
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param int|float|string $width
     * @param bool|null $min
     *
     * @return $this
     */
    public function setColWidth($col, $width, ?bool $min = false): Sheet
    {
        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                if (is_string($width) && strtolower($width) === 'auto') {
                    $this->colStyles[$colIdx]['options']['width-auto'] = true;
                }
                elseif ($width !== null) {
                    $width = ($width ? StyleManager::numFloat($width) : 0);
                    if (is_numeric($width)) {
                        if ($min) {
                            $this->colMinWidths[$colIdx] = $width;
                            if (!isset($this->colAttributes[$colIdx]['width']) || $this->colAttributes[$colIdx]['width'] < $width) {
                                $this->_setColAttributes($colIdx, ['width' => $width, 'customWidth' => '1']);
                            }
                        }
                        elseif (empty($this->colMinWidths[$colIdx]) || $this->colMinWidths[$colIdx] <= $width) {
                            $this->_setColAttributes($colIdx, ['width' => $width, 'customWidth' => '1']);
                        }
                    }
                }
            }
        }
        $this->clearSummary();

        return $this;
    }

    /**
     * Setting a multiple column's width
     *
     * @example
     * $sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
     *
     * @param array $widths
     * @param bool|null $min
     *
     * @return $this
     */
    public function setColWidths(array $widths, ?bool $min = false): Sheet
    {
        if ($widths) {
            $widths = Excel::colKeysToLetters($widths);
            foreach ($widths as $col => $width) {
                $this->setColWidth($col, $width, $min);
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
     * @param int|string|array $col Column number or column letter (or array of these)
     *
     * @return $this
     */
    public function setColAutoWidth($col): Sheet
    {
        return $this->setColWidthAuto($col);
    }

    /**
     * Setting a minimal column's width
     *
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param int|float|string $width
     *
     * @return $this
     */
    public function setColMinWidth($col, $width): Sheet
    {
        return $this->setColWidth($col, $width, true);
    }

    /**
     * Setting a multiple column's minimal width
     *
     * @example
     * $sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
     *
     * @param array $widths
     *
     * @return $this
     */
    public function setColMinWidths(array $widths): Sheet
    {
        return $this->setColWidths($widths, true);
    }

    /**
     * @param int|string|array $col Column number or column letter (or array of these)
     * @param int $outlineLevel
     *
     * @return $this
     */
    public function setColOutlineLevel($col, int $outlineLevel): Sheet
    {
        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                $attr = ['outlineLevel' => $outlineLevel];
                $this->_setColAttributes($colIdx, $attr);
            }
        }

        return $this;

    }

    /**
     * @return array
     */
    public function getColAttributes(): array
    {
        $result = [];
        if ($this->colAttributes) {
            foreach ($this->colAttributes as $colIdx => $attributes) {
                if ($attributes) {
                    if (isset($this->colAttributes[$colIdx]['min'], $this->colAttributes[$colIdx]['max']) && count($this->colAttributes[$colIdx]) === 2) {
                        // only 'min' & 'max'
                        continue;
                    }
                    $result[$colIdx] = $attributes;
                    if (!isset($result[$colIdx]['min'])) {
                        $result[$colIdx]['min'] = $colIdx + 1;
                    }
                    if (!isset($result[$colIdx]['max'])) {
                        $result[$colIdx]['max'] = $colIdx + 1;
                    }
                    if (isset($attributes['width'])) {
                        $result[$colIdx]['width'] = number_format($attributes['width'], 8, '.', '');
                    }
                    if (!empty($attributes['hidden'])) {
                        $result[$colIdx]['hidden'] = '1';
                    }
                    if (!isset($result[$colIdx]['hidden']) && !isset($result[$colIdx]['width'])) {
                        $result[$colIdx]['width'] = Excel::DEFAULT_COL_WIDTH;
                    }
                }
            }
            //ksort($result);
        }

        return $result;
    }

    /**
     * @param int $colIdx
     * @param array $settings
     *
     * @return void
     */
    public function _setColAttributes(int $colIdx, array $settings)
    {
        if (!isset($this->colAttributes[$colIdx]['min'])) {
            $this->colAttributes[$colIdx]['min'] = $colIdx + 1;
            $this->colAttributes[$colIdx]['max'] = $colIdx + 1;
        }
        foreach ($settings as $key => $val) {
            $this->colAttributes[$colIdx][$key] = $val;
        }
    }

    /**
     * @param int $colIdx
     * @param array $settings
     *
     * @return void
     */
    public function _delColAttributes(int $colIdx, array $settings)
    {
        foreach ($settings as $key) {
            if ($this->colAttributes[$colIdx][$key]) {
                unset($this->colAttributes[$colIdx][$key]);
            }
        }
    }

    protected function _setColCellStyle($col, array $style, bool $whole): Sheet
    {
        $colIndexes = Excel::colIndexRange($col);
        foreach($colIndexes as $colIdx) {
            if ($colIdx >= 0) {
                $style = StyleManager::normalize($style);
                if (!empty($this->colStyles[$colIdx])) {
                    $this->colStyles[$colIdx] = array_replace_recursive($this->colStyles[$colIdx], $style);
                }
                else {
                    $this->colStyles[$colIdx] = $style;
                }
                if ($whole) {
                    $this->colAttributes[$colIdx]['style'] = $this->excel->addStyle($style, $resultStyle);
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
    public function setColFormula($col, string $formula): Sheet
    {
        if ($this->currentColIdx) {
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
     * @param $rowNum
     * @param $key
     * @param $val
     *
     * @return void
     */
    protected function _setRowSettings($rowNum, $key, $val)
    {
        if ($rowNum <= $this->rowCountWritten) {
            ExceptionAddress::throwNew('Row number must be greater than written rows');
        }
        $rowIdx = (int)$rowNum - 1;
        if ($key === 'height') {
            $this->rowAttributes[$rowIdx]['customHeight'] = 1;
            $this->rowAttributes[$rowIdx]['ht'] = $val;
        }
        elseif ($key === 'style') {
            $this->rowAttributes[$rowIdx]['customFormat'] = 1;
            $this->rowAttributes[$rowIdx]['s'] = $val;
        }
        elseif ($key === 'hidden' || $key === 'outlineLevel' || $key === 'collapsed') {
            $this->rowAttributes[$rowIdx][$key] = $val;
        }
        else {
            $this->rowSettings[$rowIdx][$key] = $val;
        }
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
            $this->_setRowSettings($rowNum, 'height', Writer::floatStr($height));
        }
        else {
            $address = $this->_parseAddress($rowNum, null, true);
            for ($row = $address['rowNum1']; $row <= $address['rowNum2']; $row++) {
                $this->setRowHeight($row, $height);
            }
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
     * Hide/show a specific row
     *
     * @param int|array $rowNum
     * @param bool $visible
     *
     * @return $this
     */
    public function setRowVisible($rowNum, bool $visible): Sheet
    {
        if (is_array($rowNum)) {
            foreach ($rowNum as $row) {
                $this->setRowVisible($row, $visible);
            }
        }
        elseif (is_numeric($rowNum)) {
            $this->_setRowSettings($rowNum, 'hidden', $visible ? 0 : 1);
        }
        else {
            $address = $this->_parseAddress($rowNum, null, true);
            for ($row = $address['rowNum1']; $row <= $address['rowNum2']; $row++) {
                $this->setRowVisible($row, $visible);
            }
        }

        return $this;
    }

    /**
     * Hide a specific row
     *
     * @param int|array $rowNum
     *
     * @return $this
     */
    public function setRowHidden($rowNum): Sheet
    {

        return $this->setRowVisible($rowNum, false);
    }

    /**
     *
     * @example
     * setRowOutlineLevel(5, 1)
     * setRowOutlineLevel([5, 6, 7], 1)
     * setRowOutlineLevel('5:7', 1)
     *
     * @param int|array|string $rowNum
     * @param int $outlineLevel
     * @param bool|null $collapsed
     *
     * @return $this
     */
    public function setRowOutlineLevel($rowNum, int $outlineLevel, ?bool $collapsed = null): Sheet
    {
        if (is_array($rowNum)) {
            foreach ($rowNum as $row) {
                $this->setRowOutlineLevel($row, $outlineLevel, $collapsed);
            }
        }
        elseif (is_numeric($rowNum)) {
            $this->_setRowSettings($rowNum, 'outlineLevel', $outlineLevel);
            if ($collapsed !== null && !isset($this->rowOutlineCollapsed[$rowNum])) {
                $this->rowOutlineCollapsed[$outlineLevel] = $collapsed;
            }
            if (!empty($this->rowOutlineCollapsed[$outlineLevel])) {
                $this->_setRowSettings($rowNum, 'hidden', 1);
            }
        }
        else {
            $address = $this->_parseAddress($rowNum, null, true);
            for ($row = $address['rowNum1']; $row <= $address['rowNum2']; $row++) {
                $this->setRowOutlineLevel($row, $outlineLevel, $collapsed);
            }
        }

        return $this;
    }

    /**
     * @return $this
     */
    public function beginOutlineLevel(?bool $collapsed = false): Sheet
    {
        $this->rowOutlineLevel++;
        $this->rowOutlineCollapsed[$this->rowOutlineLevel] = $collapsed;

        return $this;
    }

    /**
     * @return $this
     */
    public function endOutlineLevel(): Sheet
    {
        if ($this->rowOutlineLevel > 0) {
            $collapsed = !empty($this->rowOutlineCollapsed[$this->rowOutlineLevel]);
            if ($collapsed) {
                // collapse the next row
                $this->_setRowSettings($this->currentRowIdx + 2, 'collapsed', 1);
            }
            $this->rowOutlineLevel--;
        }

        return $this;
    }

    /**
     * @return int
     */
    public function getOutlineLevel(): int
    {
        return $this->rowOutlineLevel;
    }

    /**
     * @param int $rowNum
     * @param array $rowOptions
     * @param bool $whole
     *
     * @return void
     */
    protected function _setRowOptions(int $rowNum, array $rowOptions, bool $whole)
    {
        if ($rowNum <= $this->rowCountWritten) {
            ExceptionAddress::throwNew('Row number must be greater than written rows');
        }
        $rowIdx = $rowNum - 1;
        if ($rowOptions) {
            $rowOptions = StyleManager::normalize($rowOptions);
            if (isset($rowOptions['height'])) {
                if (!$whole) {
                    Exception::throwNew('The "height" parameter can only be set for the entire row');
                }
                $this->setRowHeight($rowNum, $rowOptions['height']);
                unset($rowOptions['height']);
            }
            if ($rowOptions) {
                $styleIdx = $this->excel->addStyle($rowOptions, $resultStyle);
                if ($whole && $styleIdx > 0) {
                    $this->_setRowSettings($rowNum, 'style', $this->excel->addStyle($rowOptions, $resultStyle));
                }
                if (isset($this->rowStyles[$rowIdx])) {
                    $this->rowStyles[$rowIdx] = array_replace_recursive($this->rowStyles[$rowIdx], $rowOptions);
                }
                else {
                    $this->rowStyles[$rowIdx] = $rowOptions;
                }
            }
        }
    }

    /**
     * Use 'setRowDataStyle()' or 'setRowDataStyleArray()' instead
     * @deprecated since v.6.1
     */
    public function setRowOptions($arg1, ?array $arg2 = null): Sheet
    {
        if ($arg2 === null) {
            return $this->setRowDataStyleArray($arg1);
        }

        return $this->setRowDataStyle($arg1, $arg2);
    }

    /**
     * @param $arg1
     * @param array|null $arg2
     *
     * @return $this
     * @deprecated since v.6.1
     *
     */
    public function setRowStyles($arg1, ?array $arg2 = null): Sheet
    {
        return $this->setRowOptions($arg1, $arg2);
    }

    /**
     * Style are applied to the entire sheet row (even if it is empty)
     *
     * @param int|string|array $rowRange
     * @param array $style
     *
     * @return $this
     *
     * @example
     * $sheet->setRowStyle(3, ['height' = 20]); // options for row number 3
     * $sheet->setRowStyle('2:5', ['font-color' = '#f00']); // options for range of rows
     */
    public function setRowStyle($rowRange, array $style): Sheet
    {
        $rows = Excel::rowNumberRange($rowRange);
        foreach ($rows as $rowNum) {
            $this->_setRowOptions($rowNum, $style, true);
        }

        return $this;
    }

    /**
     * Styles are applied to the entire sheet row (even if it is empty)
     *
     * @param array $rowStyles
     *
     * @return $this
     *
     * @example
     * $sheet->setRowStyleArray([3 => $style1, 5 => $style2]); // styles for rows 3 and 5
     */
    public function setRowStyleArray(array $rowStyles): Sheet
    {
        foreach ($rowStyles as $rowNum => $style) {
            $this->_setRowOptions($rowNum, $style, true);
        }
        return $this;
    }

    /**
     * Style are applied only to non-empty cells in a row (or row range)
     *
     * @param int|string|array $rowRange
     * @param array $style
     *
     * @return $this
     *
     * @example
     * $sheet->setRowDataStyle(3, ['height' = 20]); // options for row number 3
     * $sheet->setRowDataStyle('2:5', ['font-color' = '#f00']); // options for range of rows
     */
    public function setRowDataStyle($rowRange, array $style): Sheet
    {
        $rows = Excel::rowNumberRange($rowRange);
        foreach ($rows as $rowNum) {
            $this->_setRowOptions($rowNum, $style, false);
        }

        return $this;
    }

    /**
     * Styles are applied only to non-empty cells in a rows
     *
     * @param array $rowStyles
     *
     * @return $this
     *
     * @example
     * $sheet->setRowDataStyleArray([3 => $style1, 5 => $style2]); // styles for rows 3 and 5
     */
    public function setRowDataStyleArray(array $rowStyles): Sheet
    {
        foreach ($rowStyles as $row => $style) {
            if (is_numeric($row)) {
                $rows = Excel::rowNumberRange($row);
                foreach ($rows as $rowNum) {
                    $this->_setRowOptions($rowNum, $style, false);
                }
            }
            else {
                $this->_setRowOptions($row, $style, false);
            }
        }

        return $this;
    }

    /**
     * @param int $rowIdx
     * @param array|null $rowOptions
     *
     * @return array
     */
    protected function getRowAttributes(int $rowIdx, ?array $rowOptions = []): array
    {
        $rowAttributes = $this->rowAttributes[$rowIdx] ?? [];
        if (!empty($rowOptions['height'])) {
            $rowAttributes['customHeight'] = 1;
            $rowAttributes['ht'] = Writer::floatStr($rowOptions['height']);
        }
        if (!empty($rowOptions['hidden'])) {
            $rowAttributes['hidden'] = 1;
        }
        if (!empty($rowOptions['collapsed'])) {
            $rowAttributes['collapsed'] = 1;
        }

        return $rowAttributes;
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

        $rowIdx = $this->rowCountWritten;
        if (isset($this->rowSettings[$rowIdx])) {
            $rowOptions = array_replace($this->rowSettings[$rowIdx], $rowOptions);
        }

        $rowAttributes = $this->getRowAttributes($rowIdx, $rowOptions);
        $rowAttrStr = Writer::tagAttributes($rowAttributes);

        // add auto formulas of columns
        if ($this->colFormulas && $row) {
            foreach($this->colFormulas as $colIdx => $formula) {
                if (!isset($row[$colIdx])) {
                    $row[$colIdx] = $formula;
                }
            }
            ksort($row);
        }

        if ($row || ($row === [null] && $rowAttrStr)) {
            if (empty($this->sheetStylesSummary)) {
                if ($this->defaultStyle) {
                    $this->sheetStylesSummary = [
                        'general_style' => StyleManager::mergeStyles([$this->excel->getDefaultStyle(), $this->defaultStyle]),
                        'hyperlink_style' => StyleManager::mergeStyles([$this->excel->getHyperlinkStyle(), $this->defaultStyle]),
                    ];
                }
                else {
                    $this->sheetStylesSummary = [
                        'general_style' => $this->excel->getDefaultStyle(),
                        'hyperlink_style' => $this->excel->getHyperlinkStyle(),
                    ];
                }
            }
            if ($row && $row !== [null]) {
                $this->fileWriter->write('<row r="' . ($this->rowCountWritten + 1) . '" ' . $rowAttrStr . '>');
                foreach ($row as $colIdx => $cellValue) {
                    if (!isset($this->colStylesSummary[$colIdx])) {
                        if (!isset($this->colStyles[$colIdx])) {
                            $this->colStylesSummary[$colIdx] = $this->sheetStylesSummary;
                        }
                        else {
                            $this->colStylesSummary[$colIdx] = [
                                'general_style' => StyleManager::mergeStyles([
                                    $this->sheetStylesSummary['general_style'],
                                    $this->colStyles[$colIdx],
                                ]),
                                'hyperlink_style' => StyleManager::mergeStyles([
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

                    if (isset($cellsOptions[$colIdx]['_xf_id'])) {
                        $cellStyleIdx = $cellsOptions[$colIdx]['_xf_id'];
                        $numberFormatType = $cellsOptions[$colIdx]['number_format_type'] ?? 'n_auto';
                    }
                    else {
                        // Define cell style index and number format
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
                            $cellStyle = StyleManager::mergeStyles($styleStack);
                        }
                        else {
                            $cellStyle = $styleStack ? $styleStack[0] : [];
                        }
                        if (!empty($cellStyle['format']['format-pattern']) && !empty($this->excel->getDefaultFormatStyles()[$cellStyle['format']['format-pattern']])) {
                            $cellStyle = StyleManager::mergeStyles([$this->excel->getDefaultFormatStyles()[$cellStyle['format']['format-pattern']], $cellStyle]);
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
                                if (!empty($this->excel->getHyperlinkStyle())) {
                                    $cellStyle = StyleManager::mergeStyles([$this->excel->getHyperlinkStyle(), $cellStyle]);
                                }
                            }
                            unset($cellStyle['hyperlink']);
                        }

                        $styleHash = $cellStyle ? json_encode($cellStyle) : '';
                        if (!isset($_styleCache[$styleHash])) {
                            $cellStyleIdx = $this->excel->addStyle($cellStyle, $resultStyle);
                            $_styleCache[$styleHash] = ['cell_style' => $cellStyle, 'result_style' => $resultStyle, 'style_idx' => $cellStyleIdx];
                        }
                        else {
                            $resultStyle = $_styleCache[$styleHash]['result_style'];
                            $cellStyleIdx = $_styleCache[$styleHash]['style_idx'];
                        }

                        $numberFormat = $resultStyle['number_format'];
                        $numberFormatType = $resultStyle['number_format_type'];

                        if (!empty($cellStyle['options']['width-auto']) && !($cellValue && is_string($cellValue) && $cellValue[0] === '=')) {
                            $this->_columnWidth($colIdx, $cellValue, $numberFormat, $resultStyle ?? []);
                        }

                        if (!$writer) {
                            $writer = $this->excel->getWriter();
                        }
                    }
                    if ($cellValue !== null || $cellStyleIdx !== 0 || $numberFormatType !== 'n_auto') {
                        $writer->_writeCell($this->fileWriter, $rowIdx + 1, $colIdx + 1, $cellValue, $numberFormatType, $cellStyleIdx);
                        $this->_setDimension($rowIdx + 1, $colIdx + 1);
                    }
                    $colIdx++;
                    if ($colIdx > $this->colCountWritten) {
                        $this->colCountWritten = $colIdx;
                    }
                }
                $this->fileWriter->write('</row>');
            }
            elseif ($rowAttrStr) {
                $this->fileWriter->write('<row r="' . ($this->rowCountWritten + 1) . '" ' . $rowAttrStr . '/>');
            }
        }
        $this->rowCountWritten++;
    }

    /**
     * @param int $colIdx
     * @param $cellValue
     * @param $numberFormat
     * @param $style
     */
    protected function _columnWidth(int $colIdx, $cellValue, $numberFormat, $style)
    {
        if ($cellValue) {
            $fontName = $style['font']['val']['name'] ?? Font::DEFAULT_FONT_NAME;
            $fontSize = $style['font']['val']['size'] ?? Font::DEFAULT_FONT_SIZE;
            $value = (isset($cellValue['shared_value'])) ? $cellValue['shared_value'] : $cellValue;

            $len = Font::calcTextWidth($fontName, $fontSize, $value, $numberFormat);
            if ($this->autoFilter) {
                $len += 1;
            }

            if ((empty($this->colAttributes[$colIdx]['width']) || $this->colAttributes[$colIdx]['width'] < $len) && (empty($this->colMinWidths[$colIdx]) || $this->colMinWidths[$colIdx] <= $len)) {
                $this->_setColAttributes($colIdx, ['width' => $len]);
            }
        }
    }

    /**
     * @return void
     */
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
        $this->defaultStyle = StyleManager::normalize($style);

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
        $normStyle = StyleManager::normalizeFont($font);
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

    /**
     * @return $this
     */
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


    protected function _checkOutput()
    {
        if ($this->excel->saved) {
            Exception::throwNew('The output file is already saved');
        }
    }

    /**
     * Write value to the current cell and move pointer to the next cell in the row
     *
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     */
    public function writeCell($value, ?array $styles = null): Sheet
    {
        $this->_checkOutput();

        if ($this->lastTouch['ref'] === 'row') {
            $this->_writeCurrentRow();
        }
        ///-- $styles = $styles ? Style::normalize($styles) : [];
        if ($this->currentRowIdx < $this->rowCountWritten) {
            $this->currentRowIdx = $this->rowCountWritten;
        }
        //$this->withLastCell();
        $cellAddress = [
            'row' => 1 + $this->currentRowIdx,
            'col' => 1 + $this->currentColIdx,
        ];
        $this->_setCellData($cellAddress, $value, $styles, false);
        $this->_touchEnd($this->currentRowIdx, $this->currentColIdx, 'cell');
        ++$this->currentColIdx;

        return $this;
    }

    /**
     *  Write several values into cells of one row
     *
     * @param array $values
     * @param array|null $cellStyles
     *
     * @return $this
     */
    public function writeCells(array $values, ?array $cellStyles = null): Sheet
    {
        $this->_checkOutput();

        if ($this->lastTouch['ref'] === 'row') {
            $this->_writeCurrentRow();
        }
        if ($this->currentRowIdx < $this->rowCountWritten) {
            $this->currentRowIdx = $this->rowCountWritten;
        }
        $startRowIdx = $this->currentRowIdx;
        $startColIdx = $this->currentColIdx;

        if (is_array($cellStyles)) {
            $key = array_key_first($cellStyles);
            if (!is_int($key)) {
                $cellStyles = Excel::colKeysToIndexes($cellStyles, -$startColIdx);
            }
            else {
                $cellStyles = [];
            }
        }
        else {
            $cellStyles = [];
        }

        $cellAddress = [
            'row' => 1 + $this->currentRowIdx,
            'col' => 1 + $this->currentColIdx,
        ];
        foreach ($values as $pos => $value) {
            $styles = !empty($cellStyles[$pos]) ? $cellStyles[$pos] : null;
            $this->_setCellData($cellAddress, $value, $styles, false);
            $cellAddress['col']++;
            ++$this->currentColIdx;
        }
        $this->_touchStart($startRowIdx, $startColIdx, 'area');
        $this->_touchEnd($startRowIdx, $this->currentColIdx - 1, 'area');

        return $this;
    }

    /**
     * @return $this
     */
    public function nextCell(): Sheet
    {
        $this->writeCell(null);

        return $this;
    }

    /**
     * @param array $header
     * @param array|null $rowStyle
     * @param array|null $colStyles
     *
     * @return $this
     *@example
     * $sheet->writeHeader(['title1', 'title2', 'title3']); // texts for cells of header
     * $sheet->writeHeader(['title1' => '@text', 'title2' => 'YYYY-MM-DD', 'title3' => ['format' => ..., 'font' => ...]]); // texts and formats of columns
     * $sheet->writeHeader($cellValues, $rowStyle, $colStyles); // texts and formats of columns and options of row
     *
     */
    public function writeHeader(array $header, ?array $rowStyle = null, ?array $colStyles = []): Sheet
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
                    $colStyles[$colNum + $this->offsetCol] = isset($colStyles[$colNum + $this->currentColIdx]) ? array_replace_recursive($colStyles[$colNum + $this->currentColIdx], $val) : $val;
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
            ExceptionAddress::throwNew('Row number must be greater than written rows');
        }
        else {
            $cellAddress = $address;
        }

        if (isset($address['colIndex'], $address['rowIndex'])) {
            $this->currentColIdx = (int)$address['colIndex'];
            $this->currentRowIdx = (int)$address['rowIndex'];
        }
        else {
            while ($this->currentRowIdx < $cellAddress['row'] - 1) {
                $this->nextRow();
            }
        }

        return $address;
    }

    /**
     * Write value to the specified cell and move pointer to the next cell in the row
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     *
     * @example
     * $sheet->writeTo('B5', $value); // write to single cell
     * $sheet->writeTo('B5:C7', $value); // write a value to merged cells
     * $sheet->writeTo(['col' => 2, 'row' => 5], $value); // address as array
     * $sheet->writeTo([2, 5], $value); // address as array
     *
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
            $this->currentColIdx += $address['width'];
        }
        else {
            $this->currentColIdx++;
        }
        if ($address['rowNum2'] === $address['rowNum1'] && $address['colNum2'] === $address['colNum1']) {
            $ref = 'cell';
        }
        elseif ($address['rowNum2'] === $address['rowNum1']) {
            $ref = 'row';
        }
        else {
            $ref = 'area';
        }
        $this->_touchEnd($address['rowNum2'] - 1, $address['colNum2'] - 1, $ref);

        return $this;
    }

    /**
     * Write 2d array form the specified cell
     *
     * @param $topLeftCell
     * @param array $data
     *
     * @return $this
     */
    public function writeArrayTo($topLeftCell, array $data): Sheet
    {
        if (preg_match('/^([a-z]+)(\d+)$/i', $topLeftCell, $m)) {
            $colNumBegin = Excel::colNumber($m[1]);
            $rowNum = $m[2];
            foreach ($data as $rowData) {
                $colNum = $colNumBegin;
                foreach ($rowData as $cellValue) {
                    $this->writeTo(Excel::cellAddress($rowNum, $colNum++), $cellValue);
                }
                $rowNum++;
            }
        }

        return $this;
    }

    /**
     * Merge cells
     *
     * @example
     * $sheet->mergeCells('A1:C3');
     * $sheet->mergeCells(['A1:B2', 'C1:D2']);
     *
     * @param array|string|int $rangeSet
     * @param int|null $actionMode Action in case of intersection:
     *      0 - exception;
     *      1 - replace;
     *      2 - keep;
     *      -1 - skip intersection check
     *
     * @return $this
     */
    public function mergeCells($rangeSet, ?int $actionMode = 0): Sheet
    {
        foreach((array)$rangeSet as $range) {
            if (isset($this->mergeCells[$range]) || empty($range)) {
                // cells are already merged
                continue;
            }
            $range = strtoupper($range);
            if (preg_match('/^(\$?[A-Z]+)(\$?\d+)(:(\$?[A-Z]+)(\$?\d+))?$/', $range, $matches)) {
                if (empty($matches[3])) {
                    $matches[4] = $matches[1];
                    $matches[5] = $matches[2];
                }
                $dimension = [
                    'range' => $range,
                    'rowNum1' => ($matches[2] <= Excel::MAX_ROW) ? (int)$matches[2] : -1,
                    'rowNum2' => ($matches[5] <= Excel::MAX_ROW) ? (int)$matches[5] : -1,
                    'colNum1' => Excel::colNumber($matches[1]),
                    'colNum2' => Excel::colNumber($matches[4]),
                ];

                if ($actionMode > -1) {
                    if (!(
                        $dimension['rowNum1'] > $this->mergedCellsArray['rowNum2'] ||
                        $dimension['rowNum2'] < $this->mergedCellsArray['rowNum1'] ||
                        $dimension['colNum1'] > $this->mergedCellsArray['colNum2'] ||
                        $dimension['colNum2'] < $this->mergedCellsArray['colNum1']
                    )) {
                        // is intersection
                        foreach ($this->mergeCells as $savedRange => $savedDimension) {
                            if (!(
                                $dimension['rowNum1'] > $savedDimension['rowNum2'] ||
                                $dimension['rowNum2'] < $savedDimension['rowNum1'] ||
                                $dimension['colNum1'] > $savedDimension['colNum2'] ||
                                $dimension['colNum2'] < $savedDimension['colNum1']
                            )) {
                                if ($actionMode === 1) {
                                    unset($this->mergeCells[$savedRange]);
                                }
                                elseif ($actionMode === 2) {
                                    $dimension = [];
                                    break;
                                }
                                else {
                                    ExceptionAddress::throwNew('Cannot merge cells %s because they intersect with %s', $range, $savedRange);
                                }
                            }
                        }
                    }
                }
            }
            else {
                ExceptionAddress::throwNew('Wrong range ' . print_r($range, true));
            }

            if (!empty($dimension)) {
                $this->mergeCells[$range] = $dimension;
                if ($this->mergedCellsArray['rowNum1'] > $dimension['rowNum1']) {
                    $this->mergedCellsArray['rowNum1'] = $dimension['rowNum1'];
                }
                if ($this->mergedCellsArray['rowNum2'] < $dimension['rowNum2']) {
                    $this->mergedCellsArray['rowNum2'] = $dimension['rowNum2'];
                }
                if ($this->mergedCellsArray['colNum1'] > $dimension['colNum1']) {
                    $this->mergedCellsArray['colNum1'] = $dimension['colNum1'];
                }
                if ($this->mergedCellsArray['colNum2'] < $dimension['colNum2']) {
                    $this->mergedCellsArray['colNum2'] = $dimension['colNum2'];
                }
            }
        }

        return $this;
    }

    /**
     * Merge relative cells
     *
     * @example
     * $sheet->mergeCells('C3:E8');
     * $sheet->mergeCells(3); // 3 columns of current row, equivalent of mergeCells('A5:C5') if current row is 5
     * $sheet->mergeCells(['RC3:RC5', 'RC6:RC7']); // equivalent of mergeCells(['C7:E7', 'F7:G7']) if current row is 7
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
            $dimension = $this->_rangeDimension($range, 1, 0);
            $this->mergeCells($dimension['range']);
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
     * @return int
     */
    protected function _writeCurrentRow(): int
    {
        $savedRow = $this->currentRowIdx;
        if (!empty($this->cells['values']) || !empty($this->cells['styles']) || $this->rowSettings) {
            $writer = $this->excel->getWriter();
            if (!$this->open) {
                $writer->writeSheetDataBegin($this);
            }
            $maxRowIdx = max($this->cells['values'] ? max(array_keys($this->cells['values'])) : -1,
                $this->cells['styles'] ? max(array_keys($this->cells['styles'])) : -1);
            if ($this->rowAttributes) {
                $maxRowIdx = max($maxRowIdx, max(array_keys($this->rowAttributes)));
            }
            if ($maxRowIdx >= 0) {
                // has values or styles
                if ($maxRowIdx < $this->currentRowIdx) {
                    $maxRowIdx = $this->currentRowIdx;
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
                    $rowSettings = $this->rowSettings[$rowIdx] ?? [];
                    if ($this->rowOutlineLevel) {
                        $this->setRowOutlineLevel($this->rowCountWritten + 1, $this->rowOutlineLevel);
                    }

                    if ($values || $styles) {
                        ksort($values);
                        ksort($styles);
                        $this->_writeRow($writer, $values, $rowSettings, $styles);
                    }
                    elseif ($rowSettings || !empty($this->rowAttributes[$rowIdx])) {
                        $this->_writeRow($writer, [null], $rowSettings, []);
                    }
                    else {
                        //$this->rowCount++;
                        $this->_writeRow($writer, [], [], []);
                    }
                    if (isset($this->rowStyles[$rowIdx])) {
                        unset($this->rowStyles[$rowIdx]);
                    }
                }
                $this->currentRowIdx++;
            }

            $this->currentColIdx = $this->offsetCol;

            if (isset($this->colStyles[-1])) {
                // column styles for next rows
                $this->setColDataStyle($this->currentRowIdx + 1, $this->colStyles[-1]);
                unset($this->colStyles[-1]);
            }
            $this->withLastCell();
        }

        return $this->currentRowIdx - $savedRow;
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
    public function writeRow(array $rowValues = [], ?array $rowStyle = null, ?array $cellStyles = null): Sheet
    {
        $this->_checkOutput();

        if (($this->currentColIdx > $this->offsetCol) || $this->areas) {
            $this->_writeCurrentRow();
        }

        if (!is_array($rowValues)) {
            $rowValues = [$rowValues];
        }
        else {
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
            $rowStyle = StyleManager::normalize($rowStyle);
            $this->rowStyles[$this->currentRowIdx] = $rowStyle;
            if (isset($rowStyle['options']['height'])) {
                $this->setRowHeight($this->currentRowIdx + 1, $rowStyle['options']['height']);
            }
        }

        if ($this->rowOutlineLevel) {
            $this->setRowOutlineLevel($this->currentRowIdx + 1, $this->rowOutlineLevel);
        }

        $this->lastTouch['area']['col_idx1'] = $this->lastTouch['area']['col_idx2'] = -1;
        $maxColIdx = max($cellStyles ? max(array_keys($cellStyles)) : 0, $rowValues ? max(array_keys($rowValues)) : 0);
        $this->_touchStart($this->currentRowIdx, $this->offsetCol, 'row');
        for ($colIdx = 0; $colIdx <= $maxColIdx; $colIdx++) {
            if (isset($rowValues[$colIdx]) || isset($cellStyles[$colIdx])) {
                if ($this->lastTouch['area']['col_idx1'] === -1) {
                    $this->lastTouch['area']['col_idx1'] = $colIdx;
                }
                $this->lastTouch['area']['col_idx2'] = $colIdx;
                $this->_setCellData(['col_idx' => $this->offsetCol + $colIdx, 'row_idx' => $this->currentRowIdx], $rowValues[$colIdx] ?? null, $cellStyles[$colIdx] ?? null);
            }
            $this->lastTouch['cell']['col_idx'] = ++$this->currentColIdx;
        }
        $this->_touchEnd($this->currentRowIdx, $maxColIdx, 'row');

        return $this;
    }

    /**
     * Write values from two-dimensional array
     *
     * @param array $rowArray Array of rows
     * @param array|null $rowStyle Style applied to each row
     *
     * @return $this
     */
    public function writeArray(array $rowArray = [], ?array $rowStyle = null): Sheet
    {
        foreach ($rowArray as $rowValues) {
            $this->writeRow($rowValues, $rowStyle);
        }

        return $this;
    }

    /**
     * Move to the next row
     *
     * @param array|null $style
     *
     * @return $this
     */
    public function nextRow(?array $style = []): Sheet
    {
        $this->_checkOutput();

        $ref = $this->lastTouch['ref'];
        $writtenRows = $this->_writeCurrentRow();
        if ($writtenRows) {
            if ($ref === 'row') {
                $this->currentRowIdx++;
            }
            $this->currentColIdx = $this->offsetCol;
            $this->_touch($this->currentRowIdx, $this->currentColIdx, $this->currentRowIdx, $this->currentColIdx, 'cell');
        }
        else {
            $this->currentRowIdx++;
        }
        if (!empty($style)) {
            $this->_setRowOptions($this->currentRowIdx + 1, $style, true);
        }

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
        for ($i = 0; $i <= $rowCount; $i++) {
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
        $this->_checkOutput();

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
        $this->_touch($coord[0]['row'] - 1, $coord[0]['col'] - 1, $coord[1]['row'] - 1, $coord[1]['col'] - 1, 'area');

        return $area;
    }

    /**
     * Begin a new area
     *
     * @param string|null $cellAddress Upper left cell of area
     *
     * @return Area
     */
    public function beginArea(?string $cellAddress = null): Area
    {
        if (null === $cellAddress) {
            $cellAddress = 'A' . ($this->rowCountWritten + 1);
        }
        $dimension = Excel::rangeDimension($cellAddress, true);
        if ($dimension['rowNum1'] <= $this->rowCountWritten) {
            throw new Exception("Cannot make area from $cellAddress (row number must be greater than written rows)");
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
     * @param bool|null $colOnly
     * @param bool|null $rowOnly
     *
     * @return array|null[]|null
     */
    protected function _parseAddress($cellAddress, ?bool $colOnly = false, ?bool $rowOnly = false): ?array
    {
        $result = ['row' => null, 'col' => null];
        if (is_array($cellAddress)) {
            if (!isset($cellAddress['row'], $cellAddress['col']) && count($cellAddress) > 1) {
                [$col, $row] = array_values($cellAddress);
                $cellAddress = ['row' => $row, 'col' => $col];
            }
            if (isset($cellAddress['row'], $cellAddress['col'])) {
                $result = $cellAddress;
                $result['rowIndex'] = $cellAddress['row'] - 1;
                $result['colIndex'] = $cellAddress['col'] - 1;
                $result['width'] = $result['height'] = $result['cellCount'] = 1;
                $result['rowNum1'] = $result['rowNum2'] = $cellAddress['row'];
                $result['colNum1'] = $result['colNum2'] = $cellAddress['col'];
            }
        }
        elseif (is_string($cellAddress)) {
            if ($rowOnly && preg_match('/^(\d+)(:(\d+))?$/', $cellAddress, $m)) {
                $result['rowNum1'] = (int)$m[1];
                $result['rowNum2'] = !empty($m[3]) ? (int)$m[3] : $result['rowNum1'];
                $result['rowIndex'] = $result['rowNum1'] - 1;
                $result['height'] = $result['rowNum2'] - $result['rowNum1'] + 1;
            }
            else {
                $result = $this->_rangeDimension($cellAddress);
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
        if (preg_match('/^R\[?(-?\d+)?]?C/', $cellAddress)) {
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
     * @param string|array $range1
     * @param string|array $range2
     *
     * @return bool
     */
    protected function _checkIntersection($range1, $range2): bool
    {
        $dim1 = isset($range1['rowNum1'], $range1['colNum1']) ? $range1 : $this->_rangeDimension($range1);
        $dim2 = isset($range2['rowNum1'], $range2['colNum1']) ? $range2 : $this->_rangeDimension($range2);
        if (
            ((($dim1['rowNum1'] >= $dim2['rowNum1']) && ($dim1['rowNum1'] <= $dim2['rowNum2']))
                || (($dim1['rowNum2'] >= $dim2['rowNum1']) && ($dim1['rowNum2'] <= $dim2['rowNum2'])))
            && ((($dim1['colNum1'] >= $dim2['colNum1']) && ($dim1['colNum1'] <= $dim2['colNum2']))
                || (($dim1['colNum2'] >= $dim2['colNum1']) && ($dim1['colNum2'] <= $dim2['colNum2'])))
        ) {
            return true;
        }

        return false;
    }

    /**
     * @param string|array $range
     *
     * @return string
     */
    protected function _fullRangeAddress($range): string
    {
        $absAddress = '';
        if (is_array($range) && isset($range['absAddress'])) {
            $absAddress = $range['absAddress'];
        }
        if (is_string($range)) {
            if (substr_count($range, '$') === 4) {
                $absAddress = $range;
            }
            else {
                $range = $this->_rangeDimension($range);
                $absAddress = $range['absAddress'];
            }
        }

        if ($absAddress) {
            return "'" . $this->sanitizedSheetName . "'!" . $absAddress;
        }

        return '';
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
                if ($row <= $this->rowCountWritten /* $this->currentRowIdx */) {
                    ExceptionAddress::throwNew('Row number must be greater than written rows');
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
            if (is_scalar($value)
                || ($value instanceof RichText)
                // it's a formula & value ['=A1+B2', 123]
                || (is_array($value) && !empty($value[0]) && is_string($value[0]) && ($value[0][0] === '=') && count($value) === 2)
            ) {
                $this->cells['values'][$rowIdx][$colIdx] = $value;
            }
            elseif (is_callable($value)) {
                $this->cells['values'][$rowIdx][$colIdx] = $value($this);
            }
            else {
                $addr = Excel::cellAddress($colIdx + 1, $rowIdx + 1);
                Exception::throwNew('Value for cell %s must be scalar or callable', $addr);
            }
            if ($changeCurrent) {
                $this->currentRowIdx = $rowIdx;
                $this->currentColIdx = $colIdx;
            }
        }
        if ($styles) {
            $this->cells['styles'][$rowIdx][$colIdx] = StyleManager::normalize($styles);
        }

        return $dimension;
    }

    /**
     * Set a value to the single cell or to the cell range
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     *
     * @example
     * $sheet->setValue('B5', $value);
     * $sheet->setValue('B5:C7', $value, $styles);
     * $sheet->setValue(['col' => 2, 'row' => 5], $value, $styles);
     * $sheet->setValue([2, 5], $value);
     */
    public function setValue($cellAddress, $value, ?array $styles = null): Sheet
    {
        $this->writeTo($cellAddress, $value, $styles);
        if ($this->currentColIdx) {
            $this->currentColIdx--;
        }

        return $this;
    }

    /**
     * Set a formula to the single cell or to the cell range
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     *
     * @example
     *  $sheet->setFormula('B5', '=F23');
     *  $sheet->setFormula('B5:C7', $formula, $styles);
     *  $sheet->setFormula(['col' => 2, 'row' => 5], '=R2C3+R3C4');
     *  $sheet->setFormula([2, 5], '=SUM(A4:A18)');
     */
    public function setFormula($cellAddress, $value, ?array $styles = null): Sheet
    {
        if (empty($value)) {
            $value = null;
        }
        elseif (strpos($value, '=') !== 0) {
            $value = '=' . $value;
        }

        $this->_setCellData($cellAddress, $value, $styles, true);

        return $this;
    }

    /**
     * Select a single cell or cell range in the current row
     *
     * @param string|array $cellAddress
     *
     * @return $this
     *
     * @example
     * $sheet->cell('B5')->writeCell($value);
     * $sheet->cell('B5:C7')->applyBorder('thin');
     * $sheet->cell(['col' => 2, 'row' => 5])->applyUnlock();
     * $sheet->cell([2, 5])->applyColor($color);
     */
    public function cell($cellAddress): Sheet
    {
        $dimension = $this->_setCellData($cellAddress, null, null, true);
        $this->currentRowIdx = $dimension['rowIndex'];
        $this->currentColIdx = $dimension['colIndex'];
        $this->_touchEnd($this->currentRowIdx, $this->currentColIdx, 'cell');

        return $this;
    }

    /**
     * Set style for the specific cell
     *
     * @param string $cellAddress Cell address
     * @param mixed $style Style array or object
     * @param bool|null $mergeStyles True - merge style with previous style for this cell (if exists)
     *
     * @return $this
     */
    public function setCellStyle(string $cellAddress, $style, ?bool $mergeStyles = false): Sheet
    {
        $dimension = $this->_rangeDimension($cellAddress);
        if ($dimension['rowNum1'] <= $this->rowCountWritten) {
            throw new Exception('Row number must be greater than written rows');
        }
        $style = StyleManager::normalize($style);
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
        $this->currentRowIdx = $dimension['rowIndex'];
        $this->currentColIdx = $dimension['colIndex'];
        $this->_touchEnd($this->currentRowIdx, $this->currentColIdx, 'cell');
        ++$this->currentColIdx;

        return $this;
    }

    /**
     * Add additional styles to a cell
     *
     * @param string $cellAddr
     * @param array $style
     *
     * @return $this
     */
    public function addCellStyle(string $cellAddr, array $style): Sheet
    {
        return $this->setCellStyle($cellAddr, $style, true);
    }

    /**
     * Alias for 'addCellStyle()'
     *
     * @param string $cellAddr
     * @param array $style
     *
     * @return $this
     */
    public function addStyle(string $cellAddr, array $style): Sheet
    {
        return $this->addCellStyle($cellAddr, $style, true);
    }

    /**
     * Alias for 'setCellStyle()'
     *
     * @param string $cellAddress
     * @param mixed $style
     * @param bool|null $mergeStyles
     *
     * @return $this
     */
    public function setStyle(string $cellAddress, $style, ?bool $mergeStyles = false): Sheet
    {
        return $this->setCellStyle($cellAddress, $style, $mergeStyles);
    }

    /**
     * @param string $cellAddr
     * @param string $color
     *
     * @return $this
     */
    public function setBgColor(string $cellAddr, string $color): Sheet
    {
        return $this->setCellStyle($cellAddr, ['fill-color' => $color], true);
    }

    /**
     * @param string $cellAddr
     * @param string $format
     *
     * @return $this
     */
    public function setFormat(string $cellAddr, string $format): Sheet
    {
        return $this->setCellStyle($cellAddr, ['format' => $format], true);
    }

    /**
     * @param string $range
     * @param string|array $style
     *
     * @return $this
     */
    public function setOuterBorder(string $range, $style): Sheet
    {
        $borderStyle = StyleManager::borderOptions($style);
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
            $maxRowIdx = max(array_keys($this->cells['values']) + array_keys($this->cells['styles']));
            // writes row by row
            for ($rowIdx = $this->rowCountWritten; $rowIdx <= $maxRowIdx; $rowIdx++) {
                if (isset($this->cells['values'][$rowIdx]) || isset($this->cells['styles'][$rowIdx])) {
                    $colMax = 0;
                    $rowValues = $this->cells['values'][$rowIdx] ?? [];
                    if ($rowValues && ($keyMax = max(array_keys($rowValues))) > $colMax) {
                        $colMax = $keyMax;
                    }
                    $cellStyles = $this->cells['styles'][$rowIdx] ?? [];
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
                    //$this->_writeRow($writer, [null]);
                    $this->_writeRow($writer, isset($this->rowSettings[$rowIdx]) ? [null] : []);
                }
            }
            $this->clearAreas();
        }
        $this->currentRowIdx = $this->rowCountWritten;
        $this->currentColIdx = 0;
        $this->_touchEnd($this->currentRowIdx, $this->currentColIdx, 'cell');
        $this->withLastCell();

        return $this;
    }

    /**
     * @param Writer $writer
     *
     * @return void
     */
    public function writeDataBegin(Writer $writer)
    {
        // if already initialized
        if ($this->open) {
            return;
        }

        $sheetFileName = $writer->makeTempFile($this->index . ':sheetData');
        $this->setFileWriter($writer->makeFileWriter($sheetFileName));

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

        if ($this->currentColIdx) {
            $this->_writeCurrentRow();
        }

        if ($this->rowSettings || $this->rowAttributes) {
            $maxRowSettings = max(array_merge(array_keys($this->rowSettings), array_keys($this->rowAttributes)));
            for ($rowIdx = $this->rowCountWritten; $rowIdx <= $maxRowSettings; $rowIdx++) {
                $this->_writeRow($this->excel->getWriter(), [null], [], []);
            }
        }

        for ($level = $this->rowOutlineLevel; $this->rowOutlineLevel > 0; $level--, $this->rowOutlineLevel--) {
            $this->_writeRow($this->excel->getWriter(), [null], ['outlineLevel' => $level], []);
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
     * @param int $row1
     * @param int $col1
     * @param string|null $ref
     *
     * @return void
     */
    protected function _touchStart(int $row1, int $col1, ?string $ref = null)
    {
        $this->_touch($row1, $col1, null, null, $ref);
    }

    /**
     * @param int $row2
     * @param int $col2
     * @param string|null $ref
     *
     * @return void
     */
    protected function _touchEnd(int $row2, int $col2, ?string $ref = null)
    {
        $this->_touch(null, null, $row2, $col2, $ref);
    }

    /**
     * @param int|null $row1
     * @param int|null $col1
     * @param int|null $row2
     * @param int|null $col2
     * @param string|null $ref
     *
     * @return void
     */
    protected function _touch(?int $row1, ?int $col1, ?int $row2, ?int $col2, ?string $ref = null)
    {
        // _touchStart
        if ($row1 !== null) {
            $this->lastTouch['row'] = [
                'row_idx' => $row1,
            ];
            $this->lastTouch['cell'] = [
                'row_idx' => $row1,
                'col_idx' => $col1,
            ];
            $this->lastTouch['area'] = [
                'row_idx1' => $row1,
                'row_idx2' => $row1,
                'col_idx1' => $col1,
                'col_idx2' => $col1,
            ];
            if ($ref) {
                $this->lastTouch['ref'] = $ref;
            }
        }

        // _touchEnd
        if ($row2 !== null) {
            $this->lastTouch['row'] = [
                'row_idx' => $row2,
            ];
            $this->lastTouch['cell'] = [
                'row_idx' => $row2,
                'col_idx' => $col2,
            ];
            $this->lastTouch['area']['row_idx2'] = $row2;
            $this->lastTouch['area']['col_idx2'] = $col2;

            if ($ref) {
                $this->lastTouch['ref'] = $ref;
                if ($ref === 'cell') {
                    $this->lastTouch['area']['row_idx1'] = $row2;
                    $this->lastTouch['area']['col_idx1'] = $col2;
                }
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
        $dimension = $this->_rangeDimension($range);
        if ($dimension['rowNum1'] <= $this->rowCountWritten) {
            throw new Exception('Row number must be greater than written rows');
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
     * @example
     * $sheet->addNamedRange('B3:C5', 'Demo');
     *
     * @param string $range
     * @param string $name
     *
     * @return $this
     */
    public function addNamedRange(string $range, string $name): Sheet
    {
        if ($range) {
            $dimension = $this->_rangeDimension($range);
        }
        else {
            $cell1 = Excel::cellAddress($this->lastTouch['area']['row_idx1'] + 1, $this->lastTouch['area']['col_idx1'] + 1, true);
            $cell2 = Excel::cellAddress($this->lastTouch['area']['row_idx2'] + 1, $this->lastTouch['area']['col_idx2'] + 1, true);
            if ($cell1 === $cell2) {
                $address = $cell1;
            }
            else {
                $address = $cell1 . ':' . $cell2;
            }
            $dimension = [
                'absAddress' => $address,
                'rowNum1' => $this->lastTouch['area']['row_idx1'] + 1,
                'colNum1' => $this->lastTouch['area']['col_idx1'] + 1,
                'rowNum2' => $this->lastTouch['area']['row_idx2'] + 1,
                'colNum2' => $this->lastTouch['area']['col_idx2'] + 1,
            ];

            $dimension['range'] = Excel::cellAddress($dimension['rowNum1'], $dimension['colNum1'])
                . ':' . Excel::cellAddress($dimension['rowNum1'], $dimension['colNum1']);

        }
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

            $this->excel->addDefinedName($name, $this->_fullRangeAddress($dimension));
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

    // === NOTES === //

    /**
     * Add note to the sheet
     *
     * @example
     * $sheet->addNote('A1', $noteText, $noteStyle);
     * $sheet->writeCell($cellValue)->addNote($noteText, $noteStyle);
     *
     * @param string|mixed $cell
     * @param string|array|null $comment
     * @param array $noteStyle
     *
     * @return $this
     */
    public function addNote($cell, $comment = null, array $noteStyle = []): Sheet
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
            $dimension = $this->_rangeDimension($cell);
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
                $noteStyle['fill_color'] = '#' . substr(StyleManager::normalizeColor($noteStyle['fill_color']), 2);
            }
            elseif (!empty($noteStyle['bg_color'])) {
                $noteStyle['fill_color'] = '#' . substr(StyleManager::normalizeColor($noteStyle['bg_color']), 2);
            }
            if (!empty($noteStyle['width']) && (is_int($noteStyle['width']) || is_float($noteStyle['width']))) {
                $noteStyle['width'] = number_format($noteStyle['width'], 2, '.', '') . 'pt';
            }
            if (!empty($noteStyle['height']) && (is_int($noteStyle['height']) || is_float($noteStyle['height']))) {
                $noteStyle['height'] = number_format($noteStyle['height'], 2, '.', '') . 'pt';
            }

            if ($comment instanceof RichText) {
                $text = $comment->outXml();
            }
            else {
                $text = (new RichText(htmlspecialchars($comment)))->outXml();
            }
            $this->notes[$cell] = [
                'cell' => $cell,
                'row_index' => $rowIdx,
                'col_index' => $colIdx,
                'text' => $text,
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
                $this->setBottomNodesOptions('legacyDrawing', ['r:id' => 'rId' . $this->relationshipId]);
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

    // === IMAGES === //

    /**
     * Add image to the sheet from local file, URL or image string in base64
     *
     * @example
     * $sheet->addImage('A1', 'path/to/file');
     * $sheet->addImage('A1', 'path/to/file', ['width => 100]);
     *
     * @param string $cell Cell address
     * @param string $imageFile URL, local path or image string in base64
     * @param array|null $imageStyle ['width' => ..., 'height' => ..., 'hyperlink' => ...]
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
            $dimension = $this->_rangeDimension($cell);
            $cell = $dimension['cell1'];
            $rowIdx = $dimension['rowIndex'];
            $colIdx = $dimension['colIndex'];
        }
        if ($cell) {
            if ($rowIdx >= $this->currentRowIdx && !isset($this->cells['values'][$rowIdx][$colIdx])) {
                $this->cells['values'][$rowIdx][$colIdx] = null;
            }
            $imageData = $this->excel->loadImageFile($imageFile);
            if ($imageData) {
                $imageData['cell'] = $cell;
                $imageData['row_index'] = $rowIdx;
                $imageData['col_index'] = $colIdx;
                $imageData['x'] = $imageStyle['x'] ?? 0;
                $imageData['y'] = $imageStyle['y'] ?? 0;
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
                if (!empty($imageStyle['hyperlink'])) {
                    $imageData['hyperlink'] = ['r_id' => $this->incDrawingRelsId(), 'link' => $imageStyle['hyperlink']];
                }
                $imageData['r_id'] = $this->incDrawingRelsId();
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
                    $this->setBottomNodesOptions('drawing', ['r:id' => 'rId' . $this->relationshipId]);
                }
            }
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getImages(): array
    {

        return $this->media['images'] ?? [];
    }

    /**
     * Add chart object to the specified range of cells
     *
     * @param string $range Set the position where the chart should appear in the worksheet
     * @param Chart $chart Chart object
     *
     * @return $this
     */
    public function addChart(string $range, Chart $chart): Sheet
    {
        $dimension = Excel::rangeDimension($range, true);
        $chart->setTopLeftPosition($dimension['cell1']);
        $chart->setPosition($dimension['cell2']);
        $chart->setSheet($this);

        $chart->rId = $this->incDrawingRelsId();
        $this->charts[] = $chart;
        if (!$chart->getName()) {
            $chart->setName('Chart ' . count($this->charts));
        }

        if (!isset($this->relationships['drawing'])) {
            $file = 'drawing' . $this->index . '.xml';
            $this->relationships['drawing'][++$this->relationshipId] = [
                'file' => $file,
                'link' => '../drawings/' . $file,
                'type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                'extra' => '',
            ];
            $this->setBottomNodesOptions('drawing', ['r:id' => 'rId' . $this->relationshipId]);
        }

        return $this;
    }

    /**
     * @return Chart[]
     */
    public function getCharts(): array
    {
        return $this->charts;
    }

    /**
     * Add data validation object to the specified range of cells
     *
     * @param string $range
     * @param DataValidation $validation
     *
     * @return $this
     */
    public function addDataValidation(string $range, DataValidation $validation): Sheet
    {
        $dimension = Excel::rangeDimension($range, true);
        if ($dimension['cellCount'] === 1) {
            $validation->setSqref($this, $dimension['cell1']);
        }
        else {
            $validation->setSqref($this, $dimension['localRange']);
        }
        $this->_setDimension($dimension['rowNum1'], $dimension['colNum1']);

        $this->validations[$dimension['localRange']] = $validation;

        return $this;
    }

    /**
     * @return DataValidation[]
     */
    public function getDataValidations(): array
    {
        return $this->validations;
    }

    /**
     * Add conditional formatting object to the specified range of cells
     *
     * @param string $range
     * @param Conditional|Conditional[] $conditionals
     *
     * @return $this
     */
    public function addConditionalFormatting(string $range, $conditionals): Sheet
    {
        if (!is_array($conditionals)) {
            $conditionals = [$conditionals];
        }
        foreach ($conditionals as $conditional) {
            $dimension = Excel::rangeDimension($range, true);
            if ($dimension['cellCount'] === 1) {
                $conditional->setSqref($this, $dimension['cell1']);
            }
            else {
                $conditional->setSqref($this, $dimension['localRange']);
            }
            $this->_setDimension($dimension['rowNum1'], $dimension['colNum1']);

            $this->conditionals[] = $conditional;
        }

        return $this;
    }

    /**
     * @return Conditional[]
     */
    public function getConditionalFormatting(): array
    {
        return $this->conditionals;
    }

    // === PROTECTION === //

    /**
     * Protect sheet
     *
     * @param string|null $password
     *
     * @return $this
     */
    public function protect(?string $password = null): Sheet
    {
        $this->protection['sheet'] = 1;
        if ($password) {
            $this->protection['password'] = Excel::hashPassword($password);
        }

        return $this;
    }

    /**
     * AutoFilters should be allowed to operate when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowAutoFilter(?bool $allow = true): Sheet
    {
        $this->protection['autoFilter'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Deleting columns should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowDeleteColumns(?bool $allow = true): Sheet
    {
        $this->protection['deleteColumns'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Deleting rows should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowDeleteRows(?bool $allow = true): Sheet
    {
        $this->protection['deleteRows'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Formatting cells should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowFormatCells(?bool $allow = true): Sheet
    {
        $this->protection['formatCells'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Formatting columns should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowFormatColumns(?bool $allow = true): Sheet
    {
        $this->protection['formatColumns'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Formatting rows should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowFormatRows(?bool $allow = true): Sheet
    {
        $this->protection['formatRows'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Inserting columns should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowInsertColumns(?bool $allow = true): Sheet
    {
        $this->protection['insertColumns'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Inserting hyperlinks should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowInsertHyperlinks(?bool $allow = true): Sheet
    {
        $this->protection['insertHyperlinks'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Inserting rows should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowInsertRows(?bool $allow = true): Sheet
    {
        $this->protection['insertRows'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Objects are allowed to be edited when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowEditObjects(?bool $allow = true): Sheet
    {
        $this->protection['objects'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * PivotTables should be allowed to operate when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowPivotTables(?bool $allow = true): Sheet
    {
        $this->protection['pivotTables'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Scenarios are allowed to be edited when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowEditScenarios(?bool $allow = true): Sheet
    {
        $this->protection['scenarios'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Selection of locked cells should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowSelectLockedCells(?bool $allow = true): Sheet
    {
        $this->protection['selectLockedCells'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Selection of unlocked cells should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowSelectUnlockedCells(?bool $allow = true): Sheet
    {
        $this->protection['selectUnlockedCells'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Selection of any cells should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowSelectCells(?bool $allow = true): Sheet
    {
        $this->protection['selectLockedCells'] = ($allow === false) ? 1 : 0;
        $this->protection['selectUnlockedCells'] = ($allow === false) ? 1 : 0;

        return $this;
    }

    /**
     * Sorting should be allowed when the sheet is protected
     *
     * @param bool|null $allow
     *
     * @return $this
     */
    public function allowSort(?bool $allow = true): Sheet
    {
        $this->protection['sort'] = 0;

        return $this;
    }

    /**
     * Unprotect sheet
     *
     * @return $this
     */
    public function unprotect(): Sheet
    {
        $this->protection['sheet'] = 0;

        return $this;
    }

    /**
     * @param string $key
     * @param string|float|int $value
     *
     * @return void
     */
    protected function _pageMargin(string $key, $value)
    {
        if (preg_match('/^([\d.]+)\s?(mm|cm|in)/', $value, $m)) {
            if ($m[2] === 'cm') {
                $value = (float)$m[1] * 0.39;
            }
            elseif ($m[2] === 'mm') {
                $value = (float)$m[1] * 0.039;
            }
            else {
                $value = (float)$m[1];
            }
        }
        else {
            $value = (float)$value;
        }
        $this->bottomNodesOptions['pageMargins'][$key] = number_format($value, 3, '.', '');
    }

    /**
     * Page margins for a sheet or a custom sheet view in mm|cm|in
     *
     * @param array $margins
     *
     * @return $this
     */
    public function pageMargins(array $margins): Sheet
    {
        foreach ($margins as $key => $val) {
            if (!in_array($key, ['left', 'right', 'top', 'bottom', 'header', 'footer'])) {
                Exception::throwNew('Wrong key for page margin "' . $key . '"');
            }
            $this->_pageMargin($key, $val);
        }

        return $this;
    }

    /**
     * @param array $margins
     *
     * @return $this
     */
    public function setPageMargins(array $margins): Sheet
    {
        return $this->pageMargins($margins);
    }

    /**
     * Left Page Margin in mm|cm|in
     *
     * @param string|float $value
     *
     * @return $this
     */
    public function pageMarginLeft($value): Sheet
    {
        $this->_pageMargin('left', $value);

        return $this;
    }

    /**
     * Right page margin in mm|cm|in
     *
     * @param string|float $value
     *
     * @return $this
     */
    public function pageMarginRight($value): Sheet
    {
        $this->_pageMargin('right', $value);

        return $this;
    }

    /**
     * Top Page Margin in mm|cm|in
     *
     * @param string|float $value
     *
     * @return $this
     */
    public function pageMarginTop($value): Sheet
    {
        $this->_pageMargin('top', $value);

        return $this;
    }

    /**
     * Bottom Page Margin in mm|cm|in
     *
     * @param string|float $value
     *
     * @return $this
     */
    public function pageMarginBottom($value): Sheet
    {
        $this->_pageMargin('bottom', $value);

        return $this;
    }

    /**
     * Header Page Margin in mm|cm|in
     *
     * @param string|float $value
     *
     * @return $this
     */
    public function pageMarginHeader($value): Sheet
    {
        $this->_pageMargin('header', $value);

        return $this;
    }

    /**
     * Footer Page Margin in mm|cm|in
     *
     * @param string|float $value
     *
     * @return $this
     */
    public function pageMarginFooter($value): Sheet
    {
        $this->_pageMargin('footer', $value);

        return $this;
    }

    /**
     * Set Paper size (when paperHeight and paperWidth are specified, paperSize should be ignored)
     *
     * @param int $paperSize
     *
     * @return $this
     */
    public function pagePaperSize(int $paperSize): Sheet
    {
        $this->bottomNodesOptions['pageSetup']['paperSize'] = $paperSize;

        return $this;
    }

    /**
     * Height of custom paper as a number followed by a unit identifier mm|cm|in (ex: 297mm, 11in)
     *
     * @param string|float|int $paperHeight
     *
     * @return $this
     */
    public function pagePaperHeight($paperHeight): Sheet
    {
        if ($paperHeight == (float)$paperHeight) {
            $paperHeight = number_format($paperHeight, 1, '.', '') . 'in';
        }
        $this->bottomNodesOptions['pageSetup']['paperHeight'] = $paperHeight;

        return $this;
    }

    /**
     * Width of custom paper as a number followed by a unit identifier mm|cm|in (ex: 21cm, 8.5in)
     *
     * @param string|float|int $paperWidth
     *
     * @return $this
     */
    public function pagePaperWidth($paperWidth): Sheet
    {
        if ($paperWidth == (float)$paperWidth) {
            $paperWidth = number_format($paperWidth, 1, '.', '') . 'in';
        }
        $this->bottomNodesOptions['pageSetup']['paperWidth'] = $paperWidth;

        return $this;
    }

    /**
     * @param string $range
     *
     * @return $this
     */
    public function setPrintArea(string $range): Sheet
    {
        if (strpos($range, ',')) {
            $ranges = explode(',', $range);
        }
        elseif (strpos($range, ';')) {
            $ranges = explode(';', $range);
        }
        else {
            $ranges = [$range];
        }
        $address = '';
        foreach ($ranges as $range) {
            $dimension = $this->_rangeDimension($range);
            // checking intersections
            foreach ($this->printAreas as $printArea) {
                if ($this->_checkIntersection($dimension, $printArea)) {
                    throw new Exception('Print areas should not overlap (' . $printArea['localRange'] . ' & ' . $dimension['localRange'] . ')');
                }
            }
            $this->printAreas[] = $dimension;
            if ($address) {
                $address .= ',';
            }
            $address .= $this->_fullRangeAddress($dimension);
        }
        $this->excel->addDefinedName('_xlnm.Print_Area', $address, ['localSheetId' => $this->localSheetId]);

        return $this;
    }

    /**
     * @param string|null $rowsAtTop
     * @param string|null $colsAtLeft
     *
     * @return $this
     */
    public function setPrintTitles(?string $rowsAtTop, ?string $colsAtLeft = null): Sheet
    {
        $rowsTitle = $colsTitle = null;
        if ($rowsAtTop && preg_match('/(\d+)(:(\d+))?/', $rowsAtTop, $m)) {
            $rowsTitle = "'" . $this->sanitizedSheetName . "'!" . (empty($m[3]) ? '$' . $m[1] . ':$' . $m[1] : '$' . $m[1] . ':$' . $m[3]);
        }
        if ($colsAtLeft && preg_match('/([A-Z]+)(:([A-Z]+))?/', strtoupper($colsAtLeft), $m)) {
            $colsTitle = "'" . $this->sanitizedSheetName . "'!" . (empty($m[3]) ? '$' . $m[1] . ':$' . $m[1] : '$' . $m[1] . ':$' . $m[3]);
        }
        if ($rowsTitle || $colsTitle) {
            $address = '';
            if ($colsTitle) {
                $address = $colsTitle;
            }
            if ($rowsTitle) {
                $address .= ($address ? ',' : '') . $rowsTitle;
            }
            $this->excel->addDefinedName('_xlnm.Print_Titles', $address, ['localSheetId' => $this->localSheetId]);
        }

        return $this;
    }

    /**
     * @param string $range
     *
     * @return $this
     */
    public function setPrintTopRows(string $range): Sheet
    {
        $this->printTopRows = $range;

        return $this->setPrintTitles($this->printTopRows, $this->printLeftColumns);
    }

    /**
     * @param string $range
     *
     * @return $this
     */
    public function setPrintLeftColumns(string $range): Sheet
    {
        $this->printLeftColumns = $range;

        return $this->setPrintTitles($this->printTopRows, $this->printLeftColumns);
    }

    /**
     * Show grid line in the print area
     *
     * @param bool|null $bool
     *
     * @return $this
     */
    public function setPrintGridlines(?bool $bool = true): Sheet
    {
        $this->setBottomNodeOption('printOptions', 'gridLines', $bool ? '1' : '0');

        return $this;
    }

    /**
     * Center the print area horizontally
     *
     * @param bool|null $value
     *
     * @return $this
     */
    public function setPrintHorizontalCentered(?bool $value = true): Sheet
    {
        $this->setBottomNodeOption('printOptions', 'horizontalCentered', $value ? '1' : '0');

        return $this;
    }

    /**
     * Center the print area vertically
     *
     * @param bool|null $value
     *
     * @return $this
     */
    public function setPrintVerticalCentered(?bool $value = true): Sheet
    {
        $this->setBottomNodeOption('printOptions', 'verticalCentered', $value ? '1' : '0');

        return $this;
    }

    /**
     * Center the print area horizontally and vertically
     *
     * @param bool|null $value
     *
     * @return $this
     */
    public function setPrintCentered(?bool $value = true): Sheet
    {
        $this->setPrintHorizontalCentered($value)
            ->setPrintVerticalCentered($value);

        return $this;
    }

    /**
     * @return array|array[]
     */
    public function getSheetViews(): array
    {
        $result = [];
        $paneRow = ($this->freezeRows ? $this->freezeRows + 1 : 0);
        $paneCol = ($this->freezeColumns ? $this->freezeColumns + 1 : 0);
        foreach ($this->sheetViews as $n => $sheetView) {
            $result[$n] = [
                '_attr' => $sheetView,
            ];
            if ($this->isRightToLeft()) {
                $result[$n]['_attr']['rightToLeft'] = 'true';
            }
            if ($this->active) {
                $result[$n]['_attr']['tabSelected'] = 'true';
            }
            if ($this->freezeRows && $this->freezeColumns) {
                // frozen rows and cols
                $activeCell = $this->activeCell ?? Excel::cellAddress($paneRow, $paneCol);
                $activeRef = $this->activeRef ?? $activeCell;
                $result[$n]['_items'] = [
                    [
                        '_tag' => 'pane',
                        '_attr' => ['ySplit' => $this->freezeRows, 'xSplit' => $this->freezeColumns, 'topLeftCell' => Excel::cellAddress($paneRow, $paneCol), 'activePane' => 'bottomRight', 'state' => 'frozen'],
                    ],
                    [
                        '_tag' => 'selection',
                        '_attr' => ['pane' => 'topRight', 'activeCell' => Excel::cellAddress($paneRow, 1), 'sqref' => Excel::cellAddress($paneRow, 1)],
                    ],
                    [
                        '_tag' => 'selection',
                        '_attr' => ['pane' => 'bottomLeft', 'activeCell' => Excel::cellAddress(1, $paneCol), 'sqref' => Excel::cellAddress(1, $paneCol)],
                    ],
                    [
                        '_tag' => 'selection',
                        '_attr' => ['pane' => 'bottomRight', 'activeCell' => $activeCell, 'sqref' => $activeRef],
                    ],
                ];
            }
            elseif ($this->freezeRows) {
                // frozen rows only
                $activeCell = $this->activeCell ?? Excel::cellAddress($paneRow, 1);
                $activeRef = $this->activeRef ?? $activeCell;
                $result[$n]['_items'] = [
                    [
                        '_tag' => 'pane',
                        '_attr' => ['ySplit' => $this->freezeRows, 'topLeftCell' => Excel::cellAddress($paneRow, 1), 'activePane' => 'bottomRight', 'state' => 'frozen'],
                    ],
                    [
                        '_tag' => 'selection',
                        '_attr' => ['pane' => 'bottomLeft', 'activeCell' => $activeCell, 'sqref' => $activeRef],
                    ],
                ];
            }
            elseif ($this->freezeColumns) {
                // frozen cols only
                $activeCell = $this->activeCell ?? Excel::cellAddress(1, $paneCol);
                $activeRef = $this->activeRef ?? $activeCell;
                $result[$n]['_items'] = [
                    [
                        '_tag' => 'pane',
                        '_attr' => ['xSplit' => $this->freezeColumns, 'topLeftCell' => Excel::cellAddress(1, $paneCol), 'activePane' => 'topRight', 'state' => 'frozen'],
                    ],
                    [
                        '_tag' => 'selection',
                        '_attr' => ['pane' => 'topRight', 'activeCell' => $activeCell, 'sqref' => $activeRef],
                    ],
                ];
            }
            else {
                // not frozen
                $activeCell = $this->activeCell ?? $this->minCell();
                $activeRef = $this->activeRef ?? $activeCell;
                $result[$n]['_items'] = [
                    [
                        '_tag' => 'selection',
                        '_attr' => ['pane' => 'topLeft', 'activeCell' => $activeCell, 'sqref' => $activeRef],
                    ],
                ];
            }
        }

        return $result;
    }

    public function getSheetFormatPr(): array
    {
        return $this->sheetFormatPr;
    }

    /**
     * @return array
     */
    public function getProtection(): array
    {

        return $this->protection;
    }

    /**
     * @return array
     */
    public function getPageMargins(): array
    {

        return $this->bottomNodesOptions['pageMargins'] ?? [];
    }

    public function getBottomNodesOptions(): array
    {
        // need specified order for some nodes
        $order = [
            'printOptions',
            'pageMargins',
            'pageSetup',
            'drawing',
            'legacyDrawing',
        ];
        $result = [];
        foreach ($order as $key) {
            if (isset($this->bottomNodesOptions[$key])) {
                $result[$key] = $this->bottomNodesOptions[$key];
            }
        }
        foreach ($this->bottomNodesOptions as $key => $value) {
            if (!in_array($key, $order)) {
                $result[$key] = $value;
            }
        }

        return $result;
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
        $this->setRowHeight($this->currentRowIdx + 1, $height);

        return $this;
    }

    /**
     * @param int $outlineLevel
     *
     * @return $this
     */
    public function applyRowOutlineLevel(int $outlineLevel): Sheet
    {
        $this->setRowOutlineLevel($this->currentRowIdx + 1, $outlineLevel);

        return $this;
    }

    /**
     * @param array $style
     *
     * @return $this
     */
    public function applyStyle(array $style): Sheet
    {
        $style = StyleManager::normalize($style);
        foreach ($style as $key => $options) {
            $this->_setStyleOptions([], $key, $options, true);
        }

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

            // top-left border
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

            // top-right border
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

            // bottom-right border
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

            // bottom-left border
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
            $font['font-size'] = $fontSize;
        }
        if ($fontStyle) {
            $font['font-style'] = $fontStyle;
        }
        if ($fontColor) {
            $font['font-color'] = $fontColor;
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

    /**
     * @return $this
     */
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
        $this->_setStyleOptions([], 'font', [Style::FONT_COLOR => $fontColor]);

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
     * Fill background color
     *
     * @param string $color
     * @param string|null $pattern
     *
     * @return $this
     */
    public function applyFillColor(string $color, ?string $pattern = null): Sheet
    {
        $this->_setStyleOptions([], 'fill', ['fill-color' => $color, 'fill-pattern' => $pattern ?: 'solid']);

        return $this;
    }

    /**
     * Alias of 'applyFillColor()'
     *
     * @param string $color
     * @param string|null $pattern
     *
     * @return $this
     */
    public function applyBgColor(string $color, ?string $pattern = null): Sheet
    {
        return $this->applyFillColor($color, $pattern);
    }

    /**
     * Fill background by gradient
     *
     * @param string $color1
     * @param string $color2
     * @param int|null $degree
     * @return $this
     */
    public function applyFillGradient(string $color1, string $color2, ?int $degree = null): Sheet
    {
        $this->_setStyleOptions([], 'fill', [
            'fill-pattern' => Style::FILL_GRADIENT_LINEAR,
            'fill-gradient-start' => $color1,
            'fill-gradient-end' => $color2,
            'fill-gradient-degree' => $degree ?: 0,
        ]);

        return $this;
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
     * Apply left alignment to content
     *
     * @return $this
     */
    public function applyAlignLeft(): Sheet
    {
        return $this->applyTextAlign('left');
    }

    /**
     * Apply right alignment to content
     *
     * @return $this
     */
    public function applyAlignRight(): Sheet
    {
        return $this->applyTextAlign('right');
    }

    /**
     * @param bool|null $textWrap
     *
     * @return $this
     */
    public function applyTextWrap(?bool $textWrap = true): Sheet
    {
        $this->_setStyleOptions([], 'format', ['format-text-wrap' => (int)$textWrap]);

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

    /**
     * @param int $degrees
     *
     * @return $this
     */
    public function applyTextRotation(int $degrees): Sheet
    {
        $this->_setStyleOptions([], 'format', [ 'format-text-rotation' => $degrees ] );

        return $this;
    }

    /**
     * @param int $indent
     *
     * @return $this
     */
    public function applyIndentLeft(int $indent): Sheet
    {
        $options = ['format-align-horizontal' => 'left', 'format-align-indent' => $indent];
        $this->_setStyleOptions([], 'format', $options);

        return $this;
    }

    /**
     * @param int $indent
     *
     * @return $this
     */
    public function applyIndentRight(int $indent): Sheet
    {
        $options = ['format-align-horizontal' => 'right', 'format-align-indent' => $indent];
        $this->_setStyleOptions([], 'format', $options);

        return $this;
    }

    /**
     * @param int $indent
     *
     * @return $this
     */
    public function applyIndentDistributed(int $indent): Sheet
    {
        $options = ['format-align-horizontal' => 'distributed', 'format-align-indent' => $indent];
        $this->_setStyleOptions([], 'format', $options);

        return $this;
    }

    /**
     * @param string|array $format
     *
     * @return $this
     */
    public function applyFormat($format): Sheet
    {
        if (is_array($format)) {
            $this->_setStyleOptions([], 'format', $format);
        }
        else {
            if ($format && $format[0] === '@') {
                $format = strtoupper($format);
            }
            $this->_setStyleOptions([], 'format', ['format-pattern' => $format]);
        }

        return $this;
    }

    /**
     * @param bool $unlock
     *
     * @return $this
     */
    public function applyUnlock(?bool $unlock = true): Sheet
    {
        $this->_setStyleOptions([], 'protection', ['protection-locked' => ($unlock === false) ? 1 : 0]);

        return $this;
    }

    /**
     * @return $this
     */
    public function applyHide(?bool $hide = true): Sheet
    {
        $this->_setStyleOptions([], 'protection', ['protection-hidden' => ($hide === false) ? 0 : 1]);

        return $this;
    }

    /**
     * @param string $name
     *
     * @return $this
     */
    public function applyNamedRange(string $name): Sheet
    {
        $this->addNamedRange('', $name);

        return $this;
    }

    /**
     * @param DataValidation $validation
     *
     * @return $this
     */
    public function applyDataValidation(DataValidation $validation): Sheet
    {
        $dataValidation = clone $validation;
        $address = Helper::cellAddress($this->lastTouch['cell']['row_idx'] + 1, $this->lastTouch['cell']['col_idx'] + 1);
        $this->addDataValidation($address, $dataValidation);

        return $this;
    }

    /**
     * @param Conditional|Conditional[] $conditionals
     *
     * @return $this
     */
    public function applyConditionalFormatting($conditionals): Sheet
    {
        if (is_array($conditionals)) {
            foreach ($conditionals as $conditional) {
                $this->applyConditionalFormatting($conditional);
            }
        }
        else {
            $conditional = clone $conditionals;
            $address = Helper::cellAddress($this->lastTouch['area']['row_idx1'] + 1, $this->lastTouch['area']['col_idx1'] + 1)
                . ':' . Helper::cellAddress($this->lastTouch['area']['row_idx2'] + 1, $this->lastTouch['area']['col_idx2'] + 1);
            $this->addConditionalFormatting($address, $conditional);
        }

        return $this;
    }

    /**
     * @param bool|null $absolute
     *
     * @return string
     */
    public function getLastRange(?bool $absolute = false): string
    {
        $rowIdx = $this->lastTouch['cell']['row_idx'];
        $colIdx = $this->lastTouch['cell']['col_idx'];
        $ref = $this->lastTouch['ref'];

        if ($ref === 'cell') {
            $addr = Excel::cellAddress($rowIdx + 1, $colIdx + 1, $absolute);

            return $addr . ':' . $addr;
        }
        if ($ref === 'area') {
            $addr1 = Excel::cellAddress($this->lastTouch['area']['row_idx1'] + 1, $this->lastTouch['area']['col_idx1'] + 1, $absolute);
            $addr2 = Excel::cellAddress($this->lastTouch['area']['row_idx2'] + 1, $this->lastTouch['area']['col_idx2'] + 1, $absolute);

            return $addr1 . ':' . $addr2;
        }
        // row
        if ($absolute) {
            return '$' . ($this->lastTouch['row']['row_idx'] + 1) . ':$' . ($this->lastTouch['row']['row_idx'] + 1);
        }

        return '' . ($this->lastTouch['row']['row_idx'] + 1) . ':' . ($this->lastTouch['row']['row_idx'] + 1);
    }

    /**
     * @param bool|null $absolute
     *
     * @return string
     */
    public function getLastCell(?bool $absolute = false): string
    {

        return Excel::cellAddress($this->lastTouch['area']['row_idx2'] + 1, $this->lastTouch['area']['col_idx2'] + 1, $absolute);
    }
}

// EOF
