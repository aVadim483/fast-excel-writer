<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;
use avadim\FastExcelWriter\Exception\SaveException;

/**
 * Class Writer
 *
 * @package avadim\FastExcelWriter
 */
class Writer
{
    protected $excel;

    /** @var array  */
    protected $tempFiles = [];

    /** @var string  */
    protected $tempDir = '';

    /** @var array  */
    protected $cellStyles = [];

    /** @var array  */
    protected $numberFormats = [];

    protected static $localeSettings = [];

    /**
     * Writer constructor
     *
     * @param array $options;
     */
    public function __construct($options = [])
    {
        date_default_timezone_get() || date_default_timezone_set('UTC');//php.ini missing tz, avoid warning
        if (isset($options['excel'])) {
            $this->excel = $options['excel'];
        }
        if (isset($options['temp_dir'])) {
            $this->tempDir = $options['temp_dir'];
        }
        if (!is_writable($this->tempFilename())) {
            throw new Exception('Warning: tempdir ' . sys_get_temp_dir() . ' is not writeable, use ->setTempDir()');
        }
        if (!class_exists('\ZipArchive')) {
            throw new Exception('Error: ZipArchive class does not exist');
        }
        $this->addCellStyle('GENERAL', null);
    }

    /**
     * @param string $fileName
     *
     * @return WriterBuffer
     */
    public static function makeWriteBuffer($fileName)
    {
        return new WriterBuffer($fileName);
    }

    /**
     * @param $localeData
     */
    public static function setLocaleSettings($localeData)
    {
        if (!empty($localeData['functions'])) {
            uksort($localeData['functions'], static function($a, $b) {
                return mb_strlen($b) - mb_strlen($a);
            });
        }
        if (!empty($localeData['formats'])) {
            uksort($localeData['formats'], static function($a, $b) {
                return mb_strlen($b) - mb_strlen($a);
            });
        }
        self::$localeSettings = $localeData;
    }

    /**
     * @param Excel $excel
     */
    public function setExcel($excel)
    {
        $this->excel = $excel;
    }

    /**
     * @param string $tempDir
     */
    public function setTempDir($tempDir = '')
    {
        $this->tempDir = $tempDir;
    }

    /**
     *
     */
    public function __destruct()
    {
        if (!empty($this->tempFiles)) {
            foreach ($this->tempFiles as $tempFile) {
                @unlink($tempFile);
            }
        }
    }

    /**
     * @return bool|string
     */
    public function tempFilename()
    {
        $tempPrefix = 'xlsx_writer_';
        if (!$this->tempDir) {
            $tempDir = sys_get_temp_dir();
            $filename = tempnam($tempDir, $tempPrefix);
            if (!$filename) {
                $filename = tempnam(getcwd(), $tempPrefix);
            }
        } else {
            $filename = tempnam($this->tempDir, $tempPrefix);
        }
        if ($filename) {
            $this->tempFiles[] = $filename;
        }

        return $filename;
    }

    /**
     * @param array $array
     */
    protected function _ksort(&$array)
    {
        ksort($array);
        foreach($array as $key => $val) {
            if (is_array($val)) {
                $this->_ksort($val);
                $array[$key] = $val;
            }
        }
    }

    /**
     * @param $numberFormat
     * @param $cellStyle
     *
     * @return false|int|string
     */
    public function addCellStyle($numberFormat, $cellStyle)
    {
        if (empty($cellStyle)) {
            $cellStyleString = 'null';
        } elseif (is_array($cellStyle)) {
            if (!empty($cellStyle['border'])) {
                $border = Style::normalizeBorder($cellStyle['border']);
                if ($border) {
                    $cellStyle['border'] = $border;
                } else {
                    unset($cellStyle['border']);
                }
            }
            $this->_ksort($cellStyle);
            $cellStyleString = json_encode($cellStyle);
        } elseif (!is_string($cellStyle)) {
            $cellStyleString = json_encode($cellStyle);
        } else {
            $cellStyleString = (string)$cellStyle;
        }
        $numberFormatIdx = self::addToListGetIndex($this->numberFormats, $numberFormat);
        $lookupString = $numberFormatIdx . ";" . $cellStyleString;

        return self::addToListGetIndex($this->cellStyles, $lookupString);
    }

    /**
     * @param $format
     *
     * @return array
     */
    public function defineFormatType($format)
    {
        if (is_array($format)) {
            $format = reset($format);
        }
        $numberFormat = self::numberFormatStandardized($format);
        $numberFormatType = self::determineNumberFormatType($numberFormat);
        $cellStyleIdx = $this->addCellStyle($numberFormat, null);
        $formatType = [
            'number_format' => $numberFormat, //contains excel format like 'YYYY-MM-DD HH:MM:SS'
            'number_format_type' => $numberFormatType, //contains friendly format like 'datetime'
            'default_style_idx' => $cellStyleIdx,
        ];

        return $formatType;
    }

    /**
     * @return array
     */
    public function defaultFormatType()
    {
        static $defaultFormatType;

        if (!$defaultFormatType) {
            $defaultFormatType = $this->defineFormatType('GENERAL');
        }
        return $defaultFormatType;
    }

    /**
     * @param $fileName
     * @param $overWrite
     * @param $metadata
     *
     * @return bool
     */
    public function saveToFile($fileName, $overWrite = true, $metadata = [])
    {
        $sheets = $this->excel->getSheets();
        foreach ($sheets as $sheetName => $sheet) {
            if (!$sheet->open) {
                $this->writeSheetDataBegin($sheet);
            }
            $this->writeSheetDataEnd($sheet);//making sure all footers have been written
        }

        if (!is_dir(dirname($fileName))) {
            throw new SaveException('Directory "' . dirname($fileName) . '" for output file is not exist.');
        }
        if (file_exists($fileName)) {
            if ($overWrite && is_writable($fileName)) {
                @unlink($fileName); //if the zip already exists, remove it
            } else {
                throw new SaveException('File "' . $fileName. '" is not writeable');
            }
        }
        $zip = new \ZipArchive();
        if (empty($sheets)) {
            throw new SaveException('No worksheets defined');
        }
        if (!$zip->open($fileName, \ZIPARCHIVE::CREATE)) {
            throw new SaveException('Unable to create zip "' . $fileName. '"');
        }

        $zip->addEmptyDir('docProps/');
        $zip->addFromString('docProps/app.xml', $this->_buildAppXML($metadata));
        $zip->addFromString('docProps/core.xml', $this->_buildCoreXML($metadata));

        $zip->addEmptyDir('_rels/');
        $zip->addFromString('_rels/.rels', $this->_buildRelationshipsXML());

        $zip->addEmptyDir('xl/worksheets/');
        foreach ($sheets as $sheet) {
            $zip->addFile($sheet->fileName, 'xl/worksheets/' . $sheet->xmlName);
        }
        $zip->addFromString('xl/workbook.xml', $this->_buildWorkbookXML($sheets));
        $zip->addFile($this->_writeStylesXML(), 'xl/styles.xml');  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
        $zip->addFromString('[Content_Types].xml', $this->_buildContentTypesXML($sheets));

        $zip->addEmptyDir('xl/_rels/');
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->_buildWorkbookRelsXML($sheets));
        $zip->close();

        return true;
    }

    /**
     * @param Sheet $sheet
     *
     * @return WriterBuffer
     */
    protected function _writeSheetHead($sheet)
    {
        $fileWriter = self::makeWriteBuffer($this->tempFilename());

        $fileWriter->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $fileWriter->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');

        $fileWriter->write('<sheetPr>');
        if ($sheet->getPageFit()) {
            $fileWriter->write('<pageSetUpPr fitToPage="1"/>');
        } else {
            $fileWriter->write('<pageSetUpPr fitToPage="false"/>');
        }
        $fileWriter->write('</sheetPr>');

        if ($sheet->rowCount + $sheet->colCount === 0) {
            $fileWriter->write('<dimension ref="A1"/>');
        } else {
            $maxCell = $sheet->maxCell();
            $fileWriter->write('<dimension ref="A1:' . $maxCell . '"/>');
        }

        $rightToLeftValue = $sheet->isRightToLeft() ? 'true' : 'false';

        $fileWriter->write('<sheetViews>');
        $tabSelected = ($sheet->active ? 'tabSelected="true"' : '');
        $fileWriter->write('<sheetView colorId="64" defaultGridColor="true" rightToLeft="' . $rightToLeftValue . '" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" ' . $tabSelected . ' topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');

        $paneRow = ($sheet->freezeRows ? $sheet->freezeRows + 1: 0);
        $paneCol = ($sheet->freezeColumns ? $sheet->freezeColumns + 1 : 0);
        if ($sheet->freezeRows && $sheet->freezeColumns) {
            // frozen rows and cols
            $fileWriter->write('<pane ySplit="' . $sheet->freezeRows . '" xSplit="' . $sheet->freezeColumns . '" topLeftCell="' . Excel::cellAddress($paneRow, $paneCol) . '" activePane="bottomRight" state="frozen"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress($paneRow, 1) . '" activeCellId="0" pane="topRight" sqref="' . Excel::cellAddress($paneRow, 1) . '"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress(1, $paneCol) . '" activeCellId="0" pane="bottomLeft" sqref="' . Excel::cellAddress(1, $paneCol) . '"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress($paneRow, $paneCol) . '" activeCellId="0" pane="bottomRight" sqref="' . Excel::cellAddress($paneRow, $paneCol) . '"/>');
        } elseif ($sheet->freezeRows) {
            // frozen rows only
            $fileWriter->write('<pane ySplit="' . $sheet->freezeRows . '" topLeftCell="' . Excel::cellAddress($paneRow, 1) . '" activePane="bottomLeft" state="frozen"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress($paneRow, 1) . '" activeCellId="0" pane="bottomLeft" sqref="' . Excel::cellAddress($paneRow, 1) . '"/>');
        } elseif ($sheet->freezeColumns) {
            // frozen cols only
            $fileWriter->write('<pane xSplit="' . $sheet->freezeColumns . '" topLeftCell="' . Excel::cellAddress(1, $paneCol) . '" activePane="topRight" state="frozen"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress(1, $paneCol) . '" activeCellId="0" pane="topRight" sqref="' . Excel::cellAddress(1, $paneCol) . '"/>');
        } else {
            // not frozen
            $fileWriter->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
        }
        $fileWriter->write('</sheetView>');
        $fileWriter->write('</sheetViews>');

        if (!empty($sheet->colWidths)) {
            $fileWriter->write('<cols>');
            foreach ($sheet->colWidths as $colNum => $columnWidth) {
                $fileWriter->write('<col min="' . ($colNum + 1) . '" max="' . ($colNum + 1) . '" width="' . $columnWidth . '" customWidth="1"/>');
            }
            $fileWriter->write('</cols>');
        }
        //$fileWriter->write('<col collapsed="false" hidden="false" max="1024" min="' . ($i + 1) . '" style="0" customWidth="false" width="11.5"/>');

        return $fileWriter;
    }

    /**
     * @param Sheet $sheet
     */
    public function writeSheetDataBegin($sheet)
    {
        //if already initialized
        if ($sheet->open) {
            return;
        }

        $sheetFileName = $this->tempFilename();
        $sheet->setFileWriter(self::makeWriteBuffer($sheetFileName));

        $sheet->fileWriter->write('<sheetData>');

        $sheet->open = true;
        $sheet->writeAreasRows($this);

        if ($sheet->colFormats) {
            foreach($sheet->colFormats as $colNum => $format) {
                $colIndex = $colNum + 1;
                if (!isset($sheet->columns[$colIndex])) {
                    $sheet->columns[$colIndex] = $this->defineFormatType($format);
                }

            }
        }
    }

    /**
     * @param Sheet $sheet
     */
    public function writeSheetDataEnd($sheet)
    {
        if ($sheet->close) {
            return;
        }
        $sheet->writeAreas();
        $sheet->fileWriter->flush(true);
        $sheet->fileWriter->write('</sheetData>');

        $mergedCells = $sheet->getMergedCells();
        if ($mergedCells) {
            $sheet->fileWriter->write('<mergeCells>');
            foreach ($mergedCells as $range) {
                $sheet->fileWriter->write('<mergeCell ref="' . $range . '"/>');
            }
            $sheet->fileWriter->write('</mergeCells>');
        }

        if ($sheet->autoFilter) {
            $minCell = $sheet->autoFilter;
            $maxCell = Excel::cellAddress($sheet->rowCount, $sheet->colCount);
            $sheet->fileWriter->write('<autoFilter ref="' . $minCell . ':' . $maxCell . '"/>');
        }

        $pageSetupAttr = 'orientation="' . $sheet->getPageOrientation() . '"';
        if ($sheet->getPageFit()) {
            $pageFitToWidth = $sheet->getPageFitToWidth();
            $pageFitToHeight = $sheet->getPageFitToHeight();
            if ($pageFitToWidth === 1) {
                $pageSetupAttr .= ' fitToHeight="' . $pageFitToHeight . '" ';
            } else {
                $pageSetupAttr .= ' fitToHeight="' . $pageFitToHeight . '" fitToWidth="' . $pageFitToWidth . '"';
            }
        }
        $sheet->fileWriter->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
        $sheet->fileWriter->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');

        $sheet->fileWriter->write("<pageSetup  paperSize=\"1\" useFirstPageNumber=\"1\" horizontalDpi=\"0\" verticalDpi=\"0\" $pageSetupAttr r:id=\"rId1\"/>'");

        $sheet->fileWriter->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $sheet->fileWriter->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
        $sheet->fileWriter->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
        $sheet->fileWriter->write('</headerFooter>');
        $sheet->fileWriter->write('</worksheet>');
        $sheet->fileWriter->flush(true);

        $headWriter = $this->_writeSheetHead($sheet);
        $headWriter->appendFileWriter($sheet->fileWriter, $this->tempFilename());;

        $sheet->fileWriter->close();
        $sheet->close = true;

        $sheet->resetFileWriter($headWriter);
    }

    /**
     * @param $formula
     * @param $baseAddress
     *
     * @return string
     */
    protected function _convertFormula($formula, $baseAddress)
    {
        static $functionNames = [];

        $mark = md5(microtime());
        $replace = [];
        // temporary replace strings
        if (strpos($formula, '"') !== false) {
            $replace = [[], []];
            $formula = preg_replace_callback('/"[^"]+"/', static function($matches) use ($mark, &$replace) {
                $key = '<<' . $mark . '-' . md5($matches[0]) . '>>';
                $replace[0][] = $key;
                $replace[1][] = $matches[0];
                return $key;
            }, $formula);
        }
        // change relative addresses
        $formula = preg_replace_callback('/(\W)(R\[?(-?\d+)?\]?C\[?(-?\d+)?\]?)/', static function($matches) use ($baseAddress) {
            $indexes = Excel::rangeRelOffsets($matches[2]);
            if (isset($indexes[0], $indexes[1])) {
                $row = $baseAddress[0] + $indexes[0];
                $col = $baseAddress[1] + $indexes[1];
                $cell = Excel::cellAddress($row, $col);
                if ($cell) {
                    return $matches[1] . $cell;
                }
            }
            return $matches[0];
        }, $formula);

        if (!empty(self::$localeSettings['functions']) && strpos($formula, '(')) {
            // replace national function names
            if (empty($functionNames)) {
                $functionNames = [[], []];
                foreach(self::$localeSettings['functions'] as $name => $nameEn) {
                    $functionNames[0][] = $name . '(';
                    $functionNames[1][]= $nameEn . '(';
                }
            }
            $formula = str_replace($functionNames[0], $functionNames[1], $formula);
        }

        if ($replace && !empty($replace[0])) {
            // restore strings
            $formula = str_replace($replace[0], $replace[1], $formula);
        }

        return $formula;
    }

    /**
     * @param WriterBuffer $file
     * @param              $rowNumber
     * @param              $colNumber
     * @param              $value
     * @param              $numFormatType
     * @param              $cellStyleIdx
     */
    public function writeCell(WriterBuffer $file, $rowNumber, $colNumber, $value, $numFormatType, $cellStyleIdx)
    {
        $cellName = Excel::cellAddress($rowNumber, $colNumber);

        if (!is_scalar($value) || $value === '') { //objects, array, empty; null is not scalar
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '"/>');
        } elseif (is_string($value) && $value[0] === '=') {
            // formula
            $value = $this->_convertFormula($value, [$rowNumber, $colNumber]);
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="s"><f>' . self::xmlSpecialChars($value) . '</f></c>');
        } elseif ($numFormatType === 'n_string' || ($numFormatType === 'n_numeric' && !is_numeric($value))) {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="inlineStr"><is><t xml:space="preserve">' . self::xmlSpecialChars($value) . '</t></is></c>');
        } else {
            if ($numFormatType === 'n_date' || $numFormatType === 'n_datetime') {
                $dateValue = self::convertDateTime($value);
                if ($dateValue === false) {
                    $numFormatType = 'n_auto';
                } else {
                    $value = $dateValue;
                }
            }
            if ($numFormatType === 'n_date') {
                //$file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . (int)self::convertDateTime($value) . '</v></c>');
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '"><v>' . (int)$value . '</v></c>');
            } elseif ($numFormatType === 'n_datetime') {
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . $value . '</v></c>');
            } elseif ($numFormatType === 'n_numeric') {
                //$file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::xmlSpecialChars($value) . '</v></c>');//int,float,currency
                if (!is_int($value) && !is_float($value)) {
                    $value = self::xmlSpecialChars($value);
                }
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" ><v>' . $value . '</v></c>');//int,float,currency
            } elseif ($numFormatType === 'n_auto' || 1) { //auto-detect unknown column types
                if (!is_string($value) || $value === '0' || ($value[0] !== '0' && ctype_digit($value)) || preg_match("/^-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)) {
                    //$file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::xmlSpecialChars($value) . '</v></c>');//int,float,currency
                    $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . $value . '</v></c>');//int,float,currency
                } else {
                    //implied: ($cellFormat=='string')
                    if (strpos($value, '\=') === 0 || strpos($value, '\\\\=') === 0) {
                        $value = substr($value, 1);
                    }
                    $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="inlineStr"><is><t xml:space="preserve">' . self::xmlSpecialChars($value) . '</t></is></c>');
                }
            }
        }
    }

    /**
     * @return array
     */
    protected function _styleFontIndexes()
    {
        $fills = ['', ''];//2 placeholders for static xml later
        $fonts = ['', '', '', ''];//4 placeholders for static xml later
        $borders = [''];//1 placeholder for static xml later
        $styleIndexes = [];
        foreach ($this->cellStyles as $i => $cellStyleString) {
            $semiColonPos = strpos($cellStyleString, ";");
            $numberFormatIdx = substr($cellStyleString, 0, $semiColonPos);
            $styleJsonString = substr($cellStyleString, $semiColonPos + 1);
            $style = json_decode($styleJsonString, true);

            $styleIndexes[$i] = ['num_fmt_idx' => $numberFormatIdx];//initialize entry

            // new border settings
            if (!empty($style['border']) && is_array($style['border'])) {
                $borderValue = [];
                foreach($style['border'] as $side => $options) {
                    $borderValue[$side] = $options;
                    if (!empty($options['color'])) {
                        $color = Style::normaliazeColor($options['color']);
                        if ($color) {
                            $borderValue[$side]['color'] = $color;
                        }
                    }
                }
                $styleIndexes[$i]['border_idx'] = self::addToListGetIndex($borders, $borderValue);
            }
            if (!empty($style['fill'])) {
                $color = Style::normaliazeColor($style['fill']);
                if ($color) {
                    $styleIndexes[$i]['fill_idx'] = self::addToListGetIndex($fills, $color);
                }
            }
            if (!empty($style['text-align'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['text-align'] = $style['text-align'];
            }
            if (!empty($style['vertical-align'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['vertical-align'] = $style['vertical-align'];
            }
            if (!empty($style['text-wrap'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['text-wrap'] = true;
            }

            $font = null;
            if (!empty($style['font'])) {
                $font = Style::normalaizeFont($style['font']);
            }
            if (!$font) {
                $font = Style::normalaizeFont([]);
            }
            if (isset($style['color'])) {
                $color = Style::normaliazeColor($style['color']);
                if ($color) {
                    $font['color'] = $color;
                }
            }
            $styleIndexes[$i]['font_idx'] = self::addToListGetIndex($fonts, $font);
        }
        return ['fills' => $fills, 'fonts' => $fonts, 'borders' => $borders, 'styles' => $styleIndexes];
    }

    /**
     * @param $border
     * @param $side
     *
     * @return string
     */
    protected function _makeBorderSideTag($border, $side)
    {
        if (empty($border[$side]) || empty($border[$side]['style'])) {
            $tag = "<$side/>";
        }elseif (empty($border[$side]['color'])) {
            $tag = "<$side style=\"" .  $border[$side]['style'] . '"/>';
        } else {
            $tag = "<$side style=\"" .  $border[$side]['style'] . '">';
            $tag .= '<color rgb="' . $border[$side]['color'] . '"/>';
            $tag .= "</$side>";
        }
        return $tag;
    }

    /**
     * @param $borders
     *
     * @return string
     */
    protected function _makeBordersTag($borders)
    {
        $tag = '<borders count="' . (count($borders)) . '">';
        foreach($borders as $border) {
            $tag .= '<border diagonalDown="false" diagonalUp="false">';
            $tag .= $this->_makeBorderSideTag($border, 'left');
            $tag .= $this->_makeBorderSideTag($border, 'right');
            $tag .= $this->_makeBorderSideTag($border, 'top');
            $tag .= $this->_makeBorderSideTag($border, 'bottom');
            $tag .= '<diagonal/>';
            $tag .= '</border>';
        }
        $tag .= '</borders>';

        return $tag;
    }

    /**
     * @return bool|string
     */
    protected function _writeStylesXML()
    {
        $r = $this->_styleFontIndexes();
        $fills = $r['fills'];
        $fonts = $r['fonts'];
        $borders = $r['borders'];
        $styleIndexes = $r['styles'];

        $temporaryFilename = $this->tempFilename();
        $file = new WriterBuffer($temporaryFilename);
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');

        $file->write('<numFmts count="' . count($this->numberFormats) . '">');
        foreach ($this->numberFormats as $i => $v) {
            $file->write('<numFmt numFmtId="' . (164 + $i) . '" formatCode="' . self::xmlSpecialChars($v) . '" />');
        }
        //$file->write(		'<numFmt formatCode="GENERAL" numFmtId="164"/>');
        //$file->write(		'<numFmt formatCode="[$$-1009]#,##0.00;[RED]\-[$$-1009]#,##0.00" numFmtId="165"/>');
        //$file->write(		'<numFmt formatCode="YYYY-MM-DD\ HH:MM:SS" numFmtId="166"/>');
        //$file->write(		'<numFmt formatCode="YYYY-MM-DD" numFmtId="167"/>');
        $file->write('</numFmts>');

        $file->write('<fonts count="' . (count($fonts)) . '">');
        $file->write('<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');

        foreach ($fonts as $font) {
            if (!empty($font)) { //fonts have 4 empty placeholders in array to offset the 4 static xml entries above
                $file->write('<font>');
                $file->write('<name val="' . htmlspecialchars($font['name']) . '"/><charset val="1"/><family val="' . (int)$font['family'] . '"/>');
                $file->write('<sz val="' . (int)$font['size'] . '"/>');
                if (!empty($font['color'])) {
                    $file->write('<color rgb="' . $font['color'] . '"/>');
                }
                if (!empty($font['style-bold'])) {
                    $file->write('<b val="true"/>');
                }
                if (!empty($font['style-italic'])) {
                    $file->write('<i val="true"/>');
                }
                if (!empty($font['style-underline'])) {
                    $file->write('<u val="single"/>');
                }
                if (!empty($font['style-strike'])) {
                    $file->write('<strike val="true"/>');
                }
                $file->write('</font>');
            }
        }
        $file->write('</fonts>');

        $file->write('<fills count="' . (count($fills)) . '">');
        $file->write('<fill><patternFill patternType="none"/></fill>');
        $file->write('<fill><patternFill patternType="gray125"/></fill>');
        foreach ($fills as $fill) {
            if (!empty($fill)) { //fills have 2 empty placeholders in array to offset the 2 static xml entries above
                $file->write('<fill><patternFill patternType="solid"><fgColor rgb="' . $fill . '"/><bgColor indexed="64"/></patternFill></fill>');
            }
        }
        $file->write('</fills>');
        $file->write($this->_makeBordersTag($borders));

        $file->write('<cellStyleXfs count="20">');
        $file->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
        $file->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
        $file->write('<protection hidden="false" locked="true"/>');
        $file->write('</xf>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');

        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
        $file->write('</cellStyleXfs>');

        $file->write('<cellXfs count="' . (count($styleIndexes)) . '">');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0"/>');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"/>');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="166" xfId="0"/>');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="167" xfId="0"/>');
        foreach ($styleIndexes as $v) {
            $fillIdx = isset($v['fill_idx']) ? (int)$v['fill_idx'] : 0;
            $fontIdx = isset($v['font_idx']) ? (int)$v['font_idx'] : 0;
            $borderIdx = isset($v['border_idx']) ? (int)$v['border_idx'] : 0;
            /*
            $applyFont = 'true';
            $applyAlignment = !empty($v['alignment']) ? 'true' : 'false';
            $applyBorder = isset($v['border_idx']) ? 'true' : 'false';

            $horizAlignment = $v['text-align'] ?? 'general';
            $vertAlignment = $v['vertical-align'] ?? 'bottom';
            $wrapText = !empty($v['text-wrap']) ? 'true' : 'false';
            //$file->write('<xf applyAlignment="'.$applyAlignment.'" applyBorder="'.$applyBorder.'" applyFont="'.$applyFont.'" applyProtection="false" borderId="'.($borderIdx).'" fillId="'.($fillIdx).'" fontId="'.($fontIdx).'" numFmtId="'.(164+$v['num_fmt_idx']).'" xfId="0"/>');
            $file->write('<xf applyAlignment="' . $applyAlignment . '" applyBorder="' . $applyBorder . '" applyFont="' . $applyFont . '" applyProtection="false" borderId="' . ($borderIdx) . '" fillId="' . ($fillIdx) . '" fontId="' . ($fontIdx) . '" numFmtId="' . (164 + $v['num_fmt_idx']) . '" xfId="0">');
            $file->write('	<alignment horizontal="' . $horizAlignment . '" vertical="' . $vertAlignment . '" textRotation="0" wrapText="' . $wrapText . '" indent="0" shrinkToFit="false"/>');
            $file->write('	<protection locked="true" hidden="false"/>');
            $file->write('</xf>');
            */

            $xfAttr = 'applyFont="true" ';
            if (!empty($v['alignment'])) {
                $xfAttr .= 'applyAlignment="true" ';
            }
            if (isset($v['border_idx'])) {
                $xfAttr .= 'applyBorder="true" ';
            }
            $attr = '';
            if (!empty($v['text-align'])) {
                $attr .= ' horizontal="' . $v['text-align'] . '"';
            }
            if (!empty($v['vertical-align'])) {
                $attr .= ' vertical="' . $v['vertical-align'] . '"';
            }
            if (!empty($v['text-wrap'])) {
                $attr .= ' wrapText="true"';
            }
            $file->write('<xf ' . $xfAttr . ' borderId="' . ($borderIdx) . '" fillId="' . ($fillIdx) . '" fontId="' . ($fontIdx) . '" numFmtId="' . (164 + $v['num_fmt_idx']) . '" xfId="0">');
            $file->write('	<alignment ' . $attr . '/>');
            $file->write('</xf>');
        }
        $file->write('</cellXfs>');

        $file->write('<cellStyles count="6">');
        $file->write('<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
        $file->write('<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
        $file->write('<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
        $file->write('<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
        $file->write('<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
        $file->write('<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
        $file->write('</cellStyles>');
        $file->write('</styleSheet>');
        $file->close();

        return $temporaryFilename;
    }

    /**
     * @return string
     */
    protected function _buildAppXML($metadata)
    {
        $appXml = '';
        $appXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $appXml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
        $appXml .= '<TotalTime>0</TotalTime>';
        $appXml .= '<Company>' . self::xmlSpecialChars($metadata['company'] ?? '') . '</Company>';
        $appXml .= '</Properties>';

        return $appXml;
    }

    /**
     * @return string
     */
    protected function _buildCoreXML($metadata)
    {
        $coreXml = '';
        $coreXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $coreXml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        $coreXml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . date("Y-m-d\TH:i:s.00\Z") . '</dcterms:created>';//$dateTime = '2014-10-25T15:54:37.00Z';
        $coreXml .= '<dc:title>' . self::xmlSpecialChars($metadata['title'] ?? '') . '</dc:title>';
        $coreXml .= '<dc:subject>' . self::xmlSpecialChars($metadata['subject'] ?? '') . '</dc:subject>';
        $coreXml .= '<dc:creator>' . self::xmlSpecialChars($metadata['author'] ?? '') . '</dc:creator>';
        if (!empty($metadata['keywords'])) {
            $coreXml .= '<cp:keywords>' . self::xmlSpecialChars(implode(", ", (array)$metadata['keywords'])) . '</cp:keywords>';
        }
        $coreXml .= '<dc:description>' . self::xmlSpecialChars($metadata['description'] ?? '') . '</dc:description>';
        $coreXml .= '<cp:revision>0</cp:revision>';
        $coreXml .= '</cp:coreProperties>';

        return $coreXml;
    }

    /**
     * @return string
     */
    protected function _buildRelationshipsXML()
    {
        $relsXml = '';
        $relsXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $relsXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $relsXml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        $relsXml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
        $relsXml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
        $relsXml .= "\n";
        $relsXml .= '</Relationships>';

        return $relsXml;
    }

    /**
     * @param Sheet[] $sheets
     *
     * @return string
     */
    protected function _buildWorkbookXML($sheets)
    {
        $i = 0;
        $workbookXml = '';
        $workbookXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $workbookXml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $workbookXml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $workbookXml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $workbookXml .= '<sheets>';
        $definedNames = '';
        foreach ($sheets as $sheet) {
            $sheetName = self::sanitizeSheetName($sheet->sheetName);
            $workbookXml .= '<sheet name="' . self::xmlSpecialChars($sheetName) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>';
            if ($sheet->absoluteAutoFilter) {
                $filterRange = $sheet->absoluteAutoFilter . ':' . Excel::cellAddress($sheet->rowCount, $sheet->colCount, true);
                $definedNames .= '<definedName name="_xlnm._FilterDatabase" localSheetId="' . $i . '" hidden="1">\'' . $sheetName . '\'!' . $filterRange . '</definedName>';
            }
            $i++;
        }
        $workbookXml .= '</sheets>';
        $workbookXml .= '<definedNames>';
        if ($definedNames) {
            $workbookXml .= $definedNames;
        }
        $workbookXml .= '</definedNames>';
        $workbookXml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';

        return $workbookXml;
    }

    /**
     * @param Sheet[] $sheets
     *
     * @return string
     */
    protected function _buildWorkbookRelsXML($sheets)
    {
        $i = 0;
        $wkbkrelsXml = '';
        $wkbkrelsXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $wkbkrelsXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $wkbkrelsXml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        foreach ($sheets as $sheet) {
            $wkbkrelsXml .= '<Relationship Id="rId' . ($i + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' . ($sheet->xmlName) . '"/>';
            $i++;
        }
        $wkbkrelsXml .= "\n";
        $wkbkrelsXml .= '</Relationships>';

        return $wkbkrelsXml;
    }

    /**
     * @param Sheet[] $sheets
     *
     * @return string
     */
    protected function _buildContentTypesXML($sheets)
    {
        $contentTypesXml = '';
        $contentTypesXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $contentTypesXml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $contentTypesXml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        foreach ($sheets as $sheet) {
            $contentTypesXml .= '<Override PartName="/xl/worksheets/' . ($sheet->xmlName) . '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $contentTypesXml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $contentTypesXml .= "\n";
        $contentTypesXml .= '</Types>';

        return $contentTypesXml;
    }

    /**
     * @deprecated
     *
     * @param $rowNumber     int, zero based
     * @param $columnNumber  int, zero based
     * @param $absolute      bool
     *
     * @return string Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
     */
    public static function cellAddress($rowNumber, $columnNumber, $absolute = false)
    {
        return Excel::cellAddress($rowNumber + 1, $columnNumber + 1);
    }

    /**
     * @param $filename
     *
     * @return mixed
     */
    public static function sanitizeFilename($filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
    {
        $nonprinting = array_map('chr', range(0, 31));
        $invalidChars = ['<', '>', '?', '"', ':', '|', '\\', '/', '*', '&'];
        $allInvalids = array_merge($nonprinting, $invalidChars);

        return str_replace($allInvalids, "", $filename);
    }

    /**
     * @param $sheetName
     *
     * @return string
     */
    public static function sanitizeSheetName($sheetName)
    {
        static $badChars = '\\/?*:[]';
        static $goodChars = '        ';

        $sheetName = strtr($sheetName, $badChars, $goodChars);
        $sheetName = mb_substr($sheetName, 0, 31);
        $sheetName = trim(trim(trim($sheetName), "'"));//trim before and after trimming single quotes

        return !empty($sheetName) ? $sheetName : 'Sheet' . ((mt_rand() % 900) + 100);
    }

    /**
     * @param $val
     *
     * @return string
     */
    public static function xmlSpecialChars($val)
    {
        //note, badchars does not include \t\n\r (\x09\x0a\x0d)
        static $badChars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodChars = "                              ";

        return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badChars, $goodChars);//strtr appears to be faster than str_replace
    }

    /**
     * @param $numFormat
     *
     * @return string
     */
    private static function determineNumberFormatType($numFormat)
    {
        if ($numFormat === 'GENERAL') {
            return 'n_auto';
        }
        if ($numFormat === '@') {
            return 'n_string';
        }
        if ($numFormat === '0') {
            return 'n_numeric';
        }
        if (preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_datetime';
        }
        if (preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_datetime';
        }
        if (preg_match('/[Y]{2,4}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/[D]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/[M]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/$(?![^"]*+")/', $numFormat)) {
            return 'n_numeric';
        }
        if (preg_match('/%(?![^"]*+")/', $numFormat)) {
            return 'n_numeric';
        }
        if (preg_match('/0(?![^"]*+")/', $numFormat)) {
            return 'n_numeric';
        }
        return 'n_auto';
    }

    /**
     * @param $numFormat
     *
     * @return string
     */
    private static function numberFormatStandardized($numFormat)
    {
        $stack = [];
        if (!is_scalar($numFormat) || $numFormat === 'auto' || $numFormat === '' || $numFormat === 'GENERAL') {
            return 'GENERAL';
        }
        if ($numFormat === 'string' || $numFormat === 'text') {
            return '@';
        }
        if ($numFormat === 'integer' || $numFormat === 'int') {
            return '0';
        }
        if ($numFormat === 'percent') {
            return '0%';
        }
        while (isset(self::$localeSettings['formats'][$numFormat])) {
            if (!$numFormat || isset($stack[$numFormat])) {
                break;
            }
            if (isset(self::$localeSettings['formats'][$numFormat])) {
                $numFormat = self::$localeSettings['formats'][$numFormat];
            } else {
                break;
            }
        }

        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($numFormat); $i < $ix; $i++) {
            $c = $numFormat[$i];

            if ($ignoreUntil === '' && $c === '[') {
                $ignoreUntil = ']';
            } elseif ($ignoreUntil === '' && $c === '"') {
                $ignoreUntil = '"';
            } elseif ($ignoreUntil === $c) {
                $ignoreUntil = '';
            }

            if ($ignoreUntil === '' && ($c === ' ' || $c === '-' || $c === '(' || $c === ')') && ($i === 0 || $numFormat[$i - 1] !== '_')) {
                $escaped .= "\\" . $c;
            } else {
                $escaped .= $c;
            }
        }
        return $escaped;
    }

    /**
     * @param $haystack
     * @param $needle
     *
     * @return int
     */
    public static function addToListGetIndex(&$haystack, $needle)
    {
        $existingIdx = array_search($needle, $haystack, $strict = true);
        if ($existingIdx === false) {
            $existingIdx = count($haystack);
            $haystack[] = $needle;
        }
        return $existingIdx;
    }

    /**
     * @param $dateInput
     *
     * @return int|float|bool
     */
    public static function convertDateTime($dateInput) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
    {
        if (is_int($dateInput) || ctype_digit($dateInput)) {
            // date as timestamp
            $time = (int)$dateInput;
        } elseif (preg_match('/^(\d+:\d{1,2})(:\d{1,2})?$/', $dateInput, $matches)) {
            // time only
            $time = strtotime('1900-01-00 ' . $matches[1] . ($matches[2] ?? ':00'));
        } elseif (is_string($dateInput) && $dateInput && $dateInput[0] >= '0' && $dateInput[0] <= '9') {
            //starts with a digit
            $time = strtotime($dateInput);
        } else {
            $time = 0;
        }
        if ($time && preg_match('/(\d{4})-(\d{2})-(\d{2})\s(\d+):(\d{2}):(\d{2})/', date('Y-m-d H:i:s', $time), $matches)) {
            [$junk, $year, $month, $day, $hour, $min, $sec] = $matches;
            $seconds = $sec / 86400 + $min / 1440 + $hour / 24;
        } else {
            // wrong data/time string
            return false;
        }

        //using 1900 as epoch, not 1904, ignoring 1904 special case

        # Special cases for Excel.
        if ("$year-$month-$day" === '1899-12-31') {
            return $seconds;
        }    # Excel 1900 epoch
        if ("$year-$month-$day" === '1900-01-00') {
            return $seconds;
        }    # Excel 1900 epoch
        if ("$year-$month-$day" === '1900-02-29') {
            return 60 + $seconds;
        }    # Excel false leapday

        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;

        # Set month days and check for leap year.
        $leap = (($year % 400 === 0) || (($year % 4 === 0) && ($year % 100))) ? 1 : 0;
        $mdays = [31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

        # Some boundary checks
        if ($year !== 0 || $month !== 0 || $day !== 0) {
            if ($year < $epoch || $year > 9999) {
                // wrong year
                return false;
            }
            if ($month < 1 || $month > 12) {
                // wrong month
                return false;
            }
            if ($day < 1 || $day > $mdays[$month - 1]) {
                // wrong day
                return false;
            }
        }

        # Accumulate the number of days since the epoch.
        $days = $day;    # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1));    # Add days for past months
        $days += $range * 365;                      # Add days for past years
        $days += (int)(($range) / 4);             # Add leapdays
        $days -= (int)(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += (int)(($range + $offset + $norm) / 400);  # Add 400 year leapdays
        $days -= $leap;                                      # Already counted above

        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }
}

// EOF