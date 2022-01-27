<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;

/**
 * Class Excel
 *
 * @package avadim\FastExcelWriter
 */
class Excel
{
    public const EXCEL_2007_MAX_ROW = 1048576;
    public const EXCEL_2007_MAX_COL = 16384;

    protected static $tempDir;

    /** @var array Sheet[] */
    protected $sheets = [];

    /** @var Writer */
    protected $writer;

    protected $metadata = [];

    /** @var bool */
    protected $isRightToLeft = false;


    /**
     * Excel constructor
     *
     * @param array $options
     */
    public function __construct($options = [])
    {
        if (isset($options['writer'])) {
            $writer = $options['writer'];
            $writer->setExcel($this);
            if (self::$tempDir) {
                $writer->setTempDir(self::$tempDir);
            }
        } else {
            $writerOptions = [
                'excel' => $this,
            ];
            if (self::$tempDir) {
                $writerOptions['temp_dir'] = self::$tempDir;
            }
            if (isset($options['temp_dir']) && $options['temp_dir']) {
                $writerOptions['temp_dir'] = $options['temp_dir'];
            }
            $writer = new Writer($writerOptions);
        }
        $this->writer = $writer;
        $this->setDefaultLocale();
        Style::setDefaultFont(['name' => 'Arial', 'size' => 10]);
    }

    /**
     * @param array|string $sheets
     * @param array $options
     *
     * @return Excel
     */
    public static function create($sheets = null, $options = [])
    {
        $excel = new self($options);
        if (empty($sheets)) {
            $sheets = ['Sheet1'];
        } else {
            $sheets = (array)$sheets;
        }
        foreach ($sheets as $sheetName) {
            $excel->makeSheet($sheetName);
        }

        return $excel;
    }

    /**
     * Set dir for temporary files
     *
     * @param $tempDir
     */
    public static function setTempDir($tempDir)
    {
        self::$tempDir = $tempDir;
    }

    /**
     * @param $string
     */
    public static function log($string)
    {
        error_log(date('Y-m-d H:i:s:') . rtrim(is_array($string) ? json_encode($string) : $string) . "\n");
    }

    /**
     * Set default locale from the current environment
     */
    public function setDefaultLocale()
    {
        $this->setLocale('en');
        //$this->setLocale('en_US.UTF-8');
        //$this->setLocale('ru_RU.UTF-8');
        $currentLocale = setlocale(LC_ALL, 0);
        $components = explode(';', $currentLocale);
        foreach ($components as $component) {
            if (strpos($component, '=')) {
                [$param, $locale] = explode('=', $component, 2);
                if ($locale !== 'C' && strlen($locale) > 1) {
                    $this->setLocale($locale);
                    return;
                }
            }
        }
    }

    /**
     * @param string $locale
     * @param string $dir
     *
     * @return $this
     */
    public function setLocale($locale, $dir = null)
    {
        $localeSettings = [];
        // default settings
        $aFormatSettings = [
            'formats' => [
                'date' => 'YYYY-MM-DD',
                'time' => 'HH:MM:SS',
                'datetime' => 'YYYY-MM-DD HH:MM:SS',
                'money' => '# ##0.00',
            ],
        ];

        if (empty($locale)) {
            $newLocale = false;
            $oldLocale = false;
        } else {
            if (strpos($locale, '.')) {
                [$localeName, $localePage] = explode('.', $locale);
                $locale = $localeName . '.UTF-8';
            } else {
                $locale .= '.UTF-8';
            }

            // try to set locale
            $oldLocale = setlocale(LC_CTYPE, '0');
            $newLocale = setlocale(LC_CTYPE, $locale);
            if (!$newLocale && strpos($locale, '_')) {
                [$language, $country] = explode('_', $locale);
                $newLocale = setlocale(LC_CTYPE, $country);
                if (!$newLocale) {
                    $newLocale = setlocale(LC_CTYPE, $language);
                }
            }
        }
        if ($newLocale) {
            // set date & time patterns
            $tmpLocale = setlocale(LC_TIME, '0');
            if (setlocale(LC_TIME, $newLocale)) {
                $time = strtotime('1985-6-2 13:7:9');
                $datePattern = strftime('%x', $time);
                $dateTimePattern = strftime('%c', $time);
                $timePattern = trim(str_replace($datePattern, '', $dateTimePattern));
                $datePattern = str_replace(['19', '85', '06', '02', '6', '2'], ['YY', 'YY', 'MM', 'DD', 'M', 'D'], $datePattern);
                $timePattern = str_replace(['13', '07', '09', '1', '7', '9', 'PM', 'pm'], ['H', 'MM', 'SS', 'H', 'M', 'S', 'AM/PM', 'am/pm'], $timePattern);

                $aFormatSettings['formats']['date'] = $datePattern;
                $aFormatSettings['formats']['time'] = $timePattern;
                $aFormatSettings['formats']['datetime'] = $datePattern . ' ' . $timePattern;
            }
            if ($tmpLocale) {
                setlocale(LC_TIME, $tmpLocale);
            }

            // set money pattern
            $tmpLocale = setlocale(LC_MONETARY, '0');
            if (setlocale(LC_MONETARY, $newLocale)) {
                $locale_info = localeconv();

                $moneyPattern = '0' . $locale_info['decimal_point'] . str_repeat('0', $locale_info['frac_digits']);
                if ($locale_info['thousands_sep']) {
                    $moneyPattern = '#' . $locale_info['thousands_sep'] . '##' . $moneyPattern;
                }
                $sepSpace = !empty($locale_info['p_sep_by_space']) ? str_repeat(' ', $locale_info['p_sep_by_space']) : '';
                if (!empty($locale_info['p_cs_precedes'])) {
                    $moneyPattern = '[$' . $locale_info['currency_symbol'] . ']' . $sepSpace . $moneyPattern;
                } else {
                    $moneyPattern .= $sepSpace . '[$' . $locale_info['currency_symbol'] . ']';
                }
                $aFormatSettings['formats']['money'] = $moneyPattern;
                setlocale(LC_MONETARY, $oldLocale);
            }
            if ($tmpLocale) {
                setlocale(LC_MONETARY, $tmpLocale);
            }
        }

        if ($oldLocale) {
            setlocale(LC_CTYPE, $oldLocale);
        }

        if (!empty($locale)) {
            $locale = strtolower($locale);
            if (!$dir) {
                $dir = __DIR__ . '/locale';
            }
            $file = $dir . '/' . $locale . '/settings.php';

            // try load locale settings file
            $includeFile = '';
            if (is_file($file)) {
                $includeFile = $file;
            } else {
                $file = str_replace('.utf-8/', '/', $file);
                if (is_file($file)) {
                    $includeFile = $file;
                }
            }
            if ($includeFile && ($localeData = include($includeFile))) {
                $localeSettings = array_merge($localeSettings, $localeData);
            }
            if (strpos($locale, '_')) {
                [$language, $country] = explode('_', $locale);
                $file = $dir . '/' . $language . '/settings.php';
                if (is_file($file) && ($localeData = include($file))) {
                    $localeSettings = array_merge($localeSettings, $localeData);
                }
                $file = $dir . '/' . $language . '/' . $country . '/settings.php';
                if (is_file($file) && ($localeData = include($file))) {
                    $localeSettings = array_merge($localeSettings, $localeData);
                }
                $file = str_replace('.utf-8/', '/', $file);
                if (is_file($file) && ($localeData = include($file))) {
                    $localeSettings = array_merge($localeSettings, $localeData);
                }
            }
        }
        if ($localeSettings) {
            $localeSettings = array_merge($aFormatSettings, $localeSettings);
        } else {
            $localeSettings = $aFormatSettings;
        }

        $this->writer::setLocaleSettings($localeSettings);

        return $this;
    }

    /**
     * @param string $title
     *
     * @return $this
     */
    public function setTitle($title = '')
    {
        return $this->setMetadata('title', $title);
    }

    /**
     * @param string $subject
     *
     * @return $this
     */
    public function setSubject($subject = '')
    {
        return $this->setMetadata('subject', $subject);
    }

    /**
     * @param string $author
     *
     * @return $this
     */
    public function setAuthor($author = '')
    {
        return $this->setMetadata('author', $author);
    }

    /**
     * @param string $company
     *
     * @return $this
     */
    public function setCompany($company = '')
    {
        return $this->setMetadata('company', $company);
    }

    /**
     * @param string $description
     *
     * @return $this
     */
    public function setDescription($description = '')
    {
        return $this->setMetadata('description', $description);
    }

    /**
     * @param mixed $keywords
     *
     * @return $this
     */
    public function setKeywords($keywords)
    {
        if (!$keywords) {
            $newKeywords = [];
        } else {
            $oldKeywords = $this->getMetadata('keywords');
            $newKeywords = is_array($keywords) ? $keywords : array_map('trim', explode(',', $keywords));
            $newKeywords = array_unique(array_merge($oldKeywords, $newKeywords));
        }
        return $this->setMetadata('keywords', $newKeywords);
    }

    /**
     * @param $key
     * @param $value
     *
     * @return $this
     */
    public function setMetadata($key, $value)
    {
        $this->metadata[$key] = $value;

        return $this;
    }

    /**
     * @param null $key
     *
     * @return mixed
     */
    public function getMetadata($key = null)
    {
        if ($key) {
            return $this->metadata[$key] ?? null;
        }
        return $this->metadata;
    }

    /**
     * @param bool $isRightToLeft
     */
    public function setRightToLeft($isRightToLeft = false)
    {
        $this->isRightToLeft = (bool)$isRightToLeft;
    }

    /**
     * @return bool
     */
    public function isRightToLeft()
    {
        return $this->isRightToLeft;
    }

    /**
     * Set default
     *
     * @param $font
     */
    public function setFont($font)
    {
        Style::setDefaultFont($font);
    }

    /**
     * @param $colLetter
     *
     * @return int
     */
    public static function colIndex($colLetter)
    {
        // Strip cell reference down to just letters
        $letters = preg_replace('/[^A-Z]/', '', $colLetter);

        if (mb_strlen($letters) >= 3 && $letters > 'XFD') {
            return -1;
        }
        // Iterate through each letter, starting at the back to increment the value
        for ($index = 0, $i = 0; $letters !== ''; $letters = substr($letters, 0, -1), $i++) {
            $index += (ord(substr($letters, -1)) - 64) * (26 ** $i);
        }

        return ($index <= self::EXCEL_2007_MAX_COL) ? (int)$index: -1;
    }

    /**
     * @param $colIndex
     *
     * @return string
     */
    public static function colLetter($colIndex)
    {
        static $letters = ['',
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
        ];

        if (isset($letters[$colIndex])) {
            return $letters[$colIndex];
        }

        if ($colIndex > 0 && $colIndex <= self::EXCEL_2007_MAX_COL) {
            $num = $colIndex - 1;
            for ($letter = ''; $num >= 0; $num = (int)($num / 26) - 1) {
                $letter = chr($num % 26 + 0x41) . $letter;
            }
            $letters[$colIndex] = $letter;

            return $letter;
        }
        return '';
    }

    /**
     * Create cell address by row and col indexws
     *
     * @param int $rowIndex one based
     * @param int $colIndex one based
     * @param bool $absolute
     * @param bool $absoluteRow
     *
     * @return string Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
     */
    public static function cellAddress($rowIndex, $colIndex, $absolute = false, $absoluteRow = null)
    {
        if ($rowIndex > 0 && $colIndex > 0) {
            $letter = self::colLetter($colIndex);
            if ($letter) {
                if ($absolute) {
                    if (null === $absoluteRow || true === $absoluteRow) {
                        return '$' . $letter . '$' . $rowIndex;
                    }
                    return '$' . $letter . $rowIndex;
                }
                if ($absoluteRow) {
                    return $letter . '$' . $rowIndex;
                }
                return $letter . $rowIndex;
            }
        }
        return '';
    }

    /**
     * @param array|string $range
     * @param array|string $offset
     * @param bool $exception
     *
     * @return array|bool
     */
    public static function rangeDimensionRelative($range, $offset, $exception = false)
    {
        $colIndex1 = $colIndex2 = $rowIndex1 = $rowIndex2 = null;
        if (is_array($range)) {
            $sheetName = null;
            if (count($range) === 2) {
                // [[row1, col1], [row2, col2]] or [['row' => row1, 'col' => col1], ['row' => row2, 'col' => col2]]
                [$cell1, $cell2] = $range;
                if (is_array($cell1) && is_array($cell2)) {
                    if (isset($cell1['row'], $cell1['col'])) {
                        $rowIndex1 = $cell1['row'];
                        $colIndex1 = $cell1['col'];
                    } else {
                        [$rowIndex1, $colIndex1] = $cell1;
                    }
                    if (isset($cell2['row'], $cell2['col'])) {
                        $rowIndex2 = $cell1['row'];
                        $colIndex2 = $cell1['col'];
                    } else {
                        [$rowIndex2, $colIndex2] = $cell2;
                    }
                }
            } else {
                // [row1, col1, row2, col2]
                [$rowIndex1, $colIndex1, $rowIndex2, $colIndex2] = $range;
            }
        } else {
            if (strpos($range, '!')) {
                [$sheetName, $range] = explode('!', $range);
            } else {
                $sheetName = null;
            }
            $range = strtoupper($range);
            if (preg_match('/^([A-Z]+)(\d+)(:([A-Z]+)(\d+))?$/', $range, $matches)) {
                if (empty($matches[3])) {
                    $matches[4] = $matches[1];
                    $matches[5] = $matches[2];
                }
                $colIndex1 = self::colIndex($matches[1]);
                $colIndex2 = self::colIndex($matches[4]);
                $rowIndex1 = ($matches[2] <= self::EXCEL_2007_MAX_ROW) ? (int)$matches[2] : -1;
                $rowIndex2 = ($matches[5] <= self::EXCEL_2007_MAX_ROW) ? (int)$matches[5] : -1;
            }
        }
        if ($exception && ($colIndex1 === null || $colIndex2 === null || $rowIndex1 === null || $rowIndex2 === null)) {
            throw new Exception('Wrong range ' . print_r($range, true) . '');
        }
        if (!empty($offset)) {
            $rowOffset1 = $colOffset1 = $rowOffset2 = $colOffset2 = null;
            if (is_array($offset)) {
                if (count($offset) === 4) {
                    [$rowOffset1, $colOffset1, $rowOffset2, $colOffset2] = $offset;
                } elseif (count($offset) === 2) {
                    // [row, col] or ['row' => row, 'col' => col]
                    if (isset($offset['row'], $offset['col'])) {
                        $rowOffset1 = $offset['row'];
                        $colOffset1 = $offset['col'];
                    } else {
                        [$rowOffset2, $colOffset2] = $offset;
                    }
                }
            } else {
                // 'R1C1'
                [$rowOffset1, $colOffset1, $rowOffset2, $colOffset2] = self::rangeRelOffsets($offset);
            }
            if ($exception && ($rowOffset1 === null || $colOffset1 === null || $rowOffset2 === null || $colOffset2 === null)) {
                throw new Exception('Wrong offset of range ' . print_r($offset, true) . '');
            }

            $rowIndex1 += $rowOffset1;
            $colIndex1 += $colOffset1;
            $rowIndex2 += $rowOffset2;
            $colIndex2 += $colOffset2;
        }
        if ($exception && ($colIndex1 < 0 || $colIndex2 < 0 || $rowIndex1 < 0 || $rowIndex2 < 0)) {
            throw new Exception('Wrong range ' . print_r($range, true) . '');
        }

        if ($colIndex1 > 0 && $colIndex2 > 0 && $rowIndex1 > 0 && $rowIndex2 > 0) {
            // swap indexes if need
            if ($colIndex2 < $colIndex1) {
                $idx = $colIndex1;
                $colIndex1 = $colIndex2;
                $colIndex2 = $idx;
            }
            if ($rowIndex2 < $rowIndex1) {
                $idx = $rowIndex1;
                $rowIndex1 = $rowIndex2;
                $rowIndex2 = $idx;
            }
            $cell1 = Excel::colLetter($colIndex1) . $rowIndex1;
            $cell2 = Excel::colLetter($colIndex2) . $rowIndex2;
            $localRange = $cell1 . ':' . $cell2;
            $width = $colIndex2 - $colIndex1 + 1;
            $height = $rowIndex2 - $rowIndex1 + 1;

            return [
                'range' => ($sheetName ? $sheetName . '!' : '') . $localRange,
                'sheet' => $sheetName,
                'cell1' => $cell1,
                'cell2' => $cell2,
                'localRange' => $localRange,
                'rowIndex1' => $rowIndex1,
                'colIndex1' => $colIndex1,
                'rowIndex2' => $rowIndex2,
                'colIndex2' => $colIndex2,
                'width' => $width,
                'height' => $rowIndex2 - $rowIndex1 + 1,
                'cellCount' => $width * $height,
            ];
        }

        return false;
    }

    /**
     * @param array|string $range
     * @param bool $exception
     *
     * @return array|bool
     */
    public static function rangeDimension($range, $exception = false)
    {
        return self::rangeDimensionRelative($range, null, $exception);
    }

    /**
     * Return offsets by relative address (zero based)
     *
     * @param $relAddress
     *
     * @return int[]
     */
    public static function rangeRelOffsets($relAddress)
    {
        $rowOffset1 = $colOffset1 = $rowOffset2 = $colOffset2 = null;
        $offset = strtoupper($relAddress);
        if (preg_match('/^R\[?(-?\d+)?\]?C\[?(-?\d+)?\]?(:R\[?(-?\d+)?\]?C\[?(-?\d+)?\]?)?$/', $offset, $matches)) {
            $rowOffset1 = !empty($matches[1]) ? (int)$matches[1] : 0;
            $colOffset1 = !empty($matches[2]) ? (int)$matches[2] : 0;
            if (!empty($matches[3])) {
                $rowOffset2 = !empty($matches[4]) ? (int)$matches[4] : 0;
                $colOffset2 = !empty($matches[5]) ? (int)$matches[5] : 0;
            } else {
                $rowOffset2 = $rowOffset1;
                $colOffset2 = $colOffset1;
            }
        }
        return [
            $rowOffset1,
            $colOffset1,
            $rowOffset2,
            $colOffset2,
        ];
    }

    /**
     * @param $range
     *
     * @return int[]
     */
    public static function rangeIndexes($range)
    {
        $dimension = Excel::rangeDimension($range);
        if ($dimension) {
            return ['row' => $dimension['rowIndex1'], 'col' => $dimension['colIndex1']];
        }
        return ['row' => -1, 'col' => -1];
    }

    /**
     * @return Writer
     */
    public function getWriter()
    {
        return $this->writer;
    }

    /**
     * @param $sheetName
     *
     * @return Sheet
     */
    public function makeSheet($sheetName)
    {
        $key = mb_strtolower($sheetName);
        if (!isset($this->sheets[$key])) {
            $this->sheets[$key] = new Sheet($sheetName);
            $this->sheets[$key]->book = $this;
            $this->sheets[$key]->key = $key;
            $this->sheets[$key]->index = count($this->sheets);
            $this->sheets[$key]->xmlName = 'sheet' . $this->sheets[$key]->index . '.xml';
            if (count($this->sheets) === 1) {
                $this->sheets[$key]->active = true;
            }
        }
        return $this->sheets[$key];
    }

    /**
     * @param int|string $index
     *
     * @return Sheet|null
     */
    public function getSheet($index = null)
    {
        if (null === $index) {
            return reset($this->sheets);
        }

        if (is_int($index)) {
            $keys = array_keys($this->sheets);
            if (isset($keys[--$index])) {
                $key = $keys[$index];
            } else {
                // index not found
                throw  new Exception('Sheet #' . $index . ' not found');
            }
        } else {
            $key = mb_strtolower($index);
            if (!isset($this->sheets[$key])) {
                throw  new Exception('Sheet "' . $index . '" not found');
            }
        }
        return $this->sheets[$key] ?? null;
    }
    
    /**
     * @param int|string $index
     */
    public function removeSheet($index = null): void
    {
        if (null === $index) {
            array_shift($this->sheets);
        }

        if (is_int($index)) {
            $keys = array_keys($this->sheets);
            if (!isset($keys[--$index])) {
                throw  new Exception('Sheet #' . $index . ' not found');
            }
            unset($this->sheets[$keys[$index]]);
        } else {
            $key = mb_strtolower($index);
            if (!isset($this->sheets[$key])) {
                throw  new Exception('Sheet "' . $index . '" not found');
            }
            unset($this->sheets[$key]);
        }
    }

    /**
     * @return Sheet[]
     */
    public function getSheets()
    {
        return $this->sheets;
    }

    /**
     * @param $fileName
     * @param $overWrite
     */
    public function save($fileName, $overWrite = true)
    {
        $this->writer->saveToFile($fileName, $overWrite, $this->getMetadata());
    }

}

 // EOF
