<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Exceptions\ExceptionAddress;
use avadim\FastExcelWriter\Exceptions\ExceptionFile;
use avadim\FastExcelWriter\Exceptions\ExceptionRangeName;
use avadim\FastExcelWriter\Interfaces\InterfaceBookWriter;
use avadim\FastExcelWriter\Writer\Writer;

/**
 * Class Excel
 *
 * @package avadim\FastExcelWriter
 */
class Excel implements InterfaceBookWriter
{
    public const MAX_ROW = 1048576; // max row number in Excel 2007
    public const MAX_COL = 16384; // max column number in Excel 2007
    public const MIN_ROW = 0;
    public const MIN_COL = 0;

    public const EMU_PER_PIXEL = 9525;

    public const DEFAULT_COL_WIDTH = '8.83984375';

    public const PAPERSIZE_LETTER = 1; // Letter paper (8.5 in. by 11 in.)
    public const PAPERSIZE_LETTER_SMALL = 2; // Letter small paper (8.5 in. by 11 in.)
    public const PAPERSIZE_TABLOID = 3; // Tabloid paper (11 in. by 17 in.)
    public const PAPERSIZE_LEDGER = 4; // Ledger paper (17 in. by 11 in.)
    public const PAPERSIZE_LEGAL = 5; // Legal paper (8.5 in. by 14 in.)
    public const PAPERSIZE_STATEMENT = 6; // Statement paper (5.5 in. by 8.5 in.)
    public const PAPERSIZE_EXECUTIVE = 7; // Executive paper (7.25 in. by 10.5 in.)
    public const PAPERSIZE_A3 = 8; // A3 paper (297 mm by 420 mm)
    public const PAPERSIZE_A4 = 9; // A4 paper (210 mm by 297 mm)
    public const PAPERSIZE_A4_SMALL = 10; // A4 small paper (210 mm by 297 mm)
    public const PAPERSIZE_A5 = 11; // A5 paper (148 mm by 210 mm)
    public const PAPERSIZE_B4 = 12; // B4 paper (250 mm by 353 mm)
    public const PAPERSIZE_B5 = 13; // B5 paper (176 mm by 250 mm)
    public const PAPERSIZE_FOLIO = 14; // Folio paper (8.5 in. by 13 in.)
    public const PAPERSIZE_QUARTO = 15; // Quarto paper (215 mm by 275 mm)
    public const PAPERSIZE_STANDARD_1 = 16; // Standard paper (10 in. by 14 in.)
    public const PAPERSIZE_STANDARD_2 = 17; // Standard paper (11 in. by 17 in.)
    public const PAPERSIZE_NOTE = 18; // Note paper (8.5 in. by 11 in.)
    public const PAPERSIZE_NO9_ENVELOPE = 19; // #9 envelope (3.875 in. by 8.875 in.)
    public const PAPERSIZE_NO10_ENVELOPE = 20; // #10 envelope (4.125 in. by 9.5 in.)
    public const PAPERSIZE_NO11_ENVELOPE = 21; // #11 envelope (4.5 in. by 10.375 in.)
    public const PAPERSIZE_NO12_ENVELOPE = 22; // #12 envelope (4.75 in. by 11 in.)
    public const PAPERSIZE_NO14_ENVELOPE = 23; // #14 envelope (5 in. by 11.5 in.)
    public const PAPERSIZE_C = 24; // C paper (17 in. by 22 in.)
    public const PAPERSIZE_D = 25; // D paper (22 in. by 34 in.)
    public const PAPERSIZE_E = 26; // E paper (34 in. by 44 in.)
    public const PAPERSIZE_DL_ENVELOPE = 27; // DL envelope (110 mm by 220 mm)
    public const PAPERSIZE_C5_ENVELOPE = 28; // C5 envelope (162 mm by 229 mm)
    public const PAPERSIZE_C3_ENVELOPE = 29; // C3 envelope (324 mm by 458 mm)
    public const PAPERSIZE_C4_ENVELOPE = 30; // C4 envelope (229 mm by 324 mm)
    public const PAPERSIZE_C6_ENVELOPE = 31; // C6 envelope (114 mm by 162 mm)
    public const PAPERSIZE_C65_ENVELOPE = 32; // C65 envelope (114 mm by 229 mm)
    public const PAPERSIZE_B4_ENVELOPE = 33; // B4 envelope (250 mm by 353 mm)
    public const PAPERSIZE_B5_ENVELOPE = 34; // B5 envelope (176 mm by 250 mm)
    public const PAPERSIZE_B6_ENVELOPE = 35; // B6 envelope (176 mm by 125 mm)
    public const PAPERSIZE_ITALY_ENVELOPE = 36; // Italy envelope (110 mm by 230 mm)
    public const PAPERSIZE_MONARCH_ENVELOPE = 37; // Monarch envelope (3.875 in. by 7.5 in.).
    public const PAPERSIZE_6_3_4_ENVELOPE = 38; // 6 3/4 envelope (3.625 in. by 6.5 in.)
    public const PAPERSIZE_US_STANDARD_FANFOLD = 39; // US standard fanfold (14.875 in. by 11 in.)
    public const PAPERSIZE_GERMAN_STANDARD_FANFOLD = 40; // German standard fanfold (8.5 in. by 12 in.)
    public const PAPERSIZE_GERMAN_LEGAL_FANFOLD = 41; // German legal fanfold (8.5 in. by 13 in.)
    public const PAPERSIZE_ISO_B4 = 42; // ISO B4 (250 mm by 353 mm)
    public const PAPERSIZE_JAPANESE_DOUBLE_POSTCARD = 43; // Japanese double postcard (200 mm by 148 mm)
    public const PAPERSIZE_STANDARD_PAPER_1 = 44; // Standard paper (9 in. by 11 in.)
    public const PAPERSIZE_STANDARD_PAPER_2 = 45; // Standard paper (10 in. by 11 in.)
    public const PAPERSIZE_STANDARD_PAPER_3 = 46; // Standard paper (15 in. by 11 in.)
    public const PAPERSIZE_INVITE_ENVELOPE = 47; // Invite envelope (220 mm by 220 mm)
    public const PAPERSIZE_LETTER_EXTRA_PAPER = 50; // Letter extra paper (9.275 in. by 12 in.)
    public const PAPERSIZE_LEGAL_EXTRA_PAPER = 51; // Legal extra paper (9.275 in. by 15 in.)
    public const PAPERSIZE_TABLOID_EXTRA_PAPER = 52; // Tabloid extra paper (11.69 in. by 18 in.)
    public const PAPERSIZE_A4_EXTRA_PAPER = 53; // A4 extra paper (236 mm by 322 mm)
    public const PAPERSIZE_LETTER_TRANSVERSE_PAPER = 54; // Letter transverse paper (8.275 in. by 11 in.)
    public const PAPERSIZE_A4_TRANSVERSE_PAPER = 55; // A4 transverse paper (210 mm by 297 mm)
    public const PAPERSIZE_LETTER_EXTRA_TRANSVERSE_PAPER = 56; // Letter extra transverse paper (9.275 in. by 12 in.)
    public const PAPERSIZE_SUPERA_SUPERA_A4_PAPER = 57; // SuperA/SuperA/A4 paper (227 mm by 356 mm)
    public const PAPERSIZE_SUPERB_SUPERB_A3_PAPER = 58; // SuperB/SuperB/A3 paper (305 mm by 487 mm)
    public const PAPERSIZE_LETTER_PLUS_PAPER = 59; // Letter plus paper (8.5 in. by 12.69 in.)
    public const PAPERSIZE_A4_PLUS_PAPER = 60; // A4 plus paper (210 mm by 330 mm)
    public const PAPERSIZE_A5_TRANSVERSE_PAPER = 61; // A5 transverse paper (148 mm by 210 mm)
    public const PAPERSIZE_JIS_B5_TRANSVERSE_PAPER = 62; // JIS B5 transverse paper (182 mm by 257 mm)
    public const PAPERSIZE_A3_EXTRA_PAPER = 63; // A3 extra paper (322 mm by 445 mm)
    public const PAPERSIZE_A5_EXTRA_PAPER = 64; // A5 extra paper (174 mm by 235 mm)
    public const PAPERSIZE_ISO_B5_EXTRA_PAPER = 65; // ISO B5 extra paper (201 mm by 276 mm)
    public const PAPERSIZE_A2_PAPER = 66; // A2 paper (420 mm by 594 mm)
    public const PAPERSIZE_A3_TRANSVERSE_PAPER = 67; // A3 transverse paper (297 mm by 420 mm)
    public const PAPERSIZE_A3_EXTRA_TRANSVERSE_PAPER = 68; // A3 extra transverse paper (322 mm by 445 mm)
    public const PAPERSIZE_A6 = 70; // A6 (105 mm x 148 mm)
    public const PAPERSIZE_JAPANESE_ENVELOPE_KAKU_2 = 71; // Japanese Envelope Kaku #2
    public const PAPERSIZE_JAPANESE_ENVELOPE_KAKU_3 = 72; // Japanese Envelope Kaku #3
    public const PAPERSIZE_JAPANESE_ENVELOPE_CHOU_3 = 73; // Japanese Envelope Chou #3
    public const PAPERSIZE_JAPANESE_ENVELOPE_CHOU_4 = 74; // Japanese Envelope Chou #4
    public const PAPERSIZE_LETTER_ROTATED = 75; // Letter Rotated (11in x 8 1/2 11 in)
    public const PAPERSIZE_A3_ROTATED = 76; // A3 Rotated (420 mm x 297 mm)
    public const PAPERSIZE_A4_ROTATED = 77; // A4 Rotated (297 mm x 210 mm)
    public const PAPERSIZE_A5_ROTATED = 78; // A5 Rotated (210 mm x 148 mm)
    public const PAPERSIZE_B4_JIS = 79; // B4 (JIS) Rotated (364 mm x 257 mm)
    public const PAPERSIZE_B5_JIS = 80; // B5 (JIS) Rotated (257 mm x 182 mm)
    public const PAPERSIZE_JAPANESE_POSTCARD_ROTATED = 81; // Japanese Postcard Rotated (148 mm x 100 mm)
    public const PAPERSIZE_A6_ROTATED = 83; // A6 Rotated (148 mm x 105 mm)
    public const PAPERSIZE_B6_JIS = 88; // B6 (JIS) (128 mm x 182 mm)
    public const PAPERSIZE_B6_JIS_ROTATED = 89; // B6 (JIS) Rotated (182 mm x 128 mm)


    public static array $availableImageTypes = [
        'jpg' => 'image/jpeg',
        'gif' => 'image/gif',
        'png' => 'image/png',
        'webp' => 'image/webp',
        'bmp' => 'image/bmp',
        'svg' => 'image/svg+xml',
    ];

    protected static string $tempDir;

    /** @var Writer */
    public $writer;

    /** @var StyleManager */
    public $style;

    public bool $saved = false;

    public FormulaConverter $formulaConverter;

    /** @var int  */
    protected int $maxSheetIndex = 0;

    /** @var array Sheet[] */
    protected array $sheets = [];

    protected array $themes = [];

    protected array $metadata = [];

    protected array $bookViews = [];

    protected array $definedNames = [];

    /** @var bool */
    protected bool $isRightToLeft = false;

    protected array $sharedStrings = [];
    protected int $sharedStringsCount = 0;

    protected string $fileName = '';

    protected array $media = [];

    protected array $protection = [];



    /**
     * Excel constructor
     *
     * @param array|null $options Optional parameters: ['temp_dir' => ..., 'temp_prefix' => ..., 'auto_convert_number' => ..., 'shared_string' => ...]
     */
    public function __construct(?array $options = [])
    {
        $writerOptions = [
            'excel' => $this,
        ];
        if (!empty(self::$tempDir)) {
            $writerOptions['temp_dir'] = self::$tempDir;
        }
        if (isset($options['temp_dir']) && $options['temp_dir']) {
            $writerOptions['temp_dir'] = $options['temp_dir'];
        }
        if (isset($options['temp_prefix']) && $options['temp_prefix']) {
            $writerOptions['temp_prefix'] = $options['temp_prefix'];
        }
        $writerOptions['auto_convert_number'] = !empty($options['auto_convert_number']);
        $writerOptions['shared_string'] = !empty($options['shared_string']);

        if (isset($options['writer_class'])) {
            $this->writer = $this->getObject($options['writer_class'], $writerOptions);
            $this->writer->setExcel($this);
            if (self::$tempDir) {
                $this->writer->setTempDir(self::$tempDir);
            }
        }
        else {
            $this->writer = new Writer($writerOptions);
        }

        if (isset($options['style_manager'])) {
            $this->style = $this->getObject($options['style_manager']);
        }
        else {
            $this->style = new StyleManager($options);
        }

        $this->setDefaultLocale();
        if (!empty($options['locale'])) {
            $this->setLocale($options['locale']);
        }
        $settings = $this->style->getLocaleSettings();
        $this->formulaConverter = new FormulaConverter($settings['functions'] ?? []);

        $this->bookViews = [
            [
                'activeTab' => '0',
                'firstSheet' => '0',
                'showHorizontalScroll' => 'true',
                'showSheetTabs' => 'true',
                'showVerticalScroll' => 'true',
            ]
        ];
    }


    public function __destruct()
    {
        if (!$this->saved && $this->fileName) {
            $this->writer->saveToFile($this->fileName, false, $this->getMetadata());
        }
    }


    /**
     * @param mixed $value
     * @param string $format
     *
     * @return string
     */
    public static function _formatValue($value, string $format): string
    {
        if (is_numeric($value)) {
            if (strpos($format, ';')) {
                $formats = explode(';', $format);
                if ($value > 0 && !empty($formats[0])) {
                    return self::_formatValue($value, $formats[0]);
                }
                if ($value < 0 && !empty($formats[1])) {
                    return self::_formatValue($value, $formats[1]);
                }
                if ((int)$value === 0 && !empty($formats[2])) {
                    return self::_formatValue($value, $formats[2]);
                }
                return self::_formatValue($value, '0');
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
                if (preg_match('/\[\$.+]/U', $format, $m)) {
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
     * @param string|object $class
     * @param string|array $options
     *
     * @return object
     */
    protected function getObject($class, $options = null): object
    {
        if (is_object($class)) {
            return $class;
        }

        return new $class($options);
    }

    /**
     * Create new workbook
     *
     * @param array|string|null $sheets Name of sheet or array of names
     * @param array|null $options Options
     *
     * @return Excel
     */
    public static function create($sheets = null, ?array $options = []): Excel
    {
        $excel = new self($options);
        if ($sheets) {
            if (is_array($sheets)) {
                foreach ($sheets as $sheetName) {
                    $excel->makeSheet($sheetName);
                }
            }
            else {
                $excel->makeSheet((string)$sheets);
            }
        }
        else {
            $excel->makeSheet();
        }

        return $excel;
    }

    /**
     * @param string $sheetName
     *
     * @return Sheet
     */
    public static function createSheet(string $sheetName): Sheet
    {
        return new Sheet($sheetName);
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
     * Generate UUID v4
     *
     * @return string
     */
    public static function generateUuid(): string
    {
        // xxxxxxxx-xxxx-4xxx-[8-B]xxx-xxxxxxxxxxxx
        $hash = md5(microtime());
        try {
            $uuid = substr($hash, 0, 8) . '-' . substr($hash, 8, 4)
                . '-4' . dechex(random_int(256, 4095))
                . '-' . dechex(random_int(8, 11)) . dechex(random_int(256, 4095))
                . '-' . substr($hash, 12, 12);
        }
        catch (\Throwable $e) {
            $uuid = substr($hash, 0, 8) . '-' . substr($hash, 8, 4)
                . '-4' . dechex(mt_rand(256, 4095))
                . '-' . dechex(mt_rand(8, 11)) . dechex(mt_rand(256, 4095))
                . '-' . substr($hash, 12, 12);
        }

        return strtoupper($uuid);
    }

    /**
     * @param $string
     */
    public static function log($string)
    {
        error_log(date('Y-m-d H:i:s:') . rtrim(is_array($string) ? json_encode($string) : $string) . "\n");
    }

    /**
     * Convert value (int or string) to Excel timestamp
     *
     * @param int|string $value
     *
     * @return float|bool
     */
    public static function toTimestamp($value)
    {
        $result = Writer::convertDateTime($value);
        if ($result === false) {
            Exception::throwNew('Cannot convert "' . $value . '" to Excel timestamp');
        }

        return $result;
    }

    /**
     * Set default locale from the current environment
     */
    public function setDefaultLocale()
    {
        $this->setLocale('en');
        if (extension_loaded('intl')) {
            $locale = \Locale::getDefault();
            if ($locale) {
                $this->setLocale($locale);
            }
        }
    }

    /**
     * @param string $file
     * @param array $localeSettings
     *
     * @return void
     */
    protected function loadSettings(string $file, array &$localeSettings)
    {
        if ($file && is_file($file)) {
            $localeData = include($file);
            if (!empty($localeData['formats'])) {
                $formats = [];
                foreach($localeData['formats'] as $key => $val) {
                    if ($key && is_string($key)) {
                        $newKey = strtoupper($key);
                        if ($newKey[0] !== '@') {
                            $newKey = '@' . $newKey;
                        }
                        $formats[$newKey] = $val;
                    }
                }
                $localeData['formats'] = $formats;
            }
            $localeSettings = array_merge($localeSettings, $localeData);
        }
    }

    /**
     * Set locale information
     *
     * @param string $locale
     * @param string|null $dir
     *
     * @return $this
     */
    public function setLocale(string $locale, ?string $dir = null): Excel
    {
        $localeSettings = [];
        // default settings
        $aFormatSettings = [
            'formats' => [
                '@DATE' => 'YYYY-MM-DD',
                '@TIME' => 'HH:MM:SS',
                '@DATETIME' => 'YYYY-MM-DD HH:MM:SS',
                '@MONEY' => '# ##0.00',
            ],
        ];

        if (empty($locale)) {
            $newLocale = false;
            $oldLocale = false;
        }
        else {
            if ($locale === 'en') {
                $locale = 'en_US.UTF-8';
            }
            elseif (strpos($locale, '.')) {
                [$localeName, $localePage] = explode('.', $locale);
                $locale = $localeName . '.UTF-8';
            }
            else {
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
            // set money pattern
            $tmpLocale = setlocale(LC_MONETARY, '0');
            if (setlocale(LC_MONETARY, $newLocale)) {
                $localeInfo = localeconv();

                $moneyPattern = '0' . $localeInfo['decimal_point'] . str_repeat('0', $localeInfo['frac_digits']);
                if ($localeInfo['thousands_sep']) {
                    $moneyPattern = '#' . $localeInfo['thousands_sep'] . '##' . $moneyPattern;
                }
                $sepSpace = !empty($localeInfo['p_sep_by_space']) ? str_repeat(' ', $localeInfo['p_sep_by_space']) : '';
                if (!empty($localeInfo['p_cs_precedes'])) {
                    $moneyPattern = '[$' . $localeInfo['currency_symbol'] . ']' . $sepSpace . $moneyPattern;
                }
                else {
                    $moneyPattern .= $sepSpace . '[$' . $localeInfo['currency_symbol'] . ']';
                }
                $aFormatSettings['formats']['@MONEY'] = $moneyPattern;
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
            }
            else {
                $file = str_replace('.utf-8/', '/', $file);
                if (is_file($file)) {
                    $includeFile = $file;
                }
            }

            $this->loadSettings($includeFile, $localeSettings);
            if (strpos($locale, '_')) {
                [$language, $country] = explode('_', $locale);
                $file = $dir . '/' . $language . '/settings.php';
                $this->loadSettings($file, $localeSettings);

                $file = $dir . '/' . $language . '/' . $country . '/settings.php';
                $this->loadSettings($file, $localeSettings);

                $file = str_replace('.utf-8/', '/', $file);
                $this->loadSettings($file, $localeSettings);
            }
        }
        if ($localeSettings) {
            $localeSettings = array_merge($aFormatSettings, $localeSettings);
        }
        else {
            $localeSettings = $aFormatSettings;
        }

        $this->style->setLocaleSettings($localeSettings);

        return $this;
    }

    public function setTitle(?string $title = ''): Excel
    {
        return $this->setMetaTitle($title);
    }

    public function setSubject(?string $subject = ''): Excel
    {
        return $this->setMetaSubject($subject);
    }

    public function setAuthor(?string $author = ''): Excel
    {
        return $this->setMetaAuthor($author);
    }

    public function setCompany(?string $company = ''): Excel
    {
        return $this->setMetaCompany($company);
    }

    public function setDescription(?string $description = ''): Excel
    {
        return $this->setMetaDescription($description);
    }

    public function setKeywords($keywords): Excel
    {
        return $this->setMetaKeywords($keywords);
    }

    /**
     * Set metadata 'title'
     *
     * @param string|null $title
     *
     * @return $this
     */
    public function setMetaTitle(?string $title = ''): Excel
    {
        return $this->setMetadata('title', $title);
    }

    /**
     * Set metadata 'subject'
     *
     * @param string|null $subject
     *
     * @return $this
     */
    public function setMetaSubject(?string $subject = ''): Excel
    {
        return $this->setMetadata('subject', $subject);
    }

    /**
     * Set metadata 'author'
     *
     * @param string|null $author
     *
     * @return $this
     */
    public function setMetaAuthor(?string $author = ''): Excel
    {
        return $this->setMetadata('author', $author);
    }

    /**
     * Set metadata 'company'
     *
     * @param string|null $company
     *
     * @return $this
     */
    public function setMetaCompany(?string $company = ''): Excel
    {
        return $this->setMetadata('company', $company);
    }

    /**
     * Set metadata 'description'
     *
     * @param string|null $description
     *
     * @return $this
     */
    public function setMetaDescription(?string $description = ''): Excel
    {
        return $this->setMetadata('description', $description);
    }

    /**
     * Set metadata 'keywords'
     *
     * @param mixed $keywords
     *
     * @return $this
     */
    public function setMetaKeywords($keywords): Excel
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
     * Set metadata
     *
     * @param $key
     * @param $value
     *
     * @return $this
     */
    public function setMetadata($key, $value): Excel
    {
        $this->metadata[$key] = $value;

        return $this;
    }

    /**
     * Get metadata
     *
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
    public function setRightToLeft(bool $isRightToLeft)
    {
        $this->isRightToLeft = (bool)$isRightToLeft;
    }

    /**
     * @return bool
     */
    public function isRightToLeft(): bool
    {
        return $this->isRightToLeft;
    }

    /**
     * Set default font options
     *
     * @param array $fontOptions
     *
     * @return $this
     */
    public function setDefaultFont(array $fontOptions): Excel
    {
        $this->style->setDefaultFont($fontOptions);

        return $this;
    }

    /**
     * Set default font name
     *
     * @param string $fontName
     *
     * @return $this
     */
    public function setDefaultFontName(string $fontName): Excel
    {
        $this->style->setDefaultFont(['font-name' => $fontName]);

        return $this;
    }

    /**
     * Set default style
     *
     * @param array $style
     *
     * @return $this
     */
    public function setDefaultStyle(array $style): Excel
    {
        $this->style->setDefaultStyle($style);

        return $this;
    }

    /**
     * Set active (default) sheet by case-insensitive name
     *
     * @param string $name
     *
     * @return $this
     */
    public function setActiveSheet(string $name): Excel
    {
        $tabIndex = 0;
        foreach ($this->sheets as $sheet) {
            if ($sheet->isName($name)) {
                $sheet->active = true;
                $this->bookViews[0]['activeTab'] = $tabIndex;
            }
            else {
                $sheet->active = false;
            }
            $tabIndex++;
        }
        return $this;
    }

    /**
     * @param int|string|array $rowRange
     *
     * @return int[]
     */
    public static function rowNumberRange($rowRange): array
    {
        if (is_array($rowRange)) {
            $result = [];
            foreach ($rowRange as $row) {
                $result[] = self::rowNumberRange($row);
            }
            $result = array_unique(array_filter(array_merge(...$result)));
            sort($result);
        }
        elseif (is_string($rowRange) && preg_match('/^(\d+):(\d+)$/', $rowRange, $m)) {
            $result = [];
            for ($rowNum = $m[1]; $rowNum <= $m[2]; $rowNum++) {
                $result[] = (int)$rowNum;
            }
        }
        elseif (is_numeric($rowRange)) {
            $result = [(int)$rowRange];
        }
        else {
            ExceptionAddress::throwNew('Row number (or row range) is incorrect');
        }

        return $result;
    }

    /**
     * @param int|string|array $rowRange
     *
     * @return int[]
     */
    public static function rowIndexRange($rowRange): array
    {
        $result = self::rowNumberRange($rowRange);

        foreach ($result as $key => $num) {
            $result[$key] = ($num > 0) ? $num - 1 : -1;
        }

        return $result;
    }

    /**
     * Convert letter to number (ONE based)
     *
     * @param string $colLetter
     *
     * @return int
     */
    public static function colNumber(string $colLetter): int
    {
        if ($colLetter && $colLetter[0] === '$') {
            $colLetter = substr($colLetter, 1);
        }
        return Helper::colNumber($colLetter);
    }

    /**
     * Convert letter to index (ZERO based)
     *
     * @param string $colLetter
     *
     * @return int
     */
    public static function colIndex(string $colLetter): int
    {
        if ($colLetter && $colLetter[0] === '$') {
            $colLetter = substr($colLetter, 1);
        }
        $colNumber = self::colNumber($colLetter);

        if ($colNumber > 0) {
            return $colNumber - 1;
        }

        return $colNumber;
    }

    /**
     * Convert letter range to array of numbers (ONE based)
     *
     * @param string|int|array $colLetter e.g.: 'B', 2, 'C:F', ['A', 'B', 'C']
     *
     * @return array
     */
    public static function colNumberRange($colLetter): array
    {
        $result = [];
        if (is_array($colLetter)) {
            foreach ($colLetter as $col) {
                $result[] = self::colNumber($col);
            }
        }
        elseif (is_string($colLetter)) {
            $letters = self::colLetterRange($colLetter);
            $result = self::colNumberRange($letters);
        }
        else {
            $col = self::colNumber($colLetter);
            if ($col > 0) {
                $result[] = $col;
            }
        }

        return $result;
    }

    /**
     * Convert letter range to array of numbers (ZERO based)
     *
     * @param string|int|array $colLetter e.g.: 'B', 2, 'C:F', ['A', 'B', 'C']
     *
     * @return array
     */
    public static function colIndexRange($colLetter): array
    {
        $result = self::colNumberRange($colLetter);

        foreach ($result as $key => $num) {
            $result[$key] = ($num > 0) ? $num - 1 : -1;
        }

        return $result;
    }

    /**
     * Convert values to letters array
     *
     * @param array|string $colKeys
     * @param int|null $baseNum 0 or 1
     *
     * @return array
     *
     * @example
     * $res = colLetterRange([0, 1, 2]);    // returns ['A', 'B', 'C']
     * $res = colLetterRange([1, 2, 3], 1); // returns ['A', 'B', 'C']
     * $res = colLetterRange('B, E, F');    // returns ['B', 'E', 'F']
     * $res = colLetterRange('B-E, F');     // returns ['B', 'C', 'D', 'E', 'F']
     * $res = colLetterRange('B1-E8');      // returns ['B', 'C', 'D', 'E']
     * $res = colLetterRange('B1:E8');      // returns ['B:E']
     */
    public static function colLetterRange($colKeys, ?int $baseNum = 0): array
    {
        if ($colKeys) {
            if (is_array($colKeys)) {
                $key = reset($colKeys);
                if (is_numeric($key)) {
                    $columns = [];
                    foreach ($colKeys as $key) {
                        $columns[] = Excel::colLetter($key + (1 - $baseNum));
                    }
                    return $columns;
                }
                else {
                    $columns = $colKeys;
                }
                return $columns;
            }
            elseif (is_string($colKeys)) {
                if (strpos($colKeys, ',')) {
                    $colKeys = array_map('trim', explode(',', $colKeys));
                    $columns = [];
                    foreach ($colKeys as $col) {
                        $columns[] = self::colLetterRange($col);
                    }

                    return array_merge(...$columns);
                }
                elseif (($p1 = strpos($colKeys, '-')) || ($p2 = strpos($colKeys, ':'))) {
                    [$num1, $num2] = $p1 ? explode('-', $colKeys) : explode(':', $colKeys);
                    $columns = [];
                    for ($colNum = self::colNumber($num1); $colNum <= self::colNumber($num2); $colNum++) {
                        $columns[] = self::colLetter($colNum);
                    }
                    return $columns;
                }
                elseif (preg_match('/^[1-9:]+$/', $colKeys)) {
                    [$num1, $num2] = explode(':', $colKeys);
                    return [self::colLetter($num1) . ':' . self::colLetter($num2)];
                }
                elseif (preg_match('/^[a-z1-9:]+$/i', $colKeys)) {
                    $colKeys = preg_replace('/\d+/', '', $colKeys);
                    return [strtoupper($colKeys)];
                }
            }
        }
        return [];
    }

    /**
     * Convert column number to letter
     *
     * @param int $colNumber ONE based
     *
     * @return string
     */
    public static function colLetter(int $colNumber): string
    {

        return Helper::colLetter($colNumber);
    }

    /**
     * @param array $data
     *
     * @return array
     */
    public static function colKeysToLetters(array $data): array
    {
        return array_combine(Excel::colLetterRange(array_keys($data)), array_values($data));
    }

    /**
     * @param array $data
     *
     * @return array
     */
    public static function colKeysToNumbers(array $data): array
    {
        return array_combine(Excel::colNumberRange(array_keys($data)), array_values($data));
    }

    /**
     * @param array $data
     *
     * @return array
     */
    public static function colKeysToIndexes(array $data, $offset = 0): array
    {
        $row = array_combine(Excel::colIndexRange(array_keys($data)), array_values($data));
        if ($offset) {
            $result = [];
            foreach ($row as $key => $val) {
                $result[$key + $offset] = $val;
            }
        }
        else {
            $result = $row;
        }

        return $result;
    }

    /**
     * Create cell address by row and col numbers
     *
     * @param int $rowNumber ONE based
     * @param int $colNumber ONE based
     * @param bool|null $absolute
     * @param bool|null $absoluteRow
     *
     * @return string Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
     *
     * @example
     * cellAddress(3, 3) => 'C3'
     * cellAddress(43, 27) => 'AA43'
     * cellAddress(43, 27, true) => '$AA$43'
     * cellAddress(43, 27, false, true) => 'AA$43'
     */
    public static function cellAddress(int $rowNumber, int $colNumber, ?bool $absolute = false, ?bool $absoluteRow = null): string
    {
        return Helper::cellAddress($rowNumber, $colNumber, $absolute, $absoluteRow);
    }

    /**
     * @param array|string $range
     * @param array|string $offset
     * @param bool|null $exception
     *
     * @return array|null
     */
    public static function rangeDimensionRelative($range, $offset, ?bool $exception = false): ?array
    {
        $colNum1 = $colNum2 = $rowNum1 = $rowNum2 = null;
        if (is_array($range)) {
            $sheetName = null;
            if (count($range) === 2) {
                // [[row1, col1], [row2, col2]] or [['row' => row1, 'col' => col1], ['row' => row2, 'col' => col2]]
                [$cell1, $cell2] = $range;
                if (is_array($cell1) && is_array($cell2)) {
                    if (isset($cell1['row'], $cell1['col'])) {
                        $rowNum1 = $cell1['row'];
                        $colNum1 = $cell1['col'];
                    }
                    else {
                        [$rowNum1, $colNum1] = $cell1;
                    }
                    if (isset($cell2['row'], $cell2['col'])) {
                        $rowNum2 = $cell1['row'];
                        $colNum2 = $cell1['col'];
                    }
                    else {
                        [$rowNum2, $colNum2] = $cell2;
                    }
                }
            }
            else {
                // [row1, col1, row2, col2]
                [$rowNum1, $colNum1, $rowNum2, $colNum2] = $range;
            }
        }
        else {
            if (strpos($range, '!')) {
                [$sheetName, $range] = explode('!', $range);
            }
            else {
                $sheetName = null;
            }
            $range = strtoupper($range);
            if (preg_match('/^(\$?[A-Z]+)(\$?\d+)(:(\$?[A-Z]+)(\$?\d+))?$/', $range, $matches)) {
                if (empty($matches[3])) {
                    $matches[4] = $matches[1];
                    $matches[5] = $matches[2];
                }
                $colNum1 = self::colNumber($matches[1]);
                $colNum2 = self::colNumber($matches[4]);
                $rowNum1 = ($matches[2] <= self::MAX_ROW) ? (int)$matches[2] : -1;
                $rowNum2 = ($matches[5] <= self::MAX_ROW) ? (int)$matches[5] : -1;
            }
        }
        if ($exception && ($colNum1 === null || $colNum2 === null || $rowNum1 === null || $rowNum2 === null)) {
            throw new Exception('Wrong range ' . print_r($range, true));
        }
        if (!empty($offset)) {
            $rowOffset1 = $colOffset1 = $rowOffset2 = $colOffset2 = null;
            if (is_array($offset)) {
                if (count($offset) === 4) {
                    [$rowOffset1, $colOffset1, $rowOffset2, $colOffset2] = $offset;
                }
                elseif (count($offset) === 2) {
                    // [row, col] or ['row' => row, 'col' => col]
                    if (isset($offset['row'], $offset['col'])) {
                        $rowOffset1 = $offset['row'];
                        $colOffset1 = $offset['col'];
                    }
                    else {
                        [$rowOffset2, $colOffset2] = $offset;
                    }
                }
            }
            else {
                // 'R1C1'
                [$rowOffset1, $colOffset1, $rowOffset2, $colOffset2] = self::rangeRelOffsets($offset);
            }
            if ($exception && ($rowOffset1 === null || $colOffset1 === null || $rowOffset2 === null || $colOffset2 === null)) {
                throw new Exception('Wrong offset of range ' . print_r($offset, true));
            }

            $rowNum1 += $rowOffset1;
            $colNum1 += $colOffset1;
            $rowNum2 += $rowOffset2;
            $colNum2 += $colOffset2;
        }
        if ($exception && ($colNum1 < 0 || $colNum2 < 0 || $rowNum1 < 0 || $rowNum2 < 0)) {
            throw new Exception('Wrong range ' . print_r($range, true));
        }

        if ($colNum1 > 0 && $colNum2 > 0 && $rowNum1 > 0 && $rowNum2 > 0) {
            // swap indexes if need
            if ($colNum2 < $colNum1) {
                $idx = $colNum1;
                $colNum1 = $colNum2;
                $colNum2 = $idx;
            }
            if ($rowNum2 < $rowNum1) {
                $idx = $rowNum1;
                $rowNum1 = $rowNum2;
                $rowNum2 = $idx;
            }
            $letter1 = Excel::colLetter($colNum1);
            $letter2 = Excel::colLetter($colNum2);
            $cell1 = $letter1 . $rowNum1;
            $cell2 = $letter2 . $rowNum2;
            $localRange = $cell1 . ':' . $cell2;
            $width = $colNum2 - $colNum1 + 1;
            $height = $rowNum2 - $rowNum1 + 1;
            if ($cell1 === $cell2) {
                $address = ($sheetName ? $sheetName . '!' : '') . '$' . $letter1 . '$' . $rowNum1;
            }
            else {
                $address = ($sheetName ? $sheetName . '!' : '') . '$' . $letter1 . '$' . $rowNum1 . ':$' . $letter2 . '$' . $rowNum2;
            }

            return [
                'absAddress' => $address,
                'range' => ($sheetName ? $sheetName . '!' : '') . $localRange,
                'sheet' => $sheetName,
                'cell1' => $cell1,
                'cell2' => $cell2,
                'localRange' => $localRange,
                'rowNum1' => $rowNum1,
                'colNum1' => $colNum1,
                'rowNum2' => $rowNum2,
                'colNum2' => $colNum2,
                'row' => $rowNum1,
                'col' => $colNum1,
                'rowIndex' => $rowNum1 - 1,
                'colIndex' => $colNum1 - 1,
                'width' => $width,
                'height' => $rowNum2 - $rowNum1 + 1,
                'cellCount' => $width * $height,
            ];
        }

        return null;
    }

    /**
     * @param array|string $range
     * @param bool|null $exception
     *
     * @return array|null
     */
    public static function rangeDimension($range, ?bool $exception = false): ?array
    {
        return self::rangeDimensionRelative($range, null, $exception);
    }

    /**
     * @param string $sheetName
     * @param string $address
     * @param bool|null $force
     *
     * @return string
     */
    public static function fullAddress(string $sheetName, string $address, ?bool $force = false): string
    {
        if (strpos($address, '!') && !$force) {
            return $address;
        }

        if (strpos($address, '!')) {
            [$sheetName, $cells] = explode('!', $address);
        }
        else {
            $cells = $address;
        }
        return '\'' . $sheetName . '\'!' . $cells;
    }

    /**
     * Return offsets by relative address (zero based)
     *
     * @param string $relAddress
     *
     * @return int[]
     */
    public static function rangeRelOffsets(string $relAddress): array
    {
        $rowOffset1 = $colOffset1 = $rowOffset2 = $colOffset2 = null;
        $offset = strtoupper($relAddress);
        if (preg_match('/^R\[?(-?\d+)?]?C\[?(-?\d+)?]?(:R\[?(-?\d+)?]?C\[?(-?\d+)?]?)?$/', $offset, $matches)) {
            $rowOffset1 = !empty($matches[1]) ? (int)$matches[1] : 0;
            $colOffset1 = !empty($matches[2]) ? (int)$matches[2] : 0;
            if (!empty($matches[3])) {
                $rowOffset2 = !empty($matches[4]) ? (int)$matches[4] : 0;
                $colOffset2 = !empty($matches[5]) ? (int)$matches[5] : 0;
            }
            else {
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
     * @param $pixels
     *
     * @return float|int
     */
    public static function pixelsToEMU($pixels)
    {
        return $pixels * self::EMU_PER_PIXEL;
    }

    /**
     * Create a password hash from a given string

     * This code is from phpoffice/phpspreadsheet ver.1.28 and based on the spec at:
     * https://interoperability.blob.core.windows.net/files/MS-OFFCRYPTO/[MS-OFFCRYPTO].pdf
     * 2.3.7.1 Binary Document Password Verifier Derivation Method 1
     *
     *
     * @param string $password Password to hash
     *
     * @return string
     */
    public static function hashPassword(string $password): string
    {
        if (strlen($password) > 255) {
            Exception::throwNew('Maximum allowed password length is %d characters', 255);
        }

        $verifier = 0;
        $passwordLen = strlen($password);
        $passwordArray = pack('c', $passwordLen) . $password;
        for ($i = $passwordLen; $i >= 0; --$i) {
            $intermediate1 = (($verifier & 0x4000) === 0) ? 0 : 1;
            $intermediate2 = 2 * $verifier;
            $intermediate2 = $intermediate2 & 0x7fff;
            $intermediate3 = $intermediate1 | $intermediate2;
            $verifier = $intermediate3 ^ ord($passwordArray[$i]);
        }
        $verifier ^= 0xCE4B;

        return strtoupper(dechex($verifier));
    }

    /**
     * @param string $string
     * @param bool|null $richText
     *
     * @return int
     */
    public function addSharedString(string $string, ?bool $richText = false): int
    {
        if (!isset($this->sharedStrings[$string])) {
            $this->sharedStrings[$string] = ['id' => $this->sharedStringsCount++, 'count' => 1, 'rich_text' => $richText];
        }
        else {
            $this->sharedStrings[$string]['count']++;
        }

        return $this->sharedStrings[$string]['id'];
    }

    /**
     * @return array
     */
    public function getSharedStrings(): array
    {
        return $this->sharedStrings;
    }

    /**
     * @return array
     */
    public function getThemes(): array
    {
        return $this->themes;
    }

    /**
     * @return Writer
     */
    public function getWriter()
    {
        return $this->writer;
    }

    /**
     * @return string
     */
    public function getDefaultSheetName(): string
    {
        return 'Sheet' . (++$this->maxSheetIndex);
    }

    /**
     * @param string|null $sheetName
     *
     * @return Sheet
     */
    public function makeSheet(?string $sheetName = null): Sheet
    {
        if ($sheetName === null) {
            $sheetName = $this->getDefaultSheetName();
        }
        $key = mb_strtolower($sheetName);
        if (!isset($this->sheets[$key])) {
            $sheet = static::createSheet($sheetName);
            $sheet->localSheetId = count($this->sheets);
            $this->sheets[$key] = $sheet;

            $this->sheets[$key]->excel = $this;
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
     * Returns sheet by number or name of sheet.
     * Return the first sheet if number or name omitted
     *
     * @param int|string|null $index - number or name of sheet
     *
     * @return Sheet|null
     */
    public function sheet($index = null): ?Sheet
    {
        if (null === $index) {
            return reset($this->sheets);
        }

        if (is_int($index)) {
            $keys = array_keys($this->sheets);
            if (isset($keys[--$index])) {
                $key = $keys[$index];
            }
            else {
                // index not found
                throw  new Exception('Sheet #' . ++$index . ' not found');
            }
        }
        else {
            $key = mb_strtolower($index);
            if (!isset($this->sheets[$key])) {
                throw  new Exception('Sheet "' . $index . '" not found');
            }
        }
        return $this->sheets[$key] ?? null;
    }

    /**
     * Alias of sheet()
     *
     * @param int|string|null $index - number or name of sheet
     *
     * @return Sheet|null
     */
    public function getSheet($index = null): ?Sheet
    {
        return $this->sheet($index);
    }

    /**
     * Removes sheet by index or name of sheet.
     * Removes the first sheet of index omitted
     *
     * @param int|string|null $index
     *
     * @return $this
     */
    public function removeSheet($index = null): Excel
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
        }
        else {
            $key = mb_strtolower($index);
            if (!isset($this->sheets[$key])) {
                throw  new Exception('Sheet "' . $index . '" not found');
            }
            unset($this->sheets[$key]);
        }
        $localSheetId = 0;
        foreach ($this->sheets as $sheet) {
            $sheet->localSheetId = $localSheetId++;
        }

        return $this;
    }

    /**
     * Returns all sheets
     *
     * @return Sheet[]
     */
    public function getSheets(): array
    {
        return $this->sheets;
    }

    /**
     * @param string $range
     * @param string $name
     *
     * @return $this
     */
    public function addNamedRange(string $range, string $name): Excel
    {
        if (strpos($range, '!')) {
            [$sheetName, $range] = explode('!', $range);
            $sheet = $this->getSheet($sheetName);
            if ($sheet) {
                $sheet->addNamedRange($range, $name);

                return $this;
            }
        }
        ExceptionRangeName::throwNew('Sheet name not defined in range address');
    }

    /**
     * @param string $name
     * @param string $range
     * @param array|null $attributes
     *
     * @return $this
     */
    public function addDefinedName(string $name, string $range, ?array $attributes = []): Excel
    {
        $attributes = array_replace(['name' => Writer::xmlSpecialChars($name)], $attributes);
        if ($name === '_xlnm.Print_Area' && isset($attributes['localSheetId'])) {
            // add print area
            foreach ($this->definedNames as $key => $definedName) {
                if ($definedName['_attr']['name'] === $name && isset($definedName['localSheetId']) && $definedName['localSheetId'] === $attributes['localSheetId']) {
                    $this->definedNames[$key]['_value'] .= $range;
                    return $this;
                }
            }
            $this->definedNames[] = [
                '_value' => $range,
                '_attr' => $attributes,
            ];
        }
        elseif ($name === '_xlnm.Print_Titles' && isset($attributes['localSheetId'])) {
            // set print title
            foreach ($this->definedNames as $key => $definedName) {
                if ($definedName['_attr']['name'] === $name && isset($definedName['localSheetId']) && $definedName['localSheetId'] === $attributes['localSheetId']) {
                    unset($this->definedNames[$key]);
                }
            }
            $this->definedNames[] = [
                '_value' => $range,
                '_attr' => $attributes,
            ];
        }
        else {
            $this->definedNames[$name] = [
                '_value' => $range,
                '_attr' => $attributes,
            ];
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getDefinedNames(): array
    {
        $result = $this->definedNames;
        foreach ($this->sheets as $sheet) {
            if ($sheet->absoluteAutoFilter) {
                $filterRange = $sheet->absoluteAutoFilter . ':' . Excel::cellAddress($sheet->rowCountWritten, $sheet->colCountWritten, true);
                $fullAddress = "'" . $sheet->sanitizedSheetName . "'!" . $filterRange;
                $result[] = [
                    '_value' => $fullAddress,
                    '_attr' => ['name' => '_xlnm._FilterDatabase', 'localSheetId' => $sheet->localSheetId, 'hidden' => '1'],
                ];
            }
        }
        return $result;
    }

    /**
     * @param string $imageBlob
     *
     * @return array
     */
    protected function getImageInfo(string $imageBlob): array
    {
        if (substr($imageBlob, 0, 4) === '<svg'
            && isset(self::$availableImageTypes['svg'])) {
            if (preg_match("#^<svg\s([^>]+)>#si", $imageBlob, $s)
                && preg_match("#width\s*=\s*[\"'](\d+)[\"']#si", $s[1], $w)
                && preg_match("#height\s*=\s*[\"'](\d+)[\"']#si", $s[1], $h)) {
                $result = [
                    'extension' => 'svg',
                    'width' => (int)$w[1],
                    'height' => (int)$h[1],
                    'mime' => self::$availableImageTypes['svg'],
                ];
            }
            elseif (preg_match("#viewbox=[\"']([\d\s\-.]+)[\"']#si", $imageBlob, $m)) {
                $d = explode(' ', $m[1]);
                if (isset($d[2], $d[3])) {
                    $result = [
                        'extension' => 'svg',
                        'width' => (int)$d[2],
                        'height' => (int)$d[3],
                        'mime' => self::$availableImageTypes['svg'],
                    ];
                }
            }
        }
        else {
            $info = getimagesizefromstring($imageBlob);
            if (!empty($info['mime'])) {
                $extension = array_search($info['mime'], self::$availableImageTypes);
                $result = [
                    'extension' => $extension,
                    'width' => $info[0],
                    'height' => $info[1],
                    'mime' => $info['mime'],
                ];
            }
        }

        return $result ?? [];
    }

    /**
     * @param string $imageFile URL, local path or image string in base64
     *
     * @return array|null
     */
    public function loadImageFile(string $imageFile): ?array
    {
        $imageBlob = false;
        if (preg_match('/^data:image\/(\w+);base64,/', $imageFile, $matches)) {
            // $imageType = $matches[1];
            $base64Image = substr($imageFile, strpos($imageFile, ',') + 1);
            $imageBlob = base64_decode($base64Image);

        }
        elseif (preg_match('#^https?://.+#i', $imageFile)) {
            $response = file_get_contents(
                $imageFile,
                false,
                stream_context_create([
                    'http' => [
                        'ignore_errors' => true,
                    ],
                ])
            );
            if (isset($http_response_header[0])) {
                if (preg_match('#\s404\s#', $http_response_header[0])) {
                    ExceptionFile::throwNew('Image file "%s" does not exist', $imageFile);
                }
                elseif (preg_match('#\s200\s#', $http_response_header[0])) {
                    $imageBlob = $response;
                }
            }
        }
        elseif (preg_match('#^\w+://.+#', $imageFile)) {
            $imageBlob = file_get_contents($imageFile);
        }
        else {
            if (!is_file($imageFile)) {
                ExceptionFile::throwNew('Image file "%s" does not exist', $imageFile);
            }
            $imageBlob = file_get_contents($imageFile);
        }
        if ($imageBlob === false) {
            ExceptionFile::throwNew('Cannot read file "%s"', $imageFile);
        }
        if (!$imageBlob) {
            ExceptionFile::throwNew('Image file "%s" is empty', $imageFile);
        }
        $imageHash = sha1($imageBlob);

        if (!isset($this->media['images'][$imageHash])) {
            $info = $this->getImageInfo($imageBlob);
            if ($info) {
                $imageId = (empty($this->media['images']) ? 1 : count($this->media['images']) + 1);
                $name = 'image' . $imageId . '.' . $info['extension'];
                $fileName = $this->writer->makeTempFile(null, 'xl/media/' . $name);
                if ($fileName && file_put_contents($fileName, $imageBlob)) {
                    $this->media['images'][$imageHash] =  [
                        'filename' => $fileName,
                        'original' => basename($imageFile),
                        'width' => $info['width'],
                        'height' => $info['height'],
                        'name' => $name,
                        'id' => $imageId,
                        'hash' => $imageHash,
                        'extension' => $info['extension'],
                        'mime_type' => $info['mime'],
                    ];
                }
            }
        }
        if (!isset($this->media['images'][$imageHash])) {
            ExceptionFile::throwNew('File "%s" is not image', $imageFile);
        }

        return $this->media['images'][$imageHash];
    }

    /**
     * @return array
     */
    public function getImageFiles(): array
    {
        return $this->media['images'] ?? [];
    }

    /**
     * Protect workbook
     *
     * @param string|null $password
     *
     * @return $this
     */
    public function protect(?string $password = null): Excel
    {
        $this->protection = [
            'lockStructure' => '1',
            'lockWindows' => '1',
            'lockRevision' => 'false',
        ];
        if ($password) {
            $this->protection['workbookPassword'] = Excel::hashPassword($password);
        }

        return $this;
    }

    /**
     * Unprotect workbook
     *
     * @return $this
     */
    public function unprotect(): Excel
    {
        $this->protection = [];

        return $this;
    }

    /**
     * Sets default filename for saving
     *
     * @param string $fileName
     *
     * @return $this
     */
    public function setFileName(string $fileName): Excel
    {
        if (!pathinfo($fileName, PATHINFO_EXTENSION)) {
            $fileName .= '.xlsx';
        }
        $this->fileName = $fileName;

        return $this;
    }

    /**
     * Returns default filename
     *
     * @return string
     */
    public function getFileName(): string
    {
        return $this->fileName;
    }

    /**
     * @return array|array[]
     */
    public function getBookViews(): array
    {
        return $this->bookViews;
    }

    /**
     * @return array
     */
    public function getProtection(): array
    {
        return $this->protection;
    }

    /**
     * Save generated XLSX-file
     *
     * @param string|null $fileName
     * @param bool|null $overWrite
     *
     * @return bool
     */
    public function save(?string $fileName = null, ?bool $overWrite = true): bool
    {
        if ($this->saved) {
            ExceptionFile::throwNew('The workbook is already saved');
        }

        if (!$fileName && $this->fileName) {
            $fileName = $this->fileName;
        }

        if ($this->writer->saveToFile($fileName, $overWrite, $this->getMetadata())) {
            $this->saved = true;
        }
        $this->writer->removeFiles();

        return $this->saved;
    }

    /**
     * Download generated file to client (send to browser)
     *
     * @param string|null $name
     */
    public function download(?string $name = null)
    {
        $tmpFile = $this->writer->makeTempFileName(uniqid('xlsx_writer_'));
        $this->save($tmpFile);
        if (!$name) {
            $name = basename($tmpFile) . '.xlsx';
        }
        else {
            $name = basename($name);
            if (strtolower(pathinfo($name, PATHINFO_EXTENSION)) !== 'xlsx') {
                $name .= '.xlsx';
            }
        }

        header('Cache-Control: max-age=0');
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . $name . '"');

        readfile($tmpFile);
        unlink($tmpFile);
    }

    /**
     * Alias of download()
     *
     * @param string|null $name
     *
     * @return void
     */
    public function output(?string $name = null)
    {
        $this->download($name);
    }

    public function getDefaultStyle(): array
    {
        return $this->style->getDefaultStyle();
    }

    public function getHyperlinkStyle(): array
    {
        return $this->style->getHyperlinkStyle();
    }

    public function getDefaultFormatStyles(): array
    {
        return $this->style->getDefaultFormatStyles();
    }

    public function getStyleLocaleSettings(): array
    {
        return $this->style->getLocaleSettings();
    }

    public function getStyleFonts(): array
    {
        return $this->style->getStyleFonts();
    }

    public function getStyleFills(): array
    {
        return $this->style->getStyleFills();
    }

    public function getStyleBorders(): array
    {
        return $this->style->getStyleBorders();
    }

    public function getStyleCellXfs(): array
    {
        return $this->style->getStyleCellXfs();
    }

    public function getStyleIndexedColors(): array
    {
        return $this->style->getIndexedColors();
    }


    public function getStyleNumberFormats(): array
    {
        return $this->style->_getNumberFormats();
    }

    public function addStyle($cellStyle, &$resultStyle = []): int
    {
        return $this->style->addStyle($cellStyle, $resultStyle);
    }

    public function addStyleDxfs($style, &$resultStyle = []): int
    {
        return $this->style->addDxfs($style, $resultStyle);
    }

    public function getStyleDxfs(): array
    {
        return $this->style->getStyleDxfs();
    }

}

 // EOF
