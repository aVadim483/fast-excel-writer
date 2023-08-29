<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;
use avadim\FastExcelWriter\Exception\ExceptionFile;

/**
 * Class Writer
 *
 * @package avadim\FastExcelWriter
 */
class Writer
{
    // These constants from phpoffice/phpspreadsheet ver.1.28
    const XLFNREGEXP = '/(?:_xlfn\.)?((?:_xlws\.)?('
    // functions added with Excel 2010
    . 'beta[.]dist'
    . '|beta[.]inv'
    . '|binom[.]dist'
    . '|binom[.]inv'
    . '|ceiling[.]precise'
    . '|chisq[.]dist'
    . '|chisq[.]dist[.]rt'
    . '|chisq[.]inv'
    . '|chisq[.]inv[.]rt'
    . '|chisq[.]test'
    . '|confidence[.]norm'
    . '|confidence[.]t'
    . '|covariance[.]p'
    . '|covariance[.]s'
    . '|erf[.]precise'
    . '|erfc[.]precise'
    . '|expon[.]dist'
    . '|f[.]dist'
    . '|f[.]dist[.]rt'
    . '|f[.]inv'
    . '|f[.]inv[.]rt'
    . '|f[.]test'
    . '|floor[.]precise'
    . '|gamma[.]dist'
    . '|gamma[.]inv'
    . '|gammaln[.]precise'
    . '|lognorm[.]dist'
    . '|lognorm[.]inv'
    . '|mode[.]mult'
    . '|mode[.]sngl'
    . '|negbinom[.]dist'
    . '|networkdays[.]intl'
    . '|norm[.]dist'
    . '|norm[.]inv'
    . '|norm[.]s[.]dist'
    . '|norm[.]s[.]inv'
    . '|percentile[.]exc'
    . '|percentile[.]inc'
    . '|percentrank[.]exc'
    . '|percentrank[.]inc'
    . '|poisson[.]dist'
    . '|quartile[.]exc'
    . '|quartile[.]inc'
    . '|rank[.]avg'
    . '|rank[.]eq'
    . '|stdev[.]p'
    . '|stdev[.]s'
    . '|t[.]dist'
    . '|t[.]dist[.]2t'
    . '|t[.]dist[.]rt'
    . '|t[.]inv'
    . '|t[.]inv[.]2t'
    . '|t[.]test'
    . '|var[.]p'
    . '|var[.]s'
    . '|weibull[.]dist'
    . '|z[.]test'
    // functions added with Excel 2013
    . '|acot'
    . '|acoth'
    . '|arabic'
    . '|averageifs'
    . '|binom[.]dist[.]range'
    . '|bitand'
    . '|bitlshift'
    . '|bitor'
    . '|bitrshift'
    . '|bitxor'
    . '|ceiling[.]math'
    . '|combina'
    . '|cot'
    . '|coth'
    . '|csc'
    . '|csch'
    . '|days'
    . '|dbcs'
    . '|decimal'
    . '|encodeurl'
    . '|filterxml'
    . '|floor[.]math'
    . '|formulatext'
    . '|gamma'
    . '|gauss'
    . '|ifna'
    . '|imcosh'
    . '|imcot'
    . '|imcsc'
    . '|imcsch'
    . '|imsec'
    . '|imsech'
    . '|imsinh'
    . '|imtan'
    . '|isformula'
    . '|iso[.]ceiling'
    . '|isoweeknum'
    . '|munit'
    . '|numbervalue'
    . '|pduration'
    . '|permutationa'
    . '|phi'
    . '|rri'
    . '|sec'
    . '|sech'
    . '|sheet'
    . '|sheets'
    . '|skew[.]p'
    . '|unichar'
    . '|unicode'
    . '|webservice'
    . '|xor'
    // functions added with Excel 2016
    . '|forecast[.]et2'
    . '|forecast[.]ets[.]confint'
    . '|forecast[.]ets[.]seasonality'
    . '|forecast[.]ets[.]stat'
    . '|forecast[.]linear'
    . '|switch'
    // functions added with Excel 2019
    . '|concat'
    . '|countifs'
    . '|ifs'
    . '|maxifs'
    . '|minifs'
    . '|sumifs'
    . '|textjoin'
    // functions added with Excel 365
    . '|filter'
    . '|randarray'
    . '|anchorarray'
    . '|sequence'
    . '|sort'
    . '|sortby'
    . '|unique'
    . '|xlookup'
    . '|xmatch'
    . '|arraytotext'
    . '|call'
    . '|let'
    . '|lambda'
    . '|single'
    . '|register[.]id'
    . '|textafter'
    . '|textbefore'
    . '|textsplit'
    . '|valuetotext'
    . '))\s*\(/Umui';

    const XLWSREGEXP = '/(?<!_xlws\.)('
    // functions added with Excel 365
    . 'filter'
    . '|sort'
    . ')\s*\(/mui';

    /** @var Excel */
    protected $excel;

    /** @var array */
    protected array $tempFiles = [];

    /** @var string */
    protected $tempDir = '';

    /**
     * Writer constructor
     *
     * @param array|null $options ;
     */
    public function __construct(?array $options = [])
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
    }

    /**
     * Writer destructor
     */
    public function __destruct()
    {
        if (!empty($this->tempFiles)) {
            foreach ($this->tempFiles as $tempFile) {
                if (is_file($tempFile)) {
                    @unlink($tempFile);
                }
            }
        }
    }

    /**
     * @param string $fileName
     *
     * @return WriterBuffer
     */
    public static function makeWriteBuffer(string $fileName)
    {
        return new WriterBuffer($fileName);
    }

    /**
     * @param Excel $excel
     */
    public function setExcel(Excel $excel)
    {
        $this->excel = $excel;
    }

    /**
     * @param string|null $tempDir
     */
    public function setTempDir(?string $tempDir = '')
    {
        $this->tempDir = $tempDir;
    }

    /**
     * @return bool|string
     */
    public function tempFilename($localName = null)
    {
        $tempPrefix = 'xlsx_writer_';
        if (!$this->tempDir) {
            $tempDir = sys_get_temp_dir();
            $filename = tempnam($tempDir, $tempPrefix);
            if (!$filename) {
                $filename = tempnam(getcwd(), $tempPrefix);
            }
        }
        else {
            $filename = tempnam($this->tempDir, $tempPrefix);
        }
        if ($filename) {
            if ($localName) {
                $this->tempFiles[$localName] = $filename;
            }
            else {
                $this->tempFiles[] = $filename;
            }
        }

        return $filename;
    }

    /**
     * @param string $fileName
     * @param bool|null $overWrite
     * @param array|null $metadata
     *
     * @return bool
     */
    public function saveToFile(string $fileName, ?bool $overWrite = true, ?array $metadata = []): bool
    {
        $relationShips = [
            'default' => [
                'rels' => ['content_type' => 'application/vnd.openxmlformats-package.relationships+xml'],
                'xml' => ['content_type' => 'application/xml'],
            ],
            'override' => [
                'docProps/core.xml' => [
                    'content_type' => 'application/vnd.openxmlformats-package.core-properties+xml',
                    'rel' => 'root',
                    'schema' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                ],
                'docProps/app.xml' => [
                    'content_type' => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
                    'rel' => 'root',
                    'schema' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
                ],
                'xl/workbook.xml' => [
                    'content_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    'rel' => 'root',
                    'schema' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                ],
            ],
            'rel_id' => ['workbook' => 0],
        ];

        if (!is_dir(dirname($fileName))) {
            ExceptionFile::throwNew('Directory "%s" for output file is not exist', dirname($fileName));
        }
        if (file_exists($fileName)) {
            if ($overWrite && is_writable($fileName)) {
                @unlink($fileName); //if the zip already exists, remove it
            }
            else {
                ExceptionFile::throwNew('File "%s" is not writeable', $fileName);
            }
        }

        $sheets = $this->excel->getSheets();//$this->writeSheetDataBegin($sheet);
        if (empty($sheets)) {
            ExceptionFile::throwNew('No worksheets defined');
        }
        $zip = new \ZipArchive();
        if (!$zip->open($fileName, \ZIPARCHIVE::CREATE)) {
            ExceptionFile::throwNew('Unable to create zip "%s"', $fileName);
        }

        // add sheets
        $zip->addEmptyDir('xl/worksheets/');

        // 'xl/worksheets/sheet{%n}.xml' -- workbook
        // 'xl/comments{%n}.xml'
        // 'xl/drawings/vmlDrawing{%n}.vml'
        // 'xl/drawings/drawing{%n}.xml'
        // xl/drawings/_rels/drawing{%n}.xml.rels
        $this->_writeSheetsFiles($zip, $sheets, $relationShips);

        $zip->addFile($this->_writeStylesXML(), 'xl/styles.xml');
        $relationShips['override']['xl/styles.xml'] = [
            'content_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
            'rel' => 'workbook',
            'schema' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            'r_id' => 'rId' . (++$relationShips['rel_id']['workbook']),
        ];

        // 'xl/media/image{%n}.jpg'
        $this->_writeMediaFiles($zip, $relationShips);

        // 'xl/sharedStrings.xml' -- workbook
        $this->_writeSharedStrings($zip, $relationShips);

        // 'xl/theme/theme{%n}.xml' -- workbook
        $this->_writeThemesFiles($zip, $relationShips);

        $zip->addFromString('xl/workbook.xml', $this->_buildWorkbookXML($sheets));
        $zip->addEmptyDir('xl/_rels/');
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->_buildWorkbookRelsXML($relationShips));

        $zip->addEmptyDir('docProps/');
        $zip->addFromString('docProps/app.xml', $this->_buildAppXML($metadata));
        $zip->addFromString('docProps/core.xml', $this->_buildCoreXML($metadata));

        $zip->addEmptyDir('_rels/');
        $zip->addFromString('_rels/.rels', $this->_buildRelationshipsXML($relationShips));

        $zip->addFromString('[Content_Types].xml', $this->_buildContentTypesXML($relationShips));

        $zip->close();

        return true;
    }

    /**
     * @param \ZipArchive $zip
     * @param Sheet[] $sheets
     * @param array $relationShips
     *
     * @return true
     */
    protected function _writeSheetsFiles(\ZipArchive $zip, array $sheets, array &$relationShips): bool
    {
        $dirRels = false;

        foreach ($sheets as $sheet) {
            if (!$sheet->open) {
                // open and write areas
                $this->writeSheetDataBegin($sheet);
            }
            $this->writeSheetDataEnd($sheet);//making sure all footers have been written
            $sheet->relId = 'rId' . (++$relationShips['rel_id']['workbook']);
            $relationShips['override']['xl/worksheets/' . $sheet->xmlName] = [
                'content_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                'rel' => 'workbook',
                'schema' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                'r_id' => $sheet->relId,
            ];

            $zip->addFile($sheet->fileTempName, 'xl/worksheets/' . $sheet->xmlName);
            $xmlContent = $sheet->getXmlRels();
            if ($xmlContent) {
                if (!$dirRels) {
                    $zip->addEmptyDir('xl/worksheets/_rels/');
                    $dirRels = true;
                }
                $entry = 'xl/worksheets/_rels/' . $sheet->xmlName . '.rels';
                $zip->addFromString($entry, $xmlContent);
            }

            $commentList = $sheet->getNotes();
            if ($commentList) {
                // 'xl/comments{%n}.xml'
                $entry = 'xl/comments' . $sheet->index . '.xml';
                $this->_writeCommentsFile($zip, $entry, $commentList, $relationShips);

                // 'xl/drawings/vmlDrawing{%n}.vml'
                $entry = 'xl/drawings/vmlDrawing' . $sheet->index . '.vml';
                $this->_writeCommentOldStyleShape($zip, $entry, $commentList, $relationShips);
                if (empty($relationShips['default']['vml'])) {
                    $relationShips['default']['vml'] = [
                        'content_type' => 'application/vnd.openxmlformats-officedocument.vmlDrawing',
                    ];
                }
            }
            $imageList = $sheet->getImages();
            if ($imageList) {
                // 'xl/drawings/drawing{%n}.xml'
                // 'xl/drawings/_rels/drawing{%n}.xml.rels'
                $entry = 'xl/drawings/drawing' . $sheet->index . '.xml';
                $this->_writeDrawingFile($zip, $entry, $imageList, $relationShips);
                if (empty($relationShips['override'][$entry])) {
                    $relationShips['override'][$entry] = [
                        'content_type' => 'application/vnd.openxmlformats-officedocument.drawing+xml',
                    ];
                }
            }
        }

        return true;
    }

    /**
     * @param \ZipArchive $zip
     * @param array $relationShips
     *
     * @return string|null
     */
    protected function _writeThemesFiles(\ZipArchive $zip, array &$relationShips): ?string
    {
        $error = null;

        $themes = $this->excel->getThemes();
        if ($themes) {
            $zip->addEmptyDir('xl/theme/');
            foreach ($themes as $num => $theme) {
                $file = 'xl/theme/theme' . $num . '.xml';
                if ($zip->addFromString($file, $this->_buildThemeXML())) {
                    $relationShips['override'][$file] = [
                        'content_type' => 'application/vnd.openxmlformats-officedocument.theme+xml',
                        'rel' => 'workbook',
                        'schema' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
                        'r_id' => 'rId' . (++$relationShips['rel_id']['workbook']),
                    ];
                }
                else {
                    $error = $zip->getStatusString();
                }
            }
        }

        return $error;
    }

    /**
     * @param \ZipArchive $zip
     * @param array $relationShips
     *
     * @return string|null
     */
    protected function _writeSharedStrings(\ZipArchive $zip, array &$relationShips): ?string
    {
        $error = null;
        $sharedStrings = $this->excel->getSharedStrings();
        if ($sharedStrings) {
            $uniqueCount = count($sharedStrings);
            $count = 0;
            $result = '';
            foreach ($sharedStrings as $string => $info) {
                $count += $info['count'];
                $result .= '<si><t>' . $string . '</t></si>';
            }
            $xmlSharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                . '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $count . '" uniqueCount="' . $uniqueCount . '">'
                . $result
                . '</sst>';

            $file = 'xl/sharedStrings.xml';

            if ($zip->addFromString($file, $xmlSharedStrings)) {
                $relationShips['override'][$file] = [
                    'content_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
                    'rel' => 'workbook',
                    'schema' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                    'r_id' => 'rId' . (++$relationShips['rel_id']['workbook']),
                ];
            }
            else {
                $error = $zip->getStatusString();
            }
        }

        return $error;
    }

    /**
     * @param \ZipArchive $zip
     * @param array $relationShips
     *
     * @return string|null
     */
    protected function _writeMediaFiles(\ZipArchive $zip, array &$relationShips): ?string
    {
        $error = null;

        $imageList = $this->excel->getImageFiles();
        if ($imageList) {
            foreach ($imageList as $image) {
                if ($zip->addFile($image['filename'], 'xl/media/' . $image['name'])) {
                    if (empty($relationShips['default'][$image['name']])) {
                        $relationShips['default'][$image['extension']] = [
                            'content_type' => $image['mime_type'],
                        ];
                    }
                }
                else {
                    $error = $zip->getStatusString();
                }
            }
        }

        return $error;
    }

    /**
     * @param \ZipArchive $zip
     * @param string $entry
     * @param array $commentList
     * @param array $relationShips
     *
     * @return string|null
     */
    protected function _writeCommentsFile(\ZipArchive $zip, string $entry, array $commentList, array &$relationShips): ?string
    {
        $error = null;

        if ($commentList) {
            $xmlString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            $xmlString .= '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">';
            $xmlString .= '<authors><author/></authors>';
            $xmlString .= '<commentList>';
            foreach ($commentList as $comment) {
                $xmlString .= '<comment ref="' . $comment['cell'] . '" authorId="0"  shapeId="0" xr:uid="{' . Excel::generateUuid() . '}">';
                $xmlString .= '<text><r>';
                $xmlString .= '<t xml:space="preserve">' . $comment['text'] . '</t>';
                $xmlString .= '</r></text>';
                $xmlString .= '</comment>';
            }
            $xmlString .= '</commentList>';
            $xmlString .= '</comments>';

            if ($zip->addFromString($entry, $xmlString)) {
                $relationShips['override'][$entry] = [
                    'content_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
                ];
            }
            else {
                $error = $zip->getStatusString();
            }
        }

        return $error;
    }

    /**
     * @param \ZipArchive $zip
     * @param string $entry
     * @param array $commentList
     * @param array $relationShips
     *
     * @return string|null
     */
    protected function _writeCommentOldStyleShape(\ZipArchive $zip, string $entry, array $commentList, array &$relationShips): ?string
    {
        $error = null;

        $drawingCnt = 0;
        if ($commentList) {
            $xmlDrawing = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"  xmlns:x="urn:schemas-microsoft-com:office:excel">';
            $xmlDrawing .= '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe" fillcolor="#0000FF">';
            $xmlDrawing .= '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>';
            $xmlDrawing .= '<v:fill color="#0000FF"/>';
            $xmlDrawing .= '</v:shapetype>';
            foreach ($commentList as $comment) {
                $id = 1024 + (++$drawingCnt);
                $style = 'position:absolute;margin-left:' . $comment['style']['margin_left'] . ';margin-top:'  . $comment['style']['margin_top'] . ';'
                    . 'width:'  . $comment['style']['width'] . ';height:'  . $comment['style']['height'] . ';z-index:1;';
                if (empty($comment['style']['show'])) {
                    $style .= 'visibility:hidden';
                }
                $xmlDrawing .= '<v:shape id="_x0000_s' . $id . '" type="#_x0000_t202" style="' . $style . '" fillcolor="' . $comment['style']['fill_color'] . '" o:insetmode="auto">';
                $xmlDrawing .= '<v:fill color2="' . $comment['style']['fill_color'] . '"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>';
                $xmlDrawing .= '<v:textbox style="mso-direction-alt:auto">';
                $xmlDrawing .= '<div style="text-align:' . ($this->excel->isRightToLeft() ? 'right' : 'left') . '"/>';
                $xmlDrawing .= '</v:textbox>';
                $xmlDrawing .= '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>';
                $xmlDrawing .= '<x:AutoFill>False</x:AutoFill>';
                $xmlDrawing .= '<x:Row>' . $comment['row_index'] . '</x:Row><x:Column>' . $comment['col_index'] . '</x:Column>';
                $xmlDrawing .= '</x:ClientData></v:shape>';
            }
            $xmlDrawing .= '</xml>';

            if (!$zip->addFromString($entry, $xmlDrawing)) {
                $error = $zip->getStatusString();
            }
        }

        return $error;
    }

    /**
     * @param \ZipArchive $zip
     * @param string $entry
     * @param array $imageList
     * @param array $relationShips
     *
     * @return string|null
     */
    protected function _writeDrawingFile(\ZipArchive $zip, string $entry, array $imageList, array &$relationShips): ?string
    {
        $error = null;

        $relations = [];
        $xmlDrawingString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $xmlDrawingString .= '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';

        foreach ($imageList as $image) {
            $objectId = $image['id'];
            $rId = 'rId' . $image['id'];
            $baseName = $image['original'];
            $width = $image['width'] * Excel::EMU_PER_PIXEL;
            $height = $image['height'] * Excel::EMU_PER_PIXEL;

            $xmlDrawingString .= <<<EOD
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>{$image['col_index']}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>{$image['row_index']}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="{$width}" cy="{$height}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="{$objectId}" name="{$baseName}"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{$rId}"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm rot="0"/>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
EOD;
            $relations[$image['name']] = [
                'r_id' => $rId,
                'target' => '../media/' . $image['name'],
                'schema' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            ];
        }
        $xmlDrawingString .= '</xdr:wsDr>';
        if ($zip->addFromString($entry, $xmlDrawingString) && $relations) {
            $xmlRelations = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            $xmlRelations .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
            foreach ($relations as $relData) {
                $xmlRelations .= '<Relationship Id="' . $relData['r_id'] . '" Type="' . $relData['schema'] . '" Target="' . $relData['target'] . '"/>';
            }
            $xmlRelations .= '</Relationships>';
            $entryRel = str_replace('xl/drawings/drawing', 'xl/drawings/_rels/drawing', $entry) . '.rels'; //xl/drawings/_rels/drawing' . $sheet->index . '.xml.rels';
            if (!$zip->addFromString($entryRel, $xmlRelations)) {
                $error = $zip->getStatusString();
            }
        }
        else {
            $error = $zip->getStatusString();
        }

        return $error;
    }

    /**
     * @param Sheet $sheet
     *
     * @return WriterBuffer
     */
    protected function _writeSheetHead(Sheet $sheet): WriterBuffer
    {
        $fileWriter = self::makeWriteBuffer($this->tempFilename('xl/worksheets/' . $sheet->xmlName . '-head'));

        $fileWriter->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $xmlnsLinks = [
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
            'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"',
            //'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3"',
            //'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"',
            'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"',
            //'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"',
            //'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"',
            'xr:uid="{' . Excel::generateUuid() . '}"',
        ];
        $xmlns = implode(' ', $xmlnsLinks);
        $fileWriter->write('<worksheet ' . $xmlns . '>');

        if ($sheet->getPageFit()) {
            $fileWriter->write('<sheetPr>');
            $fileWriter->write('<pageSetUpPr fitToPage="1"/>');
            $fileWriter->write('</sheetPr>');
        }
        $minCell = $sheet->minCell();
        $maxCell = $sheet->maxCell();
        if ($minCell === $maxCell) {
            $fileWriter->write('<dimension ref="' . $minCell . '"/>');
        }
        else {
            $fileWriter->write('<dimension ref="' . $minCell . ':' . $maxCell . '"/>');
        }

        $rightToLeftValue = $sheet->isRightToLeft() ? 'true' : 'false';

        $fileWriter->write('<sheetViews>');

        $tabSelected = ($sheet->active ? 'tabSelected="true"' : '');
        //$fileWriter->write('<sheetView colorId="64" defaultGridColor="true" rightToLeft="' . $rightToLeftValue . '" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" ' . $tabSelected . ' topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
        $fileWriter->write('<sheetView rightToLeft="' . $rightToLeftValue . '" ' . $tabSelected . ' topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');

        $paneRow = ($sheet->freezeRows ? $sheet->freezeRows + 1 : 0);
        $paneCol = ($sheet->freezeColumns ? $sheet->freezeColumns + 1 : 0);
        if ($sheet->freezeRows && $sheet->freezeColumns) {
            // frozen rows and cols
            $fileWriter->write('<pane ySplit="' . $sheet->freezeRows . '" xSplit="' . $sheet->freezeColumns . '" topLeftCell="' . Excel::cellAddress($paneRow, $paneCol) . '" activePane="bottomRight" state="frozen"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress($paneRow, 1) . '" activeCellId="0" pane="topRight" sqref="' . Excel::cellAddress($paneRow, 1) . '"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress(1, $paneCol) . '" activeCellId="0" pane="bottomLeft" sqref="' . Excel::cellAddress(1, $paneCol) . '"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress($paneRow, $paneCol) . '" activeCellId="0" pane="bottomRight" sqref="' . Excel::cellAddress($paneRow, $paneCol) . '"/>');
        }
        elseif ($sheet->freezeRows) {
            // frozen rows only
            $fileWriter->write('<pane ySplit="' . $sheet->freezeRows . '" topLeftCell="' . Excel::cellAddress($paneRow, 1) . '" activePane="bottomLeft" state="frozen"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress($paneRow, 1) . '" activeCellId="0" pane="bottomLeft" sqref="' . Excel::cellAddress($paneRow, 1) . '"/>');
        }
        elseif ($sheet->freezeColumns) {
            // frozen cols only
            $fileWriter->write('<pane xSplit="' . $sheet->freezeColumns . '" topLeftCell="' . Excel::cellAddress(1, $paneCol) . '" activePane="topRight" state="frozen"/>');
            $fileWriter->write('<selection activeCell="' . Excel::cellAddress(1, $paneCol) . '" activeCellId="0" pane="topRight" sqref="' . Excel::cellAddress(1, $paneCol) . '"/>');
        }
        else {
            // not frozen
            $fileWriter->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
        }
        $fileWriter->write('</sheetView>');

        $fileWriter->write('</sheetViews>');

        if (!empty($sheet->colWidths)) {
            $fileWriter->write('<cols>');
            foreach ($sheet->colWidths as $colNum => $columnWidth) {
                if (is_float($columnWidth)) {
                    // important for some locales
                    $columnWidth = str_replace(',', '.', (string)$columnWidth);
                }
                $fileWriter->write('<col min="' . ($colNum + 1) . '" max="' . ($colNum + 1) . '" width="' . $columnWidth . '" customWidth="1"/>');
            }
            $fileWriter->write('</cols>');
        }

        return $fileWriter;
    }

    /**
     * @param Sheet $sheet
     */
    public function writeSheetDataBegin(Sheet $sheet)
    {
        //if already initialized
        if ($sheet->open) {
            return;
        }

        $sheet->writeDataBegin($this);
    }

    /**
     * @param Sheet $sheet
     */
    public function writeSheetDataEnd(Sheet $sheet)
    {
        if ($sheet->close) {
            return;
        }

        $sheet->writeDataEnd();

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
            $maxCell = Excel::cellAddress($sheet->rowCountWritten, $sheet->colCountWritten);
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
        //$sheet->fileWriter->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');

        $links = $sheet->getExternalLinks();
        if ($links) {
            $sheet->fileWriter->write('<hyperlinks>');
            foreach ($links as $id => $data) {
                $sheet->fileWriter->write('<hyperlink ref="' . $data['cell'] . '" r:id="rId' . $id . '" />');
            }
            $sheet->fileWriter->write('</hyperlinks>');
        }

        $sheet->fileWriter->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');

        $sheet->fileWriter->write("<pageSetup  paperSize=\"1\" useFirstPageNumber=\"1\" horizontalDpi=\"0\" verticalDpi=\"0\" $pageSetupAttr />");

        $sheet->fileWriter->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $sheet->fileWriter->write('<oddHeader/>');
        $sheet->fileWriter->write('<oddFooter/>');
        $sheet->fileWriter->write('</headerFooter>');

        if ($rId = $sheet->getDrawingId()) {
            $sheet->fileWriter->write('<drawing r:id="rId' . $rId . '"/>');
        }

        if ($rId = $sheet->getLegacyDrawingId()) {
            $sheet->fileWriter->write('<legacyDrawing r:id="rId' . $rId . '"/>');
        }

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
    protected function _convertFormula($formula, $baseAddress): string
    {
        static $functionNames = [];

        //$formula = substr($formula, 1);

        $mark = md5(microtime());
        $replace = [];
        // temporary replace strings
        if (strpos($formula, '"') !== false) {
            $replace = [[], []];
            $formula = preg_replace_callback('/"[^"]+"/', static function ($matches) use ($mark, &$replace) {
                $key = '<<' . $mark . '-' . md5($matches[0]) . '>>';
                $replace[0][] = $key;
                $replace[1][] = $matches[0];
                return $key;
            }, $formula);
        }
        // change relative addresses: =RC[-1]*RC[-2] -> =B1*A1
        $formula = preg_replace_callback('/(\W)(R\[?(-?\d+)?\]?C\[?(-?\d+)?\]?)/', static function ($matches) use ($baseAddress) {
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

        if (!empty($this->excel->style->localeSettings['functions']) && strpos($formula, '(')) {
            // replace national function names
            if (empty($functionNames)) {
                $functionNames = [[], []];
                foreach ($this->excel->style->localeSettings['functions'] as $name => $nameEn) {
                    $functionNames[0][] = $name . '(';
                    $functionNames[1][] = $nameEn . '(';
                }
            }
            $formula = str_replace($functionNames[0], $functionNames[1], $formula);
        }

        if ($replace && !empty($replace[0])) {
            // restore strings
            $formula = str_replace($replace[0], $replace[1], $formula);
        }

        if ($formula) {
            $formula = (string) preg_replace(self::XLFNREGEXP, '_xlfn.$1(', $formula);
        }
        if ($formula) {
            $formula = (string) preg_replace(self::XLWSREGEXP, '_xlws.$1(', $formula);
        }

        return $formula;
    }

    /**
     * @param WriterBuffer $file
     * @param int $rowNumber
     * @param int $colNumber
     * @param mixed $value
     * @param mixed $numFormatType
     * @param int|null $cellStyleIdx
     */
    public function _writeCell(WriterBuffer $file, int $rowNumber, int $colNumber, $value, $numFormatType, ?int $cellStyleIdx = 0)
    {
        $cellName = Excel::cellAddress($rowNumber, $colNumber);

        if (is_array($value) && isset($value['shared_index'])) {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="s"><v>' . $value['shared_index'] . '</v></c>');
        }
        elseif (!is_scalar($value) || $value === '') { //objects, array, empty; null is not scalar
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '"/>');
        }
        elseif (is_string($value) && $value[0] === '=') {
            // formula
            $value = $this->_convertFormula($value, [$rowNumber, $colNumber]);
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="s"><f>' . self::xmlSpecialChars($value) . '</f></c>');
        }
        elseif ($numFormatType === 'n_shared_string') {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="s"><v>' . $value . '</v></c>');
        }
        elseif ($numFormatType === 'n_string' || ($numFormatType === 'n_numeric' && !is_numeric($value))) {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="inlineStr"><is><t xml:space="preserve">' . self::xmlSpecialChars($value) . '</t></is></c>');
        }
        else {
            if ($numFormatType === 'n_date' || $numFormatType === 'n_datetime') {
                $dateValue = self::convertDateTime($value);
                if ($dateValue === false) {
                    $numFormatType = 'n_auto';
                }
                else {
                    $value = $dateValue;
                }
            }
            if ($numFormatType === 'n_date') {
                //$file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . (int)self::convertDateTime($value) . '</v></c>');
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '"><v>' . (int)$value . '</v></c>');
            }
            elseif ($numFormatType === 'n_datetime') {
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . $value . '</v></c>');
            }
            elseif ($numFormatType === 'n_numeric') {
                if (!is_int($value) && !is_float($value)) {
                    $value = self::xmlSpecialChars($value);
                }
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" ><v>' . $value . '</v></c>');//int,float,currency
            }
            elseif ($numFormatType === 'n_auto' || 1) { //auto-detect unknown column types
                if (!is_string($value) || $value === '0' || ($value[0] !== '0' && preg_match('/^\d+$/', $value)) || preg_match("/^-?(0|[1-9]\d*)(\.\d+)?$/", $value)) {
                    $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . $value . '</v></c>');//int,float,currency
                }
                else {
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
     * @param array|null $border
     * @param string $side
     *
     * @return string
     */
    protected function _makeBorderSideTag(?array $border, string $side): string
    {
        if (empty($border[$side]) || empty($border[$side]['style'])) {
            $tag = "<$side/>";
        }
        elseif (empty($border[$side]['color'])) {
            $tag = "<$side style=\"" . $border[$side]['style'] . '"/>';
        }
        else {
            $tag = "<$side style=\"" . $border[$side]['style'] . '">';
            $tag .= '<color rgb="' . $border[$side]['color'] . '"/>';
            $tag .= "</$side>";
        }
        return $tag;
    }

    /**
     * @param array $borders
     *
     * @return string
     */
    protected function _makeBordersTag(array $borders): string
    {
        $tag = '<borders count="' . (count($borders)) . '">';
        foreach ($borders as $border) {
            $tag .= '<border diagonalDown="false" diagonalUp="false">';
            if (!$border) {
                $tag .= '<left/><right/><top/><bottom/>';
            }
            else {
                $tag .= $this->_makeBorderSideTag($border, 'left');
                $tag .= $this->_makeBorderSideTag($border, 'right');
                $tag .= $this->_makeBorderSideTag($border, 'top');
                $tag .= $this->_makeBorderSideTag($border, 'bottom');
            }
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
        $schemaLinks = [
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            //'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr"',
            //'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"',
            //'xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"',
            //'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"',
        ];

        $temporaryFilename = $this->tempFilename();
        $file = new WriterBuffer($temporaryFilename);
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $links = implode("\n", $schemaLinks);
        $file->write("<styleSheet $links>");
        // cell styles, table styles, and pivot styles
        //$file->write('');

        //// +++++++++++
        //// <numFmts/>
        $numberFormats = $this->excel->style->_getNumberFormats();
        if (!$numberFormats) {
            $file->write('<numFmts count="0"/>');
        } else {
            $file->write('<numFmts count="' . count($numberFormats) . '">');
            foreach ($numberFormats as $num => $fmt) {
                $file->write('<numFmt numFmtId="' . (164 + $num) . '" formatCode="' . self::xmlSpecialChars($fmt) . '" />');
            }
            $file->write('</numFmts>');
        }

        //// +++++++++++
        //// <fonts/>
        $fonts = $this->excel->style->getStyleFonts();
        if (!$fonts) {
            $file->write('<fonts count="0"/>');
        }
        else {
            $file->write('<fonts count="' . count($fonts) . '">');
            foreach ($fonts as $font) {
                if (!empty($font)) {
                    $file->write($font['tag']);
                }
            }
            $file->write('</fonts>');
        }

        //// +++++++++++
        //// <fills/>
        $fills = $this->excel->style->getStyleFills();
        if (!$fills) {
            $file->write('<fills count="0"/>');
        }
        else {
            $file->write('<fills count="' . count($fills) . '">');
            foreach ($fills as $fill) {
                $file->write($fill['tag']);
            }
            $file->write('</fills>');
        }

        //// +++++++++++
        // <borders/>
        $borders = $this->excel->style->getStyleBorders();
        if (!$borders) {
            $file->write('<borders count="0"/>');
        }
        else {
            $file->write('<borders count="' . (count($borders)) . '">');
            foreach ($borders as $border) {
                $file->write($border['tag']);
            }
            $file->write('</borders>');
        }

        //// +++++++++++
        // <cellStyleXfs/>
        $cellStyleXfs = [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
        ];
        $file->write('<cellStyleXfs count="' . count($cellStyleXfs) . '">');
        foreach ($cellStyleXfs as $cellStyleXf) {
            $file->write($cellStyleXf);
        }
        $file->write('</cellStyleXfs>');

        //// +++++++++++
        // <cellXfs/>
        $cellXfs = $this->excel->style->getStyleCellXfs();
        if (!$cellXfs) {
            $file->write('<cellXfs count="1">');
            $file->write('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
            $file->write('</cellXfs>');
        }
        else {
            $file->write('<cellXfs count="' . count($cellXfs) . '">');
            foreach ($cellXfs as $cellXf) {
                $alignmentAttr = '';
                if (!empty($cellXf['format']['format-text-rotation'])) {
                    $alignmentAttr .= ' textRotation="' . $cellXf['format']['format-text-rotation'] . '"';
                }
                if (!empty($cellXf['format']['format-align-horizontal'])) {
                    $alignmentAttr .= ' horizontal="' . $cellXf['format']['format-align-horizontal'] . '"';
                }
                if (!empty($cellXf['format']['format-align-vertical'])) {
                    $alignmentAttr .= ' vertical="' . $cellXf['format']['format-align-vertical'] . '"';
                }
                if (!empty($cellXf['format']['format-text-wrap'])) {
                    $alignmentAttr .= ' wrapText="true"';
                }

                $xfAttr = 'applyFont="true" ';
                if (isset($cellXf['_border_id'])) {
                    $xfAttr .= 'applyBorder="true" ';
                }

                $xfId = $cellXf['_xf_id'] ?? 0;
                if ($alignmentAttr) {
                    $xfAttr .= 'applyAlignment="true" ';
                    $file->write('<xf ' . $xfAttr . ' borderId="' . $cellXf['_border_id'] . '" fillId="' . $cellXf['_fill_id'] . '" '
                        . 'fontId="' . $cellXf['_font_id'] . '" numFmtId="' . (164 + $cellXf['_num_fmt_id']) . '" xfId="' . $xfId . '">');
                    $file->write('	<alignment ' . $alignmentAttr . '/>');
                    $file->write('</xf>');
                }
                else {
                    $file->write('<xf ' . $xfAttr . ' borderId="' . $cellXf['_border_id'] . '" fillId="' . $cellXf['_fill_id'] . '" '
                        . 'fontId="' . $cellXf['_font_id'] . '" numFmtId="' . (164 + $cellXf['_num_fmt_id']) . '" xfId="' . $xfId . '" />');
                }
            }

            $file->write('</cellXfs>');
        }

        //// +++++++++++
        // <cellStyles/>
        $cellStyles = [
            '<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>',
            //'<cellStyle builtinId="8" customBuiltin="false" name="Hyperlink" xfId="1" />',
            //'<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="2"/>',
            //'<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="3"/>',
            //'<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="4"/>',
            //'<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="5"/>',
            //'<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="6"/>',
        ];
        $file->write('<cellStyles count="' . count($cellStyles) . '">');
        foreach ($cellStyles as $cellStyle) {
            $file->write($cellStyle);
        }
        $file->write('</cellStyles>');

        // <dxfs/>
        $file->write('<dxfs count="0"/>');

        //// +++++++++++
        // <tableStyles/>
        $file->write('<tableStyles count="0"/>');

        if ($this->excel->style->indexedColors) {
            $file->write('<colors><indexedColors>');
            foreach ($this->excel->style->indexedColors as $color) {
                $file->write('<rgbColor rgb="' . $color . '"/>');
            }
            $file->write('</indexedColors></colors>');
        }

        $file->write('</styleSheet>');
        $file->close();

        return $temporaryFilename;
    }

    /**
     * @param array|null $metadata
     *
     * @return string
     */
    protected function _buildAppXML(?array $metadata): string
    {
        $xmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $xmlText .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
        $xmlText .= '<TotalTime>0</TotalTime>';
        $xmlText .= '<Company>' . self::xmlSpecialChars($metadata['company'] ?? '') . '</Company>';
        $xmlText .= '<HyperlinksChanged>false</HyperlinksChanged>';
        $xmlText .= '</Properties>';

        return $xmlText;
    }

    /**
     * @param array|null $metadata
     *
     * @return string
     */
    protected function _buildCoreXML(?array $metadata): string
    {
        $coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
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
     * @param array $relationShips
     *
     * @return string
     */
    protected function _buildRelationshipsXML(array $relationShips)
    {
        $xmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $xmlText .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $rId = 1;
        foreach ($relationShips['override'] as $relItem => $relData) {
            if (!empty($relData['rel']) && $relData['rel'] === 'root') {
                $relId = 'rId' . ($rId++);
                $xmlText .= '<Relationship Id="' . $relId . '" Type="' . $relData['schema'] . '" Target="' . $relItem . '"/>';
            }
        }
        $xmlText .= '</Relationships>';

        return $xmlText;
    }

    /**
     * @param Sheet[] $sheets
     *
     * @return string
     */
    protected function _buildWorkbookXML(array $sheets): string
    {
        $i = 0;
        $xmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $xmlText .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $xmlText .= '<fileVersion appName="FastExcelWriter"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $xmlText .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $xmlText .= '<sheets>';
        $definedNames = '';
        foreach ($sheets as $sheet) {
            $xmlText .= '<sheet name="' . self::xmlSpecialChars($sheet->sanitizedSheetName) . '" sheetId="' . $sheet->index . '" state="visible" r:id="' . $sheet->relId . '"/>';
            if ($sheet->absoluteAutoFilter) {
                $filterRange = $sheet->absoluteAutoFilter . ':' . Excel::cellAddress($sheet->rowCountWritten, $sheet->colCountWritten, true);
                $definedNames .= '<definedName name="_xlnm._FilterDatabase" localSheetId="' . $i . '" hidden="1">\'' . $sheet->sanitizedSheetName . '\'!' . $filterRange . '</definedName>';
                //                <definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Лист1!$A$1:$B$1</definedName>
            }
            $i++;
        }
        $xmlText .= '</sheets>';
        foreach ($sheets as $sheet) {
            foreach ($sheet->getNamedRanges() as $range) {
                $definedNames .= '<definedName name="' . self::xmlSpecialChars($range['name']) . '">\'' . $sheet->sanitizedSheetName . '\'!' . $range['range'] . '</definedName>';
            }
        }

        if ($definedNames) {
            $xmlText .= '<definedNames>' . $definedNames . '</definedNames>';
        }
        else {
            $xmlText .= '<definedNames/>';
        }

        $xmlText .= '<calcPr refMode="A1" calcId="162913"/></workbook>';

        return $xmlText;
    }

    /**
     * Сontents of /xl/_rels/workbook.xml.rels
     *
     * @param array $relationShips
     *
     * @return string
     */
    protected function _buildWorkbookRelsXML(array $relationShips): string
    {
        $xmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $xmlText .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

        foreach ($relationShips['override'] as $relItem => $relData) {
            if (!empty($relData['rel']) && $relData['rel'] === 'workbook') {
                $xmlText .= '<Relationship Id="' . $relData['r_id'] . '" Type="' . $relData['schema'] . '" Target="' . substr($relItem, 3) . '"/>';
            }
        }

        $xmlText .= '</Relationships>';

        return $xmlText;
    }

    /**
     * Contents of /[Content_Types].xml
     *
     * @param array $relationShips
     *
     * @return string
     */
    protected function _buildContentTypesXML(array $relationShips): string
    {
        $xmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $xmlText .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';

        foreach ($relationShips as $tag => $tagSet) {
            foreach ($tagSet as $name => $data) {
                if ($tag === 'default') {
                    $xmlText .= '<Default Extension="' . $name . '" ContentType="' . $data['content_type'] . '"/>';
                }
                elseif ($tag === 'override') {
                    $xmlText .= '<Override PartName="/' . $name . '" ContentType="' . $data['content_type'] . '"/>';
                }
            }
        }

        $xmlText .= '</Types>';

        return $xmlText;
    }

    /**
     * @return string
     */
    protected function _buildThemeXML(): string
    {
        return '';
    }

    /**
     * @see http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
     *
     * @param $filename
     *
     * @return string
     */
    public static function sanitizeFilename($filename): string
    {
        $nonPrinting = array_map('chr', range(0, 31));
        $invalidChars = ['<', '>', '?', '"', ':', '|', '\\', '/', '*', '&'];
        $allInvalids = array_merge($nonPrinting, $invalidChars);

        return str_replace($allInvalids, "", $filename);
    }

    /**
     * @param $sheetName
     *
     * @return string
     */
    public static function sanitizeSheetName($sheetName): string
    {
        static $badChars  = '\\/?*:[]';
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
    public static function xmlSpecialChars($val): string
    {
        //note, bad chars does not include \t\n\r (\x09\x0a\x0d)
        static $badChars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodChars = "                              ";

        return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badChars, $goodChars);//strtr appears to be faster than str_replace
    }

    /**
     * //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
     *
     * @param mixed $dateInput
     *
     * @return int|float|bool
     */
    public static function convertDateTime($dateInput)
    {
        if (is_int($dateInput) || (is_string($dateInput) && preg_match('/^\d+$/', $dateInput))) {
            // date as timestamp
            $time = (int)$dateInput;
        }
        elseif (preg_match('/^(\d+:\d{1,2})(:\d{1,2})?$/', $dateInput, $matches)) {
            // time only
            $time = strtotime('1900-01-00 ' . $matches[1] . ($matches[2] ?? ':00'));
        }
        elseif (is_string($dateInput) && $dateInput && $dateInput[0] >= '0' && $dateInput[0] <= '9') {
            //starts with a digit
            $time = strtotime($dateInput);
        }
        else {
            $time = 0;
        }

        if ($time && preg_match('/(\d{4})-(\d{2})-(\d{2})\s(\d+):(\d{2}):(\d{2})/', date('Y-m-d H:i:s', $time), $matches)) {
            [$junk, $year, $month, $day, $hour, $min, $sec] = $matches;
            $seconds = $sec / 86400 + $min / 1440 + $hour / 24;
        }
        else {
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
        }    # Excel false leap day

        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leap days and 400 for 400 year leap days.
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