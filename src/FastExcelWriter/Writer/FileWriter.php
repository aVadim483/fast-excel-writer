<?php

namespace avadim\FastExcelWriter\Writer;

use avadim\FastExcelWriter\Exceptions\Exception;

/**
 * Class WriterBuffer
 *
 * @package avadim\FastExcelWriter
 */
class FileWriter
{
    /** @var bool|resource|null  */
    protected $fd = null;

    /** @var string  */
    protected string $buffer = '';

    /** @var bool  */
    protected ?bool $checkUtf8 = false;

    protected string $fileName;
    protected ?string $openFlags;
    protected int $limit = 8191;

    protected array $elements = [];
    protected int $level = -1;
    protected bool $close = false;

    /**
     * WriterBuffer constructor.
     * @param string $filename
     * @param string|null $openFlags
     * @param bool|null $checkUtf8
     */
    public function __construct(string $filename, ?string $openFlags = 'wb', ?bool $checkUtf8 = false)
    {
        $this->fileName = $filename;
        $this->openFlags = $openFlags;
        $this->checkUtf8 = $checkUtf8;
    }

    /**
     *
     */
    public function __destruct()
    {
        if ($this->buffer || $this->fd) {
            $this->close();
        }
    }

    /**
     * @return string
     */
    public function getFileName(): string
    {
        return $this->fileName;
    }

    /**
     * @param $string
     */
    public function write($string)
    {
        $this->buffer .= $string;
        if (isset($this->buffer[$this->limit])) {
            $this->flush();
        }
    }


    private function closeElement()
    {
        if (isset($this->elements[$this->level]) && !$this->elements[$this->level]['close']) {
            if ($this->elements[$this->level]['attr']) {
                $this->write(Writer::tagAttributes($this->elements[$this->level]['attr']));
            }
            $this->write('>');
            $this->elements[$this->level]['close'] = true;
            $this->elements[$this->level]['short'] = false;
        }
    }

    public function startElement($name, ?array $attr = [])
    {
        $this->closeElement();
        $name = trim($name);
        if ($name[0] === '<') {
            $this->write($name);
            $name = substr($name, 1, strlen($name) - 2);
            if ($pos = strpos($name, ' ')) {
                $name = substr($name, 0, $pos);
            }
            $this->elements[++$this->level] = [
                'name' => $name,
                'attr' => [],
                'short' => false,
                'close' => true,
            ];
        }
        else {
            $this->elements[++$this->level] = [
                'name' => $name,
                'attr' => $attr,
                'short' => true,
                'close' => false,
            ];
            $this->write('<' . $this->elements[$this->level]['name']);
        }
    }

    /**
     * @param string $key
     * @param $val
     *
     * @return void
     */
    public function writeAttribute(string $key, $val)
    {
        $this->elements[$this->level]['attr'][$key] = $val;
    }


    public function endElement()
    {
        if ($this->elements[$this->level]['attr'] && !$this->elements[$this->level]['close']) {
            $this->write(Writer::tagAttributes($this->elements[$this->level]['attr']));
        }
        if ($this->elements[$this->level]['short']) {
            $this->write('/>');
        }
        else {
            $this->write('</' . $this->elements[$this->level]['name'] . '>');
        }
        $this->elements[$this->level]['close'] = true;
        unset($this->elements[$this->level--]);
    }

    /**
     * @param string $name
     * @param array|null $attributes
     * @param string|null $content
     *
     * @return void
     */
    public function writeElementAttr(string $name, ?array $attributes = [], ?string $content = null)
    {
        $this->closeElement();
        $name = trim($name);
        if ($name[0] === '<') {
            $this->write($name);
        }
        else {
            $this->write('<' . $name);
            if ($attributes) {
                $this->write(Writer::tagAttributes($attributes));
            }
            if ($content !== null) {
                $this->write('>' . Writer::xmlSpecialChars($content) . '</' . $name . '>');
            }
            else {
                $this->write('/>');
            }
        }
    }

    /**
     * @param string $name
     * @param string|null $content
     *
     * @return void
     */
    public function writeElement(string $name, ?string $content = null)
    {
        $this->writeElementAttr($name, [], $content);
    }

    /**
     * @param string $text
     *
     * @return void
     */
    public function writeRawData(string $text)
    {
        if (isset($this->elements[$this->level])) {
            if ($this->elements[$this->level]['attr']) {
                $this->write(Writer::tagAttributes($this->elements[$this->level]['attr']));
            }
            $this->elements[$this->level]['short'] = false;
        }
        $this->write('>' . $text);
    }

    /**
     * @param bool|null $force
     */
    public function flush(?bool $force = false)
    {
        if ($this->buffer || $force) {
            if (!$this->fd) {
                $this->fd = fopen($this->fileName, $this->openFlags);
                if ($this->fd === false) {
                    throw new Exception("Unable to open {$this->fileName} for writing");
                }
            }
            if ($this->checkUtf8 && !self::isValidUTF8($this->buffer)) {
                //Excel::log("Error, invalid UTF8 encoding detected");
                $this->checkUtf8 = false;
            }
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }

    /**
     *
     */
    public function close()
    {
        if (!$this->close) {
            $this->flush(true);
            if ($this->fd) {
                fclose($this->fd);
                $this->fd = null;
            }
            $this->close = true;
        }
    }


    public function getFileResource()
    {
        return $this->fd;
    }

    /**
     * @param $string
     * @return bool
     */
    protected static function isValidUTF8($string): bool
    {
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($string, 'UTF-8');
        }
        return (bool)preg_match("//u", $string);
    }

    /**
     * Append content of $fileWriter to this file and save all to $newFileName
     *
     * @param FileWriter $fileWriter
     * @param string $newFileName
     *
     * @return int
     */
    public function appendFileWriter(FileWriter $fileWriter, string $newFileName): int
    {
        $fdTarget = fopen($newFileName, 'wb');

        $this->close();
        $fd1 = fopen($this->getFileName(), 'rb');

        $fileWriter->close();
        $fd2 = fopen($fileWriter->getFileName(), 'rb');

        $n1 = stream_copy_to_stream($fd1, $fdTarget);
        fclose($fd1);
        $n2 = stream_copy_to_stream($fd2, $fdTarget);
        fclose($fd2);
        $this->fd = $fdTarget;
        $this->fileName = $newFileName;
        $this->close = false;

        return $n1 + $n2;
    }
}

// УЩА