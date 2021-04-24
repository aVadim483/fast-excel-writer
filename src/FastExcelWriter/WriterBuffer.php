<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;

/**
 * Class WriterBuffer
 *
 * @package avadim\FastExcelWriter
 */
class WriterBuffer
{
    /** @var bool|resource|null  */
    protected $fd = null;

    /** @var string  */
    protected $buffer = '';

    /** @var bool  */
    protected $checkUtf8 = false;

    protected $fileName;
    protected $openFlags;
    protected $limit = 8191;

    /**
     * WriterBuffer constructor.
     * @param        $filename
     * @param string $openFlags
     * @param bool   $checkUtf8
     */
    public function __construct($filename, $openFlags = 'wb', $checkUtf8 = false)
    {
        $this->fileName = $filename;
        $this->openFlags = $openFlags;
        $this->checkUtf8 = $checkUtf8;
    }

    /**
     * @return string
     */
    public function getFileName()
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

    /**
     * @param bool $force
     */
    public function flush($force = false)
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
        $this->flush(true);
        if ($this->fd) {
            fclose($this->fd);
            $this->fd = null;
        }
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
     * @return bool|int
     * /
    public function pos()
    {
        $this->flush(true);
        if ($this->fd) {
            return ftell($this->fd);
        }
        return strlen($this->buffer);
    }

    /**
     * @param $pos
     * @return int
     * /
    public function seek($pos)
    {
        $this->flush(true);
        if ($this->fd) {
            return fseek($this->fd, $pos);
        }
        return -1;
    }
*/
    public function getFileHandler()
    {
        return $this->fd;
    }

    /**
     * @param $string
     * @return bool
     */
    protected static function isValidUTF8($string)
    {
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($string, 'UTF-8') ? true : false;
        }
        return preg_match("//u", $string) ? true : false;
    }

    /**
     * @param $fileWriter
     * @param $newFileName
     *
     * @return int
     */
    public function appendFileWriter($fileWriter, $newFileName)
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

        return $n1 + $n2;
    }
}

// УЩА