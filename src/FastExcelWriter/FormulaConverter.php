<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelHelper\Helper;

class FormulaConverter
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

    protected static array $functionNames = [];
    protected array $localFunctions = [];


    /**
     * @param array|null $functions
     */
    public function __construct(?array $functions = [])
    {
        $this->localFunctions = $functions;
    }


    /**
     * @param string $formula
     * @param string|array $baseAddress
     *
     * @return string
     */
    public function normalize(string $formula, $baseAddress): string
    {
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
        $formula = preg_replace_callback('/(\W)(R\[?(-?\d+)?]?C\[?(-?\d+)?]?)/', static function ($matches) use ($baseAddress) {
            if (is_array($baseAddress)) {
                $cell = Excel::cellAddress($baseAddress[0], $baseAddress[1]);
            }
            else {
                $cell = $baseAddress;
            }
            if ($cell && ($address = Helper::RCtoA1($matches[2], $cell))) {
                return $matches[1] . $address;
            }

            return $matches[0];
        }, $formula);

        if ($this->localFunctions && strpos($formula, '(')) {
            // replace national function names
            if (empty(self::$functionNames)) {
                self::$functionNames = [[], []];
                foreach ($this->localFunctions as $name => $nameEn) {
                    self::$functionNames[0][] = '/' . $name . '\s*\(/ui';
                    self::$functionNames[1][] = $nameEn . '(';
                    if ($nameEn === 'FALSE' || $nameEn === 'TRUE') {
                        self::$functionNames[0][] = '/([\(;,])\s*' . $name . '\s*([\);,])/ui';
                        self::$functionNames[1][] = '$1' . $nameEn . '$2';
                    }
                }
            }
            //$formula = str_replace(self::$functionNames[0], self::$functionNames[1], $formula);
            $formula = preg_replace(self::$functionNames[0], self::$functionNames[1], $formula);
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

        return ($formula[0] === '=') ? mb_substr($formula, 1) : $formula;
    }



}