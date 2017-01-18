<?php
/**
 * Wordlib
 *
 * User: liangtaohy@163.com
 * Date: 17/1/3
 * Time: PM1:50
 */

namespace PhpOffice\PhpWord;

use PhpOffice\PhpWord\Reader\MsDoc;

class Wordlib
{
    static $OLEBytes = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

    static $DOSBytes = [0x31, 0xbe, 0x00, 0x00, 0x00, 0xab];

    static $MacBytes = [
        [0xfe, 0x37, 0x00, 0x1c, 0x00, 0x00],
        [0xfe, 0x37, 0x00, 0x23, 0x00, 0x00]
    ];

    static $RTFBytes = ['{', '\\', 'r', 't', 'f', '1'];

    static $Win12Bytes = [
        [0x9b, 0xa5, 0x21, 0x00],	/* Win Word 1.x */
        [0xdb, 0xa5, 0x2d, 0x00],	/* Win Word 2.0 */
    ];

    static $WPBytes = [0xff, 'W', 'P', 'C'];

    /**
     * Week Maps
     * @var array
     */
    static $WDYMaps = [
        'en'    => ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
        'ch'    => ['星期天', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'],
    ];

    /**
     * Month Maps
     * @var array
     */
    static $MONMaps = [
        'en'    => ['Undefined', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
        'ch'    => ['非法', '一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'],
    ];

    const BIT_30 = 0x40000000;
    const BIG_BLOCK_SIZE = 512;

    const VERSION_97 = '97';
    const VERSION_2000 = '2000';
    const VERSION_2002 = '2002';
    const VERSION_2003 = '2003';
    const VERSION_2007 = '2007';

    const SPRA_VALUE = 10;
    const SPRA_VALUE_OPPOSITE = 20;

    const OFFICEARTBLIPEMF = 0xF01A;
    const OFFICEARTBLIPWMF = 0xF01B;
    const OFFICEARTBLIPPICT = 0xF01C;
    const OFFICEARTBLIPJPG = 0xF01D;
    const OFFICEARTBLIPPNG = 0xF01E;
    const OFFICEARTBLIPDIB = 0xF01F;
    const OFFICEARTBLIPTIFF = 0xF029;
    const OFFICEARTBLIPJPEG = 0xF02A;

    const MSOBLIPERROR = 0x00;
    const MSOBLIPUNKNOWN = 0x01;
    const MSOBLIPEMF = 0x02;
    const MSOBLIPWMF = 0x03;
    const MSOBLIPPICT = 0x04;
    const MSOBLIPJPEG = 0x05;
    const MSOBLIPPNG = 0x06;
    const MSOBLIPDIB = 0x07;
    const MSOBLIPTIFF = 0x11;
    const MSOBLIPCMYKJPEG = 0x12;

    /**
     * Common part of the file checking functions
     * @param $data
     * @param $bytes
     * @return bool
     */
    private static function CheckBytes($data, $bytes)
    {
        $c = count($bytes);
        for($i = 0; $i < $c; $i++) {
            $n = ord($data[$i]);
            if ($n !== $bytes[$i]) {
                return false;
            }
        }
    }

    /**
     * This function checks whether the given file is or is not a file with an
     * OLE envelope (That is a document made by Word 6 or later)
     * @param $data
     * @param $fileSize
     * @return bool
     */
    public static function IsWordFileWithOLE($data, $fileSize)
    {
        $iTailLen = null;

        if (empty($data)) {
            echo "No proper data given";
            return FALSE;
        }

        if ($iTailLen < self::BIG_BLOCK_SIZE * 3) {
            echo "This file is too small to be a Word document";
            return false;
        }

        $iTailLen = intval(($fileSize % self::BIG_BLOCK_SIZE));
        switch ($iTailLen) {
            case 0:		/* No tail, as it should be */
                break;
            case 1:
            case 2:		/* Filesize mismatch or a buggy email program */
                if (intval(($fileSize % 3)) == $iTailLen) {
                    echo 'Filesize mismatch or a buggy email program';
                    return FALSE;
                }
                /*
                 * Ignore extra bytes caused by buggy email programs.
                 * They have bugs in their base64 encoding or decoding.
                 * 3 bytes -> 4 ascii chars -> 3 bytes
                 */
                break;
            default:	/* Wrong filesize for a Word document */
                return FALSE;
        }

        return self::CheckBytes($data, self::$OLEBytes);
    }

    /**
     * This function checks whether the given file is or is not a "Word for DOS"
     * document
     * @param $data
     * @param $fileSize
     * @return bool
     */
    public static function IsWordForDosFile($data, $fileSize)
    {
        if (empty($data)) {
            echo "No proper data given";
            return FALSE;
        }
        if ($fileSize < 128) {
            echo "File too small to be a Word document";
            return FALSE;
        }
        return self::CheckBytes($data, self::$DOSBytes);
    }

    /**
     * This function checks whether the given file is or is not a "Mac Word 4 or 5"
     * document
     * @param $data
     * @return bool
     */
    public static function IsMacWord45File($data)
    {
        foreach (self::$MacBytes as $macByte) {
            if (self::CheckBytes($data, $macByte)) {
                return TRUE;
            }
        }

        return false;
    }

    /**
     * This function checks whether the given file is or is not a RTF document
     * @param $data
     * @return bool
     */
    public static function IsRtfFile($data)
    {
        return self::CheckBytes($data, self::$RTFBytes);
    }

    /**
     * This function checks whether the given file is or is not a "Win Word 1 or 2"
     * document
     * @param $data
     * @param $fileSize
     * @return bool
     */
    public static function IsWinWord12File($data, $fileSize)
    {
        if ($fileSize < 384) {
            echo "This file is too small to be a Word document";
            return FALSE;
        }

        foreach (self::$Win12Bytes as $win12Byte) {
            if (self::CheckBytes($data, $win12Byte)) {
                return TRUE;
            }
        }
        return FALSE;
    }

    /**
     * This function checks whether the given file is or is not a WP document
     * @param $data
     * @return bool
     */
    public static function IsWordPerfectFile($data)
    {
        if (empty($data)) {
            echo "No proper data given";
            return FALSE;
        }
        return self::CheckBytes($data, self::$WPBytes);
    }

    /**
     * iGuessVersionNumber - guess the Word version number from first few bytes
     *
     * Returns the guessed version number or -1 when no guess it possible
     * @param $data
     * @param $lFilesize
     * @return int
     */
    public static function GuessVersionNumber($data, $lFilesize)
    {
	    if(self::IsWordForDosFile($data, $lFilesize)) {
		    return 0;
	    }
        if (self::IsWinWord12File($data, $lFilesize)) {
            return 2;
        }
        if (self::IsMacWord45File($data)) {
            return 5;
        }
        if (self::IsWordFileWithOLE($data, $lFilesize)) {
            return 6;
        }
        return -1;
    }

    /**
     * iGetVersionNumber - get the Word version number from the header
     *
     * Returns the version number or -1 when unknown
     * @param $aucHeader
     * @param $bOldMacFile
     * @return int
     */
    public static function GetVersionNumber($aucHeader, &$bOldMacFile)
    {
	    $nFib = self::getInt2d($aucHeader, 0x02);
        if ($nFib >= 0x1000) {
            /* To big: must be MacWord using Big Endian */
            $nFib = self::getInt2dBE($aucHeader, 0x02);
        }

        $bOldMacFile = FALSE;

        switch ($nFib) {
            case   0:
                echo "Word for DOS";
                return 0;
            case  28:
                echo "Word 4 for Macintosh";
                $bOldMacFile = TRUE;
                return 4;
            case  33:
                echo "Word 1.x for Windows";
                return 1;
            case  35:
                echo "Word 5 for Macintosh";
                $bOldMacFile = TRUE;
                return 5;
            case  45:
                echo "Word 2 for Windows";
                return 2;
            case 101:
            case 102:
                echo "Word 6 for Windows";
                return 6;
            case 103:
            case 104:
                $usChse = self::getInt2d($aucHeader, 0x14);

                switch ($usChse) {
                    case 0:
                        echo "Word 7 for Win95";
                        return 7;
                    case 256:
                        echo "Word 6 for Macintosh";
                        $bOldMacFile = TRUE;
                        return 6;
                    default:
                        if (self::getInt1d($aucHeader, 0x05) == 0xe0) {
                            echo "Word 7 for Win95";
                            return 7;
                        }
                        echo "Word 6 for Macintosh";
                        $bOldMacFile = TRUE;
                        return 6;
                }
            default:
                $usChse = self::getInt2d($aucHeader, 0x14);

                if ($nFib < 192) {
                    /* Unknown or unsupported version of Word */
                    echo "Unknown or unsupported version of Word";
                    return -1;
                }

                if ($usChse != 256) {
                    echo "Word97 for Win95/98/NT";
                }

                if (usChse == 256) {
                    echo "Word98 for Macintosh";
                }

                return 8;
        }
    }

    /**
     * Retrieving Text
     * @param $arrayFib
     * @param $data1Table
     * @param $dataData
     * @return stdClass
     */
    public static function GetDocumentText(&$arrayFib, &$data1Table, &$dataData)
    {
        $clx = new \stdClass();

        $clx->styles = array();

        $ulBeginTextInfo = $arrayFib['fcClx'];
        $tTextInfoLen = $arrayFib['lcbClx'];

        $aucBuffer = substr($data1Table, $ulBeginTextInfo, $tTextInfoLen);

        $lOff = 0;

        $tTextBlockList = array();
        while ($lOff < $tTextInfoLen) {
            $iType = self::getInt1d($aucBuffer, $lOff);
            $lOff++;

            if ($iType === 0x00) {
                $lOff++;
                continue;
            }

            if ($iType === 0x01) { // RgPrc element
                $cbGrpprl = self::getInt2d($aucBuffer, $lOff);
                $lOff += 2;
                $oStyle = self::readPrl($aucBuffer, $dataData, $lOff, $cbGrpprl);
                $clx->styles[] = $oStyle;
                $lOff += $oStyle->length;
                continue;
            }

            if ($iType !== 0x02) { // Not A Pcdt
                echo "Not A valid Pcdt\n";
                return false;
            }

            // handle A Pcdt
            $lcb = self::getInt4d($aucBuffer, $lOff);
            $lOff += 4;

            if ($lcb < 4) {
                echo "invalid pcdt\n";
                return false;
            }

            $lPieces = self::GetNumInLcb($lcb, 8);

            for ($i = 0; $i < $lPieces; $i++) {
                $tTextBlock = new \stdClass();

                $ulTextOffset = self::getInt4d($aucBuffer, $lOff + ($lPieces + 1) * 4 + $i * 8 + 2);
                $usPropMod = self::getInt2d($aucBuffer, $lOff + ($lPieces + 1) * 4 + $i * 8 + 6);
                $ulTotLength = self::getInt4d($aucBuffer, $lOff + ($i + 1) * 4) - self::getInt4d($aucBuffer, $lOff + $i * 4);

                if (($ulTextOffset & self::BIT_30) === 0) {
                    $bUsesUnicode = true;
                } else {
                    $bUsesUnicode = false;
                    $ulTextOffset &= ~self::BIT_30;
                    $ulTextOffset = $ulTextOffset >> 1;
                }

                $tTextBlock->ulCharPos = $ulTextOffset;
                $tTextBlock->ulLength = $ulTotLength; // character size
                $tTextBlock->bUsesUnicode = $bUsesUnicode;
                $tTextBlock->usPropMod = $usPropMod;
                $tTextBlock->startCP = self::getInt4d($aucBuffer, $lOff + $i * 4);
                $tTextBlock->endCP = self::getInt4d($aucBuffer, $lOff + ($i + 1) * 4);
                if ($bUsesUnicode) {
                    $lToGo = $ulTotLength * 2;
                } else {
                    $lToGo = $ulTotLength;
                }
                $tTextBlock->lToGo = $lToGo; // byte size
                $tTextBlockList[] = $tTextBlock;
            }
            break;
        }

        $clx->textBlockList = $tTextBlockList;

        return $clx;
    }

    /**
     * Read 16-bit unsigned integer (Big Endian)
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt2dBE($data, $pos)
    {
        return ord($data[$pos]) << 8 | ord($data[$pos + 1]);
    }

    /**
     * Read 8-bit unsigned integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt1d($data, $pos)
    {
        return ord($data[$pos]);
    }

    /**
     * Read 16-bit unsigned integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt2d($data, $pos)
    {
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8);
    }

    /**
     * Read 24-bit signed integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt3d($data, $pos)
    {
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16);
    }

    /**
     * Read 32-bit signed integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt4d($data, $pos)
    {
        // FIX: represent numbers correctly on 64-bit system
        // http://sourceforge.net/tracker/index.php?func=detail&aid=1487372&group_id=99160&atid=623334
        // Hacked by Andreas Rehm 2006 to ensure correct result of the <<24 block on 32 and 64bit systems
        $or24 = ord($data[$pos + 3]);
        if ($or24 >= 128) {
            // negative number
            $ord24 = -abs((256 - $or24) << 24);
        } else {
            $ord24 = ($or24 & 127) << 24;
        }
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | $ord24;
    }

    /**
     * The FibRgLw97 structure is the third section of the FIB. This contains an array of 4-byte values.
     *
     * https://msdn.microsoft.com/en-us/library/dd922774(v=office.12).aspx
     * @param $data
     * @param $pos
     */
    public static function GetFibRgLw97($data, &$pos)
    {
        $oFibRgLw97 = new \stdClass(); // Specifies the count of bytes of those written to the WordDocument stream of the file that have any meaning. All bytes in the WordDocument stream at offset cbMac and greater MUST be ignored.
        $oFibRgLw97->cbMac = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved1 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved2 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->ccpText = self::getInt4d($data, $pos); // A signed integer that specifies the count of CPs in the main document. This value MUST be zero, 1, or greater.
        $pos += 4;
        $oFibRgLw97->ccpFtn = self::getInt4d($data, $pos); // A signed integer that specifies the count of CPs in the footnote subdocument. This value MUST be zero, 1, or greater.
        $pos += 4;
        $oFibRgLw97->ccpHdd = self::getInt4d($data, $pos); // A signed integer that specifies the count of CPs in the header subdocument. This value MUST be zero, 1, or greater.
        $pos += 4;
        $oFibRgLw97->reserved3 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->ccpAtn = self::getInt4d($data, $pos); // A signed integer that specifies the count of CPs in the comment subdocument. This value MUST be zero, 1, or greater.
        $pos += 4;
        $oFibRgLw97->ccpEdn = self::getInt4d($data, $pos); // A signed integer that specifies the count of CPs in the endnote subdocument. This value MUST be zero, 1, or greater.
        $pos += 4;
        $oFibRgLw97->ccpTxbx = self::getInt4d($data, $pos); // A signed integer that specifies the count of CPs in the textbox subdocument of the main document. This value MUST be zero, 1, or greater.
        $pos += 4;
        $oFibRgLw97->ccpHdrTxbx = self::getInt4d($data, $pos); // A signed integer that specifies the count of CPs in the textbox subdocument of the header. This value MUST be zero, 1, or greater.
        $pos += 4;
        $oFibRgLw97->reserved4 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved5 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved6 = self::getInt4d($data, $pos); // This value MUST be equal or less than the number of data elements in PlcBteChpx, as specified by FibRgFcLcb97.fcPlcfBteChpx and FibRgFcLcb97.lcbPlcfBteChpx. This value MUST be ignored.
        $pos += 4;
        $oFibRgLw97->reserved7 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved8 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved9 = self::getInt4d($data, $pos); // This value MUST be less than or equal to the number of data elements in PlcBtePapx, as specified by FibRgFcLcb97.fcPlcfBtePapx and FibRgFcLcb97.lcbPlcfBtePapx. This value MUST be ignored.
        $pos += 4;
        $oFibRgLw97->reserved10 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved11 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved12 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved13 = self::getInt4d($data, $pos);
        $pos += 4;
        $oFibRgLw97->reserved14 = self::getInt4d($data, $pos);
    }

    private static function GetNumInLcb($lcb, $iSize)
    {
        return ($lcb - 4) / (4 + $iSize);
    }

    /**
     * Read PlcfBtePapx
     *
     * Paragraph and information about them
     *
     * https://msdn.microsoft.com/en-us/library/dd908569(v=office.12).aspx
     * @param $arrayFib
     * @param $wrkdocument
     * @param $data1Table
     * @param $dataData
     * @return array
     */
    public static function GetRecordPlcfBtePapx(&$arrayFib, &$wrkdocument, &$data1Table, &$dataData)
    {
        $arrayParagraphs = array();

        $pos = $arrayFib['fcPlcfBtePapx'];
        $isize = $arrayFib['lcbPlcfBtePapx'];

        $num = self::GetNumInLcb($isize, 4);

        $pos += 4 * ($num + 1); // skip aFCs

        $aPnBtePapx = array();

        for($i=0; $i < $num; $i++) {
            $aPnBtePapx[$i] = self::getInt4d($data1Table, $pos) & 0x3FFFFF; // 22bits, 10bit unused
            $pos += 4;
        }

        for($i = 0; $i < $num; $i++) {
            $offsetBase = $aPnBtePapx[$i] * 512; // offset: pn*512
            $offset = $offsetBase;

            $cpara = self::getInt1d($wrkdocument, $offset + 511);

            $arrayFCs = array();
            for ($j = 0; $j <= $cpara; $j++) { // the number of rgfcs is cpara + 1
                $arrayFCs[$j] = self::getInt4d($wrkdocument, $offset);
                $offset += 4;
            }

            $arrayRGBs = array();
            for ($j = 1; $j <= $cpara; $j++) {
                $arrayRGBs[$j] = self::getInt1d($wrkdocument, $offset);
                $offset += 1;
                $offset += 12; // reserved skipped (12 bytes)
            }

            $styles = array();
            for ($j = 1; $j <= $cpara; $j++) {
                $rgb = $arrayRGBs[$j];
                $offset = $offsetBase + ($rgb * 2);

                $cb = self::getInt1d($wrkdocument, $offset);
                $offset += 1;
                print_r('$cb : '.$cb.PHP_EOL);
                if ($cb == 0) {
                    $cb = self::getInt1d($wrkdocument, $offset);
                    $cb = $cb * 2;
                    $offset += 1;
                    print_r('$cb0 : '.$cb.PHP_EOL);
                } else {
                    $cb = $cb * 2 - 1;
                    print_r('$cbD : '.$cb.PHP_EOL);
                }
                $istd = self::getInt2d($wrkdocument, $offset);
                $offset += 2;
                $cb -= 2;

                if ($cb > 0) {
                    $styles[$j] = self::readPrl($wrkdocument, $dataData, $offset, $cb);
                } else {
                    $styles[$j] = null;
                }
            }

            $paragraph = new \stdClass();

            $paragraph->aFCs = $arrayFCs;
            $paragraph->aRGBs = $arrayRGBs;
            $paragraph->aStyles = $styles;

            $arrayParagraphs[] = $paragraph;
        }

        return $arrayParagraphs;
    }

    public static function GetRecordPlcfBteChpx($wrkdocument, $data1Table, $dataData, $pos, $isize)
    {
        $arrayCharacters = array();

        $num = self::GetNumInLcb($isize, 4);

        $pos += ($num + 1) * 4; // skip aFCs

        $aPnBteChpx = array();

        for ($i = 1; $i <= $num; $i++) {
            $aPnBteChpx[$i] = self::getInt4d($data1Table, $pos) & 0x3FFFFF; // 22bits, 10bit unused
            $pos += 4;
        }

        for ($i = 1; $i <= $num; $i++) {
            $offsetBase = $aPnBteChpx[$i] * 512;
            $offset = $offsetBase;

            $crun = self::getInt1d($wrkdocument, $offset + 511);

            $arrayFCs = array();
            for ($j = 0; $j <= $crun; $j++) {
                $arrayFCs[$j] = self::getInt4d($wrkdocument, $offset);
                $offset += 4;
            }

            $arrayRGBs = array();

            for ($j = 1; $j<= $crun; $j++) {
                $arrayRGBs[$j] = self::getInt1d($wrkdocument, $offset);
                $offset += 1;
            }

            $oStyles = array();

            for ($j = 1; $j<=$crun; $j++) {
                $oStyle = new \stdClass();
                $oStyle->start = $arrayFCs[$j - 1];
                $oStyle->len_in_bytes = $arrayFCs[$j] - $arrayFCs[$j - 1];

                $rgb = $arrayRGBs[$j];
                if ($rgb > 0) {
                    $posChpx = $offsetBase + $rgb * 2;

                    $cb = self::getInt1d($wrkdocument, $posChpx);
                    $posChpx += 1;

                    $oStyle->style = self::readPrl($wrkdocument, $dataData, $posChpx, $cb);
                }
                $oStyles[$j] = $oStyle;
            }

            $arrayCharacters[$i] = $oStyles;
        }

        return $arrayCharacters;
    }

    /**
     * @param $sprm
     * @return \stdClass
     */
    private static function readSprm($sprm)
    {
        $oSprm = new \stdClass();
        $oSprm->isPmd = $sprm & 0x01FF;
        $oSprm->f = ($sprm / 512) & 0x0001;
        $oSprm->sgc = ($sprm / 1024) & 0x0007;
        $oSprm->spra = ($sprm / 8192);
        return $oSprm;
    }

    /**
     * @param string $data
     * @param integer $pos
     * @param \stdClass $oSprm
     * @return array
     */
    private static function readSprmSpra($data, $pos, $oSprm)
    {
        $length = 0;
        $operand = null;

        switch(dechex($oSprm->spra)) {
            case 0x0:
                $operand = self::getInt1d($data, $pos);
                $length = 1;
                switch(dechex($operand)) {
                    case 0x00:
                        $operand = false;
                        break;
                    case 0x01:
                        $operand = true;
                        break;
                    case 0x80:
                        $operand = self::SPRA_VALUE;
                        break;
                    case 0x81:
                        $operand = self::SPRA_VALUE_OPPOSITE;
                        break;
                }
                break;
            case 0x1:
                $operand = self::getInt1d($data, $pos);
                $length = 1;
                break;
            case 0x2:
            case 0x4:
            case 0x5:
                $operand = self::getInt2d($data, $pos);
                $length = 2;
                break;
            case 0x3:
                if ($oSprm->isPmd != 0x70) {
                    $operand = self::getInt4d($data, $pos);
                    $length = 4;
                }
                break;
            case 0x7:
                $operand = self::getInt3d($data, $pos);
                $length = 3;
                break;
            default:
                // print_r('YO YO YO : '.PHP_EOL);
        }

        return array(
            'length' => $length,
            'operand' => $operand,
        );
    }

    /**
     * @param $data integer
     * @param $pos integer
     * @return \stdClass
     * @link http://msdn.microsoft.com/en-us/library/dd772849%28v=office.12%29.aspx
     */
    private static function readPrl($data, $dataData, $pos, $cbNum)
    {
        $posStart = $pos;
        $oStylePrl = new \stdClass();

        // Variables
        $sprmCPicLocation = null;
        $sprmCFData = null;
        $sprmCFSpec = null;

        do {
            // Variables
            $operand = null;

            $sprm = self::getInt2d($data, $pos);
            $oSprm = self::readSprm($sprm);
            $pos += 2;
            $cbNum -= 2;

            $arrayReturn = self::readSprmSpra($data, $pos, $oSprm);
            $pos += $arrayReturn['length'];
            $cbNum -= $arrayReturn['length'];
            $operand = $arrayReturn['operand'];

            switch(dechex($oSprm->sgc)) {
                // Paragraph property
                case 0x01:
                    switch($oSprm->isPmd) {
                        case 0x03: // sprmPJc80
                            switch($operand)
                            {
                                case 0:
                                    $oStylePrl->alignment = 'left';
                                    break;
                                case 1:
                                    $oStylePrl->alignment = 'center';
                                    break;
                                case 2:
                                    $oStylePrl->alignment = 'right';
                                    break;
                                case 3:
                                case 4:
                                case 5:
                                    $oStylePrl->alignment = 'justified';
                                    break;
                                default:
                                    $oStylePrl->alignment = 'left';
                            }
                            break;
                        case 0x0A: // sprmPIlvl
                            print_r('sprmPIlvl : '.$operand.PHP_EOL.PHP_EOL);
                            $oStylePrl->iLvl = $operand;
                            break;
                        case 0x0B: // sprmPIlfo, list format
                            print_r('sprm_IsPmd : ' . dechex($operand) . PHP_EOL.PHP_EOL);
                            $oStylePrl->iLfo = $operand;
                            if ($operand === 0x0000) {
                                $oStylePrl->bList = false;
                            } else if ($operand >= 0x0001 && $operand <= 0x07FE) {
                                $oStylePrl->bList = true;
                            } else if ($operand === 0xF801) {
                                $oStylePrl->bList = false;
                            } else if ($operand >= 0xF802 && $operand <= 0xFFFF) {
                                $oStylePrl->bList = true;
                                $oStylePrl->indentPreserved = true;
                            }
                            break;
                        default:
                            print_r('sprm_IsPmd : ' . dechex($oSprm->isPmd) .PHP_EOL.PHP_EOL);
                            break;
                    }
                    break;
                // Character property
                case 0x02:
                    if (!isset($oStylePrl->styleFont)) {
                        $oStylePrl->styleFont = array();
                    }
                    switch($oSprm->isPmd) {
                        // sprmCFRMarkIns
                        case 0x01:
                            break;
                        // sprmCFFldVanish
                        case 0x02:
                            break;
                        // sprmCPicLocation
                        case 0x03:
                            $sprmCPicLocation = $operand;
                            break;
                        // sprmCFData
                        case 0x06:
                            $sprmCFData = dechex($operand) == 0x00 ? false : true;
                            break;
                        // sprmCFItalic
                        case 0x36:
                            // By default, text is not italicized.
                            switch($operand) {
                                case false:
                                case true:
                                    $oStylePrl->styleFont['italic'] = $operand;
                                    break;
                                case self::SPRA_VALUE:
                                    $oStylePrl->styleFont['italic'] = false;
                                    break;
                                case self::SPRA_VALUE_OPPOSITE:
                                    $oStylePrl->styleFont['italic'] = true;
                                    break;
                            }
                            break;
                        // sprmCIstd
                        case 0x30:
                            //print_r('sprmCIstd : '.dechex($operand).PHP_EOL.PHP_EOL);
                            break;
                        // sprmCFBold
                        case 0x35:
                            // By default, text is not bold.
                            switch($operand) {
                                case false:
                                case true:
                                    $oStylePrl->styleFont['bold'] = $operand;
                                    break;
                                case self::SPRA_VALUE:
                                    $oStylePrl->styleFont['bold'] = false;
                                    break;
                                case self::SPRA_VALUE_OPPOSITE:
                                    $oStylePrl->styleFont['bold'] = true;
                                    break;
                            }
                            break;
                        // sprmCFStrike
                        case 0x37:
                            // By default, text is not struck through.
                            switch($operand) {
                                case false:
                                case true:
                                    $oStylePrl->styleFont['strikethrough'] = $operand;
                                    break;
                                case self::SPRA_VALUE:
                                    $oStylePrl->styleFont['strikethrough'] = false;
                                    break;
                                case self::SPRA_VALUE_OPPOSITE:
                                    $oStylePrl->styleFont['strikethrough'] = true;
                                    break;
                            }
                            break;
                        // sprmCKul
                        case 0x3E:
                            switch(dechex($operand)) {
                                case 0x00:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_NONE;
                                    break;
                                case 0x01:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_SINGLE;
                                    break;
                                case 0x02:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WORDS;
                                    break;
                                case 0x03:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOUBLE;
                                    break;
                                case 0x04:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTTED;
                                    break;
                                case 0x06:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_HEAVY;
                                    break;
                                case 0x07:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASH;
                                    break;
                                case 0x09:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTHASH;
                                    break;
                                case 0x0A:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTDOTDASH;
                                    break;
                                case 0x0B:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WAVY;
                                    break;
                                case 0x14:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTTEDHEAVY;
                                    break;
                                case 0x17:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASHHEAVY;
                                    break;
                                case 0x19:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTHASHHEAVY;
                                    break;
                                case 0x1A:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTDOTDASHHEAVY;
                                    break;
                                case 0x1B:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WAVYHEAVY;
                                    break;
                                case 0x27:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASHLONG;
                                    break;
                                case 0x2B:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WAVYDOUBLE;
                                    break;
                                case 0x37:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASHLONGHEAVY;
                                    break;
                                default:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_NONE;
                                    break;
                            }
                            break;
                        // sprmCIco
                        //@link http://msdn.microsoft.com/en-us/library/dd773060%28v=office.12%29.aspx
                        case 0x42:
                            switch(dechex($operand)) {
                                case 0x00:
                                case 0x01:
                                    $oStylePrl->styleFont['color'] = '000000';
                                    break;
                                case 0x02:
                                    $oStylePrl->styleFont['color'] = '0000FF';
                                    break;
                                case 0x03:
                                    $oStylePrl->styleFont['color'] = '00FFFF';
                                    break;
                                case 0x04:
                                    $oStylePrl->styleFont['color'] = '00FF00';
                                    break;
                                case 0x05:
                                    $oStylePrl->styleFont['color'] = 'FF00FF';
                                    break;
                                case 0x06:
                                    $oStylePrl->styleFont['color'] = 'FF0000';
                                    break;
                                case 0x07:
                                    $oStylePrl->styleFont['color'] = 'FFFF00';
                                    break;
                                case 0x08:
                                    $oStylePrl->styleFont['color'] = 'FFFFFF';
                                    break;
                                case 0x09:
                                    $oStylePrl->styleFont['color'] = '000080';
                                    break;
                                case 0x0A:
                                    $oStylePrl->styleFont['color'] = '008080';
                                    break;
                                case 0x0B:
                                    $oStylePrl->styleFont['color'] = '008000';
                                    break;
                                case 0x0C:
                                    $oStylePrl->styleFont['color'] = '800080';
                                    break;
                                case 0x0D:
                                    $oStylePrl->styleFont['color'] = '800080';
                                    break;
                                case 0x0E:
                                    $oStylePrl->styleFont['color'] = '808000';
                                    break;
                                case 0x0F:
                                    $oStylePrl->styleFont['color'] = '808080';
                                    break;
                                case 0x10:
                                    $oStylePrl->styleFont['color'] = 'C0C0C0';
                            }
                            break;
                        // sprmCHps
                        case 0x43:
                            $oStylePrl->styleFont['size'] = dechex($operand/2);
                            break;
                        // sprmCIss
                        case 0x48:
                            if (!isset($oStylePrl->styleFont['superScript'])) {
                                $oStylePrl->styleFont['superScript'] = false;
                            }
                            if (!isset($oStylePrl->styleFont['subScript'])) {
                                $oStylePrl->styleFont['subScript'] = false;
                            }
                            switch (dechex($operand)) {
                                case 0x00:
                                    // Normal text
                                    break;
                                case 0x01:
                                    $oStylePrl->styleFont['superScript'] = true;
                                    break;
                                case 0x02:
                                    $oStylePrl->styleFont['subScript'] = true;
                                    break;
                            }
                            break;
                        // sprmCRgFtc0
                        case 0x4F:
                            $oStylePrl->styleFont['name'] = '';
                            if (isset($arrayFonts[$operand])) {
                                $oStylePrl->styleFont['name'] = $arrayFonts[$operand]['main'];
                            }
                            break;
                        // sprmCRgFtc1
                        case 0x50:
                            // if the language for the text is an East Asian language
                            break;
                        // sprmCRgFtc2
                        case 0x51:
                            // if the character falls outside the Unicode character range
                            break;
                        // sprmCFSpec
                        case 0x55:
                            $sprmCFSpec = $operand;
                            break;
                        // sprmCFtcBi
                        case 0x5E:
                            break;
                        // sprmCFItalicBi
                        case 0x5D:
                            break;
                        // sprmCHpsBi
                        case 0x61:
                            break;
                        // sprmCShd80
                        //@link http://msdn.microsoft.com/en-us/library/dd923447%28v=office.12%29.aspx
                        case 0x66:
                            // $operand = self::getInt2d($data, $pos);
                            $pos += 2;
                            $cbNum -= 2;
                            // $ipat = ($operand >> 0) && bindec('111111');
                            // $icoBack = ($operand >> 6) && bindec('11111');
                            // $icoFore = ($operand >> 11) && bindec('11111');
                            break;
                        // sprmCCv
                        //@link : http://msdn.microsoft.com/en-us/library/dd952824%28v=office.12%29.aspx
                        case 0x70:
                            $red = str_pad(dechex(self::getInt1d($data, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $green = str_pad(dechex(self::getInt1d($data, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $blue = str_pad(dechex(self::getInt1d($data, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $pos += 1;
                            $oStylePrl->styleFont['color'] = $red.$green.$blue;
                            $cbNum -= 4;
                            break;
                        default:
                            // print_r('@todo Character : 0x'.dechex($oSprm->isPmd));
                            // print_r(PHP_EOL);
                    }
                    break;
                // Picture property
                case 0x03:
                    break;
                // Section property
                case 0x04:
                    if (!isset($oStylePrl->styleSection)) {
                        $oStylePrl->styleSection = array();
                    }
                    switch($oSprm->isPmd) {
                        // sprmSNfcPgn
                        case 0x0E:
                            // numbering format used for page numbers
                            break;
                        // sprmSXaPage
                        case 0x1F:
                            $oStylePrl->styleSection['pageSizeW'] = $operand;
                            break;
                        // sprmSYaPage
                        case 0x20:
                            $oStylePrl->styleSection['pageSizeH'] = $operand;
                            break;
                        // sprmSDxaLeft
                        case 0x21:
                            $oStylePrl->styleSection['marginLeft'] = $operand;
                            break;
                        // sprmSDxaRight
                        case 0x22:
                            $oStylePrl->styleSection['marginRight'] = $operand;
                            break;
                        // sprmSDyaTop
                        case 0x23:
                            $oStylePrl->styleSection['marginTop'] = $operand;
                            break;
                        // sprmSDyaBottom
                        case 0x24:
                            $oStylePrl->styleSection['marginBottom'] = $operand;
                            break;
                        // sprmSFBiDi
                        case 0x28:
                            // RTL layout
                            break;
                        // sprmSDxtCharSpace
                        case 0x30:
                            // characpter pitch
                            break;
                        // sprmSDyaLinePitch
                        case 0x31:
                            // line height
                            break;
                        // sprmSClm
                        case 0x32:
                            // document grid mode
                            break;
                        // sprmSTextFlow
                        case 0x33:
                            // text flow
                            break;
                        default:
                            // print_r('@todo Section : 0x'.dechex($oSprm->isPmd));
                            // print_r(PHP_EOL);

                    }
                    break;
                // Table property
                case 0x05:
                    break;
            }
        } while ($cbNum > 0);

        if (!is_null($sprmCPicLocation)) {
            if (!is_null($sprmCFData) && $sprmCFData == 0x01) {
                // NilPICFAndBinData
                //@todo Read Hyperlink structure
                /*$lcb = self::getInt4d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 4;
                $cbHeader = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // ignored
                $sprmCPicLocation += 62;
                // depending of the element
                // Hyperlink => HFD
                // HFD > bits
                $sprmCPicLocation += 1;
                // HFD > clsid
                $sprmCPicLocation += 16;
                // HFD > hyperlink
                //@link : http://msdn.microsoft.com/en-us/library/dd909835%28v=office.12%29.aspx
                $streamVersion = self::getInt4d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 4;
                $data = self::getInt4d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 4;
                $hlstmfAbsFromGetdataRel = ($data >> 9) & bindec('1');
                $hlstmfMonikerSavedAsStr = ($data >> 8) & bindec('1');
                $hlstmfHasFrameName = ($data >> 7) & bindec('1');
                $hlstmfHasCreationTime = ($data >> 6) & bindec('1');
                $hlstmfHasGUID = ($data >> 5) & bindec('1');
                $hlstmfHasDisplayName = ($data >> 4) & bindec('1');
                $hlstmfHasLocationStr = ($data >> 3) & bindec('1');
                $hlstmfSiteGaveDisplayName = ($data >> 2) & bindec('1');
                $hlstmfIsAbsolute = ($data >> 1) & bindec('1');
                $hlstmfHasMoniker = ($data >> 0) & bindec('1');
                for ($inc = 0; $inc <= 32; $inc++) {
                    echo ($data >> $inc) & bindec('1');
                }

                print_r('$hlstmfHasMoniker > '.$hlstmfHasMoniker.PHP_EOL);
                print_r('$hlstmfIsAbsolute > '.$hlstmfIsAbsolute.PHP_EOL);
                print_r('$hlstmfSiteGaveDisplayName > '.$hlstmfSiteGaveDisplayName.PHP_EOL);
                print_r('$hlstmfHasLocationStr > '.$hlstmfHasLocationStr.PHP_EOL);
                print_r('$hlstmfHasDisplayName > '.$hlstmfHasDisplayName.PHP_EOL);
                print_r('$hlstmfHasGUID > '.$hlstmfHasGUID.PHP_EOL);
                print_r('$hlstmfHasCreationTime > '.$hlstmfHasCreationTime.PHP_EOL);
                print_r('$hlstmfHasFrameName > '.$hlstmfHasFrameName.PHP_EOL);
                print_r('$hlstmfMonikerSavedAsStr > '.$hlstmfMonikerSavedAsStr.PHP_EOL);
                print_r('$hlstmfAbsFromGetdataRel > '.$hlstmfAbsFromGetdataRel.PHP_EOL);
                if ($streamVersion == 2) {
                    $AAA = self::getInt4d($this->dataData, $sprmCPicLocation);
                    echo 'AAAA : '.$AAA.PHP_EOL;
                    if ($hlstmfHasDisplayName == 1) {
                        echo 'displayName'.PHP_EOL;
                    }
                    if ($hlstmfHasFrameName == 1) {
                        echo 'targetFrameName'.PHP_EOL;
                    }
                    if ($hlstmfHasMoniker == 1 || $hlstmfMonikerSavedAsStr == 1) {
                        $sprmCPicLocation += 16;
                        $length = self::getInt4d($this->dataData, $sprmCPicLocation);
                        $sprmCPicLocation += 4;
                        for ($inc = 0; $inc < ($length / 2); $inc++) {
                            $chr = self::getInt2d($this->dataData, $sprmCPicLocation);
                            $sprmCPicLocation += 2;
                            print_r(chr($chr));
                        }
                        echo PHP_EOL;
                        echo 'moniker : '.$length.PHP_EOL;
                    }
                    if ($hlstmfHasMoniker == 1 || $hlstmfMonikerSavedAsStr == 1) {
                        echo 'oleMoniker'.PHP_EOL;
                    }
                    if ($hlstmfHasLocationStr == 1) {
                        echo 'location'.PHP_EOL;
                    }
                    if ($hlstmfHasGUID == 1) {
                        echo 'guid'.PHP_EOL;
                        $sprmCPicLocation += 16;
                    }
                    if ($hlstmfHasCreationTime == 1) {
                        echo 'fileTime'.PHP_EOL;
                        $sprmCPicLocation += 4;
                    }
                    echo 'HYPERLINK'.PHP_EOL;
                }*/
            } else {
                // Pictures
                //@link : http://msdn.microsoft.com/en-us/library/dd925458%28v=office.12%29.aspx
                //@link : http://msdn.microsoft.com/en-us/library/dd926136%28v=office.12%29.aspx
                // PICF : lcb
                $sprmCPicLocation += 4;
                // PICF : cbHeader
                $sprmCPicLocation += 2;
                // PICF : mfpf : mm
                $mfpfMm = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : mfpf : xExt
                $sprmCPicLocation += 2;
                // PICF : mfpf : yExt
                $sprmCPicLocation += 2;
                // PICF : mfpf : swHMF
                $sprmCPicLocation += 2;
                // PICF : innerHeader : grf
                $sprmCPicLocation += 4;
                // PICF : innerHeader : padding1
                $sprmCPicLocation += 4;
                // PICF : innerHeader : mmPM
                $sprmCPicLocation += 2;
                // PICF : innerHeader : padding2
                $sprmCPicLocation += 4;
                // PICF : picmid : dxaGoal
                $picmidDxaGoal = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dyaGoal
                $picmidDyaGoal = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : mx
                $picmidMx = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : my
                $picmidMy = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dxaReserved1
                $picmidDxaCropLeft = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dyaReserved1
                $picmidDxaCropTop = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dxaReserved2
                $picmidDxaCropRight = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dyaReserved2
                $picmidDxaCropBottom = self::getInt2d($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : fReserved
                $sprmCPicLocation += 1;
                // PICF : picmid : bpp
                $sprmCPicLocation += 1;
                // PICF : picmid : brcTop80
                $sprmCPicLocation += 4;
                // PICF : picmid : brcLeft80
                $sprmCPicLocation += 4;
                // PICF : picmid : brcBottom80
                $sprmCPicLocation += 4;
                // PICF : picmid : brcRight80
                $sprmCPicLocation += 4;
                // PICF : picmid : dxaReserved3
                $sprmCPicLocation += 2;
                // PICF : picmid : dyaReserved3
                $sprmCPicLocation += 2;
                // PICF : cProps
                $sprmCPicLocation += 2;

                if ($mfpfMm == 0x0066) {
                    // cchPicName
                    $cchPicName = self::getInt1d($dataData, $sprmCPicLocation);
                    $sprmCPicLocation += 1;

                    // stPicName
                    $stPicName = '';
                    for ($inc = 0; $inc <= $cchPicName; $inc++) {
                        $chr = self::getInt1d($dataData, $sprmCPicLocation);
                        $sprmCPicLocation += 1;
                        $stPicName .= chr($chr);
                    }
                }

                // picture (OfficeArtInlineSpContainer)
                // picture : shape
                $shapeRH = self::loadRecordHeader($dataData, $sprmCPicLocation);
                $sprmCPicLocation += 8;
                if ($shapeRH['recVer'] == 0xF && $shapeRH['recInstance'] == 0x000 && $shapeRH['recType'] == 0xF004) {
                    $sprmCPicLocation += $shapeRH['recLen'];
                }
                // picture : rgfb
                //@link : http://msdn.microsoft.com/en-us/library/dd950560%28v=office.12%29.aspx
                $fileBlockRH = self::loadRecordHeader($dataData, $sprmCPicLocation);
                while ($fileBlockRH['recType'] == 0xF007 || ($fileBlockRH['recType'] >= 0xF018 && $fileBlockRH['recType'] <= 0xF117)) {
                    $sprmCPicLocation += 8;
                    switch ($fileBlockRH['recType']) {
                        // OfficeArtFBSE
                        //@link : http://msdn.microsoft.com/en-us/library/dd944923%28v=office.12%29.aspx
                        case 0xF007:
                            // btWin32
                            $sprmCPicLocation += 1;
                            // btMacOS
                            $sprmCPicLocation += 1;
                            // rgbUid
                            $sprmCPicLocation += 16;
                            // tag
                            $sprmCPicLocation += 2;
                            // size
                            $sprmCPicLocation += 4;
                            // cRef
                            $sprmCPicLocation += 4;
                            // foDelay
                            $sprmCPicLocation += 4;
                            // unused1
                            $sprmCPicLocation += 1;
                            // cbName
                            $cbName = self::getInt1d($dataData, $sprmCPicLocation);
                            $sprmCPicLocation += 1;
                            // unused2
                            $sprmCPicLocation += 1;
                            // unused3
                            $sprmCPicLocation += 1;
                            // nameData
                            if ($cbName > 0) {
                                $nameData = '';
                                for ($inc = 0; $inc <= ($cbName / 2); $inc++) {
                                    $chr = self::getInt2d($dataData, $sprmCPicLocation);
                                    $sprmCPicLocation += 2;
                                    $nameData .= chr($chr);
                                }
                            }
                            // embeddedBlip
                            //@link : http://msdn.microsoft.com/en-us/library/dd910081%28v=office.12%29.aspx
                            $embeddedBlipRH = self::loadRecordHeader($dataData, $sprmCPicLocation);
                            switch ($embeddedBlipRH['recType']) {
                                case self::OFFICEARTBLIPJPG:
                                case self::OFFICEARTBLIPJPEG:
                                    if (!isset($oStylePrl->image)) {
                                        $oStylePrl->image = array();
                                    }
                                    $sprmCPicLocation += 8;
                                    // embeddedBlip : rgbUid1
                                    $sprmCPicLocation += 16;
                                    if ($embeddedBlipRH['recInstance'] == 0x6E1) {
                                        // rgbUid2
                                        $sprmCPicLocation += 16;
                                    }
                                    // embeddedBlip : tag
                                    $sprmCPicLocation += 1;
                                    // embeddedBlip : BLIPFileData
                                    $oStylePrl->image['data'] = substr($dataData, $sprmCPicLocation, $embeddedBlipRH['recLen']);
                                    $oStylePrl->image['format'] = 'jpg';
                                    // Image Size
                                    $iCropWidth = $picmidDxaGoal - ($picmidDxaCropLeft + $picmidDxaCropRight);
                                    $iCropHeight = $picmidDyaGoal - ($picmidDxaCropTop + $picmidDxaCropBottom);
                                    if (!$iCropWidth) {
                                        $iCropWidth = 1;
                                    }
                                    if (!$iCropHeight) {
                                        $iCropHeight = 1;
                                    }
                                    $oStylePrl->image['width'] = Drawing::twipsToPixels($iCropWidth * $picmidMx / 1000);
                                    $oStylePrl->image['height'] = Drawing::twipsToPixels($iCropHeight * $picmidMy / 1000);

                                    $sprmCPicLocation += $embeddedBlipRH['recLen'];
                                    break;
                                default:
                                    // print_r(dechex($embeddedBlipRH['recType']));
                            }
                            break;
                    }
                    $fileBlockRH = self::loadRecordHeader($dataData, $sprmCPicLocation);
                }
            }
        }

        $oStylePrl->length = $pos - $posStart;
        return $oStylePrl;
    }

    /**
     * Read a record header
     * @param string $stream
     * @param integer $pos
     * @return array
     */
    private static function loadRecordHeader($stream, $pos)
    {
        $rec = self::getInt2d($stream, $pos);
        $recType = self::getInt2d($stream, $pos + 2);
        $recLen = self::getInt4d($stream, $pos + 4);
        return array(
            'recVer' => ($rec >> 0) & bindec('1111'),
            'recInstance' => ($rec >> 4) & bindec('111111111111'),
            'recType' => $recType,
            'recLen' => $recLen,
        );
    }

    public static function buildLfoList(&$arrayFib, &$data1Table)
    {
        /* LFO (List Format Override) */

        $lfoList = array();

        $ulBeginLfoInfo = $arrayFib['fcPlfLfo'];
        $tLfoInfoLen = $arrayFib['lcbPlfLfo'];

        echo PHP_EOL . 'fcPlfLfo: ' . $ulBeginLfoInfo . ', lcbPlfLfo: ' . $tLfoInfoLen . PHP_EOL;

        $aucLfoInfo = substr($data1Table, $ulBeginLfoInfo, $tLfoInfoLen);

        $pos = 0;
        $tRecords = self::getInt4d($aucLfoInfo, $pos);

        if (4 + 16 * $tRecords > $tLfoInfoLen || $tRecords >= 0x7fff) {
            /* Just a sanity check */
            return false;
        }

        for ($i = 0; $i < $tRecords; $i++) {
            $pos = 4 + 16 * $i;
            $lfo = new \stdClass();
            $lfo->lsid = self::getInt4d($aucLfoInfo, $pos);
            $pos += 4;
            $lfo->unused1 = self::getInt4d($aucLfoInfo, $pos);
            $pos += 4;
            $lfo->unused2 = self::getInt4d($aucLfoInfo, $pos);
            $pos += 4;
            $lfo->clfolvl = self::getInt1d($aucLfoInfo, $pos); // An unsigned integer that specifies the count of LFOLVL elements that are stored in the rgLfoLvl field of the LFOData element that corresponds to this LFO structure.
            $pos += 1;
            /*
             Value
Meaning
0x00
This LFO is not used for any field. The fAutoNum of the related LSTF MUST be set to 0.
0xFC
This LFO is used for the AUTONUMLGL field (see AUTONUMLGL in flt). The fAutoNum of the related LSTF MUST be set to 1.
0xFD
This LFO is used for the AUTONUMOUT field (see AUTONUMOUT in flt). The fAutoNum of the related LSTF MUST be set to 1.
0xFE
This LFO is used for the AUTONUM field (see AUTONUM in flt). The fAutoNum of the related LSTF MUST be set to 1.
0xFF
This LFO is not used for any field. The fAutoNum of the related LSTF MUST be set to 0.

             */
            $lfo->ibstFltAutoNum = self::getInt1d($aucLfoInfo, $pos); // An unsigned integer that specifies the field that this LFO represents.
            $pos += 1;
            $lfo->grfhic = self::getInt1d($aucLfoInfo, $pos);
            $pos += 1;
            $lfo->unused3 = self::getInt1d($aucLfoInfo, $pos);
            $pos += 4;

            $lfoList[$i] = $lfo;
        }

        return $lfoList;
    }

    /**
     * Read Comments From File
     * @param $arrayFib
     * @param $data1Table
     * @return array
     */
    public static function ReadComments(&$arrayFib, &$data1Table)
    {
        $comments = array();
        $offset = $arrayFib['fcPlcfandTxt'];
        $lcbPlcfandTxt = $arrayFib['lcbPlcfandTxt'];

        if ($lcbPlcfandTxt <= 0) {
            return false;
        }

        $numCPs = (int) ($lcbPlcfandTxt / 4);

        $aCPs = array();
        for ($i = 0; $i < $numCPs; $i++) {
            $aCPs[$i] = self::getInt4d($data1Table, $offset);
            $offset += 4;
        }

        if ($aCPs[$numCPs - 2] !== $arrayFib['ccpAtn'] - 1) {
            echo "invalid second-to-last cp : " . $aCPs[$numCPs - 2] . ", ccpAtn: " . $arrayFib['ccpAtn'] . PHP_EOL;
        }

        /*for ($i = 1; $i < $numCPs; $i++) {
            $aCPs[$i] -= 1;
        }*/

        echo "comment cps: " . json_encode($aCPs) . PHP_EOL;

        // Read XSTs
        $fcGrpXstAtnOwners = $arrayFib['fcGrpXstAtnOwners'];
        $lcbGrpXstAtnOwners = $arrayFib['lcbGrpXstAtnOwners']; // An unsigned integer that specifies the size, in bytes, of the XST array

        $offset = $fcGrpXstAtnOwners;

        $author_names = array();

        while ($lcbGrpXstAtnOwners > 0) {
            $cch = self::getInt2d($data1Table, $offset);
            $offset += 2;
            $author_name = substr($data1Table, $offset, $cch * 2);
            $author_name = iconv('UCS-2LE', 'UTF-8', $author_name);
            $author_names[] = $author_name;

            $offset += $cch * 2;

            $lcbGrpXstAtnOwners -= 2 + 2 * $cch;
        }

        $fcPlcfandRef = $arrayFib['fcPlcfandRef'];
        $lcbPlcfandRef = $arrayFib['lcbPlcfandRef'];

        if ($lcbPlcfandRef <= 0) {
            return false;
        }

        $numRefCPs = self::GetNumInLcb($lcbPlcfandRef, 30);

        $refCPs = array();
        $offset = $fcPlcfandRef;
        for ($i = 0; $i <= $numRefCPs; $i++) {
            $refCPs[$i] = self::getInt4d($data1Table, $offset);
            $offset += 4;
        }

        echo "ref cps: " . json_encode($refCPs) . PHP_EOL;
        $aATRDPre10 = array();
        for ($i = 1; $i <= $numRefCPs; $i++) {
            $cch = self::getInt2d($data1Table, $offset);
            if ($cch > 9) {
                echo "bad file" . PHP_EOL;
                return false;
            }
            $initial_author = substr($data1Table, $offset + 2, $cch * 2);
            $initial_author = iconv('UCS-2LE', 'UTF-8', $initial_author);

            $xstIndex = self::getInt2d($data1Table, $offset + 20);
            $lTagBkmk = self::getInt4d($data1Table, $offset + 26);
            echo "initial_author: " . $initial_author . ", xst index: $xstIndex" . ", author name: " . $author_names[$xstIndex] . PHP_EOL;
            $offset += 30;

            $aATRDPre10[$i] = array(
                'initial_author'    => $initial_author,
                'bkmk_id'           => $lTagBkmk,
                'xst_index'         => $xstIndex,
                'author_name'       => $author_names[$xstIndex],
            );
        }

        // AtrdExtra
        $fcAtrdExtra = $arrayFib['fcAtrdExtra'];
        $lcbAtrdExtra = $arrayFib['lcbAtrdExtra'];

        $aAtrdExtra = array();
        if ($lcbAtrdExtra > 0) {
            $numAtrdExtra = $lcbAtrdExtra / 18;
            if ($numAtrdExtra !== $numRefCPs) {
                echo "invalid file - numAtrdExtra: {$numAtrdExtra}, numRefCPs: {$numRefCPs}" . PHP_EOL;
                return false;
            }

            $offset = $fcAtrdExtra;

            for ($i = 1; $i <= $numAtrdExtra; $i++) {
                $dttm = self::getInt4d($data1Table, $offset);
                $minute = $dttm & 0x3F;
                $hr = ($dttm >> 6) & 0x1F;
                $dom = ($dttm >> 11) & 0x1F;
                $mon = ($dttm >> 16) & 0xF;
                $yr = 1900 + (($dttm >> 20) & 0x1FF);
                $wdy = ($dttm >> 29) & 0x7;
                $offset += 4;
                // padding1
                $offset += 2;
                $cDepth = self::getInt4d($data1Table, $offset);
                $offset += 4;
                $diatrdParent = self::getInt4d($data1Table, $offset);
                $offset += 4;
                $offset += 4;
                echo "$minute $hr $dom $mon $yr $wdy, cDepth: {$cDepth}" . PHP_EOL;
                $aAtrdExtra[$i] = array(
                    'minute'    => $minute,
                    'hr'        => $hr,
                    'dom'       => $dom,
                    'mon'       => $mon,
                    'yr'        => $yr,
                    'wdy'       => $wdy,
                    'cDepth'    => $cDepth,
                    'diatrdParent'  => $diatrdParent,
                );
            }
        }

        for ($i = 0; $i < $numCPs - 2; $i++) {
            $comment = array();

            $comment = array_merge($comment, $aATRDPre10[$i + 1]);

            if (isset($aAtrdExtra[$i + 1]) && !empty($aAtrdExtra[$i + 1])) {
                $comment = array_merge($comment, $aAtrdExtra[$i + 1]);
            }

            $comment['start_cp'] = $aCPs[$i];
            $comment['length'] = $aCPs[$i + 1] - $aCPs[$i];
            $comment['atn_cp'] = $refCPs[$i];

            $comments[$i] = $comment;
        }

        echo "comments: " . json_encode($comments) . PHP_EOL;
        return $comments;
    }
}