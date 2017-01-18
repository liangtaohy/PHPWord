<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2016 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Reader;

use PhpOffice\Common\Drawing;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style;
use PhpOffice\PhpWord\Shared\OLERead;

define('P_HEADER_SZ',		28);
define('P_SECTIONLIST_SZ',	20);
define('P_LENGTH_SZ',		 4);
define('P_SECTION_MAX_SZ',	(2 * P_SECTIONLIST_SZ + P_LENGTH_SZ));
//define P_SECTION_SZ(x)		((x) * P_SECTIONLIST_SZ + P_LENGTH_SZ)

define('PID_CODEPAGE', 1); // code page定义
define('PID_TITLE',		 2);
define('PID_SUBJECT',    3);
define('PID_AUTHOR',		 4);
define('PID_CREATE_DTM',		12);
define('PID_LASTSAVE_DTM',	13);
define('PID_APPNAME',		18);
define('PIDSI_PAGECOUNT', 0x0E);
define('PIDSI_WORDCOUNT', 0x0F);
define('PIDSI_CHARCOUNT', 0x10);

define('PIDD_MANAGER',		14);
define('PIDD_COMPANY',		15);

/**
 * VT_ definitions
 * https://msdn.microsoft.com/en-us/library/aa380072(v=vs.85).aspx
 */
define('VT_LPSTR',		30);
define('VT_FILETIME',		64);
define('VT_I2', 2); // Two bytes representing a 2-byte signed integer value.
define('VT_I4', 3); // 4-byte signed integer value.
define('VT_INT', 22); // 4-byte signed integer value (equivalent to VT_I4).
define('VT_UINT', 23); // 4-byte unsigned integer (equivalent to VT_UI4).
define('VT_DATE', 7); // A 64-bit floating point number representing the number of days (not seconds) since December 31, 1899. For example, January 1, 1900, is 2.0, January 2, 1900, is 3.0, and so on). This is stored in the same representation as VT_R8.
define('VT_CLSID', 72); // Pointer to a class identifier (CLSID) (or other globally unique identifier (GUID)).

define('TIME_OFFSET_HI', 0x019db1de);
define('TIME_OFFSET_LO', 0xd53e8000);

define('ECHO_DEBUG_ENABLE', 0);

/**
 * Reader for Word97
 *
 * @since 0.10.0
 */
class MsDoc extends AbstractReader implements ReaderInterface
{
    /**
     * PhpWord object
     *
     * @var PhpWord
     */
    private $phpWord;

    /**
     * WordDocument Stream
     *
     * @var
     */
    private $dataWorkDocument;
    /**
     * 1Table Stream
     *
     * @var
     */
    private $data1Table;
    /**
     * Data Stream
     *
     * @var
     */
    private $dataData;
    /**
     * Object Pool Stream
     *
     * @var
     */
    private $dataObjectPool;
    /**
     * @var \stdClass[]
     */
    private $arrayCharacters = array();
    /**
     * @var array
     */
    private $arrayFib = array();
    /**
     * @var string[]
     */
    private $arrayFonts = array();
    /**
     * @var string[]
     */
    private $arrayParagraphs = array();
    /**
     * @var \stdClass[]
     */
    private $arraySections = array();

    /**
     * @var array
     */
    private $docInfo = array();

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

    const BIT_30 = 0x40000000;
    const BIG_BLOCK_SIZE = 512;

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

    public static $CodePage = array(
        936   => 'CP936',
        10008   => 'GB2312', // x-mac-chinesesimp	MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
    );

    /**
     * Loads PhpWord from file
     *
     * @param string $filename
     * @return PhpWord
     */
    public function load($filename)
    {
        $this->phpWord = new PhpWord();

        $this->loadOLE($filename);


        $docInfo = $this->getSystemInformation($this->_SummaryInformation);
        $docInfo += $this->getDocumentSummaryInfo($this->_DocumentSummaryInformation);
        $this->readFib($this->dataWorkDocument);
        $this->readFibContent();
        $this->setDocInfo($docInfo);
        /*
         * dump Txt
         */
        //echo $this->generateTxt();

        return $this->phpWord;
    }

    /**
     * Load an OLE Document
     * @param string $filename
     */
    private function loadOLE($filename)
    {
        // OLE reader
        $ole = new OLERead();
        $ole->read($filename);

        // Get WorkDocument stream
        $this->dataWorkDocument = $ole->getStream($ole->wrkdocument);
        // Get 1Table stream
        $this->data1Table = $ole->getStream($ole->wrk1Table);
        // Get Data stream
        $this->dataData = $ole->getStream($ole->wrkData);
        // Get Data stream
        $this->dataObjectPool = $ole->getStream($ole->wrkObjectPool);
        // Get Summary Information data
        $this->_SummaryInformation = $ole->getStream($ole->summaryInformation);
        // Get Document Summary Information data
        $this->_DocumentSummaryInformation = $ole->getStream($ole->docSummaryInfos);
    }

    private function setDocInfo(&$docInfo)
    {
        $this->phpWord->getDocInfo()->setTitle(isset($docInfo['title']) ? $docInfo['title'] : '');
        $this->phpWord->getDocInfo()->setSubject(isset($docInfo['subject']) ? $docInfo['subject'] : '');
        $this->phpWord->getDocInfo()->setCompany(isset($docInfo['company']) ? $docInfo['company'] : '');
        $this->phpWord->getDocInfo()->setManager(isset($docInfo['manager']) ? $docInfo['manager'] : '');
        $this->phpWord->getDocInfo()->setCreator(isset($docInfo['author']) ? $docInfo['author'] : '');
        $this->phpWord->getDocInfo()->setCreated(isset($docInfo['created']) ? $docInfo['created'] : '');
        $this->phpWord->getDocInfo()->setModified(isset($docInfo['lastModified']) ? $docInfo['lastModified'] : '');

        $this->phpWord->getDocInfo()->setMainStreamSize($this->arrayFib['ccpText']);
        $this->phpWord->getDocInfo()->setCommentSize($this->arrayFib['ccpAtn']);
    }

    private function getNumInLcb($lcb, $iSize)
    {
        return ($lcb - 4) / (4 + $iSize);
    }

    private function getArrayCP($data, $posMem, $iNum)
    {
        $arrayCP = array();
        for ($inc = 0; $inc < $iNum; $inc++) {
            $arrayCP[$inc] = self::getInt4d($data, $posMem);
            $posMem += 4;
        }
        return $arrayCP;
    }

    /**
     *
     * @link http://msdn.microsoft.com/en-us/library/dd949344%28v=office.12%29.aspx
     * @link https://igor.io/2012/09/24/binary-parsing.html
     */
    private function readFib($data)
    {
        $pos = 0;
        //----- FibBase
        // wIdent
        $pos += 2;
        // nFib
        $pos += 2;
        // unused
        $pos += 2;
        // lid : Language Identifier
        $pos += 2;
        // pnNext
        $pos += 2;

        // $mem = self::getInt2d($data, $pos);
        // $fDot = ($mem >> 15) & 1;
        // $fGlsy = ($mem >> 14) & 1;
        // $fComplex = ($mem >> 13) & 1;
        // $fHasPic = ($mem >> 12) & 1;
        // $cQuickSaves = ($mem >> 8) & bindec('1111');
        // $fEncrypted = ($mem >> 7) & 1;
        // $fWhichTblStm = ($mem >> 6) & 1;
        // $fReadOnlyRecommended = ($mem >> 5) & 1;
        // $fWriteReservation = ($mem >> 4) & 1;
        // $fExtChar = ($mem >> 3) & 1;
        // $fLoadOverride = ($mem >> 2) & 1;
        // $fFarEast = ($mem >> 1) & 1;
        // $fObfuscated = ($mem >> 0) & 1;
        $pos += 2;
        // nFibBack
        $pos += 2;
        // lKey
        $pos += 4;
        // envr
        $pos += 1;

        // $mem = self::getInt1d($data, $pos);
        // $fMac = ($mem >> 7) & 1;
        // $fEmptySpecial = ($mem >> 6) & 1;
        // $fLoadOverridePage = ($mem >> 5) & 1;
        // $reserved1 = ($mem >> 4) & 1;
        // $reserved2 = ($mem >> 3) & 1;
        // $fSpare0 = ($mem >> 0) & bindec('111');
        $pos += 1;

        // reserved3
        $pos += 2;
        // reserved4
        $pos += 2;
        // reserved5
        $pos += 4;
        // reserved6
        $pos += 4;

        //----- csw
        $pos += 2;

        //----- fibRgW
        // reserved1
        $pos += 2;
        // reserved2
        $pos += 2;
        // reserved3
        $pos += 2;
        // reserved4
        $pos += 2;
        // reserved5
        $pos += 2;
        // reserved6
        $pos += 2;
        // reserved7
        $pos += 2;
        // reserved8
        $pos += 2;
        // reserved9
        $pos += 2;
        // reserved10
        $pos += 2;
        // reserved11
        $pos += 2;
        // reserved12
        $pos += 2;
        // reserved13
        $pos += 2;
        // lidFE
        $pos += 2;

        //----- cslw
        $pos += 2;

        //----- fibRgLw
        // cbMac
        $pos += 4;
        // reserved1
        $pos += 4;
        // reserved2
        $pos += 4;
        $this->arrayFib['ccpText'] = self::getInt4d($data, $pos);
        $pos += 4;
        $this->arrayFib['ccpFtn'] = self::getInt4d($data, $pos);
        $pos += 4;
        $this->arrayFib['ccpHdd'] = self::getInt4d($data, $pos);
        $pos += 4;
        // reserved3
        $pos += 4;
        // ccpAtn
        $this->arrayFib['ccpAtn'] = self::getInt4d($data, $pos);
        $pos += 4;
        // ccpEdn
        $this->arrayFib['ccpEdn'] = self::getInt4d($data, $pos);
        $pos += 4;
        // ccpTxbx
        $this->arrayFib['ccpTxbx'] = self::getInt4d($data, $pos);
        $pos += 4;
        // ccpHdrTxbx
        $this->arrayFib['ccpHdrTxbx'] = self::getInt4d($data, $pos);
        $pos += 4;
        // reserved4
        $pos += 4;
        // reserved5
        $pos += 4;
        // reserved6
        $pos += 4;
        // reserved7
        $pos += 4;
        // reserved8
        $pos += 4;
        // reserved9
        $pos += 4;
        // reserved10
        $pos += 4;
        // reserved11
        $pos += 4;
        // reserved12
        $pos += 4;
        // reserved13
        $pos += 4;
        // reserved14
        $pos += 4;

        //----- cbRgFcLcb
        $cbRgFcLcb = self::getInt2d($data, $pos);
        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo "cbRgFcLcb: $pos," . dechex($cbRgFcLcb) . "\n";
        }
        $pos += 2;
        //----- fibRgFcLcbBlob
        switch ($cbRgFcLcb) {
            case 0x005D:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                break;
            case 0x006C:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                break;
            case 0x0088:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2002);
                break;
            case 0x00A4:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2002);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2003);
                break;
            case 0x00B7:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2002);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2003);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2007);
                break;
        }
        //----- cswNew
        $this->arrayFib['cswNew'] = self::getInt2d($data, $pos);
        $pos += 2;

        if ($this->arrayFib['cswNew'] != 0) {
            //@todo : fibRgCswNew
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "\nneed read fibRgCswNew\n";
            }
        }

        return $pos;
    }

    private function readBlockFibRgFcLcb($data, $pos, $version)
    {
        if ($version == self::VERSION_97) {
            $this->arrayFib['fcStshfOrig'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbStshfOrig'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcStshf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbStshf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffndRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffndRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffndTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffndTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfandRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfandRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfandTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfandTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfSed'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfSed'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcPad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcPad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfPhe'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfPhe'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfHdd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfHdd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBteChpx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBteChpx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBtePapx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBtePapx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfSea'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfSea'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfFfn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfFfn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldAtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldAtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcCmds'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbCmds'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPrDrvr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPrDrvr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPrEnvPort'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPrEnvPort'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPrEnvLand'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPrEnvLand'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcWss'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbWss'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDop'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDop'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfAssoc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfAssoc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcClx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbClx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAutosaveSource'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAutosaveSource'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcGrpXstAtnOwners'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbGrpXstAtnOwners'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfAtnBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfAtnBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcSpaMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcSpaMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcSpaHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcSpaHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfAtnBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfAtnBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfAtnBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfAtnBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPms'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPms'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcFormFldSttbs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbFormFldSttbs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfendRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfendRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfendTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfendTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDggInfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDggInfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfRMark'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfRMark'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfAutoCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfAutoCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfWkb'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfWkb'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfSpl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfSpl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcftxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcftxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfHdrtxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfHdrtxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffldHdrTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffldHdrTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcStwUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbStwUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbTtmbd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbTtmbd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcCookieData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbCookieData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfIntlFld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfIntlFld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRouteSlip'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRouteSlip'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbSavedBy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbSavedBy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbFnm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbFnm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlfLst'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlfLst'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlfLfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlfLfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfTxbxBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfTxbxBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfTxbxHdrBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfTxbxHdrBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDocUndoWord9'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDocUndoWord9'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUskf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUskf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcupcRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcupcRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcupcUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcupcUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbGlsyStyle'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbGlsyStyle'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlgosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlgosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcocx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcocx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBteLvc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBteLvc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['dwLowDateTime'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['dwHighDateTime'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfLvcPre10'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfLvcPre10'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfAsumy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfAsumy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfGram'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfGram'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbListNames'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbListNames'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfUssr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfUssr'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2000) {
            $this->arrayFib['fcPlcfTch'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfTch'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRmdThreading'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRmdThreading'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcMid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbMid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbRgtplc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbRgtplc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcMsoEnvelope'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbMsoEnvelope'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfLad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfLad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRgDofr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRgDofr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfCookieOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfCookieOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2002) {
            $this->arrayFib['fcUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfPgp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfPgp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfuim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfuim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlfguidUim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlfguidUim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAtrdExtra'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAtrdExtra'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlrsid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlrsid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfcookie'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfcookie'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcFactoidData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbFactoidData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDocUndo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDocUndo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfbkmkBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfbkmkBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfbkfBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfbkfBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfbklBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfbklBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPmsNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPmsNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcODSO'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbODSO'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2003) {
            $this->arrayFib['fcHplxsdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbHplxsdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcCustomXForm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbCustomXForm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbProtUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbProtUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfd'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2007) {
            $this->arrayFib['fcPlcfmthd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfmthd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcArtoData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbArtoData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused5'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused5'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused6'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused6'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcOssTheme'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbOssTheme'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcColorSchemeMapping'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbColorSchemeMapping'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        return $pos;
    }

    private function readFibContent()
    {
        // Informations about Font
        $this->readRecordSttbfFfn();

        // Informations about page
        $this->readRecordPlcfSed();

        // reading paragraphs
        //@link https://github.com/notmasteryet/CompoundFile/blob/ec118f354efebdee9102e41b5b7084fce81125b0/WordFileReader/WordDocument.cs#L86
        // $this->readRecordPlcfBtePapx();

        // reading character formattings
        //@link https://github.com/notmasteryet/CompoundFile/blob/ec118f354efebdee9102e41b5b7084fce81125b0/WordFileReader/WordDocument.cs#L94
        // $this->readRecordPlcfBteChpx();

        // $this->generatePhpWord();

        // reading paragraphs
        $this->getRecordPlcfBtePapx();
    }

    /**
     * Section and information about them
     * @link : http://msdn.microsoft.com/en-us/library/dd924458%28v=office.12%29.aspx
     */
    private function readRecordPlcfSed()
    {
        $posMem = $this->arrayFib['fcPlcfSed'];
        // PlcfSed
        // PlcfSed : aCP
        $aCP = array();
        $aCP[0] = self::getInt4d($this->data1Table, $posMem);
        $posMem += 4;
        $aCP[1] = self::getInt4d($this->data1Table, $posMem);
        $posMem += 4;

        // PlcfSed : aSed
        //@link : http://msdn.microsoft.com/en-us/library/dd950194%28v=office.12%29.aspx
        $numSed = $this->getNumInLcb($this->arrayFib['lcbPlcfSed'], 12);

        $aSed = array();
        for ($iInc = 0; $iInc < $numSed; ++$iInc) {
            // Sed : http://msdn.microsoft.com/en-us/library/dd950982%28v=office.12%29.aspx
            // fn
            $posMem += 2;
            // fnMpr
            $aSed[$iInc] = self::getInt4d($this->data1Table, $posMem);
            $posMem += 4;
            // fnMpr
            $posMem += 2;
            // fcMpr
            $posMem += 4;
        }

        foreach ($aSed as $offsetSed) {
            // Sepx : http://msdn.microsoft.com/en-us/library/dd921348%28v=office.12%29.aspx
            if (!isset($this->dataWorkDocument[$offsetSed])) {
                continue;
            }

            $cb = self::getInt2d($this->dataWorkDocument, $offsetSed);
            $offsetSed += 2;

            $oStylePrl = $this->readPrl($this->dataWorkDocument, $offsetSed, $cb);
            $offsetSed += $oStylePrl->length;

            $this->arraySections[] = $oStylePrl;
        }
    }

    /**
     * Specifies the fonts that are used in the document
     * @link : http://msdn.microsoft.com/en-us/library/dd943880%28v=office.12%29.aspx
     */
    private function readRecordSttbfFfn()
    {
        $posMem = $this->arrayFib['fcSttbfFfn'];

        $cData = self::getInt2d($this->data1Table, $posMem);
        $posMem += 2;
        $cbExtra = self::getInt2d($this->data1Table, $posMem);
        $posMem += 2;

        if ($cData < 0x7FF0 && $cbExtra == 0) {
            for ($inc = 0; $inc < $cData; $inc++) {
                // len
                $posMem += 1;
                // ffid
                $posMem += 1;
                // wWeight (400 : Normal - 700 bold)
                $posMem += 2;
                // chs
                $posMem += 1;
                // ixchSzAlt
                $ixchSzAlt = self::getInt1d($this->data1Table, $posMem);
                $posMem += 1;
                // panose
                $posMem += 10;
                // fs
                $posMem += 24;
                // xszFfn
                $xszFfn = '';
                do {
                    $char = self::getInt2d($this->data1Table, $posMem);
                    $posMem += 2;
                    if ($char > 0) {
                        $xszFfn .= chr($char);
                    }
                } while ($char != 0);
                // xszAlt
                $xszAlt = '';
                if ($ixchSzAlt > 0) {
                    do {
                        $char = self::getInt2d($this->data1Table, $posMem);
                        $posMem += 2;
                        if ($char == 0) {
                            break;
                        }
                        $xszAlt .= chr($char);
                    } while ($char != 0);
                }
                $this->arrayFonts[] = array(
                    'main' => $xszFfn,
                    'alt' => $xszAlt,
                );
            }
        }
    }

    /**
     * Read PlcfBtePapx
     *
     * Paragraph and information about them
     *
     * https://msdn.microsoft.com/en-us/library/dd908569(v=office.12).aspx
     */
    private function getRecordPlcfBtePapx()
    {
        $arrayParagraphs = array();

        $pos = $this->arrayFib['fcPlcfBtePapx'];
        $isize = $this->arrayFib['lcbPlcfBtePapx'];

        $num = $this->getNumInLcb($isize, 4);

        $pos += 4 * ($num + 1); // skip aFCs

        $aPnBtePapx = array();

        for($i=0; $i < $num; $i++) {
            $aPnBtePapx[$i] = self::getInt4d($this->data1Table, $pos) & 0x3FFFFF; // 22bits, 10bit unused
            $pos += 4;
        }

        for($i = 0; $i < $num; $i++) {
            $offsetBase = $aPnBtePapx[$i] * 512; // offset: pn*512
            $offset = $offsetBase;

            $cpara = self::getInt1d($this->dataWorkDocument, $offset + 511);

            $arrayFCs = array();
            for ($j = 0; $j <= $cpara; $j++) { // the number of rgfcs is cpara + 1
                $arrayFCs[$j] = self::getInt4d($this->dataWorkDocument, $offset);
                $offset += 4;
            }

            $arrayRGBs = array();
            for ($j = 1; $j <= $cpara; $j++) {
                $arrayRGBs[$j] = self::getInt1d($this->dataWorkDocument, $offset);
                $offset += 1;
                $offset += 12; // reserved skipped (12 bytes)
            }

            $styles = array();
            for ($j = 1; $j <= $cpara; $j++) {
                $rgb = $arrayRGBs[$j];
                $offset = $offsetBase + ($rgb * 2);

                $cb = self::getInt1d($this->dataWorkDocument, $offset);
                $offset += 1;
                if ($cb == 0) {
                    $cb = self::getInt1d($this->dataWorkDocument, $offset);
                    $cb = $cb * 2;
                    $offset += 1;
                } else {
                    $cb = $cb * 2 - 1;
                }
                $istd = self::getInt2d($this->dataWorkDocument, $offset);
                $offset += 2;
                $cb -= 2;

                if ($cb > 0) {
                    $styles[$j] = $this->readPrl($this->dataWorkDocument, $offset, $cb);
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

        $this->arrayParagraphs = $arrayParagraphs;
    }

    /**
     * Character formatting properties to text in a document
     * @link http://msdn.microsoft.com/en-us/library/dd907108%28v=office.12%29.aspx
     */
    private function readRecordPlcfBteChpx()
    {
        $posMem = $this->arrayFib['fcPlcfBteChpx'];
        $num = $this->getNumInLcb($this->arrayFib['lcbPlcfBteChpx'], 4);
        $aPnBteChpx = array();
        /*for ($inc = 0; $inc <= $num; $inc++) {
            $aPnBteChpx[$inc] = self::getInt4d($this->data1Table, $posMem);
            $posMem += 4;
        }*/

        $posMem += ($num + 1) * 4;

        for ($i = 1; $i <= $num; $i++) {
            $aPnBteChpx[$i] = self::getInt4d($this->data1Table, $posMem) & 0x3FFFFF; // 22bits, 10bit unused
            $posMem += 4;
        }

        for ($i = 1; $i <= $num; $i++) {
            $pnFkpChpx = $aPnBteChpx[$i];

            $offsetBase = $pnFkpChpx * 512;
            $offset = $offsetBase;

            // ChpxFkp
            // @link : http://msdn.microsoft.com/en-us/library/dd910989%28v=office.12%29.aspx
            $numRGFC = self::getInt1d($this->dataWorkDocument, $offset + 511);
            $arrayRGFC = array();
            for ($inc = 0; $inc <= $numRGFC; $inc++) {
                $arrayRGFC[$inc] = self::getInt4d($this->dataWorkDocument, $offset);
                $offset += 4;
            }

            $arrayRGB = array();
            for ($inc = 1; $inc <= $numRGFC; $inc++) {
                $arrayRGB[$inc] = self::getInt1d($this->dataWorkDocument, $offset);
                $offset += 1;
            }

            $start = 0;
            foreach ($arrayRGB as $keyRGB => $rgb) {
                $oStyle = new \stdClass();
                $oStyle->pos_start = $start;
                $oStyle->pos_ori_start = $arrayRGFC[$keyRGB - 1];
                //$oStyle->pos_len = (int)ceil((($arrayRGFC[$keyRGB] -1) - $arrayRGFC[$keyRGB -1]) / 2);
                $oStyle->pos_len = $arrayRGFC[$keyRGB] - $arrayRGFC[$keyRGB -1];
                $start += $oStyle->pos_len;

                if ($rgb > 0) {
                    // Chp Structure
                    // @link : http://msdn.microsoft.com/en-us/library/dd772849%28v=office.12%29.aspx
                    $posRGB = $offsetBase + $rgb * 2;

                    $cb = self::getInt1d($this->dataWorkDocument, $posRGB);
                    $posRGB += 1;

                    $oStyle->style = $this->readPrl($this->dataWorkDocument, $posRGB, $cb);
                    $posRGB += $oStyle->style->length;
                }
                $this->arrayCharacters[] = $oStyle;
            }
        }
    }

    /**
     * @param $sprm
     * @return \stdClass
     */
    private function readSprm($sprm)
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
    private function readSprmSpra($data, $pos, $oSprm)
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
    private function readPrl($data, $pos, $cbNum)
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
            $oSprm = $this->readSprm($sprm);
            $pos += 2;
            $cbNum -= 2;

            $arrayReturn = $this->readSprmSpra($data, $pos, $oSprm);
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
                            // print_r('sprmPIlvl : '.$operand.PHP_EOL.PHP_EOL);
                            $oStylePrl->iLvl = $operand;
                            break;
                        case 0x0B: // sprmPIlfo, list format
                            // print_r('sprm_IsPmd : ' . dechex($operand) . PHP_EOL.PHP_EOL);
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
                            // print_r('sprm_IsPmd : ' . dechex($oSprm->isPmd) .PHP_EOL.PHP_EOL);
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
                            if (isset($this->arrayFonts[$operand])) {
                                $oStylePrl->styleFont['name'] = $this->arrayFonts[$operand]['main'];
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
                            $red = str_pad(dechex(self::getInt1d($this->dataWorkDocument, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $green = str_pad(dechex(self::getInt1d($this->dataWorkDocument, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $blue = str_pad(dechex(self::getInt1d($this->dataWorkDocument, $pos)), 2, '0', STR_PAD_LEFT);
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
                    if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                        echo "\nPicture property\n";
                    }
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
                $mfpfMm = self::getInt2d($this->dataData, $sprmCPicLocation);
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
                $picmidDxaGoal = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dyaGoal
                $picmidDyaGoal = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : mx
                $picmidMx = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : my
                $picmidMy = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dxaReserved1
                $picmidDxaCropLeft = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dyaReserved1
                $picmidDxaCropTop = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dxaReserved2
                $picmidDxaCropRight = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                // PICF : picmid : dyaReserved2
                $picmidDxaCropBottom = self::getInt2d($this->dataData, $sprmCPicLocation);
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
                    $cchPicName = self::getInt1d($this->dataData, $sprmCPicLocation);
                    $sprmCPicLocation += 1;

                    // stPicName
                    $stPicName = '';
                    for ($inc = 0; $inc <= $cchPicName; $inc++) {
                        $chr = self::getInt1d($this->dataData, $sprmCPicLocation);
                        $sprmCPicLocation += 1;
                        $stPicName .= chr($chr);
                    }
                }

                // picture (OfficeArtInlineSpContainer)
                // picture : shape
                $shapeRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 8;
                if ($shapeRH['recVer'] == 0xF && $shapeRH['recInstance'] == 0x000 && $shapeRH['recType'] == 0xF004) {
                    $sprmCPicLocation += $shapeRH['recLen'];
                }
                // picture : rgfb
                //@link : http://msdn.microsoft.com/en-us/library/dd950560%28v=office.12%29.aspx
                $fileBlockRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
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
                            $cbName = self::getInt1d($this->dataData, $sprmCPicLocation);
                            $sprmCPicLocation += 1;
                            // unused2
                            $sprmCPicLocation += 1;
                            // unused3
                            $sprmCPicLocation += 1;
                            // nameData
                            if ($cbName > 0) {
                                $nameData = '';
                                for ($inc = 0; $inc <= ($cbName / 2); $inc++) {
                                    $chr = self::getInt2d($this->dataData, $sprmCPicLocation);
                                    $sprmCPicLocation += 2;
                                    $nameData .= chr($chr);
                                }
                            }
                            // embeddedBlip
                            //@link : http://msdn.microsoft.com/en-us/library/dd910081%28v=office.12%29.aspx
                            $embeddedBlipRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
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
                                    $oStylePrl->image['data'] = substr($this->dataData, $sprmCPicLocation, $embeddedBlipRH['recLen']);
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
                    $fileBlockRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
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
    private function loadRecordHeader($stream, $pos)
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

    private function generateTxt()
    {
        $clx = $this->getDocumentText($this->arrayFib, $this->data1Table, $this->dataData);

        $textBlockList = $clx->textBlockList;

        $num = count($textBlockList);

        $mainStream = '';
        for ($i = 0; $i < $num; $i++) {
            $tTextBlock = $textBlockList[$i];

            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo PHP_EOL . "start CP: " . $tTextBlock->startCP . ", end CP: " . $tTextBlock->endCP . PHP_EOL;

                echo PHP_EOL . "start FC: " . $tTextBlock->ulCharPos . ", end FC: " . ($tTextBlock->ulCharPos + $tTextBlock->lToGo) . PHP_EOL;
            }

            $chars = substr($this->dataWorkDocument, $tTextBlock->ulCharPos, $tTextBlock->lToGo);
            if ($tTextBlock->bUsesUnicode) {
                $subText = iconv('UCS-2LE', 'UTF-8', $chars);
            }

            if ($subText === false) {
                continue;
            }

            $subText = str_replace(chr(13), "\n", $subText);
            $mainStream .= $subText;
        }

        $comment_chars = mb_substr($mainStream, $this->arrayFib['ccpText'] + $this->arrayFib['ccpFtn'] + $this->arrayFib['ccpHdd'], $this->arrayFib['ccpAtn'], 'UTF-8');

        $mainStream = mb_substr($mainStream, 0, $this->arrayFib['ccpText'], 'UTF-8');

        $comments = $this->readComments();
        $totComments = count($comments);

        $totParas = count($this->arrayParagraphs);
        $pad = 0;
        $startCPs = array();
        for ($i =0; $i < $totParas; $i++) {
            $para = $this->arrayParagraphs[$i];
            for ($j = 0; $j < count($para->aFCs) - 1; $j++) {
                $fc = $para->aFCs[$j];
                $efc = $para->aFCs[$j + 1];
                for ($k = 0; $k < $num; $k++) {
                    $tTextBlock = $textBlockList[$k];
                    if ($fc >= $tTextBlock->ulCharPos && $fc < $tTextBlock->ulCharPos + $tTextBlock->lToGo) {
                        $dfc = $fc - $tTextBlock->ulCharPos;
                        if ($tTextBlock->bUsesUnicode) {
                            $dfc = $dfc >> 1;
                        }
                        $start_cp = $dfc + $tTextBlock->startCP;
                    }

                    if ($efc <= $tTextBlock->ulCharPos + $tTextBlock->lToGo) {
                        $dfc = $efc - $tTextBlock->ulCharPos;
                        if ($tTextBlock->bUsesUnicode) {
                            $dfc = $dfc >> 1;
                        }
                        $end_cp = $dfc + $tTextBlock->startCP;
                    }

                    if (isset($start_cp) && isset($end_cp) && $end_cp > $start_cp) {
                        $paraStyle = $para->aStyles[$j + 1];
                        if (!empty($paraStyle) && isset($paraStyle->bList) && $paraStyle->bList) {
                            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                                echo PHP_EOL . "paragraph start cp: " . $start_cp . ", end cp: " . $end_cp . PHP_EOL;
                            }
                            if (in_array($start_cp, $startCPs)) {
                                continue;
                            }
                            $startCPs[] = $start_cp;
                            $p = mb_substr($mainStream, 0, $start_cp + $pad, "UTF-8");
                            $e = mb_substr($mainStream, $start_cp + $pad, null, "UTF-8");
                            $mainStream = $p . "*\t" . $e;
                            $pad += 2;
                        }
                    }
                }
            }
        }

        $mainStreams = explode(chr(0x05), $mainStream);

        $comment_strs = array();
        for ($i = 0; $i < $totComments; $i++) {
            $comment = $comments[$i];
            $comment['chars'] = mb_substr($comment_chars, $comment['start_cp'], $comment['length'], 'UTF-8');
            $comment['chars'] = str_replace(chr(0x05), '', $comment['chars']);
            $comment['chars'] = trim($comment['chars']);
            if (empty($comment['chars'])) {
                continue;
            }
            $comment_strs[] = $this->outputCommentTxt($comment, $i + 1);
        }

        $content = '';
        foreach ($mainStreams as $key => $item) {
            $content .= $item;
            if (isset($comment_strs[$key]) && !empty($comment_strs[$key])) {
                $content .= "(" . $comment_strs[$key] . ")";
            }
        }

        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo $content;
            echo "content length: " . mb_strlen($content, "UTF-8") . PHP_EOL;
        }

        return $content;
    }

    private function outputCommentTxt(&$comment, $i)
    {
        if (empty($comment)) {
            return '';
        }

        $date_format = "%d/%d/%d %d:%d";
        $lastDay = sprintf($date_format, $comment['yr'], $comment['mon'], $comment['dom'], $comment['hr'], $comment['minute']);
        return sprintf("\033[7m批注[%d] %s:\e[0m \033[4m%s\033[0m", $i, $lastDay, $comment['chars']);
    }

    private function generatePhpWord()
    {
        foreach ($this->arraySections as $itmSection) {
            $oSection = $this->phpWord->addSection();
            $oSection->setSettings($itmSection->styleSection);

            $sHYPERLINK = '';
            foreach ($this->arrayParagraphs as $itmParagraph) {
                $textPara = $itmParagraph;
                $textParaLen = strlen($textPara);
                foreach ($this->arrayCharacters as $oCharacters) {
                    $tmp = substr($textPara, $oCharacters->pos_start, $oCharacters->pos_len);
                    $subText = iconv('UCS-2LE', 'UTF-8', $tmp);

                    //$subText = mb_substr($textPara, $oCharacters->pos_start, $oCharacters->pos_len, 'utf-8');
                    if ($subText === false) {
                        continue;
                    }
                    $subText = str_replace(chr(13), PHP_EOL, $subText);
                    $arrayText = explode(PHP_EOL, $subText);
                    if (end($arrayText) == '') {
                        array_pop($arrayText);
                    }
                    if (reset($arrayText) == '') {
                        array_shift($arrayText);
                    }

                    // Style Character
                    $styleFont = array();
                    if (isset($oCharacters->style)) {
                        if (isset($oCharacters->style->styleFont)) {
                            $styleFont = $oCharacters->style->styleFont;
                        }
                    }

                    foreach ($arrayText as $sText) {
                        // HyperLink
                        if (empty($sText) && !empty($sHYPERLINK)) {
                            $arrHYPERLINK = explode('"', $sHYPERLINK);
                            $oSection->addLink($arrHYPERLINK[1], null);
                            // print_r('>addHyperLink<'.$sHYPERLINK.'>'.ord($sHYPERLINK[0]).EOL);
                            $sHYPERLINK = '';
                        }

                        // TextBreak
                        if (empty($sText)) {
                            $oSection->addTextBreak();
                            $sHYPERLINK = '';
                            // print_r('>addTextBreak<' . EOL);
                        }

                        if (!empty($sText)) {
                            if (!empty($sHYPERLINK) && ord($sText[0]) > 20) {
                                $sHYPERLINK .= $sText;
                            }
                            if (empty($sHYPERLINK)) {
                                if (ord($sText[0]) > 20) {
                                    if (strpos(trim($sText), 'HYPERLINK "') === 0) {
                                        $sHYPERLINK = $sText;
                                    } else {
                                        $oSection->addText($sText, $styleFont);
                                        // print_r('>addText<'.$sText.'>'.ord($sText[0]).EOL);
                                    }
                                }
                                if (ord($sText[0]) == 1) {
                                    if (isset($oCharacters->style->image)) {
                                        $fileImage = tempnam(sys_get_temp_dir(), 'PHPWord_MsDoc').'.'.$oCharacters->style->image['format'];
                                        file_put_contents($fileImage, $oCharacters->style->image['data']);
                                        $oSection->addImage($fileImage, array('width' => $oCharacters->style->image['width'], 'height' => $oCharacters->style->image['height']));
                                        // print_r('>addImage<'.$fileImage.'>'.EOL);
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
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

    public static function pSectionSZ($x)
    {
        return (($x) * P_SECTIONLIST_SZ + P_LENGTH_SZ);
    }

    public static function analyseSystemInformationHeader($data)
    {
        $usLittleEndian = self::getInt2d($data, 0);

        if ($usLittleEndian !== 0xFFFE) {
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "bigendian\n";
            }
            return NULL;
        }

        $usEmpty =  self::getInt2d($data, 2);
        if ($usEmpty != 0x0000) {
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "usEmpty false\n";
            }
            return NULL;
        }

        $ulTmp = self::getInt4d($data, 4);
        $usOS = ($ulTmp >> 16);
	    $usVersion = ($ulTmp & 0xffff);

        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo "OS Version: ";
        }
        switch ($usOS) {
            case 0:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Win16";
                }
                break;
            case 1:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "MacOS";
                }
                break;
            case 2:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Win32";
                }
                break;
            default:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo $usOS;
                }
                break;
        }
        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo "usVersioin: " . $usVersion;
            echo "\n";
        }

        $tSectionCount = self::getInt4d($data, 24);

        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo "tSectionCount: {$tSectionCount}\n";
        }
	    if ($tSectionCount != 1 && $tSectionCount != 2) {
            return NULL;
        }

        $aucSecLst = substr($data, P_HEADER_SZ, self::pSectionSZ($tSectionCount));

        $ulTmp = self::getInt4d($aucSecLst, 0);
        $ulTmp = self::getInt4d($aucSecLst, 4);
        $ulTmp = self::getInt4d($aucSecLst, 8);
        $ulTmp = self::getInt4d($aucSecLst, 12);
        $ulOffset = self::getInt4d($aucSecLst, 16);

        if ($ulOffset != P_HEADER_SZ + P_SECTIONLIST_SZ &&
            $ulOffset != P_HEADER_SZ + 2 * P_SECTIONLIST_SZ) {
            return NULL;
        }

        $tLength =
            self::getInt4d($aucSecLst, $tSectionCount * P_SECTIONLIST_SZ);

        $aucBuffer = substr($data, $ulOffset, $tLength);
        return $aucBuffer;
    }

    private function getSystemInformation($data)
    {
        $aucBuffer = self::analyseSystemInformationHeader($data);

        $res = array();

        $tCount = self::getInt4d($aucBuffer, 4);

        for ($tIndex = 0; $tIndex < $tCount; $tIndex++) {
            $tPropID = self::getInt4d($aucBuffer, 8 + $tIndex * 8);
		    $ulOffset = self::getInt4d($aucBuffer, 12 + $tIndex * 8);
            $tPropType = self::getInt4d($aucBuffer, $ulOffset);

            switch ($tPropID) {
                case PID_CODEPAGE:
                    if ($tPropType === VT_I2) {
                        $res['codepage'] = self::getInt2d($aucBuffer, $ulOffset + 4);
                    }
                    break;
                case PID_TITLE:
                    if ($tPropType === VT_LPSTR) {
                        if (isset($res['title'])) {
                            $res['title'] .= self::szLpstr($ulOffset, $aucBuffer);
                        } else {
                            $res['title'] = self::szLpstr($ulOffset, $aucBuffer);
                        }

                        if (isset($res['codepage']) && !empty($res['codepage']) && isset(self::$CodePage[$res['codepage']])) {
                            $res['title'] = iconv(self::$CodePage[$res['codepage']], 'UTF-8', $res['title']);
                        }
                    }
                    break;
                case PID_SUBJECT:
                    if ($tPropType == VT_LPSTR) {
                        $res['subject'] = self::szLpstr($ulOffset, $aucBuffer);
                        if (isset($res['codepage']) && !empty($res['codepage']) && isset(self::$CodePage[$res['codepage']])) {
                            $res['subject'] = iconv(self::$CodePage[$res['codepage']], 'UTF-8', $res['subject']);
                        }
                    }
                    break;
                case PID_AUTHOR:
                    if ($tPropType == VT_LPSTR) {
                        $res['author'] = self::szLpstr($ulOffset, $aucBuffer);
                        if (isset($res['codepage']) && !empty($res['codepage']) && isset(self::$CodePage[$res['codepage']])) {
                            $res['author'] = iconv(self::$CodePage[$res['codepage']], 'UTF-8', $res['author']);
                        }
                    }
                    break;
                case PID_APPNAME:
                    if ($tPropType == VT_LPSTR) {
                        $res['appname'] = self::szLpstr($ulOffset, $aucBuffer);
                    }
                    break;
                case PID_CREATE_DTM:
                    if ($tPropType == VT_FILETIME) {
                        $res['created'] = $this->getFileTime($aucBuffer, $ulOffset);
                    }
                    break;
                case PID_LASTSAVE_DTM:
                    if ($tPropType == VT_FILETIME) {
                        $res['lastModified'] = $this->getFileTime($aucBuffer, $ulOffset);
                    }
                    break;
                case PIDSI_PAGECOUNT:
                    if ($tPropType == VT_I4) {
                        $res['pagecount'] = self::getInt4d($aucBuffer, $ulOffset + 4);
                    }
                    break;
                case PIDSI_WORDCOUNT:
                    if ($tPropType == VT_I4) {
                        $res['wordcount'] = self::getInt4d($aucBuffer, $ulOffset + 4);
                    }
                    break;
                case PIDSI_CHARCOUNT:
                    if ($tPropType == VT_I4) {
                        $res['charcount'] = self::getInt4d($aucBuffer, $ulOffset + 4);
                    }
                    break;
            }
        }

        return $res;
    }

    /**
     * getDocumentSummaryInfo - analyse the document summary information
     * @param $data
     * @return array
     */
    private function getDocumentSummaryInfo($data)
    {
        $aucBuffer = self::analyseSystemInformationHeader($data);
        $tCount = self::getInt4d($aucBuffer, 4);

        $res = array();
        for($tIndex = 0; $tIndex < $tCount; $tIndex++)
        {
            $tPropID = self::getInt4d($aucBuffer, 8 + $tIndex * 8);
            $ulOffset = self::getInt4d($aucBuffer, 12 + $tIndex * 8);
            $tPropType = self::getInt4d($aucBuffer, $ulOffset);
            switch ($tPropID) {
                case PIDD_MANAGER:
                    if ($tPropType == VT_LPSTR) {
                        $res['manager'] = self::szLpstr($ulOffset, $aucBuffer);
                    }
                    break;
                case PIDD_COMPANY:
                    if ($tPropType == VT_LPSTR) {
                        $res['company'] = self::szLpstr($ulOffset, $aucBuffer);
                    }
                    break;
            }
        }
        return $res;
    }

    public static function isspace($char)
    {
        if ($char == '\t' || $char == ' ' || $char == '\r' || $char == '\n' || $char == '\v' || $char == '\f') {
            return true;
        }

        return false;
    }

    public static function szLpstr($ulOffset, $aucBuffer)
    {
        $tSize = self::getInt4d($aucBuffer, $ulOffset + 4);

        if ($tSize === 0) {
            return NULL;
        }

        $szStart = $ulOffset + 8;

        while(self::isspace($aucBuffer[$szStart])) {
            $szStart++;
        }

        if (!isset($aucBuffer[$szStart]) || empty($aucBuffer[$szStart])) {
            return NULL;
        }

        $s = array();
        for($i=$szStart; $i < $szStart + $tSize; $i++) {
            $s[] = $aucBuffer[$i];
        }
        $c = count($s);
        for($i=$c;$i>0;$i--) {
            if (self::isspace($s[$i -1]) || ord($s[$i - 1]) === 0) {
                unset($s[$i]);
            } else {
                break;
            }
        }
        return implode("", $s);
    }

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
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "No proper data given";
            }
            return FALSE;
        }

        if ($iTailLen < self::BIG_BLOCK_SIZE * 3) {
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "This file is too small to be a Word document";
            }
            return false;
        }

        $iTailLen = intval(($fileSize % self::BIG_BLOCK_SIZE));
        switch ($iTailLen) {
            case 0:		/* No tail, as it should be */
                break;
            case 1:
            case 2:		/* Filesize mismatch or a buggy email program */
                if (intval(($fileSize % 3)) == $iTailLen) {
                    if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                        echo 'Filesize mismatch or a buggy email program';
                    }
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
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "No proper data given";
            }
            return FALSE;
        }
        if ($fileSize < 128) {
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "File too small to be a Word document";
            }
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
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "This file is too small to be a Word document";
            }
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
            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                echo "No proper data given";
            }
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
     * GetVersionNumber - get the Word version number from the header
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
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Word for DOS";
                }
                return 0;
            case  28:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Word 4 for Macintosh";
                }
                $bOldMacFile = TRUE;
                return 4;
            case  33:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Word 1.x for Windows";
                }
                return 1;
            case  35:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Word 5 for Macintosh";
                }
                $bOldMacFile = TRUE;
                return 5;
            case  45:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Word 2 for Windows";
                }
                return 2;
            case 101:
            case 102:
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Word 6 for Windows";
                }
                return 6;
            case 103:
            case 104:
                $usChse = self::getInt2d($aucHeader, 0x14);

                switch ($usChse) {
                    case 0:
                        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                            echo "Word 7 for Win95";
                        }
                        return 7;
                    case 256:
                        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                            echo "Word 6 for Macintosh";
                        }
                        $bOldMacFile = TRUE;
                        return 6;
                    default:
                        if (self::getInt1d($aucHeader, 0x05) == 0xe0) {
                            if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                                echo "Word 7 for Win95";
                            }
                            return 7;
                        }
                        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                            echo "Word 6 for Macintosh";
                        }
                        $bOldMacFile = TRUE;
                        return 6;
                }
            default:
                $usChse = self::getInt2d($aucHeader, 0x14);

                if ($nFib < 192) {
                    /* Unknown or unsupported version of Word */
                    if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                        echo "Unknown or unsupported version of Word";
                    }
                    return -1;
                }

                if ($usChse != 256) {
                    if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                        echo "Word97 for Win95/98/NT";
                    }
                }

                if (usChse == 256) {
                    if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                        echo "Word98 for Macintosh";
                    }
                }

                return 8;
        }
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

    private function buildLfoList()
    {
        /* LFO (List Format Override) */

        $lfoList = array();

        $ulBeginLfoInfo = $this->arrayFib['fcPlfLfo'];
        $tLfoInfoLen = $this->arrayFib['lcbPlfLfo'];

        $aucLfoInfo = substr($this->data1Table, $ulBeginLfoInfo, $tLfoInfoLen);

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
     * Read Comment Document
     *
     * @return array
     */
    private function readComments()
    {
        $comments = array();
        $offset = $this->arrayFib['fcPlcfandTxt'];
        $lcbPlcfandTxt = $this->arrayFib['lcbPlcfandTxt'];

        if ($lcbPlcfandTxt <= 0) {
            return false;
        }

        $numCPs = (int) ($lcbPlcfandTxt / 4);

        $aCPs = array();
        for ($i = 0; $i < $numCPs; $i++) {
            $aCPs[$i] = self::getInt4d($this->data1Table, $offset);
            $offset += 4;
        }

        if ($aCPs[$numCPs - 2] !== $this->arrayFib['ccpAtn'] - 1) {
            echo "invalid second-to-last cp : " . $aCPs[$numCPs - 2] . ", ccpAtn: " . $this->arrayFib['ccpAtn'] . PHP_EOL;
        }

        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo "comment cps: " . json_encode($aCPs) . PHP_EOL;
        }

        // Read XSTs
        $fcGrpXstAtnOwners = $this->arrayFib['fcGrpXstAtnOwners'];
        $lcbGrpXstAtnOwners = $this->arrayFib['lcbGrpXstAtnOwners']; // An unsigned integer that specifies the size, in bytes, of the XST array

        $offset = $fcGrpXstAtnOwners;

        $author_names = array();

        while ($lcbGrpXstAtnOwners > 0) {
            $cch = self::getInt2d($this->data1Table, $offset);
            $offset += 2;
            $author_name = substr($this->data1Table, $offset, $cch * 2);
            $author_name = iconv('UCS-2LE', 'UTF-8', $author_name);
            $author_names[] = $author_name;

            $offset += $cch * 2;

            $lcbGrpXstAtnOwners -= 2 + 2 * $cch;
        }

        $fcPlcfandRef = $this->arrayFib['fcPlcfandRef'];
        $lcbPlcfandRef = $this->arrayFib['lcbPlcfandRef'];

        if ($lcbPlcfandRef <= 0) {
            return false;
        }

        $numRefCPs = $this->GetNumInLcb($lcbPlcfandRef, 30);

        $refCPs = array();
        $offset = $fcPlcfandRef;
        for ($i = 0; $i <= $numRefCPs; $i++) {
            $refCPs[$i] = self::getInt4d($this->data1Table, $offset);
            $offset += 4;
        }

        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo "ref cps: " . json_encode($refCPs) . PHP_EOL;
        }
        $aATRDPre10 = array();
        for ($i = 1; $i <= $numRefCPs; $i++) {
            $cch = self::getInt2d($this->data1Table, $offset);
            if ($cch > 9) {
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "bad file" . PHP_EOL;
                }
                return false;
            }
            $initial_author = substr($this->data1Table, $offset + 2, $cch * 2);
            $initial_author = iconv('UCS-2LE', 'UTF-8', $initial_author);

            $xstIndex = self::getInt2d($this->data1Table, $offset + 20);
            $lTagBkmk = self::getInt4d($this->data1Table, $offset + 26);

            $offset += 30;

            $aATRDPre10[$i] = array(
                'initial_author'    => $initial_author,
                'bkmk_id'           => $lTagBkmk,
                'xst_index'         => $xstIndex,
                'author_name'       => $author_names[$xstIndex],
            );
        }

        // AtrdExtra
        $fcAtrdExtra = $this->arrayFib['fcAtrdExtra'];
        $lcbAtrdExtra = $this->arrayFib['lcbAtrdExtra'];

        $aAtrdExtra = array();
        if ($lcbAtrdExtra > 0) {
            $numAtrdExtra = $lcbAtrdExtra / 18;
            if ($numAtrdExtra !== $numRefCPs) {
                return false;
            }

            $offset = $fcAtrdExtra;

            for ($i = 1; $i <= $numAtrdExtra; $i++) {
                $dttm = self::getInt4d($this->data1Table, $offset);
                $minute = $dttm & 0x3F;
                $hr = ($dttm >> 6) & 0x1F;
                $dom = ($dttm >> 11) & 0x1F;
                $mon = ($dttm >> 16) & 0xF;
                $yr = 1900 + (($dttm >> 20) & 0x1FF);
                $wdy = ($dttm >> 29) & 0x7;
                $offset += 4;
                // padding1
                $offset += 2;
                $cDepth = self::getInt4d($this->data1Table, $offset);
                $offset += 4;
                $diatrdParent = self::getInt4d($this->data1Table, $offset);
                $offset += 4;
                $offset += 4;

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

        if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
            echo "comments: " . json_encode($comments) . PHP_EOL;
        }
        return $comments;
    }

    /**
     * Retrieving Text
     *
     * @return stdClass
     */
    private function getDocumentText()
    {
        $clx = new \stdClass();

        $clx->styles = array();

        $ulBeginTextInfo = $this->arrayFib['fcClx'];
        $tTextInfoLen = $this->arrayFib['lcbClx'];

        $aucBuffer = substr($this->data1Table, $ulBeginTextInfo, $tTextInfoLen);

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
                $oStyle = $this->readPrl($aucBuffer, $lOff, $cbGrpprl);
                $clx->styles[] = $oStyle;
                $lOff += $oStyle->length;
                continue;
            }

            if ($iType !== 0x02) { // Not A Pcdt
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "Not A valid Pcdt\n";
                }
                return false;
            }

            // handle A Pcdt
            $lcb = self::getInt4d($aucBuffer, $lOff);
            $lOff += 4;

            if ($lcb < 4) {
                if (defined('ECHO_DEBUG_ENABLE') && ECHO_DEBUG_ENABLE) {
                    echo "invalid pcdt\n";
                }
                return false;
            }

            $lPieces = $this->getNumInLcb($lcb, 8);

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
     * get a filetime property in seconds
     *
     * @param $aucBuffer
     * @param $ulOffset
     * @return int
     */
    private function getFileTime($aucBuffer, $ulOffset)
    {
        $ulLo = self::getInt4d($aucBuffer, $ulOffset + 4);
        $ulHi = self::getInt4d($aucBuffer, $ulOffset + 8);

        $dHi = $ulHi - TIME_OFFSET_HI;
        $dLo = $ulLo - TIME_OFFSET_LO;

        $dTmp = $dLo / 10000000.0; /* 10^7 */
        $dTmp += $dHi * 429.4967926; /* 2^32 / 10^7 */

        $tResult = $dTmp < 0.0 ? (int)($dTmp - 0.5) : (int)($dTmp + 0.5);

        return $tResult;
    }
}
