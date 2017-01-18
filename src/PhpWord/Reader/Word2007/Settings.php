<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 16/11/11
 * Time: PM7:48
 */

namespace PhpOffice\PhpWord\Reader\Word2007;

use PhpOffice\Common\XMLReader;
use PhpOffice\PhpWord\PhpWord;

class Settings extends AbstractPart
{
    public static $Settings = array();
    /**
     * Read styles.xml.
     *
     * @param \PhpOffice\PhpWord\PhpWord $phpWord
     * @return void
     */
    public function read(PhpWord $phpWord)
    {
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);

        $nodes = $xmlReader->getElements('*');
        if ($nodes->length > 0) {
            foreach ($nodes as $node) {
                switch($node->nodeName) {
                    case 'w:defaultTabStop':
                        $val = intval($xmlReader->getAttribute('w:val', $node));
                        self::$Settings['defaultTabStop'] = $val;
                        break;
                }
            }
        }
    }
}