<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 17/1/22
 * Time: PM2:26
 */

namespace PhpOffice\PhpWord\Reader;

use PhpOffice\PhpWord\PhpWord;

class PDFReader extends AbstractReader implements ReaderInterface
{
    /**
     * Loads PhpWord from file
     *
     * @param string $filename
     * @return PhpWord
     * @throws \Exception
     */
    public function load($filename)
    {
        $phpWord = new PhpWord();

        require_once(dirname(__FILE__) . '/../../../vendor/smalot/pdfparser/vendor/autoload.php');
        // Parse pdf file and build necessary objects.
        $parser = new \Smalot\PdfParser\Parser();

        $pdf    = $parser->parseFile($filename);
        $text = $pdf->getText();

        $textLen = mb_strlen($text, "UTF-8");
        $phpWord->getDocInfo()->setMainStreamSize($textLen);

        // Retrieve all details from the pdf file.
        $details  = $pdf->getDetails();

        // Loop over each property to extract values (string or array).
        foreach ($details as $property => $value) {
            if (is_array($value)) {
                $value = implode(', ', $value);
            }
            echo $property . ' => ' . $value . PHP_EOL;
            $property = strtolower($property);
            if (strcasecmp($property, "title")) {
                $phpWord->getDocInfo()->setTitle($value);
            } else if (strcasecmp($property, "creator")) {
                $phpWord->getDocInfo()->setCreator($value);
            } else if (strcasecmp($property, "creationdate")) {
                $phpWord->getDocInfo()->setCreated(strtotime($value));
            } else if (strcasecmp($property, "moddate")) {
                $phpWord->getDocInfo()->setModified(strtotime($value));
            } else if (strcasecmp($property, "Keywords")) {
                $phpWord->getDocInfo()->setKeywords($value);
            } else if (strcasecmp($property, "pages")) {
                $pages = intval($value);
            }
        }

        return $phpWord;
    }
}