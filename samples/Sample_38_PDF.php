<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 17/1/17
 * Time: PM4:57
 */

include_once 'Sample_Header.php';
define('PHPWORD_TESTS_BASE_DIR', realpath(__DIR__));
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\Writer\PDF;
// Read contents
$name = basename(__FILE__, '.php');

$source = __DIR__ . "/resources/{$name}.docx";
$source = "/Users/xlegal/Desktop/股权代持协议(有利于代持人).docx";
echo date('H:i:s'), " Reading contents from `{$source}`", EOL;
$phpWord = \PhpOffice\PhpWord\IOFactory::load($source);

define('DOMPDF_ENABLE_AUTOLOAD', false);
$file = __DIR__ . '/results/temp.pdf';

$rendererName = Settings::PDF_RENDERER_DOMPDF;
$rendererLibraryPath = realpath(PHPWORD_TESTS_BASE_DIR . '/../vendor/dompdf/dompdf');
Settings::setPdfRenderer($rendererName, $rendererLibraryPath);
$writer = new PDF($phpWord);
$writer->save($file);

include_once 'Sample_Footer.php';