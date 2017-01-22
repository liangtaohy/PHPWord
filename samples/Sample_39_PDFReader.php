<?php
include_once 'Sample_Header.php';

$name = basename(__FILE__, '.php');

$source = __DIR__ . "/resources/{$name}.pdf";

echo date('H:i:s'), " Reading contents from `{$source}`", EOL;
$time_start = microtime(true);
$phpWord = \PhpOffice\PhpWord\IOFactory::load($source, 'PDFReader');
$time_end = microtime(true);
echo "time used in ms: " . (($time_end - $time_start) * 1000) . PHP_EOL;
echo $phpWord->getDocInfo()->getTitle() . PHP_EOL;
echo "Main Document Characters & Comment Characters: " . ($phpWord->getDocInfo()->getMainStreamSize() + $phpWord->getDocInfo()->getCommentSize()). PHP_EOL;
// (Re)write contents
$writers = array('Word2007' => 'docx', 'ODText' => 'odt', 'RTF' => 'rtf', 'HTML'=> 'html');
foreach ($writers as $writer => $extension) {
    echo date('H:i:s'), " Write to {$writer} format", EOL;
    $xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, $writer);
    $xmlWriter->save("{$name}.{$extension}");
    rename("{$name}.{$extension}", "results/{$name}.{$extension}");
}

include_once 'Sample_Footer.php';
