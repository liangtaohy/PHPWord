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

namespace PhpOffice\PhpWord\Writer\HTML\Element;

use PhpOffice\PhpWord\Element\AbstractContainer as ContainerElement;

use PhpOffice\PhpWord\Style;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\Paragraph;
use PhpOffice\PhpWord\Writer\HTML\Style\Font as FontStyleWriter;
use PhpOffice\PhpWord\Writer\HTML\Style\Paragraph as ParagraphStyleWriter;

/**
 * Container element HTML writer
 *
 * @since 0.11.0
 */
class Container extends AbstractElement
{
    /**
     * Namespace; Can't use __NAMESPACE__ in inherited class (RTF)
     *
     * @var string
     */
    protected $namespace = 'PhpOffice\\PhpWord\\Writer\\HTML\\Element';

    /**
     * Write paragraph style
     *
     * @return string
     */
    private function getParagraphStyle()
    {
        /** @var \PhpOffice\PhpWord\Element\Text $element Type hint */
        $element = $this->element;
        $style = '';
        if (!method_exists($element, 'getParagraphStyle')) {
            return $style;
        }

        $paragraphStyle = $element->getParagraphStyle();
        $pStyleIsObject = ($paragraphStyle instanceof Paragraph);
        if ($pStyleIsObject) {
            $styleWriter = new ParagraphStyleWriter($paragraphStyle);
            $style = $styleWriter->write();
        }
        if ($style) {
            $attribute = $pStyleIsObject ? 'style' : 'class';
            $style = " {$attribute}=\"{$style}\"";
        }

        if ($pStyleIsObject) {
            $styleName = $paragraphStyle->getStyleName();
        }

        if (!empty($styleName)) {
            $style = " class=\"p{$styleName}\"";
        }

        return $style;
    }
    
    /**
     * Write container
     *
     * @return string
     */
    public function write()
    {
        $container = $this->element;
        if (!$container instanceof ContainerElement) {
            return '';
        }
        $containerClass = substr(get_class($container), strrpos(get_class($container), '\\') + 1);
        $withoutP = in_array($containerClass, array('TextInserted', 'TextRun', 'Footnote', 'Endnote')) ? true : false;
        $content = '';

        $elements = $container->getElements();
        foreach ($elements as $element) {
            $elementClass = get_class($element);
            $writerClass = str_replace('PhpOffice\\PhpWord\\Element', $this->namespace, $elementClass);
            if (class_exists($writerClass)) {
                /** @var \PhpOffice\PhpWord\Writer\HTML\Element\AbstractElement $writer Type hint */
                $_writerClass = substr($writerClass, strrpos($writerClass, '\\') + 1);

                if ($_writerClass === 'TextRun' && $containerClass === 'Section') {
                    $writer = new $writerClass($this->parentWriter, $element, true);
                } else {
                    $writer = new $writerClass($this->parentWriter, $element, $withoutP);
                }
                $c = $writer->write();
                $content .= $c;
            }
        }

        if ($containerClass === 'Section') {
            if (method_exists($this->element, 'getParagraphStyle')) {
                $style = $this->getParagraphStyle();
            }
            $content = "<p{$style}>" . $content . "</p>";
        }

        return $content;
    }
}
