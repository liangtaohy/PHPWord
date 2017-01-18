<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 16/11/18
 * Time: PM3:01
 */

namespace PhpOffice\PhpWord\Writer\HTML\Element;

/**
 * Class DrawingML
 * @package PhpOffice\PhpWord\Writer\HTML\Element
 */
class DrawingML extends AbstractElement
{
    /**
     * Text written after opening
     *
     * @var string
     */
    private $openingText = '';

    /**
     * Text written before closing
     *
     * @var string
     */
    private $closingText = '';

    /**
     * Opening tags
     *
     * @var string
     */
    private $openingTags = '';

    /**
     * Closing tag
     *
     * @var string
     */
    private $closingTags = '';

    public function write()
    {
        $element = $this->element;

        $output = '<img src="baidu.png" border="0">';
    }
}