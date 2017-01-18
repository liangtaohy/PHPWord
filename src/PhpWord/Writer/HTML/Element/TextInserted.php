<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 16/11/10
 * Time: PM4:41
 */

namespace PhpOffice\PhpWord\Writer\HTML\Element;

/**
 * Class TextInserted
 * @package PhpOffice\PhpWord\Writer\HTML\Element
 */
class TextInserted extends Text
{
    /**
     * Write text run
     *
     * @return string
     */
    public function write()
    {
        $content = '';

        $content .= $this->writeOpening();
        $writer = new Container($this->parentWriter, $this->element);
        $content .= $writer->write();
        $content .= $this->writeClosing();

        return $content;
    }

    /**
     * @return string
     */
    protected function writeOpening()
    {
        $commentId = $this->element->getCommentId();
        $content = "<ins class=\"ins-com\" id=\"ins-com-{$commentId}\" >";

        return $content;
    }

    /**
     * @return string
     */
    protected function writeClosing()
    {
        return "</ins>";
    }
}