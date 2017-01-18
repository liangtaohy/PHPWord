<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 16/11/11
 * Time: PM4:35
 */

namespace PhpOffice\PhpWord\Writer\HTML\Element;


class TextDeled extends Text
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
        $content .= $this->element->getDelTextContent();
        $content .= $this->writeClosing();

        return $content;
    }

    /**
     * @return string
     */
    protected function writeOpening()
    {
        $commentId = $this->element->getCommentId();
        $content = "<del class=\"del-com\" id=\"del-com-{$commentId}\" >";

        return $content;
    }

    /**
     * @return string
     */
    protected function writeClosing()
    {
        return "</del>";
    }
}