<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 16/11/11
 * Time: PM1:49
 */

namespace PhpOffice\PhpWord\Element;

/**
 * Class TextDeled
 * @package PhpOffice\PhpWord\Element
 */
class TextDeled extends AbstractContainer
{
    /**
     * @var string Container type
     */
    protected $container = 'TextDeled';

    /**
     * Unique Id for comment
     *
     * @var int
     */
    protected $commentId;

    /**
     * 作者 （注释作者）
     * @var string
     */
    protected $author;

    /**
     * 日期 （Annotation 日期）
     * @var string
     */
    protected $date;

    /**
     * 删除的文本
     * @var
     */
    protected $delTextContent;

    /**
     * TextDeled constructor.
     * @param $delTextContent
     * @param $id
     * @param string $author
     * @param string $date
     */
    public function __construct($delTextContent, $id, $author = '', $date = '')
    {
        $this->delTextContent = $delTextContent;
        $this->commentId = $id;
        $this->author = $author;
        $this->date = $date;
    }
}