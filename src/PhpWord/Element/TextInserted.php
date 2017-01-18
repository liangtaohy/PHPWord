<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 16/11/10
 * Time: PM3:47
 */

namespace PhpOffice\PhpWord\Element;


class TextInserted extends AbstractContainer
{
    /**
     * @var string Container type
     */
    protected $container = 'TextInserted';

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
     * TextInserted constructor.
     * @param $id
     * @param string $author
     * @param string $date
     */
    public function __construct($id, $author = '', $date = '')
    {
        $this->commentId = $id;
        $this->author = $author;
        $this->date = $date;
    }
}