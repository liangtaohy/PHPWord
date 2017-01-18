<?php
/**
 * Created by PhpStorm.
 * User: xlegal
 * Date: 16/11/15
 * Time: PM4:42
 */

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Element;

class DrawingML extends AbstractElement
{
    const IMAGE_MODE_INLINE = 'inline';
    const IMAGE_MODE_ANCHOR = 'anchor';

    /**
     * Image Style
     *
     * @var \PhpOffice\PhpWord\Style\Image
     */
    private $style;

    /**
     * Image Relation ID
     *
     * @var int
     */
    private $imageRelationId;

    /**
     * Has media relation flag; true for Link, Image, and Object
     *
     * @var bool
     */
    protected $mediaRelation = true;

    /**
     * Image Source
     * @var
     */
    private $imageSource;

    /**
     * Icon
     *
     * @var string
     */
    private $icon;

    /**
     * Image Mode: Inline
     *
     * @var string
     */
    private $mode;

    /**
     * DrawingML constructor.
     * @param $rid
     * @param string $mode
     * @param null $style
     */
    public function __construct($rid, $target, $mode = 'inline', $style = null)
    {
        $this->imageRelationId = $rid;
        $this->mode = $mode;
        $this->imageSource = $target;
    }

    /**
     * Get object style
     *
     * @return \PhpOffice\PhpWord\Style\Image
     */
    public function getStyle()
    {
        return $this->style;
    }

    /**
     * Get object icon
     *
     * @return string
     */
    public function getIcon()
    {
        return $this->icon;
    }

    /**
     * Get image relation ID
     *
     * @return int
     */
    public function getImageRelationId()
    {
        return $this->imageRelationId;
    }

    /**
     * Set Image Relation ID.
     *
     * @param int $rId
     * @return void
     */
    public function setImageRelationId($rId)
    {
        $this->imageRelationId = $rId;
    }
}