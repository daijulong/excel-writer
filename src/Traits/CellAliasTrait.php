<?php
/**
 * 单元格操作的快捷方式
 *
 * 一般链接到 CellTrait 中的方法
 *
 * @author daijulong <daijulong@gmail.com>
 */

namespace ExcelWriter\Traits;

trait CellAliasTrait
{

    /**
     * 取得指定单元格并设置为活动单元格
     *
     * Alias for $this->getCell()
     *
     * @access
     * @param string|int $coordinate 坐标，如'A1'，如为数字则为列号
     * @param null|int $row 行号，与列号同为数字时，单元格取列号与行号定位的单元格
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function cell ($coordinate = 'A1', $row = null)
    {
        return $this->getCell($coordinate, $row);
    }

    /**
     * 合并单元格
     *
     * Alias for $this->mergeCells()
     *
     * @access public
     * @param string|array $cells 单元格，多组合并则使用数组或字符串以“,”分隔
     * @author daijulong <daijulong@gmail.com>
     */
    public function merge ($cells)
    {
        $this->mergeCells($cells);
    }

    /**
     * 当前单元格赋值
     *
     * Alias for $this->setValue()
     *
     * @access public
     * @param mixed $value 值
     * @param string $type 值类型，str或s:字符串|f:公式|n:数字|b:布尔值|null|inlineStr|e
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function value ($value = '', $type = 'str')
    {
        return $this->setValue($value, $type);
    }

    /**
     * 设置单元格文字颜色
     *
     * Alias for $this->setFontColor()
     *
     * @access public
     * @param string $color 颜色（RGB）
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function color ($color)
    {
        return $this->setFontColor($color);
    }

    /**
     * 设置单元格文字加粗
     *
     * Alias for $this->setFontBold(true)
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function bold ()
    {
        return $this->setFontBold(true);
    }

    /**
     * 取消单元格文字加粗
     *
     * Alias for $this->setFontBold(false)
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function notBold ()
    {
        return $this->setFontBold(false);
    }

    /**
     * 设置水平左对齐
     *
     * Alias for $this->setHorizontal('left')
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function left ()
    {
        return $this->setHorizontal('left');
    }

    /**
     * 设置水平居中对齐
     *
     * Alias for $this->setHorizontal('center')
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function center ()
    {
        return $this->setHorizontal('center');
    }

    /**
     * 设置水平右对齐
     *
     * Alias for $this->setHorizontal('right')
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function right ()
    {
        return $this->setHorizontal('right');
    }

    /**
     * 设置垂直上对齐
     *
     * Alias for $this->setVertical('top')
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function top ()
    {
        return $this->setVertical('top');
    }

    /**
     * 设置垂直下对齐
     *
     * Alias for $this->setVertical('bottom')
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function bottom ()
    {
        return $this->setVertical('bottom');
    }

    /**
     * 设置垂直居中对齐
     *
     * Alias for $this->setVertical('center')
     *
     * @access public
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function centerV ()
    {
        return $this->setVertical('center');
    }

    /**
     * 设置单元格文字大小
     *
     * Alias for $this->setFontSize()
     *
     * @access public
     * @param float $size 字体大小
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function size ($size)
    {
        return $this->setFontSize($size);
    }

    /**
     * 设置单元格缩进
     *
     * Alias for $this->setIndent()
     *
     * @access public
     * @param int $indent 缩进值
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function indent ($indent)
    {
        return $this->setIndent($indent);
    }

    /**
     * 设置单元格边框
     *
     * Alias for $this->setBorder()
     *
     * @access public
     * @param string $position 边框位置，多个以‘,’分隔，outline|inside|left|top|right|bottom
     * @param string $color 边框颜色（RGB）
     * @param string $border_type 边框类型<br>
     *                  none|dashDot|dashDotDot|dashed|dotted|double|hair|medium|mediumDashDot|mediumDashDotDot|mediumDashed|slantDashDot|thick|thin
     * @param string $cells 单元格（组）,多个(组)以“,”分隔，如“A1,A2:B3,C4:D6”，如指定则同时设置这些单元格
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function border ($position = '*', $color = '', $border_type = 'thin', $cells = '')
    {
        return $this->setBorder($position, $color, $border_type, $cells);
    }

    /**
     * 取消边框
     *
     * Alias for $this->setBorderNone()
     *
     * @access public
     * @param string $position 边框位置，多个以‘,’分隔，outline|inside|left|top|right|bottom
     * @param string $cells 单元格（组）,多个(组)以“,”分隔，如“A1,A2:B3,C4:D6”，如指定则同时设置这些单元格
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function borderNone ($position = '*', $cells = '')
    {
        return $this->setBorderNone($position, $cells);
    }

    /**
     * 设置单元格文字斜体
     *
     * Alias for $this->setFontItalic()
     *
     * @access public
     * @param bool $italic 是否斜体
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function italic ($italic = true)
    {
        return $this->setFontItalic($italic);
    }

    /**
     * 设置自动换行
     *
     * Alias for $this->setWrap()
     *
     * @access public
     * @param bool $wrap 是否自动换行
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function wrap ($wrap = true)
    {
        return $this->setWrap($wrap);
    }
}