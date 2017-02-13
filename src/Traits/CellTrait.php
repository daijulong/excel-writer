<?php
/**
 * 对单元格进行操作
 *
 * 一般为当前活动单元个
 *
 * @author daijulong <daijulong@gmail.com>
 */

namespace ExcelWriter\Traits;

trait CellTrait
{
    /**
     * 取得指定单元格并设置为活动单元格
     *
     * @access
     * @param string|int $coordinate 坐标，如'A1'，如为数字则为列号
     * @param null|int $row 行号，与列号同为数字时，单元格取列号与行号定位的单元格
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function getCell ($coordinate = 'A1', $row = null)
    {
        if (is_numeric($coordinate) && is_numeric($row)) {
            $this->active_cell = $this->active_sheet->getCellByColumnAndRow($coordinate, $row);
            return $this;
        }
        $this->active_cell = $this->active_sheet->getCell($coordinate);
        return $this;
    }

    /**
     * 获取活动单元格样式
     *
     * @access private
     * @return \PHPExcel_Style
     * @author daijulong <daijulong@gmail.com>
     */
    private function getActiveCellStyle ()
    {
        return $this->active_cell->getStyle();
    }

    /**
     * 设置单元格文字样式（数组配置）
     *
     * @param array $style 样式数组
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setFontStyle ($style)
    {
        $this->getActiveCellStyle()->applyFromArray($style);
        return $this;
    }

    /**
     * 设置单元格字体
     *
     * @access public
     * @param string $font_name 字体名称
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setFontName ($font_name = '')
    {
        if ('' == $font_name) {
            return $this;
        }
        $this->getActiveCellStyle()->getFont()->setName($font_name);
        return $this;
    }

    /**
     * 设置单元格文字颜色
     *
     * @access public
     * @param string $color 颜色（RGB）
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setFontColor ($color)
    {
        $this->getActiveCellStyle()->getFont()->setColor(new \PHPExcel_Style_Color($color));
        return $this;
    }

    /**
     * 设置单元格文字大小
     *
     * @access public
     * @param float $size 字体大小
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setFontSize ($size)
    {
        $this->getActiveCellStyle()->getFont()->setSize($size);
        return $this;
    }

    /**
     * 设置单元格文字加粗
     *
     * @access public
     * @param bool $bold 是否加粗
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setFontBold ($bold = true)
    {
        $this->getActiveCellStyle()->getFont()->setBold($bold);
        return $this;
    }

    /**
     * 设置单元格文字斜体
     *
     * @access public
     * @param bool $italic 是否斜体
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setFontItalic ($italic = true)
    {
        $this->getActiveCellStyle()->getFont()->setItalic($italic);
        return $this;
    }

    /**
     * 应用单元格样式（多个一次性应用）
     *
     * @access public
     * @param array $style 样式数组
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function applyStyle ($style)
    {
        $this->getActiveCellStyle()->applyFromArray($style);
        return $this;
    }

    /**
     * 设置单元格缩进
     *
     * @access public
     * @param int $indent 缩进值
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setIndent ($indent)
    {
        $this->getActiveCellStyle()->getAlignment()->setIndent($indent);
        return $this;
    }

    /**
     * 设置水平对齐
     *
     * @access public
     * @param string $type 对齐方式 general|left|right|center|centerContinuous|justify|fill|distributed(XLSX可用)
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setHorizontal ($type = 'general')
    {
        $this->getActiveCellStyle()->getAlignment()->applyFromArray(['horizontal' => $type]);
        return $this;
    }

    /**
     * 设置垂直对齐
     *
     * @access public
     * @param string $type 对齐方式 center|top|bottom|justify|distributed(XLSX可用)
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setVertical ($type = 'center')
    {
        $this->getActiveCellStyle()->getAlignment()->applyFromArray(['vertical' => $type]);
        return $this;
    }

    /**
     * 设置自动换行
     *
     * @access public
     * @param bool $wrap 是否自动换行
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setWrap ($wrap = true)
    {
        $this->getActiveCellStyle()->getAlignment()->applyFromArray(['wrap' => $wrap]);
        return $this;
    }

    /**
     * 设置缩小到合适
     *
     * @access public
     * @param bool $shrink2fit 是否缩小到合适大小
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setShrinkFit ($shrink2fit = true)
    {
        $this->getActiveCellStyle()->getAlignment()->applyFromArray(['shrinkToFit' => $shrink2fit]);
        return $this;
    }

    /**
     * 设置单元格边框
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
    public function setBorder ($position = '*', $color = '', $border_type = 'thin', $cells = '')
    {
        $borders = $this->buildBorderStyle($position, $color, $border_type);
        if (empty($borders)) {
            return $this;
        }
        $this->getActiveCellStyle()->applyFromArray(['borders' => $borders]);
        if ($cells != '') {
            $this->setCellsBorder($cells, $position, $color, $border_type);
        }
        return $this;
    }

    /**
     * 取消边框
     *
     * @access public
     * @param string $position 边框位置，多个以‘,’分隔，outline|inside|left|top|right|bottom
     * @param string $cells 单元格（组）,多个(组)以“,”分隔，如“A1,A2:B3,C4:D6”，如指定则同时设置这些单元格
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setBorderNone ($position = '*', $cells)
    {
        $borders = $this->buildBorderStyle($position, '', 'none');
        if (empty($borders)) {
            return $this;
        }
        $this->getActiveCellStyle()->applyFromArray(['borders' => $borders]);
        if ($cells != '') {
            $this->setCellsBorder($cells, $position, '', 'none');
        }
        return $this;
    }

    /**
     * 生成边框样式数组
     *
     * @access private
     * @param string $position $position 边框位置，多个以‘,’分隔，outline|inside|left|top|right|bottom
     * @param string $color 边框颜色（RGB）
     * @param string $border_type 边框类型<br>
     *                  none|dashDot|dashDotDot|dashed|dotted|double|hair|medium|mediumDashDot|mediumDashDotDot|mediumDashed|slantDashDot|thick|thin
     * @return array
     * @author daijulong <daijulong@gmail.com>
     */
    private function buildBorderStyle ($position = '*', $color = '', $border_type = 'thin')
    {
        if ($position == '' || $position == 'all' || $position == '*') {
            $position = 'allborders';
        }
        $positions = explode(',', $position);
        $borders = [];
        foreach ($positions as $pos) {
            $borders[$pos] = [
                'style' => $border_type ?: 'thin',
                'color' => ['rgb' => $color],
            ];
        }
        return $borders;
    }

    /**
     * 当前单元格赋值
     *
     * @access public
     * @param mixed $value 值
     * @param string $type 值类型，str或s:字符串|f:公式|n:数字|b:布尔值|null|inlineStr|e
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setValue ($value = '', $type = 'str')
    {
        if (!$type) {
            $type = 'str';
        }
        $this->active_cell->setValueExplicit($value, $type);
        return $this;
    }

    /**
     * 单元格批量赋值
     *
     * @access public
     * @param string|array $cell 单元格，支持数组array(array(0=>单元格,1=>值,2=>值类型）,...)
     * @param string $value 值，$cell为string时，作为$cell的值
     * @param string $type 值类型，str或s:字符串|f:公式|n:数字|b:布尔值|null|inlineStr|e ，$cell为数组时，$type作为默认值类型
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setCellsValue ($cell, $value = '', $type = 'str')
    {
        if (is_string($cell) && '' != $cell) {
            $this->active_sheet->setCellValueExplicit($cell, $value, $type);
        } else if (is_array($cell) && !empty($cell)) {
            foreach ($cell as $_cell) {
                $this->active_sheet->setCellValueExplicit($_cell[0], $_cell[1], $_cell[2] ? $_cell[2] : $type);
            }
        }
        return $this;
    }

    /**
     * 合并单元格
     *
     * @access public
     * @param string|array $cells 单元格，多组合并则使用数组或字符串以“,”分隔
     * @author daijulong <daijulong@gmail.com>
     */
    public function mergeCells ($cells)
    {
        if (is_string($cells) && '' != $cells) {
            $cells = explode(',', $cells);
        }
        if (is_array($cells) && !empty($cells)) {
            foreach ($cells as $_cell) {
                $this->active_sheet->mergeCells($_cell);
            }
        }
    }

    /**
     * 设置单元格背景颜色
     *
     * @access public
     * @param string $color 颜色（RGB）
     * @param string $type 填充类型
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function setBackgroundColor ($color, $type = 'solid')
    {
        $this->getActiveCellStyle()->getFill()->setFillType($type);
        $this->getActiveCellStyle()->getFill()->getStartColor()->setRGB($color);
        return $this;
    }

    /**
     * 添加注释
     *
     * @access public
     * @param string $comment 注释内容
     * @return $this
     * @author daijulong <daijulong@gmail.com>
     */
    public function comment ($comment)
    {
        $comment_obj = $this->active_sheet->getComment($this->active_cell->getCoordinate());
        $comment_obj->getText()->createTextRun($comment);
        return $this;
    }
}