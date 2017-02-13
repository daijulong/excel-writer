<?php
/**
 * 对指定的多个单元格进行操作
 *
 * @author daijulong <daijulong@gmail.com>
 */

namespace ExcelWriter\Traits;


trait CellsTrait
{
    /**
     * 设置单元格边框
     *
     * 一般用于多个集中设置边框
     *
     * @access public
     * @param string $cells 单元格（组）,多个(组)以“,”分隔，如“A1,A2:B3,C4:D6”
     * @param string $position 边框位置，多个以‘,’分隔，outline|inside|left|top|right|bottom
     * @param string $color 颜色（RGB）
     * @param string $border_type 边框类型<br>
     *                  none|dashDot|dashDotDot|dashed|dotted|double|hair|medium|mediumDashDot|mediumDashDotDot|mediumDashed|slantDashDot|thick|thin
     * @return null
     * @author daijulong <daijulong@gmail.com>
     */
    public function setCellsBorder ($cells, $position = '*', $color = '', $border_type = 'thin')
    {
        if (!is_string($cells) || '' == $cells) {
            return;
        }
        if ($border_type == '') {
            $border_type = 'thin';
        }
        $borders = $this->buildBorderStyle($position, $color, $border_type);
        if (empty($borders)) {
            return;
        }
        $borders_style = ['borders' => $borders];
        $cells_arr = explode(',', $cells);
        foreach ($cells_arr as $cell) {
            $this->active_sheet->getStyle($cell)->applyFromArray($borders_style);
        }
    }

}