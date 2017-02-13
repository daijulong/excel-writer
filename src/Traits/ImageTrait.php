<?php
/**
 * 图片相关操作
 *
 * @author daijulong <daijulong@gmail.com>
 */

namespace ExcelWriter\Traits;

trait ImageTrait
{
    /**
     * 插入图片
     *
     * @access public
     * @param string $cell 单元格坐标
     * @param string $img_file 图片文件路径
     * @param int $width 显示宽度
     * @param int $height 显示高度
     * @param int $offset_x X坐标偏移量
     * @param int $offset_y Y坐标偏移量
     * @param array $options 其他选项，name，description 等
     * @return null
     * @author daijulong <daijulong@gmail.com>
     */
    public function setImage ($cell = '', $img_file, $width = 0, $height = 0, $offset_x = 0, $offset_y = 0, $options = [])
    {
        if (is_null($cell)) {
            return;
        }
        $drawing_obj = new \PHPExcel_Worksheet_Drawing();
        $drawing_obj->setCoordinates($cell);
        $drawing_obj->setPath($img_file);
        if ($width > 0) {
            $drawing_obj->setWidth($width);
        }
        if ($height > 0) {
            $drawing_obj->setHeight($height);
        }
        $drawing_obj->setOffsetX($offset_x);
        $drawing_obj->setOffsetY($offset_y);
        if (isset($options['name'])) {
            $drawing_obj->setName($options['name']);
        }
        if (isset($options['description'])) {
            $drawing_obj->setDescription($options['description']);
        }
        $drawing_obj->setWorksheet($this->active_sheet);
        unset($drawing_obj);
    }
}