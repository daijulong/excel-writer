<?php
/**
 * 对工作表中单元格的批量操作
 *
 * @author daijulong <daijulong@gmail.com>
 */

namespace ExcelWriter\Traits;


trait BatchTrait
{
    /**
     * 添加一行数据
     *
     * @access public
     * @param array $data 数据，定义值，类型，样式等
     *
     *         $data = [
     *               'A' => ['序号', 'str'],
     *               'B' => ['姓名', 'str', ['font' => ['color' => ['rgb' => '000000']]]],
     *               'C' => '班级',
     *               'D' => '数学',
     *               'E' => '语文',
     *               'F' => '英语',
     *               'G' => '物理',
     *               'H' => '化学',
     *               'I' => '总分',
     *               'J' => '平均分',
     *          ];
     *
     * @param int $row 所在行
     * @param array $default_style
     *         [
     *             'font'    => [
     *                 'name'      => 'Arial',
     *                 'bold'      => true,
     *                 'italic'    => false,
     *                 'underline' => PHPExcel_Style_Font::UNDERLINE_DOUBLE,
     *                 'strike'    => false,
     *                 'color'     => [
     *                     'rgb' => '808080'
     *                 ]
     *             ],
     *             'borders' => [
     *                 'bottom'     => [
     *                     'style' => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     'color' => [
     *                         'rgb' => '808080'
     *                     ]
     *                 ],
     *                 'top'     => [
     *                     'style' => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     'color' => [
     *                         'rgb' => '808080'
     *                     ]
     *                 ]
     *             ],
     *             'quotePrefix'    => true
     *         ]
     * @author daijulong <daijulong@gmail.com>
     */
    public function batchFillRow ($data, $row = 1, $default_style = [])
    {
        foreach ($data as $col => $cnt) {
            $_style = [];
            $type = 'str';
            if (!is_array($cnt)) {
                $value = $cnt;
            } else {
                $value = isset($cnt[0]) ? $cnt[0] : '';
                $type = isset($cnt[1]) && $cnt[1] != '' ? $cnt[1] : 'str';
                $_style = isset($cnt[2]) && is_array($cnt[2]) ? $cnt[2] : [];
            }
            $style = empty($_style) ? $default_style : array_replace_recursive($default_style, $_style);
            $cell_coll = !is_numeric($col) ? $col . $row : $col;
            $this->getCell($cell_coll, $row)->setValue($value, $type)->getActiveCellStyle()->applyFromArray($style)->getAlignment()->applyFromArray(isset($style['align']) ? $style['align'] : []);
        }
    }

    /**
     * 添加多行数据
     *
     * @access public
     * @param array $data 数据，类似数据库中查询到的多条数据记录集合，但不能有多余的列，且单条数据顺序要与列吻合
     * @param string $start_col 开始填充的列
     * @param int $start_row 开始填充的行号
     * @param array $cols_data_type 列格式，['A'=>'n','D'=>'f'...]，未指定的列将默认使用'str'
     * @param array $cols_style 每列的样式，如有未在$default_style描述的样式，在此声明即可
     *       'A' => [
     *             'font'    => [
     *                 'name'      => 'Arial',
     *                 'bold'      => true,
     *                 'italic'    => false,
     *                 'underline' => PHPExcel_Style_Font::UNDERLINE_DOUBLE,
     *                 'strike'    => false,
     *                 'color'     => [
     *                     'rgb' => '808080'
     *                 ]
     *             ],
     *             'borders' => [
     *                 'bottom'     => [
     *                     'style' => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     'color' => [
     *                         'rgb' => '808080'
     *                     ]
     *                 ],
     *                 'top'     => [
     *                     'style' => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     'color' => [
     *                         'rgb' => '808080'
     *                     ]
     *                 ]
     *             ],
     *             'quotePrefix'    => true
     *         ],
     *          'B' => [],...
     * @param  array $default_style 默认样式，每一列都将使用此样式，格式同$cols_style中描述
     * @return int 最后插入行的行号
     * @author daijulong <daijulong@gmail.com>
     */
    public function batchFillRows ($data, $start_col = 'A', $start_row = 1, $cols_data_type = [], $cols_style = [], $default_style = [])
    {
        if (empty($data)) {
            return $start_row;
        }
        if (!empty($cols_style)) {
            foreach ($cols_style as $col => $style) {
                $cols_style[$col] = empty($style) ? $default_style : array_replace_recursive($default_style, $style);
            }
        }
        $start_col_index = \PHPExcel_Cell::columnIndexFromString($start_col);
        $data = array_values($data);
        foreach ($data as $srow => $row_data) {
            $curr_row_data = array_values($row_data);
            foreach ($curr_row_data as $key => $cell_data) {
                $cell_col = \PHPExcel_Cell::stringFromColumnIndex($start_col_index + $key - 1);
                $cell_style = isset($cols_style[$cell_col]) ? $cols_style[$cell_col] : $default_style;
                $this->getCell($cell_col . $start_row)->setValue(str_replace(['#DROW#', '#CROW#'], [$srow + 1, $start_row], $cell_data), isset($cols_data_type[$cell_col]) ? $cols_data_type[$cell_col] : 'str')->getActiveCellStyle()->applyFromArray($cell_style)->getAlignment()->applyFromArray(isset($cell_style['align']) ? $cell_style['align'] : []);
            }
            $start_row++;
        }
        return $start_row - 1;
    }
}