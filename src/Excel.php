<?php
namespace ExcelWriter;

use ExcelWriter\Traits\BatchTrait;
use ExcelWriter\Traits\ImageTrait;
use ExcelWriter\Traits\SheetTrait;
use ExcelWriter\Traits\CellTrait;
use ExcelWriter\Traits\CellsTrait;
use ExcelWriter\Traits\CellAliasTrait;

class Excel
{
    use SheetTrait, CellTrait, CellsTrait, CellAliasTrait, BatchTrait, ImageTrait;

    /**
     * PHPExcel实例
     *
     * @var null|\PHPExcel
     */
    private $excel = null;

    /**
     * 当前活动工作表
     *
     * @var null|\PHPExcel_Worksheet
     */
    private $active_sheet = null;

    /**
     * 当前活动单元格
     * @var null
     */
    private $active_cell = null;


    /**
     * Excel constructor.
     *
     * @param string $title 默认创建工作表的标题
     * @param array $layout 默认创建工作表排版
     * @param array $style 默认创建工作表内容默认样式
     * @param string $tab_color 默认创建工作表TAB颜色（RGB）
     */
    public function __construct ($title = '', $layout = [], $style = [], $tab_color = '')
    {
        $this->excel = new \PHPExcel();
        $this->active_sheet = $this->excel->getActiveSheet();
        $this->getCell($this->active_sheet->getActiveCell());
        $this->applyActiveSheetSettings($title, $layout, $style, $tab_color);
    }

    /**
     * 取得当前活动工作表
     *
     * @access public
     * @return null|\PHPExcel_Worksheet
     * @author daijulong <daijulong@gmail.com>
     */
    public function getActiveSheet ()
    {
        return $this->active_sheet;
    }

    /**
     * 切换当前活动工作表
     *
     * @access public
     * @param int $sheet_id 工作表ID
     * @return null|\PHPExcel_Worksheet
     * @author daijulong <daijulong@gmail.com>
     */
    public function switchActiveSheet ($sheet_id)
    {
        $this->active_sheet = $this->excel->getSheet($sheet_id);
        $this->getCell($this->active_sheet->getActiveCell());
        $this->excel->setActiveSheetIndex($sheet_id);
        return $this->active_sheet;
    }


    /**
     * 生成并保存文件
     *
     * @access public
     * @param string $name 文件名，不包括路径与扩展名，扩展名将自动添加
     * @param string $type 文件类型，xls|xlsx|html等
     * @author daijulong <daijulong@gmail.com>
     */
    public function download ($name, $type = 'xlsx')
    {
        switch ($this->getWriterTrueType($type)) {
            case 'Excel5':
            case 'Excel2007':
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                break;
            case 'HTML':
                break;
            default:
        }
        $objWriter = \PHPExcel_IOFactory:: createWriter($this->excel, $this->getWriterTrueType($type));
        header('Content-Disposition: attachment;filename=' . $name . '.' . $type);
        header('Cache-Control: max-age=0');
        $objWriter->save('php://output');
    }

    /**
     * 生成并保存文件
     *
     * @access public
     * @param string $full_name 文件全名（包括路径和扩展名）
     * @param string $type 文件类型，xls|xlsx|pdf|html|csv等
     * @return string 文件名
     * @author daijulong <daijulong@gmail.com>
     */
    public function save ($full_name, $type = 'xlsx')
    {
        $objWriter = \PHPExcel_IOFactory:: createWriter($this->excel, $this->getWriterTrueType($type));
        $objWriter->save($full_name);
        return $full_name;
    }

    /**
     * 取得实际的写入文件类型
     *
     * @access private
     * @param string $type 类型
     * @return mixed|string
     * @author daijulong <daijulong@gmail.com>
     */
    private function getWriterTrueType ($type)
    {
        $types = [
            'xls' => 'Excel5',
            'xlsx' => 'Excel2007',
            'html' => 'HTML'
        ];
        if (isset($types[$type])) {
            return $types[$type];
        }
        return 'Excel2007';
    }

}