<?php
/**
 * 对工作表的基本操作
 *
 * @author daijulong <daijulong@gmail.com>
 */

namespace ExcelWriter\Traits;

trait SheetTrait
{
    /**
     * 纸张大小别名
     *
     * @var array
     */
    private static $papers_alias = [
        'A3' => \PHPExcel_Worksheet_PageSetup::PAPERSIZE_A3,
        'A4' => \PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4,
        'A4_SMALL' => \PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4_SMALL,
        'A5' => \PHPExcel_Worksheet_PageSetup::PAPERSIZE_A5,
        'B4' => \PHPExcel_Worksheet_PageSetup::PAPERSIZE_B4,
        'B5' => \PHPExcel_Worksheet_PageSetup::PAPERSIZE_B5,
    ];

    /**
     * 设置行高
     *
     * @access public
     * @param int $row 行号
     * @param float $hight 行高
     * @author daijulong <daijulong@gmail.com>
     */
    public function setRowHight ($row, $hight = -1)
    {
        $this->active_sheet->getRowDimension($row)->setRowHeight($hight);
    }

    /**
     * 设置列宽
     *
     * @access public
     * @param array $col_width 列=>宽度数组
     * @return bool
     * @author daijulong <daijulong@gmail.com>
     */
    public function setColsWidth ($col_width)
    {
        if (!is_array($col_width) || empty($col_width)) {
            return false;
        }
        //列宽
        foreach ($col_width as $col => $wd) {
            $this->active_sheet->getColumnDimension($col)->setWidth($wd);
        }
        return true;
    }


    /**
     * 创建一个新的工作表
     *
     * @access public
     * @param string $title 工作表标题
     * @param array $layout 工作表排版
     * @param array $style 工作表内容默认样式
     * @param string $tab_color 工作表TAB颜色（RGB）
     * @return null|\PHPExcel_Worksheet
     * @author daijulong <daijulong@gmail.com>
     */
    public function createSheet ($title = '', $layout = [], $style = [], $tab_color = '')
    {
        $this->active_sheet = $this->excel->createSheet();
        $this->applyActiveSheetSettings($title, $layout, $style, $tab_color);
        $this->getCell($this->active_sheet->getActiveCell());
        return $this->active_sheet;
    }

    /**
     * 工作表应用设置
     *
     * 用于设置工作表标题、默认排版、默认样式等
     *
     * @access public
     * @param string $title 工作表标题
     * @param array $layout 工作表排版
     * @param array $style 工作表内容默认样式
     * @param string $tab_color 工作表TAB颜色（RGB）
     * @return null|\PHPExcel_Worksheet
     * @author daijulong <daijulong@gmail.com>
     */
    public function applyActiveSheetSettings ($title = '', $layout = [], $style = [], $tab_color = '')
    {
        if ($title != '') {
            $this->active_sheet->setTitle($title);
        }
        if ($tab_color != '') {
            $this->active_sheet->getTabColor()->setRGB($tab_color);
        }
        $this->applySheetDefaultLayoutAndStyle($this->active_sheet, $layout, $style);
        return $this->active_sheet;
    }

    /**
     * 工作表应用默认排版和样式
     *
     * @access private
     * @param \PHPExcel_Worksheet $sheet 工作表
     * @param array $layout 默认排版 <br>
     *                      $layout = [
     *                           'header_footer' => '',//页眉页脚
     *                           //边距
     *                           'margin' => [
     *                               'left' => 0,
     *                               'right' => 0,
     *                               'top' => 0,
     *                               'bottom' => 0,
     *                               'header' => 0,
     *                               'footer' => 0,
     *                           ],
     *                           //默认行高
     *                           'row_height' => 0,
     *                           //默认列宽
     *                           'col_width' => 0,
     *                           //打印和纸张
     *                           'page' => [
     *                              'fit_to_width' => 1,
     *                              'fit_to_height' => 0,
     *                              'paper_size' => 'A4',
     *                              'orientation' => 'default',//打印方向，default:默认|portrait:纵向|landscape:横向
     *                              'header' => '',//页眉，可用标记，标记请查看\PHPExcel_Worksheet_HeaderFooter
     *                              'footer' => '',//页脚，可用标记
     *                           ],
     *                       ];
     * @param array $style 默认样式 <br>
     *                      $style = [
     *                           //字体
     *                           'font' => [
     *                               'name' => '',
     *                               'size' => '',
     *                               'color' => '',
     *                               'bold' => false,
     *                               'italic' => false
     *                           ],
     *                           //对齐
     *                           'align' => [
     *                               'horizontal' => '',// 见\PHPExcel_Style_Alignment
     *                               'vertical' => '',// 见\PHPExcel_Style_Alignment
     *                           ],
     *                       ];
     * @return \PHPExcel_Worksheet
     * @author daijulong <daijulong@gmail.com>
     */
    private function applySheetDefaultLayoutAndStyle (\PHPExcel_Worksheet $sheet, $layout = [], $style = [])
    {
        //页边距
        if (isset($layout['margin'])) {
            (!isset($layout['margin']['left']) || !is_numeric($layout['margin']['left'])) ?: $sheet->getPageMargins()->setLeft($layout['margin']['left']);
            !isset($layout['margin']['top']) || !is_numeric($layout['margin']['top']) ?: $sheet->getPageMargins()->setTop($layout['margin']['top']);
            !isset($layout['margin']['right']) || !is_numeric($layout['margin']['right']) ?: $sheet->getPageMargins()->setRight($layout['margin']['right']);
            !isset($layout['margin']['bottom']) || !is_numeric($layout['margin']['bottom']) ?: $sheet->getPageMargins()->setBottom($layout['margin']['bottom']);
            !isset($layout['margin']['header']) || !is_numeric($layout['margin']['header']) ?: $sheet->getPageMargins()->setHeader($layout['margin']['header']);
            !isset($layout['margin']['footer']) || !is_numeric($layout['margin']['footer']) ?: $sheet->getPageMargins()->setFooter($layout['margin']['footer']);
        }

        //样式-文字
        if (isset($style['font']) && !empty($style['font'])) {
            !isset($style['font']['name']) || $style['font']['name'] == '' ?: $sheet->getDefaultStyle()->getFont()->setName($style['font']['name']);
            !isset($style['font']['bold']) ?: $sheet->getDefaultStyle()->getFont()->setBold($style['font']['bold']);
            !isset($style['font']['color']) || $style['font']['color'] == '' ?: $sheet->getDefaultStyle()->getFont()->setColor(new \PHPExcel_Style_Color($style['font']['color']));
            !isset($style['font']['size']) || !is_numeric($style['font']['size']) ?: $sheet->getDefaultStyle()->getFont()->setSize($style['font']['size']);
        }

        //对齐
        if (isset($style['align']) && !empty($style['align'])) {
            !isset($style['align']['horizontal']) || $style['align']['horizontal'] == '' ?: $sheet->getDefaultStyle()->getAlignment()->setHorizontal($style['align']['horizontal']);
            !isset($style['align']['vertical']) || $style['align']['vertical'] == '' ?: $sheet->getDefaultStyle()->getAlignment()->setVertical($style['align']['vertical']);
        }

        //行高
        if (isset($layout['row_height']) && is_numeric($layout['row_height'])) {
            $sheet->getDefaultRowDimension()->setRowHeight($layout['row_height']);
        }

        //列宽
        if (isset($layout['col_width']) && is_numeric($layout['col_width'])) {
            $sheet->getDefaultColumnDimension()->setWidth($layout['col_width']);
        }

        //打印、纸张
        if (isset($layout['page'])) {
            !isset($layout['page']['fit_to_width']) ?: $sheet->getPageSetup()->setFitToWidth($layout['page']['fit_to_width']);
            !isset($layout['page']['fit_to_height']) ?: $sheet->getPageSetup()->setFitToHeight($layout['page']['fit_to_height']);
            if (isset($layout['page']['paper_size'])) {
                $paper_size = $layout['page']['paper_size'];
                if (!is_numeric($paper_size)) {
                    $paper_size = strtoupper($paper_size);
                    $paper_size = key_exists($paper_size, self::$papers_alias) ? self::$papers_alias[$paper_size] : '';
                }
                $sheet->getPageSetup()->setPaperSize($paper_size);
            }
            if (isset($layout['page']['orientation']) && in_array($layout['page']['orientation'], ['portrait', 'landscape'])) {
                $sheet->getPageSetup()->setOrientation($layout['page']['orientation']);
            }
        }

        //页眉页脚
        !isset($layout['page']['header']) ?: $sheet->getHeaderFooter()->setOddHeader($layout['page']['header']);
        !isset($layout['page']['footer']) ?: $sheet->getHeaderFooter()->setOddFooter($layout['page']['footer']);

        return $sheet;
    }

}