<?php
ini_set('display_errors', 'on');
require '../vendor/autoload.php';

/*
 * 绘制第一个工作表
 */

$border_color = '333333';

//默认排版
$sheet_default_layout = [
    //页边距，上下留的大点，放页眉页脚
    'margin' => [
        'left' => 0.5 / 2.54,
        'top' => 1 / 2.54,
        'right' => 0.5 / 2.54,
        'bottom' => 1 / 2.54,
        'header' => 0,
    ],
    'row_height' => 16.5,
    'page' => [
        'fit_to_width' => 1, //自动适应到页面的宽度
        'fit_to_height' => 0, //因表格较长，取消默认的高度自动适应到一张纸
        'paper_size' => 'A4',
        'orientation' => '',//打印方向，default:默认|portrait:纵向|landscape:横向
        'header' => '&R&I&K00FF00&8&A',//页眉
        'footer' => '&C 第 &P / &N 页',//页脚
    ],
];
//默认样式
$sheet_default_style = [
    'font' => [
        'name' => '微软雅黑',
        'size' => 10,
//        'color' => '00FF00',
//        'bold' => true
    ],
    'align' => [
        'horizontal' => 'left',
        'vertical' => 'center',
    ],
];
//创建Excel对象
$excel = new \ExcelWriter\Excel('单据表格演示', $sheet_default_layout, $sheet_default_style, 'FF0000');


//设置列宽
$excel->setColsWidth(['A' => 9.625, 'B' => 10.625, 'C' => 14, 'D' => 20.875, 'E' => 8.625, 'F' => 8.625, 'G' => 8.625, 'H' => 8.625, 'I' => 8.625, 'J' => 4.75, 'K' => 9.5,]);
//第一行
$excel->setRowHight(1, 20.25);//行高
$excel->merge('A1:F1,H1:I1');//合并单元格
$excel->cell('A1')->value('天下第一修车厂 - 报价单')->center()->color('FF0000')->bold();
$excel->cell('G1')->value('订单编号')->bold()->center();
$excel->cell('H1')->value('TC-150308-03')->center();
$excel->cell('J1')->value('顾问')->bold()->center();
$excel->cell('K1')->value('李大白')->center();
$excel->setCellsBorder('A1:F1,G1,H1:I1,J1,K1', '*', $border_color);//设置边框
//第二行
$excel->merge('A2:F2,G2:K2');
$excel->cell('A2')->value('车辆信息')->center();
$excel->cell('G2')->value('客户信息')->center();
$excel->setCellsBorder('A2:F2,G2:K2', '*', $border_color);
//第三行
$excel->merge('B3:C3,E3:F3,G3:H3,I3:K3');
$excel->cell('A3')->value('车辆品牌：')->right()->border('*', $border_color, 'thin')->border('right,bottom', $border_color, 'hair');
$excel->cell('B3')->value('宝马')->left()->border('*', $border_color, 'thin', 'B3:C3')->border('left,bottom', $border_color, 'hair', 'B3:C3');
$excel->cell('D3')->value('车辆型号：')->right()->border('*', $border_color, 'thin')->border('right,bottom', $border_color, 'hair');
$excel->cell('E3')->value('X5')->left()->border('*', $border_color, 'thin', 'E3:F3')->border('left,bottom', $border_color, 'hair', 'E3:F3');
$excel->cell('G3')->value('车主姓名：')->right()->border('*', $border_color, 'thin')->border('right,bottom', $border_color, 'hair', 'G3:H3');
$excel->cell('I3')->value('张三')->left()->border('*', $border_color, 'thin', 'I3:K3')->border('left,bottom', $border_color, 'hair', 'I3:K3');
//第四行
$excel->merge('B4:C4,E4:F4,G4:H4,I4:K4');
$excel->cell('A4')->value('颜色：')->right()->border('*', $border_color, 'hair')->border('left', $border_color, '');
$excel->cell('B4')->value('黑')->left()->border('*', $border_color, 'hair', 'B4:C4')->border('right', $border_color, 'thin', 'B4:C4');
$excel->cell('D4')->value('年份：')->right()->border('*', $border_color, 'hair')->border('left', $border_color, 'thin');
$excel->cell('E4')->value('2009')->left()->border('*', $border_color, 'hair', 'E4:F4')->border('right', $border_color, 'thin', 'E4:F4');
$excel->cell('G4')->value('性别：')->right()->border('*', $border_color, 'hair', 'G4:H4')->border('left', $border_color, 'thin', 'G4:H4');
$excel->cell('I4')->value('男')->left()->border('*', $border_color, 'hair', 'I4:K4')->border('right', $border_color, 'thin', 'I4:K4');
//第五行
$excel->merge('B5:C5,E5:F5,G5:H5,I5:K5');
$excel->cell('A5')->value('车架号：')->right()->border('*', $border_color, 'hair')->border('left', $border_color, 'thin');
$excel->cell('B5')->value('ABCDEFGHIJKLMNOPQ')->left()->border('*', $border_color, 'hair', 'B5:C5')->border('right', $border_color, 'thin', 'B5:C5');
$excel->cell('D5')->value('里程：')->right()->border('*', $border_color, 'hair')->border('left', $border_color, 'thin');
$excel->cell('E5')->value('20155')->left()->border('*', $border_color, 'hair', 'E5:F5')->border('right', $border_color, 'thin', 'E5:F5');
$excel->cell('G5')->value('得知本厂的渠道：')->right()->border('*', $border_color, 'hair', 'G5:H5')->border('left', $border_color, 'thin', 'G5:H5');
$excel->cell('I5')->value('朋友推荐')->left()->border('*', $border_color, 'hair', 'I5:K5')->border('right', $border_color, 'thin', 'I5:K5');
//第六行
$excel->merge('B6:C6,E6:F6,G6:H6,I6:K6');
$excel->cell('A6')->value('车牌号：')->right()->border('*', $border_color, 'thin')->border('right,top', $border_color, 'hair');
$excel->cell('B6')->value('沪X0000')->left()->border('*', $border_color, '', 'B6:C6')->border('left,top', $border_color, 'hair', 'B6:C6');
$excel->cell('D6')->value('变速箱：')->right()->border('*', $border_color, 'thin')->border('right,top', $border_color, 'hair');
$excel->cell('E6')->value('自动')->left()->border('*', $border_color, '', 'E6:F6')->border('left,top', $border_color, 'hair', 'E6:F6');
$excel->cell('G6')->value('手机号：')->right()->border('*', $border_color, '', 'G6:H6')->border('right,top', $border_color, 'hair', 'G6:H6');
$excel->cell('I6')->value('18000000000')->left()->border('*', $border_color, '', 'I6:K6')->border('left,top', $border_color, 'hair', 'I6:K6');

/* 表体 */
//第七行
$excel->merge('A7:E7,F7:K7');
$excel->cell('A7')->value('服  务  清  单')->right()->bold()->size(11)->border('*', $border_color, '', 'A7:E7')->borderNone('right');
$excel->cell('E7')->borderNone('right');
$excel->cell('F7')->value('(维修项目享受三个月五千公里质保)')->right()->bottom()->size(7)->bold()->border('*', $border_color, '', 'F7:K7')->borderNone('left');
$excel->cell('F7')->borderNone('left');
//以下因项目行数不定，故行号以变量标识
$sheet1_global_row = 8;
//////////////////////////////////////////////////项目块1  -  start
$excel->merge('A' . $sheet1_global_row . ':F' . $sheet1_global_row . ',G' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('建议星级：★★★ ')->right()->bold()->border('*', $border_color, 'thin', 'A' . $sheet1_global_row . ':F' . $sheet1_global_row)->borderNone('right');
$excel->cell('F' . $sheet1_global_row)->borderNone('right');
$excel->cell('G' . $sheet1_global_row)->value('建议立即维护，否则对日常用车有严重影响！')->right()->bottom()->size(7)->bold()->border('*', $border_color, 'thin', 'G' . $sheet1_global_row . ':K' . $sheet1_global_row)->borderNone('left');
$excel->cell('G' . $sheet1_global_row)->borderNone('left');
$sheet1_global_row++;
//块表头
$excel->merge('B' . $sheet1_global_row . ':C' . $sheet1_global_row . ',E' . $sheet1_global_row . ':F' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('No.')->center()->border('*', $border_color, 'thin')->border('right', $border_color, 'hair');
$excel->cell('B' . $sheet1_global_row)->value('项目')->center()->border('*', $border_color, 'thin', 'B' . $sheet1_global_row . ':C' . $sheet1_global_row)->border('left,right', $border_color, 'hair', 'B' . $sheet1_global_row . ',C' . $sheet1_global_row);
$excel->cell('D' . $sheet1_global_row)->value('物料')->center()->border('*', $border_color, 'thin')->border('left', $border_color, 'hair')->border('right', $border_color, 'hair');
$excel->cell('E' . $sheet1_global_row)->value('单价')->center()->border('*', $border_color, 'thin', 'E' . $sheet1_global_row . ':F' . $sheet1_global_row)->border('left,right', $border_color, 'hair', 'E' . $sheet1_global_row . ',F' . $sheet1_global_row);
$excel->cell('G' . $sheet1_global_row)->value('数量')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('H' . $sheet1_global_row)->value('工时费')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('I' . $sheet1_global_row)->value('小计')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('J' . $sheet1_global_row)->value('备注')->center()->border('*', $border_color, 'thin', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left', $border_color, 'hair', 'J' . $sheet1_global_row);
$sheet1_global_row++;

//因用到求各公式，记录起始行和结束行
$sum_1_start_row = $sheet1_global_row;
$sum_1_end_row = $sheet1_global_row;
//测试数据填充
$text_arr = range(1, rand(10, 15));
foreach ($text_arr as $i) {
    $excel->merge('B' . $sheet1_global_row . ':C' . $sheet1_global_row . ',E' . $sheet1_global_row . ':F' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
    $excel->cell('A' . $sheet1_global_row)->value($i)->center()->border('*', $border_color, 'hair')->border('left', $border_color, 'thin');
    $excel->cell('B' . $sheet1_global_row)->value('项目名称 1-' . $i)->left()->border('*', $border_color, 'hair', 'B' . $sheet1_global_row . ':C' . $sheet1_global_row);
    $excel->cell('D' . $sheet1_global_row)->value('物料 1-' . $i)->left()->border('*', $border_color, 'hair');
    $excel->cell('E' . $sheet1_global_row)->value(rand(10, 100), 'n')->right()->border('*', $border_color, 'hair', 'E' . $sheet1_global_row . ':F' . $sheet1_global_row);
    $excel->cell('G' . $sheet1_global_row)->value(rand(1, 10), 'n')->right()->border('*', $border_color, 'hair');
    $excel->cell('H' . $sheet1_global_row)->value(rand(100, 200), 'n')->right()->border('*', $border_color, 'hair');
    $excel->cell('I' . $sheet1_global_row)->value('=E' . $sheet1_global_row . '*G' . $sheet1_global_row . '+H' . $sheet1_global_row, 'f')->right()->border('*', $border_color, 'hair');
    $excel->cell('J' . $sheet1_global_row)->value('备注 1-' . $i)->left()->border('*', $border_color, 'hair', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('right', $border_color, 'thin', 'K' . $sheet1_global_row);
    $sum_1_end_row = $sheet1_global_row;
    $sheet1_global_row++;
}
$excel->merge('A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('本项合计：')->right()->bold()->border('*', $border_color, 'thin', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('top,right', $border_color, 'hair', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row);
$excel->cell('I' . $sheet1_global_row)->value('=SUM(I' . $sum_1_start_row . ':I' . $sum_1_end_row . ')', 'f')->right()->border('*', $border_color, 'hair', 'I' . $sheet1_global_row)->border('bottom', $border_color, 'thin');
$excel->cell('J' . $sheet1_global_row)->border('*', $border_color, 'thin', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left,top', $border_color, 'hair', 'J' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$sheet1_global_row++;
///////////////////////////////////////////////////////////项目块1  -  end
//////////////////////////////////////////////////项目块2  -  start
$excel->merge('A' . $sheet1_global_row . ':F' . $sheet1_global_row . ',G' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('建议星级：★★ ')->right()->bold()->border('*', $border_color, '', 'A' . $sheet1_global_row . ':F' . $sheet1_global_row)->borderNone('right');
$excel->cell('F' . $sheet1_global_row)->borderNone('right');
$excel->cell('G' . $sheet1_global_row)->value('在建议里程数内维护，否则对日常用车有严重影响！')->right()->bottom()->size(7)->bold()->border('*', $border_color, '', 'G' . $sheet1_global_row . ':K' . $sheet1_global_row)->borderNone('left');
$excel->cell('G' . $sheet1_global_row)->borderNone('left');
$sheet1_global_row++;
//块表头
$excel->merge('B' . $sheet1_global_row . ':C' . $sheet1_global_row . ',E' . $sheet1_global_row . ':F' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('No.')->center()->border('*', $border_color, 'thin')->border('right', $border_color, 'hair');
$excel->cell('B' . $sheet1_global_row)->value('项目')->center()->border('*', $border_color, '', 'B' . $sheet1_global_row . ':C' . $sheet1_global_row)->border('left,right', $border_color, 'hair', 'B' . $sheet1_global_row . ',C' . $sheet1_global_row);
$excel->cell('D' . $sheet1_global_row)->value('物料')->center()->border('*', $border_color, 'thin')->border('left', $border_color, 'hair')->border('right', $border_color, 'hair');
$excel->cell('E' . $sheet1_global_row)->value('单价')->center()->border('*', $border_color, '', 'E' . $sheet1_global_row . ':F' . $sheet1_global_row)->border('left,right', $border_color, 'hair', 'E' . $sheet1_global_row . ',F' . $sheet1_global_row);
$excel->cell('G' . $sheet1_global_row)->value('数量')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('H' . $sheet1_global_row)->value('工时费')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('I' . $sheet1_global_row)->value('小计')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('J' . $sheet1_global_row)->value('备注')->center()->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left', $border_color, 'hair', 'J' . $sheet1_global_row);
$sheet1_global_row++;
//因用到求各公式，记录起始行和结束行
$sum_2_start_row = $sheet1_global_row;
$sum_2_end_row = $sheet1_global_row;
//测试数据填充
$text_arr = range(1, rand(5, 15));
foreach ($text_arr as $i) {
    $excel->merge('B' . $sheet1_global_row . ':C' . $sheet1_global_row . ',E' . $sheet1_global_row . ':F' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
    $excel->cell('A' . $sheet1_global_row)->value($i)->center()->border('*', $border_color, 'hair')->border('left', $border_color, 'thin');
    $excel->cell('B' . $sheet1_global_row)->value('项目名称 2-' . $i)->left()->border('*', $border_color, 'hair', 'B' . $sheet1_global_row . ':C' . $sheet1_global_row);
    $excel->cell('D' . $sheet1_global_row)->value('物料 2-' . $i)->left()->border('*', $border_color, 'hair');
    $excel->cell('E' . $sheet1_global_row)->value(rand(10, 100), 'n')->right()->border('*', $border_color, 'hair', 'E' . $sheet1_global_row . ':F' . $sheet1_global_row);
    $excel->cell('G' . $sheet1_global_row)->value(rand(1, 10), 'n')->right()->border('*', $border_color, 'hair');
    $excel->cell('H' . $sheet1_global_row)->value(rand(100, 200), 'n')->right()->border('*', $border_color, 'hair');
    $excel->cell('I' . $sheet1_global_row)->value('=E' . $sheet1_global_row . '*G' . $sheet1_global_row . '+H' . $sheet1_global_row, 'f')->right()->border('*', $border_color, 'hair');
    $excel->cell('J' . $sheet1_global_row)->value('备注 2-' . $i)->left()->border('*', $border_color, 'hair', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('right', $border_color, 'thin', 'K' . $sheet1_global_row);
    $sum_2_end_row = $sheet1_global_row;
    $sheet1_global_row++;
}
$excel->merge('A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('本项合计：')->right()->bold()->border('*', $border_color, '', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('top,right', $border_color, 'hair', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',H' . $sheet1_global_row);
$excel->cell('I' . $sheet1_global_row)->value('=SUM(I' . $sum_2_start_row . ':I' . $sum_2_end_row . ')', 'f')->right()->border('*', $border_color, 'hair', 'I' . $sheet1_global_row)->border('bottom', $border_color, '');
$excel->cell('J' . $sheet1_global_row)->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left,top', $border_color, 'hair', 'J' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$sheet1_global_row++;
///////////////////////////////////////////////////////////项目块2  -  end
//////////////////////////////////////////////////项目块3  -  start
$excel->merge('A' . $sheet1_global_row . ':F' . $sheet1_global_row . ',G' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('建议星级：★ ')->right()->bold()->border('*', $border_color, '', 'A' . $sheet1_global_row . ':F' . $sheet1_global_row)->borderNone('right');
$excel->cell('F' . $sheet1_global_row)->borderNone('right');
$excel->cell('G' . $sheet1_global_row)->value('以下故障对日常用车有一定影响，但可不维护！')->right()->bottom()->size(8)->bold()->border('*', $border_color, '', 'G' . $sheet1_global_row . ':K' . $sheet1_global_row)->borderNone('left');
$excel->cell('G' . $sheet1_global_row)->borderNone('left');
$sheet1_global_row++;
//块表头
$excel->merge('B' . $sheet1_global_row . ':C' . $sheet1_global_row . ',E' . $sheet1_global_row . ':F' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('No.')->center()->border('*', $border_color, 'thin')->border('right', $border_color, 'hair');
$excel->cell('B' . $sheet1_global_row)->value('项目')->center()->border('*', $border_color, '', 'B' . $sheet1_global_row . ':C' . $sheet1_global_row)->border('left,right', $border_color, 'hair', 'B' . $sheet1_global_row . ',C' . $sheet1_global_row);
$excel->cell('D' . $sheet1_global_row)->value('物料')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('E' . $sheet1_global_row)->value('单价')->center()->border('*', $border_color, '', 'E' . $sheet1_global_row . ':F' . $sheet1_global_row)->border('left,right', $border_color, 'hair', 'E' . $sheet1_global_row . ',F' . $sheet1_global_row);
$excel->cell('G' . $sheet1_global_row)->value('数量')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('H' . $sheet1_global_row)->value('工时费')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('I' . $sheet1_global_row)->value('小计')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('J' . $sheet1_global_row)->value('备注')->center()->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left', $border_color, 'hair', 'J' . $sheet1_global_row);
$sheet1_global_row++;
//因用到求各公式，记录起始行和结束行
$sum_3_start_row = $sheet1_global_row;
$sum_3_end_row = $sheet1_global_row;
//测试数据填充
$text_arr = range(1, rand(15, 30));
foreach ($text_arr as $i) {
    $excel->merge('B' . $sheet1_global_row . ':C' . $sheet1_global_row . ',E' . $sheet1_global_row . ':F' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
    $excel->cell('A' . $sheet1_global_row)->value($i)->center()->border('*', $border_color, 'hair')->border('left', $border_color, 'thin');
    $excel->cell('B' . $sheet1_global_row)->value('项目名称 3-' . $i)->left()->border('*', $border_color, 'hair', 'B' . $sheet1_global_row . ':C' . $sheet1_global_row);
    $excel->cell('D' . $sheet1_global_row)->value('物料 3-' . $i)->left()->border('*', $border_color, 'hair');
    $excel->cell('E' . $sheet1_global_row)->value(rand(10, 100), 'n')->right()->border('*', $border_color, 'hair', 'E' . $sheet1_global_row . ':F' . $sheet1_global_row);
    $excel->cell('G' . $sheet1_global_row)->value(rand(1, 10), 'n')->right()->border('*', $border_color, 'hair');
    $excel->cell('H' . $sheet1_global_row)->value(rand(100, 200), 'n')->right()->border('*', $border_color, 'hair');
    $excel->cell('I' . $sheet1_global_row)->value('=E' . $sheet1_global_row . '*G' . $sheet1_global_row . '+H' . $sheet1_global_row, 'f')->right()->border('*', $border_color, 'hair');
    $excel->cell('J' . $sheet1_global_row)->value('备注 2-' . $i)->left()->border('*', $border_color, 'hair', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('right', $border_color, 'thin', 'K' . $sheet1_global_row);
    $sum_3_end_row = $sheet1_global_row;
    $sheet1_global_row++;
}
$excel->merge('A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('本项合计：')->right()->bold()->border('*', $border_color, '', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('top,right', $border_color, 'hair', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',H' . $sheet1_global_row);
$excel->cell('I' . $sheet1_global_row)->value('=SUM(I' . $sum_3_start_row . ':I' . $sum_3_end_row . ')', 'f')->right()->border('*', $border_color, 'hair', 'I' . $sheet1_global_row)->border('bottom', $border_color, '');
$excel->cell('J' . $sheet1_global_row)->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left,top', $border_color, 'hair', 'J' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$sheet1_global_row++;
///////////////////////////////////////////////////////////项目块3  -  end
$service_charge_sum_cells = implode(',', ['I' . ($sum_1_end_row + 1), 'I' . ($sum_2_end_row + 1), 'I' . ($sum_3_end_row + 1)]);//三个小项目块费用合计单元格
$excel->merge('A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('服务费用合计：')->right()->bold()->border('*', $border_color, 'thin', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('top,right', $border_color, 'hair', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',H' . $sheet1_global_row);
$excel->cell('I' . $sheet1_global_row)->value('=SUM(' . $service_charge_sum_cells . ')', 'f')->right()->border('*', $border_color, 'hair', 'I' . $sheet1_global_row)->border('top,bottom', $border_color, 'thin');
$excel->cell('J' . $sheet1_global_row)->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left,top', $border_color, 'hair', 'J' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$service_charge_sum_cell = 'I' . $sheet1_global_row;//服务清单大项目块费用合计，后面用到所有费用合计
$sheet1_global_row++;
/////////////////////////服务清单项目块结束


$excel->merge('A' . $sheet1_global_row . ':E' . $sheet1_global_row . ',F' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('其  他  费  用')->right()->bold()->size(11)->border('*', $border_color, '', 'A' . $sheet1_global_row . ':E' . $sheet1_global_row)->borderNone('right');
$excel->cell('E' . $sheet1_global_row)->borderNone('right');
$excel->cell('F' . $sheet1_global_row)->value('维修客户以下空白')->right()->bottom()->size(8)->bold()->border('*', $border_color, '', 'F' . $sheet1_global_row . ':K' . $sheet1_global_row)->borderNone('left');
$excel->cell('F' . $sheet1_global_row)->borderNone('left');
$sheet1_global_row++;

//////////////////////////////////////////////////其他费用项目块  -  start
//块表头
$excel->merge('B' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('No.')->center()->border('*', $border_color, 'thin')->border('right', $border_color, 'hair');
$excel->cell('B' . $sheet1_global_row)->value('项目明细')->center()->border('*', $border_color, '', 'B' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('left,right', $border_color, 'hair', 'B' . $sheet1_global_row . ',H' . $sheet1_global_row);
$excel->cell('I' . $sheet1_global_row)->value('金额')->center()->border('*', $border_color, 'thin')->border('left,right', $border_color, 'hair');
$excel->cell('J' . $sheet1_global_row)->value('备注')->center()->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left', $border_color, 'hair', 'J' . $sheet1_global_row);
$sheet1_global_row++;
//因用到求各公式，记录起始行和结束行
$sum_4_start_row = $sheet1_global_row;
$sum_4_end_row = $sheet1_global_row;
//测试数据填充
$text_arr = range(1, rand(10, 20));
foreach ($text_arr as $i) {
    $excel->merge('B' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
    $excel->cell('A' . $sheet1_global_row)->value($i)->center()->border('*', $border_color, 'hair')->border('left', $border_color, '');
    $excel->cell('B' . $sheet1_global_row)->value('其他项目名称' . $i)->left()->border('*', $border_color, 'hair', 'B' . $sheet1_global_row . ':H' . $sheet1_global_row);
    $excel->cell('I' . $sheet1_global_row)->value(rand(100, 200), 'n')->right()->border('*', $border_color, 'hair');
    $excel->cell('J' . $sheet1_global_row)->value('备注' . $i)->left()->border('*', $border_color, 'hair', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('right', $border_color, '', 'K' . $sheet1_global_row);
    $sum_4_end_row = $sheet1_global_row;
    $sheet1_global_row++;
}
$excel->merge('A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('其他费用合计：')->right()->bold()->border('*', $border_color, 'thin', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('top,right', $border_color, 'hair', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',H' . $sheet1_global_row);
$excel->cell('I' . $sheet1_global_row)->value('=SUM(I' . $sum_4_start_row . ':I' . $sum_4_end_row . ')', 'f')->right()->border('*', $border_color, 'hair', 'I' . $sheet1_global_row)->border('bottom', $border_color, 'thin');
$excel->cell('J' . $sheet1_global_row)->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left,top', $border_color, 'hair', 'J' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$others_charge_sum_cell = 'I' . $sheet1_global_row;//其他大项目块费用合计，后面用到所有费用合计
$sheet1_global_row++;
///////////////////////////////////////////////////////////其他费用项目块  -  end
$excel->merge('A' . $sheet1_global_row . ':H' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('费用合计（不含税）：')->right()->bold()->border('*', $border_color, 'thin', 'A' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('right', $border_color, 'hair', 'H' . $sheet1_global_row);
$excel->cell('I' . $sheet1_global_row)->value('=SUM(' . $service_charge_sum_cells . ',' . $others_charge_sum_cell . ')', 'f')->right()->border('*', $border_color, 'hair', 'I' . $sheet1_global_row)->border('top,bottom', $border_color, 'thin');
$excel->cell('J' . $sheet1_global_row)->border('*', $border_color, '', 'J' . $sheet1_global_row . ':K' . $sheet1_global_row)->border('left,top', $border_color, 'hair', 'J' . $sheet1_global_row . ',J' . $sheet1_global_row . ':K' . $sheet1_global_row);
$sheet1_global_row++;
///////////////////////////  页底
$excel->merge('A' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('注：如服务项目中有备注为预估价格，最后价格可能与所做的估价有15%的浮动。')->left()->bold()->border('*', $border_color, '', 'A' . $sheet1_global_row . ':K' . $sheet1_global_row);
$sheet1_global_row++;
$excel->setRowHight($sheet1_global_row, 33.9);
$excel->merge('A' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('我方仅针对以上所列项目及故障进行检查维修，未列明项目及客户未表述项目不在我方检查范围内，此类项目产生故障或由于故障造成的间接结果，我方不承担责任')->wrap()->left()->bold()->border('*', $border_color, 'thin', 'A' . $sheet1_global_row . ':K' . $sheet1_global_row);
$sheet1_global_row++;
$excel->merge('A' . $sheet1_global_row . ':H' . $sheet1_global_row);
$excel->setRowHight($sheet1_global_row, 13);
$excel->cell('A' . $sheet1_global_row)->value('by 天下第一修车厂')->italic()->size(7)->border('left', $border_color, 'thin')->border('bottom', $border_color, 'thin', 'A' . $sheet1_global_row . ':B' . $sheet1_global_row);
////跨行的，在上面一行先写
$excel->merge('I' . $sheet1_global_row . ':K' . ($sheet1_global_row + 1));
$excel->cell('I' . $sheet1_global_row)->value("官方网站：www.daijulong.com\n微信公众平台：daijulong-com\n邮箱：daijulong@gmail.com\n新浪认证微博：weibo.com/daijulong")->border('*', $border_color, 'thin', 'I' . $sheet1_global_row . ':K' . ($sheet1_global_row + 1))->borderNone('left', 'I' . $sheet1_global_row . ':K' . ($sheet1_global_row + 1));
$excel->cell('I' . $sheet1_global_row)->size(7)->wrap();
$sheet1_global_row++;
//跨行下一行行高
$excel->setRowHight($sheet1_global_row, 50);
$excel->merge('C' . $sheet1_global_row . ':D' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value('客户签字：')->left()->top()->border('*', $border_color, 'thin')->borderNone('right');
$excel->cell('C' . $sheet1_global_row)->value('维修顾问签字：')->center()->top()->borderNone('*', 'C' . $sheet1_global_row . ':D' . $sheet1_global_row)->border('bottom', $border_color, '', 'C' . $sheet1_global_row . ':D' . $sheet1_global_row);
$excel->cell('F' . $sheet1_global_row)->value('日期：')->left()->top()->borderNone('*')->border('bottom', $border_color, 'thin');
//剩余几个单元格边框
$excel->cell('B' . $sheet1_global_row)->border('*', $border_color, '')->borderNone('left,right');
$excel->cell('E' . $sheet1_global_row)->borderNone('*')->border('bottom', $border_color, 'thin');
$excel->cell('G' . $sheet1_global_row)->borderNone('*', 'G' . $sheet1_global_row . ':H' . $sheet1_global_row)->border('bottom', $border_color, 'thin', 'H' . $sheet1_global_row);
$sheet1_global_row++;
$excel->merge('A' . $sheet1_global_row . ':K' . $sheet1_global_row);
$excel->cell('A' . $sheet1_global_row)->value("第一联： 财务联（白）                                    第二联：维修顾问联（粉）                                    第三联 留存联（黄）")->center()->centerV()->border('*', $border_color, '', 'A' . $sheet1_global_row . ':K' . $sheet1_global_row);

$sheet1_global_row++;

/*
 * END 绘制第一个工作表
 */



/*
 * 绘制第二个工作表
 */
$excel->createSheet('其他演示', [], [], '0000FF');
$excel->setColsWidth(['C' => 12]);
$sheet2_global_row = 1;
$excel->merge('A' . $sheet2_global_row . ':J' . $sheet2_global_row);
$excel->cell('A' . $sheet2_global_row)->value('简单表格绘制演示-成绩表')->size(12)->center()->bold();
$sheet2_global_row++;
$header_data = [
    'A' => ['序号', 'str'],
    'B' => ['姓名', 'str', ['font' => ['color' => ['rgb' => '000000']]]],
    'C' => '班级',
    'D' => '数学',
    'E' => '语文',
    'F' => '英语',
    'G' => '物理',
    'H' => '化学',
    'I' => '总分',
    'J' => '平均分',
];
$default_style = [
    'font' => [
        'bold' => true,
        'color' => ['rgb' => 'FF0000'],
    ],
    'borders' => [
        'allborders' => [
            'style' => 'thin',
            'color' => [
                'rgb' => '808080'
            ]
        ],
    ],
    'align' => [
        'horizontal' => 'center',
    ],
];

$excel->batchFillRow($header_data, $sheet2_global_row, $default_style);
$sheet2_global_row++;

//批量添加多行内容
//可以是数据库中查询得出的数据集，经过与列对应重新组织后的数据
$content_data = [
    ['#DROW#', '张三', '高二（1）班', rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), '=SUM(D#CROW#:H#CROW#)', '=AVERAGE(D#CROW#:H#CROW#)'],
    ['#DROW#', '李四', '高二（1）班', rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), '=SUM(D#CROW#:H#CROW#)', '=AVERAGE(D#CROW#:H#CROW#)'],
    ['#DROW#', '王五', '高二（1）班', rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), '=SUM(D#CROW#:H#CROW#)', '=AVERAGE(D#CROW#:H#CROW#)'],
    ['#DROW#', '赵六', '高二（1）班', rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), rand(60, 99), '=SUM(D#CROW#:H#CROW#)', '=AVERAGE(D#CROW#:H#CROW#)'],
];
$content_data_type = ['D' => 'n', 'E' => 'n', 'F' => 'n', 'G' => 'n', 'H' => 'n', 'I' => 'f', 'J' => 'f'];
$content_style = [
    'B' => ['align' => ['horizontal' => 'left'], 'font' => ['bold' => true]],
    'C' => ['align' => ['horizontal' => 'left'],],
];
$content_default_style = [
    'borders' => [
        'allborders' => [
            'style' => 'hair',
            'color' => [
                'rgb' => 'FF0000'
            ]
        ],
    ],
    'align' => [
        'horizontal' => 'right',
    ],
];
$start_row = $sheet2_global_row;
$end_row = $excel->batchFillRows($content_data, 'A', $start_row, $content_data_type, $content_style, $content_default_style);
$sheet2_global_row = $end_row + 1;
//成绩主体外框线画为实线
$excel->setCellsBorder('A' . $start_row . ':J' . $start_row, 'top', '808080', 'thin');
$excel->setCellsBorder('A' . $start_row . ':A' . $end_row, 'left', '808080', 'thin');
$excel->setCellsBorder('J' . $start_row . ':J' . $end_row, 'right', '808080', 'thin');
$excel->setCellsBorder('A' . $end_row . ':J' . $end_row, 'bottom', '808080', 'thin');
//插入图片演示
$excel->setImage('C' . ($sheet2_global_row + 3), './demo-image.jpg', 50, 50, 10, 10, ['name' => 'img.jpg', 'description' => 'GitHub截图']);
//注释
$excel->cell('E' . ($sheet2_global_row + 1))->setValue('带注释')->comment("这个是注释的内容\r\n很不错吧？");
/*
 * END 绘制第二个工作表
 */


//将第一个工作表切换为活动工作表，否则下载的文件打开后将显示最后一个工作表
$excel->switchActiveSheet(0);

//保存文件
//$excel->save('./demo-' . date('YmdHis') . '.xlsx', 'xlsx');
//直接下载
$excel->download('demo-' . date('YmdHis'), 'xlsx');