<?php

/**

 * 根据数据绘制Excel(可以同时在一张Excel中绘制：饼图、柱状图、折线图、列表)
 *
 * 支持指定导出的文件名
 *
 * 支持渲染excel中公司logo
 *
 * 支持依次渲染多个饼图、柱状图、折线图、列表
 *
 * 饼图、柱状图、折线图 支持多列数据
 *
 * 列表支持指定每行每列数据值、文字颜色、背景颜色
 *
 * @author Winston LEE <winston.lee@esghl.com>
 *
 * Class ExcelService
 *
 * 请求示例：
 *
 * @file_name 文件名
 *
 * @title 大标题
 *
 * @logo 公司logo路径
 *
 * @data 数据
 *
 * @data.type 1饼图 2柱状图 3折线图 4列表
 *
 * @data.title 数据标题
 *
 * @data.list 数据源
 *
 * $data = [
 *      "file_name" => "测试文件",
 *      "title" => "测试标题",
 *      "logo" => "",
 *      "data" => [
 *          [
 *              "type" => 4,
 *              "mainTitle" => '子标题',
 *              "title" => ['姓名', '年龄', '性别'],
 *              "list" =>
 *              [
 *                  [
 *                      ['value' => '张三', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00'],
 *                      ['value' => '25', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00'],
 *                      ['value' => '男', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00']
 *                  ],
 *                  [
 *                      ['value' => '李四', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0'],
 *                      ['value' => '30', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0'],
 *                      ['value' => '女', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0']
 *                  ],
 *                  [
 *                      ['value' => '王五', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00'],
 *                      ['value' => '28', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00'],
 *                      ['value' => '男', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00']
 *                  ],
 *              ],
 *          ],
 *          [
 *              'title'=>'测试1',
 *              'type'=>'1',
 *              'list'=>[['省份','总数'],['Q1',0.1],['Q2',0.2],['Q3',0.3],['Q4',0.4]]
 *          ],
 *          [
 *               'title'=>'测试2',
 *               'type'=>'2',
 *               'list'=>[['省份','总数'],['Q5',12],['Q6',56],['Q7',52],['Q8',30]]
 *          ],
 *          [
 *              'title'=>'测试3',
 *              'type'=>'3',
 *              'list'=>[['省份','总数'],['Q9',12],['Q10',56],['Q11',52],['Q12',30]]
 *          ],
 *          [
 *              "type" => 4,
 *              "title" => ['姓名1', '年龄2', '性别3'],
 *              "list" =>
 *               [
 *                  [
 *                      ['value' => '张三', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00'],
 *                      ['value' => '25', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00'],
 *                      ['value' => '男', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00']
 *                  ],
 *                  [
 *                      ['value' => '李四', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0'],
 *                      ['value' => '30', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0'],
 *                      ['value' => '女', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0']
 *                  ],
 *                  [
 *                      ['value' => '王五', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00'],
 *                      ['value' => '28', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00'],
 *                      ['value' => '男', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00']
 *                  ],
 *              ],
 *          ],
 *      ],
 * ];
 * $excelExporter = new ExcelService();
 * return $excelExporter->excel($data);
 *
 */


namespace Modules\Common\Services;


use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\Layout;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Chart\Axis;
use PhpOffice\PhpSpreadsheet\Chart\GridLines;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpWord\Collection\Charts;


class ExcelService
{

    private $spreadsheet; //phpSpreadsheet实例
    private $worksheet;   //数据单元表
    private $currentRow;  //当前行
    private $fileName;    //文件名
    private $logoPath;    //logo地址
    private $mainTitle;   //大标题
    private $includeCharts;   //作图开关
    private $chartStartColumn;   //图表原始数据渲染起始列
    private $chartStartColumnHidden;   //图表原始数据隐藏标识

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->worksheet = $this->spreadsheet->getActiveSheet();
        $this->currentRow = 6;
        $this->includeCharts = false;
        $this->chartStartColumn  ='AA';
        $this->chartStartColumnHidden = false;
        $this->setLogo();
    }

    /**
     * 计算列在Excel中的值
     * @param $index
     * @return string
     */
    public function coordinateFromIndex($index) {
        $column = '';
        while ($index >= 0) {
            $column = chr(ord('A') + ($index % 26)) . $column;
            $index = (int)($index / 26) - 1;
        }

        return $column;
    }

    /**
     * 把列值计算为数字
     * @param $columnName
     * @return float|int
     */
    public function columnNameToNumber($columnName) {
        $columnNumber = 0;
        $length = strlen($columnName);

        for ($i = 0; $i < $length; $i++) {
            $columnNumber = $columnNumber * 26 + (ord($columnName[$i]) - ord('A') + 1);
        }

        return $columnNumber;
    }

    /**
     * 把数字计算为列值
     * @param $columnName
     * @return float|int
     */
    public function columnNumberToName($columnNumber) {
        $columnName = '';

        while ($columnNumber > 0) {
            $columnNumber--;
            $columnName = chr($columnNumber % 26 + ord('A')) . $columnName;
            $columnNumber = floor($columnNumber / 26);
        }

        return $columnName;
    }

    /**
     * 根据开始列值计算累加$length列值
     * @param $startCell
     * @param $length
     * @return string
     */
    public function calculateEndpoint($startCell, $length) {
        if(empty($length)){
            $this->chartStartColumn = $startCell;
            return $startCell;
        }
        // 获取起始单元格的列名
        $startColumn = preg_replace('/\d/', '', $startCell);

        // 将列名转换为数字
        $startColumnNumber = $this->columnNameToNumber($startColumn);

        // 计算结束列的数字
//        $endColumnNumber = $startColumnNumber + $length - 1;
        $endColumnNumber = $startColumnNumber + $length;

        // 将数字转换为列名
        $endColumn = $this->columnNumberToName($endColumnNumber);

        // 获取起始单元格的行号
        $startRow = preg_replace('/\D/', '', $startCell);

        // 组合终点单元格
        $endCell = $endColumn . $startRow;

        $this->chartStartColumn = $endCell;

        return $endCell;
    }

    /**
     * 根据开始列值计算累加26列值
     * @param $startCell
     * @param $length
     * @return string
     */
    function calculateEndpointByRows($startCell, $length) {
        if(empty($length)){
            $this->chartStartColumn = $startCell;
            return $startCell;
        }
        // 获取起始单元格的列名
        $startColumn = preg_replace('/\d/', '', $startCell);

        // 将列名转换为数字
        $startColumnNumber = $this->columnNameToNumber($startColumn);

        // 计算结束列的数字
        $endColumnNumber = $startColumnNumber + 26 * $length;

        // 将数字转换为列名
        $endColumn = $this->columnNumberToName($endColumnNumber);

        // 获取起始单元格的行号
        $startRow = preg_replace('/\D/', '', $startCell);

        // 组合终点单元格
        $endCell = $endColumn . $startRow;

        $this->chartStartColumn = $endCell;

        return $endCell;
    }

    /**
     * 设定导出的文件名
     * @param $fileName
     */
    public function setFileName($fileName)
    {
        $this->fileName = $fileName;
    }

    /**
     * 设定Excel的logo
     * @param $logoPath
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setLogo($logoPath = '')
    {
        $this->logoPath = $logoPath;
        if(empty($this->logoPath)){
            $this->logoPath = resource_path('image/logo/ieasyfm.png');
        }
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setDescription('Logo');
        $drawing->setPath($this->logoPath);
        $drawing->setHeight(30);
        $drawing->setCoordinates('A1');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($this->worksheet);
        $this->worksheet->mergeCells('A1:A2');
    }

    /**
     * 设定Excel的大标题
     * @param $mainTitle
     * @param $data
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setMainTitle($mainTitle, $data)
    {
        $this->mainTitle = $mainTitle;
        if(empty($this->mainTitle)){
            $this->mainTitle = $this->fileName;
        }
        $colCount = (isset($data[0]) && count($data[0]) <= 0)? 10 : count($data[0]);
        $this->worksheet->setCellValue('A4', $this->mainTitle);
        $this->worksheet->mergeCells("A4:" . Coordinate::stringFromColumnIndex($colCount) . "4");
        $this->worksheet->getStyle('A4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    }

    /**
     * 设定列表数据
     * @param $data
     */
    public function setList($data)
    {
        $dataMainTitle = getter($data,'mainTitle');
        $dataTitles = getter($data,'title');
        $dataSource = getter($data,'list');
        if(!empty($dataMainTitle)){
            $this->setListMainTitle($dataMainTitle,$dataSource);
        }
        if(!empty($dataTitles)){
            $this->setListTitle($dataTitles);
        }
        if(!empty($dataSource)){
            $this->setListData($dataSource);
        }
    }

    /**
     * 设定列表附属大标题
     * @param $dataMainTitle
     * @param $dataSource
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setListMainTitle($dataMainTitle,$dataSource)
    {
        $column = 'A';
        $colCount = (isset($dataSource[0]) && count($dataSource[0]) <=0)?10:count($dataSource[0]);
        $this->worksheet->setCellValue($column . $this->currentRow, $dataMainTitle);
        $this->worksheet->mergeCells($column . $this->currentRow.":" . Coordinate::stringFromColumnIndex($colCount) . $this->currentRow);
        $this->worksheet->getStyle($column . $this->currentRow)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $this->currentRow ++;
    }

    /**
     * 设定列表数据标题
     * @param $dataTitles
     */
    public function setListTitle($dataTitles)
    {
        $column = 'A';
        if(!empty($dataTitles)){
            foreach ($dataTitles as $title) {
                $cellLength = 15;
                $this->worksheet->getColumnDimension("$column")->setWidth("$cellLength");
                $this->worksheet->setCellValue($column . $this->currentRow, $title);
                $column++;
            }
        }

        // 设置数据标题行的居中对齐
        $lastDataTitleColumn = $this->coordinateFromIndex(count($dataTitles) - 1);
        $this->worksheet->getStyle("A".$this->currentRow.":{$lastDataTitleColumn}".$this->currentRow)
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
            ->setVertical(Alignment::VERTICAL_CENTER);
    }

    /**
     * 设定列表数据
     * @param $dataSource
     */
    public function setListData($dataSource)
    {
        $this->currentRow ++;
        if(!empty($dataSource)){
            foreach ($dataSource as $dataRow) {
                if(!empty($dataRow)){
                    $column = 'A';
                    foreach ($dataRow as $cellData) {
                        $this->worksheet->setCellValue($column . $this->currentRow, getter($cellData,'value'));

                        if (!empty(getter($cellData,'textColor'))) {
                            $this->worksheet->getStyle($column . $this->currentRow)->getFont()->getColor()->setARGB(getter($cellData,'textColor'));
                        }

                        if (!empty(getter($cellData,'backgroundColor'))) {
                            $this->worksheet->getStyle($column . $this->currentRow)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB(getter($cellData,'backgroundColor'));
                        }

                        // 设置单元格居中对齐
                        $this->worksheet->getStyle($column . $this->currentRow)
                            ->getAlignment()
                            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                            ->setVertical(Alignment::VERTICAL_CENTER);

                        $column++;
                    }
                    $this->currentRow++;
                }
            }
        }
    }

    /**
     * 设定图表数据
     * @param $dataSource
     */
    public function setChartData($dataSource)
    {
        //打开作图开关
        $this->includeCharts = true;
        //数据支撑
        $data = (isset($dataSource['list']) && !empty($dataSource['list'])) ? $dataSource['list'] : array(
            array('维度',	'A类型',	'B类型'),
            array('Q1',   12,   13),
            array('Q2',   56,   52),
            array('Q3',   52,   58),
            array('Q4',   30,   33),
        );

        //标题
        $title = (isset($dataSource['title']) && !empty($dataSource['title'])) ? $dataSource['title'] : '测试';

        //类型 1饼图2柱状图3折线图
        $type = (isset($dataSource['type']) && !empty($dataSource['type'])) ? $dataSource['type'] : '1';

        //颜色(目前phpspreadSheet版本暂不支持自定义图表颜色)
        $colors = (isset($dataSource['colors']) && !empty($dataSource['colors'])) ? $dataSource['colors'] :  ['#FF0000', '#00FF00']; // 指定颜色，例如：红色和绿色

        if($type == 1){
            //处理数据全部为0的情况
            $total = 0;
            foreach($data as $key => $item){
                if($key > 0){
                    $total += floatval($item[1]);
                }
            }

            //设定每个项目值相同 使饼图平均分配
            if($total == 0){
                foreach($data as $key => $item){
                    if($key > 0){
                        $data[$key][1] = 1;
                    }
                }
            }
        }

        //根据类型指明绘制图标对应的参数
        switch($type){
            //饼图
            case 1:
                $plotType = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::TYPE_PIECHART;
                $plotGrouping = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::GROUPING_PERCENT_STACKED;
                break;
            //柱状图
            case 2:
                $plotType = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::TYPE_BARCHART;
                $plotGrouping = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::GROUPING_CLUSTERED;
                break;
            //折线图
            case 3:
                $plotType = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::TYPE_LINECHART;
                $plotGrouping = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::GROUPING_STANDARD;
                break;
            //饼图
            default:
                $plotType = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::TYPE_PIECHART;
                $plotGrouping = \PhpOffice\PhpSpreadsheet\Chart\DataSeries::GROUPING_PERCENT_STACKED;
                break;
        }

        $this->currentRow++;

        $this->chartStartColumn =  $this->calculateEndpointByRows($this->chartStartColumn,1);

        $saveChartStartColumn = array($this->chartStartColumn);

        //设定数据
        $this->worksheet->fromArray(
            $data,'',$this->chartStartColumn.'1'
        );

        //设置图表比例对应的名称
        $xAxisTickValues1 = array(
            new \PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues('String', 'Worksheet!$'.$this->chartStartColumn.'$2:$'.$this->chartStartColumn.'$'.count($data), NULL, (count($data) - 1)),	    //4	 Q1 to Q4
        );

        //设置图表比例显示的来源
        $dataSeriesLabels1 = array();
        // 设置作图区域数据
        $dataSeriesValues1 = array();
        //循环渲染多列数据
        foreach(getter($data,0) as $key => $item){
            if($key > 0){
                $this->chartStartColumn = $this->calculateEndpoint($this->chartStartColumn,1);
                array_push($saveChartStartColumn,$this->chartStartColumn);
                //设置图表比例显示的来源
                $dataSeriesLabels1[] = new \PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues('String', 'Worksheet!$'.$this->chartStartColumn.'$1', NULL, 1);
                // 设置作图区域数据
                $dataSeriesValues1[] = new \PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues('Number', 'Worksheet!$'.$this->chartStartColumn.'$2:$'.$this->chartStartColumn.'$'.count($data), NULL, (count($data) - 1)); //4	 Q1 to Q4
            }
        }

        // 构建数据
        $series1 = new \PhpOffice\PhpSpreadsheet\Chart\DataSeries(
        //这里设置生成图表类型（折线图，柱状图，饼状图…）
            $plotType,				   // plotType
            //设置当前图表展示类型（堆积，非堆积，百分比），根据不同图表设置不同。
            $plotGrouping, // plotGrouping (Pie charts don't have any grouping)
            //这里需要返回一个数组，有多少条数据需要多少长度的数组。
            range(0, count($dataSeriesValues1)-1),					// plotOrder
            //标签
            $dataSeriesLabels1,										// plotLabel
            //X轴数据
            $xAxisTickValues1,										// plotCategory
            //绘图所需数据
            $dataSeriesValues1,										// plotValues
        );

        //这里有个重要的配置，主要对图表柱状图绘图方式的设置。（横着展示还是竖着展示）
        if($type == 2){
            $series1->setPlotDirection(\PhpOffice\PhpSpreadsheet\Chart\DataSeries::DIRECTION_COL);
        }

        //折线图设定平滑曲线(无效)
        if($type == 3){
            $series1->setSmoothLine(true);
        }

        $layout1 = new \PhpOffice\PhpSpreadsheet\Chart\Layout();
        $layout1->setShowVal(TRUE); //设置是否显示原始值
        $layout1->setShowPercent(TRUE); //设置是否显示比例

        // 给数据系列分配一个做图区域
        $plotArea1 = new \PhpOffice\PhpSpreadsheet\Chart\PlotArea($layout1, array($series1));
        //设置图表图例
        $legend1 = new \PhpOffice\PhpSpreadsheet\Chart\Legend(\PhpOffice\PhpSpreadsheet\Chart\Legend::POSITION_RIGHT, NULL, false);
        // 设置图表标题
        $title1 = new \PhpOffice\PhpSpreadsheet\Chart\Title($title);
        //创建图形
        $chart1 = new \PhpOffice\PhpSpreadsheet\Chart\Chart(
            'chart',	// name
            $title1,		// title
            $legend1,		// legend
            $plotArea1,		// plotArea
            true,			// plotVisibleOnly
            0,				// displayBlanksAs
            NULL,			// xAxisLabel
            NULL			// yAxisLabel		- Pie charts don't have a Y-Axis
        );

        //设置图形绘制区域范围
        $chart1->setTopLeftPosition('A'.$this->currentRow);
        $chart1->setBottomRightPosition('J'. ($this->currentRow + 20));

        $this->currentRow += 22;

        //增加图表
        $this->worksheet->addChart($chart1);

        //隐藏原始数据(默认关闭)
        if(!empty($saveChartStartColumn) && $this->chartStartColumnHidden){
            foreach($saveChartStartColumn as $item)
            {
                $this->worksheet->getColumnDimension("$item")->setVisible(false);
            }
        }

    }


    /**
     * 渲染数据
     * @param $dataSources
     * [
     *      "file_name" => "测试文件",
     *      "title" => "测试标题",
     *      "logo" => "",
     *      "data" => [
     *            [
     *               "type" => 4,
     *                "title" => ['姓名', '年龄', '性别'],
     *                "list" =>
     * *                [
     *                      [
     *                          ['value' => '张三', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00'],
     *                          ['value' => '25', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00'],
     *                          ['value' => '男', 'textColor' => 'FF000000', 'backgroundColor' => 'FFFFFF00']
     *                      ],
     *                      [
     *                          ['value' => '李四', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0'],
     *                          ['value' => '30', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0'],
     *                          ['value' => '女', 'textColor' => 'FF000000', 'backgroundColor' => 'FFA0A0A0']
     *                      ],
     *                      [
     *                          ['value' => '王五', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00'],
     *                          ['value' => '28', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00'],
     *                          ['value' => '男', 'textColor' => 'FF000000', 'backgroundColor' => 'FF00FF00']
     *                      ],
     *                  ],
     *              ],
     *            [
     *               'title'=>'测试',
     *                'type'=>'2',
     *                'list'=>[['省份','总数',...],['Q1',12,...],['Q2',56,...],['Q3',52,...],['Q4',30,...]]],
     *       ],
     * ]
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderData($dataSources)
    {
        $file_name = getter($dataSources,'file_name');
        $title = getter($dataSources,'title');
        $data = getter($dataSources,'data');
        $logo = getter($dataSources,'logo');
        if(!empty($file_name)){
            $this->setFileName($file_name);
        }
        if(!empty($title)){
            $this->setMainTitle($title,$data);
        }
        if(!empty($logo)){
            $this->setLogo($logo);
        }
        if(!empty($data)){
            foreach($data as $item)
            {
                if(!empty($item))
                {
                    $type = getter($item,'type');
                    if($type != 4){
                        $this->setChartData($item);
                    }else{
                        $this->setList($item);
                    }
                }
            }
        }
    }

    /**
     * 导出文件
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function export()
    {
        // 导出的文件名
        $fileName = $this->fileName .date('YmdHis') . '.xlsx';
        $writer =IOFactory::createWriter($this->spreadsheet, 'Xlsx');
        $writer->setIncludeCharts(TRUE); //打开做图开关
        $writer->save(storage_path("app") . "/" . $fileName);
        //判断文件是否存在
        if (!file_exists(storage_path("app") . "/" . $fileName)) {
            return message(MESSAGE_FAILED, false);
        }
        // 移动文件
        copy(storage_path("app") . "/" . $fileName, UPLOAD_TEMP_PATH . "/" . $fileName);
        // 下载地址
        $fileUrl = get_image_url(str_replace(ATTACHMENT_PATH, "", UPLOAD_TEMP_PATH) . "/" . $fileName);
        return message(MESSAGE_OK, true, $fileUrl);
    }

    /**
     * 渲染数据并导出
     * @param $dataSources
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function excel($dataSources)
    {
        $this->renderData($dataSources);
        return $this->export();
    }

}
