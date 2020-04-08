<?php

namespace Zhuoxin\Excel;

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xls;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

/*
基于phpspreadsheet表格操作类

单独使用类：
1.安装：
    composer require phpoffice/phpspreadsheet
2.使用：
    导出Excel表格： PhpSpreadsheet::exportExcel()
    导入Excel数据： PhpSpreadsheet::importExcel()

原文链接：https://blog.csdn.net/DestinyLordC/article/details/84071456
基于原文基础更改封装 --- Zhuo新
*/

class PhpSpreadsheet
{
    /**
     * 导出Excel表格
     *
     * @param  array   $datas         二维数组 [0=>['data'=>'数据1']，1=>['data'=>'数据2']] 二维数据范围 A-Z
     * @param  array   $data_hreader  数据表第一行 例：['用户名', '积分', '排名']
     * @param  string  $fileName      导出文件名称
     * @param  array   $options       操作选项，例如
     *                                bool   print       设置打印格式
     *                                string freezePane  锁定行数，例如表头为第一行，则锁定表头输入A2
     *                                array  setARGB     设置背景色，例如['A1', 'C1']
     *                                array  setWidth    设置宽度，例如['A' => 30, 'C' => 20]
     *                                int    initWidth   设置每列默认宽度，初始为25
     *                                bool   setBorder   设置单元格边框
     *                                array  mergeCells  设置合并单元格，例如['A1:J1' => 'A1:J1']
     *                                array  formula     设置公式，例如['F2' => '=IF(D2>0,E42/D2,0)']
     *                                array  format      设置格式，整列设置，例如['A' => 'General']
     *                                array  alignCenter 设置居中样式，例如['A1', 'A2']
     *                                array  bold        设置加粗样式，例如['A1', 'A2']
     *                                string savePath    保存路径，设置后则文件保存到服务器，不通过浏览器下载
     *
     * @return bool
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    static function exportExcel(
        array $datas,
        $data_hreader = [],
        $fileName = '',
        array $options = []
    ) {
        try {
            if (empty($datas)) {
                return false;
            }
            set_time_limit(0);
            /** @var Spreadsheet $objSpreadsheet */
            $objSpreadsheet = new Spreadsheet();
            /* 设置默认文字居左，上下居中 */
            $styleArray = [
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_LEFT,
                    'vertical'   => Alignment::VERTICAL_CENTER,
                ],
            ];
            $objSpreadsheet->getDefaultStyle()->applyFromArray($styleArray);
            /* 设置Excel Sheet */
            $activeSheet = $objSpreadsheet->setActiveSheetIndex(0);

            /* 打印设置 */
            if (isset($options['print']) && $options['print']) {
                /* 设置打印为A4效果 */
                $activeSheet->getPageSetup()->setPaperSize(
                    PageSetup:: PAPERSIZE_A4
                );
                /* 设置打印时边距 */
                $pValue = 1 / 2.54;
                $activeSheet->getPageMargins()->setTop($pValue / 2);
                $activeSheet->getPageMargins()->setBottom($pValue * 2);
                $activeSheet->getPageMargins()->setLeft($pValue / 2);
                $activeSheet->getPageMargins()->setRight($pValue / 2);
            }

            /*合并数据表第一行文字*/
            if ($data_hreader) {
                array_unshift($datas, $data_hreader);
            }

            /*根据数据列长度 组装表格列数*/
            $cellBase = [];
            $initCell = 'A';
            for ($i = 0; $i < count($datas[0]); $i++) {
                array_push($cellBase, $initCell);
                $initCell++;
            }

            /*组装数据结构*/
            $newDatas = [];
            foreach ($datas as $key => $data) {
                $i         = 0;
                $cliderArr = [];
                foreach ($data as $item) {
                    $cliderArr[$cellBase[$i].($key + 1)] = $item;
                    $i++;
                }
                array_push($newDatas, $cliderArr);
            }
            unset($datas);
            /* 行数据处理 */
            foreach ($newDatas as $data) {
                foreach ($data as $sKey => $sItem) {
                    /* 默认文本格式 */
                    $pDataType = DataType::TYPE_STRING;
                    /* 设置单元格格式 */
                    if (isset($options['format'])
                        && !empty($options['format'])
                    ) {
                        $colRow = Coordinate::coordinateFromString($sKey);
                        /* 存在该列格式并且有特殊格式 */
                        if (isset($options['format'][$colRow[0]])
                            && NumberFormat::FORMAT_GENERAL
                            != $options['format'][$colRow[0]]
                        ) {
                            $activeSheet->getStyle($sKey)->getNumberFormat()
                                ->setFormatCode($options['format'][$colRow[0]]);
                            if (false !== strpos(
                                    $options['format'][$colRow[0]], '0.00'
                                )
                                && is_numeric(
                                    str_replace(['￥', ','], '', $sItem)
                                )
                            ) {
                                /* 数字格式转换为数字单元格 */
                                $pDataType = DataType::TYPE_NUMERIC;
                                $sItem     = str_replace(
                                    ['￥', ','], '', $sItem
                                );
                            }
                        } elseif (is_int($sItem)) {
                            $pDataType = DataType::TYPE_NUMERIC;
                        }
                    }
                    $activeSheet->setCellValueExplicit(
                        $sKey, $sItem, $pDataType
                    );
                    //                $activeSheet->setCellValue($sKey, $sItem);
                    /* 存在:形式的合并行列，列入A1:B2，则对应合并 */
                    if (false !== strstr($sKey, ":")) {
                        $options['mergeCells'][$sKey] = $sKey;
                    }
                }
            }
            unset($newDatas);

            /* 设置锁定行 */
            if (isset($options['freezePane'])
                && !empty($options['freezePane'])
            ) {
                $activeSheet->freezePane($options['freezePane']);
                unset($options['freezePane']);
            }

            // 设置初始化宽度
            $initWidth = isset($options['initWidth'])
            && !empty($options['initWidth']) ? $options['initWidth'] : 25;
            foreach ($cellBase as $swItem) {
                $activeSheet->getColumnDimension($swItem)->setWidth(
                    $initWidth
                );
            }
            /* 设置宽度 */
            if (isset($options['setWidth']) && !empty($options['setWidth'])) {
                foreach ($options['setWidth'] as $swKey => $swItem) {
                    $activeSheet->getColumnDimension($swKey)->setWidth($swItem);
                }
                unset($options['setWidth']);
            }

            /* 设置背景色 */
            if (isset($options['setARGB']) && !empty($options['setARGB'])) {
                foreach ($options['setARGB'] as $sItem) {
                    $activeSheet->getStyle($sItem)
                        ->getFill()->setFillType(Fill::FILL_SOLID)
                        ->getStartColor()->setARGB(Color::COLOR_YELLOW);
                }
                unset($options['setARGB']);
            }

            /* 设置公式 */
            if (isset($options['formula']) && !empty($options['formula'])) {
                foreach ($options['formula'] as $fKey => $fItem) {
                    $activeSheet->setCellValue($fKey, $fItem);
                }

                unset($options['formula']);
            }

            /* 合并行列处理 */
            if (isset($options['mergeCells'])
                && !empty($options['mergeCells'])
            ) {
                $activeSheet->setMergeCells($options['mergeCells']);
                unset($options['mergeCells']);
            }

            /* 设置居中 */
            if (isset($options['alignCenter'])
                && !empty($options['alignCenter'])
            ) {
                $styleArray = [
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_CENTER,
                        'vertical'   => Alignment::VERTICAL_CENTER,
                    ],
                ];

                foreach ($options['alignCenter'] as $acItem) {
                    $activeSheet->getStyle($acItem)->applyFromArray(
                        $styleArray
                    );
                }

                unset($options['alignCenter']);
            }

            /* 设置加粗 */
            if (isset($options['bold']) && !empty($options['bold'])) {
                foreach ($options['bold'] as $bItem) {
                    $activeSheet->getStyle($bItem)->getFont()->setBold(true);
                }

                unset($options['bold']);
            }

            /* 设置单元格边框，整个表格设置即可，必须在数据填充后才可以获取到最大行列 */
            if (isset($options['setBorder']) && $options['setBorder']) {
                $border    = [
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN, // 设置border样式
                            'color'       => ['argb' => 'FF000000'],
                            // 设置border颜色
                        ],
                    ],
                ];
                $setBorder = 'A1:'.$activeSheet->getHighestColumn()
                    .$activeSheet->getHighestRow();
                $activeSheet->getStyle($setBorder)->applyFromArray($border);
                unset($options['setBorder']);
            }

            /* 导出名称 */
            $fileName = !empty($fileName) ? $fileName :
                (date('YmdHis').'.xlsx');

            if (!isset($options['savePath'])) {
                /* 直接导出Excel，无需保存到本地，输出07Excel文件 */
                header(
                    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                );
                header(
                    "Content-Disposition:attachment;filename=".iconv(
                        "utf-8", "GB2312//TRANSLIT", $fileName
                    )
                );
                header('Cache-Control: max-age=0');//禁止缓存
                $savePath = 'php://output';
            } else {
                $savePath = $options['savePath'];
            }

            ob_clean();
            ob_start();
            $objWriter = IOFactory::createWriter($objSpreadsheet, 'Xlsx');
            $objWriter->save($savePath);
            /* 释放内存 */
            $objSpreadsheet->disconnectWorksheets();
            unset($objSpreadsheet);
            ob_end_flush();
            return true;
        } catch (Exception $e) {
            return false;
        }
    }


    /**
     * 导入Excel数据
     *
     * @param  string  $file       文件地址
     * @param  int     $sheet      工作表sheet(传0则获取第一个sheet)
     * @param  int     $columnCnt  列数(传0则自动获取最大列)
     * @param  array   $options    操作选项
     *                             array mergeCells 合并单元格数组
     *                             array formula    公式数组
     *                             array format     单元格格式数组
     *
     * @return array
     * @throws \Exception
     */
    static function importExcel(
        $file = '',
        $sheet = 0,
        $columnCnt = 0,
        &$options = []
    ) {
        try {
            /* 转码 */
            $file = iconv("utf-8", "gb2312", $file);
            if (empty($file) OR !file_exists($file)) {
                throw new \Exception('文件不存在!');
            }

            /** @var Xlsx $objRead */
            $objRead = IOFactory::createReader('Xlsx');

            if (!$objRead->canRead($file)) {
                /** @var Xls $objRead */
                $objRead = IOFactory::createReader('Xls');

                if (!$objRead->canRead($file)) {
                    throw new \Exception('只支持导入Excel文件！');
                }
            }

            /* 如果不需要获取特殊操作，则只读内容，可以大幅度提升读取Excel效率 */
            empty($options) && $objRead->setReadDataOnly(true);
            /* 建立excel对象 */
            $obj = $objRead->load($file);
            /* 获取指定的sheet表 */
            $currSheet = $obj->getSheet($sheet);

            if (isset($options['mergeCells'])) {
                /* 读取合并行列 */
                $options['mergeCells'] = $currSheet->getMergeCells();
            }

            if (0 == $columnCnt) {
                /* 取得最大的列号 */
                $columnH = $currSheet->getHighestColumn();
                /* 兼容原逻辑，循环时使用的是小于等于 */
                $columnCnt = Coordinate::columnIndexFromString($columnH);
            }

            /* 获取总行数 */
            $rowCnt = $currSheet->getHighestRow();
            $data   = [];

            /* 读取内容 */
            for ($_row = 1; $_row <= $rowCnt; $_row++) {
                $isNull = true;

                for ($_column = 1; $_column <= $columnCnt; $_column++) {
                    $cellName = Coordinate::stringFromColumnIndex($_column);
                    $cellId   = $cellName.$_row;
                    $cell     = $currSheet->getCell($cellId);

                    if (isset($options['format'])) {
                        /* 获取格式 */
                        $format = $cell->getStyle()->getNumberFormat()
                            ->getFormatCode();
                        /* 记录格式 */
                        $options['format'][$_row][$cellName] = $format;
                    }

                    if (isset($options['formula'])) {
                        /* 获取公式，公式均为=号开头数据 */
                        $formula = $currSheet->getCell($cellId)->getValue();

                        if (0 === strpos($formula, '=')) {
                            $options['formula'][$cellName.$_row] = $formula;
                        }
                    }

                    if (isset($format) && 'm/d/yyyy' == $format) {
                        /* 日期格式翻转处理 */
                        $cell->getStyle()->getNumberFormat()->setFormatCode(
                            'yyyy/mm/dd'
                        );
                    }

                    $data[$_row][] = trim(
                        $currSheet->getCell($cellId)->getFormattedValue()
                    );

                    if (!empty($data[$_row])) {
                        $isNull = false;
                    }
                }

                /* 判断是否整行数据为空，是的话删除该行数据 */
                if ($isNull) {
                    unset($data[$_row]);
                }
            }

            return $data;
        } catch (\Exception $e) {
            throw $e;
        }
    }

}