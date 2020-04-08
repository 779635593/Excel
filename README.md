### _基于PhpSpreadsheet做的导出导入Excel类_

> 教程：

- 安装：
``composer require zhuoxin/excel``
- 使用：
``use Zhuoxin\Excel\PhpSpreadsheet``

- 导出Excel表格
>``PhpSpreadsheet::exportExcel()``

     /*
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
     */

- 导入Excel数据
> ``PhpSpreadsheet::importExcel()``

    /**
     * @param  string  $file       文件地址
     * @param  int     $sheet      工作表sheet(传0则获取第一个sheet)
     * @param  int     $columnCnt  列数(传0则自动获取最大列)
     * @param  array   $options    操作选项
     *                             array mergeCells 合并单元格数组
     *                             array formula    公式数组
     *                             array format     单元格格式数组
     */

- Github：https://github.com/779635593/Excel.git     
> 原文链接：https://blog.csdn.net/DestinyLordC/article/details/84071456
  基于原文基础更改封装 --- Zhuo新