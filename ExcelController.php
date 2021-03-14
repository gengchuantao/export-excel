<?php
namespace app\common\controller;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use XLSXWriter;
class ExcelController
{
    /**
     * 导出excel表
     * $data：要导出excel表的数据，接受一个二维数组
     * $name：excel表的表名
     * $head：excel表的表头，接受一个一维数组
     * $key：$data中对应表头的键的数组，接受一个一维数组
     * 备注：此函数缺点是，表头（对应列数）不能超过52；
     *循环不够灵活，一个单元格中不方便存放两个数据库字段的值
     */

    /**

     * @ description 文件导出

     * @ date 2019-05-06

     * @ array $field_name 字段名称 汉字(索引数组) ['产品','姓名']

     * @ array $data 数据 ['a' => data, 'b' => data]

     * @ array $field_column 数据中的下标名称 字段数据 (索引数组) ['a','b']

     * @ string $file_name 文件名称

     * @ array $arr 需要转换为数字的$field_column中的key(索引数组)

     * @ return  file

     */
    public function outdata($name,$head = [],$data = [],$keys = [])
    {
        if (empty($name)|| empty($head)|| empty($data)|| empty($keys))return false;
        $count = count($head);  //计算表头数量
        $xlsTitle = iconv('utf-8', 'gb2312', $name);//文件名称
        $fileName = $xlsTitle.date('_YmdHis');
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //设置header
        $i= 0;
        foreach ($head as $value) {
            $cellName= self::stringFromColumnIndex($i). "1";
            $sheet->setCellValue($cellName, $value)->calculateColumnWidths();
            $sheet->getColumnDimension(self::stringFromColumnIndex($i))->setWidth(15);
            $i++;
        }

        //设置value
        $len= count($keys);
        foreach ($data as $key=> $item) {
            $row= 2 + ($key* 1);
            for ($i= 0; $i< $len; $i++) {
                $sheet->setCellValue(self::stringFromColumnIndex($i). $row, $item[$keys[$i]] );
            }
            ob_flush();
        }
        ob_end_clean();
        ob_start();
        // 设置输出头部信息
        header('Content-Encoding: UTF-8');
        header("Content-Type: text/csv; charset=UTF-8");
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '.csv"');
        header('Content-Description: File Transfer');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Cache-Control: max-age=0');
        echo chr(0xEF).chr(0xBB).chr(0xBF);
        $writer = new Csv($spreadsheet);
        $writer->save('php://output');
        ob_flush();
        flush();
        unset($spreadsheet);
        unset($writer);
        exit();
    }
    public static function stringFromColumnIndex($pColumnIndex = 0)

    {

        static $_indexCache= array();

        if (!isset($_indexCache[$pColumnIndex])) {

            if ($pColumnIndex < 26) {

                $_indexCache[$pColumnIndex]= chr(65 + $pColumnIndex);

            }elseif ($pColumnIndex < 702) {

                $_indexCache[$pColumnIndex]= chr(64 + ($pColumnIndex / 26)). chr(65 + $pColumnIndex % 26);

            }else {

                $_indexCache[$pColumnIndex]= chr(64 + (($pColumnIndex - 26)/ 676)). chr(65 + ((($pColumnIndex - 26)% 676)/ 26)). chr(65 + $pColumnIndex % 26);

            }

        }

        return $_indexCache[$pColumnIndex];

    }
    public function export($name,$head = [],$data = [],$keys = [])
    {

        $writer = new XLSXWriter();
        if (empty($name)|| empty($head)|| empty($data)|| empty($keys))return false;
        //定义文件名
        $filename = $name.date('_YmdHis');
        //定义工作表名称
        $sheet1 = $name;
        //对每列指定数据类型，对应单元格的数据类型
        /*foreach ($head as $key => $item){
            $col_style[] = $key ==5 ? 'price': 'string';
        }*/
        //设置列格式，suppress_row: 去掉会多出一行数据；widths: 指定每列宽度
        //$writer->writeSheetHeader($sheet1, $col_style, ['suppress_row'=>true,'widths'=>[20,20,20,20,20,20,20,20]] );
        //$writer->writeSheetHeader($sheet1, $col_style, ['suppress_row'=>true,'widths'=>[]] );
        //写入第二行的数据，顺便指定样式
        //$writer->writeSheetRow($sheet1, [$name], ['height'=>32,'font-size'=>16,'font-style'=>'bold','halign'=>'center','valign'=>'center']);/*设置标题头，指定样式*/
        /*设置字段，指定样式*/
        $styles1 = array( 'font'=>'宋体','font-size'=>9,'font-style'=>'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom');
        //写入第一行字段
        $writer->writeSheetRow($sheet1,$head,$styles1);
        // 最后是数据，foreach写入
        foreach ($data as $value) {
            foreach ($value as $item) {
                $temp[] = $item;
            }
            $rows[] = $temp;
            unset($temp);
        }
        //定义数据格式
        $styles2 = ['height'=>13,'font'=>'宋体','font-size'=>9,'border'=>'left,right,top,bottom'];
        //逐条写入数据
        foreach($rows as $row){
            $writer->writeSheetRow($sheet1,$row,$styles2);
        }
        //合并单元格，第一行的大标题需要合并单元格
        //$writer->markMergedCell($sheet1, $start_row=0, $start_col=0, $end_row=0, $end_col=7);
        //设置 header，用于浏览器下载
        $ua = isset ( $_SERVER ["HTTP_USER_AGENT"] ) ? $_SERVER ["HTTP_USER_AGENT"] : '';
        if (preg_match ( "/Trident/", $ua )) {
            //判断是否为IE11浏览器
            $filename=urlencode($filename);
            //定义编码，解决IE下输出文件名乱码
            header('Content-Encoding: GB2312');
            //以下为设置下载类型
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
            header('Content-Transfer-Encoding: binary');
        } else {
            //定义编码
            header('Content-Encoding: UTF-8');
            //以下为设置下载类型
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
            header('Content-Transfer-Encoding: binary');
        }
        header('Content-Description: File Transfer');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Cache-Control: max-age=0');
        //输出文档
        $writer->writeToStdOut();
        exit(0);
    }
    public function exports($name,$head = [],$data = [],$keys = [])
    {

        $writer = new XLSXWriter();
        if (empty($name)|| empty($head)|| empty($data)|| empty($keys))return false;
        //定义文件名
        $filename = $name.date('_YmdHis');
        //定义工作表名称
        $sheet1 = $name;
        /*设置字段，指定样式*/
        $styles1 = array( 'font'=>'宋体','font-size'=>9,'font-style'=>'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom');
        //写入第一行字段
        $writer->writeSheetHeader($sheet1,$head,$styles1);
        // 最后是数据，foreach写入
        foreach ($data as $value) {
            foreach ($value as $item) {
                $temp[] = $item;
            }
            $rows[] = $temp;
            unset($temp);
        }
        //定义数据格式
        $styles2 = ['height'=>13,'font'=>'宋体','font-size'=>9,'border'=>'left,right,top,bottom'];
        //逐条写入数据
        foreach($rows as $row){
            $writer->writeSheetRow($sheet1,$row,$styles2);
        }
        //设置 header，用于浏览器下载
        $ua = isset ( $_SERVER ["HTTP_USER_AGENT"] ) ? $_SERVER ["HTTP_USER_AGENT"] : '';
        if (preg_match ( "/Trident/", $ua )) {
            //判断是否为IE11浏览器
            $filename=urlencode($filename);
            //定义编码，解决IE下输出文件名乱码
            header('Content-Encoding: GB2312');
            //以下为设置下载类型
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
            header('Content-Transfer-Encoding: binary');
        } else {
            //定义编码
            header('Content-Encoding: UTF-8');
            //以下为设置下载类型
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
            header('Content-Transfer-Encoding: binary');
        }
        header('Content-Description: File Transfer');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Cache-Control: max-age=0');
        //输出文档
        $writer->writeToStdOut();
        exit(0);
    }

}