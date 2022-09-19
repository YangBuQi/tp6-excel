<?php


namespace app\Excel;


use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use think\exception\ValidateException;
use think\facade\Filesystem;

class Excel
{
    //需要先安装composer扩展命令：composer require phpoffice/phpspreadsheet

    /**
     * Excel导入
     * Author：ly
     * Date：2022/9/16
     * Time：11:32
     */
    public function ImportExcel($filename = "")
    {
        $file[] = $filename;

        try {
            // 验证文件大小，名称等是否正确
            validate(['file' => 'filesize:51200|fileExt:xls,xlsx'])
                ->check($file);
            // 将文件保存到本地
            $savename = Filesystem::disk('public')->putFile('file', $file[0]);
            // 截取后缀
            $fileExtendName = substr(strrchr($savename, '.'), 1);
            // 有Xls和Xlsx格式两种
            if ($fileExtendName == 'xlsx') {
                $objReader = IOFactory::createReader('Xlsx');
            } else {
                $objReader = IOFactory::createReader('Xls');
            }
            // 设置文件为只读
            $objReader->setReadDataOnly(TRUE);
            // 读取文件，tp6默认上传的文件，在runtime的相应目录下，可根据实际情况自己更改
            $objPHPExcel = $objReader->load(public_path() . 'storage/' . $savename);
            //excel中的第一张sheet
            $sheet = $objPHPExcel->getSheet(0);
            // 取得总行数
            $highestRow = $sheet->getHighestRow();
            // 取得总列数
            $highestColumn = $sheet->getHighestColumn();
            Coordinate::columnIndexFromString($highestColumn);
            $lines = $highestRow - 1;
            if ($lines <= 0) {
                return "数据为空数组";
            }
            // 直接取出excle中的数据
            $data = $objPHPExcel->getActiveSheet()->toArray();
            // 删除第一个元素（表头）
            array_shift($data);
            //导入成功后，删除文件
            unlink(public_path() . 'storage/' . $savename);
            // 返回结果
            return $data;
        } catch (ValidateException $e) {
            return $e->getMessage();
        }
    }

    /**
     * 导出
     * Author：ly
     * Date：2022/9/16
     * Time：11:39
     */
    public function ExportExcel($header = [], $type = true, $data = [], $fileName = "TP6")
    {
        // 实例化类
        $preadsheet = new Spreadsheet();
        // 创建sheet
        $sheet = $preadsheet->getActiveSheet();
        // 循环设置表头数据
        foreach ($header as $k => $v) {
            $sheet->setCellValue($k, $v);
        }
        // 生成数据
        $sheet->fromArray($data, null, "A2");
        // 样式设置
        $sheet->getDefaultColumnDimension()->setWidth(12);
        // 设置下载与后缀
        if ($type) {
            header("Content-Type:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            $type = "Xlsx";
            $suffix = "xlsx";
        } else {
            header("Content-Type:application/vnd.ms-excel");
            $type = "Xls";
            $suffix = "xls";
        }
        ob_end_clean();//清楚缓存区
        // 激活浏览器窗口
        header("Content-Disposition:attachment;filename=$fileName.$suffix");
        //缓存控制
        header("Cache-Control:max-age=0");
        // 调用方法执行下载
        $writer = IOFactory::createWriter($preadsheet, $type);
        // 数据流
        $writer->save("php://output");
    }

}