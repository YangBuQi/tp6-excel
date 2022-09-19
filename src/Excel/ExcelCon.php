<?php
declare (strict_types = 1);

namespace app\controller;

use app\Excel\Excel;
use think\facade\Db;
use think\Request;

class ExcelController
{
    /**
     * Excel导入
     * Author：ly
     * Date：2022/9/16
     * Time：19:03
     */
   public function ImportExcel(Request $request)
   {
       // 接收文件上传信息
       $files = $request->file("file");
       // 调用类库，读取excel中的内容,返回导入的数据
       $data = (new Excel())->ImportExcel($files);

       dd($data);   //  导入的数据，二维数组
   }

    /**
     * Excel导出
     * Author：ly
     * Date：2022/9/16
     * Time：19:03
     */
   public function ExportExcel()
   {
       //查询表字段名称（字段名称为英文，导出数据需要的是字段的注释）
       $Field = Db::query("show full columns from user");  //user为表名，根据表名获取
       for($i=65;$i<91;$i++){
           //strtoupper将字符转换为大写，chr() 函数从指定 ASCII 值返回字符
           $Alphabet[] =  strtoupper(chr($i));//输出大写字母
       }
       foreach ($Field as $item=>$v)
       {
           $Comment[] = $v['Comment']; //获取到表注释
       }
       for ($i=0;$i<count($Comment);$i++)
       {
           $header[] = $Alphabet[$i].'1';  //循环拼接头部格式下标
       }
       //合并两个数组来创建一个新数组，一个数组元素为键名，另一个数组元素为键值
       $NewHeader = array_combine($header,$Comment);
       //从数据库里查值
       $data = Db::name('user')->select()->toArray();
       // 保存文件的类型
       $type= true;
       // 设置下载文件保存的名称
       $fileName = '信息导出'.time();
       // 调用方法导出excel
       (new Excel())->ExportExcel($NewHeader,$type,$data,$fileName);
   }
}
