# phpspreadsheet

## 1. composer install

```json
"require": {
		"phpoffice/phpspreadsheet": "^1.5",
	},
```

## 2. download excel

```php
require VENDORPATH.'autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

public function download(){
    ini_set ('memory_limit', '1024M');
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $title = '订单信息报表';
    $title_array = ['序号','用户昵称','手机号','商品名称&数量','收货地址'];
    $export_array[] = $title1_array;
    $result = []; //查询出的二维数组
    foreach($result as $vv){
        $export_array[] = array_values($vv);
    }
    
    foreach ($export_array as $key1=>$sub_data) { //列
        foreach ($sub_data as $key2=>$item) { //行
            $sheet->setCellValueExplicitByColumnAndRow($key2+1, $key1+1,$item,'s');
        }
    }
}
unset($data);
$writer = new Xlsx($spreadsheet);
unset($spreadsheet);
$file_name = '订单信息报表'.date('YmdHis');
//        $writer->save(FILEPATH.'excel/'.$fileName);
header("Pragma: public");
header("Expires: 0");
header('Access-Control-Allow-Origin:*');
header('Access-Control-Allow-Headers:content-type');
header('Access-Control-Allow-Credentials:true');
header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
header("Content-Type:application/force-download");
header("Content-Type:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header("Content-Type:application/octet-stream");
header("Content-Type:application/download");;
header("Content-Disposition:attachment;filename=".$file_name.'.xlsx');
header("Content-Transfer-Encoding:binary");
$writer->save('php://output');
exit();
```

## 3. download excel with cell style

```php
require VENDORPATH.'autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;   


// 导出订单信息
public function exportOrder($data){
        ini_set ('memory_limit', '1024M');
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $title = '商品订单信息报表';
        $sql = "SELECT SQL_CALC_FOUND_ROWS * from orders ";
        $sql .= " order by o.pay_time DESC ";
        $result = $this->db->query($sql)->result_array();
        $title1_array = [$title,'','','','','','','','','','']; // 文档标题
        $title2_array = ['序号','用户昵称','手机号','商品名称&数量','收货地址',]; // 文档列名
        $export_array[] = $title1_array;
        $export_array[] = $title2_array;
        $order_no = 0; //组织序号
        foreach($result as &$item){
            $order_no++;
            array_unshift($item,$order_no);
            $item['comment_status'] = $item['comment_status'] == 1  ? '已评价' : '待评价';
            $item['is_finish'] = $item['is_finish'] == 1  ? '交易完成' : '未完成';
            if($item['order_status'] == 1){
                $item['order_status'] = '正常';
            }elseif($item['order_status'] == 2){
                $item['order_status'] = '退款/退货';
            }elseif($item['order_status'] == 3){
                $item['order_status'] = '已退款';
            }
            unset($item['order_id']);
            unset($item['user_id']);
            unset($item['reciever']);
            unset($item['recieve_mobile']);

            $export_array[] = array_values($item);
         }
        $sheet->getDefaultColumnDimension()->setWidth(20); //设置默认列宽
    	// 以下针对每列设置行宽
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(8);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(30);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(40);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(10);
        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(30);
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(10);
        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(13);
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('Z')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('AA')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('AB')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('AC')->setWidth(25);
        $spreadsheet->getActiveSheet()->getColumnDimension('AD')->setWidth(10);
        $spreadsheet->getActiveSheet()->getColumnDimension('AE')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('AF')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('AG')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('AH')->setWidth(15);
        $spreadsheet->getActiveSheet()->getColumnDimension('AI')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('AJ')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('AK')->setWidth(20);
        $sheet->mergeCells('A1:AK1');//单元格合并合并
        $sheet->getStyle('A1:AK1')->getFont()->setBold(true)->setName('宋体')
            ->setSize(24);//单元格内字体样式,颜色
        $sheet->getStyle('A:AK')->getFont()->setName('Times New Roman');
        $sheet->getStyle('A1:AK1')->getFont()->setBold(true);
        $sheet->getStyle('A2:AK2')->getFont()->setBold(true);;
        $sheet->getRowDimension('1')->setRowHeight(60);//设置行高
        $sheet->getRowDimension('2')->setRowHeight(30);
        $sheet->getDefaultRowDimension()->setRowHeight(40);
        $sheet->getStyle('B:AK')->getAlignment()->setWrapText(true);//自动换行
        $styleArray = [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER, // 水平居中
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ], // 垂直居中
        ];
        $sheet->getStyle('A1:AK2')->applyFromArray($styleArray);
        $sheet->getStyle('A:AK')->applyFromArray($styleArray);
    
        foreach ($export_array as $key1=>$sub_data) { //列
            foreach ($sub_data as $key2=>$item) { //行
                $sheet->setCellValueExplicitByColumnAndRow($key2+1, $key1+1,$item,'s');
            }
        }
//        unset($data);
        $writer = new Xlsx($spreadsheet);
        unset($spreadsheet);
        $file_name = '订单信息报表'.date('YmdHis');
//        $writer->save(FILEPATH.'excel/'.$fileName); 保存文件至服务器
        header("Pragma: public");
        header("Expires: 0");
        header('Access-Control-Allow-Origin:*');
        header('Access-Control-Allow-Headers:content-type');
        header('Access-Control-Allow-Credentials:true');
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");;
        header("Content-Disposition:attachment;filename=".$file_name.'.xlsx');
        header("Content-Transfer-Encoding:binary");
        $writer->save('php://output');
        exit();
    }

```

## 4. download excel with images

```php
require VENDORPATH.'autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
// excel导出带图片的单元格
// 
```

