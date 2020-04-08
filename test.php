<?php

require "vendor/autoload.php";

use Zhuoxin\Excel\PhpSpreadsheet;

// 导出
$datas = [
    [
        'name' => '小明',
        'sex'  => '男'
    ],
    [
        'name' => '小王',
        'sex'  => '女'
    ]
];

$data_hreader = [
    '姓名',
    '性别'
];

PhpSpreadsheet::exportExcel($datas, $data_hreader);
