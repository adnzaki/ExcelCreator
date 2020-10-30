<?php

require 'ExcelCreator.php';

// create new object
$excel = new ExcelCreator();

// initialize PHPSpreadsheet
$spreadsheet = $excel->spreadsheet;

$spreadsheet->setActiveSheetIndex(0);

// prepare column's header
$header = [
    'No.', 'Nama Siswa', 'Nilai', 'Catatan'
];

// data to be filled into cells
$data = [
    ['Nama',        'Tempat Lahir'],
    ['Adnan Zaki',  'Jakarta'],
    ['Dien Azizah', 'Bojonegoro']
];

// fill cells with header and data as provided
$excel->fillCell($header);
$excel->fillCell($data);

// set first column's width
$excel->setColumnWidth('A', 14);

// set other columns' width to be automatic
// or you can set multiple columns with the custom same size
// by passing second argument of this method with the size value
$columns = ['B', 'C', 'D'];
$excel->setMultipleColumnsWidth($columns);

// set row height of the header
$excel->setRowHeight('1', 30);

// configure style for header
$headerStyle = [
    'fill' => [
        'fillType' => $excel->fill::FILL_SOLID,
        'color' => ['argb' => 'FFFFFF00'],
    ],
    'font' => [
        'name' => 'Arial',
        'size' => 11,
        'bold' => true,
    ],
    'border' => [
        'borderStyle' => $excel->border::BORDER_THIN,
        'color' => $excel->color::COLOR_BLACK,
    ],
];

// configure style for data rows
$dataStyle = [            
    'font' => [
        'name' => 'Arial',
        'size' => 10,
    ],
    'border' => $headerStyle['border'] // same as header's border
];

// apply styles
$excel->applyStyle($headerStyle, 'A1:D1');
$excel->applyStyle($dataStyle, 'A2:D10');

// this content-type set with CodeIgniter4 controller
// use traditional PHP's header() function if you do not use CodeIgniter4
$this->response->setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

// save to client's browser
$excel->save('halo anak muda.xlsx');