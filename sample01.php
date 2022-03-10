<?php

require 'vendor/autoload.php';

// create new object
$excel = new ExcelCreator();

// initialize PHPSpreadsheet
$spreadsheet = $excel->spreadsheet;

$spreadsheet->setActiveSheetIndex(0);

// prepare column's header
$header = [
    'No', 'Nama', 'Tempat Lahir', 'Hobi'
];

// data to be filled into cells
$data = [
    ['1', 'Adnan Zaki',  'Jakarta', 'Coding'],
    ['2', 'Dien Azizah', 'Bojonegoro', 'Reading']
];

// fill cells with header and data as provided
$excel->fillCell($header);
$excel->fillCell($data, 'A2');

// set first column's width
$excel->setColumnWidth('A', 6);

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
    'alignment' => [
        'vertical' => $excel->alignment::VERTICAL_CENTER,
        'horizontal' => $excel->alignment::HORIZONTAL_CENTER
    ],
    'font' => [
        'name' => 'Arial',
        'size' => 11,
        'bold' => true,
    ],
    'borders' => [
        'top' => ['borderStyle' => $excel->border::BORDER_THIN],
        'bottom' => ['borderStyle' => $excel->border::BORDER_THIN],
        'right' => ['borderStyle' => $excel->border::BORDER_THIN],
        'left' => ['borderStyle' => $excel->border::BORDER_THIN],
        'color' => $excel->color::COLOR_BLACK,
    ],
];

// configure style for data rows
$dataStyle = [            
    'font' => [
        'name' => 'Arial',
        'size' => 10,
    ],
    'borders' => $headerStyle['borders'] // same as header's border
];

$numStyle = [
    'alignment' => [
        'horizontal' => $excel->alignment::HORIZONTAL_CENTER
    ],
    'borders' => $headerStyle['borders']
];

// apply styles
$excel->applyStyle($headerStyle, 'A1:D1');
$excel->applyStyle($dataStyle, 'A2:D10');
$excel->applyStyle($numStyle, 'A2:A10');

// you do not need to define header('Content-type: ....') 
// if your Excel filetype is "Xlsx", and you can leave the second argument of
// $excel->save() blank as well. Otherwise, you have to define the header()
// and pass the excel filetype in the second argument.
$excel->save('halo anak muda.xlsx');