# ExcelCreator
### <i>A simple way to work with PHPSpreadsheet</i>

## Introduction
<strong>ExcelCreator</strong> is an additional tool that enables you to use PHPSpreadsheet
more easily. ExcelCreator simplifies method calls like `$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(12)` into `$excel->setColumnWidth('A', 12)`.
You may think it does not really help you, but imagine that you have to write code like you usually use in PHPSpreadsheet many times. If you are a lazy programmer, ExcelCreator is a perfect tool to access commonly-used features in PHPSpreadsheet in a convenient way.

## Notes
- ExcelCreator built for CodeIgniter4 library, but you can still use it everywhere by modifying the `require` section on top line of the class.
- This documentation assumes that you have already experienced using PHPSpreadsheet

## Usage
This section will guide you how to transform PHPSpreadsheet original use into ExcelCreator
- Initializing ExcelCreator<br>
Assumes you have imported/included ExcelCreator, simply make an object 
```
$excel = new ExcelCreator();
```

- Call the Xlsx writer
```
$excel->writer($excel->spreadsheet);
```

- Call the Xlsx reader
```
$excel->reader($excel->spreadsheet);
```

- Saving excel file to client's browser not get more simple<br>
```
$excel->save('hello world.xlsx'); // save in Excel 2007 format
// or
$excel->save('hello world.xlsx', 'Xls'); // save in Xls format
```
Note that you have to set Content-Type header before use this method. 
And also the second parameter should follow the PHPSpreadsheet file type.

- Apply styles into cells<br>
The best way to apply style is using array and pass it as parameter
```
$dataStyle = [            
    'font' => [
        'name' => 'Arial',
        'size' => 10,
    ],
    'border' => [
        'borderStyle' => $excel->border::BORDER_THIN,
        'color' => $excel->color::COLOR_BLACK,
    ],
];

$excel->applyStyle($dataStyle, 'A2:D10');
```
Styles that have been supported in ExcelCreator are `Alignment`, `Border`, `Color`, `Fill` and `Font`

- Fill cells with data<br>
ExcelCreator provides simple way to fill cells with your data/value. ExcelCreator uses method chaining from PHPSpreadsheet `Spreadsheet::getActiveSheet()->fromArray($value);`
```
$data = [
    ['Name',        'Place of birth'],
    ['Adnan Zaki',  'Jakarta'],
    ['Dien Azizah', 'Bojonegoro']
];

$excel->fillCell($data); // fill cell from A1
// or 
$excel->fillCell($data, 'A2'); // fill cell from A2
```

- Wrapping text<br>
Wrapping a cell is get easier
```
$excel->wrapText('B5');
```
Note: Wrapped text can be overridden if you set style array after `wrapText()` without
defining wrapText in alignment.


- Merge and unmerge cells<br>
Although merging and unmerging cells in PHPSpreadsheet is easy, but we make it more simple.
```
$excel->mergeCells('A1:B1');
$excel->unmergeCells('A1:B1');
```

- Setting column's width<br>
Have you ever get tired of writing `$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(12);`
just to set column's width? Here is how ExcelCreator makes it easy for you:
```
$excel->setColumnWidth('D', 12);

// or

$excel->setColumnWidth('D'); // will create auto size
```
Or if you would like to set some columns with the same size:
```
$columns = ['A', 'B', 'C'];
$excel->setMultipleColumnsWidth($columns, 12);

or

$excel->setMultipleColumnsWidth($columns); // will create auto size for those columns
```
And if you would like to set default column's width
```
$excel->setDefaultColumnWidth(12);
```

- Setting row's height<br>
Setting a row's height is much similar like setting column's width
```
$excel->setRowHeight('5', 20);
```
Or if you would like to set some rows with the same height:
```
$excel->setMultipleRowsHeight('1-5', 20);
```
But now `setMultipleRowsHeight()` supports more multiple rows:
```
$excel->setMultipleRowsHeight(['1' => 40, '2' => 30 '3-6' => 20]);
```
And if you would like to set default row's height
```
$excel->setDefaultRowHeight(20);
```
- Set default font<br>
Setting default font is as easy as follow:
```
$excel->setDefaultFont('Arial', 12)
```