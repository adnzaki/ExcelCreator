# ExcelCreator

### <i>A simple and elegant way to work with PHPSpreadsheet</i>

## Introduction

<strong>ExcelCreator</strong> is a modern wrapper around PHPSpreadsheet that simplifies the most commonly used tasks in reading and writing Excel files. It provides two main classes: `Writer` for creating and formatting spreadsheets, and `Reader` for loading and reading spreadsheet content.

## Installation

Install ExcelCreator using [Composer](https://getcomposer.org/):

### With existing `composer.json`

```json
{
    "require": {
        "adnzaki/excel-creator": "^2.0.1"
    }
}
```

Then run:

```bash
composer update
```

### Without `composer.json`

```bash
composer require adnzaki/excel-creator
```

## Install the Latest Source Code

If you prefer the latest development version:

```json
{
    "require": {
        "adnzaki/excel-creator": "dev-main"
    }
}
```

Then run `composer update`.

## Usage

This library provides two main classes under the namespace `ExcelTools`:

### 📝 Writing Excel Files with `Writer`

```php
use ExcelTools\Writer;

$excel = new Writer();
```

* Saving file to browser:

```php
$excel->save('hello_world.xlsx');
```

* Apply styles to cell range:

```php
$style = [
    'font' => [
        'name' => 'Arial',
        'size' => 10,
    ],
    'borders' => [
        'allBorders' => [
            'borderStyle' => $excel->border::BORDER_THIN,
            'color' => ['argb' => $excel->color::COLOR_BLACK],
        ],
    ]
];
$excel->applyStyle($style, 'A2:D10');
```

* Fill data:

```php
$data = [
    ['Name', 'City'],
    ['Zaki', 'Jakarta'],
    ['Dien', 'Bojonegoro']
];
$excel->fillCell($data, 'A1');
```

* Text wrapping:

```php
$excel->wrapText('B5');
```

* Merge & unmerge:

```php
$excel->mergeCells('A1:B1');
$excel->unmergeCells('A1:B1');
```

* Column widths:

```php
$excel->setColumnWidth('A', 20);
$excel->setMultipleColumnsWidth(['B', 'C'], 25);
$excel->setDefaultColumnWidth(15);
```

* Row heights:

```php
$excel->setRowHeight(3, 25);
$excel->setMultipleRowsHeight(['1' => 40, '3-5' => 25]);
$excel->setDefaultRowHeight(20);
```

* Set default font:

```php
$excel->setDefaultFont('Calibri', 11);
```

---

### 📖 Reading Excel Files with `Reader`

```php
use ExcelTools\Reader;

$reader = new Reader();
$reader->loadFromFile('SampleFile.xlsx');
```

* Get all data from the active sheet:

```php
$data = $reader->getSheetData(true); // true = first row as headers
print_r($data);
```

* Get sheet names:

```php
$sheetNames = $reader->getSheetNames();
```

* Switch to another sheet:

```php
$reader->setActiveSheet('Sheet2');
```

## License

MIT

## Author

Adnan Zaki – [https://github.com/adnzaki](https://github.com/adnzaki)
