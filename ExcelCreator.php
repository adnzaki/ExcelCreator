<?php

/**
 * ExcelCreator class is a simple and easy way to use PHPSpreadsheet in CodeIgniter4.
 * It simplifies some PHPSpreadsheet functionality and still let you
 * access all of PHPSpreadsheet features at once.
 * This class enables you to access commonly-used features of PHPSpreadsheet
 * without have to import each class into your code.
 * Just make an object of this class and you can access classes such as
 * Spreadsheet object itself, the writer, styles, etc.
 * 
 * @package     Library
 * @author      Adnan Zaki
 * @copyright   Wolestech DevTeam
 */

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;

class ExcelCreator
{
    /**
     * The Spreadsheet object
     * 
     * @var object
     */
    public $spreadsheet;

    /**
     * The Color object
     * 
     * @var object
     */
    public $color;

    /**
     * The Alighment style object
     * 
     * @var object
     */
    public $alignment;

    /**
     * The Border style object
     * 
     * @var object
     */
    public $border;

    /**
     * The Fill style object
     * 
     * @var object
     */
    public $fill;

    /**
     * The Font style object
     * 
     * @var object
     */
    public $font;

    /**
     * Constructor
     * 
     * @return void
     */
    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->alignment = new Alignment();
        $this->border = new Border();
        $this->color = new Color();
        $this->fill = new Fill();
        $this->font = new Font();
    }

    /**
     * Run the XlsxWriter() object
     * 
     * @param object $spreadsheet -> The Spreadsheet object
     * 
     * @return PhpOffice\PhpSpreadsheet\Writer\Xlsx
     */
    public function writer($spreadsheet)
    {
        return new XlsxWriter($spreadsheet);
    }

    /**
     * Run the XlsxReader() object
     * 
     * @param object $spreadsheet -> The Spreadsheet object
     * 
     * @return PhpOffice\PhpSpreadsheet\Reader\Xlsx
     */
    public function reader($spreadsheet)
    {
        return new XlsxReader($spreadsheet);
    }

    /**
     * A short way to save excel file
     * Note that user has to set content-type header first
     * before use this method
     * 
     * @param string $filename
     * @param string $excelType
     * 
     * @return void
     */
    public function save($filename, $excelType = 'Xlsx')
    {
        $contentType = [
            'Xlsx'  => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Xls'   => 'application/vnd.ms-excel'
        ];

        header('Content-type: ' . $contentType[$excelType]);
        header("Content-Disposition: attachment;filename={$filename}");
        header('Cache-Control: max-age=0');
        $writer = IOFactory::createWriter($this->spreadsheet, $excelType);

        ob_end_clean();
        $writer->save('php://output');
    }

    /**
     * A convenient way to apply style to PHPSpreadsheet
     * in a compact and simple code
     * Although there are many functions to set style in PHPSpreadsheet,
     * but it is recommended to apply from array as it has 
     * more simple and readable code
     * 
     * @param array $style
     * @param string $range
     * 
     * @return void
     */
    public function applyStyle($style, $range)
    {
        $styleConfig = new Style();
        $styleConfig->applyFromArray($style);
        $this->spreadsheet->getActiveSheet()->duplicateStyle($styleConfig, $range);
    }

    /**
     * Fill cell value from array
     * 
     * @param array $value
     * @param string $startCell
     * 
     * @return void
     */
    public function fillCell($value, $startCell = 'A1')
    {
        $this->spreadsheet->getActiveSheet()->fromArray($value, null, $startCell);
    }

    /**
     * Wrap a text
     * 
     * @param string $cell
     * 
     * @return void
     */
    public function wrapText($cell)
    {
        $this->spreadsheet->getActiveSheet()->getStyle($cell)->getAlignment()->setWrapText(true);
    }

    /**
     * Merge Cells
     * 
     * @param string $cells
     * 
     * @return void
     */
    public function mergeCells($cells)
    {
        $this->spreadsheet->getActiveSheet()->mergeCells($cells);
    }

    /**
     * Unmerge Cells
     * 
     * @param string $cells
     * 
     * @return void
     */
    public function unmergeCells($cells)
    {
        $this->spreadsheet->getActiveSheet()->unmergeCells($cells);
    }

    /**
     * Set column width
     * 
     * @param string $dimension
     * @param int|null $width
     * 
     * @return void
     */
    public function setColumnWidth($dimension, $width = null)
    {
        if($width === null)
        {
            $this->spreadsheet
                ->getActiveSheet()
                ->getColumnDimension($dimension)
                ->setAutoSize(true);
        }
        else 
        {
            $this->spreadsheet
                ->getActiveSheet()
                ->getColumnDimension($dimension)
                ->setWidth($width);
        }
    }

    /**
     * Set default column width
     * 
     * @param int $width
     * 
     * @return void
     */
    public function setDefaultColumnWidth($width)
    {
        $this->spreadsheet
            ->getActiveSheet()
            ->getDefaultColumnDimension()
            ->setWidth($width);
    }

    /**
     * Set multiple columns width
     * Sometimes you would like to make width of some columns
     * in the same size, and this is a short way to do that
     * 
     * @param array $dimensions
     * @param int|null $width
     * 
     * @return void
     */
    public function setMultipleColumnsWidth($dimensions = [], $width = null)
    {
        foreach($dimensions as $val)
        {
            $this->setColumnWidth($val, $width);
        }
    }

    /**
     * Set row height
     * 
     * @param string $dimension
     * @param int $height
     * 
     * @return void
     */
    public function setRowHeight($dimension, $height)
    {
        $this->spreadsheet
            ->getActiveSheet()
            ->getRowDimension($dimension)
            ->setRowHeight($height);
    }

    /**
     * Set default row height
     * 
     * @param int $height
     * 
     * @return void
     */
    public function setDefaultRowHeight($height)
    {
        $this->spreadsheet
            ->getActiveSheet()
            ->getDefaultRowDimension()
            ->setRowHeight($height);
    }

    /**
     * Set multiple rows with the same height
     * The $dimensions now supports string value like '1-5'
     * or array like ['1' => 40, '3-5' => 25]
     * 
     * @param string|array $dimensions
     * @param int|null $height
     * 
     * @return void
     */
    public function setMultipleRowsHeight($dimensions, $height = null)
    {
        if(is_array($dimensions))
        {
            foreach($dimensions as $key => $val)
            {
                $this->_setMultipleRowsHeight($key, $val);
            }
        }
        else
        {
            $this->_setMultipleRowsHeight($dimensions, $height);
        }
        
    }

    /**
     * This is the excecutor of setMultipleRowsHeight
     * 
     * @param string $dimensionRange
     * @param int $height
     * 
     * @return void
     */
    private function _setMultipleRowsHeight($dimensionRange, $height)
    {
        if(! strpos($dimensionRange, '-'))
        {
            $this->setRowHeight($dimensionRange, $height);
        }
        else
        {
            $range = explode('-', $dimensionRange);
            for($i = $range[0]; $i <= $range[1]; $i++)
            {
                $this->setRowHeight($i, $height);
            }
        }
    }

    /**
     * Set default font to a worksheet
     * 
     * @param string $font
     * @param int $size
     * 
     * @return void
     */
    public function setDefaultFont($font, $size)
    {
        $this->spreadsheet->getDefaultStyle()
            ->getFont()
            ->setName($font)
            ->setSize($size);
    }
}
