<?php

namespace ExcelTools;

/**
 * ExcelReader class is a dedicated class to handle reading Excel files
 * using PHPSpreadsheet, and it extends from ExcelCreator to maintain
 * structural harmony.
 * 
 * @package     Library
 * @author      Adnan Zaki
 * @copyright   Wolestech DevTeam
 */

use ExcelTools\Writer;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Reader extends Writer
{
    /**
     * Load spreadsheet from file
     *
     * @param string $filepath
     * @return void
     */
    public function loadFromFile($filepath)
    {
        $reader = IOFactory::createReaderForFile($filepath);
        $this->spreadsheet = $reader->load($filepath);
    }

    /**
     * Get data from active worksheet
     *
     * @param bool $withHeaderRow
     * @return array
     */
    public function getSheetData($withHeaderRow = false)
    {
        $sheet = $this->spreadsheet->getActiveSheet();
        $highestRow = $sheet->getHighestRow();
        $highestColumn = $sheet->getHighestColumn();
        $data = [];

        $headers = [];
        if ($withHeaderRow) {
            $headers = $sheet->rangeToArray("A1:{$highestColumn}1", null, true, false)[0];
            $startRow = 2;
        } else {
            $startRow = 1;
        }

        for ($row = $startRow; $row <= $highestRow; $row++) {
            $rowData = [];
            $cells = $sheet->rangeToArray("A{$row}:{$highestColumn}{$row}", null, true, false)[0];
            foreach ($cells as $colIndex => $cell) {
                $cellCoord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex + 1) . $row;
                $cellObj = $sheet->getCell($cellCoord);

                // Cek apakah cell ini bertipe tanggal
                if (\PhpOffice\PhpSpreadsheet\Shared\Date::isDateTime($cellObj)) {
                    $timestamp = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($cell);
                    $cell = date('Y-m-d', $timestamp);
                }

                if ($withHeaderRow) {
                    $rowData[$headers[$colIndex]] = $cell;
                } else {
                    $rowData[] = $cell;
                }
            }
            $data[] = $rowData;
        }

        return $data;
    }


    /**
     * Get list of worksheet names
     * 
     * @return array
     */
    public function getSheetNames()
    {
        return $this->spreadsheet->getSheetNames();
    }

    /**
     * Set worksheet by index or name
     * 
     * @param int|string $identifier
     * @return void
     */
    public function setActiveSheet($identifier)
    {
        if (is_numeric($identifier)) {
            $this->spreadsheet->setActiveSheetIndex($identifier);
        } else {
            $this->spreadsheet->setActiveSheetIndexByName($identifier);
        }
    }
}
