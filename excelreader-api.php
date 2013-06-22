<?php

require_once(__DIR__ . '/PHPExcel.php');
require_once(__DIR__ . '/PHPExcel/IOFactory.php');

class ExcelReader {

    public static function error($msg, $die = true) {
        echo "\n";
        echo $msg;
        echo "\n";

        if ($die === true) {
            die();
        }
    }

    public static function read_excel_file($file = NULL, $file_type = NULL, &$col_maxlength = NULL, $active_sheet_index = 0) {

        // if file type is a text file or csv file, try importing as CSV
        // else, try importing as Excel5 and Excel

        if (!$file) {
            self::error("Filename Missing");
        }

        if (!file_exists($file)) {
            self::error("File not found");
        }

        switch ($file_type) {
            case "text/csv":
            case "text/plain":
                try {
                    $reader = PHPExcel_IOFactory::createReader("CSV");

                    $workbook = $reader->
                                    load($file);
                } catch (Exception $e_csv) {
                    return FALSE;
                }

                break;
            case "application/vnd.oasis.opendocument.spreadsheet":
                try {
                    $reader = PHPExcel_IOFactory::createReader("OOCalc");

                    $workbook = $reader->load($file);
                } catch (Exception $e_csv) {
                    return FALSE;
                }
                break;
            case "application/vnd.ms-excel":
            case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            case "application/zip":
                try {
                    $reader = PHPExcel_IOFactory::createReader("Excel5");
                    $workbook = $reader->load($file);
                } catch (Exception $e_excel5) {
                    try {
                        $reader = PHPExcel_IOFactory::createReader("Excel2007");
                        $workbook = $reader->load($file);
                    } catch (Exception $e_excel2007) {
                        return FALSE;
                    }
                }
                break;
            default:
                return FALSE;
        }
        
  
    		if(!$active_sheet_index){
    			$active_sheet_index = 0;
    		}
    
    		$workbook->setActiveSheetIndex($active_sheet_index);
    
    
            $objWorksheet = $workbook->getActiveSheet();
    
            $returnset = array();
    
            $row_maxlength = array();
            $col_maxlength = array();
    
            $all_columns = array();
    
            foreach ($objWorksheet->getRowIterator() as $row) {
    
                $cellIterator = $row->getCellIterator();
    
                $cellIterator->setIterateOnlyExistingCells(false);
    
                $rowset = array();
    
                $row_index = $row->getRowIndex();
    
                foreach ($cellIterator as $cell_position => $cell) {
                    $value = $cell->getValue();
                    $col_index = $cell->getColumn();
    
                    if ($value != "" AND $value != NULL) {
                        $rowset[$col_index] = $value;
    
                        $col_maxlength[$col_index][] = strlen($value . "");
    
                        $all_columns[$row_index][] = $col_index;
                    }
                }
    
                $col_maxlength[0][] = strlen($row_index . "");
    
                $returnset[$row_index] = $rowset;
            }
    
            foreach ($col_maxlength as $col_pos => $col_lengths) {
                $col_maxlength[$col_pos] = max($col_lengths);
            }
    
            return $returnset;
        }
    
        public static function padding($val, $maxlen, $char = " ") {
            return str_pad($val, $maxlen, $char, STR_PAD_RIGHT);
        }
    }
