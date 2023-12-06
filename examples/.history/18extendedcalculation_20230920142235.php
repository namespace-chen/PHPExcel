<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2015 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

/** PHPExcel */
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


// List functions
echo date('H:i:s') . " List implemented functions\n";
$objCalc = PHPExcel_Calculation::getInstance();
print_r($objCalc->listFunctionNames());

// Create new PHPExcel object
echo date('H:i:s') . " Create new PHPExcel object\n";
$objPHPExcel = new PHPExcel();

// Add some data, we will use some formulas here
echo date('H:i:s') . " Add some data\n";
$objXLSX->getActiveSheet()->setCellValue('A14', 'Count:');

$objXLSX->getActiveSheet()->setCellValue('B1', 'Range 1');
$objXLSX->getActiveSheet()->setCellValue('B2', 2);
$objXLSX->getActiveSheet()->setCellValue('B3', 8);
$objXLSX->getActiveSheet()->setCellValue('B4', 10);
$objXLSX->getActiveSheet()->setCellValue('B5', True);
$objXLSX->getActiveSheet()->setCellValue('B6', False);
$objXLSX->getActiveSheet()->setCellValue('B7', 'Text String');
$objXLSX->getActiveSheet()->setCellValue('B9', '22');
$objXLSX->getActiveSheet()->setCellValue('B10', 4);
$objXLSX->getActiveSheet()->setCellValue('B11', 6);
$objXLSX->getActiveSheet()->setCellValue('B12', 12);

$objXLSX->getActiveSheet()->setCellValue('B14', '=COUNT(B2:B12)');

$objXLSX->getActiveSheet()->setCellValue('C1', 'Range 2');
$objXLSX->getActiveSheet()->setCellValue('C2', 1);
$objXLSX->getActiveSheet()->setCellValue('C3', 2);
$objXLSX->getActiveSheet()->setCellValue('C4', 2);
$objXLSX->getActiveSheet()->setCellValue('C5', 3);
$objXLSX->getActiveSheet()->setCellValue('C6', 3);
$objXLSX->getActiveSheet()->setCellValue('C7', 3);
$objXLSX->getActiveSheet()->setCellValue('C8', '0');
$objXLSX->getActiveSheet()->setCellValue('C9', 4);
$objXLSX->getActiveSheet()->setCellValue('C10', 4);
$objXLSX->getActiveSheet()->setCellValue('C11', 4);
$objXLSX->getActiveSheet()->setCellValue('C12', 4);

$objXLSX->getActiveSheet()->setCellValue('C14', '=COUNT(C2:C12)');

$objXLSX->getActiveSheet()->setCellValue('D1', 'Range 3');
$objXLSX->getActiveSheet()->setCellValue('D2', 2);
$objXLSX->getActiveSheet()->setCellValue('D3', 3);
$objXLSX->getActiveSheet()->setCellValue('D4', 4);

$objXLSX->getActiveSheet()->setCellValue('D5', '=((D2 * D3) + D4) & " should be 10"');

$objXLSX->getActiveSheet()->setCellValue('E1', 'Other functions');
$objXLSX->getActiveSheet()->setCellValue('E2', '=PI()');
$objXLSX->getActiveSheet()->setCellValue('E3', '=RAND()');
$objXLSX->getActiveSheet()->setCellValue('E4', '=RANDBETWEEN(5, 10)');

$objXLSX->getActiveSheet()->setCellValue('E14', 'Count of both ranges:');
$objXLSX->getActiveSheet()->setCellValue('F14', '=COUNT(B2:C12)');

// Calculated data
echo date('H:i:s') . " Calculated data\n";
echo 'Value of B14 [=COUNT(B2:B12)]: ' . $objXLSX->getActiveSheet()->getCell('B14')->getCalculatedValue() . "\r\n";


// Echo memory peak usage
echo date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB\r\n";

// Echo done
echo date('H:i:s') . " Done" , EOL;
