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

/** Include PHPExcel */
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


// Create new PHPExcel object
echo date('H:i:s') , " Create new PHPExcel object" , EOL;
$objPHPExcel = new PHPExcel();

// Set document properties
echo date('H:i:s') , " Set document properties" , EOL;
$objXLSX->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Maarten Balliauw")
							 ->setTitle("Office 2007 XLSX Test Document")
							 ->setSubject("Office 2007 XLSX Test Document")
							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
							 ->setKeywords("office 2007 openxml php")
							 ->setCategory("Test result file");


// Create a first sheet, representing sales data
echo date('H:i:s') , " Add some data" , EOL;
$objXLSX->setActiveSheetIndex(0);
$objXLSX->getActiveSheet()->setCellValue('B1', 'Invoice');
$objXLSX->getActiveSheet()->setCellValue('D1', PHPExcel_Shared_Date::PHPToExcel( gmmktime(0,0,0,date('m'),date('d'),date('Y')) ));
$objXLSX->getActiveSheet()->getStyle('D1')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_XLSX15);
$objXLSX->getActiveSheet()->setCellValue('E1', '#12566');

$objXLSX->getActiveSheet()->setCellValue('A3', 'Product Id');
$objXLSX->getActiveSheet()->setCellValue('B3', 'Description');
$objXLSX->getActiveSheet()->setCellValue('C3', 'Price');
$objXLSX->getActiveSheet()->setCellValue('D3', 'Amount');
$objXLSX->getActiveSheet()->setCellValue('E3', 'Total');

$objXLSX->getActiveSheet()->setCellValue('A4', '1001');
$objXLSX->getActiveSheet()->setCellValue('B4', 'PHP for dummies');
$objXLSX->getActiveSheet()->setCellValue('C4', '20');
$objXLSX->getActiveSheet()->setCellValue('D4', '1');
$objXLSX->getActiveSheet()->setCellValue('E4', '=IF(D4<>"",C4*D4,"")');

$objXLSX->getActiveSheet()->setCellValue('A5', '1012');
$objXLSX->getActiveSheet()->setCellValue('B5', 'OpenXML for dummies');
$objXLSX->getActiveSheet()->setCellValue('C5', '22');
$objXLSX->getActiveSheet()->setCellValue('D5', '2');
$objXLSX->getActiveSheet()->setCellValue('E5', '=IF(D5<>"",C5*D5,"")');

$objXLSX->getActiveSheet()->setCellValue('E6', '=IF(D6<>"",C6*D6,"")');
$objXLSX->getActiveSheet()->setCellValue('E7', '=IF(D7<>"",C7*D7,"")');
$objXLSX->getActiveSheet()->setCellValue('E8', '=IF(D8<>"",C8*D8,"")');
$objXLSX->getActiveSheet()->setCellValue('E9', '=IF(D9<>"",C9*D9,"")');

$objXLSX->getActiveSheet()->setCellValue('D11', 'Total excl.:');
$objXLSX->getActiveSheet()->setCellValue('E11', '=SUM(E4:E9)');

$objXLSX->getActiveSheet()->setCellValue('D12', 'VAT:');
$objXLSX->getActiveSheet()->setCellValue('E12', '=E11*0.21');

$objXLSX->getActiveSheet()->setCellValue('D13', 'Total incl.:');
$objXLSX->getActiveSheet()->setCellValue('E13', '=E11+E12');

// Add comment
echo date('H:i:s') , " Add comments" , EOL;

$objXLSX->getActiveSheet()->getComment('E11')->setAuthor('PHPExcel');
$objCommentRichText = $objXLSX->getActiveSheet()->getComment('E11')->getText()->createTextRun('PHPExcel:');
$objCommentRichText->getFont()->setBold(true);
$objXLSX->getActiveSheet()->getComment('E11')->getText()->createTextRun("\r\n");
$objXLSX->getActiveSheet()->getComment('E11')->getText()->createTextRun('Total amount on the current invoice, excluding VAT.');

$objXLSX->getActiveSheet()->getComment('E12')->setAuthor('PHPExcel');
$objCommentRichText = $objXLSX->getActiveSheet()->getComment('E12')->getText()->createTextRun('PHPExcel:');
$objCommentRichText->getFont()->setBold(true);
$objXLSX->getActiveSheet()->getComment('E12')->getText()->createTextRun("\r\n");
$objXLSX->getActiveSheet()->getComment('E12')->getText()->createTextRun('Total amount of VAT on the current invoice.');

$objXLSX->getActiveSheet()->getComment('E13')->setAuthor('PHPExcel');
$objCommentRichText = $objXLSX->getActiveSheet()->getComment('E13')->getText()->createTextRun('PHPExcel:');
$objCommentRichText->getFont()->setBold(true);
$objXLSX->getActiveSheet()->getComment('E13')->getText()->createTextRun("\r\n");
$objXLSX->getActiveSheet()->getComment('E13')->getText()->createTextRun('Total amount on the current invoice, including VAT.');
$objXLSX->getActiveSheet()->getComment('E13')->setWidth('100pt');
$objXLSX->getActiveSheet()->getComment('E13')->setHeight('100pt');
$objXLSX->getActiveSheet()->getComment('E13')->setMarginLeft('150pt');
$objXLSX->getActiveSheet()->getComment('E13')->getFillColor()->setRGB('EEEEEE');


// Add rich-text string
echo date('H:i:s') , " Add rich-text string" , EOL;
$objRichText = new PHPExcel_RichText();
$objRichText->createText('This invoice is ');

$objPayable = $objRichText->createTextRun('payable within thirty days after the end of the month');
$objPayable->getFont()->setBold(true);
$objPayable->getFont()->setItalic(true);
$objPayable->getFont()->setColor( new PHPExcel_Style_Color( PHPExcel_Style_Color::COLOR_DARKGREEN ) );

$objRichText->createText(', unless specified otherwise on the invoice.');

$objXLSX->getActiveSheet()->getCell('A18')->setValue($objRichText);

// Merge cells
echo date('H:i:s') , " Merge cells" , EOL;
$objXLSX->getActiveSheet()->mergeCells('A18:E22');
$objXLSX->getActiveSheet()->mergeCells('A28:B28');		// Just to test...
$objXLSX->getActiveSheet()->unmergeCells('A28:B28');	// Just to test...

// Protect cells
echo date('H:i:s') , " Protect cells" , EOL;
$objXLSX->getActiveSheet()->getProtection()->setSheet(true);	// Needs to be set to true in order to enable any worksheet protection!
$objXLSX->getActiveSheet()->protectCells('A3:E13', 'PHPExcel');

// Set cell number formats
echo date('H:i:s') , " Set cell number formats" , EOL;
$objXLSX->getActiveSheet()->getStyle('E4:E13')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);

// Set column widths
echo date('H:i:s') , " Set column widths" , EOL;
$objXLSX->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objXLSX->getActiveSheet()->getColumnDimension('D')->setWidth(12);
$objXLSX->getActiveSheet()->getColumnDimension('E')->setWidth(12);

// Set fonts
echo date('H:i:s') , " Set fonts" , EOL;
$objXLSX->getActiveSheet()->getStyle('B1')->getFont()->setName('Candara');
$objXLSX->getActiveSheet()->getStyle('B1')->getFont()->setSize(20);
$objXLSX->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
$objXLSX->getActiveSheet()->getStyle('B1')->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_SINGLE);
$objXLSX->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);

$objXLSX->getActiveSheet()->getStyle('D1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objXLSX->getActiveSheet()->getStyle('E1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);

$objXLSX->getActiveSheet()->getStyle('D13')->getFont()->setBold(true);
$objXLSX->getActiveSheet()->getStyle('E13')->getFont()->setBold(true);

// Set alignments
echo date('H:i:s') , " Set alignments" , EOL;
$objXLSX->getActiveSheet()->getStyle('D11')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objXLSX->getActiveSheet()->getStyle('D12')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objXLSX->getActiveSheet()->getStyle('D13')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$objXLSX->getActiveSheet()->getStyle('A18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY);
$objXLSX->getActiveSheet()->getStyle('A18')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objXLSX->getActiveSheet()->getStyle('B5')->getAlignment()->setShrinkToFit(true);

// Set thin black border outline around column
echo date('H:i:s') , " Set thin black border outline around column" , EOL;
$styleThinBlackBorderOutline = array(
	'borders' => array(
		'outline' => array(
			'style' => PHPExcel_Style_Border::BORDER_THIN,
			'color' => array('argb' => 'FF000000'),
		),
	),
);
$objXLSX->getActiveSheet()->getStyle('A4:E10')->applyFromArray($styleThinBlackBorderOutline);


// Set thick brown border outline around "Total"
echo date('H:i:s') , " Set thick brown border outline around Total" , EOL;
$styleThickBrownBorderOutline = array(
	'borders' => array(
		'outline' => array(
			'style' => PHPExcel_Style_Border::BORDER_THICK,
			'color' => array('argb' => 'FF993300'),
		),
	),
);
$objXLSX->getActiveSheet()->getStyle('D13:E13')->applyFromArray($styleThickBrownBorderOutline);

// Set fills
echo date('H:i:s') , " Set fills" , EOL;
$objXLSX->getActiveSheet()->getStyle('A1:E1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objXLSX->getActiveSheet()->getStyle('A1:E1')->getFill()->getStartColor()->setARGB('FF808080');

// Set style for header row using alternative method
echo date('H:i:s') , " Set style for header row using alternative method" , EOL;
$objXLSX->getActiveSheet()->getStyle('A3:E3')->applyFromArray(
		array(
			'font'    => array(
				'bold'      => true
			),
			'alignment' => array(
				'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
			),
			'borders' => array(
				'top'     => array(
 					'style' => PHPExcel_Style_Border::BORDER_THIN
 				)
			),
			'fill' => array(
	 			'type'       => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
	  			'rotation'   => 90,
	 			'startcolor' => array(
	 				'argb' => 'FFA0A0A0'
	 			),
	 			'endcolor'   => array(
	 				'argb' => 'FFFFFFFF'
	 			)
	 		)
		)
);

$objXLSX->getActiveSheet()->getStyle('A3')->applyFromArray(
		array(
			'alignment' => array(
				'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
			),
			'borders' => array(
				'left'     => array(
 					'style' => PHPExcel_Style_Border::BORDER_THIN
 				)
			)
		)
);

$objXLSX->getActiveSheet()->getStyle('B3')->applyFromArray(
		array(
			'alignment' => array(
				'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
			)
		)
);

$objXLSX->getActiveSheet()->getStyle('E3')->applyFromArray(
		array(
			'borders' => array(
				'right'     => array(
 					'style' => PHPExcel_Style_Border::BORDER_THIN
 				)
			)
		)
);

// Unprotect a cell
echo date('H:i:s') , " Unprotect a cell" , EOL;
$objXLSX->getActiveSheet()->getStyle('B1')->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED);

// Add a hyperlink to the sheet
echo date('H:i:s') , " Add a hyperlink to an external website" , EOL;
$objXLSX->getActiveSheet()->setCellValue('E26', 'www.phpexcel.net');
$objXLSX->getActiveSheet()->getCell('E26')->getHyperlink()->setUrl('http://www.phpexcel.net');
$objXLSX->getActiveSheet()->getCell('E26')->getHyperlink()->setTooltip('Navigate to website');
$objXLSX->getActiveSheet()->getStyle('E26')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

echo date('H:i:s') , " Add a hyperlink to another cell on a different worksheet within the workbook" , EOL;
$objXLSX->getActiveSheet()->setCellValue('E27', 'Terms and conditions');
$objXLSX->getActiveSheet()->getCell('E27')->getHyperlink()->setUrl("sheet://'Terms and conditions'!A1");
$objXLSX->getActiveSheet()->getCell('E27')->getHyperlink()->setTooltip('Review terms and conditions');
$objXLSX->getActiveSheet()->getStyle('E27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

// Add a drawing to the worksheet
echo date('H:i:s') , " Add a drawing to the worksheet" , EOL;
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Logo');
$objDrawing->setDescription('Logo');
$objDrawing->setPath('./images/officelogo.jpg');
$objDrawing->setHeight(36);
$objDrawing->setWorksheet($objXLSX->getActiveSheet());

// Add a drawing to the worksheet
echo date('H:i:s') , " Add a drawing to the worksheet" , EOL;
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Paid');
$objDrawing->setDescription('Paid');
$objDrawing->setPath('./images/paid.png');
$objDrawing->setCoordinates('B15');
$objDrawing->setOffsetX(110);
$objDrawing->setRotation(25);
$objDrawing->getShadow()->setVisible(true);
$objDrawing->getShadow()->setDirection(45);
$objDrawing->setWorksheet($objXLSX->getActiveSheet());

// Add a drawing to the worksheet
echo date('H:i:s') , " Add a drawing to the worksheet" , EOL;
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('PHPExcel logo');
$objDrawing->setDescription('PHPExcel logo');
$objDrawing->setPath('./images/phpexcel_logo.gif');
$objDrawing->setHeight(36);
$objDrawing->setCoordinates('D24');
$objDrawing->setOffsetX(10);
$objDrawing->setWorksheet($objXLSX->getActiveSheet());

// Play around with inserting and removing rows and columns
echo date('H:i:s') , " Play around with inserting and removing rows and columns" , EOL;
$objXLSX->getActiveSheet()->insertNewRowBefore(6, 10);
$objXLSX->getActiveSheet()->removeRow(6, 10);
$objXLSX->getActiveSheet()->insertNewColumnBefore('E', 5);
$objXLSX->getActiveSheet()->removeColumn('E', 5);

// Set header and footer. When no different headers for odd/even are used, odd header is assumed.
echo date('H:i:s') , " Set header/footer" , EOL;
$objXLSX->getActiveSheet()->getHeaderFooter()->setOddHeader('&L&BInvoice&RPrinted on &D');
$objXLSX->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' . $objXLSX->getProperties()->getTitle() . '&RPage &P of &N');

// Set page orientation and size
echo date('H:i:s') , " Set page orientation and size" , EOL;
$objXLSX->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
$objXLSX->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

// Rename first worksheet
echo date('H:i:s') , " Rename first worksheet" , EOL;
$objXLSX->getActiveSheet()->setTitle('Invoice');


// Create a new worksheet, after the default sheet
echo date('H:i:s') , " Create a second Worksheet object" , EOL;
$objXLSX->createSheet();

// Llorem ipsum...
$sLloremIpsum = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Vivamus eget ante. Sed cursus nunc semper tortor. Aliquam luctus purus non elit. Fusce vel elit commodo sapien dignissim dignissim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Curabitur accumsan magna sed massa. Nullam bibendum quam ac ipsum. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Proin augue. Praesent malesuada justo sed orci. Pellentesque lacus ligula, sodales quis, ultricies a, ultricies vitae, elit. Sed luctus consectetuer dolor. Vivamus vel sem ut nisi sodales accumsan. Nunc et felis. Suspendisse semper viverra odio. Morbi at odio. Integer a orci a purus venenatis molestie. Nam mattis. Praesent rhoncus, nisi vel mattis auctor, neque nisi faucibus sem, non dapibus elit pede ac nisl. Cras turpis.';

// Add some data to the second sheet, resembling some different data types
echo date('H:i:s') , " Add some data" , EOL;
$objXLSX->setActiveSheetIndex(1);
$objXLSX->getActiveSheet()->setCellValue('A1', 'Terms and conditions');
$objXLSX->getActiveSheet()->setCellValue('A3', $sLloremIpsum);
$objXLSX->getActiveSheet()->setCellValue('A4', $sLloremIpsum);
$objXLSX->getActiveSheet()->setCellValue('A5', $sLloremIpsum);
$objXLSX->getActiveSheet()->setCellValue('A6', $sLloremIpsum);

// Set the worksheet tab color
echo date('H:i:s') , " Set the worksheet tab color" , EOL;
$objXLSX->getActiveSheet()->getTabColor()->setARGB('FF0094FF');;

// Set alignments
echo date('H:i:s') , " Set alignments" , EOL;
$objXLSX->getActiveSheet()->getStyle('A3:A6')->getAlignment()->setWrapText(true);

// Set column widths
echo date('H:i:s') , " Set column widths" , EOL;
$objXLSX->getActiveSheet()->getColumnDimension('A')->setWidth(80);

// Set fonts
echo date('H:i:s') , " Set fonts" , EOL;
$objXLSX->getActiveSheet()->getStyle('A1')->getFont()->setName('Candara');
$objXLSX->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);
$objXLSX->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
$objXLSX->getActiveSheet()->getStyle('A1')->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_SINGLE);

$objXLSX->getActiveSheet()->getStyle('A3:A6')->getFont()->setSize(8);

// Add a drawing to the worksheet
echo date('H:i:s') , " Add a drawing to the worksheet" , EOL;
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Terms and conditions');
$objDrawing->setDescription('Terms and conditions');
$objDrawing->setPath('./images/termsconditions.jpg');
$objDrawing->setCoordinates('B14');
$objDrawing->setWorksheet($objXLSX->getActiveSheet());

// Set page orientation and size
echo date('H:i:s') , " Set page orientation and size" , EOL;
$objXLSX->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objXLSX->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

// Rename second worksheet
echo date('H:i:s') , " Rename second worksheet" , EOL;
$objXLSX->getActiveSheet()->setTitle('Terms and conditions');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objXLSX->setActiveSheetIndex(0);
