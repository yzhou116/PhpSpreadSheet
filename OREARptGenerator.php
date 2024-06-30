<?php
require_once './vendor/autoload.php';
require_once './OREARptClass.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Style;

/*
When you try to test the code, please update text file rpt.orea.mem.2022-03-30.oreaMthRpt.txt and rpt.orea.pub.2022-03-30.oreaMthRpt
at the today's date, for example, if today is 2022-03-31, then change the name of these two files to 
rpt.orea.mem.2022-03-31.oreaMthRpt.txt and rpt.orea.pub.2022-03-31.oreaMthRpt. 
https://green.korahlimited.com/ccrChat/ccrOrea/bin/tmp/
Thanks. 
*/


$spreadSheet = new Spreadsheet();
$sheet = $spreadSheet->getActiveSheet()->setTitle("Categories");;
$oreaObj = new OREARpt();
$oreaPubObj = new OREARpt();

$sheet->setCellValue('A1', 'Instance');
$sheet->setCellValue('A2', 'OREA Member');
$sheet->setCellValue('B1', 'Category');

$oreaObj->setTitleStyle($spreadSheet,'A1:C1');
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(1)->setAutoSize(false);
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth('30');
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(2)->setAutoSize(true);
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(3)->setAutoSize(true);
$spreadSheet->getActiveSheet()->getStyle('A1:C1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$now = new \DateTime('now');
$month = $now->format('M');
$year = $now->format('Y');
$currentDate = date("Y-m-d");
$sheet->setCellValue('C1', 'Category Hits in '.$month.' ,'.$year);
//$intentions = new Spreadsheet();

//read data from text file

$orea = $oreaObj->readTxtFile('rpt.orea.mem.'.$currentDate.'.oreaMthRpt.txt');
$oreaPub = $oreaPubObj->readTxtFile('rpt.orea.pub.'.$currentDate.'.oreaMthRpt.txt');
$oreaPubCate = $oreaPubObj->readSpecificData('Category');
$oreaPubIntention = $oreaPubObj->readSpecificData('Intention');
$oreaCate = $oreaObj->readSpecificData("Category");
$oreaIntention = $oreaObj->readSpecificData("Intention");
$oreaStandardFrom = $oreaObj->readSpecificData("Standard Forms Explained");
$oreaTopicInDetial = $oreaObj->readSpecificData("Topics in detail");
$oreaFormGeneral = $oreaObj->readSpecificData("Standard Forms General");
$FormsWO = $oreaObj->readSpecificData("Forms w o");
$GTCount = $oreaObj->readSpecificData("GTCount");
$ETCount = $oreaObj->readSpecificData("ETCount");
$TDCount = $oreaObj->readSpecificData("TDCount");
$SFTCount = $oreaObj->readSpecificData("SFTCount");
$countCateoty = 0;
//start to convert category data to array and write to excel
$oreaPubCate = $oreaPubObj->removeVerticalBar($oreaPubCate );
$oreaCate = $oreaObj->removeVerticalBar($oreaCate);
$oreaStandardFrom = $oreaObj->removeVerticalBar($oreaStandardFrom);
$oreaTopicInDetial = $oreaObj->removeVerticalBar($oreaTopicInDetial);
$oreaFormGeneral = $oreaObj->removeVerticalBar($oreaFormGeneral);
$FormsWO = $oreaObj->removeVerticalBar($FormsWO);
$oreaStandardFromNum = $oreaObj->getNumber($oreaStandardFrom);
$oreaTopicInDetailNum = $oreaObj->getNumber($oreaTopicInDetial);
$oreaFormGeneralNum = $oreaObj->getNumber($oreaFormGeneral);
$FormsWONum = $oreaObj->getNumber($FormsWO);
$finalRow = $oreaObj->writeCategory($sheet,$oreaCate);
$cateFromArr = [$oreaStandardFromNum,$oreaTopicInDetailNum,$oreaFormGeneralNum,$FormsWONum];
$finalRow = $oreaObj->writeCategoryForms($sheet,$finalRow, $cateFromArr);
$spreadSheet->getActiveSheet()->mergeCells('A2:A'.strval($finalRow -1));
$countOreaM = $finalRow;
$sheet->setCellValue('A'.strval($countOreaM), 'OREA Pub');
//var_dump("This is final row" + strval($finalRow));
$oreaObj->setStyleOreaType($spreadSheet,'A2:A'.strval($finalRow -1), 'member');
$spreadSheet->getActiveSheet()->getStyle('A2:A'.strval($finalRow -1))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
$finalRow = $oreaPubObj->writeCategory($sheet,$oreaPubCate, $finalRow);


$spreadSheet->getActiveSheet()->mergeCells('A'.strval($countOreaM).':A'.strval($finalRow));

$oreaObj->setStyleOreaType($spreadSheet,'A'.strval($countOreaM).':A'.strval($finalRow ), "Pub");
$spreadSheet->getActiveSheet()->getStyle('A'.strval($countOreaM).':A'.strval($finalRow ))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

$spreadSheet->getActiveSheet()->getStyle('C2:C'.strval($finalRow))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$oreaObj->setTextStyle($spreadSheet, 'B2:C'.strval($finalRow));

//start to convert intention data to array and write to excel
$sheet = $spreadSheet->createSheet();
$spreadSheet->setActiveSheetIndex(1);
$sheet->setCellValue('A2', 'OREA Member');

$oreaObj->setTitleStyle($spreadSheet, 'A1:D1');
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(1)->setAutoSize(false);
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth('30');
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(2)->setAutoSize(true);
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(3)->setAutoSize(false);
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(3)->setWidth('60');
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(4)->setAutoSize(false);
$spreadSheet->getActiveSheet()->getColumnDimensionByColumn(4)->setWidth('53');
$spreadSheet->getActiveSheet()->getStyle('A1:D1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$sheet->setTitle('Intentions');
$sheet->setCellValue('A1', 'Instance');
$sheet->setCellValue('B1', 'Type');
$sheet->setCellValue('C1', 'Intention');
$sheet->setCellValue('D1', 'Intention Hits in '.$month.' ,'.$year);



$oreaPubIntention = $oreaPubObj->removeVerticalBar($oreaPubIntention );
$oreaIntention = $oreaObj->removeVerticalBar($oreaIntention);

$finalRow = $oreaObj->writeIntention($sheet,$oreaIntention);

$spreadSheet->getActiveSheet()->mergeCells('A2:A'.strval($finalRow));
$countOreaM = $finalRow+1;
$sheet->setCellValue('A'.strval($countOreaM), 'OREA Pub');


$oreaObj->setStyleOreaType($spreadSheet,'A2:A'.strval($finalRow -1), 'member');
$spreadSheet->getActiveSheet()->getStyle('A2:A'.strval($finalRow -1))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);


$finalRow = $oreaPubObj->writeIntention($sheet,$oreaPubIntention, $finalRow+1);
//var_dump("This is final row" + strval($finalRow));
$spreadSheet->getActiveSheet()->mergeCells('A'.strval($countOreaM).':A'.strval($finalRow));


$oreaObj->setStyleOreaType($spreadSheet,'A'.strval($countOreaM).':A'.strval($finalRow ), "Pub");
$spreadSheet->getActiveSheet()->getStyle('A'.strval($countOreaM).':A'.strval($finalRow ))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);


$spreadSheet->getActiveSheet()->getStyle('D2:D'.strval($finalRow))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$spreadSheet->getActiveSheet()->getStyle('B2:B'.strval($finalRow))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

$oreaObj->setTextStyle($spreadSheet,'B2:D'.strval($finalRow));
$finalRow  = $finalRow+1;
//add data count to intention table
$GTCount = $oreaObj->removeVerticalBar($GTCount);
$ETCount = $oreaObj->removeVerticalBar($ETCount);
$TDCount = $oreaObj->removeVerticalBar($TDCount);
$SFTCount = $oreaObj->removeVerticalBar($SFTCount);
$GTCount  = $GTCount[0][1];
$ETCount  = $ETCount[0][1];
$TDCount  = $TDCount[0][1];
$SFTCount  = $SFTCount[0][1];
$dataCountArry = [$GTCount,$ETCount,$TDCount,$SFTCount];
//$dataCountArry = [$GTCount,$ETCount,$TDCount,0 ];
$oreaObj->addDataCount($sheet, $finalRow, $dataCountArry,$spreadSheet);

//add forms details
$sheet = $spreadSheet->createSheet();
$spreadSheet->setActiveSheetIndex(2);

$sheet->setTitle('Standard Forms Explained');
$oreaObj->addFormSheet($sheet,$oreaStandardFrom,$spreadSheet);


$sheet = $spreadSheet->createSheet();
$spreadSheet->setActiveSheetIndex(3);
$sheet->setTitle('Standard Forms-Topic in detail');
$oreaObj->addFormSheet($sheet,$oreaTopicInDetial,$spreadSheet);


$sheet = $spreadSheet->createSheet();
$spreadSheet->setActiveSheetIndex(4);
$sheet->setTitle('Standard Forms - General');
$oreaObj->addFormSheet($sheet,$oreaFormGeneral,$spreadSheet);


//$sheet = new Worksheet();
$sheet = $spreadSheet->createSheet();
$spreadSheet->setActiveSheetIndex(5);
$sheet->setTitle('Forms w o Forms Explained File');
$oreaObj->addFormSheet($sheet,$FormsWO,$spreadSheet);



//save xlsx file
$currentDate = date("Ymd");
$prevDate = date("Ymd",strtotime("-1 month"));
print_r($countOreaM) ;
$writer = new Xlsx($spreadSheet);
$writer->save('_OREA ccR - Statistics '.$prevDate.' - '.$currentDate.'.xlsx');