<?php
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Style;

class OREARpt{

  
   private $OreaArray;
function __construct() {
    $this->OreaArray  = [];
}
public  function readTxtFile($path){
    
    $handle = fopen($path, "r");
    if ($handle) {
        while (($line = fgets($handle)) !== false) {
           array_push( $this->OreaArray,$line );
        }
    
        fclose($handle);
    } else {
        return false;
    }
    return  $this->OreaArray;
   }
   public function readSpecificData($info){
       $countPlus = 0;
       if($info=="Category"){
        $countPlus = 8;
       }else if($info=="Intention"){
        $countPlus = 5;
       }else if($info=="Standard Forms Explained"){
        $countPlus = 14;
       }else if($info=="Topics in detail"){
        $countPlus = 20;
       }else if($info=="Standard Forms General"){
        $countPlus = 26;
       }
       else if($info=="Forms w o"){
        $countPlus = 32;
       }
       else if($info=="GTCount"){
        $countPlus = 41;
       }
       else if($info=="ETCount"){
        $countPlus = 44;
       }
       else if($info=="TDCount"){
        $countPlus = 47;
       }
       else if($info=="SFTCount"){
        $countPlus = 50;
       }
       $counter = 0;
       $oreaCate = [];
       if(sizeof($this->OreaArray) > 0){
       
        for($i = 0; $i < sizeof($this->OreaArray); $i++){
            if($this->OreaArray[$i][0] == '+'){
                $counter++;
            }
            if($counter == $countPlus){
                for($j = $i + 1; $j < sizeof($this->OreaArray); $j++){
                    if($this->OreaArray[$j][0] != '+'){
                        array_push($oreaCate,$this->OreaArray[$j] );
                    }else{
                        break;
                    }
                   
                }
                break;
            }
        }
        return $oreaCate;
       }else{
           return false;
       }
   }
   public function removeVerticalBar($array){
       $res = [];
       for($i = 0; $i < sizeof($array); $i++){
        $res[$i] = preg_split("/\|/" ,$array[$i]);
       }
       return $res;

   }
   public function getNumber($array){
       $totalCount = 0;
       for($i = 0; $i < sizeof($array); $i++){
           $tempNum = intval(trim($array[$i][3]));
           $totalCount = $totalCount +  $tempNum;
       };
       return $totalCount;

   }
   public function writeCategory($sheet, $array, $row = 0){
       $finalRow = 0;
       if($row == 0 ){
        for($i = 0; $i < sizeof($array); $i++){
            $hitTimes = round(floatval($array[$i][3]));
          // $hitTimes = round(floatval($array[$i][2]));
            if( $hitTimes > 0){
                $sheet->setCellValue('B'.strval($i+2),$array[$i][2]);
                $sheet->setCellValue('C'.strval($i+2),$hitTimes);
                if($finalRow < $i+2){
                    $finalRow = $i + 2;
                }
                
            }
           }
       }else{
        for($i = 0; $i < sizeof($array); $i++){
            $hitTimes = round(floatval($array[$i][3]));
       //   $hitTimes = round(floatval($array[$i][2]));
            if( $hitTimes > 0){
                $sheet->setCellValue('B'.strval($row),$array[$i][2]);
                $sheet->setCellValue('C'.strval($row),$hitTimes);
                if($finalRow < $row){
                    $finalRow = $row;
                }
                $row++;
                
            }
                
           }
       }
    
       return $finalRow;
   }
   public function writeCategoryForms($sheet, $row, $array){
       $row++;
       $sheet->setCellValue('B'. strval($row), 'Standard Forms Explained');
       $sheet->setCellValue('C'. strval($row), $array[0]);
       $row++;
       $sheet->setCellValue('B'. strval($row), 'Standard Forms - Topic in details');
       $sheet->setCellValue('C'. strval($row), $array[1]);
       $row++;
       $sheet->setCellValue('B'. strval($row), 'Standard Forms - General');
       $sheet->setCellValue('C'. strval($row), $array[2]);
       return $row+1;
   }
   public function writeIntention($sheet, $array, $row = 0){
    $finalRow = 0;
   
    if($row == 0 ){
      
     for($i = 0; $i < sizeof($array); $i++){
       
         $hitTimes = round(floatval($array[$i][3]));
       // $hitTimes = round(floatval($array[$i][2]));
        
         if( $hitTimes > 0){
             $sheet->setCellValue('B'.strval($i+2),$array[$i][1]);
             $sheet->setCellValue('C'.strval($i+2),$array[$i][2]);
             $sheet->setCellValue('D'.strval($i+2),$hitTimes);
             if($finalRow < $i+2){
                 $finalRow = $i + 2;
             }
             
         }
        }
    }else{
     for($i = 0; $i < sizeof($array); $i++){
      
         $hitTimes = round(floatval($array[$i][3]));
     // $hitTimes = round(floatval($array[$i][2]));
         if( $hitTimes > 0){
             $sheet->setCellValue('B'.strval($row),$array[$i][1]);
             $sheet->setCellValue('C'.strval($row),$array[$i][2]);
             $sheet->setCellValue('D'.strval($row),$hitTimes);
             if($finalRow < $row){
                 $finalRow = $row;
             }
             $row++;
             
         }
             
        }
       
    }
  //  var_dump("This is final row" + $finalRow);
   
  
 
    return $finalRow;
}
public function addDataCount($sheet, $row, $array,$spreadSheet){
 
    $this->setTitleStyle( $spreadSheet,'A'.$row.':D'.$row );
    $now = new \DateTime('now');
    $now2 = new \DateTime('now');
    $fullmonth = $now2->format('F');
    $month = $now->format('M');
    $year = $now->format('Y');
    $day = $now->format('D');
    $firstRow = $row + 1;
    $sheet->setCellValue('A'. strval($row),'#');
    $sheet->setCellValue('C'. strval($row),'Criteria');
    $sheet->setCellValue('D'. strval($row),'# of Training Data as per the criteria '.$fullmonth.' ,'.$year);
    $row++;
    $sheet->setCellValue('A'. strval($row),'1');
    $sheet->setCellValue('C'. strval($row),"Total number of training data added on the bot for all the intentions in the category 'Standard Forms - General' as of ". $month.' '.$day. ' ,'.$year);
    $spreadSheet->getActiveSheet()->getStyle('C'. strval($row))->getAlignment()->setWrapText(true);
    $sheet->setCellValue('D'. strval($row),$array[0]);
    $row++;
    $sheet->setCellValue('A'. strval($row),'2');
    $sheet->setCellValue('C'. strval($row),"Total number of training data added on the bot for all the intentions in the category 'Standard Forms - Expliained' as of ". $month.' '.$day. ' ,'.$year);
    $spreadSheet->getActiveSheet()->getStyle('C'. strval($row))->getAlignment()->setWrapText(true);
    $sheet->setCellValue('D'. strval($row),$array[1]);
    $row++;
    $sheet->setCellValue('A'. strval($row),'3');
    $sheet->setCellValue('C'. strval($row),"Total number of training data added on the bot for all the intentions in the category 'OREA Chat Dialog Tree >> 02-Standard Forms' as of ". $month.' '.$day. ' ,'.$year);
    $spreadSheet->getActiveSheet()->getStyle('C'. strval($row))->getAlignment()->setWrapText(true);
    $sheet->setCellValue('D'. strval($row),$array[3]);
    $row++;
    $sheet->setCellValue('A'. strval($row),'4');
    $sheet->setCellValue('C'. strval($row),"Total number of training data added on the bot for all the intentions in the category 'Standard Forms - Topics in Details' as of ". $month.' '.$day. ' ,'.$year);
    $spreadSheet->getActiveSheet()->getStyle('C'. strval($row))->getAlignment()->setWrapText(true);
    $sheet->setCellValue('D'. strval($row),$array[2]);

    $spreadSheet->getActiveSheet()->getStyle('A'.$firstRow .':A'.strval($row))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);;
    $spreadSheet->getActiveSheet()->getStyle('D'.$firstRow .':D'.strval($row))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);;


    $this->setTextStyle($spreadSheet, 'B'.$firstRow .':B'.strval($row));
    $this->setTextStyle($spreadSheet, 'C'.$firstRow .':C'.strval($row));
    $this->setTextStyle($spreadSheet, 'D'.$firstRow .':D'.strval($row));

}
public function addFormSheet($sheet, $array,$spreadSheet){
    $styleText = array(
        'font'  => array(
            'size'  => 12,
            'name'  => 'Time New Roman'
        )
    );
    for($i = 0; $i < sizeof($array); $i++){
        $sheet->setCellValue('A'.strval($i + 1),trim($array[$i][2]));
        $sheet->setCellValue('B'.strval($i + 1),trim($array[$i][3]));
    }
    $spreadSheet->getActiveSheet()->getStyle('A1:B'.strval(sizeof($array)))->applyFromArray($styleText);
    $spreadSheet->getActiveSheet()->getColumnDimensionByColumn(1)->setAutoSize(true);
    $spreadSheet->getActiveSheet()->getColumnDimensionByColumn(2)->setAutoSize(false);
    $spreadSheet->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth('15');
    $spreadSheet->getActiveSheet()->getStyle('B1:B'.strval(sizeof($array)))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
}
public function setTextStyle($spreadSheet, $range){
    $styleText = array(
        'font'  => array(
            'size'  => 12,
            'name'  => 'Time New Roman'
        ),
        'borders' => array(
            'inside' => array(
                'borderStyle' => Border::BORDER_THIN,
                'color' => array('argb' => '#fcba03'),
            ),
            'outline' => array(
                'borderStyle' => Border::BORDER_THIN,
                'color' => array('argb' => '#fcba03'),
            ),
        ),
    );
    $spreadSheet->getActiveSheet()->getStyle($range)->applyFromArray($styleText);
}
public function setTitleStyle($spreadSheet, $range){
    $styleTitle = array(
        'font'  => array(
            'bold'  => true,
            'size'  => 14,
            'name'  => 'Time New Roman'
        ),
        'borders' => array(
            'inside' => array(
                'borderStyle' => Border::BORDER_THIN,
                'color' => array('argb' => '#fcba03'),
            ),
            'outline' => array(
                'borderStyle' => Border::BORDER_THIN,
                'color' => array('argb' => '#fcba03'),
            ),
        ),
        'fill' => array(
            'fillType' => Fill::FILL_SOLID,
            'startColor' => array('rgb' => 'FFD700')
        )
    );
    $spreadSheet->getActiveSheet()->getStyle($range)->applyFromArray($styleTitle);
}
public function setStyleOreaType($spreadSheet, $range, $type){
    $styleOrea = [];
    if(strcmp($type, "Pub") === 0){
        $styleOrea = array(
            'font'  => array(
                'bold'  => true,
                'size'  => 14,
                'name'  => 'Time New Roman'
            ),
            'borders' => array(
                'inside' => array(
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => array('argb' => '#fcba03'),
                ),
                'outline' => array(
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => array('argb' => '#fcba03'),
                ),
            ),
            'fill' => array(
                'fillType' => Fill::FILL_SOLID,
                'startColor' => array('rgb' => '89CFF0')
            )
        );
    }else{
        $styleOrea  = array(
            'font'  => array(
                'bold'  => true,
                'size'  => 14,
                'name'  => 'Time New Roman'
            ),
            'borders' => array(
                'inside' => array(
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => array('argb' => '#fcba03'),
                ),
                'outline' => array(
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => array('argb' => '#fcba03'),
                ),
            ),
            'fill' => array(
                'fillType' => Fill::FILL_SOLID,
                'startColor' => array('rgb' => '0096FF')
            )
        );
    }
     $spreadSheet->getActiveSheet()->getStyle($range)->applyFromArray($styleOrea);
}

}