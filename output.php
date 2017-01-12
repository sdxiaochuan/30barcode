<?php
//var_dump($_FILES);

define('BARCODE', true);
//autoload
include("barcode_lib.php");

class Barcode
{
   public $number;
   public $encoding;
   public $scale;

   protected $_encoder;

   function __construct($encoding, $number=null, $scale=null)
   {
      $this->number = ($number==null) ? $this->_random() : $number;
      $this->scale = ($scale==null || $scale<4) ? 4 : $scale;

      // Reflection Class : Method

      $this->_encoder = new EAN13($this->number, $this->scale);
      $this->_encoder->display();
   }


   private function _random()
   {
     return substr(number_format(time() * rand(),0,'',''),0,12);
   }
}


if ($_POST) {
    
    include_once './PHPExcel_1.8.0/Classes/PHPExcel.php';
    
    $objPHPExcel = new PHPExcel();
    
    $rendererName = PHPExcel_Settings::PDF_RENDERER_TCPDF;
    $rendererLibraryPath = './TCPDF-master';
    
    if (!PHPExcel_Settings::setPdfRenderer(
        $rendererName,
        $rendererLibraryPath
    )) {
        die(
            'NOTICE: Please set the $rendererName and $rendererLibraryPath values' .
            EOL .
            'at the top of this script as appropriate for your directory structure'
        );
    }
    
    $objPHPExcel->getActiveSheet()->getPageMargins()->setTop(0);
    $objPHPExcel->getActiveSheet()->getPageMargins()->setRight(0);
    $objPHPExcel->getActiveSheet()->getPageMargins()->setLeft(0);
    $objPHPExcel->getActiveSheet()->getPageMargins()->setBottom(0);
    
    // $folder = $_POST['folder'];
    $products = count($_POST['ean']);
    
    $units = array("A1", "B1", "C1",
                   "A2", "B2", "C2",
                   "A3", "B3", "C3",
                   "A4", "B4", "C4",
                   "A5", "B5", "C5",
                   "A6", "B6", "C6",
                   "A7", "B7", "C7",
                   "A8", "B8", "C8",
                   "A9", "B9", "C9",
                   "A10", "B10", "C10");
    $current = 0;

    $objPHPExcel->setActiveSheetIndex(0);
    
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(32);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(32);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(32);
    
    for ($n = 1; $n <= 10; $n++) {
        $objPHPExcel->getActiveSheet()->getRowDimension($n)->setRowHeight(72.7);
    }
    
    
    
    for ($n = 0; $n < $products; $n ++) {
        
        if (isset($_POST['ean'][$n]) && !empty($_POST['ean'][$n])) {
            //$image =  $_FILES["barcode"]["tmp_name"][$n];

            //Create png files
            new Barcode("EAN-13", $_POST['ean'][$n]);

            $image = "./tmp/" . $_POST['ean'][$n] . ".png";
//            $check = getimagesize($_FILES["barcode"][$n]["tmp_name"]);
//            if($check !== false) {
//                echo "File is an image - " . $check["mime"] . ".";
//                $uploadOk = 1;
//            } else {
//                echo "File is not an image.";
//                $uploadOk = 0;
//            }
            
            if (!file_exists($image)) {
                echo $image;
                exit;
                continue;
            }
            
            $name = $_POST['name'][$n];
            $name = strlen($name) > 24 ? substr($name, 0, 24) : $name;
            
//            exit;
            for ($count = 0; $count < $_POST['count'][$n]; $count++) {
                
                $unit = $units[$current++];
                
                $objDrawing = new PHPExcel_Worksheet_Drawing();
                $objDrawing->setPath($image);
                $objDrawing->setResizeProportional(false);
                $objDrawing->setHeight(55);
                $objDrawing->setWidth(210);
                $objDrawing->setOffsetX(0);
                $objDrawing->setOffsetY(10);
                $objDrawing->setCoordinates($unit);
                $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
                
                $objPHPExcel->getActiveSheet()->setCellValue($unit, $name);
                
                $objPHPExcel->getActiveSheet()->getStyle($unit)->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getStyle($unit)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $objPHPExcel->getActiveSheet()->getStyle($unit)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM);
                        
            }
        }
        
    }
    
    $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_LETTER);
    
    $margin = $objPHPExcel->getActiveSheet()->getPageMargins();
    
    
    /*********************************************************************************************
     *  IMPORTANT: top=bottom=left=right=0.25, header=0.2, bottom=0
     *********************************************************************************************/
    //set margin -- Not working
    $margin->setTop(0.25)->setBottom(0.25)->setLeft(0.25)->setRight(0.25)->setHeader(0.2)->setBottom(0);
    
    //set border -- Not working
    $objPHPExcel->getActiveSheet()->getStyle('A1:C10')->applyFromArray(
        array(
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_NONE
                )
            )
        ));
    
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="01simple.xls"');  
    header('Cache-Control: max-age=0');  
   
    $objWriter = PHPExcel_IOFactory:: createWriter($objPHPExcel, 'Excel5');
    $objWriter->save( 'php://output');
    exit;
    
//    header('Content-Type: application/pdf');  
//    header('Content-Disposition: attachment;filename="01simple.pdf"');  
//    header('Cache-Control: max-age=0');  
//    
//    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');
//    $objWriter->writeAllSheets();
//    $objWriter->save('php://output');

//    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');  
//    $objWriter->save('/tmp/output.pdf'); 
    
    exit;
    
}
?>