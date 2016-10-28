<?php

//agrego algunas líneas para mostrar todos los errores (ya que esto es una prueba y quiero ver si algo falla)

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);


//incluyo la librería PHPExcel (ojo con el path)
require_once 'Classes/PHPExcel.php';

//creo el objeto:
$objPHPExcel = new PHPExcel();

//configuro las propiedades del documento
$objPHPExcel->getProperties()->setCreator("Fernando Mosquera")
        ->setLastModifiedBy("Fernando Mosquera")
        ->setTitle("Prueba Excel PHPCentral")
        ->setSubject("Prueba Excel PHPCentral")
        ->setDescription("Este es un documento de prueba excel para testear la librería PHPExcel.");

//Agrego la cabecera
$objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', 'Marca')
        ->setCellValue('B1', 'Modelo')
        ->setCellValue('C1', 'Precio Lista');

//Invento un array con info (podría venir de una base de datos):
$autos = array();
$autos[] = array('marca'=>'Suzuki', 'modelo'=>'Suzuki Swift 1.3 GL', 'precio'=>14290.00);
$autos[] = array('marca'=>'Chevrolet', 'modelo'=>'Chevrolet Cruze 1.8 AT', 'precio'=>22740.00);
$autos[] = array('marca'=>'Mazda', 'modelo'=>'Mazda 3MT 1.6 4x2 GS STD', 'precio'=>20990.00);

//Agrego información al excel en base al array inventado:
foreach ($autos as $nroRegistro => $auto){
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.($nroRegistro+2), $auto['marca'])
        ->setCellValue('B'.($nroRegistro+2), $auto['modelo'])
        ->setCellValue('C'.($nroRegistro+2), $auto['precio']);
}

//Nombro la hoja del excel:
$objPHPExcel->getActiveSheet()->setTitle('Precios de autos');


//Seteo que la hoja 0 sea la primera hoja que se muestre al abrir el documento (podrías haber más de una hoja y hacer que se muestre otra hoja en vez de la primera)
$objPHPExcel->setActiveSheetIndex(0);

// Guardo el documento en formato Excel 2007 (*.xlsx)
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('archivo_excel_2007.xlsx');


// Guardo el documento en formato Excel 97 (*.xls)
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('archivo_excel_97.xls');

?>