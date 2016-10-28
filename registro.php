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
$objPHPExcel->getProperties()->setCreator("Soledad Gutierrez")
        ->setLastModifiedBy("Soledad Gutierrez")
        ->setTitle("Prueba Excel PHPCentral")
        ->setSubject("Prueba Excel PHPCentral")
        ->setDescription("Este es un documento de prueba excel para testear la librería PHPExcel.");

//Agrego la cabecera
$objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', 'Comida')
        ->setCellValue('B1', 'Tipo')
        ->setCellValue('C1', 'Precio');

//Invento un array con info (podría venir de una base de datos):
$alimentos = array();
$alimentos[] = array('comida'=>'Arroz chaufa', 'tipo'=>'Comida peruana', 'precio'=>20.00);
$alimentos[] = array('comida'=>'Ceviche', 'tipo'=>'Comida típica', 'precio'=>50.00);
$alimentos[] = array('comida'=>'Arroz con pollo', 'tipo'=>'Comida tradicional', 'precio'=>15.00);

//Agrego información al excel en base al array inventado:
foreach ($alimentos as $nroRegistro => $alimento){
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.($nroRegistro+2), $alimento['comida'])
        ->setCellValue('B'.($nroRegistro+2), $alimento['tipo'])
        ->setCellValue('C'.($nroRegistro+2), $alimento['precio']);
}

//Nombro la hoja del excel:
$objPHPExcel->getActiveSheet()->setTitle('Precios de comida');


//Seteo que la hoja 0 sea la primera hoja que se muestre al abrir el documento (podrías haber más de una hoja y hacer que se muestre otra hoja en vez de la primera)
$objPHPExcel->setActiveSheetIndex(0);

// Guardo el documento en formato Excel 2007 (*.xlsx)
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('archivo_soledad_excel_2007.xlsx');


// Guardo el documento en formato Excel 97 (*.xls)
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('archivo_soledad_excel_97.xls');

?>