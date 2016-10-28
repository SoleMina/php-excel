<?php

require_once 'Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();


$objPHPExcel->
	getProperties()
		->setCreator("TEDnologia.com")
		->setLastModifiedBy("TEDnologia.com")
		->setTitle("Exportar Excel con PHP")
		->setSubject("Documento de prueba")
		->setDescription("Documento generado con PHPExcel")
		->setKeywords("usuarios phpexcel")
		->setCategory("reportes");

$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Nombre')
            ->setCellValue('B1', 'E-mail')
            ->setCellValue('C1', 'Twitter')
            ->setCellValue('D1', 'Edad')
            ->setCellValue('A2', 'Karina')
            ->setCellValue('B2', 'pradokspg@gmail.com')
            ->setCellValue('C2', '@karinaprado')
            ->setCellValue('D2', 23);

$objPHPExcel->getActiveSheet()->setTitle('Usuarios');
$objPHPExcel->setActiveSheetIndex(0);

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="01simple.xls"');
header('Cache-Control: max-age=0');


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;

?>