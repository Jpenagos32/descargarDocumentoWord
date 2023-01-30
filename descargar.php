<?php

use PhpOffice\PhpWord\Style\Language;
require "vendor/autoload.php";

//Crear el nuevo documento
$phpWord = new \PhpOffice\PhpWord\PhpWord();


// Agregar una seccion vacia al documento
$seccion = $phpWord->addSection();

// Configuraciones por defecto del documento
//Añadir fuente personalizada o deseada
$fuente = 'Arial11';
$phpWord->addFontStyle(
    $fuente, ['name' => 'Arial', 'size' => 11]
);

//Centrar texto
$centrado = 'miEstilo';
$phpWord->addParagraphStyle($centrado, ['align' => 'center']);

// Añadir el lenguaje español al documento
$phpWord->getSettings()->setThemeFontLang(new Language(Language::ES_ES));


// Secciones del documento
//Añadir los titulos del documento 
$seccion->addText(
    "Acueducto y Alcantarillado de Popayán S.A. E. S.P",
    $fuente, $centrado
);


$seccion->addText('Empresa de Servicios Públicos',
    $fuente, $centrado
);

$seccion->addText('NIT 891.500.117-1',
    $fuente, $centrado
);

$seccion->addText(
    'Listado de Toma de Lecturas Mes: 2023-02 Ciclo: 4',
    ['name' => 'Arial', 'size' => 11, 'bold' => true], $centrado
);

// Crear estilo de tabla
$estiloTabla = ['borderColor' => '000000', 'borderSize' => 1, 'cellMargin' => 10];
$phpWord->addTableStyle('estilo', $estiloTabla);

// Agregar array con datos provisionales
$datosPersonas = array(
    array('Camilo Rodriguez 1', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 2', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 3', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 4', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 5', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 6', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 7', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 8', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 9', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 10', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 11', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 12', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 13', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 14', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 15', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 16', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 17', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 18', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 19', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 20', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 21', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 22', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 23', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 24', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 25', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 26', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 27', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 28', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 29', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 30', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 31', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 32', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 33', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 34', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 35', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 36', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 37', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 38', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 39', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______'),
    array('Camilo Rodriguez 40', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_______')
);

//Crear en formato columnas de word


// Crear tabla
/* $tabla = $seccion->addTable('estilo');
$tabla->addRow();
$celda = $tabla->addCell();
$celda->addText("Nombre");
$celda = $tabla->addCell();
$celda->addText("Direccion");
$celda =$tabla->addCell();
$celda->addText('Codigo');
$celda =$tabla->addCell();
$celda->addText('U');
$celda =$tabla->addCell();
$celda->addText('ObsMtuo');
$celda =$tabla->addCell();
$celda->addText('No.Medi');
$celda =$tabla->addCell();
$celda->addText('Actual'); */


//Guardar el documento 
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('Listado de Toma de Lecturas.docx');

//Descargar el documento en la ruta de descargas
//Los nombres deben coincidir con el que se asignó en el objeto save
header("Content-Disposition: attachment; filename=Listado de Toma de Lecturas.docx");
header("Content-type: application/msword ");
readfile("Listado de Toma de Lecturas.docx");

?>