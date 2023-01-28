<?php

require "vendor/autoload.php";

//Crear el nuevo documento
$phpWord = new \PhpOffice\PhpWord\PhpWord();


// Agregar una seccion vacia al documento
$seccion = $phpWord->addSection();

//Añadir elemento de texto con fuente personalizada
$fuente = 'Arial11';
$phpWord->addFontStyle(
    $fuente, ['name' => 'Arial', 'size' => 11]
);

//Centrar texto
$centrado = 'miEstilo';
$phpWord->addParagraphStyle($centrado, ['align' => 'center']);

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

// Crear tabla
$tabla = $seccion->addTable('estilo');
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
$celda->addText('Actual');

$tabla->addRow();
$celda =$tabla->addCell();
$celda->addText('RADIO SUPER TRANSMISORES');
$celda =$tabla->addCell();
$celda->addText('Kra 17A # 57N-253 EL UVO');
$celda =$tabla->addCell();
$celda->addText('02/05/0680/00 ');
$celda =$tabla->addCell();
$celda->addText('1');
$celda =$tabla->addCell();
$celda->addText('--');
$celda =$tabla->addCell();
$celda->addText('0044060-2013');
$celda =$tabla->addCell();
$celda->addText('---------');

$tabla->addRow();
$celda =$tabla->addCell();
$celda->addText('DILFREDO RIOS HERRERA');
$celda =$tabla->addCell();
$celda->addText('Kra 17A # 57N-295 EL UVO');
$celda =$tabla->addCell();
$celda->addText('02/05/0700/00');
$celda =$tabla->addCell();
$celda->addText('1');
$celda =$tabla->addCell();
$celda->addText('--');
$celda =$tabla->addCell();
$celda->addText('0282253-2020');
$celda =$tabla->addCell();
$celda->addText('---------');

$tabla->addRow();
$celda =$tabla->addCell();
$celda->addText('TOTALIZADOR LA RIOJA');
$celda =$tabla->addCell();
$celda->addText('Kra 17A # 57N-61 LA RIOJA');
$celda =$tabla->addCell();
$celda->addText('02/05/0750/00');
$celda =$tabla->addCell();
$celda->addText('7');
$celda =$tabla->addCell();
$celda->addText('--');
$celda =$tabla->addCell();
$celda->addText('03W-160142');
$celda =$tabla->addCell();
$celda->addText('---------');

$tabla->addRow();
$celda =$tabla->addCell();
$celda->addText('TEMPORAL GONZALEZ RODRIG');
$celda =$tabla->addCell();
$celda->addText('EZ S. EN C. Kra 17A # 57N-61 LA RIOJA');
$celda =$tabla->addCell();
$celda->addText('02/05/0800/00');
$celda =$tabla->addCell();
$celda->addText('7');
$celda =$tabla->addCell();
$celda->addText('--');
$celda =$tabla->addCell();
$celda->addText('--0274712-2020');
$celda =$tabla->addCell();
$celda->addText('---------');




//Guardar el documento 
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('Listado de Toma de Lecturas.docx');

//Descargar el documento en la ruta de descargas
//Los nombres deben coincidir con el que se asignó en el objeto save
header("Content-Disposition: attachment; filename=Listado de Toma de Lecturas.docx");
header("Content-type: application/msword ");
readfile("Listado de Toma de Lecturas.docx");

?>