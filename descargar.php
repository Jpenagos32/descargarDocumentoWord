<?php

use PhpOffice\PhpWord\Style\Language;
require "vendor/autoload.php";


//Crear el nuevo documento
$phpWord = new \PhpOffice\PhpWord\PhpWord();
\PhpOffice\PhpWord\Settings::setDefaultPaper('Letter');

// Agregar una seccion vacia al documento
$seccion = $phpWord->addSection();

###################################################################################################################################################

// Configuraciones por defecto del documento
//Añadir fuente personalizada o deseada
$fuente = 'Arial11';
$phpWord->addFontStyle(
    $fuente, 
    [
        'name' => 'Arial',
        'size' => 11,
        'bold' => false
    ]
);

$fuenteCodigo = [
    'name' => 'Verdana',
    'size' => 11,
    'bold' => false,
];

//Centrar texto
$centrado = 'miEstilo';
$phpWord->addParagraphStyle($centrado, ['align' => 'center']);

// Añadir el lenguaje español al documento
$phpWord->getSettings()->setThemeFontLang(new Language(Language::ES_ES));

// Margenes por defecto en el documento
$margenDocumento = $seccion->getStyle();
$margenDocumento->setMarginLeft(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1));
$margenDocumento->setMarginRight(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1));
$margenDocumento->setMarginTop(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1));
$margenDocumento->setMarginBottom(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1));

// Documento de solo lectura Comentar para desactivar el modo
$proteccionDocumento = $phpWord->getSettings()->getDocumentProtection();
$proteccionDocumento->setEditing('readOnly');

###############################################################################################################################

//Añadir los titulos del documento como header para todas las páginas
$header = $seccion->addHeader();
$header->addText(
    "Acueducto y Alcantarillado de Popayán S.A. E. S.P",
    $fuente, $centrado
);

$header->addText('Empresa de Servicios Públicos',
    $fuente, $centrado
);

$header->addText('NIT 891.500.117-1',
    $fuente, $centrado
);

$header->addText(
    'Listado de Toma de Lecturas Mes: 2023-02 Ciclo: 4',
    ['name' => 'Arial', 'size' => 11, 'bold' => true], $centrado
);

###########################################################################################################################################

// Crear estilo de tabla
$estiloTabla = [
    'borderColor' => 'ffffff',
    'borderSize' => 0,
    'position' => 'vertAnchor',
    'cellMarginRight' => 85,
    // 'width' => 2000 * 2000,
    'unit' => 'pct',
    'align' => 'center',
    'layout' => 'autofit'
];
// $phpWord->addTableStyle('estilo', $estiloTabla);

$estiloCelda = [
    'valign' => 'center'
];

$estiloParrafo = [
    'alignment' => 'center',
];

$estiloFila = [
    'cantSplit' => false,
    'exactHeight' => true
];
$estiloFilaHeader = [
    'tblHeader' => true,
    'cantSplit' => false,
    'exactHeight' => true
];

#####################################################################################################################################################

// Agregar array con datos provisionales
$datosPersonas = array(
    array('Camilo Rodriguez 1', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 2', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 3', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 4', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 5', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 6', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 7', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 8', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 9', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 10', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 11', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 12', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 13', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 14', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 15', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 16', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 17', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 18', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 19', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 20', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 21', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 22', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 23', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 24', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 25', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 26', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 27', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 28', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 29', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 30', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 31', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 32', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 33', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 34', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 35', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 36', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 37', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 38', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 39', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________'),
    array('Camilo Rodriguez 40', 'Kra 17A # 57N-253 EL UVO', '02/05/0680/00', 1, '--', '0044060-2013', '_________')
);

###################################################################################################################################################

// Agregar la tabla del Header
$tabla = $seccion->addTable($estiloTabla);
$tabla->addRow(854, $estiloFilaHeader);
$celda = $tabla->addCell();
$celda->addText("Nombre", $fuente, $estiloParrafo);
$celda = $tabla->addCell();
$celda->addText("Direccion", $fuente, $estiloParrafo);
$celda =$tabla->addCell();
$celda->addText('Codigo', $fuente, $estiloParrafo);
$celda =$tabla->addCell();
$celda->addText('U', $fuente, $estiloParrafo);
$celda =$tabla->addCell();
$celda->addText('ObsMtuo', $fuente, $estiloParrafo);
$celda =$tabla->addCell();
$celda->addText('No.Medi', $fuente, $estiloParrafo);
$celda =$tabla->addCell();
$celda->addText('Actual', $fuente, $estiloParrafo);



// Obtener la informacion del arreglo para crear la tabla
foreach ($datosPersonas as $dato) {
    $tabla = $seccion->addTable($estiloTabla);
    $tabla->addRow(760, $estiloFila); # Cambiar alto de las celdas
    foreach ($dato as $valor) {
        $tipoDeLetra = strpos($valor, '/') ? $fuenteCodigo : $fuente;
        $valorEspaciado = " " . $valor;
        $espaciadoCodigo = strpos($valor, '/')? $valorEspaciado : $valor;
        $celda = $tabla->addCell();
        $celda->addText($espaciadoCodigo, $tipoDeLetra, $estiloParrafo);
    }
}

##############################################################################################################################################

//Guardar el documento 
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('Listado de Toma de Lecturas.docx');

//Descargar el documento en la ruta de descargas
//Los nombres deben coincidir con el que se asignó en el objeto save
header("Content-Disposition: attachment; filename=Listado de Toma de Lecturas.docx");
header("Content-type: application/msword ");
readfile("Listado de Toma de Lecturas.docx");

?>