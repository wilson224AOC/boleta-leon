<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

if ($_SERVER['REQUEST_METHOD'] !== 'POST' || empty($_FILES['planilla']) || empty($_POST['dni'])) {
    http_response_code(400);
    die('Solicitud no válida.');
}

$dniBuscado = trim((string)$_POST['dni']);

$tmpFile = $_FILES['planilla']['tmp_name'];
$destino = dirname(__FILE__) . '/uploads/' . uniqid() . '_gratificacion_preview.xlsx';
move_uploaded_file($tmpFile, $destino);

$reader = new Xlsx();
$reader->setReadDataOnly(true);
$spreadsheet = $reader->load($destino);

// La hoja de gratificaciones suele llamarse "Gratificación (2)" o similar;
// si no se encuentra, se usa la hoja activa.
$hoja = null;
foreach ($spreadsheet->getSheetNames() as $nombreHoja) {
    if (stripos($nombreHoja, 'gratific') !== false) {
        $hoja = $spreadsheet->getSheetByName($nombreHoja);
        break;
    }
}
if (!$hoja) {
    $hoja = $spreadsheet->getActiveSheet();
}

function getCellValueGratPreview($hoja, $col, $row) {
    $cell = $hoja->getCell($col . $row);
    $val  = $cell->getValue();
    if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
        $cached = $cell->getCalculatedValueString();
        return $cached !== '' ? $cached : 0;
    }
    return $val;
}

$columnas = [
    'dni'                  => 'A',
    'nombre'               => 'B',
    'f_nacimiento'         => 'C',
    'puesto'               => 'D',
    'correo'               => 'E',
    'afp_onp'              => 'F',
    'cussp'                => 'G',
    'f_ingreso'            => 'H',
    'al'                   => 'I',
    'asig_familiar_flag'   => 'J',
    'basico'               => 'K',
    'asignacion_familiar'  => 'L',
    'otras_rem'            => 'M',
    'rem_computable'       => 'N',
    'bono_ley29351'        => 'O',
    'total_base'           => 'P',
    'meses'                => 'Q',
    'importe'              => 'R',
];

$campos_numericos = [
    'basico', 'asignacion_familiar', 'otras_rem', 'rem_computable',
    'bono_ley29351', 'total_base', 'meses', 'importe',
];

function limpiarNumeroGratPreview($valor) {
    if (is_numeric($valor)) return floatval($valor);
    $valor = preg_replace('/[^0-9.\-]/', '', str_replace(',', '', (string)$valor));
    return floatval($valor);
}

function formatearFechaGratPreview($valor) {
    if (empty($valor)) return '';
    if (is_numeric($valor)) {
        $ts = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($valor);
        return date('d/m/Y', $ts);
    }
    $ts = strtotime($valor);
    if ($ts) return date('d/m/Y', $ts);
    return $valor;
}

$trabajador = null;
$maxRow = $hoja->getHighestRow();

// Los datos empiezan en la fila 3 (fila 1 = fecha de corte global, fila 2 = encabezados)
for ($row = 3; $row <= $maxRow; $row++) {
    $dni = getCellValueGratPreview($hoja, 'A', $row);
    if (empty($dni)) continue;
    if (trim((string)$dni) !== $dniBuscado) continue;

    $trabajador = [];
    foreach ($columnas as $campo => $col) {
        $val = getCellValueGratPreview($hoja, $col, $row);
        if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
            $val = 0;
        }
        if (in_array($campo, $campos_numericos)) {
            $trabajador[$campo] = limpiarNumeroGratPreview($val);
        } else {
            $trabajador[$campo] = $val ?? '';
        }
    }
    $trabajador['f_nacimiento'] = formatearFechaGratPreview($trabajador['f_nacimiento']);
    $trabajador['f_ingreso']    = formatearFechaGratPreview($trabajador['f_ingreso']);
    $trabajador['al']           = formatearFechaGratPreview($trabajador['al']);
    break;
}

unlink($destino);

if (!$trabajador) {
    http_response_code(404);
    die('No se encontró al trabajador con DNI ' . htmlspecialchars($dniBuscado) . ' en la planilla.');
}

require 'generar_pdf_gratificacion.php';

$rutaPdf = generarConstanciaGratificacionPDF($trabajador);

// Se devuelve inline (no attachment) para que el navegador lo muestre en un iframe
// en vez de descargarlo.
header('Content-Type: application/pdf');
header('Content-Disposition: inline; filename="' . basename($rutaPdf) . '"');
header('Content-Length: ' . filesize($rutaPdf));
header('Cache-Control: no-store');
readfile($rutaPdf);

if (file_exists($rutaPdf)) {
    unlink($rutaPdf);
}
exit;