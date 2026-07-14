<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

if ($_SERVER['REQUEST_METHOD'] !== 'POST' || empty($_FILES['planilla'])) {
    die('Acceso no válido.');
}

// DNIs seleccionados desde el formulario
$dnisSeleccionados = [];
if (!empty($_POST['dnis_seleccionados'])) {
    $dnisSeleccionados = array_filter(array_map('trim', explode(',', $_POST['dnis_seleccionados'])));
}

$tmpFile = $_FILES['planilla']['tmp_name'];
$destino = dirname(__FILE__) . '/uploads/' . uniqid() . '_gratificacion.xlsx';
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

function getCellValueGrat($hoja, $col, $row) {
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

function limpiarNumeroGrat($valor) {
    if (is_numeric($valor)) return floatval($valor);
    $valor = preg_replace('/[^0-9.\-]/', '', str_replace(',', '', (string)$valor));
    return floatval($valor);
}

function formatearFechaGrat($valor) {
    if (empty($valor)) return '';
    if (is_numeric($valor)) {
        $ts = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($valor);
        return date('d/m/Y', $ts);
    }
    $ts = strtotime($valor);
    if ($ts) return date('d/m/Y', $ts);
    return $valor;
}

$trabajadores = [];
$maxRow = $hoja->getHighestRow();

// Los datos empiezan en la fila 3 (fila 1 = fecha de corte global, fila 2 = encabezados)
for ($row = 3; $row <= $maxRow; $row++) {
    $dni = getCellValueGrat($hoja, 'A', $row);
    if (empty($dni)) continue;

    $dniStr = trim((string)$dni);
    if (!empty($dnisSeleccionados) && !in_array($dniStr, $dnisSeleccionados)) {
        continue;
    }

    $trabajador = [];
    foreach ($columnas as $campo => $col) {
        $val = getCellValueGrat($hoja, $col, $row);
        if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
            $val = 0;
        }
        if (in_array($campo, $campos_numericos)) {
            $trabajador[$campo] = limpiarNumeroGrat($val);
        } else {
            $trabajador[$campo] = $val ?? '';
        }
    }

    $trabajador['f_nacimiento'] = formatearFechaGrat($trabajador['f_nacimiento']);
    $trabajador['f_ingreso']    = formatearFechaGrat($trabajador['f_ingreso']);
    $trabajador['al']           = formatearFechaGrat($trabajador['al']);

    $trabajadores[] = $trabajador;
}

if (empty($trabajadores)) {
    die('No se encontraron trabajadores en la planilla de gratificación.');
}

require 'generar_pdf_gratificacion.php';

$archivos_pdf = [];
foreach ($trabajadores as $t) {
    $archivos_pdf[] = generarConstanciaGratificacionPDF($t);
}

$zipName = count($archivos_pdf) === 1
    ? 'gratificacion_' . ($trabajadores[0]['dni'] ?? 'trabajador') . '_' . date('Ymd') . '.zip'
    : 'gratificaciones_' . date('Ymd_His') . '.zip';

$zipPath = dirname(__FILE__) . '/pdfs/' . $zipName;
$zip = new ZipArchive();
$zip->open($zipPath, ZipArchive::CREATE);
foreach ($archivos_pdf as $pdf) {
    $zip->addFile($pdf, basename($pdf));
}
$zip->close();

foreach ($archivos_pdf as $pdf) {
    if (file_exists($pdf)) unlink($pdf);
}
unlink($destino);

header('Content-Type: application/zip');
header('Content-Disposition: attachment; filename="' . $zipName . '"');
header('Content-Length: ' . filesize($zipPath));
if (!empty($_POST['download_token'])) {
    $token = preg_replace('/[^0-9]/', '', $_POST['download_token']);
    setcookie('descarga_lista_' . $token, '1', time() + 60, '/');
}
readfile($zipPath);
unlink($zipPath);
exit;