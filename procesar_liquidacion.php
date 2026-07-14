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
$destino = dirname(__FILE__) . '/uploads/' . uniqid() . '_liquidacion.xlsx';
move_uploaded_file($tmpFile, $destino);

$reader = new Xlsx();
$reader->setReadDataOnly(true);
$spreadsheet = $reader->load($destino);

$hoja = null;
foreach ($spreadsheet->getSheetNames() as $nombreHoja) {
    if (stripos($nombreHoja, 'liquidaci') !== false) {
        $hoja = $spreadsheet->getSheetByName($nombreHoja);
        break;
    }
}
if (!$hoja) {
    $hoja = $spreadsheet->getActiveSheet();
}

function getCellValueLiq($hoja, $col, $row) {
    $cell = $hoja->getCell($col . $row);
    $val  = $cell->getValue();
    if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
        $cached = $cell->getCalculatedValueString();
        return $cached !== '' ? $cached : 0;
    }
    return $val;
}

// Mapeo de columnas de la hoja "Liquidación"
$columnas = [
    'dni'              => 'A',
    'nombre'           => 'B',
    'f_nacimiento'     => 'C',
    'puesto'           => 'D',
    'afp_onp'          => 'E',
    'cussp'            => 'F',
    'f_ingreso'        => 'G',
    'f_cese'           => 'H',
    'dias_asistidos'   => 'I',
    'basico'           => 'K',
    'asig_familiar'    => 'L',
    'bono_asistencia'  => 'M',
    'rem_computable'   => 'N',
    'bono29351'        => 'V',
    'grat_base'        => 'W',
    'grat_meses'       => 'X',
    'grat_total'       => 'Y',
    'cts_meses'        => 'Z',
    'cts_dias'         => 'AA',
    'cts_monto'        => 'AB',
    'vac_meses'        => 'AC',
    'vac_dias'         => 'AD',
    'vac_monto'        => 'AE',
    'descuento_afp'    => 'AF',
    'descuento_segafp' => 'AG',
    'descuento_onp'    => 'AH',
    'comedor'          => 'AI',
    'essalud'          => 'AJ',
    'sueldo'           => 'AK',
    'neto_liquidacion' => 'AL',
    // Opcionales: si la planilla las agrega, se usan; si no, quedan vacías y
    // se aplica el valor por defecto más abajo.
    'regimen'          => 'AN',
    'causa'            => 'AO',
];

$campos_numericos = [
    'basico', 'asig_familiar', 'bono_asistencia', 'rem_computable',
    'bono29351', 'grat_base', 'grat_meses', 'grat_total',
    'cts_meses', 'cts_dias', 'cts_monto',
    'vac_meses', 'vac_dias', 'vac_monto',
    'descuento_afp', 'descuento_segafp', 'descuento_onp',
    'comedor', 'essalud', 'sueldo', 'neto_liquidacion',
];

function limpiarNumeroLiq($valor) {
    if (is_numeric($valor)) return floatval($valor);
    $valor = preg_replace('/[^0-9.\-]/', '', str_replace(',', '', (string)$valor));
    return floatval($valor);
}

function formatearFechaLiq($valor) {
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

for ($row = 4; $row <= $maxRow; $row++) {
    $dni = getCellValueLiq($hoja, 'A', $row);
    if (empty($dni)) continue;

    $dniStr = trim((string)$dni);
    if (!empty($dnisSeleccionados) && !in_array($dniStr, $dnisSeleccionados)) {
        continue;
    }

    $trabajador = [];
    foreach ($columnas as $campo => $col) {
        $val = getCellValueLiq($hoja, $col, $row);
        if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
            $val = 0;
        }
        if (in_array($campo, $campos_numericos)) {
            $trabajador[$campo] = limpiarNumeroLiq($val);
        } else {
            $trabajador[$campo] = $val ?? '';
        }
    }

    $trabajador['dni']          = $dniStr;
    $trabajador['f_nacimiento'] = formatearFechaLiq($trabajador['f_nacimiento']);
    $trabajador['f_ingreso']    = formatearFechaLiq($trabajador['f_ingreso']);
    $trabajador['f_cese']       = formatearFechaLiq($trabajador['f_cese']);

    // Valores por defecto si la planilla no trae estas columnas
    if (empty($trabajador['regimen'])) $trabajador['regimen'] = 'REGIMEN GENERAL REMYPE';
    if (empty($trabajador['causa']))   $trabajador['causa']   = 'RENUNCIA';

    $trabajadores[] = $trabajador;
}

if (empty($trabajadores)) {
    die('No se encontraron trabajadores con los criterios seleccionados.');
}

require 'generar_pdf_liquidacion.php';

$archivos_pdf = [];
foreach ($trabajadores as $t) {
    $archivos_pdf[] = generarLiquidacionPDF($t);
}

$zipName = count($archivos_pdf) === 1
    ? 'liquidacion_' . ($trabajadores[0]['dni'] ?? 'trabajador') . '_' . date('Ymd') . '.zip'
    : 'liquidaciones_' . date('Ymd_His') . '.zip';

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