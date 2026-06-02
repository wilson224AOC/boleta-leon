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
$destino = dirname(__FILE__) . '/uploads/' . uniqid() . '_planilla.xlsx';
move_uploaded_file($tmpFile, $destino);

$reader = new Xlsx();
$reader->setReadDataOnly(true);
$spreadsheet = $reader->load($destino);

$hoja = $spreadsheet->getSheetByName('Planilla Mensual');
if (!$hoja) {
    $hoja = $spreadsheet->getActiveSheet();
}

function getCellValue($hoja, $col, $row) {
    $cell = $hoja->getCell($col . $row);
    $val  = $cell->getValue();
    if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
        $cached = $cell->getCalculatedValueString();
        return $cached !== '' ? $cached : 0;
    }
    return $val;
}

$columnas = [
    'tipo'            => 'B',
    'dni'             => 'C',
    'nombre'          => 'D',
    'f_nacimiento'    => 'E',
    'puesto'          => 'F',
    'afp_onp'         => 'G',
    'cussp'           => 'H',
    'banco'           => 'I',
    'cuenta'          => 'J',
    'f_ingreso'       => 'L',
    'periodo'         => 'O',
    'asig_familiar'   => 'P',
    'basico'          => 'Q',
    'dias_trab'       => 'V',
    'dias_descanso'   => 'W',
    'Vacaciones'      => 'X',
    'dias_feriados'   => 'Y',
    'horas_extras'    => 'AD',
    'monto_basico'    => 'AE',
    'monto_asig'      => 'AF',
    'feriado_lab'     => 'AG',
    'bono_asistencia' => 'AH',
    'bono_horas'      => 'AI',
    'rem_bruta'       => 'AJ',
    'movilidad'       => 'AK',
    'viaticos'        => 'AL',
    'afp_10'          => 'AM',
    'seg_afp'         => 'AN',
    'onp'             => 'AO',
    'renta_5ta'       => 'AP',
    'otros_desc'      => 'AQ',
    'adelantos'       => 'AR',
    'adelanto_quincena' => 'AS',
    'total_desc'      => 'AT',
    'neto'            => 'AU',
    'total_no_rem'    => 'AV',
    'essalud'         => 'AW',
    'total_a_pagar'   => 'AX',
    'f_cese'          => 'AY',
    'correo'          => 'AZ',
];

$campos_numericos = [
    'basico','dias_trab','dias_descanso','horas_extras',
    'monto_basico','monto_asig','feriado_lab','bono_asistencia','bono_horas',
    'rem_bruta','movilidad','viaticos','afp_10','seg_afp','onp',
    'renta_5ta','otros_desc','adelantos','total_desc','neto',
    'total_no_rem','essalud','total_a_pagar','extra_pagar'
];

function limpiarNumero($valor) {
    if (is_numeric($valor)) return floatval($valor);
    $valor = preg_replace('/[^0-9.\-]/', '', str_replace(',', '', $valor));
    return floatval($valor);
}

function formatearFecha($valor) {
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

for ($row = 3; $row <= $maxRow; $row++) {
    $tipo = getCellValue($hoja, 'B', $row);
    $dni  = getCellValue($hoja, 'C', $row);

    if (empty($dni)) continue;
    if (strtoupper(trim($tipo)) !== 'PLANILLA') continue;

    // Filtrar solo los DNIs seleccionados (si se enviaron)
    $dniStr = trim((string)$dni);
    if (!empty($dnisSeleccionados) && !in_array($dniStr, $dnisSeleccionados)) {
        continue;
    }

    $trabajador = [];
    foreach ($columnas as $campo => $col) {
        $val = getCellValue($hoja, $col, $row);
        if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
            $val = 0;
        }
        if (in_array($campo, $campos_numericos)) {
            $trabajador[$campo] = limpiarNumero($val);
        } else {
            $trabajador[$campo] = $val ?? '';
        }
    }

    $trabajador['f_nacimiento'] = formatearFecha($trabajador['f_nacimiento']);
    $trabajador['f_ingreso']    = formatearFecha($trabajador['f_ingreso']);
    $trabajador['f_cese']       = formatearFecha($trabajador['f_cese']);

    
    $periodo = $trabajador['periodo'] ?? '';
    if (is_string($periodo)) {
        $periodo = preg_replace('/(\d{4})\d+$/', '$1', $periodo);
        $trabajador['periodo'] = strtoupper(trim($periodo));
    }

    $trabajadores[] = $trabajador;
}

if (empty($trabajadores)) {
    die('No se encontraron trabajadores con los criterios seleccionados.');
}

require 'generar_pdf.php';

$archivos_pdf = [];
foreach ($trabajadores as $t) {
    $archivos_pdf[] = generarBoletaPDF($t);
}

$zipName = count($archivos_pdf) === 1
    ? 'boleta_' . ($trabajadores[0]['dni'] ?? 'trabajador') . '_' . date('Ymd') . '.zip'
    : 'boletas_' . date('Ymd_His') . '.zip';

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