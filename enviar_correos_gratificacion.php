<?php
    require 'vendor/autoload.php';

    require_once __DIR__ . '/PHPMailer/PHPMailer.php';
    require_once __DIR__ . '/PHPMailer/SMTP.php';
    require_once __DIR__ . '/PHPMailer/Exception.php';

    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PHPMailer\PHPMailer\PHPMailer;
    use PHPMailer\PHPMailer\Exception;

    header('Content-Type: application/json');

    if ($_SERVER['REQUEST_METHOD'] !== 'POST' || empty($_FILES['planilla'])) {
        echo json_encode(['ok' => false, 'mensaje' => 'Datos incompletos.']);
        exit;
    }

    $destinatarios = json_decode($_POST['destinatarios'] ?? '[]', true);
    if (empty($destinatarios)) {
        echo json_encode(['ok' => false, 'mensaje' => 'No se recibieron destinatarios.']);
        exit;
    }

    // ── Leer planilla ──────────────────────────────────────────────
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

    function getCellValGrat($hoja, $col, $row) {
        $cell = $hoja->getCell($col . $row);
        $val  = $cell->getValue();
        if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
            $cached = $cell->getCalculatedValueString();
            return $cached !== '' ? $cached : 0;
        }
        return $val;
    }

    // Reutilizamos la misma lógica de procesar_gratificacion.php
    require_once 'generar_pdf_gratificacion.php';

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

    function limpiarNumGrat($v) {
        if (is_numeric($v)) return floatval($v);
        return floatval(preg_replace('/[^0-9.\-]/', '', str_replace(',', '', (string)$v)));
    }
    function fmtFechaGrat($v) {
        if (empty($v)) return '';
        if (is_numeric($v)) {
            $ts = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($v);
            return date('d/m/Y', $ts);
        }
        $ts = strtotime($v);
        return $ts ? date('d/m/Y', $ts) : $v;
    }

    // Indexar destinatarios por DNI para búsqueda rápida
    $dniSet = array_column($destinatarios, 'dni');

    $maxRow = $hoja->getHighestRow();
    $enviados = 0;
    $errores  = [];
    $pdfsTemp = [];

    // Los datos empiezan en la fila 3 (fila 1 = fecha de corte global, fila 2 = encabezados)
    for ($row = 3; $row <= $maxRow; $row++) {
        $dni = trim((string)getCellValGrat($hoja, 'A', $row));
        if (empty($dni)) continue;
        if (!in_array($dni, $dniSet)) continue;

        // Buscar correo del destinatario
        $correo = '';
        foreach ($destinatarios as $d) {
            if ($d['dni'] === $dni) { $correo = $d['correo']; break; }
        }
        if (empty($correo)) continue;

        // Construir array del trabajador
        $t = [];
        foreach ($columnas as $campo => $col) {
            $val = getCellValGrat($hoja, $col, $row);
            if (is_string($val) && strlen($val) > 0 && $val[0] === '=') $val = 0;
            $t[$campo] = in_array($campo, $campos_numericos) ? limpiarNumGrat($val) : ($val ?? '');
        }
        $t['f_nacimiento'] = fmtFechaGrat($t['f_nacimiento']);
        $t['f_ingreso']    = fmtFechaGrat($t['f_ingreso']);
        $t['al']           = fmtFechaGrat($t['al']);

        // Generar PDF en carpeta temporal
        $pdfPath = generarConstanciaGratificacionPDF($t);
        $pdfsTemp[] = $pdfPath;

        // ── Enviar email con PHPMailer ─────────────────────────────
        $mail = new PHPMailer(true);
        try {
            // ╔══════════════════════════════════════════════╗
            // ║  CONFIGURA AQUÍ TUS DATOS SMTP               ║
            // ╚══════════════════════════════════════════════╝

$mail->isSMTP();
$mail->Host       = 'smtp.gmail.com';
$mail->SMTPAuth   = true;
$mail->Username   = 'aoberluis@agroimportadoraleon.com';
$mail->Password   = 'mebf rwzk uqjq bdpd';
$mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
$mail->Port       = 587;

$mail->SMTPDebug = 0;

$mail->SMTPOptions = [
    'ssl' => [
        'verify_peer'       => false,
        'verify_peer_name'  => false,
        'allow_self_signed' => true
    ]
];

$mail->CharSet = 'UTF-8';

        $mail->setFrom('aoberluis@agroimportadoraleon.com', 'Agroimportadora Leon S.A.C.');
        $mail->addAddress($correo, $t['nombre']);

            $mail->Subject = 'Constancia de Gratificación – ' . $t['al'];
            $mail->isHTML(true);
            $mail->Body = "
                <div style='font-family:Arial,sans-serif;max-width:500px;margin:0 auto;'>
                    <h2 style='color:#1a6bbf;'>Agroimportadora Leon S.A.C.</h2>
                    <p>Estimado/a <strong>{$t['nombre']}</strong>,</p>
                    <p>Adjuntamos su constancia de pago de gratificación correspondiente al periodo con fecha de corte <strong>{$t['al']}</strong>.</p>
                    <p>Saludos</p>
                    <p style='color:#666;font-size:13px;margin-top:24px;'>
                        Este es un correo automático, por favor no responder.
                    </p>
                </div>
            ";
            $mail->AltBody = "Estimado/a {$t['nombre']}, adjuntamos su constancia de gratificación al {$t['al']}.";
            $mail->addAttachment($pdfPath, 'Constancia_Gratificacion_' . $dni . '.pdf');

            $mail->send();
            $enviados++;
        } catch (Exception $e) {
            $errores[] = $t['nombre'] . ': ' . $mail->ErrorInfo;
        }
    }

    // Limpiar PDFs temporales
    foreach ($pdfsTemp as $p) { if (file_exists($p)) unlink($p); }
    unlink($destino);

    // Respuesta
    if (empty($errores)) {
        echo json_encode(['ok' => true, 'mensaje' => "Se enviaron {$enviados} constancia(s) correctamente."]);
    } else {
        $msg = "Enviados: {$enviados}. Errores: " . implode(' | ', $errores);
        echo json_encode(['ok' => count($errores) < count($destinatarios), 'mensaje' => $msg]);
    }
    exit;