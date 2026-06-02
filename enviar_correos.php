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
    $destino = dirname(__FILE__) . '/uploads/' . uniqid() . '_planilla.xlsx';
    move_uploaded_file($tmpFile, $destino);

    $reader = new Xlsx();
    $reader->setReadDataOnly(true);
    $spreadsheet = $reader->load($destino);
    $hoja = $spreadsheet->getSheetByName('Planilla Mensual') ?? $spreadsheet->getActiveSheet();

    function getCellVal($hoja, $col, $row) {
        $cell = $hoja->getCell($col . $row);
        $val  = $cell->getValue();
        if (is_string($val) && strlen($val) > 0 && $val[0] === '=') {
            $cached = $cell->getCalculatedValueString();
            return $cached !== '' ? $cached : 0;
        }
        return $val;
    }

    // Reutilizamos la misma lógica de procesar.php
    require_once 'generar_pdf.php';

    $columnas = [
        'tipo'               => 'B',  'dni'             => 'C',
        'nombre'             => 'D',  'f_nacimiento'    => 'E',
        'puesto'             => 'F',  'afp_onp'         => 'G',
        'cussp'              => 'H',  'banco'           => 'I',
        'cuenta'             => 'J',  'f_ingreso'       => 'L',
        'periodo'            => 'O',  'asig_familiar'   => 'P',
        'basico'             => 'Q',  'dias_trab'       => 'V',
        'dias_descanso'      => 'W',  'Vacaciones'      => 'X',
        'dias_feriados'      => 'Y',  'horas_extras'       => 'AD', 
        'monto_basico'    => 'AE',
        'monto_asig'         => 'AF', 'feriado_lab'     => 'AG',
        'bono_asistencia'    => 'AH', 'bono_horas'      => 'AI',
        'rem_bruta'          => 'AJ', 'movilidad'       => 'AK',
        'viaticos'           => 'AL', 'afp_10'          => 'AM',
        'seg_afp'            => 'AN', 'onp'             => 'AO',
        'renta_5ta'          => 'AP', 'otros_desc'      => 'AQ',
        'adelantos'          => 'AR', 'adelanto_quincena' => 'AS',
        'total_desc'         => 'AT', 'neto'            => 'AU',
        'total_no_rem'       => 'AV', 'essalud'         => 'AW',
        'total_a_pagar'      => 'AX', 'correo'             => 'AZ',
        'f_cese'            => 'AY',
    ];

    $campos_numericos = [
        'basico','dias_trab','dias_descanso','horas_extras',
        'monto_basico','monto_asig','feriado_lab','bono_asistencia','bono_horas',
        'rem_bruta','movilidad','viaticos','afp_10','seg_afp','onp',
        'renta_5ta','otros_desc','adelantos','adelanto_quincena','total_desc',
        'neto','total_no_rem','essalud','total_a_pagar','extra_pagar'
    ];

    function limpiarNum($v) {
        if (is_numeric($v)) return floatval($v);
        return floatval(preg_replace('/[^0-9.\-]/', '', str_replace(',', '', $v)));
    }
    function fmtFecha($v) {
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

    for ($row = 3; $row <= $maxRow; $row++) {
        $tipo = getCellVal($hoja, 'B', $row);
        $dni  = trim((string)getCellVal($hoja, 'C', $row));
        if (empty($dni) || strtoupper(trim($tipo)) !== 'PLANILLA') continue;
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
            $val = getCellVal($hoja, $col, $row);
            if (is_string($val) && strlen($val) > 0 && $val[0] === '=') $val = 0;
            $t[$campo] = in_array($campo, $campos_numericos) ? limpiarNum($val) : ($val ?? '');
        }
        $t['f_nacimiento'] = fmtFecha($t['f_nacimiento']);
        $t['f_ingreso']    = fmtFecha($t['f_ingreso']);
        if (is_string($t['periodo'])) {
            $t['periodo'] = strtoupper(trim(preg_replace('/(\d{4})\d+$/', '$1', $t['periodo'])));
        }

        // Generar PDF en carpeta temporal
        $pdfPath = generarBoletaPDF($t);
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

            $mail->Subject = 'Boleta de Remuneraciones – ' . $t['periodo'];
            $mail->isHTML(true);
            $mail->Body = "
                <div style='font-family:Arial,sans-serif;max-width:500px;margin:0 auto;'>
                    <h2 style='color:#1a6bbf;'>Agroimportadora Leon S.A.C.</h2>
                    <p>Estimado/a <strong>{$t['nombre']}</strong>,</p>
                    <p>Adjuntamos su boleta de remuneraciones correspondiente al periodo <strong>{$t['periodo']}</strong>.</p>
                    <p>Saludos</p>
                    <p style='color:#666;font-size:13px;margin-top:24px;'>
                        Este es un correo automático, por favor no responder.
                    </p>
                </div>
            ";
            $mail->AltBody = "Estimado/a {$t['nombre']}, adjuntamos su boleta del periodo {$t['periodo']}.";
            $mail->addAttachment($pdfPath, 'Boleta_' . $t['periodo'] . '_' . $dni . '.pdf');

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
        echo json_encode(['ok' => true, 'mensaje' => "Se enviaron {$enviados} boleta(s) correctamente."]);
    } else {
        $msg = "Enviados: {$enviados}. Errores: " . implode(' | ', $errores);
        echo json_encode(['ok' => count($errores) < count($destinatarios), 'mensaje' => $msg]);
    }
    exit;