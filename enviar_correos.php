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
        'tipo'                  => 'B',
        'dni'                   => 'C',
        'nombre'                => 'D',
        'f_nacimiento'          => 'E',
        'puesto'                => 'F',
        'afp_onp'               => 'G',
        'cussp'                 => 'H',
        'banco'                 => 'I',
        'cuenta'                => 'J',
        'f_ingreso'             => 'L',
        'periodo'               => 'O',
        'asig_familiar'         => 'P',
        'basico'                => 'Q',
        'dias_trab'             => 'V',
        'dias_descanso'         => 'X',
        'Vacaciones'            => 'W',
        'dias_feriados'         => 'Y',
        'feriado_lab_dias'     => 'Z',
        'descanso_medico_dias'  => 'AA',
        'horas_extras'          => 'AE',
        'monto_basico'          => 'AF',
        'monto_asig'            => 'AG',
        'feriado_lab'           => 'AH',
        'bono_asistencia'       => 'AI',
        'bono_horas'            => 'AJ',
        'descanso_medico_monto' => 'AK',
        'rem_bruta'             => 'AL',
        'movilidad'             => 'AM',
        'viaticos'              => 'AN',
        'afp_10'                => 'AO',
        'seg_afp'               => 'AP',
        'onp'                   => 'AQ',
        'renta_5ta'             => 'AR',
        'otros_desc'            => 'AS',
        'adelantos'             => 'AT',
        'adelanto_quincena'     => 'AU',
        'total_desc'            => 'AV',
        'neto'                  => 'AW',
        'total_no_rem'          => 'AX',
        'essalud'               => 'AY',
        'total_a_pagar'         => 'AZ',
        'f_cese'                => 'BA',
        'correo'                => 'BB',
    ];

    $campos_numericos = [
        'basico','dias_trab','dias_descanso','horas_extras','feriado_lab_dias',
        'descanso_medico_dias','descanso_medico_monto',
        'monto_basico','monto_asig','feriado_lab','bono_asistencia','bono_horas',
        'rem_bruta','movilidad','viaticos','afp_10','seg_afp','onp',
        'renta_5ta','otros_desc','adelantos','adelanto_quincena','total_desc',
        'neto','total_no_rem','essalud','total_a_pagar'
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

    // Ver nota en procesar.php: el "Periodo" puede venir de una fórmula cuyo
    // valor cacheado en el Excel quedó en inglés (aunque Excel lo muestre en
    // español al abrirlo). Traducimos el mes para que la boleta salga en español.
    function traducirMesPeriodo($texto) {
        static $meses = [
            'JANUARY' => 'ENERO', 'FEBRUARY' => 'FEBRERO', 'MARCH' => 'MARZO',
            'APRIL' => 'ABRIL', 'MAY' => 'MAYO', 'JUNE' => 'JUNIO',
            'JULY' => 'JULIO', 'AUGUST' => 'AGOSTO', 'SEPTEMBER' => 'SEPTIEMBRE',
            'OCTOBER' => 'OCTUBRE', 'NOVEMBER' => 'NOVIEMBRE', 'DECEMBER' => 'DICIEMBRE',
        ];
        foreach ($meses as $en => $es) {
            $texto = preg_replace('/\b' . $en . '\b/i', $es, $texto);
        }
        return $texto;
    }

    // Indexar destinatarios por DNI para búsqueda rápida
    $dniSet = array_column($destinatarios, 'dni');

    $maxRow = $hoja->getHighestRow();
    $enviados = 0;
    $errores  = [];
    $pdfsTemp = [];
    $extrasTemp = [];

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
        $t['f_cese']       = fmtFecha($t['f_cese']);
        if (is_string($t['periodo'])) {
            $t['periodo'] = traducirMesPeriodo(strtoupper(trim(preg_replace('/(\d{4})\d+$/', '$1', $t['periodo']))));
        }

        // Generar PDF en carpeta temporal
        $pdfPath = generarBoletaPDF($t);
        $pdfsTemp[] = $pdfPath;

        // ── Documento extra (opcional) para este trabajador ─────────
        // El frontend lo envía con la clave "extra_<dni>" cuando el usuario
        // adjuntó un Word o PDF adicional para esa fila.
        $tieneExtra          = false;
        $extraPath           = '';
        $extraNombreOriginal = '';
        $extraFileKey = 'extra_' . $dni;
        if (isset($_FILES[$extraFileKey]) && $_FILES[$extraFileKey]['error'] === UPLOAD_ERR_OK) {
            $extraNombreOriginal = $_FILES[$extraFileKey]['name'];
            $extExtra = strtolower(pathinfo($extraNombreOriginal, PATHINFO_EXTENSION));
            if (in_array($extExtra, ['pdf', 'doc', 'docx'])) {
                $extraPath = dirname(__FILE__) . '/uploads/' . uniqid() . '_extra.' . $extExtra;
                if (move_uploaded_file($_FILES[$extraFileKey]['tmp_name'], $extraPath)) {
                    $tieneExtra = true;
                    $extrasTemp[] = $extraPath;
                }
            }
        }

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

            $mail->isHTML(true);

            if ($tieneExtra) {
                // Mensaje para boleta + documento adicional
                $mail->Subject = 'Boleta de Remuneraciones y Documento Adicional – ' . $t['periodo'];
                $mail->Body = "
                    <div style='font-family:Arial,sans-serif;max-width:500px;margin:0 auto;'>
                        <h2 style='color:#1a6bbf;'>Agroimportadora Leon S.A.C.</h2>
                        <p>Estimado/a <strong>{$t['nombre']}</strong>,</p>
                        <p>Adjuntamos su boleta de remuneraciones correspondiente al periodo <strong>{$t['periodo']}</strong>, junto con un documento adicional.</p>
                        <p>Saludos</p>
                        <p style='color:#666;font-size:13px;margin-top:24px;'>
                            Este es un correo automático, por favor no responder.
                        </p>
                    </div>
                ";
                $mail->AltBody = "Estimado/a {$t['nombre']}, adjuntamos su boleta del periodo {$t['periodo']} junto con un documento adicional.";
            } else {
                // Mensaje solo boleta (el mismo de siempre)
                $mail->Subject = 'Boleta de Remuneraciones – ' . $t['periodo'];
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
            }

            $mail->addAttachment($pdfPath, 'Boleta_' . $t['periodo'] . '_' . $dni . '.pdf');
            if ($tieneExtra) {
                $mail->addAttachment($extraPath, $extraNombreOriginal);
            }

            $mail->send();
            $enviados++;
        } catch (Exception $e) {
            $errores[] = $t['nombre'] . ': ' . $mail->ErrorInfo;
        }
    }

    // Limpiar PDFs y documentos extra temporales
    foreach ($pdfsTemp as $p) { if (file_exists($p)) unlink($p); }
    foreach ($extrasTemp as $p) { if (file_exists($p)) unlink($p); }
    unlink($destino);

    // Respuesta
    if (empty($errores)) {
        echo json_encode(['ok' => true, 'mensaje' => "Se enviaron {$enviados} boleta(s) correctamente."]);
    } else {
        $msg = "Enviados: {$enviados}. Errores: " . implode(' | ', $errores);
        echo json_encode(['ok' => count($errores) < count($destinatarios), 'mensaje' => $msg]);
    }
    exit;