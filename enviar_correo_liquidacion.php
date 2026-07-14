<?php
require_once __DIR__ . '/PHPMailer/PHPMailer.php';
require_once __DIR__ . '/PHPMailer/SMTP.php';
require_once __DIR__ . '/PHPMailer/Exception.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

header('Content-Type: application/json');

if ($_SERVER['REQUEST_METHOD'] !== 'POST' || empty($_FILES['archivo']) || empty($_POST['correo'])) {
    echo json_encode(['ok' => false, 'mensaje' => 'Datos incompletos.']);
    exit;
}

$correo = trim($_POST['correo']);
$nombre = trim($_POST['nombre'] ?? '');

if (!filter_var($correo, FILTER_VALIDATE_EMAIL)) {
    echo json_encode(['ok' => false, 'mensaje' => 'El correo ingresado no es válido.']);
    exit;
}

if ($_FILES['archivo']['error'] !== UPLOAD_ERR_OK) {
    echo json_encode(['ok' => false, 'mensaje' => 'Error al subir el archivo.']);
    exit;
}

$nombreOriginal = $_FILES['archivo']['name'];
$ext = strtolower(pathinfo($nombreOriginal, PATHINFO_EXTENSION));
$extsPermitidas = ['pdf', 'doc', 'docx', 'xls', 'xlsx'];
if (!in_array($ext, $extsPermitidas)) {
    echo json_encode(['ok' => false, 'mensaje' => 'Formato de archivo no permitido. Usa PDF, Word o Excel.']);
    exit;
}

$rutaTemp = dirname(__FILE__) . '/uploads/' . uniqid() . '_liq.' . $ext;
if (!move_uploaded_file($_FILES['archivo']['tmp_name'], $rutaTemp)) {
    echo json_encode(['ok' => false, 'mensaje' => 'No se pudo procesar el archivo.']);
    exit;
}

$destinatarioNombre = $nombre !== '' ? $nombre : 'Estimado/a';

$mail = new PHPMailer(true);
try {
    $mail->isSMTP();
    $mail->Host       = 'smtp.gmail.com';
    $mail->SMTPAuth   = true;
    $mail->Username   = 'aoberluis@agroimportadoraleon.com';
    $mail->Password   = 'mebf rwzk uqjq bdpd';
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
    $mail->Port       = 587;
    $mail->SMTPDebug  = 0;

    $mail->SMTPOptions = [
        'ssl' => [
            'verify_peer'       => false,
            'verify_peer_name'  => false,
            'allow_self_signed' => true
        ]
    ];

    $mail->CharSet = 'UTF-8';

    $mail->setFrom('aoberluis@agroimportadoraleon.com', 'Agroimportadora Leon S.A.C.');
    $mail->addAddress($correo, $nombre);

    $mail->isHTML(true);
    $mail->Subject = 'Documento de Liquidación – Agroimportadora Leon S.A.C.';
    $mail->Body = "
        <div style='font-family:Arial,sans-serif;max-width:500px;margin:0 auto;'>
            <h2 style='color:#1a6bbf;'>Agroimportadora Leon S.A.C.</h2>
            <p>Estimado/a <strong>{$destinatarioNombre}</strong>,</p>
            <p>Adjuntamos su documento de liquidación.</p>
            <p>Saludos</p>
            <p style='color:#666;font-size:13px;margin-top:24px;'>
                Este es un correo automático, por favor no responder.
            </p>
        </div>
    ";
    $mail->AltBody = "Estimado/a {$destinatarioNombre}, adjuntamos su documento de liquidación.";

    $mail->addAttachment($rutaTemp, $nombreOriginal);

    $mail->send();
    unlink($rutaTemp);

    echo json_encode(['ok' => true, 'mensaje' => "Documento de liquidación enviado correctamente a {$correo}."]);
} catch (Exception $e) {
    if (file_exists($rutaTemp)) unlink($rutaTemp);
    echo json_encode(['ok' => false, 'mensaje' => 'Error al enviar: ' . $mail->ErrorInfo]);
}
exit;