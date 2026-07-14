<?php
require_once 'vendor/autoload.php';

/**
 * Calcula el periodo de gratificación (semestre) a partir de la fecha "Al".
 * Si el mes de la fecha "Al" es <= 6 => semestre Ene-Jun.
 * Si es > 6 => semestre Jul-Dic.
 * Devuelve ['inicio' => DateTime, 'fin' => DateTime]
 */
function calcularPeriodoGratificacion(DateTime $al): array {
    $anio = (int)$al->format('Y');
    $mes  = (int)$al->format('n');
    if ($mes <= 6) {
        $inicio = new DateTime("$anio-01-01");
    } else {
        $inicio = new DateTime("$anio-07-01");
    }
    return ['inicio' => $inicio, 'fin' => $al];
}

/**
 * Periodo computable = desde la fecha más tardía entre (fecha de ingreso, inicio del semestre)
 * hasta el fin del semestre.
 */
function calcularPeriodoComputable(DateTime $fIngreso, DateTime $inicioPeriodo, DateTime $finPeriodo): array {
    $inicioComputable = ($fIngreso > $inicioPeriodo) ? $fIngreso : $inicioPeriodo;
    return ['inicio' => $inicioComputable, 'fin' => $finPeriodo];
}

function generarConstanciaGratificacionPDF(array $t): string {
    $pdf = new TCPDF('P', 'mm', 'A4', true, 'UTF-8');
    $pdf->SetCreator('Sistema Boletas');
    $pdf->SetAuthor('RRHH');
    $pdf->setPrintHeader(false);
    $pdf->setPrintFooter(false);
    $pdf->SetMargins(10, 10, 10);
    $pdf->AddPage();

    // ============================================================
    // LOGO + CABECERA EMPRESA
    // ============================================================
    $logoPath = dirname(__FILE__) . '/logo_leon_primero.jpeg';
    if (file_exists($logoPath)) {
        $pdf->Image($logoPath, 10, 8, 0, 18, 'JPEG');
    }

    $pdf->SetFont('helvetica', 'B', 12);
    $pdf->SetXY(0, 10);
    $pdf->Cell(0, 6, 'AGROIMPORTADORA LEON S.A.C.', 0, 1, 'C');

    $pdf->SetFont('helvetica', '', 9);
    $pdf->SetX(0);
    $pdf->Cell(0, 5, 'RUC 20608596179', 0, 1, 'C');

    $pdf->SetX(0);
    $pdf->Cell(0, 5, 'URB. COVICORTI MZA. P3 LOTE. 52', 0, 1, 'C');

    $pdf->Ln(0.5);
    $pdf->SetLineWidth(0.3);
    $yLinea1 = $pdf->GetY();
    $pdf->Line(10, $yLinea1, 200, $yLinea1);
    $pdf->Ln(0.5);

    // ============================================================
    // TÍTULO
    // ============================================================
    $pdf->SetFont('helvetica', 'B', 12);
    $pdf->Cell(0, 6, 'CONSTANCIA DE PAGO DE GRATIFICACION', 0, 1, 'C');

    $yLinea2 = $pdf->GetY();
    $pdf->Line(10, $yLinea2, 200, $yLinea2);
    $pdf->Ln(4);

    // ============================================================
    // FECHAS / PERIODOS
    // ============================================================
    $fIngreso = DateTime::createFromFormat('d/m/Y', $t['f_ingreso']) ?: new DateTime();
    $al       = DateTime::createFromFormat('d/m/Y', $t['al']) ?: new DateTime();

    $periodo    = calcularPeriodoGratificacion($al);
    $computable = calcularPeriodoComputable($fIngreso, $periodo['inicio'], $periodo['fin']);

    // ============================================================
    // DATOS DEL TRABAJADOR
    // ============================================================
    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(45, 5.5, 'APELLIDOS Y NOMBRES:', 0, 0);
    $pdf->SetFont('helvetica', '', 9);
    $pdf->Cell(0, 5.5, $t['nombre'], 0, 1);

    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(45, 5.5, 'CARGO:', 0, 0);
    $pdf->SetFont('helvetica', '', 9);
    $pdf->Cell(0, 5.5, $t['puesto'], 0, 1);

    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(45, 5.5, 'DNI:', 0, 0);
    $pdf->SetFont('helvetica', '', 9);
    $pdf->Cell(55, 5.5, $t['dni'], 0, 0);
    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(35, 5.5, 'FECHA DE INGRESO:', 0, 0);
    $pdf->SetFont('helvetica', '', 9);
    $pdf->Cell(0, 5.5, $t['f_ingreso'], 0, 1);

    $pdf->Ln(1.5);

    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(50, 5.5, 'PERIODO DE GRATIFICACION:', 0, 0);
    $pdf->Cell(3, 5.5, '', 0, 0);
    $pdf->SetFont('helvetica', '', 9);
    $pdf->Cell(30, 5.5, $periodo['inicio']->format('d/m/Y'), 0, 0);
    $pdf->Cell(10, 5.5, 'al', 0, 0, 'C');
    $pdf->Cell(0, 5.5, $periodo['fin']->format('d/m/Y'), 0, 1);

    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(50, 5.5, 'PERIODO COMPUTABLE:', 0, 0);
    $pdf->Cell(3, 5.5, '', 0, 0);
    $pdf->SetFont('helvetica', '', 9);
    $pdf->Cell(30, 5.5, $computable['inicio']->format('d/m/Y'), 0, 0);
    $pdf->Cell(10, 5.5, 'al', 0, 0, 'C');
    $pdf->Cell(0, 5.5, $computable['fin']->format('d/m/Y'), 0, 1);

    $meses = (float)$t['meses'];
    $pdf->SetX(63);
    $pdf->Cell(0, 5.5, number_format($meses, 0) . ' meses', 0, 1);
    $pdf->SetX(63);
    $pdf->Cell(0, 5.5, '0 dias', 0, 1);

    $pdf->Ln(3);

    // ============================================================
    // REMUNERACION COMPUTABLE
    // ============================================================
    $pdf->SetFont('helvetica', 'B', 9.5);
    $pdf->Cell(0, 6, 'REMUNERACION COMPUTABLE', 0, 1);

    $labelW = 90; $valW = 0;
    $pdf->SetFont('helvetica', '', 8.5);
    $pdf->Cell(10, 5, '', 0, 0);
    $pdf->Cell($labelW - 10, 5, 'Remuneracion Basica', 0, 0);
    $pdf->Cell($valW, 5, number_format((float)$t['basico'], 2), 0, 1, 'R');

    $pdf->Cell(10, 5, '', 0, 0);
    $pdf->Cell($labelW - 10, 5, 'Asignacion Familiar', 0, 0);
    $pdf->Cell($valW, 5, number_format((float)$t['asignacion_familiar'], 2), 0, 1, 'R');

    $pdf->Cell(10, 5, '', 0, 0);
    $pdf->Cell($labelW - 10, 5, 'Otras remuneraciones', 0, 0);
    $pdf->Cell($valW, 5, number_format((float)$t['otras_rem'], 2), 0, 1, 'R');

    $pdf->Ln(3);

    // ============================================================
    // CALCULO DEL DEPOSITO
    // ============================================================
    $pdf->SetFont('helvetica', 'B', 9.5);
    $pdf->Cell(0, 6, 'CALCULO DEL DEPOSITO', 0, 1);

    $pdf->SetFont('helvetica', '', 8.5);
    $pdf->Cell(10, 5, '', 0, 0);
    $pdf->Cell($labelW - 40, 5, 'Meses', 0, 0);
    $pdf->Cell(95, 5, number_format($meses, 0), 0, 0, 'C');
    $pdf->Cell($valW, 5, number_format((float)$t['rem_computable'], 2), 0, 1, 'R');

    $pdf->Cell(10, 5, '', 0, 0);
    $pdf->Cell($labelW - 40, 5, 'Dias', 0, 0);
    $pdf->Cell(95, 5, '0', 0, 0, 'C');
    $pdf->Cell($valW, 5, '0.00', 0, 1, 'R');

    $pdf->Ln(1);
    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(10, 5.5, '', 0, 0);
    $pdf->Cell($labelW - 40, 5.5, 'TOTAL Gratificacion Bruto:', 0, 0);
    $anchoCaja      = 58;
    $margenDerecho  = $pdf->getMargins()['right'];
    $pdf->SetX($pdf->getPageWidth() - $margenDerecho - $anchoCaja);
    $pdf->Cell(14, 5.5, 'S/', 'TB', 0, 'L');
    $pdf->Cell($anchoCaja - 14, 5.5, number_format((float)$t['rem_computable'], 2), 'TB', 1, 'R');

    $pdf->Ln(3);

    // ============================================================
    // OTROS INGRESOS
    // ============================================================
    $pdf->SetFont('helvetica', 'B', 9.5);
    $pdf->Cell(0, 6, 'OTROS INGRESOS', 0, 1);

    $pdf->SetFont('helvetica', '', 8.5);
    $pdf->Cell(10, 5, '', 0, 0);
    $pdf->MultiCell($labelW - 40, 5, 'Bonificacion Extraordinaria - Ley 29351', 0, 'L', false, 0);
    $pdf->Cell(30, 5, '', 0, 0);
    $pdf->Cell($valW, 5, number_format((float)$t['bono_ley29351'], 2), 0, 1, 'R');

    $pdf->Ln(3);

    $totalFinal = (float)$t['rem_computable'] + (float)$t['bono_ley29351'];

    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(10, 6, '', 0, 0);
    $pdf->MultiCell($labelW - 40, 3, "TOTAL 1/2 Gratificacion\nPercibida (REMYPE)", 0, 'L', false, 0);
    $pdf->SetX($pdf->getPageWidth() - $margenDerecho - $anchoCaja);
    $pdf->Cell(14, 6, 'S/', 'TB', 0, 'L');
    $pdf->Cell($anchoCaja - 14, 6, number_format((float)$t['importe'], 2), 'TB', 1, 'R');

    // ============================================================
    // FIRMAS — al fondo de la página
    // ============================================================
    $yFirma = 260;

    $firmaPath = dirname(__FILE__) . '/firmaagro (1).jpg';
    if (file_exists($firmaPath)) {
        $pdf->Image($firmaPath, 39, $yFirma - 18, 35, 0, 'JPG');
    }

    $pdf->Line(20,  $yFirma, 95,  $yFirma);
    $pdf->Line(105, $yFirma, 190, $yFirma);

    $pdf->SetXY(20, $yFirma + 2);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(75, 4, 'EMPLEADOR', 0, 0, 'C');
    $pdf->SetXY(20, $yFirma + 6);
    $pdf->Cell(75, 4, 'RUC 20608596179', 0, 0, 'C');

    $pdf->SetXY(105, $yFirma + 2);
    $pdf->Cell(85, 4, $t['nombre'], 0, 1, 'C');
    $pdf->SetXY(105, $yFirma + 6);
    $pdf->Cell(85, 4, $t['dni'], 0, 0, 'C');

    // ============================================================
    // GUARDAR
    // ============================================================
    $safe = preg_replace('/[^A-Za-z0-9]/', '_', $t['nombre']);
    $per  = preg_replace('/[^A-Za-z0-9]/', '_', $periodo['fin']->format('Y_m'));
    $nombreArchivo = dirname(__FILE__) . '/pdfs/gratificacion_' . $safe . '_' . $per . '.pdf';
    $pdf->Output($nombreArchivo, 'F');
    return $nombreArchivo;
}