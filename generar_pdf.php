<?php
require_once 'vendor/autoload.php';

function nombreBanco($sigla) {
    $bancos = [
        'BCP'        => 'BCP - BANCO DE CREDITO DEL PERU',
        'BBVA'       => 'BBVA - BANCO CONTINENTAL',
        'SCOTIABANK' => 'SCOTIABANK PERU',
        'INTERBANK'  => 'INTERBANK',
        'BN'         => 'BN - BANCO DE LA NACION',
        'BANBIF'     => 'BANBIF - BANCO INTERAMERICANO DE FINANZAS',
        'PICHINCHA'  => 'BANCO PICHINCHA',
        'MIBANCO'    => 'MIBANCO',
        'GNB'        => 'GNB - BANCO GNB PERU',
        'ALFIN'      => 'ALFIN BANCO',
        'CITIBANK'   => 'CITIBANK PERU',
        'COMERCIO'   => 'BANCO DE COMERCIO',
    ];
    $sigla = strtoupper(trim($sigla));
    return $bancos[$sigla] ?? $sigla;
}

function generarBoletaPDF(array $t): string {
    $pdf = new TCPDF('P', 'mm', 'A4', true, 'UTF-8');
    $pdf->SetCreator('Sistema Boletas');
    $pdf->SetAuthor('RRHH');
    $pdf->setPrintHeader(false);
    $pdf->setPrintFooter(false);
    $pdf->SetMargins(10, 10, 10);
    $pdf->AddPage();

    $pageW = 190;

    // ============================================================
    // LOGO + CABECERA EMPRESA
    // ============================================================
    $logoPath = dirname(__FILE__) . '/logo_leon_primero.jpeg';
    if (file_exists($logoPath)) {
        $pdf->Image($logoPath, 10, 8, 0, 14, 'JPEG');
    }

    $pdf->SetXY(38, 9);
    $pdf->SetFont('helvetica', 'B', 11);
    $pdf->Cell(0, 5, 'AGROIMPORTADORA  LEON S.A.C.', 0, 1);

    $pdf->SetXY(38, 14);
    $pdf->SetFont('helvetica', '', 8);
    $pdf->Cell(0, 4, 'R.U.C.:  20608596179', 0, 1);

    $pdf->SetXY(38, 18);
    $pdf->Cell(0, 4, 'URB.  COVICORTI  MZA.  P3  LOTE.  52', 0, 1);

    $pdf->SetY(26);

    // ============================================================
    // TÍTULO
    // ============================================================
    $pdf->SetFont('helvetica', 'B', 12);
    $pdf->Cell(0, 7, 'BOLETA  DE  REMUNERACIONES', 0, 1, 'C');
    $pdf->SetFont('helvetica', '', 9);
    $pdf->Cell(0, 4, $t['periodo'], 0, 1, 'C');
    $pdf->Ln(2);

    // ============================================================
    // DATOS DEL TRABAJADOR
    // ============================================================
    $lw = 24; $vw = 46; $lw2 = 10; $vw2 = 40; $lw3 = 26;

    // Fila 1
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw,  4.5, 'CÓDIGO.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($vw,  4.5, $t['dni'], 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw2, 4.5, 'DNI', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($vw2, 4.5, $t['dni'], 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw3, 4.5, 'BASICO.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0,    4.5, number_format((float)$t['basico'], 1), 0, 1);

    // Fila 2
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw, 4.5, 'TRABAJADOR.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($vw + $lw2 + $vw2, 4.5, $t['nombre'], 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw3, 4.5, 'FEC.NACIM.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0, 4.5, $t['f_nacimiento'], 0, 1);

    // Fila 3
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw, 4.5, 'CARGO.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($vw + $lw2 + $vw2, 4.5, $t['puesto'], 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw3, 4.5, 'FEC. INGRESO.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0, 4.5, $t['f_ingreso'], 0, 1);

    // Fila 4
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw, 4.5, 'S.P.P.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($vw, 4.5, $t['afp_onp'], 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw2 + 4, 4.5, 'CUSSP.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($vw2 - 4, 4.5, $t['cussp'], 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw3, 4.5, 'FEC.CESE:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0, 4.5, $t['f_cese'], 0, 1);

    // Fila 5
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($lw, 4.5, 'CALIFICACIÓN.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0, 4.5, 'Ninguno', 0, 1);

    $pdf->Ln(2);

    // ============================================================
    // DÍAS Y HORAS
    // ============================================================
    $diasLab = (float)$t['dias_trab'] + (float)$t['dias_descanso'] + (float)$t['Vacaciones'] + (float)$t['dias_feriados'] + (float)$t['feriado_lab_dias'];
    $hrNorm  = $diasLab * 8;
    $hrExtra = (float)$t['horas_extras'];

    $DL = 20; $DV = 14;
    $HL = 24; $HV = 12;
    $LL = 34; $LV = 12;
    $VL = 28; $VV = $pageW - $DL - $DV - $HL - $HV - $LL - $LV - $VL;

    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($DL, 4.5, 'D.LABOR. :', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($DV, 4.5, number_format($diasLab, 1), 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($HL, 4.5, 'D. NO LAB. :', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($HV, 4.5, '0.00', 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($LL, 4.5, 'D.LIC.CON GOC.HAB.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($LV, 4.5, '0.00', 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($VL, 4.5, 'D.VACACIONES:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0, 4.5, '0.00', 0, 1);

    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($DL, 4.5, 'HR.NORM. :', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($DV, 4.5, number_format($hrNorm, 2), 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($HL, 4.5, 'D.SUSP.DISCIPL. :', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($HV, 4.5, '0.00', 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($LL, 4.5, 'D.LIC. SIN GOC.HAB.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($LV, 4.5, '0.00', 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($VL, 4.5, 'D.DESC.MED.:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0, 4.5, number_format((float)($t['descanso_medico_dias'] ?? 0), 2), 0, 1);

    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($DL, 4.5, 'HR.Ex. :', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($DV, 4.5, '0.00', 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($HL, 4.5, 'HR.Ex35. :', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($HV, 4.5, '0.00', 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($LL, 4.5, 'D.LIC.PATERNIDAD:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell($LV, 4.5, '0.00', 0, 0);
    $pdf->SetFont('helvetica', 'B', 7.5);
    $pdf->Cell($VL, 4.5, 'D. SUBSIDIO:', 0, 0);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(0, 4.5, '0.00', 0, 1);

    $pdf->Ln(2);

    // ============================================================
    // TABLA PRINCIPAL
    // ============================================================
    $c1n = 38; $c1v = 18; $c1 = $c1n + $c1v;
    $c2n = 28; $c2v = 14; $c2 = $c2n + $c2v;
    $c3n = 34; $c3v = 16; $c3 = $c3n + $c3v;
    $c4n = 26; $c4v = 16; $c4 = $c4n + $c4v;

    $rowH = 5.5;

    $pdf->SetFont('helvetica', 'B', 7.5);
    $yH = $pdf->GetY();
    $pdf->Cell($c1, 7, 'REMUNERACIONES',           1, 0, 'C');
    $pdf->Cell($c2, 7, 'NO REMUNERATIVOS',          1, 0, 'C');
    $pdf->Cell($c3, 7, 'DESCUENTOS AL TRABAJADOR',  1, 0, 'C');
    $pdf->MultiCell($c4, 3.5, "CONTRIBUCIONES DEL\nEMPLEADOR", 1, 'C');
    $pdf->SetXY(10, $yH + 7);

    $pdf->SetFont('helvetica', '', 7.5);

    $fmt = function($v) {
        return ($v !== '' && (float)$v != 0) ? number_format((float)$v, 2) : '';
    };

    // ============================================================
    // COMPACTAR cada columna: solo filas con valor, luego rellenar
    // ============================================================

    // Columna 1 — Remuneraciones
    $rem_all = [
        ['BASICO',           $t['monto_basico']],
        ['ASIG. FAMILIAR',   $t['monto_asig']],
        ['BONO ASISTENCIA',  $t['bono_asistencia']],
        ['FERIADO LAB.',     $t['feriado_lab']],
        ['BONO PRODUCCIÓN.', $t['bono_horas']],
        ['DESC. MEDICO',     $t['descanso_medico_monto'] ?? 0],
    ];

    // Columna 2 — No remunerativos
    $no_rem_all = [
        ['MOVILIDAD',    $t['movilidad']],
        ['ALIMENTACIÓN', $t['viaticos']],
    ];

    // Columna 3 — Descuentos
    $desc_all = [
        ['FONDO AFP',        $t['afp_10']],
        ['SEGURO AFP',       $t['seg_afp']],
        ['ONP',              $t['onp']],
        ['RENTA 5TA.',       $t['renta_5ta']],
        ['OTROS DESCUENTOS', $t['otros_desc']],
        ['ADELANTOS VIATICOS', $t['adelantos']],
        ['ADELANTO QUINCENA',  $t['adelanto_quincena']],
    ];

    // Columna 4 — Contribuciones
    $contrib_all = [
        ['ESSALUD', $t['essalud']],
    ];

    // Filtrar solo los que tienen valor
    $hasVal = function($item) { return $item[1] !== '' && (float)$item[1] != 0; };

    $col1 = array_values(array_filter($rem_all,     $hasVal));
    $col2 = array_values(array_filter($no_rem_all,  $hasVal));
    $col3 = array_values(array_filter($desc_all,    $hasVal));
    $col4 = array_values(array_filter($contrib_all, $hasVal));

    // Número de filas = la columna más larga
    $totalFilas = max(count($col1), count($col2), count($col3), count($col4));
    // Mínimo 5 filas para que la tabla no quede muy pequeña
    $totalFilas = max($totalFilas, 5);

    // Rellenar con filas vacías hasta $totalFilas
    $pad = function(array $col, int $n) {
        while (count($col) < $n) $col[] = ['', ''];
        return $col;
    };

    $col1 = $pad($col1, $totalFilas);
    $col2 = $pad($col2, $totalFilas);
    $col3 = $pad($col3, $totalFilas);
    $col4 = $pad($col4, $totalFilas);

    $x1 = 10;
    $x2 = $x1 + $c1;
    $x3 = $x2 + $c2;
    $x4 = $x3 + $c3;

    $yStart = $pdf->GetY();

    for ($i = 0; $i < $totalFilas; $i++) {
        $pdf->Cell($c1n, $rowH, $col1[$i][0], 0, 0, 'L');
        $pdf->Cell($c1v, $rowH, $fmt($col1[$i][1]), 0, 0, 'R');
        $pdf->Cell($c2n, $rowH, $col2[$i][0], 0, 0, 'L');
        $pdf->Cell($c2v, $rowH, $fmt($col2[$i][1]), 0, 0, 'R');
        $pdf->Cell($c3n, $rowH, $col3[$i][0], 0, 0, 'L');
        $pdf->Cell($c3v, $rowH, $fmt($col3[$i][1]), 0, 0, 'R');
        $pdf->Cell($c4n, $rowH, $col4[$i][0], 0, 0, 'L');
        $pdf->Cell($c4v, $rowH, $fmt($col4[$i][1]), 0, 1, 'R');
    }

    $totalAltoDatos = $totalFilas * $rowH;
    $pdf->Rect($x1, $yStart, $c1, $totalAltoDatos, 'D');
    $pdf->Rect($x2, $yStart, $c2, $totalAltoDatos, 'D');
    $pdf->Rect($x3, $yStart, $c3, $totalAltoDatos, 'D');
    $pdf->Rect($x4, $yStart, $c4, $totalAltoDatos, 'D');

    $yTot = $pdf->GetY();
    $pdf->SetFont('helvetica', 'B', 7.5);

    $pdf->Cell($c1n, $rowH, 'Total Ingresos S/',   0, 0, 'L');
    $pdf->Cell($c1v, $rowH, number_format((float)$t['rem_bruta'],    2), 0, 0, 'R');
    $pdf->Cell($c2n, $rowH, 'Total No Remu. S/',   0, 0, 'L');
    $pdf->Cell($c2v, $rowH, number_format((float)$t['total_no_rem'], 2), 0, 0, 'R');
    $pdf->Cell($c3n, $rowH, 'Total Descuentos S/', 0, 0, 'L');
    $pdf->Cell($c3v, $rowH, number_format((float)$t['total_desc'],   2), 0, 0, 'R');
    $pdf->Cell($c4n, $rowH, 'Total Aportacion S/', 0, 0, 'L');
    $pdf->Cell($c4v, $rowH, number_format((float)$t['essalud'],      2), 0, 1, 'R');

    $pdf->Rect($x1, $yTot, $c1, $rowH, 'D');
    $pdf->Rect($x2, $yTot, $c2, $rowH, 'D');
    $pdf->Rect($x3, $yTot, $c3, $rowH, 'D');
    $pdf->Rect($x4, $yTot, $c4, $rowH, 'D');

    $pdf->Ln();
    $pdf->Ln(4);

    // ============================================================
    // TIPO PAGO + NETO A PAGAR
    // ============================================================
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(30, 5, 'Tipo Pago:', 0, 0);
    $pdf->Cell(110, 5, nombreBanco($t['banco']) . '  Cta. Cte. ' . $t['cuenta'], 0, 0);

    $pdf->SetFont('helvetica', 'B', 10);
    $pdf->Cell(24, 8, 'Neto a Pagar', 1, 0, 'C');
    $pdf->SetFont('helvetica', 'B', 12);
    $pdf->Cell(26, 8, number_format((float)$t['total_a_pagar'], 2), 1, 1, 'C');

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

    $pdf->SetXY(105, $yFirma + 2);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(85, 4, $t['nombre'], 0, 1, 'C');

    $pdf->SetXY(105, $yFirma + 6);
    $pdf->Cell(85, 4, 'DNI ' . $t['dni'], 0, 0, 'C');

    // ============================================================
    // GUARDAR
    // ============================================================
    $safe = preg_replace('/[^A-Za-z0-9]/', '_', $t['nombre']);
    $per  = preg_replace('/[^A-Za-z0-9]/', '_', $t['periodo']);
    $nombreArchivo = dirname(__FILE__) . '/pdfs/boleta_' . $safe . '_' . $per . '.pdf';
    $pdf->Output($nombreArchivo, 'F');
    return $nombreArchivo;
}