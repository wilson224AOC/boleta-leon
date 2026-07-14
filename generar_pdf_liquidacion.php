<?php
require_once 'vendor/autoload.php';

// ============================================================
// NÚMERO A LETRAS (para el monto en el recibo)
// ============================================================
function _numLiqBloque(int $n): string {
    $unidades = ['', 'UNO','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE',
        'DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISEIS','DIECISIETE','DIECIOCHO','DIECINUEVE','VEINTE'];
    $decenas  = ['','DIEZ','VEINTE','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA'];
    $centenas = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS','SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS'];

    if ($n == 0) return '';
    if ($n == 100) return 'CIEN';

    $texto = '';
    $c = intdiv($n, 100);
    $r = $n % 100;
    if ($c > 0) $texto .= $centenas[$c];
    if ($r > 0) {
        if ($texto !== '') $texto .= ' ';
        if ($r <= 20) {
            $texto .= $unidades[$r];
        } else {
            $d = intdiv($r, 10);
            $u = $r % 10;
            $texto .= $decenas[$d];
            if ($u > 0) $texto .= ' Y ' . $unidades[$u];
        }
    }
    return $texto;
}

function numeroALetrasLiq(int $num): string {
    if ($num == 0) return 'CERO';
    if ($num < 0)  return 'MENOS ' . numeroALetrasLiq(-$num);

    $millones = intdiv($num, 1000000);
    $resto    = $num % 1000000;
    $miles    = intdiv($resto, 1000);
    $cientos  = $resto % 1000;

    $partes = [];
    if ($millones > 0) {
        $partes[] = ($millones == 1) ? 'UN MILLON' : (_numLiqBloque($millones) . ' MILLONES');
    }
    if ($miles > 0) {
        $partes[] = ($miles == 1) ? 'MIL' : (_numLiqBloque($miles) . ' MIL');
    }
    if ($cientos > 0) {
        $partes[] = _numLiqBloque($cientos);
    }
    return trim(implode(' ', $partes));
}

function montoEnLetrasLiq(float $monto): string {
    $entero   = (int)floor($monto + 0.0000001);
    $centavos = (int)round(($monto - $entero) * 100);
    if ($centavos >= 100) { $entero++; $centavos -= 100; }
    return numeroALetrasLiq($entero) . ' Y ' . str_pad((string)$centavos, 2, '0', STR_PAD_LEFT) . '/100';
}

// ============================================================
// FECHAS
// ============================================================
function fechaLargaLiq(string $fechaDMY, bool $mayus = false): string {
    if (empty($fechaDMY)) return '';
    $partes = explode('/', $fechaDMY);
    if (count($partes) !== 3) return $fechaDMY;
    [$d, $m, $y] = $partes;
    $meses = [1=>'Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    $mesNombre = $meses[(int)$m] ?? '';
    $texto = sprintf('%02d de %s de %s', (int)$d, $mesNombre, $y);
    return $mayus ? mb_strtoupper($texto, 'UTF-8') : $texto;
}

function sumarDiasFechaLiq(string $fechaDMY, int $dias): string {
    if (empty($fechaDMY)) return '';
    $dt = DateTime::createFromFormat('d/m/Y', $fechaDMY);
    if (!$dt) return $fechaDMY;
    $dt->modify(($dias >= 0 ? '+' : '') . $dias . ' day');
    return $dt->format('d/m/Y');
}

// ============================================================
// CABECERA (se repite en ambas páginas)
// ============================================================
function dibujarCabeceraLiquidacion($pdf, array $t) {
    $pdf->SetXY(10, 10);
    $pdf->SetFont('helvetica', 'B', 12);
    $pdf->Cell(190, 6, 'LIQUIDACIÓN  DE  BENEFICIOS  SOCIALES', 0, 1, 'C');

    $y = 20;
    $filas = [
        ['APELLIDOS Y NOMBRES', $t['nombre']],
        ['FECHA DE INGRESO',    fechaLargaLiq($t['f_ingreso'])],
        ['CARGO',               $t['puesto']],
        ['REGIMEN',             $t['regimen']],
        ['FECHA DE CESE',       fechaLargaLiq($t['f_cese'])],
        ['CAUSA DE RETIRO',     $t['causa']],
    ];

    $pdf->SetXY(10, $y);
    foreach ($filas as $f) {
        $pdf->SetX(10);
        $pdf->SetFont('helvetica', 'B', 8.5);
        $pdf->Cell(42, 4.3, $f[0], 0, 0);
        $pdf->SetFont('helvetica', '', 8.5);
        $pdf->Cell(0, 4.3, $f[1], 0, 1);
    }

    $yLinea = $pdf->GetY() + 1.5;
    $pdf->Line(10, $yLinea, 200, $yLinea);
    $pdf->SetY($yLinea + 4);
}

// ============================================================
// BLOQUE "REMUNERACIÓN COMPUTABLE" (se repite 3 veces en pág. 1)
// ============================================================
function dibujarRemComputableLiq($pdf, array $t, string $labelBono, string $labelTotal) {
    $pdf->SetX(30);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell(0, 4.3, 'REMUNERACIÓN COMPUTABLE', 0, 1);

    $x = 40; $lw = 42; $sw = 8; $vw = 22;

    $filas = [
        ['BASICO',       (float)$t['basico']],
        ['ASIG.FAMILIAR', (float)$t['asig_familiar']],
        [$labelBono,      (float)$t['bono_asistencia']],
    ];
    foreach ($filas as $f) {
        $pdf->SetX($x);
        $pdf->SetFont('helvetica', '', 8);
        $pdf->Cell($lw, 4.3, $f[0], 0, 0);
        $pdf->Cell($sw, 4.3, 'S/.', 0, 0, 'R');
        $pdf->Cell($vw, 4.3, number_format($f[1], 2), 0, 1, 'R');
    }

    $yLinea = $pdf->GetY() + 0.5;
    $pdf->Line($x, $yLinea, $x + $lw + $sw + $vw, $yLinea);
    $pdf->SetY($yLinea + 1);

    $pdf->SetX($x);
    $pdf->SetFont('helvetica', 'BU', 8);
    $pdf->Cell($lw, 4.3, $labelTotal, 0, 0);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell($sw, 4.3, 'S/.', 0, 0, 'R');
    $pdf->Cell($vw, 4.3, number_format((float)$t['rem_computable'], 2), 0, 1, 'R');
    $pdf->Ln(2.5);
}

// ============================================================
// SECCIÓN CON DETALLE Y TOTAL (Otros ingresos / Descuentos, pág. 2)
// ============================================================
function dibujarSeccionDetalleLiq($pdf, string $titulo, float $total, array $items) {
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 9);
    $pdf->Cell(148, 5, $titulo, 0, 0);
    $pdf->Cell(10, 5, 'S/.', 0, 0, 'R');
    $pdf->Cell(22, 5, number_format($total, 2), 0, 1, 'R');

    foreach ($items as $it) {
        $pdf->SetX(20);
        $pdf->SetFont('helvetica', '', 8);
        $pdf->Cell(48, 4.3, $it['label'], 0, 0);
        $pdf->Cell(20, 4.3, $it['pct'] ?? '', 0, 0);
        $pdf->Cell(10, 4.3, 'S/.', 0, 0, 'R');
        $pdf->Cell(22, 4.3, number_format($it['monto'], 2), 0, 1, 'R');
    }
    $pdf->Ln(3);
}

// ============================================================
// GENERAR PDF COMPLETO
// ============================================================
function generarLiquidacionPDF(array $t): string {
    $pdf = new TCPDF('P', 'mm', 'A4', true, 'UTF-8');
    $pdf->SetCreator('Sistema Liquidaciones');
    $pdf->SetAuthor('RRHH');
    $pdf->setPrintHeader(false);
    $pdf->setPrintFooter(false);
    $pdf->SetMargins(10, 10, 10);
    $pdf->SetAutoPageBreak(false);

    $N = (float)$t['rem_computable'];

    // ============================================================
    // PÁGINA 1 — CTS, VACACIONES TRUNCAS, GRATIFICACIÓN
    // ============================================================
    $pdf->AddPage();
    dibujarCabeceraLiquidacion($pdf, $t);

    // ── COMPENSACIÓN POR TIEMPO DE SERVICIOS ──
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 9.5);
    $pdf->Cell(0, 5, 'COMPENSACIÓN POR TIEMPO DE SERVICIOS', 0, 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell(45, 4.3, 'DEPOSITOS YA EFECTUADOS', 0, 0);
    $pdf->SetFont('helvetica', '', 8);
    $pdf->Cell(0, 4.3, 'Ningún Periodo', 0, 1);
    $pdf->Ln(1);

    dibujarRemComputableLiq($pdf, $t, 'BONO ASISTENCIA', 'TOTAL BASE CTS');

    $ctsMeses = (int)round((float)$t['cts_meses']);
    $ctsDias  = (int)round((float)$t['cts_dias']);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell(0, 4.3, 'SEMESTRE QUE SE LIQUIDA', 0, 1);
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', '', 8);
    $pdf->Cell(0, 4.3, 'Del ' . $t['f_ingreso'] . ' al ' . $t['f_cese'] . ' : ' . $ctsMeses . ' meses ' . str_pad((string)$ctsDias, 2, '0', STR_PAD_LEFT) . ' dias', 0, 1);
    $pdf->Ln(1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell(0, 4.3, 'CÁLCULO', 0, 1);

    $ctsMesesVal = $ctsMeses > 0 ? ($N / 12 / 2 * $ctsMeses) : 0;
    $ctsDiasVal  = $ctsDias  > 0 ? ($N / 12 / 30 / 2 * $ctsDias) : 0;

    $colW = 92;
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'BU', 8);
    $pdf->Cell($colW, 4.3, 'Por los Meses Completos:', 0, 0);
    $pdf->Cell($colW, 4.3, 'Por los Días:', 0, 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', '', 8);
    $pdf->Cell(9, 4.3, 'S/.', 0, 0);
    $pdf->Cell($colW - 9, 4.3, number_format($N, 2) . ' / 12 / 2  x ' . $ctsMeses . ' meses', 0, 0);
    $pdf->Cell(9, 4.3, 'S/.', 0, 0);
    $pdf->Cell($colW - 9, 4.3, number_format($N, 2) . ' / 12 / 30 / 2  x ' . $ctsDias . ' días', 0, 1);

    $pdf->SetX(10);
    $pdf->Cell($colW, 4.3, number_format($ctsMesesVal, 2), 0, 0, 'R');
    $pdf->Cell($colW, 4.3, number_format($ctsDiasVal, 2), 0, 1, 'R');

    $yLinea = $pdf->GetY() + 0.5;
    $pdf->Line(10, $yLinea, 190, $yLinea);
    $pdf->SetY($yLinea + 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'BU', 8.5);
    $pdf->Cell(60, 4.5, 'TOTAL C.T.S.', 0, 0);
    $pdf->SetFont('helvetica', 'B', 8.5);
    $pdf->Cell(10, 4.5, 'S/.', 0, 0, 'R');
    $pdf->Cell(30, 4.5, number_format((float)$t['cts_monto'], 2), 0, 1, 'R');
    $pdf->Ln(3);

    // ── VACACIONES TRUNCAS ──
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 9.5);
    $pdf->Cell(0, 5, 'VACACIONES TRUNCAS', 0, 1);

    dibujarRemComputableLiq($pdf, $t, 'BONO DE ASISTENCIA', 'TOTAL BASE VACACIONES');

    $vacMeses = (int)round((float)$t['vac_meses']);
    $vacDias  = (int)round((float)$t['vac_dias']);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell(0, 4.3, 'PERIODOS VACACIONALES ADEUDADOS', 0, 1);
    $pdf->Ln(0.5);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', '', 8);
    $pdf->Cell(90, 4.3, 'Vacaciones Truncas  Del ' . $t['f_ingreso'] . ' al ' . $t['f_cese'], 0, 0);
    $pdf->Cell(30, 4.3, $vacMeses . ' meses, ' . $vacDias . ' dias', 0, 0);
    $pdf->Cell(10, 4.3, 'S/.', 0, 0, 'R');
    $pdf->Cell(30, 4.3, number_format((float)$t['vac_monto'], 2), 0, 1, 'R');

    $yLinea = $pdf->GetY() + 0.5;
    $pdf->Line(10, $yLinea, 190, $yLinea);
    $pdf->SetY($yLinea + 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'BU', 8.5);
    $pdf->Cell(120, 4.5, 'TOTAL VACAC. TRUNCAS', 0, 0);
    $pdf->SetFont('helvetica', 'B', 8.5);
    $pdf->Cell(10, 4.5, 'S/.', 0, 0, 'R');
    $pdf->Cell(30, 4.5, number_format((float)$t['vac_monto'], 2), 0, 1, 'R');
    $pdf->Ln(3);

    // ── GRATIFICACIÓN ──
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 9.5);
    $pdf->Cell(0, 5, 'GRATIFICACIÓN', 0, 1);

    dibujarRemComputableLiq($pdf, $t, 'BONO ASISTENCIA', 'TOTAL BASE GRATIFICACIONES');

    $gratMeses = (int)round((float)$t['grat_meses']);
    $inicioGrat = sumarDiasFechaLiq($t['f_ingreso'], 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell(0, 4.3, 'PERIODO COMPUTABLE', 0, 1);
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', '', 8);
    $pdf->Cell(0, 4.3, 'Del ' . $inicioGrat . ' al ' . $t['f_cese'] . ' : ' . $gratMeses . ' meses 0 dias', 0, 1);
    $pdf->Ln(1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 8);
    $pdf->Cell(0, 4.3, 'CÁLCULO', 0, 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'BU', 8);
    $pdf->Cell(0, 4.3, 'Por los Meses Completos:', 0, 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', '', 8);
    $pdf->Cell(9, 4.3, 'S/.', 0, 0);
    $pdf->Cell(0, 4.3, number_format($N, 2) . ' / 6 / 2 x ' . $gratMeses . ' meses', 0, 1);

    $pdf->SetX(10);
    $pdf->Cell(90, 4.3, number_format((float)$t['grat_base'], 2), 0, 1, 'R');

    $yLinea = $pdf->GetY() + 0.5;
    $pdf->Line(10, $yLinea, 190, $yLinea);
    $pdf->SetY($yLinea + 1);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'BU', 8.5);
    $pdf->Cell(60, 4.5, 'TOTAL GRATIFIC.', 0, 0);
    $pdf->SetFont('helvetica', 'B', 8.5);
    $pdf->Cell(10, 4.5, 'S/.', 0, 0, 'R');
    $pdf->Cell(30, 4.5, number_format((float)$t['grat_base'], 2), 0, 1, 'R');

    // ============================================================
    // PÁGINA 2 — OTROS INGRESOS, DESCUENTOS, NETO A PAGAR, FIRMAS
    // ============================================================
    $pdf->AddPage();
    dibujarCabeceraLiquidacion($pdf, $t);

    // Otros ingresos (Bonificación Ley 29351)
    $itemsIngresos = [];
    if ((float)$t['bono29351'] != 0) {
        $itemsIngresos[] = ['label' => 'BONIF.LEY 29351', 'monto' => (float)$t['bono29351']];
    }
    dibujarSeccionDetalleLiq($pdf, 'OTROS INGRESOS', (float)$t['bono29351'], $itemsIngresos);

    // Descuentos (según AFP u ONP)
    $esOnp = (strtoupper(trim($t['afp_onp'])) === 'ONP');
    $itemsDesc = [];
    $totalDesc = 0;

    if ($esOnp) {
        $itemsDesc[] = ['label' => 'FONDO AFP/ONP', 'monto' => (float)$t['descuento_onp']];
        $totalDesc  += (float)$t['descuento_onp'];
    } else {
        $itemsDesc[] = ['label' => 'FONDO AFP/ONP', 'monto' => (float)$t['descuento_afp']];
        $itemsDesc[] = ['label' => 'SEGURO AFP', 'pct' => '1.37 %', 'monto' => (float)$t['descuento_segafp']];
        $totalDesc  += (float)$t['descuento_afp'] + (float)$t['descuento_segafp'];
    }
    if ((float)$t['comedor'] != 0) {
        $itemsDesc[] = ['label' => 'COMEDOR', 'monto' => (float)$t['comedor']];
        $totalDesc  += (float)$t['comedor'];
    }

    dibujarSeccionDetalleLiq($pdf, 'DESCUENTOS ' . $t['afp_onp'], $totalDesc, $itemsDesc);

    $pdf->Ln(4);

    // Neto a pagar
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', 'B', 11);
    $pdf->Cell(148, 7, 'NETO A PAGAR', 0, 0);
    $pdf->Cell(10, 7, 'S/.', 0, 0, 'R');
    $pdf->SetFont('helvetica', 'B', 12);
    $pdf->Cell(22, 7, number_format((float)$t['neto_liquidacion'], 2), 0, 1, 'R');
    $pdf->Ln(6);

    // Declaración
    $montoLetras = montoEnLetrasLiq((float)$t['neto_liquidacion']);
    $pdf->SetX(10);
    $pdf->SetFont('helvetica', '', 8.5);
    $pdf->MultiCell(190, 4.3,
        'He recibido de AGROIMPORTADORA LEON S.A.C. la suma de: ' . $montoLetras .
        ' Soles por concepto de mi liquidación por tiempo de servicio, a la cual la encuentro conforme y a mi entera satisfacción, no teniendo reclamo alguno por formular',
        0, 'J');
    $pdf->Ln(4);

    $pdf->SetX(10);
    $pdf->SetFont('helvetica', '', 8.5);
    $pdf->Cell(0, 4.3, 'TRUJILLO, ' . fechaLargaLiq($t['f_cese'], true), 0, 1, 'R');

    // Firmas
    $yFirma = 260;

    $firmaPath = dirname(__FILE__) . '/firmaagro (1).jpg';
    if (file_exists($firmaPath)) {
        $pdf->Image($firmaPath, 39, $yFirma - 18, 35, 0, 'JPG');
    }

    $pdf->Line(20,  $yFirma, 95,  $yFirma);
    $pdf->Line(105, $yFirma, 190, $yFirma);

    $pdf->SetXY(20, $yFirma + 2);
    $pdf->SetFont('helvetica', '', 7.5);
    $pdf->Cell(75, 4, 'FIRMA Y SELLO DEL EMPLEADOR', 0, 0, 'C');

    $pdf->SetXY(105, $yFirma + 2);
    $pdf->Cell(85, 4, 'FIRMA Y HUELLA DIGITAL', 0, 1, 'C');

    $pdf->SetXY(105, $yFirma + 6);
    $pdf->Cell(85, 4, $t['nombre'], 0, 0, 'C');

    $pdf->SetXY(105, $yFirma + 10);
    $pdf->Cell(85, 4, $t['dni'], 0, 0, 'C');

    // ============================================================
    // GUARDAR
    // ============================================================
    $safe = preg_replace('/[^A-Za-z0-9]/', '_', $t['nombre']);
    $nombreArchivo = dirname(__FILE__) . '/pdfs/liquidacion_' . $safe . '_' . $t['dni'] . '.pdf';
    $pdf->Output($nombreArchivo, 'F');
    return $nombreArchivo;
}