<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Generador de Boletas – Agroimportadora Leon S.A.C.</title>
    <style>
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            min-height: 100vh;
            background: #0d1b2a;
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 40px 16px 60px;
        }

        /* ── HEADER ─────────────────────────── */
        .header {
            width: 100%;
            max-width: 1400px;
            display: flex;
            align-items: center;
            gap: 18px;
            margin-bottom: 32px;
        }
        .header img {
            height: 64px;
            border-radius: 8px;
            background: #fff;
            padding: 4px 8px;
            object-fit: contain;
        }
        .header-text { flex: 1 1 auto; }
        .header-text h1 { font-size: 22px; font-weight: 700; color: #ffffff; line-height: 1.2; }
        .header-text p  { font-size: 13px; color: #7a9cbf; margin-top: 3px; }

        /* ── TOGGLE DE MODO (arriba a la derecha) ───────────────── */
        .mode-toggle {
            display: flex;
            gap: 4px;
            background: #152336;
            border: 1px solid #1e3450;
            border-radius: 999px;
            padding: 3px;
            flex-shrink: 0;
        }
        .mode-btn {
            border: none;
            background: transparent;
            color: #7a9cbf;
            font-size: 12px;
            font-weight: 600;
            padding: 7px 14px;
            border-radius: 999px;
            cursor: pointer;
            transition: background .2s, color .2s;
            white-space: nowrap;
        }
        .mode-btn:hover { color: #c5d8ef; }
        .mode-btn.active { background: #1a6bbf; color: #fff; }

        /* ── LAYOUT DOS COLUMNAS ────────────── */
        .layout {
            width: 100%;
            max-width: 1400px;
            display: flex;
            gap: 24px;
            align-items: flex-start;
        }

        /* ── CARD ───────────────────────────── */
        .card {
            background: #152336;
            border: 1px solid #1e3450;
            border-radius: 16px;
            padding: 36px 40px;
        }
        .card-left  { flex: 1 1 0; min-width: 0; }
        .card-right {
            width: 420px;
            flex-shrink: 0;
            display: none; /* se muestra al cargar Excel */
        }

        .section-title {
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: .08em;
            color: #4a90d9;
            margin-bottom: 14px;
        }

        /* ── DROP ZONE ──────────────────────── */
        .drop-zone {
            border: 2px dashed #2a4a6e;
            border-radius: 12px;
            background: #0f1e2e;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: border-color .25s, background .25s;
            position: relative;
        }
        .drop-zone:hover, .drop-zone.dragover { border-color: #4a90d9; background: #132334; }
        .drop-zone input[type="file"] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
        .drop-icon  { font-size: 42px; margin-bottom: 12px; }
        .drop-label { font-size: 15px; color: #c5d8ef; font-weight: 500; }
        .drop-sub   { font-size: 12px; color: #4a6a8a; margin-top: 6px; }
        .file-chosen { margin-top: 14px; font-size: 13px; color: #4a90d9; font-weight: 600; display: none; }

        /* ── PROGRESO ───────────────────────── */
        .progress-wrap { display: none; margin-top: 18px; }
        .progress-bar-bg { background: #0f1e2e; border-radius: 99px; height: 6px; overflow: hidden; }
        .progress-bar-fill { height: 100%; width: 0%; background: #4a90d9; border-radius: 99px; transition: width .4s; }
        .progress-label { font-size: 12px; color: #4a6a8a; margin-top: 6px; text-align: center; }

        /* ── DIVIDER ────────────────────────── */
        .divider { height: 1px; background: #1e3450; margin: 28px 0; }

        /* ── TABLA ──────────────────────────── */
        #preview-section { display: none; }

        .toolbar { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin-bottom: 14px; }
        .toolbar-left { display: flex; align-items: center; gap: 10px; }
        .badge-count { background: #1a3a5c; color: #7ab8e8; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 99px; }

        .search-input {
            background: #0f1e2e;
            border: 1px solid #2a4a6e;
            border-radius: 8px;
            color: #c5d8ef;
            padding: 7px 12px;
            font-size: 13px;
            outline: none;
            width: 200px;
            transition: border-color .2s;
        }
        .search-input:focus { border-color: #4a90d9; }
        .search-input::placeholder { color: #3a5a7a; }

        .table-wrap { overflow-x: auto; border-radius: 10px; border: 1px solid #1e3450; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        thead { background: #0f1e2e; }
        thead th { color: #4a90d9; font-weight: 600; text-align: left; padding: 11px 14px; white-space: nowrap; border-bottom: 1px solid #1e3450; }
        thead th:first-child { width: 40px; text-align: center; }
        tbody tr { border-bottom: 1px solid #172840; transition: background .15s; }
        tbody tr:last-child { border-bottom: none; }
        tbody tr:hover { background: #112030; }
        tbody tr.selected { background: #0e2240; }
        tbody td { padding: 10px 14px; color: #c5d8ef; white-space: nowrap; }
        tbody td:first-child { text-align: center; }

        .pill { display: inline-block; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 99px; }
        .pill-active { background: #0d2d1a; color: #3ecf7a; }

        input[type="checkbox"] { accent-color: #4a90d9; width: 16px; height: 16px; cursor: pointer; }

        .btn-preview {
            border: 1px solid #2a4a6e;
            background: #0f1e2e;
            color: #7ab8e8;
            font-size: 12px;
            font-weight: 600;
            padding: 5px 10px;
            border-radius: 8px;
            cursor: pointer;
            white-space: nowrap;
            transition: background .15s, border-color .15s;
        }
        .btn-preview:hover { background: #132334; border-color: #4a90d9; }
        .btn-preview.active { background: #0e2240; border-color: #4a90d9; color: #fff; }

        /* ── BOTTOM BAR ─────────────────────── */
        .bottom-bar { display: none; margin-top: 24px; background: #0f1e2e; border: 1px solid #1e3450; border-radius: 12px; padding: 16px 20px; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px; }
        .bottom-bar.visible { display: flex; }
        .selected-info { font-size: 14px; color: #7ab8e8; }
        .selected-info strong { color: #fff; }

        /* ── BOTONES ────────────────────────── */
        .btn {
            border: none;
            border-radius: 10px;
            padding: 12px 24px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: background .2s, transform .1s;
        }
        .btn:active { transform: scale(.98); }
        .btn:disabled { opacity: .4; cursor: not-allowed; }
        .btn-blue  { background: #1a6bbf; color: #fff; }
        .btn-blue:hover:not(:disabled)  { background: #1557a0; }
        .btn-green { background: #1a7a40; color: #fff; }
        .btn-green:hover:not(:disabled) { background: #156030; }

        /* ── SPINNER ────────────────────────── */
        .spinner { display: none; width: 18px; height: 18px; border: 2px solid rgba(255,255,255,.3); border-top-color: #fff; border-radius: 50%; animation: spin .7s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }

        /* ── ALERTS ─────────────────────────── */
        .alert { border-radius: 10px; padding: 13px 18px; font-size: 13px; margin-top: 16px; display: none; }
        .alert-error   { background: #2a1010; border: 1px solid #5a2020; color: #f08080; }
        .alert-success { background: #0a2a14; border: 1px solid #1a5a30; color: #5cc87a; }

        /* ── PANEL CORREOS ──────────────────── */
        .email-list { display: flex; flex-direction: column; gap: 8px; max-height: 480px; overflow-y: auto; margin-bottom: 16px; }
        .email-item {
            display: flex;
            align-items: center;
            gap: 10px;
            background: #0f1e2e;
            border: 1px solid #1e3450;
            border-radius: 8px;
            padding: 10px 12px;
            transition: background .15s;
        }
        .email-item:hover { background: #112030; }
        .email-item.selected-email { border-color: #4a90d9; background: #0e2240; }
        .email-item-info { flex: 1; min-width: 0; }
        .email-item-name { font-size: 13px; font-weight: 600; color: #c5d8ef; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .email-item-addr { font-size: 11px; color: #4a90d9; margin-top: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .email-item-addr.no-email { color: #5a3a3a; font-style: italic; }

        /* ── ADJUNTAR DOCUMENTO EXTRA POR TRABAJADOR ───────────── */
        .email-item-extra { flex-shrink: 0; }
        .btn-adjuntar {
            display: flex;
            align-items: center;
            gap: 5px;
            font-size: 11px;
            font-weight: 600;
            color: #7ab8e8;
            background: #132334;
            border: 1px solid #2a4a6e;
            border-radius: 7px;
            padding: 6px 9px;
            cursor: pointer;
            white-space: nowrap;
            transition: background .15s, border-color .15s;
        }
        .btn-adjuntar:hover { background: #16324a; border-color: #4a90d9; }
        .btn-adjuntar.tiene-archivo { border-color: #3ecf7a; color: #3ecf7a; background: #0d2416; }
        .btn-adjuntar .extra-label { max-width: 110px; overflow: hidden; text-overflow: ellipsis; }
        .btn-quitar-extra {
            border: none; background: transparent; color: #f08080; font-size: 13px;
            cursor: pointer; padding: 0 2px; line-height: 1;
        }

        .email-toolbar { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; flex-wrap: wrap; gap: 8px; }
        .badge-email { background: #1a3a5c; color: #7ab8e8; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 99px; }

        .email-status { font-size: 12px; margin-top: 10px; padding: 10px 14px; border-radius: 8px; display: none; }
        .email-status.ok  { background: #0a2a14; border: 1px solid #1a5a30; color: #5cc87a; display: block; }
        .email-status.err { background: #2a1010; border: 1px solid #5a2020; color: #f08080; display: block; }

        /* ── FORM LIQUIDACIÓN MANUAL ────────── */
        .form-label {
            display: block;
            font-size: 12px;
            color: #7ab8e8;
            margin-bottom: 6px;
            font-weight: 600;
        }
        .form-input {
            width: 100%;
            background: #0f1e2e;
            border: 1px solid #2a4a6e;
            border-radius: 8px;
            color: #c5d8ef;
            padding: 10px 12px;
            font-size: 13px;
            outline: none;
            transition: border-color .2s;
        }
        .form-input:focus { border-color: #4a90d9; }
        .form-field { margin-bottom: 14px; }

        /* ── LOADING OVERLAY ────────────────── */
        #loading-overlay { display: none; position: fixed; inset: 0; background: rgba(13,27,42,.85); z-index: 999; flex-direction: column; align-items: center; justify-content: center; gap: 16px; }
        #loading-overlay.active { display: flex; }
        .loading-spinner { width: 48px; height: 48px; border: 4px solid rgba(74,144,217,.3); border-top-color: #4a90d9; border-radius: 50%; animation: spin .8s linear infinite; }
        .loading-text { color: #c5d8ef; font-size: 16px; font-weight: 500; }
        .loading-sub  { color: #4a6a8a; font-size: 13px; margin-top: -10px; }
    </style>
</head>
<body>

<!-- Loading Overlay -->
<div id="loading-overlay">
    <div class="loading-spinner"></div>
    <div class="loading-text" id="loading-text">Generando boletas PDF…</div>
    <div class="loading-sub" id="loading-sub">Esto puede tomar unos segundos.</div>
</div>

<!-- Header -->
<div class="header">
    <img src="logo_leon_primero.jpeg" alt="Logo Agroimportadora Leon S.A.C.">
    <div class="header-text">
        <h1 id="titulo-principal">Agroimportadora Leon S.A.C.</h1>
        <p id="subtitulo-principal">Sistema de generación de boletas de pago</p>
    </div>
    <div class="mode-toggle">
        <button type="button" class="mode-btn active" id="btn-modo-normal" onclick="cambiarModo('normal')">📄 Boletas</button>
        <button type="button" class="mode-btn" id="btn-modo-grat" onclick="cambiarModo('grat')">🎁 Gratificación</button>
        <button type="button" class="mode-btn" id="btn-modo-liq" onclick="cambiarModo('liquidacion')">🧾 Liquidación</button>
    </div>
</div>

<!-- Layout dos columnas -->
<div class="layout">

    <!-- ══ PANEL IZQUIERDO ══════════════════════════════════════ -->
    <div class="card card-left">

        <div class="section-title" id="titulo-carga">📊 Cargar planilla</div>
        <div class="drop-zone" id="drop-zone">
            <input type="file" id="file-input" accept=".xlsx,.xls">
            <div class="drop-icon">📊</div>
            <div class="drop-label" id="drop-label">Arrastra tu planilla Excel aquí</div>
            <div class="drop-sub" id="drop-sub">o haz clic para seleccionar el archivo (.xlsx)</div>
            <div class="file-chosen" id="file-chosen">📎 <span id="file-name-display"></span></div>
        </div>
        <div class="progress-wrap" id="progress-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="progress-fill"></div></div>
            <div class="progress-label" id="progress-label">Leyendo archivo…</div>
        </div>
        <div class="alert alert-error" id="alert-error"></div>

        <div class="divider" id="divider-preview" style="display:none;"></div>

        <div id="preview-section">
            <div class="toolbar">
                <div class="toolbar-left">
                    <div class="section-title" style="margin-bottom:0;" id="titulo-trabajadores">👥 Trabajadores encontrados</div>
                    <span class="badge-count" id="badge-count">0</span>
                </div>
                <input type="text" class="search-input" id="search-input" placeholder="🔍 Buscar por nombre o DNI…">
            </div>

            <div class="table-wrap">
                <table id="workers-table">
                    <thead>
                        <tr>
                            <th><input type="checkbox" id="check-all" title="Seleccionar todos"></th>
                            <th>#</th>
                            <th>DNI</th>
                            <th>Nombre</th>
                            <th>Puesto</th>
                            <th id="th-periodo">Periodo</th>
                            <th id="th-neto">Neto</th>
                            <th>Estado</th>
                            <th id="th-acciones" style="display:none;">Vista previa</th>
                        </tr>
                    </thead>
                    <tbody id="workers-body"></tbody>
                </table>
            </div>

            <div class="bottom-bar" id="bottom-bar">
                <div class="selected-info">
                    <strong id="selected-count">0</strong> de <span id="total-count">0</span> seleccionados
                </div>
                <button class="btn btn-blue" id="btn-generate" onclick="generarBoletas()" disabled>
                    <div class="spinner" id="btn-spinner"></div>
                    <span id="btn-generate-label">📄 Generar Boletas PDF</span>
                </button>
            </div>
        </div>

    </div><!-- /card-left -->

    <!-- ══ PANEL DERECHO — ENVÍO POR CORREO / VISTA PREVIA ══════ -->
    <div class="card card-right" id="card-right">

        <!-- Sub-panel: envío por correo (modo Boletas) -->
        <div id="panel-correos">
            <div class="section-title">✉️ Enviar boletas por correo</div>
            <p style="font-size:12px; color:#4a6a8a; margin-bottom:16px; line-height:1.6;">
                Selecciona los trabajadores a los que quieres enviar su boleta. Solo aparecen los que tienen correo registrado en la planilla (columna BB).
            </p>

            <div class="email-toolbar">
                <div style="display:flex;align-items:center;gap:8px;">
                    <input type="checkbox" id="check-all-email" title="Seleccionar todos con correo">
                    <span style="font-size:13px;color:#7ab8e8;">Seleccionar todos</span>
                </div>
                <span class="badge-email" id="badge-email">0 con correo</span>
            </div>

            <div class="email-list" id="email-list"></div>

            <button class="btn btn-green" id="btn-send-email" onclick="enviarCorreos()" disabled style="width:100%; justify-content:center;">
                <div class="spinner" id="email-spinner"></div>
                <span id="email-btn-text">✉️ Enviar boletas seleccionadas</span>
            </button>

            <div class="email-status" id="email-status"></div>
        </div>

        <!-- Sub-panel: vista previa (modo Gratificación) -->
        <div id="panel-preview" style="display:none;">
            <div class="section-title">👁️ Vista previa</div>
            <p style="font-size:12px; color:#4a6a8a; margin-bottom:16px; line-height:1.6;">
                Haz clic en "Ver" junto a un trabajador para previsualizar aquí su constancia, sin descargarla.
            </p>

            <div id="preview-placeholder" style="font-size:13px; color:#4a6a8a; text-align:center; padding:60px 10px; border:1px dashed #2a4a6e; border-radius:10px;">
                Selecciona "👁️ Ver" en una fila de la tabla.
            </div>

            <div id="preview-loading" style="display:none; font-size:13px; color:#7ab8e8; text-align:center; padding:60px 10px;">
                Generando vista previa…
            </div>

            <iframe id="preview-iframe" style="display:none; width:100%; height:560px; border:1px solid #1e3450; border-radius:10px; background:#fff;"></iframe>

            <div id="preview-name" style="display:none; margin-top:10px; font-size:13px; color:#7ab8e8; font-weight:600;"></div>
        </div>

        <!-- Sub-panel: envío por correo (modo Gratificación) -->
        <div id="panel-correos-grat" style="display:none;">
            <div class="divider"></div>
            <div class="section-title">✉️ Enviar constancias por correo</div>
            <p style="font-size:12px; color:#4a6a8a; margin-bottom:16px; line-height:1.6;">
                Selecciona los trabajadores a los que quieres enviar su constancia de gratificación. Solo aparecen los que tienen correo registrado en la planilla (columna E).
            </p>

            <div class="email-toolbar">
                <div style="display:flex;align-items:center;gap:8px;">
                    <input type="checkbox" id="check-all-email-grat" title="Seleccionar todos con correo">
                    <span style="font-size:13px;color:#7ab8e8;">Seleccionar todos</span>
                </div>
                <span class="badge-email" id="badge-email-grat">0 con correo</span>
            </div>

            <div class="email-list" id="email-list-grat"></div>

            <button class="btn btn-green" id="btn-send-email-grat" onclick="enviarCorreosGratificacion()" disabled style="width:100%; justify-content:center;">
                <div class="spinner" id="email-spinner-grat"></div>
                <span id="email-btn-text-grat">✉️ Enviar constancias seleccionadas</span>
            </button>

            <div class="email-status" id="email-status-grat"></div>
        </div>

        <!-- Sub-panel: envío por correo (modo Liquidación) — manual, con archivo propio -->
        <div id="panel-correo-liquidacion" style="display:none;">
            <div class="section-title">✉️ Enviar documento de liquidación</div>
            <p style="font-size:12px; color:#4a6a8a; margin-bottom:16px; line-height:1.6;">
                Escribe el correo del destinatario y adjunta el archivo que deseas enviar. No se usa el PDF generado automáticamente, sino el que tú cargues aquí.
            </p>

            <div class="form-field">
                <label class="form-label" for="correo-liquidacion">Correo destinatario</label>
                <input type="email" id="correo-liquidacion" class="form-input" placeholder="correo@ejemplo.com">
            </div>

            <div class="form-field">
                <label class="form-label" for="nombre-liquidacion">Nombre del trabajador (opcional)</label>
                <input type="text" id="nombre-liquidacion" class="form-input" placeholder="Nombre del trabajador">
            </div>

            <div class="form-field">
                <label class="form-label">Archivo a enviar</label>
                <div class="drop-zone" id="drop-zone-liquidacion" style="padding:24px 14px;">
                    <input type="file" id="archivo-liquidacion" accept=".pdf,.doc,.docx,.xls,.xlsx">
                    <div class="drop-icon" style="font-size:28px; margin-bottom:8px;">📎</div>
                    <div class="drop-label" style="font-size:13px;">Arrastra el archivo aquí</div>
                    <div class="drop-sub">o haz clic para seleccionarlo</div>
                    <div class="file-chosen" id="file-chosen-liquidacion">📎 <span id="file-name-display-liquidacion"></span></div>
                </div>
            </div>

            <button class="btn btn-green" id="btn-send-liquidacion" onclick="enviarCorreoLiquidacion()" disabled style="width:100%; justify-content:center;">
                <div class="spinner" id="liquidacion-spinner"></div>
                <span id="liquidacion-btn-text">✉️ Enviar documento</span>
            </button>

            <div class="email-status" id="liquidacion-status"></div>
        </div>

    </div><!-- /card-right -->

</div><!-- /layout -->

<!-- Form oculto para generar PDFs (la acción se cambia según el modo: procesar.php / procesar_gratificacion.php / procesar_liquidacion.php) -->
<form id="form-boletas" action="procesar.php" method="POST" enctype="multipart/form-data" style="display:none;">
    <input type="file" name="planilla" id="form-file">
    <input type="hidden" name="dnis_seleccionados" id="dnis-seleccionados">
    <input type="hidden" name="download_token" id="download-token">
</form>

<script>
    let trabajadoresData = [];
    let archivoSeleccionado = null;
    let modoActual = 'normal'; // 'normal' = boletas mensuales | 'grat' = gratificación | 'liquidacion'
    let documentosExtra = {}; // dni -> File (documento extra adjunto para boletas)
    let archivoLiquidacion = null; // archivo libre para el envío manual de liquidación

    // ── Traducción de meses (fórmulas de Excel cacheadas en inglés) ──
    function traducirMesPeriodo(texto) {
        const meses = {
            'JANUARY': 'ENERO', 'FEBRUARY': 'FEBRERO', 'MARCH': 'MARZO',
            'APRIL': 'ABRIL', 'MAY': 'MAYO', 'JUNE': 'JUNIO',
            'JULY': 'JULIO', 'AUGUST': 'AGOSTO', 'SEPTEMBER': 'SEPTIEMBRE',
            'OCTOBER': 'OCTUBRE', 'NOVEMBER': 'NOVIEMBRE', 'DECEMBER': 'DICIEMBRE',
        };
        let t = (texto || '').toString().trim();
        for (const [en, es] of Object.entries(meses)) {
            t = t.replace(new RegExp('\\b' + en + '\\b', 'gi'), es);
        }
        return t;
    }

    // ── Cambio de modo (Boletas ⇄ Gratificación ⇄ Liquidación) ──
    function cambiarModo(nuevoModo) {
        if (nuevoModo === modoActual) return;
        modoActual = nuevoModo;

        document.getElementById('btn-modo-normal').classList.toggle('active', nuevoModo === 'normal');
        document.getElementById('btn-modo-grat').classList.toggle('active', nuevoModo === 'grat');
        document.getElementById('btn-modo-liq').classList.toggle('active', nuevoModo === 'liquidacion');

        // Reset de estado
        trabajadoresData = [];
        archivoSeleccionado = null;
        documentosExtra = {};
        document.getElementById('file-input').value = '';
        document.getElementById('file-chosen').style.display = 'none';
        document.getElementById('preview-section').style.display = 'none';
        document.getElementById('divider-preview').style.display = 'none';
        document.getElementById('card-right').style.display = 'none';
        document.getElementById('panel-correos-grat').style.display = 'none';
        document.getElementById('email-list-grat').innerHTML = '';
        document.getElementById('email-status-grat').className = 'email-status';
        document.getElementById('email-status-grat').textContent = '';
        limpiarPreviaPDF();
        ocultarError();

        // Reset del panel manual de liquidación
        document.getElementById('panel-correo-liquidacion').style.display = 'none';
        document.getElementById('correo-liquidacion').value = '';
        document.getElementById('nombre-liquidacion').value = '';
        document.getElementById('archivo-liquidacion').value = '';
        document.getElementById('file-chosen-liquidacion').style.display = 'none';
        document.getElementById('liquidacion-status').className = 'email-status';
        document.getElementById('liquidacion-status').textContent = '';
        document.getElementById('btn-send-liquidacion').disabled = true;
        archivoLiquidacion = null;

        document.getElementById('th-acciones').style.display = (nuevoModo === 'grat') ? 'table-cell' : 'none';

        if (nuevoModo === 'grat') {
            document.getElementById('titulo-principal').textContent = 'Agroimportadora Leon S.A.C.';
            document.getElementById('subtitulo-principal').textContent = 'Sistema de generación de constancias de gratificación';
            document.getElementById('titulo-carga').textContent = '🎁 Cargar planilla de gratificación';
            document.getElementById('drop-label').textContent = 'Arrastra tu planilla de gratificación aquí';
            document.getElementById('drop-sub').textContent = 'o haz clic para seleccionar el archivo (.xlsx)';
            document.getElementById('titulo-trabajadores').textContent = '👥 Trabajadores encontrados';
            document.getElementById('th-periodo').textContent = 'Periodo (Al)';
            document.getElementById('th-neto').textContent = 'Importe';
            document.getElementById('btn-generate-label').textContent = '🎁 Generar Constancias PDF';
        } else if (nuevoModo === 'liquidacion') {
            document.getElementById('titulo-principal').textContent = 'Agroimportadora Leon S.A.C.';
            document.getElementById('subtitulo-principal').textContent = 'Sistema de generación de liquidaciones de beneficios sociales';
            document.getElementById('titulo-carga').textContent = '🧾 Cargar planilla de liquidación';
            document.getElementById('drop-label').textContent = 'Arrastra tu planilla de liquidación aquí';
            document.getElementById('drop-sub').textContent = 'o haz clic para seleccionar el archivo (.xlsx)';
            document.getElementById('titulo-trabajadores').textContent = '👥 Trabajadores encontrados';
            document.getElementById('th-periodo').textContent = 'Fecha de cese';
            document.getElementById('th-neto').textContent = 'Neto liquidación';
            document.getElementById('btn-generate-label').textContent = '🧾 Generar Liquidaciones PDF';
        } else {
            document.getElementById('titulo-principal').textContent = 'Agroimportadora Leon S.A.C.';
            document.getElementById('subtitulo-principal').textContent = 'Sistema de generación de boletas de pago';
            document.getElementById('titulo-carga').textContent = '📊 Cargar planilla';
            document.getElementById('drop-label').textContent = 'Arrastra tu planilla Excel aquí';
            document.getElementById('drop-sub').textContent = 'o haz clic para seleccionar el archivo (.xlsx)';
            document.getElementById('titulo-trabajadores').textContent = '👥 Trabajadores encontrados';
            document.getElementById('th-periodo').textContent = 'Periodo';
            document.getElementById('th-neto').textContent = 'Neto';
            document.getElementById('btn-generate-label').textContent = '📄 Generar Boletas PDF';
        }
    }

    // ── Drop Zone (planilla principal) ─────────────────────────
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', e => {
        e.preventDefault(); dropZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) procesarArchivo(file);
    });
    fileInput.addEventListener('change', () => { if (fileInput.files[0]) procesarArchivo(fileInput.files[0]); });

    function procesarArchivo(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!['xlsx','xls'].includes(ext)) { mostrarError('El archivo debe ser .xlsx o .xls'); return; }
        archivoSeleccionado = file;
        documentosExtra = {};
        document.getElementById('file-name-display').textContent = file.name;
        document.getElementById('file-chosen').style.display = 'block';
        ocultarError();
        leerExcelCliente(file);
    }

    // ── Leer Excel con SheetJS ──────────────────────────────────
    function leerExcelCliente(file) {
        const progressWrap  = document.getElementById('progress-wrap');
        const progressFill  = document.getElementById('progress-fill');
        const progressLabel = document.getElementById('progress-label');

        progressWrap.style.display = 'block';
        progressFill.style.width = '20%';
        progressLabel.textContent = 'Cargando archivo…';

        const cargar = () => {
            const reader = new FileReader();
            reader.onprogress = e => {
                if (e.lengthComputable) {
                    progressFill.style.width = Math.round((e.loaded / e.total) * 50 + 20) + '%';
                }
            };
            reader.onload = e => {
                progressFill.style.width = '80%';
                progressLabel.textContent = 'Procesando datos…';
                try {
                    const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
                    let sheetName;
                    if (modoActual === 'grat') {
                        sheetName = wb.SheetNames.find(n => /gratific/i.test(n)) || wb.SheetNames[0];
                    } else if (modoActual === 'liquidacion') {
                        sheetName = wb.SheetNames.find(n => /liquidaci/i.test(n)) || wb.SheetNames[0];
                    } else {
                        sheetName = wb.SheetNames.includes('Planilla Mensual') ? 'Planilla Mensual' : wb.SheetNames[0];
                    }
                    const ws = wb.Sheets[sheetName];
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

                    trabajadoresData = [];

                    if (modoActual === 'grat') {
                        // Fila 1 = fecha de corte global, fila 2 = encabezados, datos desde fila 3 (índice 2)
                        for (let i = 2; i < rows.length; i++) {
                            const row = rows[i];
                            const dni = (row[0] || '').toString().trim();
                            if (!dni) continue;

                            const correo = (row[4] || '').toString().trim();
                            const importe = parseFloat(row[17]) || 0; // columna R
                            let alTexto = row[8];
                            if (alTexto instanceof Date) {
                                alTexto = alTexto.toLocaleDateString('es-PE');
                            }

                            trabajadoresData.push({
                                dni,
                                nombre:  (row[1] || '').toString().trim(), // columna B
                                puesto:  (row[3] || '').toString().trim(), // columna D
                                periodo: traducirMesPeriodo((alTexto || '').toString().trim()),
                                neto: importe,
                                correo
                            });
                        }
                    } else if (modoActual === 'liquidacion') {
                        // Encabezados en fila 3 (índice 2), datos desde fila 4 (índice 3)
                        for (let i = 3; i < rows.length; i++) {
                            const row = rows[i];
                            const dni = (row[0] || '').toString().trim();
                            if (!dni) continue;

                            let fCese = row[7]; // columna H
                            if (fCese instanceof Date) {
                                fCese = fCese.toLocaleDateString('es-PE');
                            }
                            const neto = parseFloat(row[37]) || 0; // columna AL

                            trabajadoresData.push({
                                dni,
                                nombre:  (row[1] || '').toString().trim(), // columna B
                                puesto:  (row[3] || '').toString().trim(), // columna D
                                periodo: (fCese  || '').toString().trim(), // ya es fecha formateada, no necesita traducción
                                neto,
                                correo: ''
                            });
                        }
                    } else {
                        for (let i = 2; i < rows.length; i++) {
                            const row  = rows[i];
                            const tipo = (row[1] || '').toString().trim().toUpperCase();
                            const dni  = (row[2] || '').toString().trim();
                            if (!dni || tipo !== 'PLANILLA') continue;

                            // Columna BB = índice 53 (A=0, B=1, … BB=53)
                            const correo = (row[53] || '').toString().trim();
                            const neto   = parseFloat(row[48]) || 0;

                            trabajadoresData.push({
                                dni,
                                nombre:  (row[3]  || '').toString().trim(),
                                puesto:  (row[5]  || '').toString().trim(),
                                periodo: traducirMesPeriodo((row[14] || '').toString().trim()),
                                neto,
                                correo
                            });
                        }
                    }

                    progressFill.style.width = '100%';
                    progressLabel.textContent = `${trabajadoresData.length} trabajadores encontrados.`;
                    setTimeout(() => { progressWrap.style.display = 'none'; mostrarTabla(); }, 600);
                } catch(err) {
                    mostrarError('Error al leer el archivo: ' + err.message);
                    progressWrap.style.display = 'none';
                }
            };
            reader.readAsBinaryString(file);
        };

        if (!window.XLSX) {
            const s = document.createElement('script');
            s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
            s.onload = cargar;
            document.head.appendChild(s);
        } else { cargar(); }
    }

    // ── Tabla principal ─────────────────────────────────────────
    function mostrarTabla() {
        document.getElementById('divider-preview').style.display = 'block';
        document.getElementById('preview-section').style.display = 'block';
        if (trabajadoresData.length === 0) {
            let msg;
            if (modoActual === 'grat') {
                msg = 'No se encontraron trabajadores en la planilla de gratificación.';
            } else if (modoActual === 'liquidacion') {
                msg = 'No se encontraron trabajadores en la planilla de liquidación.';
            } else {
                msg = 'No se encontraron trabajadores con tipo PLANILLA.';
            }
            mostrarError(msg);
            return;
        }
        document.getElementById('badge-count').textContent = trabajadoresData.length;
        document.getElementById('total-count').textContent = trabajadoresData.length;
        renderTabla(trabajadoresData);
        actualizarBottomBar();

        if (modoActual === 'normal') {
            document.getElementById('card-right').style.display = 'block';
            document.getElementById('panel-correos').style.display = 'block';
            document.getElementById('panel-preview').style.display = 'none';
            document.getElementById('panel-correos-grat').style.display = 'none';
            document.getElementById('panel-correo-liquidacion').style.display = 'none';
            mostrarPanelCorreos();
        } else if (modoActual === 'grat') {
            document.getElementById('card-right').style.display = 'block';
            document.getElementById('panel-correos').style.display = 'none';
            document.getElementById('panel-preview').style.display = 'block';
            document.getElementById('panel-correos-grat').style.display = 'block';
            document.getElementById('panel-correo-liquidacion').style.display = 'none';
            limpiarPreviaPDF();
            mostrarPanelCorreosGrat();
        } else {
            // Liquidación: panel de envío manual de correo (correo escrito a mano + archivo propio)
            document.getElementById('card-right').style.display = 'block';
            document.getElementById('panel-correos').style.display = 'none';
            document.getElementById('panel-preview').style.display = 'none';
            document.getElementById('panel-correos-grat').style.display = 'none';
            document.getElementById('panel-correo-liquidacion').style.display = 'block';
        }
    }

    function renderTabla(datos) {
        const tbody = document.getElementById('workers-body');
        tbody.innerHTML = '';
        datos.forEach((t, idx) => {
            const tr = document.createElement('tr');
            tr.dataset.dni = t.dni;
            tr.innerHTML = `
                <td><input type="checkbox" class="worker-cb" value="${esc(t.dni)}" ></td>
                <td style="color:#4a6a8a;">${idx + 1}</td>
                <td>${esc(t.dni)}</td>
                <td style="font-weight:500;">${esc(t.nombre)}</td>
                <td style="color:#7ab8e8;">${esc(t.puesto)}</td>
                <td>${esc(t.periodo)}</td>
                <td style="font-weight:600;color:#5cc87a;">S/ ${t.neto.toFixed(2)}</td>
                <td><span class="pill pill-active">Activo</span></td>
                <td style="display:${modoActual === 'grat' ? 'table-cell' : 'none'};">
                    <button type="button" class="btn-preview" data-dni="${esc(t.dni)}" onclick="verPrevia('${esc(t.dni)}', this)">👁️ Ver</button>
                </td>
            `;
            tr.querySelector('.worker-cb').addEventListener('change', function() {
                tr.classList.toggle('selected', this.checked);
                onCheckChange();
            });
            tbody.appendChild(tr);
        });
    }

    function esc(str) {
        return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }

    document.getElementById('check-all').addEventListener('change', function() {
        document.querySelectorAll('.worker-cb').forEach(cb => {
            cb.checked = this.checked;
            cb.closest('tr').classList.toggle('selected', this.checked);
        });
        onCheckChange();
    });

    function onCheckChange() {
        const total   = document.querySelectorAll('.worker-cb').length;
        const checked = document.querySelectorAll('.worker-cb:checked').length;
        document.getElementById('check-all').checked       = (checked === total && total > 0);
        document.getElementById('check-all').indeterminate = (checked > 0 && checked < total);
        document.getElementById('selected-count').textContent = checked;
        actualizarBottomBar();
    }

    function actualizarBottomBar() {
        const checked = document.querySelectorAll('.worker-cb:checked').length;
        document.getElementById('bottom-bar').classList.toggle('visible', trabajadoresData.length > 0);
        document.getElementById('btn-generate').disabled = (checked === 0);
    }

    document.getElementById('search-input').addEventListener('input', function() {
        const q = this.value.toLowerCase();
        const filtrados = trabajadoresData.filter(t =>
            t.nombre.toLowerCase().includes(q) || t.dni.includes(q)
        );
        renderTabla(filtrados);
        document.getElementById('badge-count').textContent = filtrados.length;
    });

    // ── Panel de correos (Boletas) ──────────────────────────────
    function mostrarPanelCorreos() {
        const panel      = document.getElementById('card-right');
        const emailList  = document.getElementById('email-list');
        const conCorreo  = trabajadoresData.filter(t => t.correo && t.correo.includes('@'));
        
        document.getElementById('badge-email').textContent = `${conCorreo.length} con correo`;
        panel.style.display = 'block';
        emailList.innerHTML = '';

        trabajadoresData.forEach(t => {
            const tieneCorreo = t.correo && t.correo.includes('@');
            const div = document.createElement('div');
            div.className = 'email-item';
            div.dataset.dni = t.dni;
            const archivoExtra = documentosExtra[t.dni];
            div.innerHTML = `
                <input type="checkbox" class="email-cb" value="${esc(t.dni)}" 
                    data-correo="${esc(t.correo)}"
                    ${!tieneCorreo ? 'disabled' : ''}>
                <div class="email-item-info">
                    <div class="email-item-name">${esc(t.nombre)}</div>
                    <div class="email-item-addr ${!tieneCorreo ? 'no-email' : ''}">
                        ${tieneCorreo ? esc(t.correo) : 'Sin correo registrado'}
                    </div>
                </div>
                <div class="email-item-extra">
                    <label class="btn-adjuntar ${archivoExtra ? 'tiene-archivo' : ''}" for="extra-${esc(t.dni)}" title="Adjuntar un documento extra (Word o PDF) que se enviará junto a la boleta">
                        📎 <span class="extra-label">${archivoExtra ? esc(archivoExtra.name) : 'Adjuntar doc.'}</span>
                        ${archivoExtra ? `<span class="btn-quitar-extra" data-dni="${esc(t.dni)}" title="Quitar documento">✕</span>` : ''}
                    </label>
                    <input type="file" id="extra-${esc(t.dni)}" class="extra-file-input" data-dni="${esc(t.dni)}" accept=".pdf,.doc,.docx" style="display:none;">
                </div>
            `;
            if (tieneCorreo) {
                div.querySelector('.email-cb').addEventListener('change', function() {
                    div.classList.toggle('selected-email', this.checked);
                    actualizarBtnEmail();
                });
            }
            div.querySelector('.extra-file-input').addEventListener('change', function(e) {
                const archivo = e.target.files[0];
                if (!archivo) return;
                const extOk = /\.(pdf|doc|docx)$/i.test(archivo.name);
                if (!extOk) { mostrarError('El documento extra debe ser PDF o Word (.pdf, .doc, .docx).'); e.target.value = ''; return; }
                documentosExtra[t.dni] = archivo;
                actualizarBotonAdjuntar(div, t.dni);
            });
            const btnQuitarInicial = div.querySelector('.btn-quitar-extra');
            if (btnQuitarInicial) {
                btnQuitarInicial.addEventListener('click', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    delete documentosExtra[t.dni];
                    div.querySelector('.extra-file-input').value = '';
                    actualizarBotonAdjuntar(div, t.dni);
                });
            }
            emailList.appendChild(div);
        });

        actualizarBtnEmail();
    }

    // Actualiza en el DOM solo el botón "Adjuntar" de una fila, sin re-renderizar
    // todo el panel (para no perder la selección de checkboxes de las demás filas).
    function actualizarBotonAdjuntar(div, dni) {
        const archivoExtra = documentosExtra[dni];
        const label = div.querySelector('.btn-adjuntar');
        label.classList.toggle('tiene-archivo', !!archivoExtra);
        label.innerHTML = `
            📎 <span class="extra-label">${archivoExtra ? esc(archivoExtra.name) : 'Adjuntar doc.'}</span>
            ${archivoExtra ? `<span class="btn-quitar-extra" data-dni="${esc(dni)}" title="Quitar documento">✕</span>` : ''}
        `;
        const btnQuitar = label.querySelector('.btn-quitar-extra');
        if (btnQuitar) {
            btnQuitar.addEventListener('click', function(e) {
                e.preventDefault();
                e.stopPropagation();
                delete documentosExtra[dni];
                div.querySelector('.extra-file-input').value = '';
                actualizarBotonAdjuntar(div, dni);
            });
        }
    }

    document.getElementById('check-all-email').addEventListener('change', function() {
        document.querySelectorAll('.email-cb:not(:disabled)').forEach(cb => {
            cb.checked = this.checked;
            cb.closest('.email-item').classList.toggle('selected-email', this.checked);
        });
        actualizarBtnEmail();
    });

    function actualizarBtnEmail() {
        const checked = document.querySelectorAll('.email-cb:checked').length;
        document.getElementById('btn-send-email').disabled = (checked === 0);
        const total = document.querySelectorAll('.email-cb:not(:disabled)').length;
        document.getElementById('check-all-email').checked       = (checked === total && total > 0);
        document.getElementById('check-all-email').indeterminate = (checked > 0 && checked < total);
    }

    // ── Panel de correos (Gratificación) ──────────────────────────
    function mostrarPanelCorreosGrat() {
        const emailList = document.getElementById('email-list-grat');
        const conCorreo = trabajadoresData.filter(t => t.correo && t.correo.includes('@'));

        document.getElementById('badge-email-grat').textContent = `${conCorreo.length} con correo`;
        emailList.innerHTML = '';

        trabajadoresData.forEach(t => {
            const tieneCorreo = t.correo && t.correo.includes('@');
            const div = document.createElement('div');
            div.className = 'email-item';
            div.dataset.dni = t.dni;
            div.innerHTML = `
                <input type="checkbox" class="email-cb-grat" value="${esc(t.dni)}" 
                    data-correo="${esc(t.correo)}"
                    ${!tieneCorreo ? 'disabled' : ''}>
                <div class="email-item-info">
                    <div class="email-item-name">${esc(t.nombre)}</div>
                    <div class="email-item-addr ${!tieneCorreo ? 'no-email' : ''}">
                        ${tieneCorreo ? esc(t.correo) : 'Sin correo registrado'}
                    </div>
                </div>
            `;
            if (tieneCorreo) {
                div.querySelector('.email-cb-grat').addEventListener('change', function() {
                    div.classList.toggle('selected-email', this.checked);
                    actualizarBtnEmailGrat();
                });
            }
            emailList.appendChild(div);
        });

        actualizarBtnEmailGrat();
    }

    document.getElementById('check-all-email-grat').addEventListener('change', function() {
        document.querySelectorAll('.email-cb-grat:not(:disabled)').forEach(cb => {
            cb.checked = this.checked;
            cb.closest('.email-item').classList.toggle('selected-email', this.checked);
        });
        actualizarBtnEmailGrat();
    });

    function actualizarBtnEmailGrat() {
        const checked = document.querySelectorAll('.email-cb-grat:checked').length;
        document.getElementById('btn-send-email-grat').disabled = (checked === 0);
        const total = document.querySelectorAll('.email-cb-grat:not(:disabled)').length;
        document.getElementById('check-all-email-grat').checked       = (checked === total && total > 0);
        document.getElementById('check-all-email-grat').indeterminate = (checked > 0 && checked < total);
    }

    // ── Generar PDFs ────────────────────────────────────────────
    function generarBoletas() {
        const seleccionados = [...document.querySelectorAll('.worker-cb:checked')].map(cb => cb.value);
        if (seleccionados.length === 0) return;

        document.getElementById('dnis-seleccionados').value = seleccionados.join(',');
        const token = Date.now();
        document.cookie = `descarga_token=${token}; path=/`;

        document.getElementById('download-token').value = token;

        const dt = new DataTransfer();
        dt.items.add(archivoSeleccionado);
        document.getElementById('form-file').files = dt.files;

        const form = document.getElementById('form-boletas');
        if (modoActual === 'grat') {
            form.action = 'procesar_gratificacion.php';
        } else if (modoActual === 'liquidacion') {
            form.action = 'procesar_liquidacion.php';
        } else {
            form.action = 'procesar.php';
        }

        let textoCarga;
        if (modoActual === 'grat') {
            textoCarga = 'Generando constancias PDF…';
        } else if (modoActual === 'liquidacion') {
            textoCarga = 'Generando liquidaciones PDF…';
        } else {
            textoCarga = 'Generando boletas PDF…';
        }
        document.getElementById('loading-text').textContent = textoCarga;
        document.getElementById('loading-sub').textContent  = 'Esto puede tomar unos segundos.';
        document.getElementById('loading-overlay').classList.add('active');
        form.submit();

        const interval = setInterval(() => {
            if (document.cookie.includes(`descarga_lista_${token}`)) {
                clearInterval(interval);
                document.getElementById('loading-overlay').classList.remove('active');
                document.cookie = `descarga_token=${token}; max-age=0; path=/`;
                document.cookie = `descarga_lista_${token}=1; max-age=0; path=/`;
            }
        }, 500);
        setTimeout(() => { clearInterval(interval); document.getElementById('loading-overlay').classList.remove('active'); }, 30000);
    }

    // ── Vista previa (Gratificación) ──────────────────────────────
    let btnPreviaActivo = null;

    function limpiarPreviaPDF() {
        const iframe = document.getElementById('preview-iframe');
        if (iframe && iframe.dataset.blobUrl) {
            URL.revokeObjectURL(iframe.dataset.blobUrl);
            iframe.removeAttribute('src');
            iframe.dataset.blobUrl = '';
        }
        if (iframe) iframe.style.display = 'none';

        const nameEl = document.getElementById('preview-name');
        if (nameEl) { nameEl.style.display = 'none'; nameEl.textContent = ''; }

        const placeholder = document.getElementById('preview-placeholder');
        if (placeholder) placeholder.style.display = 'block';

        const loading = document.getElementById('preview-loading');
        if (loading) loading.style.display = 'none';

        if (btnPreviaActivo) { btnPreviaActivo.classList.remove('active'); btnPreviaActivo = null; }
    }

    async function verPrevia(dni, btn) {
        if (!archivoSeleccionado) {
            mostrarError('Primero carga la planilla de gratificación.');
            return;
        }

        if (btnPreviaActivo) btnPreviaActivo.classList.remove('active');
        btnPreviaActivo = btn;
        btn.classList.add('active');

        document.getElementById('card-right').style.display = 'block';
        document.getElementById('panel-preview').style.display = 'block';
        document.getElementById('panel-correos').style.display = 'none';

        document.getElementById('preview-placeholder').style.display = 'none';
        document.getElementById('preview-iframe').style.display = 'none';
        document.getElementById('preview-name').style.display = 'none';
        document.getElementById('preview-loading').style.display = 'block';

        try {
            const formData = new FormData();
            formData.append('planilla', archivoSeleccionado);
            formData.append('dni', dni);

            const resp = await fetch('vista_previa_gratificacion.php', { method: 'POST', body: formData });
            if (!resp.ok) {
                const txt = await resp.text();
                throw new Error(txt || `Error ${resp.status}`);
            }

            const blob = await resp.blob();
            const url  = URL.createObjectURL(blob);

            const iframe = document.getElementById('preview-iframe');
            if (iframe.dataset.blobUrl) URL.revokeObjectURL(iframe.dataset.blobUrl);
            iframe.src = url;
            iframe.dataset.blobUrl = url;
            iframe.style.display = 'block';

            const t = trabajadoresData.find(x => x.dni === dni);
            const nameEl = document.getElementById('preview-name');
            nameEl.textContent = t ? `📄 ${t.nombre}` : '';
            nameEl.style.display = 'block';
        } catch (err) {
            mostrarError('No se pudo generar la vista previa: ' + err.message);
            document.getElementById('preview-placeholder').style.display = 'block';
        } finally {
            document.getElementById('preview-loading').style.display = 'none';
        }
    }

    // ── Enviar correos (Boletas) ────────────────────────────────
    async function enviarCorreos() {
        const seleccionados = [...document.querySelectorAll('.email-cb:checked')].map(cb => ({
            dni:    cb.value,
            correo: cb.dataset.correo
        }));
        if (seleccionados.length === 0) return;

        const btn = document.getElementById('btn-send-email');
        const spinner = document.getElementById('email-spinner');
        const txtBtn  = document.getElementById('email-btn-text');
        const status  = document.getElementById('email-status');

        btn.disabled    = true;
        spinner.style.display = 'block';
        txtBtn.textContent    = 'Enviando…';
        status.className      = 'email-status';

        document.getElementById('loading-text').textContent = 'Enviando correos…';
        document.getElementById('loading-sub').textContent  = 'Generando y enviando boletas por email.';
        document.getElementById('loading-overlay').classList.add('active');

        try {
            const formData = new FormData();
            formData.append('planilla', archivoSeleccionado);
            formData.append('destinatarios', JSON.stringify(seleccionados));

            // Adjuntar el documento extra (si existe) de cada trabajador seleccionado.
            // Se manda con la clave "extra_<dni>" para que el backend sepa a quién pertenece.
            seleccionados.forEach(s => {
                const archivoExtra = documentosExtra[s.dni];
                if (archivoExtra) {
                    formData.append('extra_' + s.dni, archivoExtra, archivoExtra.name);
                }
            });

            const resp = await fetch('enviar_correos.php', { method: 'POST', body: formData });
            const data = await resp.json();

            document.getElementById('loading-overlay').classList.remove('active');

            if (data.ok) {
                status.textContent  = `✅ ${data.mensaje}`;
                status.className    = 'email-status ok';
            } else {
                status.textContent  = `⚠️ ${data.mensaje}`;
                status.className    = 'email-status err';
            }
        } catch(err) {
            document.getElementById('loading-overlay').classList.remove('active');
            status.textContent = '⚠️ Error de conexión: ' + err.message;
            status.className   = 'email-status err';
        }

        spinner.style.display = 'none';
        txtBtn.textContent    = '✉️ Enviar boletas seleccionadas';
        btn.disabled = false;
    }

    // ── Enviar correos (Gratificación) ────────────────────────────
    async function enviarCorreosGratificacion() {
        const seleccionados = [...document.querySelectorAll('.email-cb-grat:checked')].map(cb => ({
            dni:    cb.value,
            correo: cb.dataset.correo
        }));
        if (seleccionados.length === 0) return;

        const btn = document.getElementById('btn-send-email-grat');
        const spinner = document.getElementById('email-spinner-grat');
        const txtBtn  = document.getElementById('email-btn-text-grat');
        const status  = document.getElementById('email-status-grat');

        btn.disabled    = true;
        spinner.style.display = 'block';
        txtBtn.textContent    = 'Enviando…';
        status.className      = 'email-status';

        document.getElementById('loading-text').textContent = 'Enviando correos…';
        document.getElementById('loading-sub').textContent  = 'Generando y enviando constancias por email.';
        document.getElementById('loading-overlay').classList.add('active');

        try {
            const formData = new FormData();
            formData.append('planilla', archivoSeleccionado);
            formData.append('destinatarios', JSON.stringify(seleccionados));

            const resp = await fetch('enviar_correos_gratificacion.php', { method: 'POST', body: formData });
            const data = await resp.json();

            document.getElementById('loading-overlay').classList.remove('active');

            if (data.ok) {
                status.textContent  = `✅ ${data.mensaje}`;
                status.className    = 'email-status ok';
            } else {
                status.textContent  = `⚠️ ${data.mensaje}`;
                status.className    = 'email-status err';
            }
        } catch(err) {
            document.getElementById('loading-overlay').classList.remove('active');
            status.textContent = '⚠️ Error de conexión: ' + err.message;
            status.className   = 'email-status err';
        }

        spinner.style.display = 'none';
        txtBtn.textContent    = '✉️ Enviar constancias seleccionadas';
        btn.disabled = false;
    }

    // ── Envío manual de correo (Liquidación) — correo escrito a mano + archivo propio ──
    const dropZoneLiq  = document.getElementById('drop-zone-liquidacion');
    const fileInputLiq = document.getElementById('archivo-liquidacion');

    dropZoneLiq.addEventListener('dragover', e => { e.preventDefault(); dropZoneLiq.classList.add('dragover'); });
    dropZoneLiq.addEventListener('dragleave', () => dropZoneLiq.classList.remove('dragover'));
    dropZoneLiq.addEventListener('drop', e => {
        e.preventDefault(); dropZoneLiq.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) seleccionarArchivoLiquidacion(file);
    });
    fileInputLiq.addEventListener('change', () => {
        if (fileInputLiq.files[0]) seleccionarArchivoLiquidacion(fileInputLiq.files[0]);
    });

    function seleccionarArchivoLiquidacion(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!['pdf','doc','docx','xls','xlsx'].includes(ext)) {
            mostrarError('El archivo debe ser PDF, Word o Excel.');
            return;
        }
        archivoLiquidacion = file;
        document.getElementById('file-name-display-liquidacion').textContent = file.name;
        document.getElementById('file-chosen-liquidacion').style.display = 'block';
        actualizarBtnLiquidacion();
    }

    document.getElementById('correo-liquidacion').addEventListener('input', actualizarBtnLiquidacion);

    function actualizarBtnLiquidacion() {
        const correo = document.getElementById('correo-liquidacion').value.trim();
        const valido = correo.includes('@') && archivoLiquidacion;
        document.getElementById('btn-send-liquidacion').disabled = !valido;
    }

    async function enviarCorreoLiquidacion() {
        const correo = document.getElementById('correo-liquidacion').value.trim();
        const nombre = document.getElementById('nombre-liquidacion').value.trim();
        if (!correo || !archivoLiquidacion) return;

        const btn = document.getElementById('btn-send-liquidacion');
        const spinner = document.getElementById('liquidacion-spinner');
        const txtBtn = document.getElementById('liquidacion-btn-text');
        const status = document.getElementById('liquidacion-status');

        btn.disabled = true;
        spinner.style.display = 'block';
        txtBtn.textContent = 'Enviando…';
        status.className = 'email-status';

        document.getElementById('loading-text').textContent = 'Enviando correo…';
        document.getElementById('loading-sub').textContent = 'Enviando el documento de liquidación.';
        document.getElementById('loading-overlay').classList.add('active');

        try {
            const formData = new FormData();
            formData.append('correo', correo);
            formData.append('nombre', nombre);
            formData.append('archivo', archivoLiquidacion, archivoLiquidacion.name);

            const resp = await fetch('enviar_correo_liquidacion.php', { method: 'POST', body: formData });
            const data = await resp.json();

            document.getElementById('loading-overlay').classList.remove('active');

            if (data.ok) {
                status.textContent = `✅ ${data.mensaje}`;
                status.className = 'email-status ok';
            } else {
                status.textContent = `⚠️ ${data.mensaje}`;
                status.className = 'email-status err';
            }
        } catch (err) {
            document.getElementById('loading-overlay').classList.remove('active');
            status.textContent = '⚠️ Error de conexión: ' + err.message;
            status.className = 'email-status err';
        }

        spinner.style.display = 'none';
        txtBtn.textContent = '✉️ Enviar documento';
        btn.disabled = false;
    }

    // ── Utilidades ──────────────────────────────────────────────
    function mostrarError(msg) {
        const el = document.getElementById('alert-error');
        el.textContent = '⚠️ ' + msg;
        el.style.display = 'block';
    }
    function ocultarError() { document.getElementById('alert-error').style.display = 'none'; }
</script>
</body>
</html>