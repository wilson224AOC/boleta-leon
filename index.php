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
        .header-text h1 { font-size: 22px; font-weight: 700; color: #ffffff; line-height: 1.2; }
        .header-text p  { font-size: 13px; color: #7a9cbf; margin-top: 3px; }

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

        .email-toolbar { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; flex-wrap: wrap; gap: 8px; }
        .badge-email { background: #1a3a5c; color: #7ab8e8; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 99px; }

        .email-status { font-size: 12px; margin-top: 10px; padding: 10px 14px; border-radius: 8px; display: none; }
        .email-status.ok  { background: #0a2a14; border: 1px solid #1a5a30; color: #5cc87a; display: block; }
        .email-status.err { background: #2a1010; border: 1px solid #5a2020; color: #f08080; display: block; }

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
        <h1>Agroimportadora Leon S.A.C.</h1>
        <p>Sistema de generación de boletas de pago</p>
    </div>
</div>

<!-- Layout dos columnas -->
<div class="layout">

    <!-- ══ PANEL IZQUIERDO ══════════════════════════════════════ -->
    <div class="card card-left">

        <div class="section-title">📊 Cargar planilla</div>
        <div class="drop-zone" id="drop-zone">
            <input type="file" id="file-input" accept=".xlsx,.xls">
            <div class="drop-icon">📊</div>
            <div class="drop-label">Arrastra tu planilla Excel aquí</div>
            <div class="drop-sub">o haz clic para seleccionar el archivo (.xlsx)</div>
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
                    <div class="section-title" style="margin-bottom:0;">👥 Trabajadores encontrados</div>
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
                            <th>Periodo</th>
                            <th>Neto</th>
                            <th>Estado</th>
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
                    <span>📄 Generar Boletas PDF</span>
                </button>
            </div>
        </div>

    </div><!-- /card-left -->

    <!-- ══ PANEL DERECHO — ENVÍO POR CORREO ═════════════════════ -->
    <div class="card card-right" id="card-right">

        <div class="section-title">✉️ Enviar boletas por correo</div>
        <p style="font-size:12px; color:#4a6a8a; margin-bottom:16px; line-height:1.6;">
            Selecciona los trabajadores a los que quieres enviar su boleta. Solo aparecen los que tienen correo registrado en la planilla (columna AY).
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

    </div><!-- /card-right -->

</div><!-- /layout -->

<!-- Form oculto para generar PDFs -->
<form id="form-boletas" action="procesar.php" method="POST" enctype="multipart/form-data" style="display:none;">
    <input type="file" name="planilla" id="form-file">
    <input type="hidden" name="dnis_seleccionados" id="dnis-seleccionados">
    <input type="hidden" name="download_token" id="download-token">
</form>

<script>
    let trabajadoresData = [];
    let archivoSeleccionado = null;

    // ── Drop Zone ───────────────────────────────────────────────
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
                    const sheetName = wb.SheetNames.includes('Planilla Mensual') ? 'Planilla Mensual' : wb.SheetNames[0];
                    const ws = wb.Sheets[sheetName];
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

                    trabajadoresData = [];
                    for (let i = 2; i < rows.length; i++) {
                        const row  = rows[i];
                        const tipo = (row[1] || '').toString().trim().toUpperCase();
                        const dni  = (row[2] || '').toString().trim();
                        if (!dni || tipo !== 'PLANILLA') continue;

                        // Columna AY = índice 50 (A=0, B=1, … AY=50)
                        const correo = (row[51] || '').toString().trim();
                        const neto   = parseFloat(row[47]) || 0;

                        trabajadoresData.push({
                            dni,
                            nombre:  (row[3]  || '').toString().trim(),
                            puesto:  (row[5]  || '').toString().trim(),
                            periodo: (row[14] || '').toString().trim(),
                            neto,
                            correo
                        });
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
            mostrarError('No se encontraron trabajadores con tipo PLANILLA.');
            return;
        }
        document.getElementById('badge-count').textContent = trabajadoresData.length;
        document.getElementById('total-count').textContent = trabajadoresData.length;
        renderTabla(trabajadoresData);
        actualizarBottomBar();
        mostrarPanelCorreos();
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

    // ── Panel de correos ────────────────────────────────────────
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
            `;
            if (tieneCorreo) {
                div.querySelector('.email-cb').addEventListener('change', function() {
                    div.classList.toggle('selected-email', this.checked);
                    actualizarBtnEmail();
                });
            }
            emailList.appendChild(div);
        });

        actualizarBtnEmail();
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

        document.getElementById('loading-text').textContent = 'Generando boletas PDF…';
        document.getElementById('loading-sub').textContent  = 'Esto puede tomar unos segundos.';
        document.getElementById('loading-overlay').classList.add('active');
        document.getElementById('form-boletas').submit();

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

    // ── Enviar correos ──────────────────────────────────────────
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