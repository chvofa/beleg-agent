/* ── Beleg-Agent Web UI – Client-Side JS ──────────────────────────────── */

// ── SSE Connection ──────────────────────────────────────────────────────

let eventSource = null;

function connectSSE() {
    if (eventSource) eventSource.close();
    eventSource = new EventSource('/api/events');

    eventSource.addEventListener('status', (e) => {
        const data = JSON.parse(e.data);
        updateStatusIndicator(data.agent);
    });

    eventSource.addEventListener('upload', (e) => {
        const data = JSON.parse(e.data);
        if (data.files && data.files.length > 0) {
            showToast(`${data.files.length} Datei(en) hochgeladen`);
        }
    });

    eventSource.addEventListener('task_output', (e) => {
        const data = JSON.parse(e.data);
        appendTerminal(data.task_id, data.line);
    });

    eventSource.addEventListener('task_complete', (e) => {
        const data = JSON.parse(e.data);
        const status = data.status === 'done' ? 'Abgeschlossen' : 'Fehler';
        showToast(`Task ${status}`);
        // Refresh active task buttons
        document.querySelectorAll('.task-running').forEach(el => {
            el.classList.remove('task-running');
            el.disabled = false;
        });
    });

    eventSource.onerror = () => {
        setTimeout(connectSSE, 5000);
    };
}

function updateStatusIndicator(status) {
    const dot = document.getElementById('status-dot');
    const text = document.getElementById('status-text');
    if (!dot) return;

    dot.className = 'dot';
    if (status === 'running' || status === 'restarted') {
        dot.classList.add('green');
        if (text) text.textContent = 'Läuft';
    } else if (status === 'stopped') {
        dot.classList.add('red');
        if (text) text.textContent = 'Gestoppt';
    }
}


// ── Toast Notifications ─────────────────────────────────────────────────

function showToast(message) {
    const container = document.getElementById('toast-container') || createToastContainer();
    const toast = document.createElement('div');
    toast.className = 'alert alert-success';
    toast.style.cssText = 'animation: fadeIn 0.2s; margin-bottom: 0.5rem;';
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 4000);
}

function createToastContainer() {
    const c = document.createElement('div');
    c.id = 'toast-container';
    c.style.cssText = 'position:fixed;top:1rem;right:1rem;z-index:999;max-width:350px;';
    document.body.appendChild(c);
    return c;
}


// ── Agent Control ───────────────────────────────────────────────────────

async function agentAction(action) {
    const res = await fetch(`/agent/${action}`, { method: 'POST' });
    const data = await res.json();
    if (data.ok) {
        showToast(`Agent: ${action}`);
        setTimeout(() => location.reload(), 1000);
    }
}


// ── File Upload ─────────────────────────────────────────────────────────

function initUploadZone() {
    const zone = document.getElementById('upload-zone');
    const input = document.getElementById('upload-input');
    const list = document.getElementById('upload-list');
    if (!zone) return;

    zone.addEventListener('click', () => input.click());

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
    });

    zone.addEventListener('dragleave', () => {
        zone.classList.remove('dragover');
    });

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        uploadFiles(e.dataTransfer.files);
    });

    input.addEventListener('change', () => {
        uploadFiles(input.files);
        input.value = '';
    });

    async function uploadFiles(files) {
        const formData = new FormData();
        for (const f of files) {
            formData.append('file', f);
        }
        list.innerHTML = '<li>Lade hoch...</li>';
        try {
            const res = await fetch('/api/upload', { method: 'POST', body: formData });
            const data = await res.json();
            if (data.ok && data.files) {
                list.innerHTML = data.files
                    .map(f => `<li>${f} <span class="badge badge-green">Hochgeladen</span></li>`)
                    .join('');
                showToast(`${data.files.length} Datei(en) in Inbox`);
            } else {
                list.innerHTML = `<li class="cell-red">${data.error || 'Fehler'}</li>`;
            }
        } catch (err) {
            list.innerHTML = `<li class="cell-red">Upload fehlgeschlagen</li>`;
        }
    }
}


// ── Number Formatting ───────────────────────────────────────────────────

function formatBetrag(raw, waehrung) {
    const val = parseFloat(raw);
    if (isNaN(val)) return raw || '';
    const prefix = val < 0 ? '-' : '';
    const abs = Math.abs(val).toFixed(2);
    const [intPart, dec] = abs.split('.');
    const formatted = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, "'");
    const result = prefix + formatted + '.' + dec;
    return waehrung ? result + ' ' + waehrung : result;
}


// ── Protocol Table ──────────────────────────────────────────────────────

let protocolData = [];
let sortCol = null;
let sortAsc = true;

async function loadProtocol() {
    const tbody = document.getElementById('protocol-body');
    if (!tbody) return;

    const res = await fetch('/api/protocol');
    protocolData = await res.json();
    // Originalindex merken (für Beleg-Viewer)
    protocolData.forEach((row, i) => row._idx = i);
    buildFilterOptions(protocolData);
    renderProtocol(getFilteredData());
}

function buildFilterOptions(rows) {
    const filters = {
        'filter-typ': 'Typ',
        'filter-zahlungsart': 'Zahlungsart',
        'filter-waehrung': 'Währung',
    };
    for (const [id, col] of Object.entries(filters)) {
        const el = document.getElementById(id);
        if (!el) continue;
        const values = [...new Set(rows.map(r => r[col]).filter(Boolean))].sort();
        const firstOpt = el.options[0].textContent;
        el.innerHTML = `<option value="">${firstOpt}</option>` +
            values.map(v => `<option value="${v}">${v}</option>`).join('');
    }
}

function escapeHtml(s) {
    const d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
}

function renderProtocol(rows) {
    const tbody = document.getElementById('protocol-body');
    if (!tbody) return;

    tbody.innerHTML = rows.map(row => {
        const abg = row.Abgeglichen === 'Ja';
        const conf = parseFloat(row.Confidence_Score) || 0;
        const bem = row.Bemerkungen || '';
        const hasFile = row.hat_datei;
        const icon = hasFile
            ? `<span class="beleg-icon" title="Beleg anzeigen" onclick="openBeleg(${row._idx})">&#128196;</span>`
            : `<span class="beleg-icon muted" title="Keine Datei">&#128196;</span>`;
        return `<tr>
            <td style="text-align:center; padding:0.4rem">${icon}</td>
            <td>${row.Datum_Rechnung || ''}</td>
            <td class="cell-muted">${row.Valutadatum || ''}</td>
            <td>${row.Rechnungssteller || ''}</td>
            <td>${row.Typ || ''}</td>
            <td style="text-align:right; font-variant-numeric:tabular-nums">${formatBetrag(row.Betrag, row.Währung)}</td>
            <td>${row.Zahlungsart || ''}</td>
            <td class="${abg ? 'cell-green' : 'cell-red'}">${row.Abgeglichen || 'Nein'}</td>
            <td class="${conf >= 0.85 ? '' : conf >= 0.6 ? 'cell-yellow' : 'cell-red'}">${conf ? (conf * 100).toFixed(0) + '%' : ''}</td>
            <td class="cell-bemerkung" ${bem ? `title="${escapeHtml(bem)}"` : ''}>${escapeHtml(bem)}</td>
        </tr>`;
    }).join('');

    const count = document.getElementById('protocol-count');
    if (count) count.textContent = `${rows.length} Belege`;
}

function getFilteredData() {
    const q = (document.getElementById('protocol-search')?.value || '').toLowerCase();
    const fTyp = document.getElementById('filter-typ')?.value || '';
    const fZahl = document.getElementById('filter-zahlungsart')?.value || '';
    const fAbg = document.getElementById('filter-abgeglichen')?.value || '';
    const fWhr = document.getElementById('filter-waehrung')?.value || '';

    return protocolData.filter(row => {
        if (fTyp && row.Typ !== fTyp) return false;
        if (fZahl && row.Zahlungsart !== fZahl) return false;
        if (fWhr && row.Währung !== fWhr) return false;
        if (fAbg) {
            const ist = row.Abgeglichen === 'Ja' ? 'Ja' : 'Nein';
            if (ist !== fAbg) return false;
        }
        if (q) {
            return Object.values(row).some(v => String(v).toLowerCase().includes(q));
        }
        return true;
    });
}

function filterProtocol() {
    renderProtocol(getFilteredData());
}

async function openBeleg(idx) {
    // Versuch: direkt als PDF im neuen Tab öffnen
    const res = await fetch(`/api/belege/view/${idx}`);
    const contentType = res.headers.get('Content-Type') || '';
    if (contentType.includes('application/pdf') || contentType.includes('image/')) {
        // PDF/Bild: in neuem Tab öffnen
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        window.open(url, '_blank');
    } else {
        // JSON-Antwort (OneDrive-Download oder Fehler)
        const data = await res.json();
        if (data.ok) {
            showToast(data.hinweis || 'Datei wird geöffnet');
        } else {
            showToast(data.error || 'Fehler beim Öffnen');
        }
    }
}

function resetFilters() {
    const ids = ['protocol-search', 'filter-typ', 'filter-zahlungsart', 'filter-abgeglichen', 'filter-waehrung'];
    ids.forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
    renderProtocol(protocolData);
}

function updateSortArrows() {
    document.querySelectorAll('th.sortable .sort-arrow').forEach(el => {
        const col = el.closest('th').dataset.col;
        if (col === sortCol) {
            el.textContent = sortAsc ? ' \u25B2' : ' \u25BC';
        } else {
            el.textContent = '';
        }
    });
}

function sortProtocol(colName) {
    if (sortCol === colName) {
        sortAsc = !sortAsc;
    } else {
        sortCol = colName;
        sortAsc = true;
    }
    const data = getFilteredData();
    data.sort((a, b) => {
        let va = a[colName] || '';
        let vb = b[colName] || '';
        if (!isNaN(va) && !isNaN(vb) && va !== '' && vb !== '') {
            va = parseFloat(va) || 0;
            vb = parseFloat(vb) || 0;
        }
        if (va < vb) return sortAsc ? -1 : 1;
        if (va > vb) return sortAsc ? 1 : -1;
        return 0;
    });
    renderProtocol(data);
    updateSortArrows();
}


// ── Reconciliation ──────────────────────────────────────────────────────

async function uploadCSV() {
    const input = document.getElementById('csv-input');
    if (!input?.files.length) return;

    const formData = new FormData();
    formData.append('file', input.files[0]);
    const res = await fetch('/api/reconciliation/upload', { method: 'POST', body: formData });
    const data = await res.json();
    if (data.ok) {
        showToast(`CSV hochgeladen: ${data.file}`);
    }
    input.value = '';
}

async function startReconciliation(type) {
    const btn = event.target;
    btn.disabled = true;
    btn.classList.add('task-running');
    const terminal = document.getElementById(`terminal-${type}`);
    if (terminal) terminal.textContent = 'Starte...\n';

    const res = await fetch(`/api/reconciliation/${type}`, { method: 'POST' });
    const data = await res.json();
    if (data.task_id && terminal) {
        terminal.dataset.taskId = data.task_id;
    }
}

function appendTerminal(taskId, line) {
    document.querySelectorAll('.terminal').forEach(el => {
        if (el.dataset.taskId === taskId) {
            el.textContent += line + '\n';
            el.scrollTop = el.scrollHeight;
        }
    });
}


// ── Review Queue ────────────────────────────────────────────────────────

async function loadReview() {
    const container = document.getElementById('review-list');
    if (!container) return;

    const res = await fetch('/api/review');
    const files = await res.json();
    const reviewFiles = files.filter(f => f.is_pruefen || f.is_duplikat);

    if (reviewFiles.length === 0) {
        container.innerHTML = '<div class="card"><p>Keine Dateien zur Prüfung.</p></div>';
        return;
    }

    container.innerHTML = reviewFiles.map(f => `
        <div class="review-card" data-filename="${f.name}">
            <div>
                <div class="filename">${f.name}</div>
                <div class="meta">${f.size > 1024 ? (f.size / 1024).toFixed(0) + ' KB' : f.size + ' B'} &middot; ${f.modified}</div>
                <div style="margin-top:0.5rem">
                    <span class="badge ${f.is_duplikat ? 'badge-yellow' : 'badge-red'}">
                        ${f.is_duplikat ? 'DUPLIKAT' : 'PRÜFEN'}
                    </span>
                </div>
            </div>
            <div class="review-form">
                <div class="form-row">
                    <div class="form-group">
                        <label>Rechnungssteller</label>
                        <input name="rechnungssteller" placeholder="z.B. Amazon">
                    </div>
                    <div class="form-group">
                        <label>Datum</label>
                        <input name="rechnungsdatum" type="date">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Betrag</label>
                        <input name="betrag" type="number" step="0.01" placeholder="0.00">
                    </div>
                    <div class="form-group">
                        <label>Währung</label>
                        <select name="waehrung">
                            <option>CHF</option><option>EUR</option><option>USD</option><option>GBP</option>
                        </select>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Typ</label>
                        <select name="typ">
                            <option>Rechnung</option><option>Gutschrift</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Zahlungsart</label>
                        <select name="zahlungsart">
                            <option value="">-</option>
                            <option>KK CHF</option><option>KK EUR</option>
                            <option>Überweisung</option><option>eBill</option>
                        </select>
                    </div>
                </div>
            </div>
            <div class="btn-group" style="flex-direction:column">
                <button class="btn btn-primary btn-sm" onclick="approveReview('${f.name}', this)">Freigeben</button>
                <button class="btn btn-danger btn-sm" onclick="rejectReview('${f.name}', this)">Löschen</button>
            </div>
        </div>
    `).join('');
}

async function approveReview(filename, btn) {
    const card = btn.closest('.review-card');
    const data = {};
    card.querySelectorAll('input, select').forEach(el => {
        if (el.name) data[el.name] = el.value;
    });

    if (!data.rechnungssteller || !data.rechnungsdatum || !data.betrag) {
        showToast('Bitte alle Pflichtfelder ausfüllen');
        return;
    }

    const res = await fetch(`/api/review/${encodeURIComponent(filename)}/approve`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
    });
    const result = await res.json();
    if (result.ok) {
        card.remove();
        showToast('Beleg freigegeben');
    } else {
        showToast(result.error || 'Fehler');
    }
}

async function rejectReview(filename, btn) {
    if (!confirm('Datei wirklich löschen?')) return;
    const card = btn.closest('.review-card');
    const res = await fetch(`/api/review/${encodeURIComponent(filename)}/reject`, {
        method: 'POST',
    });
    const result = await res.json();
    if (result.ok) {
        card.remove();
        showToast('Datei gelöscht');
    }
}


// ── Logs ────────────────────────────────────────────────────────────────

async function loadLogs() {
    const viewer = document.getElementById('log-viewer');
    if (!viewer) return;

    const res = await fetch('/api/logs?lines=200');
    const data = await res.json();
    viewer.textContent = data.lines.join('\n');
    viewer.scrollTop = viewer.scrollHeight;
}


// ── Settings ────────────────────────────────────────────────────────────

async function loadSettings() {
    const form = document.getElementById('settings-form');
    if (!form) return;

    const res = await fetch('/api/settings');
    const data = await res.json();

    form.querySelector('[name="ablage_stammpfad"]').value = data.ablage_stammpfad || '';
    form.querySelector('[name="bank_profil"]').value = data.bank_profil || 'ubs';
    form.querySelector('[name="confidence_auto"]').value = data.confidence_auto || 0.85;
    form.querySelector('[name="confidence_rueckfrage"]').value = data.confidence_rueckfrage || 0.60;
}

async function saveSettings(e) {
    e.preventDefault();
    const form = document.getElementById('settings-form');
    const data = {
        ablage_stammpfad: form.querySelector('[name="ablage_stammpfad"]').value,
        bank_profil: form.querySelector('[name="bank_profil"]').value,
        confidence_auto: parseFloat(form.querySelector('[name="confidence_auto"]').value),
        confidence_rueckfrage: parseFloat(form.querySelector('[name="confidence_rueckfrage"]').value),
    };

    const res = await fetch('/api/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
    });
    const result = await res.json();
    showToast(result.hinweis || 'Gespeichert');
}


// ── Export ───────────────────────────────────────────────────────────────

function exportProtocol(format) {
    const von = document.getElementById('export-von')?.value || '';
    const bis = document.getElementById('export-bis')?.value || '';
    window.location.href = `/api/protocol/export?format=${format}&von=${von}&bis=${bis}`;
}

function downloadBelege() {
    const von = document.getElementById('export-von')?.value || '';
    const bis = document.getElementById('export-bis')?.value || '';
    window.location.href = `/api/belege/download?von=${von}&bis=${bis}`;
}


// ── Init ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
    connectSSE();
    initUploadZone();
    loadProtocol();
    loadReview();
    loadLogs();
    loadSettings();
});
