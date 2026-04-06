"""Скрипт для добавления JS кода высвобождения в app.js."""
import sys

js_code = r"""

// ============================================================
// Загрузка ДС на высвобождение
// ============================================================

const vysvUploadArea = document.getElementById('vysv-upload-area');
const vysvFileInput = document.getElementById('vysv-file-input');
const vysvUploadBtn = document.getElementById('vysv-upload-btn');
let vysvSelectedFile = null;

vysvUploadArea.addEventListener('click', () => vysvFileInput.click());
vysvUploadArea.addEventListener('dragover', (e) => { e.preventDefault(); vysvUploadArea.classList.add('dragover'); });
vysvUploadArea.addEventListener('dragleave', () => { vysvUploadArea.classList.remove('dragover'); });
vysvUploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    vysvUploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) handleVysvFile(files[0]);
});
vysvFileInput.addEventListener('change', () => {
    if (vysvFileInput.files.length > 0) handleVysvFile(vysvFileInput.files[0]);
});

function handleVysvFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'doc' && ext !== 'docx') {
        showAlert('error', 'Поддерживаются только файлы DOC и DOCX');
        return;
    }
    vysvSelectedFile = file;
    vysvUploadArea.querySelector('.upload-text').textContent = file.name;
    vysvUploadArea.querySelector('.upload-hint').textContent = formatFileSize(file.size);
    vysvUploadBtn.disabled = false;
}

const VYSV_UPLOAD_STEPS = [
    { id: 'vysv-step-upload', key: 'upload', label: 'Загрузка файла на сервер' },
    { id: 'vysv-step-parse',  key: 'parse',  label: 'Парсинг документа' },
    { id: 'vysv-step-check',  key: 'check',  label: 'Проверка данных' },
    { id: 'vysv-step-save',   key: 'save',   label: 'Сохранение результатов' },
];

function renderVysvProgress(container) {
    container.innerHTML = '<div class="upload-progress">' +
        VYSV_UPLOAD_STEPS.map(s =>
            '<div class="upload-progress-step" id="' + s.id + '">' +
            '<span class="upload-step-icon"><i class="fas fa-circle" style="color: #d9d9d9; font-size: 10px;"></i></span>' +
            '<span class="upload-step-label" style="color: var(--text-secondary);">' + s.label + '</span>' +
            '</div>'
        ).join('') +
        '</div>';
}

function setVysvStepStatus(key, status, label) {
    const step = VYSV_UPLOAD_STEPS.find(s => s.key === key);
    if (!step) return;
    const el = document.getElementById(step.id);
    if (!el) return;
    const iconEl = el.querySelector('.upload-step-icon');
    const labelEl = el.querySelector('.upload-step-label');
    if (status === 'progress') {
        iconEl.innerHTML = '<i class="fas fa-spinner fa-spin" style="color: #1890ff;"></i>';
        labelEl.style.color = 'var(--text-primary)';
        labelEl.style.fontWeight = '500';
    } else if (status === 'done') {
        iconEl.innerHTML = '<i class="fas fa-check-circle" style="color: #52c41a;"></i>';
        labelEl.style.color = 'var(--text-primary)';
        labelEl.style.fontWeight = 'normal';
    } else if (status === 'error') {
        iconEl.innerHTML = '<i class="fas fa-times-circle" style="color: #ff4d4f;"></i>';
        labelEl.style.color = '#ff4d4f';
        labelEl.style.fontWeight = 'normal';
    } else if (status === 'warn') {
        iconEl.innerHTML = '<i class="fas fa-exclamation-circle" style="color: #faad14;"></i>';
        labelEl.style.color = 'var(--text-primary)';
        labelEl.style.fontWeight = 'normal';
    }
    if (label) labelEl.textContent = label;
}

async function uploadVysvFile() {
    if (!vysvSelectedFile) return;
    vysvUploadBtn.disabled = true;
    vysvUploadBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Проверка...';
    const container = document.getElementById('vysv-upload-result');
    renderVysvProgress(container);
    setVysvStepStatus('upload', 'progress', 'Загрузка файла на сервер...');
    const formData = new FormData();
    formData.append('file', vysvSelectedFile);
    try {
        await new Promise((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            xhr.open('POST', '/api/upload-vysvobozhdenie');
            let sseBuffer = '';
            function parseSseBuffer(text) {
                const events = text.split('\n\n');
                const remaining = events.pop();
                for (const eventStr of events) {
                    if (!eventStr.trim()) continue;
                    const eventMatch = eventStr.match(/^event: (\w+)/m);
                    const dataMatch = eventStr.match(/^data: (.+)$/m);
                    if (!eventMatch || !dataMatch) continue;
                    const eventType = eventMatch[1];
                    let data;
                    try { data = JSON.parse(dataMatch[1]); } catch { continue; }
                    if (eventType === 'progress') {
                        setVysvStepStatus(data.step, data.status, data.label);
                    } else if (eventType === 'result') {
                        displayVysvResult(data);
                        loadStats();
                        loadAgreements();
                    } else if (eventType === 'error') {
                        container.innerHTML = '<div class="alert alert-error"><i class="fas fa-exclamation-circle"></i><div>' + data.message + '</div></div>';
                    }
                }
                return remaining;
            }
            xhr.upload.onprogress = (e) => {
                if (e.lengthComputable) {
                    const pct = Math.round(e.loaded / e.total * 100);
                    setVysvStepStatus('upload', 'progress', 'Загрузка файла на сервер... ' + pct + '%');
                }
            };
            let readOffset = 0;
            xhr.onprogress = () => {
                const newChunk = xhr.responseText.slice(readOffset);
                readOffset = xhr.responseText.length;
                sseBuffer += newChunk;
                sseBuffer = parseSseBuffer(sseBuffer);
            };
            xhr.onload = () => {
                if (xhr.status >= 200 && xhr.status < 300) {
                    if (sseBuffer.trim()) parseSseBuffer(sseBuffer + '\n\n');
                    resolve();
                } else {
                    reject(new Error('HTTP ' + xhr.status));
                }
            };
            xhr.onerror = () => reject(new Error('Ошибка сети'));
            xhr.send(formData);
        });
    } catch (error) {
        showAlert('error', 'Ошибка: ' + error.message);
    } finally {
        vysvUploadBtn.disabled = false;
        vysvUploadBtn.innerHTML = '<i class="fas fa-check"></i> Проверить';
    }
}

function displayVysvResult(result) {
    const container = document.getElementById('vysv-upload-result');
    if (result.detail) {
        container.innerHTML = '<div class="alert alert-error"><i class="fas fa-exclamation-circle"></i><div>' + result.detail + '</div></div>';
        return;
    }
    const fmt = function(n) { return n != null ? n.toLocaleString('ru-RU', {minimumFractionDigits: 2}) : '—'; };
    const errorsHtml = result.errors && result.errors.length > 0
        ? '<div class="alert alert-error"><i class="fas fa-times-circle"></i><div><strong>Ошибки (' + result.errors.length + '):</strong><ul style="margin: 8px 0 0 16px;">' + result.errors.map(function(e) { return '<li>' + e + '</li>'; }).join('') + '</ul></div></div>'
        : '';
    const warningsHtml = result.warnings && result.warnings.length > 0
        ? '<div class="alert alert-warning"><i class="fas fa-exclamation-triangle"></i><div><strong>Предупреждения (' + result.warnings.length + '):</strong><ul style="margin: 8px 0 0 16px;">' + result.warnings.map(function(w) { return '<li>' + w + '</li>'; }).join('') + '</ul></div></div>'
        : '';
    const statusHtml = result.is_valid
        ? '<div class="alert alert-success"><i class="fas fa-check-circle"></i> Проверка пройдена успешно</div>'
        : '<div class="alert alert-error"><i class="fas fa-times-circle"></i> Найдены ошибки</div>';
    const checksHtml = buildCheckSummary(result.checks || []);
    const general = (result.data && result.data.general) || {};
    const infoRows = [
        ['№ ДС', general.ds_number || '—'],
        ['Контракт', general.contract_short_number ? 'ГК ' + general.contract_short_number : (general.contract_number || '—')],
        ['Закрытый этап', general.closed_stage != null ? 'Этап ' + general.closed_stage : '—'],
        ['Сумма закрытого этапа', general.closed_amount != null ? fmt(general.closed_amount) + ' руб.' : '—'],
        ['Новая цена контракта', general.new_contract_price != null ? fmt(general.new_contract_price) + ' руб.' : '—'],
    ];
    const infoTableHtml = '<div style="margin-top: 16px;"><div style="font-weight: 600; margin-bottom: 8px;">Основные данные документа</div><table style="width: 100%; border-collapse: collapse;">' +
        infoRows.map(function(r) { return '<tr><td style="padding: 4px 12px 4px 0; color: var(--text-secondary); white-space: nowrap;">' + r[0] + '</td><td style="padding: 4px 0; font-weight: 500;">' + r[1] + '</td></tr>'; }).join('') +
        '</table></div>';
    const stages1 = (result.data && result.data.stages_table1) || [];
    const itogo_km = result.data && result.data.itogo_km;
    const itogo_price = result.data && result.data.itogo_price;
    const stagesInfoHtml = stages1.length > 0
        ? '<div style="margin-top: 12px; font-size: 13px; color: var(--text-secondary);">Таблица этапов: ' + stages1.length + ' строк' + (itogo_km != null ? ' | ИТОГО км: ' + fmt(itogo_km) : '') + (itogo_price != null ? ' | ИТОГО цена: ' + fmt(itogo_price) + ' руб.' : '') + '</div>'
        : '';
    container.innerHTML =
        '<div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; flex-wrap: wrap; gap: 8px;">' +
        '<div><strong>ДС №' + (result.ds_number || '?') + ' (высвобождение)</strong>' +
        (result.contract_number ? ' <span class="badge badge-default">ГК ' + result.contract_number + '</span>' : '') +
        '</div></div>' +
        statusHtml + errorsHtml + warningsHtml + checksHtml + infoTableHtml + stagesInfoHtml;
}
"""

with open('static/js/app.js', 'ab') as f:
    f.write(js_code.encode('utf-8'))

print(f'Added {len(js_code)} bytes of JS')
