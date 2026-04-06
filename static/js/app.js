// Theme
function toggleTheme() {
    document.body.classList.toggle('dark-mode');
    localStorage.setItem('theme', document.body.classList.contains('dark-mode') ? 'dark' : 'light');
}

// Load saved theme
if (localStorage.getItem('theme') === 'dark') {
    document.body.classList.add('dark-mode');
}

// Loading Overlay
function showLoadingOverlay(message = 'Загрузка...') {
    // Удаляем старый оверлей если есть
    hideLoadingOverlay();

    const overlay = document.createElement('div');
    overlay.id = 'loading-overlay';
    overlay.innerHTML = `
        <div style="
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 10000;
        ">
            <div style="
                background: var(--card-bg, #fff);
                padding: 32px 48px;
                border-radius: 12px;
                text-align: center;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
            ">
                <div class="spinner" style="margin: 0 auto 16px;"></div>
                <div style="font-size: 16px; color: var(--text-primary, #333);">${message}</div>
                <div style="font-size: 13px; color: var(--text-secondary, #666); margin-top: 8px;">
                    Это может занять некоторое время
                </div>
            </div>
        </div>
    `;
    document.body.appendChild(overlay);
}

function hideLoadingOverlay() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.remove();
    }
}

// Tabs
document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

        tab.classList.add('active');
        document.getElementById('tab-' + tab.dataset.tab).classList.add('active');

        // Load data for tab
        if (tab.dataset.tab === 'agreements') {
            loadAgreements();
        } else if (tab.dataset.tab === 'calendar') {
            loadCalendar();
        } else if (tab.dataset.tab === 'stages') {
            loadStages();
        } else if (tab.dataset.tab === 'table-checks') {
            // Загружаем список ГК один раз (если ещё не загружен)
            const contractSel = document.getElementById('tc-contract');
            if (contractSel && contractSel.options.length <= 1) {
                tcLoadContracts();
            }
        } else if (tab.dataset.tab === 'acts') {
            loadActsHistory();
        }
    });
});

// File Upload
const uploadArea = document.getElementById('upload-area');
const fileInput = document.getElementById('file-input');
const uploadBtn = document.getElementById('upload-btn');
let selectedFile = null;

function resetUploadArea() {
    selectedFile = null;
    fileInput.value = '';
    uploadArea.querySelector('.upload-text').textContent = 'Перетащите файл сюда или нажмите для выбора';
    uploadArea.querySelector('.upload-hint').textContent = 'Архив (ZIP, RAR) — изменение параметров, документ (DOC, DOCX) — высвобождение';
    uploadBtn.style.display = '';
    uploadBtn.disabled = true;
    uploadBtn.innerHTML = '<i class="fas fa-check"></i> Проверить';
    const resetBtn = document.getElementById('reset-btn');
    if (resetBtn) resetBtn.style.display = 'none';
    const badge = document.getElementById('upload-file-type');
    if (badge) badge.style.display = 'none';
    const resultContainer = document.getElementById('upload-result');
    if (resultContainer) {
        resultContainer.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon"><i class="fas fa-inbox"></i></div>
                <div>Загрузите архив для проверки</div>
            </div>
        `;
    }
    lastUploadResult = null;
    lastAgreementId = null;
}

function showResetButton() {
    uploadBtn.style.display = 'none';
    const resetBtn = document.getElementById('reset-btn');
    if (resetBtn) resetBtn.style.display = '';
}

uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
        handleFile(fileInput.files[0]);
    }
});

function handleFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    const isArchive = ext === 'zip' || ext === 'rar';
    const isDoc = ext === 'doc' || ext === 'docx';
    if (!isArchive && !isDoc) {
        showAlert('error', 'Поддерживаются ZIP, RAR, DOC и DOCX файлы');
        return;
    }

    selectedFile = file;
    uploadArea.querySelector('.upload-text').textContent = file.name;
    uploadArea.querySelector('.upload-hint').textContent = formatFileSize(file.size);
    uploadBtn.disabled = false;

    const badge = document.getElementById('upload-file-type');
    if (badge) {
        badge.textContent = isArchive ? 'Архив (изменение параметров)' : 'Документ (высвобождение)';
        badge.style.display = 'block';
    }
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

const UPLOAD_STEPS = [
    { id: 'step-upload', key: 'upload', label: 'Загрузка файла на сервер' },
    { id: 'step-extract', key: 'extract', label: 'Распаковка и парсинг архива' },
    { id: 'step-km', key: 'km', label: 'Поиск данных км' },
    { id: 'step-check', key: 'check', label: 'Проверка данных' },
    { id: 'step-save', key: 'save', label: 'Сохранение результатов' },
];

function renderUploadProgress(container) {
    container.innerHTML = `
        <div class="upload-progress">
            ${UPLOAD_STEPS.map(s => `
                <div class="upload-progress-step" id="${s.id}">
                    <span class="upload-step-icon"><i class="fas fa-circle" style="color: #d9d9d9; font-size: 10px;"></i></span>
                    <span class="upload-step-label" style="color: var(--text-secondary);">${s.label}</span>
                </div>
            `).join('')}
        </div>
    `;
}

function setStepStatus(key, status, label) {
    const step = UPLOAD_STEPS.find(s => s.key === key);
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
    } else if (status === 'warn') {
        iconEl.innerHTML = '<i class="fas fa-exclamation-circle" style="color: #faad14;"></i>';
        labelEl.style.color = 'var(--text-primary)';
        labelEl.style.fontWeight = 'normal';
    } else if (status === 'skipped') {
        iconEl.innerHTML = '<i class="fas fa-minus-circle" style="color: #d9d9d9;"></i>';
        labelEl.style.color = 'var(--text-secondary)';
        labelEl.style.fontWeight = 'normal';
    } else if (status === 'error') {
        iconEl.innerHTML = '<i class="fas fa-times-circle" style="color: #ff4d4f;"></i>';
        labelEl.style.color = '#ff4d4f';
        labelEl.style.fontWeight = 'normal';
    }

    if (label) labelEl.textContent = label;
}

async function uploadFile() {
    if (!selectedFile) return;

    uploadBtn.disabled = true;
    uploadBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Проверка...';

    const container = document.getElementById('upload-result');
    renderUploadProgress(container);

    setStepStatus('upload', 'progress', 'Загрузка файла на сервер...');

    const formData = new FormData();
    formData.append('file', selectedFile);

    const ext = selectedFile.name.split('.').pop().toLowerCase();
    const endpoint = (ext === 'doc' || ext === 'docx') ? '/api/upload-vysvobozhdenie' : '/api/upload';

    try {
        // Используем XHR: upload.onprogress — процент загрузки,
        // xhr.onprogress — потоковое чтение SSE-ответа
        await new Promise((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            xhr.open('POST', endpoint);
            let sseBuffer = '';

            function parseSseBuffer(text) {
                const events = text.split('\n\n');
                const remaining = events.pop(); // возможно неполный фрагмент
                for (const eventStr of events) {
                    if (!eventStr.trim()) continue;
                    const eventMatch = eventStr.match(/^event: (\w+)/m);
                    const dataMatch = eventStr.match(/^data: (.+)$/m);
                    if (!eventMatch || !dataMatch) continue;
                    const eventType = eventMatch[1];
                    let data;
                    try { data = JSON.parse(dataMatch[1]); } catch { continue; }
                    if (eventType === 'progress') {
                        setStepStatus(data.step, data.status, data.label);
                    } else if (eventType === 'result') {
                        lastUploadResult = data;
                        if (data.has_vysvobozhdenie && data.vysvobozhdenie) {
                            showVysvobozhdenieDialog(data);
                        } else {
                            displayUploadResult(data);
                        }
                        showResetButton();
                        loadStats();
                        loadAgreements();
                    } else if (eventType === 'error') {
                        container.innerHTML = `
                            <div class="alert alert-error">
                                <i class="fas fa-exclamation-circle"></i>
                                <div>${data.message}</div>
                            </div>
                        `;
                    }
                }
                return remaining;
            }

            // Прогресс загрузки файла на сервер
            xhr.upload.onprogress = (e) => {
                if (e.lengthComputable) {
                    const pct = Math.round(e.loaded / e.total * 100);
                    setStepStatus('upload', 'progress', `Загрузка файла на сервер... ${pct}%`);
                }
            };

            // Потоковое чтение SSE-ответа
            let readOffset = 0;
            xhr.onprogress = () => {
                const newChunk = xhr.responseText.slice(readOffset);
                readOffset = xhr.responseText.length;
                sseBuffer += newChunk;
                sseBuffer = parseSseBuffer(sseBuffer);
            };

            xhr.onload = () => {
                if (xhr.status >= 200 && xhr.status < 300) {
                    // Разбираем остаток буфера
                    if (sseBuffer.trim()) parseSseBuffer(sseBuffer + '\n\n');
                    resolve();
                } else {
                    reject(new Error(`HTTP ${xhr.status}`));
                }
            };
            xhr.onerror = () => reject(new Error('Ошибка сети'));
            xhr.send(formData);
        });

    } catch (error) {
        showAlert('error', 'Ошибка: ' + error.message);
    } finally {
        // Сбрасываем только выбор файла; кнопки управляются через showResetButton() при успехе
        selectedFile = null;
        fileInput.value = '';
        uploadArea.querySelector('.upload-text').textContent = 'Перетащите файл сюда или нажмите для выбора';
        uploadArea.querySelector('.upload-hint').textContent = 'Архив (ZIP, RAR) — изменение параметров, документ (DOC, DOCX) — высвобождение';
        const badge = document.getElementById('upload-file-type');
        if (badge) badge.style.display = 'none';
        // Если результат не получен (ошибка сети/сервера) — возвращаем кнопку "Проверить"
        if (!lastUploadResult) {
            uploadBtn.style.display = '';
            uploadBtn.disabled = true;
            uploadBtn.innerHTML = '<i class="fas fa-check"></i> Проверить';
        }
    }
}

let lastUploadResult = null;
let lastAgreementId = null;
let currentViewedAgreement = null;

// ============================================================================
// Диалог подтверждения высвобождения в архиве
// ============================================================================

function showVysvobozhdenieDialog(result) {
    const vysv = result.vysvobozhdenie;
    if (!vysv) return;

    const container = document.getElementById('upload-result');
    const closedStage = vysv.closed_stage || '-';
    const closedAmount = vysv.closed_amount
        ? vysv.closed_amount.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
        : '-';
    const newPrice = vysv.new_contract_price
        ? vysv.new_contract_price.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
        : '-';

    container.innerHTML = `
        <div style="padding: 16px; background: #fffbe6; border: 1px solid #ffe58f; border-radius: 8px; margin-bottom: 16px;">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px;">
                <i class="fas fa-exclamation-triangle" style="color: #d48806; font-size: 24px;"></i>
                <div>
                    <strong style="font-size: 16px; color: #d48806;">Обнаружено положение о высвобождении</strong>
                    <div style="font-size: 13px; color: #666; margin-top: 4px;">
                        В архиве с ДС на изменение параметров найден раздел о высвобождении средств
                    </div>
                </div>
            </div>
            <div style="background: #fff; padding: 12px; border-radius: 6px; border: 1px solid #d9d9d9;">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 200px; color: var(--text-secondary);">Закрываемый этап:</td>
                        <td style="font-weight: 600;">${closedStage}</td>
                    </tr>
                    <tr>
                        <td style="color: var(--text-secondary);">Сумма высвобождения:</td>
                        <td style="font-weight: 600; font-family: monospace;">${closedAmount} руб.</td>
                    </tr>
                    <tr>
                        <td style="color: var(--text-secondary);">Новая цена контракта:</td>
                        <td style="font-weight: 600; font-family: monospace;">${newPrice} руб.</td>
                    </tr>
                </table>
            </div>
            <div style="margin-top: 16px; display: flex; gap: 12px;">
                <button class="btn btn-primary" onclick="confirmVysvobozhdenie(true)">
                    <i class="fas fa-check"></i> Использовать данные о высвобождении
                </button>
                <button class="btn btn-default" onclick="confirmVysvobozhdenie(false)">
                    <i class="fas fa-times"></i> Только изменение параметров
                </button>
            </div>
        </div>
    `;
}

function confirmVysvobozhdenie(useVysvobozhdenie) {
    if (!lastUploadResult) return;

    if (!useVysvobozhdenie && lastUploadResult.vysvobozhdenie) {
        const modifiedResult = JSON.parse(JSON.stringify(lastUploadResult));
        delete modifiedResult.vysvobozhdenie;
        modifiedResult.has_vysvobozhdenie = false;
        if (modifiedResult.data) {
            modifiedResult.data = JSON.parse(JSON.stringify(lastUploadResult.data));
            delete modifiedResult.data.vysvobozhdenie;
        }
        lastUploadResult = modifiedResult;
    }

    displayUploadResult(lastUploadResult);
}

function displayUploadResult(result) {
    lastUploadResult = result;
    lastAgreementId = result.agreement_id || null;
    const container = document.getElementById('upload-result');

    if (result.detail) {
        container.innerHTML = `
            <div class="alert alert-error">
                <i class="fas fa-exclamation-circle"></i>
                <div>${result.detail}</div>
            </div>
        `;
        return;
    }

    const errorsHtml = result.errors && result.errors.length > 0
        ? `<div class="alert alert-error">
            <i class="fas fa-times-circle"></i>
            <div>
                <strong>Ошибки (${result.errors.length}):</strong>
                <ul style="margin: 8px 0 0 16px;">
                    ${result.errors.map(e => `<li>${e}</li>`).join('')}
                </ul>
            </div>
        </div>`
        : '';

    const warningsHtml = result.warnings && result.warnings.length > 0
        ? `<div class="alert alert-warning">
            <i class="fas fa-exclamation-triangle"></i>
            <div>
                <strong>Предупреждения (${result.warnings.length}):</strong>
                <ul style="margin: 8px 0 0 16px;">
                    ${result.warnings.map(w => `<li>${w}</li>`).join('')}
                </ul>
            </div>
        </div>`
        : '';

    const statusHtml = result.is_valid
        ? '<div class="alert alert-success"><i class="fas fa-check-circle"></i> Проверка пройдена успешно</div>'
        : '<div class="alert alert-error"><i class="fas fa-times-circle"></i> Найдены ошибки</div>';

    // Build data tables
    let tablesHtml = '';
    if (result.data) {
        tablesHtml = buildDataTables(result.data, result.checks || []);
    }

    const kmData = result.data && result.data.km_data;
    const kmFilename = kmData
        ? `КМ_Приложение${kmData.appendix_number || ''}_ДС${kmData.ds_number || result.ds_number || 'data'}.json`
        : '';

    container.innerHTML = `
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; flex-wrap: wrap; gap: 8px;">
            <div>
                <strong>ДС №${result.ds_number || '?'}</strong>
                ${result.contract_number ? `<span class="badge badge-default">ГК ${result.contract_number}</span>` : ''}
            </div>
            ${lastAgreementId ? `
            <button class="btn btn-warning btn-sm" onclick="startManualEdit()">
                <i class="fas fa-pen"></i> Ручное редактирование
            </button>
            ` : ''}
        </div>
        ${statusHtml}
        ${errorsHtml}
        ${warningsHtml}
        ${tablesHtml}
        ${result.data ? `
            <div style="margin-top: 16px; display: flex; gap: 8px; flex-wrap: wrap;">
                <button class="btn btn-default btn-sm" onclick="downloadJson(lastUploadResult.data, 'ДС_${lastUploadResult.ds_number || 'data'}.json')">
                    <i class="fas fa-download"></i> Скачать JSON с общей информацией
                </button>
                ${kmData ? `
                <button class="btn btn-default btn-sm" onclick="downloadJson(lastUploadResult.data.km_data, '${kmFilename}')">
                    <i class="fas fa-download"></i> Скачать JSON с расчётами
                </button>
                ` : ''}
            </div>
        ` : ''}
    `;
}

function buildCheckSummary(checks) {
    if (!checks || checks.length === 0) return '';

    return `
    <div class="check-summary">
        ${checks.map(c => {
            const cssClass = !c.ok ? 'check-fail' : (c.has_warnings ? 'check-warn' : 'check-ok');
            const icon = !c.ok ? 'fa-times-circle' : (c.has_warnings ? 'fa-exclamation-circle' : 'fa-check-circle');
            const detail = c.detail || (c.errors.length > 0 ? `${c.errors.length} ошибок` : (c.warnings.length > 0 ? `${c.warnings.length} предупреждений` : ''));
            return `
                <div class="check-summary-item ${cssClass}">
                    <i class="fas ${icon}"></i>
                    <div>
                        <div class="check-summary-label">${c.label}</div>
                        ${detail ? `<div class="check-summary-detail">${detail}</div>` : ''}
                    </div>
                </div>
            `;
        }).join('')}
    </div>
    `;
}

function buildDataTables(data, checks) {
    let html = '';

    // General info table
    const general = data.general || {};

    html += buildCheckSummary(checks);

    html += `
        <div class="collapsible-section" data-section="general">
            <div class="collapsible-header" onclick="toggleSection(this)">
                <h4><i class="fas fa-info-circle"></i> Общие данные</h4>
                <span class="collapsible-toggle"><i class="fas fa-chevron-down"></i></span>
            </div>
            <div class="collapsible-content">
                <table>
                    <tbody>
                        <tr><td style="width: 200px;"><strong>Номер ДС</strong></td><td>${general.ds_number || '-'}</td></tr>
                        <tr><td><strong>Номер контракта</strong></td><td>${general.contract_number || '-'}</td></tr>
                        ${general.sum_text ? `<tr><td><strong>Сумма</strong></td><td>${formatNumber(general.sum_text)} руб.</td></tr>` : ''}
                        ${general.probeg_sravnenie ? `<tr><td><strong>Пробег</strong></td><td>${formatNumber(general.probeg_sravnenie)} км</td></tr>` : ''}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    // Changes table
    const changesWithMoney = data.change_with_money || [];
    const changesWithoutMoney = data.change_without_money || [];
    const changesWithMoneyNoAppendix = data.change_with_money_no_appendix || [];
    const changesWithoutMoneyNoAppendix = data.change_without_money_no_appendix || [];
    
    // Изменения с приложениями
    const changesWithAppendix = [
        ...changesWithMoney.map(c => ({...c, type: 'Объёмы'})),
        ...changesWithoutMoney.map(c => ({...c, type: 'Без объёмов'}))
    ];

    // Изменения без приложений
    const changesNoAppendix = [
        ...changesWithMoneyNoAppendix.map(c => ({...c, type: 'Объёмы', has_appendix: false})),
        ...changesWithoutMoneyNoAppendix.map(c => ({...c, type: 'Без объёмов', has_appendix: false}))
    ];

    // Изменения без приложений
    if (changesNoAppendix.length > 0) {
        html += `
            <div class="collapsible-section" data-section="changes-without-appendix">
                <div class="collapsible-header" onclick="toggleSection(this)">
                    <h4><i class="fas fa-exchange-alt"></i> Изменения без приложений (${changesNoAppendix.length})</h4>
                    <span class="collapsible-toggle"><i class="fas fa-chevron-down"></i></span>
                </div>
                <div class="collapsible-content">
                    <div style="overflow-x: auto;">
                        <table>
                            <thead>
                                <tr>
                                    <th>Маршрут</th>
                                    <th>Тип</th>
                                    <th>Тип дня</th>
                                    <th>Дата с</th>
                                    <th>Дата по</th>
                                    <th>Дата на</th>
                                    <th>Описание</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${changesNoAppendix.map(c => `
                                    <tr>
                                        <td><strong>${c.route || '-'}</strong></td>
                                        <td><span class="badge ${c.type === 'Объёмы' ? 'badge-success' : 'badge-default'}">${c.type}</span></td>
                                        <td>${c.day_type || '-'}</td>
                                        <td>${c.date_from || '-'}</td>
                                        <td>${c.date_to || '-'}</td>
                                        <td>${c.date_on || '-'}</td>
                                        <td style="font-size: 13px;">${c.point || '-'}</td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;
    }

    // Изменения с приложениями
    if (changesWithAppendix.length > 0) {
        html += `
            <div class="collapsible-section" data-section="changes-with-appendix">
                <div class="collapsible-header" onclick="toggleSection(this)">
                    <h4><i class="fas fa-exchange-alt"></i> Изменения с приложениями (${changesWithAppendix.length})</h4>
                    <span class="collapsible-toggle"><i class="fas fa-chevron-down"></i></span>
                </div>
                <div class="collapsible-content">
                    <div style="overflow-x: auto;">
                        <table>
                            <thead>
                                <tr>
                                    <th>Маршрут</th>
                                    <th>Прил.</th>
                                    <th>Тип</th>
                                    <th>Тип дня</th>
                                    <th>Дата с</th>
                                    <th>Дата по</th>
                                    <th>Дата на</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${changesWithAppendix.map(c => `
                                    <tr>
                                        <td><strong>${c.route || '-'}</strong></td>
                                        <td>${c.appendix || '-'}</td>
                                        <td><span class="badge ${c.type === 'Объёмы' ? 'badge-success' : 'badge-default'}">${c.type}</span></td>
                                        <td>${c.day_type || '-'}</td>
                                        <td>${c.date_from || '-'}</td>
                                        <td>${c.date_to || '-'}</td>
                                        <td>${c.date_on || '-'}</td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;
    }

    // Appendices table
    const appendices = data.appendices || {};
    const appendixKeys = Object.keys(appendices);

    if (appendixKeys.length > 0) {
        html += `
            <div class="collapsible-section" data-section="appendices">
                <div class="collapsible-header" onclick="toggleSection(this)">
                    <h4><i class="fas fa-file-alt"></i> Приложения (${appendixKeys.length})</h4>
                    <span class="collapsible-toggle"><i class="fas fa-chevron-down"></i></span>
                </div>
                <div class="collapsible-content">
        `;

        appendixKeys.forEach(key => {
            try {
                const app = appendices[key];
                if (!app) return;

                const numTypes = app.num_of_types || 0;

                html += `
                    <div style="margin-bottom: 16px; padding: 16px; background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px;">
                        <div style="margin-bottom: 12px;">
                            <strong>Приложение № ${app.appendix_num || key}</strong>
                            <span class="badge badge-default">${app.route || '-'}</span>
                            ${app.date_from ? `<span style="color: #8c8c8c; margin-left: 8px;">с ${app.date_from}</span>` : ''}
                            ${app.date_to ? `<span style="color: #8c8c8c;"> по ${app.date_to}</span>` : ''}
                            ${app.date_on ? `<span style="color: #8c8c8c; margin-left: 8px;">на ${app.date_on}</span>` : ''}
                        </div>
                        <div style="margin-bottom: 8px; font-size: 13px; color: #8c8c8c;">
                            Протяжённость: от НП ${app.length_forward || '-'} км, от КП ${app.length_reverse || '-'} км
                        </div>
                `;

                const lengthFwd = app.length_forward || 0;
                const lengthRev = app.length_reverse || 0;

                // Check if there are winter/summer periods
                const hasWinter = app.period_winter && app.period_winter.num_of_types > 0;
                const hasSummer = app.period_summer && app.period_summer.num_of_types > 0;

                if (numTypes > 0 || hasWinter || hasSummer) {
                    if (hasWinter || hasSummer) {
                    // Render winter period if exists
                    if (hasWinter) {
                        const winter = app.period_winter;
                        const winterLengthFwd = app.length_forward || 0;
                        const winterLengthRev = app.length_reverse || 0;

                        html += `
                            <div style="margin-top: 12px; padding: 12px; background: var(--bg-light); border-left: 3px solid #1890ff; border-radius: 4px;">
                                <strong style="color: #1890ff;"><i class="fas fa-snowflake"></i> Зимний период</strong>
                                ${winter.date_from ? `<span style="color: #8c8c8c; font-size: 12px; margin-left: 8px;">с ${winter.date_from}</span>` : ''}
                                ${winter.date_to ? `<span style="color: #8c8c8c; font-size: 12px;"> по ${winter.date_to}</span>` : ''}
                                <table style="font-size: 13px; margin-top: 8px;">
                                    <thead>
                                        <tr>
                                            <th>Тип дня</th>
                                            <th>Рейсы от НП</th>
                                            <th>Рейсы от КП</th>
                                            <th>Рейсы всего</th>
                                            <th>Пробег от НП</th>
                                            <th>Пробег от КП</th>
                                            <th>Пробег всего</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                        `;

                        for (let i = 1; i <= winter.num_of_types; i++) {
                            const typeName = winter[`type_${i}_name`] || `Тип ${i}`;
                            const fwdNum = winter[`type_${i}_forward_number`] || 0;
                            const revNum = winter[`type_${i}_reverse_number`] || 0;
                            const sumNum = winter[`type_${i}_sum_number`] || 0;

                            let fwdPrb = winter[`type_${i}_forward_probeg`];
                            let revPrb = winter[`type_${i}_reverse_probeg`];
                            let sumPrb = winter[`type_${i}_sum_probeg`];

                            if (fwdPrb === undefined || fwdPrb === null) {
                                fwdPrb = winterLengthFwd * fwdNum;
                            }
                            if (revPrb === undefined || revPrb === null) {
                                revPrb = winterLengthRev * revNum;
                            }
                            if (sumPrb === undefined || sumPrb === null) {
                                sumPrb = fwdPrb + revPrb;
                            }

                            html += `
                                <tr>
                                    <td>${typeName}</td>
                                    <td>${formatNumber(fwdNum)}</td>
                                    <td>${formatNumber(revNum)}</td>
                                    <td><strong>${formatNumber(sumNum)}</strong></td>
                                    <td>${formatNumber(fwdPrb)}</td>
                                    <td>${formatNumber(revPrb)}</td>
                                    <td><strong>${formatNumber(sumPrb)}</strong></td>
                                </tr>
                            `;
                        }

                        html += '</tbody></table></div>';
                    }

                    // Render summer period if exists
                    if (hasSummer) {
                        const summer = app.period_summer;
                        const summerLengthFwd = app.length_forward || 0;
                        const summerLengthRev = app.length_reverse || 0;

                        html += `
                            <div style="margin-top: 12px; padding: 12px; background: var(--bg-light); border-left: 3px solid #fa8c16; border-radius: 4px;">
                                <strong style="color: #fa8c16;"><i class="fas fa-sun"></i> Летний период</strong>
                                ${summer.date_from ? `<span style="color: #8c8c8c; font-size: 12px; margin-left: 8px;">с ${summer.date_from}</span>` : ''}
                                ${summer.date_to ? `<span style="color: #8c8c8c; font-size: 12px;"> по ${summer.date_to}</span>` : ''}
                                <table style="font-size: 13px; margin-top: 8px;">
                                    <thead>
                                        <tr>
                                            <th>Тип дня</th>
                                            <th>Рейсы от НП</th>
                                            <th>Рейсы от КП</th>
                                            <th>Рейсы всего</th>
                                            <th>Пробег от НП</th>
                                            <th>Пробег от КП</th>
                                            <th>Пробег всего</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                        `;

                        for (let i = 1; i <= summer.num_of_types; i++) {
                            const typeName = summer[`type_${i}_name`] || `Тип ${i}`;
                            const fwdNum = summer[`type_${i}_forward_number`] || 0;
                            const revNum = summer[`type_${i}_reverse_number`] || 0;
                            const sumNum = summer[`type_${i}_sum_number`] || 0;

                            let fwdPrb = summer[`type_${i}_forward_probeg`];
                            let revPrb = summer[`type_${i}_reverse_probeg`];
                            let sumPrb = summer[`type_${i}_sum_probeg`];

                            if (fwdPrb === undefined || fwdPrb === null) {
                                fwdPrb = summerLengthFwd * fwdNum;
                            }
                            if (revPrb === undefined || revPrb === null) {
                                revPrb = summerLengthRev * revNum;
                            }
                            if (sumPrb === undefined || sumPrb === null) {
                                sumPrb = fwdPrb + revPrb;
                            }

                            html += `
                                <tr>
                                    <td>${typeName}</td>
                                    <td>${formatNumber(fwdNum)}</td>
                                    <td>${formatNumber(revNum)}</td>
                                    <td><strong>${formatNumber(sumNum)}</strong></td>
                                    <td>${formatNumber(fwdPrb)}</td>
                                    <td>${formatNumber(revPrb)}</td>
                                    <td><strong>${formatNumber(sumPrb)}</strong></td>
                                </tr>
                            `;
                        }

                        html += '</tbody></table></div>';
                    }
                } else {
                    // Render regular table (no winter/summer periods)
                    html += `
                        <table style="font-size: 13px;">
                            <thead>
                                <tr>
                                    <th>Тип дня</th>
                                    <th>Рейсы от НП</th>
                                    <th>Рейсы от КП</th>
                                    <th>Рейсы всего</th>
                                    <th>Пробег от НП</th>
                                    <th>Пробег от КП</th>
                                    <th>Пробег всего</th>
                                </tr>
                            </thead>
                            <tbody>
                    `;

                    for (let i = 1; i <= numTypes; i++) {
                        const typeName = app[`type_${i}_name`] || `Тип ${i}`;
                        const fwdNum = app[`type_${i}_forward_number`] || 0;
                        const revNum = app[`type_${i}_reverse_number`] || 0;
                        const sumNum = app[`type_${i}_sum_number`] || 0;

                        // Пробег: берём из JSON или рассчитываем как length × number
                        let fwdPrb = app[`type_${i}_forward_probeg`];
                        let revPrb = app[`type_${i}_reverse_probeg`];
                        let sumPrb = app[`type_${i}_sum_probeg`];

                        if (fwdPrb === undefined || fwdPrb === null) {
                            fwdPrb = lengthFwd * fwdNum;
                        }
                        if (revPrb === undefined || revPrb === null) {
                            revPrb = lengthRev * revNum;
                        }
                        if (sumPrb === undefined || sumPrb === null) {
                            sumPrb = fwdPrb + revPrb;
                        }

                        html += `
                            <tr>
                                <td>${typeName}</td>
                                <td>${formatNumber(fwdNum)}</td>
                                <td>${formatNumber(revNum)}</td>
                                <td><strong>${formatNumber(sumNum)}</strong></td>
                                <td>${formatNumber(fwdPrb)}</td>
                                <td>${formatNumber(revPrb)}</td>
                                <td><strong>${formatNumber(sumPrb)}</strong></td>
                            </tr>
                        `;
                    }

                    html += '</tbody></table>';
                }
            }
            } catch (error) {
                console.error('Error rendering appendix:', key, error);
                html += `<div style="color: #ff4d4f; padding: 8px;">Ошибка при отображении приложения: ${error.message}</div>`;
            }

            html += '</div>';
        });

        html += '</div></div>';
    }

    return html;
}

    // Toggle collapsible sections
function toggleSection(header) {
    const section = header.closest('.collapsible-section');
    section.classList.toggle('collapsed');
}

// Download JSON
function downloadJson(data, filename = 'data.json') {
    const jsonString = JSON.stringify(data, null, 2);
    const blob = new Blob([jsonString], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function formatNumber(num) {
    if (num === null || num === undefined) return '-';
    if (typeof num === 'number') {
        return num.toLocaleString('ru-RU', { maximumFractionDigits: 2 });
    }
    return num;
}

// Stats
async function loadStats() {
    try {
        const response = await fetch('/api/stats');
        const stats = await response.json();

        document.getElementById('stat-contracts').textContent = stats.contracts || 0;
        document.getElementById('stat-agreements').textContent = stats.agreements || 0;
        document.getElementById('stat-checked').textContent = stats.agreements_checked || 0;

        // Load route params count
        const routesResponse = await fetch('/api/contracts');
        const contracts = await routesResponse.json();
        document.getElementById('stat-routes').textContent = contracts.length > 0 ? '✓' : '0';

    } catch (error) {
        console.error('Error loading stats:', error);
    }
}

// Agreements
let agreementsData = [];
let agrSortBy = null;
let agrSortOrder = 'asc';

function agrSort(field) {
    if (agrSortBy === field) {
        agrSortOrder = agrSortOrder === 'asc' ? 'desc' : 'asc';
    } else {
        agrSortBy = field;
        agrSortOrder = 'asc';
    }
    document.querySelectorAll('#tab-agreements th.sortable').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        th.querySelector('i').className = 'fas fa-sort';
    });
    const activeTh = document.querySelector('#tab-agreements th[data-sort="' + field + '"]');
    if (activeTh) {
        activeTh.classList.add(agrSortOrder === 'asc' ? 'sort-asc' : 'sort-desc');
        activeTh.querySelector('i').className = agrSortOrder === 'asc' ? 'fas fa-sort-up' : 'fas fa-sort-down';
    }
    renderAgreementsTable();
}

function sortAgreements(data) {
    if (!agrSortBy) return data;
    const sorted = [...data];
    sorted.sort((a, b) => {
        let va = a[agrSortBy] ?? '';
        let vb = b[agrSortBy] ?? '';
        if (typeof va === 'string') va = va.toLowerCase();
        if (typeof vb === 'string') vb = vb.toLowerCase();
        if (va < vb) return agrSortOrder === 'asc' ? -1 : 1;
        if (va > vb) return agrSortOrder === 'asc' ? 1 : -1;
        return 0;
    });
    return sorted;
}

function renderAgreementsTable() {
    const tbody = document.getElementById('agreements-table');
    const data = sortAgreements(agreementsData);

    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="empty-state">Нет загруженных соглашений</td></tr>';
        return;
    }

    tbody.innerHTML = data.map(a => {
        let typeBadge;
        if (a.ds_type === 'Высвобождение') {
            typeBadge = '<span class="badge" style="background:#fff0f6; color:#c41d7f; border:1px solid #ffadd2;" title="Высвобождение">В</span>';
        } else if (a.has_embedded_vysvobozhdenie) {
            typeBadge = '<span class="badge" style="background:#e6f7ff; color:#096dd9; border:1px solid #91d5ff;" title="Изменение параметров">ИП</span>'
                      + ' <span class="badge" style="background:#fff0f6; color:#c41d7f; border:1px solid #ffadd2;" title="Высвобождение">В</span>';
        } else {
            typeBadge = '<span class="badge" style="background:#e6f7ff; color:#096dd9; border:1px solid #91d5ff;" title="Изменение параметров">ИП</span>';
        }
        return `
        <tr>
            <td>${a.id}</td>
            <td>${a.contract_number || '-'}</td>
            <td>ДС №${a.ds_number}</td>
            <td>${typeBadge}</td>
            <td>
                <span class="badge ${a.status === 'checked' ? 'badge-success' : a.status === 'applied' ? 'badge-success' : 'badge-warning'}">
                    ${a.status === 'checked' ? 'Проверен' : a.status === 'applied' ? 'Применён' : 'Черновик'}
                </span>
            </td>
            <td>
                ${a.errors_count > 0 ? `<span class="badge badge-error" title="Ошибок: ${a.errors_count}">${a.errors_count} ош.</span>` : ''}
                ${a.warnings_count > 0 ? `<span class="badge badge-warning" title="Предупреждений: ${a.warnings_count}">${a.warnings_count} пред.</span>` : ''}
                ${a.has_seasonal_warnings ? `<span class="badge badge-warning" style="background:#fff7e6; color:#d48806; border:1px solid #ffd591;" title="Есть предупреждения по сезонным графикам"><i class="fas fa-snowflake"></i> граф.</span>` : ''}
                ${!a.errors_count && !a.warnings_count && !a.has_seasonal_warnings ? '-' : ''}
            </td>
            <td>${new Date(a.created_at).toLocaleDateString('ru-RU')}</td>
            <td>
                <button class="btn btn-default" onclick="viewAgreement(${a.id})" title="Просмотр">
                    <i class="fas fa-eye"></i>
                </button>
                ${a.status !== 'applied' ? `
                    <button class="btn btn-success" onclick="applyAgreement(${a.id})" title="Применить">
                        <i class="fas fa-check"></i>
                    </button>
                ` : `
                    <button class="btn btn-warning" onclick="unapplyAgreement(${a.id})" title="Отменить применение">
                        <i class="fas fa-undo"></i>
                    </button>
                `}
                <button class="btn btn-danger" onclick="deleteAgreement(${a.id}, ${a.status === 'applied'})" title="Удалить">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        </tr>`;
    }).join('');
}

async function loadAgreements() {
    const tbody = document.getElementById('agreements-table');

    // Заполняем список контрактов один раз при первом вызове
    const filterContract = document.getElementById('agr-filter-contract');
    if (filterContract && filterContract.options.length <= 1) {
        try {
            const r = await fetch('/api/contracts');
            const contracts = await r.json();
            contracts.forEach(c => {
                const opt = document.createElement('option');
                opt.value = c.number;
                opt.textContent = 'ГК ' + c.number;
                filterContract.appendChild(opt);
            });
        } catch (e) {}
    }

    try {
        const params = new URLSearchParams();
        const contract = document.getElementById('agr-filter-contract')?.value;
        const dsType = document.getElementById('agr-filter-ds-type')?.value;
        if (contract) params.set('contract_number', contract);
        if (dsType) params.set('ds_type', dsType);

        const response = await fetch('/api/agreements/?' + params.toString());
        if (!response.ok) throw new Error('HTTP ' + response.status);
        agreementsData = await response.json();
        renderAgreementsTable();

        // Загружаем историю высвобождений с тем же фильтром по контракту
        loadVysvobozhdeniHistory(contract || null);
    } catch (error) {
        tbody.innerHTML = `<tr><td colspan="8" class="empty-state">Ошибка загрузки: ${error.message}</td></tr>`;
    }
}

async function viewAgreement(id) {
    try {
        const response = await fetch(`/api/agreements/${id}`);
        const agreement = await response.json();

        currentViewedAgreement = agreement;

        const tablesHtml = agreement.json_data ? buildDataTables(agreement.json_data) : '';

        const validationErrors = agreement.check_errors || [];
        const validationWarnings = agreement.check_warnings || [];
        const seasonalWarnings = validationWarnings.filter(w =>
            w.includes('возможна ошибка в графиках') || w.includes('изменён сезонный график')
        );

        document.getElementById('modal-title').textContent = `ДС №${agreement.ds_number}`;
        document.getElementById('modal-body').innerHTML = `
            <div style="margin-bottom: 16px;">
                <strong>Контракт:</strong> ГК ${agreement.contract_number || '-'}<br>
                <strong>Статус:</strong> ${agreement.status}<br>
                <strong>Загружен:</strong> ${new Date(agreement.created_at).toLocaleString('ru-RU')}
            </div>

            ${validationErrors.length > 0 ? `
                <div class="alert alert-error">
                    <strong><i class="fas fa-exclamation-circle"></i> Ошибки валидации:</strong>
                    <ul>${validationErrors.map(e => `<li>${e}</li>`).join('')}</ul>
                </div>
            ` : ''}

            ${seasonalWarnings.length > 0 ? `
                <div class="alert alert-warning" style="background:#fffbe6; border-color:#ffe58f;">
                    <strong><i class="fas fa-snowflake" style="color:#d48806;"></i> Сезонные графики:</strong>
                    <ul style="margin-top:6px;">${seasonalWarnings.map(w => `<li>${w}</li>`).join('')}</ul>
                </div>
            ` : ''}

            ${tablesHtml}

            <div style="margin-top: 16px; display: flex; gap: 8px; flex-wrap: wrap;">
                <button class="btn btn-default btn-sm" onclick="downloadJson(currentViewedAgreement.json_data, 'ДС_${agreement.ds_number}.json')">
                    <i class="fas fa-download"></i> Скачать JSON с общей информацией
                </button>
                ${agreement.json_data?.km_data ? `
                <button class="btn btn-default btn-sm" onclick="downloadJson(currentViewedAgreement.json_data.km_data, 'КМ_Приложение${agreement.json_data.km_data.appendix_number || ''}_ДС${agreement.json_data.km_data.ds_number || agreement.ds_number}.json')">
                    <i class="fas fa-download"></i> Скачать JSON с расчётами
                </button>
                ` : ''}
            </div>
        `;

        openModal();
    } catch (error) {
        showAlert('error', 'Ошибка: ' + error.message);
    }
}

function getMonthName(month) {
    const names = ['', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                   'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    return names[month] || month;
}

async function applyAgreement(id) {
    if (!confirm('Применить ДС? Данные будут добавлены в базу и будут пересчитаны этапы.')) return;

    // Показываем оверлей загрузки
    showLoadingOverlay('Применение ДС и пересчёт этапов...');

    try {
        const response = await fetch(`/api/agreements/${id}/apply`, { method: 'POST' });
        const result = await response.json();

        hideLoadingOverlay();

        if (result.success) {
            // Если есть расхождения по км, показываем модальное окно
            if (result.km_discrepancies && result.km_discrepancies.length > 0) {
                showKmDiscrepanciesModal(result.message, result.km_discrepancies, result.seasonal_changes);
            } else if (result.seasonal_changes && result.seasonal_changes.length > 0) {
                showSeasonalChangesModal(result.message, result.seasonal_changes);
            } else {
                showAlert('success', result.message || 'ДС успешно применён');
            }
            loadAgreements();
        } else {
            showAlert('error', result.detail || 'Ошибка');
        }
    } catch (error) {
        hideLoadingOverlay();
        showAlert('error', 'Ошибка: ' + error.message);
    }
}

function showKmDiscrepanciesModal(message, discrepancies, seasonalChanges) {
    document.getElementById('modal-title').innerHTML = '<i class="fas fa-exclamation-triangle" style="color: var(--warning-color);"></i> Расхождения по километрам';
    document.getElementById('modal-body').innerHTML = `
        <div class="alert alert-success" style="margin-bottom: 16px;">
            <i class="fas fa-check"></i> ${message}
        </div>
        <div class="alert alert-warning">
            <strong><i class="fas fa-calculator"></i> Обнаружены расхождения между рассчитанными и указанными в Excel км:</strong>
            <p style="font-size: 13px; margin-top: 8px; color: var(--text-secondary);">
                Сравнение выполнено для периодов с марта 2026 года.
            </p>
        </div>
        <div class="table-container" style="margin-top: 16px;">
            <table>
                <thead>
                    <tr>
                        <th>Маршрут</th>
                        <th>Период</th>
                        <th style="text-align: right;">Рассчитано</th>
                        <th style="text-align: right;">В Excel</th>
                        <th style="text-align: right;">Разница</th>
                    </tr>
                </thead>
                <tbody>
                    ${discrepancies.map(d => {
                        // Парсим строку ошибки для извлечения данных
                        const match = d.match(/Маршрут ([^,]+), ([^:]+): рассчитано ([\d.]+), в Excel ([\d.]+), разница ([+-]?[\d.]+)/);
                        if (match) {
                            const diff = parseFloat(match[5]);
                            const diffClass = diff > 0 ? 'color: var(--error-color);' : 'color: var(--success-color);';
                            return `
                                <tr>
                                    <td><strong>${match[1]}</strong></td>
                                    <td>${match[2]}</td>
                                    <td style="text-align: right; font-family: monospace;">${match[3]}</td>
                                    <td style="text-align: right; font-family: monospace;">${match[4]}</td>
                                    <td style="text-align: right; font-family: monospace; ${diffClass}">${match[5]}</td>
                                </tr>
                            `;
                        }
                        return `<tr><td colspan="5">${d}</td></tr>`;
                    }).join('')}
                </tbody>
            </table>
        </div>
        ${seasonalChanges && seasonalChanges.length > 0 ? `
        <div style="margin-top:14px; background:#fffbe6; border:1px solid #ffe58f; border-radius:6px; padding:14px 16px;">
            <strong><i class="fas fa-snowflake" style="color:#d48806;"></i> Изменения сезонного графика:</strong>
            <ul style="margin-top:8px; margin-bottom:0; padding-left:20px;">
                ${seasonalChanges.map(c => `<li style="margin-bottom:4px;">${c}</li>`).join('')}
            </ul>
        </div>` : ''}
    `;
    document.getElementById('modal-footer').innerHTML = `
        <button class="btn btn-default" onclick="closeModal()">Закрыть</button>
    `;
    document.getElementById('modal').classList.add('active');
}

function showSeasonalChangesModal(message, seasonalChanges) {
    document.getElementById('modal-title').innerHTML = '<i class="fas fa-snowflake" style="color: #d48806;"></i> Изменения сезонного графика';
    document.getElementById('modal-body').innerHTML = `
        <div class="alert alert-success" style="margin-bottom: 16px;">
            <i class="fas fa-check"></i> ${message}
        </div>
        <div style="background:#fffbe6; border:1px solid #ffe58f; border-radius:6px; padding:14px 16px;">
            <strong><i class="fas fa-snowflake" style="color:#d48806;"></i> Обнаружены изменения сезонного графика:</strong>
            <ul style="margin-top:8px; margin-bottom:0; padding-left:20px;">
                ${seasonalChanges.map(c => `<li style="margin-bottom:4px;">${c}</li>`).join('')}
            </ul>
            <p style="margin-top:10px; margin-bottom:0; font-size:12px; color:#666;">
                Пересчёт выполнен с учётом новых параметров на указанные периоды.
                Проверьте результаты во вкладке «Проверка таблиц».
            </p>
        </div>
    `;
    document.getElementById('modal-footer').innerHTML = `
        <button class="btn btn-default" onclick="closeModal()">Закрыть</button>
    `;
    document.getElementById('modal').classList.add('active');
}

async function deleteAgreement(id, isApplied = false) {
    const msg = isApplied
        ? 'Удалить применённый ДС? Связанные параметры маршрутов тоже будут удалены.'
        : 'Удалить ДС? Это действие нельзя отменить.';
    if (!confirm(msg)) return;

    try {
        const url = isApplied ? `/api/agreements/${id}?force=true` : `/api/agreements/${id}`;
        const response = await fetch(url, { method: 'DELETE' });
        const result = await response.json();

        if (result.success) {
            showAlert('success', 'ДС удалён');
            loadAgreements();
            loadStats();
        } else {
            showAlert('error', result.detail || 'Ошибка');
        }
    } catch (error) {
        showAlert('error', 'Ошибка: ' + error.message);
    }
}

async function unapplyAgreement(id) {
    if (!confirm('Отменить применение ДС? Связанные параметры маршрутов будут удалены.')) return;

    showLoadingOverlay('Отмена применения ДС...');

    try {
        const response = await fetch(`/api/agreements/${id}/unapply`, { method: 'POST' });
        const result = await response.json();

        hideLoadingOverlay();

        if (result.success) {
            showAlert('success', result.message || 'Применение отменено');
            loadAgreements();
        } else {
            showAlert('error', result.detail || 'Ошибка');
        }
    } catch (error) {
        hideLoadingOverlay();
        showAlert('error', 'Ошибка: ' + error.message);
    }
}

// Calendar
const weekdays = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];

function initCalendar() {
    document.getElementById('calendar-month').addEventListener('change', loadCalendar);
    document.getElementById('calendar-year').addEventListener('change', loadCalendar);
    loadCalendar();
    loadSeasonPeriods();
}

// ============================================================
// Сезонность маршрутов (RouteSeasonPeriod — конкретные даты)
// ============================================================

let _seasonPeriodEditId = null; // null = новый, число = редактирование

function _fmtDate(iso) {
    // "2025-11-16" → "16.11.2025"
    const [y, m, d] = iso.split('-');
    return `${d}.${m}.${y}`;
}

async function loadSeasonPeriods() {
    const container = document.getElementById('season-periods-list');
    container.innerHTML = '<div class="loading"><div class="spinner"></div></div>';

    const route = document.getElementById('season-filter-route')?.value || '';
    const url = '/api/calendar/season-periods' + (route ? `?route=${encodeURIComponent(route)}` : '');

    try {
        const resp = await fetch(url);
        const periods = await resp.json();
        renderSeasonPeriods(periods);
    } catch (e) {
        container.innerHTML = `<div style="color:#cf1322;">Ошибка загрузки: ${e.message}</div>`;
    }
}

// Обновляет список маршрутов в фильтре
function _updateSeasonRouteFilter(periods) {
    const sel = document.getElementById('season-filter-route');
    if (!sel) return;
    const current = sel.value;
    const routes = [...new Set(periods.map(p => p.route))].sort();
    sel.innerHTML = '<option value="">Все маршруты</option>' +
        routes.map(r => `<option value="${r}" ${r === current ? 'selected' : ''}>${r}</option>`).join('');
}

function renderSeasonPeriods(periods) {
    _updateSeasonRouteFilter(periods);
    const container = document.getElementById('season-periods-list');
    if (!periods || periods.length === 0) {
        container.innerHTML = '<div style="color:var(--text-secondary);font-size:13px;padding:8px 0;">Нет записей.</div>';
        return;
    }

    // Группируем по году date_from
    const byYear = {};
    for (const p of periods) {
        const year = p.date_from.split('-')[0];
        if (!byYear[year]) byYear[year] = [];
        byYear[year].push(p);
    }
    const years = Object.keys(byYear).sort();

    const blocks = years.map(year => {
        const yp = byYear[year];
        yp.sort((a, b) => a.route.localeCompare(b.route, 'ru') || a.date_from.localeCompare(b.date_from));

        const cnt = yp.length;
        const cntLabel = cnt === 1 ? '1 период' : cnt < 5 ? `${cnt} периода` : `${cnt} периодов`;

        const rows = yp.map(p => {
            const isWinter = p.season === 'winter';
            const seasonLabel = isWinter
                ? '<i class="fas fa-snowflake" style="color:#1890ff;font-size:11px;"></i> Зима'
                : '<i class="fas fa-sun" style="color:#fa8c16;font-size:11px;"></i> Лето';
            return `
                <tr>
                    <td><strong>${p.route}</strong></td>
                    <td>${seasonLabel}</td>
                    <td>${_fmtDate(p.date_from)}</td>
                    <td>${_fmtDate(p.date_to)}</td>
                    <td style="white-space:nowrap;">
                        <button class="btn" style="font-size:12px;padding:2px 8px;"
                            onclick="seasonPeriodEdit(${p.id},'${p.route}','${p.season}','${p.date_from}','${p.date_to}')">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn" style="font-size:12px;padding:2px 8px;color:#cf1322;"
                            onclick="seasonPeriodDelete(${p.id})">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                </tr>`;
        }).join('');

        return `
            <div style="margin-bottom:4px;">
                <div onclick="seasonToggleYear('${year}')" style="
                    cursor:pointer; padding:8px 14px;
                    background:var(--bg-secondary,#f5f5f5);
                    border:1px solid var(--border-color,#e8e8e8);
                    border-radius:6px; display:flex; align-items:center; gap:10px;
                    user-select:none; font-size:13px;">
                    <i id="season-year-icon-${year}" class="fas fa-chevron-right"
                        style="font-size:10px; color:var(--text-secondary); transition:transform 0.15s;"></i>
                    <span style="font-weight:600; font-size:14px;">${year}</span>
                    <span style="color:var(--text-secondary); font-size:12px;">${cntLabel}</span>
                </div>
                <div id="season-year-content-${year}" style="display:none; margin-top:2px; border:1px solid var(--border-color,#e8e8e8); border-radius:0 0 6px 6px; overflow:hidden;">
                    <table class="table" style="font-size:13px; margin:0;">
                        <thead>
                            <tr>
                                <th>Маршрут</th><th>Сезон</th><th>Начало</th><th>Конец</th><th></th>
                            </tr>
                        </thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>
            </div>`;
    }).join('');

    container.innerHTML = blocks;
}

function seasonToggleYear(year) {
    const content = document.getElementById(`season-year-content-${year}`);
    const icon = document.getElementById(`season-year-icon-${year}`);
    if (!content) return;
    const open = content.style.display !== 'none';
    content.style.display = open ? 'none' : '';
    if (icon) icon.style.transform = open ? '' : 'rotate(90deg)';
}

function seasonPeriodShowAddForm() {
    _seasonPeriodEditId = null;
    document.getElementById('season-period-form-title').textContent = 'Новый период';
    document.getElementById('spf-route').value = '';
    document.getElementById('spf-route').disabled = false;
    document.getElementById('spf-season').value = 'winter';
    document.getElementById('spf-date-from').value = '';
    document.getElementById('spf-date-to').value = '';
    document.getElementById('season-period-form-error').style.display = 'none';
    document.getElementById('season-period-form').style.display = '';
    document.getElementById('spf-route').focus();
}

function seasonPeriodEdit(id, route, season, dateFrom, dateTo) {
    _seasonPeriodEditId = id;
    document.getElementById('season-period-form-title').textContent = `Редактировать период — ${route}`;
    document.getElementById('spf-route').value = route;
    document.getElementById('spf-route').disabled = true;
    document.getElementById('spf-season').value = season;
    document.getElementById('spf-date-from').value = dateFrom;
    document.getElementById('spf-date-to').value = dateTo;
    document.getElementById('season-period-form-error').style.display = 'none';
    document.getElementById('season-period-form').style.display = '';
    document.getElementById('spf-date-from').focus();
}

function seasonPeriodCancelForm() {
    document.getElementById('season-period-form').style.display = 'none';
    _seasonPeriodEditId = null;
}

async function seasonPeriodSaveForm() {
    const route = document.getElementById('spf-route').value.trim();
    const season = document.getElementById('spf-season').value;
    const dateFrom = document.getElementById('spf-date-from').value;
    const dateTo = document.getElementById('spf-date-to').value;
    const errEl = document.getElementById('season-period-form-error');

    if (!route || !dateFrom || !dateTo) {
        errEl.textContent = 'Заполните все поля';
        errEl.style.display = '';
        return;
    }

    const params = new URLSearchParams({ season, date_from: dateFrom, date_to: dateTo });
    let url, method;
    if (_seasonPeriodEditId !== null) {
        url = `/api/calendar/season-periods/${_seasonPeriodEditId}?${params}`;
        method = 'PUT';
    } else {
        params.set('route', route);
        url = `/api/calendar/season-periods?${params}`;
        method = 'POST';
    }

    try {
        const resp = await fetch(url, { method });
        const data = await resp.json();
        if (!resp.ok) {
            errEl.textContent = data.detail || 'Ошибка сохранения';
            errEl.style.display = '';
            return;
        }
        seasonPeriodCancelForm();
        await loadSeasonPeriods();
    } catch (e) {
        errEl.textContent = `Ошибка: ${e.message}`;
        errEl.style.display = '';
    }
}

async function seasonPeriodDelete(id) {
    if (!confirm('Удалить этот период?')) return;
    try {
        const resp = await fetch(`/api/calendar/season-periods/${id}`, { method: 'DELETE' });
        if (resp.ok) await loadSeasonPeriods();
    } catch (e) {
        alert(`Ошибка: ${e.message}`);
    }
}

async function loadCalendar() {
    const month = document.getElementById('calendar-month').value;
    const year = document.getElementById('calendar-year').value;

    const grid = document.getElementById('calendar-grid');
    grid.innerHTML = weekdays.map(d => `<div class="calendar-header">${d}</div>`).join('');

    try {
        const response = await fetch(`/api/calendar/?year=${year}&month=${month}`);
        const data = await response.json();

        // Calculate first day offset
        const firstDay = new Date(year, month - 1, 1);
        const startOffset = (firstDay.getDay() + 6) % 7; // Monday = 0

        // Add empty cells
        for (let i = 0; i < startOffset; i++) {
            grid.innerHTML += '<div class="calendar-day" style="visibility: hidden;"></div>';
        }

        // Add days
        const overrideMap = {};
        data.overrides.forEach(o => {
            if (!overrideMap[o.date]) overrideMap[o.date] = [];
            overrideMap[o.date].push(o);
        });

        data.days.forEach(day => {
            const classes = ['calendar-day'];
            const dayNum = parseInt(day.date.split('-')[2]);
            const weekday = day.weekday;

            if (weekday >= 6) classes.push('weekend');
            if (day.is_holiday) classes.push('holiday');
            if (overrideMap[day.date]) classes.push('override');

            const today = new Date().toISOString().split('T')[0];
            if (day.date === today) classes.push('today');

            grid.innerHTML += `<div class="${classes.join(' ')}" title="${day.note || ''}" data-date="${day.date}" onclick="openDayEditor('${day.date}', ${day.weekday}, ${day.is_holiday}, ${day.treat_as || 'null'})">${dayNum}</div>`;
        });

        // Load overrides list
        loadOverrides(data.overrides);

    } catch (error) {
        console.error('Error loading calendar:', error);
    }
}

function loadOverrides(overrides) {
    const container = document.getElementById('overrides-list');

    if (!overrides || overrides.length === 0) {
        container.innerHTML = '<div class="empty-state">Нет переопределений в этом месяце</div>';
        return;
    }

    // Group by date
    const grouped = {};
    overrides.forEach(o => {
        if (!grouped[o.date]) grouped[o.date] = [];
        grouped[o.date].push(o);
    });

    container.innerHTML = Object.entries(grouped).sort(([a],[b]) => a.localeCompare(b)).map(([date, items]) => `
        <div style="margin-bottom: 10px; padding: 10px 12px; background: #fafafa; border-radius: 6px;">
            <strong style="font-size: 13px;">${new Date(date).toLocaleDateString('ru-RU')}</strong>
            <div style="margin-top: 6px; font-size: 13px;">
                ${items.map(i => `
                    <div style="margin: 3px 0; display: flex; align-items: center; gap: 8px;">
                        <span style="flex: 1;">
                            <span style="color: #8c8c8c; font-size: 12px;">ГК ${i.contract_number}</span>
                            Маршрут <strong>${i.route}</strong> → <span class="badge badge-warning">${getDayName(i.treat_as)}</span>
                            ${i.source_agreement_id ? '' : `<span style="color: #52c41a; font-size: 11px; margin-left: 4px;">✎ ручное</span>`}
                        </span>
                        <button onclick="deleteOverride('${date}', '${i.route}', '${i.contract_number}')" title="Удалить" style="
                            background: none; border: none; cursor: pointer;
                            color: #ff4d4f; font-size: 16px; line-height: 1;
                            padding: 2px 4px; border-radius: 4px; flex-shrink: 0;
                        " onmouseover="this.style.background='#fff1f0'" onmouseout="this.style.background='none'">✕</button>
                    </div>
                `).join('')}
            </div>
        </div>
    `).join('');
}

async function deleteOverride(dateStr, route, contractNumber) {
    if (!confirm(`Удалить изменение для маршрута ${route} (ГК ${contractNumber}) на ${new Date(dateStr).toLocaleDateString('ru-RU')}?`)) return;

    const params = new URLSearchParams({ date_str: dateStr, contract_number: contractNumber, route });
    const response = await fetch(`/api/calendar/override?${params}`, { method: 'DELETE' });

    if (response.ok) {
        loadCalendar();
    } else {
        alert('Ошибка при удалении');
    }
}

async function showRouteOverrides() {
    const route = document.getElementById('override-route-filter').value.trim();
    if (!route) {
        showAlert('warning', 'Введите номер маршрута');
        return;
    }

    const response = await fetch(`/api/calendar/overrides-by-route?route=${encodeURIComponent(route)}`);
    if (!response.ok) { showAlert('error', 'Ошибка загрузки'); return; }
    const data = await response.json();

    document.getElementById('modal-title').textContent = `Перераспределения маршрута ${route}`;

    if (!data.overrides || data.overrides.length === 0) {
        document.getElementById('modal-body').innerHTML = `<div class="empty-state">Нет перераспределений для маршрута ${route}</div>`;
    } else {
        document.getElementById('modal-body').innerHTML = data.overrides.map(i => `
            <div style="display: flex; align-items: center; gap: 8px; padding: 6px 0; border-bottom: 1px solid #f0f0f0; font-size: 13px;">
                <span style="width: 100px; color: #595959;">${new Date(i.date).toLocaleDateString('ru-RU')}</span>
                <span style="color: #8c8c8c; width: 60px;">ГК ${i.contract_number}</span>
                <span>→ <span class="badge badge-warning">${getDayNameFull(i.treat_as)}</span></span>
                ${i.source_agreement_id ? '' : `<span style="color: #52c41a; font-size: 11px;">✎ ручное</span>`}
            </div>
        `).join('');
    }

    document.getElementById('modal-footer').innerHTML = `<button class="btn btn-default" onclick="closeModal()">Закрыть</button>`;
    openModal();
}

function getDayName(num) {
    const days = {1: 'Пн', 2: 'Вт', 3: 'Ср', 4: 'Чт', 5: 'Пт', 6: 'Сб', 7: 'Вс'};
    return days[num] || num;
}

function getDayNameFull(num) {
    const days = {
        1: 'Понедельник',
        2: 'Вторник',
        3: 'Среда',
        4: 'Четверг',
        5: 'Пятница',
        6: 'Суббота',
        7: 'Воскресенье'
    };
    return days[num] || `День ${num}`;
}

// Day editor
let currentEditDate = null;
let currentEditWeekday = null;

function openDayEditor(dateStr, weekday, isHoliday, treatAs) {
    currentEditDate = dateStr;
    currentEditWeekday = weekday;

    const displayDate = new Date(dateStr).toLocaleDateString('ru-RU', {
        weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
    });
    const currentDayType = treatAs || weekday;

    document.getElementById('modal-title').textContent = `Редактирование дня`;
    document.getElementById('modal-body').innerHTML = `
        <div style="margin-bottom: 16px;">
            <strong style="font-size: 15px;">${displayDate}</strong>
            <div style="margin-top: 6px; color: #8c8c8c; font-size: 13px;">
                Текущий тип: <span class="badge badge-default">${getDayNameFull(currentDayType)}</span>
                ${isHoliday ? '<span class="badge badge-error" style="margin-left: 8px;">Праздник</span>' : ''}
                ${treatAs ? '<span class="badge badge-warning" style="margin-left: 8px;">Переопределён</span>' : ''}
            </div>
        </div>

        <div class="form-group">
            <label class="form-label">Применить к</label>
            <select class="form-select" id="edit-scope" onchange="toggleRouteInput()">
                <option value="base">Всем маршрутам (базовый календарь)</option>
                <option value="contract">Конкретному ГК (все его маршруты)</option>
                <option value="route">Конкретному маршруту</option>
            </select>
        </div>

        <div id="contract-input-group" style="display: none;">
            <div class="form-group">
                <label class="form-label">ГК</label>
                <select class="form-select" id="edit-contract">
                    <option value="">— Загрузка... —</option>
                </select>
            </div>
        </div>

        <div id="route-input-group" style="display: none;">
            <div class="form-group">
                <label class="form-label">Маршрут</label>
                <input type="text" class="form-input" id="edit-route" placeholder="Например: 120">
                <div style="margin-top: 4px; font-size: 12px; color: #8c8c8c;">ГК определится автоматически</div>
            </div>
        </div>

        <div class="form-group">
            <label class="form-label">Считать как</label>
            <select class="form-select" id="edit-treat-as">
                <option value="">— Не менять —</option>
                <option value="1">Понедельник (рабочий)</option>
                <option value="2">Вторник (рабочий)</option>
                <option value="3">Среда (рабочий)</option>
                <option value="4">Четверг (рабочий)</option>
                <option value="5">Пятница (рабочий)</option>
                <option value="6">Суббота (выходной)</option>
                <option value="7">Воскресенье (выходной)</option>
            </select>
        </div>

        <div class="form-group" id="holiday-group">
            <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                <input type="checkbox" id="edit-is-holiday" ${isHoliday ? 'checked' : ''}>
                Праздничный день
            </label>
        </div>

        <div class="form-group">
            <label class="form-label">Примечание</label>
            <input type="text" class="form-input" id="edit-note" placeholder="Например: По производственному календарю">
        </div>
    `;

    // Загружаем список ГК
    fetch('/api/contracts').then(r => r.json()).then(contracts => {
        const sel = document.getElementById('edit-contract');
        if (!sel) return;
        sel.innerHTML = contracts.map(c => `<option value="${c.number}">ГК ${c.number}</option>`).join('');
    }).catch(() => {});

    document.getElementById('modal-footer').innerHTML = `
        <button class="btn btn-default" onclick="closeModal()">Отмена</button>
        <button class="btn btn-primary" onclick="saveDayEdit()">
            <i class="fas fa-save"></i> Сохранить
        </button>
    `;

    openModal();
}

function toggleRouteInput() {
    const scope = document.getElementById('edit-scope').value;
    document.getElementById('contract-input-group').style.display = scope === 'contract' ? 'block' : 'none';
    document.getElementById('route-input-group').style.display = scope === 'route' ? 'block' : 'none';
    document.getElementById('holiday-group').style.display = scope === 'base' ? 'block' : 'none';
}

async function saveDayEdit() {
    const scope = document.getElementById('edit-scope').value;
    const treatAs = document.getElementById('edit-treat-as').value;
    const note = document.getElementById('edit-note').value;

    if (!treatAs) {
        showAlert('warning', 'Выберите тип дня');
        return;
    }

    try {
        if (scope === 'base') {
            // Базовый календарь
            const isHoliday = document.getElementById('edit-is-holiday').checked;

            const params = new URLSearchParams({
                treat_as: treatAs,
                is_holiday: isHoliday,
                note: note
            });

            const response = await fetch(`/api/calendar/base/${currentEditDate}?${params}`, {
                method: 'POST'
            });

            const result = await response.json();
            if (result.success) {
                showAlert('success', 'День обновлён для всех маршрутов');
                closeModal();
                loadCalendar();
            } else {
                showAlert('error', result.detail || 'Ошибка сохранения');
            }
        } else if (scope === 'contract') {
            // Переопределение для всего ГК
            const contract = document.getElementById('edit-contract').value;
            if (!contract) { showAlert('warning', 'Выберите ГК'); return; }

            const params = new URLSearchParams({
                date_str: currentEditDate, contract_number: contract,
                treat_as: treatAs, source_text: note || 'Ручное изменение'
            });
            const response = await fetch(`/api/calendar/override-by-contract?${params}`, { method: 'POST' });
            const result = await response.json();
            if (result.success) {
                showAlert('success', `День обновлён для всех маршрутов ГК ${contract} (${result.routes_updated} маршрутов)`);
                closeModal();
                loadCalendar();
            } else {
                showAlert('error', result.detail || 'Ошибка сохранения');
            }
        } else {
            // Переопределение для конкретного маршрута (ГК определяется автоматически)
            const route = document.getElementById('edit-route').value.trim();
            if (!route) { showAlert('warning', 'Укажите номер маршрута'); return; }

            const params = new URLSearchParams({
                date_str: currentEditDate, route: route,
                treat_as: treatAs, source_text: note || 'Ручное изменение'
            });
            const response = await fetch(`/api/calendar/override-by-route?${params}`, { method: 'POST' });
            const result = await response.json();
            if (result.success) {
                showAlert('success', `День обновлён для маршрута ${route}`);
                closeModal();
                loadCalendar();
            } else {
                showAlert('error', result.detail || 'Ошибка сохранения');
            }
        }
    } catch (error) {
        showAlert('error', 'Ошибка: ' + error.message);
    }
}

function formatDateRange(dateFrom, dateTo) {
    const from = new Date(dateFrom).toLocaleDateString('ru-RU', {day: '2-digit', month: '2-digit'});
    const to = new Date(dateTo).toLocaleDateString('ru-RU', {day: '2-digit', month: '2-digit'});
    if (dateFrom === dateTo) return from;
    return `${from} — ${to}`;
}

// Calculations
async function calculateKm() {
    const route = document.getElementById('calc-route').value;
    const dateFrom = document.getElementById('calc-date-from').value;
    const dateTo = document.getElementById('calc-date-to').value;

    if (!route || !dateFrom || !dateTo) {
        showAlert('warning', 'Заполните все поля');
        return;
    }

    const resultCard = document.getElementById('calc-result');
    const resultContent = document.getElementById('calc-result-content');

    resultCard.style.display = 'block';
    resultContent.innerHTML = '<div class="loading"><div class="spinner"></div></div>';

    try {
        const response = await fetch(`/api/calculations/route-auto/${route}?date_from=${dateFrom}&date_to=${dateTo}`);
        const data = await response.json();

        if (data.detail) {
            resultContent.innerHTML = `<div class="alert alert-error">${data.detail}</div>`;
            return;
        }

        // Сегменты — источники данных для расчёта
        const segmentsHtml = data.segments && data.segments.length > 0 ? `
            <div style="margin-bottom: 24px; padding: 16px; background: var(--bg-light); border-radius: 8px; border-left: 4px solid #1890ff;">
                <h4 style="margin: 0 0 12px 0;"><i class="fas fa-database"></i> Источники данных для расчёта</h4>
                <table style="font-size: 13px;">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Период</th>
                            <th>Источник</th>
                            <th>Протяж. НП</th>
                            <th>Протяж. КП</th>
                            <th>Дней</th>
                            <th>Рейсов</th>
                            <th>Км</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${data.segments.map(s => `
                            <tr>
                                <td><code>${s.route_params_id || '-'}</code></td>
                                <td>${formatDateRange(s.date_from, s.date_to)}</td>
                                <td>${s.source || '<span style="color: #8c8c8c;">—</span>'}</td>
                                <td>${s.length_forward}</td>
                                <td>${s.length_reverse}</td>
                                <td>${s.days_count}</td>
                                <td>${s.forward_trips + s.reverse_trips}</td>
                                <td><strong>${s.total_km.toLocaleString('ru-RU')}</strong></td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        ` : '';

        // Календарные переопределения из ДС
        const overridesHtml = data.calendar_overrides && data.calendar_overrides.length > 0 ? `
            <div style="margin-bottom: 24px; padding: 16px; background: var(--bg-light); border-radius: 8px; border-left: 4px solid #faad14;">
                <h4 style="margin: 0 0 12px 0;"><i class="fas fa-calendar-alt" style="color: #faad14;"></i> Переопределения календаря</h4>
                <table style="font-size: 13px;">
                    <thead>
                        <tr>
                            <th>Источник</th>
                            <th>Период</th>
                            <th>Дней</th>
                            <th>Считать как</th>
                            <th>Основание</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${data.calendar_overrides.map(o => `
                            <tr>
                                <td>${o.ds_number ? `<span class="badge badge-warning">ДС${o.ds_number}</span>` : '<span style="color: #8c8c8c;">вручную</span>'}</td>
                                <td>${o.dates}</td>
                                <td>${o.count}</td>
                                <td><span class="badge badge-default">${getDayNameFull(o.treat_as)}</span></td>
                                <td style="font-size: 12px; color: var(--text-secondary); max-width: 400px;">${o.source_text || '—'}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        ` : '';

        const fmtRub = v => v != null ? v.toLocaleString('ru-RU', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '—';
        const rubCards = data.price_available ? `
                <div class="stat-card">
                    <div class="stat-label">Всего руб.</div>
                    <div class="stat-value info">${fmtRub(data.total_rub)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Руб. от НП</div>
                    <div class="stat-value">${fmtRub(data.forward_rub)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Руб. от КП</div>
                    <div class="stat-value">${fmtRub(data.reverse_rub)}</div>
                </div>` : `
                <div class="stat-card" style="grid-column: span 3;">
                    <div class="stat-label">Стоимость, руб.</div>
                    <div class="stat-value" style="font-size:13px; color: var(--text-secondary);">Нет данных внешней БД</div>
                </div>`;

        resultContent.innerHTML = `
            <div class="stats-grid" style="margin-bottom: 24px;">
                <div class="stat-card">
                    <div class="stat-label">Всего км</div>
                    <div class="stat-value info">${data.total_km.toLocaleString('ru-RU')}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Км от НП</div>
                    <div class="stat-value">${data.total_forward_km.toLocaleString('ru-RU')}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Км от КП</div>
                    <div class="stat-value">${data.total_reverse_km.toLocaleString('ru-RU')}</div>
                </div>
                ${rubCards}
            </div>

            ${segmentsHtml}

            ${overridesHtml}

            <h4 style="margin-bottom: 12px;">По типам дней:</h4>
            <table>
                <thead>
                    <tr>
                        <th>Тип дня</th>
                        <th>Дней</th>
                        <th>Рейсов от НП</th>
                        <th>Рейсов от КП</th>
                        <th>Км от НП</th>
                        <th>Км от КП</th>
                    </tr>
                </thead>
                <tbody>
                    ${Object.entries(data.by_day_type).map(([type, d]) => `
                        <tr>
                            <td>${getDayName(parseInt(type))}</td>
                            <td>${d.days_count}</td>
                            <td>${d.forward_trips}</td>
                            <td>${d.reverse_trips}</td>
                            <td>${d.forward_km.toFixed(2)}</td>
                            <td>${d.reverse_km.toFixed(2)}</td>
                        </tr>
                    `).join('')}
                    ${(() => {
                        const vals = Object.values(data.by_day_type);
                        const sum = (fn) => vals.reduce((acc, d) => acc + fn(d), 0);
                        return `<tr style="font-weight:600; background:var(--bg-secondary);">
                            <td>Итого</td>
                            <td>${sum(d => d.days_count)}</td>
                            <td>${sum(d => d.forward_trips)}</td>
                            <td>${sum(d => d.reverse_trips)}</td>
                            <td>${sum(d => d.forward_km).toFixed(2)}</td>
                            <td>${sum(d => d.reverse_km).toFixed(2)}</td>
                        </tr>`;
                    })()}
                </tbody>
            </table>
        `;

    } catch (error) {
        resultContent.innerHTML = `<div class="alert alert-error">Ошибка: ${error.message}</div>`;
    }
}

// Modal
function openModal() {
    document.getElementById('modal').classList.add('active');
}

function closeModal() {
    document.getElementById('modal').classList.remove('active');
}

// Alerts
function showAlert(type, message) {
    const alert = document.createElement('div');
    alert.className = `alert alert-${type}`;
    alert.style.cssText = 'position: fixed; top: 80px; right: 24px; z-index: 1000; max-width: 400px; animation: slideIn 0.3s ease;';
    alert.innerHTML = `<i class="fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'times-circle' : 'exclamation-triangle'}"></i> ${message}`;

    document.body.appendChild(alert);

    setTimeout(() => {
        alert.style.opacity = '0';
        setTimeout(() => alert.remove(), 300);
    }, 4000);
}

// Export volumes
async function loadExportRoutes() {
    const contract = document.getElementById('export-contract').value;
    const infoDiv = document.getElementById('export-routes-info');
    const countSpan = document.getElementById('export-routes-count');
    const listDiv = document.getElementById('export-routes-list');

    infoDiv.style.display = 'block';
    listDiv.innerHTML = '<div class="spinner" style="width: 20px; height: 20px;"></div>';

    try {
        const response = await fetch(`/api/calculations/routes/${contract}`);
        const data = await response.json();

        if (data.detail) {
            listDiv.innerHTML = `<span style="color: var(--error-color);">${data.detail}</span>`;
            return;
        }

        countSpan.textContent = data.routes_count;
        listDiv.textContent = data.routes.join(', ');

    } catch (error) {
        listDiv.innerHTML = `<span style="color: var(--error-color);">Ошибка: ${error.message}</span>`;
    }
}

function _getExportParams() {
    return {
        contract: document.getElementById('export-contract').value,
        startMonth: document.getElementById('export-start-month').value,
        startYear: document.getElementById('export-start-year').value,
        endMonth: document.getElementById('export-end-month').value,
        endYear: document.getElementById('export-end-year').value,
    };
}

async function _downloadExportFile(url, fallbackFilename, btnId, btnLabel) {
    const btn = document.getElementById(btnId);
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Формирование...';

    try {
        const response = await fetch(url);

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || 'Ошибка экспорта');
        }

        const blob = await response.blob();
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = fallbackFilename;

        if (contentDisposition) {
            const match = contentDisposition.match(/filename\*=UTF-8''(.+)/);
            if (match) {
                filename = decodeURIComponent(match[1]);
            }
        }

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        showAlert('success', 'Файл успешно сформирован');

    } catch (error) {
        showAlert('error', error.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = btnLabel;
    }
}

async function exportMonthlyVolumes() {
    const p = _getExportParams();
    const url = `/api/calculations/export/${p.contract}?start_year=${p.startYear}&start_month=${p.startMonth}&end_year=${p.endYear}&end_month=${p.endMonth}`;
    await _downloadExportFile(url, `Объёмы_ГК${p.contract}.xlsx`, 'export-btn-km', '<i class="fas fa-download"></i> Выгрузить км');
}

async function exportRubVolumes() {
    const p = _getExportParams();
    const url = `/api/calculations/export-rub/${p.contract}?start_year=${p.startYear}&start_month=${p.startMonth}&end_year=${p.endYear}&end_month=${p.endMonth}`;
    await _downloadExportFile(url, `Объёмы_руб_ГК${p.contract}.xlsx`, 'export-btn-rub', '<i class="fas fa-ruble-sign"></i> Выгрузить руб.');
}

async function exportCombinedVolumes() {
    const p = _getExportParams();
    const url = `/api/calculations/export-combined/${p.contract}?start_year=${p.startYear}&start_month=${p.startMonth}&end_year=${p.endYear}&end_month=${p.endMonth}`;
    await _downloadExportFile(url, `Объёмы_км_и_руб_ГК${p.contract}.xlsx`, 'export-btn-combined', '<i class="fas fa-file-excel"></i> Выгрузить км и руб.');
}

// Manual edit functions
function startManualEdit() {
    const data = lastUploadResult && lastUploadResult.data;
    if (!data) return;

    const general = data.general || {};
    const appendices = data.appendices || {};

    let formHtml = '';

    // General section
    formHtml += `
        <div style="margin-bottom: 20px; padding: 12px; background: var(--bg-secondary); border-radius: 6px;">
            <h4 style="margin: 0 0 12px 0;"><i class="fas fa-info-circle"></i> Общие данные</h4>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
                <div class="form-group" style="margin: 0;">
                    <label class="form-label">Сумма</label>
                    <input type="text" class="form-input" data-path="general.sum_text" data-dtype="text" value="${general.sum_text || ''}">
                </div>
                <div class="form-group" style="margin: 0;">
                    <label class="form-label">Пробег, км</label>
                    <input type="number" step="0.01" class="form-input" data-path="general.probeg_sravnenie" data-dtype="number" value="${general.probeg_sravnenie || ''}">
                </div>
            </div>
        </div>
    `;

    // Appendices
    for (const [key, app] of Object.entries(appendices)) {
        if (!app) continue;

        const hasWinter = app.period_winter && app.period_winter.num_of_types > 0;
        const hasSummer = app.period_summer && app.period_summer.num_of_types > 0;

        formHtml += `
            <div style="margin-bottom: 20px; padding: 12px; background: var(--bg-secondary); border-radius: 6px;">
                <h4 style="margin: 0 0 12px 0;">
                    <i class="fas fa-file-alt"></i>
                    Приложение ${app.appendix_num || key}
                    <span class="badge badge-default" style="margin-left: 8px;">${app.route || '-'}</span>
                </h4>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 12px;">
                    <div class="form-group" style="margin: 0;">
                        <label class="form-label">Протяжённость от НП, км</label>
                        <input type="number" step="0.001" class="form-input"
                            data-path="appendices.${key}.length_forward" data-dtype="number"
                            value="${app.length_forward || ''}">
                    </div>
                    <div class="form-group" style="margin: 0;">
                        <label class="form-label">Протяжённость от КП, км</label>
                        <input type="number" step="0.001" class="form-input"
                            data-path="appendices.${key}.length_reverse" data-dtype="number"
                            value="${app.length_reverse || ''}">
                    </div>
                </div>
        `;

        if (hasWinter || hasSummer) {
            if (hasWinter) {
                formHtml += buildEditPeriodHtml(key, 'period_winter', app.period_winter, 'Зимний период');
            }
            if (hasSummer) {
                formHtml += buildEditPeriodHtml(key, 'period_summer', app.period_summer, 'Летний период');
            }
        } else {
            const numTypes = app.num_of_types || 0;
            if (numTypes > 0) {
                formHtml += buildEditTypesHtml(`appendices.${key}`, app, numTypes);
            }
        }

        formHtml += '</div>';
    }

    // Change blocks
    const changeBlocks = [
        { key: 'change_with_money_no_appendix',    label: 'Изменения без приложений (с оплатой)',    hasAppendix: false },
        { key: 'change_without_money_no_appendix', label: 'Изменения без приложений (без оплаты)',   hasAppendix: false },
        { key: 'change_with_money',                label: 'Изменения с приложениями (с оплатой)',    hasAppendix: true  },
        { key: 'change_without_money',             label: 'Изменения с приложениями (без оплаты)',   hasAppendix: true  },
    ];

    for (const block of changeBlocks) {
        const items = data[block.key];
        if (!items || items.length === 0) continue;

        formHtml += `
            <div style="margin-bottom: 20px; padding: 12px; background: var(--bg-secondary); border-radius: 6px;">
                <h4 style="margin: 0 0 12px 0;"><i class="fas fa-exchange-alt"></i> ${block.label}</h4>
        `;

        items.forEach((item, idx) => {
            const base = `${block.key}.${idx}`;
            formHtml += `
                <div style="border: 1px solid var(--border-color); border-radius: 4px; padding: 10px; margin-bottom: 10px;">
                    <div style="font-size: 12px; color: var(--text-secondary); margin-bottom: 8px;">Запись ${idx + 1}</div>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px;">
                        <div class="form-group" style="margin: 0;">
                            <label class="form-label" style="font-size: 12px;">Маршрут</label>
                            <input type="text" class="form-input" style="font-size: 12px; padding: 4px 6px;"
                                data-path="${base}.route" data-dtype="text" value="${item.route || ''}">
                        </div>
                        <div class="form-group" style="margin: 0;">
                            <label class="form-label" style="font-size: 12px;">Тип дня</label>
                            <input type="text" class="form-input" style="font-size: 12px; padding: 4px 6px;"
                                data-path="${base}.day_type" data-dtype="text" value="${item.day_type || ''}">
                        </div>
                        <div class="form-group" style="margin: 0;">
                            <label class="form-label" style="font-size: 12px;">Дата с</label>
                            <input type="text" class="form-input" style="font-size: 12px; padding: 4px 6px;"
                                data-path="${base}.date_from" data-dtype="text" value="${item.date_from || ''}">
                        </div>
                        <div class="form-group" style="margin: 0;">
                            <label class="form-label" style="font-size: 12px;">Дата по</label>
                            <input type="text" class="form-input" style="font-size: 12px; padding: 4px 6px;"
                                data-path="${base}.date_to" data-dtype="text" value="${item.date_to || ''}">
                        </div>
                        <div class="form-group" style="margin: 0;">
                            <label class="form-label" style="font-size: 12px;">Дата на</label>
                            <input type="text" class="form-input" style="font-size: 12px; padding: 4px 6px;"
                                data-path="${base}.date_on" data-dtype="text" value="${item.date_on || ''}">
                        </div>
                        ${block.hasAppendix ? `
                        <div class="form-group" style="margin: 0;">
                            <label class="form-label" style="font-size: 12px;">Приложение</label>
                            <input type="text" class="form-input" style="font-size: 12px; padding: 4px 6px;"
                                data-path="${base}.appendix" data-dtype="text" value="${item.appendix || ''}">
                        </div>
                        ` : ''}
                        <div class="form-group" style="margin: 0; grid-column: 1 / -1;">
                            <label class="form-label" style="font-size: 12px;">Пункт</label>
                            <input type="text" class="form-input" style="font-size: 12px; padding: 4px 6px;"
                                data-path="${base}.point" data-dtype="text" value="${item.point || ''}">
                        </div>
                    </div>
                </div>
            `;
        });

        formHtml += '</div>';
    }

    document.getElementById('modal-title').textContent = 'Ручное редактирование данных ДС';
    document.getElementById('modal-body').style.maxHeight = '65vh';
    document.getElementById('modal-body').style.overflowY = 'auto';
    document.getElementById('modal-body').innerHTML = formHtml;
    document.getElementById('modal-footer').innerHTML = `
        <button class="btn btn-default" onclick="closeModal()">Отмена</button>
        <button class="btn btn-primary" onclick="saveManualEdit()">
            <i class="fas fa-save"></i> Сохранить
        </button>
    `;
    openModal();
}

function buildEditPeriodHtml(appKey, periodKey, periodData, label) {
    const numTypes = periodData.num_of_types || 0;
    if (numTypes === 0) return '';
    return `
        <div style="margin-top: 10px;">
            <strong style="font-size: 13px;">${label}</strong>
            ${buildEditTypesHtml(`appendices.${appKey}.${periodKey}`, periodData, numTypes)}
        </div>
    `;
}

function buildEditTypesHtml(basePath, data, numTypes) {
    let html = `
        <table style="font-size: 13px; margin-top: 8px; width: 100%;">
            <thead>
                <tr>
                    <th>Тип дня</th>
                    <th style="width: 90px;">Рейсы НП</th>
                    <th style="width: 90px;">Рейсы КП</th>
                    <th style="width: 90px;">Всего</th>
                </tr>
            </thead>
            <tbody>
    `;
    for (let i = 1; i <= numTypes; i++) {
        const typeName = data[`type_${i}_name`] || '';
        const fwdNum = data[`type_${i}_forward_number`] ?? 0;
        const revNum = data[`type_${i}_reverse_number`] ?? 0;
        const sumNum = data[`type_${i}_sum_number`] ?? 0;
        html += `
            <tr>
                <td>
                    <input type="text" class="form-input" style="width: 100%; font-size: 12px; padding: 4px 6px;"
                        data-path="${basePath}.type_${i}_name" data-dtype="text" value="${typeName}">
                </td>
                <td>
                    <input type="number" class="form-input" style="width: 80px; font-size: 12px; padding: 4px 6px;"
                        data-path="${basePath}.type_${i}_forward_number" data-dtype="number" value="${fwdNum}" min="0">
                </td>
                <td>
                    <input type="number" class="form-input" style="width: 80px; font-size: 12px; padding: 4px 6px;"
                        data-path="${basePath}.type_${i}_reverse_number" data-dtype="number" value="${revNum}" min="0">
                </td>
                <td>
                    <input type="number" class="form-input" style="width: 80px; font-size: 12px; padding: 4px 6px;"
                        data-path="${basePath}.type_${i}_sum_number" data-dtype="number" value="${sumNum}" min="0">
                </td>
            </tr>
        `;
    }
    html += '</tbody></table>';
    return html;
}

async function saveManualEdit() {
    if (!lastUploadResult || !lastAgreementId) {
        showAlert('error', 'Нет данных для сохранения');
        closeModal();
        return;
    }

    const updatedData = JSON.parse(JSON.stringify(lastUploadResult.data));

    document.querySelectorAll('#modal-body [data-path]').forEach(input => {
        const path = input.getAttribute('data-path');
        const dtype = input.getAttribute('data-dtype');
        const raw = input.value;
        const value = dtype === 'number' ? (raw === '' ? null : parseFloat(raw)) : raw;
        setDeepValue(updatedData, path, value);
    });

    try {
        const response = await fetch(`/api/agreements/${lastAgreementId}/json`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(updatedData)
        });

        if (response.ok) {
            lastUploadResult.data = updatedData;
            closeModal();
            // Re-render the result with updated data
            const container = document.getElementById('upload-result');
            const tablesHtml = buildDataTables(updatedData);
            // Just update the tables section without full re-render
            showAlert('success', 'Данные успешно сохранены');
            // Trigger full re-display
            displayUploadResult(lastUploadResult);
        } else {
            const err = await response.json();
            showAlert('error', err.detail || 'Ошибка сохранения');
        }
    } catch (error) {
        showAlert('error', 'Ошибка: ' + error.message);
    }
}

function setDeepValue(obj, path, value) {
    const parts = path.split('.');
    let current = obj;
    for (let i = 0; i < parts.length - 1; i++) {
        if (current === null || current === undefined || current[parts[i]] === undefined) return;
        current = current[parts[i]];
    }
    current[parts[parts.length - 1]] = value;
}

// Auto-load routes when contract changes
document.getElementById('export-contract')?.addEventListener('change', () => {
    document.getElementById('export-routes-info').style.display = 'none';
});

// =============================================================================
// Вкладка «Проверка таблиц»
// =============================================================================


async function tcLoadContracts() {
    const sel = document.getElementById('tc-contract');
    try {
        const res = await fetch('/api/contracts');
        const list = await res.json();
        list.sort((a, b) => a.number.localeCompare(b.number, undefined, { numeric: true }));
        list.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.number;
            opt.textContent = `ГК ${c.number}`;
            sel.appendChild(opt);
        });
    } catch (e) {
        console.error('tcLoadContracts error', e);
    }
}

const STATUS_LABEL = { checked: 'Проверен', applied: 'Применён', draft: 'Черновик' };

async function tcLoadAgreements() {
    const contractFilter = document.getElementById('tc-contract').value;
    const sel = document.getElementById('tc-agreement');
    sel.innerHTML = '<option value="">— выберите ДС —</option>';

    if (!contractFilter) return;

    try {
        const res = await fetch(`/api/table-checks/agreements?contract_number=${contractFilter}`);
        const list = await res.json();

        const checked = list.filter(a => a.status === 'checked');
        const applied = list.filter(a => a.status === 'applied');

        if (checked.length > 0) {
            const grp = document.createElement('optgroup');
            grp.label = 'Требуют проверки';
            checked.forEach(a => {
                const opt = document.createElement('option');
                opt.value = a.id;
                opt.textContent = `ДС №${a.number} — ${STATUS_LABEL[a.status] ?? a.status}`;
                grp.appendChild(opt);
            });
            sel.appendChild(grp);
        }

        if (applied.length > 0) {
            const grp = document.createElement('optgroup');
            grp.label = 'Применённые';
            applied.forEach(a => {
                const opt = document.createElement('option');
                opt.value = a.id;
                opt.textContent = `ДС №${a.number} — ${STATUS_LABEL[a.status] ?? a.status}`;
                grp.appendChild(opt);
            });
            sel.appendChild(grp);
        }

        // Приоритет автовыбора: сначала checked (ждут проверки), иначе первый applied
        const priority = checked.length > 0 ? checked : applied;
        if (priority.length > 0) {
            sel.value = priority[0].id;
        }
    } catch (e) {
        console.error('tcLoadAgreements error', e);
    }
}

async function tcRunChecks() {
    const agreementId = document.getElementById('tc-agreement').value;
    if (!agreementId) {
        showAlert('error', 'Выберите дополнительное соглашение');
        return;
    }

    const btn = document.getElementById('tc-run-btn');
    document.getElementById('tc-empty').style.display = 'none';
    document.getElementById('tc-results').style.display = 'none';
    document.getElementById('tc-prev-banner').style.display = 'none';
    document.getElementById('tc-reference-block').style.display = 'none';
    document.getElementById('tc-loading').style.display = '';
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Проверяю…';

    try {
        const [res, seasonalRes] = await Promise.all([
            fetch(`/api/table-checks/${agreementId}`),
            fetch(`/api/table-checks/agreement/${agreementId}/seasonal`),
        ]);
        if (!res.ok) {
            const err = await res.json();
            showAlert('error', err.detail || 'Ошибка загрузки данных');
            document.getElementById('tc-empty').style.display = '';
            return;
        }
        const data = await res.json();
        const seasonalData = seasonalRes.ok ? await seasonalRes.json() : { rows: [] };
        tcRender(data, seasonalData);
    } catch (e) {
        showAlert('error', 'Ошибка: ' + e.message);
        document.getElementById('tc-empty').style.display = '';
    } finally {
        document.getElementById('tc-loading').style.display = 'none';
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-search"></i> Проверить';
    }
}

function tcRender(data, seasonalData) {
    // --- Баннер предыдущего ДС ---
    const banner = document.getElementById('tc-prev-banner');
    const refBlock = document.getElementById('tc-reference-block');

    if (data.prev_source === 'agreement') {
        banner.innerHTML = `
            <div style="padding: 10px 16px; background: #f6ffed; border: 1px solid #b7eb8f; border-radius: 6px; font-size: 13px;">
                <i class="fas fa-check-circle" style="color: #52c41a;"></i>
                Сравнение с <strong>ДС №${data.prev_info.number}</strong> (найден в системе, статус: ${data.prev_info.status})
            </div>`;
        banner.style.display = '';
        refBlock.style.display = 'none';
    } else if (data.prev_source === 'reference') {
        banner.innerHTML = `
            <div style="padding: 10px 16px; background: #fffbe6; border: 1px solid #ffe58f; border-radius: 6px; font-size: 13px;">
                <i class="fas fa-database" style="color: #faad14;"></i>
                Сравнение с <strong>эталонными данными ДС №${data.prev_info.number}</strong>
                ${data.prev_info.note ? `<span style="color: #666;"> — ${data.prev_info.note}</span>` : ''}
            </div>`;
        banner.style.display = '';
        refBlock.style.display = 'none';
        tcLoadExistingRefs(data.agreement.contract_number);
    } else {
        banner.innerHTML = `
            <div style="padding: 10px 16px; background: #fff2f0; border: 1px solid #ffccc7; border-radius: 6px; font-size: 13px;">
                <i class="fas fa-exclamation-circle" style="color: #ff4d4f;"></i>
                Предыдущий ДС <strong>не найден</strong> в системе. Проверки 1 и 3 недоступны.
            </div>`;
        banner.style.display = '';

        // Показываем форму ввода эталона
        const contractNumber = data.agreement.contract_number;
        document.getElementById('tc-ref-ds-number').value = '';
        document.getElementById('tc-ref-initial-km').value = '';
        document.getElementById('tc-ref-probeg').value = '';
        document.getElementById('tc-ref-note').value = '';
        document.getElementById('tc-ref-save-status').textContent = '';
        // Сохраняем контракт для последующего сохранения
        refBlock.dataset.contractNumber = contractNumber;
        refBlock.dataset.agreementId = data.agreement.id;
        refBlock.style.display = '';
        tcLoadExistingRefs(contractNumber);
    }

    // Вспомогательная функция: форматирует подпись «Приложение №N к ДС №M»
    function tcAppendixLabel(t) {
        if (!t) return '';
        const parts = [];
        if (t.appendix_number) parts.push(`Приложение №${t.appendix_number}`);
        if (t.ds_number) parts.push(`к ДС №${t.ds_number}`);
        return parts.join(' ');
    }

    // --- Таблица 1: Расчет изменения объема ---
    document.getElementById('tc-appendix-raschet').textContent = tcAppendixLabel(data.tables.table_raschet_izm_objema);
    tcRenderTable(
        data.tables.table_raschet_izm_objema.headers,
        data.tables.table_raschet_izm_objema.rows,
        'tc-table-raschet-head',
        'tc-table-raschet-body',
    );
    tcRenderChecks(data.tables.table_raschet_izm_objema.checks, 'tc-checks-raschet');

    // --- Таблица 2: Этапы сроки (со встроенными проверками) ---
    document.getElementById('tc-appendix-etapy-sroki').textContent = tcAppendixLabel(data.tables.table_etapy_sroki);
    tcRenderEtapySroki(data.tables.table_etapy_sroki);

    // --- Таблица 3: Финансирование ---
    document.getElementById('tc-appendix-finansirovanie').textContent = tcAppendixLabel(data.tables.table_finansirovanie_po_godam);
    tcRenderTable(
        data.tables.table_finansirovanie_po_godam.headers,
        data.tables.table_finansirovanie_po_godam.rows,
        'tc-table-finansirovanie-head',
        'tc-table-finansirovanie-body',
    );
    tcRenderChecks(data.tables.table_finansirovanie_po_godam.checks, 'tc-checks-finansirovanie');

    // --- Таблица 4: Этапы с авансами ---
    document.getElementById('tc-appendix-avans').textContent = tcAppendixLabel(data.tables.table_etapy_avans);
    tcRenderTableAvans(
        data.tables.table_etapy_avans.headers,
        data.tables.table_etapy_avans.rows,
        'tc-table-avans-head',
        'tc-table-avans-body',
    );
    tcRenderChecks(data.tables.table_etapy_avans.checks, 'tc-checks-avans');

    // --- Блок 5: Проверка объемов по маршрутам ---
    const kmAppendix = data.km_appendix_number ? `Приложение №${data.km_appendix_number}` : '';
    document.getElementById('tc-appendix-km-routes').textContent = kmAppendix;
    tcRenderKmRoutes(data.km_by_routes, data.km_routes_not_applicable, data.km_total_vs_probeg);

    // --- Блок 6: Сезонные графики ---
    tcRenderSeasonal(seasonalData || { rows: [] });

    document.getElementById('tc-results').style.display = '';
}

function tcRenderSeasonal(data) {
    const card = document.getElementById('tc-seasonal-card');
    const content = document.getElementById('tc-seasonal-content');
    const rows = data.rows || [];

    card.style.display = '';

    if (rows.length === 0) {
        content.innerHTML = `<div style="padding:12px 0; color:var(--text-secondary); font-size:13px;">
            <i class="fas fa-info-circle"></i> В контракте нет маршрутов с сезонными графиками (207, 305, 1КР, 2КР, 3КР).
        </div>`;
        return;
    }

    const tableRows = rows.map(r => {
        if (r.no_ds_changes) {
            // Маршрут есть в контракте, но в этом ДС изменений не было
            const winterRef = r.winter ? r.winter.ref_range : '—';
            const summerRef = r.summer ? r.summer.ref_range : '—';
            return `
                <tr style="color:var(--text-secondary);">
                    <td><strong>${r.route}</strong></td>
                    <td>—</td>
                    <td style="font-style:italic;">—</td>
                    <td>${winterRef}</td>
                    <td style="font-style:italic;">—</td>
                    <td>${summerRef}</td>
                    <td style="font-style:italic; color:var(--text-secondary);">нет изменений в этом ДС</td>
                </tr>
            `;
        }

        if (r.no_period_data) {
            // Маршрут есть в приложении ДС, но без указания конкретных дат периодов
            return `
                <tr>
                    <td><strong>${r.route}</strong></td>
                    <td>${r.appendix}</td>
                    <td colspan="4" style="color:var(--text-secondary); font-style:italic;">
                        нет изменений в датах сезонных графиков
                    </td>
                    <td><span style="color:#52c41a;"><i class="fas fa-check"></i></span></td>
                </tr>
            `;
        }

        const winterDs = r.winter ? r.winter.ds_range : '—';
        const winterRef = r.winter ? (r.winter.ref_range || 'не задан') : '—';
        const winterMismatch = r.winter && r.winter.mismatch;

        const summerDs = r.summer ? r.summer.ds_range : '—';
        const summerRef = r.summer ? (r.summer.ref_range || 'не задан') : '—';
        const summerMismatch = r.summer && r.summer.mismatch;

        const mismatchStyle = 'background: #fff2f0; color: #cf1322; font-weight: 500;';

        let changeHtml = '';
        if (r.seasonal_change) {
            changeHtml = `<span style="display:inline-block; padding:2px 8px; background:#fffbe6; border:1px solid #ffe58f; border-radius:4px; font-size:12px; color:#d48806;">
                <i class="fas fa-exclamation-triangle"></i>
                изменён сезонный график за ${r.seasonal_change.year} год:
                ${r.seasonal_change.date_from} – ${r.seasonal_change.date_to}
            </span>`;
        }

        return `
            <tr>
                <td><strong>${r.route}</strong></td>
                <td>${r.appendix}</td>
                <td style="${winterMismatch ? mismatchStyle : ''}">${winterDs}${winterMismatch ? ' ⚠' : ''}</td>
                <td>${winterRef}</td>
                <td style="${summerMismatch ? mismatchStyle : ''}">${summerDs}${summerMismatch ? ' ⚠' : ''}</td>
                <td>${summerRef}</td>
                <td>${changeHtml || '<span style="color:#52c41a;"><i class="fas fa-check"></i></span>'}</td>
            </tr>
        `;
    }).join('');

    const hasMismatches = rows.some(r => (r.winter && r.winter.mismatch) || (r.summer && r.summer.mismatch));
    const hasChanges = rows.some(r => r.seasonal_change);

    let summaryHtml = '';
    if (hasMismatches) {
        summaryHtml += `<div style="padding:10px 14px; background:#fff7e6; border:1px solid #ffd591; border-radius:6px; margin-bottom:12px; font-size:13px;">
            <i class="fas fa-exclamation-circle" style="color:#d48806;"></i>
            <strong>Возможна ошибка в графиках</strong> — периоды в ДС отличаются от эталонных значений (ячейки выделены красным).
        </div>`;
    }
    if (hasChanges) {
        summaryHtml += `<div style="padding:10px 14px; background:#fffbe6; border:1px solid #ffe58f; border-radius:6px; margin-bottom:12px; font-size:13px;">
            <i class="fas fa-info-circle" style="color:#d48806;"></i>
            <strong>Изменение сезонного графика</strong> — параметры вступают в силу в середине сезонного периода. Пересчёт выполнен на остаток сезона.
        </div>`;
    }

    content.innerHTML = `
        ${summaryHtml}
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Маршрут</th>
                        <th>Приложение</th>
                        <th>Зима (ДС)</th>
                        <th>Зима (эталон)</th>
                        <th>Лето (ДС)</th>
                        <th>Лето (эталон)</th>
                        <th>Статус</th>
                    </tr>
                </thead>
                <tbody>${tableRows}</tbody>
            </table>
        </div>
    `;
}

function tcRenderTable(headers, rows, headId, bodyId) {
    const thead = document.getElementById(headId);
    const tbody = document.getElementById(bodyId);

    // Добавляем класс tc-table к родительскому <table>
    const table = thead.closest('table');
    if (table) table.classList.add('tc-table');

    thead.innerHTML = headers.map(h => `<th>${h}</th>`).join('');
    tbody.innerHTML = rows.map(row => {
        const isTotal = row.some(cell => String(cell).toUpperCase() === 'ИТОГО');
        const rowStyle = isTotal ? ' style="font-weight: 600; background: var(--bg-secondary);"' : '';
        return `<tr${rowStyle}>${row.map(cell => `<td>${cell ?? ''}</td>`).join('')}</tr>`;
    }).join('');
}

// Рендер таблицы "Этапы с авансами" с объединением ячеек в столбце rub6 (индекс 6)
// для строк с одинаковым номером этапа и одинаковым значением
function tcRenderTableAvans(headers, rows, headId, bodyId) {
    const thead = document.getElementById(headId);
    const tbody = document.getElementById(bodyId);
    const table = thead.closest('table');
    if (table) table.classList.add('tc-table');

    thead.innerHTML = headers.map(h => `<th>${h}</th>`).join('');

    const MERGE_COL = 6;
    const rowspans = new Array(rows.length).fill(1);
    const skipCell = new Array(rows.length).fill(false);

    for (let i = 0; i < rows.length; i++) {
        if (skipCell[i]) continue;
        const stage = String(rows[i][0] ?? '').trim();
        if (!/^\d+$/.test(stage)) continue; // только строки с номером этапа
        const val = String(rows[i][MERGE_COL] ?? '').trim();
        if (!val) continue;

        let span = 1;
        for (let j = i + 1; j < rows.length; j++) {
            if (String(rows[j][0] ?? '').trim() === stage && String(rows[j][MERGE_COL] ?? '').trim() === val) {
                span++;
                skipCell[j] = true;
            } else {
                break;
            }
        }
        rowspans[i] = span;
    }

    tbody.innerHTML = rows.map((row, i) => {
        const isTotal = row.some(cell => String(cell).toUpperCase() === 'ИТОГО');
        const rowStyle = isTotal ? ' style="font-weight: 600; background: var(--bg-secondary);"' : '';
        const cells = row.map((cell, colIdx) => {
            if (colIdx === MERGE_COL) {
                if (skipCell[i]) return '';
                const rsAttr = rowspans[i] > 1 ? ` rowspan="${rowspans[i]}"` : '';
                return `<td${rsAttr} style="vertical-align: middle;">${cell ?? ''}</td>`;
            }
            return `<td>${cell ?? ''}</td>`;
        }).join('');
        return `<tr${rowStyle}>${cells}</tr>`;
    }).join('');
}

function tcCheckBadge(check) {
    if (!check) return '<td></td>';
    let icon, color, title = '';
    if (check.ok === true) {
        icon = 'fa-check-circle'; color = '#52c41a';
        title = check.expected_fmt ? `Ожидалось: ${check.expected_fmt}` : '';
    } else if (check.ok === false) {
        icon = 'fa-times-circle'; color = '#ff4d4f';
        title = check.message || (check.expected_fmt ? `Ожидалось: ${check.expected_fmt}, фактически: ${check.actual_fmt}` : '');
    } else {
        icon = 'fa-question-circle'; color = '#faad14';
        title = check.message || '';
    }
    return `<td style="text-align:center; white-space:nowrap;" title="${title}">
        <i class="fas ${icon}" style="color:${color}; font-size:14px;"></i>
        ${check.ok === false && check.message ? `<span style="font-size:11px; color:#ff4d4f; display:block;">${check.message}</span>` : ''}
    </td>`;
}

function tcRenderEtapySroki(tableData) {
    const thead = document.getElementById('tc-table-etapy-sroki-head');
    const tbody = document.getElementById('tc-table-etapy-sroki-body');
    const table = thead.closest('table');
    if (table) table.classList.add('tc-table');

    const rowChecks = tableData.row_checks || [];
    const itogo_check = tableData.itogo_check;

    // Заголовки + 2 колонки проверок
    thead.innerHTML = tableData.headers.map(h => `<th>${h}</th>`).join('')
        + '<th style="text-align:center; white-space:nowrap;">Пров. км</th>'
        + '<th style="text-align:center; white-space:nowrap;">Пров. руб.</th>';

    tbody.innerHTML = tableData.rows.map((row, idx) => {
        const isTotal = row.some(cell => String(cell).toUpperCase() === 'ИТОГО');
        const rowStyle = isTotal ? ' style="font-weight: 600; background: var(--bg-secondary);"' : '';
        const cells = row.map(cell => `<td>${cell ?? ''}</td>`).join('');

        if (isTotal) {
            // Для строки ИТОГО — показываем арифметическую проверку
            if (itogo_check) {
                return `<tr${rowStyle}>${cells}${tcCheckBadge(itogo_check.km)}${tcCheckBadge(itogo_check.price)}</tr>`;
            }
            return `<tr${rowStyle}>${cells}<td></td><td></td></tr>`;
        }

        const rc = rowChecks[idx];
        if (!rc) return `<tr${rowStyle}>${cells}<td></td><td></td></tr>`;

        // Добавляем пометку «Закрыт» или «Расчётный» для ячеек
        const kmBadge = tcCheckBadge(rc.km);
        const priceBadge = tcCheckBadge(rc.price);

        const closedMark = rc.closed
            ? ' style="background: rgba(0,0,0,0.02);"'
            : '';

        return `<tr${closedMark}>${cells}${kmBadge}${priceBadge}</tr>`;
    }).join('');

    // Блок с пояснением цветов
    const legend = document.getElementById('tc-etapy-sroki-legend');
    if (legend) {
        legend.innerHTML = `
            <span style="font-size:12px; color:var(--text-secondary);">
                <i class="fas fa-check-circle" style="color:#52c41a;"></i> Совпадает &nbsp;
                <i class="fas fa-times-circle" style="color:#ff4d4f;"></i> Расхождение &nbsp;
                <i class="fas fa-question-circle" style="color:#faad14;"></i> Нет данных &nbsp;
                ${!tableData.price_available ? '<i class="fas fa-exclamation-triangle" style="color:#faad14;"></i> Внешняя БД недоступна — проверка руб. пропущена' : ''}
            </span>`;
    }
}

function tcRenderChecks(checks, containerId) {
    const container = document.getElementById(containerId);
    if (!checks || checks.length === 0) {
        container.innerHTML = '';
        return;
    }

    const html = checks.map(c => {
        let icon, color, bg, border;
        if (c.ok === true) {
            icon = 'fa-check-circle'; color = '#52c41a'; bg = '#f6ffed'; border = '#b7eb8f';
        } else if (c.ok === false) {
            icon = 'fa-times-circle'; color = '#ff4d4f'; bg = '#fff2f0'; border = '#ffccc7';
        } else {
            icon = 'fa-question-circle'; color = '#faad14'; bg = '#fffbe6'; border = '#ffe58f';
        }

        let detail = '';
        if (c.ok === false && c.expected !== null && c.actual !== null) {
            detail = `<span style="color: #666; margin-left: 8px; font-size: 12px;">
                Ожидалось: <strong>${c.expected_fmt ?? c.expected}</strong>,
                фактически: <strong>${c.actual_fmt ?? c.actual}</strong>
            </span>`;
        } else if (c.ok === false && c.message) {
            detail = `<span style="color: #666; margin-left: 8px; font-size: 12px;">${c.message}</span>`;
        } else if (c.ok === null && c.message) {
            detail = `<span style="color: #666; margin-left: 8px; font-size: 12px;">${c.message}</span>`;
        } else if (c.ok === true && c.expected !== null) {
            detail = `<span style="color: #8c8c8c; margin-left: 8px; font-size: 12px;">${c.expected_fmt ?? c.expected}</span>`;
        }

        return `
            <div style="display: flex; align-items: center; padding: 8px 12px; margin-bottom: 6px;
                        background: ${bg}; border: 1px solid ${border}; border-radius: 6px; font-size: 13px;">
                <i class="fas ${icon}" style="color: ${color}; margin-right: 10px; font-size: 15px; flex-shrink: 0;"></i>
                <span>${c.name}</span>
                ${detail}
            </div>`;
    }).join('');

    container.innerHTML = html;
}

function tcRenderKmRoutes(checks, notApplicable, totalChecks) {
    const container = document.getElementById('tc-checks-km-routes');
    if (notApplicable) {
        container.innerHTML = '<div style="display:flex; align-items:center; gap:8px; padding: 10px 14px; background:#f0f5ff; border:1px solid #adc6ff; border-radius:6px; font-size:13px; color:#2f54eb;"><i class="fas fa-info-circle"></i> Проверка объёмов по маршрутам для ГК252 не требуется — данные о пробеге хранятся в таблице «Объёмы работ» основного документа ДС.</div>';
        return;
    }

    const fmtKm = v => Number(v).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    // --- Блок: сумма КМ vs probeg ---
    let totalHtml = '';
    if (totalChecks && totalChecks.length > 0) {
        const rows = totalChecks.map(t => {
            let icon, color;
            if (t.ok === true)       { icon = 'fa-check-circle'; color = '#52c41a'; }
            else if (t.ok === false) { icon = 'fa-times-circle'; color = '#ff4d4f'; }
            else                     { icon = 'fa-question-circle'; color = '#faad14'; }

            const valStr = t.expected != null
                ? `${fmtKm(t.km_total)} км = ${fmtKm(t.expected)} км`
                : `${fmtKm(t.km_total)} км — значение отсутствует`;

            const errStr = t.ok === false
                ? `<span style="color:#ff4d4f; font-size:12px; margin-left:8px;">${t.message}</span>`
                : '';

            return `<tr>
                <td style="padding:4px 8px; font-size:13px;">${t.label}</td>
                <td style="padding:4px 8px; font-size:13px; font-variant-numeric:tabular-nums;">${valStr}${errStr}</td>
                <td style="padding:4px 8px; text-align:center;">
                    <i class="fas ${icon}" style="color:${color};"></i>
                </td>
            </tr>`;
        }).join('');

        totalHtml = `<div style="margin-bottom:12px; border:1px solid #d9d9d9; border-radius:6px; overflow:hidden;">
            <div style="padding:6px 10px; background:var(--bg-secondary); font-size:12px; font-weight:600; color:var(--text-secondary); border-bottom:1px solid #d9d9d9;">
                Сумма КМ по всем периодам vs значения пробега из ДС
            </div>
            <table style="width:100%; border-collapse:collapse;">
                <tbody>${rows}</tbody>
            </table>
        </div>`;
    }

    if (!checks || checks.length === 0) {
        container.innerHTML = totalHtml + '<div style="color: var(--text-secondary); font-size: 13px; padding: 8px 0;">Нет данных КМ по маршрутам</div>';
        return;
    }

    const html = checks.map(c => {
        let icon, color, bg, border;
        if (c.ok === true) {
            icon = 'fa-check-circle'; color = '#52c41a'; bg = '#f6ffed'; border = '#b7eb8f';
        } else if (c.ok === false) {
            icon = 'fa-times-circle'; color = '#ff4d4f'; bg = '#fff2f0'; border = '#ffccc7';
        } else {
            icon = 'fa-question-circle'; color = '#faad14'; bg = '#fffbe6'; border = '#ffe58f';
        }

        const sourceLabel = c.period_label
            ? (c.closed
                ? '<span style="font-size:11px; background:#f0f0f0; padding:1px 5px; border-radius:3px; margin-left:6px; color:#555;">закрыт</span>'
                : '<span style="font-size:11px; background:#e6f4ff; padding:1px 5px; border-radius:3px; margin-left:6px; color:#1677ff;">расчётный</span>')
            : '';

        const periodStr = c.period_label ? `Период ${c.period_label}` : 'Проверка КМ по маршрутам';

        const fmtKm = v => v.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

        let statusMsg = '';
        if (c.ok === true && c.closed) {
            statusMsg = `<span style="color: #52c41a; margin-left: 8px; font-size: 12px;">общее значение км совпадает</span>`;
        } else if (c.ok === true && !c.closed) {
            statusMsg = `<span style="color: #52c41a; margin-left: 8px; font-size: 12px;">все маршруты совпадают</span>`;
        } else if (c.ok === false && c.closed && c.message) {
            statusMsg = `<span style="color: #ff4d4f; margin-left: 8px; font-size: 12px;">${c.message}</span>`;
        } else if (c.ok === null && c.message) {
            statusMsg = `<span style="color: #666; margin-left: 8px; font-size: 12px;">${c.message}</span>`;
        }

        let errorsHtml = '';
        if (!c.closed && c.errors && c.errors.length > 0) {
            errorsHtml = '<div style="margin-top: 6px; padding-left: 25px;">'
                + c.errors.map(e => {
                    const diffStr = e.diff > 0 ? `+${fmtKm(e.diff)}` : fmtKm(e.diff);
                    return `<div style="font-size: 12px; color: #ff4d4f; margin-bottom: 2px;">
                        Маршрут <strong>${e.route}</strong>: в ДС — ${fmtKm(e.json_km)} км,
                        ожидалось — ${fmtKm(e.expected_km)} км
                        (разница: ${diffStr} км)
                    </div>`;
                }).join('')
                + '</div>';
        }

        let totalCheckHtml = '';
        if (c.total_check) {
            const tc = c.total_check;
            let tcIcon, tcColor, tcText;
            if (tc.ok === true) {
                tcIcon = 'fa-check-circle'; tcColor = '#52c41a';
                tcText = `сумма маршрутов = total (${fmtKm(tc.json_total)} км)`;
            } else if (tc.ok === false) {
                tcIcon = 'fa-times-circle'; tcColor = '#ff4d4f';
                const diffStr = tc.diff > 0 ? `+${fmtKm(tc.diff)}` : fmtKm(tc.diff);
                tcText = `сумма маршрутов (${fmtKm(tc.sum_routes)}) ≠ total (${fmtKm(tc.json_total)}), разница: ${diffStr} км`;
            } else {
                tcIcon = 'fa-question-circle'; tcColor = '#faad14';
                tcText = tc.message || 'нет данных для проверки суммы';
            }
            totalCheckHtml = `<div style="margin-top: 5px; padding-left: 25px; font-size: 12px; color: ${tcColor};">
                <i class="fas ${tcIcon}" style="color: ${tcColor}; margin-right: 4px;"></i>${tcText}
            </div>`;
        }

        return `
            <div style="padding: 8px 12px; margin-bottom: 6px; background: ${bg}; border: 1px solid ${border}; border-radius: 6px; font-size: 13px;">
                <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 4px;">
                    <i class="fas ${icon}" style="color: ${color}; margin-right: 6px; font-size: 15px; flex-shrink: 0;"></i>
                    <span>${periodStr}</span>
                    ${sourceLabel}
                    ${statusMsg}
                </div>
                ${totalCheckHtml}
                ${errorsHtml}
            </div>`;
    }).join('');

    container.innerHTML = totalHtml + html;
}

function tcToggleSection(sectionId) {
    const section = document.getElementById(sectionId);
    const icon = document.getElementById(sectionId + '-icon');
    if (!section) return;
    const hidden = section.style.display === 'none';
    section.style.display = hidden ? '' : 'none';
    if (icon) {
        icon.classList.toggle('fa-chevron-down', !hidden);
        icon.classList.toggle('fa-chevron-up', hidden);
    }
}

async function tcSaveReference() {
    const refBlock = document.getElementById('tc-reference-block');
    const contractNumber = refBlock.dataset.contractNumber;
    const agreementId = refBlock.dataset.agreementId;

    const dsNumber = document.getElementById('tc-ref-ds-number').value.trim();
    const initialKm = parseFloat(document.getElementById('tc-ref-initial-km').value);
    const probeg = parseFloat(document.getElementById('tc-ref-probeg').value);
    const note = document.getElementById('tc-ref-note').value.trim();

    if (!dsNumber) { showAlert('error', 'Укажите номер предыдущего ДС'); return; }
    if (isNaN(initialKm) && isNaN(probeg)) { showAlert('error', 'Укажите хотя бы одно значение km'); return; }

    const statusEl = document.getElementById('tc-ref-save-status');
    statusEl.textContent = 'Сохраняю...';

    try {
        const res = await fetch('/api/table-checks/references', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contract_number: contractNumber,
                reference_ds_number: dsNumber,
                initial_km: isNaN(initialKm) ? null : initialKm,
                probeg_etapy: isNaN(probeg) ? null : probeg,
                note: note || null,
            }),
        });

        if (res.ok) {
            statusEl.textContent = 'Сохранено';
            // Перезапускаем проверку
            await tcRunChecks();
        } else {
            const err = await res.json();
            statusEl.textContent = '';
            showAlert('error', err.detail || 'Ошибка сохранения');
        }
    } catch (e) {
        statusEl.textContent = '';
        showAlert('error', 'Ошибка: ' + e.message);
    }
}

async function tcLoadExistingRefs(contractNumber) {
    const container = document.getElementById('tc-existing-refs');
    if (!container) return;
    try {
        const res = await fetch(`/api/table-checks/references?contract_number=${contractNumber}`);
        const refs = await res.json();
        if (refs.length === 0) { container.innerHTML = ''; return; }

        container.innerHTML = `
            <div style="font-size: 12px; font-weight: 600; color: var(--text-secondary); margin-bottom: 8px;">
                Сохранённые эталоны для ГК ${contractNumber}:
            </div>
            ${refs.map(r => `
                <div style="display: flex; align-items: center; gap: 12px; padding: 6px 10px;
                            background: var(--bg-secondary); border-radius: 4px; margin-bottom: 4px; font-size: 12px;">
                    <span>ДС №${r.reference_ds_number}</span>
                    ${r.initial_km != null ? `<span>Объём: ${r.initial_km.toLocaleString('ru-RU', {maximumFractionDigits: 2})} км</span>` : ''}
                    ${r.probeg_etapy != null ? `<span>Пробег: ${r.probeg_etapy.toLocaleString('ru-RU', {maximumFractionDigits: 2})} км</span>` : ''}
                    ${r.note ? `<span style="color: var(--text-secondary);">${r.note}</span>` : ''}
                    <button class="btn btn-sm" style="margin-left: auto; color: #ff4d4f; border-color: #ffccc7;"
                            onclick="tcDeleteRef(${r.id}, '${contractNumber}')">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>`).join('')}`;
    } catch (e) {
        console.error('tcLoadExistingRefs error', e);
    }
}

async function tcDeleteRef(refId, contractNumber) {
    if (!confirm('Удалить эталонную запись?')) return;
    try {
        const res = await fetch(`/api/table-checks/references/${refId}`, { method: 'DELETE' });
        if (res.ok) {
            tcLoadExistingRefs(contractNumber);
        } else {
            showAlert('error', 'Ошибка удаления');
        }
    } catch (e) {
        showAlert('error', 'Ошибка: ' + e.message);
    }
}

// ─── История высвобождений ───────────────────────────────────────────────────

let vysvHistoryVisible = true;

function toggleVysvHistory() {
    vysvHistoryVisible = !vysvHistoryVisible;
    document.getElementById('vysv-history-body').style.display = vysvHistoryVisible ? '' : 'none';
    const chevron = document.getElementById('vysv-history-chevron');
    chevron.style.transform = vysvHistoryVisible ? '' : 'rotate(-90deg)';
}

async function loadVysvobozhdeniHistory(contractNumber) {
    const tbody = document.getElementById('vysv-history-table');
    if (!tbody) return;

    try {
        const params = new URLSearchParams();
        if (contractNumber) params.set('contract_number', contractNumber);
        const res = await fetch('/api/agreements/vysvobozhdenie-history?' + params.toString());
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const data = await res.json();
        renderVysvobozhdeniHistory(data);
    } catch (e) {
        tbody.innerHTML = `<tr><td colspan="8" class="empty-state">Ошибка загрузки: ${e.message}</td></tr>`;
    }
}

function renderVysvobozhdeniHistory(data) {
    const tbody = document.getElementById('vysv-history-table');
    const countEl = document.getElementById('vysv-history-count');

    if (countEl) {
        countEl.textContent = data.length > 0 ? `(${data.length})` : '';
    }

    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="empty-state">Нет данных о высвобождениях</td></tr>';
        return;
    }

    const fmt = (num) => num != null
        ? Number(num).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₽'
        : '—';

    tbody.innerHTML = data.map(v => {
        const typeBadge = v.is_standalone
            ? '<span class="badge" style="background:#fff0f6; color:#c41d7f; border:1px solid #ffadd2;">Отдельный ДС</span>'
            : '<span class="badge" style="background:#fff7e6; color:#d46b08; border:1px solid #ffd591;"><i class="fas fa-unlock-alt"></i> В архиве ДС</span>';
        return `<tr style="cursor:pointer;" onclick="viewAgreement(${v.agreement_id})" title="Открыть ДС №${v.ds_number}">
            <td>${v.contract_number || '—'}</td>
            <td>ДС №${v.ds_number}</td>
            <td>${typeBadge}</td>
            <td style="text-align:center;">${v.closed_stage != null ? v.closed_stage + ' этап' : '—'}</td>
            <td style="text-align:right; font-variant-numeric: tabular-nums;">${fmt(v.closed_amount)}</td>
            <td style="text-align:right; font-variant-numeric: tabular-nums;">${fmt(v.new_contract_price)}</td>
            <td>${new Date(v.created_at).toLocaleDateString('ru-RU')}</td>
            <td>
                <button class="btn btn-default" onclick="event.stopPropagation(); viewAgreement(${v.agreement_id})" title="Просмотр">
                    <i class="fas fa-eye"></i>
                </button>
            </td>
        </tr>`;
    }).join('');
}
