// ============================================================================
// Диалог подтверждения высвобождения в архиве
// ============================================================================

function showVysvobozhdenieDialog(result) {
    const vysv = result.vysvobozhdenie;
    if (!vysv) return;

    const container = document.getElementById('upload-result');
    const closedStage = vysv.closed_stage || '-';
    const closedAmount = vysv.closed_amount ? vysv.closed_amount.toLocaleString('ru-RU', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '-';
    const newPrice = vysv.new_contract_price ? vysv.new_contract_price.toLocaleString('ru-RU', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '-';

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

    // Если пользователь выбрал использовать высвобождение - показываем результат как есть
    // Если выбрал только изменение параметров - удаляем данные высвобождения из результата
    if (!useVysvobozhdenie && lastUploadResult.vysvobozhdenie) {
        // Создаем копию результата без данных высвобождения
        const modifiedResult = JSON.parse(JSON.stringify(lastUploadResult));
        delete modifiedResult.vysvobozhdenie;
        modifiedResult.has_vysvobozhdenie = false;
        modifiedResult.data = JSON.parse(JSON.stringify(lastUploadResult.data));
        delete modifiedResult.data.vysvobozhdenie;
        
        lastUploadResult = modifiedResult;
    }

    // Показываем обычный результат
    displayUploadResult(lastUploadResult);
}