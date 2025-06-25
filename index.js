
import jsPDF from 'jspdf';
import 'jspdf-autotable';

// --- GitHub Gist Configuration ---
const GIST_ID = 'ac0032048e7c8cef632893c7ed5907a5';
// !!! SEGURANÇA: O GIST_TOKEN NÃO DEVE SER EMBUTIDO DIRETAMENTE NO CÓDIGO DO LADO DO CLIENTE EM PRODUÇÃO! !!!
// !!! ISTO É UM RISCO DE SEGURANÇA GRAVE. USE UM BACKEND COMO PROXY PARA CHAMADAS À API DO GITHUB.  !!!
const GIST_TOKEN = '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'; // Replace with your actual token if needed for testing
const GIST_FILENAME = 'rdo_data.json';
const AUTO_SAVE_DEBOUNCE_DELAY = 3000; // ms
let autoSaveTimeoutId = null;

// --- ImgBB Configuration ---
// !!! SEGURANÇA: A CHAVE DA API IMGBB NÃO DEVE SER EMBUTIDA DIRETAMENTE NO CÓDIGO DO LADO DO CLIENTE EM PRODUÇÃO! !!!
// !!! ISTO É UM RISCO DE SEGURANÇA. CONSIDERE USAR UM BACKEND COMO PROXY. !!!
const IMGBB_API_KEY = '931646ab5a014e23cfdcd29053c87438';
const IMGBB_UPLOAD_URL = 'https://api.imgbb.com/1/upload';


// --- Global state for dirty form tracking ---
let isFormDirty = false;
let saveProgressButtonElement = null;
let efetivoWasJustClearedByButton = false;
let hasEfetivoBeenLoadedAtLeastOnce = false;

// --- Status Options ---
const statusOptions = ["Concluído", "Em Andamento", "Iniciado", "Paralisado", "Parado", ""];

// --- Localização Options ---
const localizacaoOptions = [
    "Ampliação do Sistema de Digestão de Lodo",
    "Canteiro de Obras",
    "Estação Elevatória de Lodo",
    "Isolamento Térmico dos 8 Digestores",
    "Plano de Gestão da Qualidade",
    "Pré Operação e Operação Assistida",
    "Sis. de Aquec. e Purif. dos Digestores",
    "Subst. Sist. Homogeneização de Lodo",
    "Rede Externa",
    ""
];

// --- Tipo de Serviço Options ---
const tipoServicoOptions = [
    "Arquitetura", "Automação", "Canteiro", "Controle Tecnológico", "Contratação",
    "Desmontagem", "Demolição", "Estruturas", "Equipamentos", "Fornecimento",
    "Fundações", "Geotécnico", "Hidráulico - Sanitárias", "Hidromecânico",
    "Impermeabilização", "Instalações Auxiliares", "Instalações Elétricas",
    "Instrumentação", "Intalações Elétricas Prediais", "Movimento de Entulho",
    "Movimento de Solo", "Montagem", "Obra Civil", "Revestimento",
    "Revest. Térmico", "Viário e Urbanismo", "Treinamento", "Teste",
    ""
];

// --- Serviço Options ---
const servicoOptions = [
    "Aço", "Alimentação", "Alvenaria", "Andaime", "Apicoamento", "Arrasamento", "Assentamento", "Aterramento", "Aterro",
    "Base", "Binder", "Bota Espera", "Bota Fora", "Cabeamento", "CAUQ", "Cimbramento", "Compactação", "Concreto",
    "Contenção", "Corte Verde", "Cura Concreto", "Demolição", "Desbaste", "Desforma", "Desmobilização", "Desmontagem",
    "Drenagem", "Ensaio", "Envelope de Concreto", "Envoltória de Areia", "Equipamento", "Escada", "Escavação",
    "Estanqueidade", "Execução", "Esgotamento", "Fabricação", "Fechamento", "Forma", "Fornecimento", "Fresagem",
    "Furo", "Fundação", "Geogrelha", "Grauteamento", "Impermeabilização", "Imprimação", "Infraestrutura", "Injeção",
    "Inspeção", "Instalação", "Instalação Elétrica", "Junta", "Laje Alveolar", "Lastro", "Limpeza", "Locação", "Magro",
    "Manutenção", "Mobilização", "Montagem", "Nivelamento", "Operação", "Passivação", "Pintura", "Posto Obra",
    "Pulverização", "Prova de Carga", "Reaterro", "Reboco", "Recup. Estrutural", "Regularização", "Remoção", "Rufos",
    "Regularização Mecânica", "Revest. Térmico", "Segurança", "Solda", "Sondagem", "Sub-Base", "Sub-Leito", "Suporte",
    "Telhado", "Testes", "Topografia",
    ""
];

// --- Month Year Picker State ---
let currentPickerYear = new Date().getFullYear();
let currentPickerMonth = new Date().getMonth(); // 0-indexed

// --- Responsive Scaling ---
const RDO_ORIGINAL_DESIGN_WIDTH = 1100; // px
const RDO_SIDEBAR_MARGIN = 20; // px, same as .efetivo-actions-sidebar margin-left
const originalRdoSheetHeights = new Map(); // Stores original height of each rdo-day-wrapper

// --- Helper Functions ---
function getElementByIdSafe(id, parent = document) {
    return parent.querySelector(`#${id}`);
}

function getInputValue(id, parent = document) {
    const el = parent.querySelector(`#${id}`);
    if (el && typeof el.value === 'string') {
        return el.value.trim();
    }
    return '';
}

function getTextAreaValue(id, parent = document) {
    const el = parent.querySelector(`#${id}`);
    if (el && typeof el.value === 'string') {
        return el.value;
    }
    return '';
}


function setInputValue(id, value, parent = document) {
    const el = parent.querySelector(`#${id}`);
    if (el) {
        el.value = value;
    }
}

function parseDateDDMMYYYY(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') {
        return null;
    }
    const parts = dateStr.split('/');
    if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1000 && year < 3000 && month >= 0 && month <= 11 && day > 0 && day <= 31) {
            const d = new Date(year, month, day);
            if (d.getFullYear() === year && d.getMonth() === month && d.getDate() === day) {
                return d;
            }
        }
    }
    console.warn(`Invalid date format for parseDateDDMMYYYY: ${dateStr}. Expected DD/MM/YYYY.`);
    return null;
}

function formatDateToYYYYMMDD(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
        console.warn("Invalid date passed to formatDateToYYYYMMDD:", date);
        return "";
    }
    return date.toISOString().split('T')[0];
}

function getDayOfWeekFromDate(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
        console.warn("Invalid date passed to getDayOfWeekFromDate:", date);
        return "";
    }
    const daysOfWeek = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
    return daysOfWeek[date.getDay()];
}

function calculateDaysBetween(startDate, endDate) {
    if (!(startDate instanceof Date) || isNaN(startDate.getTime()) || !(endDate instanceof Date) || isNaN(endDate.getTime())) {
        console.warn("Invalid dates passed to calculateDaysBetween:", startDate, endDate);
        return NaN;
    }
    const MS_PER_DAY = 1000 * 60 * 60 * 24;
    const utc1 = Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    const utc2 = Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
    return Math.floor((utc2 - utc1) / MS_PER_DAY);
}


function getNextDateInfo(currentDateStr) {
    const currentDate = new Date(currentDateStr + 'T00:00:00');
    currentDate.setDate(currentDate.getDate() + 1);
    const nextDate = formatDateToYYYYMMDD(currentDate);
    const dayOfWeek = getDayOfWeekFromDate(currentDate);
    return { nextDate, dayOfWeek };
}

function getDaysInMonth(year, month_1_indexed) {
    if (month_1_indexed < 1 || month_1_indexed > 12 || year < 1000 || year > 3000) {
        return NaN;
    }
    return new Date(year, month_1_indexed, 0).getDate();
}

// --- Image Upload (ImgBB) ---
async function dataUrlToFile(dataUrl, filename) {
    try {
        const res = await fetch(dataUrl);
        const blob = await res.blob();
        return new File([blob], filename, { type: blob.type });
    } catch (error) {
        console.error("Error converting data URL to File:", error);
        throw error;
    }
}

async function uploadImageToImgBB(file, context = "image") {
    if (!IMGBB_API_KEY) {
        console.error("ImgBB API Key is missing.");
        updateSaveProgressMessage("Erro: Chave da API ImgBB não configurada.", true);
        return null;
    }
    if (!(file instanceof File)) {
        console.error("Invalid file object provided to uploadImageToImgBB");
        updateSaveProgressMessage("Erro: Arquivo de imagem inválido para upload.", true);
        return null;
    }

    const formData = new FormData();
    formData.append('image', file);
    formData.append('key', IMGBB_API_KEY);

    // Keep specific message for direct image selection action
    updateSaveProgressMessage(`Enviando ${context} para ImgBB...`, false);

    try {
        const response = await fetch(IMGBB_UPLOAD_URL, {
            method: 'POST',
            body: formData,
        });

        const responseData = await response.json();

        if (response.ok && responseData.success && responseData.data && responseData.data.url) {
            updateSaveProgressMessage(`${context} enviada com sucesso!`, false); // Keep specific for direct action
            return responseData.data.url;
        } else {
            console.error('ImgBB Upload Error:', responseData);
            updateSaveProgressMessage(`Falha ao enviar ${context}: ${responseData?.error?.message || response.statusText}`, true);
            return null;
        }
    } catch (error) {
        console.error('Network or other error during ImgBB upload:', error);
        updateSaveProgressMessage(`Erro de rede ao enviar ${context}: ${error.message}`, true);
        return null;
    }
}

// --- Loading Spinner Control ---
function showLoadingSpinner() {
    const overlay = getElementByIdSafe('loading-overlay');
    if (overlay) {
        overlay.style.display = 'flex';
    }
}

function hideLoadingSpinner() {
    const overlay = getElementByIdSafe('loading-overlay');
    if (overlay) {
        overlay.style.display = 'none';
    }
}

// --- Confirmation Modal Control ---
let confirmationModalOverlay = null;
let confirmClearRdoButton = null;
let cancelClearRdoButton = null;

function showConfirmationModal() {
    if (!confirmationModalOverlay) {
        confirmationModalOverlay = getElementByIdSafe('confirmation-modal-overlay');
    }
    if (confirmationModalOverlay) {
        confirmationModalOverlay.style.display = 'flex';
        if(confirmClearRdoButton) confirmClearRdoButton.focus();
    }
}

function hideConfirmationModal() {
    if (confirmationModalOverlay) {
        confirmationModalOverlay.style.display = 'none';
    }
    const clearRdoBtn = getElementByIdSafe('clearRdoButton');
    if (clearRdoBtn) clearRdoBtn.focus();
}


// --- Dirty Form State Management & Auto-Save ---
function updateSaveButtonState() {
    if (!saveProgressButtonElement) {
        saveProgressButtonElement = getElementByIdSafe('saveProgressButton');
    }
    if (saveProgressButtonElement) {
        saveProgressButtonElement.disabled = !isFormDirty;
        saveProgressButtonElement.setAttribute('aria-disabled', String(!isFormDirty));
    }
}

async function autoSaveToGist() {
    console.log("Auto-save triggered...");
    updateSaveProgressMessage("Salvando progresso...", false); // Generic start message
    const rdoData = serializeRdoDataForGist();
    if (rdoData) {
        try {
            const success = await updateGistContent(JSON.stringify(rdoData));
            if (success) {
                isFormDirty = false;
                updateSaveButtonState();
                // "Progresso salvo!" message is handled by updateGistContent
            } else {
                // Error message handled by updateGistContent
            }
        } catch (error) {
            console.error("Error during autoSaveToGist process:", error);
            updateSaveProgressMessage(`Erro no salvamento automático: ${error.message}`, true);
        }
    }
}

const debouncedAutoSave = () => {
    clearTimeout(autoSaveTimeoutId);
    autoSaveTimeoutId = setTimeout(autoSaveToGist, AUTO_SAVE_DEBOUNCE_DELAY);
};

function markFormAsDirty() {
    if (!isFormDirty) {
        isFormDirty = true;
        updateSaveButtonState();
    }
    debouncedAutoSave();
}

// --- Core RDO Day Functionality ---
let rdoDayCounter = 0;
let pristineRdoDay0WrapperHtml = '';

const shiftControlOptions = {
    tempo: { options: ["", "B", "L", "F"], ariaLabelPart: 'Condição do tempo' },
    trabalho: { options: ["", "N", "P", "T"], ariaLabelPart: 'Condição de trabalho' }
};

const globalMasterFieldBases = ['contratada', 'contrato_num', 'prazo', 'inicio_obra'];
const contractFieldMappingsBase = [
    { masterIdBase: 'contratada', slaveIdBases: ['contratada_pt', 'contratada_pt2'] },
    { masterIdBase: 'data', slaveIdBases: ['data_pt', 'data_pt2'] },
    { masterIdBase: 'contrato_num', slaveIdBases: ['contrato_num_pt', 'contrato_num_pt2'] },
    { masterIdBase: 'dia_semana', slaveIdBases: ['dia_semana_pt', 'dia_semana_pt2'] },
    { masterIdBase: 'prazo', slaveIdBases: ['prazo_pt', 'prazo_pt2'] },
    { masterIdBase: 'inicio_obra', slaveIdBases: ['inicio_obra_pt', 'inicio_obra_pt2'] },
    { masterIdBase: 'decorridos', slaveIdBases: ['decorridos_pt', 'decorridos_pt2'] },
    { masterIdBase: 'restantes', slaveIdBases: ['restantes_pt', 'restantes_pt2'] },
];

function generateDropdownHTML(idBase, namePrefix, optionsArray, ariaLabelPrefix, dayIndex, rowCount) {
    let optionsHTML = '';
    optionsArray.forEach(optionText => {
        const displayText = optionText === "" ? "(Limpar)" : optionText;
        optionsHTML += `<li class="status-option" data-value="${optionText}" role="option" aria-selected="false">${displayText}</li>`;
    });

    return `
        <div class="status-select-container">
            <div class="status-display" id="${idBase}_display" tabindex="0" role="button" aria-haspopup="listbox" aria-expanded="false" aria-label="${ariaLabelPrefix}, Dia ${dayIndex + 1}, Linha ${rowCount + 1}"></div>
            <input type="hidden" name="${namePrefix}_day${dayIndex}[]" id="${idBase}_hidden">
            <ul class="status-dropdown" role="listbox" id="${idBase}_dropdown" style="display: none;">
                ${optionsHTML}
            </ul>
        </div>
    `;
}

function addActivityRow(tableId, dayIndex, pageContainerElement) {
    if (!pageContainerElement) {
        console.warn(`addActivityRow: pageContainerElement not provided for table ${tableId}, day ${dayIndex}. Rows not added.`);
        return;
    }
    const tableBody = getElementByIdSafe(`${tableId}_day${dayIndex}`, pageContainerElement)?.getElementsByTagName('tbody')[0];
    if (!tableBody) {
        console.warn(`addActivityRow: Table body for ${tableId}_day${dayIndex} not found in provided page container (ID: ${pageContainerElement.id}). Rows not added.`);
        return;
    }


    const newRow = tableBody.insertRow();
    newRow.style.height = '20px';
    const tableBodyRowCount = tableBody.rows.length - 1; // 0-indexed for IDs

    const cell1 = newRow.insertCell(); // Localização
    const cell2 = newRow.insertCell(); // Tipo
    const cell3 = newRow.insertCell(); // Serviços
    const cell4 = newRow.insertCell(); // Status
    const cell5 = newRow.insertCell(); // Observações

    const localizacaoCellIdBase = `localizacao_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell1.classList.add('status-cell');
    cell1.innerHTML = generateDropdownHTML(localizacaoCellIdBase, 'localizacao', localizacaoOptions, 'Localização da atividade', dayIndex, tableBodyRowCount);

    const tipoServicoCellIdBase = `tipo_servico_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell2.classList.add('status-cell');
    cell2.innerHTML = generateDropdownHTML(tipoServicoCellIdBase, 'tipo_servico', tipoServicoOptions, 'Tipo de serviço da atividade', dayIndex, tableBodyRowCount);

    const servicoCellIdBase = `servico_desc_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell3.classList.add('status-cell');
    cell3.innerHTML = generateDropdownHTML(servicoCellIdBase, 'servico_desc', servicoOptions, 'Serviço da atividade', dayIndex, tableBodyRowCount);

    const statusCellIdBase = `status_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}`;
    cell4.classList.add('status-cell');
    cell4.innerHTML = generateDropdownHTML(statusCellIdBase, 'status', statusOptions, 'Status da atividade', dayIndex, tableBodyRowCount);

    cell5.innerHTML = `<textarea name="observacoes_atividade_day${dayIndex}[]" id="obs_atividade_day${dayIndex}_${tableId.replace(/Table_?/, '')}_row${tableBodyRowCount}" aria-label="Observações da atividade, Dia ${dayIndex + 1}, Linha ${tableBodyRowCount + 1}"></textarea>`;


    [
        { cell: cell1, baseId: localizacaoCellIdBase },
        { cell: cell2, baseId: tipoServicoCellIdBase },
        { cell: cell3, baseId: servicoCellIdBase },
        { cell: cell4, baseId: statusCellIdBase }
    ].forEach(config => {
        const displayElement = config.cell.querySelector(`#${config.baseId}_display`);
        const hiddenInputElement = config.cell.querySelector(`#${config.baseId}_hidden`);
        const dropdownElement = config.cell.querySelector(`#${config.baseId}_dropdown`);

        if (displayElement && hiddenInputElement && dropdownElement) {
            displayElement.addEventListener('click', (event) => {
                event.stopPropagation();
                document.querySelectorAll('.status-dropdown').forEach(dd => {
                    if (dd.id !== dropdownElement.id) {
                        (dd).style.display = 'none';
                        const otherDisplayId = dd.id.replace('_dropdown', '_display');
                        const otherDisplay = document.getElementById(otherDisplayId);
                        if (otherDisplay) otherDisplay.setAttribute('aria-expanded', 'false');
                    }
                });

                const isCurrentlyOpen = dropdownElement.style.display === 'block';
                dropdownElement.style.display = isCurrentlyOpen ? 'none' : 'block';
                displayElement.setAttribute('aria-expanded', String(!isCurrentlyOpen));
                 if (!isCurrentlyOpen) {
                    const currentVal = hiddenInputElement.value;
                    let foundSelected = false;
                    dropdownElement.querySelectorAll('.status-option').forEach(opt => {
                        if (opt.dataset.value === currentVal) {
                            opt.setAttribute('aria-selected', 'true');
                            foundSelected = true;
                        } else {
                            opt.setAttribute('aria-selected', 'false');
                        }
                    });
                     if (!foundSelected) {
                         const emptyOpt = Array.from(dropdownElement.querySelectorAll('.status-option')).find(o => o.dataset.value === "");
                         if(emptyOpt) emptyOpt.setAttribute('aria-selected', 'true');
                    }
                }
            });

            displayElement.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); displayElement.click(); }
                else if (e.key === 'Escape' && dropdownElement.style.display === 'block') {
                    dropdownElement.style.display = 'none';
                    displayElement.setAttribute('aria-expanded', 'false');
                    displayElement.focus();
                }
            });

            dropdownElement.querySelectorAll('.status-option').forEach(optionElement => {
                optionElement.addEventListener('click', () => {
                    const selectedValue = optionElement.dataset.value || "";
                    displayElement.textContent = selectedValue === "" ? "" : optionElement.textContent;
                    hiddenInputElement.value = selectedValue;
                    hiddenInputElement.dispatchEvent(new Event('input', { bubbles: true }));
                    markFormAsDirty();

                    dropdownElement.querySelectorAll('.status-option').forEach(opt => opt.setAttribute('aria-selected', 'false'));
                    optionElement.setAttribute('aria-selected', 'true');

                    dropdownElement.style.display = 'none';
                    displayElement.setAttribute('aria-expanded', 'false');
                    displayElement.focus();
                });
            });
            hiddenInputElement.addEventListener('input', markFormAsDirty);
        }
    });

    const obsTextarea = cell5.querySelector('textarea');
    if(obsTextarea) {
        obsTextarea.addEventListener('input', markFormAsDirty);
    }
}


const shiftControlUpdaters = new Map();


function createCombinedShiftControl(
    cellIdBase,
    turnoIdBase,
    initialTempoValue,
    initialTrabalhoValue,
    ariaTurnoSuffixBase,
    dayIndex,
    parentElement
) {
    const cellId = `${cellIdBase}_day${dayIndex}`;
    const cell = getElementByIdSafe(cellId, parentElement);
    if (!cell) {
        console.warn(`Shift control cell with ID ${cellId} not found within provided parent (Parent ID: ${parentElement?.id}, Day Index: ${dayIndex}). Control not created.`);
        return;
    }

    const turnoId = `${turnoIdBase}_day${dayIndex}`;
    const ariaTurnoSuffix = `${ariaTurnoSuffixBase} Dia ${dayIndex + 1}`;

    const tempoConfig = shiftControlOptions.tempo;
    const trabalhoConfig = shiftControlOptions.trabalho;

    let currentTempoIndex = Math.max(0, tempoConfig.options.indexOf(initialTempoValue));
    let currentTrabalhoIndex = Math.max(0, trabalhoConfig.options.indexOf(initialTrabalhoValue));

    const controlWrapper = document.createElement('div');
    controlWrapper.className = 'weather-toggle-control';
    controlWrapper.setAttribute('role', 'group');
    controlWrapper.setAttribute('aria-label', `Controles de condição para ${ariaTurnoSuffix}`);

    const tempoHiddenInput = document.createElement('input');
    tempoHiddenInput.type = 'hidden';
    tempoHiddenInput.name = `tempo_${turnoId}_value`;
    tempoHiddenInput.id = `tempo_${turnoId}_hidden`;


    const trabalhoHiddenInput = document.createElement('input');
    trabalhoHiddenInput.type = 'hidden';
    trabalhoHiddenInput.name = `trabalho_${turnoId}_value`;
    trabalhoHiddenInput.id = `trabalho_${turnoId}_hidden`;


    const topHalf = document.createElement('div');
    topHalf.className = 'toggle-half toggle-top';
    topHalf.setAttribute('role', 'button');
    topHalf.tabIndex = 0;

    const bottomHalf = document.createElement('div');
    bottomHalf.className = 'toggle-half toggle-bottom';
    bottomHalf.setAttribute('role', 'button');
    bottomHalf.tabIndex = 0;

    const updateAriaLabels = () => {
        const tempoValue = tempoConfig.options[currentTempoIndex];
        const trabalhoValue = trabalhoConfig.options[currentTrabalhoIndex];
        const tempoDisplayValue = tempoValue === "" ? "(Não selecionado)" : tempoValue;
        const trabalhoDisplayValue = trabalhoValue === "" ? "(Não selecionado)" : trabalhoValue;
        topHalf.setAttribute('aria-label', `${tempoConfig.ariaLabelPart} para ${ariaTurnoSuffix}: ${tempoDisplayValue}. Clique para alterar.`);
        bottomHalf.setAttribute('aria-label', `${trabalhoConfig.ariaLabelPart} para ${ariaTurnoSuffix}: ${trabalhoDisplayValue}. Clique para alterar.`);
    };

    const updateVisualsAndState = (newTempoValue, newTrabalhoValue) => {
        currentTempoIndex = Math.max(0, tempoConfig.options.indexOf(newTempoValue));
        currentTrabalhoIndex = Math.max(0, trabalhoConfig.options.indexOf(newTrabalhoValue));

        const currentTempoDisplayValue = tempoConfig.options[currentTempoIndex];
        const currentTrabalhoDisplayValue = trabalhoConfig.options[currentTrabalhoIndex];

        topHalf.textContent = currentTempoDisplayValue;
        tempoHiddenInput.value = currentTempoDisplayValue;
        bottomHalf.textContent = currentTrabalhoDisplayValue;
        trabalhoHiddenInput.value = currentTrabalhoDisplayValue;
        updateAriaLabels();
    };

    shiftControlUpdaters.set(cellId, { update: updateVisualsAndState });
    updateVisualsAndState(initialTempoValue, initialTrabalhoValue);

    topHalf.addEventListener('click', () => {
        currentTempoIndex = (currentTempoIndex + 1) % tempoConfig.options.length;
        const newTempoValue = tempoConfig.options[currentTempoIndex];
        topHalf.textContent = newTempoValue;
        tempoHiddenInput.value = newTempoValue;
        updateAriaLabels();
        markFormAsDirty();
    });
    topHalf.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); topHalf.click(); } });

    bottomHalf.addEventListener('click', () => {
        currentTrabalhoIndex = (currentTrabalhoIndex + 1) % trabalhoConfig.options.length;
        const newTrabalhoValue = trabalhoConfig.options[currentTrabalhoIndex];
        bottomHalf.textContent = newTrabalhoValue;
        trabalhoHiddenInput.value = newTrabalhoValue;
        updateAriaLabels();
        markFormAsDirty();
    });
    bottomHalf.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); bottomHalf.click(); } });

    controlWrapper.appendChild(tempoHiddenInput);
    controlWrapper.appendChild(trabalhoHiddenInput);
    controlWrapper.appendChild(topHalf);
    controlWrapper.appendChild(bottomHalf);

    cell.innerHTML = '';
    cell.appendChild(controlWrapper);
}

function applyMultiColumnLayoutIfNecessary(tableBodyElement) {
    if (!tableBodyElement) return;
    if (tableBodyElement.getElementsByTagName('tr').length > 18) {
        tableBodyElement.classList.add('multi-column-items');
    }
}


function resetReportPhotoSlotState(
    photoInput, photoPreview,
    photoPlaceholder, clearButton,
    labelElement
) {
    if (photoPreview) {
        photoPreview.src = '#';
        photoPreview.style.display = 'none';
        photoPreview.removeAttribute('data-natural-width');
        photoPreview.removeAttribute('data-natural-height');
        photoPreview.removeAttribute('data-is-effectively-loaded');
        photoPreview.style.width = '100%';
        photoPreview.style.height = '100%';
        photoPreview.style.objectFit = 'cover';
        photoPreview.style.position = 'relative';
        photoPreview.style.transform = 'none';
        photoPreview.style.left = 'auto';
        photoPreview.style.top = 'auto';
    }
    if (photoPlaceholder) photoPlaceholder.style.display = 'block';
    if (clearButton) clearButton.style.display = 'none';
    if (labelElement) {
        labelElement.classList.remove('drag-over');
        labelElement.style.cursor = 'pointer';
    }
    if (photoInput) {
        photoInput.value = '';
        photoInput.disabled = false;
    }
}

function setReportPhotoSlotImage(
    photoInput, photoPreview,
    photoPlaceholder, clearButton,
    labelElement,
    imageUrl, naturalWidth, naturalHeight, file
) {
    if (!photoPreview || !labelElement || !photoPlaceholder || !clearButton || !photoInput) return;

    if (!imageUrl || imageUrl === '#' || imageUrl.startsWith('about:blank')) {
        resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
        return;
    }

    const tempImg = new Image();
    tempImg.onload = () => {
        const nw = naturalWidth ? parseInt(naturalWidth) : tempImg.naturalWidth;
        const nh = naturalHeight ? parseInt(naturalHeight) : tempImg.naturalHeight;

        if (!nw || !nh) {
            console.warn("Could not determine image dimensions for setReportPhotoSlotImage.");
            resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
            return;
        }

        photoPreview.src = imageUrl;
        photoPreview.style.display = 'block';
        photoPreview.style.width = '100%';
        photoPreview.style.height = '100%';
        photoPreview.style.objectFit = 'cover';
        photoPreview.style.position = 'relative';
        photoPreview.style.transform = 'none';
        photoPreview.style.left = 'auto';
        photoPreview.style.top = 'auto';
        photoPreview.dataset.naturalWidth = nw.toString();
        photoPreview.dataset.naturalHeight = nh.toString();
        photoPreview.dataset.isEffectivelyLoaded = 'true';


        photoPlaceholder.style.display = 'none';
        clearButton.style.display = 'block';

        if (imageUrl.startsWith('http')) {
             photoInput.value = '';
        } else if (file) {

        }

        photoInput.disabled = true;
        labelElement.style.cursor = 'default';
    };
    tempImg.onerror = () => {
        console.warn("Error loading image in setReportPhotoSlotImage: ", imageUrl);
        resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
    };
    tempImg.src = imageUrl;
}


async function setupReportPhotoSlot(
    inputIdBase,
    previewIdBase,
    placeholderIdBase,
    clearButtonIdBase,
    labelIdBase,
    dayIndex,
    photoSlotNumber,
    parentElement
) {
    const daySuffix = `_day${dayIndex}`;
    const photoInput = getElementByIdSafe(`${inputIdBase}${daySuffix}`, parentElement);
    const photoPreview = getElementByIdSafe(`${previewIdBase}${daySuffix}`, parentElement);
    const photoPlaceholder = getElementByIdSafe(`${placeholderIdBase}${daySuffix}`, parentElement);
    const clearButton = getElementByIdSafe(`${clearButtonIdBase}${daySuffix}`, parentElement);
    const labelElement = getElementByIdSafe(`${labelIdBase}${daySuffix}`, parentElement);

    if (!photoInput || !photoPreview || !photoPlaceholder || !clearButton || !labelElement) {
        console.warn(`Missing elements for report photo slot: ${inputIdBase}${daySuffix} (Slot ${photoSlotNumber}) in parent. Parent ID: ${parentElement ? parentElement.id : 'N/A'}. Day Index: ${dayIndex}`);
        return;
    }

    labelElement.style.cursor = 'pointer';
    photoInput.disabled = false;

    const processFile = async (file) => {
        if (file && file.type.startsWith('image/')) {
            const reader = new FileReader();
            reader.onload = async function(e_reader) {
                if (e_reader.target?.result) {
                    setReportPhotoSlotImage(
                        photoInput, photoPreview, photoPlaceholder, clearButton, labelElement,
                        e_reader.target.result, undefined, undefined, file
                    );

                    const imgbbUrl = await uploadImageToImgBB(file, `Relatório Foto ${photoSlotNumber} (Dia ${dayIndex + 1})`);
                    if (imgbbUrl) {
                        setReportPhotoSlotImage(
                            photoInput, photoPreview, photoPlaceholder, clearButton, labelElement,
                            imgbbUrl, undefined, undefined, null
                        );
                        markFormAsDirty();
                    } else {
                        resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
                        // User feedback for specific image upload failure is kept.
                    }
                }
            }
            reader.readAsDataURL(file);
        } else if (file) {
            alert("Por favor, solte apenas arquivos de imagem.");
            if (labelElement) labelElement.classList.remove('drag-over');
        }
    };

    photoInput.addEventListener('change', async (event) => {
        const file = (event.target).files?.[0];
        if (file) await processFile(file);
    });

    clearButton.addEventListener('click', () => {
        resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
        markFormAsDirty();
    });

    labelElement.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); labelElement.classList.add('drag-over'); });
    labelElement.addEventListener('dragleave', (e) => { e.preventDefault(); e.stopPropagation(); labelElement.classList.remove('drag-over'); });
    labelElement.addEventListener('drop', async (e) => {
        e.preventDefault(); e.stopPropagation(); labelElement.classList.remove('drag-over');
        if (photoInput.disabled && photoPreview.style.display !== 'none') {
            return;
        }
        const files = e.dataTransfer?.files;
        if (files && files.length > 0) {
            await processFile(files[0]);
        }
    });

    resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
}


function updateAllSignaturePreviews(imageUrl) {
    const isImageEffectivelyLoaded = imageUrl &&
                                     imageUrl !== '#' &&
                                     !imageUrl.startsWith('about:blank') &&
                                     !imageUrl.endsWith('/#') &&
                                     imageUrl.startsWith('http');

    for (let d = 0; d <= rdoDayCounter; d++) {
        const daySuffix = `_day${d}`;
        const dayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        if (!dayWrapper) continue;

        const previewIdsAndContainers = [
            { previewId: `consbemPhotoPreview${daySuffix}`, containerId: `consbemPhotoContainer${daySuffix}`, labelFor: `consbemPhoto${daySuffix}`, clearId: `clearConsbemPhoto${daySuffix}`, inputId: `consbemPhoto${daySuffix}` },
            { previewId: `consbemPhotoPreview_pt${daySuffix}`, containerId: `consbemPhotoContainer_pt${daySuffix}`, labelFor: `consbemPhoto_pt${daySuffix}`, clearId: `clearConsbemPhoto_pt${daySuffix}`, inputId: `consbemPhoto_pt${daySuffix}` },
            { previewId: `consbemPhotoPreview_pt2${daySuffix}`, containerId: `consbemPhotoContainer_pt2${daySuffix}`, labelFor: `consbemPhoto_pt2${daySuffix}`, clearId: `clearConsbemPhoto_pt2${daySuffix}`, inputId: `consbemPhoto_pt2${daySuffix}` }
        ];

        previewIdsAndContainers.forEach(pc => {
            const preview = getElementByIdSafe(pc.previewId, dayWrapper);
            const container = getElementByIdSafe(pc.containerId, dayWrapper);
            const uploadButton = dayWrapper.querySelector(`label[for="${pc.labelFor}"]`);
            const clearButton = getElementByIdSafe(pc.clearId, dayWrapper);
            const inputElement = getElementByIdSafe(pc.inputId, dayWrapper);

            if (preview && container && uploadButton && clearButton && inputElement) {
                if (isImageEffectivelyLoaded && imageUrl) {
                    preview.src = imageUrl;
                    preview.style.display = 'block';
                    preview.dataset.isEffectivelyLoaded = 'true';
                    container.style.display = 'inline-block';
                    inputElement.value = '';

                    uploadButton.style.display = 'none';
                    inputElement.disabled = true;
                    const isDay0Page1Consbem = d === 0 && pc.inputId === `consbemPhoto${daySuffix}`;
                    clearButton.style.display = isDay0Page1Consbem ? 'block' : 'none';

                } else {
                    preview.src = '#';
                    preview.style.display = 'none';
                    preview.dataset.isEffectivelyLoaded = 'false';
                    container.style.display = 'none';
                    clearButton.style.display = 'none';
                    inputElement.value = '';

                    if (d === 0 && pc.inputId === `consbemPhoto${daySuffix}`) {
                        uploadButton.style.display = 'inline-block';
                        inputElement.disabled = false;
                    } else {
                        uploadButton.style.display = 'none';
                        inputElement.disabled = true;
                    }
                }
            }
        });
    }
}

async function setupPhotoAttachmentForDay(dayIndex) {
    const daySuffix = `_day${dayIndex}`;

    const masterPhotoInput = getElementByIdSafe(`consbemPhoto${daySuffix}`);
    const masterPhotoPreview = getElementByIdSafe(`consbemPhotoPreview${daySuffix}`);
    const masterClearPhotoButton = getElementByIdSafe(`clearConsbemPhoto${daySuffix}`);

    if (dayIndex === 0) {
        if (!masterPhotoInput || !masterPhotoPreview || !masterClearPhotoButton) {
            console.warn("Master signature elements for Day 0 not fully found. Signature functionality may be impaired.");
            return;
        }

        masterPhotoInput.addEventListener('change', async function(event) {
            const file = (event.target).files?.[0];
            if (file) {
                const imgbbUrl = await uploadImageToImgBB(file, "Assinatura CONSBEM"); // Specific message for this action is fine
                if (imgbbUrl) {
                    updateAllSignaturePreviews(imgbbUrl);
                    markFormAsDirty();
                } else {
                    updateAllSignaturePreviews(null);
                    masterPhotoInput.value = '';
                }
            } else {
                updateAllSignaturePreviews(null);
                markFormAsDirty();
            }
        });

        masterClearPhotoButton.addEventListener('click', function() {
            masterPhotoInput.value = '';
            updateAllSignaturePreviews(null);
            markFormAsDirty();
        });

        const day0Wrapper = getElementByIdSafe('rdo_day_0_wrapper');
        if (day0Wrapper) {
            ['_pt', '_pt2'].forEach(pagePrefix => {
                const slaveUploadLabel = day0Wrapper.querySelector(`label[for="consbemPhoto${pagePrefix}${daySuffix}"]`);
                if (slaveUploadLabel) slaveUploadLabel.style.display = 'none';

                const slaveClearBtn = getElementByIdSafe(`clearConsbemPhoto${pagePrefix}${daySuffix}`, day0Wrapper);
                if (slaveClearBtn) slaveClearBtn.style.display = 'none';

                const slaveInput = getElementByIdSafe(`consbemPhoto${pagePrefix}${daySuffix}`, day0Wrapper);
                if (slaveInput) slaveInput.disabled = true;
            });
        }

    } else {
        ['', '_pt', '_pt2'].forEach(pagePrefix => {
            const dayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
            if(dayWrapper) {
                const uploadButton = dayWrapper.querySelector(`label[for="consbemPhoto${pagePrefix}${daySuffix}"]`);
                if (uploadButton) uploadButton.style.display = 'none';

                const clearButton = getElementByIdSafe(`clearConsbemPhoto${pagePrefix}${daySuffix}`, dayWrapper);
                if (clearButton) clearButton.style.display = 'none';

                const inputElement = getElementByIdSafe(`consbemPhoto${pagePrefix}${daySuffix}`, dayWrapper);
                if (inputElement) inputElement.disabled = true;
            }
        });
    }
}

// --- Debounced Efetivo Total Calculation ---
const calculateTotalsDebounceTimeouts = new Map();
const DEBOUNCE_CALCULATE_TOTALS_DELAY = 300;

function debouncedCalculateSectionTotalsForDay(pageContainer, dayIndex) {
    if (calculateTotalsDebounceTimeouts.has(dayIndex)) {
        clearTimeout(calculateTotalsDebounceTimeouts.get(dayIndex));
    }
    calculateTotalsDebounceTimeouts.set(dayIndex, setTimeout(() => {
        calculateSectionTotalsForDay(pageContainer, dayIndex);
        calculateTotalsDebounceTimeouts.delete(dayIndex);
    }, DEBOUNCE_CALCULATE_TOTALS_DELAY));
}

function calculateSectionTotalsForDay(dayWrapperOrPageContainer, dayIndex) {
    const laborSections = [
        { columnClass: 'labor-direta', totalInputSelectorSuffix: '.total-field input.section-total-input' },
        { columnClass: 'labor-indireta', totalInputSelectorSuffix: '.total-field input.section-total-input' },
        { columnClass: 'equipamentos', totalInputSelectorSuffix: '.total-field input.section-total-input' }
    ];

    laborSections.forEach(section => {
        const columnElement = dayWrapperOrPageContainer.querySelector('.' + section.columnClass);
        if (columnElement) {
            let sectionTotal = 0;
            const quantityInputs = columnElement.querySelectorAll('.labor-table .quantity-input');
            quantityInputs.forEach(input => {
                sectionTotal += parseInt(input.value) || 0;
            });
            const totalSectionInput = columnElement.querySelector(section.totalInputSelectorSuffix);
            if (totalSectionInput) {
                totalSectionInput.value = sectionTotal.toString();
            }
        }
    });
    checkAndToggleEfetivoAlertForDay(dayIndex);
    updateClearEfetivoButtonVisibility();
}


// --- Rainfall Data Fetching ---
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
const REQUEST_DELAY_MS = 1000;

let isGloballyRateLimited = false;
let globalRateLimitRetryAfterTimestamp = 0;
const GLOBAL_RATE_LIMIT_COOLDOWN_MS = 61000;



let globalFetchMessageContainer = null;
let globalSaveMessageContainer = null;
let fetchMessageTimeoutId = null;
let saveProgressMessageTimeoutId = null;
const MESSAGE_AUTO_HIDE_DURATION = 7000;


let rainfallProgressBarWrapper = null;
let rainfallProgressBar = null;
let rainfallProgressText = null;
let rainfallProgressDetailsText = null;


function updateFetchMessage(message, isError = false) {
    if (!globalFetchMessageContainer) {
        globalFetchMessageContainer = getElementByIdSafe('fetchRainfallMessageContainer');
    }

    if (fetchMessageTimeoutId) {
        clearTimeout(fetchMessageTimeoutId);
        fetchMessageTimeoutId = null;
    }

    if (globalFetchMessageContainer) {
        if (message && message.trim() !== "") {
            globalFetchMessageContainer.textContent = message;
            globalFetchMessageContainer.style.display = 'block';
            globalFetchMessageContainer.style.color = isError ? '#f44336' : '#212529';

            if (!isError) {
                fetchMessageTimeoutId = window.setTimeout(() => {
                    updateFetchMessage(null);
                }, MESSAGE_AUTO_HIDE_DURATION);
            }
        } else {
            globalFetchMessageContainer.textContent = '';
            globalFetchMessageContainer.style.display = 'none';
        }
    }
}

function updateSaveProgressMessage(message, isError = false) {
    if (!globalSaveMessageContainer) {
        globalSaveMessageContainer = getElementByIdSafe('saveProgressMessageContainer');
    }

    if (saveProgressMessageTimeoutId) {
        clearTimeout(saveProgressMessageTimeoutId);
        saveProgressMessageTimeoutId = null;
    }

    if (globalSaveMessageContainer) {
        if (message && message.trim() !== "") {
            globalSaveMessageContainer.innerHTML = message.replace(/\n/g, '<br>');
            globalSaveMessageContainer.style.display = 'block';
            globalSaveMessageContainer.style.color = isError ? '#d32f2f' : '#1b5e20';

            const lowerCaseMessage = message.toLowerCase();
            const persistentMessages = [
                "salvando progresso...", "limpando rdo...", "carregando progresso...", "processando imagens...", "enviando", // Keep "enviando" for ImgBB direct actions
                // Exclude PDF generation messages as they have their own progress bar
            ];
            const isPersistent = persistentMessages.some(pm => lowerCaseMessage.includes(pm));

            if (!isPersistent && !isError) {
                saveProgressMessageTimeoutId = window.setTimeout(() => {
                    updateSaveProgressMessage(null);
                }, MESSAGE_AUTO_HIDE_DURATION);
            }
        } else {
            globalSaveMessageContainer.textContent = '';
            globalSaveMessageContainer.style.display = 'none';
        }
    }
}


function updateRainfallStatus(dayIndex, status, customMessage) {
    const statusEl = getElementByIdSafe(`indice_pluv_status_day${dayIndex}`);
    if (statusEl) {
        let messageToDisplay = customMessage;
        let tooltipText = '';

        if (messageToDisplay === undefined) {
            switch (status) {
                case 'loading': messageToDisplay = '...'; break;
                case 'success': messageToDisplay = '✓'; break;
                case 'error': messageToDisplay = 'X'; break;
                case 'idle': messageToDisplay = '-'; break;
                default: messageToDisplay = '-';
            }
        }
        statusEl.textContent = messageToDisplay;

        statusEl.className = 'rainfall-status';
        if (status !== 'idle') {
            statusEl.classList.add(status);
        }

        switch (messageToDisplay) {
            case 'P':
                tooltipText = "Previsão de chuva. Este valor é uma estimativa futura.";
                statusEl.classList.add('forecast');
                statusEl.classList.remove('success');
                break;
            case '✓':
                tooltipText = "Dados de chuva obtidos com sucesso.";
                statusEl.classList.remove('forecast');
                break;
            case 'X':
                tooltipText = "Erro ao buscar dados de chuva.";
                statusEl.classList.remove('forecast');
                break;
            case '-':
                tooltipText = "Dados de chuva não disponíveis ou fora do período de busca.";
                statusEl.classList.remove('forecast');
                break;
            case '...':
                tooltipText = "Buscando dados de chuva...";
                statusEl.classList.remove('forecast');
                break;
            case 'RL':
                tooltipText = "Limite de requisições da API atingido. Tente mais tarde.";
                statusEl.classList.remove('forecast');
                statusEl.classList.add('error');
                break;
            case 'Data Inválida':
                tooltipText = "Data inválida para busca de chuva.";
                statusEl.classList.remove('forecast');
                break;
            default:
                tooltipText = "Status dos dados pluviométricos.";
                statusEl.classList.remove('forecast');
        }
        statusEl.title = tooltipText;
    }
}

async function fetchRainfallData(dateStrYYYYMMDD, dayIndex) {
    const inputEl = getElementByIdSafe(`indice_pluv_valor_day${dayIndex}`);
    const dayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
    const page1Container = dayWrapper?.querySelector('.rdo-container');


    if (!inputEl || !dayWrapper || !page1Container) {
        console.warn(`Rainfall input or day/page1 wrapper for day ${dayIndex} not found. Skipping fetch.`);
        return 'error';
    }

    const systemToday = new Date();
    systemToday.setHours(0, 0, 0, 0);
    const requestedDateObj = new Date(dateStrYYYYMMDD + 'T00:00:00');

    const maxForecastDays = 15;
    const maxForecastDate = new Date(systemToday);
    maxForecastDate.setDate(systemToday.getDate() + maxForecastDays);
    const minApiDate = new Date('2016-01-01T00:00:00');

    if (requestedDateObj > maxForecastDate) {
        setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
        updateRainfallStatus(dayIndex, 'idle', '-');
        return 'idle_due_to_proactive_range';
    }

    if (requestedDateObj < minApiDate) {
        setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
        updateRainfallStatus(dayIndex, 'idle', '-');
        return 'idle_due_to_proactive_range';
    }

    updateRainfallStatus(dayIndex, 'loading');

    const lat = -23.5144;
    const lon = -46.8446;
    const apiUrl = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&daily=precipitation_sum&timezone=America/Sao_Paulo&start_date=${dateStrYYYYMMDD}&end_date=${dateStrYYYYMMDD}`;

    try {
        const response = await fetch(apiUrl);
        if (!response.ok) {
            let apiErrorMsg = `API error: ${response.status}`;
            let isOutOfRangeErrorByApi = false;
            let isRateLimitErrorByApi = response.status === 429;

            try {
                const errorJson = await response.json();
                if (errorJson && errorJson.reason) {
                    apiErrorMsg += ` - ${errorJson.reason}`;
                    if (errorJson.reason.toLowerCase().includes("out of allowed range") || errorJson.reason.toLowerCase().includes("is out of range for daily data")) {
                        isOutOfRangeErrorByApi = true;
                    }
                }
            } catch (e) { /* ignore parsing error of error response */ }

            if (isRateLimitErrorByApi) {
                console.warn(`Rate limit error for day ${dayIndex} (${dateStrYYYYMMDD}): ${apiErrorMsg}`);
                updateRainfallStatus(dayIndex, 'error', 'RL');
                return 'rate_limited';
            }
            if (isOutOfRangeErrorByApi) {
                setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
                updateRainfallStatus(dayIndex, 'idle', '-');
                return 'idle_due_to_api_range';
            }
            console.error(`API error for day ${dayIndex} (${dateStrYYYYMMDD}): ${apiErrorMsg}`);
            updateRainfallStatus(dayIndex, 'error');
            return 'error';
        }
        const jsonData = await response.json();

        if (jsonData && jsonData.daily && Array.isArray(jsonData.daily.precipitation_sum) && jsonData.daily.precipitation_sum.length > 0) {
            const precipitation = jsonData.daily.precipitation_sum[0];
            if (precipitation !== null && typeof precipitation === 'number') {
                setInputValue(`indice_pluv_valor_day${dayIndex}`, precipitation.toFixed(1), page1Container);

                if (requestedDateObj > systemToday) {
                    updateRainfallStatus(dayIndex, 'success', 'P');
                } else {
                    updateRainfallStatus(dayIndex, 'success');
                }

                let newTempoValue = "";
                if (precipitation < 1) {
                    newTempoValue = "B";
                } else if (precipitation < 20) {
                    newTempoValue = "L";
                } else {
                    newTempoValue = "F";
                }

                const turnosToUpdate = [
                    { turnoIdBase: 't1', cellIdBase: 'turno1_control_cell' },
                    { turnoIdBase: 't2', cellIdBase: 'turno2_control_cell' }
                ];

                turnosToUpdate.forEach(turnoConfig => {
                    const cellId = `${turnoConfig.cellIdBase}_day${dayIndex}`;
                    const updater = shiftControlUpdaters.get(cellId);
                    const currentTrabalhoValue = getInputValue(`trabalho_${turnoConfig.turnoIdBase}_day${dayIndex}_hidden`, page1Container);

                    if (updater) {
                        updater.update(newTempoValue, currentTrabalhoValue);
                    }
                });


                markFormAsDirty();
                return 'success';
            } else {
                setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
                updateRainfallStatus(dayIndex, 'idle', '-');
                return 'idle';
            }
        } else {
            console.warn(`Precipitation data array not found or empty in API response for ${dateStrYYYYMMDD}. Response:`, jsonData);
            setInputValue(`indice_pluv_valor_day${dayIndex}`, '', page1Container);
            updateRainfallStatus(dayIndex, 'idle', '-');
            return 'idle';
        }
    } catch (error) {
        console.error(`Network or other error fetching rainfall data for day ${dayIndex} (${dateStrYYYYMMDD}):`, error);
        updateRainfallStatus(dayIndex, 'error');
        return 'error';
    }
}


async function fetchAllRainfallDataForAllDaysOnClick() {
    const fetchButton = getElementByIdSafe('fetchRainfallButton');
    if (!fetchButton) {
        console.error("Fetch rainfall button not found.");
        return;
    }

    if (isGloballyRateLimited && Date.now() < globalRateLimitRetryAfterTimestamp) {
        updateFetchMessage(`Limite de API atingido. Tente novamente após ${new Date(globalRateLimitRetryAfterTimestamp).toLocaleTimeString()}.`, true);
        return;
    }
    isGloballyRateLimited = false;

    fetchButton.disabled = true;

    const totalDaysToFetch = rdoDayCounter + 1;
    let processedDaysCount = 0;

    if (rainfallProgressBarWrapper && rainfallProgressBar && rainfallProgressText && rainfallProgressDetailsText && totalDaysToFetch > 0) {
        rainfallProgressBar.style.width = '0%';
        rainfallProgressText.textContent = `0/${totalDaysToFetch}`;
        rainfallProgressDetailsText.textContent = 'Iniciando busca...';
        rainfallProgressBarWrapper.style.display = 'block';
        updateFetchMessage(null);
    } else {
        updateFetchMessage("Buscando dados... Essa ação pode demorar um pouco.");
    }

    let successCount = 0;
    let errorCount = 0;
    let rateLimitHitInThisRun = false;

    for (let d = 0; d <= rdoDayCounter; d++) {
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;
        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector(`.rdo-day-wrapper`);
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (!rdoDayWrapper) continue;

        const page1Container = rdoDayWrapper.querySelector('.rdo-container');
        if (!page1Container) continue;

        const dateEl = getElementByIdSafe(`data_day${d}`, page1Container);
        const dateStr = dateEl ? dateEl.value : '';

        if (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
            try {
                const fetchStatus = await fetchRainfallData(dateStr, d);

                switch (fetchStatus) {
                    case 'success':
                        successCount++;
                        break;
                    case 'rate_limited':
                        rateLimitHitInThisRun = true;
                        isGloballyRateLimited = true;
                        globalRateLimitRetryAfterTimestamp = Date.now() + GLOBAL_RATE_LIMIT_COOLDOWN_MS;
                        errorCount++;
                        break;
                    case 'error':
                        errorCount++;
                        break;
                }

            } catch (e) {
                console.error(`Unhandled exception during fetchRainfallData call for day ${d}:`, e);
                updateRainfallStatus(d, 'error');
                errorCount++;
            }
        } else {
            console.warn(`Invalid or missing date for day ${d}. Skipping rainfall fetch. Date was: ${dateStr}`);
            updateRainfallStatus(d, 'idle', 'Data Inválida');
        }

        processedDaysCount++;
        if (rainfallProgressBarWrapper && rainfallProgressBar && rainfallProgressText && rainfallProgressDetailsText && totalDaysToFetch > 0) {
            const progressPercent = (processedDaysCount / totalDaysToFetch) * 100;
            rainfallProgressBar.style.width = `${progressPercent}%`;
            rainfallProgressText.textContent = `${processedDaysCount}/${totalDaysToFetch}`;

            if (rateLimitHitInThisRun) {
            } else if (processedDaysCount < totalDaysToFetch) {
                rainfallProgressDetailsText.textContent = `Processando dia ${processedDaysCount + 1} de ${totalDaysToFetch}...`;
            } else {
                rainfallProgressDetailsText.textContent = 'Finalizando...';
            }
        }

        if (rateLimitHitInThisRun) {
            break;
        }
        if (d < rdoDayCounter) {
            await delay(REQUEST_DELAY_MS);
        }
    }

    fetchButton.disabled = false;
    let finalUserMessage = null;
    let isFinalUserMessageError = false;

    if (rateLimitHitInThisRun) {
        finalUserMessage = `Busca interrompida. Limite de API atingido. Tente o restante após ${new Date(globalRateLimitRetryAfterTimestamp).toLocaleTimeString()}.`;
        isFinalUserMessageError = true;
    } else if (errorCount > 0) {
        finalUserMessage = `Busca concluída. ${successCount} dados obtidos, ${errorCount} erros.`;
        isFinalUserMessageError = true;
    } else if (successCount > 0 || rdoDayCounter >=0 ) {
        finalUserMessage = "Busca de dados pluviométricos concluída.";
    } else {
        finalUserMessage = "Nenhum dado pluviométrico para buscar ou dados não disponíveis.";
    }
    updateFetchMessage(finalUserMessage, isFinalUserMessageError);


    if (rainfallProgressDetailsText) {
        if (rateLimitHitInThisRun) {
            rainfallProgressDetailsText.textContent = "Interrompido (limite API).";
        } else if (errorCount > 0) {
            rainfallProgressDetailsText.textContent = `Concluído com ${errorCount} erro(s).`;
        } else if (successCount > 0 || totalDaysToFetch > 0) {
            rainfallProgressDetailsText.textContent = "Concluído com sucesso!";
        } else {
            rainfallProgressDetailsText.textContent = "Nenhum dado para buscar.";
        }
    }

    if (rainfallProgressBarWrapper && totalDaysToFetch > 0) {
        setTimeout(() => {
            if (rainfallProgressBarWrapper) {
                rainfallProgressBarWrapper.style.display = 'none';
            }
        }, 5000);
    }
}

// --- Efetivo Spreadsheet Loading ---

const commonWordsToFilter = ['de', 'do', 'da', 'a', 'o', 'e', 'para', 'c', 's', 'geral', 'civil'];


function normalizeStringForMatching(str) {
    if (!str) return "";
    let normalized = str.toLowerCase();
    normalized = normalized.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    normalized = normalized.replace(/\bmec\s*[\ufffd]\s*nico\b/g, "mecanico");

    normalized = normalized.replace(/\bop\s+retro\s+escav\b/g, "operador retroescavadeira");
    normalized = normalized.replace(/\bretro\s+escav\b/g, "retroescavadeira");
    normalized = normalized.replace(/\baux\s+serv\s+gerais\b/g, "auxiliar servico geral");
    normalized = normalized.replace(/\bserv\s+gerais\b/g, "servico geral");
    normalized = normalized.replace(/\btec\s+seg\s+trab\b/g, "tecnico seguranca trabalho");
    normalized = normalized.replace(/\bseg\s+trab\b/g, "seguranca trabalho");
    normalized = normalized.replace(/\bassist\s+adm\b/g, "assistente administrativo");

    normalized = normalized.replace(/\bauxiliar\s+tec\s+eng\b/g, "auxiliar engenharia");
    normalized = normalized.replace(/\bengenheiro\s+civil\b/g, "engenheiro planejamento");
    normalized = normalized.replace(/\bengenheiro\s+qualidade\b/g, "engenheiro producao qualidade");

    normalized = normalized.replace(/\baux\s+tec\s+eng\b/g, "auxiliar tecnico engenharia");
    normalized = normalized.replace(/\btec\s+eng\b/g, "tecnico engenharia");
    normalized = normalized.replace(/\bop\s+trat\s+esgoto\b/g, "operador tratamento esgoto");
    normalized = normalized.replace(/\btrat\s+esgoto\b/g, "tratamento esgoto");
    normalized = normalized.replace(/\baux\s+de\s+eng\b/g, "auxiliar engenharia");
    normalized = normalized.replace(/\baux\s+de\s+engenharia\b/g, "auxiliar engenharia");


    normalized = normalized.replace(/[.,\/()"'\ufffd]/g, " ");

    const wordMappings = {
        "enc": "encarregado", "encarr": "encarregado",
        "obras": "obra",
        "aux": "auxiliar", "auxil": "auxiliar",
        "serv": "servico", "servs": "servico", "servicos": "servico",
        "op": "operador", "operad": "operador",
        "eng": "engenheiro", "engenh": "engenheiro", "engenharia": "engenharia",
        "elet": "eletrico", "eletric": "eletrico",
        "mec": "mecanico", "mecanic": "mecanico", "mecanica": "mecanico",
        "tec": "tecnico", "tecn": "tecnico",
        "seg": "seguranca",
        "trab": "trabalho", "trabs": "trabalho",
        "adm": "administrativo", "admin": "administrativo",
        "prod": "producao", "produc": "producao",
        "plan": "planejamento",
        "qual": "qualidade",
        "mont": "montador", "montad": "montador",
        "est": "estrutura", "estrut": "estrutura",
        "maq": "maquina", "maqs": "maquina", "maquinas": "maquina",
        "retroescavadeira": "maquina",
        "topog": "topografia", "topogr": "topografo",
        "cad": "cadista",
        "des": "desenhista",
        "coord": "coordenador",
        "contratos": "contrato", "contrato": "contrato",
        "contab": "contabilidade",
        "lab": "laboratorio", "laborat": "laboratorista",
        "sup": "supervisor", "superv": "supervisor",
        "aj": "ajudante", "ajud": "ajudante", "ajudante": "ajudante",
        "gerais": "geral",
        "asg": "auxiliar servico geral",
        "gte": "gerente", "gerente": "gerente",
        "assist": "assistente",
        "i": "i", "ii": "ii", "iii": "iii", "iv": "iv", "v": "v",
        "tratamento": "tratamento",
        "esgoto": "esgoto"
    };

    let words = normalized.split(/\s+/).filter(w => w.length > 0);
    words = words.map(word => wordMappings[word] || word);
    words = words.filter(word => !commonWordsToFilter.includes(word));

    normalized = words.join(" ");
    normalized = normalized.replace(/[^a-z0-9\s]/g, "");
    normalized = normalized.replace(/\s+/g, " ");
    return normalized.trim();
}


function parseCSV(csvString) {
    const lines = csvString.split(/\r\n|\n/);
    if (lines.length < 1) {
        updateSaveProgressMessage("Arquivo CSV vazio.", true);
        return [];
    }

    const headerLine = lines[0].trim();
    if (!headerLine) {
        updateSaveProgressMessage("Cabeçalho do CSV não encontrado.", true);
        return [];
    }

    let headers = [];
    let detectedDelimiter = '';
    const normalizedFuncaoVariants = ['função', 'funcao', 'fun��o', 'cargo'].map(normalizeStringForMatching);
    const commaHeadersOriginal = headerLine.split(',');
    const commaHeadersNormalized = commaHeadersOriginal.map(h => normalizeStringForMatching(h.trim()));
    const dataCommaIndex = commaHeadersNormalized.indexOf(normalizeStringForMatching('data'));
    const funcaoCommaIndex = commaHeadersNormalized.findIndex(h => normalizedFuncaoVariants.includes(h));

    if (dataCommaIndex !== -1 && funcaoCommaIndex !== -1) {
        headers = commaHeadersOriginal.map(h => h.trim());
        detectedDelimiter = ',';
    } else {
        const semicolonHeadersOriginal = headerLine.split(';');
        const semicolonHeadersNormalized = semicolonHeadersOriginal.map(h => normalizeStringForMatching(h.trim()));
        const dataSemicolonIndex = semicolonHeadersNormalized.indexOf(normalizeStringForMatching('data'));
        const funcaoSemicolonIndex = semicolonHeadersNormalized.findIndex(h => normalizedFuncaoVariants.includes(h));

        if (dataSemicolonIndex !== -1 && funcaoSemicolonIndex !== -1) {
            headers = semicolonHeadersOriginal.map(h => h.trim());
            detectedDelimiter = ';';
        } else {
            console.error("CSV headers 'Data' or 'Função'/'Cargo' (or variants) not found. Header line was:", headerLine);
            updateSaveProgressMessage("Erro no cabeçalho do CSV. Verifique 'Data' e 'Função'.", true);
            return [];
        }
    }

    const finalDataHeaderIndex = headers.map(h => normalizeStringForMatching(h)).indexOf(normalizeStringForMatching('data'));
    const finalFuncaoHeaderIndex = headers.map(h => normalizeStringForMatching(h)).findIndex(hNorm => normalizedFuncaoVariants.includes(hNorm));

    if (finalDataHeaderIndex === -1 || finalFuncaoHeaderIndex === -1) {
        console.error("Internal error: CSV headers 'Data' or 'Função' resolved to -1 after delimiter detection. Detected Original Headers:", headers);
        updateSaveProgressMessage("Erro interno ao processar cabeçalhos do CSV.", true);
        return [];
    }

    const data = [];
    for (let i = 1; i < lines.length; i++) {
        const lineContent = lines[i].trim();
        if (!lineContent) continue;
        const values = lineContent.split(detectedDelimiter);
        if (values.length >= Math.max(finalDataHeaderIndex, finalFuncaoHeaderIndex) + 1) {
            const entry = {};
            entry['data'] = (values[finalDataHeaderIndex] || "").trim();
            entry['funcao'] = (values[finalFuncaoHeaderIndex] || "").trim();
            if (entry['data'] && entry['funcao']) {
                 data.push(entry);
            }
        } else {
             console.warn(`Skipping malformed CSV line (expected at least ${Math.max(finalDataHeaderIndex, finalFuncaoHeaderIndex) + 1} values, got ${values.length} using delimiter '${detectedDelimiter}'): ${lineContent}`);
        }
    }
    return data;
}


async function populateEfetivoFromCSVData(csvData) {
    if (csvData.length === 0) {
        return;
    }

    const aggregatedData = new Map();
    const csvDatesNotFoundInRdoOrInvalid = new Set();
    const csvFuncoesNotFoundForProcessedDate = new Map();

    for (const entry of csvData) {
        const dataStrDDMMYYYY = (entry['data'] || "").trim();
        const funcaoFromCsvOriginal = (entry['funcao'] || "").trim();
        if (!dataStrDDMMYYYY || !funcaoFromCsvOriginal) {
            if (dataStrDDMMYYYY && !funcaoFromCsvOriginal) console.warn(`CSV entry for date ${dataStrDDMMYYYY} has empty função.`);
            else if (!dataStrDDMMYYYY && funcaoFromCsvOriginal) console.warn(`CSV entry for função ${funcaoFromCsvOriginal} has empty data.`);
            continue;
        }
        const funcaoFromCsvNormKey = normalizeStringForMatching(funcaoFromCsvOriginal);
        if(!funcaoFromCsvNormKey) continue;
        const parsedDate = parseDateDDMMYYYY(dataStrDDMMYYYY);
        if (!parsedDate) {
            csvDatesNotFoundInRdoOrInvalid.add(dataStrDDMMYYYY);
            continue;
        }
        const dateYYYYMMDD = formatDateToYYYYMMDD(parsedDate);
        if (!aggregatedData.has(dateYYYYMMDD)) {
            aggregatedData.set(dateYYYYMMDD, new Map());
        }
        const funcoesForDate = aggregatedData.get(dateYYYYMMDD);
        funcoesForDate.set(funcaoFromCsvNormKey, (funcoesForDate.get(funcaoFromCsvNormKey) || 0) + 1);
    }

    if (aggregatedData.size === 0 && csvData.length > 0) {
        let errorMsg = "Nenhuma data válida ou função encontrada no CSV para processar.";
        if (csvDatesNotFoundInRdoOrInvalid.size > 0) {
            errorMsg = `Datas inválidas ou não encontradas no CSV. Verifique o formato DD/MM/YYYY.`;
        }
        updateSaveProgressMessage(errorMsg, true);
        return;
    }

    let rdosEffectivelyUpdatedCount = 0;
    let totalFuncaoAssignments = 0;
    const rdoDatesProcessedThisRun = new Set();
    if (csvData.length > 0 || aggregatedData.size > 0) {
        hasEfetivoBeenLoadedAtLeastOnce = true;
    }

    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWrapperCandidate1 = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;
        if (dayWrapperCandidate1) {
            rdoDayWrapper = dayWrapperCandidate1.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }
        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container');
        }
        if (!currentDayPage1Container) {
            console.warn(`populateEfetivoFromCSVData: Page 1 container for RDO day index ${d} not found.`);
            continue;
        }
        const dataDayEl = getElementByIdSafe(`data_day${d}`, currentDayPage1Container);
        const rdoDateStrYYYYMMDD = dataDayEl ? dataDayEl.value : '';
        if (!rdoDateStrYYYYMMDD || !aggregatedData.has(rdoDateStrYYYYMMDD)) {
            continue;
        }
        csvData.forEach(csvEntry => {
            const parsedCsvDate = parseDateDDMMYYYY(csvEntry['data']);
            if (parsedCsvDate && formatDateToYYYYMMDD(parsedCsvDate) === rdoDateStrYYYYMMDD) {
                csvDatesNotFoundInRdoOrInvalid.delete(csvEntry['data']);
            }
        });
        const funcoesToApplyOnThisRdo = aggregatedData.get(rdoDateStrYYYYMMDD);
        if (!funcoesToApplyOnThisRdo || funcoesToApplyOnThisRdo.size === 0) continue;
        let itemsAppliedToThisRdoDay = 0;
        const laborTables = currentDayPage1Container.querySelectorAll('.labor-table');
        laborTables.forEach(table => {
            table.querySelectorAll('.quantity-input').forEach(input => input.value = '');
        });
        for (const [funcaoCsvKeyNorm, count] of funcoesToApplyOnThisRdo) {
            let foundFuncaoInRdoDay = false;
            for (const table of laborTables) {
                const rows = table.querySelectorAll('tbody tr');
                for (const row of rows) {
                    const itemNameCell = row.querySelector('.item-name-cell');
                    const quantityInput = row.querySelector('.quantity-input');
                    if (itemNameCell && quantityInput) {
                        const itemNameFromTableOriginal = (itemNameCell.textContent || "").trim();
                        const itemNameFromTableNorm = normalizeStringForMatching(itemNameFromTableOriginal);
                        if (itemNameFromTableNorm === funcaoCsvKeyNorm) {
                            quantityInput.value = count.toString();
                            itemsAppliedToThisRdoDay++;
                            foundFuncaoInRdoDay = true;
                            break;
                        } else {
                            const rdoSignificantWords = itemNameFromTableNorm.split(" ").filter(w => w.length > 1 && !commonWordsToFilter.includes(w));
                            const csvSignificantWords = funcaoCsvKeyNorm.split(" ").filter(w => w.length > 1 && !commonWordsToFilter.includes(w));
                            let matchedByFallback = false;
                            if (rdoSignificantWords.length > 0 && csvSignificantWords.length > 0 && rdoSignificantWords.length === csvSignificantWords.length) {
                                const rdoWordSet = new Set(rdoSignificantWords);
                                const allCsvWordsFoundInRdo = csvSignificantWords.every(csvWord => rdoWordSet.has(csvWord));
                                if (allCsvWordsFoundInRdo) {
                                    matchedByFallback = true;
                                }
                            }
                            if (matchedByFallback) {
                                quantityInput.value = count.toString();
                                itemsAppliedToThisRdoDay++;
                                foundFuncaoInRdoDay = true;
                                break;
                            }
                        }
                    }
                }
                if (foundFuncaoInRdoDay) break;
            }
            if (!foundFuncaoInRdoDay) {
                if (!csvFuncoesNotFoundForProcessedDate.has(rdoDateStrYYYYMMDD)) {
                    csvFuncoesNotFoundForProcessedDate.set(rdoDateStrYYYYMMDD, new Set());
                }
                const originalFuncaoForMessage = csvData.find(entry => {
                    const originalCsvEntryFuncao = (entry['funcao'] || "").trim();
                    return normalizeStringForMatching(originalCsvEntryFuncao) === funcaoCsvKeyNorm;
                })?.['funcao'] || funcaoCsvKeyNorm;
                csvFuncoesNotFoundForProcessedDate.get(rdoDateStrYYYYMMDD).add(originalFuncaoForMessage);
            }
        }
        if (itemsAppliedToThisRdoDay > 0 || funcoesToApplyOnThisRdo.size > 0) {
            calculateSectionTotalsForDay(currentDayPage1Container, d);
            totalFuncaoAssignments += itemsAppliedToThisRdoDay;
            if (!rdoDatesProcessedThisRun.has(rdoDateStrYYYYMMDD)) {
                rdosEffectivelyUpdatedCount++;
                rdoDatesProcessedThisRun.add(rdoDateStrYYYYMMDD);
            }
        }
    }

    let message = "";
    let isError = false;

    if (totalFuncaoAssignments > 0) {
        markFormAsDirty();
        message = "Dados do CSV processados.";
    } else if (csvData.length > 0 && aggregatedData.size > 0 && rdosEffectivelyUpdatedCount > 0) {
        markFormAsDirty();
        message = "Dados do CSV processados. Alguns itens não corresponderam.";
    } else if (csvData.length > 0 && aggregatedData.size > 0 && rdosEffectivelyUpdatedCount === 0) {
        message = "Nenhuma data do CSV correspondeu aos RDOs.";
        isError = true;
    } else if (csvData.length > 0 && aggregatedData.size === 0 && csvDatesNotFoundInRdoOrInvalid.size === 0) {
        message = "Nenhum dado válido no CSV.";
        isError = true;
    }

    const notFoundDetailsForConsole = [];
    if (csvDatesNotFoundInRdoOrInvalid.size > 0) {
        notFoundDetailsForConsole.push(`Datas do CSV não encontradas nos RDOs ou em formato inválido (DD/MM/YYYY): ${Array.from(csvDatesNotFoundInRdoOrInvalid).join(', ')}.`);
    }
    csvFuncoesNotFoundForProcessedDate.forEach((funcoesSet, dateYYYYMMDD) => {
        const originalDateForMessage = csvData.find(entry => {
            const parsed = parseDateDDMMYYYY(entry['data']);
            return parsed && formatDateToYYYYMMDD(parsed) === dateYYYYMMDD;
        })?.['data'] || dateYYYYMMDD;
        notFoundDetailsForConsole.push(`Para a data ${originalDateForMessage} (nos RDOs): Funções do CSV não encontradas: ${Array.from(funcoesSet).join(', ')}.`);
    });

    if (notFoundDetailsForConsole.length > 0) {
        console.warn("Detalhes do carregamento de efetivo (CSV):\n" + notFoundDetailsForConsole.join("\n"));
        if (totalFuncaoAssignments === 0 && !message) {
             isError = true;
             message = "Dados do CSV não aplicados. Verifique o arquivo.";
        } else if (totalFuncaoAssignments > 0 && message === "Dados do CSV processados.") {
             message = "Dados do CSV processados. Alguns itens não corresponderam."; // Downgrade success message if there were misses
        }
    }


    if (message) {
        updateSaveProgressMessage(message, isError);
    } else if (csvData.length === 0) {
        // No action, no message if CSV was empty initially
    } else {
        // If no specific message was set but CSV had content, it implies parsing issues caught earlier (e.g., bad header)
        if (!globalSaveMessageContainer?.textContent?.includes("Cabeçalho")) { // Avoid overriding header error
             updateSaveProgressMessage("Arquivo CSV vazio ou inválido.", true);
        }
    }
    efetivoWasJustClearedByButton = false;
    updateAllEfetivoAlerts();
    updateClearEfetivoButtonVisibility();
}


function handleEfetivoFileUpload(event) {
    updateSaveProgressMessage(null);
    const input = event.target;
    if (!input.files || input.files.length === 0) {
        updateSaveProgressMessage("Nenhum arquivo selecionado.", true);
        return;
    }

    const file = input.files[0];
    if (file.type !== 'text/csv' && !file.name.toLowerCase().endsWith('.csv')) {
        updateSaveProgressMessage("Formato de arquivo inválido. Por favor, selecione um arquivo .csv.", true);
        input.value = '';
        return;
    }

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const csvString = e.target?.result;
            if (!csvString.trim()) {
                 updateSaveProgressMessage("Arquivo CSV vazio.", true);
                 input.value = '';
                 return;
            }
            const parsedData = parseCSV(csvString);
            if (parsedData.length > 0) {
                 await populateEfetivoFromCSVData(parsedData);
            } else if (csvString.trim() && !globalSaveMessageContainer?.textContent?.includes("Cabeçalho")) {
                // If parseCSV returned empty but there was content and no header error message shown by parseCSV
                updateSaveProgressMessage("Nenhuma linha de dados válida no CSV.", true);
            }
        } catch (error) {
            console.error("Erro ao processar arquivo CSV:", error);
            updateSaveProgressMessage(`Erro ao processar arquivo CSV: ${error.message}`, true);
        } finally {
            input.value = '';
        }
    };
    reader.onerror = () => {
        console.error("Erro ao ler o arquivo.");
        updateSaveProgressMessage("Erro ao ler o arquivo.", true);
        input.value = '';
    };
    reader.readAsText(file);
}

function triggerEfetivoFileUpload() {
    updateSaveProgressMessage(null);
    const fileUploadElement = getElementByIdSafe('efetivoSpreadsheetUpload');
    if (fileUploadElement) {
        fileUploadElement.click();
    } else {
        console.error("Elemento de upload de arquivo de efetivo não encontrado.");
        updateSaveProgressMessage("Erro: Funcionalidade de upload não está disponível.", true);
    }
}

function checkIfAnyEfetivoExists() {
    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;

        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container');
        }

        if (!currentDayPage1Container) {
            continue;
        }

        const quantityInputs = currentDayPage1Container.querySelectorAll('.labor-table .quantity-input');
        for (const input of quantityInputs) {
            const value = parseInt(input.value, 10);
            if (!isNaN(value) && value > 0) {
                return true;
            }
        }
    }
    return false;
}

function updateClearEfetivoButtonVisibility() {
    const clearButton = getElementByIdSafe('clearEfetivoButton');
    if (!clearButton) return;

    if (checkIfAnyEfetivoExists()) {
        clearButton.style.display = 'flex';
    } else {
        clearButton.style.display = 'none';
    }
}


async function clearAllEfetivo() {
    updateSaveProgressMessage(null);
    let initialCheckChangesMade = false;

    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;

        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }
        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container');
        }

        if (currentDayPage1Container) {
            const quantityInputs = currentDayPage1Container.querySelectorAll('.labor-table .quantity-input');
            for (const input of quantityInputs) {
                if (input.value !== '') {
                    initialCheckChangesMade = true;
                    break;
                }
            }
        }
        if (initialCheckChangesMade) break;
    }

    if (initialCheckChangesMade) {
        efetivoWasJustClearedByButton = true;
        hasEfetivoBeenLoadedAtLeastOnce = true;
    } else {
        efetivoWasJustClearedByButton = false;
    }

    let actualChangesInLoop = false;
    for (let d = 0; d <= rdoDayCounter; d++) {
        let currentDayPage1Container = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        let rdoDayWrapper = null;

        if (dayWithButtonLayoutWrapper) {
            rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            rdoDayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (rdoDayWrapper) {
            currentDayPage1Container = rdoDayWrapper.querySelector('.rdo-container');
        }

        if (!currentDayPage1Container) {
            console.warn(`clearAllEfetivo: Page 1 container for RDO day index ${d} not found.`);
            continue;
        }

        const laborTables = currentDayPage1Container.querySelectorAll('.labor-table');
        laborTables.forEach(table => {
            table.querySelectorAll('.quantity-input').forEach(input => {
                if (input.value !== '') {
                    input.value = '';
                    actualChangesInLoop = true;
                }
            });
        });
        calculateSectionTotalsForDay(currentDayPage1Container, d);
    }

    if (actualChangesInLoop) {
        markFormAsDirty();
        updateSaveProgressMessage("Efetivo limpo.", false);
    } else {
        updateSaveProgressMessage("Nenhum efetivo para limpar.", false);
    }
    updateClearEfetivoButtonVisibility();
}


async function handleCopySaturdayEfetivo(targetSundayDayIndex) {
    try {
        updateSaveProgressMessage(null);
        let targetSundayPage1Container = null;
        const targetSundayDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${targetSundayDayIndex}`);

        if (targetSundayDayWithButtonLayoutWrapper) {
            const rdoWrapper = targetSundayDayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
            if (rdoWrapper) targetSundayPage1Container = rdoWrapper.querySelector('.rdo-container');
        } else {
            const rdoWrapper = getElementByIdSafe(`rdo_day_${targetSundayDayIndex}_wrapper`);
            if (rdoWrapper) targetSundayPage1Container = rdoWrapper.querySelector('.rdo-container');
        }

        if (!targetSundayPage1Container) {
            console.error(`handleCopySaturdayEfetivo: Target Sunday Page 1 container for day ${targetSundayDayIndex} not found.`);
            updateSaveProgressMessage("Erro: Domingo de destino não encontrado.", true);
            return;
        }

        const sundayDateEl = getElementByIdSafe(`data_day${targetSundayDayIndex}`, targetSundayPage1Container);
        const sundayDateStr = sundayDateEl ? sundayDateEl.value : '';

        if (!sundayDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(sundayDateStr)) {
            console.error(`handleCopySaturdayEfetivo: Invalid date for target Sunday day ${targetSundayDayIndex}: ${sundayDateStr}`);
            updateSaveProgressMessage("Erro: Data do domingo inválida.", true);
            return;
        }

        const sundayDate = new Date(sundayDateStr + 'T00:00:00');
        if (isNaN(sundayDate.getTime())) {
            console.error(`handleCopySaturdayEfetivo: Could not parse date for target Sunday day ${targetSundayDayIndex}: ${sundayDateStr}`);
            updateSaveProgressMessage("Erro ao processar data do domingo.", true);
            return;
        }

        const saturdayDate = new Date(sundayDate);
        saturdayDate.setDate(sundayDate.getDate() - 1);
        const saturdayDateStrYYYYMMDD = formatDateToYYYYMMDD(saturdayDate);

        let sourceSaturdayDayIndex = -1;
        let sourceSaturdayPage1Container = null;

        for (let d = 0; d <= rdoDayCounter; d++) {
            let currentDayPage1Container = null;
            const currentDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
            if (currentDayWithButtonLayoutWrapper) {
                const rdoWrapper = currentDayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
                if (rdoWrapper) currentDayPage1Container = rdoWrapper.querySelector('.rdo-container');
            } else {
                const rdoWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
                if (rdoWrapper) currentDayPage1Container = rdoWrapper.querySelector('.rdo-container');
            }

            if (currentDayPage1Container) {
                const currentDateEl = getElementByIdSafe(`data_day${d}`, currentDayPage1Container);
                const currentDateValue = currentDateEl ? currentDateEl.value : '';

                if (currentDateValue === saturdayDateStrYYYYMMDD) {
                    sourceSaturdayDayIndex = d;
                    sourceSaturdayPage1Container = currentDayPage1Container;
                    break;
                }
            }
        }

        if (sourceSaturdayDayIndex === -1 || !sourceSaturdayPage1Container) {
            updateSaveProgressMessage("Sábado anterior não encontrado.", true);
            return;
        }

        const sourceLaborSection = sourceSaturdayPage1Container.querySelector('.labor-section');
        const targetLaborSection = targetSundayPage1Container.querySelector('.labor-section');

        if (!sourceLaborSection || !targetLaborSection) {
            console.error("handleCopySaturdayEfetivo: Labor sections not found in source or target.");
            updateSaveProgressMessage("Erro: Seções de efetivo não encontradas.", true);
            return;
        }

        const sourceQuantityInputs = sourceLaborSection.querySelectorAll('.labor-table .quantity-input');
        const targetQuantityInputs = targetLaborSection.querySelectorAll('.labor-table .quantity-input');

        if (sourceQuantityInputs.length !== targetQuantityInputs.length) {
            console.warn("handleCopySaturdayEfetivo: Mismatch in number of labor inputs between Saturday and Sunday. Copy might be incomplete.");
            updateSaveProgressMessage("Aviso: Discrepância nos campos de efetivo.", true);
        }

        let copied = false;
        sourceQuantityInputs.forEach((sourceInput, index) => {
            if (targetQuantityInputs[index]) {
                targetQuantityInputs[index].value = sourceInput.value;
                copied = true;
            }
        });

        if (copied) {
            efetivoWasJustClearedByButton = false;
            hasEfetivoBeenLoadedAtLeastOnce = true;
            calculateSectionTotalsForDay(targetSundayPage1Container, targetSundayDayIndex);
            markFormAsDirty();
            updateSaveProgressMessage("Efetivo copiado.");
        } else {
            updateSaveProgressMessage("Nenhum efetivo para copiar.", true);
        }
        updateAllEfetivoAlerts();
        updateClearEfetivoButtonVisibility();

    } catch (error) {
        console.error("Error in handleCopySaturdayEfetivo:", error);
        updateSaveProgressMessage("Ocorreu um erro inesperado ao copiar.", true);
    }
}

function updateEfetivoCopyButtonVisibility(dayIndex, originalRdoDayWrapper) {
    if (!originalRdoDayWrapper) {
        console.warn(`Original RDO day wrapper not provided for efetivo button visibility, day ${dayIndex}`);
        return originalRdoDayWrapper;
    }

    const page1ContainerForDate = originalRdoDayWrapper.querySelector(`#rdo_day_${dayIndex}_page_0`);
    let dayOfWeek = '';
    if (page1ContainerForDate) {
        const dateInputForDay = getElementByIdSafe(`data_day${dayIndex}`, page1ContainerForDate);
        if (dateInputForDay) {
            const dateVal = dateInputForDay.value;
            if (dateVal && /^\d{4}-\d{2}-\d{2}$/.test(dateVal)) {
                 const dateObj = new Date(dateVal + 'T00:00:00');
                if (!isNaN(dateObj.getTime())) {
                    dayOfWeek = getDayOfWeekFromDate(dateObj);
                }
            }
        }
    }

    const dayWithButtonContainerId = `day_with_button_container_day${dayIndex}`;
    const sidebarContainerId = `efetivo_actions_sidebar_day${dayIndex}`;
    const buttonId = `copySaturdayEfetivoButton_day${dayIndex}`;

    if (dayOfWeek === 'domingo') {
        let dayWithButtonContainerToUse;

        if (originalRdoDayWrapper.parentElement &&
            originalRdoDayWrapper.parentElement.id === dayWithButtonContainerId &&
            originalRdoDayWrapper.parentElement.classList.contains('day-with-button-container')) {
            dayWithButtonContainerToUse = originalRdoDayWrapper.parentElement;
        } else {
            dayWithButtonContainerToUse = document.createElement('div');
            dayWithButtonContainerToUse.id = dayWithButtonContainerId;
            dayWithButtonContainerToUse.className = 'day-with-button-container';
            dayWithButtonContainerToUse.appendChild(originalRdoDayWrapper);
        }

        let sidebarContainer = dayWithButtonContainerToUse.querySelector(`#${sidebarContainerId}`);
        if (!sidebarContainer) {
            sidebarContainer = document.createElement('div');
            sidebarContainer.id = sidebarContainerId;
            sidebarContainer.className = 'efetivo-actions-sidebar';
            if (originalRdoDayWrapper.nextSibling !== sidebarContainer) {
                 dayWithButtonContainerToUse.appendChild(sidebarContainer);
            }
        }

        let copyButton = sidebarContainer.querySelector(`#${buttonId}`);
        if (!copyButton) {
            copyButton = document.createElement('button');
            copyButton.id = buttonId;
            copyButton.textContent = 'Copiar Efetivo do Sábado Anterior';
            copyButton.className = 'header-action-button copy-saturday-button';
            copyButton.type = 'button';
            copyButton.addEventListener('click', () => handleCopySaturdayEfetivo(dayIndex));
            sidebarContainer.appendChild(copyButton);
        }

        originalRdoDayWrapper.style.margin = '0';
        originalRdoDayWrapper.style.marginBottom = '0';
        return dayWithButtonContainerToUse;
    } else {
        originalRdoDayWrapper.style.margin = '';
        originalRdoDayWrapper.style.marginBottom = '';
        return originalRdoDayWrapper;
    }
}

function findNextDayElementInDom(currentDayIndex, parentWrapper) {
    const nextDayIndex = currentDayIndex + 1;
    let nextElement = parentWrapper.querySelector(`#day_with_button_container_day${nextDayIndex}`);
    if (!nextElement) {
        nextElement = parentWrapper.querySelector(`#rdo_day_${nextDayIndex}_wrapper`);
    }
    return nextElement;
}

async function updateAllDateRelatedFieldsFromDay0() {
    const allDaysWrapper = getElementByIdSafe('rdo-all-days-wrapper');
    if (!allDaysWrapper) {
        console.error("updateAllDateRelatedFieldsFromDay0: #rdo-all-days-wrapper not found. Global date synchronization cannot proceed.");
        return;
    }

    const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
    if (!day0Page1Container) {
        console.error("updateAllDateRelatedFieldsFromDay0: Day 0 Page 1 container (#rdo_day_0_page_0) not found. Global date synchronization cannot proceed.");
        return;
    }

    const dataDay0Input = day0Page1Container.querySelector('#data_day0');
    if (!dataDay0Input) {
        console.error("updateAllDateRelatedFieldsFromDay0: Master date input #data_day0 not found within #rdo_day_0_page_0. Date updates aborted.");
        return;
    }

    let day0DataStr = dataDay0Input.value.trim();
    const day0DefaultDataStr = dataDay0Input.defaultValue.trim();

    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(day0DataStr)) {
        if (dateRegex.test(day0DefaultDataStr)) {
            day0DataStr = day0DefaultDataStr;
            dataDay0Input.value = day0DataStr;
        } else {
            const today = new Date();
            day0DataStr = formatDateToYYYYMMDD(new Date(today.getFullYear(), today.getMonth(), 1));
            dataDay0Input.value = day0DataStr;
            console.warn(`updateAllDateRelatedFieldsFromDay0: Invalid date in #data_day0. Using current month's first day: ${day0DataStr}`);
        }
    }

    const parts = day0DataStr.split('-');
    if (parts[2] !== '01') {
        day0DataStr = `${parts[0]}-${parts[1]}-01`;
        dataDay0Input.value = day0DataStr;
        console.warn(`updateAllDateRelatedFieldsFromDay0: #data_day0 was not 1st of month. Corrected to: ${day0DataStr}`);
    }


    const day0PrazoStr = getInputValue('prazo_day0', day0Page1Container);
    const day0InicioObraStr = getInputValue('inicio_obra_day0', day0Page1Container);
    const day0Contratada = getInputValue('contratada_day0', day0Page1Container);
    const day0ContratoNum = getInputValue('contrato_num_day0', day0Page1Container);
    const day0RdoNumFull = getInputValue('numero_rdo_day0', day0Page1Container);
    const day0RdoNumParts = day0RdoNumFull.split('-');
    const day0BaseRdoNumStr = day0RdoNumParts[0];
    const day0BaseRdoNumInt = parseInt(day0BaseRdoNumStr, 10);

    const baseDateForDay0 = new Date(day0DataStr + 'T00:00:00');
    if (isNaN(baseDateForDay0.getTime())) {
        console.error(`updateAllDateRelatedFieldsFromDay0: Invalid base date calculated for Day 0 from string: "${day0DataStr}T00:00:00". Date updates aborted.`);
        return;
    }

    const inicioObraDateObj = parseDateDDMMYYYY(day0InicioObraStr);
    const prazoDateObj = parseDateDDMMYYYY(day0PrazoStr);

    for (let d = 0; d <= rdoDayCounter; d++) {
        const daySuffix = `_day${d}`;

        let liveElementForDayD = document.getElementById(`day_with_button_container_day${d}`) || document.getElementById(`rdo_day_${d}_wrapper`);
        let contentWrapperForDayD;

        if (liveElementForDayD && liveElementForDayD.id === `day_with_button_container_day${d}`) {
            contentWrapperForDayD = liveElementForDayD.querySelector(`.rdo-day-wrapper[id="rdo_day_${d}_wrapper"]`);
        } else if (liveElementForDayD && liveElementForDayD.id === `rdo_day_${d}_wrapper`) {
            contentWrapperForDayD = liveElementForDayD;
        } else {
            contentWrapperForDayD = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (!contentWrapperForDayD || contentWrapperForDayD.id !== `rdo_day_${d}_wrapper`) {
            console.warn(`updateAllDateRelatedFieldsFromDay0: Content wrapper (rdo_day_${d}_wrapper) for day index ${d} could not be confirmed. Skipping processing for this day.`);
            continue;
        }

        const currentPage1Container = contentWrapperForDayD.querySelector(`#rdo_day_${d}_page_0`);
        if (!currentPage1Container) {
            console.warn(`updateAllDateRelatedFieldsFromDay0: Page 1 container (rdo_day_${d}_page_0) for day index ${d} not found within content wrapper. Skipping detailed updates for this day.`);
            continue;
        }

        const currentDateForThisDay = new Date(baseDateForDay0.getTime());
        currentDateForThisDay.setDate(baseDateForDay0.getDate() + d);
        const newDateStrYYYYMMDD = formatDateToYYYYMMDD(currentDateForThisDay);
        const newDayOfWeek = getDayOfWeekFromDate(currentDateForThisDay);

        setInputValue(`data${daySuffix}`, newDateStrYYYYMMDD, currentPage1Container);
        setInputValue(`dia_semana${daySuffix}`, newDayOfWeek, currentPage1Container);

        let decorridosStr = "N/A", restantesStr = "N/A";
        if (inicioObraDateObj) {
            const decorridosVal = calculateDaysBetween(inicioObraDateObj, currentDateForThisDay) + 1;
            decorridosStr = !isNaN(decorridosVal) && decorridosVal >= 0 ? decorridosVal.toString() : "0";
            if (prazoDateObj) {
                const totalPrazoDays = calculateDaysBetween(inicioObraDateObj, prazoDateObj) + 1;
                if (!isNaN(totalPrazoDays) && totalPrazoDays >= 0) {
                    const restantesVal = totalPrazoDays - decorridosVal;
                    restantesStr = !isNaN(restantesVal) && restantesVal >= 0 ? restantesVal.toString() : "0";
                }
            }
        }
        setInputValue(`decorridos${daySuffix}`, decorridosStr, currentPage1Container);
        setInputValue(`restantes${daySuffix}`, restantesStr, currentPage1Container);

        if (d > 0) {
            setInputValue(`contratada${daySuffix}`, day0Contratada, currentPage1Container);
            setInputValue(`contrato_num${daySuffix}`, day0ContratoNum, currentPage1Container);
            setInputValue(`prazo${daySuffix}`, day0PrazoStr, currentPage1Container);
            setInputValue(`inicio_obra${daySuffix}`, day0InicioObraStr, currentPage1Container);
        }

        if (!isNaN(day0BaseRdoNumInt)) {
            const currentDayRdoBaseNum = day0BaseRdoNumInt + d;
            setInputValue(`numero_rdo${daySuffix}`, `${currentDayRdoBaseNum}-A`, currentPage1Container);
        } else if (d > 0) {
            setInputValue(`numero_rdo${daySuffix}`, `-A`, currentPage1Container);
        }

        contractFieldMappingsBase.forEach(fieldPair => {
            const masterValue = getInputValue(fieldPair.masterIdBase + daySuffix, currentPage1Container);
            fieldPair.slaveIdBases.forEach(slaveIdBase => {
                setInputValue(slaveIdBase + daySuffix, masterValue, contentWrapperForDayD);
            });
        });

        const currentMasterRdoNumFull = getInputValue(`numero_rdo${daySuffix}`, currentPage1Container);
        const currentMasterRdoNumParts = currentMasterRdoNumFull.split('-');
        const currentMasterBaseRdoStr = currentMasterRdoNumParts[0];
        if (currentMasterBaseRdoStr) {
            setInputValue(`numero_rdo_pt${daySuffix}`, `${currentMasterBaseRdoStr}-B`, contentWrapperForDayD);
            setInputValue(`numero_rdo_pt2${daySuffix}`, `${currentMasterBaseRdoStr}-C`, contentWrapperForDayD);
        }

        const newTopLevelElementForDayD = updateEfetivoCopyButtonVisibility(d, contentWrapperForDayD);

        if (liveElementForDayD === newTopLevelElementForDayD) {
            if (allDaysWrapper && newTopLevelElementForDayD && (!newTopLevelElementForDayD.parentElement || newTopLevelElementForDayD.parentElement !== allDaysWrapper)) {
                const nextDayElement = findNextDayElementInDom(d, allDaysWrapper);
                if (nextDayElement) {
                    allDaysWrapper.insertBefore(newTopLevelElementForDayD, nextDayElement);
                } else {
                    allDaysWrapper.appendChild(newTopLevelElementForDayD);
                }
            }
        } else {
            if (newTopLevelElementForDayD && allDaysWrapper) {
                const nextDayElement = findNextDayElementInDom(d, allDaysWrapper);
                if (nextDayElement) {
                    allDaysWrapper.insertBefore(newTopLevelElementForDayD, nextDayElement);
                } else {
                    allDaysWrapper.appendChild(newTopLevelElementForDayD);
                }
            }
            if (liveElementForDayD && liveElementForDayD !== newTopLevelElementForDayD) {
                if (liveElementForDayD.parentElement === allDaysWrapper) {
                    allDaysWrapper.removeChild(liveElementForDayD);
                } else if (!liveElementForDayD.parentElement && (!newTopLevelElementForDayD || !newTopLevelElementForDayD.contains(liveElementForDayD))) {
                    liveElementForDayD.remove();
                }
            }
        }

        const statusEl = getElementByIdSafe(`indice_pluv_status_day${d}`);
        if (statusEl && !statusEl.textContent?.trim()) {
             updateRainfallStatus(d, 'idle');
        }
    }
}


async function setupContractDataSyncForDay(dayIndex, dayWrapperElement, page1Container) {
    const daySuffix = `_day${dayIndex}`;

    if (!dayWrapperElement || !page1Container) {
        console.error(`Day wrapper or page1Container not provided for day ${dayIndex}. Contract data sync aborted.`);
        return;
    }

    if (dayIndex === 0) {
        const day0DateDrivingFields = ['prazo_day0', 'inicio_obra_day0'];
        day0DateDrivingFields.forEach(masterId => {
            const masterEl = getElementByIdSafe(masterId, page1Container);
            if (masterEl) {
                masterEl.readOnly = false;
                const eventType = (masterId.includes('prazo') || masterId.includes('inicio_obra')) ? 'change' : 'input';
                masterEl.addEventListener(eventType, async () => {
                    markFormAsDirty();
                    await updateAllDateRelatedFieldsFromDay0();
                });
            }
        });

        const dataDay0InputEl = getElementByIdSafe('data_day0', page1Container);
        const monthYearPickerIcon = page1Container.querySelector('.date-input-icon-wrapper .date-picker-indicator-icon');
        const monthYearPickerElement = getElementByIdSafe('monthYearPicker_day0', page1Container);

        if (dataDay0InputEl && monthYearPickerIcon && monthYearPickerElement) {
            const currentYearDisplay = monthYearPickerElement.querySelector('.current-year');
            const prevYearButton = monthYearPickerElement.querySelector('.prev-year');
            const nextYearButton = monthYearPickerElement.querySelector('.next-year');
            const monthsGrid = monthYearPickerElement.querySelector('.months-grid');
            const okButton = monthYearPickerElement.querySelector('.picker-ok');
            const cancelButton = monthYearPickerElement.querySelector('.picker-cancel');

            const renderPicker = () => {
                if(currentYearDisplay) currentYearDisplay.textContent = currentPickerYear.toString();
                if(monthsGrid) {
                    monthsGrid.querySelectorAll('button').forEach(btn => {
                        btn.classList.remove('selected');
                        if (parseInt(btn.dataset.month || '-1') === currentPickerMonth) {
                            btn.classList.add('selected');
                        }
                    });
                }
            };

            const openPicker = () => {
                const currentDateVal = dataDay0InputEl.value;
                if (currentDateVal && /^\d{4}-\d{2}-\d{2}$/.test(currentDateVal)) {
                    const [year, month] = currentDateVal.split('-').map(Number);
                    currentPickerYear = year;
                    currentPickerMonth = month - 1;
                } else {
                    const today = new Date();
                    currentPickerYear = today.getFullYear();
                    currentPickerMonth = today.getMonth();
                }
                renderPicker();
                monthYearPickerElement.style.display = 'block';
                monthYearPickerIcon.setAttribute('aria-expanded', 'true');
                prevYearButton.focus();
            };

            const closePicker = () => {
                monthYearPickerElement.style.display = 'none';
                monthYearPickerIcon.setAttribute('aria-expanded', 'false');
                monthYearPickerIcon.focus();
            };

            monthYearPickerIcon.addEventListener('click', (e) => {
                e.stopPropagation();
                if (monthYearPickerElement.style.display === 'block') {
                    closePicker();
                } else {
                    openPicker();
                }
            });
             monthYearPickerIcon.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    monthYearPickerIcon.click();
                }
            });


            if(prevYearButton) prevYearButton.addEventListener('click', () => { currentPickerYear--; renderPicker(); });
            if(nextYearButton) nextYearButton.addEventListener('click', () => { currentPickerYear++; renderPicker(); });

            if(monthsGrid) {
                monthsGrid.querySelectorAll('button').forEach(btn => {
                    btn.addEventListener('click', () => {
                        currentPickerMonth = parseInt(btn.dataset.month || '0');
                        renderPicker();
                    });
                });
            }

            if(okButton) {
                okButton.addEventListener('click', async () => {
                    const newDay = '01';
                    const newMonth = (currentPickerMonth + 1).toString().padStart(2, '0');
                    const newYear = currentPickerYear.toString();
                    const newDateYYYYMMDD = `${newYear}-${newMonth}-${newDay}`;

                    dataDay0InputEl.value = newDateYYYYMMDD;
                    markFormAsDirty();
                    await adjustRdoDaysForMonth(newDateYYYYMMDD);
                    closePicker();
                });
            }

            if(cancelButton) cancelButton.addEventListener('click', closePicker);

            document.addEventListener('click', (event) => {
                if (monthYearPickerElement.style.display === 'block' &&
                    !monthYearPickerElement.contains(event.target) &&
                    !monthYearPickerIcon.contains(event.target)) {
                    closePicker();
                }
            });
             monthYearPickerElement.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    closePicker();
                }
            });

        } else {
            console.warn("Month/Year picker elements for Day 0 not fully found. Custom picker disabled.");
        }


        ['contratada_day0', 'contrato_num_day0'].forEach(globalMasterId => {
            const masterEl = getElementByIdSafe(globalMasterId, page1Container);
            if (masterEl) {
                masterEl.readOnly = false;
                masterEl.addEventListener('input', async () => {
                    markFormAsDirty();
                    await updateAllDateRelatedFieldsFromDay0();
                });
                 masterEl.addEventListener('change', async () => {
                    markFormAsDirty();
                    await updateAllDateRelatedFieldsFromDay0();
                });
            }
        });
    } else {
        contractFieldMappingsBase.forEach(fieldPair => {
            const masterElement = getElementByIdSafe(fieldPair.masterIdBase + daySuffix, page1Container);
            if (masterElement) masterElement.readOnly = true;
        });
        const rdoNumInput = getElementByIdSafe(`numero_rdo${daySuffix}`, page1Container);
        if (rdoNumInput) rdoNumInput.readOnly = true;
         const dataInput = getElementByIdSafe(`data${daySuffix}`, page1Container);
        if(dataInput) dataInput.readOnly = true;
    }

    const masterRdoInputDay0 = getElementByIdSafe('numero_rdo_day0', page1Container);
    if (dayIndex === 0 && masterRdoInputDay0) {
        masterRdoInputDay0.readOnly = false;
        const rdoChangeHandler = async () => {
            markFormAsDirty();
            await updateAllDateRelatedFieldsFromDay0();
        };
        masterRdoInputDay0.addEventListener('input', rdoChangeHandler);
        masterRdoInputDay0.addEventListener('change', rdoChangeHandler);
    }

    contractFieldMappingsBase.forEach(fieldPair => {
        fieldPair.slaveIdBases.forEach(slaveIdBase => {
            const slaveElement = getElementByIdSafe(slaveIdBase + daySuffix, dayWrapperElement);
            if (slaveElement) slaveElement.readOnly = true;
        });
    });
    const rdoPtInput = getElementByIdSafe(`numero_rdo_pt${daySuffix}`, dayWrapperElement);
    if (rdoPtInput) rdoPtInput.readOnly = true;
    const rdoPt2Input = getElementByIdSafe(`numero_rdo_pt2${daySuffix}`, dayWrapperElement);
    if (rdoPt2Input) rdoPt2Input.readOnly = true;
}


// --- Day Navigation Panel ---
const dayTabsContainer = getElementByIdSafe('day-tabs-container');
const dayTabs = [];
let navObserver = null;
let lastActivatedByScroll = -1;
const EXTRA_SCROLL_PADDING_TOP = 11;

function setActiveTab(activeIndex) {
    if (activeIndex < 0 || activeIndex >= dayTabs.length) {
        dayTabs.forEach(tab => {
            if (tab) {
                tab.classList.remove('active');
                tab.removeAttribute('aria-current');
            }
        });
        const firstValidTab = dayTabs.find(tab => tab);
        if (firstValidTab) {
            firstValidTab.classList.add('active');
            firstValidTab.setAttribute('aria-current', 'true');
        }
        return;
    }


    dayTabs.forEach((tab, index) => {
        if (tab) {
            if (index === activeIndex) {
                tab.classList.add('active');
                tab.setAttribute('aria-current', 'true');
            } else {
                tab.classList.remove('active');
                tab.removeAttribute('aria-current');
            }
        }
    });
}

function createDayTab(dayIndex) {
    if (!dayTabsContainer) return;

    if (document.getElementById(`day-tab-${dayIndex}`)) {
        return;
    }

    const tab = document.createElement('button');
    tab.id = `day-tab-${dayIndex}`;
    tab.className = 'day-tab';
    tab.type = 'button';
    tab.setAttribute('aria-controls', `rdo_day_${dayIndex}_wrapper`);
    tab.setAttribute('role', 'tab');

    const dayNumberText = document.createTextNode(`${dayIndex + 1}`);
    tab.appendChild(dayNumberText);

    const alertIconSpan = document.createElement('span');
    alertIconSpan.className = 'efetivo-alert-icon';
    alertIconSpan.setAttribute('aria-hidden', 'true');
    alertIconSpan.textContent = '⚠️';
    tab.appendChild(alertIconSpan);


    tab.addEventListener('click', () => {
        const targetDayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${dayIndex}`);
        const targetRdoDayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
        const elementToScrollTo = targetDayWithButtonLayoutWrapper || targetRdoDayWrapper;

        if (elementToScrollTo) {
            const rootStyles = getComputedStyle(document.documentElement);
            const consbemHeaderHeight = parseInt(rootStyles.getPropertyValue('--consbem-header-height').trim() || '60', 10);
            const rdoActionBarHeight = parseInt(rootStyles.getPropertyValue('--rdo-action-bar-height').trim() || '60', 10);
            const totalHeaderOffset = consbemHeaderHeight + rdoActionBarHeight;

            const elementPosition = elementToScrollTo.getBoundingClientRect().top + window.scrollY;
            const scrollToPosition = elementPosition - totalHeaderOffset - EXTRA_SCROLL_PADDING_TOP;


            window.scrollTo({
                top: scrollToPosition,
                behavior: 'smooth'
            });
        }
    });

    dayTabsContainer.appendChild(tab);
    dayTabs[dayIndex] = tab;

     if (dayTabs.filter(t => t).length === 1 && !dayTabsContainer.querySelector('.day-tab.active')) {
        setActiveTab(dayIndex);
    }
}

function setupNavigationObserver() {
    if (navObserver) {
        navObserver.disconnect();
    }
    const appContentCanvas = getElementByIdSafe('app-content-canvas');
    if (!appContentCanvas) {
        console.warn("App content canvas not found for navigation observer.");
        return;
    }

    const consbemHeaderHeight = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--consbem-header-height') || '60');
    const rdoActionBarHeight = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--rdo-action-bar-height') || '70');
    const topOffset = consbemHeaderHeight + rdoActionBarHeight + EXTRA_SCROLL_PADDING_TOP + 5;


    const observerOptions = {
        root: null,
        rootMargin: `-${topOffset}px 0px -${window.innerHeight - topOffset - 100}px 0px`,
        threshold: 0.01
    };

    navObserver = new IntersectionObserver((entries) => {
        let highestVisibleEntry = null;

        entries.forEach(entry => {
            if (entry.isIntersecting) {
                if (!highestVisibleEntry || entry.boundingClientRect.top < highestVisibleEntry.boundingClientRect.top) {
                    highestVisibleEntry = entry;
                }
            }
        });

        if (highestVisibleEntry) {
            let dayIndexStr;
            const targetEl = highestVisibleEntry.target;

            if (targetEl.classList.contains('day-with-button-container')) {
                const idParts = targetEl.id.split('_');
                dayIndexStr = idParts[idParts.length -1]?.replace('day','');
            } else if (targetEl.classList.contains('rdo-day-wrapper')) {
                dayIndexStr = targetEl.dataset.dayIndex;
            }


            if (dayIndexStr) {
                const dayIndex = parseInt(dayIndexStr, 10);
                if (dayIndex !== lastActivatedByScroll || !dayTabs[dayIndex]?.classList.contains('active')) {
                    setActiveTab(dayIndex);
                    lastActivatedByScroll = dayIndex;
                }
            }
        }
    }, observerOptions);

    const dayWrappers = document.querySelectorAll('.rdo-day-wrapper, .day-with-button-container');
    dayWrappers.forEach(wrapper => {
        if (navObserver) {
            navObserver.observe(wrapper);
        }
    });
}

function attachDirtyListenersToDay(dayWrapperElement) {
    const inputs = dayWrapperElement.querySelectorAll(
        'input[type="text"], input[type="date"], input[type="number"], input[type="month"], textarea, select, input[type="file"]'
    );
    inputs.forEach(input => {
        if (input.type === 'hidden' && (input.name.startsWith('status_day') || input.name.startsWith('localizacao_day') || input.name.startsWith('tipo_servico_day') || input.name.startsWith('servico_desc_day')) ) {
            return;
        }
        if (input.classList.contains('quantity-input')) {
            return;
        }
        if (input.id.startsWith('consbemPhoto_') || input.id.startsWith('report_photo_')) {
            return;
        }


        const eventType = (input.type === 'date' || input.type === 'month' || input.tagName.toLowerCase() === 'select' || input.type === 'file') ? 'change' : 'input';
        input.addEventListener(eventType, markFormAsDirty);
    });
}

function checkAndToggleEfetivoAlertForDay(dayIndex) {
    const tab = getElementByIdSafe(`day-tab-${dayIndex}`);
    if (!tab) return;

    const alertIcon = tab.querySelector('.efetivo-alert-icon');
    if (!alertIcon) return;

    const hideAlert = () => {
        alertIcon.classList.remove('visible');
    };
    const showAlert = () => {
        alertIcon.classList.add('visible');
    };

    let rdoDayWrapper = null;
    const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${dayIndex}`);
    if (dayWithButtonLayoutWrapper) {
        rdoDayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
    } else {
        rdoDayWrapper = getElementByIdSafe(`rdo_day_${dayIndex}_wrapper`);
    }

    if (!rdoDayWrapper) {
        hideAlert();
        return;
    }
    const page1Container = rdoDayWrapper.querySelector(`#rdo_day_${dayIndex}_page_0`);
    if (!page1Container) {
        hideAlert();
        return;
    }

    if (!hasEfetivoBeenLoadedAtLeastOnce) {
        hideAlert();
        return;
    }

    const quantityInputs = page1Container.querySelectorAll('.labor-table .quantity-input');
    let efetivoPresent = false;
    if (quantityInputs.length > 0) {
        for (const input of quantityInputs) {
            const value = parseInt(input.value, 10);
            if (!isNaN(value) && value > 0) {
                efetivoPresent = true;
                break;
            }
        }
    }

    if (efetivoPresent) {
        hideAlert();
    } else {
        if (efetivoWasJustClearedByButton) {
            hideAlert();
        } else {
            showAlert();
        }
    }
}


function updateAllEfetivoAlerts() {
    for (let d = 0; d <= rdoDayCounter; d++) {
        checkAndToggleEfetivoAlertForDay(d);
    }
}


async function initializeDayInstance(dayIndex, dayWrapperElement) {
    dayWrapperElement.dataset.dayIndex = dayIndex.toString();

    const pageContainers = dayWrapperElement.querySelectorAll('.rdo-container');
    pageContainers.forEach((pageContainer, pageInDayIdx) => {
        let pageIndicator = pageContainer.querySelector('.page-indicator');
        if (!pageIndicator) {
            pageIndicator = document.createElement('div');
            pageIndicator.className = 'page-indicator';
            if (pageContainer.firstChild) {
                pageContainer.insertBefore(pageIndicator, pageContainer.firstChild);
            } else {
                pageContainer.appendChild(pageIndicator);
            }
        }
        pageIndicator.textContent = `DIA ${dayIndex + 1} / FOLHA ${pageInDayIdx + 1}`;
    });

    const page1Container = dayWrapperElement.querySelector(`#rdo_day_${dayIndex}_page_0`);
    const page2Container = dayWrapperElement.querySelector(`#rdo_day_${dayIndex}_page_1`);
    const page3Container = dayWrapperElement.querySelector(`#rdo_day_${dayIndex}_page_2`);


    if (!page1Container) {
        console.error(`Page 1 container (rdo_day_${dayIndex}_page_0) not found in dayWrapperElement for day ${dayIndex}. Initialization incomplete.`);
        return;
    }
    if (!page2Container) {
         console.warn(`Page 2 container (rdo_day_${dayIndex}_page_1) not found for day ${dayIndex}. Some features might not initialize fully.`);
    }
    if (!page3Container) {
         console.warn(`Page 3 container (rdo_day_${dayIndex}_page_2) not found for day ${dayIndex}. Some features might not initialize fully.`);
    }

    createCombinedShiftControl('turno1_control_cell', 't1', 'B', 'N', '1º Turno', dayIndex, page1Container);
    createCombinedShiftControl('turno2_control_cell', 't2', 'B', 'N', '2º Turno', dayIndex, page1Container);
    createCombinedShiftControl('turno3_control_cell', 't3', "", "", '3º Turno', dayIndex, page1Container);


    const laborColumns = page1Container.querySelectorAll('.labor-section .labor-column');
    laborColumns.forEach(column => {
        let category = '';
        if (column.classList.contains('labor-direta')) category = 'mod';
        else if (column.classList.contains('labor-indireta')) category = 'moi';
        else if (column.classList.contains('equipamentos')) category = 'equip';

        if (category) {
            const quantityInputs = column.querySelectorAll('.labor-table .quantity-input');
            quantityInputs.forEach((input, itemIndex) => {
                const row = input.closest('tr');
                if (row) {
                    const itemNameCell = row.querySelector('.item-name-cell');
                    if (itemNameCell) {
                        let itemName = (itemNameCell.textContent || '').trim();
                        let normalizedItemName = itemName.toLowerCase()
                                                      .replace(/\s+/g, '_')
                                                      .replace(/[().\/"]/g, '')
                                                      .replace(/[^a-z0-9_]/g, '');
                        if (normalizedItemName.length > 50) {
                            normalizedItemName = normalizedItemName.substring(0, 50);
                        }
                        if (normalizedItemName && !/^_(_)*$/.test(normalizedItemName)) {
                             input.id = `efetivo_day${dayIndex}_${category}_${normalizedItemName}`;
                        } else {
                            input.id = `efetivo_day${dayIndex}_${category}_item${itemIndex}`;
                        }
                    } else {
                         input.id = `efetivo_day${dayIndex}_${category}_row${itemIndex}_input`;
                    }
                } else {
                     input.id = `efetivo_day${dayIndex}_${category}_orphan${itemIndex}_input`;
                }
            });
        }
    });

    const laborSectionPage1Elements = page1Container.querySelectorAll('.labor-section .labor-column');
    if (laborSectionPage1Elements.length > 0) {
        laborSectionPage1Elements.forEach(columnElement => {
            const quantityInputs = columnElement.querySelectorAll('.labor-table .quantity-input');
             if (dayIndex === 0) {
                quantityInputs.forEach(input => {
                     if (!input.value) input.value = '';
                });
            }
            applyMultiColumnLayoutIfNecessary(columnElement.querySelector('.labor-table tbody'));
        });
        page1Container.addEventListener('input', (event) => {
            const target = event.target;
            if (target && target.classList.contains('quantity-input')) {
                efetivoWasJustClearedByButton = false;
                hasEfetivoBeenLoadedAtLeastOnce = true;
                markFormAsDirty();
                debouncedCalculateSectionTotalsForDay(page1Container, dayIndex);
            }
        });
        calculateSectionTotalsForDay(page1Container, dayIndex);
    }


    await setupPhotoAttachmentForDay(dayIndex);

    const activitiesTableBodyP1 = getElementByIdSafe(`activitiesTable_day${dayIndex}`, page1Container)?.getElementsByTagName('tbody')[0];
    if (activitiesTableBodyP1 && activitiesTableBodyP1.rows.length === 0) {
        for (let i = 0; i < 24; i++) addActivityRow('activitiesTable', dayIndex, page1Container);
    }

    if (page2Container) {
        const activitiesTableBodyP2 = getElementByIdSafe(`activitiesTable_pt_day${dayIndex}`, page2Container)?.getElementsByTagName('tbody')[0];
        if (activitiesTableBodyP2 && activitiesTableBodyP2.rows.length === 0) {
            for (let i = 0; i < 22; i++) addActivityRow('activitiesTable_pt', dayIndex, page2Container);
        }
    }


    if (page3Container) {
        for (let i = 1; i <= 4; i++) {
            await setupReportPhotoSlot(
                `report_photo_${i}_pt2_input`, `report_photo_${i}_pt2_preview`,
                `report_photo_${i}_pt2_placeholder`,
                `report_photo_${i}_pt2_clear_button`,
                `report_photo_${i}_pt2_label`, dayIndex, i,
                page3Container
            );
            const captionInput = getElementByIdSafe(`report_photo_${i}_pt2_caption_day${dayIndex}`, page3Container);
            if (captionInput) {
                captionInput.addEventListener('input', markFormAsDirty);
            }
        }
    }

    await setupContractDataSyncForDay(dayIndex, dayWrapperElement, page1Container);
    attachDirtyListenersToDay(dayWrapperElement);

    updateRainfallStatus(dayIndex, 'idle');

    // Store original height for responsive scaling
    if (!originalRdoSheetHeights.has(dayWrapperElement.id)) {
        originalRdoSheetHeights.set(dayWrapperElement.id, dayWrapperElement.offsetHeight);
    }
}


function updateIdsRecursive(element, oldDaySuffix, newDaySuffix) {
    const oldDayNumStr = oldDaySuffix.replace('_day', '');
    const newDayNumStr = newDaySuffix.replace('_day', '');

    const regexPatternFull = new RegExp(`_day_${oldDayNumStr}_`, 'g');
    const replacementFull = `_day_${newDayNumStr}_`;
    const regexPatternSuffixOnly = new RegExp(`${oldDaySuffix}\\b`, 'g');
    const replacementSuffixOnly = newDaySuffix;

    ['id', 'for', 'name', 'aria-controls', 'aria-labelledby', 'aria-describedby'].forEach(attrName => {
        let attrValue = element.getAttribute(attrName);
        if (attrValue) {
            if (attrValue.includes(`_day_${oldDayNumStr}_`)) {
                attrValue = attrValue.replace(regexPatternFull, replacementFull);
            }
            if (attrValue.match(regexPatternSuffixOnly)) {
                attrValue = attrValue.replace(regexPatternSuffixOnly, replacementSuffixOnly);
            }
            element.setAttribute(attrName, attrValue);
        }
    });

    const ariaLabelAttribute = element.getAttribute('aria-label');
    const oldDayNumForLabel = parseInt(oldDayNumStr, 10);
    const newDayNumForLabel = parseInt(newDayNumStr, 10);

    if (ariaLabelAttribute) {
        let newAriaLabel = ariaLabelAttribute;
        if (newAriaLabel.includes(`_day_${oldDayNumStr}_`)) {
            newAriaLabel = newAriaLabel.replace(regexPatternFull, replacementFull);
        }
        if (newAriaLabel.match(regexPatternSuffixOnly)) {
            newAriaLabel = newAriaLabel.replace(regexPatternSuffixOnly, replacementSuffixOnly);
        }
        if (newAriaLabel.includes(`Dia ${oldDayNumForLabel + 1}`)) {
            newAriaLabel = newAriaLabel.replace(`Dia ${oldDayNumForLabel + 1}`, `Dia ${newDayNumForLabel + 1}`);
        }
        element.setAttribute('aria-label', newAriaLabel);
    }

    element.childNodes.forEach(child => {
        if (child.nodeType === Node.ELEMENT_NODE) {
            updateIdsRecursive(child, oldDaySuffix, newDaySuffix);
        }
    });
}

function resetDayInstanceContent(dayWrapperElement, dayIndex) {
    const daySuffix = `_day${dayIndex}`;

    dayWrapperElement.querySelectorAll('.labor-table .quantity-input').forEach(input => input.value = '');

    const activityFieldsSelectors = [
        `[id^="activitiesTable_day${dayIndex}"] textarea[id^="obs_atividade_"]`,
        `[id^="activitiesTable_pt_day${dayIndex}"] textarea[id^="obs_atividade_"]`
    ];
    dayWrapperElement.querySelectorAll(activityFieldsSelectors.join(', '))
        .forEach(el => el.value = '');


    const dropdownContainers = dayWrapperElement.querySelectorAll('.status-select-container');
    dropdownContainers.forEach(container => {
        const display = container.querySelector('.status-display');
        const hiddenInput = container.querySelector('input[type="hidden"]');
        const dropdown = container.querySelector('.status-dropdown');

        if (display) {
            display.textContent = '';
            display.setAttribute('aria-expanded', 'false');
        }
        if (hiddenInput) {
            hiddenInput.value = '';
        }
        if (dropdown) {
            dropdown.style.display = 'none';
            dropdown.querySelectorAll('.status-option').forEach(opt => {
                opt.setAttribute('aria-selected', (opt.dataset.value === "") ? 'true' : 'false');
            });
        }
    });

    const page2Container = dayWrapperElement.querySelectorAll('.rdo-container')[1];
    if (page2Container) {
        const page2ObservationsTextarea = getElementByIdSafe(`observacoes_comentarios_pt${daySuffix}`, page2Container);
        if (page2ObservationsTextarea) page2ObservationsTextarea.value = '';
    }

    const page3Container = dayWrapperElement.querySelectorAll('.rdo-container')[2];
    if (page3Container) {
        for (let i = 1; i <= 4; i++) {
            const captionInput = getElementByIdSafe(`report_photo_${i}_pt2_caption${daySuffix}`, page3Container);
            if (captionInput) captionInput.value = '';

            const photoInput = getElementByIdSafe(`report_photo_${i}_pt2_input${daySuffix}`, page3Container);
            const preview = getElementByIdSafe(`report_photo_${i}_pt2_preview${daySuffix}`, page3Container);
            const placeholder = getElementByIdSafe(`report_photo_${i}_pt2_placeholder${daySuffix}`, page3Container);
            const clearButton = getElementByIdSafe(`report_photo_${i}_pt2_clear_button${daySuffix}`, page3Container);
            const label = getElementByIdSafe(`report_photo_${i}_pt2_label${daySuffix}`, page3Container);
            resetReportPhotoSlotState(photoInput, preview, placeholder, clearButton, label);
        }
    }


    const obsClima = dayWrapperElement.querySelector(`#obs_clima${daySuffix}`);
    if(obsClima) obsClima.value = '';

    const indicePluvInput = dayWrapperElement.querySelector(`#indice_pluv_valor${daySuffix}`);
    if (indicePluvInput) indicePluvInput.value = '';
    updateRainfallStatus(dayIndex, 'idle');

    const page1Container = dayWrapperElement.querySelectorAll('.rdo-container')[0];
    if (page1Container) {
        calculateSectionTotalsForDay(page1Container, dayIndex);
    }
    checkAndToggleEfetivoAlertForDay(dayIndex);
}

async function addNewRdoDay(newDayIndex) {
    if (!pristineRdoDay0WrapperHtml) {
        console.error("Pristine HTML for Day 0 wrapper not available! Cannot add new day.");
        return null;
    }

    const newDayWrapper = document.createElement('div');
    newDayWrapper.className = 'rdo-day-wrapper';
    newDayWrapper.innerHTML = pristineRdoDay0WrapperHtml;
    newDayWrapper.id = `rdo_day_${newDayIndex}_wrapper`;

    updateIdsRecursive(newDayWrapper, '_day0', `_day${newDayIndex}`);
    newDayWrapper.dataset.dayIndex = newDayIndex.toString();

    newDayWrapper.querySelectorAll('.rdo-container').forEach(container => {
        container.classList.add('rdo-container-subsequent-day');
    });

    resetDayInstanceContent(newDayWrapper, newDayIndex);
    return newDayWrapper;
}


async function removeLastRdoDay() {
    if (rdoDayCounter < 1) return;

    const lastDayIndex = rdoDayCounter;

    const dayWithButtonContainerToRemove = document.getElementById(`day_with_button_container_day${lastDayIndex}`);
    const rdoDayWrapperItself = getElementByIdSafe(`rdo_day_${lastDayIndex}_wrapper`);

    let elementToRemoveFromObserver = null;
    let elementToRemoveFromDOM = null;

    if (dayWithButtonContainerToRemove) {
        elementToRemoveFromObserver = dayWithButtonContainerToRemove;
        elementToRemoveFromDOM = dayWithButtonContainerToRemove;
    } else if (rdoDayWrapperItself) {
        elementToRemoveFromObserver = rdoDayWrapperItself;
        elementToRemoveFromDOM = rdoDayWrapperItself;
    }


    const tabToRemove = getElementByIdSafe(`day-tab-${lastDayIndex}`);

    if (elementToRemoveFromObserver && navObserver) {
        navObserver.unobserve(elementToRemoveFromObserver);
    }
    if (elementToRemoveFromDOM) {
        originalRdoSheetHeights.delete(elementToRemoveFromDOM.id); // Remove from height cache
        elementToRemoveFromDOM.remove();
    } else {
        console.warn(`Could not find wrapper to remove for day ${lastDayIndex}`);
    }


    if (tabToRemove) {
        tabToRemove.remove();
        dayTabs[lastDayIndex] = undefined;
    }

    rdoDayCounter--;

    if (tabToRemove && tabToRemove.classList.contains('active')) {
        setActiveTab(rdoDayCounter);
    } else if (dayTabs.filter(t=>t).length > 0 && !dayTabsContainer?.querySelector('.day-tab.active')) {
        setActiveTab(rdoDayCounter);
    }
    markFormAsDirty();
    updateClearEfetivoButtonVisibility();
}


async function adjustRdoDaysForMonth(baseDateStrYYYYMMDD) {
    if (!baseDateStrYYYYMMDD || !/^\d{4}-\d{2}-\d{2}$/.test(baseDateStrYYYYMMDD)) {
        console.warn("adjustRdoDaysForMonth: Invalid or empty date format provided (expected YYYY-MM-01):", baseDateStrYYYYMMDD);
        const today = new Date();
        baseDateStrYYYYMMDD = formatDateToYYYYMMDD(new Date(today.getFullYear(), today.getMonth(), 1));
        console.warn(`Using fallback date for month adjustment: ${baseDateStrYYYYMMDD}`);
        const day0DateInput = getElementByIdSafe('data_day0');
        if (day0DateInput) {
            day0DateInput.value = baseDateStrYYYYMMDD;
        }
    }

    const dateParts = baseDateStrYYYYMMDD.split('-');
    const year = parseInt(dateParts[0]);
    const month_1_indexed = parseInt(dateParts[1]);

    const numDaysInTargetMonth = getDaysInMonth(year, month_1_indexed);

    if (isNaN(numDaysInTargetMonth) || numDaysInTargetMonth <= 0) {
        console.warn(`adjustRdoDaysForMonth: Could not determine days in month for ${year}-${month_1_indexed}.`);
        return;
    }

    let currentNumRdos = rdoDayCounter + 1;
    let daysAdjusted = false;
    const allDaysWrapper = getElementByIdSafe('rdo-all-days-wrapper');
    if (!allDaysWrapper) {
        console.error("adjustRdoDaysForMonth: #rdo-all-days-wrapper not found.");
        return;
    }

    const fragment = document.createDocumentFragment();

    while (currentNumRdos < numDaysInTargetMonth) {
        rdoDayCounter++;
        const newDayWrapper = await addNewRdoDay(rdoDayCounter);
        if (newDayWrapper) {
            await initializeDayInstance(rdoDayCounter, newDayWrapper);
            const finalElementForFragment = updateEfetivoCopyButtonVisibility(rdoDayCounter, newDayWrapper);
            fragment.appendChild(finalElementForFragment);
            createDayTab(rdoDayCounter);
        } else {
            rdoDayCounter--;
        }
        currentNumRdos = rdoDayCounter + 1;
        daysAdjusted = true;
    }
    if (fragment.childNodes.length > 0) {
        allDaysWrapper.appendChild(fragment);
    }


    while (currentNumRdos > numDaysInTargetMonth && currentNumRdos > 1) {
        await removeLastRdoDay();
        currentNumRdos = rdoDayCounter + 1;
        daysAdjusted = true;
    }

    await updateAllDateRelatedFieldsFromDay0();

    setupNavigationObserver();
    applyResponsiveScaling(); // Apply scaling after DOM changes

    const activeTabIndex = dayTabs.findIndex(tab => tab?.classList.contains('active'));
    if (numDaysInTargetMonth > 0) {
        if (activeTabIndex >= numDaysInTargetMonth || activeTabIndex === -1) {
            setActiveTab(Math.min(numDaysInTargetMonth - 1, Math.max(0, activeTabIndex)));
        }
    } else if (dayTabs.filter(t=>t).length > 0) {
        setActiveTab(0);
    }

    if (dayTabs.filter(t=>t).length > 0 && !dayTabsContainer?.querySelector('.day-tab.active')) {
        setActiveTab(0);
    }

    if (daysAdjusted) {
        markFormAsDirty();
    }
    updateClearEfetivoButtonVisibility();
    updateAllEfetivoAlerts();
}


// --- PDF Generation Constants & Helpers ---
const margin = 7;
const pageWidth = 210;
const pageHeight = 297;
const contentWidth = pageWidth - (2 * margin);
const lineWeight = 0.1;
const lineWeightStrong = 0.1;
const photoFrameLineWeight = 0.1;

async function imageUrlToDataUrl(url) {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Failed to fetch image ${url}: ${response.statusText}`);
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

const addImageToPdfAsync = async (pdf, selector, x, y, w, h, parent) => {
    const imgElement = parent.querySelector(selector);
    if (imgElement && imgElement.src && !imgElement.src.endsWith('#') && !imgElement.src.startsWith('about:blank')) {
        const isStaticHeaderLogo = imgElement.classList.contains('sabesp-logo') || imgElement.classList.contains('consorcio-logo');
        if (isStaticHeaderLogo || imgElement.dataset.isEffectivelyLoaded === 'true') {
            try {
                let imageDataUrl = imgElement.src;

                if (imageDataUrl.startsWith('http') && !imageDataUrl.startsWith('data:image')) {
                    try { imageDataUrl = await imageUrlToDataUrl(imageDataUrl); }
                    catch (fetchError) { console.error(`Failed to convert image URL ${imgElement.src}:`, fetchError); pdf.text(`[Img Load Err]`, x + 2, y + h / 2); return; }
                }


                let naturalW = 0, naturalH = 0;

                const imgNaturalWidthAttr = imgElement.dataset.naturalWidth;
                const imgNaturalHeightAttr = imgElement.dataset.naturalHeight;

                if (imgNaturalWidthAttr && imgNaturalHeightAttr) {
                    naturalW = parseFloat(imgNaturalWidthAttr);
                    naturalH = parseFloat(imgNaturalHeightAttr);
                } else {
                    const tempImgLoad = new Image();
                    tempImgLoad.src = imageDataUrl;
                    await new Promise((resolve, reject) => {
                         tempImgLoad.onload = () => resolve();
                         tempImgLoad.onerror = (err) => { console.error(`Temp Img Load Error for ${selector}`, err); reject(err); };
                    });
                    naturalW = tempImgLoad.naturalWidth;
                    naturalH = tempImgLoad.naturalHeight;
                    if (naturalW && naturalH && imgElement.dataset) {
                        imgElement.dataset.naturalWidth = naturalW.toString();
                        imgElement.dataset.naturalHeight = naturalH.toString();
                    }
                }

                 if(!naturalW || !naturalH) {
                    try {
                        const props = pdf.getImageProperties(imageDataUrl);
                        naturalW = props.width;
                        naturalH = props.height;
                    } catch (e) {
                         console.error(`Error getting image properties for ${selector} via jsPDF:`, e);
                         pdf.text(`[Img Meta Err]`, x + 2, y + h / 2); return;
                    }
                }
                 if(!naturalW || !naturalH) {
                    console.error(`Could not determine dimensions for image ${selector}. Skipping addImage.`);
                    pdf.text(`[Img Dim Err]`, x + 2, y + h / 2); return;
                }

                let format = 'JPEG';
                const mimeTypeMatch = imageDataUrl.match(/^data:image\/([a-zA-Z]+);base64,/);
                if (mimeTypeMatch && mimeTypeMatch[1]) { format = mimeTypeMatch[1].toUpperCase(); }
                else {
                    try {
                        const propsForFormat = pdf.getImageProperties(imageDataUrl);
                        format = propsForFormat.fileType || 'JPEG';
                    } catch (e) {
                        console.warn(`Could not determine image type for ${selector} via jsPDF, defaulting to JPEG. Error:`, e);
                        format = 'JPEG';
                    }
                }
                const supportedFormats = ['PNG', 'JPEG', 'WEBP'];
                if (!supportedFormats.includes(format)) { console.warn(`Unsupported image format ${format} for ${selector}, defaulting to JPEG.`); format = 'JPEG';}


                if (imgElement.classList.contains('report-photo-image-pt2')) {
                    const htmlImg = new Image();
                    htmlImg.crossOrigin = "anonymous";

                    const imgLoadPromise = new Promise((resolve, reject) => {
                        htmlImg.onload = () => resolve();
                        htmlImg.onerror = (err) => {
                            console.error(`Failed to load image into HTMLImageElement for canvas processing: ${selector}`, err);
                            reject(new Error(`HTMLImageElement load error for ${selector}`));
                        };
                        htmlImg.src = imageDataUrl;
                    });
                    await imgLoadPromise;

                    const dpi = 150;
                    const cvWidthPx = (w / 25.4) * dpi;
                    const cvHeightPx = (h / 25.4) * dpi;

                    const tempCanvas = document.createElement('canvas');
                    tempCanvas.width = cvWidthPx;
                    tempCanvas.height = cvHeightPx;
                    const ctx = tempCanvas.getContext('2d');

                    if (!ctx) {
                        console.error(`Could not get 2D context for pre-cropping image ${selector}. Falling back to direct addImage with clipping (might not be perfect 'cover').`);
                        const imageAspectRatioFallback = naturalW / naturalH;
                        const boxAspectRatioFallback = w / h;
                        let drawWfb, drawHfb, drawXfb, drawYfb;
                        if (imageAspectRatioFallback > boxAspectRatioFallback) {
                            drawHfb = h; drawWfb = h * imageAspectRatioFallback; drawXfb = x + (w - drawWfb) / 2; drawYfb = y;
                        } else {
                            drawWfb = w; drawHfb = w / imageAspectRatioFallback; drawXfb = x; drawYfb = y + (h - drawHfb) / 2;
                        }
                        pdf.saveGraphicsState();
                        pdf.addImage(imageDataUrl, format, drawXfb, drawYfb, drawWfb, drawHfb);
                        pdf.restoreGraphicsState();
                    } else {
                        const imageAspect = naturalW / naturalH;
                        const canvasAspect = cvWidthPx / cvHeightPx;

                        let srcX = 0, srcY = 0, srcW = naturalW, srcH = naturalH;

                        if (imageAspect > canvasAspect) {
                            srcH = naturalH;
                            srcW = naturalH * canvasAspect;
                            srcX = (naturalW - srcW) / 2;
                        } else {
                            srcW = naturalW;
                            srcH = naturalW / canvasAspect;
                            srcY = (naturalH - srcH) / 2;
                        }
                        ctx.drawImage(htmlImg, srcX, srcY, srcW, srcH, 0, 0, cvWidthPx, cvHeightPx);
                        const croppedImageDataUrl = tempCanvas.toDataURL(format === 'PNG' ? 'image/png' : 'image/jpeg', 0.9);
                        pdf.addImage(croppedImageDataUrl, format === 'PNG' ? 'PNG' : 'JPEG', x, y, w, h);
                    }
                } else {
                    const imageAspectRatio = naturalW / naturalH;
                    const boxAspectRatio = w / h;
                    let drawW, drawH, drawX, drawY;
                    if (imageAspectRatio > boxAspectRatio) {
                        drawW = w;
                        drawH = w / imageAspectRatio;
                    } else {
                        drawH = h;
                        drawW = h * imageAspectRatio;
                    }
                    drawX = x + (w - drawW) / 2;
                    drawY = y + (h - drawH) / 2;
                    pdf.addImage(imageDataUrl, format, drawX, drawY, drawW, drawH);
                }

            } catch (e) { console.error("Error adding image " + selector + " to PDF:", e); pdf.text(`[Img Proc Err]`, x + 2, y + h / 2); }
        }
    }
};


async function drawPdfHeader(pdf, container, currentY) {
    const headerElement = container.querySelector('.header');
    if (!headerElement) return currentY;
    const panelTopY = currentY, headerPanelHeight = 16;
    const logoHeight = 12, logoSabespWidth = 25, logoConsorcioWidth = 55, logoConsorcioHeight = 16;

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, panelTopY, margin + contentWidth, panelTopY);
    pdf.line(margin, panelTopY, margin, panelTopY + headerPanelHeight);
    pdf.line(margin + contentWidth, panelTopY, margin + contentWidth, panelTopY + headerPanelHeight);

    const internalPaddingX = 1;
    const sabespLogoY = panelTopY + 2;
    await addImageToPdfAsync(pdf, '.sabesp-logo', margin - 3, sabespLogoY, logoSabespWidth, logoHeight, headerElement);

    const consorcioLogoX = pageWidth - margin - logoConsorcioWidth - internalPaddingX;
    const consorcioLogoY = panelTopY + (headerPanelHeight - logoConsorcioHeight) / 2;
    await addImageToPdfAsync(pdf, '.consorcio-logo', consorcioLogoX, consorcioLogoY, logoConsorcioWidth, logoConsorcioHeight, headerElement);

    const titleCenterX = pageWidth / 2 - 20;
    pdf.setFontSize(10); pdf.setFont('helvetica', 'bold');
    pdf.text("RDO - RELATÓRIO DIÁRIO DE OBRAS", titleCenterX, panelTopY + 5.6, { align: 'center' });
    pdf.setFontSize(7);
    pdf.text("SABESP - COMPANHIA DE SANEAMENTO BÁSICO DO ESTADO DE SÃO PAULO", titleCenterX, panelTopY + 9.1, { align: 'center' });
    pdf.text("AMPLIAÇÃO DA CAPACIDADE DE TRATAMENTO DA FASE SÓLIDA-ETE BARUERI PARA 16M³/S", titleCenterX, panelTopY + 12.1, { align: 'center' });

    return panelTopY + headerPanelHeight;
}

function drawContractInfo(pdf, pageContainer, currentY, dayIndex, pageInDayIdx) {
    const daySuffix = `_day${dayIndex}`;
    let idSuffixForPageFields = daySuffix;
    let rdoNumIdForPage = `numero_rdo${daySuffix}`;

    if (pageInDayIdx === 1) {
        idSuffixForPageFields = `_pt${daySuffix}`;
        rdoNumIdForPage = `numero_rdo_pt${daySuffix}`;
    } else if (pageInDayIdx === 2) {
        idSuffixForPageFields = `_pt2${daySuffix}`;
        rdoNumIdForPage = `numero_rdo_pt2${daySuffix}`;
    }

    const rdoNumero = getInputValue(rdoNumIdForPage, pageContainer);
    const contratada = getInputValue(`contratada${idSuffixForPageFields}`, pageContainer);

    const dataInputEl = getElementByIdSafe(`data${idSuffixForPageFields}`, pageContainer);
    let dataVal = dataInputEl ? dataInputEl.value : '';
    let displayData = dataVal;

    if (dataVal.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const dateParts = dataVal.split('-');
        displayData = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
    }


    const contratoNum = getInputValue(`contrato_num${idSuffixForPageFields}`, pageContainer);
    const diaSemana = getInputValue(`dia_semana${idSuffixForPageFields}`, pageContainer);
    const prazo = getInputValue(`prazo${idSuffixForPageFields}`, pageContainer);
    const inicioObra = getInputValue(`inicio_obra${idSuffixForPageFields}`, pageContainer);
    const decorridos = getInputValue(`decorridos${idSuffixForPageFields}`, pageContainer);
    const restantes = getInputValue(`restantes${idSuffixForPageFields}`, pageContainer);

    const boxHeight = 17;
    const mainInfoWidth = contentWidth - 33;
    const rdoBoxWidth = 33;
    const boxTopY = currentY;
    const boxBottomY = currentY + boxHeight;

    pdf.setLineWidth(lineWeight);
    pdf.line(margin, boxTopY, margin + contentWidth, boxTopY);

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, boxTopY, margin, boxBottomY);
    pdf.line(margin + mainInfoWidth, boxTopY, margin + mainInfoWidth, boxBottomY);
    pdf.line(margin, boxBottomY, margin + mainInfoWidth, boxBottomY);

    pdf.setLineWidth(lineWeightStrong);
    const rdoBoxRealX = margin + mainInfoWidth;
    pdf.line(rdoBoxRealX + rdoBoxWidth, boxTopY, rdoBoxRealX + rdoBoxWidth, boxBottomY);
    pdf.line(rdoBoxRealX, boxBottomY, rdoBoxRealX + rdoBoxWidth, boxBottomY);

    pdf.setFontSize(7); pdf.setFont('helvetica', 'normal');
    const rowHeight = boxHeight / 3, textBaselineOffset = rowHeight - 1.5;
    pdf.setLineWidth(lineWeight);

    pdf.text(`CONTRATADA: ${contratada}`, margin + 1.5, currentY + textBaselineOffset);
    pdf.line(margin + mainInfoWidth * 0.65, currentY, margin + mainInfoWidth * 0.65, currentY + rowHeight);
    pdf.text(`DATA: ${displayData}`, margin + mainInfoWidth * 0.65 + 1.5, currentY + textBaselineOffset);

    pdf.line(margin, currentY + rowHeight, margin + mainInfoWidth, currentY + rowHeight);
    pdf.text(`CONTRATO Nº: ${contratoNum}`, margin + 1.5, currentY + rowHeight + textBaselineOffset);
    pdf.line(margin + mainInfoWidth * 0.65, currentY + rowHeight, margin + mainInfoWidth * 0.65, currentY + rowHeight * 2);
    pdf.text(`DIA DA SEMANA: ${diaSemana}`, margin + mainInfoWidth * 0.65 + 1.5, currentY + rowHeight + textBaselineOffset);

    pdf.line(margin, currentY + rowHeight * 2, margin + mainInfoWidth, currentY + rowHeight * 2);
    const textYRow3 = currentY + rowHeight * 2 + textBaselineOffset;
    const segmentWidth = (mainInfoWidth - 3) / 4;
    pdf.text(`PRAZO: ${prazo}`, margin + 1.5, textYRow3);
    pdf.text(`INÍCIO: ${inicioObra}`, margin + 1.5 + segmentWidth, textYRow3);
    pdf.text(`DECORRIDOS: ${decorridos}`, margin + 1.5 + (2 * segmentWidth), textYRow3);
    pdf.text(`RESTANTES: ${restantes}`, margin + 1.5 + (3 * segmentWidth), textYRow3);

    const rdoPdfBoxX = margin + mainInfoWidth;
    pdf.setFontSize(8); pdf.setFont('helvetica', 'bold');
    pdf.text("Número:", rdoPdfBoxX + rdoBoxWidth / 2, currentY + 4.5, { align: 'center' });
    pdf.setFontSize(20); pdf.setFont('helvetica', 'bold'); pdf.setTextColor(255, 0, 0);
    pdf.text(rdoNumero, rdoPdfBoxX + rdoBoxWidth / 2, currentY + 11.5, { align: 'center' });
    pdf.setTextColor(0, 0, 0); pdf.setFont('helvetica', 'normal');

    return currentY + boxHeight;
}

function drawWeatherConditions(pdf, pageContainer, currentY, dayIndex) {
    const weatherEl = pageContainer.querySelector('.weather-conditions');
    if (!weatherEl) return currentY;
    const daySuffix = `_day${dayIndex}`;
    const sectionHeight = 25, titleBgColor = '#E0E0E0', titleHeight = 5;
    const modeloPanelWidth = 35, turnosObsCombinedWidth = contentWidth - modeloPanelWidth;
    const turnosTitleSectionWidth = turnosObsCombinedWidth * 0.45, obsTitleSectionWidth = turnosObsCombinedWidth * 0.55;

    const modeloPanelX = margin;
    const modeloPanelY = currentY;
    const turnosObsBaseX = modeloPanelX + modeloPanelWidth;

    pdf.setFillColor(titleBgColor);
    pdf.rect(modeloPanelX, modeloPanelY, modeloPanelWidth, titleHeight, 'F');
    pdf.rect(turnosObsBaseX, modeloPanelY, turnosObsCombinedWidth, titleHeight, 'F');

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(modeloPanelX, modeloPanelY, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY);
    pdf.line(modeloPanelX, modeloPanelY + sectionHeight, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY + sectionHeight);
    pdf.line(modeloPanelX, modeloPanelY, modeloPanelX, modeloPanelY + sectionHeight);
    pdf.line(turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY + sectionHeight);

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(modeloPanelX, modeloPanelY + titleHeight, turnosObsBaseX + turnosObsCombinedWidth, modeloPanelY + titleHeight);

    pdf.setLineWidth(lineWeight);
    pdf.line(modeloPanelX + modeloPanelWidth, modeloPanelY + titleHeight, modeloPanelX + modeloPanelWidth, modeloPanelY + sectionHeight);

    const modeloContentY = modeloPanelY + titleHeight;
    const modeloContentHeight = sectionHeight - titleHeight;
    const yShiftDown = 0.75;
    pdf.setFont('helvetica', 'bold'); pdf.setFontSize(9); pdf.setTextColor(0,0,0);
    pdf.text("Modelo", modeloPanelX + 10, modeloContentY + modeloContentHeight / 2 + 6, { align: 'center', angle: 90 });

    const labelsX = modeloPanelX + 8;
    pdf.setFont('helvetica', 'normal'); pdf.setFontSize(7);
    pdf.text("Tempo", labelsX, modeloContentY + (modeloContentHeight / 3) + yShiftDown, { maxWidth: 15 });
    pdf.text("Trabalho", labelsX, modeloContentY + (modeloContentHeight / 3 * 2) + yShiftDown, { maxWidth: 15 });

    const staticCircleRadius = 6, staticCircleCenterX = (modeloPanelX + modeloPanelWidth) - 3 - staticCircleRadius;
    const staticCircleCenterY = modeloContentY + modeloContentHeight / 2 + yShiftDown;
    const circleTextBaselineOffsetY = 1.8;

    pdf.setLineWidth(lineWeight);
    pdf.circle(staticCircleCenterX, staticCircleCenterY, staticCircleRadius, 'S');
    pdf.line(staticCircleCenterX - staticCircleRadius, staticCircleCenterY, staticCircleCenterX + staticCircleRadius, staticCircleCenterY);
    pdf.setFontSize(10); pdf.setFont('helvetica', 'bold');
    pdf.text("B", staticCircleCenterX, staticCircleCenterY - staticCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' });
    pdf.text("N", staticCircleCenterX, staticCircleCenterY + staticCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' });

    pdf.setTextColor(0,0,0); pdf.setFontSize(9); pdf.setFont('helvetica', 'bold');
    pdf.text('CONDIÇÕES CLIMÁTICAS E DE TRABALHO', turnosObsBaseX + turnosTitleSectionWidth / 2, modeloPanelY + 3.5, { align: 'center' });

    const indicePluvTitleX = turnosObsBaseX + turnosTitleSectionWidth;
    pdf.setFont('helvetica', 'bold'); pdf.setFontSize(7);
    const indicePluv = getInputValue(`indice_pluv_valor${daySuffix}`, weatherEl);
    pdf.text(`Índice Pluviométrico: ${indicePluv} mm`, indicePluvTitleX + 1.5, modeloPanelY + 3.2, { maxWidth: obsTitleSectionWidth - 3 });
    pdf.setFont('helvetica', 'normal'); pdf.setFontSize(7);

    const turnosObsContentY = modeloPanelY + titleHeight;
    const turnosObsContentHeight = sectionHeight - titleHeight;
    const turnoDetailHeaderHeight = 7, turnoLabelTextYBaselineOffset = 3.0, turnoTimeTextYBaselineOffset = 6.0;

    const turnoLabelCells = weatherEl.querySelectorAll('.turno-label-cell');
    const turnosDataFromDOM = Array.from(turnoLabelCells).slice(0, 3).map(thCell => {
        const label = Array.from(thCell.childNodes).find(node => node.nodeType === Node.TEXT_NODE)?.textContent?.trim() || '';
        const time = thCell.querySelector('.turno-time-cell')?.textContent?.trim() || '';
        return { label, time };
    });

    const turnoColWidth = turnosTitleSectionWidth / 3;
    turnosDataFromDOM.forEach((turnoData, index) => {
        const turnoColActualX = turnosObsBaseX + index * turnoColWidth;
        pdf.setFontSize(7); pdf.setFont('helvetica', 'bold');
        pdf.text(turnoData.label, turnoColActualX + turnoColWidth / 2, turnosObsContentY + turnoLabelTextYBaselineOffset, { align: 'center' });
        if (turnoData.time) {
            pdf.setFontSize(6); pdf.setFont('helvetica', 'normal');
            pdf.text(turnoData.time, turnoColActualX + turnoColWidth / 2, turnosObsContentY + turnoTimeTextYBaselineOffset, { align: 'center' });
        }
    });

    const circleAreaYStart = turnosObsContentY + turnoDetailHeaderHeight;
    const circleAreaHeight = turnosObsContentHeight - turnoDetailHeaderHeight;
    const dynamicCircleRadius = 6;
    const turnosControlValueGetters = [
        { cellIdBase: 'turno1_control_cell', turnoIdBase: 't1' },
        { cellIdBase: 'turno2_control_cell', turnoIdBase: 't2' },
        { cellIdBase: 'turno3_control_cell', turnoIdBase: 't3' }
    ];
    turnosControlValueGetters.forEach((turnoCtrl, index) => {
        const turnoColActualX = turnosObsBaseX + index * turnoColWidth;
        const topValue = getInputValue(`tempo_${turnoCtrl.turnoIdBase}${daySuffix}_hidden`, pageContainer);
        const bottomValue = getInputValue(`trabalho_${turnoCtrl.turnoIdBase}${daySuffix}_hidden`, pageContainer);

        const circleCenterX = turnoColActualX + turnoColWidth / 2;
        const circleCenterY = circleAreaYStart + (circleAreaHeight / 2);
        pdf.setLineWidth(lineWeight);
        pdf.circle(circleCenterX, circleCenterY, dynamicCircleRadius, 'S');
        pdf.line(circleCenterX - dynamicCircleRadius, circleCenterY, circleCenterX + dynamicCircleRadius, circleCenterY);
        pdf.setFontSize(10); pdf.setFont('helvetica', 'bold');
        pdf.text(topValue, circleCenterX, circleCenterY - dynamicCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' });
        pdf.text(bottomValue, circleCenterX, circleCenterY + dynamicCircleRadius / 2 + circleTextBaselineOffsetY, { align: 'center' });
    });

    pdf.setLineWidth(lineWeight);
    pdf.line(indicePluvTitleX, turnosObsContentY, indicePluvTitleX, turnosObsContentY + turnosObsContentHeight);

    pdf.setFontSize(7); pdf.setFont('helvetica', 'normal');
    const obsClima = getTextAreaValue(`obs_clima${daySuffix}`, weatherEl);
    pdf.text(`Obs: ${obsClima}`, indicePluvTitleX + 1.5, turnosObsContentY + 3.5, { maxWidth: obsTitleSectionWidth - 3, align: 'left' });

    return currentY + sectionHeight;
}

function drawLegend(pdf, pageContainer, currentY) {
    const legendElement = pageContainer.querySelector('.legend');
    if (!legendElement) return currentY;

    const legendText = (legendElement.textContent || '').replace(/\s\s+/g, ' ').trim();
    const textHorizontalPadding = 2, textVerticalPadding = 1.5, legendFontSize = 6.5;

    pdf.setFontSize(legendFontSize); pdf.setFont('helvetica', 'normal');
    const textLines = pdf.splitTextToSize(legendText, contentWidth - (2 * textHorizontalPadding));
    const textDimensions = pdf.getTextDimensions(textLines, { fontSize: legendFontSize, maxWidth: contentWidth - (2 * textHorizontalPadding) });
    const legendTextActualHeight = textDimensions.h;
    const totalSectionHeight = legendTextActualHeight + (2 * textVerticalPadding);

    pdf.setLineWidth(lineWeightStrong); pdf.rect(margin, currentY, contentWidth, totalSectionHeight, 'S');

    const ascent = (legendFontSize * (25.4 / 72)) * 0.8;
    let firstLineBaselineY = currentY + textVerticalPadding + ascent;
    const singleLineHeight = legendTextActualHeight / textLines.length;

    const prefixToBold = "Legendas:";
    for (let i = 0; i < textLines.length; i++) {
        const line = textLines[i], lineY = firstLineBaselineY + (i * singleLineHeight);
        let currentTextX = margin + textHorizontalPadding;

        if (i === 0 && line.startsWith(prefixToBold)) {
            const boldPart = prefixToBold;
            const normalPart = line.substring(boldPart.length);
            pdf.setFont('helvetica', 'bold'); pdf.setFontSize(legendFontSize);
            pdf.text(boldPart, currentTextX, lineY);
            const boldPartWidth = pdf.getTextWidth(boldPart); currentTextX += boldPartWidth;
            pdf.setFont('helvetica', 'normal'); pdf.setFontSize(legendFontSize);
            pdf.text(normalPart, currentTextX, lineY, { maxWidth: contentWidth - (2 * textHorizontalPadding) - boldPartWidth });
        } else {
            pdf.setFont('helvetica', 'normal'); pdf.setFontSize(legendFontSize);
            pdf.text(line, currentTextX, lineY, { maxWidth: contentWidth - (2 * textHorizontalPadding) });
        }
    }
    return currentY + totalSectionHeight;
}

function drawLaborColumns(pdf, pageContainer, currentY, dayIndex) {
    const laborSectionEl = pageContainer.querySelector('.labor-section');
    if (!laborSectionEl) return currentY;

    const columnsConfig = [
        { selector: '.labor-direta', proportion: 1, isMultiSubColumn: false, totalLabel: "Total M.O.D = " },
        { selector: '.labor-indireta', proportion: 2, isMultiSubColumn: true, totalLabel: "Total M.O.I = " },
        { selector: '.equipamentos', proportion: 2, isMultiSubColumn: true, totalLabel: "Total Equip. = " }
    ];
    const totalProportions = columnsConfig.reduce((sum, col) => sum + col.proportion, 0);

    const titleHeight = 5, itemRowHeight = 3.5, itemsPerVisualRowStack = 18;
    const itemFontSize = 6, textBaselineInRow = itemRowHeight - 1.285;
    const qtyDisplayWidthInPdf = 10, itemNamePaddingLeft = 2;

    pdf.setFont('helvetica', 'normal');
    let columnSectionStartY = currentY;
    const colContentsHeight = itemsPerVisualRowStack * itemRowHeight;
    const totalRowHeight = 5;
    const totalSectionHeight = titleHeight + colContentsHeight + totalRowHeight;

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, columnSectionStartY, margin, columnSectionStartY + totalSectionHeight);
    pdf.line(margin + contentWidth, columnSectionStartY, margin + contentWidth, columnSectionStartY + totalSectionHeight);
    pdf.line(margin, columnSectionStartY + totalSectionHeight, margin + contentWidth, columnSectionStartY + totalSectionHeight);

    let currentCumulativeX = margin;

    columnsConfig.forEach((colConfig, idx) => {
        const colEl = laborSectionEl.querySelector(colConfig.selector);
        if (!colEl) return;

        const currentColWidth = (contentWidth / totalProportions) * colConfig.proportion;
        const title = (colEl.querySelector('h4')?.textContent || '').trim();

        pdf.setLineWidth(lineWeight);
        pdf.setFillColor('#E0E0E0'); pdf.setFontSize(9); pdf.setFont('helvetica', 'bold');
        pdf.rect(currentCumulativeX, columnSectionStartY, currentColWidth, titleHeight, 'FD');
        pdf.text(title, currentCumulativeX + currentColWidth / 2, columnSectionStartY + 3.5, { align: 'center' });

        let itemAreaY = columnSectionStartY + titleHeight;
        const table = colEl.querySelector('.labor-table');
        pdf.setFontSize(itemFontSize); pdf.setFont('helvetica', 'normal'); pdf.setLineWidth(lineWeight);

        const quantityInputs = Array.from(table.querySelectorAll('tbody tr .quantity-input'));
        const itemNames = Array.from(table.querySelectorAll('tbody tr .item-name-cell')).map(el => el.textContent?.trim() || '');

        if (colConfig.isMultiSubColumn) {
            const subColWidth = (currentColWidth - lineWeight) / 2;

            let subColStartX = currentCumulativeX;
            for (let i = 0; i < itemsPerVisualRowStack; i++) {
                if (i >= quantityInputs.length / 2 && i >= itemNames.length / 2) break;
                const currentItemY = itemAreaY + i * itemRowHeight;
                const qty = quantityInputs[i]?.value || "";
                const name = itemNames[i] || "";

                pdf.text(qty, subColStartX + qtyDisplayWidthInPdf - 2, currentItemY + textBaselineInRow, { align: 'right', maxWidth: qtyDisplayWidthInPdf - 4 });
                pdf.line(subColStartX + qtyDisplayWidthInPdf, currentItemY, subColStartX + qtyDisplayWidthInPdf, currentItemY + itemRowHeight);
                pdf.text(name, subColStartX + qtyDisplayWidthInPdf + itemNamePaddingLeft, currentItemY + textBaselineInRow, { maxWidth: subColWidth - qtyDisplayWidthInPdf - itemNamePaddingLeft - 1 });

                const isLastTrueItemInSubColumn = i === Math.floor(Math.min(quantityInputs.length, itemNames.length) / 2) -1 && (Math.min(quantityInputs.length,itemNames.length)/2 <= itemsPerVisualRowStack);
                if (i < itemsPerVisualRowStack - 1 && !isLastTrueItemInSubColumn && i < (Math.min(quantityInputs.length,itemNames.length)/2) -1) {
                     pdf.line(subColStartX, currentItemY + itemRowHeight, subColStartX + subColWidth, currentItemY + itemRowHeight);
                }
            }

            subColStartX = currentCumulativeX + subColWidth + lineWeight;
            for (let i = 0; i < itemsPerVisualRowStack; i++) {
                const itemGlobalIndex = itemsPerVisualRowStack + i;
                if (itemGlobalIndex >= quantityInputs.length && itemGlobalIndex >= itemNames.length) break;
                const currentItemY = itemAreaY + i * itemRowHeight;
                const qty = quantityInputs[itemGlobalIndex]?.value || "";
                const name = itemNames[itemGlobalIndex] || "";

                pdf.text(qty, subColStartX + qtyDisplayWidthInPdf - 2, currentItemY + textBaselineInRow, { align: 'right', maxWidth: qtyDisplayWidthInPdf - 4 });
                pdf.line(subColStartX + qtyDisplayWidthInPdf, currentItemY, subColStartX + qtyDisplayWidthInPdf, currentItemY + itemRowHeight);
                pdf.text(name, subColStartX + qtyDisplayWidthInPdf + itemNamePaddingLeft, currentItemY + textBaselineInRow, { maxWidth: subColWidth - qtyDisplayWidthInPdf - itemNamePaddingLeft - 1 });

                const isLastTrueItemInThisSubColumn = (itemsPerVisualRowStack + i) === (Math.min(quantityInputs.length, itemNames.length) -1) && (itemsPerVisualRowStack + i < itemsPerVisualRowStack * 2);

                if (i < itemsPerVisualRowStack - 1 && !isLastTrueItemInThisSubColumn && (itemsPerVisualRowStack + i < (Math.min(quantityInputs.length, itemNames.length) -1 ))) {
                     pdf.line(subColStartX, currentItemY + itemRowHeight, subColStartX + subColWidth, currentItemY + itemRowHeight);
                }
            }
            pdf.line(currentCumulativeX + subColWidth, itemAreaY, currentCumulativeX + subColWidth, itemAreaY + colContentsHeight);
        } else {
            for (let i = 0; i < itemsPerVisualRowStack; i++) {
                if (i >= quantityInputs.length && i >= itemNames.length) break;
                const currentItemY = itemAreaY + i * itemRowHeight;
                const qty = quantityInputs[i]?.value || "";
                const name = itemNames[i] || "";

                pdf.text(qty, currentCumulativeX + qtyDisplayWidthInPdf - 2, currentItemY + textBaselineInRow, { align: 'right', maxWidth: qtyDisplayWidthInPdf - 4 });
                pdf.line(currentCumulativeX + qtyDisplayWidthInPdf, currentItemY, currentCumulativeX + qtyDisplayWidthInPdf, currentItemY + itemRowHeight);
                pdf.text(name, currentCumulativeX + qtyDisplayWidthInPdf + itemNamePaddingLeft, currentItemY + textBaselineInRow, { maxWidth: currentColWidth - qtyDisplayWidthInPdf - itemNamePaddingLeft - 1 });

                const isLastTrueItemOverall = i === Math.min(quantityInputs.length, itemNames.length) -1 && Math.min(quantityInputs.length, itemNames.length) <= itemsPerVisualRowStack;
                if (i < itemsPerVisualRowStack - 1 && !isLastTrueItemOverall && i < Math.min(quantityInputs.length, itemNames.length) -1 ) {
                    pdf.line(currentCumulativeX, currentItemY + itemRowHeight, currentCumulativeX + currentColWidth, currentItemY + itemRowHeight);
                }
            }
        }

        const totalRowY = columnSectionStartY + titleHeight + colContentsHeight;
        pdf.setFontSize(7);
        pdf.setFont('helvetica', 'bold');
        pdf.setLineWidth(lineWeight);

        pdf.rect(currentCumulativeX, totalRowY, currentColWidth, totalRowHeight, 'S');

        const totalValue = (colEl.querySelector('.section-total-input'))?.value || "0";
        const totalText = colConfig.totalLabel + totalValue;

        const textPaddingRight = 2;
        pdf.text(totalText, currentCumulativeX + currentColWidth - textPaddingRight, totalRowY + 3.5, {
            align: 'right',
            maxWidth: currentColWidth - (textPaddingRight * 2)
        });

        if (idx < columnsConfig.length - 1) {
            pdf.setLineWidth(lineWeightStrong);
            pdf.line(currentCumulativeX + currentColWidth, columnSectionStartY, currentCumulativeX + currentColWidth, columnSectionStartY + totalSectionHeight);
        }
        currentCumulativeX += currentColWidth;
    });
    return columnSectionStartY + totalSectionHeight;
}


function drawActivityTable(pdf, pageContainer, currentY, dayIndex, pageInDayIdx, targetTotalHeightForSection) {
    let tableIdBase = "activitiesTable";
    let numRowsToDraw = 24;
    const sectionTitle = "DESCRIÇÃO DAS ATIVIDADES EXECUTADAS";

    if (pageInDayIdx === 1) {
        tableIdBase = "activitiesTable_pt";
        numRowsToDraw = 22;
    }

    const tableElement = getElementByIdSafe(`${tableIdBase}_day${dayIndex}`, pageContainer);
    if (!tableElement) return currentY;

    const tableHead = tableElement.tHead;
    const tableBody = tableElement.tBodies[0];
    if (!tableHead || !tableBody) return currentY;

    const titleHeight = 5, headerRowHeight = 5;
    let activityRowHeight = 5.25;
    const MIN_ACTIVITY_ROW_PDF_HEIGHT = 3.5;

    if (pageInDayIdx === 0 && targetTotalHeightForSection !== undefined && targetTotalHeightForSection > 0) {
        const heightForRowsArea = targetTotalHeightForSection - titleHeight - headerRowHeight;
        if (heightForRowsArea > 0 && numRowsToDraw > 0) {
            activityRowHeight = Math.max(heightForRowsArea / numRowsToDraw, MIN_ACTIVITY_ROW_PDF_HEIGHT);
        } else if (heightForRowsArea <=0) {
            activityRowHeight = MIN_ACTIVITY_ROW_PDF_HEIGHT;
        }
    }

    const fontSizeHeader = 7, fontSizeBody = 6.5;

    const colWidthsConfig = [
        { proportion: 24.14, title: 'Localização', dataKey: `localizacao_day${dayIndex}[]` },
        { proportion: 13.79, title: 'Tipo', dataKey: `tipo_servico_day${dayIndex}[]` },
        { proportion: 13.79, title: 'Serviço', dataKey: `servico_desc_day${dayIndex}[]` },
        { proportion: 10.34, title: 'Status', dataKey: `status_day${dayIndex}[]` },
        { proportion: 37.93, title: 'Observações', dataKey: `observacoes_atividade_day${dayIndex}[]` },
    ];

    const totalProportions = colWidthsConfig.reduce((sum, col) => sum + col.proportion, 0);
    const calculatedColWidths = colWidthsConfig.map(c => (contentWidth / totalProportions) * c.proportion);

    const initialCurrentY = currentY;

    pdf.setLineWidth(lineWeightStrong);
    pdf.setFillColor('#E0E0E0'); pdf.setFontSize(8); pdf.setFont('helvetica', 'bold');
    pdf.rect(margin, currentY, contentWidth, titleHeight, 'FD');
    pdf.text(sectionTitle, margin + contentWidth / 2, currentY + 3.5, { align: 'center' });
    currentY += titleHeight;

    const tableHeaderY = currentY;
    pdf.setLineWidth(lineWeight); pdf.setFillColor('#E0E0E0'); pdf.setFontSize(fontSizeHeader); pdf.setFont('helvetica', 'bold');
    pdf.rect(margin, currentY, contentWidth, headerRowHeight, 'FD');
    let currentX = margin;
    colWidthsConfig.forEach((col, idx) => {
        pdf.text(col.title, currentX + calculatedColWidths[idx] / 2, currentY + 3.5, { align: 'center', maxWidth: calculatedColWidths[idx] - 2 });
        if (idx < colWidthsConfig.length -1) pdf.line(currentX + calculatedColWidths[idx], currentY, currentX + calculatedColWidths[idx], currentY + headerRowHeight);
        currentX += calculatedColWidths[idx];
    });
    currentY += headerRowHeight;

    pdf.setFontSize(fontSizeBody); pdf.setFont('helvetica', 'normal');
    pdf.setLineWidth(lineWeight);
    const rows = Array.from(tableBody.rows).slice(0, numRowsToDraw);

    const pointToMm = 0.352778;
    const lineHeightFactor = 1.15;
    const singleLineRenderHeight = fontSizeBody * pointToMm * lineHeightFactor;

    rows.forEach((row, rowIndex) => {
        const rowY = currentY + rowIndex * activityRowHeight;
        currentX = margin;

        colWidthsConfig.forEach((colConfig, cellIndex) => {
            const cellWidth = calculatedColWidths[cellIndex];
            let text = "";

            if (colConfig.dataKey.startsWith(`status_day${dayIndex}`) ||
                colConfig.dataKey.startsWith(`localizacao_day${dayIndex}`) ||
                colConfig.dataKey.startsWith(`tipo_servico_day${dayIndex}`) ||
                colConfig.dataKey.startsWith(`servico_desc_day${dayIndex}`)) {
                const hiddenInput = row.cells[cellIndex]?.querySelector('input[type="hidden"]');
                if (hiddenInput) {
                    text = hiddenInput.value;
                }
            } else {
                const inputEl = row.cells[cellIndex]?.querySelector('input[type="text"], textarea');
                if (inputEl) {
                    text = inputEl.value;
                }
            }

            const textLines = pdf.splitTextToSize(text, cellWidth - 2);

            const textBlockRenderHeight = (textLines.length > 0) ? (textLines.length - 1) * singleLineRenderHeight + (fontSizeBody * pointToMm) : 0;
            const yForFirstLine = rowY + (activityRowHeight - textBlockRenderHeight) / 2 + (fontSizeBody * pointToMm * 0.8);

            let currentDrawingY = yForFirstLine;
            for (const line of textLines) {
                 if (currentDrawingY < rowY + activityRowHeight - (fontSizeBody*pointToMm*0.2) ) {
                    pdf.text(line, currentX + 1, currentDrawingY, { maxWidth: cellWidth - 2 });
                 }
                currentDrawingY += singleLineRenderHeight;
            }

            if (cellIndex < colWidthsConfig.length - 1) {
                pdf.line(currentX + cellWidth, rowY, currentX + cellWidth, rowY + activityRowHeight);
            }
            currentX += cellWidth;
        });

        if (!(pageInDayIdx === 0 && rowIndex === rows.length - 1)) {
            pdf.line(margin, rowY + activityRowHeight, margin + contentWidth, rowY + activityRowHeight);
        }
    });

    const finalTableY = currentY + (numRowsToDraw * activityRowHeight);

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, tableHeaderY, margin, finalTableY);
    pdf.line(margin + contentWidth, tableHeaderY, margin + contentWidth, finalTableY);

    return initialCurrentY + (titleHeight + headerRowHeight + (numRowsToDraw * activityRowHeight));
}


function drawObservationsSection(
    pdf,
    pageContainer,
    currentY,
    dayIndex,
    targetTotalHeightForSection
) {
    const daySuffix = `_pt_day${dayIndex}`;
    const obsTextarea = getElementByIdSafe(`observacoes_comentarios${daySuffix}`, pageContainer);
    if (!obsTextarea) return currentY;

    const sectionTitle = "OUTRAS OBSERVAÇÕES E COMENTÁRIOS";
    const titleHeight = 5;
    const fontSizeTitle = 9, fontSizeBody = 7;
    const padding = 1.5;
    const pdfLineHeightMm = 6;
    const MIN_OBSERVATIONS_CONTENT_BOX_HEIGHT_PDF = 65;

    let actualSectionHeight;
    let contentBoxHeight;
    const initialCurrentY = currentY;

    if (targetTotalHeightForSection !== undefined && targetTotalHeightForSection > titleHeight) {
        actualSectionHeight = targetTotalHeightForSection;
        contentBoxHeight = actualSectionHeight - titleHeight;
    } else {
        const obsText = obsTextarea.value;
        pdf.setFontSize(fontSizeBody);
        pdf.setFont('helvetica', 'normal');

        const pointToMm = 0.352778;
        const lineHeightFactorObs = 1.15;
        const singleLineHeightWithFactorObs = fontSizeBody * pointToMm * lineHeightFactorObs;
        const textLines = pdf.splitTextToSize(obsText, contentWidth - (2 * padding));
        const requiredHeightForTextActual = (textLines.length > 0) ? (textLines.length -1) * singleLineHeightWithFactorObs + (fontSizeBody * pointToMm) : 0;

        const calculatedContentHeightFromText = requiredHeightForTextActual + (2 * padding);
        contentBoxHeight = Math.max(calculatedContentHeightFromText, MIN_OBSERVATIONS_CONTENT_BOX_HEIGHT_PDF);
        actualSectionHeight = contentBoxHeight + titleHeight;
    }

    pdf.setLineWidth(lineWeightStrong);
    pdf.setFillColor('#E0E0E0');
    pdf.setFontSize(fontSizeTitle);
    pdf.setFont('helvetica', 'bold');
    pdf.rect(margin, currentY, contentWidth, titleHeight, 'FD');
    pdf.text(sectionTitle, margin + contentWidth / 2, currentY + 3.5, { align: 'center' });
    currentY += titleHeight;

    const contentBoxTopY = currentY;

    pdf.setLineWidth(lineWeight);
    pdf.rect(margin, contentBoxTopY, contentWidth, contentBoxHeight, 'S');

    pdf.setLineWidth(lineWeight);
    pdf.setDrawColor(0, 0, 0);

    let lineRenderY = contentBoxTopY + pdfLineHeightMm;
    const contentBoxBottomYForLines = contentBoxTopY + contentBoxHeight;

    while (lineRenderY < contentBoxBottomYForLines - 0.5) {
        pdf.line(margin, lineRenderY, margin + contentWidth, lineRenderY);
        lineRenderY += pdfLineHeightMm;
    }

    pdf.setDrawColor(0, 0, 0);
    pdf.setLineWidth(lineWeight);

    const obsText = obsTextarea.value;
    pdf.setFontSize(fontSizeBody);
    pdf.setFont('helvetica', 'normal');
    const textLines = pdf.splitTextToSize(obsText, contentWidth - (2 * padding));

    const pointToMm = 0.352778;
    const lineHeightFactorObs = 1.15;
    const singleLineHeightWithFactorObs = fontSizeBody * pointToMm * lineHeightFactorObs;
    const textPrintStartY = contentBoxTopY + padding + (fontSizeBody * pointToMm * 0.8);

    let currentDrawingY = textPrintStartY;
    for (const line of textLines) {
        pdf.text(line, margin + padding, currentDrawingY, {
            maxWidth: contentWidth - (2 * padding)
        });
        currentDrawingY += singleLineHeightWithFactorObs;
        if (currentDrawingY > contentBoxTopY + contentBoxHeight - padding) break;
    }

    return initialCurrentY + actualSectionHeight;
}


async function drawReportPhotosSection(
    pdf,
    pageContainer,
    currentY,
    dayIndex
) {
    const daySuffix = `_day${dayIndex}`;
    const sectionTitle = "RELATÓRIO FOTOGRÁFICO";
    const titleHeight = 5;
    const photoGridMarginBelowTitle = 2;

    const titleActualStartY = currentY;

    pdf.setLineWidth(lineWeightStrong);
    pdf.setFillColor('#E0E0E0');
    pdf.setFontSize(9);
    pdf.setFont('helvetica', 'bold');
    pdf.rect(margin, titleActualStartY, contentWidth, titleHeight, 'FD');
    pdf.text(sectionTitle, margin + contentWidth / 2, titleActualStartY + 3.5, { align: 'center' });

    const photoGridContentStartY = titleActualStartY + titleHeight + photoGridMarginBelowTitle;

    const photoGridPadding = 2;
    const photoFrameImgHeight = 70;
    const captionAreaHeight = 12;
    const interPhotoColumnGap = 10;
    const interPhotoRowGap = 0;
    const imageDisplayPadding = 0.5;

    const photoGridEffectiveWidth = contentWidth - (2 * photoGridPadding);
    const photoCellWidth = (photoGridEffectiveWidth - interPhotoColumnGap) / 2;
    const photoCellHeight = photoFrameImgHeight + captionAreaHeight;
    const photoGridEffectiveHeight = (photoCellHeight * 2) + interPhotoRowGap;

    const gridStartX = margin + photoGridPadding;
    const panelLineStartY = titleActualStartY + titleHeight;

    const signatureSectionHeightConstant = 25;
    const panelLineEndY = pageHeight - margin - signatureSectionHeightConstant;

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, panelLineStartY, margin, panelLineEndY);
    pdf.line(margin + contentWidth, panelLineStartY, margin + contentWidth, panelLineEndY);


    for (let rowIndex = 0; rowIndex < 2; rowIndex++) {
        for (let colIndex = 0; colIndex < 2; colIndex++) {
            const photoNum = rowIndex * 2 + colIndex + 1;

            const cellStartX = gridStartX + (colIndex * (photoCellWidth + interPhotoColumnGap));
            const cellStartY = photoGridContentStartY + (rowIndex * (photoCellHeight + interPhotoRowGap));

            const imageX = cellStartX + imageDisplayPadding;
            const imageY = cellStartY + imageDisplayPadding;
            const imageW = photoCellWidth - (2 * imageDisplayPadding);
            const imageH = photoFrameImgHeight - (2 * imageDisplayPadding);

            const imgPreviewId = `#report_photo_${photoNum}_pt2_preview${daySuffix}`;
            await addImageToPdfAsync(pdf, imgPreviewId, imageX, imageY, imageW, imageH, pageContainer);

            pdf.setLineWidth(photoFrameLineWeight);
            pdf.rect(imageX, imageY, imageW, imageH, 'S');

            const captionContentRenderAreaY = cellStartY + photoFrameImgHeight + 1.5;
            const captionFontSize = 6.5;
            const pointToMm = 0.352778;
            const captionFontSizeInMm = captionFontSize * pointToMm;
            const captionLineHeightFactor = 1.1;
            const captionBaselineToBaselineSeparation = captionFontSizeInMm * captionLineHeightFactor;

            const captionLabel = `Foto ${String(photoNum).padStart(2, '0')}:`;
            pdf.setFontSize(captionFontSize);
            pdf.setFont('helvetica', 'bold');

            const labelBaselineY = captionContentRenderAreaY + captionFontSizeInMm * 0.8;
            pdf.text(captionLabel, cellStartX + 1.5, labelBaselineY);

            pdf.setFont('helvetica', 'normal');
            const captionLabelWidth = pdf.getTextWidth(captionLabel);
            const captionTextValue = getInputValue(`report_photo_${photoNum}_pt2_caption${daySuffix}`, pageContainer);
            const captionTextStartX = cellStartX + 1.5 + captionLabelWidth + 1;
            const captionTextMaxWidth = photoCellWidth - (1.5 + captionLabelWidth + 1 + 1.5);

            if (captionTextValue.trim() !== "") {
                const textLines = pdf.splitTextToSize(captionTextValue, captionTextMaxWidth);
                let currentDrawingTextY = labelBaselineY;
                for (const line of textLines) {
                    if (currentDrawingTextY < (cellStartY + photoFrameImgHeight + captionAreaHeight) - (captionFontSizeInMm * 0.1)) {
                         pdf.text(line, currentDrawingTextY, currentDrawingTextY, { maxWidth: captionTextMaxWidth });
                    }
                    currentDrawingTextY += captionBaselineToBaselineSeparation;
                }
            }
        }
    }
    const finalYAfterSection = titleActualStartY + titleHeight + photoGridMarginBelowTitle + photoGridEffectiveHeight;
    return finalYAfterSection;
}


async function drawSignatureSection(pdf, pageContainer, currentY, dayIndex, pageInDayIdx) {
    const daySuffix = `_day${dayIndex}`;
    const sectionTopY = currentY;
    const sectionHeight = 25;
    const titleHeight = 5;
    const contentAreaHeight = sectionHeight - titleHeight;
    const sectionBottomY = sectionTopY + sectionHeight;
    const sectionRightX = margin + contentWidth;

    const singleSignatureAreaWidth = contentWidth / 2;
    const actualSignatureLineWidth = singleSignatureAreaWidth * 0.7;

    pdf.setLineWidth(lineWeight);
    pdf.line(margin, sectionTopY, sectionRightX, sectionTopY);

    pdf.setLineWidth(lineWeightStrong);
    pdf.line(margin, sectionTopY, margin, sectionBottomY);
    pdf.line(sectionRightX, sectionTopY, sectionRightX, sectionBottomY);
    pdf.line(margin, sectionBottomY, sectionRightX, sectionBottomY);

    pdf.setLineWidth(lineWeight);
    pdf.setFillColor('#E0E0E0');
    pdf.rect(margin, sectionTopY, contentWidth, titleHeight, 'FD');

    pdf.setFontSize(8);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(0, 0, 0);
    pdf.text("ASSINATURAS", margin + contentWidth / 2, sectionTopY + titleHeight / 2 + 1.5, { align: 'center' });

    pdf.setLineWidth(lineWeight);

    const contentTopY = sectionTopY + titleHeight;

    const leftSigAreaX = margin;
    const imgMaxH = contentAreaHeight * 0.55;
    const imgMaxW = actualSignatureLineWidth;
    const imgX = leftSigAreaX + (singleSignatureAreaWidth - imgMaxW) / 2;
    const imgY = contentTopY + 1.5;

    let consbemPreviewSelector = `#consbemPhotoPreview${daySuffix}`;
    if (pageInDayIdx === 1) {
        consbemPreviewSelector = `#consbemPhotoPreview_pt${daySuffix}`;
    } else if (pageInDayIdx === 2) {
        consbemPreviewSelector = `#consbemPhotoPreview_pt2${daySuffix}`;
    }

    await addImageToPdfAsync(pdf, consbemPreviewSelector, imgX, imgY, imgMaxW, imgMaxH, pageContainer);

    const leftLineY = imgY + imgMaxH + 2;
    const leftLineXStart = leftSigAreaX + (singleSignatureAreaWidth - actualSignatureLineWidth) / 2;
    const leftLineXEnd = leftLineXStart + actualSignatureLineWidth;
    pdf.line(leftLineXStart, leftLineY, leftLineXEnd, leftLineY);

    const leftLabelY = leftLineY + 3.5;
    pdf.setFontSize(7);
    pdf.setFont('helvetica', 'normal');
    pdf.text("RESPONSAVEL CONSBEM", leftSigAreaX + singleSignatureAreaWidth / 2, leftLabelY, { align: 'center' });

    const rightSigAreaX = margin + singleSignatureAreaWidth;
    const rightLineY = leftLineY;
    const rightLineXStart = rightSigAreaX + (singleSignatureAreaWidth - actualSignatureLineWidth) / 2;
    const rightLineXEnd = rightLineXStart + actualSignatureLineWidth;
    pdf.line(rightLineXStart, rightLineY, rightLineXEnd, rightLineY);

    const rightLabelY = leftLabelY;
    pdf.text("FISCALIZAÇÃO SABESP", rightSigAreaX + singleSignatureAreaWidth / 2, rightLabelY, { align: 'center' });

    pdf.setTextColor(0, 0, 0);

    return sectionTopY + sectionHeight;
}

async function generateSingleDayPdf(pdf, dayIndex, dayWrapperElement, isFirstDay, progressUpdateFn) {
    const pageContainers = Array.from(dayWrapperElement.querySelectorAll('.rdo-container'));
    const totalPagesInDay = pageContainers.length;

    for (let pageIdx = 0; pageIdx < totalPagesInDay; pageIdx++) {
        const pageContainer = pageContainers[pageIdx];
        if (!isFirstDay || pageIdx > 0) {
            pdf.addPage();
        }

        if (progressUpdateFn) {
             // For a more granular progress, you might pass pageIdx and totalPagesInDay
            // For now, the main progress is per day, so this call might be redundant if already called per day
            // Or, it could update a sub-detail text:
            // progressUpdateFn(dayIndex + 1, `Processando Dia ${dayIndex + 1}, Página ${pageIdx + 1}/${totalPagesInDay}...`);
        }


        let currentY = margin;
        currentY = await drawPdfHeader(pdf, pageContainer, currentY);
        currentY = drawContractInfo(pdf, pageContainer, currentY, dayIndex, pageIdx);

        const signatureSectionHeight = 25;

        if (pageIdx === 0) {
            currentY = drawWeatherConditions(pdf, pageContainer, currentY, dayIndex);
            currentY = drawLegend(pdf, pageContainer, currentY);
            currentY = drawLaborColumns(pdf, pageContainer, currentY, dayIndex);

            const signatureSectionTargetY = pageHeight - margin - signatureSectionHeight;
            const availableHeightForFullActivitySection = signatureSectionTargetY - currentY;

            currentY = drawActivityTable(pdf, pageContainer, currentY, dayIndex, 0, availableHeightForFullActivitySection);
            currentY = signatureSectionTargetY;

        } else if (pageIdx === 1) {
             currentY = drawActivityTable(pdf, pageContainer, currentY, dayIndex, 1);

             const signatureSectionTargetY = pageHeight - margin - signatureSectionHeight;
             const availableHeightForObservationsSection = signatureSectionTargetY - currentY;
             currentY = drawObservationsSection(pdf, pageContainer, currentY, dayIndex, availableHeightForObservationsSection);
             currentY = signatureSectionTargetY;

        } else if (pageIdx === 2) {
            currentY = await drawReportPhotosSection(pdf, pageContainer, currentY, dayIndex);
            currentY = pageHeight - margin - signatureSectionHeight;
        }

        await drawSignatureSection(pdf, pageContainer, currentY, dayIndex, pageIdx);
        await delay(50); // Small delay per page to allow UI to breathe if needed
    }
}


async function generatePdf() {
    const generatePdfButton = getElementByIdSafe('generatePdfButton');
    if (generatePdfButton) generatePdfButton.disabled = true;

    // --- PDF Progress Bar Elements ---
    const pdfProgressBarWrapperEl = getElementByIdSafe('pdfProgressBarWrapper');
    const pdfProgressDetailsTextEl = getElementByIdSafe('pdfProgressDetailsText');
    const pdfProgressBarEl = getElementByIdSafe('pdfProgressBar');
    const pdfProgressBarContainerEl = getElementByIdSafe('pdfProgressBarContainer');
    const pdfProgressTextEl = getElementByIdSafe('pdfProgressText');
    // --- End PDF Progress Bar Elements ---

    const totalDaysToProcess = rdoDayCounter + 1;
    let generationError = null;

    if (pdfProgressBarWrapperEl && pdfProgressDetailsTextEl && pdfProgressBarEl && pdfProgressBarContainerEl && pdfProgressTextEl) {
        pdfProgressDetailsTextEl.textContent = 'Iniciando geração do PDF...';
        pdfProgressBarEl.style.width = '0%';
        pdfProgressTextEl.textContent = `0/${totalDaysToProcess}`;
        pdfProgressBarContainerEl.setAttribute('aria-valuemax', totalDaysToProcess.toString());
        pdfProgressBarContainerEl.setAttribute('aria-valuenow', '0');
        pdfProgressBarWrapperEl.style.display = 'block';
    } else {
        // Fallback to old message if new elements are not found
        updateSaveProgressMessage("Gerando PDF... Por favor, aguarde.", false);
    }


    const pdf = new jsPDF('p', 'mm', 'a4');
    pdf.setProperties({ title: 'Relatório Diário de Obras - RDO' });

    try {
        for (let d = 0; d <= rdoDayCounter; d++) {
            const currentDayNumber = d + 1;
            if (pdfProgressBarWrapperEl && pdfProgressDetailsTextEl && pdfProgressBarEl && pdfProgressBarContainerEl && pdfProgressTextEl) {
                pdfProgressDetailsTextEl.textContent = `Processando Dia ${currentDayNumber} de ${totalDaysToProcess}...`;
                const progressPercentage = (currentDayNumber / totalDaysToProcess) * 100;
                pdfProgressBarEl.style.width = `${progressPercentage}%`;
                pdfProgressTextEl.textContent = `${currentDayNumber}/${totalDaysToProcess}`;
                pdfProgressBarContainerEl.setAttribute('aria-valuenow', currentDayNumber.toString());
            }

            let dayWrapper;
            const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
            if (dayWithButtonLayoutWrapper) {
                dayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
            } else {
                dayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
            }

            if (dayWrapper) {
                await generateSingleDayPdf(pdf, d, dayWrapper, d === 0);
            } else {
                console.warn(`generatePdf: Wrapper for day ${d} not found. Skipping this day.`);
            }
            if (d < rdoDayCounter) {
                await delay(100); // Give a bit of breathing room between days
            }
        }

        const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
        let filename = "RDO_Completo.pdf";
        if (day0Page1Container) {
            const rdoNum = getInputValue('numero_rdo_day0', day0Page1Container);
            const dataDay0El = getElementByIdSafe('data_day0', day0Page1Container);
            const dataVal = dataDay0El ? dataDay0El.value : '';

            if (rdoNum && dataVal && /^\d{4}-\d{2}-\d{2}$/.test(dataVal)) {
                filename = `RDO_${rdoNum.replace('-A', '')}_${dataVal}.pdf`;
            } else if (rdoNum) {
                 filename = `RDO_${rdoNum.replace('-A', '')}.pdf`;
            } else if (dataVal && /^\d{4}-\d{2}-\d{2}$/.test(dataVal)) {
                 filename = `RDO_Ref_${dataVal}.pdf`;
            }
        }

        pdf.save(filename);
        if (pdfProgressDetailsTextEl) {
            pdfProgressDetailsTextEl.textContent = 'PDF gerado com sucesso!';
             if (pdfProgressBarEl) pdfProgressBarEl.style.width = '100%';
        }
        updateSaveProgressMessage("PDF gerado com sucesso!", false); // Keep general success message


    } catch (error) {
        console.error("Error generating PDF:", error);
        generationError = error;
        if (pdfProgressDetailsTextEl) {
            pdfProgressDetailsTextEl.textContent = 'Erro ao gerar PDF.';
        }
        updateSaveProgressMessage("Erro ao gerar PDF. Verifique o console.", true); // Keep general error message
    } finally {
        if (generatePdfButton) generatePdfButton.disabled = false;

        if (pdfProgressBarWrapperEl) {
            setTimeout(() => {
                pdfProgressBarWrapperEl.style.display = 'none';
                // Reset for next time (optional, could be done on show)
                if(pdfProgressDetailsTextEl) pdfProgressDetailsTextEl.textContent = 'Gerando PDF...';
                if(pdfProgressBarEl) pdfProgressBarEl.style.width = '0%';
                if(pdfProgressTextEl) pdfProgressTextEl.textContent = '0/0';
                if(pdfProgressBarContainerEl) {
                    pdfProgressBarContainerEl.setAttribute('aria-valuenow', '0');
                }
            }, generationError ? 5000 : 3000); // Longer display for error
        }
    }
}

// --- Save/Load Progress (Gist Integration) ---
const CURRENT_SAVE_VERSION = 3;


function collectDayData(dayIndex, dayWrapperElement) {
    const dayData = {
        reportPhotos_pt2: []
    };
    const inputs = dayWrapperElement.querySelectorAll(
        'input[type="text"], input[type="date"], input[type="month"], input[type="number"], textarea, select, input[type="hidden"]'
    );
    inputs.forEach(input => {
        if (input.id) {
            dayData[input.id] = (input).value;
        }
    });

    const page3Container = dayWrapperElement.querySelectorAll('.rdo-container')[2];
    if (page3Container) {
        for (let i = 1; i <= 4; i++) {
            const preview = getElementByIdSafe(`report_photo_${i}_pt2_preview_day${dayIndex}`, page3Container);
            const captionInput = getElementByIdSafe(`report_photo_${i}_pt2_caption_day${dayIndex}`, page3Container);
            let imageUrl = null;
            if (preview && preview.style.display !== 'none' && preview.src && preview.src.startsWith('http') && preview.dataset.isEffectivelyLoaded === 'true') {
                imageUrl = preview.src;
            }
            dayData.reportPhotos_pt2.push({
                url: imageUrl,
                caption: captionInput ? captionInput.value : ''
            });
        }
    }
    return dayData;
}

function serializeRdoDataForGist() {
    let masterSignatureImgBBUrl = null;
    const day0Preview = getElementByIdSafe('consbemPhotoPreview_day0');
    if (day0Preview && day0Preview.src && day0Preview.src.startsWith('http') && day0Preview.dataset.isEffectivelyLoaded === 'true') {
        masterSignatureImgBBUrl = day0Preview.src;
    }


    const formData = {
        version: CURRENT_SAVE_VERSION,
        rdoDayCounter: rdoDayCounter,
        days: [],
        masterSignatureImgBBUrl: masterSignatureImgBBUrl,
        hasEfetivoBeenLoadedAtLeastOnce: hasEfetivoBeenLoadedAtLeastOnce,
        efetivoWasJustClearedByButton: efetivoWasJustClearedByButton
    };

    for (let d = 0; d <= rdoDayCounter; d++) {
        let dayWrapper = null;
        const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
        if (dayWithButtonLayoutWrapper) {
            dayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
        } else {
            dayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
        }

        if (dayWrapper) {
            formData.days.push(collectDayData(d, dayWrapper));
        } else {
            console.warn(`Serialize RDO Data: Day wrapper for day index ${d} not found. Data for this day will not be included.`);
             if (d === 0) {
                const d0El = document.getElementById('rdo_day_0_wrapper');
                const d0ContainerEl = document.getElementById('day_with_button_container_day0');
                console.log(`Debug Day 0 Serialization: rdo_day_0_wrapper exists? ${!!d0El}. day_with_button_container_day0 exists? ${!!d0ContainerEl}`);
                if (d0ContainerEl) {
                    console.log(`Debug Day 0 Serialization: rdo_day_0_wrapper inside container? ${!!d0ContainerEl.querySelector('.rdo-day-wrapper')}`);
                }
            }
        }
    }
    return formData;
}


async function fetchGistContent() {
    updateSaveProgressMessage("Carregando progresso...", false);
    try {
        const response = await fetch(`https://api.github.com/gists/${GIST_ID}`, {
            method: 'GET',
            headers: {
                'Authorization': `token ${GIST_TOKEN}`,
                'Accept': 'application/vnd.github.v3+json',
            },
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ message: response.statusText }));
            console.error(`Erro ao buscar Gist: ${response.status}`, errorData);
            updateSaveProgressMessage(`Erro ao carregar progresso.`, true);
            return null;
        }

        const gistData = await response.json();
        if (gistData.files && gistData.files[GIST_FILENAME] && gistData.files[GIST_FILENAME].content) {
            updateSaveProgressMessage("Progresso carregado.", false);
            return JSON.parse(gistData.files[GIST_FILENAME].content);
        } else {
            updateSaveProgressMessage("Nenhum progresso salvo encontrado.", false);
            return null;
        }
    } catch (error) {
        console.error("Erro de rede ou ao processar Gist:", error);
        updateSaveProgressMessage("Erro de rede ao carregar progresso.", true);
        return null;
    }
}

async function updateGistContent(rdoDataJsonString) {
    // Message "Salvando progresso..." will be shown by caller (autoSaveToGist or manualSaveProgress)
    try {
        const response = await fetch(`https://api.github.com/gists/${GIST_ID}`, {
            method: 'PATCH',
            headers: {
                'Authorization': `token ${GIST_TOKEN}`,
                'Accept': 'application/vnd.github.v3+json',
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                files: {
                    [GIST_FILENAME]: {
                        content: rdoDataJsonString,
                    },
                },
            }),
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ message: response.statusText }));
            console.error(`Erro ao atualizar Gist: ${response.status}`, errorData);
            updateSaveProgressMessage("Erro ao salvar progresso.", true);
            return false;
        }
        updateSaveProgressMessage("Progresso salvo!", false);
        return true;
    } catch (error) {
        console.error("Erro de rede ao atualizar Gist:", error);
        updateSaveProgressMessage("Erro de rede ao salvar progresso.", true);
        return false;
    }
}


function manualSaveProgress() {
    updateSaveProgressMessage("Salvando progresso...", false);
    const rdoData = serializeRdoDataForGist();
    if (rdoData) {
        const jsonString = JSON.stringify(rdoData, null, 2);
        updateGistContent(jsonString).then(success => {
            if (success) {
                isFormDirty = false;
                updateSaveButtonState();
                 // "Progresso salvo!" message is handled by updateGistContent
            }
        });
    }
}

async function applyGistDataToRdo(dataFromGist) {
    try {
        if (!dataFromGist || typeof dataFromGist.version !== 'number' || dataFromGist.version > CURRENT_SAVE_VERSION) {
            updateSaveProgressMessage("Dados salvos inválidos. Iniciando com padrão.", true);
            const day0Page1ContainerInitial = getElementByIdSafe('rdo_day_0_page_0');
            if (day0Page1ContainerInitial) {
                const day0DateInputInitial = getElementByIdSafe('data_day0', day0Page1ContainerInitial);
                let dateForMonthAdjustmentInitial = day0DateInputInitial ? day0DateInputInitial.value : null;
                 if (day0DateInputInitial) {
                     if (!dateForMonthAdjustmentInitial || !/^\d{4}-\d{2}-\d{2}$/.test(dateForMonthAdjustmentInitial)) {
                        const today = new Date();
                        dateForMonthAdjustmentInitial = formatDateToYYYYMMDD(new Date(today.getFullYear(), today.getMonth(), 1));
                        day0DateInputInitial.value = dateForMonthAdjustmentInitial;
                    }
                    await adjustRdoDaysForMonth(dateForMonthAdjustmentInitial);
                } else {
                    await updateAllDateRelatedFieldsFromDay0();
                }
            }
            return false;
        }

        hasEfetivoBeenLoadedAtLeastOnce = dataFromGist.hasEfetivoBeenLoadedAtLeastOnce || false;
        efetivoWasJustClearedByButton = dataFromGist.efetivoWasJustClearedByButton || false;

        const allDaysWrapper = getElementByIdSafe('rdo-all-days-wrapper');
        if (allDaysWrapper) {
            for (let d = rdoDayCounter; d > 0; d--) {
                const dayWithButtonContainerToRemove = document.getElementById(`day_with_button_container_day${d}`);
                const rdoDayWrapperItself = getElementByIdSafe(`rdo_day_${d}_wrapper`);
                let elementToRemoveFromObserver = null;
                let elementToRemoveFromDOM = null;

                if (dayWithButtonContainerToRemove) {
                    elementToRemoveFromObserver = dayWithButtonContainerToRemove;
                    elementToRemoveFromDOM = dayWithButtonContainerToRemove;
                } else if (rdoDayWrapperItself) {
                    elementToRemoveFromObserver = rdoDayWrapperItself;
                    elementToRemoveFromDOM = rdoDayWrapperItself;
                }

                if (elementToRemoveFromObserver && navObserver) navObserver.unobserve(elementToRemoveFromObserver);
                if (elementToRemoveFromDOM) {
                    originalRdoSheetHeights.delete(elementToRemoveFromDOM.id);
                    elementToRemoveFromDOM.remove();
                }

                const tabToRemove = getElementByIdSafe(`day-tab-${d}`);
                if (tabToRemove) tabToRemove.remove();
            }
        }
        dayTabs.length = Math.min(dayTabs.length, 1);
        rdoDayCounter = 0;
        if (dayTabsContainer) {
            const firstTab = dayTabsContainer.querySelector('#day-tab-0');
            dayTabsContainer.innerHTML = '';
            if (firstTab) dayTabsContainer.appendChild(firstTab);
        }
        if (dayTabs.length > 0 && dayTabs[0]) setActiveTab(0);


        let day0Wrapper = getElementByIdSafe('rdo_day_0_wrapper');
        if (!day0Wrapper) {
             const day0InSundayContainer = document.getElementById('day_with_button_container_day0');
             if (day0InSundayContainer) {
                 day0Wrapper = day0InSundayContainer.querySelector('.rdo-day-wrapper');
                 if (day0Wrapper && allDaysWrapper && day0InSundayContainer.parentElement !== allDaysWrapper) {
                    allDaysWrapper.appendChild(day0InSundayContainer);
                 }
             }
        }
         if (!day0Wrapper && allDaysWrapper) {
            console.error("applyGistDataToRdo: Day 0 wrapper still not found after Sunday check. Attempting re-creation from pristine state.");
             day0Wrapper = document.createElement('div');
             day0Wrapper.className = 'rdo-day-wrapper';
             day0Wrapper.innerHTML = pristineRdoDay0WrapperHtml;
             day0Wrapper.id = 'rdo_day_0_wrapper';
             updateIdsRecursive(day0Wrapper, '_day0', `_day0`);
             day0Wrapper.dataset.dayIndex = "0";
             allDaysWrapper.insertBefore(day0Wrapper, allDaysWrapper.firstChild);
        }

        if (!day0Wrapper) {
            updateSaveProgressMessage("Erro crítico: RDO Dia 0 não encontrado. Não é possível carregar.", true);
            return false;
        }


        resetDayInstanceContent(day0Wrapper, 0);
        await initializeDayInstance(0, day0Wrapper);
        const finalDay0Element = updateEfetivoCopyButtonVisibility(0, day0Wrapper);
        if (finalDay0Element !== day0Wrapper && allDaysWrapper) {
            if (day0Wrapper.parentElement && day0Wrapper.parentElement !== allDaysWrapper) {
                 day0Wrapper.parentElement.replaceWith(finalDay0Element);
            } else if (!day0Wrapper.parentElement && finalDay0Element.parentElement !== allDaysWrapper) {
                 allDaysWrapper.insertBefore(finalDay0Element, allDaysWrapper.firstChild);
            }
        } else if (finalDay0Element === day0Wrapper && day0Wrapper.parentElement !== allDaysWrapper && allDaysWrapper) {
            allDaysWrapper.insertBefore(day0Wrapper, allDaysWrapper.firstChild);
        }

        let signatureUrlFromGist = dataFromGist.masterSignatureImgBBUrl || dataFromGist.masterSignatureBase64;
        if (signatureUrlFromGist && signatureUrlFromGist.startsWith('data:image')) {
            try {
                const signatureFile = await dataUrlToFile(signatureUrlFromGist, 'signature_day0.png');
                const imgbbUrl = await uploadImageToImgBB(signatureFile, "Assinatura CONSBEM (migrada)");
                if (imgbbUrl) {
                    signatureUrlFromGist = imgbbUrl;
                    markFormAsDirty();
                } else {
                    signatureUrlFromGist = null;
                    updateSaveProgressMessage("Erro ao processar assinatura salva.", true);
                }
            } catch (e) {
                signatureUrlFromGist = null;
                updateSaveProgressMessage("Erro ao processar assinatura salva.", true);
                console.error("Error migrating signature from base64:", e);
            }
        }
        updateAllSignaturePreviews(signatureUrlFromGist);


        const targetRdoDayCountFromGist = (dataFromGist.rdoDayCounter !== undefined ? dataFromGist.rdoDayCounter : (dataFromGist.days ? dataFromGist.days.length - 1 : 0)) + 1;

        if (dataFromGist.days && dataFromGist.days.length > 0) {
            const day0Data = dataFromGist.days[0];
            if (day0Data) {
                Object.keys(day0Data).forEach(key => {
                    const value = day0Data[key];
                    const element = day0Wrapper.querySelector(`#${key}`);
                    if (key === 'data_day0') {
                        const dataDay0El = day0Wrapper.querySelector('#data_day0');
                        if (dataDay0El && dataDay0El.type === 'date') {
                            if (value && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
                                dataDay0El.value = value;
                            } else if (value && /^\d{4}-\d{2}$/.test(value)) {
                                dataDay0El.value = value + "-01";
                            } else {
                                const today = new Date();
                                dataDay0El.value = formatDateToYYYYMMDD(new Date(today.getFullYear(), today.getMonth(), 1));
                            }
                        } else if (element) {
                             setInputValue(key, value, day0Wrapper);
                        }
                    } else if (element) {
                        if (element.type === 'file') { /* Skip */ }
                        else if (key.startsWith('tempo_') || key.startsWith('trabalho_')) { /* Skip */ }
                        else if (key.includes('_naturalWidth') || key.includes('_naturalHeight')) { /* Skip */ }
                        else if (key.startsWith('report_photo_') && key.includes('_pt2_input_day0')) { /* Skip */ }
                        else if (typeof value === 'string') {
                             setInputValue(key, value, day0Wrapper);
                        }
                    }
                });
                ['t1', 't2', 't3'].forEach(turnoIdBase => {
                    const tempoValue = day0Data[`tempo_${turnoIdBase}_day0_hidden`];
                    const trabalhoValue = day0Data[`trabalho_${turnoIdBase}_day0_hidden`];
                    const cellId = `${turnoIdBase === 't1' ? 'turno1' : turnoIdBase === 't2' ? 'turno2' : 'turno3'}_control_cell_day0`;
                    const updater = shiftControlUpdaters.get(cellId);
                    if (updater && typeof tempoValue === 'string' && typeof trabalhoValue === 'string') {
                        updater.update(tempoValue, trabalhoValue);
                    }
                });
                const dropdownHiddenInputsDay0 = day0Wrapper.querySelectorAll(
                    'input[type="hidden"][id^="status_day0_"], input[type="hidden"][id^="localizacao_day0_"], input[type="hidden"][id^="tipo_servico_day0_"], input[type="hidden"][id^="servico_desc_day0_"]'
                );
                dropdownHiddenInputsDay0.forEach(hiddenInput => {
                     if (day0Data[hiddenInput.id] !== undefined) {
                        hiddenInput.value = day0Data[hiddenInput.id];
                    }
                    const displayId = hiddenInput.id.replace('_hidden', '_display');
                    const displayElement = getElementByIdSafe(displayId, day0Wrapper);
                    if (displayElement) {
                        const currentValue = hiddenInput.value;
                        const optionText = statusOptions.includes(currentValue) ? currentValue : localizacaoOptions.includes(currentValue) ? currentValue : tipoServicoOptions.includes(currentValue) ? currentValue : servicoOptions.includes(currentValue) ? currentValue : "";
                        displayElement.textContent = optionText === "" ? "" : optionText;
                        const dropdownId = hiddenInput.id.replace('_hidden', '_dropdown');
                        const dropdownElement = getElementByIdSafe(dropdownId, day0Wrapper);
                        if (dropdownElement) {
                            dropdownElement.querySelectorAll('.status-option').forEach(opt => {
                                opt.setAttribute('aria-selected', (opt.dataset.value === currentValue) ? 'true' : 'false');
                            });
                        }
                    }
                });
                const page3ContainerDay0 = day0Wrapper.querySelectorAll('.rdo-container')[2];
                 if (page3ContainerDay0 && day0Data.reportPhotos_pt2 && Array.isArray(day0Data.reportPhotos_pt2)) {
                    day0Data.reportPhotos_pt2.forEach(async (photoData, i) => {
                        if (photoData && photoData.url && photoData.url.startsWith('http')) {
                            const photoInput = getElementByIdSafe(`report_photo_${i+1}_pt2_input_day0`, page3ContainerDay0);
                            const preview = getElementByIdSafe(`report_photo_${i+1}_pt2_preview_day0`, page3ContainerDay0);
                            const placeholder = getElementByIdSafe(`report_photo_${i+1}_pt2_placeholder_day0`, page3ContainerDay0);
                            const clearButton = getElementByIdSafe(`report_photo_${i+1}_pt2_clear_button_day0`, page3ContainerDay0);
                            const label = getElementByIdSafe(`report_photo_${i+1}_pt2_label_day0`, page3ContainerDay0);
                            setReportPhotoSlotImage(photoInput, preview, placeholder, clearButton, label, photoData.url);
                        }
                        const captionInput = getElementByIdSafe(`report_photo_${i+1}_pt2_caption_day0`, page3ContainerDay0);
                        if (captionInput && photoData) captionInput.value = photoData.caption || '';
                    });
                }
            }
        }


        const fragment = document.createDocumentFragment();
        for (let d = 1; d < targetRdoDayCountFromGist; d++) {
            rdoDayCounter = d;
            const newDayWrapper = await addNewRdoDay(d);
            if (newDayWrapper) {
                 await initializeDayInstance(d, newDayWrapper);

                 if (dataFromGist.days && dataFromGist.days[d]) {
                    const dayData = dataFromGist.days[d];
                     Object.keys(dayData).forEach(key => {
                        const value = dayData[key];
                         const element = newDayWrapper.querySelector(`#${key}`);
                         if (element) {
                            if (element.type === 'file') { /* Skip */ }
                            else if (key.startsWith('tempo_') || key.startsWith('trabalho_')) { /* Skip */ }
                            else if (key.includes('_naturalWidth') || key.includes('_naturalHeight')) { /* Skip */ }
                            else if (key.startsWith('report_photo_') && key.includes(`_pt2_input_day${d}`)) { /* Skip */ }
                            else if (typeof value === 'string') {
                                 setInputValue(key, value, newDayWrapper);
                            }
                        }
                    });
                     ['t1', 't2', 't3'].forEach(turnoIdBase => {
                        const tempoValue = dayData[`tempo_${turnoIdBase}_day${d}_hidden`];
                        const trabalhoValue = dayData[`trabalho_${turnoIdBase}_day${d}_hidden`];
                        const cellId = `${turnoIdBase === 't1' ? 'turno1' : turnoIdBase === 't2' ? 'turno2' : 'turno3'}_control_cell_day${d}`;
                        const updater = shiftControlUpdaters.get(cellId);
                        if (updater && typeof tempoValue === 'string' && typeof trabalhoValue === 'string') {
                            updater.update(tempoValue, trabalhoValue);
                        }
                    });
                    const dropdownHiddenInputs = newDayWrapper.querySelectorAll(
                         `input[type="hidden"][id^="status_day${d}_"], input[type="hidden"][id^="localizacao_day${d}_"], input[type="hidden"][id^="tipo_servico_day${d}_"], input[type="hidden"][id^="servico_desc_day${d}_"]`
                    );
                    dropdownHiddenInputs.forEach(hiddenInput => {
                         if (dayData[hiddenInput.id] !== undefined) {
                            hiddenInput.value = dayData[hiddenInput.id];
                        }
                        const displayId = hiddenInput.id.replace('_hidden', '_display');
                        const displayElement = getElementByIdSafe(displayId, newDayWrapper);
                        if (displayElement) {
                            const currentValue = hiddenInput.value;
                            const optionText = statusOptions.includes(currentValue) ? currentValue : localizacaoOptions.includes(currentValue) ? currentValue : tipoServicoOptions.includes(currentValue) ? currentValue : servicoOptions.includes(currentValue) ? currentValue : "";
                            displayElement.textContent = optionText === "" ? "" : optionText;
                            const dropdownId = hiddenInput.id.replace('_hidden', '_dropdown');
                            const dropdownElement = getElementByIdSafe(dropdownId, newDayWrapper);
                            if (dropdownElement) {
                                dropdownElement.querySelectorAll('.status-option').forEach(opt => {
                                    opt.setAttribute('aria-selected', (opt.dataset.value === currentValue) ? 'true' : 'false');
                                });
                            }
                        }
                    });
                    const page3ContainerDayD = newDayWrapper.querySelectorAll('.rdo-container')[2];
                    if (page3ContainerDayD && dayData.reportPhotos_pt2 && Array.isArray(dayData.reportPhotos_pt2)) {
                        dayData.reportPhotos_pt2.forEach(async (photoData, i) => {
                            if (photoData && photoData.url && photoData.url.startsWith('http')) {
                                const photoInput = getElementByIdSafe(`report_photo_${i+1}_pt2_input_day${d}`, page3ContainerDayD);
                                const preview = getElementByIdSafe(`report_photo_${i+1}_pt2_preview_day${d}`, page3ContainerDayD);
                                const placeholder = getElementByIdSafe(`report_photo_${i+1}_pt2_placeholder_day${d}`, page3ContainerDayD);
                                const clearButton = getElementByIdSafe(`report_photo_${i+1}_pt2_clear_button_day${d}`, page3ContainerDayD);
                                const label = getElementByIdSafe(`report_photo_${i+1}_pt2_label_day${d}`, page3ContainerDayD);
                                setReportPhotoSlotImage(photoInput, preview, placeholder, clearButton, label, photoData.url);
                            }
                             const captionInput = getElementByIdSafe(`report_photo_${i+1}_pt2_caption_day${d}`, page3ContainerDayD);
                            if (captionInput && photoData) captionInput.value = photoData.caption || '';
                        });
                    }
                }
                const finalElementForFragment = updateEfetivoCopyButtonVisibility(d, newDayWrapper);
                fragment.appendChild(finalElementForFragment);
                createDayTab(d);
            }
        }
        rdoDayCounter = Math.max(0, targetRdoDayCountFromGist -1);

        if (allDaysWrapper && fragment.childNodes.length > 0) {
            allDaysWrapper.appendChild(fragment);
        }

        await updateAllDateRelatedFieldsFromDay0();

        for (let d = 0; d < targetRdoDayCountFromGist; d++) {
            let dayWrapperForTotals = null;
            const dayWithButtonWrapper = document.getElementById(`day_with_button_container_day${d}`);
            if (dayWithButtonWrapper) {
                dayWrapperForTotals = dayWithButtonWrapper.querySelector('.rdo-day-wrapper');
            } else {
                dayWrapperForTotals = getElementByIdSafe(`rdo_day_${d}_wrapper`);
            }
            if (dayWrapperForTotals) {
                 const page1Container = dayWrapperForTotals.querySelectorAll('.rdo-container')[0];
                 if(page1Container) calculateSectionTotalsForDay(page1Container, d);
            }
        }

        updateAllSignaturePreviews(signatureUrlFromGist);
        setupNavigationObserver();
        applyResponsiveScaling(); // Apply scaling after Gist data is applied

        updateAllEfetivoAlerts();
        updateClearEfetivoButtonVisibility();
        isFormDirty = false;
        updateSaveButtonState();
        updateSaveProgressMessage("Progresso aplicado!", false);
        return true;

    } catch (error) {
        console.error("Error applying Gist data:", error);
        updateSaveProgressMessage("Erro ao aplicar progresso salvo.", true);
        return false;
    }
}

// --- Responsive Scaling Implementation ---
function applyResponsiveScaling() {
    const appContentCanvas = getElementByIdSafe('app-content-canvas');
    if (!appContentCanvas) return;

    const canvasWidth = appContentCanvas.offsetWidth;
    let scaleFactor = canvasWidth / RDO_ORIGINAL_DESIGN_WIDTH;
    if (scaleFactor >= 1) {
        scaleFactor = 1; // Do not scale up
    }

    const rdoDayWrappers = document.querySelectorAll('.rdo-day-wrapper');
    rdoDayWrappers.forEach(sheet => {
        let originalHeight = originalRdoSheetHeights.get(sheet.id);
        if (originalHeight === undefined) { // If not cached, get it and cache it
            originalHeight = sheet.offsetHeight;
            originalRdoSheetHeights.set(sheet.id, originalHeight);
        }

        const scaledHeight = originalHeight * scaleFactor;
        sheet.style.transform = `scale(${scaleFactor})`;
        sheet.style.height = `${scaledHeight}px`;

        // Adjust parent .day-with-button-container if it exists
        const parentContainer = sheet.closest('.day-with-button-container');
        if (parentContainer) {
            parentContainer.style.height = `${scaledHeight}px`;
            const sidebar = parentContainer.querySelector('.efetivo-actions-sidebar');
            if (sidebar) {
                // Calculate the horizontal offset needed to keep the sidebar next to the scaled sheet
                // The sheet scales from its center, so its perceived right edge moves.
                // Original right edge is at RDO_ORIGINAL_DESIGN_WIDTH / 2 from its center.
                // Scaled right edge is at (RDO_ORIGINAL_DESIGN_WIDTH * scaleFactor) / 2 from its center.
                // The difference is how much the right edge "moved" inwards relative to the unscaled position.
                const unscaledHalfWidth = RDO_ORIGINAL_DESIGN_WIDTH / 2;
                const scaledHalfWidth = (RDO_ORIGINAL_DESIGN_WIDTH * scaleFactor) / 2;
                const edgeShift = unscaledHalfWidth - scaledHalfWidth;

                // The sidebar's `left: 100%` is relative to the *unscaled* width of the sheet's original position (due to transform-origin).
                // We want to move it left by the amount the edge shifted, minus its own margin.
                const translateX = -edgeShift;
                sidebar.style.transform = `translateX(${translateX}px)`;
            }
        }
    });
}

function debounce(func, delay) {
    let timeout;
    return function(...args) {
        const context = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), delay);
    };
}
const debouncedApplyResponsiveScaling = debounce(applyResponsiveScaling, 150);


// --- Event Listeners and Initial Setup ---
document.addEventListener('DOMContentLoaded', async () => {
    showLoadingSpinner(); // Show spinner before any loading

    rainfallProgressBarWrapper = getElementByIdSafe('rainfallProgressBarWrapper');
    rainfallProgressBar = getElementByIdSafe('rainfallProgressBar');
    rainfallProgressText = getElementByIdSafe('rainfallProgressText');
    rainfallProgressDetailsText = getElementByIdSafe('rainfallProgressDetailsText');

    confirmationModalOverlay = getElementByIdSafe('confirmation-modal-overlay');
    confirmClearRdoButton = getElementByIdSafe('confirmClearRdoButton');
    cancelClearRdoButton = getElementByIdSafe('cancelClearRdoButton');


    const initialDay0Wrapper = getElementByIdSafe('rdo_day_0_wrapper');
    if (initialDay0Wrapper) {
        pristineRdoDay0WrapperHtml = initialDay0Wrapper.innerHTML;
        // Store initial height for day 0
        originalRdoSheetHeights.set(initialDay0Wrapper.id, initialDay0Wrapper.offsetHeight);
        createDayTab(0);
        setActiveTab(0);

    } else {
        console.error("CRITICAL: Initial rdo_day_0_wrapper not found in HTML. Application cannot start correctly.");
        updateSaveProgressMessage("Erro crítico na inicialização da página. Verifique o console.", true);
        hideLoadingSpinner(); // Hide spinner on critical error
        return;
    }


    const generatePdfButton = getElementByIdSafe('generatePdfButton');
    if (generatePdfButton) generatePdfButton.addEventListener('click', generatePdf);

    const fetchRainfallButton = getElementByIdSafe('fetchRainfallButton');
    if (fetchRainfallButton) fetchRainfallButton.addEventListener('click', fetchAllRainfallDataForAllDaysOnClick);

    saveProgressButtonElement = getElementByIdSafe('saveProgressButton');
    if (saveProgressButtonElement) {
        saveProgressButtonElement.addEventListener('click', manualSaveProgress);
    }

    const efetivoUploadInput = getElementByIdSafe('efetivoSpreadsheetUpload');
    if (efetivoUploadInput) efetivoUploadInput.addEventListener('change', handleEfetivoFileUpload);

    const loadEfetivoButton = getElementByIdSafe('loadEfetivoSpreadsheetButton');
    if (loadEfetivoButton) loadEfetivoButton.addEventListener('click', triggerEfetivoFileUpload);

    const clearEfetivoBtn = getElementByIdSafe('clearEfetivoButton');
    if (clearEfetivoBtn) clearEfetivoBtn.addEventListener('click', clearAllEfetivo);

    const clearRdoBtn = getElementByIdSafe('clearRdoButton');
    if (clearRdoBtn) {
        clearRdoBtn.addEventListener('click', () => {
            showConfirmationModal();
        });
    }

    if (confirmClearRdoButton) {
        confirmClearRdoButton.addEventListener('click', async () => {
            hideConfirmationModal();
            showLoadingSpinner();
            updateSaveProgressMessage("Limpando RDO...", false); // Show processing message

            try {
                let anyFieldChanged = false;

                for (let d = 0; d <= rdoDayCounter; d++) {
                    let dayWrapper = null;
                    const dayWithButtonLayoutWrapper = document.getElementById(`day_with_button_container_day${d}`);
                    if (dayWithButtonLayoutWrapper) {
                        dayWrapper = dayWithButtonLayoutWrapper.querySelector('.rdo-day-wrapper');
                    } else {
                        dayWrapper = getElementByIdSafe(`rdo_day_${d}_wrapper`);
                    }

                    if (dayWrapper) {
                        const page1Container = dayWrapper.querySelector(`#rdo_day_${d}_page_0`);
                        if (page1Container) {
                            const indicePluvInput = getElementByIdSafe(`indice_pluv_valor_day${d}`, page1Container);
                            if (indicePluvInput && indicePluvInput.value !== '') {
                                indicePluvInput.value = '';
                                anyFieldChanged = true;
                            }

                            const currentRainfallStatusEl = getElementByIdSafe(`indice_pluv_status_day${d}`, page1Container);
                            if (currentRainfallStatusEl && (currentRainfallStatusEl.textContent !== '-' || currentRainfallStatusEl.classList.length > 1) ) {
                                 anyFieldChanged = true;
                            }
                            updateRainfallStatus(d, 'idle');

                            const shiftsToUpdate = [
                                { cellIdBase: 'turno1_control_cell', turnoIdBase: 't1', targetTempo: "B", targetTrabalho: "N" },
                                { cellIdBase: 'turno2_control_cell', turnoIdBase: 't2', targetTempo: "B", targetTrabalho: "N" },
                                { cellIdBase: 'turno3_control_cell', turnoIdBase: 't3', targetTempo: "", targetTrabalho: "" }
                            ];

                            shiftsToUpdate.forEach(shift => {
                                const tempoHiddenInput = getElementByIdSafe(`tempo_${shift.turnoIdBase}_day${d}_hidden`, page1Container);
                                const trabalhoHiddenInput = getElementByIdSafe(`trabalho_${shift.turnoIdBase}_day${d}_hidden`, page1Container);
                                let shiftSpecificChange = false;

                                if (tempoHiddenInput && tempoHiddenInput.value !== shift.targetTempo) {
                                    shiftSpecificChange = true;
                                }
                                if (trabalhoHiddenInput && trabalhoHiddenInput.value !== shift.targetTrabalho) {
                                    shiftSpecificChange = true;
                                }

                                const cellId = `${shift.cellIdBase}_day${d}`;
                                const updater = shiftControlUpdaters.get(cellId);
                                if (updater) {
                                    updater.update(shift.targetTempo, shift.targetTrabalho);
                                    if (shiftSpecificChange) {
                                        anyFieldChanged = true;
                                    }
                                } else {
                                     console.warn(`Updater for ${cellId} not found during RDO clear.`);
                                }
                            });

                            const laborTables = page1Container.querySelectorAll('.labor-table');
                            laborTables.forEach(table => {
                                table.querySelectorAll('.quantity-input').forEach(input => {
                                    if (input.value !== '') {
                                        input.value = '';
                                        anyFieldChanged = true;
                                    }
                                });
                            });
                            calculateSectionTotalsForDay(page1Container, d);

                            const obsClimaTextarea = getElementByIdSafe(`obs_clima_day${d}`, page1Container);
                            if (obsClimaTextarea && obsClimaTextarea.value !== '') {
                                obsClimaTextarea.value = '';
                                anyFieldChanged = true;
                            }

                            const activitiesTableP1 = page1Container.querySelector(`#activitiesTable_day${d}`);
                            if (activitiesTableP1) {
                                const rows = activitiesTableP1.querySelectorAll('tbody tr');
                                rows.forEach(row => {
                                    const dropdownCells = row.querySelectorAll('.status-cell');
                                    dropdownCells.forEach(cell => {
                                        const display = cell.querySelector('.status-display');
                                        const hiddenInput = cell.querySelector('input[type="hidden"]');
                                        if (display && display.textContent !== '') {
                                            display.textContent = '';
                                            anyFieldChanged = true;
                                        }
                                        if (hiddenInput && hiddenInput.value !== '') {
                                            hiddenInput.value = '';
                                        }
                                    });
                                    const obsTextareaActivity = row.querySelector('textarea[id^="obs_atividade_"]');
                                    if (obsTextareaActivity && obsTextareaActivity.value !== '') {
                                        obsTextareaActivity.value = '';
                                        anyFieldChanged = true;
                                    }
                                });
                            }
                        }

                        const page2Container = dayWrapper.querySelector(`#rdo_day_${d}_page_1`);
                        if (page2Container) {
                            const activitiesTableP2 = page2Container.querySelector(`#activitiesTable_pt_day${d}`);
                            if (activitiesTableP2) {
                                const rows = activitiesTableP2.querySelectorAll('tbody tr');
                                rows.forEach(row => {
                                    const dropdownCells = row.querySelectorAll('.status-cell');
                                    dropdownCells.forEach(cell => {
                                        const display = cell.querySelector('.status-display');
                                        const hiddenInput = cell.querySelector('input[type="hidden"]');
                                        if (display && display.textContent !== '') {
                                            display.textContent = '';
                                            anyFieldChanged = true;
                                        }
                                        if (hiddenInput && hiddenInput.value !== '') {
                                            hiddenInput.value = '';
                                        }
                                    });
                                    const obsTextareaActivity = row.querySelector('textarea[id^="obs_atividade_"]');
                                    if (obsTextareaActivity && obsTextareaActivity.value !== '') {
                                        obsTextareaActivity.value = '';
                                        anyFieldChanged = true;
                                    }
                                });
                            }
                            const obsComentariosTextarea = getElementByIdSafe(`observacoes_comentarios_pt_day${d}`, page2Container);
                            if (obsComentariosTextarea && obsComentariosTextarea.value !== '') {
                                obsComentariosTextarea.value = '';
                                anyFieldChanged = true;
                            }
                        }

                        const page3Container = dayWrapper.querySelector(`#rdo_day_${d}_page_2`);
                         if (page3Container) {
                            for (let i = 1; i <= 4; i++) {
                                const daySuffixForPhoto = `_day${d}`;
                                const photoInput = getElementByIdSafe(`report_photo_${i}_pt2_input${daySuffixForPhoto}`, page3Container);
                                const photoPreview = getElementByIdSafe(`report_photo_${i}_pt2_preview${daySuffixForPhoto}`, page3Container);
                                const photoPlaceholder = getElementByIdSafe(`report_photo_${i}_pt2_placeholder${daySuffixForPhoto}`, page3Container);
                                const clearButton = getElementByIdSafe(`report_photo_${i}_pt2_clear_button${daySuffixForPhoto}`, page3Container);
                                const labelElement = getElementByIdSafe(`report_photo_${i}_pt2_label${daySuffixForPhoto}`, page3Container);
                                const captionInput = getElementByIdSafe(`report_photo_${i}_pt2_caption${daySuffixForPhoto}`, page3Container);

                                if (photoPreview && photoPreview.style.display !== 'none' && photoPreview.dataset.isEffectivelyLoaded === 'true') {
                                    anyFieldChanged = true;
                                }
                                 if (captionInput && captionInput.value !== '') {
                                    captionInput.value = '';
                                    anyFieldChanged = true;
                                }
                                resetReportPhotoSlotState(photoInput, photoPreview, photoPlaceholder, clearButton, labelElement);
                            }
                        }
                    }
                }

                const masterPhotoInputDay0 = getElementByIdSafe('consbemPhoto_day0');
                const masterPhotoPreviewDay0 = getElementByIdSafe('consbemPhotoPreview_day0');

                if (masterPhotoInputDay0 && masterPhotoPreviewDay0) {
                     const wasSignaturePresent = masterPhotoPreviewDay0.style.display !== 'none' &&
                                               masterPhotoPreviewDay0.dataset.isEffectivelyLoaded === 'true' &&
                                               masterPhotoPreviewDay0.src.startsWith('http');
                    if (wasSignaturePresent) {
                        masterPhotoInputDay0.value = '';
                        updateAllSignaturePreviews(null);
                        anyFieldChanged = true;
                    }
                }

                if (anyFieldChanged) {
                    markFormAsDirty();
                    updateSaveProgressMessage("RDO limpo.", false);
                } else {
                    updateSaveProgressMessage("Nenhum campo para limpar.", false);
                }
                updateClearEfetivoButtonVisibility(); // Crucially update after RDO clear
                updateAllEfetivoAlerts();
            } finally {
                hideLoadingSpinner(); // Ensure spinner is hidden after processing
            }
        });
    }


    if (cancelClearRdoButton) {
        cancelClearRdoButton.addEventListener('click', () => {
            hideConfirmationModal();
            updateSaveProgressMessage("Limpeza cancelada.", false);
        });
    }
    // Handle Esc key for confirmation modal
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            if (confirmationModalOverlay && confirmationModalOverlay.style.display === 'flex') {
                hideConfirmationModal();
                updateSaveProgressMessage("Limpeza cancelada.", false);
            }
        }
    });


    updateSaveButtonState();

    try {
        const loadedData = await fetchGistContent();
        if (loadedData) {
            await applyGistDataToRdo(loadedData);
        } else {
            await initializeDayInstance(0, initialDay0Wrapper);
            const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0', initialDay0Wrapper);
            if (day0Page1Container) {
                const day0DateInputInitial = getElementByIdSafe('data_day0', day0Page1Container);
                let dateForMonthAdjustment = day0DateInputInitial ? day0DateInputInitial.value : null;

                if (day0DateInputInitial) {
                    if (!dateForMonthAdjustment || !/^\d{4}-\d{2}-\d{2}$/.test(dateForMonthAdjustment)) {
                        const today = new Date();
                        dateForMonthAdjustment = formatDateToYYYYMMDD(new Date(today.getFullYear(), today.getMonth(), 1));
                        day0DateInputInitial.value = dateForMonthAdjustment;
                    }
                    const parts = dateForMonthAdjustment.split('-');
                    if (parts[2] !== '01') {
                         dateForMonthAdjustment = `${parts[0]}-${parts[1]}-01`;
                         day0DateInputInitial.value = dateForMonthAdjustment;
                    }
                    await adjustRdoDaysForMonth(dateForMonthAdjustment);
                } else {
                    await updateAllDateRelatedFieldsFromDay0();
                }
            } else {
                 await updateAllDateRelatedFieldsFromDay0();
            }
        }
    } catch (e) {
        console.error("Error during initial Gist load or default setup:", e);
        updateSaveProgressMessage("Erro na inicialização.", true);
        await initializeDayInstance(0, initialDay0Wrapper);
         const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0', initialDay0Wrapper);
         if (day0Page1Container) {
             const day0DateInputInitial = getElementByIdSafe('data_day0', day0Page1Container);
             let dateForMonthAdjustment = day0DateInputInitial ? day0DateInputInitial.value : null;
              if (day0DateInputInitial) {
                 if (!dateForMonthAdjustment || !/^\d{4}-\d{2}-\d{2}$/.test(dateForMonthAdjustment)) {
                    const today = new Date();
                    dateForMonthAdjustment = formatDateToYYYYMMDD(new Date(today.getFullYear(), today.getMonth(), 1));
                    day0DateInputInitial.value = dateForMonthAdjustment;
                }
                const parts = dateForMonthAdjustment.split('-');
                 if (parts[2] !== '01') {
                     dateForMonthAdjustment = `${parts[0]}-${parts[1]}-01`;
                     day0DateInputInitial.value = dateForMonthAdjustment;
                }
                await adjustRdoDaysForMonth(dateForMonthAdjustment);
             } else {
                 await updateAllDateRelatedFieldsFromDay0();
             }
         } else {
            await updateAllDateRelatedFieldsFromDay0();
         }
    } finally {
        hideLoadingSpinner(); // Hide spinner after all initial loading is done or fails
    }

    setupNavigationObserver();
    applyResponsiveScaling(); // Initial call
    window.addEventListener('resize', debouncedApplyResponsiveScaling); // Debounced call on resize

    updateClearEfetivoButtonVisibility();
    updateAllEfetivoAlerts();
    isFormDirty = false;
    updateSaveButtonState();

    document.addEventListener('click', (event) => {
        const openDropdowns = document.querySelectorAll('.status-dropdown[style*="display: block"]');
        const openMonthYearPicker = document.querySelector('.month-year-picker[style*="display: block"]');

        let clickedInsideADropdownOrItsDisplay = false;
        let clickedInsideMonthYearPickerOrIcon = false;

        openDropdowns.forEach(dropdown => {
            const displayId = dropdown.id.replace('_dropdown', '_display');
            const displayElement = document.getElementById(displayId);
            if (dropdown.contains(event.target) || (displayElement && displayElement.contains(event.target))) {
                clickedInsideADropdownOrItsDisplay = true;
            }
        });

        if(openMonthYearPicker) {
            const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
            const monthYearPickerIcon = day0Page1Container?.querySelector('.date-input-icon-wrapper .date-picker-indicator-icon');
            if (openMonthYearPicker.contains(event.target) || (monthYearPickerIcon && monthYearPickerIcon.contains(event.target))) {
                clickedInsideMonthYearPickerOrIcon = true;
            }
        }


        if (!clickedInsideADropdownOrItsDisplay) {
            openDropdowns.forEach(dropdown => {
                dropdown.style.display = 'none';
                const displayId = dropdown.id.replace('_dropdown', '_display');
                const displayElement = document.getElementById(displayId);
                if (displayElement) {
                    displayElement.setAttribute('aria-expanded', 'false');
                }
            });
        }

        if(openMonthYearPicker && !clickedInsideMonthYearPickerOrIcon) {
            openMonthYearPicker.style.display = 'none';
            const day0Page1Container = getElementByIdSafe('rdo_day_0_page_0');
            const monthYearPickerIcon = day0Page1Container?.querySelector('.date-input-icon-wrapper .date-picker-indicator-icon');
            if(monthYearPickerIcon) monthYearPickerIcon.setAttribute('aria-expanded', 'false');
        }
    });

});
