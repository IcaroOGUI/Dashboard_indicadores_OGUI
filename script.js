Chart.register(ChartDataLabels);


const today = new Date();
const dateOptions = { day: '2-digit', month: '2-digit', year: 'numeric' };
document.getElementById('header-date').textContent = today.toLocaleDateString('pt-BR', dateOptions);

let charts = {
    progress: null,
    status: null,
    capitulo: null,
    respondente: null
};

const fileInput = document.getElementById('file-upload');
const fileNameDisplay = document.getElementById('file-name');
const uploadArea = document.getElementById('upload-area');
const dashboardContainer = document.getElementById('dashboard-container');
const errorMessage = document.getElementById('error-message');
const errorText = document.getElementById('error-text');
const btnReset = document.getElementById('btn-reset');

const statusColorMap = {
    'finalizado': '#10B981',
    'validado ógui': '#3B82F6',
    'ajustar': '#30CFC7',
    'respondido': '#F59E0B',
    'sem resposta': '#C7C7C7'
};
const defaultColors = ['#8B5CF6', '#EC4899', '#06B6D4', '#14B8A6', '#F97316'];

function getColorForStatus(statusName, index) {
    const key = statusName.toLowerCase();
    return statusColorMap[key] || defaultColors[index % defaultColors.length];
}

function getTextColorForStatus(statusName) {
    const key = statusName.toLowerCase();
    if (key === 'sem resposta') return '#1E293B';
    return '#ffffff';
}

fileInput.addEventListener('change', handleFileSelect, false);
btnReset.addEventListener('click', resetDashboard, false);

function resetDashboard() {
    uploadArea.classList.remove('hidden');
    dashboardContainer.classList.add('hidden');
    btnReset.classList.add('hidden');
    fileInput.value = '';
    fileNameDisplay.textContent = '';
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    fileNameDisplay.textContent = file.name;
    hideError();

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const rows = XLSX.utils.sheet_to_json(worksheet, { range: 1, defval: null });

            if (rows.length === 0) {
                showError("A planilha parece estar vazia ou não tem dados após a primeira linha.");
                return;
            }

            processExcelData(rows);

        } catch (error) {
            console.error("Erro:", error);
            showError("Erro ao processar o arquivo Excel. Verifique se os cabeçalhos estão na segunda linha.");
        }
    };

    reader.readAsArrayBuffer(file);
}

function processExcelData(rows) {
    const keys = Object.keys(rows[0] || {});
    const normalizeStr = (str) => str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();

    const statusColName = keys.find(k => normalizeStr(k) === 'status');
    const capituloColName = keys.find(k => normalizeStr(k) === 'capitulo');
    const respondenteColName = keys.find(k => normalizeStr(k) === 'respondente');

    if (!statusColName) {
        showError("Coluna 'Status' não encontrada na segunda linha do Excel.");
        return;
    }

    let totalValidRows = 0;
    let finalizadoCount = 0;
    const statusCounts = {};
    const capituloData = {};
    const respondenteData = {};

    rows.forEach(row => {
        let statusRaw = row[statusColName];

        if (statusRaw === null || statusRaw === undefined || statusRaw.toString().trim() === "") {
            return;
        }

        totalValidRows++;
        const statusValue = statusRaw.toString().trim();
        const statusKey = statusValue.toLowerCase();

        if (statusKey === 'finalizado') finalizadoCount++;

        statusCounts[statusValue] = (statusCounts[statusValue] || 0) + 1;

        const aggregateDynamicData = (colName, dataObj, fallbackLabel) => {
            if (colName) {
                let val = row[colName];
                let label = (val !== null && val !== undefined && val.toString().trim() !== "") ? val.toString().trim() : fallbackLabel;

                if (!dataObj[label]) {
                    dataObj[label] = { total: 0 };
                }

                dataObj[label][statusValue] = (dataObj[label][statusValue] || 0) + 1;
                dataObj[label].total++;
            }
        };

        aggregateDynamicData(capituloColName, capituloData, "Sem Capítulo");
        aggregateDynamicData(respondenteColName, respondenteData, "Sem Respondente");
    });

    if (totalValidRows === 0) {
        showError("Não foram encontrados dados válidos na coluna 'Status'.");
        return;
    }

    updateDashboard(totalValidRows, finalizadoCount, statusCounts, capituloData, respondenteData);
}

function updateDashboard(total, finalizado, statusCounts, capituloData, respondenteData) {
    uploadArea.classList.add('hidden');
    btnReset.classList.remove('hidden');

    dashboardContainer.classList.remove('hidden');
    dashboardContainer.classList.add('flex');

    document.getElementById('stat-total').textContent = total;
    document.getElementById('stat-finalizado').textContent = finalizado;
    document.getElementById('stat-pendente').textContent = total - finalizado;

    const progressPercentage = ((finalizado / total) * 100).toFixed(1);
    document.getElementById('progress-percentage').textContent = `${progressPercentage}%`;

    Chart.defaults.font.family = "'Inter', sans-serif";
    Chart.defaults.color = '#64748B';
    Chart.defaults.animation = false; 
    Chart.defaults.devicePixelRatio = Math.max(2, window.devicePixelRatio || 1); 

    Object.values(charts).forEach(chart => {
        if (chart) chart.destroy();
    });

    const uniqueStatuses = Object.keys(statusCounts);

    
    const ctxProgress = document.getElementById('progressChart').getContext('2d');
    charts.progress = new Chart(ctxProgress, {
        type: 'doughnut',
        data: {
            labels: ['Finalizado', 'Outros'],
            datasets: [{
                data: [finalizado, total - finalizado],
                backgroundColor: [statusColorMap['finalizado'], '#E2E8F0'],
                borderWidth: 0,
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false, rotation: -90, circumference: 180, cutout: '80%',
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 10, padding: 10, font: { size: 11 } } },
                datalabels: { display: false },
                tooltip: { callbacks: { label: context => ` ${context.label}: ${context.parsed} (${((context.parsed / total) * 100).toFixed(1)}%)` } }
            }
        }
    });

    
    const ctxStatus = document.getElementById('statusChart').getContext('2d');
    const dataStatus = uniqueStatuses.map(label => statusCounts[label]);
    const bgStatusColors = uniqueStatuses.map((label, index) => getColorForStatus(label, index));

    charts.status = new Chart(ctxStatus, {
        type: 'bar',
        data: {
            
            labels: uniqueStatuses.map(l => l.split(' ')),
            datasets: [{ data: dataStatus, backgroundColor: bgStatusColors, borderRadius: 4, barPercentage: 0.6 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: {
                    anchor: 'end', align: 'top', color: '#334155',
                    font: { weight: 'bold', size: 11 },
                    formatter: value => value
                },
                tooltip: {
                    callbacks: {
                        title: items => uniqueStatuses[items[0].dataIndex],
                        label: context => ` ${context.parsed.y} (${((context.parsed.y / total) * 100).toFixed(1)}%)`
                    }
                }
            },
            scales: {
                y: { beginAtZero: true, grace: '10%', grid: { display: false }, ticks: { display: false }, border: { display: false } },
                x: {
                    grid: { display: false },
                    border: { display: false },
                    
                    ticks: { font: { size: 11 }, maxRotation: 0, minRotation: 0, autoSkip: false }
                }
            },
            layout: { padding: { top: 20 } }
        }
    });

    
    stackedData = { capitulo: capituloData, respondente: respondenteData, statuses: uniqueStatuses };
    renderStackedCharts();

    
    document.fonts.ready.then(buildPrintSnapshots);
}


let stackedData = null;
let viewMode = 'all'; 

const isFinalizado = (name) => name.toLowerCase() === 'finalizado';

function createStackedChart(canvasId, dataObj, minHeight) {
    const statuses = stackedData.statuses;
    const ctx = document.getElementById(canvasId).getContext('2d');

    const labels = Object.keys(dataObj).sort((a, b) => dataObj[b].total - dataObj[a].total);

    
    const wrapper = ctx.canvas.parentElement;
    wrapper.style.height = Math.min(800, Math.max(minHeight, labels.length * 38 + 70)) + 'px';

    let datasets;

    if (viewMode === 'simple') {
        
        const finalizadoOf = (label) => statuses
            .filter(isFinalizado)
            .reduce((sum, st) => sum + (dataObj[label][st] || 0), 0);

        datasets = [
            {
                label: 'Finalizado',
                data: labels.map(finalizadoOf),
                backgroundColor: statusColorMap['finalizado'],
                borderRadius: 4,
                barPercentage: 0.75,
                maxBarThickness: 28,
                datalabels: { color: '#ffffff' }
            },
            {
                label: 'Outros Status',
                data: labels.map(l => dataObj[l].total - finalizadoOf(l)),
                backgroundColor: '#CBD5E1',
                borderRadius: 4,
                barPercentage: 0.75,
                maxBarThickness: 28,
                datalabels: { color: '#1E293B' }
            }
        ];
    } else {
        
        datasets = statuses.map((statusName, index) => ({
            label: statusName,
            data: labels.map(labelItem => dataObj[labelItem][statusName] || 0),
            backgroundColor: getColorForStatus(statusName, index),
            borderRadius: 4,
            barPercentage: 0.75,
            maxBarThickness: 28,
            datalabels: { color: getTextColorForStatus(statusName) }
        }));
    }

    return new Chart(ctx, {
        type: 'bar',
        data: { labels: labels, datasets: datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            layout: { padding: { right: 0 } },
            plugins: {
                legend: {
                    position: 'top',
                    align: 'start',
                    labels: { usePointStyle: true, boxWidth: 6, font: { size: 11 }, padding: 15 }
                },
                datalabels: {
                    font: { weight: 'bold', size: 10 },
                    formatter: value => value > 0 ? value : ''
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    callbacks: {
                        afterBody: function(tooltipItems) {
                            let totalItem = 0;
                            tooltipItems.forEach(item => totalItem += item.parsed.x);
                            return `\nTotal: ${totalItem}`;
                        }
                    }
                }
            },
            scales: {
                x: { display: false, stacked: true },
                y: {
                    stacked: true,
                    grid: { display: false, drawBorder: false },
                    border: { display: false },
                    ticks: { font: { size: 10 } }
                }
            }
        }
    });
}

function renderStackedCharts() {
    if (!stackedData) return;

    if (charts.capitulo) { charts.capitulo.destroy(); charts.capitulo = null; }
    if (charts.respondente) { charts.respondente.destroy(); charts.respondente = null; }

    if (Object.keys(stackedData.capitulo).length > 0) {
        charts.capitulo = createStackedChart('capituloChart', stackedData.capitulo, 150);
    }
    if (Object.keys(stackedData.respondente).length > 0) {
        charts.respondente = createStackedChart('respondenteChart', stackedData.respondente, 150);
    }
}

function setViewMode(mode) {
    viewMode = mode;
    document.querySelectorAll('.mode-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.mode === mode);
    });
    renderStackedCharts();
    buildPrintSnapshots(); 
}

document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.addEventListener('click', () => setViewMode(btn.dataset.mode));
});

function showError(message) {
    errorText.textContent = message;
    errorMessage.classList.remove('hidden');
    fileInput.value = '';
}

function hideError() {
    errorMessage.classList.add('hidden');
}


function redrawCharts() {
    void dashboardContainer.offsetWidth; 
    Object.values(charts).forEach(chart => { if (chart) chart.resize(); });
}

function buildPrintSnapshots() {
    if (dashboardContainer.classList.contains('hidden')) return;

    document.querySelectorAll('.chart-snapshot').forEach(img => img.remove());

    dashboardContainer.classList.add('print-mode');
    redrawCharts();

    Object.values(charts).forEach(chart => {
        if (!chart) return;
        const img = new Image();
        img.className = 'chart-snapshot';
        img.alt = '';
        img.src = chart.toBase64Image('image/png', 1);
        chart.canvas.insertAdjacentElement('afterend', img);
    });

    dashboardContainer.classList.remove('print-mode');
    redrawCharts();
}