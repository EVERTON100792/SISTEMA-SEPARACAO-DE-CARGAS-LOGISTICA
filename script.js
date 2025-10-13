// ================================================================================================
//  A LÓGICA DO GRÁFICO FOI ATUALIZADA PARA MELHORAR A VISUALIZAÇÃO DOS NÚMEROS E A ESTÉTICA.
//  O RESTANTE DO CÓDIGO JAVASCRIPT PERMANECE IDÊNTICO E TOTALMENTE FUNCIONAL.
// ================================================================================================
let resumoChart = null;

function updateAndRenderChart() {
    const vehicleCounts = { fiorino: 0, van: 0, tresQuartos: 0, toco: 0 };
    const vehicleWeights = { fiorino: 0, van: 0, tresQuartos: 0, toco: 0 };

    for (const loadId in activeLoads) {
        const load = activeLoads[loadId];
        if (load.vehicleType === 'fiorino') {
            vehicleCounts.fiorino++;
            vehicleWeights.fiorino += load.totalKg;
        } else if (load.vehicleType === 'van') {
            vehicleCounts.van++;
            vehicleWeights.van += load.totalKg;
        } else if (load.vehicleType === 'tresQuartos') {
            vehicleCounts.tresQuartos++;
            vehicleWeights.tresQuartos += load.totalKg;
        }
    }

    vehicleCounts.toco = Object.keys(gruposToco).length;
    for (const cf in gruposToco) {
        vehicleWeights.toco += gruposToco[cf].totalKg;
    }

    const totalVarejoCount = vehicleCounts.fiorino + vehicleCounts.van + vehicleCounts.tresQuartos + vehicleCounts.toco;
    const totalVarejoWeight = vehicleWeights.fiorino + vehicleWeights.van + vehicleWeights.tresQuartos + vehicleWeights.toco;

    const countSeriesData = [vehicleCounts.fiorino, vehicleCounts.van, vehicleCounts.tresQuartos, vehicleCounts.toco, totalVarejoCount];
    const weightSeriesData = [vehicleWeights.fiorino, vehicleWeights.van, vehicleWeights.tresQuartos, vehicleWeights.toco, totalVarejoWeight];
    const categories = ['Fiorino', 'Van', '3/4', 'Toco', 'Total Varejo'];

    const series = [
        { name: 'Quantidade', type: 'column', data: countSeriesData },
        { name: 'Peso (kg)', type: 'column', data: weightSeriesData }
    ];

    if (resumoChart) {
        resumoChart.updateSeries(series);
    } else {
        const options = {
            series: series,
            chart: {
                height: 350,
                type: 'line',
                stacked: false,
                toolbar: { show: false },
                foreColor: 'var(--dark-text-secondary)',
                animations: {
                    enabled: true,
                    easing: 'easeinout',
                    speed: 800,
                    animateGradually: {
                        enabled: true,
                        delay: 150
                    },
                    dynamicAnimation: {
                        enabled: true,
                        speed: 350
                    }
                }
            },
            fill: {
                type: 'gradient',
                gradient: {
                    shade: 'dark',
                    type: "vertical",
                    shadeIntensity: 0.5,
                    gradientToColors: undefined,
                    inverseColors: false,
                    opacityFrom: 0.85,
                    opacityTo: 0.95,
                    stops: [0, 100]
                }
            },
            stroke: {
                width: [0, 0],
                curve: 'smooth'
            },
            plotOptions: {
                bar: {
                    columnWidth: '60%',
                    borderRadius: 5
                }
            },
            colors: ['#6a5acd', '#00e396'],
            dataLabels: {
                enabled: true,
                formatter: function (val, { seriesIndex }) {
                    if (seriesIndex === 0) { // Quantidade
                        return val.toFixed(0);
                    } else { // Peso
                        if (val >= 1000) {
                            return (val / 1000).toFixed(1) + 'k';
                        }
                        return val.toFixed(0);
                    }
                },
                offsetY: -25,
                style: {
                    fontSize: '13px',
                    fontWeight: 'bold',
                    colors: ['#fff']
                },
                background: {
                  enabled: true,
                  foreColor: '#fff',
                  borderRadius: 3,
                  padding: 5,
                  opacity: 0.7,
                  borderWidth: 1,
                  borderColor: '#666',
                },
            },
            grid: {
                borderColor: 'var(--dark-border)',
                strokeDashArray: 4,
                yaxis: { lines: { show: true } }
            },
            xaxis: {
                categories: categories,
                labels: { 
                    style: { 
                        fontWeight: 600,
                        fontSize: '13px'
                    } 
                }
            },
            yaxis: [
                {
                    seriesName: 'Quantidade',
                    axisTicks: { show: true },
                    axisBorder: { show: true, color: '#6a5acd' },
                    labels: {
                        style: { colors: '#6a5acd' },
                        formatter: function (val) {
                            return val.toFixed(0);
                        }
                    },
                    title: {
                        text: "Quantidade de Veículos",
                        style: { color: '#6a5acd', fontWeight: 600 }
                    },
                },
                {
                    seriesName: 'Peso (kg)',
                    opposite: true,
                    axisTicks: { show: true },
                    axisBorder: { show: true, color: '#00e396' },
                    labels: {
                        style: { colors: '#00e396' },
                        formatter: function (val) {
                            return (val / 1000).toFixed(1) + "k kg";
                        }
                    },
                    title: {
                        text: "Peso Total (kg)",
                        style: { color: '#00e396', fontWeight: 600 }
                    }
                }
            ],
            tooltip: {
                theme: 'dark',
                shared: true,
                intersect: false,
                y: {
                    formatter: function (val, { seriesIndex }) {
                        if (seriesIndex === 0) { // Quantidade
                            return val.toFixed(0) + " veículos";
                        } else { // Peso
                            return val.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " kg";
                        }
                    }
                }
            },
            legend: {
                position: 'top',
                horizontalAlign: 'right',
                offsetY: -10
            }
        };

        const chartContainer = document.querySelector("#chart-resumo-veiculos");
        if (chartContainer) {
            chartContainer.innerHTML = '';
            resumoChart = new ApexCharts(chartContainer, options);
            resumoChart.render();
        }
    }

    const container = document.getElementById('card-resumo-geral-container');
    if (container) container.style.display = 'block';
}

const agendamentoClientCodes = new Set([
    '1398', '1494', '4639', '4872', '5546', '6896', 'D11238', '17163', '19622', '20350', '22545', '23556', '23761', '24465', '29302', '32462', '32831', '32851', '32869', '32905', '33039', '33046', '33047', '33107', '33388', '33392', '33400', '33401', '33403', '33406', '33420', '33494', '33676', '33762', '33818', '33859', '33907', '33971', '34011', '34096', '34167', '34425', '34511', '34810', '34981', '35050', '35054', '35798', '36025', '36580', '36792', '36853', '36945', '37101', '37589', '37634', '38207', '38448', '38482', '38564', '38681', '38735', 'D38896', '39081', '39177', '39620', '40144', '40442', '40702', '40844', '41233', '42200', '42765', '47244', '47253', '47349', '50151', '50816', '51993', '52780', '53134', '58645', '60900', '61182', '61315', '61316', '61317', '61318', '61324', '63080', '63500', '63705', '64288', '66590', '67660', '67884', '69281', '69286', '69318', '70968', '71659', '73847', '76019', '76580', '77475', '77520', '78895', '79838', '80727', '81353', 'DB3183', '83184', '83634', '85534', 'DB6159', '86350', '86641', '89073', '89151', '90373', '92017', '95092', '95660', '96758', '98227', '99268', '100087', '101246', '101253', '101346', '103518', '105394', '106198', '109288', '110023', '110894', '111145', '111154', '111302', '112207', '112670', '117028', '117123', '120423', '120455', '120473', '120533', '121747', '122155', '122785', '123815', '124320', '125228', '126430', '131476', '132397', '133916', '135395', '135928', '136086', '136260', '137919', '138825', '139013', '139329', '139611', '143102', '44192', '144457', '145014', '145237', '145322', '146644', '146988', '148071', '149598', '150503', '151981', '152601', '152835', '152925', '153289', '154423', '154778', '154808', '155177', '155313', '155368', '155419', '155475', '155823', '155888', '156009', '156585', '156696', '157403', '158235', '159168', '160382', '160982', '161737', '162499', '162789', '163234', '163382', '163458', '164721', '164779', '164780', '164924', '165512', '166195', '166337', '166353', '166468', '166469', '167353', '167810', '167819', '168464', '169863', '169971', '170219', '170220', '170516', '171147', '171160', '171191', '171200', '171320', '171529', '171642', '171863', '172270', '172490', '172656', '172859', '173621', '173964', '173977', '174249', '174593', '174662', '174901', '175365', '175425', '175762', '175767', '175783', '176166', '176278', '176453', '176747', '177327', '177488', '177529', '177883', '177951', '177995', '178255', '178377', '178666', '179104', '179510', '179542', '179690', '180028', '180269', '180342', '180427', '180472', '180494', '180594', '180772', '181012', '181052', '181179', '182349', '182885', '182901', '183011', '183016', '183046', '183048', '183069', '183070', '183091', '183093', '183477', '183676', '183787', '184011', '184038', '189677', '190163', '190241', '190687', '190733', '191148', '191149', '191191', '191902', '191972', '192138', '192369', '192638', '192713', '193211', '193445', '193509', '194432', '194508', '194750', '194751', '194821', '194831', '195287', '195338', '195446', '196118', '196405', '196446', '196784', '197168', '197249', '197983', '198187', '198438', '198747', '198796', '198895', '198907', '198908', '199172', '199615', '199625', '199650', '199651', '199713', '199733', '199927', '199991', '200091', '200194', '200239', '200253', '200382', '200404', '200597', '200917', '201294', '201754', '201853', '201936', '201948', '201956', '201958', '201961', '201974', '202022', '202187', '202199', '202714', '203072', '203093', '203201', '203435', '203436', '203451', '203512', '203769', '204895', '204910', '204911', '204913', '204914', '204915', '204917', '204971', '204979', '205108', '205220', '205744', '205803', '206116', '206163', '206208', '206294', '206380', '206628', '206730', '206731', '206994', '207024', '207029', '207403', '207689', '207902', '208489', '208613', '208622', '208741', '208822', '208844', '208853', '208922', '209002', '209004', '209248', '209281', '209321', '209322', '209684', '210124', '210230', '210490', '210747', '210759', '210819', '210852', '211059', '211110', '211276', '211277', '211279', '211332', '211411', '212401', '212417', '212573', '212900', '213188', '213189', '213190', '213202', '213203', '213242', '213442', '213454', '213855', '213909', '213910', '213967', '214046', '214150', '214387', '214433', '214442', '214594', '214746', '215022', '215116', '215160', '215161', '215493', '215494', '215651', '215687', '215733', '215777', '215942', '216112', '216393', '216400', '216630', '216684', '217190', '217283', '217310', '217343', '217545', '217605', '217828', '217871', '217872', '217877', '217949', '217965', '218169', '218196', '218383', '218486', '218578', '218580', '218640', '218820', '218845', '219539', '219698', '219715', '219884', '220158', '220183', '220645', '220950', '221023', '221248', '221251', '222164', '222165', '223025', '223379', '223525', '223703', '223727', '223877', '223899', '223900', '223954', '224956', '224957', '224958', '224959', '224961', '224962', '225112', '225408', '225449', '225904', '226903', '226939', '227190', '227387', '228589', '228693', '228695'
]);

const specialClientNames = ['IRMAOS MUFFATO S.A', 'FINCO & FINCO', 'BOM DIA'];

// ================================================================================================
//  INÍCIO DO CÓDIGO JAVASCRIPT PRINCIPAL
// ================================================================================================

function deepClone(obj) {
    if (obj === null || typeof obj !== 'object') {
        return obj;
    }
    if (obj instanceof Date) {
        return new Date(obj.getTime());
    }
    if (Array.isArray(obj)) {
        const arrCopy = [];
        for (let i = 0; i < obj.length; i++) {
            arrCopy[i] = deepClone(obj[i]);
        }
        return arrCopy;
    }
    const objCopy = {};
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            objCopy[key] = deepClone(obj[key]);
        }
    }
    return objCopy;
}

function normalizeClientId(id) {
    if (id === null || typeof id === 'undefined') return '';
    return String(id).trim().replace(/^0+/, '');
}

function checkAgendamento(pedido) {
    const normalizedCode = normalizeClientId(pedido.Cliente);
    pedido.Agendamento = agendamentoClientCodes.has(normalizedCode) ? 'Sim' : 'Não';
}

const isSpecialClient = (p) => p.Nome_Cliente && specialClientNames.includes(p.Nome_Cliente.toUpperCase().trim());

let planilhaData = [];
let originalColumnHeaders = [];
let pedidosGeraisAtuais = [];
let gruposToco = {};
let gruposPorCFGlobais = {};
let pedidosComCFNumericoIsolado = [];
let pedidosManualmenteBloqueadosAtuais = [];
let pedidosPrioritarios = [];
let pedidosBloqueados = new Set();
let pedidosEspeciaisProcessados = new Set();
let pedidosSemCorte = new Set();
let pedidosVendaAntecipadaProcessados = new Set();
let rota1SemCarga = [];
let pedidosFuncionarios = [];
let pedidosCarretaSemCF = [];
let pedidosTransferencias = [];
let pedidosCondorTruck = [];
let tocoPedidoIds = new Set();
let currentLeftoversForPrinting = [];
let activeLoads = {};
let manualLoadInProgress = null;

const defaultConfigs = {
    fiorinoMinCapacity: 300, fiorinoMaxCapacity: 500, fiorinoCubage: 1.5, fiorinoHardMaxCapacity: 560, fiorinoHardCubage: 1.7,
    vanMinCapacity: 1100, vanMaxCapacity: 1560, vanCubage: 5.0, vanHardMaxCapacity: 1600, vanHardCubage: 5.6,
    tresQuartosMinCapacity: 2300, tresQuartosMaxCapacity: 4100, tresQuartosCubage: 15.0,
    tocoMinCapacity: 5000, tocoMaxCapacity: 8500, tocoCubage: 30.0
};

function saveConfigurations() {
    const configStatus = document.getElementById('configStatus');
    configStatus.innerHTML = '<p class="text-info">Salvando configurações...</p>';
    try {
        const configs = {};
        Object.keys(defaultConfigs).forEach(key => {
            const element = document.getElementById(key);
            if (element) {
                configs[key] = parseFloat(element.value);
            }
        });
        localStorage.setItem('vehicleConfigs', JSON.stringify(configs));
        configStatus.innerHTML = '<p class="text-success">Configurações salvas!</p>';
        setTimeout(() => { configStatus.innerHTML = ''; }, 3000);
    } catch (error) {
        console.error("Erro ao salvar no localStorage:", error);
        configStatus.innerHTML = `<p class="text-danger">Erro ao salvar: ${error.message}</p>`;
    }
}

function loadConfigurations() {
    try {
        const savedConfigs = localStorage.getItem('vehicleConfigs');
        const configs = savedConfigs ? JSON.parse(savedConfigs) : defaultConfigs;
        Object.keys(defaultConfigs).forEach(key => {
            const element = document.getElementById(key);
            if (element) {
                element.value = configs[key] !== undefined ? configs[key] : defaultConfigs[key];
            }
        });
    } catch (error) {
        console.error("Erro ao carregar do localStorage:", error);
        const configStatus = document.getElementById('configStatus');
        configStatus.innerHTML = `<p class="text-warning">Não foi possível carregar configurações. Usando valores padrão.</p>`;
    }
}

function setupConfigButtons() {
    document.getElementById('saveConfig').addEventListener('click', saveConfigurations);
    document.getElementById('resetAllConfigs').addEventListener('click', () => {
        Object.keys(defaultConfigs).forEach(key => {
            document.getElementById(key).value = defaultConfigs[key];
        });
        saveConfigurations();
    });
    ['Fiorino', 'Van', 'TresQuartos', 'Toco'].forEach(type => {
        document.getElementById(`reset${type}`).addEventListener('click', () => {
            Object.keys(defaultConfigs).forEach(key => {
                if (key.toLowerCase().startsWith(type.toLowerCase())) {
                    document.getElementById(key).value = defaultConfigs[key];
                }
            });
            saveConfigurations();
        });
    });
}


document.addEventListener('DOMContentLoaded', () => {
    try {
        loadConfigurations();
    } catch (e) {
        console.error("Falha crítica ao carregar configurações.", e);
    }

    setupConfigButtons();
    document.getElementById('limparResultadosBtn').addEventListener('click', limparTudo);
    document.getElementById('pedidoSearchInput').addEventListener('keyup', (event) => {
        if (event.key === 'Enter') {
            buscarPedido();
        }
    });

    document.querySelectorAll('#sidebar-nav a').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            document.querySelector('#sidebar-nav a.active')?.classList.remove('active');
            this.classList.add('active');
            const targetElement = document.querySelector(this.getAttribute('href'));
            if (targetElement) {
                targetElement.scrollIntoView({ behavior: 'smooth' });
            }
        });
    });
});

const rotaVeiculoMap = {
    '11101': { type: 'fiorino', title: 'Rota 11101' }, '11301': { type: 'fiorino', title: 'Rota 11301' }, '11311': { type: 'fiorino', title: 'Rota 11311' }, '11561': { type: 'fiorino', title: 'Rota 11561' }, '11721': { type: 'fiorino', title: 'Rotas 11721 & 11731', combined: ['11731'] }, '11731': { type: 'fiorino', title: 'Rotas 11721 & 11731', combined: ['11721'] },
    '11102': { type: 'van', title: 'Rota 11102' }, '11331': { type: 'van', title: 'Rota 11331' }, '11341': { type: 'van', title: 'Rota 11341' }, '11342': { type: 'van', title: 'Rota 11342' }, '11351': { type: 'van', title: 'Rota 11351' }, '11521': { type: 'van', title: 'Rota 11521' }, '11531': { type: 'van', title: 'Rota 11531' }, '11551': { type: 'van', title: 'Rota 11551' }, '11571': { type: 'van', title: 'Rota 11571' }, '11701': { type: 'van', title: 'Rota 11701' }, '11711': { type: 'van', title: 'Rota 11711' },
    '11361': { type: 'tresQuartos', title: 'Rota 11361' }, '11501': { type: 'tresQuartos', title: 'Rotas 11501, 11502 & 11511', combined: ['11502', '11511'] }, '11502': { type: 'tresQuartos', title: 'Rotas 11501, 11502 & 11511', combined: ['11501', '11511'] }, '11511': { type: 'tresQuartos', title: 'Rotas 11501, 11502 & 11511', combined: ['11501', '11502'] }, '11541': { type: 'tresQuartos', title: 'Rota 11541' }
};

const fileInput = document.getElementById('fileInput');
const processarBtn = document.getElementById('processarBtn');
const statusDiv = document.getElementById('status');
const isNumeric = (str) => str && /^\d+$/.test(String(str).trim());

fileInput.addEventListener('change', (event) => { handleFile(event.target.files[0]); });

function handleFile(file) {
    if (!file) return;
    limparTudo();
    statusDiv.innerHTML = '<p class="text-info">Carregando planilha...</p>';
    processarBtn.disabled = true;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellDates:true});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            let headerRowIndex = -1;
            for (let i = 0; i < rawData.length; i++) {
                const row = rawData[i];
                if (row && row.some(cell => String(cell).trim().toLowerCase() === 'cod_rota')) {
                    headerRowIndex = i;
                    originalColumnHeaders = row.map(h => h ? String(h).trim() : '');
                    break;
                }
            }

            if (headerRowIndex === -1) throw new Error("Não foi possível encontrar a linha de cabeçalho com 'Cod_Rota'.");
            
            const dataRows = rawData.slice(headerRowIndex + 1);
            planilhaData = dataRows.map(row => { 
                const pedido = {}; 
                originalColumnHeaders.forEach((header, i) => { 
                    if (header) { 
                        if ((header.toLowerCase() === 'predat' || header.toLowerCase() === 'dat_ped')) {
                            let cellValue = row[i];
                            if (typeof cellValue === 'number') {
                                const date = new Date(Math.round((cellValue - 25569) * 86400 * 1000));
                                pedido[header] = !isNaN(date.getTime()) ? date : '';
                            } else if (cellValue instanceof Date) {
                               pedido[header] = !isNaN(cellValue.getTime()) ? cellValue : '';
                            } else {
                                pedido[header] = cellValue !== undefined ? cellValue : '';
                            }
                        } else {
                            pedido[header] = row[i] !== undefined ? row[i] : ''; 
                        }
                    } 
                }); 
                pedido.Quilos_Saldo = parseFloat(String(pedido.Quilos_Saldo).replace(',', '.')) || 0;
                pedido.Cubagem = parseFloat(String(pedido.Cubagem).replace(',', '.')) || 0;
                return pedido; 
            });
            planilhaData.forEach(checkAgendamento);
            statusDiv.innerHTML = `<p class="text-success">Planilha "${file.name}" carregada.</p>`;
            processarBtn.disabled = false;
            
            popularFiltrosDeRota();

            if (document.getElementById('autoProcessCheckbox').checked) {
                processar();
            }

        } catch (error) {
            statusDiv.innerHTML = `<p class="text-danger"><strong>Erro:</strong> ${error.message}</p>`;
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function popularFiltrosDeRota() {
    const container = document.getElementById('filtro-rota-container');
    const rotaInicioSelect = document.getElementById('rotaInicioSelect');
    const rotaFimSelect = document.getElementById('rotaFimSelect');

    rotaInicioSelect.innerHTML = '<option value="">Rota de Início...</option>';
    rotaFimSelect.innerHTML = '<option value="">Rota de Fim...</option>';

    if (planilhaData.length === 0) {
        container.style.display = 'none';
        return;
    }

    const rotas = [...new Set(planilhaData.map(p => String(p.Cod_Rota || ''))) ].filter(Boolean);
    rotas.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    rotas.forEach(rota => {
        rotaInicioSelect.innerHTML += `<option value="${rota}">${rota}</option>`;
        rotaFimSelect.innerHTML += `<option value="${rota}">${rota}</option>`;
    });

    container.style.display = 'block';
}

function limparFiltroDeRota() {
    document.getElementById('rotaInicioSelect').selectedIndex = 0;
    document.getElementById('rotaFimSelect').selectedIndex = 0;
}

/**
 * NOVA FUNÇÃO DE BUSCA
 * Busca em todos os dados por número do pedido, nome do cliente ou cidade.
 * Exibe uma lista de resultados, cada um com um botão para destacar.
 */
function buscarPedido() {
    const searchTerm = document.getElementById('pedidoSearchInput').value.trim().toLowerCase();
    const searchResultDiv = document.getElementById('search-result');
    searchResultDiv.innerHTML = '';
    if (!searchTerm) return;

    if (planilhaData.length === 0) {
        searchResultDiv.innerHTML = '<p class="text-warning">Por favor, carregue e processe a planilha primeiro.</p>';
        return;
    }

    // Agrupa todos os pedidos de todas as fontes em um único array para busca
    const allPedidosMap = new Map();
    const allSources = [
        { data: pedidosGeraisAtuais, location: 'Pedidos Disponíveis (Varejo)' },
        { data: currentLeftoversForPrinting, location: 'Sobras Finais' },
        ...Object.values(activeLoads).map(load => ({ data: load.pedidos, location: `Carga ${load.numero || load.id}` })),
        { data: pedidosManualmenteBloqueadosAtuais, location: 'Bloqueado Manualmente' },
        { data: rota1SemCarga, location: 'Rota 1 para Alteração' },
        { data: pedidosComCFNumericoIsolado, location: 'Filtrados (Bloqueio Varejo)' },
        { data: pedidosFuncionarios, location: 'Pedidos de Funcionários' },
        { data: pedidosTransferencias, location: 'Pedidos de Transferência' },
        ...Object.values(gruposPorCFGlobais).map(group => ({ data: group.pedidos, location: `Truck/Carreta (CF: ${group.pedidos[0]?.CF})` }))
    ];

    allSources.forEach(source => {
        source.data.forEach(pedido => {
            if (!allPedidosMap.has(String(pedido.Num_Pedido))) {
                allPedidosMap.set(String(pedido.Num_Pedido), { pedido, location: source.location });
            }
        });
    });

    const resultados = [];
    for (const [_, value] of allPedidosMap) {
        const p = value.pedido;
        if (String(p.Num_Pedido).toLowerCase().includes(searchTerm) ||
            String(p.Nome_Cliente).toLowerCase().includes(searchTerm) ||
            String(p.Cidade).toLowerCase().includes(searchTerm)) {
            resultados.push(value);
        }
    }

    if (resultados.length > 0) {
        let resultHtml = `<div class="list-group">`;
        resultados.forEach(res => {
            resultHtml += `
                <div class="list-group-item list-group-item-action flex-column align-items-start">
                    <div class="d-flex w-100 justify-content-between">
                        <h6 class="mb-1">Pedido: ${res.pedido.Num_Pedido}</h6>
                        <button class="btn btn-sm btn-outline-info" onclick="highlightPedido('${res.pedido.Num_Pedido}')">Destacar</button>
                    </div>
                    <p class="mb-1 small">Cliente: ${res.pedido.Nome_Cliente}</p>
                    <small class="text-muted">Local: ${res.location}</small>
                </div>`;
        });
        resultHtml += `</div>`;
        searchResultDiv.innerHTML = resultHtml;
    } else {
        searchResultDiv.innerHTML = '<div class="alert alert-warning">Nenhum pedido encontrado com o termo informado.</div>';
    }
}

function priorizarPedido(numPedido) { if (!pedidosPrioritarios.includes(String(numPedido))) { pedidosPrioritarios.push(String(numPedido)); buscarPedido(); } }

/**
 * NOVA FUNÇÃO DE DESTAQUE
 * Localiza o pedido, troca para a aba correta (se necessário) e rola a tela.
 */
function highlightPedido(numPedido) {
    document.querySelectorAll('.search-highlight').forEach(el => el.classList.remove('search-highlight'));

    const row = document.getElementById(`pedido-${numPedido}`);
    if (!row) {
        alert("Não foi possível encontrar a linha do pedido na interface. Pode estar em um grupo recolhido.");
        return;
    }

    const tabPane = row.closest('.tab-pane');
    if (tabPane && !tabPane.classList.contains('show')) {
        const tabButton = document.querySelector(`button[data-bs-target="#${tabPane.id}"]`);
        if (tabButton) {
            const tab = new bootstrap.Tab(tabButton);
            tab.show();
            // Espera a animação da aba terminar antes de rolar
            setTimeout(() => {
                highlightAndScroll(row);
            }, 350);
        }
    } else {
        highlightAndScroll(row);
    }
}

function highlightAndScroll(row) {
    const accordionCollapse = row.closest('.accordion-collapse');
    if (accordionCollapse && !accordionCollapse.classList.contains('show')) {
        const collapse = new bootstrap.Collapse(accordionCollapse);
        collapse.show();
    }
    
    row.scrollIntoView({ behavior: 'smooth', block: 'center' });
    row.classList.add('search-highlight');
    setTimeout(() => {
        row.classList.remove('search-highlight');
    }, 2500); // Remove o destaque após 2.5 segundos
}


function atualizarListaBloqueados() {
    const divLista = document.getElementById('lista-pedidos-bloqueados');
    if (pedidosBloqueados.size === 0) {
        divLista.innerHTML = '<div class="empty-state"><i class="bi bi-shield-slash"></i><p>Nenhum pedido bloqueado.</p></div>';
        return;
    }
    const list = document.createElement('ul');
    list.className = 'list-group list-group-flush';
    pedidosBloqueados.forEach(numPedido => {
        const item = document.createElement('li');
        item.className = 'list-group-item d-flex justify-content-between align-items-center py-1 bg-transparent';
        item.innerHTML = `<span>${numPedido}</span> <button class="btn btn-sm btn-outline-secondary" onclick="desbloquearPedido('${numPedido}')">Desbloquear</button>`;
        list.appendChild(item);
    });
    divLista.innerHTML = '';
    divLista.appendChild(list);
}

function bloquearPedido() {
    const input = document.getElementById('bloquearPedidoInput');
    const numPedido = input.value.trim();
    if (numPedido) {
        pedidosBloqueados.add(numPedido);
        input.value = '';
        atualizarListaBloqueados();
    }
}

function desbloquearPedido(numPedido) {
    pedidosBloqueados.delete(numPedido);
    atualizarListaBloqueados();
}

function atualizarListaSemCorte() {
    const divLista = document.getElementById('lista-pedidos-sem-corte');
    if (pedidosSemCorte.size === 0) {
        divLista.innerHTML = '<div class="empty-state"><i class="bi bi-scissors"></i><p>Nenhum pedido marcado.</p></div>';
        return;
    }
    const list = document.createElement('ul');
    list.className = 'list-group list-group-flush';
    pedidosSemCorte.forEach(numPedido => {
        const item = document.createElement('li');
        item.className = 'list-group-item d-flex justify-content-between align-items-center py-1 bg-transparent';
        item.innerHTML = `<span>${numPedido}</span> <button class="btn btn-sm btn-outline-secondary" onclick="removerMarcacaoSemCorte('${numPedido}')">Remover</button>`;
        list.appendChild(item);
    });
    divLista.innerHTML = '';
    divLista.appendChild(list);
}

function marcarPedidosSemCorte() {
    const input = document.getElementById('semCorteInput');
    const numeros = input.value.split(/[\n,; ]+/).map(n => n.trim()).filter(Boolean);
    numeros.forEach(num => pedidosSemCorte.add(num));
    input.value = '';
    atualizarListaSemCorte();
}

function removerMarcacaoSemCorte(numPedido) {
    pedidosSemCorte.delete(numPedido);
    atualizarListaSemCorte();
}

function resetarEstadoGlobal() {
    pedidosGeraisAtuais = [];
    gruposToco = {};
    gruposPorCFGlobais = {};
    pedidosComCFNumericoIsolado = [];
    rota1SemCarga = [];
    pedidosFuncionarios = [];
    pedidosCarretaSemCF = [];
    pedidosTransferencias = [];
    pedidosCondorTruck = [];
    pedidosManualmenteBloqueadosAtuais = [];
    pedidosPrioritarios = [];
    tocoPedidoIds.clear();
    currentLeftoversForPrinting = [];
    activeLoads = {};
}

function limparTudo(){
    resetarEstadoGlobal();
    originalColumnHeaders = []; 
    pedidosEspeciaisProcessados.clear();
    pedidosBloqueados.clear();
    pedidosSemCorte.clear();
    pedidosVendaAntecipadaProcessados.clear();
    
    const idsParaLimpar = [
        'resultado-venda-antecipada', 'resultado-carga-especial', 'resultado-bloqueados', 'resultado-geral', 
        'resultado-fiorino-geral', 'resultado-van-geral', 'resultado-34-geral',
        'resultado-rota1', 'resultado-toco', 
        'resultado-cf-accordion-container', 'resultado-cf-numerico', 'search-result', 
        'resultado-funcionarios-tab', 'resultado-transferencias-tab'
    ];
    const emptyStateMessages = {
        'resultado-geral': '<div class="empty-state"><i class="bi bi-file-earmark-excel"></i><p>Nenhum pedido de varejo disponível.</p></div>',
        'botoes-fiorino': '<div class="empty-state"><i class="bi bi-box"></i><p>Nenhuma rota de Fiorino disponível.</p></div>',
        'botoes-van': '<div class="empty-state"><i class="bi bi-truck-front-fill"></i><p>Nenhuma rota de Van disponível.</p></div>',
        'botoes-34': '<div class="empty-state"><i class="bi bi-truck-flatbed"></i><p>Nenhuma rota de 3/4 disponível.</p></div>'
    };

    idsParaLimpar.forEach(id => {
        const el = document.getElementById(id);
        if(el) el.innerHTML = '';
    });
    Object.keys(emptyStateMessages).forEach(id => {
        const el = document.getElementById(id);
        if(el) el.innerHTML = emptyStateMessages[id];
    });

    if (resumoChart) {
        resumoChart.destroy();
        resumoChart = null;
    }
    document.getElementById('card-resumo-geral-container').style.display = 'none';
    document.getElementById('filtro-rota-container').style.display = 'none';

    document.getElementById('pedidosEspeciaisInput').value = '';
    document.getElementById('bloquearPedidoInput').value = '';
    document.getElementById('pedidoSearchInput').value = '';
    document.getElementById('semCorteInput').value = '';
    document.getElementById('vendaAntecipadaInput').value = '';
    atualizarListaBloqueados();
    atualizarListaSemCorte();
    limparFiltroDeRota();
    
    fileInput.value = '';
    planilhaData = [];
    processarBtn.disabled = true;
    statusDiv.innerHTML = '';
}

function processar() {
    statusDiv.innerHTML = '<div class="d-flex align-items-center justify-content-center"><div class="spinner-border spinner-border-sm text-primary me-2" role="status"></div><span class="text-primary">Processando...</span></div>';
    processarBtn.disabled = true;
    
    setTimeout(() => {
        const divsDeResultado = {
            geral: document.getElementById('resultado-geral'),
            toco: document.getElementById('resultado-toco'),
            cfNumerico: document.getElementById('resultado-cf-numerico'),
            rota1: document.getElementById('resultado-rota1'),
            bloqueados: document.getElementById('resultado-bloqueados'),
            cfAccordion: document.getElementById('resultado-cf-accordion-container'),
            funcionarios: document.getElementById('resultado-funcionarios-tab'),
            transferencias: document.getElementById('resultado-transferencias-tab')
        };
        
        try {
            if (planilhaData.length === 0) { throw new Error("Nenhum dado de planilha carregado."); }
            
            resetarEstadoGlobal();

            const rotaInicio = document.getElementById('rotaInicioSelect').value;
            const rotaFim = document.getElementById('rotaFimSelect').value;
            let dadosParaProcessar = [...planilhaData];

            if (rotaInicio && rotaFim) {
                dadosParaProcessar = planilhaData.filter(p => {
                    const rotaPedido = String(p.Cod_Rota || '');
                    return rotaPedido.localeCompare(rotaInicio, undefined, { numeric: true }) >= 0 &&
                           rotaPedido.localeCompare(rotaFim, undefined, { numeric: true }) <= 0;
                });
            }
            
            Object.values(divsDeResultado).forEach(div => { if (div) div.innerHTML = ''; });
            ['resultado-fiorino-geral', 'resultado-van-geral', 'resultado-34-geral'].forEach(id => {
                const el = document.getElementById(id);
                if(el) el.innerHTML = '';
            });

            // Separação inicial de listas especiais
            const pedidosExcluidos = new Set();
            dadosParaProcessar.forEach(p => {
                const coluna5 = String(p.Coluna5 || '').toUpperCase();
                if(coluna5.includes('TBL FUNCIONARIO')) {
                    pedidosFuncionarios.push(p);
                    pedidosExcluidos.add(p.Num_Pedido);
                } else if (coluna5.includes('TBL ESP CARRETA') && !isNumeric(p.CF)) {
                    pedidosCarretaSemCF.push(p);
                    pedidosExcluidos.add(p.Num_Pedido);
                } else if (coluna5.includes('TRANSF. TODESCH')) {
                    pedidosTransferencias.push(p);
                    pedidosExcluidos.add(p.Num_Pedido);
                } else if (coluna5.includes('CONDOR (TRUCK)') || coluna5.includes('CONDOR TOD TRUC')) {
                    pedidosCondorTruck.push(p);
                    pedidosExcluidos.add(p.Num_Pedido);
                }
            });
            
            displayGenericAccordionList(divsDeResultado.funcionarios, pedidosFuncionarios, "Pedidos de Funcionários", "Func", 'Pedidos com a tag "TBL FUNCIONARIO" na Coluna 5.');
            displayGenericAccordionList(divsDeResultado.transferencias, pedidosTransferencias, "Pedidos de Transferência", "Transf", 'Pedidos com a tag "TRANSF. TODESCH" na Coluna 5.');

            let dadosFiltrados = dadosParaProcessar.filter(p => !pedidosExcluidos.has(p.Num_Pedido));
            
            let pedidosDisponiveis = [];
            dadosFiltrados.filter(p => p.Num_Pedido).forEach(p => {
                if (pedidosBloqueados.has(String(p.Num_Pedido)) || pedidosEspeciaisProcessados.has(String(p.Num_Pedido)) || pedidosVendaAntecipadaProcessados.has(String(p.Num_Pedido))) {
                    pedidosManualmenteBloqueadosAtuais.push(p);
                } else {
                    pedidosDisponiveis.push(p);
                }
            });
            displayPedidosBloqueados(divsDeResultado.bloqueados, pedidosManualmenteBloqueadosAtuais);
            
            rota1SemCarga = pedidosDisponiveis.filter(p => {
                const codRota = String(p.Cod_Rota || '').trim();
                return codRota === '1' && !isNumeric(p.CF) && ['TBL 08', 'TBL TODESCHINI', 'PROMO BOLINHO'].some(termo => String(p.Coluna5 || '').toUpperCase().includes(termo));
            });
            const rota1PedidoIds = new Set(rota1SemCarga.map(p => p.Num_Pedido));
            displayRota1(divsDeResultado.rota1, rota1SemCarga);

            const pedidosParaProcessamentoGeral = pedidosDisponiveis.filter(p => !rota1PedidoIds.has(p.Num_Pedido));
            const pedidos = pedidosParaProcessamentoGeral.filter(p => String(p.Coluna4) != '500');

            // Lógica de Toco e CFs
            const clientesComBloqueio = new Set(pedidos.filter(p => String(p['BLOQ.']).trim()).map(p => normalizeClientId(p.Cliente)));
            const pedidosTocoBase = pedidos.filter(p => (p.Coluna4 && String(p.Coluna4).toUpperCase().includes('TOCO')) || (p.Coluna5 && String(p.Coluna5).toUpperCase().includes('TOCO')));
            const cfCounts = pedidosTocoBase.reduce((acc, p) => { if (p.CF && isNumeric(p.CF)) acc[p.CF] = (acc[p.CF] || 0) + 1; return acc; }, {});
            const cfsRepetidos = Object.keys(cfCounts).filter(cf => cfCounts[cf] > 1);
            const pedidosTocoFiltrados = pedidosTocoBase.filter(p => cfsRepetidos.includes(String(p.CF)));
            gruposToco = pedidosTocoFiltrados.reduce((acc, p) => {
                const cf = p.CF;
                if (!acc[cf]) acc[cf] = { pedidos: [], totalKg: 0, totalCubagem: 0 };
                acc[cf].pedidos.push(p);
                acc[cf].totalKg += p.Quilos_Saldo;
                acc[cf].totalCubagem += p.Cubagem;
                return acc;
            }, {});
            displayToco(divsDeResultado.toco, gruposToco);
            tocoPedidoIds = new Set(pedidosTocoFiltrados.map(p => String(p.Num_Pedido)));

            // Separação final entre Varejo e Cargas Grandes (CF)
            let pedidosParaProcessamentoVarejo = [];
            gruposPorCFGlobais = {};
            pedidosComCFNumericoIsolado = [];
            
            pedidos.forEach(p => {
                if (tocoPedidoIds.has(String(p.Num_Pedido)) || ['21', '23', '17'].includes(String(p.Coluna4))) return;
                
                if (isNumeric(p.CF)) {
                    const cf = String(p.CF).trim();
                    if (!gruposPorCFGlobais[cf]) gruposPorCFGlobais[cf] = { pedidos: [], totalKg: 0, totalCubagem: 0 };
                    gruposPorCFGlobais[cf].pedidos.push(p);
                    gruposPorCFGlobais[cf].totalKg += p.Quilos_Saldo;
                    gruposPorCFGlobais[cf].totalCubagem += p.Cubagem;
                } else {
                    if (clientesComBloqueio.has(normalizeClientId(p.Cliente))) {
                        pedidosComCFNumericoIsolado.push(p);
                    } else {
                        pedidosParaProcessamentoVarejo.push(p);
                    }
                }
            });
            
            displayCargasCfAccordion(divsDeResultado.cfAccordion, gruposPorCFGlobais, pedidosCarretaSemCF, pedidosCondorTruck);
            displayPedidosCFNumerico(divsDeResultado.cfNumerico, pedidosComCFNumericoIsolado);

            pedidosGeraisAtuais = [...pedidosParaProcessamentoVarejo]; 
            const gruposGerais = pedidosGeraisAtuais.reduce((acc, p) => {
                const rota = p.Cod_Rota; if (!acc[rota]) { acc[rota] = { pedidos: [], totalKg: 0 }; } acc[rota].pedidos.push(p); acc[rota].totalKg += p.Quilos_Saldo; return acc;
            }, {});
            displayGerais(divsDeResultado.geral, gruposGerais);
            
            updateAndRenderChart();
            
            statusDiv.innerHTML = `<p class="text-success">Processamento concluído!</p>`;
        } catch (error) { 
            statusDiv.innerHTML = `<p class="text-danger"><strong>Ocorreu um erro:</strong></p><pre>${error.stack}</pre>`; 
            console.error(error); 
        } finally {
            processarBtn.disabled = false;
        }
    }, 50);
}


function isOverdue(predat) {
    if (!predat) return false;
    const date = predat instanceof Date ? predat : new Date(predat);
    if (isNaN(date)) return false; 
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return date < today;
}

/**
 * FUNÇÃO DE EXPORTAÇÃO MODIFICADA
 * Agora inclui pedidos bloqueados com "C" que estão atrasados.
 */
function exportarPedidosAtrasados() {
    if (pedidosGeraisAtuais.length === 0) {
        alert("Por favor, processe a planilha primeiro.");
        return;
    }

    const pedidosAtrasados = pedidosGeraisAtuais.filter(p => 
        isOverdue(p.Predat) && 
        (!String(p['BLOQ.']).trim() || String(p['BLOQ.']).trim().toUpperCase() === 'C')
    );

    if (pedidosAtrasados.length === 0) {
        alert("Nenhum pedido em atraso (ou bloqueado com 'C') foi encontrado.");
        return;
    }

    alert(`${pedidosAtrasados.length} pedidos em atraso serão exportados.`);

    const header = [ 'Cliente', 'Nome_Cliente', 'Cidade', 'UF', 'Num_Pedido', 'Quilos_Saldo', 'Predat', 'Dat_Ped', 'Coluna5', 'BLOQ.' ];
    
    const dataToExport = pedidosAtrasados.map(p => {
        const newP = {...p};
        if (newP.Predat instanceof Date) { newP.Predat = newP.Predat.toLocaleDateString('pt-BR', { timeZone: 'UTC' }); }
        if (newP.Dat_Ped instanceof Date) { newP.Dat_Ped = newP.Dat_Ped.toLocaleDateString('pt-BR', { timeZone: 'UTC' }); }
        
        let filteredP = {};
        header.forEach(col => { filteredP[col] = newP[col]; });
        return filteredP;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport, { header: header });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Pedidos Atrasados");
    XLSX.writeFile(workbook, "pedidos_atrasados.xlsx");
}

function createTable(pedidos, columnsToDisplay, sourceId = '') {
    if (!pedidos || pedidos.length === 0) return '';
    const colunasExibir = columnsToDisplay || ['Cod_Rota', 'Cliente', 'Nome_Cliente', 'Agendamento', 'Num_Pedido', 'Quilos_Saldo', 'Cubagem', 'Cidade', 'Predat', 'BLOQ.', 'Coluna4', 'Coluna5', 'CF'];
    let table = '<div class="table-responsive"><table class="table table-sm table-bordered table-striped table-hover"><thead><tr>';
    colunasExibir.forEach(c => table += `<th>${c.replace('_', ' ')}</th>`);
    table += '</tr></thead><tbody>';
    pedidos.forEach(p => {
        const isPriorityRow = pedidosPrioritarios.includes(String(p.Num_Pedido));
        const clienteIdNormalizado = normalizeClientId(p.Cliente);
        table += `<tr id="pedido-${p.Num_Pedido}" 
                        class="${isPriorityRow ? 'table-primary' : ''}" 
                        data-cliente-id="${clienteIdNormalizado}" 
                        data-pedido-id="${p.Num_Pedido}"
                        onclick="highlightClientRows(event)"
                        draggable="true"
                        ondragstart="dragStart(event, '${p.Num_Pedido}', '${clienteIdNormalizado}', '${sourceId}')">`;
        colunasExibir.forEach(c => {
            let cellContent = p[c] === undefined || p[c] === null ? '' : p[c];
            if (c === 'Num_Pedido') {
                const isPriority = pedidosPrioritarios.includes(String(p.Num_Pedido));
                const isSemCorte = pedidosSemCorte.has(String(p.Num_Pedido));
                const priorityBadge = isPriority ? ' <span class="badge bg-primary">Prioridade</span>' : '';
                const semCorteBadge = isSemCorte ? ' <span class="badge bg-transparent" title="Pedido Sem Corte"><i class="bi bi-scissors text-warning"></i></span>' : '';
                table += `<td>${cellContent}${priorityBadge}${semCorteBadge}</td>`;
            } else if (c === 'Agendamento' && cellContent === 'Sim') {
                table += `<td><span class="badge bg-warning text-dark">${cellContent}</span></td>`;
            } else if (c === 'Predat' || c === 'Dat_Ped') {
                const dateObj = cellContent instanceof Date ? cellContent : new Date(cellContent);
                let formattedDate = (dateObj instanceof Date && !isNaN(dateObj)) ? dateObj.toLocaleDateString('pt-BR', { timeZone: 'UTC' }) : (cellContent || '');
                
                if (c === 'Predat' && isOverdue(p.Predat)) {
                    table += `<td><span class="text-danger fw-bold">${formattedDate}</span></td>`;
                } else {
                    table += `<td>${formattedDate}</td>`;
                }
            } else if (c === 'Quilos_Saldo' || c === 'Cubagem') {
                table += `<td>${(cellContent || 0).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>`;
            }
            else {
                table += `<td>${cellContent}</td>`;
            }
        });
        table += '</tr>';
    });
    table += '</tbody></table></div>';
    return table;
}

function displayGerais(div, grupos) {
    if (Object.keys(grupos).length === 0) { 
        div.innerHTML = '<div class="empty-state"><i class="bi bi-file-earmark-excel"></i><p>Nenhum pedido de varejo disponível.</p></div>'; 
        document.getElementById('botoes-fiorino').innerHTML = '<div class="empty-state"><i class="bi bi-box"></i><p>Nenhuma rota de Fiorino disponível.</p></div>';
        document.getElementById('botoes-van').innerHTML = '<div class="empty-state"><i class="bi bi-truck-front-fill"></i><p>Nenhuma rota de Van disponível.</p></div>';
        document.getElementById('botoes-34').innerHTML = '<div class="empty-state"><i class="bi bi-truck-flatbed"></i><p>Nenhuma rota de 3/4 disponível.</p></div>';
        return; 
    }
    const rotasDisponiveis = new Set(Object.keys(grupos));
    const botoes = { fiorino: '', van: '', tresQuartos: '' };
    const addedButtons = new Set();
    rotasDisponiveis.forEach(rota => {
        const config = rotaVeiculoMap[rota];
        if (config && !addedButtons.has(rota)) {
            let rotaValue = `'${rota}'`;
            if (config.combined) {
                const combinedRoutes = [rota, ...config.combined];
                rotaValue = `[${combinedRoutes.map(r => `'${r}'`).join(', ')}]`;
                combinedRoutes.forEach(r => addedButtons.add(r));
            }
            const vehicleType = config.type;
            const colorClass = vehicleType === 'fiorino' ? 'success' : (vehicleType === 'van' ? 'primary' : 'warning');
            const functionCall = `separarCargasGeneric(${rotaValue}, 'resultado-${vehicleType === 'tresQuartos' ? '34' : vehicleType}-geral', '${config.title}', '${vehicleType}', this)`;
            botoes[vehicleType] += `<button id="btn-${vehicleType}-${rota}" class="btn btn-outline-${colorClass} mt-2 me-2" onclick="${functionCall}">${config.title}</button>`;
        }
    });
    document.getElementById('botoes-fiorino').innerHTML = botoes.fiorino || '<div class="empty-state"><i class="bi bi-box"></i><p>Nenhuma rota de Fiorino encontrada.</p></div>';
    document.getElementById('botoes-van').innerHTML = botoes.van || '<div class="empty-state"><i class="bi bi-truck-front-fill"></i><p>Nenhuma rota de Van encontrada.</p></div>';
    document.getElementById('botoes-34').innerHTML = botoes.tresQuartos || '<div class="empty-state"><i class="bi bi-truck-flatbed"></i><p>Nenhuma rota de 3/4 encontrada.</p></div>';
    let accordionHtml = '<div class="accordion accordion-flush" id="accordionGeral">';
    Object.keys(grupos).sort().forEach((rota, index) => {
        const grupo = grupos[rota];
        const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const veiculo = rotaVeiculoMap[rota]?.type || 'N/D';
        const veiculoNome = veiculo.replace('tresQuartos', '3/4').replace(/^\w/, c => c.toUpperCase());
        accordionHtml += `<div class="accordion-item"><h2 class="accordion-header" id="headingGeral${index}"><button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseGeral${index}"><strong>Rota: ${rota} (${veiculoNome})</strong> &nbsp; <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${grupo.pedidos.length}</span> <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKgFormatado} kg</span></button></h2><div id="collapseGeral${index}" class="accordion-collapse collapse" data-bs-parent="#accordionGeral"><div class="accordion-body">${createTable(grupo.pedidos, null, 'geral')}</div></div></div>`;
    });
    accordionHtml += '</div>'; div.innerHTML = accordionHtml;
}

function displayPedidosBloqueados(div, pedidos) {
    if (pedidos.length === 0) {
        div.innerHTML = '<div class="empty-state"><i class="bi bi-shield-check"></i><p>Nenhum pedido bloqueado manualmente.</p></div>';
        return;
    }
    const totalKg = pedidos.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
    const totalKgFormatado = totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    let html = `<div class="alert alert-danger"><strong>Total Bloqueado:</strong> ${pedidos.length} pedidos / ${totalKgFormatado} kg</div>`;
    html += createTable(pedidos);
    div.innerHTML = html;
}

function displayPedidosCFNumerico(div, pedidos) {
    if (pedidos.length === 0) { div.innerHTML = '<div class="empty-state"><i class="bi bi-funnel"></i><p>Nenhum pedido de varejo filtrado por regra.</p></div>'; return; }
    const grupos = pedidos.reduce((acc, p) => {
        const rota = p.Cod_Rota; if (!acc[rota]) { acc[rota] = { pedidos: [], totalKg: 0 }; } acc[rota].pedidos.push(p); acc[rota].totalKg += p.Quilos_Saldo; return acc;
    }, {});
    let accordionHtml = '<div class="accordion accordion-flush" id="accordionCF">';
    Object.keys(grupos).sort().forEach((rota, index) => {
        const grupo = grupos[rota];
        const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        accordionHtml += `<div class="accordion-item"><h2 class="accordion-header" id="headingCF${index}"><button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseCF${index}"><strong>Rota: ${rota}</strong> &nbsp; <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${grupo.pedidos.length}</span> <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKgFormatado} kg</span></button></h2><div id="collapseCF${index}" class="accordion-collapse collapse" data-bs-parent="#accordionCF"><div class="accordion-body">${createTable(pedidos)}</div></div></div>`;
    });
    accordionHtml += '</div>'; div.innerHTML = accordionHtml;
}

function displayRota1(div, pedidos) {
    if (!pedidos || pedidos.length === 0) {
        div.innerHTML = '<div class="empty-state"><i class="bi bi-check-circle"></i><p>Nenhum pedido da Rota 1 para alteração.</p></div>';
        return;
    }
    const html = `
        <div class="d-flex justify-content-end mb-2 no-print">
            <button class="btn btn-sm btn-outline-warning" onclick="imprimirGeneric('resultado-rota1', 'Pedidos Rota 1 para Alteração')">
                <i class="bi bi-printer-fill me-1"></i>Imprimir Lista
            </button>
        </div>
        ${createTable(pedidos, ['Num_Pedido', 'Cliente', 'Nome_Cliente', 'Quilos_Saldo', 'Cidade', 'Predat', 'CF', 'Coluna5'])}
    `;
    div.innerHTML = html;
}

function createPrintWindow(title) {
    const printWindow = window.open('', '', 'height=800,width=1200');
    printWindow.document.write('<html><head><title>' + title + '</title>');
    printWindow.document.write('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">');
    printWindow.document.write(`<style>body { margin: 20px; -webkit-print-color-adjust: exact; print-color-adjust: exact; } .card { break-inside: avoid; margin-bottom: 1rem; page-break-inside: avoid; } .no-print { display: none !important; } .bg-success { background-color: #198754 !important; color: white !important; } .bg-primary { background-color: #0d6efd !important; color: white !important; } .bg-warning { background-color: #ffc107 !important; color: black !important; } .bg-danger { background-color: #dc3545 !important; color: white !important; } .table-responsive { overflow: visible !important; } table, th, td { border: 1px solid #dee2e6 !important; } .table-primary, .table-primary > th, .table-primary > td { --bs-table-bg: #cfe2ff !important; color: #000 !important; } h1, h2, h3, h4, h5 { margin-top: 1rem; margin-bottom: 0.5rem; }</style></head><body>`);
    return printWindow;
}

function imprimirGeneric(divId, title) {
    const divToPrint = document.getElementById(divId);
    if (!divToPrint) return;

    const printWindow = createPrintWindow(title);
    const contentToPrint = divToPrint.cloneNode(true);
    contentToPrint.querySelectorAll('.no-print').forEach(btn => btn.remove());

    printWindow.document.body.innerHTML = `<h3>${title}</h3>` + contentToPrint.innerHTML;
    printWindow.document.close();
    printWindow.focus(); 
    setTimeout(() => { 
        printWindow.print(); 
        printWindow.close(); 
    }, 500);
}

function imprimirSobras(title) {
    if (currentLeftoversForPrinting.length === 0) { alert("Nenhuma sobra para imprimir."); return; }
    const totalKgFormatado = currentLeftoversForPrinting.reduce((sum, p) => sum + p.Quilos_Saldo, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const printWindow = createPrintWindow(title);
    let contentToPrint = `<h3>${title} - Total: ${totalKgFormatado} kg</h3>` + createTable(currentLeftoversForPrinting);
    printWindow.document.body.innerHTML = contentToPrint;
    printWindow.document.close();
    printWindow.focus(); 
    setTimeout(() => { printWindow.print(); printWindow.close(); }, 500);
}

function imprimirCargasGeneric(divId, title) {
    const divToPrint = document.getElementById(divId);
    if (!divToPrint) return;
    
    const completedLoadsCards = divToPrint.querySelectorAll('.card .card-header:not(.bg-danger)');
    
    if (completedLoadsCards.length === 0) { alert(`Nenhuma carga concluída para imprimir na seção "${title}".`); return; }
    
    const printWindow = createPrintWindow('Imprimir: ' + title);
    let contentToPrint = `<h3>${title}</h3>`;
    completedLoadsCards.forEach(header => { contentToPrint += header.closest('.card.mb-3').outerHTML; });
    printWindow.document.body.innerHTML = contentToPrint;
    printWindow.document.close();
    printWindow.focus(); 
    setTimeout(() => { 
        printWindow.print(); 
        printWindow.close(); 
    }, 500);
}

function imprimirTocoIndividual(cf) {
    if (!gruposToco[cf]) { alert(`Nenhuma carga Toco encontrada para o CF: ${cf}`); return; }
    const grupo = gruposToco[cf];
    const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const totalCubagemFormatado = grupo.totalCubagem.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const printWindow = createPrintWindow('Imprimir Carga Toco CF: ' + cf);
    let contentToPrint = `<h3>Carga Toco CF: ${cf} - Total: ${totalKgFormatado} kg / ${totalCubagemFormatado} m³</h3>` + createTable(grupo.pedidos);
    printWindow.document.body.innerHTML = contentToPrint;
    printWindow.document.close();
    printWindow.focus(); 
    setTimeout(() => { printWindow.print(); printWindow.close(); }, 500);
}

function imprimirCargaManualIndividual(loadId) {
    const cardToPrint = document.getElementById(loadId);
    if (!activeLoads[loadId] || !cardToPrint) {
        alert(`Erro: Carga com ID ${loadId} não encontrada para impressão.`);
        return;
    }
    const title = `Impressão - Carga Manual ${activeLoads[loadId].numero}`;
    const printWindow = createPrintWindow(title);
    
    printWindow.document.body.innerHTML = `<h3>${title}</h3>` + cardToPrint.outerHTML;
    printWindow.document.close();
    printWindow.focus(); 
    setTimeout(() => { 
        printWindow.print(); 
        printWindow.close(); 
    }, 500);
}

function getVehicleConfig(vehicleType) {
    return {
        minKg: parseFloat(document.getElementById(`${vehicleType}MinCapacity`).value),
        softMaxKg: parseFloat(document.getElementById(`${vehicleType}MaxCapacity`).value),
        softMaxCubage: parseFloat(document.getElementById(`${vehicleType}Cubage`).value),
        hardMaxKg: parseFloat(document.getElementById(`${vehicleType}HardMaxCapacity`)?.value || document.getElementById(`${vehicleType}MaxCapacity`).value),
        hardMaxCubage: parseFloat(document.getElementById(`${vehicleType}HardCubage`)?.value || document.getElementById(`${vehicleType}Cubage`).value),
    };
}

function isMoveValid(load, groupToAdd, vehicleType) {
    const config = getVehicleConfig(vehicleType);
    if ((load.totalKg + groupToAdd.totalKg) > config.hardMaxKg) return false;
    if ((load.totalCubagem + groupToAdd.totalCubagem) > config.hardMaxCubage) return false;
    if (groupToAdd.isSpecial) {
        const specialClientIdsInLoad = new Set(load.pedidos.filter(isSpecialClient).map(p => normalizeClientId(p.Cliente)));
        const groupToAddClientId = normalizeClientId(groupToAdd.pedidos[0].Cliente);
        if (!specialClientIdsInLoad.has(groupToAddClientId) && specialClientIdsInLoad.size >= 2) {
            return false;
        }
    }
    if (groupToAdd.pedidos.some(p => p.Agendamento === 'Sim') && load.pedidos.some(p => p.Agendamento === 'Sim')) return false;
    return true;
}

function runHeuristicOptimization(packableGroups, vehicleType) {
    const strategies = [
        { name: 'priority-weight-desc', sorter: (a, b) => (b.pedidos.some(p => pedidosPrioritarios.includes(String(p.Num_Pedido))) - a.pedidos.some(p => pedidosPrioritarios.includes(String(p.Num_Pedido)))) || (b.totalKg - a.totalKg) },
        { name: 'scheduled-weight-desc', sorter: (a, b) => (b.pedidos.some(p => p.Agendamento === 'Sim') - a.pedidos.some(p => p.Agendamento === 'Sim')) || (b.totalKg - a.totalKg) },
        { name: 'weight-desc', sorter: (a, b) => b.totalKg - a.totalKg },
        { name: 'weight-asc', sorter: (a, b) => a.totalKg - b.totalKg }
    ];
    let bestResult = null;
    for (const strategy of strategies) {
        const sortedGroups = [...packableGroups].sort(strategy.sorter);
        const result = createSolutionFromHeuristic(sortedGroups, vehicleType);
        const leftoverWeight = result.leftovers.reduce((sum, g) => sum + g.totalKg, 0);
        if (bestResult === null || leftoverWeight < bestResult.leftoverWeight) {
            bestResult = { ...result, leftoverWeight: leftoverWeight, strategy: strategy.name };
        }
    }
    console.log(`Nível 1: Melhor estratégia para ${vehicleType} foi ${bestResult.strategy} com ${bestResult.leftoverWeight.toFixed(2)}kg de sobra.`);
    return bestResult;
}

function createSolutionFromHeuristic(itemsParaEmpacotar, vehicleType) {
    const config = getVehicleConfig(vehicleType);
    let loads = [];
    let leftoverItems = [];
    itemsParaEmpacotar.forEach(item => {
        if (item.totalKg > config.hardMaxKg || item.totalCubagem > config.hardMaxCubage) {
            leftoverItems.push(item); return;
        }
        
        let bestFit = null;
        for (const load of loads) {
            if (isMoveValid(load, item, vehicleType)) {
                const remainingCapacity = config.hardMaxKg - (load.totalKg + item.totalKg);
                if (bestFit === null || remainingCapacity < bestFit.remainingCapacity) {
                    bestFit = { load: load, remainingCapacity: remainingCapacity };
                }
            }
        }
        if (bestFit) {
            bestFit.load.pedidos.push(...item.pedidos);
            bestFit.load.totalKg += item.totalKg;
            bestFit.load.totalCubagem += item.totalCubagem;
            bestFit.load.usedHardLimit = bestFit.load.totalKg > config.softMaxKg || bestFit.load.totalCubagem > config.softMaxCubage;
        } else {
            loads.push({
                pedidos: [...item.pedidos],
                totalKg: item.totalKg,
                totalCubagem: item.totalCubagem,
                isSpecial: item.isSpecial,
                usedHardLimit: (item.totalKg > config.softMaxKg || item.totalCubagem > config.softMaxCubage)
            });
        }
    });
    
    let finalLoads = [];
    let unplacedGroups = [];
    loads.forEach(load => {
        const hasPriority = load.pedidos.some(p => pedidosPrioritarios.includes(String(p.Num_Pedido)));
        if (load.pedidos.length > 0 && (load.totalKg >= config.minKg || (hasPriority && vehicleType !== 'tresQuartos'))) {
            finalLoads.push(load);
        } else if (load.pedidos.length > 0) {
            const clientGroupsInFailedLoad = Object.values(load.pedidos.reduce((acc, p) => {
                const clienteId = normalizeClientId(p.Cliente);
                if (!acc[clienteId]) { acc[clienteId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(p) }; }
                acc[clienteId].pedidos.push(p);
                acc[clienteId].totalKg += p.Quilos_Saldo;
                acc[clienteId].totalCubagem += p.Cubagem;
                return acc;
            }, {}));
            unplacedGroups.push(...clientGroupsInFailedLoad);
        }
    });
    
    return { loads: finalLoads, leftovers: [...leftoverItems, ...unplacedGroups] };
}

function getSolutionEnergy(solution, vehicleType) {
    const config = getVehicleConfig(vehicleType);
    const leftoverWeight = solution.leftovers.reduce((sum, group) => sum + group.totalKg, 0);
    const loadPenalty = solution.loads.reduce((sum, load) => {
        if (load.totalKg > 0 && load.totalKg < config.minKg) {
            return sum + 1000 + (config.minKg - load.totalKg);
        }
        return sum;
    }, 0);
    return leftoverWeight + loadPenalty;
}

function calculateDisplaySobras(solution, vehicleType) {
    const config = getVehicleConfig(vehicleType);
    let totalSobras = solution.leftovers.reduce((sum, group) => sum + group.totalKg, 0);
    solution.loads.forEach(load => {
        if (load.totalKg > 0 && load.totalKg < config.minKg) {
            totalSobras += load.totalKg;
        }
    });
    return totalSobras;
}

function runSimulatedAnnealing(packableGroups, vehicleType, uiText) {
     return new Promise(resolve => {
            const initialTemp = 1000;
            const coolingRate = 0.993;
            const iterationsPerTemp = 200;

            const initialSortedGroups = [...packableGroups].sort((a, b) => b.totalKg - a.totalKg);
            let bestSolution = createSolutionFromHeuristic(initialSortedGroups, vehicleType);
            let currentSolution = deepClone(bestSolution);
            
            let currentTemp = initialTemp;
            
            const progressBar = document.getElementById('processing-progress-bar');
            const statusText = document.getElementById('processing-status-text');
            const detailsText = document.getElementById('processing-details-text');

            function doTemperatureStep() {
                for (let i = 0; i < iterationsPerTemp; i++) {
                    let neighborSolution = deepClone(currentSolution);
                    let moveMade = false;
                    const moveType = Math.random();
                    if (moveType < 0.7 && neighborSolution.leftovers.length > 0) {
                        const leftoverIndex = Math.floor(Math.random() * neighborSolution.leftovers.length);
                        const groupToPlace = neighborSolution.leftovers[leftoverIndex];
                        const targetLoadIndex = neighborSolution.loads.length > 0 ? Math.floor(Math.random() * (neighborSolution.loads.length + 1)) : 0;

                        if (targetLoadIndex < neighborSolution.loads.length) {
                            const targetLoad = neighborSolution.loads[targetLoadIndex];
                            if(isMoveValid(targetLoad, groupToPlace, vehicleType)) {
                                targetLoad.pedidos.push(...groupToPlace.pedidos);
                                targetLoad.totalKg += groupToPlace.totalKg;
                                targetLoad.totalCubagem += groupToPlace.totalCubagem;
                                neighborSolution.leftovers.splice(leftoverIndex, 1);
                                moveMade = true;
                            }
                        } else if (isMoveValid({pedidos:[], totalKg:0, totalCubagem:0}, groupToPlace, vehicleType)) { 
                            neighborSolution.loads.push(groupToPlace);
                            neighborSolution.leftovers.splice(leftoverIndex, 1);
                            moveMade = true;
                        }
                    } else if (neighborSolution.loads.length > 0) {
                        const loadIndex = Math.floor(Math.random() * neighborSolution.loads.length);
                        const load = neighborSolution.loads[loadIndex];
                        
                        if (load.pedidos.length > 0) {
                            const clientGroupsInLoad = Object.values(load.pedidos.reduce((acc, p) => {
                                const cId = normalizeClientId(p.Cliente);
                                if (!acc[cId]) acc[cId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(p) };
                                acc[cId].pedidos.push(p); acc[cId].totalKg += p.Quilos_Saldo; acc[cId].totalCubagem += p.Cubagem;
                                return acc;
                            }, {}));

                            if(clientGroupsInLoad.length > 0) {
                                const groupIndexToRemove = Math.floor(Math.random() * clientGroupsInLoad.length);
                                const groupToMove = clientGroupsInLoad[groupIndexToRemove];
                                
                                const idsToRemove = new Set(groupToMove.pedidos.map(p => p.Num_Pedido));
                                load.pedidos = load.pedidos.filter(p => !idsToRemove.has(p.Num_Pedido));
                                load.totalKg -= groupToMove.totalKg;
                                load.totalCubagem -= groupToMove.totalCubagem;
                                if(load.pedidos.length === 0) neighborSolution.loads.splice(loadIndex,1);

                                neighborSolution.leftovers.push(groupToMove);
                                moveMade = true;
                            }
                        }
                    }
                    
                    if(moveMade){
                        const currentEnergy = getSolutionEnergy(currentSolution, vehicleType);
                        const neighborEnergy = getSolutionEnergy(neighborSolution, vehicleType);
                        if (neighborEnergy < currentEnergy || Math.random() < Math.exp((currentEnergy - neighborEnergy) / currentTemp)) {
                            currentSolution = neighborSolution;
                            if (getSolutionEnergy(currentSolution, vehicleType) < getSolutionEnergy(bestSolution, vehicleType)) {
                                bestSolution = deepClone(currentSolution);
                            }
                        }
                    }
                }

                currentTemp *= coolingRate;
                const progress = Math.min(100, 100 * (1 - Math.log(currentTemp) / Math.log(initialTemp)));
                progressBar.style.width = `${progress}%`;
                statusText.textContent = uiText;
                const displaySobras = calculateDisplaySobras(bestSolution, vehicleType);
                detailsText.textContent = `Sobra Atual: ${displaySobras.toFixed(2)} kg`;

                if (currentTemp > 1) {
                    requestAnimationFrame(doTemperatureStep);
                } else {
                    const config = getVehicleConfig(vehicleType);
                    let finalLoads = [];
                    let finalLeftovers = [...bestSolution.leftovers];
                    bestSolution.loads.forEach(load => {
                        if (load.pedidos.length > 0 && load.totalKg >= config.minKg) {
                            finalLoads.push(load);
                        } else if (load.pedidos.length > 0) {
                            const groups = Object.values(load.pedidos.reduce((acc, p) => { 
                                const cId = normalizeClientId(p.Cliente);
                                if (!acc[cId]) acc[cId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(p) };
                                acc[cId].pedidos.push(p); acc[cId].totalKg += p.Quilos_Saldo; acc[cId].totalCubagem += p.Cubagem;
                                return acc; 
                            }, {}));
                            finalLeftovers.push(...groups);
                        }
                    });
                    console.log(`Otimização para ${vehicleType} finalizada com ${finalLeftovers.reduce((s, g) => s + g.totalKg, 0).toFixed(2)}kg de sobra.`);
                    resolve({ loads: finalLoads, leftovers: finalLeftovers });
                }
            }
            
            requestAnimationFrame(doTemperatureStep);
       });
}

function refinarCargasComTrocas(initialLoads, initialLeftovers, vehicleType) {
    let loads = deepClone(initialLoads);
    let leftovers = deepClone(initialLeftovers);
    let improvementMade;

    if (leftovers.length === 0) {
        return { refinedLoads: loads, remainingLeftovers: leftovers };
    }
    
    let attempts = 0;
    const maxAttempts = (leftovers.length * loads.length) || 1; 

    do {
        improvementMade = false;
        attempts++;
        for (let i = 0; i < leftovers.length; i++) {
            const leftoverGroup = leftovers[i];
            for (let j = 0; j < loads.length; j++) {
                const load = loads[j];
                const clientGroupsInLoad = Object.values(load.pedidos.reduce((acc, pedido) => {
                    const clienteId = normalizeClientId(pedido.Cliente);
                    if (!acc[clienteId]) acc[clienteId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(pedido) };
                    acc[clienteId].pedidos.push(pedido);
                    acc[clienteId].totalKg += pedido.Quilos_Saldo;
                    acc[clienteId].totalCubagem += pedido.Cubagem;
                    return acc;
                }, {}));
                for (let k = 0; k < clientGroupsInLoad.length; k++) {
                    const groupToSwapOut = clientGroupsInLoad[k];
                    
                    const tempLoadAfterRemoval = {
                        pedidos: load.pedidos.filter(p => !groupToSwapOut.pedidos.some(gp => gp.Num_Pedido === p.Num_Pedido)),
                        totalKg: load.totalKg - groupToSwapOut.totalKg,
                        totalCubagem: load.totalCubagem - groupToSwapOut.totalCubagem,
                    };
                    
                    if (isMoveValid(tempLoadAfterRemoval, leftoverGroup, vehicleType)) {
                        const idsToRemove = new Set(groupToSwapOut.pedidos.map(p => p.Num_Pedido));
                        load.pedidos = load.pedidos.filter(p => !idsToRemove.has(p.Num_Pedido));
                        load.totalKg -= groupToSwapOut.totalKg;
                        load.totalCubagem -= groupToSwapOut.totalCubagem;
                        load.pedidos.push(...leftoverGroup.pedidos);
                        load.totalKg += leftoverGroup.totalKg;
                        load.totalCubagem += leftoverGroup.totalCubagem;
                        leftovers.splice(i, 1); 
                        leftovers.push(groupToSwapOut);
                        console.log(`POLIMENTO (Nível 3): Encaixou grupo de ${leftoverGroup.totalKg.toFixed(2)}kg trocando por um de ${groupToSwapOut.totalKg.toFixed(2)}kg.`);
                        improvementMade = true;
                        break;
                    }
                }
                if (improvementMade) break;
            }
            if (improvementMade) break;
        }
    } while (improvementMade && leftovers.length > 0 && attempts < maxAttempts);

    if(attempts >= maxAttempts) console.warn("Polimento (Nível 3) interrompido para evitar loop infinito.");
    return { refinedLoads: loads, remainingLeftovers: leftovers };
}

async function separarCargasGeneric(routeOrRoutes, divId, title, vehicleType, buttonElement) {
    const allRouteButtons = document.querySelectorAll('#botoes-fiorino button, #botoes-van button, #botoes-34 button');
    allRouteButtons.forEach(btn => btn.disabled = true);

    const resultadoDiv = document.getElementById(divId);
    
    const autoGeneratedContent = resultadoDiv.querySelector('.resultado-container');
    if (autoGeneratedContent) { autoGeneratedContent.remove(); }

    if (buttonElement) {
        currentLeftoversForPrinting = [];
        const parentContainer = buttonElement.parentElement;
        parentContainer.querySelectorAll('button').forEach(btn => {
            btn.classList.remove('active', 'btn-success', 'btn-primary', 'btn-warning');
            const colorClass = btn.id.includes('fiorino') ? 'success' : (btn.id.includes('van') ? 'primary' : 'warning');
            btn.classList.add(`btn-outline-${colorClass}`);
            btn.innerHTML = btn.innerHTML.replace('<i class="bi bi-check-circle-fill me-2"></i>', '');
        });
        const colorClass = vehicleType === 'fiorino' ? 'success' : (vehicleType === 'van' ? 'primary' : 'warning');
        buttonElement.classList.remove(`btn-outline-${colorClass}`);
        buttonElement.classList.add(`btn-${colorClass}`, 'active');
        buttonElement.innerHTML = `<i class="bi bi-check-circle-fill me-2"></i>${title}`;
    }

    if (planilhaData.length === 0) {
        resultadoDiv.innerHTML = '<p class="text-danger">Nenhum dado de planilha carregado.</p>'; 
        allRouteButtons.forEach(btn => btn.disabled = false);
        return;
    }

    const routes = Array.isArray(routeOrRoutes) ? routeOrRoutes : [String(routeOrRoutes)];
    let pedidosRota = pedidosGeraisAtuais.filter(p => routes.includes(String(p.Cod_Rota)));

    const clientGroupsMap = pedidosRota.reduce((acc, pedido) => {
        const clienteId = normalizeClientId(pedido.Cliente);
        if (!acc[clienteId]) { acc[clienteId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(pedido) }; }
        acc[clienteId].pedidos.push(pedido);
        acc[clienteId].totalKg += pedido.Quilos_Saldo;
        acc[clienteId].totalCubagem += pedido.Cubagem;
        return acc;
    }, {});
    const packableGroups = Object.values(clientGroupsMap);

    const optimizationLevel = document.getElementById('optimizationLevelSelect').value;
    let optimizationResult;
    
    const modal = new bootstrap.Modal(document.getElementById('processing-modal'));
    
    try {
        if(optimizationLevel !== '1') {
            document.getElementById('processing-progress-bar').style.width = '0%';
            modal.show();
        } else {
            resultadoDiv.insertAdjacentHTML('beforeend', '<div id="spinner-temp-container" class="d-flex align-items-center justify-content-center p-5"><div class="spinner-border text-primary" role="status"></div><span class="ms-3">Analisando...</span></div>');
        }

        switch (optimizationLevel) {
            case '1':
                optimizationResult = runHeuristicOptimization(packableGroups, vehicleType);
                break;
            case '3':
                const saResultForPolish = await runSimulatedAnnealing(packableGroups, vehicleType, 'Otimizando... (Nível 3 - Fase 1/2)');
                document.getElementById('processing-status-text').textContent = 'Otimizando... (Nível 3 - Fase 2/2)';
                document.getElementById('processing-details-text').textContent = 'Aplicando polimento com trocas locais.';
                const polished = refinarCargasComTrocas(saResultForPolish.loads, saResultForPolish.leftovers, vehicleType);
                optimizationResult = { loads: polished.refinedLoads, leftovers: polished.remainingLeftovers };
                break;
            case '4':
                const saResultForReconstruction = await runSimulatedAnnealing(packableGroups, vehicleType, 'Otimizando... (Nível 4 - Fase 1/3)');
                document.getElementById('processing-status-text').textContent = 'Otimizando... (Nível 4 - Fase 2/3)';
                document.getElementById('processing-details-text').textContent = 'Reconstruindo cargas ineficientes.';
                const reconstructed = await refinarComReconstrucao(saResultForReconstruction.loads, saResultForReconstruction.leftovers, vehicleType);
                document.getElementById('processing-status-text').textContent = 'Otimizando... (Nível 4 - Fase 3/3)';
                document.getElementById('processing-details-text').textContent = 'Aplicando polimento final.';
                const finalPolished = refinarCargasComTrocas(reconstructed.refinedLoads, reconstructed.remainingLeftovers, vehicleType);
                optimizationResult = { loads: finalPolished.refinedLoads, leftovers: finalPolished.remainingLeftovers };
                break;
            case '2':
            default:
                optimizationResult = await runSimulatedAnnealing(packableGroups, vehicleType, 'Otimizando... (Nível 2)');
                break;
        }
    } finally {
        allRouteButtons.forEach(btn => btn.disabled = false);
        if(optimizationLevel !== '1') {
            modal.hide();
        } else {
            document.getElementById('spinner-temp-container')?.remove();
        }
    }

    const { refinedLoads, remainingLeftovers } = refineLoadsWithSimpleFit(optimizationResult.loads, optimizationResult.leftovers);
    let finalValidLoads = [];
    let finalLeftoverGroups = [...remainingLeftovers];
    refinedLoads.forEach(load => {
        const config = getVehicleConfig(vehicleType);
        const hasPriority = load.pedidos.some(p => pedidosPrioritarios.includes(String(p.Num_Pedido)));
        if (load.totalKg >= config.minKg || (hasPriority && vehicleType !== 'tresQuartos')) {
            finalValidLoads.push(load);
        } else {
            finalLeftoverGroups.push(...Object.values(load.pedidos.reduce((acc, p) => {
                const cId = normalizeClientId(p.Cliente);
                if (!acc[cId]) acc[cId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(p) };
                acc[cId].pedidos.push(p); acc[cId].totalKg += p.Quilos_Saldo; acc[cId].totalCubagem += p.Cubagem;
                return acc;
            }, {})));
        }
    });

    currentLeftoversForPrinting = finalLeftoverGroups.flatMap(group => group.pedidos);

    finalValidLoads.forEach((load, index) => {
        load.numero = `${vehicleType.charAt(0).toUpperCase()}${index + 1}`;
        const loadId = `${vehicleType}-${Date.now()}-${index}`;
        load.id = loadId;
        load.vehicleType = vehicleType;
        activeLoads[loadId] = load;
    });

    const vehicleInfo = {
        fiorino: { name: 'Fiorino', colorClass: 'bg-success', textColor: 'text-white', icon: 'bi-box-seam-fill' },
        van: { name: 'Van', colorClass: 'bg-primary', textColor: 'text-white', icon: 'bi-truck-front-fill' },
        tresQuartos: { name: '3/4', colorClass: 'bg-warning', textColor: 'text-dark', icon: 'bi-truck-flatbed' }
    };
    
    let html = `<h5 class="mt-3">Cargas para <strong>${title}</strong></h5>`;
    
    if(finalValidLoads.length === 0){
        html += `<div class="alert alert-secondary">Nenhuma carga foi formada para esta rota.</div>`;
    } else {
        finalValidLoads.forEach(load => {
            html += renderLoadCard(load, vehicleType, vehicleInfo[vehicleType]);
        });
    }
    
    if (currentLeftoversForPrinting.length > 0) {
        const finalLeftoverKg = currentLeftoversForPrinting.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
        html += `<div id="leftovers-card-${divId}" class="drop-zone-card" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event)" data-load-id="leftovers" data-vehicle-type="leftovers">
                     <h5 class="mt-4">Sobras Finais: ${finalLeftoverKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} kg</h5>
                     <div class="card mb-3">
                         <div class="card-header bg-danger text-white d-flex align-items-center">
                             Pedidos Restantes
                             <div class="ms-auto">
                                <button id="start-manual-load-btn" class="btn btn-success ms-auto no-print" onclick="startManualLoadBuilder()"><i class="bi bi-plus-circle-fill me-1"></i>Criar Carga Manual</button>
                                <button class="btn btn-info ms-2 no-print" onclick="imprimirSobras('Sobras Finais de ${title}')"><i class="bi bi-printer-fill me-1"></i>Imprimir</button>
                             </div>
                         </div>
                         <div class="card-body">${createTable(currentLeftoversForPrinting, ['Num_Pedido', 'Quilos_Saldo', 'Agendamento', 'Cubagem', 'Predat', 'Cliente', 'Nome_Cliente', 'Cidade', 'CF'], 'leftovers')}</div>
                     </div>
                </div>`;
    }
    resultadoDiv.innerHTML = `<div class="resultado-container">${html}</div>`;
    updateAndRenderChart();
}

function refineLoadsWithSimpleFit(initialLoads, initialLeftovers) {
    let refinedLoads = deepClone(initialLoads);
    let remainingLeftovers = deepClone(initialLeftovers);

    for (let i = remainingLeftovers.length - 1; i >= 0; i--) {
        const leftoverGroup = remainingLeftovers[i];
        
        for (const load of refinedLoads) {
            const vehicleType = load.vehicleType;
            if (!vehicleType) continue; 
            
            if (isMoveValid(load, leftoverGroup, vehicleType)) {
                load.pedidos.push(...leftoverGroup.pedidos);
                load.totalKg += leftoverGroup.totalKg;
                load.totalCubagem += leftoverGroup.totalCubagem;
                remainingLeftovers.splice(i, 1);
                break; 
            }
        }
    }
    return { refinedLoads, remainingLeftovers };
}

async function refinarComReconstrucao(initialLoads, initialLeftovers, vehicleType) {
    let loads = deepClone(initialLoads);
    if (loads.length < 2) {
        console.log("POLIMENTO (Nível 4): Poucas cargas para reconstruir.");
        return { refinedLoads: loads, remainingLeftovers: initialLeftovers };
    }
    
    loads.sort((a, b) => a.totalKg - b.totalKg);
    const worstLoad = loads[0]; 
    const config = getVehicleConfig(vehicleType);
    if (worstLoad.totalKg >= config.softMaxKg) {
        console.log("POLIMENTO (Nível 4): Carga menos cheia já está otimizada.");
        return { refinedLoads: initialLoads, remainingLeftovers: initialLeftovers };
    }
    
    const groupsToReallocate = Object.values(worstLoad.pedidos.reduce((acc, p) => {
        const cId = normalizeClientId(p.Cliente);
        if (!acc[cId]) acc[cId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(p) };
        acc[cId].pedidos.push(p); 
        acc[cId].totalKg += p.Quilos_Saldo; 
        acc[cId].totalCubagem += p.Cubagem;
        return acc;
    }, {}));

    const newLeftovers = [...initialLeftovers, ...groupsToReallocate];
    const remainingLoads = loads.slice(1);
    
    const { refinedLoads: reconstructedLoads, remainingLeftovers: finalLeftovers } = refineLoadsWithSimpleFit(remainingLoads, newLeftovers);
    
    const originalSobras = initialLeftovers.reduce((sum, g) => sum + g.totalKg, 0);
    const newSobras = finalLeftovers.reduce((sum, g) => sum + g.totalKg, 0);

    if (newSobras < originalSobras) {
        console.log(`POLIMENTO (Nível 4): Reconstrução bem-sucedida! Sobra reduzida de ${originalSobras.toFixed(2)}kg para ${newSobras.toFixed(2)}kg.`);
        return { refinedLoads: reconstructedLoads, remainingLeftovers: finalLeftovers };
    } else {
        console.log("POLIMENTO (Nível 4): Reconstrução não melhorou. Revertendo.");
        return { refinedLoads: initialLoads, remainingLeftovers: initialLeftovers };
    }
}

function renderLoadCard(load, vehicleType, vInfo) {
    load.pedidos.sort((a, b) => (String(a.Cliente).localeCompare(String(b.Cliente), undefined, { numeric: true })) || (String(a.Num_Pedido).localeCompare(String(b.Num_Pedido), undefined, { numeric: true })));
    const totalKgFormatado = load.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const totalCubagemFormatado = (load.totalCubagem || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const isPriorityLoad = load.pedidos.some(p => pedidosPrioritarios.includes(String(p.Num_Pedido)));
    const priorityBadge = isPriorityLoad ? '<span class="badge bg-light text-dark ms-3">CARGA COM PRIORIDADE</span>' : '';
    const hardLimitBadge = load.usedHardLimit ? '<span class="badge bg-danger-subtle text-danger-emphasis ms-3"><i class="bi bi-exclamation-triangle-fill"></i> CAPACIDADE EXTRA</span>' : '';

    const config = getVehicleConfig(vehicleType);
    const maxKg = config.hardMaxKg;
    const isOverloaded = maxKg > 0 && load.totalKg > maxKg;
    const pesoPercentual = maxKg > 0 ? (load.totalKg / maxKg) * 100 : 0;
    let progressColor = isOverloaded || pesoPercentual > 100 ? 'bg-danger' : (pesoPercentual > 95 ? 'bg-danger' : (pesoPercentual > 75 ? 'bg-warning' : 'bg-success'));

    const progressBar = `<div class="progress mt-2" role="progressbar" style="height: 10px;"><div class="progress-bar ${progressColor}" style="width: ${Math.min(pesoPercentual, 100)}%"></div></div>`;
    const headerColorClass = isOverloaded ? 'bg-danger' : vInfo.colorClass;
    const printButton = String(load.id).startsWith('manual-') ? `<button class="btn btn-sm btn-outline-info mb-3 no-print" onclick="imprimirCargaManualIndividual('${load.id}')"><i class="bi bi-printer-fill me-1"></i>Imprimir</button>` : '';

    return `<div id="${load.id}" class="card mb-3 drop-zone-card ${isPriorityLoad ? 'border-primary' : ''}" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event)" data-load-id="${load.id}" data-vehicle-type="${vehicleType}"><div class="card-header ${headerColorClass} ${vInfo.textColor}"><i class="bi ${vInfo.icon} me-2"></i>${vInfo.name} #${load.numero} - <i class="bi bi-database me-1"></i>Total: ${totalKgFormatado} kg / <i class="bi bi-rulers me-1"></i>${totalCubagemFormatado} m³ ${priorityBadge} ${hardLimitBadge}</div><div class="card-body">${printButton}${progressBar}${createTable(load.pedidos, null, load.id)}</div></div>`;
}

// Funções de atalho
function separarCargasFiorino(routes, divId, title, buttonElement) { separarCargasGeneric(routes, divId, title, 'fiorino', buttonElement); }
function separarCargasVan(routes, divId, title, buttonElement) { separarCargasGeneric(routes, divId, title, 'van', buttonElement); }
function separarCargas34(routes, divId, title, buttonElement) { separarCargasGeneric(routes, divId, title, 'tresQuartos', buttonElement); }

function displayToco(div, grupos) {
    if (Object.keys(grupos).length === 0) { div.innerHTML = '<div class="empty-state"><i class="bi bi-inboxes-fill"></i><p>Nenhuma carga "Toco" encontrada.</p></div>'; return; }
    
    const maxKg = parseFloat(document.getElementById('tocoMaxCapacity').value);
    let accordionHtml = '<div class="accordion accordion-flush" id="accordionToco">';
    
    Object.keys(grupos).sort().forEach((cf, index) => {
        const grupo = grupos[cf]; 
        const loadId = `toco-${cf}`;
        grupo.id = loadId;
        grupo.vehicleType = 'toco';
        activeLoads[loadId] = grupo;
        grupo.pedidos.sort((a,b) => String(a.Cliente).localeCompare(String(b.Cliente),undefined,{numeric:true})||String(a.Num_Pedido).localeCompare(String(b.Num_Pedido),undefined,{numeric:true}));
        const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const isOverloaded = grupo.totalKg > maxKg;
        const weightBadge = isOverloaded
            ? `<span class="badge bg-danger ms-2"><i class="bi bi-exclamation-triangle-fill me-1"></i>${totalKgFormatado} kg (ACIMA!)</span>`
            : `<span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKgFormatado} kg</span>`;
        
        const pesoPercentual = (grupo.totalKg / maxKg) * 100;
        let progressColor = isOverloaded || pesoPercentual > 100 ? 'bg-danger' : (pesoPercentual > 75 ? 'bg-warning' : 'bg-success');
        const progressBar = `<div class="progress mb-3" role="progressbar" style="height: 10px;"><div class="progress-bar ${progressColor}" style="width: ${Math.min(pesoPercentual, 100)}%"></div></div>`;
        
        accordionHtml += `<div class="accordion-item"><h2 class="accordion-header" id="headingToco${index}"><button class="accordion-button collapsed ${isOverloaded ? 'bg-danger' : ''}" type="button" data-bs-toggle="collapse" data-bs-target="#collapseToco${index}"><strong>CF: ${cf}</strong> &nbsp; <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${grupo.pedidos.length}</span> ${weightBadge}</button></h2><div id="collapseToco${index}" class="accordion-collapse collapse" data-bs-parent="#accordionToco"><div class="accordion-body drop-zone-card" id="${loadId}" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event)" data-load-id="${loadId}" data-vehicle-type="toco"><button class="btn btn-sm btn-outline-info mb-3 no-print" onclick="imprimirTocoIndividual('${cf}')"><i class="bi bi-printer-fill me-1"></i>Imprimir</button>${progressBar}${createTable(grupo.pedidos, null, loadId)}</div></div></div>`;
    });
    accordionHtml += '</div>'; div.innerHTML = accordionHtml;
}

function displayCargasCfAccordion(div, grupos, pedidosCarreta, pedidosCondor) {
    let todosOsGrupos = {...grupos};
    const chaveCarreta = "Pedidos de Carreta/Truck sem CF";
    const chaveCondor = "Pedidos Condor (Truck)";

    if (pedidosCarreta && pedidosCarreta.length > 0) {
        todosOsGrupos[chaveCarreta] = { pedidos: pedidosCarreta, totalKg: pedidosCarreta.reduce((s, p) => s + p.Quilos_Saldo, 0), totalCubagem: pedidosCarreta.reduce((s, p) => s + p.Cubagem, 0) };
        gruposPorCFGlobais[chaveCarreta] = todosOsGrupos[chaveCarreta];
    }
    if (pedidosCondor && pedidosCondor.length > 0) {
        todosOsGrupos[chaveCondor] = { pedidos: pedidosCondor, totalKg: pedidosCondor.reduce((s, p) => s + p.Quilos_Saldo, 0), totalCubagem: pedidosCondor.reduce((s, p) => s + p.Cubagem, 0) };
        gruposPorCFGlobais[chaveCondor] = todosOsGrupos[chaveCondor];
    }

    if (Object.keys(todosOsGrupos).length === 0) {
        div.innerHTML = '<div class="empty-state"><i class="bi bi-truck"></i><p>Nenhum grupo de carga maior encontrado.</p></div>';
        return;
    }
    
    let accordionHtml = '<div class="accordion accordion-flush" id="accordionCargasPorCF">';
    const sortedKeys = Object.keys(todosOsGrupos).sort((a,b) => a.localeCompare(b, undefined, {numeric: true}));

    sortedKeys.forEach((cf, index) => {
        const grupo = todosOsGrupos[cf];
        grupo.pedidos.sort((a,b) => String(a.Cliente).localeCompare(String(b.Cliente),undefined,{numeric:true})||String(a.Num_Pedido).localeCompare(String(b.Num_Pedido),undefined,{numeric:true}));
        const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const totalCubagemFormatado = grupo.totalCubagem.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const collapseId = `collapseCF-Mesa-${String(cf).replace(/\s|\(|\)|\//g, '')}`;

        accordionHtml += `
            <div class="accordion-item">
                <h2 class="accordion-header" id="headingCargaCFMesa${index}">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#${collapseId}">
                        <strong>${isNumeric(cf) ? "CF: " : ""}${cf}</strong> &nbsp;
                        <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKgFormatado} kg</span>
                        <span class="badge bg-light text-dark ms-2"><i class="bi bi-rulers me-1"></i>${totalCubagemFormatado} m³</span>
                    </button>
                </h2>
                <div id="${collapseId}" class="accordion-collapse collapse" data-bs-parent="#accordionCargasPorCF">
                    <div class="accordion-body">
                        <button class="btn btn-sm btn-outline-info mb-3 no-print" onclick="imprimirCargaCFIndividual('${cf}')"><i class="bi bi-printer-fill me-1"></i>Imprimir Carga</button>
                        ${createTable(grupo.pedidos)}
                    </div>
                </div>
            </div>`;
    });
    accordionHtml += '</div>';
    div.innerHTML = accordionHtml;
}


function imprimirCargaCFIndividual(cf) {
    if (!gruposPorCFGlobais || !gruposPorCFGlobais[cf]) {
        alert(`Nenhuma carga encontrada para: ${cf}`);
        return;
    }
    const grupo = gruposPorCFGlobais[cf];
    const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const totalCubagemFormatado = grupo.totalCubagem.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const printWindow = createPrintWindow(`Imprimir Carga: ${cf}`);
    let contentToPrint = `<h3>Carga: ${cf} - Total: ${totalKgFormatado} kg / ${totalCubagemFormatado} m³</h3>` + createTable(grupo.pedidos);
    printWindow.document.body.innerHTML = contentToPrint;
    printWindow.document.close();
    printWindow.focus();
    setTimeout(() => { printWindow.print(); printWindow.close(); }, 500);
}

function montarCargaPredefinida(inputId, resultadoId, processedSet, nomeCarga) {
    const resultadoDiv = document.getElementById(resultadoId);
    resultadoDiv.innerHTML = '';
    if (planilhaData.length === 0) {
        resultadoDiv.innerHTML = '<div class="alert alert-warning">Carregue a planilha primeiro.</div>';
        return;
    }
    const numerosPedidos = document.getElementById(inputId).value.split('\n').map(n => n.trim()).filter(Boolean);
    if (numerosPedidos.length === 0) {
        resultadoDiv.innerHTML = `<div class="alert alert-warning">Nenhum pedido inserido para a ${nomeCarga}.</div>`;
        return;
    }
    const pedidosSelecionados = [];
    const pedidosNaoEncontrados = [];
    numerosPedidos.forEach(num => {
        const pedidoEncontrado = planilhaData.find(p => String(p.Num_Pedido) === num);
        if (pedidoEncontrado) pedidosSelecionados.push(pedidoEncontrado);
        else pedidosNaoEncontrados.push(num);
    });
    if (pedidosNaoEncontrados.length > 0) {
        resultadoDiv.innerHTML = `<div class="alert alert-danger">Pedidos não encontrados: ${pedidosNaoEncontrados.join(', ')}.</div>`;
        return;
    }
    const totalKg = pedidosSelecionados.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
    const totalCubagem = pedidosSelecionados.reduce((sum, p) => sum + p.Cubagem, 0);
    const veiculos = [
        { tipo: 'fiorino', nome: 'Fiorino', maxKg: getVehicleConfig('fiorino').hardMaxKg, maxCubagem: getVehicleConfig('fiorino').hardMaxCubage },
        { tipo: 'van', nome: 'Van', maxKg: getVehicleConfig('van').hardMaxKg, maxCubagem: getVehicleConfig('van').hardMaxCubage },
        { tipo: 'tresQuartos', nome: '3/4', maxKg: getVehicleConfig('tresQuartos').hardMaxKg, maxCubagem: getVehicleConfig('tresQuartos').hardMaxCubage },
    ];
    const veiculoEscolhido = veiculos.find(v => totalKg <= v.maxKg && totalCubagem <= v.maxCubagem);
    
    if (veiculoEscolhido) {
        processedSet.clear();
        pedidosSelecionados.forEach(p => processedSet.add(String(p.Num_Pedido)));
        const load = { id: `${nomeCarga.toLowerCase().replace(/\s+/g, '-')}-1`, pedidos: pedidosSelecionados, totalKg, totalCubagem, numero: nomeCarga, vehicleType: veiculoEscolhido.tipo };
        activeLoads[load.id] = load;
        
        const vehicleInfoMap = {
            fiorino: { name: 'Fiorino', colorClass: 'bg-success', textColor: 'text-white', icon: 'bi-box-seam-fill' },
            van: { name: 'Van', colorClass: 'bg-primary', textColor: 'text-white', icon: 'bi-truck-front-fill' },
            tresQuartos: { name: '3/4', colorClass: 'bg-warning', textColor: 'text-dark', icon: 'bi-truck-flatbed' }
        };
        resultadoDiv.innerHTML = `
            <div class="alert alert-success d-flex justify-content-between align-items-center">
                <div><strong>Carga ${nomeCarga} montada!</strong> ${pedidosSelecionados.length} pedidos agrupados em um(a) <strong>${veiculoEscolhido.nome}</strong>.</div>
                <button class="btn btn-light btn-sm no-print" onclick="imprimirGeneric('${resultadoId}', 'Carga ${nomeCarga}')"><i class="bi bi-printer-fill me-1"></i> Imprimir</button>
            </div>
            ${renderLoadCard(load, veiculoEscolhido.tipo, vehicleInfoMap[veiculoEscolhido.tipo])}`;
    } else {
        processedSet.clear();
        resultadoDiv.innerHTML = `<div class="alert alert-danger"><strong>Não foi possível montar a carga.</strong> O total excede a capacidade dos veículos. (Peso: ${totalKg.toFixed(2)} kg, Cubagem: ${totalCubagem.toFixed(2)} m³)</div>`;
    }
}

function highlightClientRows(event) {
    const clickedRow = event.target.closest('tr');
    if (!clickedRow || !clickedRow.dataset.clienteId) return;

    const clienteId = clickedRow.dataset.clienteId;
    const isAlreadyHighlighted = clickedRow.classList.contains('client-highlight');

    document.querySelectorAll('tr.client-highlight').forEach(row => row.classList.remove('client-highlight'));
    if (!isAlreadyHighlighted) {
        document.querySelectorAll(`tr[data-cliente-id='${clienteId}']`).forEach(row => row.classList.add('client-highlight'));
    }
}

function dragStart(event, pedidoId, clienteId, sourceId) {
    event.dataTransfer.setData("text/plain", JSON.stringify({ pedidoId, clienteId, sourceId }));
    event.dataTransfer.effectAllowed = "move";
}

function dragOver(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
    event.target.closest('.drop-zone-card')?.classList.add('drag-over');
}

function dragLeave(event) {
    event.target.closest('.drop-zone-card')?.classList.remove('drag-over');
}

function drop(event) {
    event.preventDefault();
    const dropZoneCard = event.target.closest('.drop-zone-card');
    if (!dropZoneCard) return;
    dropZoneCard.classList.remove('drag-over');

    const { clienteId, sourceId } = JSON.parse(event.dataTransfer.getData("text/plain"));
    const targetId = dropZoneCard.dataset.loadId;

    if (sourceId === targetId) return; 

    const getSourceLoad = (id) => {
        if (id === 'leftovers') return { pedidos: currentLeftoversForPrinting };
        if (id === 'geral') return { pedidos: pedidosGeraisAtuais };
        if (id === 'manual-builder') return manualLoadInProgress;
        return activeLoads[id];
    };
    const getTargetLoad = (id) => {
        if (id === 'leftovers') return { pedidos: currentLeftoversForPrinting, totalKg: 0, totalCubagem: 0 };
        if (id === 'manual-builder') return manualLoadInProgress;
        return activeLoads[id];
    };

    const sourceLoad = getSourceLoad(sourceId);
    const targetLoad = getTargetLoad(targetId);
    
    if (!sourceLoad || !targetLoad) return console.error("ERRO: Carga de origem ou destino não encontrada.");

    const clientOrdersToMove = sourceLoad.pedidos.filter(p => normalizeClientId(p.Cliente) === clienteId);
    if (clientOrdersToMove.length === 0) return;

    const orderIdsToMove = new Set(clientOrdersToMove.map(p => p.Num_Pedido));
    const clientBlockKg = clientOrdersToMove.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
    const clientBlockCubagem = clientOrdersToMove.reduce((sum, p) => sum + p.Cubagem, 0);
    
    if (targetId !== 'leftovers' && targetId !== 'manual-builder') {
        const config = getVehicleConfig(targetLoad.vehicleType);
        if ((targetLoad.totalKg + clientBlockKg > config.hardMaxKg) || (targetLoad.totalCubagem + clientBlockCubagem > config.hardMaxCubage)) {
            if (!confirm(`AVISO: Mover este grupo excede a capacidade do veículo. Deseja continuar?`)) return;
        }
    }
    
    // Remover da origem
    if (sourceId === 'leftovers') currentLeftoversForPrinting = sourceLoad.pedidos.filter(p => !orderIdsToMove.has(p.Num_Pedido));
    else if (sourceId === 'geral') pedidosGeraisAtuais = sourceLoad.pedidos.filter(p => !orderIdsToMove.has(p.Num_Pedido));
    else {
        sourceLoad.pedidos = sourceLoad.pedidos.filter(p => !orderIdsToMove.has(p.Num_Pedido));
        sourceLoad.totalKg -= clientBlockKg;
        sourceLoad.totalCubagem -= clientBlockCubagem;
    }

    // Adicionar ao destino
    if (targetId === 'leftovers') currentLeftoversForPrinting.push(...clientOrdersToMove);
    else {
        targetLoad.pedidos.push(...clientOrdersToMove);
        targetLoad.totalKg += clientBlockKg;
        targetLoad.totalCubagem += clientBlockCubagem;
    }
    
    // Atualizar UI
    if (sourceId === 'manual-builder') updateManualBuilderUI(); else updateLoadUI(sourceId);
    if (targetId === 'manual-builder') updateManualBuilderUI(); else updateLoadUI(targetId);

    updateAndRenderChart();
}

function updateLoadUI(loadId) {
    const activeTabPane = document.querySelector('.tab-pane.active');
    
    if (loadId === 'leftovers' || loadId === 'geral') {
        const geraisDiv = document.getElementById('resultado-geral');
        if (geraisDiv) {
            const gruposGerais = pedidosGeraisAtuais.reduce((acc, p) => {
                const rota = p.Cod_Rota; if (!acc[rota]) { acc[rota] = { pedidos: [], totalKg: 0 }; } acc[rota].pedidos.push(p); acc[rota].totalKg += p.Quilos_Saldo; return acc;
            }, {});
            displayGerais(geraisDiv, gruposGerais);
        }
        
        const leftoversCard = activeTabPane ? activeTabPane.querySelector('[data-load-id="leftovers"]') : null;
        if (!leftoversCard) return;

        if (currentLeftoversForPrinting.length === 0) {
            leftoversCard.innerHTML = `
                <h5 class="mt-4">Sobras Finais: 0,00 kg</h5>
                <div class="card mb-3"><div class="card-header bg-danger text-white">Pedidos Restantes</div><div class="card-body"><p class="text-muted text-center py-3">Todos os pedidos foram alocados.</p></div></div>`;
        } else {
            const totalKgUnassigned = currentLeftoversForPrinting.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
            leftoversCard.innerHTML = `
                <h5 class="mt-4">Sobras Finais: ${totalKgUnassigned.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} kg</h5>
                <div class="card mb-3"><div class="card-header bg-danger text-white d-flex align-items-center">Pedidos Restantes<div class="ms-auto"><button id="start-manual-load-btn" class="btn btn-success ms-auto no-print" onclick="startManualLoadBuilder()"><i class="bi bi-plus-circle-fill me-1"></i>Criar Carga Manual</button><button class="btn btn-info ms-2 no-print" onclick="imprimirSobras('Sobras Finais')"><i class="bi bi-printer-fill me-1"></i>Imprimir</button></div></div><div class="card-body">${createTable(currentLeftoversForPrinting, ['Num_Pedido', 'Quilos_Saldo', 'Agendamento', 'Cubagem', 'Predat', 'Cliente', 'Nome_Cliente', 'Cidade', 'CF'], 'leftovers')}</div></div>`;
        }
        return;
    }

    const load = activeLoads[loadId];
    if (!load) return;

    const cardElement = document.getElementById(loadId);
    if (!cardElement) return;

    const vInfoMap = {
        fiorino: { name: 'Fiorino', colorClass: 'bg-success', textColor: 'text-white', icon: 'bi-box-seam-fill' },
        van: { name: 'Van', colorClass: 'bg-primary', textColor: 'text-white', icon: 'bi-truck-front-fill' },
        tresQuartos: { name: '3/4', colorClass: 'bg-warning', textColor: 'text-dark', icon: 'bi-truck-flatbed' },
        toco: { name: 'Toco', colorClass: 'bg-dark', textColor: 'text-white', icon: 'bi-inboxes-fill'}
    };
    const vInfo = vInfoMap[load.vehicleType];
    
    if(load.vehicleType !== 'toco') {
        cardElement.outerHTML = renderLoadCard(load, load.vehicleType, vInfo);
    } else {
        const accordionBody = cardElement;
        const headerButton = accordionBody.closest('.accordion-item').querySelector('.accordion-button');
        const maxKg = getVehicleConfig('toco').hardMaxKg;
        const totalKgFormatado = load.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const isOverloaded = load.totalKg > maxKg;
        
        accordionBody.querySelector('.table-responsive').outerHTML = createTable(load.pedidos, null, loadId);

        const weightBadge = headerButton.querySelector('.badge.bg-info, .badge.bg-danger');
        if(weightBadge) {
            weightBadge.outerHTML = isOverloaded
            ? `<span class="badge bg-danger ms-2"><i class="bi bi-exclamation-triangle-fill me-1"></i>${totalKgFormatado} kg (ACIMA!)</span>`
            : `<span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKgFormatado} kg</span>`;
        }
        headerButton.classList.toggle('bg-danger', isOverloaded);
    }
}

function startManualLoadBuilder() {
    if (document.getElementById('manual-load-builder-wrapper')) return; 

    manualLoadInProgress = { pedidos: [], totalKg: 0, totalCubagem: 0, vehicleType: 'fiorino' };
    
    const activeTabPane = document.querySelector('.tab-pane.active');
    if (!activeTabPane) return alert("Erro: Nenhuma aba de trabalho está ativa.");

    const builderWrapper = document.createElement('div');
    builderWrapper.id = 'manual-load-builder-wrapper';
    builderWrapper.className = 'p-3 border-top border-secondary'; 
    
    builderWrapper.innerHTML = `
        <div class="card border-info shadow-lg" id="manual-load-card">
            <div class="card-header bg-info text-dark"><h5 class="mb-0"><i class="bi bi-tools me-2"></i>Painel de Montagem Manual</h5></div>
            <div class="card-body">
                <div class="row align-items-center mb-3">
                    <div class="col-md-4"><label for="manualVehicleType" class="form-label">Veículo:</label><select id="manualVehicleType" class="form-select" onchange="updateManualBuilderUI()"><option value="fiorino">Fiorino</option><option value="van">Van</option><option value="tresQuartos">3/4</option></select></div>
                    <div class="col-md-5"><p class="mb-1"><strong>Peso:</strong> <span id="manualLoadKg">0,00</span> kg</p><p class="mb-0"><strong>Cubagem:</strong> <span id="manualLoadCubage">0,00</span> m³</p></div>
                    <div class="col-md-3 text-end"><button class="btn btn-danger me-2" onclick="cancelManualLoad()"><i class="bi bi-x-circle"></i></button><button id="finalizeManualLoadBtn" class="btn btn-success" onclick="finalizeManualLoad()" disabled><i class="bi bi-check-circle"></i> Criar</button></div>
                </div>
                <div id="manual-progress-bar-container"></div>
                <div id="manual-drop-zone" class="p-3 border rounded drop-zone-card" style="background-color: var(--dark-bg); min-height: 150px;" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event)" data-load-id="manual-builder">
                    <p class="text-muted text-center" id="manual-drop-text">Arraste os pedidos da lista de "Sobras" para cá.</p>
                    <div id="manual-load-table-container"></div>
                </div>
            </div>
        </div>`;
    
    activeTabPane.appendChild(builderWrapper);
    
    const startBtn = activeTabPane.querySelector('#start-manual-load-btn');
    if (startBtn) startBtn.style.display = 'none';
    updateManualBuilderUI(); 
}

function updateManualBuilderUI() {
    if (!manualLoadInProgress) return;
    manualLoadInProgress.vehicleType = document.getElementById('manualVehicleType').value;
    document.getElementById('manualLoadKg').textContent = manualLoadInProgress.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    document.getElementById('manualLoadCubage').textContent = manualLoadInProgress.totalCubagem.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    
    const tableContainer = document.getElementById('manual-load-table-container');
    document.getElementById('manual-drop-text').style.display = manualLoadInProgress.pedidos.length > 0 ? 'none' : 'block';
    tableContainer.innerHTML = manualLoadInProgress.pedidos.length > 0 ? createTable(manualLoadInProgress.pedidos, null, 'manual-builder') : '';
    
    const config = getVehicleConfig(manualLoadInProgress.vehicleType);
    document.getElementById('finalizeManualLoadBtn').disabled = manualLoadInProgress.totalKg < config.minKg;
    
    const pesoPercentual = config.hardMaxKg > 0 ? (manualLoadInProgress.totalKg / config.hardMaxKg) * 100 : 0;
    let progressColor = 'bg-secondary';
    if(manualLoadInProgress.totalKg >= config.minKg) progressColor = 'bg-success';
    if (pesoPercentual > 75) progressColor = 'bg-warning';
    if (pesoPercentual > 95) progressColor = 'bg-danger';
    document.getElementById('manual-progress-bar-container').innerHTML = `<div class="progress mb-3" style="height: 10px;"><div class="progress-bar ${progressColor}" style="width: ${Math.min(pesoPercentual, 100)}%"></div></div>`;
}

function finalizeManualLoad() {
    if (!manualLoadInProgress || manualLoadInProgress.pedidos.length === 0) return;
    const vehicleType = manualLoadInProgress.vehicleType;
    const newLoad = { ...manualLoadInProgress, id: `manual-${vehicleType}-${Date.now()}`, numero: `M-${Object.keys(activeLoads).filter(k => k.startsWith('manual')).length + 1}` };
    activeLoads[newLoad.id] = newLoad;
    
    const vehicleInfoMap = { fiorino: { name: 'Fiorino', colorClass: 'bg-success', textColor: 'text-white', icon: 'bi-box-seam-fill' }, van: { name: 'Van', colorClass: 'bg-primary', textColor: 'text-white', icon: 'bi-truck-front-fill' }, tresQuartos: { name: '3/4', colorClass: 'bg-warning', textColor: 'text-dark', icon: 'bi-truck-flatbed' } };
    const newCardHTML = renderLoadCard(newLoad, vehicleType, vehicleInfoMap[vehicleType]);
    
    const activeTabPane = document.querySelector('.tab-pane.active');
    activeTabPane?.querySelector('.resultado-container')?.insertAdjacentHTML('beforeend', newCardHTML);
    
    document.getElementById('manual-load-builder-wrapper')?.remove();
    manualLoadInProgress = null;
    activeTabPane?.querySelector('#start-manual-load-btn')?.style.display = 'inline-block';
    
    updateAndRenderChart();
}

function cancelManualLoad() {
    if (!manualLoadInProgress) return;
    currentLeftoversForPrinting.push(...manualLoadInProgress.pedidos);
    document.getElementById('manual-load-builder-wrapper')?.remove();
    manualLoadInProgress = null;
    updateLoadUI('leftovers'); 
    document.querySelector('.tab-pane.active #start-manual-load-btn')?.style.display = 'inline-block';
}

function displayGenericAccordionList(div, pedidos, title, accordionIdPrefix, description) {
    if (!div) return;
    if (pedidos.length === 0) {
        div.innerHTML = `<div class="empty-state"><i class="bi bi-info-circle"></i><p>Nenhum pedido de ${title.toLowerCase()} encontrado.</p></div>`;
        return;
    }
    pedidos.sort((a,b) => String(a.Cliente).localeCompare(String(b.Cliente),undefined,{numeric:true})||String(a.Num_Pedido).localeCompare(String(b.Num_Pedido),undefined,{numeric:true}));
    const totalKg = pedidos.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
    const printAreaId = `${accordionIdPrefix}-print-area`;

    div.innerHTML = `
        <div class="accordion accordion-flush" id="accordion${accordionIdPrefix}">
            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse${accordionIdPrefix}">
                        <strong>${title}</strong>
                        <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${pedidos.length} Pedidos</span>
                        <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKg.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})} kg</span>
                    </button>
                </h2>
                <div id="collapse${accordionIdPrefix}" class="accordion-collapse collapse" data-bs-parent="#accordion${accordionIdPrefix}">
                    <div class="accordion-body" id="${printAreaId}">
                        <div class="d-flex justify-content-between align-items-center mb-3 no-print">
                            <p class="text-muted small mb-0">${description}</p>
                            <button class="btn btn-sm btn-outline-info" onclick="imprimirGeneric('${printAreaId}', '${title}')"><i class="bi bi-printer-fill me-1"></i>Imprimir</button>
                        </div>
                        ${createTable(pedidos, ['Num_Pedido', 'Cliente', 'Nome_Cliente', 'Quilos_Saldo', 'Cidade', 'Predat', 'Coluna5', 'BLOQ.'])}
                    </div>
                </div>
            </div>
        </div>`;
}
