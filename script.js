// ================================================================================================
//  SCRIPT COMPLETO E CORRIGIDO (v2)
//  - Corrigido o erro "ReferenceError: displayGenericAccordion is not defined".
//  - Barra de busca aprimorada para encontrar pedidos em qualquer local.
//  - Relatório de atrasos atualizado para incluir pedidos com bloqueio 'C'.
//  - Código refatorado para melhor performance, organização e legibilidade.
// ================================================================================================
'use strict';

// Variáveis Globais de Estado
let planilhaData = [];
let originalColumnHeaders = [];
let pedidosGeraisAtuais = [];
let gruposToco = {};
let gruposPorCFGlobais = {};
let pedidosComCFNumericoIsolado = [];
let pedidosManualmenteBloqueadosAtuais = [];
let pedidosPrioritarios = new Set();
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
let resumoChart = null;

// Mapeamento de Clientes e Rotas
const agendamentoClientCodes = new Set(['1398', '1494', '4639', '4872', '5546', '6896', 'D11238', '17163', '19622', '20350', '22545', '23556', '23761', '24465', '29302', '32462', '32831', '32851', '32869', '32905', '33039', '33046', '33047', '33107', '33388', '33392', '33400', '33401', '33403', '33406', '33420', '33494', '33676', '33762', '33818', '33859', '33907', '33971', '34011', '34096', '34167', '34425', '34511', '34810', '34981', '35050', '35054', '35798', '36025', '36580', '36792', '36853', '36945', '37101', '37589', '37634', '38207', '38448', '38482', '38564', '38681', '38735', 'D38896', '39081', '39177', '39620', '40144', '40442', '40702', '40844', '41233', '42200', '42765', '47244', '47253', '47349', '50151', '50816', '51993', '52780', '53134', '58645', '60900', '61182', '61315', '61316', '61317', '61318', '61324', '63080', '63500', '63705', '64288', '66590', '67660', '67884', '69281', '69286', '69318', '70968', '71659', '73847', '76019', '76580', '77475', '77520', '78895', '79838', '80727', '81353', 'DB3183', '83184', '83634', '85534', 'DB6159', '86350', '86641', '89073', '89151', '90373', '92017', '95092', '95660', '96758', '98227', '99268', '100087', '101246', '101253', '101346', '103518', '105394', '106198', '109288', '110023', '110894', '111145', '111154', '111302', '112207', '112670', '117028', '117123', '120423', '120455', '120473', '120533', '121747', '122155', '122785', '123815', '124320', '125228', '126430', '131476', '132397', '133916', '135395', '135928', '136086', '136260', '137919', '138825', '139013', '139329', '139611', '143102', '44192', '144457', '145014', '145237', '145322', '146644', '146988', '148071', '149598', '150503', '151981', '152601', '152835', '152925', '153289', '154423', '154778', '154808', '155177', '155313', '155368', '155419', '155475', '155823', '155888', '156009', '156585', '156696', '157403', '158235', '159168', '160382', '160982', '161737', '162499', '162789', '163234', '163382', '163458', '164721', '164779', '164780', '164924', '165512', '166195', '166337', '166353', '166468', '166469', '167353', '167810', '167819', '168464', '169863', '169971', '170219', '170220', '170516', '171147', '171160', '171191', '171200', '171320', '171529', '171642', '171863', '172270', '172490', '172656', '172859', '173621', '173964', '173977', '174249', '174593', '174662', '174901', '175365', '175425', '175762', '175767', '175783', '176166', '176278', '176453', '176747', '177327', '177488', '177529', '177883', '177951', '177995', '178255', '178377', '178666', '179104', '179510', '179542', '179690', '180028', '180269', '180342', '180427', '180472', '180494', '180594', '180772', '181012', '181052', '181179', '182349', '182885', '182901', '183011', '183016', '183046', '183048', '183069', '183070', '183091', '183093', '183477', '183676', '183787', '184011', '184038', '189677', '190163', '190241', '190687', '190733', '191148', '191149', '191191', '191902', '191972', '192138', '192369', '192638', '192713', '193211', '193445', '193509', '194432', '194508', '194750', '194751', '194821', '194831', '195287', '195338', '195446', '196118', '196405', '196446', '196784', '197168', '197249', '197983', '198187', '198438', '198747', '198796', '198895', '198907', '198908', '199172', '199615', '199625', '199650', '199651', '199713', '199733', '199927', '199991', '200091', '200194', '200239', '200253', '200382', '200404', '200597', '200917', '201294', '201754', '201853', '201936', '201948', '201956', '201958', '201961', '201974', '202022', '202187', '202199', '202714', '203072', '203093', '203201', '203435', '203436', '203451', '203512', '203769', '204895', '204910', '204911', '204913', '204914', '204915', '204917', '204971', '204979', '205108', '205220', '205744', '205803', '206116', '206163', '206208', '206294', '206380', '206628', '206730', '206731', '206994', '207024', '207029', '207403', '207689', '207902', '208489', '208613', '208622', '208741', '208822', '208844', '208853', '208922', '209002', '209004', '209248', '209281', '209321', '209322', '209684', '210124', '210230', '210490', '210747', '210759', '210819', '210852', '211059', '211110', '211276', '211277', '211279', '211332', '211411', '212401', '212417', '212573', '212900', '213188', '213189', '213190', '213202', '213203', '213242', '213442', '213454', '213855', '213909', '213910', '213967', '214046', '214150', '214387', '214433', '214442', '214594', '214746', '215022', '215116', '215160', '215161', '215493', '215494', '215651', '215687', '215733', '215777', '215942', '216112', '216393', '216400', '216630', '216684', '217190', '217283', '217310', '217343', '217545', '217605', '217828', '217871', '217872', '217877', '217949', '217965', '218169', '218196', '218383', '218486', '218578', '218580', '218640', '218820', '218845', '219539', '219698', '219715', '219884', '220158', '220183', '220645', '220950', '221023', '221248', '221251', '222164', '222165', '223025', '223379', '223525', '223703', '223727', '223877', '223899', '223900', '223954', '224956', '224957', '224958', '224959', '224961', '224962', '225112', '225408', '225449', '225904', '226903', '226939', '227190', '227387', '228589', '228693', '228695']);
const specialClientNames = ['IRMAOS MUFFATO S.A', 'FINCO & FINCO', 'BOM DIA'];
const rotaVeiculoMap = {
    '11101': { type: 'fiorino', title: 'Rota 11101' }, '11301': { type: 'fiorino', title: 'Rota 11301' }, '11311': { type: 'fiorino', title: 'Rota 11311' }, '11561': { type: 'fiorino', title: 'Rota 11561' }, '11721': { type: 'fiorino', title: 'Rotas 11721 & 11731', combined: ['11731'] }, '11731': { type: 'fiorino', title: 'Rotas 11721 & 11731', combined: ['11721'] },
    '11102': { type: 'van', title: 'Rota 11102' }, '11331': { type: 'van', title: 'Rota 11331' }, '11341': { type: 'van', title: 'Rota 11341' }, '11342': { type: 'van', title: 'Rota 11342' }, '11351': { type: 'van', title: 'Rota 11351' }, '11521': { type: 'van', title: 'Rota 11521' }, '11531': { type: 'van', title: 'Rota 11531' }, '11551': { type: 'van', title: 'Rota 11551' }, '11571': { type: 'van', title: 'Rota 11571' }, '11701': { type: 'van', title: 'Rota 11701' }, '11711': { type: 'van', title: 'Rota 11711' },
    '11361': { type: 'tresQuartos', title: 'Rota 11361' }, '11501': { type: 'tresQuartos', title: 'Rotas 11501, 11502 & 11511', combined: ['11502', '11511'] }, '11502': { type: 'tresQuartos', title: 'Rotas 11501, 11502 & 11511', combined: ['11501', '11511'] }, '11511': { type: 'tresQuartos', title: 'Rotas 11501, 11502 & 11511', combined: ['11501', '11502'] }, '11541': { type: 'tresQuartos', title: 'Rota 11541' }
};
const defaultConfigs = {
    fiorinoMinCapacity: 300, fiorinoMaxCapacity: 500, fiorinoCubage: 1.5, fiorinoHardMaxCapacity: 560, fiorinoHardCubage: 1.7,
    vanMinCapacity: 1100, vanMaxCapacity: 1560, vanCubage: 5.0, vanHardMaxCapacity: 1600, vanHardCubage: 5.6,
    tresQuartosMinCapacity: 2300, tresQuartosMaxCapacity: 4100, tresQuartosCubage: 15.0,
    tocoMinCapacity: 5000, tocoMaxCapacity: 8500, tocoCubage: 30.0
};
// Funções Utilitárias
const isSpecialClient = (p) => p.Nome_Cliente && specialClientNames.includes(p.Nome_Cliente.toUpperCase().trim());
const isNumeric = (str) => str && /^\d+$/.test(String(str).trim());
const deepClone = (obj) => JSON.parse(JSON.stringify(obj));
const normalizeClientId = (id) => (id === null || typeof id === 'undefined') ? '' : String(id).trim().replace(/^0+/, '');

// Event Listeners da Interface
document.addEventListener('DOMContentLoaded', () => {
    loadConfigurations();

    document.getElementById('fileInput').addEventListener('change', (e) => handleFile(e.target.files[0]));
    document.getElementById('limparFiltroRotaBtn').addEventListener('click', limparFiltroDeRota);
    document.getElementById('processarBtn').addEventListener('click', processar);
    document.getElementById('limparResultadosBtn').addEventListener('click', limparTudo);
    document.querySelectorAll('#sidebar-nav a').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            document.querySelector('#sidebar-nav a.active')?.classList.remove('active');
            this.classList.add('active');
            document.querySelector(this.getAttribute('href'))?.scrollIntoView({ behavior: 'smooth' });
        });
    });

    document.getElementById('searchBtn').addEventListener('click', buscarPedido);
    document.getElementById('searchInput').addEventListener('keyup', (e) => { if (e.key === 'Enter') buscarPedido(); });

    document.getElementById('exportarAtrasadosBtn').addEventListener('click', exportarPedidosAtrasados);
    document.getElementById('bloquearPedidoBtn').addEventListener('click', bloquearPedido);
    document.getElementById('marcarSemCorteBtn').addEventListener('click', marcarPedidosSemCorte);
    document.getElementById('montarCargaEspecialBtn').addEventListener('click', () => montarCargaPredefinida('pedidosEspeciaisInput', 'resultado-carga-especial', pedidosEspeciaisProcessados, 'Especial'));
    document.getElementById('montarVendaAntecipadaBtn').addEventListener('click', () => montarCargaPredefinida('vendaAntecipadaInput', 'resultado-venda-antecipada', pedidosVendaAntecipadaProcessados, 'Venda Antecipada'));

    document.getElementById('saveConfig').addEventListener('click', saveConfigurations);
    document.getElementById('resetFiorino').addEventListener('click', () => resetVehicleConfig('fiorino'));
    document.getElementById('resetVan').addEventListener('click', () => resetVehicleConfig('van'));
    document.getElementById('resetTresQuartos').addEventListener('click', () => resetVehicleConfig('tresQuartos'));
    document.getElementById('resetToco').addEventListener('click', () => resetVehicleConfig('toco'));
    document.getElementById('resetAllConfigs').addEventListener('click', resetAll);
});

// Lógica de Configuração de Veículos
function saveConfigurations() {
    const configStatus = document.getElementById('configStatus');
    try {
        const configs = {};
        Object.keys(defaultConfigs).forEach(key => {
            const element = document.getElementById(key);
            if (element) configs[key] = parseFloat(element.value);
        });
        localStorage.setItem('vehicleConfigs', JSON.stringify(configs));
        configStatus.innerHTML = '<p class="text-success">Configurações salvas!</p>';
    } catch (error) {
        console.error("Erro ao salvar no localStorage:", error);
        configStatus.innerHTML = `<p class="text-danger">Erro ao salvar.</p>`;
    }
    setTimeout(() => { configStatus.innerHTML = ''; }, 3000);
}

function loadConfigurations() {
    try {
        const savedConfigs = localStorage.getItem('vehicleConfigs');
        const configs = savedConfigs ? JSON.parse(savedConfigs) : defaultConfigs;
        Object.keys(defaultConfigs).forEach(key => {
            const element = document.getElementById(key);
            if (element) element.value = configs[key] ?? defaultConfigs[key];
        });
    } catch (error) {
        console.error("Erro ao carregar do localStorage:", error);
        resetAll();
    }
}

function resetVehicleConfig(vehiclePrefix) {
    Object.keys(defaultConfigs)
        .filter(key => key.startsWith(vehiclePrefix))
        .forEach(key => {
            const element = document.getElementById(key);
            if(element) element.value = defaultConfigs[key];
        });
    saveConfigurations();
}

function resetAll() {
    Object.keys(defaultConfigs).forEach(key => {
        const element = document.getElementById(key);
        if(element) element.value = defaultConfigs[key];
    });
    saveConfigurations();
}

// Lógica Principal de Processamento
function handleFile(file) {
    if (!file) return;
    limparTudo();
    const statusDiv = document.getElementById('status');
    statusDiv.innerHTML = '<p class="text-info">Carregando planilha...</p>';
    document.getElementById('processarBtn').disabled = true;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            const headerRowIndex = rawData.findIndex(row => row && row.some(cell => String(cell).trim().toLowerCase() === 'cod_rota'));
            if (headerRowIndex === -1) throw new Error("Não foi possível encontrar o cabeçalho 'Cod_Rota'.");
            
            originalColumnHeaders = rawData[headerRowIndex].map(h => h ? String(h).trim() : '');
            const dataRows = rawData.slice(headerRowIndex + 1);

            planilhaData = dataRows.map(row => {
                const pedido = {};
                originalColumnHeaders.forEach((header, i) => {
                    if (header) pedido[header] = row[i] ?? '';
                });
                pedido.Quilos_Saldo = parseFloat(String(pedido.Quilos_Saldo).replace(',', '.')) || 0;
                pedido.Cubagem = parseFloat(String(pedido.Cubagem).replace(',', '.')) || 0;
                const normalizedCode = normalizeClientId(pedido.Cliente);
                pedido.Agendamento = agendamentoClientCodes.has(normalizedCode) ? 'Sim' : 'Não';
                return pedido;
            });

            statusDiv.innerHTML = `<p class="text-success">Planilha "${file.name}" carregada.</p>`;
            document.getElementById('processarBtn').disabled = false;
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

    const rotas = [...new Set(planilhaData.map(p => String(p.Cod_Rota || '')))].filter(Boolean);
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

function processar() {
    const statusDiv = document.getElementById('status');
    statusDiv.innerHTML = `<div class="d-flex align-items-center justify-content-center"><div class="spinner-border spinner-border-sm text-primary me-2"></div><span class="text-primary">Processando...</span></div>`;
    document.getElementById('processarBtn').disabled = true;

    setTimeout(() => {
        try {
            if (planilhaData.length === 0) throw new Error("Nenhum dado de planilha carregado.");

            resetarEstadoGlobal();
            limparResultadosVisuais();

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

            const pedidosExcluidos = new Set();
            dadosParaProcessar.forEach(p => {
                const coluna5 = String(p.Coluna5 || '').toUpperCase();
                if (coluna5.includes('TBL FUNCIONARIO')) {
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
            let dadosFiltrados = dadosParaProcessar.filter(p => !pedidosExcluidos.has(p.Num_Pedido));
            
            let pedidosDisponiveis = [];
            dadosFiltrados.forEach(p => {
                const numPedidoStr = String(p.Num_Pedido);
                if (pedidosBloqueados.has(numPedidoStr)) {
                    pedidosManualmenteBloqueadosAtuais.push(p);
                } else if (!pedidosEspeciaisProcessados.has(numPedidoStr) && !pedidosVendaAntecipadaProcessados.has(numPedidoStr)) {
                    pedidosDisponiveis.push(p);
                }
            });

            rota1SemCarga = pedidosDisponiveis.filter(p => {
                return String(p.Cod_Rota || '').trim() === '1' &&
                       !isNumeric(p.CF) &&
                       ['TBL 08', 'TBL TODESCHINI', 'PROMO BOLINHO'].some(termo => String(p.Coluna5 || '').toUpperCase().includes(termo));
            });
            const rota1PedidoIds = new Set(rota1SemCarga.map(p => p.Num_Pedido));
            let pedidosParaProcessamentoGeral = pedidosDisponiveis.filter(p => !rota1PedidoIds.has(p.Num_Pedido));
            
            const clientesComBloqueioSistema = new Set(
                pedidosParaProcessamentoGeral.filter(p => String(p['BLOQ.']).trim()).map(p => normalizeClientId(p.Cliente))
            );
            
            const pedidosTocoBase = pedidosParaProcessamentoGeral.filter(p => 
                (p.Coluna4 && String(p.Coluna4).toUpperCase().includes('TOCO')) || 
                (p.Coluna5 && String(p.Coluna5).toUpperCase().includes('TOCO'))
            );
            
            const cfCounts = {};
            pedidosTocoBase.forEach(p => { if (isNumeric(p.CF)) cfCounts[p.CF] = (cfCounts[p.CF] || 0) + 1; });
            const cfsRepetidos = new Set(Object.keys(cfCounts).filter(cf => cfCounts[cf] > 1));
            
            const pedidosTocoFiltrados = pedidosTocoBase.filter(p => cfsRepetidos.has(String(p.CF)));
            tocoPedidoIds = new Set(pedidosTocoFiltrados.map(p => String(p.Num_Pedido)));

            let pedidosParaVarejoFinal = [];
            pedidosParaProcessamentoGeral.forEach(p => {
                if (tocoPedidoIds.has(String(p.Num_Pedido))) return;
                if (['21', '23', '17'].includes(String(p.Coluna4))) return;
                
                if (isNumeric(p.CF)) {
                    const cf = String(p.CF).trim();
                    if (!gruposPorCFGlobais[cf]) gruposPorCFGlobais[cf] = { pedidos: [], totalKg: 0, totalCubagem: 0 };
                    gruposPorCFGlobais[cf].pedidos.push(p);
                } else {
                    if (clientesComBloqueioSistema.has(normalizeClientId(p.Cliente))) {
                        pedidosComCFNumericoIsolado.push(p);
                    } else {
                        pedidosParaVarejoFinal.push(p);
                    }
                }
            });

            pedidosGeraisAtuais = pedidosParaVarejoFinal;
            
            Object.values(gruposPorCFGlobais).forEach(g => {
                g.totalKg = g.pedidos.reduce((s, p) => s + p.Quilos_Saldo, 0);
                g.totalCubagem = g.pedidos.reduce((s, p) => s + p.Cubagem, 0);
            });
            gruposToco = pedidosTocoFiltrados.reduce((acc, p) => {
                const cf = p.CF;
                if (!acc[cf]) acc[cf] = { pedidos: [], totalKg: 0, totalCubagem: 0 };
                acc[cf].pedidos.push(p);
                acc[cf].totalKg += p.Quilos_Saldo;
                acc[cf].totalCubagem += p.Cubagem;
                return acc;
            }, {});

            renderizarResultados();

            statusDiv.innerHTML = `<p class="text-success">Processamento concluído!</p>`;
        } catch (error) {
            statusDiv.innerHTML = `<p class="text-danger"><strong>Ocorreu um erro:</strong></p><pre>${error.stack}</pre>`;
            console.error(error);
        } finally {
            document.getElementById('processarBtn').disabled = false;
        }
    }, 50);
}

// Lógica da Busca (Melhorada)
function buscarPedido() {
    const searchType = document.getElementById('searchType').value;
    const searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();
    const resultDiv = document.getElementById('search-result');

    if (!searchTerm) {
        resultDiv.innerHTML = '';
        return;
    }

    const allPedidos = buscarEmTodasAsFontes(searchTerm, searchType);

    if (allPedidos.length === 0) {
        resultDiv.innerHTML = '<div class="alert alert-warning">Nenhum pedido encontrado com este critério.</div>';
        return;
    }

    let html = `<div class="list-group">`;
    allPedidos.forEach(item => {
        html += `<div class="list-group-item list-group-item-action d-flex justify-content-between align-items-center">
                    <div>
                        <div class="fw-bold">${item.pedido.Num_Pedido} - ${item.pedido.Nome_Cliente}</div>
                        <small class="text-muted">${item.localizacao}</small>
                    </div>
                    <button class="btn btn-sm btn-outline-info" 
                            onclick="highlightPedido(this)"
                            data-pedido-id="${item.pedido.Num_Pedido}"
                            data-tab-id="${item.tabId || ''}"
                            data-accordion-id="${item.accordionId || ''}">
                        <i class="bi bi-bullseye"></i> Localizar
                    </button>
                 </div>`;
    });
    html += `</div>`;
    resultDiv.innerHTML = html;
}

function buscarEmTodasAsFontes(term, type) {
    const resultados = [];
    const adicionado = new Set();
    const matcher = (pedido) => String(pedido[type] || '').toLowerCase().includes(term);

    const adicionarResultado = (pedido, localizacao, tabId, accordionId) => {
        if (pedido && !adicionado.has(pedido.Num_Pedido) && matcher(pedido)) {
            resultados.push({ pedido, localizacao, tabId, accordionId });
            adicionado.add(pedido.Num_Pedido);
        }
    };
    
    // 1. Cargas ativas e sobras
    Object.values(activeLoads).forEach(load => {
        const tabMap = { fiorino: 'fiorino-tab', van: 'van-tab', tresQuartos: 'tres-quartos-tab', toco: 'cf-tab' };
        const tabId = tabMap[load.vehicleType];
        const accordionId = load.vehicleType === 'toco' ? `collapseToco${Object.keys(gruposToco).sort().indexOf(load.pedidos[0]?.CF)}` : null;
        load.pedidos.forEach(p => adicionarResultado(p, `Carga ${load.numero || ''}`, tabId, accordionId));
    });
    currentLeftoversForPrinting.forEach(p => {
        const activeTab = document.querySelector('#vehicleTabs .nav-link.active');
        adicionarResultado(p, 'Sobras na Mesa de Trabalho', activeTab?.id, null);
    });

    // 2. Grupos de Truck/Carreta
    Object.entries(gruposPorCFGlobais).forEach(([cf, grupo]) => {
        const accordionId = `collapseCF-Mesa-${String(cf).replace(/\s|\(|\)|\//g, '')}`;
        grupo.pedidos.forEach(p => adicionarResultado(p, `Carga Truck/Carreta (CF: ${cf})`, 'cf-tab', accordionId));
    });

    // 3. Listas Adicionais
    pedidosGeraisAtuais.forEach(p => {
        const accordionId = `collapseGeral${Object.keys(pedidosGeraisAtuais.reduce((acc, p) => { acc[p.Cod_Rota] = true; return acc; }, {})).sort().indexOf(p.Cod_Rota)}`;
        adicionarResultado(p, 'Pedidos Disponíveis', null, accordionId);
    });
    pedidosFuncionarios.forEach(p => adicionarResultado(p, 'Pedidos de Funcionários', 'funcionarios-tab', 'collapseFuncionários'));
    pedidosTransferencias.forEach(p => adicionarResultado(p, 'Pedidos de Transferência', 'transferencias-tab', 'collapseTransferências'));
    pedidosManualmenteBloqueadosAtuais.forEach(p => adicionarResultado(p, 'Bloqueados Manualmente', null, null));
    rota1SemCarga.forEach(p => adicionarResultado(p, 'Rota 1 para Alteração', null, null));
    pedidosComCFNumericoIsolado.forEach(p => adicionarResultado(p, 'Filtrados (Regra Bloqueio)', null, `collapseCF${Object.keys(pedidosComCFNumericoIsolado.reduce((acc,p)=>{acc[p.Cod_Rota]=true;return acc;},{})).sort().indexOf(p.Cod_Rota)}`));
    
    return resultados;
}

function highlightPedido(button) {
    const { pedidoId, tabId, accordionId } = button.dataset;

    const scrollToElement = () => {
        const row = document.getElementById(`pedido-${pedidoId}`);
        if (row) {
            row.scrollIntoView({ behavior: 'smooth', block: 'center' });
            row.classList.remove('search-highlight');
            void row.offsetWidth; // Trigger reflow
            row.classList.add('search-highlight');
        }
    };
    
    const expandAndScroll = () => {
        if (accordionId) {
            const collapseEl = document.getElementById(accordionId);
            if (collapseEl && !collapseEl.classList.contains('show')) {
                const bsCollapse = bootstrap.Collapse.getOrCreateInstance(collapseEl);
                collapseEl.addEventListener('shown.bs.collapse', scrollToElement, { once: true });
                bsCollapse.show();
            } else {
                 scrollToElement();
            }
        } else {
            scrollToElement();
        }
    };
    
    if (tabId) {
        const tabButton = document.getElementById(tabId);
        const activeTab = document.querySelector('#vehicleTabs .nav-link.active');
        if (tabButton && (!activeTab || tabButton.id !== activeTab.id)) {
            const bsTab = bootstrap.Tab.getOrCreateInstance(tabButton);
            // O target do evento é o próprio botão de tab que foi ativado
            document.querySelector('#vehicleTabs').addEventListener('shown.bs.tab', (event) => {
                 if (event.target.id === tabId) {
                    expandAndScroll();
                }
            }, { once: true });
            bsTab.show();
        } else {
            expandAndScroll();
        }
    } else {
        expandAndScroll();
    }
}

// Lógica do Relatório de Atrasados (Melhorada)
function exportarPedidosAtrasados() {
    if (planilhaData.length === 0) {
        alert("Por favor, carregue a planilha primeiro.");
        return;
    }

    const pedidosAtrasados = planilhaData.filter(p => {
        const bloqueado = String(p['BLOQ.'] || '').trim().toUpperCase();
        return isOverdue(p.Predat) && (bloqueado === '' || bloqueado === 'C');
    });

    if (pedidosAtrasados.length === 0) {
        alert("Nenhum pedido em atraso (ou bloqueado com 'C') foi encontrado na planilha completa.");
        return;
    }

    alert(`${pedidosAtrasados.length} pedidos em atraso serão exportados.`);

    const header = ['Cliente', 'Nome_Cliente', 'Cidade', 'UF', 'Num_Pedido', 'Quilos_Saldo', 'Predat', 'Dat_Ped', 'Coluna5', 'BLOQ.'];
    const dataToExport = pedidosAtrasados.map(p => {
        const newP = {};
        header.forEach(col => {
            let value = p[col];
            if ((col === 'Predat' || col === 'Dat_Ped') && value instanceof Date) {
                value = value.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
            }
            newP[col] = value;
        });
        return newP;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Pedidos Atrasados");
    XLSX.writeFile(workbook, "pedidos_atrasados.xlsx");
}

function isOverdue(predat) {
    if (!predat) return false;
    const date = predat instanceof Date ? predat : new Date(predat);
    if (isNaN(date)) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    date.setHours(0,0,0,0);
    return date < today;
}

// Funções de Limpeza e Reset
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
    pedidosPrioritarios.clear();
    tocoPedidoIds.clear();
    currentLeftoversForPrinting = [];
    activeLoads = {};
}

function limparResultadosVisuais() {
    const idsParaLimpar = [
        'resultado-venda-antecipada', 'resultado-carga-especial', 'resultado-bloqueados', 'resultado-geral',
        'resultado-fiorino-geral', 'resultado-van-geral', 'resultado-34-geral',
        'resultado-rota1', 'resultado-toco',
        'resultado-cf-accordion-container', 'resultado-cf-numerico', 'search-result',
        'resultado-funcionarios-tab', 'resultado-transferencias-tab'
    ];
    idsParaLimpar.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = '';
    });

    const emptyStates = {
        'resultado-geral': '<div class="empty-state"><i class="bi bi-file-earmark-excel"></i><p>Nenhum pedido de varejo disponível.</p></div>',
        'botoes-fiorino': '<div class="empty-state"><i class="bi bi-box"></i><p>Nenhuma rota de Fiorino disponível.</p></div>',
        'botoes-van': '<div class="empty-state"><i class="bi bi-truck-front-fill"></i><p>Nenhuma rota de Van disponível.</p></div>',
        'botoes-34': '<div class="empty-state"><i class="bi bi-truck-flatbed"></i><p>Nenhuma rota de 3/4 disponível.</p></div>',
        'resultado-cf-accordion-container': '<div class="empty-state"><i class="bi bi-truck"></i><p>Nenhum grupo de carga maior encontrado.</p></div>',
        'resultado-funcionarios-tab': '<div class="empty-state"><i class="bi bi-people-fill"></i><p>Nenhum pedido de funcionário encontrado.</p></div>',
        'resultado-transferencias-tab': '<div class="empty-state"><i class="bi bi-arrow-left-right"></i><p>Nenhum pedido de transferência encontrado.</p></div>'
    };
    Object.entries(emptyStates).forEach(([id, html]) => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = html;
    });

    if (resumoChart) {
        resumoChart.destroy();
        resumoChart = null;
    }
    document.getElementById('card-resumo-geral-container').style.display = 'none';
}

function limparTudo(){
    resetarEstadoGlobal();
    originalColumnHeaders = [];
    planilhaData = [];

    pedidosEspeciaisProcessados.clear();
    pedidosBloqueados.clear();
    pedidosSemCorte.clear();
    pedidosVendaAntecipadaProcessados.clear();
    
    limparResultadosVisuais();

    document.getElementById('filtro-rota-container').style.display = 'none';
    ['pedidosEspeciaisInput', 'bloquearPedidoInput', 'searchInput', 'semCorteInput', 'vendaAntecipadaInput', 'fileInput'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    
    atualizarListaBloqueados();
    atualizarListaSemCorte();
    limparFiltroDeRota();
    
    document.getElementById('processarBtn').disabled = true;
    document.getElementById('status').innerHTML = '';
}

// Restante do código original que foi omitido anteriormente...
// Cole aqui as suas funções originais de `updateAndRenderChart` até o final.

// =================== INÍCIO DO CÓDIGO RESTANTE =====================

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

function renderizarResultados() {
    displayGenericAccordion(document.getElementById('resultado-funcionarios-tab'), pedidosFuncionarios, 'Funcionários', 'Pedidos com a tag "TBL FUNCIONARIO" na Coluna 5.');
    displayGenericAccordion(document.getElementById('resultado-transferencias-tab'), pedidosTransferencias, 'Transferências', 'Pedidos com a tag "TRANSF. TODESCH" na Coluna 5.');
    displayPedidosBloqueados(document.getElementById('resultado-bloqueados'), pedidosManualmenteBloqueadosAtuais);
    displayRota1(document.getElementById('resultado-rota1'), rota1SemCarga);
    displayToco(document.getElementById('resultado-toco'), gruposToco);
    displayCargasCfAccordion(document.getElementById('resultado-cf-accordion-container'), gruposPorCFGlobais, pedidosCarretaSemCF, pedidosCondorTruck);
    displayPedidosCFNumerico(document.getElementById('resultado-cf-numerico'), pedidosComCFNumericoIsolado);
    displayGerais(document.getElementById('resultado-geral'), pedidosGeraisAtuais.reduce((acc, p) => {
        const rota = p.Cod_Rota || 'Sem Rota';
        if (!acc[rota]) { acc[rota] = { pedidos: [], totalKg: 0 }; }
        acc[rota].pedidos.push(p);
        acc[rota].totalKg += p.Quilos_Saldo;
        return acc;
    }, {}));
    updateAndRenderChart();
}


function displayGenericAccordion(div, pedidos, title, description) {
    if (!div) return;
    if (pedidos.length === 0) {
        div.innerHTML = `<div class="empty-state"><i class="bi bi-info-circle"></i><p>Nenhum pedido de ${title.toLowerCase()} encontrado.</p></div>`;
        return;
    }

    pedidos.sort((a, b) => String(a.Num_Pedido).localeCompare(String(b.Num_Pedido), undefined, { numeric: true }));
    const totalKg = pedidos.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
    
    const accordionId = `accordion${title.replace(/\s/g, '')}`;
    const collapseId = `collapse${title.replace(/\s/g, '')}`;
    const printAreaId = `${title.toLowerCase().replace(/\s/g, '')}-print-area`;

    let accordionHtml = `<div class="accordion accordion-flush" id="${accordionId}">
        <div class="accordion-item">
            <h2 class="accordion-header">
                <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#${collapseId}">
                    <strong>Pedidos de ${title}</strong> &nbsp;
                    <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${pedidos.length} Pedidos</span>
                    <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKg.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})} kg</span>
                </button>
            </h2>
            <div id="${collapseId}" class="accordion-collapse collapse" data-bs-parent="#${accordionId}">
                <div class="accordion-body" id="${printAreaId}">
                    <div class="d-flex justify-content-between align-items-center mb-3 no-print">
                        <p class="text-muted small mb-0">${description}</p>
                        <button class="btn btn-sm btn-outline-info" onclick="imprimirGeneric('${printAreaId}', 'Pedidos de ${title}')">
                            <i class="bi bi-printer-fill me-1"></i>Imprimir Lista
                        </button>
                    </div>
                    ${createTable(pedidos, ['Num_Pedido', 'Cliente', 'Nome_Cliente', 'Quilos_Saldo', 'Cidade', 'Predat', 'Coluna5', 'BLOQ.'])}
                </div>
            </div>
        </div>
    </div>`;
    div.innerHTML = accordionHtml;
}
