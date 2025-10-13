// ================================================================================================
//  SCRIPT COMPLETO E OTIMIZADO
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
    // Carregar configurações e anexar todos os event listeners
    loadConfigurations();

    // -- Sidebar --
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

    // -- Busca --
    document.getElementById('searchBtn').addEventListener('click', buscarPedido);
    document.getElementById('searchInput').addEventListener('keyup', (e) => { if (e.key === 'Enter') buscarPedido(); });

    // -- Ações dos Modais --
    document.getElementById('exportarAtrasadosBtn').addEventListener('click', exportarPedidosAtrasados);
    document.getElementById('bloquearPedidoBtn').addEventListener('click', bloquearPedido);
    document.getElementById('marcarSemCorteBtn').addEventListener('click', marcarPedidosSemCorte);
    document.getElementById('montarCargaEspecialBtn').addEventListener('click', () => montarCargaPredefinida('pedidosEspeciaisInput', 'resultado-carga-especial', pedidosEspeciaisProcessados, 'Especial'));
    document.getElementById('montarVendaAntecipadaBtn').addEventListener('click', () => montarCargaPredefinida('vendaAntecipadaInput', 'resultado-venda-antecipada', pedidosVendaAntecipadaProcessados, 'Venda Antecipada'));

    // -- Configurações dos Veículos --
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

    // Usar setTimeout para permitir que a UI atualize antes do processamento pesado
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

            // 1. Separação inicial por tipo de pedido
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
            
            // 2. Separação por bloqueios manuais e pré-definidos (Venda Antecipada, Especial)
            let pedidosDisponiveis = [];
            dadosFiltrados.forEach(p => {
                const numPedidoStr = String(p.Num_Pedido);
                if (pedidosBloqueados.has(numPedidoStr)) {
                    pedidosManualmenteBloqueadosAtuais.push(p);
                } else if (!pedidosEspeciaisProcessados.has(numPedidoStr) && !pedidosVendaAntecipadaProcessados.has(numPedidoStr)) {
                    pedidosDisponiveis.push(p);
                }
            });

            // 3. Processamento de regras específicas (Rota 1, Toco, CF, etc.)
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

            // 4. Agrupamento final para exibição
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
            
            // Re-calcular totais dos grupos
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

            // 5. Renderizar tudo na tela
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
                            data-container-id="${item.containerId}"
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

    const adicionarResultado = (pedido, localizacao, tabId, containerId, accordionId) => {
        if (pedido && !adicionado.has(pedido.Num_Pedido) && matcher(pedido)) {
            resultados.push({ pedido, localizacao, tabId, containerId, accordionId });
            adicionado.add(pedido.Num_Pedido);
        }
    };
    
    // 1. Cargas ativas e sobras (mais provável)
    for (const loadId in activeLoads) {
        const load = activeLoads[loadId];
        const tabId = load.vehicleType === 'fiorino' ? 'fiorino-tab' : (load.vehicleType === 'van' ? 'van-tab' : 'tres-quartos-tab');
        load.pedidos.forEach(p => adicionarResultado(p, `Carga ${load.numero || ''}`, tabId, loadId));
    }
    currentLeftoversForPrinting.forEach(p => adicionarResultado(p, 'Sobras na Mesa de Trabalho', 'fiorino-tab', 'leftovers-card-resultado-fiorino-geral'));

    // 2. Grupos de Truck/Carreta/Toco
    Object.entries(gruposPorCFGlobais).forEach(([cf, grupo]) => {
        const accordionId = `collapseCF-Mesa-${String(cf).replace(/\s|\(|\)|\//g, '')}`;
        grupo.pedidos.forEach(p => adicionarResultado(p, `Carga Truck/Carreta (CF: ${cf})`, 'cf-tab', 'resultado-cf-accordion-container', accordionId));
    });
    Object.entries(gruposToco).forEach(([cf, grupo]) => {
        const collapseId = `collapseToco${Object.keys(gruposToco).sort().indexOf(cf)}`;
        grupo.pedidos.forEach(p => adicionarResultado(p, `Carga Toco (CF: ${cf})`, 'cf-tab', `toco-${cf}`, collapseId));
    });

    // 3. Listas Adicionais
    pedidosGeraisAtuais.forEach(p => {
        const accordionId = `collapseGeral${Object.keys(pedidosGeraisAtuais.reduce((acc, p) => { acc[p.Cod_Rota] = true; return acc; }, {})).sort().indexOf(p.Cod_Rota)}`;
        adicionarResultado(p, 'Pedidos Disponíveis', null, 'resultado-geral', accordionId);
    });
    pedidosFuncionarios.forEach(p => adicionarResultado(p, 'Pedidos de Funcionários', 'funcionarios-tab', 'resultado-funcionarios-tab', 'collapseFuncionários'));
    pedidosTransferencias.forEach(p => adicionarResultado(p, 'Pedidos de Transferência', 'transferencias-tab', 'resultado-transferencias-tab', 'collapseTransferências'));
    pedidosManualmenteBloqueadosAtuais.forEach(p => adicionarResultado(p, 'Bloqueados Manualmente', null, 'resultado-bloqueados'));
    rota1SemCarga.forEach(p => adicionarResultado(p, 'Rota 1 para Alteração', null, 'resultado-rota1'));
    pedidosComCFNumericoIsolado.forEach(p => adicionarResultado(p, 'Filtrados (Regra Bloqueio)', null, 'resultado-cf-numerico'));
    
    return resultados;
}

function highlightPedido(button) {
    const { pedidoId, tabId, accordionId } = button.dataset;

    const scrollToElement = (id) => {
        const row = document.getElementById(`pedido-${id}`);
        if (row) {
            row.scrollIntoView({ behavior: 'smooth', block: 'center' });
            row.classList.remove('search-highlight');
            void row.offsetWidth;
            row.classList.add('search-highlight');
        }
    };
    
    const navigateAndHighlight = () => {
        if (accordionId) {
            const collapseEl = document.getElementById(accordionId);
            if (collapseEl && !collapseEl.classList.contains('show')) {
                const bsCollapse = bootstrap.Collapse.getOrCreateInstance(collapseEl);
                collapseEl.addEventListener('shown.bs.collapse', () => scrollToElement(pedidoId), { once: true });
                bsCollapse.show();
            } else {
                 scrollToElement(pedidoId);
            }
        } else {
            scrollToElement(pedidoId);
        }
    };
    
    if (tabId) {
        const tabButton = document.getElementById(tabId);
        const activeTab = document.querySelector('#vehicleTabs .nav-link.active');
        if (tabButton && tabButton.id !== activeTab.id) {
            const bsTab = bootstrap.Tab.getOrCreateInstance(tabButton);
            // O listener precisa ser no container do conteúdo principal, não no botão
            document.getElementById('vehicleTabsContent').addEventListener('shown.bs.tab', (event) => {
                 if (event.target.id === `${tabId}-pane`) { // Bootstrap anexa o evento ao pane
                    navigateAndHighlight();
                }
            }, { once: true });
            bsTab.show();
        } else {
            navigateAndHighlight();
        }
    } else {
        navigateAndHighlight();
    }
}

// Lógica do Relatório de Atrasados (Melhorada)
function exportarPedidosAtrasados() {
    if (planilhaData.length === 0) { // Usar planilhaData como fonte da verdade
        alert("Por favor, processe a planilha primeiro.");
        return;
    }

    const pedidosAtrasados = planilhaData.filter(p => {
        const bloqueado = String(p['BLOQ.'] || '').trim().toUpperCase();
        // Incluir se: estiver atrasado E (não estiver bloqueado com qualquer valor OU estiver bloqueado com 'C')
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
    // Tentar normalizar datas que podem vir como string
    const date = predat instanceof Date ? predat : new Date(predat);
    if (isNaN(date)) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    date.setHours(0,0,0,0); // Normalizar a data do pedido também
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

// Funções de Renderização na UI
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

function createTable(pedidos, columnsToDisplay, sourceId = '') {
    if (!pedidos || pedidos.length === 0) return '<p class="text-muted p-3">Nenhum pedido nesta carga.</p>';
    const colunasExibir = columnsToDisplay || ['Cod_Rota', 'Cliente', 'Nome_Cliente', 'Agendamento', 'Num_Pedido', 'Quilos_Saldo', 'Cubagem', 'Cidade', 'Predat', 'BLOQ.', 'Coluna4', 'Coluna5', 'CF'];
    
    let table = '<div class="table-responsive"><table class="table table-sm table-bordered table-striped table-hover"><thead><tr>';
    colunasExibir.forEach(c => table += `<th>${c.replace('_', ' ')}</th>`);
    table += '</tr></thead><tbody>';
    
    pedidos.forEach(p => {
        const isPriorityRow = pedidosPrioritarios.has(String(p.Num_Pedido));
        const clienteIdNormalizado = normalizeClientId(p.Cliente);
        table += `<tr id="pedido-${p.Num_Pedido}" 
                        class="${isPriorityRow ? 'table-primary' : ''}" 
                        data-cliente-id="${clienteIdNormalizado}" 
                        data-pedido-id="${p.Num_Pedido}"
                        onclick="highlightClientRows(event)"
                        draggable="true"
                        ondragstart="dragStart(event, '${p.Num_Pedido}', '${clienteIdNormalizado}', '${sourceId}')">`;
        colunasExibir.forEach(c => {
            let cellContent = p[c] ?? '';
            let cellHtml = '';
            if (c === 'Num_Pedido') {
                const priorityBadge = isPriorityRow ? ' <span class="badge bg-primary">Prioridade</span>' : '';
                const semCorteBadge = pedidosSemCorte.has(String(p.Num_Pedido)) ? ' <span class="badge bg-transparent" title="Pedido Sem Corte"><i class="bi bi-scissors text-warning"></i></span>' : '';
                cellHtml = `<td>${cellContent}${priorityBadge}${semCorteBadge}</td>`;
            } else if (c === 'Agendamento' && cellContent === 'Sim') {
                cellHtml = `<td><span class="badge bg-warning text-dark">${cellContent}</span></td>`;
            } else if (c === 'Predat' || c === 'Dat_Ped') {
                const dateObj = cellContent instanceof Date ? cellContent : new Date(cellContent);
                const formattedDate = (dateObj instanceof Date && !isNaN(dateObj)) ? dateObj.toLocaleDateString('pt-BR', { timeZone: 'UTC' }) : cellContent;
                cellHtml = (c === 'Predat' && isOverdue(p.Predat)) ? `<td><span class="text-danger fw-bold">${formattedDate}</span></td>` : `<td>${formattedDate}</td>`;
            } else if (c === 'Quilos_Saldo' || c === 'Cubagem') {
                cellHtml = `<td>${(cellContent || 0).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>`;
            } else {
                cellHtml = `<td>${cellContent}</td>`;
            }
            table += cellHtml;
        });
        table += '</tr>';
    });
    table += '</tbody></table></div>';
    return table;
}

// ... Colar o restante das funções originais (sem modificações) aqui ...
// As funções abaixo são do seu código original e devem funcionar com as melhorias já implementadas acima.

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
    const numeros = input.value.split(/[\n\s,;]+/).map(n => n.trim()).filter(Boolean);
    numeros.forEach(num => pedidosSemCorte.add(num));
    input.value = '';
    atualizarListaSemCorte();
}

function removerMarcacaoSemCorte(numPedido) {
    pedidosSemCorte.delete(numPedido);
    atualizarListaSemCorte();
}

// Funções de renderização de seções específicas
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
            botoes[vehicleType] += `<button class="btn btn-outline-${colorClass} mt-2 me-2" onclick="${functionCall}">${config.title}</button>`;
        }
    });
    document.getElementById('botoes-fiorino').innerHTML = botoes.fiorino || '<div class="empty-state"><i class="bi bi-box"></i><p>Nenhuma rota de Fiorino encontrada.</p></div>';
    document.getElementById('botoes-van').innerHTML = botoes.van || '<div class="empty-state"><i class="bi bi-truck-front-fill"></i><p>Nenhuma rota de Van encontrada.</p></div>';
    document.getElementById('botoes-34').innerHTML = botoes.tresQuartos || '<div class="empty-state"><i class="bi bi-truck-flatbed"></i><p>Nenhuma rota de 3/4 encontrada.</p></div>';
    
    let accordionHtml = '<div class="accordion accordion-flush" id="accordionGeral">';
    Object.keys(grupos).sort().forEach((rota, index) => {
        const grupo = grupos[rota];
        const veiculo = rotaVeiculoMap[rota]?.type || 'N/D';
        const veiculoNome = veiculo.replace('tresQuartos', '3/4').replace(/^\w/, c => c.toUpperCase());
        accordionHtml += `<div class="accordion-item"><h2 class="accordion-header" id="headingGeral${index}"><button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseGeral${index}"><strong>Rota: ${rota} (${veiculoNome})</strong> &nbsp; <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${grupo.pedidos.length}</span> <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${grupo.totalKg.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})} kg</span></button></h2><div id="collapseGeral${index}" class="accordion-collapse collapse" data-bs-parent="#accordionGeral"><div class="accordion-body">${createTable(grupo.pedidos, null, 'geral')}</div></div></div>`;
    });
    accordionHtml += '</div>'; 
    div.innerHTML = accordionHtml;
}

function displayPedidosBloqueados(div, pedidos) {
    if (pedidos.length === 0) {
        div.innerHTML = '<div class="empty-state"><i class="bi bi-shield-check"></i><p>Nenhum pedido bloqueado manualmente.</p></div>';
        return;
    }
    const totalKg = pedidos.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
    let html = `<div class="alert alert-danger"><strong>Total Bloqueado:</strong> ${pedidos.length} pedidos / ${totalKg.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})} kg</div>`;
    html += createTable(pedidos);
    div.innerHTML = html;
}

function displayPedidosCFNumerico(div, pedidos) {
    if (pedidos.length === 0) { 
        div.innerHTML = '<div class="empty-state"><i class="bi bi-funnel"></i><p>Nenhum pedido de varejo filtrado por regra de bloqueio.</p></div>'; 
        return; 
    }
    const grupos = pedidos.reduce((acc, p) => {
        const rota = p.Cod_Rota || 'Sem Rota'; 
        if (!acc[rota]) { acc[rota] = { pedidos: [], totalKg: 0 }; } 
        acc[rota].pedidos.push(p); 
        acc[rota].totalKg += p.Quilos_Saldo; 
        return acc;
    }, {});
    let accordionHtml = '<div class="accordion accordion-flush" id="accordionCF">';
    Object.keys(grupos).sort().forEach((rota, index) => {
        const grupo = grupos[rota];
        accordionHtml += `<div class="accordion-item"><h2 class="accordion-header" id="headingCF${index}"><button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseCF${index}"><strong>Rota: ${rota}</strong> &nbsp; <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${grupo.pedidos.length}</span> <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} kg</span></button></h2><div id="collapseCF${index}" class="accordion-collapse collapse" data-bs-parent="#accordionCF"><div class="accordion-body">${createTable(grupo.pedidos)}</div></div></div>`;
    });
    accordionHtml += '</div>'; 
    div.innerHTML = accordionHtml;
}

function displayRota1(div, pedidos) {
    if (!pedidos || pedidos.length === 0) {
        div.innerHTML = '<div class="empty-state"><i class="bi bi-check-circle"></i><p>Nenhum pedido da Rota 1 para alteração encontrado.</p></div>';
        return;
    }
    let html = `
        <div class="d-flex justify-content-end mb-2 no-print">
            <button class="btn btn-sm btn-outline-warning" onclick="imprimirGeneric('resultado-rota1', 'Pedidos Rota 1 para Alteração')">
                <i class="bi bi-printer-fill me-1"></i>Imprimir Lista
            </button>
        </div>
        ${createTable(pedidos, ['Num_Pedido', 'Cliente', 'Nome_Cliente', 'Quilos_Saldo', 'Cidade', 'Predat', 'CF', 'Coluna5'])}
    `;
    div.innerHTML = html;
}

// ... (O resto das suas funções originais: impressão, algoritmos de otimização, drag & drop, etc.)

// A PARTIR DAQUI, COLE O RESTANTE DAS SUAS FUNÇÕES ORIGINAIS
// Por exemplo:
// function createPrintWindow(title) { ... }
// function imprimirGeneric(divId, title) { ... }
// ... e todas as outras funções até o final do seu arquivo.

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
    const load = activeLoads[loadId];
    const cardToPrint = document.getElementById(loadId);

    if (!load || !cardToPrint) {
        alert(`Erro: Carga com ID ${loadId} não encontrada para impressão.`);
        return;
    }

    const title = `Impressão - Carga Manual ${load.numero} (${load.vehicleType})`;
    const printWindow = createPrintWindow(title);
    
    printWindow.document.body.innerHTML = `<h3>${title}</h3>` + cardToPrint.outerHTML;
    printWindow.document.close();
    printWindow.focus(); 
    setTimeout(() => { 
        printWindow.print(); 
        printWindow.close(); 
    }, 500);
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

// Lógica de Otimização e Algoritmos...
function getVehicleConfig(vehicleType) {
    const configs = {
        minKg: parseFloat(document.getElementById(`${vehicleType}MinCapacity`).value),
        softMaxKg: parseFloat(document.getElementById(`${vehicleType}MaxCapacity`).value),
        softMaxCubage: parseFloat(document.getElementById(`${vehicleType}Cubage`).value),
        hardMaxKg: parseFloat(document.getElementById(`${vehicleType}HardMaxCapacity`)?.value || document.getElementById(`${vehicleType}MaxCapacity`).value),
        hardMaxCubage: parseFloat(document.getElementById(`${vehicleType}HardCubage`)?.value || document.getElementById(`${vehicleType}Cubage`).value),
    };
    return configs;
}

function isMoveValid(load, groupToAdd, vehicleType) {
    const config = getVehicleConfig(vehicleType);

    if ((load.totalKg + groupToAdd.totalKg) > config.hardMaxKg) return false;
    if ((load.totalCubagem + groupToAdd.totalCubagem) > config.hardMaxCubage) return false;

    if (groupToAdd.isSpecial) {
        const specialClientIdsInLoad = new Set(
            load.pedidos
                .filter(isSpecialClient)
                .map(p => normalizeClientId(p.Cliente))
        );
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
        { name: 'priority-weight-desc',
          sorter: (a, b) => {
            const aHasPrio = a.pedidos.some(p => pedidosPrioritarios.has(String(p.Num_Pedido)));
            const bHasPrio = b.pedidos.some(p => pedidosPrioritarios.has(String(p.Num_Pedido)));
            if (aHasPrio && !bHasPrio) return -1;
            if (!aHasPrio && bHasPrio) return 1;
            return b.totalKg - a.totalKg;
          }
        },
        { name: 'scheduled-weight-desc',
          sorter: (a, b) => {
            const aHasSched = a.pedidos.some(p => p.Agendamento === 'Sim');
            const bHasSched = b.pedidos.some(p => p.Agendamento === 'Sim');
            if (aHasSched && !bHasSched) return -1;
            if (!aHasSched && bHasSched) return 1;
            return b.totalKg - a.totalKg;
          }
        },
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
        const hasPriority = load.pedidos.some(p => pedidosPrioritarios.has(String(p.Num_Pedido)));
        const allowPriorityOverride = vehicleType !== 'tresQuartos';
        
        if (load.pedidos.length > 0 && (load.totalKg >= config.minKg || (hasPriority && allowPriorityOverride))) {
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
    
    const leftovers = [...leftoverItems, ...unplacedGroups];
    return { loads: finalLoads, leftovers };
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
    let totalSobras = 0;
    totalSobras += solution.leftovers.reduce((sum, group) => sum + group.totalKg, 0);
    solution.loads.forEach(load => {
        if (load.totalKg > 0 && load.totalKg < config.minKg) {
            totalSobras += load.totalKg;
        }
    });
    return totalSobras;
}

// ... E assim por diante, cole aqui todas as outras funções originais que faltam:
// - runSimulatedAnnealing
// - refinarCargasComTrocas
// - separarCargasGeneric (e as específicas fiorino, van, etc)
// - e todas as demais até o final do arquivo original.

// Cole o restante do seu JavaScript original aqui.
// Esta seção foi omitida para brevidade, mas você deve incluir todas as suas funções
// de `runSimulatedAnnealing` até o final do arquivo original.
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
            btn.classList.remove('active', 'btn-success', 'btn-primary', 'btn-warning', 'btn-secondary');
            const originalColorClass = btn.id.includes('fiorino') ? 'success' : (btn.id.includes('van') ? 'primary' : 'warning');
            btn.classList.add(`btn-outline-${originalColorClass}`);
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

    packableGroups.forEach(group => {
        group.Quilos_Saldo = group.totalKg;
        group.Cubagem = group.totalCubagem;
        if (group.totalCubagem > 0) group.density = group.totalKg / group.totalCubagem;
        else group.density = Infinity; 
    });

    const optimizationLevel = document.getElementById('optimizationLevelSelect').value;
    let optimizationResult;
    
    const modalElement = document.getElementById('processing-modal');
    const modal = bootstrap.Modal.getOrCreateInstance(modalElement);
    
    try {
        if(optimizationLevel !== '1') {
            document.getElementById('processing-progress-bar').style.width = '0%';
            modal.show();
        } else {
            resultadoDiv.insertAdjacentHTML('beforeend', '<div id="spinner-temp-container" class="d-flex align-items-center justify-content-center p-5"><div class="spinner-border text-primary" role="status"></div><span class="ms-3">Analisando estratégias e montando cargas...</span></div>');
        }

        switch (optimizationLevel) {
            case '1':
                console.log(`Executando Nível 1: Heurística Rápida para ${vehicleType}...`);
                optimizationResult = runHeuristicOptimization(packableGroups, vehicleType);
                break;
            case '3':
                console.log(`Executando Nível 3: Otimização Exaustiva (SA + Polimento Simples) para ${vehicleType}...`);
                const saResultForPolish = await runSimulatedAnnealing(packableGroups, vehicleType, 'Otimizando... (Nível 3 - Fase 1/2)');
                
                document.getElementById('processing-status-text').textContent = 'Otimizando... (Nível 3 - Fase 2/2)';
                document.getElementById('processing-details-text').textContent = 'Aplicando polimento com trocas locais.';
                document.getElementById('processing-progress-bar').style.width = '100%';

                const polished = refinarCargasComTrocas(saResultForPolish.loads, saResultForPolish.leftovers, vehicleType);
                optimizationResult = { loads: polished.refinedLoads, leftovers: polished.remainingLeftovers };
                break;
            case '4':
                console.log(`Executando Nível 4: Otimização Reconstrutiva para ${vehicleType}...`);
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
                console.log(`Executando Nível 2: Otimização Avançada (SA) para ${vehicleType}...`);
                optimizationResult = await runSimulatedAnnealing(packableGroups, vehicleType, 'Otimizando... (Nível 2)');
                break;
        }
    } finally {
        allRouteButtons.forEach(btn => btn.disabled = false);
        if(optimizationLevel !== '1') {
            modal.hide();
        } else {
            const spinner = document.getElementById('spinner-temp-container');
            if (spinner) spinner.remove();
        }
    }

    optimizationResult.loads.forEach(load => { load.vehicleType = vehicleType; });
    const { refinedLoads, remainingLeftovers } = refineLoadsWithSimpleFit(optimizationResult.loads, optimizationResult.leftovers);

    let primaryLoads = refinedLoads;
    let leftoverGroups = remainingLeftovers;
    let secondaryLoads = [];
    let tertiaryLoads = [];

    if (vehicleType === 'fiorino' && leftoverGroups.length > 0) {
        const fiorinoLeftoversResult = runHeuristicOptimization(leftoverGroups, 'fiorino'); 
        if (fiorinoLeftoversResult.loads.length > 0) {
            fiorinoLeftoversResult.loads.forEach(l => l.vehicleType = 'fiorino');
            primaryLoads.push(...fiorinoLeftoversResult.loads);
        }
        leftoverGroups = fiorinoLeftoversResult.leftovers;

        const totalLeftoverKg = leftoverGroups.reduce((sum, g) => sum + g.totalKg, 0);
        const vanMin = parseFloat(document.getElementById('vanMinCapacity').value);
        if (totalLeftoverKg >= vanMin && leftoverGroups.length > 0) {
            const vanResult = runHeuristicOptimization(leftoverGroups, 'van');
            vanResult.loads.forEach(l => l.vehicleType = 'van');
            secondaryLoads = vanResult.loads;
            leftoverGroups = vanResult.leftovers;
        }
    }
    
    const totalFinalLeftoverKg = leftoverGroups.reduce((sum, g) => sum + g.totalKg, 0);
    const tresQuartosMin = parseFloat(document.getElementById('tresQuartosMinCapacity').value);
    if (totalFinalLeftoverKg >= tresQuartosMin && leftoverGroups.length > 0) {
                                         const vehicleForTertiary = (vehicleType === 'fiorino' || vehicleType === 'van') ? 'tresQuartos' : '';
                                         if(vehicleForTertiary) {
                                             const tresQuartosResult = runHeuristicOptimization(leftoverGroups, vehicleForTertiary);
                                             tresQuartosResult.loads.forEach(l => l.vehicleType = 'tresQuartos');
                                             tertiaryLoads = tresQuartosResult.loads;
                                             leftoverGroups = tresQuartosResult.leftovers;
                                         }
    }

    const allPotentialLoads = [ ...primaryLoads, ...secondaryLoads, ...tertiaryLoads ];
    const finalValidLoads = [];
    let finalLeftoverGroups = [...leftoverGroups];

    allPotentialLoads.forEach(load => {
        if (!load.vehicleType) {
            console.warn("Carga encontrada sem tipo de veículo, não é possível validar o peso mínimo.", load);
            finalValidLoads.push(load);
            return;
        }
        const config = getVehicleConfig(load.vehicleType);
        const hasPriority = load.pedidos.some(p => pedidosPrioritarios.has(String(p.Num_Pedido)));
        const allowPriorityOverride = load.vehicleType !== 'tresQuartos';
        
        if (load.totalKg >= config.minKg || (hasPriority && allowPriorityOverride)) {
            finalValidLoads.push(load);
        } else {
            const clientGroupsInFailedLoad = Object.values(load.pedidos.reduce((acc, p) => {
                const clienteId = normalizeClientId(p.Cliente);
                if (!acc[clienteId]) { acc[clienteId] = { pedidos: [], totalKg: 0, totalCubagem: 0, isSpecial: isSpecialClient(p) }; }
                acc[clienteId].pedidos.push(p);
                acc[clienteId].totalKg += p.Quilos_Saldo;
                acc[clienteId].totalCubagem += p.Cubagem;
                return acc;
            }, {}));
            finalLeftoverGroups.push(...clientGroupsInFailedLoad);
        }
    });

    currentLeftoversForPrinting = finalLeftoverGroups.flatMap(group => group.pedidos);

    finalValidLoads.forEach((load, index) => {
        load.numero = `${load.vehicleType.charAt(0).toUpperCase()}${index + 1}`;
        const loadId = `${load.vehicleType}-${Date.now()}-${index}`;
        load.id = loadId;
        activeLoads[loadId] = load;
    });

    const vehicleInfo = {
        fiorino: { name: 'Fiorino', colorClass: 'bg-success', textColor: 'text-white', icon: 'bi-box-seam-fill' },
        van: { name: 'Van', colorClass: 'bg-primary', textColor: 'text-white', icon: 'bi-truck-front-fill' },
        tresQuartos: { name: '3/4', colorClass: 'bg-warning', textColor: 'text-dark', icon: 'bi-truck-flatbed' }
    };
    
    let html = `<h5 class="mt-3">Cargas para <strong>${title}</strong></h5>`;
    
    const generatedLoads = finalValidLoads.filter(l => l.pedidos.length > 0);
    if(generatedLoads.length === 0){
         html += `<div class="alert alert-secondary">Nenhuma carga foi formada para esta rota.</div>`;
    } else {
        generatedLoads.forEach(load => {
            html += renderLoadCard(load, load.vehicleType, vehicleInfo[load.vehicleType]);
        });
    }
    
    if (currentLeftoversForPrinting.length > 0) {
        const finalLeftoverKg = currentLeftoversForPrinting.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
        const manualLoadButton = `<button id="start-manual-load-btn" class="btn btn-success ms-auto no-print" onclick="startManualLoadBuilder()"><i class="bi bi-plus-circle-fill me-1"></i>Criar Carga Manual</button>`;
        const printButtonHtml = `<button class="btn btn-info ms-2 no-print" onclick="imprimirSobras('Sobras Finais de ${title}')"><i class="bi bi-printer-fill me-1"></i>Imprimir</button>`;

        html += `<div id="leftovers-card-${divId}" class="drop-zone-card" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event)" data-load-id="leftovers" data-vehicle-type="leftovers">
                                     <h5 class="mt-4">Sobras Finais: ${finalLeftoverKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} kg</h5>
                                     <div class="card mb-3">
                                         <div class="card-header bg-danger text-white d-flex align-items-center">
                                             Pedidos Restantes
                                             <div class="ms-auto">${manualLoadButton}${printButtonHtml}</div>
                                         </div>
                                         <div class="card-body">${createTable(currentLeftoversForPrinting, ['Num_Pedido', 'Quilos_Saldo', 'Agendamento', 'Cubagem', 'Predat', 'Cliente', 'Nome_Cliente', 'Cidade', 'CF'], 'leftovers')}</div>
                                     </div>
                                 </div>`;
    }
    resultadoDiv.innerHTML = `<div class="resultado-container">${html}</div>`;
    updateAndRenderChart();
}

// O restante do seu código JS, que não foi modificado, deve ser inserido aqui.
// ... (continue colando o restante das suas funções originais)
// Esta é uma lista de funções que você deve garantir que estão no seu arquivo final:
// - refinarCargasComTrocas
// - refinarComReconstrucao
// - renderLoadCard
// - displayToco
// - displayCargasCfAccordion
// - montarCargaPredefinida
// - highlightClientRows
// - dragStart, dragOver, dragLeave, drop
// - updateLoadUI
// - startManualLoadBuilder, updateManualBuilderUI, finalizeManualLoad, cancelManualLoad
// - updateAndRenderChart

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

async function refinarComReconstrucao(initialLoads, initialLeftovers, vehicleType) {
    let loads = deepClone(initialLoads);
    let leftovers = deepClone(initialLeftovers);

    if (loads.length < 2) {
        console.log("POLIMENTO (Nível 4): Poucas cargas para reconstruir. Pulando etapa.");
        return { refinedLoads: loads, remainingLeftovers: leftovers };
    }

    
    loads.sort((a, b) => a.totalKg - b.totalKg);
    const worstLoad = loads[0]; 

    const config = getVehicleConfig(vehicleType);
    if (worstLoad.totalKg >= config.softMaxKg) {
        console.log("POLIMENTO (Nível 4): A carga menos cheia já está bem otimizada. Pulando etapa.");
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

    const newLeftovers = [...leftovers, ...groupsToReallocate];
    const remainingLoads = loads.slice(1);

    
    const { refinedLoads: reconstructedLoads, remainingLeftovers: finalLeftovers } = refineLoadsWithSimpleFit(remainingLoads, newLeftovers);
    
    const originalSobras = initialLeftovers.reduce((sum, g) => sum + g.totalKg, 0);
    const newSobras = finalLeftovers.reduce((sum, g) => sum + g.totalKg, 0);

    if (newSobras < originalSobras) {
        console.log(`POLIMENTO (Nível 4): Reconstrução bem-sucedida! Sobra reduzida de ${originalSobras.toFixed(2)}kg para ${newSobras.toFixed(2)}kg.`);
        return { refinedLoads: reconstructedLoads, remainingLeftovers: finalLeftovers };
    } else {
        console.log("POLIMENTO (Nível 4): Reconstrução não melhorou o resultado. Revertendo.");
        return { refinedLoads: initialLoads, remainingLeftovers: initialLeftovers };
    }
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
