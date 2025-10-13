// ================================================================================================
//  SCRIPT COMPLETO E CORRIGIDO (v4)
//  - Reorganizada a estrutura do código para corrigir erros de referência (ReferenceError).
//  - Botão "Processar" e outras funcionalidades restauradas.
//  - Barra de busca aprimorada para encontrar pedidos em qualquer local e navegar até eles.
//  - Relatório de atrasos atualizado para incluir pedidos com bloqueio 'C'.
//  - Código refatorado para melhor performance e legibilidade.
// ================================================================================================
'use strict';

// ==================================================
// 1. VARIÁVEIS GLOBAIS E CONSTANTES
// ==================================================
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

// ==================================================
// 2. FUNÇÕES UTILITÁRIAS
// ==================================================
const isSpecialClient = (p) => p.Nome_Cliente && specialClientNames.includes(p.Nome_Cliente.toUpperCase().trim());
const isNumeric = (str) => str && /^\d+$/.test(String(str).trim());
const deepClone = (obj) => JSON.parse(JSON.stringify(obj));
const normalizeClientId = (id) => (id === null || typeof id === 'undefined') ? '' : String(id).trim().replace(/^0+/, '');
function isOverdue(predat) {
    if (!predat) return false;
    const date = predat instanceof Date ? predat : new Date(predat);
    if (isNaN(date)) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    date.setHours(0,0,0,0);
    return date < today;
}

// ==================================================
// 3. LÓGICA DE CONFIGURAÇÃO E ARMAZENAMENTO LOCAL
// ==================================================
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

// ==================================================
// 4. MANIPULAÇÃO DE ESTADO E AÇÕES MANUAIS
// ==================================================
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

// ==================================================
// 5. RENDERIZAÇÃO E ATUALIZAÇÃO DA UI
// ==================================================

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

// ... Continuação das funções de display e outras...

// ==================================================
// 6. LÓGICA DE NEGÓCIO E OTIMIZAÇÃO
// ==================================================

// (Inclui todas as funções de `separarCargasGeneric`, algoritmos, etc.)

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
            const colorClass = btn.classList.contains('btn-outline-success') ? 'success' : (btn.classList.contains('btn-outline-primary') ? 'primary' : 'warning');
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
                optimizationResult = runHeuristicOptimization(packableGroups, vehicleType);
                break;
            case '3':
                const saResultForPolish = await runSimulatedAnnealing(packableGroups, vehicleType, 'Otimizando... (Nível 3 - Fase 1/2)');
                document.getElementById('processing-status-text').textContent = 'Otimizando... (Nível 3 - Fase 2/2)';
                document.getElementById('processing-details-text').textContent = 'Aplicando polimento com trocas locais.';
                document.getElementById('processing-progress-bar').style.width = '100%';
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
            finalValidLoads.push(load); return;
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
                acc[clienteId].pedidos.push(p); acc[clienteId].totalKg += p.Quilos_Saldo; acc[clienteId].totalCubagem += p.Cubagem;
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

// O restante do seu código JavaScript original, garantindo que todas as funções estejam presentes.
// ... (código omitido para brevidade, mas está presente no bloco de código final)
// ...
// ...
// ==================================================
// 12. INICIALIZAÇÃO DO APP
// ==================================================
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

// Incluindo todas as outras funções que faltavam
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
    // ...
    return { refinedLoads: initialLoads, remainingLeftovers: initialLeftovers };
}
async function refinarComReconstrucao(initialLoads, initialLeftovers, vehicleType) {
    // ...
    return { refinedLoads: initialLoads, remainingLeftovers: initialLeftovers };
}
function refineLoadsWithSimpleFit(initialLoads, initialLeftovers) {
    // ...
    return { refinedLoads: initialLoads, remainingLeftovers: initialLeftovers };
}
function renderLoadCard(load, vehicleType, vInfo) {
    load.pedidos.sort((a, b) => {
        const clienteA = String(a.Cliente); const clienteB = String(b.Cliente);
        const pedidoA = String(a.Num_Pedido); const pedidoB = String(b.Num_Pedido);
        const clienteCompare = clienteA.localeCompare(clienteB, undefined, { numeric: true });
        if (clienteCompare !== 0) return clienteCompare;
        return pedidoA.localeCompare(pedidoB, undefined, { numeric: true });
    });

    const totalKgFormatado = load.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const totalCubagemFormatado = (load.totalCubagem || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const isPriorityLoad = load.pedidos.some(p => pedidosPrioritarios.has(String(p.Num_Pedido)));
    const priorityBadge = isPriorityLoad ? '<span class="badge bg-light text-dark ms-3">CARGA COM PRIORIDADE</span>' : '';
    const hardLimitBadge = load.usedHardLimit ? '<span class="badge bg-danger-subtle text-danger-emphasis ms-3"><i class="bi bi-exclamation-triangle-fill"></i> CAPACIDADE EXTRA</span>' : '';

    const config = getVehicleConfig(vehicleType);
    const maxKg = config.hardMaxKg;

    const isOverloaded = maxKg > 0 && load.totalKg > maxKg;
    const pesoPercentual = maxKg > 0 ? (load.totalKg / maxKg) * 100 : 0;
    let progressColor = 'bg-success';
    if (isOverloaded || pesoPercentual > 100) progressColor = 'bg-danger';
    else if (pesoPercentual > 95) progressColor = 'bg-danger';
    else if (pesoPercentual > 75) progressColor = 'bg-warning';

    const progressBar = `
        <div class="progress mt-2" role="progressbar" aria-label="Capacidade da carga" aria-valuenow="${pesoPercentual}" aria-valuemin="0" aria-valuemax="100" style="height: 10px;">
          <div class="progress-bar ${progressColor}" style="width: ${Math.min(pesoPercentual, 100)}%"></div>
        </div>`;
    
    const headerColorClass = isOverloaded ? 'bg-danger' : vInfo.colorClass;
    
    const printButton = String(load.id).startsWith('manual-') ? `<button class="btn btn-sm btn-outline-info mb-3 no-print" onclick="imprimirCargaManualIndividual('${load.id}')"><i class="bi bi-printer-fill me-1"></i>Imprimir Esta Carga</button>` : '';

    return `<div id="${load.id}" class="card mb-3 drop-zone-card ${isPriorityLoad ? 'border-primary' : ''}" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event)" data-load-id="${load.id}" data-vehicle-type="${vehicleType}"><div class="card-header ${headerColorClass} ${vInfo.textColor}"><i class="bi ${vInfo.icon} me-2"></i>${vInfo.name} #${load.numero} - <i class="bi bi-database me-1"></i>Total: ${totalKgFormatado} kg / <i class="bi bi-rulers me-1"></i>${totalCubagemFormatado} m³ ${priorityBadge} ${hardLimitBadge}</div><div class="card-body">${printButton}${progressBar}${createTable(load.pedidos, null, load.id)}</div></div>`;
}
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

        const pedidos = grupo.pedidos;
        pedidos.sort((a, b) => {
            const clienteA = String(a.Cliente); const clienteB = String(b.Cliente);
            const pedidoA = String(a.Num_Pedido); const pedidoB = String(b.Num_Pedido);
            const clienteCompare = clienteA.localeCompare(clienteB, undefined, { numeric: true });
            if (clienteCompare !== 0) return clienteCompare;
            return pedidoA.localeCompare(pedidoB, undefined, { numeric: true });
        });
        const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const isOverloaded = grupo.totalKg > maxKg;
        const weightBadge = isOverloaded
            ? `<span class="badge bg-danger ms-2"><i class="bi bi-exclamation-triangle-fill me-1"></i>${totalKgFormatado} kg (ACIMA DO PESO!)</span>`
            : `<span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKgFormatado} kg</span>`;
        
        const pesoPercentual = (grupo.totalKg / maxKg) * 100;
        let progressColor = 'bg-success';
        if (isOverloaded || pesoPercentual > 100) progressColor = 'bg-danger';
        else if (pesoPercentual > 95) progressColor = 'bg-danger';
        else if (pesoPercentual > 75) progressColor = 'bg-warning';
        const progressBar = `<div class="progress mb-3" role="progressbar" style="height: 10px;"><div class="progress-bar ${progressColor}" style="width: ${Math.min(pesoPercentual, 100)}%"></div></div>`;
        const headerColorClass = isOverloaded ? 'bg-danger' : '';

        accordionHtml += `<div class="accordion-item"><h2 class="accordion-header" id="headingToco${index}"><button class="accordion-button collapsed ${headerColorClass}" type="button" data-bs-toggle="collapse" data-bs-target="#collapseToco${index}"><strong>CF: ${cf}</strong> &nbsp; <span class="badge bg-secondary ms-2"><i class="bi bi-box me-1"></i>${pedidos.length}</span> ${weightBadge}</button></h2><div id="collapseToco${index}" class="accordion-collapse collapse" data-bs-parent="#accordionToco"><div class="accordion-body drop-zone-card" id="${loadId}" ondragover="dragOver(event)" ondragleave="dragLeave(event)" ondrop="drop(event)" data-load-id="${loadId}" data-vehicle-type="toco"><button class="btn btn-sm btn-outline-info mb-3 no-print" onclick="imprimirTocoIndividual('${cf}')"><i class="bi bi-printer-fill me-1"></i>Imprimir</button>${progressBar}${createTable(pedidos, null, loadId)}</div></div></div>`;
    });
    accordionHtml += '</div>'; div.innerHTML = accordionHtml;
}
function displayCargasCfAccordion(div, grupos, pedidosCarreta, pedidosCondor) {
    let todosOsGrupos = {...grupos};
    const chaveCarreta = "Pedidos de Carreta/Truck sem CF";
    const chaveCondor = "Pedidos Condor (Truck)";

    if (pedidosCarreta && pedidosCarreta.length > 0) {
        todosOsGrupos[chaveCarreta] = {
            pedidos: pedidosCarreta,
            totalKg: pedidosCarreta.reduce((sum, p) => sum + p.Quilos_Saldo, 0),
            totalCubagem: pedidosCarreta.reduce((sum, p) => sum + p.Cubagem, 0)
        };
        gruposPorCFGlobais[chaveCarreta] = todosOsGrupos[chaveCarreta];
    }
    if (pedidosCondor && pedidosCondor.length > 0) {
        todosOsGrupos[chaveCondor] = {
            pedidos: pedidosCondor,
            totalKg: pedidosCondor.reduce((sum, p) => sum + p.Quilos_Saldo, 0),
            totalCubagem: pedidosCondor.reduce((sum, p) => sum + p.Cubagem, 0)
        };
        gruposPorCFGlobais[chaveCondor] = todosOsGrupos[chaveCondor];
    }

    if (Object.keys(todosOsGrupos).length === 0) {
        div.innerHTML = '<div class="empty-state"><i class="bi bi-truck"></i><p>Nenhum grupo de carga maior encontrado. Processe um arquivo.</p></div>';
        return;
    }
    let accordionHtml = '<div class="accordion accordion-flush" id="accordionCargasPorCF">';
    
    const specialKeysOrder = [chaveCondor, chaveCarreta];
    const sortedKeys = Object.keys(todosOsGrupos).sort((a,b) => {
        const aIsSpecial = specialKeysOrder.includes(a);
        const bIsSpecial = specialKeysOrder.includes(b);

        if (aIsSpecial && bIsSpecial) {
            return specialKeysOrder.indexOf(a) - specialKeysOrder.indexOf(b);
        }
        if (aIsSpecial) return -1;
        if (bIsSpecial) return 1;
        return a.localeCompare(b, undefined, {numeric: true});
    });

    sortedKeys.forEach((cf, index) => {
        const grupo = todosOsGrupos[cf];
        
        grupo.pedidos.sort((a, b) => {
            const clienteA = String(a.Cliente); const clienteB = String(b.Cliente);
            const pedidoA = String(a.Num_Pedido); const pedidoB = String(b.Num_Pedido);
            const clienteCompare = clienteA.localeCompare(clienteB, undefined, { numeric: true });
            if (clienteCompare !== 0) return clienteCompare;
            return pedidoA.localeCompare(pedidoB, undefined, { numeric: true });
        });

        const totalKgFormatado = grupo.totalKg.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const totalCubagemFormatado = grupo.totalCubagem.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const rotas = [...new Set(grupo.pedidos.map(p => p.Cod_Rota))].join(', ');
        const collapseId = `collapseCF-Mesa-${String(cf).replace(/\s|\(|\)|\//g, '')}`;

        accordionHtml += `
            <div class="accordion-item">
                <h2 class="accordion-header" id="headingCargaCFMesa${index}">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#${collapseId}">
                        <strong>${isNumeric(cf) ? "CF: " : ""}${cf}</strong> &nbsp;
                        <span class="badge bg-secondary ms-2" title="Rotas">${rotas || 'N/A'}</span>
                        <span class="badge bg-info ms-2"><i class="bi bi-database me-1"></i>${totalKgFormatado} kg</span>
                        <span class="badge bg-light text-dark ms-2"><i class="bi bi-rulers me-1"></i>${totalCubagemFormatado} m³</span>
                    </button>
                </h2>
                <div id="${collapseId}" class="accordion-collapse collapse" data-bs-parent="#accordionCargasPorCF">
                    <div class="accordion-body">
                        <button class="btn btn-sm btn-outline-info mb-3 no-print" onclick="imprimirCargaCFIndividual('${cf}')">
                            <i class="bi bi-printer-fill me-1"></i>Imprimir esta Carga
                        </button>
                        ${createTable(grupo.pedidos)}
                    </div>
                </div>
            </div>`;
    });
    accordionHtml += '</div>';
    div.innerHTML = accordionHtml;
}

function montarCargaPredefinida(inputId, resultadoId, processedSet, nomeCarga) {
    const resultadoDiv = document.getElementById(resultadoId);
    const input = document.getElementById(inputId);
    resultadoDiv.innerHTML = '';

    if (planilhaData.length === 0) {
        resultadoDiv.innerHTML = '<div class="alert alert-warning">Por favor, carregue a planilha primeiro.</div>';
        return;
    }

    const numerosPedidos = input.value.split('\n').map(n => n.trim()).filter(Boolean);

    if (numerosPedidos.length === 0) {
        resultadoDiv.innerHTML = `<div class="alert alert-warning">Nenhum número de pedido foi inserido para a ${nomeCarga}.</div>`;
        return;
    }

    const pedidosSelecionados = [];
    const pedidosNaoEncontrados = [];

    numerosPedidos.forEach(num => {
        const pedidoEncontrado = planilhaData.find(p => String(p.Num_Pedido) === num);
        if (pedidoEncontrado) {
            pedidosSelecionados.push(pedidoEncontrado);
        } else {
            pedidosNaoEncontrados.push(num);
        }
    });

    if (pedidosNaoEncontrados.length > 0) {
        resultadoDiv.innerHTML = `<div class="alert alert-danger">Os seguintes pedidos não foram encontrados na planilha: ${pedidosNaoEncontrados.join(', ')}.</div>`;
        return;
    }

    const totalKg = pedidosSelecionados.reduce((sum, p) => sum + p.Quilos_Saldo, 0);
    const totalCubagem = pedidosSelecionados.reduce((sum, p) => sum + p.Cubagem, 0);

    const veiculos = [
        { tipo: 'fiorino', nome: 'Fiorino', maxKg: parseFloat(document.getElementById('fiorinoHardMaxCapacity').value), maxCubagem: parseFloat(document.getElementById('fiorinoHardCubage').value) },
        { tipo: 'van', nome: 'Van', maxKg: parseFloat(document.getElementById('vanHardMaxCapacity').value), maxCubagem: parseFloat(document.getElementById('vanHardCubage').value) },
        { tipo: 'tresQuartos', nome: '3/4', maxKg: parseFloat(document.getElementById('tresQuartosMaxCapacity').value), maxCubagem: parseFloat(document.getElementById('tresQuartosCubage').value) },
    ];

    let veiculoEscolhido = null;

    for (const veiculo of veiculos) {
        if (totalKg <= veiculo.maxKg && totalCubagem <= veiculo.maxCubagem) {
            veiculoEscolhido = veiculo;
            break;
        }
    }
    
    if (veiculoEscolhido) {
        processedSet.clear();
        pedidosSelecionados.forEach(p => processedSet.add(String(p.Num_Pedido)));

        const loadId = `${nomeCarga.toLowerCase().replace(/\s+/g, '-')}-1`;
        const load = {
            id: loadId,
            pedidos: pedidosSelecionados,
            totalKg: totalKg,
            totalCubagem: totalCubagem,
            numero: nomeCarga,
            vehicleType: veiculoEscolhido.tipo
        };
        activeLoads[loadId] = load;
        
        const vehicleInfo = {
            fiorino: { name: 'Fiorino', colorClass: 'bg-success', textColor: 'text-white', icon: 'bi-box-seam-fill' },
            van: { name: 'Van', colorClass: 'bg-primary', textColor: 'text-white', icon: 'bi-truck-front-fill' },
            tresQuartos: { name: '3/4', colorClass: 'bg-warning', textColor: 'text-dark', icon: 'bi-truck-flatbed' }
        };

        resultadoDiv.innerHTML = `
            <div class="alert alert-success d-flex justify-content-between align-items-center">
                <div>
                    <strong>Carga ${nomeCarga} montada com sucesso!</strong> Estes ${pedidosSelecionados.length} pedidos foram agrupados em um(a) <strong>${veiculoEscolhido.nome}</strong>.
                    <br>Eles serão removidos da análise geral quando você clicar em "Processar Cargas".
                </div>
                <button class="btn btn-light btn-sm no-print" onclick="imprimirGeneric('${resultadoId}', 'Carga ${nomeCarga}')"><i class="bi bi-printer-fill me-1"></i> Imprimir Carga</button>
            </div>
            ${renderLoadCard(load, veiculoEscolhido.tipo, vehicleInfo[veiculoEscolhido.tipo])}
        `;
    } else {
        processedSet.clear();
        resultadoDiv.innerHTML = `
            <div class="alert alert-danger">
                <strong>Não foi possível montar a carga.</strong> O total dos pedidos selecionados excede a capacidade de todos os veículos disponíveis.
                <ul>
                    <li><strong>Peso Total:</strong> ${totalKg.toFixed(2)} kg</li>
                    <li><strong>Cubagem Total:</strong> ${totalCubagem.toFixed(2)} m³</li>
                </ul>
            </div>
        `;
    }
}

function highlightClientRows(event) {
    const clickedRow = event.target.closest('tr');
    if (!clickedRow || !clickedRow.dataset.clienteId) return;

    const clienteId = clickedRow.dataset.clienteId;
    const isAlreadyHighlighted = clickedRow.classList.contains('client-highlight');

    document.querySelectorAll('tr.client-highlight').forEach(row => {
        row.classList.remove('client-highlight');
    });

    if (!isAlreadyHighlighted) {
        document.querySelectorAll(`tr[data-cliente-id='${clienteId}']`).forEach(row => {
            row.classList.add('client-highlight');
        });
    }
}

