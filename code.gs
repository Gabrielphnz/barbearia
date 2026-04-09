// ==================== CONFIGURAÇÕES GLOBAIS ====================
const SHEETS = {
  USUARIOS: 'USUARIOS',
  CLIENTES: 'CLIENTES',
  AGENDA: 'AGENDA',
  FINANCEIRO: 'Total', 
  EVOLUCAO: 'EVOLUCAO',
  ESTOQUE_PRODUTOS: 'ESTOQUE_PRODUTOS',
  ESTOQUE_MOV: 'ESTOQUE_MOV',
  LOG_EXCLUSOES: 'LOG_EXCLUSOES'
};

// ==================== AUTENTICAÇÃO ====================
function autenticar(login, senha) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USUARIOS);
  if(!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == login && data[i][2] == senha) {
      return { success: true, nivel: data[i][3], nome: data[i][0] };
    }
  }
  return { success: false };
}

// ==================== CONFIGURAÇÕES (NOVO MOTOR) ====================
function getConfiguracoesCompletas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('config') || ss.getSheets().find(s => s.getName().toLowerCase().trim() === 'config');
  let config = { tiposCorte: [], precos: [], barbeiros: [], formasPagamento: [] };
  if (!sheet) return config;

  const data = sheet.getDataRange().getValues();
  let colEstilos = -1, rowEstilos = -1, colPrecos = -1, colBarbeiros = -1, rowBarbeiros = -1, colFormas = -1, rowFormas = -1;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      let val = String(data[r][c]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
      if (val === 'ESTILOS' || val === 'TIPO DE CORTE' || val === 'SERVICOS') { colEstilos = c; rowEstilos = r; }
      else if (val === 'PRECOS' || val === 'VALORES' || val === 'PRECO' || val === 'PREÇOS') { colPrecos = c; }
      else if (val === 'BARBEIROS' || val === 'PROFISSIONAIS') { colBarbeiros = c; rowBarbeiros = r; }
      else if (val.includes('FORMAS PAG') || val.includes('PAGAMENTO') || val === 'FORMAS DE PAGAMENTO') { colFormas = c; rowFormas = r; }
    }
  }

  if (colEstilos !== -1) {
    if (colPrecos === -1) colPrecos = colEstilos + 1;
    for (let r = rowEstilos + 1; r < data.length; r++) {
      let estilo = String(data[r][colEstilos]).trim();
      if (!estilo) break; 
      let preco = parseFloat(String(data[r][colPrecos]).replace(',', '.')) || 0;
      config.tiposCorte.push(estilo);
      config.precos.push({ tipo: estilo, preco: preco });
    }
  }

  if (colBarbeiros !== -1) {
    for (let r = rowBarbeiros + 1; r < data.length; r++) {
      let b = String(data[r][colBarbeiros]).trim();
      if (!b) break;
      config.barbeiros.push(b);
    }
  }

  if (colFormas !== -1) {
    for (let r = rowFormas + 1; r < data.length; r++) {
      let f = String(data[r][colFormas]).trim();
      if (!f) break;
      config.formasPagamento.push(f);
    }
  }
  return config;
}

function salvarConfiguracoesAPI(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('config') || ss.getSheets().find(s => s.getName().toLowerCase().trim() === 'config');
  if (!sheet) sheet = ss.insertSheet('config'); 
  sheet.clear(); 
  
  sheet.getRange(1, 1).setValue('ESTILOS').setFontWeight('bold');
  sheet.getRange(1, 2).setValue('PREÇOS').setFontWeight('bold');
  for (let i = 0; i < config.precos.length; i++) {
    sheet.getRange(i + 2, 1).setValue(config.precos[i].tipo);
    sheet.getRange(i + 2, 2).setValue(config.precos[i].preco);
  }
  
  sheet.getRange(1, 4).setValue('BARBEIROS').setFontWeight('bold');
  for (let i = 0; i < config.barbeiros.length; i++) {
    sheet.getRange(i + 2, 4).setValue(config.barbeiros[i]);
  }
  
  sheet.getRange(1, 6).setValue('FORMAS DE PAGAMENTO').setFontWeight('bold');
  for (let i = 0; i < config.formasPagamento.length; i++) {
    sheet.getRange(i + 2, 6).setValue(config.formasPagamento[i]);
  }
  return true;
}

// ==================== CLIENTES ====================
function getClientes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLIENTES);
  if (!sheet || sheet.getLastRow() < 2) return []; 
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const clientes = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i].join('').trim()) continue; 
    let cliente = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      cliente[headers[j]] = (val === null || val === undefined) ? '' : String(val).trim();
    }
    clientes.push(cliente);
  }
  return clientes.sort((a, b) => String(a.Nome || a.Cliente || '').localeCompare(String(b.Nome || b.Cliente || '')));
}

function salvarClienteCompleto(dados, usuario) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLIENTES);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let linhaExistente = -1;
  
  if (dados.ID && sheet.getLastRow() > 1) {
    const colIndex = headers.findIndex(h => String(h).toUpperCase().trim() === 'ID');
    const colId = colIndex !== -1 ? colIndex + 1 : 1;
    const ids = sheet.getRange(2, colId, sheet.getLastRow()-1, 1).getValues().flat();
    linhaExistente = ids.findIndex(id => String(id) === String(dados.ID)) + 2;
  }
  
  const novaLinha = new Array(headers.length).fill('');
  headers.forEach((h, i) => {
    let hUpper = String(h).toUpperCase().trim();
    if (hUpper === 'ID') novaLinha[i] = dados.ID || new Date().getTime();
    else if (hUpper === 'NOME' || hUpper === 'CLIENTE') novaLinha[i] = dados.Nome;
    else if (hUpper === 'APELIDO') novaLinha[i] = dados.Apelido || '';
    else if (hUpper === 'CELULAR' || hUpper === 'TELEFONE') novaLinha[i] = dados.Celular;
    else if (hUpper === 'CPF') novaLinha[i] = dados.CPF || '';
    else if (hUpper === 'NASCIMENTO') novaLinha[i] = dados.Nascimento || '';
    else if (hUpper === 'ENDERECO' || hUpper === 'ENDEREÇO') novaLinha[i] = dados.Endereco || '';
    else if (hUpper === 'BAIRRO') novaLinha[i] = dados.Bairro || '';
    else if (hUpper === 'CIDADE') novaLinha[i] = dados.Cidade || '';
    else if (hUpper === 'EMAIL') novaLinha[i] = dados.Email || '';
    else if (hUpper === 'OBS' || hUpper.includes('OBSERVA')) novaLinha[i] = dados.Obs || '';
  });
  
  if (linhaExistente > 1) sheet.getRange(linhaExistente, 1, 1, novaLinha.length).setValues([novaLinha]);
  else sheet.appendRow(novaLinha);
  return { success: true, clienteNome: dados.Nome };
}

// ==================== AGENDA ====================
function getAgendamentos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AGENDA);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const eventos = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i].join('').trim()) continue;
    let ev = {};
    for (let j = 0; j < headers.length; j++) ev[headers[j]] = data[i][j];
    
    let dataStr = '';
    if (ev.Data instanceof Date) dataStr = Utilities.formatDate(ev.Data, Session.getScriptTimeZone(), "yyyy-MM-dd");
    else if (typeof ev.Data === 'string') {
      if(ev.Data.includes('/')) { let p = ev.Data.split('/'); if(p.length === 3) dataStr = `${p[2]}-${p[1]}-${p[0]}`; }
      else dataStr = ev.Data.split('T')[0];
    }
    
    let horaStr = '00:00';
    if (ev.Hora instanceof Date) horaStr = Utilities.formatDate(ev.Hora, Session.getScriptTimeZone(), "HH:mm");
    else if (typeof ev.Hora === 'string' && ev.Hora.trim() !== '') horaStr = ev.Hora.substring(0, 5);
    
    if (dataStr && dataStr.length >= 8) {
      eventos.push({
        id: ev.ID || new Date().getTime() + i,
        data: dataStr, hora: horaStr,
        cliente: ev.Cliente || '', barbeiro: ev.Barbeiro || '',
        servico: ev['Tipo de corte'] || ev['Servico'] || 'Corte',
        celular: ev.Celular || '', status: ev.Status || 'Agendado',
        obs: ev.Obs || ev['Observações'] || ''
      });
    }
  }
  return eventos.sort((a, b) => a.hora.localeCompare(b.hora));
}

function salvarAgendamento(agendamento, usuario) {
  const sheetAgenda = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AGENDA);
  const headers = sheetAgenda.getRange(1, 1, 1, sheetAgenda.getLastColumn()).getValues()[0];
  let linhaExistente = -1;
  
  if (agendamento.id && sheetAgenda.getLastRow() > 1) {
    const ids = sheetAgenda.getRange(2, 1, sheetAgenda.getLastRow()-1, 1).getValues().flat();
    linhaExistente = ids.findIndex(id => String(id) === String(agendamento.id)) + 2;
  }
  
  const novaLinha = [];
  for (let i = 0; i < headers.length; i++) {
    let h = String(headers[i]).toUpperCase().trim();
    let valor = '';
    if (h === 'ID') valor = agendamento.id || new Date().getTime();
    else if (h === 'DATA') valor = agendamento.data;
    else if (h === 'HORA') valor = agendamento.hora;
    else if (h === 'CLIENTE') valor = agendamento.cliente;
    else if (h === 'BARBEIRO') valor = agendamento.barbeiro;
    else if (h === 'TIPO DE CORTE' || h === 'SERVICO' || h === 'SERVIÇO') valor = agendamento.servico;
    else if (h === 'STATUS') valor = agendamento.status;
    else if (h === 'CELULAR') valor = agendamento.celular;
    else if (h === 'OBS' || h.includes('OBSERVA')) valor = agendamento.obs;
    novaLinha.push(valor);
  }
  
  if (linhaExistente > 1) sheetAgenda.getRange(linhaExistente, 1, 1, novaLinha.length).setValues([novaLinha]);
  else sheetAgenda.appendRow(novaLinha);

  if (agendamento.status === 'Concluído' && agendamento.valorLancamento !== undefined && agendamento.valorLancamento !== '') {
    lancarNoFinanceiro(agendamento, usuario);
    registrarEvolucao(agendamento.cliente, agendamento.barbeiro, agendamento.servico, agendamento.obs);
  }
  return { success: true };
}

function excluirAgendamento(id, usuario) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AGENDA);
  if(sheet.getLastRow() <= 1) return { success: false };
  const ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat();
  const index = ids.findIndex(v => String(v) === String(id));
  if (index !== -1) {
    sheet.deleteRow(index+2);
    return { success: true };
  }
  return { success: false };
}

// ==================== FINANCEIRO BLINDADO ====================
function getSheetFinanceiro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.FINANCEIRO);
  if (sheet) return sheet;
  return ss.getSheets().find(s => ['total', 'atendimentos', 'financeiro', 'caixa'].includes(s.getName().toLowerCase().trim()));
}

function getHeaderRowIndex(dataFin) {
  for (let r = 0; r < Math.min(5, dataFin.length); r++) {
    if (String(dataFin[r][0]).toUpperCase().trim() === 'DATA') return r;
  }
  return 0;
}

function lancarNoFinanceiro(agendamento, usuario) {
  const sheet = getSheetFinanceiro();
  if (!sheet || sheet.getLastColumn() < 1) return; 

  const dataFin = sheet.getDataRange().getValues();
  if (dataFin.length === 0) return;

  const headerRowIdx = getHeaderRowIndex(dataFin);
  const headers = dataFin[headerRowIdx] || sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let novaLinha = new Array(headers.length).fill('');
  let dataBR = agendamento.data;
  if(dataBR && dataBR.includes('-')) dataBR = dataBR.split('-').reverse().join('/');
  const valorNum = parseFloat(agendamento.valorLancamento.toString().replace(',', '.'));
  const formaPag = String(agendamento.formaPagamento || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();

  let colunaValorEncontrada = false;

  headers.forEach((h, i) => {
    let hUpper = String(h).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
    if (hUpper === 'DATA') novaLinha[i] = dataBR;
    else if (hUpper === 'BARBEIRO') novaLinha[i] = agendamento.barbeiro;
    else if (hUpper === 'CLIENTES' || hUpper === 'CLIENTE' || hUpper === 'NOME') novaLinha[i] = agendamento.cliente;
    else if (hUpper === 'TIPO DE CORTE' || hUpper === 'SERVICO') novaLinha[i] = agendamento.servico;
    else if (hUpper.includes('OBSERVA')) novaLinha[i] = agendamento.obs;
    else if (hUpper === 'FORMA DE PAG' || hUpper.includes('FORMA PAG')) novaLinha[i] = agendamento.formaPagamento;
    else if (hUpper === 'ID CONTROLE') novaLinha[i] = agendamento.id || new Date().getTime(); // Rastreabilidade
    else if (formaPag && hUpper.includes(formaPag)) { novaLinha[i] = valorNum; colunaValorEncontrada = true; }
  });

  if (!colunaValorEncontrada) {
    headers.forEach((h, i) => {
      if (!colunaValorEncontrada && String(h).toUpperCase().includes('VALOR')) {
        novaLinha[i] = valorNum;
        colunaValorEncontrada = true;
      }
    });
  }
  sheet.appendRow(novaLinha);
}

function getDadosFinanceiros(filtroDataInicio, filtroDataFim, barbeiro = null) {
  const sheet = getSheetFinanceiro();
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];
  
  const headerRowIdx = getHeaderRowIndex(data);
  const headers = data[headerRowIdx];
  const transacoes = [];
  const inicio = filtroDataInicio ? new Date(filtroDataInicio) : null;
  const fim = filtroDataFim ? new Date(filtroDataFim) : null;
  
  for (let i = data.length - 1; i > headerRowIdx; i--) {
    if (!data[i].join('').trim()) continue;
    let t = { linha: i + 1, id: '' }; // Exporta a linha real da planilha para permitir edição
    let dataTransacao = null;
    
    for (let j = 0; j < headers.length; j++) {
      let h = String(headers[j]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
      let val = data[i][j];
      
      if (h === 'DATA') {
        if (val instanceof Date) {
          dataTransacao = val;
          t.data = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (typeof val === 'string') {
          let partes = val.split('/');
          if (partes.length === 3) dataTransacao = new Date(`${partes[2]}-${partes[1]}-${partes[0]}`);
          t.data = val;
        }
      }
      else if (h === 'CLIENTES' || h === 'CLIENTE') t.cliente = val;
      else if (h === 'BARBEIRO') t.barbeiro = val;
      else if (h === 'TIPO DE CORTE' || h === 'SERVICO') t.servico = val;
      else if (h.includes('FORMA DE PAG')) t.forma = val;
      else if (h === 'ID CONTROLE') t.id = val;
      else if (h.includes('VALOR') && val && !isNaN(parseFloat(val))) t.valor = (t.valor || 0) + parseFloat(val);
    }
    
    if (t.valor || t.forma) { // Permite mostrar mesmo se o valor for 0 mas tiver forma
      if (inicio && dataTransacao < inicio) continue;
      if (fim && dataTransacao > fim) continue;
      if (barbeiro && t.barbeiro !== barbeiro) continue;
      transacoes.push(t);
    }
  }
  return transacoes;
}

// NOVA FUNÇÃO: EDITA O FINANCEIRO DIRETO NA LINHA
function atualizarEdicaoFinanceiro(dadosLancamento) {
  const sheet = getSheetFinanceiro();
  if (!sheet) return { success: false };
  
  const linha = dadosLancamento.linha;
  const dataFin = sheet.getDataRange().getValues();
  const headerRowIdx = getHeaderRowIndex(dataFin);
  const headers = dataFin[headerRowIdx];
  
  // Zera as colunas de valor antigas na linha para não somar/duplicar
  for (let j = 0; j < headers.length; j++) {
    let hUpper = String(headers[j]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
    if (hUpper.includes('VALOR')) sheet.getRange(linha, j + 1).setValue('');
  }
  
  const valorNum = parseFloat(dadosLancamento.valor);
  const formaPag = String(dadosLancamento.forma).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();

  // Reescreve a linha
  headers.forEach((h, j) => {
    let hUpper = String(h).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
    let celula = sheet.getRange(linha, j + 1);
    
    if (hUpper === 'DATA') celula.setValue(dadosLancamento.data);
    else if (hUpper === 'BARBEIRO') celula.setValue(dadosLancamento.barbeiro);
    else if (hUpper === 'CLIENTES' || hUpper === 'CLIENTE') celula.setValue(dadosLancamento.cliente);
    else if (hUpper === 'TIPO DE CORTE' || hUpper === 'SERVICO') celula.setValue(dadosLancamento.servico);
    else if (hUpper === 'FORMA DE PAG' || hUpper.includes('FORMA PAG')) celula.setValue(dadosLancamento.forma);
    else if (formaPag && hUpper.includes(formaPag)) celula.setValue(valorNum);
  });
  
  return { success: true };
}

// ==================== ESTOQUE E VENDAS (NOVO MOTOR) ====================
function getProdutos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ESTOQUE_PRODUTOS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const produtos = [];
  for (let i = 1; i < data.length; i++) {
    if (!String(data[i][1]).trim()) continue;
    let p = {};
    for (let j = 0; j < headers.length; j++) {
      p[String(headers[j]).toUpperCase().trim()] = data[i][j];
    }
    produtos.push(p);
  }
  return produtos;
}

// NOVA FUNÇÃO: Realiza a Venda do Produto
function realizarVendaProduto(venda, usuario) {
  const sheetProd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ESTOQUE_PRODUTOS);
  const produtos = getProdutos();
  const idxProd = produtos.findIndex(p => p.PRODUTO === venda.produto);
  
  if (idxProd !== -1) {
    const linha = idxProd + 2;
    // Localiza as colunas cruciais
    const headers = sheetProd.getRange(1, 1, 1, sheetProd.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase().trim());
    const colAtual = headers.indexOf('ESTOQUE ATUAL') + 1;
    
    if (colAtual > 0) {
      const qtdAtual = parseFloat(sheetProd.getRange(linha, colAtual).getValue()) || 0;
      sheetProd.getRange(linha, colAtual).setValue(qtdAtual - venda.quantidade);
    }
    
    // Registra Movimento
    const sheetMov = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ESTOQUE_MOV);
    if (sheetMov) {
      sheetMov.appendRow([
        venda.data, usuario, venda.produto, 'Venda', venda.quantidade, venda.precoUnitario, venda.total, venda.barbeiro, 'Venda Balcão'
      ]);
    }
    
    // Lança no Financeiro
    lancarNoFinanceiro({
      data: venda.data,
      cliente: 'Venda Avulsa',
      barbeiro: venda.barbeiro,
      servico: 'Produto: ' + venda.produto,
      formaPagamento: venda.formaPagamento,
      valorLancamento: venda.total,
      obs: `${venda.quantidade}x ${venda.produto}`
    }, usuario);
    
    return { success: true };
  }
  return { success: false, msg: 'Produto não encontrado' };
}

// (Mantenha as funções antigas salvarProduto e movimentarEstoque originais aqui se quiser gerenciar estoque normal)
function salvarProduto(produto, usuario) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ESTOQUE_PRODUTOS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let linhaExistente = -1;
  if (produto.ID && sheet.getLastRow() > 1) {
    const colId = headers.findIndex(h => String(h).toUpperCase().trim() === 'ID') + 1;
    if(colId > 0) {
      const ids = sheet.getRange(2, colId, sheet.getLastRow()-1, 1).getValues().flat();
      linhaExistente = ids.findIndex(id => String(id) === String(produto.ID)) + 2;
    }
  }
  const novaLinha = new Array(headers.length).fill('');
  headers.forEach((h, i) => {
    let hUpper = h.toString().toUpperCase().trim();
    if (hUpper === 'ID') novaLinha[i] = produto.ID || new Date().getTime();
    else if (hUpper === 'PRODUTO') novaLinha[i] = produto.nome;
    else if (hUpper === 'CATEGORIA') novaLinha[i] = produto.categoria;
    else if (hUpper === 'UNIDADE') novaLinha[i] = produto.unidade;
    else if (hUpper === 'ESTOQUE MINIMO') novaLinha[i] = produto.estoqueMinimo;
    else if (hUpper === 'ESTOQUE ATUAL' && !produto.ID) novaLinha[i] = 0;
    else if (hUpper === 'CUSTO') novaLinha[i] = produto.custo || 0;
    else if (hUpper === 'PRECO' || hUpper === 'PREÇO') novaLinha[i] = produto.preco || 0;
  });
  if (linhaExistente > 1) sheet.getRange(linhaExistente, 1, 1, novaLinha.length).setValues([novaLinha]);
  else sheet.appendRow(novaLinha);
}

// ==================== DASHBOARD ANALYTICS (GRÁFICOS) ====================
function getMetricasDashboard() {
  const financeiro = getDadosFinanceiros(null, null, null);
  const hojeDate = new Date();
  const hojeStr = Utilities.formatDate(hojeDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  let faturamentoHoje = 0;
  let faturamentoMes = 0;
  let cortesMaisVendidos = {};
  let rankingBarbeiros = {};
  let clientesAtendidos = new Set();
  
 let mesesMap = new Map(); // chave "MM/AAAA"
financeiro.forEach(t => {
  if (t.data) {
    let partes = t.data.split('/');
    if (partes.length === 3) {
      let mesAno = `${partes[1]}/${partes[2]}`;
      mesesMap.set(mesAno, (mesesMap.get(mesAno) || 0) + (t.valor || 0));
    }
  }
});
// ordenar e pegar últimos 6 meses
let entradas = Array.from(mesesMap.entries()).sort((a,b) => {
  let [ma, ya] = a[0].split('/'), [mb, yb] = b[0].split('/');
  return new Date(ya, ma-1) - new Date(yb, mb-1);
});
let ultimos = entradas.slice(-6);
let faturamentoMensalChart = { labels: ultimos.map(e => e[0]), dados: ultimos.map(e => e[1]) };

  financeiro.forEach(t => {
    // Totais Básicos
    if (t.data === hojeStr) faturamentoHoje += t.valor || 0;
    
    if(t.data) {
      const dataTrans = t.data.split('/');
      if (dataTrans.length === 3) {
        const mesTrans = dataTrans[1];
        const anoTrans = dataTrans[2];
        const mesAno = `${mesTrans}/${anoTrans}`;
        
        if (parseInt(mesTrans) === (hojeDate.getMonth() + 1) && parseInt(anoTrans) === hojeDate.getFullYear()) {
          faturamentoMes += t.valor || 0;
          if (t.cliente && t.cliente !== 'Venda Avulsa') clientesAtendidos.add(t.cliente);
        }

        // Gráfico de Faturamento (Soma por Mês/Ano)
        if(!mesesMap[mesAno]) mesesMap[mesAno] = 0;
        mesesMap[mesAno] += t.valor || 0;
      }
    }

    // Top Cortes/Serviços
    if (t.servico) {
      if(!cortesMaisVendidos[t.servico]) cortesMaisVendidos[t.servico] = 0;
      cortesMaisVendidos[t.servico]++;
    }

let servicosOrdenados = Object.entries(cortesMaisVendidos)
  .map(([nome, qtd]) => [nome, Number(qtd) || 0])
  .filter(([, qtd]) => qtd > 0)
  .sort((a, b) => b[1] - a[1])
  .slice(0, 5);

let chartServicos = {
  labels: servicosOrdenados.map(s => s[0]),
  dados: servicosOrdenados.map(s => s[1])
};

    // Top Barbeiros
    if (t.barbeiro) {
      if(!rankingBarbeiros[t.barbeiro]) rankingBarbeiros[t.barbeiro] = 0;
      rankingBarbeiros[t.barbeiro] += t.valor || 0;
    }
  });

  // Prepara dados pro Chart.js (Faturamento Mensal ordenado)
  const sortedMeses = Object.keys(mesesMap).sort((a, b) => {
    let [ma, ya] = a.split('/'); let [mb, yb] = b.split('/');
    return new Date(ya, ma - 1) - new Date(yb, mb - 1);
  }).slice(-6); // Pega os últimos 6 meses

  sortedMeses.forEach(m => {
    faturamentoMensalChart.labels.push(m);
    faturamentoMensalChart.dados.push(mesesMap[m]);
  });

  // Prepara Pizza (Top 5 Serviços)
  let servicosOrdenados = Object.entries(cortesMaisVendidos).sort((a,b) => b[1] - a[1]).slice(0,5);
  let chartServicos = { labels: servicosOrdenados.map(s => s[0]), dados: servicosOrdenados.map(s => s[1]) };

  // Prepara Barras (Ranking Barbeiros)
  let barbeirosOrdenados = Object.entries(rankingBarbeiros).sort((a,b) => b[1] - a[1]);
  let chartBarbeiros = { labels: barbeirosOrdenados.map(b => b[0]), dados: barbeirosOrdenados.map(b => b[1]) };

  const ticketMedio = clientesAtendidos.size > 0 ? (faturamentoMes / clientesAtendidos.size) : 0;
  
  const produtos = getProdutos();
  const baixoEstoque = produtos.filter(p => (p['ESTOQUE ATUAL'] || 0) <= (p['ESTOQUE MINIMO'] || 0)).map(p => p.PRODUTO);

  return { 
    faturamentoHoje, 
    faturamentoMes, 
    ticketMedio, 
    clientesUnicosMes: clientesAtendidos.size,
    baixoEstoque,
    graficos: {
      faturamentoMensal: faturamentoMensalChart,
      topServicos: chartServicos,
      topBarbeiros: chartBarbeiros
    }
  };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Barbearia Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
