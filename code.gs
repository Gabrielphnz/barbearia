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

// ==================== FINANCEIRO (SIMPLES E DIRETO) ====================
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

// Lança o valor na coluna VALOR e o texto na coluna FORMA DE PAGAMENTO
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
  const formaPag = agendamento.formaPagamento; // Pega exatamente o texto que veio do select (ex: "PIX EMANUEL")

  headers.forEach((h, i) => {
    let hUpper = String(h).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
    
    if (hUpper === 'DATA') novaLinha[i] = dataBR;
    else if (hUpper === 'BARBEIRO') novaLinha[i] = agendamento.barbeiro;
    else if (hUpper === 'CLIENTES' || hUpper === 'CLIENTE' || hUpper === 'NOME') novaLinha[i] = agendamento.cliente;
    else if (hUpper === 'TIPO DE CORTE' || hUpper === 'SERVICO') novaLinha[i] = agendamento.servico;
    else if (hUpper.includes('OBSERVA')) novaLinha[i] = agendamento.obs;
    else if (hUpper === 'FORMA DE PAG' || hUpper.includes('FORMA PAG')) novaLinha[i] = formaPag; // Coloca a string da forma de pagamento
    else if (hUpper === 'ID CONTROLE') novaLinha[i] = agendamento.id || new Date().getTime(); 
  });

  // Acha a coluna de Valor (a primeira que tiver a palavra VALOR no cabeçalho)
  let idxValor = headers.findIndex(h => String(h).toUpperCase().trim() === 'VALOR' || String(h).toUpperCase().trim() === 'VALOR (R$)');
  if (idxValor === -1) idxValor = headers.findIndex(h => String(h).toUpperCase().includes('VALOR'));
  
  if (idxValor !== -1) {
    novaLinha[idxValor] = valorNum; // Coloca o dinheiro aqui
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
  
  let idxValor = headers.findIndex(h => String(h).toUpperCase().trim() === 'VALOR' || String(h).toUpperCase().trim() === 'VALOR (R$)');
  if (idxValor === -1) idxValor = headers.findIndex(h => String(h).toUpperCase().includes('VALOR'));

  for (let i = data.length - 1; i > headerRowIdx; i--) {
    if (!data[i].join('').trim()) continue;
    let t = { linha: i + 1, id: '', valor: 0 }; 
    let dataTransacao = null;
    
    for (let j = 0; j < headers.length; j++) {
      let h = String(headers[j]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
      let val = data[i][j];
      
      // BLINDAGEM: Converte TUDO para texto (String) antes de enviar para a tela
      if (h === 'DATA') {
        if (val instanceof Date) { dataTransacao = val; t.data = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy"); } 
        else if (typeof val === 'string') { let p = val.split('/'); if (p.length === 3) dataTransacao = new Date(`${p[2]}-${p[1]}-${p[0]}`); t.data = val; }
        else { t.data = String(val); }
      }
      else if (h === 'CLIENTES' || h === 'CLIENTE') t.cliente = String(val);
      else if (h === 'BARBEIRO') t.barbeiro = String(val);
      else if (h === 'TIPO DE CORTE' || h === 'SERVICO') t.servico = String(val);
      else if (h.includes('FORMA DE PAG')) t.forma = String(val);
      else if (h === 'ID CONTROLE') t.id = String(val);
    }
    
    if (idxValor !== -1) {
      let v = data[i][idxValor];
      if (v && !isNaN(parseFloat(v))) t.valor = parseFloat(v);
    }
    
    if (t.valor || t.forma) { 
      if (inicio && dataTransacao < inicio) continue;
      if (fim && dataTransacao > fim) continue;
      if (barbeiro && t.barbeiro !== barbeiro) continue;
      transacoes.push(t);
    }
  }
  return transacoes;
}


function atualizarEdicaoFinanceiro(dadosLancamento) {
  const sheet = getSheetFinanceiro();
  if (!sheet) return { success: false };
  const linha = dadosLancamento.linha;
  const dataFin = sheet.getDataRange().getValues();
  const headerRowIdx = getHeaderRowIndex(dataFin);
  const headers = dataFin[headerRowIdx];
  
  const valorNum = parseFloat(dadosLancamento.valor);
  const formaPag = dadosLancamento.forma; // Mantém a string exata

  let idxValor = headers.findIndex(h => String(h).toUpperCase().trim() === 'VALOR' || String(h).toUpperCase().trim() === 'VALOR (R$)');
  if (idxValor === -1) idxValor = headers.findIndex(h => String(h).toUpperCase().includes('VALOR'));

  headers.forEach((h, j) => {
    let hUpper = String(h).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
    let celula = sheet.getRange(linha, j + 1);
    
    if (hUpper === 'DATA') celula.setValue(dadosLancamento.data);
    else if (hUpper === 'BARBEIRO') celula.setValue(dadosLancamento.barbeiro);
    else if (hUpper === 'CLIENTES' || hUpper === 'CLIENTE') celula.setValue(dadosLancamento.cliente);
    else if (hUpper === 'TIPO DE CORTE' || hUpper === 'SERVICO') celula.setValue(dadosLancamento.servico);
    else if (hUpper === 'FORMA DE PAG' || hUpper.includes('FORMA PAG')) celula.setValue(formaPag);
    
    // Limpa colunas de valor que não são a principal (caso tenha sobrado lixo antigo)
    else if (hUpper.includes('VALOR') && j !== idxValor) celula.setValue('');
  });
  
  if (idxValor !== -1) {
    sheet.getRange(linha, idxValor + 1).setValue(valorNum);
  }
  
  return { success: true };
}

function obterFichaCompleta(nomeCliente) {
  const clientes = getClientes();
  const dadosPessoais = clientes.find(c => String(c.Nome || c.Cliente).toLowerCase() === String(nomeCliente).toLowerCase()) || {};
  const historico = []; let totalGasto = 0;
  const sheetFinanceiro = getSheetFinanceiro();
  
  if (sheetFinanceiro && sheetFinanceiro.getLastRow() > 1) {
    const dataFin = sheetFinanceiro.getDataRange().getValues();
    const headerRowIdx = getHeaderRowIndex(dataFin);
    const headFin = dataFin[headerRowIdx];
    
    let idxValor = headFin.findIndex(h => String(h).toUpperCase().trim() === 'VALOR' || String(h).toUpperCase().trim() === 'VALOR (R$)');
    if (idxValor === -1) idxValor = headFin.findIndex(h => String(h).toUpperCase().includes('VALOR'));

    for (let i = dataFin.length - 1; i > headerRowIdx; i--) {
      if (!dataFin[i].join('').trim()) continue;
      let isMatch = false; let t = { valor: 0 };
      
      headFin.forEach((h, j) => {
        let hUp = String(h).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim(); 
        let val = dataFin[i][j];
        if ((hUp === 'CLIENTES' || hUp === 'CLIENTE' || hUp === 'NOME') && String(val).toLowerCase() === String(nomeCliente).toLowerCase()) isMatch = true;
        
        if (hUp === 'DATA') t.data = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy") : val;
        if (hUp === 'BARBEIRO') t.barbeiro = val;
        if (hUp === 'TIPO DE CORTE' || hUp === 'SERVICO') t.servico = val;
        if (hUp.includes('FORMA DE PAG')) t.forma = val;
      });
      
      if (idxValor !== -1) {
        let v = dataFin[i][idxValor];
        if (v && !isNaN(parseFloat(v))) t.valor = parseFloat(v);
      }
      
      if (isMatch && t.valor > 0) { historico.push(t); totalGasto += t.valor; }
    }
  }
  return { dados: dadosPessoais, historico: historico, resumo: { totalGasto: totalGasto, ultimaVisita: historico.length > 0 ? historico[0].data : 'Nunca', qtdVisitas: historico.length } };
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
    
    const sheetFin = getSheetFinanceiro();
    if (sheetFin) {
      const headersFin = sheetFin.getRange(1, 1, 1, sheetFin.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase().trim());
      const novaLinhaFin = new Array(headersFin.length).fill('');
      headersFin.forEach((h, idx) => {
        if (h === 'DATA') novaLinhaFin[idx] = new Date(venda.data + 'T12:00:00'); // Evita bug de timezone
        else if (h.includes('VALOR')) novaLinhaFin[idx] = parseFloat(venda.total);
        else if (h === 'TIPO DE CORTE' || h === 'SERVICO') novaLinhaFin[idx] = venda.produto;
        else if (h.includes('FORMA DE PAG')) novaLinhaFin[idx] = venda.formaPagamento;
        else if (h === 'BARBEIRO') novaLinhaFin[idx] = venda.barbeiro;
        else if (h === 'CLIENTES' || h === 'CLIENTE') novaLinhaFin[idx] = venda.cliente || '';
        else if (h === 'USUÁRIO') novaLinhaFin[idx] = usuario;
      });
      sheetFin.appendRow(novaLinhaFin);
    }
    
    // --- NOVO: REGISTRO NA EVOLUÇÃO (FICHA DO CLIENTE) ---
    // Se a venda teve um cliente selecionado, envia pro histórico
    if (venda.cliente && venda.cliente.trim() !== '') {
        const obsVenda = `Compra de Produto PDV: ${venda.quantidade}x`;
        registrarEvolucao(venda.cliente, venda.barbeiro, venda.produto, obsVenda);
    }

    return { success: true };
    
      }
  return { success: false, msg: 'Produto não encontrado' };
}

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
    
    // CORREÇÃO: Agora ele lê a variável "estoqueAtual" enviada pela tela
    else if (hUpper === 'ESTOQUE ATUAL') novaLinha[i] = produto.estoqueAtual || 0; 
    
    else if (hUpper === 'CUSTO') novaLinha[i] = produto.custo || 0;
    else if (hUpper === 'PRECO' || hUpper === 'PREÇO') novaLinha[i] = produto.preco || 0;
  });
  
  if (linhaExistente > 1) sheet.getRange(linhaExistente, 1, 1, novaLinha.length).setValues([novaLinha]);
  else sheet.appendRow(novaLinha);
}

// ==================== DASHBOARD & GRÁFICOS ====================
function getMetricasDashboard() {
  const financeiro = getDadosFinanceiros(null, null, null);
  const hojeDate = new Date();
  const hojeStr = Utilities.formatDate(hojeDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  let faturamentoHoje = 0, faturamentoMes = 0;
  let cortesMaisVendidos = {}, rankingBarbeiros = {};
  let clientesAtendidos = new Set();
  let faturamentoMensalChart = { labels: [], dados: [] };
  let mesesMap = {};

  financeiro.forEach(t => {
    if (t.data === hojeStr) faturamentoHoje += t.valor || 0;
    
    if(t.data) {
      const dataTrans = t.data.split('/');
      if (dataTrans.length === 3) {
        const mesTrans = dataTrans[1], anoTrans = dataTrans[2], mesAno = `${mesTrans}/${anoTrans}`;
        
        if (parseInt(mesTrans) === (hojeDate.getMonth() + 1) && parseInt(anoTrans) === hojeDate.getFullYear()) {
          faturamentoMes += t.valor || 0;
          if (t.cliente && t.cliente !== 'Venda Avulsa') clientesAtendidos.add(t.cliente);
        }

        // Soma os valores por mês para o Gráfico de Linha
        if(!mesesMap[mesAno]) mesesMap[mesAno] = 0; 
        mesesMap[mesAno] += t.valor || 0;
      }
    }

    // Conta os serviços para o Gráfico de Pizza
    if (t.servico) { 
      if(!cortesMaisVendidos[t.servico]) cortesMaisVendidos[t.servico] = 0; 
      cortesMaisVendidos[t.servico]++; 
    }

    // Soma o faturamento por barbeiro para o Gráfico de Barras
    if (t.barbeiro) { 
      if(!rankingBarbeiros[t.barbeiro]) rankingBarbeiros[t.barbeiro] = 0; 
      rankingBarbeiros[t.barbeiro] += t.valor || 0; 
    }
  });

  // Ordena os meses cronologicamente (pega os últimos 6)
  const sortedMeses = Object.keys(mesesMap).sort((a, b) => { 
    let [ma, ya] = a.split('/'); let [mb, yb] = b.split('/'); 
    return new Date(ya, ma - 1) - new Date(yb, mb - 1); 
  }).slice(-6);
  
  sortedMeses.forEach(m => { 
    faturamentoMensalChart.labels.push(m); 
    faturamentoMensalChart.dados.push(mesesMap[m]); 
  });

  let servicosOrdenados = Object.entries(cortesMaisVendidos).sort((a,b) => b[1] - a[1]).slice(0,5);
  let chartServicos = { labels: servicosOrdenados.map(s => s[0]), dados: servicosOrdenados.map(s => s[1]) };

  let barbeirosOrdenados = Object.entries(rankingBarbeiros).sort((a,b) => b[1] - a[1]);
  let chartBarbeiros = { labels: barbeirosOrdenados.map(b => b[0]), dados: barbeirosOrdenados.map(b => b[1]) };

  const ticketMedio = clientesAtendidos.size > 0 ? (faturamentoMes / clientesAtendidos.size) : 0;

  // Retorna os dados E a variável graficos que faltava
  return { 
    faturamentoHoje, 
    faturamentoMes, 
    ticketMedio, 
    clientesUnicosMes: clientesAtendidos.size, 
    baixoEstoque: [], 
    graficos: { 
      faturamentoMensal: faturamentoMensalChart, 
      topServicos: chartServicos, 
      topBarbeiros: chartBarbeiros 
    } 
  };
}

// NOVO: Exclusão de Financeiro integrada com Agenda
function excluirLancamentoFinanceiro(linha, idControle) {
  const sheetFin = getSheetFinanceiro();
  if (!sheetFin) return { success: false, msg: 'Aba financeiro não encontrada' };

  // 1. Deleta a linha do financeiro
  sheetFin.deleteRow(linha);

  let temAgenda = false;

  // 2. Se esse lançamento veio da agenda, desfaz o status
  if (idControle && String(idControle).trim() !== '') {
    const sheetAgenda = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AGENDA);
    if (sheetAgenda) {
      const dataAgenda = sheetAgenda.getDataRange().getValues();
      const headersAgenda = dataAgenda[0].map(h => String(h).toUpperCase().trim());
      const colId = headersAgenda.indexOf('ID');
      const colStatus = headersAgenda.indexOf('STATUS');

      if (colId !== -1 && colStatus !== -1) {
        // i = 1 para pular o cabeçalho
        for (let i = 1; i < dataAgenda.length; i++) {
          if (String(dataAgenda[i][colId]) === String(idControle)) {
            // Reverte o status
            sheetAgenda.getRange(i + 1, colStatus + 1).setValue('Agendado');
            temAgenda = true;
            break;
          }
        }
      }
    }
  }

  return { success: true, temAgenda: temAgenda, idAgenda: idControle };
}

// ==================== EVOLUÇÃO / HISTÓRICO ====================
function registrarEvolucao(cliente, barbeiro, estilo, observacoes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EVOLUCAO);
  if (sheet) {
    sheet.appendRow([
      new Date().getTime(), 
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"), 
      cliente, 
      barbeiro, 
      estilo, 
      observacoes
    ]);
    
    // OBRIGATÓRIO: Força o Google a salvar imediatamente na planilha
    SpreadsheetApp.flush(); 
  }
}

// ==================== FINALIZAÇÃO COM CHECKOUT E EVOLUÇÃO ====================
function finalizarAgendamentoCheckout(dados, usuario) {
  const sheetAgenda = getSheetAgenda();
  if (!sheetAgenda) return { success: false, msg: 'Aba Agenda não encontrada' };
  
  const dataAgenda = sheetAgenda.getDataRange().getValues();
  const headersAgenda = dataAgenda[0].map(h => String(h).toUpperCase().trim());
  const colId = headersAgenda.indexOf('ID');
  const colStatus = headersAgenda.indexOf('STATUS');

  if (colId !== -1 && colStatus !== -1) {
    for (let i = 1; i < dataAgenda.length; i++) {
      if (String(dataAgenda[i][colId]) === String(dados.id)) {
        
        // 1. Muda status da Agenda para Concluído
        sheetAgenda.getRange(i + 1, colStatus + 1).setValue('Concluído');
        
        // Pega as observações da agenda para salvar no histórico
        let obs = '';
        const colObs = headersAgenda.indexOf('OBSERVAÇÕES');
        if (colObs !== -1) obs = dataAgenda[i][colObs];

        // 2. Grava no Financeiro
        const idControle = new Date().getTime();
        sheetAgenda.getRange(i + 1, colId + 1).setValue(idControle); // Atualiza ID para vincular

        const sheetFin = getSheetFinanceiro();
        if (sheetFin) {
          const headersFin = sheetFin.getRange(1, 1, 1, sheetFin.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase().trim());
          const novaLinha = new Array(headersFin.length).fill('');
          
          headersFin.forEach((h, idx) => {
            if (h === 'DATA') novaLinha[idx] = new Date();
            else if (h.includes('VALOR')) novaLinha[idx] = parseFloat(dados.valor);
            else if (h === 'TIPO DE CORTE' || h === 'SERVICO') novaLinha[idx] = dados.servico;
            else if (h.includes('FORMA DE PAG')) novaLinha[idx] = dados.forma;
            else if (h === 'BARBEIRO') novaLinha[idx] = dados.barbeiro;
            else if (h === 'CLIENTES' || h === 'CLIENTE') novaLinha[idx] = dados.cliente;
            else if (h === 'ID CONTROLE') novaLinha[idx] = idControle;
            else if (h === 'USUÁRIO') novaLinha[idx] = usuario;
          });
          
          sheetFin.appendRow(novaLinha);
        }

        // 3. GRAVA NA EVOLUÇÃO DO CLIENTE!
        registrarEvolucao(dados.cliente, dados.barbeiro, dados.servico, obs);

        break;
      }
    }
  }
  return { success: true };
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Barbearia Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
