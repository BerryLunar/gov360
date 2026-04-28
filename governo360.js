/** 
 * Gestão de Projetos - SMMF
 * Autor: Luana Halcsik Leite - 41331 - 2026
 * Objetivo: Consolidação de dados e apoio ao Governo 360
 */

// --- NOMES DAS ABAS ---
var ABA_PROJETOS  = 'PROJETOS';
var ABA_ETAPAS    = 'ETAPAS';
var ABA_AGENDA    = 'AGENDA';
var ABA_DASHBOARD = 'DASHBOARD';
 
// --- COLUNAS: PROJETOS ---
var P = {
  ID:                 1,  // A
  PROGRAMA:           2,  // B
  CICLO:              3,  // C  ← texto/número livre, NUNCA sobrescrever
  TIPO:               4,  // D
  INICIO_PLANEJADO:   5,  // E  ← data, NUNCA sobrescrever
  INICIO_REAL:        6,  // F  ← data, NUNCA sobrescrever
  FIM_PLANEJADO:      7,  // G  ← data, NUNCA sobrescrever
  FIM_REAL:           8,  // H  ← data, NUNCA sobrescrever
  RESPONSAVEL:        9,  // I
  CONTATO:           10,  // J
  STATUS:            11,  // K  ← script escreve
  PERC_ETAPAS:       12,  // L  ← script escreve
  PERC_TEMPO:        13,  // M  ← script escreve
  PERC_GERAL:        14,  // N  ← script escreve
  ULTIMA_ATUALIZACAO:15,  // O  ← script escreve
  OBSERVACOES:       16,  // P
  DESCRICAO:         17   // Q
};
 
// --- COLUNAS: ETAPAS ---
var E = {
  ID_PROJETO:        1,  // A
  PROGRAMA:          2,  // B
  ETAPA:             3,  // C
  TIPO:              4,  // D
  INICIO_PLANEJADO:  5,  // E
  INICIO_REAL:       6,  // F
  FIM_PLANEJADO:     7,  // G
  FIM_REAL:          8,  // H
  QTD_PREVISTA:      9,  // I
  QTD_ATUAL:        10,  // J
  STATUS:           11,  // K
  DATA_STATUS:      12,  // L  ← script (OnEdit)
  OBSERVACOES:      13   // M
};
 
// --- STATUS VÁLIDOS PARA ETAPAS ---
// Aceita qualquer valor que indique conclusão
var STATUS_CONCLUIDO  = ['Concluído'];
var STATUS_EM_ANDAMENTO = ['Andamento'];
 
// --- COLUNAS: AGENDA ---
var A = {
  ID_PROJETO:        1,  // A
  PROGRAMA:          2,  // B
  ATUALIZACAO:       3,  // C
  FONTE:             4,  // D
  DATA_ATIVIDADE:    5,  // E
  DATA_LANCAMENTO:   6,  // F
  OBSERVACOES:       7   // G
};
 
// --- PALETA DE CORES ---
var COR = {
  PRIMARIO:    '#5B2C8D',
  SECUNDARIO:  '#8E44AD',
  DESTAQUE:    '#D7BDE2',
  VERDE:       '#1E8449',
  AMARELO:     '#D4AC0D',
  VERMELHO:    '#C0392B',
  AZUL:        '#1A5276',
  CINZA_CLARO: '#F2F2F2',
  BRANCO:      '#FFFFFF',
  TEXTO_CLARO: '#FFFFFF',
  TEXTO_ESCURO:'#1A1A2E'
};
 
// ================================================
// MENU
// ================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Sistema')
    .addItem('🔄 Atualizar Informações', 'atualizarTudo')
    .addToUi();
}
 
// ================================================
// ONEDIT — dispara automaticamente ao editar
// IMPORTANTE: só recalcula colunas K, L, M, N, O
// Nunca toca em A-J (dados do usuário)
// ================================================
function onEdit(e) {
  var sheet = e.range.getSheet();
  var col   = e.range.getColumn();
  var linha = e.range.getRow();
 
  if (linha === 1) return;
 
  // Na aba PROJETOS: recalcula apenas se o usuário editou colunas de INPUT (A a J e P, Q)
  // Não recalcula se ele editou as colunas que o script mesmo escreve (K-O)
  if (sheet.getName() === ABA_PROJETOS) {
    var colunasInput = [P.ID, P.PROGRAMA, P.TIPO,
                        P.INICIO_PLANEJADO, P.INICIO_REAL,
                        P.FIM_PLANEJADO, P.FIM_REAL,
                        P.RESPONSAVEL, P.CONTATO,
                        P.OBSERVACOES, P.DESCRICAO];
    if (colunasInput.indexOf(col) === -1) return; // editou coluna de output → ignora
 
    var id = sheet.getRange(linha, P.ID).getValue();
    if (!id) return;
    _recalcularLinhaProjeto(sheet, linha);
    return;
  }
 
  // Na aba ETAPAS: quando STATUS muda, grava DATA_STATUS
  if (sheet.getName() === ABA_ETAPAS && col === E.STATUS) {
  var novoStatus = e.range.getValue();
  // Sempre registra data ao alterar status
  if (novoStatus) {
    sheet.getRange(linha, E.DATA_STATUS).setValue(new Date());
  } else {
    sheet.getRange(linha, E.DATA_STATUS).clearContent();
  }
}
}

// ================================================
// RECALCULAR UMA LINHA DE PROJETO
// Lê apenas as colunas necessárias, escreve apenas K, L, M, N, O
// NUNCA lê nem escreve colunas A-J (preserva Ciclo, datas do usuário)
// ================================================
function _recalcularLinhaProjeto(sheetProjetos, linhaSheet) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var abaEtapas = ss.getSheetByName(ABA_ETAPAS);
  var etapas    = abaEtapas.getDataRange().getValues();
  var hoje      = new Date();
  hoje.setHours(0, 0, 0, 0);
 
  // Lê apenas os campos necessários individualmente para evitar confusão de índices
  var id              = sheetProjetos.getRange(linhaSheet, P.ID).getValue();
  var inicioPlanejado = sheetProjetos.getRange(linhaSheet, P.INICIO_PLANEJADO).getValue();
  var inicioReal      = sheetProjetos.getRange(linhaSheet, P.INICIO_REAL).getValue();
  var fimPlanejado    = sheetProjetos.getRange(linhaSheet, P.FIM_PLANEJADO).getValue();
  var fimReal         = sheetProjetos.getRange(linhaSheet, P.FIM_REAL).getValue();
 
  if (!id) return;
 
  // --- % TEMPO ---
  var percTempo           = 0;
  var encerradoPorFimReal = false;
 
  if (fimReal instanceof Date && !isNaN(fimReal)) {
    percTempo           = 100;
    encerradoPorFimReal = true;
  } else {
    var dtInicio = (inicioReal instanceof Date && !isNaN(inicioReal))
      ? new Date(inicioReal)
      : ((inicioPlanejado instanceof Date && !isNaN(inicioPlanejado)) ? new Date(inicioPlanejado) : null);
    var dtFim = (fimPlanejado instanceof Date && !isNaN(fimPlanejado))
      ? new Date(fimPlanejado) : null;
 
    if (dtInicio && dtFim) {
      dtInicio.setHours(0, 0, 0, 0);
      dtFim.setHours(0, 0, 0, 0);
      var duracaoTotal   = dtFim - dtInicio;
      var duracaoPassada = hoje - dtInicio;
      if (duracaoTotal <= 0)            percTempo = 0;
      else if (duracaoPassada <= 0)     percTempo = 0;
      else if (hoje >= dtFim)           percTempo = 100;
      else percTempo = Math.round((duracaoPassada / duracaoTotal) * 100);
    }
  }
 
  // --- % ETAPAS ---
  var resultado = _calcularPercEtapas(etapas, id, hoje);
  var percEtapas = resultado.percEtapas;
  var temAtrasada = resultado.temAtrasada;
  var concluidas  = resultado.concluidas;
  var totalEtapas = resultado.totalEtapas;
 
  var percGeral = Math.round((percEtapas + percTempo) / 2);
 
  // --- STATUS ---
  var status;
  if (encerradoPorFimReal || percGeral >= 100) {
    status = 'Encerrado';
  } else if (temAtrasada) {
    status = 'Atrasado';
  } else if (percTempo > 0 || concluidas > 0 || percEtapas > 0) {
  status = 'Em Execução';
}
 
  // --- ESCREVE APENAS as colunas de output (K, L, M, N, O) ---
  sheetProjetos.getRange(linhaSheet, P.STATUS)            .setValue(status);
  sheetProjetos.getRange(linhaSheet, P.PERC_ETAPAS)       .setValue(percEtapas);
  sheetProjetos.getRange(linhaSheet, P.PERC_TEMPO)        .setValue(percTempo);
  sheetProjetos.getRange(linhaSheet, P.PERC_GERAL)        .setValue(percGeral);
  sheetProjetos.getRange(linhaSheet, P.ULTIMA_ATUALIZACAO).setValue(new Date());
}
 
// ================================================
// HELPER: calcula % de etapas de um projeto
// ================================================
function _calcularPercEtapas(etapas, id, hoje) {
  var minhasEtapas = [];
  for (var i = 1; i < etapas.length; i++) {
    if (etapas[i][E.ID_PROJETO - 1] === id) minhasEtapas.push(etapas[i]);
  }
 
  var totalEtapas  = minhasEtapas.length;
  var concluidas   = 0;
  var temAtrasada  = false;
  var somaPerc     = 0;
  var etapasComQtd = 0;
 
  for (var j = 0; j < minhasEtapas.length; j++) {
    var et      = minhasEtapas[j];
    var st      = String(et[E.STATUS - 1] || '');
    var fimEt   = et[E.FIM_PLANEJADO - 1];
    var qtdPrev = et[E.QTD_PREVISTA - 1];
    var qtdAtual= et[E.QTD_ATUAL - 1];
 
    // Considera "concluída" qualquer etapa com status do grupo de conclusão
    var ehConcluida = STATUS_CONCLUIDO.indexOf(st) !== -1;
 
    if (ehConcluida) {
      concluidas++;
    } else if (fimEt instanceof Date && !isNaN(fimEt)) {
      var df = new Date(fimEt);
      df.setHours(0, 0, 0, 0);
      if (hoje > df) temAtrasada = true;
    }
 
    // Se tem qtd prevista e atual, usa proporção
    if (qtdPrev && Number(qtdPrev) > 0 && qtdAtual !== '' && qtdAtual !== null) {
      somaPerc += (Number(qtdAtual) / Number(qtdPrev)) * 100;
      etapasComQtd++;
    }
  }
 
  var percEtapas;
  if (etapasComQtd > 0) {
    // Média das proporções de cada etapa (divide por total para penalizar etapas sem qtd)
    percEtapas = Math.round(somaPerc / totalEtapas);
  } else {
    percEtapas = totalEtapas > 0
      ? Math.round((concluidas / totalEtapas) * 100)
      : 0;
  }
 
  return {
    percEtapas: Math.min(100, Math.max(0, percEtapas)),
    temAtrasada: temAtrasada,
    concluidas: concluidas,
    totalEtapas: totalEtapas
  };
}
 
// ================================================
// ATUALIZAR TUDO
// ================================================
function atualizarTudo() {
  atualizarProjetos();
  atualizarDashboard();
  SpreadsheetApp.getUi().alert('✅ Atualização Concluída!');
}
 
// ================================================
// PROJETOS — recalcula % Etapas, % Tempo, % Geral e Status
// CUIDADO: usa getRange individual por coluna para NÃO confundir
// colunas de data com números seriais
// ================================================
function atualizarProjetos() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var abaProjetos = ss.getSheetByName(ABA_PROJETOS);
  var abaEtapas   = ss.getSheetByName(ABA_ETAPAS);
 
  if (!abaProjetos || !abaEtapas) {
    Logger.log('Abas não encontradas.');
    return;
  }
 
  var totalLinhas = abaProjetos.getLastRow();
  var etapas      = abaEtapas.getDataRange().getValues();
  var hoje        = new Date();
  hoje.setHours(0, 0, 0, 0);
 
  // Lê apenas colunas de INPUT em lote (mais rápido que célula por célula)
  // Colunas A (ID), E (Início Plan), F (Início Real), G (Fim Plan), H (Fim Real)
  if (totalLinhas < 2) return;
 
  var colIds     = abaProjetos.getRange(2, P.ID,              totalLinhas - 1, 1).getValues();
  var colIniPlan = abaProjetos.getRange(2, P.INICIO_PLANEJADO, totalLinhas - 1, 1).getValues();
  var colIniReal = abaProjetos.getRange(2, P.INICIO_REAL,      totalLinhas - 1, 1).getValues();
  var colFimPlan = abaProjetos.getRange(2, P.FIM_PLANEJADO,    totalLinhas - 1, 1).getValues();
  var colFimReal = abaProjetos.getRange(2, P.FIM_REAL,         totalLinhas - 1, 1).getValues();
 
  // Arrays para gravar em lote
  var outStatus = [], outPercEt = [], outPercTp = [], outPercGe = [], outData = [];
 
  for (var i = 0; i < colIds.length; i++) {
    var id = colIds[i][0];
    if (!id) {
      outStatus.push(['']);
      outPercEt.push(['']);
      outPercTp.push(['']);
      outPercGe.push(['']);
      outData.push(['']);
      continue;
    }
 
    var inicioPlanejado = colIniPlan[i][0];
    var inicioReal      = colIniReal[i][0];
    var fimPlanejado    = colFimPlan[i][0];
    var fimReal         = colFimReal[i][0];
 
    // --- % TEMPO ---
    var percTempo           = 0;
    var encerradoPorFimReal = false;
 
    if (fimReal instanceof Date && !isNaN(fimReal)) {
      percTempo           = 100;
      encerradoPorFimReal = true;
    } else {
      var dtInicio = (inicioReal instanceof Date && !isNaN(inicioReal))
        ? new Date(inicioReal)
        : ((inicioPlanejado instanceof Date && !isNaN(inicioPlanejado)) ? new Date(inicioPlanejado) : null);
      var dtFim = (fimPlanejado instanceof Date && !isNaN(fimPlanejado))
        ? new Date(fimPlanejado) : null;
 
      if (dtInicio && dtFim) {
        dtInicio.setHours(0, 0, 0, 0);
        dtFim.setHours(0, 0, 0, 0);
        var duracaoTotal   = dtFim - dtInicio;
        var duracaoPassada = hoje - dtInicio;
        if (duracaoTotal <= 0)        percTempo = 0;
        else if (duracaoPassada <= 0) percTempo = 0;
        else if (hoje >= dtFim)       percTempo = 100;
        else percTempo = Math.round((duracaoPassada / duracaoTotal) * 100);
      }
    }
 
    // --- % ETAPAS ---
    var resultado   = _calcularPercEtapas(etapas, id, hoje);
    var percEtapas  = resultado.percEtapas;
    var temAtrasada = resultado.temAtrasada;
    var concluidas  = resultado.concluidas;
 
    var percGeral = Math.round((percEtapas + percTempo) / 2);
 
    // --- STATUS ---
    var status;
    if (encerradoPorFimReal || percGeral >= 100) {
      status = 'Encerrado';
    } else if (temAtrasada) {
      status = 'Atrasado';
    } else if (percTempo > 0 || concluidas > 0) {
      status = 'Em Execução';
    } else {
      status = 'Planejamento';
    }
 
    outStatus.push([status]);
    outPercEt.push([percEtapas]);
    outPercTp.push([percTempo]);
    outPercGe.push([percGeral]);
    outData.push([hoje]);
  }
 
  // Grava tudo de uma vez — apenas colunas K, L, M, N, O
  var n = outStatus.length;
  if (n === 0) return;
  abaProjetos.getRange(2, P.STATUS,            n, 1).setValues(outStatus);
  abaProjetos.getRange(2, P.PERC_ETAPAS,       n, 1).setValues(outPercEt);
  abaProjetos.getRange(2, P.PERC_TEMPO,        n, 1).setValues(outPercTp);
  abaProjetos.getRange(2, P.PERC_GERAL,        n, 1).setValues(outPercGe);
  abaProjetos.getRange(2, P.ULTIMA_ATUALIZACAO,n, 1).setValues(outData)
    .setNumberFormat('dd/MM/yyyy');
 
  // Formata % como número com símbolo (sem converter para decimal)
  abaProjetos.getRange(2, P.PERC_ETAPAS, n, 1).setNumberFormat('0"%"');
  abaProjetos.getRange(2, P.PERC_TEMPO,  n, 1).setNumberFormat('0"%"');
  abaProjetos.getRange(2, P.PERC_GERAL,  n, 1).setNumberFormat('0"%"');
}
  
// ================================================
// DASHBOARD
// ================================================
function atualizarDashboard() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var abaDash = ss.getSheetByName(ABA_DASHBOARD);
  if (!abaDash) abaDash = ss.insertSheet(ABA_DASHBOARD);
 
  abaDash.clearContents();
  abaDash.clearFormats();
  var charts = abaDash.getCharts();
  charts.forEach(function(chart) {
  abaDash.removeChart(chart);
});
 
  abaDash.getRange(1, 1, 80, 14).setBackground(COR.CINZA_CLARO);
 
  var projetos = ss.getSheetByName(ABA_PROJETOS).getDataRange().getValues();
  var etapas   = ss.getSheetByName(ABA_ETAPAS).getDataRange().getValues();
 
  var larguras = [220, 100, 100, 100, 20, 140, 140, 140, 140, 140, 140, 140];
  for (var c = 0; c < larguras.length; c++) {
    abaDash.setColumnWidth(c + 1, larguras[c]);
  }
 
  _painelResumo(abaDash, projetos);
  _dadosStatusProjetos(abaDash, projetos);
  var proximaLinha = _dadosProgressoProjetos(abaDash, projetos);
  var linhaEtapas  = _dadosEtapas(abaDash, etapas, proximaLinha);
 
  _graficoStatus(abaDash);
  _graficoProgresso(abaDash, projetos);
  _graficoEtapas(abaDash, linhaEtapas);
}
 
function _painelResumo(aba, projetos) {
  var hoje = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'
  );
 
  var total = 0, emExecucao = 0, atrasados = 0, encerrados = 0, planejamento = 0;
  for (var i = 1; i < projetos.length; i++) {
    var linha = projetos[i];
    if (!linha[P.ID - 1]) continue;
    total++;
    var s = linha[P.STATUS - 1];
    if (s === 'Em Execução')  emExecucao++;
    if (s === 'Atrasado')     atrasados++;
    if (s === 'Encerrado')    encerrados++;
    if (s === 'Planejamento') planejamento++;
  }
 
  var titulo = aba.getRange(1, 1, 1, 4);
  titulo.merge()
    .setValue('📊 PAINEL DE PROJETOS — SMMF')
    .setFontSize(15).setFontWeight('bold')
    .setFontColor(COR.TEXTO_CLARO)
    .setBackground(COR.PRIMARIO)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  aba.setRowHeight(1, 36);
 
  aba.getRange(2, 1, 1, 4).merge()
    .setValue('Atualizado em: ' + hoje)
    .setFontColor('#555555').setFontSize(9)
    .setBackground(COR.CINZA_CLARO)
    .setHorizontalAlignment('center');
 
  var cabsVisiveis = ['Total', 'Em Execução', 'Atrasados', 'Encerrados'];
  var valsVisiveis = [total,   emExecucao,    atrasados,   encerrados];
  var coresVisiveis= [COR.PRIMARIO, COR.AMARELO, COR.VERMELHO, COR.VERDE];
 
  for (var c = 0; c < cabsVisiveis.length; c++) {
    var colIdx = c + 1;
    aba.getRange(3, colIdx)
      .setValue(cabsVisiveis[c])
      .setFontWeight('bold').setFontSize(9)
      .setFontColor(COR.TEXTO_CLARO)
      .setBackground(coresVisiveis[c])
      .setHorizontalAlignment('center');
 
    aba.getRange(4, colIdx)
      .setValue(valsVisiveis[c])
      .setFontSize(20).setFontWeight('bold')
      .setFontColor(coresVisiveis[c])
      .setBackground(COR.BRANCO)
      .setHorizontalAlignment('center');
    aba.setRowHeight(4, 40);
  }
 
  aba.getRange(5, 1, 1, 4).merge()
    .setValue('Planejamento: ' + planejamento + ' projeto(s)')
    .setFontSize(9).setFontColor('#555555')
    .setBackground(COR.CINZA_CLARO)
    .setHorizontalAlignment('left');
 
  aba.getRange(6, 1, 1, 4).setBackground(COR.PRIMARIO);
  aba.setRowHeight(6, 4);
}
 
function _dadosStatusProjetos(aba, projetos) {
  aba.getRange(7, 1, 1, 2).merge()
    .setValue('STATUS DOS PROGRAMAS')
    .setFontWeight('bold').setFontSize(9)
    .setFontColor(COR.TEXTO_CLARO)
    .setBackground(COR.SECUNDARIO)
    .setHorizontalAlignment('center');
 
  var cabCols = ['Status', 'Qtd'];
  for (var h = 0; h < cabCols.length; h++) {
    aba.getRange(8, h + 1)
      .setValue(cabCols[h])
      .setFontWeight('bold').setFontSize(9)
      .setFontColor(COR.TEXTO_ESCURO)
      .setBackground(COR.DESTAQUE)
      .setHorizontalAlignment('center');
  }
  _bordaBranca(aba, 8, 1, 1, 2);
 
  var cont = { 'Planejamento': 0, 'Em Execução': 0, 'Atrasado': 0, 'Encerrado': 0 };
  for (var i = 1; i < projetos.length; i++) {
    var s = projetos[i][P.STATUS - 1];
    if (cont.hasOwnProperty(s)) cont[s]++;
  }
 
  var coresStatus = {
    'Planejamento': COR.AZUL,
    'Em Execução':  COR.AMARELO,
    'Atrasado':     COR.VERMELHO,
    'Encerrado':    COR.VERDE
  };
 
  var linha = 9;
  for (var status in cont) {
    aba.getRange(linha, 1)
      .setValue(status).setFontSize(9)
      .setFontColor(coresStatus[status] || COR.TEXTO_ESCURO)
      .setBackground(COR.BRANCO).setFontWeight('bold');
    aba.getRange(linha, 2)
      .setValue(cont[status]).setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground(COR.BRANCO);
    _bordaBranca(aba, linha, 1, 1, 2);
    linha++;
  }
}
 
function _dadosProgressoProjetos(aba, projetos) {
  aba.getRange(14, 1, 1, 4).merge()
    .setValue('PROGRESSO POR PROGRAMA')
    .setFontWeight('bold').setFontSize(9)
    .setFontColor(COR.TEXTO_CLARO)
    .setBackground(COR.SECUNDARIO)
    .setHorizontalAlignment('center');
 
  var cabCols = ['Programa', '% Etapas', '% Tempo', '% Geral'];
  for (var h = 0; h < cabCols.length; h++) {
    aba.getRange(15, h + 1)
      .setValue(cabCols[h])
      .setFontWeight('bold').setFontSize(9)
      .setFontColor(COR.TEXTO_ESCURO)
      .setBackground(COR.DESTAQUE)
      .setHorizontalAlignment('center');
  }
  _bordaBranca(aba, 15, 1, 1, 4);
 
  var linha = 16;
  for (var i = 1; i < projetos.length; i++) {
    var proj = projetos[i];
    if (!proj[P.ID - 1]) continue;
 
    var nome = String(proj[P.PROGRAMA - 1] || '');
    if (nome.length > 30) nome = nome.substring(0, 28) + '…';
 
    // Lê os percentuais como número puro
    // A coluna no sheet pode ter formato '0"%"' mas getValues() retorna o número
    var pEtapas = Number(proj[P.PERC_ETAPAS - 1]) || 0;
    var pTempo  = Number(proj[P.PERC_TEMPO  - 1]) || 0;
    var pGeral  = Number(proj[P.PERC_GERAL  - 1]) || 0;
 
    aba.getRange(linha, 1).setValue(nome).setFontSize(9).setBackground(COR.BRANCO);
    aba.getRange(linha, 2).setValue(pEtapas).setFontSize(9)
      .setHorizontalAlignment('center').setBackground(COR.BRANCO)
      .setNumberFormat('0"%"');
    aba.getRange(linha, 3).setValue(pTempo).setFontSize(9)
      .setHorizontalAlignment('center').setBackground(COR.BRANCO)
      .setNumberFormat('0"%"');
    aba.getRange(linha, 4).setValue(pGeral).setFontSize(9)
      .setHorizontalAlignment('center').setBackground(COR.BRANCO)
      .setNumberFormat('0"%"');
 
    _bordaBranca(aba, linha, 1, 1, 4);
    linha++;
  }
 
  return linha;
}
 
function _dadosEtapas(aba, etapas, linhaInicio) {
  var LINHA_ETAPAS = linhaInicio + 2;
 
  aba.getRange(LINHA_ETAPAS - 1, 1, 1, 2).merge()
    .setValue('STATUS DAS ETAPAS')
    .setFontWeight('bold').setFontSize(9)
    .setFontColor(COR.TEXTO_CLARO)
    .setBackground(COR.SECUNDARIO)
    .setHorizontalAlignment('center');
 
  aba.getRange(LINHA_ETAPAS, 1).setValue('Status da Etapa')
    .setFontWeight('bold').setFontSize(9)
    .setBackground(COR.DESTAQUE).setFontColor(COR.TEXTO_ESCURO);
  aba.getRange(LINHA_ETAPAS, 2).setValue('Qtd')
    .setFontWeight('bold').setFontSize(9)
    .setBackground(COR.DESTAQUE).setFontColor(COR.TEXTO_ESCURO)
    .setHorizontalAlignment('center');
  _bordaBranca(aba, LINHA_ETAPAS, 1, 1, 2);
 
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
 
  var cont = { 'Concluído': 0, 'Em andamento': 0, 'Vencida': 0, 'Não iniciada': 0 };
 
  for (var i = 1; i < etapas.length; i++) {
    var et = etapas[i];
    if (!et[E.ID_PROJETO - 1]) continue;
 
    var status   = String(et[E.STATUS - 1] || '');
    var fimEt    = et[E.FIM_PLANEJADO - 1];
    var inicioEt = et[E.INICIO_PLANEJADO - 1];
 
    var ehConcluida = STATUS_CONCLUIDO.indexOf(status) !== -1;
    var ehAndamento = STATUS_EM_ANDAMENTO.indexOf(status) !== -1;
 
    if (ehConcluida) {
      cont['Concluído']++;
    } else if (fimEt instanceof Date && !isNaN(fimEt) && hoje > new Date(fimEt)) {
      cont['Vencida']++;
    } else if (ehAndamento || (inicioEt instanceof Date && !isNaN(inicioEt) && hoje >= new Date(inicioEt))) {
      cont['Em andamento']++;
    } else {
      cont['Não iniciada']++;
    }
  }
 
  var coresEt = {
    'Concluído':    COR.VERDE,
    'Em andamento': COR.AMARELO,
    'Vencida':      COR.VERMELHO,
    'Não iniciada': COR.AZUL
  };
 
  var linha = LINHA_ETAPAS + 1;
  for (var s in cont) {
  if (cont[s] === 0) continue; // NÃO mostra categorias vazias
    aba.getRange(linha, 1).setValue(s).setFontSize(9)
      .setFontColor(coresEt[s] || COR.TEXTO_ESCURO)
      .setFontWeight('bold').setBackground(COR.BRANCO);
    aba.getRange(linha, 2).setValue(cont[s]).setFontSize(11)
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground(COR.BRANCO);
    _bordaBranca(aba, linha, 1, 1, 2);
    linha++;
  }
 
  return LINHA_ETAPAS;
}
 
function _graficoStatus(aba) {
  var rangeStatus = aba.getRange(8, 1, 5, 2);
 
  aba.insertChart(
    aba.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(rangeStatus)
      .setPosition(1, 6, 0, 0)
      .setOption('title', 'Status dos Programas')
      .setOption('titleTextStyle', { color: COR.PRIMARIO, fontSize: 12, bold: true })
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: '' })
      .setOption('vAxis', { title: 'Qtd', minValue: 0, format: '0',
                            gridlines: { color: COR.CINZA_CLARO, count: 5 } })
      .setOption('backgroundColor', COR.CINZA_CLARO)
      .setOption('chartArea', { backgroundColor: COR.BRANCO,
                                left: 50, top: 40, width: '80%', height: '70%' })
      .setOption('width', 420).setOption('height', 260)
      .setOption('colors', [COR.AZUL, COR.AMARELO, COR.VERMELHO, COR.VERDE])
      .build()
  );
}
 
function _graficoProgresso(aba, projetos) {
  var total = 0;
  for (var i = 1; i < projetos.length; i++) {
    if (projetos[i][P.ID - 1]) total++;
  }
  if (total === 0) return;

  var rangeGrafico = aba.getRange(16, 1, total, 4);

  // Remove gráficos antigos dessa área
  aba.getCharts().forEach(function(chart) {
    var pos = chart.getContainerInfo();
    if (pos.getAnchorColumn() >= 9) {
      aba.removeChart(chart);
    }
  });

  aba.insertChart(
    aba.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(rangeGrafico)
      .setPosition(1, 9, 0, 0) // ← COMEÇA NA COLUNA J (9)
      .setOption('title', 'Progresso por Programa (%)')
      .setOption('legend', { position: 'top' })
      .setOption('hAxis', { minValue: 0, maxValue: 100 })
      .setOption('width', 520)
      .setOption('height', 320)
      .setOption('colors', [COR.SECUNDARIO, COR.AZUL, COR.PRIMARIO])
      .build()
  );
}
 
function _graficoEtapas(aba, linhaEtapas) {
  // Cabeçalho está em linhaEtapas
  // Dados reais começam na linha seguinte
  var rangeEtapas = aba.getRange(linhaEtapas + 1, 1, 4, 2);

  // Remove gráficos antigos dessa área (evita duplicação bugada)
  var charts = aba.getCharts();
  charts.forEach(function(chart) {
    var pos = chart.getContainerInfo();
    if (pos.getAnchorColumn() === 6 && pos.getAnchorRow() >= 14) {
      aba.removeChart(chart);
    }
  });

  aba.insertChart(
    aba.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(rangeEtapas)
      .setPosition(14, 6, 0, 0)
      .setOption('title', 'Status das Etapas')
      .setOption('legend', { position: 'none' })
      .setOption('vAxis', { minValue: 0 })
      .setOption('width', 420)
      .setOption('height', 260)
      .build()
  );
}
 
// ================================================
// HELPER: bordas brancas
// ================================================
function _bordaBranca(aba, linha, col, numLinhas, numCols) {
  aba.getRange(linha, col, numLinhas, numCols)
    .setBorder(true, true, true, true, true, true,
               COR.BRANCO, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
 
// ================================================
// GATILHO DIÁRIO
// ================================================
function rodarDiariamente() {
  atualizarProjetos();
  atualizarDashboard();
}