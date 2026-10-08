/**
 * ============================================================
 *  🛠️ SETUP — monta a planilha inteira.
 *  Execute `setup` uma vez. Rodar de novo reconstrói os painéis
 *  (Resumo, Investimentos, Lançamentos) SEM apagar seus dados:
 *  Configurações e Banco de Dados existentes são preservados.
 * ============================================================
 */


const PADRAO_CATEGORIAS = [
  ['Salário', 'Receita', ''], ['Renda extra', 'Receita', ''], ['Rendimentos', 'Receita', ''],
  ['Reembolsos', 'Receita', ''],
  ['Alimentação', 'Despesa', 1200], ['Moradia', 'Despesa', 1500], ['Transporte', 'Despesa', 400],
  ['Saúde', 'Despesa', 300], ['Lazer', 'Despesa', 400], ['Compras', 'Despesa', 500],
  ['Contas', 'Despesa', 500], ['Assinaturas', 'Despesa', 100], ['Educação', 'Despesa', 300],
  ['Cuidados pessoais', 'Despesa', 200], ['Outros', 'Despesa', 200],
  ['Investimentos', 'Investimento', ''], ['Metas', 'Investimento', ''],
  ['Pagamento de fatura', 'Transferência', ''], ['Entre contas', 'Transferência', '']
];

const PADRAO_SUBCATEGORIAS = [
  ['Salário', 'Mensal'], ['Salário', '13º'], ['Salário', 'Férias'],
  ['Renda extra', 'Freelance'], ['Renda extra', 'Vendas'],
  ['Alimentação', 'Mercado'], ['Alimentação', 'Restaurante'], ['Alimentação', 'Delivery'], ['Alimentação', 'Padaria'],
  ['Moradia', 'Aluguel'], ['Moradia', 'Condomínio'], ['Moradia', 'Manutenção'],
  ['Transporte', 'Combustível'], ['Transporte', 'App de transporte'], ['Transporte', 'Transporte público'],
  ['Saúde', 'Farmácia'], ['Saúde', 'Consultas'], ['Saúde', 'Plano de saúde'], ['Saúde', 'Academia'],
  ['Lazer', 'Passeios'], ['Lazer', 'Viagens'], ['Lazer', 'Cinema e shows'],
  ['Compras', 'Roupas'], ['Compras', 'Eletrônicos'], ['Compras', 'Casa'], ['Compras', 'Presentes'],
  ['Contas', 'Internet'], ['Contas', 'Energia'], ['Contas', 'Telefone'], ['Contas', 'Água'],
  ['Assinaturas', 'Streaming'], ['Assinaturas', 'Aplicativos'],
  ['Educação', 'Cursos'], ['Educação', 'Livros'],
  ['Investimentos', 'Cripto'], ['Investimentos', 'FIIs'], ['Investimentos', 'Renda fixa']
];

const PADRAO_CONTAS = ['Conta corrente', 'Carteira', 'Conta digital', 'Outros'];
const PADRAO_FORMAS = ['Pix', 'Débito', 'Crédito', 'Dinheiro', 'Boleto', 'Transferência'];

const PADRAO_CARTOES = [
  ['Cartão 1', 5000, 25, 3, 'Conta corrente'],
  ['Cartão 2', 3000, 3, 10, 'Conta digital']
];

const PADRAO_FIXAS = [
  ['Internet', 'Contas', 100, 10, 'Conta corrente', 'Débito', ''],
  ['Energia', 'Contas', 180, 15, 'Conta corrente', 'Boleto', ''],
  ['Telefone', 'Contas', 60, 12, 'Conta corrente', 'Débito', ''],
  ['Academia', 'Saúde', 110, 5, 'Conta digital', 'Pix', ''],
  ['Assinaturas', 'Assinaturas', 55, 20, 'Conta corrente', 'Débito', ''],
  ['Outros', 'Outros', 50, 25, 'Conta corrente', 'Pix', '']
];

const PADRAO_ATIVOS = [
  ['Bitcoin', 'Cripto', 'CURRENCY:BTCBRL', ''],
  ['MXRF11', 'FII', 'BVMF:MXRF11', ''],
  ['HGLG11', 'FII', 'BVMF:HGLG11', '']
];

/* ============================================================
 *  ORQUESTRAÇÃO
 * ============================================================ */

function setup() {
  const ss = planilha_();
  ss.setSpreadsheetLocale('pt_BR');
  ss.setSpreadsheetTimeZone(FUSO);

  SpreadsheetApp.flush();

  // Só preserva Configurações / Banco de Dados se já estiverem MONTADAS (não basta a aba existir)
  const cfgExistente = ss.getSheetByName(ABA.CFG);
  const bdExistente = ss.getSheetByName(ABA.BD);
  const cfgNova = !cfgExistente || cfgExistente.getRange('B10').getValue() !== 'Categoria';
  const bdNovo = !bdExistente || bdExistente.getRange('A1').getValue() !== 'ID';

  const sh = {
    res: prepararAba_(ss, ABA.RESUMO, true),
    inv: prepararAba_(ss, ABA.INV, true),
    lanc: prepararAba_(ss, ABA.LANC, true),
    cfg: prepararAba_(ss, ABA.CFG, cfgNova),
    bd: prepararAba_(ss, ABA.BD, bdNovo)
  };
  ordenarAbas_(ss);
  garantirColunas_(sh.cfg, CFG.AUX_ATIVOS);
  garantirColunas_(sh.res, 20);
  garantirColunas_(sh.bd, BD_NCOLS);
  if (sh.bd.getMaxRows() < 3000) sh.bd.insertRowsAfter(sh.bd.getMaxRows(), 3000 - sh.bd.getMaxRows());

  criarNomes_(ss, sh);
  SpreadsheetApp.flush();

  if (cfgNova) montarConfiguracoes_(sh.cfg);
  else garantirMetas_(sh.cfg);
  montarAuxConfig_(sh.cfg);
  if (bdNovo) montarBancoDeDados_(sh.bd);
  montarLancamentos_(sh.lanc);
  montarResumo_(sh.res, sh.cfg);
  montarInvestimentos_(sh.inv);
  removerAbasVazias_(ss);

  try {
    instalarGatilhos_();
  } catch (e) {
    Logger.log('Não foi possível criar o gatilho mensal: ' + e);
  }

  SpreadsheetApp.flush();
  verificarFormulas_(ss, sh.res);
  renderResumo_();
  registrarSnapshot();
  ss.setActiveSheet(sh.res);
  ss.toast('Sua planilha está pronta! 💗', 'Finanças', 6);
}

/**
 * Autoteste: confere se as fórmulas estão sendo entendidas pela planilha.
 * Se algo falhar, interrompe com um diagnóstico para enviar ao suporte.
 */
function verificarFormulas_(ss, shRes) {
  const testes = [
    ['soma com vírgula', '=SUM(1,2)', '3'],
    ['intervalo nomeado', '=N(CFG_DIA_INICIO)', String(Number(ss.getSheetByName(ABA.CFG).getRange(CFG_DIA_INICIO_A1).getValue()))],
    ['FILTER com nome', '=IFERROR(INDEX(FILTER(META_NOME,META_NOME<>""),1),"vazio")', null]
  ];
  const celulas = shRes.getRange(20, 20, testes.length, 1);
  celulas.setFormulas(fxs_(testes.map(t => [t[1]])));
  SpreadsheetApp.flush();
  const vistos = celulas.getDisplayValues().map(r => r[0]);
  celulas.clearContent();

  const falhas = testes
    .map((t, i) => ({ nome: t[0], esperado: t[2], visto: vistos[i] }))
    .filter(t => t.visto.charAt(0) === '#' || (t.esperado !== null && t.visto !== t.esperado));
  if (falhas.length) {
    const nomes = ss.getNamedRanges().map(n => n.getName());
    throw new Error('DIAGNÓSTICO — ' + falhas.map(f => f.nome + ': "' + f.visto + '"').join(' | ') +
      ' | localidade: ' + ss.getSpreadsheetLocale() + ' | nomes criados: ' + nomes.length +
      ' | CFG_DIA_INICIO existe: ' + (nomes.indexOf('CFG_DIA_INICIO') >= 0));
  }
}

function prepararAba_(ss, nome, limpar) {
  let sh = ss.getSheetByName(nome);
  if (!sh) return ss.insertSheet(nome);
  if (limpar) limparAba_(sh);
  return sh;
}

function limparAba_(sh) {
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.setConditionalFormatRules([]);
  sh.getBandings().forEach(b => b.remove());
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const tudo = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  tudo.breakApart();
  tudo.clearDataValidations();
  sh.clear();
  sh.showColumns(1, sh.getMaxColumns());
  sh.setFrozenRows(0);
  sh.setRowHeights(1, sh.getMaxRows(), 21);
}

function ordenarAbas_(ss) {
  [ABA.RESUMO, ABA.INV, ABA.LANC, ABA.CFG, ABA.BD].forEach((nome, i) => {
    ss.setActiveSheet(ss.getSheetByName(nome));
    ss.moveActiveSheet(i + 1);
  });
}

function removerAbasVazias_(ss) {
  const nossas = Object.keys(ABA).map(k => ABA[k]);
  ss.getSheets().forEach(s => {
    if (nossas.indexOf(s.getName()) < 0 && s.getLastRow() === 0 && s.getLastColumn() === 0) ss.deleteSheet(s);
  });
}

function garantirColunas_(sh, n) {
  const max = sh.getMaxColumns();
  if (max < n) sh.insertColumnsAfter(max, n - max);
}

function instalarGatilhos_() {
  const existe = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'registrarSnapshot');
  if (!existe) ScriptApp.newTrigger('registrarSnapshot').timeBased().onMonthDay(1).atHour(9).create();
}

/** Intervalos nomeados: deixam as fórmulas legíveis (ex.: SUMIFS(BD_VALOR, BD_TIPO, "Receita")). */
function criarNomes_(ss, sh) {
  const nomes = {};
  const n = CFG_LINHA_FIM - CFG_LINHA_INI + 1;
  const cfgCol = col => sh.cfg.getRange(CFG_LINHA_INI, col, n, 1);

  nomes.CFG_DIA_INICIO = sh.cfg.getRange(CFG_DIA_INICIO_A1);
  nomes.CFG_ANO_INI = sh.cfg.getRange(CFG_ANO_INI_A1);
  nomes.LST_CATEGORIAS = cfgCol(CFG.CAT_NOME);
  nomes.CAT_TIPO = cfgCol(CFG.CAT_TIPO);
  nomes.CAT_ORC = cfgCol(CFG.CAT_ORC);
  nomes.LST_CONTAS = cfgCol(CFG.CONTA);
  nomes.LST_FORMAS = cfgCol(CFG.FORMA);
  nomes.CART_NOME = cfgCol(CFG.CART_NOME);
  nomes.CART_LIMITE = cfgCol(CFG.CART_LIMITE);
  nomes.CART_FECH = cfgCol(CFG.CART_FECH);
  nomes.CART_VENC = cfgCol(CFG.CART_VENC);
  nomes.FIXA_NOME = cfgCol(CFG.FIXA_NOME);
  nomes.FIXA_VALOR = cfgCol(CFG.FIXA_VALOR);
  nomes.ATV_NOME = cfgCol(CFG.ATV_NOME);
  nomes.ATV_TIPO = cfgCol(CFG.ATV_TIPO);
  nomes.ATV_PRECO = cfgCol(CFG.ATV_PRECO);
  nomes.META_NOME = cfgCol(CFG.META_NOME);
  nomes.META_VALOR = cfgCol(CFG.META_VALOR);
  nomes.META_INICIAL = cfgCol(CFG.META_INICIAL);
  nomes.META_PRAZO = cfgCol(CFG.META_PRAZO);
  nomes.AUX_CATS = cfgCol(CFG.AUX_CATS);
  nomes.AUX_SUBS = cfgCol(CFG.AUX_SUBS);
  nomes.AUX_ATIVOS = cfgCol(CFG.AUX_ATIVOS);

  const maxBD = sh.bd.getMaxRows();
  Object.keys(BD_NOMES).forEach(k => { nomes[k] = sh.bd.getRange(2, BD_NOMES[k], maxBD - 1, 1); });

  nomes.RES_REF = sh.res.getRange('T1');
  nomes.RES_INICIO = sh.res.getRange('T2');
  nomes.RES_FIM = sh.res.getRange('T3');
  nomes.FIXAS_PEND = sh.res.getRange('T9');
  nomes.KPI_ENTRADAS = sh.res.getRange('B5');
  nomes.KPI_SAIDAS = sh.res.getRange('E5');
  nomes.KPI_SALDO = sh.res.getRange('J5');
  nomes.KPI_INVESTIDO = sh.res.getRange('N5');
  nomes.KPI_FATURA = sh.res.getRange('B9');
  nomes.KPI_PARCFUT = sh.res.getRange('E9');

  // Atualiza os nomes que já existem e cria os que faltam (sem tentar remover nomes inexistentes)
  const existentes = {};
  ss.getNamedRanges().forEach(nr => { existentes[nr.getName()] = nr; });
  Object.keys(nomes).forEach(k => {
    if (existentes[k]) existentes[k].setRange(nomes[k]);
    else ss.setNamedRange(k, nomes[k]);
  });
}

/* ============================================================
 *  HELPERS DE ESTILO
 * ============================================================ */

function estiloAba_(sh, larguras, corAba) {
  garantirColunas_(sh, larguras.length);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())
    .setFontFamily(FONTE).setFontSize(10).setFontColor(COR.TEXTO).setVerticalAlignment('middle');
  larguras.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setHiddenGridlines(true);
  sh.setTabColor(corAba);
}

/** Faixa ameixa no topo, com título e subtítulo. */
function faixaTitulo_(sh, ultimaCol, colTitulo, titulo, subtitulo) {
  sh.getRange(1, 2, 2, ultimaCol - 1).setBackground(COR.AMEIXA);
  sh.getRange(1, 2, 1, colTitulo - 1).merge().setValue(titulo)
    .setFontSize(20).setFontWeight('bold').setFontColor(COR.BRANCO);
  const sub = sh.getRange(2, 2, 1, colTitulo - 1).merge().setFontColor(COR.ROSE).setFontSize(10);
  if (String(subtitulo).charAt(0) === '=') sub.setFormula(fx_(subtitulo));
  else sub.setValue(subtitulo);
  sh.setRowHeight(1, 46);
  sh.setRowHeight(2, 28);
}

function secao_(sh, a1, texto) {
  const r = sh.getRange(a1);
  r.merge().setValue(texto).setFontSize(12).setFontWeight('bold').setFontColor(COR.AMEIXA)
    .setVerticalAlignment('bottom')
    .setBorder(null, null, true, null, null, null, COR.ROSA, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(r.getRow(), 30);
}

/** Linha de cabeçalho de tabela. `celulas` = [[a1, texto], ...] (a1 pode ser um intervalo a mesclar). */
function cabecalho_(sh, celulas) {
  celulas.forEach(par => {
    const r = sh.getRange(par[0]);
    if (r.getNumColumns() > 1) r.merge();
    r.setValue(par[1]);
    r.setBackground(COR.ROSE_CLARO).setFontColor(COR.AMEIXA).setFontWeight('bold')
      .setFontSize(9).setHorizontalAlignment('center').setWrap(true);
  });
}

/** Zebra (linhas alternadas) em uma área. */
function zebra_(sh, linha, col, nLin, nCol) {
  const cores = [];
  for (let i = 0; i < nLin; i++) cores.push(new Array(nCol).fill(i % 2 ? COR.ZEBRA : COR.BRANCO));
  sh.getRange(linha, col, nLin, nCol).setBackgrounds(cores)
    .setBorder(null, null, true, null, null, true, COR.CINZA, SpreadsheetApp.BorderStyle.SOLID);
}

function escrever_(sh, a1, v) {
  const r = sh.getRange(a1);
  if (typeof v === 'string' && v.charAt(0) === '=') r.setFormula(fx_(v));
  else r.setValue(v);
  return r;
}

/** Cartão de KPI com rótulo, valor e nota (3 linhas). */
function card_(sh, linha, colIni, colFim, rotulo, formulaValor, formulaNota, fundo, formato) {
  const larg = colFim - colIni + 1;
  const tudo = sh.getRange(linha, colIni, 3, larg);
  tudo.setBackground(fundo);
  const rot = sh.getRange(linha, colIni, 1, larg).merge();
  rot.setValue(rotulo).setFontSize(9).setFontWeight('bold').setFontColor(COR.AMEIXA_CLARA)
    .setHorizontalAlignment('center').setVerticalAlignment('bottom');
  const val = sh.getRange(linha + 1, colIni, 1, larg).merge();
  val.setFormula(fx_(formulaValor)).setFontSize(18).setFontWeight('bold').setFontColor(COR.AMEIXA)
    .setHorizontalAlignment('center');
  if (formato) val.setNumberFormat(formato);
  const nota = sh.getRange(linha + 2, colIni, 1, larg).merge();
  nota.setFormula(fx_(formulaNota)).setFontSize(8).setFontStyle('italic').setFontColor(COR.TEXTO_SUAVE)
    .setHorizontalAlignment('center').setVerticalAlignment('top');
  tudo.setBorder(true, true, true, true, false, false, COR.BRANCO, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(linha, 24);
  sh.setRowHeight(linha + 1, 38);
  sh.setRowHeight(linha + 2, 22);
}

function grafico_(sh, tipo, intervalos, linha, col, largura, altura, titulo, opcoes) {
  let b = sh.newChart().setChartType(tipo)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setNumHeaders(1)
    .setPosition(linha, col, 4, 4)
    .setOption('title', titulo)
    .setOption('titleTextStyle', { color: COR.AMEIXA, fontSize: 13, bold: true, fontName: FONTE })
    .setOption('fontName', FONTE)
    .setOption('backgroundColor', COR.BRANCO)
    .setOption('legend', { position: 'bottom', textStyle: { color: COR.TEXTO, fontSize: 10 } })
    .setOption('width', largura)
    .setOption('height', altura);
  intervalos.forEach(r => { b = b.addRange(sh.getRange(r)); });
  Object.keys(opcoes || {}).forEach(k => { b = b.setOption(k, opcoes[k]); });
  sh.insertChart(b.build());
}

function validacaoLista_(valores) {
  return SpreadsheetApp.newDataValidation().requireValueInList(valores, true).setAllowInvalid(false).build();
}

function validacaoIntervalo_(range, permitirOutros) {
  return SpreadsheetApp.newDataValidation().requireValueInRange(range, true)
    .setAllowInvalid(!!permitirOutros).build();
}

function regraCor_(ranges, formula, fundo, fonte) {
  let b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(fx_(formula)).setRanges(ranges);
  if (fundo) b = b.setBackground(fundo);
  if (fonte) b = b.setFontColor(fonte);
  return b.build();
}

/* ============================================================
 *  ABA CONFIGURAÇÕES
 * ============================================================ */

function montarConfiguracoes_(sh) {
  const larg = [20, 230, 110, 130, 20, 130, 150, 20, 140, 20, 140, 20, 140, 100, 105, 110, 140, 20,
    130, 120, 110, 100, 130, 110, 130, 20, 110, 90, 160, 100, 100, 20, 130, 110, 150, 100, 20, 130, 130, 130];
  estiloAba_(sh, larg, COR.LAVANDA);
  faixaTitulo_(sh, CFG.META_PRAZO, 11, '⚙️ Configurações',
    'Tudo o que você cadastrar aqui aparece automaticamente nos menus do formulário e no painel.');

  // Parâmetros
  secao_(sh, 'B4:D4', '🧭 Parâmetros');
  escrever_(sh, 'B5', 'Dia de início do mês financeiro');
  escrever_(sh, 'C5', 26);
  escrever_(sh, 'D5', '26 = ciclo de 26 a 25 • 1 = mês do calendário');
  escrever_(sh, 'B6', 'Primeiro ano dos seletores');
  escrever_(sh, 'C6', new Date().getFullYear() - 1);
  escrever_(sh, 'D6', 'Anos disponíveis no Resumo');
  sh.getRange('B5:B6').setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('C5:C6').setBackground(COR.BLUSH).setHorizontalAlignment('center').setFontWeight('bold')
    .setBorder(true, true, true, true, null, null, COR.ROSA, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('D5:D6').setFontSize(8).setFontStyle('italic').setFontColor(COR.TEXTO_SUAVE);
  sh.getRange('C5').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 28)
    .setHelpText('Entre 1 e 28').build());
  sh.getRange('B7:Q7').merge()
    .setValue('💡 Para adicionar um item, escreva na próxima linha vazia da tabela. Para remover, apague o conteúdo da linha. ' +
      'Depois de mudar fechamento/vencimento de cartão ou o ciclo, use o menu 💗 Finanças → Recalcular.')
    .setFontSize(9).setFontStyle('italic').setFontColor(COR.AMEIXA_CLARA).setWrap(true);
  sh.setRowHeight(7, 34);

  const L = CFG_LINHA_INI;
  const n = CFG_LINHA_FIM - CFG_LINHA_INI + 1;
  const blocos = [
    { titulo: '🏷️ Categorias', col: CFG.CAT_NOME, cab: ['Categoria', 'Tipo', 'Orçamento mensal'], dados: PADRAO_CATEGORIAS },
    { titulo: '🔖 Subcategorias', col: CFG.SUB_CAT, cab: ['Categoria', 'Subcategoria'], dados: PADRAO_SUBCATEGORIAS },
    { titulo: '🏦 Contas', col: CFG.CONTA, cab: ['Conta'], dados: PADRAO_CONTAS.map(x => [x]) },
    { titulo: '💸 Formas pgto', col: CFG.FORMA, cab: ['Forma de pagamento'], dados: PADRAO_FORMAS.map(x => [x]) },
    { titulo: '💳 Cartões de crédito', col: CFG.CART_NOME,
      cab: ['Cartão', 'Limite', 'Dia fechamento', 'Dia vencimento', 'Conta de pagamento'], dados: PADRAO_CARTOES },
    { titulo: '📌 Contas fixas', col: CFG.FIXA_NOME,
      cab: ['Conta fixa', 'Categoria', 'Valor previsto', 'Dia vencimento', 'Conta', 'Forma pgto', 'Cartão (se crédito)'],
      dados: PADRAO_FIXAS },
    { titulo: '📈 Ativos', col: CFG.ATV_NOME,
      cab: ['Ativo', 'Tipo', 'Ticker (GOOGLEFINANCE)', 'Preço manual', 'Preço atual'], dados: PADRAO_ATIVOS },
    blocoMetas_(padraoMetas_())
  ];

  blocos.forEach(b => blocoConfig_(sh, b));
  sh.setRowHeight(L - 1, 34);

  // Formatos
  const col = c => sh.getRange(L, c, n, 1);
  [CFG.CAT_ORC, CFG.CART_LIMITE, CFG.FIXA_VALOR, CFG.ATV_MANUAL, CFG.ATV_PRECO]
    .forEach(c => col(c).setNumberFormat(FORMATO_MOEDA));
  formatarMetas_(sh);
  [CFG.CAT_TIPO, CFG.CART_FECH, CFG.CART_VENC, CFG.FIXA_DIA, CFG.ATV_TIPO]
    .forEach(c => col(c).setHorizontalAlignment('center'));

  // Validações (menus suspensos dentro da própria configuração)
  const dia = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 31).setAllowInvalid(false).build();
  col(CFG.CAT_TIPO).setDataValidation(validacaoLista_(TIPOS));
  col(CFG.SUB_CAT).setDataValidation(validacaoIntervalo_(col(CFG.CAT_NOME), true));
  col(CFG.FIXA_CAT).setDataValidation(validacaoIntervalo_(col(CFG.CAT_NOME), true));
  col(CFG.CART_CONTA).setDataValidation(validacaoIntervalo_(col(CFG.CONTA), true));
  col(CFG.FIXA_CONTA).setDataValidation(validacaoIntervalo_(col(CFG.CONTA), true));
  col(CFG.FIXA_FORMA).setDataValidation(validacaoIntervalo_(col(CFG.FORMA), true));
  col(CFG.FIXA_CARTAO).setDataValidation(validacaoIntervalo_(col(CFG.CART_NOME), true));
  [CFG.CART_FECH, CFG.CART_VENC, CFG.FIXA_DIA].forEach(c => col(c).setDataValidation(dia));
}

/** Título, cabeçalho, zebra e dados iniciais de uma tabela da aba Configurações. */
function blocoConfig_(sh, b) {
  const L = CFG_LINHA_INI;
  const n = CFG_LINHA_FIM - CFG_LINHA_INI + 1;
  const w = b.cab.length;
  secao_(sh, sh.getRange(L - 2, b.col, 1, w).getA1Notation(), b.titulo);
  sh.getRange(L - 1, b.col, 1, w).setValues([b.cab]).setBackground(COR.AMEIXA).setFontColor(COR.BRANCO)
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setWrap(true);
  zebra_(sh, L, b.col, n, w);
  if (b.dados.length) {
    const linhas = b.dados.map(r => r.slice(0, w));
    sh.getRange(L, b.col, linhas.length, w).setValues(linhas);
  }
}

function blocoMetas_(dados) {
  return { titulo: '🎯 Metas', col: CFG.META_NOME, cab: ['Meta', 'Valor da meta', 'Já guardado (antes)', 'Prazo'], dados: dados };
}

function formatarMetas_(sh) {
  const n = CFG_LINHA_FIM - CFG_LINHA_INI + 1;
  const col = c => sh.getRange(CFG_LINHA_INI, c, n, 1);
  [CFG.META_VALOR, CFG.META_INICIAL].forEach(c => col(c).setNumberFormat(FORMATO_MOEDA));
  col(CFG.META_PRAZO).setNumberFormat(FORMATO_DATA)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
}

/**
 * Planilhas montadas por versões antigas não tinham a tabela de Metas.
 * Cria a tabela (vazia) sem tocar no resto da aba Configurações.
 */
function garantirMetas_(sh) {
  const L = CFG_LINHA_INI;
  if (sh.getRange(L - 1, CFG.META_NOME).getValue() === 'Meta') return;
  const n = CFG_LINHA_FIM - CFG_LINHA_INI + 1;
  const area = sh.getRange(L - 2, CFG.META_NOME, n + 2, CFG.META_PRAZO - CFG.META_NOME + 1);
  const ocupada = area.getValues().some(r => r.some(v => v !== ''));
  if (ocupada) throw new Error('A área da tabela de Metas (colunas AG:AJ) em Configurações já tem conteúdo. ' +
    'Mova ou apague esse conteúdo e rode "Reconstruir painéis" de novo.');

  [20, 130, 110, 150, 100].forEach((w, i) => sh.setColumnWidth(CFG.META_NOME - 1 + i, w));
  sh.getRange(1, 2, 2, CFG.META_PRAZO - 1).setBackground(COR.AMEIXA);
  area.setFontFamily(FONTE).setFontSize(10).setFontColor(COR.TEXTO).setVerticalAlignment('middle');
  blocoConfig_(sh, blocoMetas_([]));
  formatarMetas_(sh);
}

function padraoMetas_() {
  const a = new Date().getFullYear();
  return [
    ['🚗 Carro', 40000, 0, new Date(a + 2, 11, 1)],
    ['✈️ Viagem', 8000, 0, new Date(a + 1, 6, 1)],
    ['📱 Tablet', 3000, 0, new Date(a + 1, 2, 1)]
  ];
}

/** Colunas auxiliares (ocultas) que alimentam os menus dependentes do formulário. */
function montarAuxConfig_(sh) {
  const L = CFG_LINHA_INI;
  const n = CFG_LINHA_FIM - CFG_LINHA_INI + 1;

  // Preço atual: manual (se preenchido) ou cotação do GOOGLEFINANCE. Regravado a cada setup.
  const formulasPreco = [];
  for (let r = L; r <= CFG_LINHA_FIM; r++) {
    formulasPreco.push(['=IF(AA' + r + '="","",IF(AD' + r + '<>"",AD' + r + ',' +
      'IF(AC' + r + '="","informe o ticker",IFERROR(GOOGLEFINANCE(AC' + r + '),"sem cotação"))))']);
  }
  sh.getRange(L, CFG.ATV_PRECO, n, 1).setFormulas(fxs_(formulasPreco))
    .setNumberFormat(FORMATO_MOEDA).setFontColor(COR.AMEIXA_CLARA);

  sh.showColumns(CFG.AUX_CATS, 3);
  sh.getRange(L, CFG.AUX_CATS, n, 3).clearContent();
  sh.getRange(L - 1, CFG.AUX_CATS, 1, 3).setValues([['aux: categorias do tipo', 'aux: subcategorias', 'aux: ativos e metas']]);
  sh.getRange(L, CFG.AUX_CATS).setFormula(fx_(
    '=IF(\'' + ABA.LANC + '\'!$E$5="",FILTER($B$11:$B$60,$B$11:$B$60<>""),' +
    'IFERROR(FILTER($B$11:$B$60,$C$11:$C$60=\'' + ABA.LANC + '\'!$E$5),""))'));
  sh.getRange(L, CFG.AUX_SUBS).setFormula(fx_(
    '=IFERROR(FILTER($G$11:$G$60,$F$11:$F$60=\'' + ABA.LANC + '\'!$E$6),"")'));
  sh.getRange(L, CFG.AUX_ATIVOS).setFormula(fx_(
    '={IFERROR(FILTER($AG$11:$AG$60,$AG$11:$AG$60<>""),"");IFERROR(FILTER($AA$11:$AA$60,$AA$11:$AA$60<>""),"")}'));
  sh.hideColumns(CFG.AUX_CATS, 3);
}

/* ============================================================
 *  ABA BANCO DE DADOS
 * ============================================================ */

function montarBancoDeDados_(sh) {
  const larg = [110, 130, 90, 200, 100, 120, 120, 120, 100, 100, 100, 80, 60, 75, 70, 80, 110, 100, 110, 90,
    110, 100, 100, 110, 220];
  estiloAba_(sh, larg, COR.AMEIXA_CLARA);
  sh.setHiddenGridlines(false);
  const max = sh.getMaxRows();

  sh.getRange(1, 1, 1, BD_NCOLS).setValues([BD_CABECALHO])
    .setFontWeight('bold').setFontColor(COR.BRANCO).setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);

  const colunas = (c, fmt) => sh.getRange(2, c, max - 1, 1).setNumberFormat(fmt);
  colunas(C.REG, 'dd/mm/yyyy hh:mm');
  colunas(C.DATA, FORMATO_DATA);
  colunas(C.VENC, FORMATO_DATA);
  colunas(C.MESREF, 'mmm/yyyy');
  colunas(C.VALOR, FORMATO_MOEDA);
  colunas(C.VALOR_TOTAL, FORMATO_MOEDA);
  colunas(C.PRECO, FORMATO_MOEDA);
  colunas(C.QTD, '0.########');

  const banda = sh.getRange(1, 1, max, BD_NCOLS).applyRowBanding(SpreadsheetApp.BandingTheme.PINK, true, false);
  banda.setHeaderRowColor(COR.AMEIXA).setFirstRowColor(COR.BRANCO).setSecondRowColor(COR.ZEBRA);

  sh.protect().setDescription('Banco de Dados — alimentado automaticamente pela planilha').setWarningOnly(true);
}

/* ============================================================
 *  ABA LANÇAMENTOS (formulário)
 * ============================================================ */

function montarLancamentos_(sh) {
  estiloAba_(sh, [20, 170, 240, 170, 240, 110, 24, 340], COR.ROSA);
  faixaTitulo_(sh, 8, 7, '📝 Novo lançamento',
    'Preencha os campos, marque a caixinha ✅ e pronto — o resto da planilha se atualiza sozinho.');

  const cfg = planilha_().getSheetByName(ABA.CFG);
  const cfgCol = c => cfg.getRange(CFG_LINHA_INI, c, CFG_LINHA_FIM - CFG_LINHA_INI + 1, 1);

  // ---------- Blocos do formulário ----------
  secao_(sh, 'B4:E4', '🧾 Dados do lançamento');
  secao_(sh, 'B12:E12', '💳 Compra parcelada');
  secao_(sh, 'B16:E16', '📈 Investimento / Meta');

  const rotulos = {
    B5: 'Data *', B6: 'Descrição *', B7: 'Valor (R$) *', B8: 'Conta *', B9: 'Forma de pagamento',
    B10: 'Despesa fixa?', D5: 'Tipo *', D6: 'Categoria *', D7: 'Subcategoria', D8: 'Status',
    D9: 'Cartão (se crédito)', D10: 'Observações',
    B13: 'Compra parcelada?', B14: 'Valor TOTAL da compra', D13: 'Nº de parcelas', D14: 'Parcela inicial',
    B17: 'Ativo ou meta', D17: 'Quantidade (cotas, BTC…)'
  };
  Object.keys(rotulos).forEach(a1 => {
    sh.getRange(a1).setValue(rotulos[a1]).setFontWeight('bold').setFontColor(COR.AMEIXA).setBackground(COR.BLUSH)
      .setHorizontalAlignment('right');
  });

  const entradas = ['C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'E5', 'E6', 'E7', 'E8', 'E9', 'E10',
    'C13', 'C14', 'E13', 'E14', 'C17', 'E17'];
  entradas.forEach(a1 => {
    sh.getRange(a1).setBackground(COR.BRANCO).setFontSize(11).setHorizontalAlignment('left')
      .setBorder(true, true, true, true, null, null, COR.ROSE, SpreadsheetApp.BorderStyle.SOLID);
  });
  [5, 6, 7, 8, 9, 10, 13, 14, 17].forEach(r => sh.setRowHeight(r, 30));
  sh.setRowHeight(11, 10);
  sh.setRowHeight(15, 10);

  // ---------- Validações ----------
  const v = (a1, regra) => sh.getRange(a1).setDataValidation(regra);
  v(FORM.data, SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  v(FORM.tipo, validacaoLista_(TIPOS));
  v(FORM.cat, validacaoIntervalo_(cfgCol(CFG.AUX_CATS), false));
  v(FORM.sub, validacaoIntervalo_(cfgCol(CFG.AUX_SUBS), true));
  v(FORM.conta, validacaoIntervalo_(cfgCol(CFG.CONTA), false));
  v(FORM.forma, validacaoIntervalo_(cfgCol(CFG.FORMA), false));
  v(FORM.status, validacaoLista_(STATUS));
  v(FORM.cartao, validacaoIntervalo_(cfgCol(CFG.CART_NOME), false));
  v(FORM.fixa, validacaoLista_(SIM_NAO));
  v(FORM.parcelado, validacaoLista_(SIM_NAO));
  v(FORM.valor, SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(false).build());
  v(FORM.valorTotal, SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(false).build());
  v(FORM.nParc, SpreadsheetApp.newDataValidation().requireNumberBetween(2, 120).setAllowInvalid(false).build());
  v(FORM.parcIni, SpreadsheetApp.newDataValidation().requireNumberBetween(1, 120).setAllowInvalid(false).build());
  v(FORM.ativo, validacaoIntervalo_(cfgCol(CFG.AUX_ATIVOS), true));
  v(FORM.qtd, SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());

  sh.getRange(FORM.data).setNumberFormat(FORMATO_DATA);
  sh.getRange(FORM.valor).setNumberFormat(FORMATO_MOEDA);
  sh.getRange(FORM.valorTotal).setNumberFormat(FORMATO_MOEDA);
  sh.getRange(FORM.qtd).setNumberFormat('0.########');
  sh.getRange(FORM.nParc + ':' + FORM.parcIni).setNumberFormat('0');

  // ---------- Prévia ----------
  secao_(sh, 'B19:E19', '👀 Prévia');
  sh.getRange('B20:E20').merge().setFormula(fx_(
    '=LET(d,$C$5,cart,$E$9,dia,CFG_DIA_INICIO,' +
    'f,IFERROR(INDEX(CART_FECH,MATCH(cart,CART_NOME,0)),0),' +
    'v,IFERROR(INDEX(CART_VENC,MATCH(cart,CART_NOME,0)),0),' +
    'IF(d="","📅 Informe a data para ver a prévia",' +
    'IF($C$9="Crédito",' +
    'IF(f=0,"💳 Escolha o cartão para ver em qual fatura a compra vai cair",' +
    'LET(mf,MONTH(d)+IF(DAY(d)>=MIN(f,DAY(EOMONTH(d,0))),1,0),mv,mf+IF(v>f,0,1),' +
    'venc,DATE(YEAR(d),mv,MIN(v,DAY(EOMONTH(DATE(YEAR(d),mv,1),0)))),' +
    'ref,IF(AND(dia>1,DAY(venc)>=dia),DATE(YEAR(venc),MONTH(venc)+1,1),DATE(YEAR(venc),MONTH(venc),1)),' +
    '"💳 Cai na fatura do "&cart&" que vence em "&TEXT(venc,"dd/mm/yyyy")&"  •  entra no Resumo de "&TEXT(ref,"mm/yyyy"))),' +
    'LET(ref,IF(AND(dia>1,DAY(d)>=dia),DATE(YEAR(d),MONTH(d)+1,1),DATE(YEAR(d),MONTH(d),1)),' +
    '"📅 Entra no Resumo de "&TEXT(ref,"mm/yyyy")&IF(dia>1,"  (ciclo de "&dia&" a "&(dia-1)&")","")))))'));
  sh.getRange('B21:E21').merge().setFormula(fx_(
    '=IF($C$13<>"Sim","",LET(n,N($E$13),ini,IF($E$14="",1,$E$14),tot,IF(N($C$14)>0,$C$14,N($C$7)),' +
    'IF(OR(n<2,tot<=0),"🧾 Informe o nº de parcelas e o valor total da compra",' +
    '"🧾 "&n&"x de R$ "&FIXED(tot/n,2)&"  •  serão lançadas as parcelas "&ini&" a "&n&' +
    '" ("&(n-ini+1)&" lançamentos, R$ "&FIXED((n-ini+1)*tot/n,2)&")")))'));
  sh.getRange('B20:E21').setBackground(COR.LAVANDA_CLARA).setFontColor(COR.AMEIXA).setFontStyle('italic')
    .setFontSize(10).setWrap(true);
  sh.setRowHeight(20, 30);
  sh.setRowHeight(21, 30);

  // ---------- "Botões" (caixas de seleção) ----------
  sh.getRange('B23:D23').merge().setValue('✅  REGISTRAR LANÇAMENTO   ➜   marque a caixinha')
    .setBackground(COR.ROSA).setFontColor(COR.BRANCO).setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('right');
  sh.getRange(FORM.btnRegistrar).insertCheckboxes().setValue(false)
    .setBackground(COR.ROSA).setFontColor(COR.BRANCO).setFontSize(20).setHorizontalAlignment('left');
  sh.setRowHeight(23, 44);
  sh.getRange('B24:E24').merge().setFontWeight('bold').setFontSize(10).setWrap(true).setHorizontalAlignment('center');
  sh.setRowHeight(24, 34);
  sh.getRange('B25:D25').merge().setValue('🧹  Limpar formulário   ➜')
    .setBackground(COR.LAVANDA_CLARA).setFontColor(COR.AMEIXA).setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(FORM.btnLimpar).insertCheckboxes().setValue(false)
    .setBackground(COR.LAVANDA_CLARA).setFontColor(COR.AMEIXA_CLARA).setFontSize(14).setHorizontalAlignment('left');
  sh.setRowHeight(25, 30);

  // ---------- Últimos lançamentos ----------
  secao_(sh, 'B27:F27', '🕑 Últimos lançamentos');
  cabecalho_(sh, [['B28', 'Data'], ['C28', 'Descrição'], ['D28', 'Tipo'], ['E28', 'Categoria'], ['F28', 'Valor']]);
  zebra_(sh, 29, 2, 10, 5);
  sh.getRange('B29').setFormula(fx_(
    '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({BD_DATA,' +
    'BD_DESC&IF(BD_PARCELA<>""," ("&BD_PARCELA&"/"&BD_TOTPARC&")",""),BD_TIPO,BD_CAT,BD_VALOR,ROW(BD_ID)},' +
    'BD_ID<>""),6,FALSE),10,5),"Nenhum lançamento ainda"))'));
  sh.getRange('B29:B38').setNumberFormat(FORMATO_DATA).setHorizontalAlignment('center');
  sh.getRange('F29:F38').setNumberFormat(FORMATO_MOEDA);

  // ---------- Painel de ajuda ----------
  sh.getRange('H4:H25').merge().setValue(
    '💡 COMO USAR\n\n' +
    '1. Preencha os campos. Os destacados em rosa são obrigatórios.\n\n' +
    '2. O Tipo filtra as Categorias, e a Categoria filtra as Subcategorias.\n\n' +
    '3. Crédito: escolha o cartão — a prévia mostra em qual fatura a compra vai cair.\n\n' +
    '4. Parcelado: escolha "Sim", informe o valor TOTAL, o nº de parcelas e a parcela inicial ' +
    '(use mais que 1 se a compra já estava em andamento, ex.: começa na 4/12).\n\n' +
    '5. Investimento: escolha o ativo ou a meta e, se quiser, a quantidade.\n\n' +
    '6. Marque a caixinha ✅. Banco de Dados, Resumo, faturas, parcelas e saldo se atualizam sozinhos.\n\n' +
    '✨ Contas fixas: marque ✔ no checklist do Resumo. Se registrar aqui com "Despesa fixa? = Sim" e a ' +
    'descrição igual ao nome da conta fixa, o checklist também é marcado.\n\n' +
    '🔁 Transferências (entre contas, pagamento de fatura) não contam como gasto.')
    .setBackground(COR.LAVANDA_CLARA).setFontColor(COR.AMEIXA).setFontSize(10).setWrap(true)
    .setVerticalAlignment('top')
    .setBorder(true, true, true, true, null, null, COR.LAVANDA, SpreadsheetApp.BorderStyle.SOLID);

  // ---------- Formatação condicional ----------
  const r = a1 => sh.getRange(a1);
  const regras = [
    // obrigatórios em branco
    regraCor_([r('C5')], '=$C$5=""', COR.ROSE_CLARO),
    regraCor_([r('C6')], '=$C$6=""', COR.ROSE_CLARO),
    regraCor_([r('C7')], '=AND($C$7="",$C$13<>"Sim")', COR.ROSE_CLARO),
    regraCor_([r('C8')], '=AND($C$8="",$C$9<>"Crédito")', COR.ROSE_CLARO),
    regraCor_([r('C9')], '=AND($C$9="",$E$5="Despesa")', COR.ROSE_CLARO),
    regraCor_([r('E5')], '=$E$5=""', COR.ROSE_CLARO),
    regraCor_([r('E6')], '=$E$6=""', COR.ROSE_CLARO),
    regraCor_([r('E9')], '=AND($E$9="",$C$9="Crédito")', COR.ROSE_CLARO),
    regraCor_([r('C14')], '=AND($C$13="Sim",$C$14="",$C$7="")', COR.ROSE_CLARO),
    regraCor_([r('E13')], '=AND($C$13="Sim",$E$13="")', COR.ROSE_CLARO),
    // campos que não se aplicam ficam acinzentados
    regraCor_([r('E9')], '=$C$9<>"Crédito"', COR.CINZA, COR.TEXTO_SUAVE),
    regraCor_([r('C14'), r('E13'), r('E14')], '=$C$13<>"Sim"', COR.CINZA, COR.TEXTO_SUAVE),
    regraCor_([r('C17'), r('E17')], '=$E$5<>"Investimento"', COR.CINZA, COR.TEXTO_SUAVE),
    // tipo colorido na lista de últimos lançamentos
    regraCor_([r('D29:D38')], '=D29="Receita"', null, COR.MENTA_ESCURA),
    regraCor_([r('D29:D38')], '=D29="Despesa"', null, COR.CORAL_ESCURO),
    regraCor_([r('D29:D38')], '=D29="Investimento"', null, COR.AMEIXA_CLARA)
  ];
  sh.setConditionalFormatRules(regras);

  limparFormulario_(sh);
  sh.setActiveSelection(FORM.data);
}

/* ============================================================
 *  ABA RESUMO (dashboard)
 * ============================================================ */

function montarResumo_(sh, shCfg) {
  //          A   B    C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S    T
  estiloAba_(sh, [20, 160, 80, 80, 80, 80, 80, 80, 24, 80, 80, 80, 80, 80, 80, 80, 80, 20, 110, 110], COR.ROSA);

  // ---------- Cabeçalho + seletores ----------
  faixaTitulo_(sh, 17, 10, '💗 Meu Painel Financeiro',
    '="Período de "&TEXT(RES_INICIO,"dd/mm/yyyy")&" a "&TEXT(RES_FIM,"dd/mm/yyyy")&"   •   ciclo configurável em Configurações"');
  sh.getRange('L1:M1').merge().setValue('📅 Mês');
  sh.getRange('L2:M2').merge().setValue('🗓️ Ano');
  sh.getRange('L1:M2').setFontColor(COR.BRANCO).setFontWeight('bold').setHorizontalAlignment('right');
  const diaIni = Number(shCfg.getRange(CFG_DIA_INICIO_A1).getValue()) || 1;
  const anoIni = Number(shCfg.getRange(CFG_ANO_INI_A1).getValue()) || new Date().getFullYear();
  const atual = cicloMes(new Date(), diaIni);
  const anos = [];
  for (let a = anoIni; a <= Math.max(anoIni, atual.getFullYear()) + 6; a++) anos.push(String(a));
  sh.getRange('N1:O1').merge().setValue(MESES[atual.getMonth()]).setDataValidation(validacaoLista_(MESES));
  sh.getRange('N2:O2').merge().setValue(atual.getFullYear()).setDataValidation(validacaoLista_(anos));
  sh.getRange('N1:O2').setBackground(COR.BRANCO).setFontColor(COR.AMEIXA).setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, COR.AMEIXA, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange('P1:Q2').merge().setValue('Troque o mês e o ano para navegar ✨')
    .setFontColor(COR.ROSE).setFontSize(8).setWrap(true).setHorizontalAlignment('center');

  // ---------- Auxiliares (colunas S:T, ocultas) ----------
  const aux = {
    S1: 'Mês ref', T1: '=IFERROR(DATE(N2,MATCH(N1,{"' + MESES.join('","') + '"},0),1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
    S2: 'Início do ciclo', T2: '=IF(CFG_DIA_INICIO>1,DATE(YEAR(T1),MONTH(T1)-1,CFG_DIA_INICIO),T1)',
    S3: 'Fim do ciclo', T3: '=IF(CFG_DIA_INICIO>1,DATE(YEAR(T1),MONTH(T1),CFG_DIA_INICIO-1),EOMONTH(T1,0))',
    S4: 'Destino da renda', T4: 'Valor',
    S5: 'Investido', T5: '=KPI_INVESTIDO',
    S6: 'Gastos', T6: '=KPI_SAIDAS',
    S7: 'Sobra livre', T7: '=MAX(0,KPI_SALDO)',
    S9: 'Fixas pendentes', T9: '=SUMIFS(M14:M25,P14:P25,"<>✅ Paga",J14:J25,"<>")',
    S10: 'Compras parceladas ativas',
    T10: '=IFERROR(COUNTA(UNIQUE(FILTER(BD_GRUPO,BD_PARCELADO="Sim",BD_MESREF>=RES_REF))),0)'
  };
  Object.keys(aux).forEach(a1 => escrever_(sh, a1, aux[a1]));
  sh.getRange('T1:T3').setNumberFormat(FORMATO_DATA);
  sh.getRange('T5:T9').setNumberFormat(FORMATO_MOEDA);

  // ---------- KPIs ----------
  card_(sh, 4, 2, 4, '💰 ENTRADAS', '=SUMIFS(BD_VALOR,BD_TIPO,"Receita",BD_MESREF,RES_REF)',
    '="Receitas do período"', COR.MENTA_CLARA, FORMATO_MOEDA);
  card_(sh, 4, 5, 8, '💸 SAÍDAS', '=SUMIFS(BD_VALOR,BD_TIPO,"Despesa",BD_MESREF,RES_REF)',
    '="Despesas, incluindo as faturas do mês"', COR.BLUSH, FORMATO_MOEDA);
  card_(sh, 4, 10, 13, '⚖️ SALDO DO MÊS', '=KPI_ENTRADAS-KPI_SAIDAS-KPI_INVESTIDO',
    '=IF(KPI_SALDO>=0,"💚 Entradas − saídas − investido","⚠️ Saiu mais do que entrou")', COR.LAVANDA_CLARA, FORMATO_MOEDA);
  card_(sh, 4, 14, 17, '📈 VALOR INVESTIDO', '=SUMIFS(BD_VALOR,BD_TIPO,"Investimento",BD_MESREF,RES_REF)',
    '="Taxa de investimento: "&FIXED(IFERROR(KPI_INVESTIDO/KPI_ENTRADAS,0)*100,1)&"% da renda"', COR.PESSEGO, FORMATO_MOEDA);
  card_(sh, 8, 2, 4, '💳 FATURA DO CARTÃO',
    '=SUMIFS(BD_VALOR,BD_FORMA,"Crédito",BD_TIPO,"Despesa",BD_MESREF,RES_REF)',
    '="Faturas que vencem neste mês"', COR.BLUSH, FORMATO_MOEDA);
  card_(sh, 8, 5, 8, '🧾 PARCELAS FUTURAS', '=SUMIFS(BD_VALOR,BD_PARCELADO,"Sim",BD_MESREF,">"&RES_REF)',
    '=T10&" compra(s) parcelada(s) ativa(s)"', COR.LAVANDA_CLARA, FORMATO_MOEDA);
  card_(sh, 8, 10, 13, '✅ CONTAS FIXAS', '=COUNTIF(P14:P25,"✅*")&" de "&COUNTA(J14:J25)&" pagas"',
    '="R$ "&FIXED(FIXAS_PEND,2)&" ainda pendentes"', COR.MENTA_CLARA, null);
  card_(sh, 8, 14, 17, '🔮 SALDO PREVISTO', '=KPI_SALDO-FIXAS_PEND',
    '="Saldo depois de pagar as fixas pendentes"', COR.PESSEGO, FORMATO_MOEDA);
  sh.setRowHeight(7, 10);
  sh.setRowHeight(11, 14);

  // ---------- Cartões e contas fixas ----------
  secao_(sh, 'B12:H12', '💳 Cartões de crédito');
  secao_(sh, 'J12:Q12', '✅ Contas fixas do mês');
  cabecalho_(sh, [['B13', 'Cartão'], ['C13', 'Limite'], ['D13', 'Fatura do mês'], ['E13', 'Status'],
    ['F13', 'Uso do limite'], ['G13', 'Disponível'], ['H13', 'Paga? ✔']]);
  cabecalho_(sh, [['J13:K13', 'Conta'], ['L13', 'Categoria'], ['M13', 'Valor'], ['N13', 'Vencimento'],
    ['O13', 'Paga? ✔'], ['P13:Q13', 'Status']]);
  sh.setRowHeight(13, 30);
  zebra_(sh, 14, 2, RES.CART_MAX, 7);
  zebra_(sh, 14, 10, RES.FIXA_MAX, 8);
  sh.getRange(14, 10, RES.FIXA_MAX, 2).mergeAcross();
  sh.getRange(14, 16, RES.FIXA_MAX, 2).mergeAcross();
  sh.getRange('C14:D19').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('G14:G19').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('E14:E19').setHorizontalAlignment('center').setFontSize(9);
  sh.getRange('B14:B19').setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('H14:H19').setHorizontalAlignment('center').setFontColor(COR.ROSA);
  sh.getRange('J14:J25').setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('L14:L25').setFontSize(9);
  sh.getRange('M14:M25').setNumberFormat(FORMATO_MOEDA).setBackground(COR.BLUSH);
  sh.getRange('N14:N25').setNumberFormat('dd/mm').setHorizontalAlignment('center');
  sh.getRange('O14:O25').setHorizontalAlignment('center').setFontColor(COR.ROSA);
  sh.getRange('P14:P25').setHorizontalAlignment('center').setFontSize(9);

  escrever_(sh, 'B21', 'Total das faturas do mês').setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('B21:C21').merge();
  escrever_(sh, 'D21', '=SUM(D14:D19)').setNumberFormat(FORMATO_MOEDA).setFontWeight('bold');
  sh.getRange('E21:F21').merge().setValue('Disponível total').setFontWeight('bold').setFontColor(COR.AMEIXA)
    .setHorizontalAlignment('right');
  escrever_(sh, 'G21', '=SUM(G14:G19)').setNumberFormat(FORMATO_MOEDA).setFontWeight('bold');
  sh.getRange('B21:H21').setBackground(COR.ROSE_CLARO);
  sh.getRange('B23:H26').merge().setValue(
    '💡 Marque ✔ em "Paga?" quando pagar a fatura: o pagamento vira uma transferência no Banco de Dados ' +
    '(sem contar o gasto duas vezes) e o limite é liberado. Nas contas fixas, ajuste o valor do mês ' +
    '(ex.: energia) e marque ✔ — o lançamento é criado automaticamente.')
    .setFontSize(9).setFontStyle('italic').setFontColor(COR.AMEIXA_CLARA).setWrap(true)
    .setBackground(COR.LAVANDA_CLARA).setVerticalAlignment('middle');

  sh.getRange('J26:M26').merge().setFormula(fx_('="✅ Contas fixas: "&COUNTIF(P14:P25,"✅*")&" de "&COUNTA(J14:J25)&" pagas"'));
  sh.getRange('N26:Q26').merge().setFormula(fx_('="⏳ R$ "&FIXED(FIXAS_PEND,2)&" ainda pendentes"'));
  sh.getRange('J26:Q26').setBackground(COR.ROSE_CLARO).setFontWeight('bold').setFontColor(COR.AMEIXA)
    .setHorizontalAlignment('center');

  // ---------- Gastos por categoria ----------
  secao_(sh, 'B28:H28', '🏷️ Gastos por categoria');
  secao_(sh, 'J28:Q28', '🥧 Distribuição dos gastos');
  cabecalho_(sh, [['B29', 'Categoria'], ['C29', 'Valor gasto'], ['D29', '% dos gastos'], ['E29', '% da renda'],
    ['F29', 'Orçamento'], ['G29:H29', 'Uso do orçamento']]);
  sh.setRowHeight(29, 30);
  zebra_(sh, 30, 2, 15, 7);
  sh.getRange(30, 7, 15, 2).mergeAcross();
  // Crédito à vista entra pela data da compra (período do ciclo); parcelas e demais formas, pelo mês de referência
  sh.getRange('B30').setFormula(fx_(
    '=IFERROR(LET(cats,FILTER(LST_CATEGORIAS,CAT_TIPO="Despesa"),' +
    'ok,ARRAYFORMULA(IF((BD_FORMA="Crédito")*(BD_PARCELADO<>"Sim"),(BD_DATA>=RES_INICIO)*(BD_DATA<=RES_FIM),' +
    'BD_MESREF=RES_REF)*(BD_TIPO="Despesa")),' +
    'v,ARRAYFORMULA(SUMIF(IF(ok,BD_CAT,""),cats,BD_VALOR)),' +
    'ARRAY_CONSTRAIN(SORT(FILTER({cats,v},v>0),2,FALSE),15,2)),{"Sem gastos neste mês",0})'));
  const fCat = [];
  for (let r = 30; r <= 44; r++) {
    fCat.push([
      '=IF($B' + r + '="","",IFERROR($C' + r + '/SUM($C$30:$C$44),0))',
      '=IF($B' + r + '="","",IFERROR($C' + r + '/KPI_ENTRADAS,0))',
      '=IF($B' + r + '="","",LET(o,IFERROR(INDEX(CAT_ORC,MATCH($B' + r + ',LST_CATEGORIAS,0)),""),IF(o="","",o)))',
      '=IF(OR($B' + r + '="",N($F' + r + ')=0),"",SPARKLINE($C' + r + ',{"charttype","bar";"max",$F' + r +
      ';"color1",IF($C' + r + '>$F' + r + ',"' + COR.CORAL + '","' + COR.ROSA + '")}))'
    ]);
  }
  sh.getRange('D30:G44').setFormulas(fxs_(fCat.map(f => [f[0], f[1], f[2], f[3]])));
  sh.getRange('C30:C44').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('D30:E44').setNumberFormat(FORMATO_PCT).setHorizontalAlignment('center');
  sh.getRange('F30:F44').setNumberFormat(FORMATO_MOEDA).setFontColor(COR.TEXTO_SUAVE);
  escrever_(sh, 'B45', 'Total');
  escrever_(sh, 'C45', '=SUM(C30:C44)');
  escrever_(sh, 'D45', '=IF(C45>0,1,0)');
  escrever_(sh, 'E45', '=IFERROR(C45/KPI_ENTRADAS,0)');
  escrever_(sh, 'F45', '=SUM(F30:F44)');
  sh.getRange('B45:H45').setBackground(COR.ROSE_CLARO).setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('C45').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('D45:E45').setNumberFormat(FORMATO_PCT).setHorizontalAlignment('center');
  sh.getRange('F45').setNumberFormat(FORMATO_MOEDA);

  // ---------- Parcelas do mês ----------
  secao_(sh, 'B47:H47', '🧾 Parcelas do mês e compras parceladas ativas');
  cabecalho_(sh, [['B48', 'Descrição'], ['C48', 'Cartão / forma'], ['D48', 'Parcela'], ['E48', 'Valor'],
    ['F48', 'Restantes'], ['G48', 'Falta pagar'], ['H48', '']]);
  sh.setRowHeight(48, 30);
  zebra_(sh, 49, 2, 10, 7);
  sh.getRange('B49').setFormula(fx_(
    '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({BD_DESC,IF(BD_CARTAO="",BD_FORMA,BD_CARTAO),' +
    'BD_PARCELA&"/"&BD_TOTPARC,BD_VALOR,BD_TOTPARC-BD_PARCELA,(BD_TOTPARC-BD_PARCELA)*BD_VALOR},' +
    'BD_PARCELADO="Sim",BD_MESREF=RES_REF),1,TRUE),10,6),"Nenhuma parcela neste mês"))'));
  sh.getRange('D49:D58').setHorizontalAlignment('center').setFontWeight('bold').setFontColor(COR.ROSA);
  sh.getRange('E49:E58').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('F49:F58').setHorizontalAlignment('center');
  sh.getRange('G49:G58').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('B59:D59').merge().setValue('Total comprometido no mês');
  escrever_(sh, 'E59', '=SUM(E49:E58)').setNumberFormat(FORMATO_MOEDA);
  escrever_(sh, 'F59', '=SUM(F49:F58)').setHorizontalAlignment('center');
  escrever_(sh, 'G59', '=SUM(G49:G58)').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('B59:H59').setBackground(COR.ROSE_CLARO).setFontWeight('bold').setFontColor(COR.AMEIXA);

  // ---------- Despesas futuras ----------
  secao_(sh, 'J47:Q47', '📆 Despesas futuras (próximos 6 meses)');
  cabecalho_(sh, [['J48:K48', 'Mês'], ['L48', 'Parcelas'], ['M48', 'Fixas previstas'], ['N48', 'Total'],
    ['O48', '% da renda atual'], ['P48:Q48', 'Peso']]);
  zebra_(sh, 49, 10, 6, 8);
  sh.getRange(49, 10, 7, 2).mergeAcross();
  sh.getRange(49, 16, 6, 2).mergeAcross();
  const fFut = [];
  for (let i = 1; i <= 6; i++) {
    const r = 48 + i;
    fFut.push(['=EDATE(RES_REF,' + i + ')', '', '=SUMIFS(BD_VALOR,BD_PARCELADO,"Sim",BD_MESREF,$J' + r + ')',
      '=SUM(FIXA_VALOR)', '=L' + r + '+M' + r, '=IFERROR(N' + r + '/KPI_ENTRADAS,0)',
      '=IF(MAX($N$49:$N$54)=0,"",SPARKLINE(N' + r + ',{"charttype","bar";"max",MAX($N$49:$N$54);"color1","' + COR.LAVANDA + '"}))']);
  }
  sh.getRange('J49:J54').setFormulas(fxs_(fFut.map(f => [f[0]])));
  sh.getRange('L49:P54').setFormulas(fxs_(fFut.map(f => [f[2], f[3], f[4], f[5], f[6]])));
  sh.getRange('J49:J54').setNumberFormat('mmm/yyyy').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('L49:N54').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('O49:O54').setNumberFormat(FORMATO_PCT).setHorizontalAlignment('center');
  escrever_(sh, 'J55', 'Total');
  escrever_(sh, 'L55', '=SUM(L49:L54)');
  escrever_(sh, 'M55', '=SUM(M49:M54)');
  escrever_(sh, 'N55', '=SUM(N49:N54)');
  sh.getRange('J55:Q55').setBackground(COR.ROSE_CLARO).setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('L55:N55').setNumberFormat(FORMATO_MOEDA);

  const stats = [
    ['J57', '🛍️ Compras parceladas ativas', '=T10', '0'],
    ['J58', '🔢 Parcelas que ainda faltam', '=COUNTIFS(BD_PARCELADO,"Sim",BD_MESREF,">"&RES_REF)', '0'],
    ['J59', '💰 Valor que ainda falta pagar', '=KPI_PARCFUT', FORMATO_MOEDA]
  ];
  stats.forEach(s => {
    const lin = sh.getRange(s[0]).getRow();
    sh.getRange(lin, 10, 1, 5).merge().setValue(s[1]).setFontColor(COR.AMEIXA).setFontWeight('bold');
    sh.getRange(lin, 15, 1, 3).merge().setFormula(fx_(s[2])).setNumberFormat(s[3])
      .setHorizontalAlignment('right').setFontWeight('bold').setFontColor(COR.ROSA);
    sh.getRange(lin, 10, 1, 8).setBackground(COR.BLUSH);
  });

  // ---------- Evolução 12 meses ----------
  secao_(sh, 'B78:H78', '📅 Evolução dos últimos 12 meses');
  cabecalho_(sh, [['B79', 'Mês'], ['C79', 'Receitas'], ['D79', 'Despesas'], ['E79', 'Investimentos'],
    ['F79', 'Saldo'], ['G79', '% investido'], ['H79', '']]);
  zebra_(sh, 80, 2, 12, 7);
  const fEvo = [];
  for (let i = 0; i < 12; i++) {
    const off = i - 11;
    const mes = 'EDATE(RES_REF,' + off + ')';
    const r = 80 + i;
    fEvo.push([
      '=TEXT(' + mes + ',"mmm/yy")',
      '=SUMIFS(BD_VALOR,BD_TIPO,"Receita",BD_MESREF,' + mes + ')',
      '=SUMIFS(BD_VALOR,BD_TIPO,"Despesa",BD_MESREF,' + mes + ')',
      '=SUMIFS(BD_VALOR,BD_TIPO,"Investimento",BD_MESREF,' + mes + ')',
      '=C' + r + '-D' + r + '-E' + r,
      '=IFERROR(E' + r + '/C' + r + ',0)'
    ]);
  }
  sh.getRange('B80:G91').setFormulas(fxs_(fEvo));
  sh.getRange('B80:B91').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('C80:F91').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('G80:G91').setNumberFormat(FORMATO_PCT).setHorizontalAlignment('center');

  // ---------- Maiores gastos ----------
  secao_(sh, 'J78:Q78', '🔝 Maiores gastos do mês');
  cabecalho_(sh, [['J79', 'Data'], ['K79:N79', 'Descrição'], ['O79:P79', 'Categoria'], ['Q79', 'Valor']]);
  zebra_(sh, 80, 10, 5, 8);
  sh.getRange(80, 11, 5, 4).mergeAcross();
  sh.getRange(80, 15, 5, 2).mergeAcross();
  const top = 'SORT(FILTER({BD_DATA,BD_DESC,BD_CAT,BD_VALOR},BD_TIPO="Despesa",BD_MESREF=RES_REF),4,FALSE)';
  for (let i = 1; i <= 5; i++) {
    const r = 79 + i;
    sh.getRange('J' + r).setFormula(fx_('=IFERROR(INDEX(' + top + ',' + i + ',1),"")'));
    sh.getRange('K' + r).setFormula(fx_('=IFERROR(INDEX(' + top + ',' + i + ',2),"")'));
    sh.getRange('O' + r).setFormula(fx_('=IFERROR(INDEX(' + top + ',' + i + ',3),"")'));
    sh.getRange('Q' + r).setFormula(fx_('=IFERROR(INDEX(' + top + ',' + i + ',4),"")'));
  }
  sh.getRange('J80:J84').setNumberFormat('dd/mm').setHorizontalAlignment('center');
  sh.getRange('Q80:Q84').setNumberFormat(FORMATO_MOEDA).setFontWeight('bold').setFontColor(COR.CORAL_ESCURO);

  // ---------- Indicadores extras ----------
  secao_(sh, 'J86:Q86', '✨ Indicadores extras');
  const extras = [
    [87, '📊 Média de gastos (últimos 12 meses)', '=AVERAGE(D80:D91)', FORMATO_MOEDA],
    [88, '🏆 Categoria com maior gasto', '=IF(N(C30)>0,B30,"—")', '@'],
    [89, '🐷 Taxa de poupança (saldo ÷ renda)', '=IFERROR(KPI_SALDO/KPI_ENTRADAS,0)', FORMATO_PCT],
    [90, '📆 Gasto médio por dia', '=IFERROR(KPI_SAIDAS/(RES_FIM-RES_INICIO+1),0)', FORMATO_MOEDA],
    [91, '💳 Uso do limite total dos cartões', '=IFERROR(1-G21/SUM(C14:C19),0)', FORMATO_PCT]
  ];
  extras.forEach(x => {
    sh.getRange(x[0], 10, 1, 5).merge().setValue(x[1]).setFontColor(COR.AMEIXA);
    sh.getRange(x[0], 15, 1, 3).merge().setFormula(fx_(x[2])).setNumberFormat(x[3])
      .setHorizontalAlignment('right').setFontWeight('bold').setFontColor(COR.ROSA);
  });
  zebra_(sh, 87, 10, 5, 8);

  // ---------- Formatação condicional ----------
  const r = a1 => sh.getRange(a1);
  sh.setConditionalFormatRules([
    regraCor_([r('J5:M5')], '=$J$5<0', null, COR.CORAL_ESCURO),
    regraCor_([r('J5:M5')], '=$J$5>0', null, COR.MENTA_ESCURA),
    regraCor_([r('N9:Q9')], '=$N$9<0', null, COR.CORAL_ESCURO),
    regraCor_([r('F80:F91')], '=F80<0', null, COR.CORAL_ESCURO),
    regraCor_([r('F80:F91')], '=F80>0', null, COR.MENTA_ESCURA),
    regraCor_([r('E14:E19')], '=LEFT(E14,1)="✅"', COR.MENTA_CLARA, COR.MENTA_ESCURA),
    regraCor_([r('E14:E19')], '=E14="🔒 Fechada"', COR.PESSEGO, COR.CORAL_ESCURO),
    regraCor_([r('P14:Q25')], '=LEFT($P14,1)="✅"', COR.MENTA_CLARA, COR.MENTA_ESCURA),
    regraCor_([r('P14:Q25')], '=LEFT($P14,2)="⚠️"', COR.PESSEGO, COR.CORAL_ESCURO),
    regraCor_([r('B30:H44')], '=AND(N($F30)>0,$C30>$F30)', null, COR.CORAL_ESCURO)
  ]);

  // ---------- Gráficos ----------
  secao_(sh, 'B61:Q61', '📊 Visão gráfica');
  grafico_(sh, Charts.ChartType.PIE, ['B29:C44'], 29, 10, 640, 370, 'Gastos por categoria', {
    pieHole: 0.45, colors: PALETA_GRAFICOS, pieSliceText: 'percentage',
    legend: { position: 'right', textStyle: { color: COR.TEXTO, fontSize: 10 } },
    chartArea: { left: 20, top: 40, width: '90%', height: '80%' }
  });
  grafico_(sh, Charts.ChartType.COLUMN, ['J48:J54', 'L48:M54'], 62, 2, 640, 310, 'Despesas futuras já comprometidas', {
    isStacked: true, colors: [COR.ROSA, COR.LAVANDA],
    vAxis: { format: 'R$ #,##0', gridlines: { color: '#F3E6EC' }, textStyle: { color: COR.TEXTO_SUAVE } },
    hAxis: { format: 'MMM/yy', textStyle: { color: COR.TEXTO_SUAVE } }
  });
  grafico_(sh, Charts.ChartType.PIE, ['S4:T7'], 62, 10, 640, 310, 'Para onde foi a renda do mês', {
    pieHole: 0.55, colors: [COR.ROSA, COR.LAVANDA, COR.MENTA], pieSliceText: 'percentage',
    legend: { position: 'right', textStyle: { color: COR.TEXTO, fontSize: 11 } }
  });
  grafico_(sh, Charts.ChartType.LINE, ['B79:B91', 'D79:D91'], 93, 2, 640, 320, 'Evolução dos gastos', {
    colors: [COR.ROSA], curveType: 'function', pointSize: 7, lineWidth: 3,
    vAxis: { format: 'R$ #,##0', gridlines: { color: '#F3E6EC' }, textStyle: { color: COR.TEXTO_SUAVE } },
    hAxis: { textStyle: { color: COR.TEXTO_SUAVE } }
  });
  grafico_(sh, Charts.ChartType.COLUMN, ['B79:E91'], 93, 10, 640, 320, 'Receitas × Despesas × Investimentos', {
    colors: [COR.MENTA, COR.CORAL, COR.AMEIXA_CLARA],
    vAxis: { format: 'R$ #,##0', gridlines: { color: '#F3E6EC' }, textStyle: { color: COR.TEXTO_SUAVE } },
    hAxis: { textStyle: { color: COR.TEXTO_SUAVE } }
  });

  sh.hideColumns(19, 2);
}

/* ============================================================
 *  ABA INVESTIMENTOS
 * ============================================================ */

/**
 * Soma uma coluna do BD para um ativo/meta, ignorando espaços extras
 * (inclusive o espaço "invisível" de textos colados) e maiúsculas/minúsculas.
 */
function somaAtivo_(colunaBD, celulaNome) {
  return 'SUM(IFERROR(FILTER(' + colunaBD + ',BD_TIPO="Investimento",' +
    'TRIM(SUBSTITUTE(BD_ATIVO,CHAR(160)," "))=TRIM(SUBSTITUTE(' + celulaNome + ',CHAR(160)," "))),0))';
}

function montarInvestimentos_(sh) {
  estiloAba_(sh, [20].concat(new Array(16).fill(85)).concat([20]), COR.MENTA);
  faixaTitulo_(sh, 17, 10, '📈 Investimentos & Metas',
    '="Cotações atualizadas automaticamente  •  mês do Resumo: "&TEXT(RES_REF,"mm/yyyy")');

  // ---------- KPIs ----------
  card_(sh, 4, 2, 5, '💼 TOTAL INVESTIDO', '=E31+SUM(E10:E17)', '="Carteira + metas"', COR.LAVANDA_CLARA, FORMATO_MOEDA);
  card_(sh, 4, 6, 9, '🏦 PATRIMÔNIO ATUAL', '=H31+SUM(E10:E17)', '="Carteira a preço de mercado + metas"',
    COR.MENTA_CLARA, FORMATO_MOEDA);
  card_(sh, 4, 10, 13, '📊 RENTABILIDADE GERAL', '=I31',
    '=FIXED(IFERROR(I31/E31,0)*100,2)&"% sobre o valor aportado na carteira"', COR.BLUSH, FORMATO_MOEDA);
  card_(sh, 4, 14, 17, '🗓️ APORTADO NO MÊS', '=SUMIFS(BD_VALOR,BD_TIPO,"Investimento",BD_MESREF,RES_REF)',
    '="Mês selecionado no Resumo: "&TEXT(RES_REF,"mm/yyyy")', COR.PESSEGO, FORMATO_MOEDA);
  sh.setRowHeight(7, 14);

  // ---------- Metas ----------
  secao_(sh, 'B8:Q8', '🎯 Metas financeiras');
  cabecalho_(sh, [['B9:C9', 'Meta'], ['D9', 'Valor da meta'], ['E9', 'Acumulado'], ['F9', 'Falta'],
    ['G9', '% concluído'], ['H9:J9', 'Progresso'], ['K9', 'Prazo'], ['L9', 'Meses restantes'],
    ['M9', 'Guardar por mês'], ['N9:Q9', 'Situação']]);
  sh.setRowHeight(9, 30);
  zebra_(sh, 10, 2, 8, 16);
  sh.getRange(10, 2, 8, 2).mergeAcross();
  sh.getRange(10, 8, 8, 3).mergeAcross();
  sh.getRange(10, 14, 8, 4).mergeAcross();
  for (let i = 0; i < 8; i++) {
    const r = 10 + i;
    const m = 'MATCH($B' + r + ',META_NOME,0)';
    sh.getRange('B' + r).setFormula(fx_('=IFERROR(INDEX(FILTER(META_NOME,META_NOME<>""),' + (i + 1) + '),"")'));
    sh.getRange('D' + r + ':G' + r).setFormulas(fxs_([[
      '=IF($B' + r + '="","",INDEX(META_VALOR,' + m + '))',
      '=ARRAYFORMULA(IF($B' + r + '="","",N(INDEX(META_INICIAL,' + m + '))+' + somaAtivo_('BD_VALOR', '$B' + r) + '))',
      '=IF($B' + r + '="","",MAX(0,D' + r + '-E' + r + '))',
      '=IF($B' + r + '="","",IFERROR(MIN(1,E' + r + '/D' + r + '),0))'
    ]]));
    sh.getRange('H' + r).setFormula(fx_('=IF($B' + r + '="","",SPARKLINE(G' + r + ',{"charttype","bar";"max",1;"color1",IF(G' + r +
      '>=1,"' + COR.MENTA + '","' + COR.ROSA + '");"color2","' + COR.ROSE + '"}))'));
    sh.getRange('K' + r + ':N' + r).setFormulas(fxs_([[
      '=IF($B' + r + '="","",LET(p,INDEX(META_PRAZO,' + m + '),IF(p="","",p)))',
      '=IF(OR($B' + r + '="",K' + r + '=""),"",MAX(1,(YEAR(K' + r + ')-YEAR(TODAY()))*12+MONTH(K' + r + ')-MONTH(TODAY())))',
      '=IF(OR($B' + r + '="",L' + r + '=""),"",IF(F' + r + '=0,0,F' + r + '/L' + r + '))',
      '=IF($B' + r + '="","",IF(G' + r + '>=1,"🎉 Meta concluída!",IF(L' + r + '="","💪 Defina um prazo","💪 Guarde "&"R$ "&FIXED(M' + r + ',2)&" por mês")))'
    ]]));
  }
  sh.getRange('B10:B17').setFontWeight('bold').setFontColor(COR.AMEIXA).setFontSize(11);
  sh.getRange('D10:F17').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('G10:G17').setNumberFormat(FORMATO_PCT).setHorizontalAlignment('center').setFontWeight('bold')
    .setFontColor(COR.ROSA);
  sh.getRange('K10:K17').setNumberFormat('mmm/yyyy').setHorizontalAlignment('center');
  sh.getRange('L10:L17').setHorizontalAlignment('center');
  sh.getRange('M10:M17').setNumberFormat(FORMATO_MOEDA).setFontWeight('bold');
  sh.getRange('N10:N17').setFontSize(9).setFontColor(COR.AMEIXA_CLARA);
  [10, 11, 12, 13, 14, 15, 16, 17].forEach(r => sh.setRowHeight(r, 28));

  // ---------- Carteira ----------
  secao_(sh, 'B19:K19', '💼 Carteira de investimentos');
  secao_(sh, 'L19:Q19', '🥧 Distribuição da carteira');
  cabecalho_(sh, [['B20', 'Tipo'], ['C20', 'Ativo'], ['D20', 'Quantidade'], ['E20', 'Valor aportado'],
    ['F20', 'Preço médio'], ['G20', 'Preço atual'], ['H20', 'Total acumulado'], ['I20', 'Rentab. R$'],
    ['J20', 'Rentab. %'], ['K20', '% da carteira']]);
  sh.setRowHeight(20, 30);
  zebra_(sh, 21, 2, 10, 10);
  const fCart = [];
  for (let i = 0; i < 10; i++) {
    const r = 21 + i;
    const m = 'MATCH($C' + r + ',ATV_NOME,0)';
    fCart.push([
      '=IF($C' + r + '="","",INDEX(ATV_TIPO,' + m + '))',
      '=IFERROR(INDEX(FILTER(ATV_NOME,ATV_NOME<>""),' + (i + 1) + '),"")',
      '=ARRAYFORMULA(IF($C' + r + '="","",' + somaAtivo_('BD_QTD', '$C' + r) + '))',
      '=ARRAYFORMULA(IF($C' + r + '="","",' + somaAtivo_('BD_VALOR', '$C' + r) + '))',
      '=IF(OR($C' + r + '="",N(D' + r + ')=0),"",E' + r + '/D' + r + ')',
      '=IF($C' + r + '="","",INDEX(ATV_PRECO,' + m + '))',
      '=IF($C' + r + '="","",IF(OR(N(G' + r + ')=0,N(D' + r + ')=0),E' + r + ',D' + r + '*G' + r + '))',
      '=IF($C' + r + '="","",H' + r + '-E' + r + ')',
      '=IF($C' + r + '="","",IFERROR(I' + r + '/E' + r + ',0))',
      '=IF($C' + r + '="","",IFERROR(H' + r + '/$H$31,0))'
    ]);
  }
  sh.getRange('B21:K30').setFormulas(fxs_(fCart));
  sh.getRange('B21:B30').setHorizontalAlignment('center').setFontSize(9);
  sh.getRange('C21:C30').setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('D21:D30').setNumberFormat('0.########');
  sh.getRange('E21:I30').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('J21:K30').setNumberFormat(FORMATO_PCT).setHorizontalAlignment('center');
  escrever_(sh, 'B31', 'Total');
  escrever_(sh, 'E31', '=SUM(E21:E30)');
  escrever_(sh, 'H31', '=SUM(H21:H30)');
  escrever_(sh, 'I31', '=SUM(I21:I30)');
  escrever_(sh, 'J31', '=IFERROR(I31/E31,0)');
  escrever_(sh, 'K31', '=IF(H31>0,1,0)');
  sh.getRange('B31:K31').setBackground(COR.ROSE_CLARO).setFontWeight('bold').setFontColor(COR.AMEIXA);
  sh.getRange('E31:I31').setNumberFormat(FORMATO_MOEDA);
  sh.getRange('J31:K31').setNumberFormat(FORMATO_PCT).setHorizontalAlignment('center');

  // ---------- Aportes e histórico ----------
  secao_(sh, 'B33:F33', '🧾 Aportes realizados');
  cabecalho_(sh, [['B34', 'Data'], ['C34', 'Ativo / meta'], ['D34', 'Quantidade'], ['E34', 'Valor'], ['F34', 'Preço unit.']]);
  sh.getRange('B35').setFormula(fx_(
    '=IFERROR(SORT(FILTER({BD_DATA,BD_ATIVO,BD_QTD,BD_VALOR,BD_PRECO},BD_TIPO="Investimento"),1,FALSE),' +
    '"Nenhum aporte registrado ainda")'));
  sh.getRange('B35:B400').setNumberFormat(FORMATO_DATA).setHorizontalAlignment('center');
  sh.getRange('D35:D400').setNumberFormat('0.########');
  sh.getRange('E35:F400').setNumberFormat(FORMATO_MOEDA);

  secao_(sh, 'H33:J33', '📸 Histórico do patrimônio');
  cabecalho_(sh, [['H34', 'Mês'], ['I34', 'Total investido'], ['J34', 'Patrimônio']]);
  sh.getRange('H35:H400').setNumberFormat('mmm/yyyy').setHorizontalAlignment('center');
  sh.getRange('I35:J400').setNumberFormat(FORMATO_MOEDA);
  secao_(sh, 'L33:Q33', '📈 Evolução do patrimônio');

  // ---------- Formatação condicional ----------
  const r = a1 => sh.getRange(a1);
  sh.setConditionalFormatRules([
    regraCor_([r('I21:J31')], '=N(I21)<0', null, COR.CORAL_ESCURO),
    regraCor_([r('I21:J31')], '=N(I21)>0', null, COR.MENTA_ESCURA),
    regraCor_([r('J5:M5')], '=$J$5<0', null, COR.CORAL_ESCURO)
  ]);

  // ---------- Gráficos ----------
  grafico_(sh, Charts.ChartType.PIE, ['C20:C30', 'H20:H30'], 20, 12, 510, 250, 'Distribuição da carteira', {
    pieHole: 0.5, colors: PALETA_GRAFICOS, pieSliceText: 'percentage',
    legend: { position: 'right', textStyle: { color: COR.TEXTO, fontSize: 10 } }
  });
  grafico_(sh, Charts.ChartType.LINE, ['H34:J120'], 34, 12, 510, 300, 'Evolução do patrimônio', {
    colors: [COR.LAVANDA, COR.ROSA], curveType: 'function', pointSize: 6, lineWidth: 3,
    vAxis: { format: 'R$ #,##0', gridlines: { color: '#F3E6EC' }, textStyle: { color: COR.TEXTO_SUAVE } },
    hAxis: { format: 'MMM/yy', textStyle: { color: COR.TEXTO_SUAVE } }
  });
}
