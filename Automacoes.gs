/**
 * ============================================================
 *  💗 AUTOMAÇÕES — menu, gatilhos, registro de lançamentos,
 *  faturas, parcelas, contas fixas e atualização do Resumo.
 * ============================================================
 */


function onOpen() {
  SpreadsheetApp.getUi().createMenu('💗 Finanças')
    .addItem('✅ Registrar lançamento', 'registrarLancamento')
    .addItem('🧹 Limpar formulário', 'limparFormulario')
    .addSeparator()
    .addItem('🔄 Atualizar Resumo', 'atualizarResumo')
    .addItem('🧮 Recalcular faturas e meses de referência', 'recalcularReferencias')
    .addItem('📸 Registrar snapshot do patrimônio', 'registrarSnapshot')
    .addSeparator()
    .addItem('🧪 Carregar dados de exemplo', 'carregarDadosExemplo')
    .addItem('🗑️ Remover dados de exemplo', 'removerDadosExemplo')
    .addItem('⚠️ Apagar TODOS os lançamentos', 'limparDados')
    .addSeparator()
    .addItem('🛠️ Reconstruir painéis (mantém seus dados)', 'setup')
    .addToUi();
}

/**
 * Gatilho simples: funciona no computador e no app de celular.
 * Trata os "botões" (caixas de seleção) e mantém tudo sincronizado.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const rg = e.range;
  const sh = rg.getSheet();
  const aba = sh.getName();
  const umaCelula = rg.getNumRows() === 1 && rg.getNumColumns() === 1;
  const a1 = rg.getA1Notation();
  const lin = rg.getRow();
  const col = rg.getColumn();
  const marcado = e.value === 'TRUE';

  try {
    if (aba === ABA.LANC && umaCelula) {
      if (a1 === FORM.btnRegistrar && marcado) {
        rg.setValue(false);
        registrarLancamento();
      } else if (a1 === FORM.btnLimpar && marcado) {
        rg.setValue(false);
        limparFormulario();
      } else if (a1 === FORM.tipo) {
        sh.getRange(FORM.cat).clearContent();
        sh.getRange(FORM.sub).clearContent();
        if (e.value !== 'Despesa') sh.getRange(FORM.cartao).clearContent();
      } else if (a1 === FORM.cat) {
        sh.getRange(FORM.sub).clearContent();
      } else if (a1 === FORM.forma && e.value !== 'Crédito') {
        sh.getRange(FORM.cartao).clearContent();
      }
      return;
    }

    if (aba === ABA.RESUMO && umaCelula) {
      if (a1 === RES.MES || a1 === RES.ANO) {
        renderResumo_();
        return;
      }
      const ref = refResumo_(sh, lerConfig_().diaInicio);

      // ✔ Pagar fatura
      if (col === RES.COL_CART_PAGA && lin >= RES.CART_INI && lin < RES.CART_INI + RES.CART_MAX) {
        const cartao = String(sh.getRange(lin, 2).getValue()).trim();
        if (!cartao) return;
        const r = pagarFatura_(cartao, ref, marcado, 'Pagamento fatura');
        if (r && r.ok === false) {
          rg.setValue(false);
          planilha_().toast(r.msg, '💳 Fatura', 6);
        } else if (marcado) {
          planilha_().toast('Fatura de ' + cartao + ' paga: ' + fmtMoeda_(r.total), '💳 Fatura', 5);
        }
        renderResumo_();
        return;
      }

      // ✔ Conta fixa paga
      if (col === RES.COL_FIXA_PAGA && lin >= RES.FIXA_INI && lin < RES.FIXA_INI + RES.FIXA_MAX) {
        const nome = String(sh.getRange(lin, 10).getValue()).trim();
        if (!nome) return;
        const valor = sh.getRange(lin, RES.COL_FIXA_VALOR).getValue();
        try {
          pagarFixa_(nome, ref, valor, marcado, 'Conta fixa');
        } catch (err) {
          rg.setValue(!marcado);
          throw err;
        }
        renderResumo_();
        return;
      }

      // Valor da conta fixa alterado depois de paga → atualiza o lançamento
      if (col === RES.COL_FIXA_VALOR && lin >= RES.FIXA_INI && lin < RES.FIXA_INI + RES.FIXA_MAX) {
        const nome = String(sh.getRange(lin, 10).getValue()).trim();
        if (nome) atualizarValorFixa_(nome, ref, rg.getValue());
      }
      return;
    }

    if (aba === ABA.CFG) {
      const ultimaLin = lin + rg.getNumRows() - 1;
      const ultimaCol = col + rg.getNumColumns() - 1;
      const tocaCartoesOuFixas = ultimaLin >= CFG_LINHA_INI && col <= CFG.FIXA_CARTAO && ultimaCol >= CFG.CART_NOME;
      const tocaCiclo = a1 === CFG_DIA_INICIO_A1;
      if (tocaCartoesOuFixas || tocaCiclo) renderResumo_();
    }
  } catch (err) {
    planilha_().toast(err.message, '⚠️ Ops!', 8);
  }
}

/* ============================================================
 *  FUNÇÕES DE DATA (puras — testáveis fora do Google)
 * ============================================================ */

function diasNoMes_(ano, mes) {
  return new Date(ano, mes + 1, 0).getDate();
}

/** Cria uma data ajustando o dia ao tamanho do mês (31/02 → 28/02). Aceita mês fora de 0..11. */
function dataSegura_(ano, mes, dia) {
  const base = new Date(ano, mes, 1);
  const a = base.getFullYear();
  const m = base.getMonth();
  return new Date(a, m, Math.min(dia, diasNoMes_(a, m)));
}

function somarMeses_(data, n) {
  return dataSegura_(data.getFullYear(), data.getMonth() + n, data.getDate());
}

function mesmoMes_(a, b) {
  return a instanceof Date && b instanceof Date &&
    a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
}

function semHora_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * Mês financeiro (1º dia do mês) de uma data.
 * Com diaInicio = 26: 25/10 → Out; 26/10 → Nov. Com diaInicio = 1: mês do calendário.
 */
function cicloMes(data, diaInicio) {
  const d = new Date(data);
  if (diaInicio > 1 && d.getDate() >= diaInicio) return new Date(d.getFullYear(), d.getMonth() + 1, 1);
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

/**
 * Data de vencimento da fatura em que uma compra no cartão vai cair.
 * Compras a partir do dia de fechamento entram na fatura seguinte.
 */
function vencimentoFatura(dataCompra, diaFech, diaVenc) {
  const d = new Date(dataCompra);
  const ano = d.getFullYear();
  const mes = d.getMonth();
  const fechEfetivo = Math.min(diaFech, diasNoMes_(ano, mes));
  const mesFech = mes + (d.getDate() >= fechEfetivo ? 1 : 0);
  const mesVenc = diaVenc > diaFech ? mesFech : mesFech + 1;
  return dataSegura_(ano, mesVenc, diaVenc);
}

/** Data de fechamento da fatura que vence em `venc`. */
function fechamentoDaFatura(venc, diaFech, diaVenc) {
  const mes = diaVenc > diaFech ? venc.getMonth() : venc.getMonth() - 1;
  return dataSegura_(venc.getFullYear(), mes, diaFech);
}

/** Data em que cai o dia `dia` dentro do ciclo do mês `ref`. */
function vencimentoNoCiclo(ref, dia, diaInicio) {
  const mes = (diaInicio > 1 && dia >= diaInicio) ? ref.getMonth() - 1 : ref.getMonth();
  return dataSegura_(ref.getFullYear(), mes, dia);
}

/**
 * Divide uma compra em parcelas. Os centavos que sobram vão para a última.
 * Retorna só as parcelas de `ini` até `n` (para compras já em andamento).
 */
function dividirParcelas(total, n, ini) {
  const centavos = Math.round(total * 100);
  const base = Math.floor(centavos / n);
  const ultima = centavos - base * (n - 1);
  const out = [];
  for (let k = ini; k <= n; k++) {
    out.push({ k: k, valor: (k === n ? ultima : base) / 100, offset: k - ini });
  }
  return out;
}

function fmtAnoMes_(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function fmtMesAno_(d) {
  return MESES[d.getMonth()].slice(0, 3) + '/' + d.getFullYear();
}

function fmtMoeda_(v) {
  const n = Number(v) || 0;
  const partes = Math.abs(n).toFixed(2).split('.');
  return (n < 0 ? '-' : '') + 'R$ ' + partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + partes[1];
}

function idFixa_(nome, ref) { return 'FIXA-' + nome + '-' + fmtAnoMes_(ref); }
function idFatura_(cartao, ref) { return 'FAT-' + cartao + '-' + fmtAnoMes_(ref); }

/* ============================================================
 *  LEITURA DE CONFIGURAÇÕES E DO BANCO DE DADOS
 * ============================================================ */

function planilha_() {
  return SpreadsheetApp.getActive();
}

function lerConfig_() {
  const sh = planilha_().getSheetByName(ABA.CFG);
  const n = CFG_LINHA_FIM - CFG_LINHA_INI + 1;
  const v = sh.getRange(CFG_LINHA_INI, 1, n, CFG.META_PRAZO).getValues();
  const txt = (r, c) => String(r[c - 1]).trim();
  const num = (r, c) => Number(r[c - 1]) || 0;

  const cartoes = {};
  const fixas = [];
  const ativosEMetas = [];
  v.forEach(r => {
    [CFG.META_NOME, CFG.ATV_NOME].forEach(c => { if (txt(r, c)) ativosEMetas.push(txt(r, c)); });
    const nomeCartao = txt(r, CFG.CART_NOME);
    if (nomeCartao) {
      cartoes[nomeCartao] = {
        nome: nomeCartao,
        limite: num(r, CFG.CART_LIMITE),
        fech: num(r, CFG.CART_FECH) || 1,
        venc: num(r, CFG.CART_VENC) || 10,
        conta: txt(r, CFG.CART_CONTA)
      };
    }
    const nomeFixa = txt(r, CFG.FIXA_NOME);
    if (nomeFixa) {
      fixas.push({
        nome: nomeFixa,
        cat: txt(r, CFG.FIXA_CAT),
        valor: num(r, CFG.FIXA_VALOR),
        dia: num(r, CFG.FIXA_DIA) || 1,
        conta: txt(r, CFG.FIXA_CONTA),
        forma: txt(r, CFG.FIXA_FORMA),
        cartao: txt(r, CFG.FIXA_CARTAO)
      });
    }
  });

  const dia = Number(sh.getRange(CFG_DIA_INICIO_A1).getValue());
  return {
    diaInicio: dia >= 1 && dia <= 28 ? dia : 1,
    cartoes: cartoes,
    listaCartoes: Object.keys(cartoes),
    fixas: fixas,
    ativosEMetas: ativosEMetas
  };
}

/** Mapa ID → { linha, valor } de todos os lançamentos. */
function mapaBD_() {
  const sh = planilha_().getSheetByName(ABA.BD);
  const m = new Map();
  const last = sh.getLastRow();
  if (last < 2) return m;
  sh.getRange(2, 1, last - 1, C.VALOR).getValues().forEach((r, i) => {
    if (r[0] !== '') m.set(String(r[0]), { linha: i + 2, valor: r[C.VALOR - 1] });
  });
  return m;
}

function contexto_() {
  return { cfg: lerConfig_(), ids: new Set(mapaBD_().keys()) };
}

function novoId_(prefixo) {
  return prefixo + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

/* ============================================================
 *  GRAVAÇÃO NO BANCO DE DADOS
 * ============================================================ */

/**
 * Transforma um lançamento (objeto) em uma ou mais linhas do BD.
 * Compras parceladas viram uma linha por parcela, cada uma com sua fatura.
 */
function montarLinhas_(f, ctx, origem) {
  const cfg = ctx.cfg;
  const credito = f.forma === 'Crédito' && f.tipo === 'Despesa';
  const cartao = credito ? cfg.cartoes[f.cartao] : null;
  if (credito && !cartao) throw new Error('Cartão "' + f.cartao + '" não está cadastrado em Configurações.');

  const parcelado = f.parcelado === 'Sim';
  const total = parcelado ? (Number(f.valorTotal) || Number(f.valor)) : Number(f.valor);
  const partes = parcelado
    ? dividirParcelas(total, Number(f.nParc), Number(f.parcIni) || 1)
    : [{ k: '', valor: total, offset: 0 }];
  const grupo = parcelado ? novoId_('G') : '';
  const vencBase = credito ? vencimentoFatura(f.data, cartao.fech, cartao.venc) : null;
  const agora = new Date();
  const qtd = Number(f.qtd) || 0;

  return partes.map(p => {
    let data = f.data;
    let venc = '';
    let mesRef;
    let status = f.status || 'Pago';
    if (credito) {
      venc = somarMeses_(vencBase, p.offset);
      mesRef = cicloMes(venc, cfg.diaInicio);
      status = ctx.ids.has(idFatura_(cartao.nome, mesRef)) ? 'Pago' : 'Pendente';
    } else {
      if (p.offset) data = somarMeses_(f.data, p.offset);
      mesRef = cicloMes(data, cfg.diaInicio);
    }

    const l = new Array(BD_NCOLS).fill('');
    l[C.ID - 1] = (f.id && !parcelado) ? f.id : novoId_('L');
    l[C.REG - 1] = agora;
    l[C.DATA - 1] = data;
    l[C.DESC - 1] = f.desc;
    l[C.TIPO - 1] = f.tipo;
    l[C.CAT - 1] = f.cat;
    l[C.SUB - 1] = f.sub || '';
    l[C.CONTA - 1] = f.conta || (cartao ? cartao.conta : '');
    l[C.FORMA - 1] = f.forma || '';
    l[C.CARTAO - 1] = cartao ? cartao.nome : '';
    l[C.VALOR - 1] = p.valor;
    l[C.STATUS - 1] = status;
    l[C.FIXA - 1] = f.fixa === 'Sim' ? 'Sim' : 'Não';
    l[C.PARC - 1] = parcelado ? 'Sim' : 'Não';
    l[C.PARC_N - 1] = parcelado ? p.k : '';
    l[C.PARC_TOT - 1] = parcelado ? Number(f.nParc) : '';
    l[C.VALOR_TOTAL - 1] = parcelado ? total : '';
    l[C.GRUPO - 1] = grupo;
    l[C.VENC - 1] = venc;
    l[C.MESREF - 1] = mesRef;
    l[C.ATIVO - 1] = f.tipo === 'Investimento' ? (f.ativo || '') : '';
    l[C.QTD - 1] = f.tipo === 'Investimento' && qtd ? qtd : '';
    l[C.PRECO - 1] = f.tipo === 'Investimento' && qtd ? p.valor / qtd : '';
    l[C.ORIGEM - 1] = origem;
    l[C.OBS - 1] = f.obs || '';
    return l;
  });
}

function comTrava_(fn) {
  let trava = null;
  try {
    trava = LockService.getDocumentLock();
    trava.waitLock(20000);
  } catch (e) {
    trava = null;
  }
  try {
    return fn();
  } finally {
    if (trava) trava.releaseLock();
  }
}

/** Garante que a aba tenha pelo menos `ultima` + 1 linhas (inserindo dentro dos intervalos nomeados). */
function garantirLinhas_(sh, ultima) {
  const max = sh.getMaxRows();
  if (max > ultima) return;
  sh.insertRowsBefore(max, ultima - max + 200);
}

function gravarLinhas_(linhas) {
  if (!linhas.length) return;
  const sh = planilha_().getSheetByName(ABA.BD);
  comTrava_(() => {
    const ini = Math.max(sh.getLastRow(), 1) + 1;
    garantirLinhas_(sh, ini + linhas.length);
    sh.getRange(ini, 1, linhas.length, BD_NCOLS).setValues(linhas);
  });
}

function removerPorIds_(ids) {
  const alvo = new Set(ids);
  const sh = planilha_().getSheetByName(ABA.BD);
  comTrava_(() => {
    const last = sh.getLastRow();
    if (last < 2) return;
    const col = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = col.length - 1; i >= 0; i--) {
      if (alvo.has(String(col[i][0]))) sh.deleteRow(i + 2);
    }
  });
}

/** Mantém só as linhas que passam no filtro (reescreve o BD de uma vez). */
function filtrarBD_(manter) {
  const sh = planilha_().getSheetByName(ABA.BD);
  comTrava_(() => {
    const last = sh.getLastRow();
    if (last < 2) return;
    const dados = sh.getRange(2, 1, last - 1, BD_NCOLS).getValues();
    const ficam = dados.filter(manter);
    sh.getRange(2, 1, last - 1, BD_NCOLS).clearContent();
    if (ficam.length) sh.getRange(2, 1, ficam.length, BD_NCOLS).setValues(ficam);
  });
}

/* ============================================================
 *  FORMULÁRIO (aba Lançamentos)
 * ============================================================ */

function lerFormulario_(sh) {
  const vals = sh.getRange('B5:E17').getValues();
  const pega = a1 => {
    const m = /^([A-Z])(\d+)$/.exec(a1);
    return vals[Number(m[2]) - 5][m[1].charCodeAt(0) - 66];
  };
  const txt = a1 => String(pega(a1)).trim();
  const data = pega(FORM.data);
  return {
    data: data instanceof Date ? semHora_(data) : null,
    desc: txt(FORM.desc),
    valor: pega(FORM.valor),
    conta: txt(FORM.conta),
    forma: txt(FORM.forma),
    fixa: txt(FORM.fixa) || 'Não',
    tipo: txt(FORM.tipo),
    cat: txt(FORM.cat),
    sub: txt(FORM.sub),
    status: txt(FORM.status) || 'Pago',
    cartao: txt(FORM.cartao),
    obs: txt(FORM.obs),
    parcelado: txt(FORM.parcelado) || 'Não',
    valorTotal: pega(FORM.valorTotal),
    nParc: pega(FORM.nParc),
    parcIni: pega(FORM.parcIni) || 1,
    ativo: txt(FORM.ativo),
    qtd: pega(FORM.qtd)
  };
}

function validarFormulario_(f, cfg) {
  const faltando = [];
  if (!f.data) faltando.push('Data');
  if (!f.desc) faltando.push('Descrição');
  if (!f.tipo) faltando.push('Tipo');
  if (!f.cat) faltando.push('Categoria');
  if (f.tipo === 'Despesa' && !f.forma) faltando.push('Forma de pagamento');
  if (!f.conta && f.forma !== 'Crédito') faltando.push('Conta');
  if (f.tipo === 'Investimento' && !f.ativo) faltando.push('Ativo ou meta');

  const erros = [];
  if (faltando.length) erros.push('Preencha: ' + faltando.join(', '));

  if (f.forma === 'Crédito') {
    if (f.tipo !== 'Despesa') erros.push('Crédito só pode ser usado em Despesas');
    else if (!f.cartao) erros.push('Escolha o cartão utilizado');
    else if (!cfg.cartoes[f.cartao]) erros.push('Cartão não cadastrado em Configurações');
  }

  if (f.parcelado === 'Sim') {
    const n = Number(f.nParc);
    const ini = Number(f.parcIni);
    const total = Number(f.valorTotal) || Number(f.valor);
    if (!Number.isInteger(n) || n < 2) erros.push('Nº de parcelas deve ser 2 ou mais');
    else if (!Number.isInteger(ini) || ini < 1 || ini > n) erros.push('Parcela inicial deve estar entre 1 e ' + n);
    if (!(total > 0)) erros.push('Informe o valor total da compra');
  } else if (!(Number(f.valor) > 0)) {
    erros.push('Informe um valor maior que zero');
  }
  return erros;
}

function mensagem_(sh, texto, ok) {
  sh.getRange(FORM.msg).setValue(texto).setFontColor(ok ? COR.MENTA_ESCURA : COR.CORAL_ESCURO);
}

/** Lê o formulário, valida e envia para o Banco de Dados. */
function registrarLancamento() {
  verificarInstalacao_();
  const ss = planilha_();
  const sh = ss.getSheetByName(ABA.LANC);
  try {
    const f = lerFormulario_(sh);
    const ctx = contexto_();
    // Meta/ativo escolhido como subcategoria (ex.: Metas → Carro) vale como "Ativo ou meta"
    if (f.tipo === 'Investimento' && !f.ativo) {
      const nome = ctx.cfg.ativosEMetas.find(x => x.toLowerCase() === f.sub.toLowerCase());
      if (nome) f.ativo = nome;
    }
    const erros = validarFormulario_(f, ctx.cfg);
    if (erros.length) {
      mensagem_(sh, '⚠️ ' + erros.join(' • '), false);
      return;
    }

    // Despesa fixa com o mesmo nome de uma conta fixa → marca o checklist do Resumo também
    if (f.fixa === 'Sim' && f.parcelado !== 'Sim') {
      const fx = ctx.cfg.fixas.find(x => x.nome.toLowerCase() === f.desc.toLowerCase());
      if (fx) {
        const id = idFixa_(fx.nome, cicloMes(f.data, ctx.cfg.diaInicio));
        if (!ctx.ids.has(id)) f.id = id;
      }
    }

    const linhas = montarLinhas_(f, ctx, 'Formulário');
    gravarLinhas_(linhas);

    const primeira = linhas[0];
    let texto;
    if (f.parcelado === 'Sim') {
      texto = '✔ ' + f.desc + ': ' + linhas.length + ' parcela(s) de ' + fmtMoeda_(primeira[C.VALOR - 1]) +
        ' registradas (' + primeira[C.PARC_N - 1] + '/' + f.nParc + ' a ' + f.nParc + '/' + f.nParc + ')';
    } else {
      texto = '✔ ' + f.desc + ' — ' + fmtMoeda_(primeira[C.VALOR - 1]) + ' registrado no mês ' +
        fmtMesAno_(primeira[C.MESREF - 1]);
    }
    if (primeira[C.VENC - 1]) texto += ' • fatura ' + f.cartao + ' vence ' +
      Utilities.formatDate(primeira[C.VENC - 1], FUSO, 'dd/MM/yyyy');

    limparFormulario_(sh);
    mensagem_(sh, texto, true);
    ss.toast(texto, '💗 Lançamento registrado', 5);
    if (f.id) renderResumo_();
  } catch (err) {
    mensagem_(sh, '⚠️ ' + err.message, false);
  }
}

function limparFormulario() {
  verificarInstalacao_();
  const sh = planilha_().getSheetByName(ABA.LANC);
  limparFormulario_(sh);
  sh.getRange(FORM.msg).clearContent();
}

function limparFormulario_(sh) {
  const padrao = {};
  Object.keys(FORM).forEach(k => {
    if (['btnRegistrar', 'btnLimpar', 'msg'].indexOf(k) < 0) padrao[FORM[k]] = '';
  });
  padrao[FORM.data] = semHora_(new Date());
  padrao[FORM.tipo] = 'Despesa';
  padrao[FORM.status] = 'Pago';
  padrao[FORM.fixa] = 'Não';
  padrao[FORM.parcelado] = 'Não';
  padrao[FORM.parcIni] = 1;
  Object.keys(padrao).forEach(a1 => sh.getRange(a1).setValue(padrao[a1]));
}

/* ============================================================
 *  FATURAS E CONTAS FIXAS
 * ============================================================ */

/**
 * Marca (ou desmarca) a fatura de um cartão como paga.
 * - As compras daquela fatura passam a "Pago" (liberando o limite).
 * - É criada uma Transferência "Pagamento de fatura" (não conta como gasto).
 */
function pagarFatura_(nomeCartao, ref, pagar, origem) {
  const sh = planilha_().getSheetByName(ABA.BD);
  const cartao = lerConfig_().cartoes[nomeCartao];
  if (!cartao) throw new Error('Cartão "' + nomeCartao + '" não encontrado em Configurações.');

  let total = 0;
  const last = sh.getLastRow();
  if (last >= 2) {
    const dados = sh.getRange(2, 1, last - 1, BD_NCOLS).getValues();
    const status = dados.map(r => [r[C.STATUS - 1]]);
    const indices = [];
    dados.forEach((r, i) => {
      if (r[C.TIPO - 1] === 'Despesa' && r[C.FORMA - 1] === 'Crédito' &&
        r[C.CARTAO - 1] === nomeCartao && mesmoMes_(r[C.MESREF - 1], ref)) {
        total += Number(r[C.VALOR - 1]) || 0;
        indices.push(i);
      }
    });
    if (pagar && total <= 0) return { ok: false, msg: 'A fatura de ' + nomeCartao + ' neste mês não tem lançamentos.' };
    indices.forEach(i => { status[i][0] = pagar ? 'Pago' : 'Pendente'; });
    if (indices.length) sh.getRange(2, C.STATUS, status.length, 1).setValues(status);
  } else if (pagar) {
    return { ok: false, msg: 'Ainda não há lançamentos.' };
  }

  const id = idFatura_(nomeCartao, ref);
  removerPorIds_([id]);
  if (pagar) {
    const l = new Array(BD_NCOLS).fill('');
    l[C.ID - 1] = id;
    l[C.REG - 1] = new Date();
    l[C.DATA - 1] = semHora_(new Date());
    l[C.DESC - 1] = 'Pagamento da fatura ' + nomeCartao + ' (' + fmtMesAno_(ref) + ')';
    l[C.TIPO - 1] = 'Transferência';
    l[C.CAT - 1] = 'Pagamento de fatura';
    l[C.CONTA - 1] = cartao.conta;
    l[C.FORMA - 1] = 'Transferência';
    l[C.CARTAO - 1] = nomeCartao;
    l[C.VALOR - 1] = Math.round(total * 100) / 100;
    l[C.STATUS - 1] = 'Pago';
    l[C.FIXA - 1] = 'Não';
    l[C.PARC - 1] = 'Não';
    l[C.MESREF - 1] = ref;
    l[C.ORIGEM - 1] = origem || 'Pagamento fatura';
    gravarLinhas_([l]);
  }
  return { ok: true, total: total };
}

/** Marca (ou desmarca) uma conta fixa como paga no mês `ref`. */
function pagarFixa_(nome, ref, valor, pagar, origem) {
  const ctx = contexto_();
  const fx = ctx.cfg.fixas.find(x => x.nome === nome);
  if (!fx) throw new Error('Conta fixa "' + nome + '" não encontrada em Configurações.');
  const id = idFixa_(nome, ref);

  if (!pagar) {
    removerPorIds_([id]);
    return;
  }
  if (ctx.ids.has(id)) return; // já está paga

  const v = Number(valor) || fx.valor;
  if (!(v > 0)) throw new Error('Informe o valor de "' + nome + '" antes de marcar como paga.');

  let forma = fx.forma || 'Débito';
  let cartao = '';
  if (forma === 'Crédito') {
    if (ctx.cfg.cartoes[fx.cartao]) cartao = fx.cartao;
    else forma = 'Débito';
  }
  const f = {
    id: id, data: vencimentoNoCiclo(ref, fx.dia, ctx.cfg.diaInicio), desc: fx.nome,
    tipo: 'Despesa', cat: fx.cat || 'Contas', sub: '', conta: fx.conta || 'Conta corrente',
    forma: forma, cartao: cartao, valor: v, status: 'Pago', fixa: 'Sim', parcelado: 'Não',
    obs: 'Marcada no checklist do Resumo'
  };
  gravarLinhas_(montarLinhas_(f, ctx, origem || 'Conta fixa'));
}

function atualizarValorFixa_(nome, ref, valor) {
  const reg = mapaBD_().get(idFixa_(nome, ref));
  if (!reg || !(Number(valor) > 0)) return;
  planilha_().getSheetByName(ABA.BD).getRange(reg.linha, C.VALOR).setValue(Number(valor));
}

/* ============================================================
 *  RESUMO — partes que dependem de script (checklists)
 * ============================================================ */

function refResumo_(sh, diaInicio) {
  const v = sh.getRange(RES.MES + ':' + RES.ANO).getValues();
  const idx = MESES.indexOf(String(v[0][0]).trim());
  const ano = Number(v[1][0]);
  if (idx < 0 || !ano) return cicloMes(new Date(), diaInicio);
  return new Date(ano, idx, 1);
}

function atualizarResumo() {
  verificarInstalacao_();
  renderResumo_();
  planilha_().toast('Resumo atualizado ✨', '💗 Finanças', 3);
}

/**
 * Redesenha as tabelas de cartões e contas fixas do Resumo para o mês
 * selecionado (nomes, valores e caixas de seleção). O resto é fórmula.
 */
function renderResumo_() {
  const sh = planilha_().getSheetByName(ABA.RESUMO);
  if (!sh) return;
  const cfg = lerConfig_();
  const ref = refResumo_(sh, cfg.diaInicio);
  const bd = mapaBD_();

  // ---------- Cartões (B..H) ----------
  const c0 = RES.CART_INI;
  const nC = RES.CART_MAX;
  sh.getRange(c0, RES.COL_CART_PAGA, nC, 1).removeCheckboxes();
  sh.getRange(c0, 2, nC, 7).clearContent();
  const cartoes = cfg.listaCartoes.slice(0, nC);
  if (cartoes.length) {
    const nomes = [];
    const formulas = [];
    const pagos = [];
    cartoes.forEach((nome, i) => {
      const c = cfg.cartoes[nome];
      const r = c0 + i;
      const fech = fechamentoDaFatura(vencimentoNoCiclo(ref, c.venc, cfg.diaInicio), c.fech, c.venc);
      const dataFech = 'DATE(' + fech.getFullYear() + ',' + (fech.getMonth() + 1) + ',' + fech.getDate() + ')';
      nomes.push([nome]);
      formulas.push([
        '=IFERROR(INDEX(CART_LIMITE,MATCH($B' + r + ',CART_NOME,0)),0)',
        '=SUMIFS(BD_VALOR,BD_CARTAO,$B' + r + ',BD_FORMA,"Crédito",BD_TIPO,"Despesa",BD_MESREF,RES_REF)',
        '=IF(COUNTIF(BD_ID,"FAT-"&$B' + r + '&"-"&TEXT(RES_REF,"yyyy-mm"))>0,"✅ Paga",IF($D' + r +
          '=0,"— sem gastos",IF(TODAY()>=' + dataFech + ',"🔒 Fechada","🟢 Aberta")))',
        '=IF($C' + r + '>0,SPARKLINE(MAX(0,$C' + r + '-$G' + r + '),{"charttype","bar";"max",$C' + r +
          ';"color1","' + COR.ROSA + '";"color2","' + COR.ROSE + '"}),"")',
        '=$C' + r + '-SUMIFS(BD_VALOR,BD_CARTAO,$B' + r + ',BD_FORMA,"Crédito",BD_STATUS,"Pendente")'
      ]);
      pagos.push([bd.has(idFatura_(nome, ref))]);
    });
    sh.getRange(c0, 2, cartoes.length, 1).setValues(nomes);
    sh.getRange(c0, 3, cartoes.length, 5).setFormulas(fxs_(formulas));
    const cb = sh.getRange(c0, RES.COL_CART_PAGA, cartoes.length, 1);
    cb.insertCheckboxes();
    cb.setValues(pagos);
  }

  // ---------- Contas fixas (J..Q) ----------
  const f0 = RES.FIXA_INI;
  const nF = RES.FIXA_MAX;
  sh.getRange(f0, RES.COL_FIXA_PAGA, nF, 1).removeCheckboxes();
  sh.getRange(f0, 10, nF, 8).clearContent();
  const fixas = cfg.fixas.slice(0, nF);
  if (fixas.length) {
    const valores = [];
    const pagos = [];
    const status = [];
    fixas.forEach((fx, i) => {
      const r = f0 + i;
      const reg = bd.get(idFixa_(fx.nome, ref));
      valores.push([fx.nome, '', fx.cat, reg ? reg.valor : fx.valor, vencimentoNoCiclo(ref, fx.dia, cfg.diaInicio)]);
      pagos.push([!!reg]);
      status.push(['=IF(COUNTIF(BD_ID,"FIXA-"&$J' + r + '&"-"&TEXT(RES_REF,"yyyy-mm"))>0,"✅ Paga",' +
        'IF(TODAY()>$N' + r + ',"⚠️ Atrasada","⏳ Pendente"))']);
    });
    sh.getRange(f0, 10, fixas.length, 5).setValues(valores);
    const cb = sh.getRange(f0, RES.COL_FIXA_PAGA, fixas.length, 1);
    cb.insertCheckboxes();
    cb.setValues(pagos);
    sh.getRange(f0, 16, fixas.length, 1).setFormulas(fxs_(status));
  }
}

/* ============================================================
 *  MANUTENÇÃO
 * ============================================================ */

/**
 * Recalcula vencimento de fatura e mês de referência de todos os lançamentos.
 * Use depois de mudar o fechamento/vencimento de um cartão ou o ciclo do mês.
 */
function recalcularReferencias() {
  verificarInstalacao_();
  const ss = planilha_();
  const sh = ss.getSheetByName(ABA.BD);
  const last = sh.getLastRow();
  if (last < 2) return;
  const cfg = lerConfig_();
  const dados = sh.getRange(2, 1, last - 1, BD_NCOLS).getValues();

  // menor parcela de cada grupo (para calcular o deslocamento de cada parcela)
  const menorParcela = {};
  dados.forEach(r => {
    const g = r[C.GRUPO - 1];
    const k = Number(r[C.PARC_N - 1]);
    if (g && k && (!menorParcela[g] || k < menorParcela[g])) menorParcela[g] = k;
  });

  let alterados = 0;
  const saida = dados.map(r => {
    let venc = r[C.VENC - 1];
    let mesRef = r[C.MESREF - 1];
    const id = String(r[C.ID - 1]);
    const data = r[C.DATA - 1];
    if (id.indexOf('FAT-') === 0 || !(data instanceof Date)) return [venc, mesRef];
    const cartao = cfg.cartoes[r[C.CARTAO - 1]];
    if (r[C.FORMA - 1] === 'Crédito' && cartao) {
      const g = r[C.GRUPO - 1];
      const offset = g ? Number(r[C.PARC_N - 1]) - menorParcela[g] : 0;
      venc = somarMeses_(vencimentoFatura(data, cartao.fech, cartao.venc), offset);
      mesRef = cicloMes(venc, cfg.diaInicio);
    } else {
      mesRef = cicloMes(data, cfg.diaInicio);
    }
    if (!mesmoMes_(mesRef, r[C.MESREF - 1])) alterados++;
    return [venc, mesRef];
  });
  sh.getRange(2, C.VENC, saida.length, 1).setValues(saida.map(s => [s[0]]));
  sh.getRange(2, C.MESREF, saida.length, 1).setValues(saida.map(s => [s[1]]));
  renderResumo_();
  ss.toast(alterados + ' lançamento(s) mudaram de mês de referência.', '🧮 Recalculado', 6);
}

function limparDados() {
  verificarInstalacao_();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Apagar todos os lançamentos?',
    'Isso remove TODAS as linhas do Banco de Dados. Configurações e metas são mantidas.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  filtrarBD_(() => false);
  renderResumo_();
  planilha_().toast('Banco de Dados zerado.', '🗑️ Pronto', 4);
}

/** Interrompe com uma mensagem clara se o setup ainda não criou as abas. */
function verificarInstalacao_() {
  const ss = planilha_();
  const faltando = Object.keys(ABA).map(k => ABA[k]).filter(nome => !ss.getSheetByName(nome));
  if (faltando.length) {
    throw new Error('A planilha ainda não foi montada (faltam as abas: ' + faltando.join(', ') +
      '). Abra Extensões → Apps Script, escolha a função "setup" e clique em ▶ Executar.');
  }
}

/* ============================================================
 *  FÓRMULAS NA LOCALIDADE DA PLANILHA
 *  O código escreve fórmulas no padrão americano (vírgulas).
 *  Em localidades com vírgula decimal (ex.: pt_BR), o Google
 *  espera ";" entre argumentos e "\" entre colunas de { }.
 * ============================================================ */

let USA_PONTO_VIRGULA_ = null;

function usaPontoVirgula_() {
  if (USA_PONTO_VIRGULA_ === null) {
    const loc = String(planilha_().getSpreadsheetLocale() || 'en_US');
    USA_PONTO_VIRGULA_ = !/^(en|ja|zh|ko|th|he|iw|hi|ms|fil|tl)([_-]|$)/i.test(loc);
  }
  return USA_PONTO_VIRGULA_;
}

/** Converte uma fórmula do padrão americano para a localidade da planilha. */
function fx_(formula) {
  if (typeof formula !== 'string' || formula.charAt(0) !== '=' || !usaPontoVirgula_()) return formula;
  return converterFormula_(formula);
}

function fxs_(matriz) {
  return matriz.map(linha => linha.map(fx_));
}

/** Pura (testável): "," → ";" entre argumentos e "," → "\" entre colunas de { }, fora de textos. */
function converterFormula_(f) {
  const pilha = [];
  let out = '';
  let emTexto = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f.charAt(i);
    if (ch === '"') emTexto = !emTexto;
    if (emTexto || ch === '"') { out += ch; continue; }
    if (ch === '(' || ch === '{') pilha.push(ch);
    else if (ch === ')' || ch === '}') pilha.pop();
    if (ch === ',') out += pilha[pilha.length - 1] === '{' ? '\\' : ';';
    else out += ch;
  }
  return out;
}
