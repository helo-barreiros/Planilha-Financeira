/**
 * ============================================================
 *  🧪 DADOS DE EXEMPLO — para ver o painel funcionando.
 *  Tudo é marcado com Origem = "Exemplo" e pode ser removido
 *  pelo menu 💗 Finanças → Remover dados de exemplo.
 * ============================================================
 */

function carregarDadosExemplo() {
  verificarInstalacao_();
  const ss = planilha_();
  const shBD = ss.getSheetByName(ABA.BD);
  const last = shBD.getLastRow();
  if (last >= 2) {
    const origens = shBD.getRange(2, C.ORIGEM, last - 1, 1).getValues();
    if (origens.some(o => o[0] === 'Exemplo')) {
      ss.toast('Os dados de exemplo já estão carregados. Remova-os antes de carregar de novo.', '🧪 Exemplo', 6);
      return;
    }
  }

  const ctx = contexto_();
  const cfg = ctx.cfg;
  if (!cfg.listaCartoes.length) throw new Error('Cadastre pelo menos um cartão em Configurações.');
  const c1 = cfg.listaCartoes[0];
  const c2 = cfg.listaCartoes[1] || c1;
  const ativos = listaConfig_(CFG.ATV_NOME);
  const metas = listaConfig_(CFG.META_NOME);

  const refAtual = cicloMes(new Date(), cfg.diaInicio);
  const dia = (ref, d) => vencimentoNoCiclo(ref, d, cfg.diaInicio);
  const linhas = [];
  const add = f => {
    const base = { status: 'Pago', fixa: 'Não', parcelado: 'Não', conta: 'Conta corrente', sub: '', obs: '' };
    montarLinhas_(Object.assign(base, f), ctx, 'Exemplo').forEach(l => linhas.push(l));
  };

  // [descrição, tipo, categoria, subcategoria, forma, cartão, dia, valores dos meses -3, -2, -1, 0]
  const recorrentes = [
    ['Salário', 'Receita', 'Salário', 'Mensal', 'Transferência', '', 5, [5200, 5200, 5200, 5200]],
    ['Freela de design', 'Receita', 'Renda extra', 'Freelance', 'Pix', '', 18, [0, 650, 0, 480]],
    ['Supermercado', 'Despesa', 'Alimentação', 'Mercado', 'Débito', '', 8, [430, 455, 398, 412]],
    ['Padaria e feira', 'Despesa', 'Alimentação', 'Padaria', 'Pix', '', 16, [85, 92, 78, 88]],
    ['Delivery', 'Despesa', 'Alimentação', 'Delivery', 'Crédito', c1, 12, [120, 95, 140, 110]],
    ['Restaurante', 'Despesa', 'Alimentação', 'Restaurante', 'Crédito', c2, 22, [180, 210, 160, 0]],
    ['App de transporte', 'Despesa', 'Transporte', 'App de transporte', 'Pix', '', 14, [96, 110, 84, 70]],
    ['Farmácia', 'Despesa', 'Saúde', 'Farmácia', 'Crédito', c2, 19, [65, 0, 120, 45]],
    ['Cinema e passeio', 'Despesa', 'Lazer', 'Cinema e shows', 'Crédito', c1, 24, [90, 150, 60, 0]],
    ['Roupas', 'Despesa', 'Compras', 'Roupas', 'Crédito', c1, 3, [0, 260, 0, 180]],
    ['Curso online', 'Despesa', 'Educação', 'Cursos', 'Pix', '', 10, [0, 0, 197, 0]],
    ['Skincare', 'Despesa', 'Cuidados pessoais', '', 'Crédito', c2, 7, [140, 0, 95, 0]]
  ];

  for (let o = -3; o <= 0; o++) {
    const ref = somarMeses_(refAtual, o);
    recorrentes.forEach(x => {
      const valor = x[7][o + 3];
      if (!valor) return;
      add({
        data: dia(ref, x[6]), desc: x[0], tipo: x[1], cat: x[2], sub: x[3], forma: x[4], cartao: x[5],
        valor: valor, conta: x[1] === 'Receita' ? 'Conta corrente' : (x[4] === 'Pix' ? 'Conta digital' : 'Conta corrente')
      });
    });

    // Investimentos e metas
    if (ativos[0]) add({ data: dia(ref, 6), desc: 'Aporte ' + ativos[0], tipo: 'Investimento', cat: 'Investimentos',
      sub: 'Cripto', forma: 'Pix', conta: 'Conta digital', valor: 300, ativo: ativos[0], qtd: 0.00085 });
    if (ativos[1] && o % 2 === 0) add({ data: dia(ref, 15), desc: 'Aporte ' + ativos[1], tipo: 'Investimento',
      cat: 'Investimentos', sub: 'FIIs', forma: 'Pix', conta: 'Conta digital', valor: 98, ativo: ativos[1], qtd: 10 });
    if (ativos[2] && o === -1) add({ data: dia(ref, 15), desc: 'Aporte ' + ativos[2], tipo: 'Investimento',
      cat: 'Investimentos', sub: 'FIIs', forma: 'Pix', conta: 'Conta digital', valor: 160, ativo: ativos[2], qtd: 1 });
    if (metas[0]) add({ data: dia(ref, 6), desc: 'Guardar para ' + metas[0], tipo: 'Investimento', cat: 'Metas',
      forma: 'Pix', conta: 'Conta digital', valor: 400, ativo: metas[0] });
    if (metas[1]) add({ data: dia(ref, 6), desc: 'Guardar para ' + metas[1], tipo: 'Investimento', cat: 'Metas',
      forma: 'Pix', conta: 'Conta digital', valor: 250, ativo: metas[1] });
    if (metas[2] && o === -2) add({ data: dia(ref, 6), desc: 'Guardar para ' + metas[2], tipo: 'Investimento',
      cat: 'Metas', forma: 'Pix', conta: 'Conta digital', valor: 500, ativo: metas[2] });

    // Contas fixas: meses anteriores todas pagas; mês atual só as 3 primeiras
    cfg.fixas.forEach((fx, i) => {
      if (o === 0 && i >= 3) return;
      if (!(fx.valor > 0)) return;
      let forma = fx.forma || 'Débito';
      if (forma === 'Crédito' && !cfg.cartoes[fx.cartao]) forma = 'Débito';
      add({
        id: idFixa_(fx.nome, ref), data: dia(ref, fx.dia), desc: fx.nome, tipo: 'Despesa', cat: fx.cat || 'Contas',
        forma: forma, cartao: forma === 'Crédito' ? fx.cartao : '', conta: fx.conta || 'Conta corrente',
        valor: fx.valor, fixa: 'Sim'
      });
    });
  }

  // Compras parceladas
  add({ data: dia(somarMeses_(refAtual, -3), 14), desc: 'Notebook', tipo: 'Despesa', cat: 'Compras', sub: 'Eletrônicos',
    forma: 'Crédito', cartao: c2, parcelado: 'Sim', valorTotal: 3600, nParc: 12, parcIni: 1 });
  add({ data: dia(somarMeses_(refAtual, -1), 9), desc: 'Tênis de corrida', tipo: 'Despesa', cat: 'Compras',
    sub: 'Roupas', forma: 'Crédito', cartao: c1, parcelado: 'Sim', valorTotal: 450, nParc: 3, parcIni: 1 });
  add({ data: dia(somarMeses_(refAtual, -2), 20), desc: 'Geladeira (já estava na 4ª parcela)', tipo: 'Despesa',
    cat: 'Compras', sub: 'Casa', forma: 'Crédito', cartao: c1, parcelado: 'Sim', valorTotal: 2400, nParc: 10, parcIni: 4 });

  gravarLinhas_(linhas);

  // Faturas dos meses anteriores pagas
  for (let o = -3; o <= -1; o++) {
    const ref = somarMeses_(refAtual, o);
    [c1, c2].filter((c, i, arr) => arr.indexOf(c) === i).forEach(c => pagarFatura_(c, ref, true, 'Exemplo'));
  }

  renderResumo_();
  ss.toast(linhas.length + ' lançamentos de exemplo criados. Explore o Resumo! 💗', '🧪 Exemplo', 6);
}

function removerDadosExemplo() {
  verificarInstalacao_();
  filtrarBD_(r => r[C.ORIGEM - 1] !== 'Exemplo');
  renderResumo_();
  planilha_().toast('Dados de exemplo removidos.', '🧪 Exemplo', 4);
}

/** Valores não vazios de uma coluna da aba Configurações. */
function listaConfig_(col) {
  const sh = planilha_().getSheetByName(ABA.CFG);
  return sh.getRange(CFG_LINHA_INI, col, CFG_LINHA_FIM - CFG_LINHA_INI + 1, 1).getValues()
    .map(r => String(r[0]).trim()).filter(String);
}
