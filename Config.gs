/**
 * ============================================================
 *  💗 PLANILHA DE CONTROLE FINANCEIRO — CONSTANTES
 * ============================================================
 *  Tudo que é compartilhado entre os arquivos: nomes das abas,
 *  colunas do Banco de Dados, posições do formulário, paleta.
 */

const ABA = {
  RESUMO: 'Resumo',
  INV: 'Investimentos',
  LANC: 'Lançamentos',
  CFG: 'Configurações',
  BD: 'Banco de Dados'
};

const FUSO = 'America/Sao_Paulo';
const FONTE = 'Montserrat';

const COR = {
  AMEIXA: '#5E3A5C',
  AMEIXA_CLARA: '#8A5A86',
  ROSA: '#C77DA0',
  ROSE: '#F3C6D7',
  ROSE_CLARO: '#F8DCE7',
  BLUSH: '#FCEFF4',
  ZEBRA: '#FDF7FA',
  LAVANDA: '#CDB4DB',
  LAVANDA_CLARA: '#F3ECF8',
  PESSEGO: '#FDF1EA',
  MENTA: '#7CC4A4',
  MENTA_ESCURA: '#3F8F6B',
  MENTA_CLARA: '#EAF6F0',
  CORAL: '#E07A7A',
  CORAL_ESCURO: '#C0504D',
  TEXTO: '#4A4A4A',
  TEXTO_SUAVE: '#8C7F88',
  CINZA: '#EDE7EB',
  BRANCO: '#FFFFFF'
};

/** Cores usadas nos gráficos (em ordem). */
const PALETA_GRAFICOS = ['#C77DA0', '#8A5A86', '#F2A7C3', '#CDB4DB', '#7CC4A4',
  '#F6C28B', '#9AB7E0', '#E07A7A', '#B5838D', '#6D597A', '#FFCAD4', '#A3C4BC'];

const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

const TIPOS = ['Receita', 'Despesa', 'Investimento', 'Transferência'];
const STATUS = ['Pago', 'Pendente'];
const SIM_NAO = ['Sim', 'Não'];

const FORMATO_MOEDA = '"R$" #,##0.00';
const FORMATO_DATA = 'dd/mm/yyyy';
const FORMATO_PCT = '0.0%';

/* ---------- Banco de Dados ---------- */

const BD_CABECALHO = ['ID', 'Registrado em', 'Data', 'Descrição', 'Tipo', 'Categoria',
  'Subcategoria', 'Conta', 'Forma pgto', 'Cartão', 'Valor', 'Status', 'Fixa?',
  'Parcelado?', 'Parcela nº', 'Total parcelas', 'Valor total compra', 'Grupo parcelamento',
  'Vencimento fatura', 'Mês ref', 'Ativo/Meta', 'Quantidade', 'Preço unitário',
  'Origem', 'Observações'];

/** Número da coluna (1-based) de cada campo do Banco de Dados. */
const C = {
  ID: 1, REG: 2, DATA: 3, DESC: 4, TIPO: 5, CAT: 6, SUB: 7, CONTA: 8, FORMA: 9,
  CARTAO: 10, VALOR: 11, STATUS: 12, FIXA: 13, PARC: 14, PARC_N: 15, PARC_TOT: 16,
  VALOR_TOTAL: 17, GRUPO: 18, VENC: 19, MESREF: 20, ATIVO: 21, QTD: 22, PRECO: 23,
  ORIGEM: 24, OBS: 25
};
const BD_NCOLS = BD_CABECALHO.length;

/** Intervalos nomeados do BD (usados nas fórmulas das outras abas). */
const BD_NOMES = {
  BD_ID: C.ID, BD_DATA: C.DATA, BD_DESC: C.DESC, BD_TIPO: C.TIPO, BD_CAT: C.CAT,
  BD_CONTA: C.CONTA, BD_FORMA: C.FORMA, BD_CARTAO: C.CARTAO, BD_VALOR: C.VALOR,
  BD_STATUS: C.STATUS, BD_PARCELADO: C.PARC, BD_PARCELA: C.PARC_N, BD_TOTPARC: C.PARC_TOT,
  BD_GRUPO: C.GRUPO, BD_VENC: C.VENC, BD_MESREF: C.MESREF, BD_ATIVO: C.ATIVO,
  BD_QTD: C.QTD, BD_PRECO: C.PRECO, BD_ORIGEM: C.ORIGEM
};

/* ---------- Configurações ---------- */

const CFG_LINHA_INI = 11;   // primeira linha de dados das tabelas
const CFG_LINHA_FIM = 60;   // última linha de dados das tabelas

/** Colunas (1-based) das tabelas da aba Configurações. */
const CFG = {
  CAT_NOME: 2, CAT_TIPO: 3, CAT_ORC: 4,                  // B, C, D
  SUB_CAT: 6, SUB_NOME: 7,                               // F, G
  CONTA: 9,                                              // I
  FORMA: 11,                                             // K
  CART_NOME: 13, CART_LIMITE: 14, CART_FECH: 15, CART_VENC: 16, CART_CONTA: 17, // M..Q
  FIXA_NOME: 19, FIXA_CAT: 20, FIXA_VALOR: 21, FIXA_DIA: 22, FIXA_CONTA: 23,
  FIXA_FORMA: 24, FIXA_CARTAO: 25,                       // S..Y
  ATV_NOME: 27, ATV_TIPO: 28, ATV_TICKER: 29, ATV_MANUAL: 30, ATV_PRECO: 31, // AA..AE
  META_NOME: 33, META_VALOR: 34, META_INICIAL: 35, META_PRAZO: 36,            // AG..AJ
  AUX_CATS: 38, AUX_SUBS: 39, AUX_ATIVOS: 40             // AL..AN (ocultas)
};
const CFG_DIA_INICIO_A1 = 'C5';
const CFG_ANO_INI_A1 = 'C6';

/* ---------- Formulário (aba Lançamentos) ---------- */

const FORM = {
  data: 'C5', desc: 'C6', valor: 'C7', conta: 'C8', forma: 'C9', fixa: 'C10',
  tipo: 'E5', cat: 'E6', sub: 'E7', status: 'E8', cartao: 'E9', obs: 'E10',
  parcelado: 'C13', valorTotal: 'C14', nParc: 'E13', parcIni: 'E14',
  ativo: 'C17', qtd: 'E17',
  btnRegistrar: 'E23', msg: 'B24', btnLimpar: 'E25'
};

/* ---------- Resumo ---------- */

const RES = {
  MES: 'N1', ANO: 'N2',
  CART_INI: 14, CART_MAX: 6,      // linhas 14..19, colunas B..H
  FIXA_INI: 14, FIXA_MAX: 12,     // linhas 14..25, colunas J..Q
  COL_CART_PAGA: 8,               // H
  COL_FIXA_VALOR: 13,             // M
  COL_FIXA_PAGA: 15               // O
};

/* ---------- Investimentos ---------- */

const INV = {
  HIST_LINHA_INI: 35,  // primeira linha de dados do histórico do patrimônio
  HIST_COL: 8          // H
};
