/**
 * ============================================================
 *  📸 INVESTIMENTOS — histórico mensal do patrimônio.
 *  Um gatilho (criado pelo setup) roda todo dia 1º às 9h e grava
 *  o total investido e o patrimônio do mês. Também pode ser
 *  executado pelo menu 💗 Finanças → Registrar snapshot.
 * ============================================================
 */

function registrarSnapshot() {
  const ss = planilha_();
  const sh = ss.getSheetByName(ABA.INV);
  if (!sh) return;
  SpreadsheetApp.flush();

  const investido = Number(sh.getRange('B5').getValue()) || 0;
  const patrimonio = Number(sh.getRange('F5').getValue()) || 0;
  const hoje = new Date();
  const mes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

  const ini = INV.HIST_LINHA_INI;
  const vals = sh.getRange(ini, INV.HIST_COL, 400 - ini, 1).getValues();
  let alvo = vals.length - 1;
  for (let i = 0; i < vals.length; i++) {
    const d = vals[i][0];
    if (d === '' || mesmoMes_(d, mes)) {
      alvo = i;
      break;
    }
  }
  sh.getRange(ini + alvo, INV.HIST_COL, 1, 3).setValues([[mes, investido, patrimonio]]);

  try {
    ss.toast('Patrimônio de ' + fmtMesAno_(mes) + ': ' + fmtMoeda_(patrimonio), '📸 Snapshot registrado', 4);
  } catch (e) {
    // executando por gatilho, sem interface aberta
  }
}
