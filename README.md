# 💗 Planilha de Controle Financeiro Pessoal (Google Sheets)

Planilha integrada com 5 abas: **Resumo**, **Investimentos**, **Lançamentos**, **Configurações** e **Banco de Dados**.
Você registra cada movimentação **uma única vez** no formulário. O resto se atualiza sozinho:

```
Lançamento ➜ Banco de Dados ➜ Resumo ➜ Fatura do cartão ➜ Parcelas ➜ Saldo ➜ Investimentos/Metas
```

Tudo é montado por um script do **Google Apps Script**. Não precisa instalar nada no computador.

---

## 📁 Arquivos do projeto

| Arquivo | O que tem |
|---|---|
| `Config.gs` | Constantes compartilhadas: nomes das abas, colunas do Banco de Dados, posições do formulário e do Resumo, cores. |
| `Setup.gs` | Função `setup`, que monta as 5 abas, os intervalos nomeados, as fórmulas, os gráficos e os dados iniciais. |
| `Automacoes.gs` | Menu 💗 Finanças, gatilho `onEdit`, registro de lançamentos, parcelas, faturas, contas fixas e recálculos. |
| `Investimentos.gs` | Snapshot mensal do patrimônio (`registrarSnapshot`). |
| `Exemplo.gs` | Carrega e remove os dados de exemplo. |
| `appsscript.json` | Manifesto do projeto (fuso `America/Sao_Paulo`, runtime V8). |

---

## 🚀 Instalação (cerca de 5 minutos)

1. Abra o [Google Sheets](https://sheets.new) e crie uma **planilha em branco**. Dê um nome a ela, por exemplo "Minhas Finanças".
2. No menu, abra **Extensões → Apps Script**.
3. Apague o conteúdo do arquivo `Código.gs` que já vem aberto.
4. Crie os arquivos abaixo com **+ → Script** e cole o conteúdo de cada um. O nome precisa ser igual, e o `.gs` é colocado automaticamente:

   | Arquivo no Apps Script | Copiar de |
   |---|---|
   | `Config` | `Config.gs` |
   | `Automacoes` | `Automacoes.gs` |
   | `Setup` | `Setup.gs` |
   | `Investimentos` | `Investimentos.gs` |
   | `Exemplo` | `Exemplo.gs` |

   > Dá para colar tudo dentro do próprio `Código.gs` em vez de criar 5 arquivos. A ordem não importa.

5. Opcional: em **⚙️ Configurações do projeto**, marque *"Exibir arquivo de manifesto appsscript.json"* e cole o conteúdo de `appsscript.json`. Isso garante o fuso de São Paulo.
6. Clique em **💾 Salvar**. Na barra de cima, escolha a função **`setup`** e clique em **▶ Executar**.
7. O Google vai pedir autorização. Clique em **Revisar permissões**, escolha sua conta e depois **Avançado → Acessar (não seguro)** → **Permitir**.
   *O aviso aparece porque o script é seu e não foi publicado. Ele só acessa esta planilha.*
   *O `setup` também define a localidade da planilha como **Português (Brasil)**, o fuso como **São Paulo** e cria o gatilho que grava o patrimônio todo dia 1º.*
8. Volte para a planilha e **recarregue a página (F5)**. O menu **💗 Finanças** vai aparecer.
9. Quer ver tudo funcionando antes de usar de verdade? Clique em **💗 Finanças → 🧪 Carregar dados de exemplo**. Para tirar, use **🗑️ Remover dados de exemplo**.

### 🔄 Atualizando o script
Para instalar uma versão nova do código, cole os arquivos atualizados no Apps Script e rode **💗 Finanças → 🛠️ Reconstruir painéis** (ou a função `setup` de novo). Resumo, Investimentos e Lançamentos são recriados. **Configurações e Banco de Dados são mantidos.**

---

## 🧭 Como usar no dia a dia

### 📝 Registrar um lançamento (aba Lançamentos)
1. Preencha os campos. Os destacados em **rosa** são obrigatórios.
   - O **Tipo** filtra as **Categorias**, e a Categoria filtra as **Subcategorias**.
   - Na forma **Crédito**, escolha o cartão. A **Prévia** mostra a fatura em que a compra vai cair e o mês em que ela entra no Resumo.
   - Os campos que não se aplicam ficam **cinza**.
2. Marque a caixinha **✅ REGISTRAR LANÇAMENTO**. Também funciona no **app de celular**.
3. Uma mensagem verde confirma o registro, e o formulário é limpo.

### 🧾 Compras parceladas
- Escolha **Compra parcelada? = Sim** e informe o **valor TOTAL**, o **nº de parcelas** e a **parcela inicial**.
- Exemplo: Notebook de R$ 3.600 em 12x vira 12 lançamentos de R$ 300, cada um na sua fatura. No Resumo aparece como `Notebook | 4/12 | R$ 300`.
- **Compra que já estava em andamento:** se você está na 4ª de 12, use *parcela inicial = 4*. A parcela 4 cai na fatura da data informada, e as seguintes vão para os meses seguintes.

### 💳 Cartões e faturas
- Cada cartão tem seu **fechamento** e **vencimento** na aba Configurações.
- Compra **no dia do fechamento ou depois** vai para a fatura seguinte.
- A compra no crédito conta como gasto **no mês em que a fatura vence**. Assim o saldo do mês bate com o dinheiro que realmente sai da conta.
- No Resumo, marque **✔ Paga?** no cartão quando pagar a fatura. Duas coisas acontecem:
  - as compras daquela fatura passam para "Pago" e o limite é liberado;
  - é criada uma **Transferência** "Pagamento de fatura", que **não** conta como gasto de novo.
- Quando **todas** as faturas do mês estão pagas, o card **Fatura do cartão** passa a mostrar a soma das faturas **em aberto** (as que vencem no mês seguinte), para você acompanhar o que já está gastando.

### ✅ Contas fixas
- O checklist do Resumo lista as contas fixas cadastradas em Configurações.
- Ajuste o **Valor** do mês se precisar (a energia, por exemplo) e marque **✔**. O lançamento é criado automaticamente no Banco de Dados e o saldo é recalculado.
- Desmarcar remove o lançamento.
- O contador mostra *"Contas fixas: 4 de 6 pagas"* e *"R$ X ainda pendentes"*.

### 📅 Mês financeiro (ciclo 26 → 25)
- No Resumo, escolha o **Mês** e o **Ano**. Com o ciclo 26, "Outubro" vai de **26/09 a 25/10**.
- Para voltar ao mês do calendário, mude *Dia de início do mês financeiro* para **1** em Configurações e rode **💗 Finanças → 🧮 Recalcular**.

### 📈 Investimentos e metas
- Registre um aporte como **Tipo = Investimento** e escolha o **ativo** (Bitcoin, MXRF11…) ou a **meta** (🚗 Carro, ✈️ Viagem…). Se quiser, informe a quantidade.
- A **carteira** calcula quantidade, preço médio, valor atual, rentabilidade e % de cada ativo.
- A cotação vem do `GOOGLEFINANCE`. Se algum ativo não tiver cotação (alguns FIIs), preencha o **Preço manual** em Configurações.
- As **metas** mostram o acumulado, o quanto falta, a barra de progresso e **quanto guardar por mês** até o prazo.
- A **evolução do patrimônio** é gravada automaticamente todo dia 1º. Para gravar na hora, use *💗 Finanças → 📸 Registrar snapshot*.

---

## ⚙️ Personalizando (aba Configurações)

Para adicionar um item, escreva na próxima linha vazia da tabela. Para remover, apague o conteúdo da linha.

| Tabela | Para que serve |
|---|---|
| **Categorias** (+ Tipo + Orçamento) | Menus do formulário. O orçamento aparece como barra de uso no Resumo. |
| **Subcategorias** | Menu dependente da categoria. |
| **Contas** / **Formas de pagamento** | Menus do formulário. |
| **Cartões** | Nome, limite, dia de fechamento, dia de vencimento e conta de pagamento. |
| **Contas fixas** | Checklist mensal do Resumo (valor previsto, dia, conta e forma). |
| **Ativos** | Carteira de investimentos (ticker do GOOGLEFINANCE ou preço manual). |
| **Metas** | Valor, quanto já tinha guardado antes da planilha e prazo. |

> Os cartões vêm com valores de exemplo: **Cartão 1** fecha dia 25 e vence dia 03, e **Cartão 2** fecha dia 03 e vence dia 10. Ajuste para os seus cartões e, se já houver lançamentos, rode **🧮 Recalcular faturas e meses de referência**.

---

## 📊 O que tem no Resumo

- **8 indicadores:** Entradas, Saídas, Saldo do mês, Valor investido (com taxa de investimento), Fatura do cartão, Parcelas futuras, Contas fixas pagas e Saldo previsto.
- **Cartões:** fatura do mês, status (Aberta/Fechada/Paga), barra de uso do limite e limite disponível.
- **Contas fixas:** checklist com valor, vencimento e status (Paga, Pendente ou Atrasada).
- **Gastos por categoria:** valor, % dos gastos, % da renda, orçamento e barra de uso do orçamento.
- **Parcelas do mês:** parcela k/N, parcelas restantes e quanto falta pagar.
- **Despesas futuras:** quanto já está comprometido nos próximos 6 meses (parcelas + fixas).
- **Evolução de 12 meses** e **maiores gastos do mês**.
- **Gráficos:** gastos por categoria · destino da renda · despesas futuras · evolução dos gastos · receitas × despesas × investimentos.
- **Indicadores extras:** média de gastos, maior categoria, taxa de poupança, gasto médio por dia e uso total dos limites.

---

## 🧰 Menu 💗 Finanças

| Item | O que faz |
|---|---|
| ✅ Registrar lançamento | Mesmo efeito da caixinha do formulário |
| 🧹 Limpar formulário | Volta o formulário ao padrão |
| 🔄 Atualizar Resumo | Redesenha os checklists de cartões e contas fixas |
| 🧮 Recalcular faturas e meses de referência | Use depois de mudar fechamento/vencimento ou o ciclo |
| 📸 Registrar snapshot do patrimônio | Grava o patrimônio do mês no histórico |
| 🧪 / 🗑️ Dados de exemplo | Carrega ou remove os dados de teste |
| ⚠️ Apagar TODOS os lançamentos | Zera o Banco de Dados (pede confirmação) |
| 🛠️ Reconstruir painéis | Recria Resumo, Investimentos e Lançamentos **sem apagar** Configurações nem o Banco de Dados |

---

## ❓ Dúvidas comuns

- **O menu 💗 Finanças não apareceu.** Recarregue a página. Se ainda não aparecer, rode `onOpen` uma vez pelo editor do Apps Script.
- **Marquei a caixinha e nada aconteceu.** Confirme que o `setup` foi executado e autorizado. A mensagem de erro, se houver, aparece logo abaixo do botão.
- **Corrigir ou apagar um lançamento.** Edite ou apague a linha direto na aba **Banco de Dados**. Ela mostra um aviso antes de editar, só para evitar acidentes. Numa compra parcelada, use o filtro pela coluna *Grupo parcelamento* para achar todas as parcelas.
- **Transferências** (entre contas, pagamento de fatura) não entram em entradas nem em saídas.
- **As fórmulas mostram erro ou o `setup` parou com "DIAGNÓSTICO".** O `setup` testa se a planilha entende as fórmulas. Se aparecer essa mensagem, confira se a localidade está em *Arquivo → Configurações → Português (Brasil)* e rode o `setup` de novo.
- **Um ativo aparece como "sem cotação".** O `GOOGLEFINANCE` não tem esse ticker. Preencha o **Preço manual** na tabela Ativos em Configurações.
