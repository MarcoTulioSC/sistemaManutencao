const ID_PLANILHA = "1UAsgzfc3PAZdKPU4YyKasVV6A5GsqlMSbc0arfbBwqE";
const ABA = "Solicitação Manutenção";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("indexFichas");
}

/* =====================================
   LINHA ATUAL = COLUNA C VAZIA
===================================== */
function obterLinhaAtual() {
  const sh = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(ABA);
  const lastRow = sh.getLastRow();

  // Se só tem cabeçalho
  if (lastRow < 2) return 2;

  // 🔹 lê colunas C (SR) e I (Matrícula)
  const dados = sh.getRange(2, 3, lastRow - 1, 7).getValues();
  // índices: C=0, I=6

  // 🔹 percorre de baixo para cima
  for (let i = dados.length - 1; i >= 0; i--) {
    const codigoSR = String(dados[i][0]).trim();
    const matricula = String(dados[i][6]).trim();

    if (codigoSR || matricula) {
      return i + 3; // linha seguinte à última usada
    }
  }

  // nenhuma linha usada
  return 2;
}

/* =====================================
   CRIAR LINHA INICIAL
===================================== */
function criarLinhaInicial(data) {
  const sh = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(ABA);

  const linha = obterLinhaAtual();
  const codigo = linha - 2; // 🔹 CÓDIGO = LINHA - 2

  sh.getRange(linha, 1).setValue(codigo);        // A
  sh.getRange(linha, 2).setValue("Não enviada"); // B
  sh.getRange(linha, 8).setValue(data);          // H

  return { linha, codigo };
}

/* =====================================
   FINALIZAR SOLICITAÇÃO
===================================== */
function finalizarSolicitacao(linha, d) {
  const sh = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(ABA);

  if (sh.getRange(linha, 3).getValue()) {
    throw new Error("Linha já utilizada.");
  }

  sh.getRange(linha, 2).setValue("Enviada");        // B
  sh.getRange(linha, 3).setValue(d.codigoSR);      // C (SR)
  sh.getRange(linha, 4).setValue(d.veiculo);       // D
  sh.getRange(linha, 5).setValue(d.garagem);       // E
  sh.getRange(linha, 6).setValue(d.descricao);     // F
  sh.getRange(linha, 7).setValue(d.tipoProblema);  // G
  sh.getRange(linha, 9).setValue(d.matricula);     // I
  sh.getRange(linha, 8).setValue(d.data); // H → DATA
}

/* =====================================
   LISTAS
===================================== */
function obterListas() {
  const sh = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(ABA);
  const lr = sh.getLastRow();

  return {
    garagens: [...new Set(sh.getRange(2,5,lr-1,1).getValues().flat().filter(String))],
    tipos: [...new Set(sh.getRange(2,7,lr-1,1).getValues().flat().filter(String))]
  };
}

/* =====================================
   MATRÍCULA → NOME
===================================== */
function obterNomePorMatricula(m) {
  const sh = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(ABA);
  const dados = sh.getRange(2,9,sh.getLastRow()-1,2).getValues();

  for (let r of dados) {
    if (String(r[0]).trim() === String(m).trim()) {
      return r[1];
    }
  }
  return "";
}

/* =====================================
   VEÍCULO → GARAGEM
===================================== */
function obterGaragemPorVeiculo(veiculo) {
  if (!veiculo) return "";

  const v = String(veiculo).replace(/\D/g, "").replace(/^0+/, "");
  if (!v) return "";

  const sh = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(ABA);
  const lastRow = sh.getLastRow();

  const dados = sh.getRange(2, 4, lastRow - 1, 2).getValues(); // D e E

  for (let r of dados) {
    const vPlan = String(r[0]).replace(/\D/g, "").replace(/^0+/, "");
    if (vPlan === v) {
      return r[1]; // garagem
    }
  }
  return "";
}
