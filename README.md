// Cole este código no Apps Script da planilha "Controle_Leads_CEPGAS"
// (depois de abri-la/convertê-la no Google Sheets).
// Ele grava tanto o formulário de entrada quanto o resultado do diagnóstico
// na mesma aba "Leads CEPGAS", usando o e-mail do lead para casar as duas informações.

var ABA = "Leads CEPGAS";
var LINHA_CABECALHO = 2;   // linha com os nomes das colunas (a linha 1 é o título dos grupos)
var LINHA_INICIO_DADOS = 3;

// Colunas (conforme o cabeçalho atual da planilha)
var COL_NOME = 1;              // A - Nome do Lead
var COL_EMPRESA = 2;           // B - Empresa
var COL_CONTATO = 5;           // E - Contato
var COL_PERFIL = 6;            // F - Perfil Comportamental
var COL_DIMENSAO = 7;          // G - Dimensão Predominante
var COL_DATA_DIAGNOSTICO = 8;  // H - Data do Diagnóstico
var COL_ULTIMO_CONTATO = 11;   // K - Última Data de Contato

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA);
  var dados = JSON.parse(e.postData.contents);

  if (dados.action === "lead") {
    registrarLead(sheet, dados);
  } else if (dados.action === "resultado") {
    registrarResultado(sheet, dados);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function registrarLead(sheet, dados) {
  var linha = sheet.getLastRow() + 1;
  if (linha < LINHA_INICIO_DADOS) linha = LINHA_INICIO_DADOS;

  var contato = (dados.telefone || "") + (dados.email ? " / " + dados.email : "");

  sheet.getRange(linha, COL_NOME).setValue(dados.nome || "");
  sheet.getRange(linha, COL_EMPRESA).setValue(dados.empresa || "");
  sheet.getRange(linha, COL_CONTATO).setValue(contato);
  sheet.getRange(linha, COL_ULTIMO_CONTATO).setValue(new Date());
}

function registrarResultado(sheet, dados) {
  var linha = encontrarLinhaPorEmail(sheet, dados.email);

  if (linha === -1) {
    // Não achou o lead (ex: alguém abriu o diagnóstico direto, sem passar pelo formulário)
    linha = sheet.getLastRow() + 1;
    if (linha < LINHA_INICIO_DADOS) linha = LINHA_INICIO_DADOS;
    sheet.getRange(linha, COL_NOME).setValue(dados.nome || "");
    sheet.getRange(linha, COL_CONTATO).setValue(dados.email || "");
  }

  sheet.getRange(linha, COL_PERFIL).setValue(dados.perfil || "");
  sheet.getRange(linha, COL_DIMENSAO).setValue(dados.dimensao || "");
  sheet.getRange(linha, COL_DATA_DIAGNOSTICO).setValue(new Date());
}

function encontrarLinhaPorEmail(sheet, email) {
  if (!email) return -1;
  var ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < LINHA_INICIO_DADOS) return -1;

  var valores = sheet.getRange(LINHA_INICIO_DADOS, COL_CONTATO, ultimaLinha - LINHA_INICIO_DADOS + 1, 1).getValues();
  for (var i = 0; i < valores.length; i++) {
    var contato = String(valores[i][0] || "");
    if (contato.indexOf(email) !== -1) {
      return LINHA_INICIO_DADOS + i;
    }
  }
  return -1;
}
