/************************************************************
 * MAISBOOK → GERAL ENVIOS
 *
 * I = código do MaisBook (é ele que vai para a API)
 * B = data de escolha, só se a célula estiver vazia; em Finalizado nunca é mexida
 * C = nome do cliente, só se a célula estiver vazia
 * D = contrato (0001/0964 vira 964, como a coluna já está)
 * F = pacote do MaisBook, no texto da lista da planilha
 * G = curso, só se a célula estiver vazia
 *
 * A coluna L é LOTE. Ela não trava o status.
 * Finalizado não volta para Aguardando Aprovação.
 * Aguardando Aprovação ↔ Com Modificações pode ir e voltar.
 * IMPRESSÃO só vira Finalizado quando status_selecao = finalizado.
 ************************************************************/

var CONFIG = {
  ABA: "GERAL ENVIOS",
  PRIMEIRA_LINHA: 19,
  LIMITE_POR_EXECUCAO:1,
  INTERVALO_MS: 300,

  COLUNA_DATA: 2,            // B
  COLUNA_NOME: 3,            // C
  COLUNA_CONTRATO: 4,        // D
  COLUNA_SGE: 5,             // E
  COLUNA_PLANO: 6,           // F
  COLUNA_CURSO: 7,           // G
  COLUNA_CODIGO_MAISBOOK: 9, // I
  COLUNA_STATUS: 11,         // K
  COLUNA_PRAZO_ENVIO: 12,    // O = B + 17
  COLUNA_ENVIO: 16,          // P
  COLUNA_APROVACAO: 17,      // Q
  COLUNA_PRAZO_ENTREGA: 18   // R = Q + 60
};

var API_URL = "https://maisbook.me/wp-json/maisbook/v1/clientes/";
var CHAVE_API = "MAISBOOK_API_KEY";
var CHAVE_PROXIMA_LINHA = "PROXIMA_LINHA";

var ROTULO = [
  "Sem álbum",
  "AGUARDANDO FOTO",
  "EM DIAGRAMAÇÃO",
  "DIAGRAMADO",
  "Aguardando Aprovação",
  "Com Modificações",
  "Aprovado",
  "IMPRESSÃO",
  "Finalizado"
];

var SEM_ALBUM = 0;
var AGUARDANDO_FOTO = 1;
var AGUARDANDO = 4;
var MODIFICACOES = 5;
var APROVADO = 6;
var IMPRESSAO = 7;
var FINALIZADO = 8;
var FORA = -2;


function salvarChaveAPI() {
  var chave = "COLE_SUA_CHAVE_API_AQUI";

  if (!chave || chave === "COLE_SUA_CHAVE_API_AQUI") {
    throw new Error("Insira sua chave da API.");
  }

  PropertiesService.getScriptProperties().setProperty(CHAVE_API, chave);
  Logger.log("Chave salva.");
}


function configurarAutomacao() {
  var funcoes = ["aoEditarMaisBook", "atualizarLote30"];

  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (funcoes.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.newTrigger("aoEditarMaisBook")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger("atualizarLote30")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("Automação configurada.");
}


function atualizarLote30() {
  processarFila({
    limite: CONFIG.LIMITE_POR_EXECUCAO,
    tempoMaxMs: 4 * 60 * 1000,
    circular: true
  });
}


function atualizarAgora() {
  processarFila({
    limite: 500,
    tempoMaxMs: 5 * 60 * 1000,
    circular: false,
    desdeOInicio: true
  });
  Logger.log("Atualização manual encerrada.");
}


function reiniciarFila() {
  PropertiesService.getScriptProperties()
    .setProperty(CHAVE_PROXIMA_LINHA, String(CONFIG.PRIMEIRA_LINHA));
  Logger.log("Fila reiniciada na linha " + CONFIG.PRIMEIRA_LINHA);
}


function prepararPrazos() {
  var aba = obterAba();
  var ultima = aba.getLastRow();
  var trocadas = 0;

  for (var linha = CONFIG.PRIMEIRA_LINHA; linha <= ultima; linha++) {
    if (corrigirPrazo(aba, linha)) trocadas++;
  }

  Logger.log("Prazos ajustados: " + trocadas);
}


function testarRetornoAPI() {
  var aba = obterAba();
  var codigo = codigoDaLinha(aba, CONFIG.PRIMEIRA_LINHA);

  if (!codigo) throw new Error("Não há código do MaisBook (coluna I) na linha " + CONFIG.PRIMEIRA_LINHA);

  var resposta = buscarCliente(codigo);
  Logger.log("Código MaisBook: " + codigo + " | HTTP: " + resposta.codigo);
  Logger.log(resposta.texto);
}


function aoEditarMaisBook(e) {
  if (!e || !e.range) return;

  var aba = e.range.getSheet();
  if (aba.getName() !== CONFIG.ABA) return;

  var inicioColuna = e.range.getColumn();
  var fimColuna = inicioColuna + e.range.getNumColumns() - 1;
  if (inicioColuna > CONFIG.COLUNA_CODIGO_MAISBOOK || fimColuna < CONFIG.COLUNA_CODIGO_MAISBOOK) return;

  var linha = Math.max(e.range.getRow(), CONFIG.PRIMEIRA_LINHA);
  var fim = e.range.getRow() + e.range.getNumRows() - 1;
  var feitos = 0;

  for (; linha <= fim && feitos < 15; linha++) {
    var codigo = codigoDaLinha(aba, linha);
    if (!codigo) continue;
    consultarAluno(aba, linha, codigo);
    feitos++;
    Utilities.sleep(CONFIG.INTERVALO_MS);
  }
}


function processarFila(opcoes) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log("Outra execução está em andamento.");
    return;
  }

  try {
    var aba = obterAba();
    var ultima = aba.getLastRow();
    if (ultima < CONFIG.PRIMEIRA_LINHA) return;

    var props = PropertiesService.getScriptProperties();
    var linha = opcoes.desdeOInicio
      ? CONFIG.PRIMEIRA_LINHA
      : Number(props.getProperty(CHAVE_PROXIMA_LINHA));

    if (!linha || linha < CONFIG.PRIMEIRA_LINHA || linha > ultima) {
      linha = CONFIG.PRIMEIRA_LINHA;
    }

    var consultados = 0;
    var verificadas = 0;
    var total = ultima - CONFIG.PRIMEIRA_LINHA + 1;
    var inicio = Date.now();

    while (consultados < opcoes.limite && verificadas < total) {
      if (opcoes.tempoMaxMs && Date.now() - inicio > opcoes.tempoMaxMs) break;

      var codigo = codigoDaLinha(aba, linha);
      corrigirPrazo(aba, linha);

      if (codigo) {
        consultarAluno(aba, linha, codigo);
        consultados++;
        Utilities.sleep(CONFIG.INTERVALO_MS);
      }

      linha++;
      verificadas++;

      if (linha > ultima) {
        linha = CONFIG.PRIMEIRA_LINHA;
        if (!opcoes.circular) break;
      }
    }

    props.setProperty(CHAVE_PROXIMA_LINHA, String(linha));
    Logger.log("Consultas: " + consultados + ". Próxima linha: " + linha);
  } finally {
    lock.releaseLock();
  }
}


function codigoDaLinha(aba, linha) {
  return limpar(aba.getRange(linha, CONFIG.COLUNA_CODIGO_MAISBOOK).getDisplayValue());
}


function consultarAluno(aba, linha, codigo) {
  var resposta = buscarCliente(codigo);

  if (resposta.codigo !== 200) {
    Logger.log("HTTP " + resposta.codigo + " | Código MaisBook " + codigo + " | " + resposta.texto);
    return;
  }

  var dados;
  try {
    dados = JSON.parse(resposta.texto);
  } catch (erro) {
    Logger.log("JSON inválido no código MaisBook " + codigo);
    return;
  }

  if (!dados || dados.ok === false) return;
  aplicarNaLinha(aba, linha, dados);
}


function buscarCliente(codigo) {
  var chave = PropertiesService.getScriptProperties().getProperty(CHAVE_API);
  if (!chave) throw new Error("Chave da API não encontrada.");

  var resposta = UrlFetchApp.fetch(API_URL + encodeURIComponent(limpar(codigo)), {
    method: "get",
    headers: { "X-Maisbook-Api-Key": chave },
    muteHttpExceptions: true
  });

  return {
    codigo: resposta.getResponseCode(),
    texto: resposta.getContentText()
  };
}


function aplicarNaLinha(aba, linha, dados) {
  var valores = aba.getRange(linha, 1, 1, CONFIG.COLUNA_PRAZO_ENTREGA).getDisplayValues()[0];
  var formulas = aba.getRange(linha, 1, 1, CONFIG.COLUNA_PRAZO_ENTREGA).getFormulas()[0];
  var info = destinoDaApi(dados);
  var status = decidirStatus(valores[CONFIG.COLUNA_STATUS - 1], info.destino, info.selecao);

  if (status) {
    aba.getRange(linha, CONFIG.COLUNA_STATUS).setValue(status);
    Logger.log("Linha " + linha + ": " + status);
  }

  var contrato = contratoCurto(dados.numero_contrato);
  var contratoAtual = limpar(valores[CONFIG.COLUNA_CONTRATO - 1]);
  if (
    contrato &&
    contratoAtual !== contrato &&
    !limpar(formulas[CONFIG.COLUNA_CONTRATO - 1])
  ) {
    var celulaContrato = aba.getRange(linha, CONFIG.COLUNA_CONTRATO);
    celulaContrato.setValue(/^\d+$/.test(contrato) ? Number(contrato) : contrato);
  }

  var plano = extrairPlano(dados);
  var planoAtual = limpar(valores[CONFIG.COLUNA_PLANO - 1]);
  if (
    plano &&
    normalizar(planoAtual) !== normalizar(plano) &&
    !limpar(formulas[CONFIG.COLUNA_PLANO - 1])
  ) {
    aba.getRange(linha, CONFIG.COLUNA_PLANO).setValue(plano);
  }

  var nome = limpar(dados.nome_completo);
  if (nome && celulaVazia(valores[CONFIG.COLUNA_NOME - 1], formulas[CONFIG.COLUNA_NOME - 1])) {
    aba.getRange(linha, CONFIG.COLUNA_NOME).setValue(nome);
  }

  var curso = extrairCurso(dados);
  if (curso && celulaVazia(valores[CONFIG.COLUNA_CURSO - 1], formulas[CONFIG.COLUNA_CURSO - 1])) {
    aba.getRange(linha, CONFIG.COLUNA_CURSO).setValue(curso);
  }

  // na seleção finalizada a API manda a data da finalização, não a da escolha
  var finalizado =
    indiceDaEtapa(status || valores[CONFIG.COLUNA_STATUS - 1]) === FINALIZADO ||
    info.selecao === "finalizado";
  var dataSelecao = converterData(dados.status_selecao && dados.status_selecao.data);
  if (
    !finalizado &&
    dataSelecao &&
    celulaVazia(valores[CONFIG.COLUNA_DATA - 1], formulas[CONFIG.COLUNA_DATA - 1])
  ) {
    gravarData(aba, linha, CONFIG.COLUNA_DATA, dataSelecao);
  }

  var dataAlbum = converterData(dados.status_album && dados.status_album.data);

  if (
    info.album === "disponivel" &&
    dataAlbum &&
    celulaVazia(valores[CONFIG.COLUNA_ENVIO - 1], formulas[CONFIG.COLUNA_ENVIO - 1])
  ) {
    gravarData(aba, linha, CONFIG.COLUNA_ENVIO, dataAlbum);
  }

  if (
    info.album === "aprovado" &&
    dataAlbum &&
    celulaVazia(valores[CONFIG.COLUNA_APROVACAO - 1], formulas[CONFIG.COLUNA_APROVACAO - 1])
  ) {
    gravarData(aba, linha, CONFIG.COLUNA_APROVACAO, dataAlbum);
  }
}


function destinoDaApi(dados) {
  var album = normalizar(codigoDe(dados, "status_album"));
  var selecao = normalizar(codigoDe(dados, "status_selecao"));
  var destino = -1;

  if (album === "sem_album" || album === "pendente") destino = SEM_ALBUM;
  else if (album === "disponivel") destino = AGUARDANDO;
  else if (album === "com_comentarios") destino = MODIFICACOES;
  else if (album === "aprovado") destino = APROVADO;

  if (
    (selecao === "pendente" || selecao === "selecao_incompleta") &&
    (destino < 0 || destino <= AGUARDANDO_FOTO)
  ) {
    destino = AGUARDANDO_FOTO;
  }

  return { destino: destino, selecao: selecao, album: album };
}


function decidirStatus(statusAtual, destino, codigoSelecao) {
  var indice = indiceDaEtapa(statusAtual);

  if (indice === FORA || indice === FINALIZADO) return "";

  if (indice === IMPRESSAO) {
    return codigoSelecao === "finalizado" ? ROTULO[FINALIZADO] : "";
  }

  if (destino < 0 || destino === indice || !podeGravar(destino)) return "";

  var trocaDeAjuste =
    (indice === AGUARDANDO || indice === MODIFICACOES) &&
    (destino === AGUARDANDO || destino === MODIFICACOES);

  if (trocaDeAjuste) return ROTULO[destino];
  if (indice >= APROVADO && destino < indice) return "";
  if (indice < 0 || destino > indice) return ROTULO[destino];

  return "";
}


function podeGravar(indice) {
  return (
    indice === SEM_ALBUM ||
    indice === AGUARDANDO_FOTO ||
    indice === AGUARDANDO ||
    indice === MODIFICACOES ||
    indice === APROVADO
  );
}


function indiceDaEtapa(texto) {
  var chave = normalizar(texto);
  if (!chave) return -1;

  for (var i = 0; i < ROTULO.length; i++) {
    if (normalizar(ROTULO[i]) === chave) return i;
  }

  return FORA;
}


function contratoCurto(valor) {
  var texto = limpar(valor);
  if (!texto) return "";

  var parte = texto.split("/").pop();
  parte = parte.replace(/^0+/, "");
  return parte || texto;
}


var PLANOS_PLANILHA = {
  studiodaymoscou: "Studio Day Moscou 30x20",
  studiodayberlim: "Studio Day Berlim 24x34",
  essencialbook: "Essencial Book 20x20",
  bali: "Bali 20x20",
  primebook: "Prime Book 30X20",
  diamondbook: "Diamond Book 24X34",
  moscou: "Moscou 30X20",
  pacoteberlim: "Pacote Berlim 24x34"
};


function extrairPlano(dados) {
  var pacote = dados && dados.pacote;
  if (!pacote) return "";

  if (typeof pacote === "string") {
    return PLANOS_PLANILHA[chavePlano(pacote)] || "";
  }

  var peloLabel = PLANOS_PLANILHA[chavePlano(pacote.label)];
  if (peloLabel) return peloLabel;
  return PLANOS_PLANILHA[chavePlano(pacote.codigo)] || "";
}


function chavePlano(texto) {
  return normalizar(texto)
    .replace(/[^a-z0-9]/g, "")
    .replace(/(30x20|24x34|20x20)$/, "");
}


function corrigirPrazo(aba, linha) {
  var envio = corrigirFormula(aba, linha, CONFIG.COLUNA_PRAZO_ENVIO, "B", 17);
  var entrega = corrigirFormula(aba, linha, CONFIG.COLUNA_PRAZO_ENTREGA, "Q", 60);
  return envio || entrega;
}


function corrigirFormula(aba, linha, coluna, origem, dias) {
  var celula = aba.getRange(linha, coluna);
  var atual = String(celula.getFormula() || "").replace(/\s+/g, "").toUpperCase();
  var simples = "=" + origem + linha + "+" + dias;

  if (atual !== simples) return false;

  celula.setFormula(
    '=SE(ÉNÚMERO(' + origem + linha + ");" + origem + linha + "+" + dias + ';"" )'
  );
  return true;
}


function extrairCurso(dados) {
  var valor = dados && dados.curso;
  if (typeof valor === "string") return limpar(valor);
  if (valor && typeof valor.nome === "string") return limpar(valor.nome);
  return "";
}


function codigoDe(dados, campo) {
  var bloco = dados && dados[campo];
  if (!bloco || typeof bloco !== "object") return "";
  return bloco.codigo || "";
}


function gravarData(aba, linha, coluna, data) {
  var celula = aba.getRange(linha, coluna);
  celula.setValue(data);
  celula.setNumberFormat("dd/MM/yyyy");
}


function celulaVazia(valor, formula) {
  return !limpar(valor) && !limpar(formula);
}


function limpar(texto) {
  return String(texto || "")
    .replace(/[\u200B\u00A0]/g, "")
    .trim();
}


function normalizar(texto) {
  return limpar(texto)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}


function converterData(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return isNaN(valor.getTime()) ? null : valor;

  var texto = limpar(valor);
  if (!texto) return null;

  var brasileira = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (brasileira) {
    return dataValida(Number(brasileira[3]), Number(brasileira[2]), Number(brasileira[1]));
  }

  var iso = texto.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return dataValida(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  }

  var data = new Date(texto);
  return isNaN(data.getTime()) ? null : data;
}


function dataValida(ano, mes, dia) {
  var data = new Date(ano, mes - 1, dia);
  if (
    data.getFullYear() !== ano ||
    data.getMonth() !== mes - 1 ||
    data.getDate() !== dia
  ) {
    return null;
  }
  return data;
}


function obterAba() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA);
  if (!aba) throw new Error("Aba não encontrada: " + CONFIG.ABA);
  return aba;
}
