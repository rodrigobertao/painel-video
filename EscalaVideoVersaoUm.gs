/**
 * Script Google Apps Script (GS) para gerar o painel de eventos.
 * Versão final com correções mínimas:
 * - Observação agora é por idAgenda + colaborador
 * - Mantidas todas as demais funcionalidades
 */

const COL_ID_AGENDA = 0;
const COL_DATA_EVENTO = 1;
const COL_LOCAL = 2;
const COL_HORA_PREVISTA = 3;
const COL_NOME_EVENTO = 4;
const COL_TIPO_EVENTO = 5;
const COL_SERVICO_VIDEO = 8;
const COL_COLABORADOR_VIDEO = 9;
const COL_INDICACAO_CANCELAMENTO = 10;
const COL_HORA_LANCAMENTO_CANCELADO = 11;
const COL_HORA_INICIO_EVENTO = 12;
const COL_HORA_TERMINO_EVENTO = 13;
const COL_HABILIT_HORA_EVENTO = 14;

const COL_DTV_ID_AGENDA = 0;
const COL_DTV_COLABORADOR = 2;

function normalizeName(s) {
  if (s === null || s === undefined) return "";
  try {
    return String(s)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
  } catch (e) {
    return String(s).replace(/\s+/g, " ").trim().toUpperCase();
  }
}

function doGet(e) {
  try {
    // 🔥 Se for pedido por JSON (usado no GitHub Pages)
    if (e && e.parameter && e.parameter.action === 'getEventos') {
      const dados = getEventosParaPainel();
      return ContentService
        .createTextOutput(JSON.stringify(dados))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 🔥 Se for abrir o painel (HTML normal dentro do Apps Script)
    const eventos = gerarPainel();
    const template = HtmlService.createTemplateFromFile('template');
    template.eventos = eventos;
    return template.evaluate()
      .setTitle('Tabela de Escala do Vídeo')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>Erro ao carregar o Painel</h1><p>' + error.toString() + '</p>');
  }
}


function getEventosParaPainel() {
  return gerarPainel();
}

function gerarPainel() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sheetDados = spreadsheet.getSheetByName('Dados');
  if (!sheetDados) throw new Error('Planilha "Dados" não encontrada.');

  const lastRowDados = sheetDados.getLastRow();
  const dados = (lastRowDados > 1)
    ? sheetDados.getRange(2, 1, lastRowDados - 1, sheetDados.getLastColumn()).getValues()
    : [];

  // ---- DTV ----
  const sheetDTV = spreadsheet.getSheetByName('DTV');
  let fichasEnviadas = new Map();
  if (sheetDTV) {
    const totalRows = sheetDTV.getLastRow();
    let dadosDTV = [];
    if (totalRows > 1) {
      dadosDTV = sheetDTV.getRange(2, 1, totalRows - 1, sheetDTV.getLastColumn()).getValues();
    }
    fichasEnviadas = criarMapaFichas(dadosDTV);
  }

  // ---- OBSERVAÇÃO (idAgenda + colaborador) ----
  const sheetObs = spreadsheet.getSheetByName('Observacao');
  let obsMap = new Map();

  if (sheetObs) {
    const totalObsRows = sheetObs.getLastRow();
    if (totalObsRows > 1) {
      const obsData = sheetObs.getRange(2, 1, totalObsRows - 1, sheetObs.getLastColumn()).getValues();

      obsData.forEach(r => {
        const idObs = r[0];
        const colab = r[2];
        const obsTxt = r[3];

        if (idObs && colab && obsTxt) {
          const chave = `${String(idObs).trim()}|${normalizeName(colab)}`;
          obsMap.set(chave, obsTxt);
        }
      });
    }
  }

  // ---- PROCESSAR EVENTOS ----
  let eventosProcessados = processarEFiltrarEventos(dados, fichasEnviadas, obsMap);

  // ---- ORDENAR ----
  eventosProcessados.sort((a, b) => a.colaborador.localeCompare(b.colaborador));

  // ---- REMOVER REPETIÇÃO DE NOME ----
  let ultimo = "";
  eventosProcessados = eventosProcessados.map(ev => {
    const atual = normalizeName(ev.colaborador);
    if (atual === ultimo) return { ...ev, colaborador: "" };
    ultimo = atual;
    return ev;
  });

  // ---- SALVAR ----
  const sheetDestino = spreadsheet.getSheetByName('escalaVideo') || spreadsheet.insertSheet('escalaVideo');
  sheetDestino.clearContents();
  sheetDestino.appendRow(['colaborador', 'Status', 'local', 'horaPrevista', 'servico', 'tipoEvento', 'nomeEvento']);

  if (eventosProcessados.length > 0) {
    const dadosParaPlanilha = eventosProcessados.map(ev => [
      ev.colaborador,
      ev.status,
      ev.local,
      ev.horaPrevistaFormatada,
      ev.servico,
      ev.tipoEvento,
      ev.nomeEvento
    ]);
    sheetDestino.getRange(2, 1, dadosParaPlanilha.length, dadosParaPlanilha[0].length).setValues(dadosParaPlanilha);
  }

  return eventosProcessados;
}

function criarMapaFichas(dadosDTV) {
  const fichas = new Map();

  dadosDTV.forEach(linha => {
    const idAgenda = linha[COL_DTV_ID_AGENDA];
    const colaborador = linha[COL_DTV_COLABORADOR];

    if (idAgenda && colaborador) {
      const chave = `${String(idAgenda).trim()}|${normalizeName(colaborador)}`;
      fichas.set(chave, true);
    }
  });

  return fichas;
}

function processarEFiltrarEventos(dados, fichasEnviadas, obsMap) {
  const eventos = [];

  dados.forEach(linha => {
    const idAgenda = linha[COL_ID_AGENDA];
    const colaboradoresString = linha[COL_COLABORADOR_VIDEO];
    const servicoVideo = linha[COL_SERVICO_VIDEO];
    const indicacaoCancelamento = linha[COL_INDICACAO_CANCELAMENTO];
    const horaTerminoEvento = linha[COL_HORA_TERMINO_EVENTO];
    const horaLancCancelado = linha[COL_HORA_LANCAMENTO_CANCELADO];

    if (servicoVideo === 'Serviço não solicitado') return;

    let eventoFinalizado = !!(horaTerminoEvento && String(horaTerminoEvento).trim() !== '');

    let horaPrevista = linha[COL_HORA_PREVISTA];
    let horaPrevistaFormatada = "";
    let dataHoraPrevista = null;

    if (horaPrevista instanceof Date) {
      horaPrevistaFormatada =
        horaPrevista.getHours().toString().padStart(2, '0') + ":" +
        horaPrevista.getMinutes().toString().padStart(2, '0');

      const dataEvento = linha[COL_DATA_EVENTO];
      if (dataEvento instanceof Date) {
        dataHoraPrevista = new Date(
          dataEvento.getFullYear(),
          dataEvento.getMonth(),
          dataEvento.getDate(),
          horaPrevista.getHours(),
          horaPrevista.getMinutes()
        );
      }
    } else {
      horaPrevistaFormatada = String(horaPrevista).substring(0, 5);
    }

    const colaboradoresArray = (colaboradoresString ? String(colaboradoresString).split(',') : [])
      .map(s => s.trim()).filter(s => s !== "");

    colaboradoresArray.forEach(colaborador => {

      const chaveFicha = `${String(idAgenda).trim()}|${normalizeName(colaborador)}`;
      const fichaEnviada = fichasEnviadas.has(chaveFicha);

      //-------------------------------------------------------------------
      // ✔ NOVA REGRA CORRETA: Evento cancelado com > 1h ANTES → NÃO EXIBE
      //-------------------------------------------------------------------
      if (indicacaoCancelamento === "Cancelada" && horaLancCancelado && dataHoraPrevista instanceof Date) {

        let dataCancel = null;

        if (horaLancCancelado instanceof Date) {
          dataCancel = horaLancCancelado;
        } else {
          const t = new Date(horaLancCancelado);
          if (!isNaN(t.getTime())) dataCancel = t;
        }

        if (dataCancel instanceof Date) {
          const diffMs = dataHoraPrevista.getTime() - dataCancel.getTime();

          if (diffMs > 60 * 60 * 1000) {
            // Cancelado com mais de uma hora de antecedência → IGNORA TOTAL
            return;
          }
        }
      }
      //-------------------------------------------------------------------

      //--------------------------------------------------------------
      // ✔ Status "Atenção" somente se faltar < 1 hora
      //--------------------------------------------------------------
      let status = "";

      if (dataHoraPrevista instanceof Date) {
        const agora = new Date();
        const diffMs = dataHoraPrevista.getTime() - agora.getTime();

        if (diffMs > 0 && diffMs < 60 * 60 * 1000) {
          status = "Atenção";
        }
      }
      //--------------------------------------------------------------

      if (eventoFinalizado) {
        if (fichaEnviada) return;
        status = "Aguardando Ficha";
      } else if (linha[COL_HORA_INICIO_EVENTO]) {
        status = "Em andamento";
      } else if (indicacaoCancelamento === "Cancelada") {
        if (fichaEnviada) return;
        status = "Aguardando Ficha";
      }

      const chaveObs = `${String(idAgenda).trim()}|${normalizeName(colaborador)}`;
      const nomeEventoFinal = obsMap.get(chaveObs) || linha[COL_NOME_EVENTO];
      const isObs = obsMap.has(chaveObs);

      eventos.push({
        idAgenda,
        colaborador,
        status,
        local: linha[COL_LOCAL],
        horaPrevistaFormatada,
        servico: servicoVideo,
        tipoEvento: linha[COL_TIPO_EVENTO],
        nomeEvento: nomeEventoFinal,
        isObservacao: isObs
      });
    });
  });

  return eventos;
}

