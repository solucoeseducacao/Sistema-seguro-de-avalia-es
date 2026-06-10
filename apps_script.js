// ═══════════════════════════════════════════════════════════════
// SISTEMA SEGURO DE AVALIAÇÕES — Apps Script (Backend)
// Prof. Felipe Vigneron Azevedo
// Cole este código no Google Apps Script da sua planilha
// ═══════════════════════════════════════════════════════════════

const SENHA_PROF = "Jm10sa06@!"; // Troque antes de usar

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const acao = d.acao;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (acao === "salvarProva")     return resp(salvarProva(ss, d));
    if (acao === "listarProvas")    return resp(listarProvas(ss, d));
    if (acao === "buscarProva")     return resp(buscarProva(ss, d));
    if (acao === "salvarResultado") return resp(salvarResultado(ss, d));
    if (acao === "cadastrarTurma")  return resp(cadastrarTurma(ss, d));
    if (acao === "listarAlunos")    return resp(listarAlunos(ss, d));
    if (acao === "iniciarSessao")   return resp(iniciarSessao(ss, d));
    if (acao === "encerrarSessao")  return resp(encerrarSessao(ss, d));
    if (acao === "listarResultados")return resp(listarResultados(ss, d));

    return resp({ ok: false, erro: "Ação desconhecida: " + acao });
  } catch (err) {
    return resp({ ok: false, erro: err.toString() });
  }
}

function doGet(e) {
  // GET para listar provas (aluno busca pelo código)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const acao = e.parameter.acao;
    if (acao === "buscarProva") return resp(buscarProva(ss, e.parameter));
    if (acao === "listarAlunos") return resp(listarAlunos(ss, e.parameter));
    if (acao === "listarResultados") return resp(listarResultados(ss, e.parameter));
    if (acao === "listarProvas") return resp(listarProvas(ss, e.parameter));
    return resp({ ok: false, erro: "Ação GET desconhecida" });
  } catch (err) {
    return resp({ ok: false, erro: err.toString() });
  }
}

function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ABA: PROVAS ──────────────────────────────────────────────
function getAbaProvas(ss) {
  let aba = ss.getSheetByName("__provas__");
  if (!aba) {
    aba = ss.insertSheet("__provas__");
    aba.appendRow(["id","codigo","nome","disciplina","tipo","dataProva","descricao","ano","semestre",
                   "tempoMinutos","penalidade1","horaInicio","horaFim","questoes","encerrada","criadaEm"]);
    aba.setFrozenRows(1);
    aba.hideSheet();
  }
  return aba;
}

function salvarProva(ss, d) {
  if (d.senha !== SENHA_PROF) return { ok: false, erro: "Senha inválida" };
  const aba = getAbaProvas(ss);
  const dados = aba.getDataRange().getValues();
  const idx = dados.findIndex((r, i) => i > 0 && r[0] === d.prova.id);
  const row = [
    d.prova.id, d.prova.codigo, d.prova.nome, d.prova.disciplina,
    d.prova.tipo, d.prova.dataProva || "", d.prova.descricao || "",
    d.prova.ano || "2026", d.prova.semestre || "",
    d.prova.tempoMinutos, d.prova.penalidade1,
    d.prova.horaInicio || "", d.prova.horaFim || "",
    JSON.stringify(d.prova.questoes),
    d.prova.encerrada ? "1" : "0",
    d.prova.criadaEm || new Date().toLocaleString("pt-BR")
  ];
  if (idx >= 1) aba.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  else aba.appendRow(row);
  return { ok: true };
}

function listarProvas(ss, d) {
  if (d.senha !== SENHA_PROF) return { ok: false, erro: "Senha inválida" };
  const aba = getAbaProvas(ss);
  const dados = aba.getDataRange().getValues();
  if (dados.length <= 1) return { ok: true, provas: [] };
  const provas = dados.slice(1).map(r => ({
    id: r[0], codigo: r[1], nome: r[2], disciplina: r[3],
    tipo: r[4], dataProva: r[5], descricao: r[6], ano: r[7], semestre: r[8],
    tempoMinutos: r[9], penalidade1: r[10],
    horaInicio: r[11], horaFim: r[12],
    questoes: JSON.parse(r[13] || "[]"),
    encerrada: r[14] === "1", criadaEm: r[15]
  }));
  return { ok: true, provas };
}

function buscarProva(ss, d) {
  const aba = getAbaProvas(ss);
  const dados = aba.getDataRange().getValues();
  const row = dados.find((r, i) => i > 0 && r[1] === d.codigo);
  if (!row) return { ok: false, erro: "Prova não encontrada" };
  if (row[14] === "1") return { ok: false, erro: "Prova encerrada" };
  // verificar janela de tempo
  if (row[11] && row[12]) {
    const agora = new Date();
    const hoje = Utilities.formatDate(agora, "America/Sao_Paulo", "yyyy-MM-dd");
    const ini = new Date(hoje + "T" + row[11] + ":00");
    const fim = new Date(hoje + "T" + row[12] + ":00");
    if (agora < ini) return { ok: false, erro: "Prova ainda não iniciou. Horário: " + row[11] };
    if (agora > fim) return { ok: false, erro: "Prova encerrada. Horário limite: " + row[12] };
  }
  return {
    ok: true,
    prova: {
      id: row[0], codigo: row[1], nome: row[2], disciplina: row[3],
      tipo: row[4], dataProva: row[5], descricao: row[6], ano: row[7], semestre: row[8],
      tempoMinutos: row[9], penalidade1: row[10],
      horaInicio: row[11], horaFim: row[12],
      questoes: JSON.parse(row[13] || "[]"),
      encerrada: false
    }
  };
}

// ── ABA: SESSÕES ATIVAS (controle de login duplo) ────────────
function getAbaSessoes(ss) {
  let aba = ss.getSheetByName("__sessoes__");
  if (!aba) {
    aba = ss.insertSheet("__sessoes__");
    aba.appendRow(["provaId","matricula","sessionToken","inicio"]);
    aba.setFrozenRows(1);
    aba.hideSheet();
  }
  return aba;
}

function iniciarSessao(ss, d) {
  const aba = getAbaSessoes(ss);
  const dados = aba.getDataRange().getValues();
  // Verificar se já existe sessão ativa para provaId + matricula
  const limite = new Date(Date.now() - 3 * 60 * 60 * 1000); // 3h atrás
  for (let i = dados.length - 1; i >= 1; i--) {
    const r = dados[i];
    if (r[0] === d.provaId && r[1] === d.matricula) {
      const inicio = new Date(r[3]);
      if (inicio > limite) {
        return { ok: false, erro: "Você já está realizando esta prova em outro dispositivo." };
      }
      // Sessão expirada — limpa
      aba.deleteRow(i + 1);
    }
  }
  const token = Utilities.getUuid();
  aba.appendRow([d.provaId, d.matricula, token, new Date().toISOString()]);
  return { ok: true, token };
}

function encerrarSessao(ss, d) {
  const aba = getAbaSessoes(ss);
  const dados = aba.getDataRange().getValues();
  for (let i = dados.length - 1; i >= 1; i--) {
    if (dados[i][0] === d.provaId && dados[i][1] === d.matricula) {
      aba.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// ── ABA: ALUNOS ──────────────────────────────────────────────
function getAbaAlunos(ss) {
  let aba = ss.getSheetByName("__alunos__");
  if (!aba) {
    aba = ss.insertSheet("__alunos__");
    aba.appendRow(["disciplina","semestre","matricula","nome"]);
    aba.setFrozenRows(1);
    aba.hideSheet();
  }
  return aba;
}

function cadastrarTurma(ss, d) {
  if (d.senha !== SENHA_PROF) return { ok: false, erro: "Senha inválida" };
  const aba = getAbaAlunos(ss);
  const dados = aba.getDataRange().getValues();
  // Remove alunos anteriores da mesma disciplina+semestre
  for (let i = dados.length - 1; i >= 1; i--) {
    if (dados[i][0] === d.disciplina && dados[i][1] === d.semestre) {
      aba.deleteRow(i + 1);
    }
  }
  // Insere novos
  for (const al of d.alunos) {
    aba.appendRow([d.disciplina, d.semestre, al.matricula, al.nome]);
  }
  return { ok: true, cadastrados: d.alunos.length };
}

function listarAlunos(ss, d) {
  const aba = getAbaAlunos(ss);
  const dados = aba.getDataRange().getValues();
  const alunos = dados.slice(1)
    .filter(r => r[0] === d.disciplina && r[1] === d.semestre)
    .map(r => ({ matricula: String(r[2]), nome: r[3] }));
  return { ok: true, alunos };
}

// ── ABA: RESULTADOS ──────────────────────────────────────────
function salvarResultado(ss, d) {
  const r = d.resultado;

  // ── Validação de token ──────────────────────────────────────
  // Aceita token enviado como d.token ou dentro de d.resultado._token
  const tokenRecebido = d.token || r._token || "";
  if (tokenRecebido) {
    const abaSessoes = getAbaSessoes(ss);
    const sessoes = abaSessoes.getDataRange().getValues();
    const limite = new Date(Date.now() - 3 * 60 * 60 * 1000);
    const sessaoValida = sessoes.some((row, i) =>
      i > 0 &&
      row[0] === r.provaId &&
      row[2] === tokenRecebido &&
      new Date(row[3]) > limite
    );
    if (!sessaoValida) {
      return { ok: false, erro: "Sessão inválida ou expirada. Resultado não aceito." };
    }
  }
  // ── Fim validação de token ──────────────────────────────────

  // 1. Aba raw por disciplina+tipo
  const nomeAba = limpar(r.disciplina) + " — " + r.tipo;
  let aba = ss.getSheetByName(nomeAba);
  if (!aba) {
    aba = ss.insertSheet(nomeAba);
    aba.appendRow(["Matrícula","Nome","Acertos","Total","Nota Bruta","Penalidade","Nota Final","Violações","Motivo","Data","Prova","Semestre"]);
    aba.getRange(1,1,1,12).setFontWeight("bold").setBackground("#d9ead3");
    aba.setFrozenRows(1);
  }
  // Evitar duplicata (mesmo aluno + mesma prova)
  const dados = aba.getDataRange().getValues();
  const jaExiste = dados.some((row, i) => i > 0 && row[0] === r.matricula && row[10] === r.provaId);
  if (jaExiste) return { ok: false, erro: "Resultado já registrado para este aluno nesta prova." };

  aba.appendRow([r.matricula, r.nome, r.acertos, r.total, r.notaBruta,
    r.penalidade, r.notaFinal, r.violacoes, r.motivo,
    r.data, r.provaId, r.semestre || ""]);

  // 2. Se for Quiz → atualizar aba consolidada
  if (r.tipo === "Quiz") atualizarConsolidadoQuiz(ss, r);

  // Encerrar sessão automaticamente
  encerrarSessao(ss, { provaId: r.provaId, matricula: r.matricula });

  return { ok: true };
}

function atualizarConsolidadoQuiz(ss, r) {
  const nomeAba = limpar(r.disciplina) + " — Quiz Consolidado";
  let aba = ss.getSheetByName(nomeAba);
  const colunaQuiz = "Quiz " + (r.dataProva || r.data.substring(0,10));

  if (!aba) {
    aba = ss.insertSheet(nomeAba);
    aba.appendRow(["Matrícula","Nome"]);
    aba.getRange(1,1,1,2).setFontWeight("bold").setBackground("#cfe2f3");
    aba.setFrozenRows(1);
    aba.setFrozenColumns(2);
  }

  const dados = aba.getDataRange().getValues();
  const cabecalho = dados[0];

  // Encontrar ou criar coluna para este quiz
  let colQuiz = cabecalho.indexOf(colunaQuiz);
  if (colQuiz === -1) {
    // Nova coluna para este quiz
    const ultimaCol = cabecalho.length + 1;
    aba.getRange(1, ultimaCol).setValue(colunaQuiz).setFontWeight("bold").setBackground("#cfe2f3");
    colQuiz = ultimaCol - 1;
    // Recalcula fórmulas nas colunas de resumo
    atualizarFormulasQuiz(aba);
  }

  // Encontrar ou criar linha do aluno
  let linhaAluno = -1;
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(r.matricula)) { linhaAluno = i + 1; break; }
  }
  if (linhaAluno === -1) {
    aba.appendRow([r.matricula, r.nome]);
    linhaAluno = aba.getLastRow();
  }

  // Gravar nota
  aba.getRange(linhaAluno, colQuiz + 1).setValue(parseFloat(r.notaFinal));
  atualizarFormulasQuiz(aba);
}

function atualizarFormulasQuiz(aba) {
  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return;
  const cabecalho = dados[0];

  // Colunas de quiz = todas exceto Matrícula, Nome, Média, Valor 3.0, Valor 4.0
  const colsFixas = ["Matrícula","Nome","Média","Valor 3.0","Valor 4.0"];
  const colsQuiz = cabecalho
    .map((h, i) => ({ h, i }))
    .filter(x => !colsFixas.includes(x.h) && String(x.h).startsWith("Quiz"));

  if (!colsQuiz.length) return;

  // Garantir colunas de resumo ao final
  let iMedia = cabecalho.indexOf("Média");
  let iV3 = cabecalho.indexOf("Valor 3.0");
  let iV4 = cabecalho.indexOf("Valor 4.0");
  const totalCols = cabecalho.length;

  if (iMedia === -1) {
    iMedia = totalCols;
    aba.getRange(1, iMedia + 1).setValue("Média").setFontWeight("bold").setBackground("#fff2cc");
  }
  if (iV3 === -1) {
    iV3 = Math.max(totalCols, iMedia + 1);
    aba.getRange(1, iV3 + 1).setValue("Valor 3.0").setFontWeight("bold").setBackground("#f4cccc");
  }
  if (iV4 === -1) {
    iV4 = Math.max(totalCols, iV3 + 1);
    aba.getRange(1, iV4 + 1).setValue("Valor 4.0").setFontWeight("bold").setBackground("#d9ead3");
  }

  // Inserir fórmulas para cada aluno
  const letras = colsQuiz.map(c => colLetra(c.i + 1));
  for (let i = 1; i < dados.length; i++) {
    const lin = i + 1;
    if (!dados[i][0]) continue; // linha vazia
    const refs = letras.map(l => l + lin).join(",");
    const n = colsQuiz.length;
    // Média = AVERAGE das colunas de quiz (ignora vazias com IFERROR)
    aba.getRange(lin, iMedia + 1).setFormula(`=IFERROR(AVERAGE(${refs}),"")`);
    aba.getRange(lin, iV3 + 1).setFormula(`=IFERROR(${colLetra(iMedia+1)}${lin}*0.3,"")`);
    aba.getRange(lin, iV4 + 1).setFormula(`=IFERROR(${colLetra(iMedia+1)}${lin}*0.4,"")`);
  }
}

function colLetra(n) {
  let s = "";
  while (n > 0) { s = String.fromCharCode(65 + (n-1)%26) + s; n = Math.floor((n-1)/26); }
  return s;
}

function listarResultados(ss, d) {
  if (d.senha !== SENHA_PROF) return { ok: false, erro: "Senha inválida" };
  const nomeAba = limpar(d.disciplina) + " — " + d.tipo;
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) return { ok: true, resultados: [] };
  const dados = aba.getDataRange().getValues();
  const resultados = dados.slice(1)
    .filter(r => !d.provaId || r[10] === d.provaId)
    .map(r => ({
      matricula: r[0], nome: r[1], acertos: r[2], total: r[3],
      notaBruta: r[4], penalidade: r[5], notaFinal: r[6],
      violacoes: r[7], motivo: r[8], data: r[9], provaId: r[10], semestre: r[11]
    }));
  return { ok: true, resultados };
}

function limpar(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g,"")
    .replace(/[^a-zA-Z0-9 \-]/g,"").trim().substring(0,30);
}
