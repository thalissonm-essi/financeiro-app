// Abastece by ESSI: abastecimento e manutenção dos veículos, rodando no Google Apps Script (grátis).
// Projeto separado do Anota, ligado a uma planilha própria.
// Propriedades do script (Configurações do projeto): ANTHROPIC_API_KEY. MOTO_KEY é criada pelo ativar().

const P = PropertiesService.getScriptProperties();
const cfg = k => P.getProperty(k);
const fmtData = d => Utilities.formatDate(d, "America/Sao_Paulo", "yyyy-MM-dd");
const COMBUSTIVEIS = ["Gasolina comum", "Gasolina aditivada", "Etanol", "Outro"];
const ABAS = {
  VEICULOS: ["ID", "Nome", "Tipo", "Ano", "Criado em"],
  ABASTECIMENTOS: ["ID", "Veículo", "Data", "Km", "Litros", "Valor", "R$/litro", "Combustível", "Tanque cheio", "Fotos"],
  MANUTENCAO: ["ID", "Veículo", "Item", "A cada (km)", "Último km", "Última data", "Valor", "Obs"],
};
const aba = nome => SpreadsheetApp.getActive().getSheetByName(nome);
const json = x => ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);
const limpa = (s, n) => String(s || "").replace(/^[=+\-@\s]+/, "").slice(0, n);  // texto que começa com = vira fórmula
const id_ = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);

// O app manda tudo por POST (corpo text/plain com JSON, evita bloqueio de CORS)
function doPost(e) {
  let u;
  try { u = JSON.parse(e.postData.contents); } catch (err) { return json({ erro: "pedido inválido" }); }
  if (!u.k || u.k !== cfg("MOTO_KEY")) return json({ erro: "negado" });
  try {
    if (u.ler) return json(ler_(u.ler));
    if (u.abastecimento) { salvarAbast_(u.abastecimento); return json(tudo_()); }
    if (u.apagar) { apagarLinha_("ABASTECIMENTOS", u.apagar); return json(tudo_()); }
    if (u.veiculo) { salvarVeiculo_(u.veiculo); return json(tudo_()); }
    if (u.manut) { salvarManut_(u.manut); return json(tudo_()); }
    if (u.apagarManut) { apagarLinha_("MANUTENCAO", u.apagarManut); return json(tudo_()); }
    return json(tudo_());
  } catch (err) { return json({ erro: err.message }); }
}
function doGet() { return ContentService.createTextOutput("Abastece ativo."); }

// ---------- Leitura ----------
function linhas_(nome) { return aba(nome).getDataRange().getValues().slice(1).filter(r => r[0]); }
const dt_ = v => v instanceof Date ? fmtData(v) : String(v || "");
function tudo_() {
  return {
    veiculos: linhas_("VEICULOS").map(r => ({ id: r[0], nome: r[1], tipo: r[2], ano: String(r[3]) })),
    abastecimentos: linhas_("ABASTECIMENTOS").map(r => ({ id: r[0], veiculoId: r[1], data: dt_(r[2]), km: +r[3], litros: +r[4],
      valor: +r[5], precoLitro: +r[6], combustivel: r[7], tanqueCheio: r[8] === true || r[8] === "Sim", fotos: String(r[9] || "").split(" ").filter(String) })),
    manutencao: linhas_("MANUTENCAO").map(r => ({ id: r[0], veiculoId: r[1], item: r[2], intervaloKm: +r[3] || 0, ultimoKm: +r[4] || 0,
      ultimaData: dt_(r[5]), valor: +r[6] || 0, obs: r[7] })),
  };
}

// ---------- Escrita ----------
function salvarAbast_(a) {
  const km = +a.km, litros = +a.litros, valor = +a.valor;
  if (!(km > 0 && km < 2e6 && litros > 0 && litros < 500 && valor > 0 && valor < 1e5)) throw new Error("Confira km, litros e valor.");
  if (!linhas_("VEICULOS").some(r => r[0] === a.veiculoId)) throw new Error("Veículo não encontrado.");
  const data = /^\d{4}-\d{2}-\d{2}$/.test(a.data) ? a.data : fmtData(new Date());
  const fotos = (a.imagens || []).slice(0, 3).map((b64, i) => salvarFoto_(b64, data + "-" + km + "-" + (i + 1))).join(" ");
  const [y, m, d] = data.split("-").map(Number);
  aba("ABASTECIMENTOS").appendRow([id_("a"), a.veiculoId, new Date(y, m - 1, d), km, litros, valor, +(valor / litros).toFixed(3),
    COMBUSTIVEIS.includes(a.combustivel) ? a.combustivel : "Outro", a.tanqueCheio ? "Sim" : "Não", fotos]);
}

// Fotos vão para a pasta "Abastece fotos" no seu Drive (só você vê). Na planilha fica o ID do arquivo.
function salvarFoto_(b64, nome) {
  if (String(b64).length > 4e6) return "";
  const it = DriveApp.getFoldersByName("Abastece fotos"), pasta = it.hasNext() ? it.next() : DriveApp.createFolder("Abastece fotos");
  return pasta.createFile(Utilities.newBlob(Utilities.base64Decode(b64), "image/jpeg", nome + ".jpg")).getId();
}

function salvarVeiculo_(v) {
  const nome = limpa(v.nome, 60);
  if (!nome) throw new Error("Dê um nome ao veículo.");
  aba("VEICULOS").appendRow([id_("v"), nome, limpa(v.tipo, 20) || "Moto", limpa(v.ano, 12), new Date()]);
}

// Cria item novo, ou (com id) registra que a manutenção foi feita agora
function salvarManut_(m) {
  const ws = aba("MANUTENCAO");
  const [y, mo, d] = (/^\d{4}-\d{2}-\d{2}$/.test(m.ultimaData) ? m.ultimaData : fmtData(new Date())).split("-").map(Number);
  const linha = [m.veiculoId, limpa(m.item, 60), +m.intervaloKm || 0, +m.ultimoKm || 0, new Date(y, mo - 1, d), +m.valor || 0, limpa(m.obs, 200)];
  if (!linha[1]) throw new Error("Dê um nome à manutenção.");
  if (m.id) {
    const i = ws.getDataRange().getValues().findIndex(r => r[0] === m.id);
    if (i > 0) return ws.getRange(i + 1, 2, 1, 7).setValues([linha]);
  }
  ws.appendRow([id_("m")].concat(linha));
}

function apagarLinha_(nome, id) {
  const ws = aba(nome), v = ws.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === id) ws.deleteRow(i + 1);
}

// ---------- IA: lê fotos da bomba e do painel ----------
function ler_(l) {
  limite_("ler", 40);
  const imgs = (l.imagens || []).slice(0, 3).filter(b => String(b).length < 4e6);
  if (!imgs.length) throw new Error("Mande ao menos uma foto.");
  const ult = linhas_("ABASTECIMENTOS").filter(r => r[1] === l.veiculoId).map(r => +r[3]).sort((a, b) => b - a)[0];
  const prompt = `Estas fotos são de um abastecimento no Brasil. Podem mostrar o visor da bomba (total em R$, litros, preço por litro) e/ou o painel do veículo (odômetro).
Use o odômetro TOTAL, nunca o parcial (TRIP). Vírgula ou ponto podem ser separador decimal na bomba.${ult ? `\nO último km registrado foi ${ult}; o novo deve ser maior.` : ""}
Responda SÓ com JSON: {"valor": número|null, "litros": número|null, "precoLitro": número|null, "km": número|null, "combustivel": ${JSON.stringify(COMBUSTIVEIS)} ou null, "obs": "dúvida de leitura em poucas palavras, ou vazio"}.
Use null para o que não estiver visível. Não invente.`;
  const content = imgs.map(b => ({ type: "image", source: { type: "base64", media_type: "image/jpeg", data: b } })).concat({ type: "text", text: prompt });
  const r = claudeJson_([{ role: "user", content }], 400);
  if (r.litros == null && r.valor && r.precoLitro) r.litros = +(r.valor / r.precoLitro).toFixed(3);
  return r;
}

function claudeJson_(messages, max) {
  const r = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post", contentType: "application/json", muteHttpExceptions: true,
    headers: { "x-api-key": cfg("ANTHROPIC_API_KEY"), "anthropic-version": "2023-06-01" },
    payload: JSON.stringify({ model: "claude-haiku-5-5", max_tokens: max, messages }) });
  const corpo = JSON.parse(r.getContentText());
  if (!corpo.content) throw new Error("Claude: " + r.getContentText().slice(0, 200));
  const t = corpo.content.find(c => c.type === "text").text;  // o 1º bloco pode ser "thinking"
  return JSON.parse(t.slice(t.indexOf("{"), t.lastIndexOf("}") + 1));
}

// Limite por hora (protege a conta da Anthropic de loop). Contador no cache.
function limite_(nome, max) {
  const c = CacheService.getScriptCache(), k = "lim" + nome + new Date().getHours(), n = +(c.get(k) || 0) + 1;
  c.put(k, String(n), 3600);
  if (n > max) throw new Error("Limite de leituras por hora atingido. Preencha à mão ou tente mais tarde.");
}

// ---------- Rodar 1x no editor ----------
// Cria as abas, a chave do app e já traz a Elite e os 3 primeiros abastecimentos.
function ativar() {
  const ss = SpreadsheetApp.getActive();
  Object.entries(ABAS).forEach(([nome, cab]) => {
    const ws = ss.getSheetByName(nome) || ss.insertSheet(nome);
    if (!ws.getRange("A1").getValue()) {
      ws.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold").setBackground("#d9520b").setFontColor("#fff");
      ws.setFrozenRows(1);
    }
  });
  if (!cfg("MOTO_KEY")) P.setProperty("MOTO_KEY", Utilities.getUuid());
  if (!linhas_("VEICULOS").length) {
    aba("VEICULOS").appendRow(["elite125", "Honda Elite 125", "Moto", "2020/21", new Date()]);
    aba("ABASTECIMENTOS").appendRow(["a-55020", "elite125", new Date(2026, 9, 18), 55020, 4.85, 29.54, 6.091, "Gasolina comum", "Sim", ""]);
    aba("ABASTECIMENTOS").appendRow(["a-55182", "elite125", new Date(2026, 9, 23), 55182, 4.80, 28.18, 5.871, "Gasolina comum", "Sim", ""]);
    aba("ABASTECIMENTOS").appendRow(["a-55335", "elite125", new Date(2026, 9, 27), 55335, 4.557, 27.30, 5.99, "Gasolina comum", "Sim", ""]);
  }
  DriveApp.getRootFolder();  // pede a permissão do Drive agora, não na primeira foto
  Logger.log("Chave do app: " + cfg("MOTO_KEY"));
}
