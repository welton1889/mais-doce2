// ===== Mais Doce - Backend do estoque (Google Apps Script) =====
const EMAILS = "weltonpereiracorreia1889@gmail.com,pricy2@gmail.com,weltonpereiracorreia@gmail.com";
const CAB_ESTOQUE = ["Tipo", "Marca", "Sabor", "Quantidade", "Validade", "Aviso enviado"];
const CAB_MOV = ["Data/Hora", "Movimento", "Tipo", "Marca", "Sabor", "Quantidade", "Validade"];

function aba_(nome, cab) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(nome) || ss.insertSheet(nome);
  if (s.getLastRow() === 0) {
    s.appendRow(cab);
    s.getRange(1, 1, 1, cab.length).setFontWeight("bold").setBackground("#f8c8d8");
    s.setFrozenRows(1);
    s.getRange("E:E").setNumberFormat("@");
  }
  return s;
}
function dataTxt_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(v || "");
}
function lerItens_() {
  const s = aba_("Estoque", CAB_ESTOQUE);
  return s.getDataRange().getValues().slice(1).map((r, i) => ({
    linha: i + 2, tipo: r[0], marca: r[1], sabor: r[2], qtd: Number(r[3]) || 0, validade: dataTxt_(r[4]), aviso: r[5]
  }));
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return json_({ itens: lerItens_().map(({ tipo, marca, sabor, qtd, validade }) => ({ tipo, marca, sabor, qtd, validade })) });
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action === "check") return json_({ ok: true, enviados: verificarVencidos() });
  if (d.action === "move") {
    const s = aba_("Estoque", CAB_ESTOQUE);
    const it = lerItens_().find(i => i.tipo === d.tipo && i.marca === d.marca && i.sabor === d.sabor && i.validade === (d.validade || ""));
    const delta = d.mov === "Entrada" ? d.qtd : -d.qtd;
    if (it) s.getRange(it.linha, 4).setValue(Math.max(0, it.qtd + delta));
    else s.appendRow([d.tipo, d.marca, d.sabor, Math.max(0, delta), d.validade || "", ""]);
    aba_("Movimentos", CAB_MOV).appendRow([new Date(), d.mov, d.tipo, d.marca, d.sabor, d.qtd, d.validade || ""]);
    verificarVencidos();
    return json_({ ok: true });
  }
  return json_({ ok: false });
}
// Envia e-mail dos produtos vencidos (uma vez por item)
function verificarVencidos() {
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = aba_("Estoque", CAB_ESTOQUE);
  const venc = lerItens_().filter(i => i.qtd > 0 && i.validade && i.validade < hoje && i.aviso !== "SIM");
  if (!venc.length) return 0;
  const corpo = venc.map(i => `• ${i.tipo} - ${i.marca} - ${i.sabor} | Qtd: ${i.qtd} | Venceu em: ${i.validade}`).join("\n");
  MailApp.sendEmail(EMAILS, "⚠ Mais Doce: produto(s) VENCIDO(S) no estoque", "Os itens abaixo estão vencidos:\n\n" + corpo);
  venc.forEach(i => s.getRange(i.linha, 6).setValue("SIM"));
  return venc.length;
}
// Rode UMA vez para verificar vencidos automaticamente todo dia às 8h
function ativarVerificacaoDiaria() {
  ScriptApp.newTrigger("verificarVencidos").timeBased().everyDays(1).atHour(8).create();
}
