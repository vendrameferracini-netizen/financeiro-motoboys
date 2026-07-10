import {
  isSupabaseConfigured,
  supabase,
  TABLES,
  getSupabaseSession,
  signInSupabase,
  signOutSupabase,
  loadProfiles,
  saveProfile,
  loadCloudState,
  saveCloudRecord,
  deleteCloudRecord,
  saveCloudSettings,
  saveBackupToCloud,
  logCloudChange
} from "./supabaseClient.js";

const workbook = window.EMBEDDED_WORKBOOK || { sheetCount: 0, sheets: [] };
const syncReport = window.SYNC_REPORT || null;
const sheets = workbook.sheets || [];
const BRL = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const NUM = new Intl.NumberFormat("pt-BR");
const STORE_KEY = "motoboyFinanceiro.baseSocios.v1";
const MIGRATION_SNAPSHOT_KEY = `${STORE_KEY}.preSupabaseMigration`;
const MIGRATION_MARK_KEY = `${STORE_KEY}.migrationStatus`;
const PARTNERS = ["GIL", "SALES", "GUILHERME"];
const RESPONSIBLES = [...PARTNERS, "BASE"];
const PARTNER_SHEETS = { GIL: "GIL", SALES: "SALES", GUILHERME: "GUILHERME M" };
const PERMISSION_MESSAGE = "Você não possui permissão para realizar esta ação.";
const VIEWS_BY_ROLE = {
  admin: ["dashboard", "motoboys", "freelancers", "socios", "base", "lancamentos", "descontos", "descontos-base", "despesas", "fechamentos", "recibos", "relatorios", "backup", "configuracoes", "usuarios"],
  GIL: ["dashboard", "motoboys", "socios", "base", "descontos", "despesas", "fechamentos", "recibos", "relatorios"],
  SALES: ["dashboard", "motoboys", "socios", "base", "descontos", "despesas", "fechamentos", "recibos", "relatorios"],
  GUILHERME: ["dashboard", "motoboys", "socios", "base", "descontos", "despesas", "fechamentos", "recibos", "relatorios"],
  OPERADOR: ["lancamentos"]
};
const CLEAN_OPERATIONAL_MODE = true;
const CLEAN_OPERATIONAL_VERSION = "real-use-v1";
const BASE_RATE_ML = 8;
const BASE_RATE_SHOPEE = 5;
const DEFAULT_NO_COLLECTION_ML = 6;
const DEFAULT_NO_COLLECTION_SHOPEE = 4;

let state = loadState();
let imported = { riders: [], daily: [], baseEntries: [], discounts: [], payments: [], ignored: [], auditBlocks: [], discountOrigins: [], duplicateCount: 0, guilhermeAuditRows: [] };
let editingRiderId = "";
let editingDiscountId = "";
let editingExpenseId = "";
let editingDailyId = "";
let selectedRider = "";
let selectedPartner = "GIL";
let supabaseSession = null;
let supabaseProfile = null;
let supabaseOnline = false;
let lastSupabaseSync = "";
let localStateForMigration = null;
let profiles = [];

const $ = (id) => document.getElementById(id);

function defaultState() {
  return { riders: [], daily: [], baseEntries: [], discounts: [], expenses: [], paid: {}, payments: [], receipts: [], basePaid: {}, config: { ml: 8, shopee: 5, avulso: 0 }, audit: [], cleanOperational: true, lastBackupAt: "" };
}

function normalizeStateData(data) {
  const defaults = defaultState();
  const parsed = data && typeof data === "object" ? data : {};
  return {
    ...defaults,
    ...parsed,
    riders: Array.isArray(parsed.riders) ? parsed.riders : [],
    daily: Array.isArray(parsed.daily) ? parsed.daily : [],
    baseEntries: Array.isArray(parsed.baseEntries) ? parsed.baseEntries : [],
    discounts: Array.isArray(parsed.discounts) ? parsed.discounts : [],
    expenses: Array.isArray(parsed.expenses) ? parsed.expenses : [],
    payments: Array.isArray(parsed.payments) ? parsed.payments : [],
    receipts: Array.isArray(parsed.receipts) ? parsed.receipts : [],
    paid: parsed.paid && typeof parsed.paid === "object" ? parsed.paid : {},
    basePaid: parsed.basePaid && typeof parsed.basePaid === "object" ? parsed.basePaid : {},
    config: { ...defaults.config, ...(parsed.config || {}) },
    audit: Array.isArray(parsed.audit) ? parsed.audit : [],
    lastBackupAt: parsed.lastBackupAt || ""
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (raw) return normalizeStateData(JSON.parse(raw));
  } catch {}
  return defaultState();
}

function saveState(action, detail, meta = null) {
  if (action) state.audit.unshift({ at: new Date().toISOString(), action, detail, ...(meta ? { meta } : {}) });
  localStorage.setItem(STORE_KEY, JSON.stringify(state));
  if (supabaseOnline && supabaseProfile?.role === "admin") {
    saveCloudSettings(state)
      .then(() => logCloudChange(action || "sync", detail || "", meta || {}))
      .catch((error) => showSupabaseError(error));
  }
}

function showSupabaseError(error) {
  const message = error?.message || "Sem conexão com Supabase. Não foi possível salvar.";
  console.error("Erro Supabase:", error);
  const target = $("supabaseFeedback");
  if (target) {
    target.textContent = message.includes("Supabase") ? message : `Sem conexão com Supabase. Não foi possível salvar. ${message}`;
    target.className = "feedback error";
  }
}

function showDailyFeedback(message, type = "ok") {
  const target = $("dailyFeedback");
  if (!target) return;
  target.textContent = message;
  target.className = `feedback ${type}`;
}

async function persistRecord(bucket, record, options = {}) {
  if (!requireWrite(bucket, record)) {
    if (options.throwOnError) throw new Error(PERMISSION_MESSAGE);
    return record;
  }
  if (!supabaseOnline) {
    const error = new Error("Sem conexão com Supabase. Não foi possível salvar.");
    showSupabaseError(error);
    if (options.throwOnError) throw error;
    return record;
  }
  try {
    if (!record.partner && !record.responsible && !record.owner && supabaseProfile?.role && supabaseProfile.role !== "admin") record.responsible = supabaseProfile.role;
    const saved = await saveCloudRecord(bucket, record);
    Object.assign(record, saved);
    lastSupabaseSync = new Date().toISOString();
    renderSupabaseStatus();
    return record;
  } catch (error) {
    showSupabaseError(error);
    if (options.throwOnError) throw error;
    return record;
  }
}

async function removeCloudRecord(bucket, record) {
  if (!requireWrite(bucket, record)) return;
  if (!supabaseOnline) {
    showSupabaseError(new Error("Sem conexão com Supabase. Não foi possível salvar."));
    return;
  }
  try {
    await deleteCloudRecord(bucket, record);
    lastSupabaseSync = new Date().toISOString();
    renderSupabaseStatus();
  } catch (error) {
    showSupabaseError(error);
  }
}

function resetOperationalBuckets() {
  state.daily = [];
  state.baseEntries = [];
  state.discounts = [];
  state.expenses = [];
  state.paid = {};
  state.payments = [];
  state.receipts = [];
  state.basePaid = {};
}

function prepareCleanOperationalState() {
  if (!CLEAN_OPERATIONAL_MODE || state.cleanOperationalVersion === CLEAN_OPERATIONAL_VERSION) return;
  resetOperationalBuckets();
  state.cleanOperational = true;
  state.cleanOperationalVersion = CLEAN_OPERATIONAL_VERSION;
  saveState("Modo limpo aplicado", "Dados operacionais antigos foram zerados; cadastros, valores e regras foram preservados.");
}

function applyCleanOperationalMode(model) {
  if (!CLEAN_OPERATIONAL_MODE) return model;
  return {
    ...model,
    daily: [],
    baseEntries: [],
    discounts: [],
    payments: [],
    auditBlocks: [],
    discountOrigins: [],
    duplicateCount: 0,
    guilhermeAuditRows: [],
    ignored: [audit("Versao limpa", "Lancamentos operacionais importados foram zerados. A planilha permanece apenas como base de cadastros/regras.")]
  };
}

function seedRidersFromWorkbook() {
  const existing = new Map((state.riders || []).filter((r) => !isInvalidRiderName(r.name)).map((r) => [normalize(r.name), {
    id: r.id || uid("rider"),
    name: r.name,
    sheet: r.sheet || r.name,
    region: r.region || "",
    collection: r.collection || "com coleta",
    rateMl: Number.isFinite(Number(r.rateMl)) ? Number(r.rateMl) : state.config.ml,
    rateShopee: Number.isFinite(Number(r.rateShopee)) ? Number(r.rateShopee) : state.config.shopee,
    rateAvulso: Number.isFinite(Number(r.rateAvulso)) ? Number(r.rateAvulso) : state.config.avulso,
    note: r.note || "",
    active: r.active !== false
  }]));
  imported.riders.filter((r) => !r.partner).forEach((r) => {
    const key = normalize(r.name);
    if (!existing.has(key)) {
      existing.set(key, {
        id: uid("rider"),
        name: r.name,
        sheet: r.sheet || r.name,
        region: "",
        collection: "com coleta",
        rateMl: state.config.ml,
        rateShopee: state.config.shopee,
        rateAvulso: state.config.avulso,
        note: "",
        active: true
      });
    }
  });
  state.riders = [...existing.values()].sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));
  saveState();
}

function initializeWorkbookData() {
  prepareCleanOperationalState();
  imported = applyCleanOperationalMode(buildImportedModelV2());
  seedRidersFromWorkbook();
  selectedRider = allRiders()[0]?.name || "";
}

function uid(prefix) { return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`; }
function normalize(text) { return String(text || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
function riderDisplayName(row = {}) { return String(row.name || row.nome || row.rider || row.motoboy || row.motoboyName || "").trim(); }
function compareRiderName(a, b) { return riderDisplayName(a).localeCompare(riderDisplayName(b), "pt-BR", { sensitivity: "base" }); }
function sortByRiderName(rows = []) { return [...rows].sort(compareRiderName); }
function appRole() {
  if (!isSupabaseConfigured) return "admin";
  if (!supabaseSession) return "anon";
  const role = String(supabaseProfile?.role || "").trim();
  if (role === "admin") return "admin";
  const normalized = normalizeResponsible(role);
  if (normalized === "BASE" && normalize(role).includes("operador")) return "OPERADOR";
  return normalized;
}
function isAdminRole() { return appRole() === "admin"; }
function isOperatorRole() { return appRole() === "OPERADOR"; }
function isPartnerRole() { return PARTNERS.includes(appRole()); }
function permittedViews() { return VIEWS_BY_ROLE[appRole()] || []; }
function canView(view) { return permittedViews().includes(view); }
function permissionDenied() {
  window.alert(PERMISSION_MESSAGE);
  ["backupFeedback", "riderFeedback", "userFeedback", "supabaseFeedback"].forEach((id) => {
    const el = $(id);
    if (el) {
      el.textContent = PERMISSION_MESSAGE;
      el.className = "feedback error";
    }
  });
  return false;
}
function todayKey() { return new Date().toISOString().slice(0, 10); }
function ownsResponsible(value) {
  if (isAdminRole()) return true;
  if (!isPartnerRole()) return false;
  return normalizeResponsible(value) === appRole();
}
function canWriteRecord(bucket, record = {}) {
  if (isAdminRole()) return true;
  if (isOperatorRole()) return bucket === "daily";
  if (!isPartnerRole()) return false;
  if (bucket === "baseEntries") return ownsResponsible(record.partner || record.responsible);
  if (bucket === "discounts") return ownsResponsible(record.partner || record.responsible) && normalizeResponsible(record.partner || record.responsible) !== "BASE";
  if (bucket === "expenses") return ownsResponsible(record.responsible);
  if (bucket === "payments" || bucket === "receipts") return ownsResponsible(record.partner || record.responsible);
  return false;
}
function requireWrite(bucket, record = {}) {
  return canWriteRecord(bucket, record) || permissionDenied();
}
function managerAlias(text) {
  const key = normalize(text).replace(/[^a-z0-9]+/g, " ").trim();
  if (!key) return "";
  const hits = [];
  if (/\b(gm|g m|guilherme m|guilherme)\b/.test(key)) hits.push("GUILHERME");
  if (/\bsales\b/.test(key)) hits.push("SALES");
  if (/\bgil\b/.test(key)) hits.push("GIL");
  const unique = [...new Set(hits)];
  return unique.length === 1 ? unique[0] : "";
}
function money(value) { return BRL.format(Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100); }
function num(value) { return NUM.format(Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100); }
function escapeHtml(value) { return String(value ?? "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[ch]); }
function audit(type, detail) { return { type, detail }; }
function isHeaderText(text) { return /^(nome|data|vale|ocorrencia|ocorrência|valor|motorista|qtd|saida|saída|total dia)$/i.test(String(text || "").trim()); }
function isTotalText(text) { return /total|geral|quinzena|valor total|total entrada|total de pacotes/i.test(String(text || "")); }
function isNumericLike(text) { const value = String(text || "").trim(); return value !== "" && /^-?\d+([.,]\d+)?$/.test(value); }
function isExcelSerialLike(text) { const n = Number(String(text || "").trim().replace(",", ".")); return Number.isFinite(n) && n >= 20000 && n <= 80000; }
function isInvalidRiderName(name) {
  const value = String(name || "").trim();
  if (!value) return true;
  if (isHeaderText(value) || isTotalText(value) || isNumericLike(value) || isExcelSerialLike(value)) return true;
  if (/system\.xml|^vazia|^modelo$/i.test(value)) return true;
  return !/[a-zA-ZÀ-ÿ]{2,}/.test(value);
}

function parseMoney(raw) {
  if (raw == null || raw === "") return 0;
  let text = String(raw).replace(/\s/g, "").replace("R$", "");
  if (text.includes(",")) text = text.replace(/\./g, "").replace(",", ".");
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(text)) text = text.replace(/\./g, "");
  const value = Number(text);
  return Number.isFinite(value) ? value : 0;
}
function uniqueKey(parts) { return parts.map((part) => normalize(String(part ?? ""))).join("|"); }
function colName(index) {
  let n = index + 1, out = "";
  while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); }
  return out;
}
function moneyKey(value) { return Math.round(Number(value || 0) * 100) / 100; }

function getSheet(name) { return sheets.find((sheet) => sheet.name === name); }
function matrix(sheet) { return (Array.isArray(sheet?.rows) ? sheet.rows : []).map((row) => (row.cells || []).map((cell) => String(cell?.value ?? ""))); }
function isSystemSheet(name) { return ["financeiro", "relatorio de pagamento", "graficos", "empresa"].includes(normalize(name)); }
function isPartnerSheetName(name) { return Object.values(PARTNER_SHEETS).includes(name); }
function isRealRiderSheet(name) { return !isSystemSheet(name) && !isPartnerSheetName(name) && !isInvalidRiderName(name); }
function excelDate(serial) {
  const n = Number(serial);
  if (!Number.isFinite(n) || n < 20000 || n > 80000) return "";
  return new Date(Date.UTC(1899, 11, 30 + n)).toISOString().slice(0, 10);
}
function displayDate(date) {
  if (!date) return "";
  const [y, m, d] = date.split("-");
  return `${d}/${m}/${y}`;
}

function buildImportedModel() {
  const payments = paymentRecords();
  const riders = new Map();
  const daily = [];
  const baseEntries = [];
  const discounts = [];

  payments.forEach((p) => riders.set(normalize(p.name), { name: p.name, sheet: p.sheet, partner: isPartner(p.name, p.sheet) }));

  sheets.filter((sheet) => !isSystemSheet(sheet.name)).forEach((sheet) => {
    if (Object.values(PARTNER_SHEETS).includes(sheet.name)) {
      const partner = Object.keys(PARTNER_SHEETS).find((key) => PARTNER_SHEETS[key] === sheet.name);
      partnerBaseRows(sheet).forEach((row, index) => baseEntries.push({ id: `imp-base-${partner}-${index}`, source: "planilha", partner, ...row }));
      partnerDiscountRows(sheet, partner).forEach((row, index) => discounts.push({ id: `imp-partner-disc-${partner}-${index}`, source: "planilha", ...row }));
      return;
    }
    const payment = payments.find((p) => p.sheet === sheet.name);
    const rider = payment?.name || sheet.name;
    if (!riders.has(normalize(rider))) riders.set(normalize(rider), { name: rider, sheet: sheet.name, partner: false });
    deliveryRows(sheet).forEach((row, index) => {
      const ml = Number(row.ml || 0);
      const shopee = Number(row.shopee || 0);
      const gross = parseMoney(row.valueMl) + parseMoney(row.valueShopee);
      if (row.date && (ml || shopee || gross)) daily.push({ id: `imp-day-${sheet.name}-${index}`, source: "planilha", date: row.date, rider, ml, shopee, avulso: 0, rateMl: ml ? parseMoney(row.valueMl) / ml : 0, rateShopee: shopee ? parseMoney(row.valueShopee) / shopee : 0, rateAvulso: 0, gross, responsible: row.responsible || "", note: row.note || "" });
      const value = parseMoney(row.discountValue) + parseMoney(row.vale);
      if (value > 0) discounts.push({ id: `imp-disc-${sheet.name}-${index}`, source: "planilha", date: row.date || "2026-06-15", partner: partnerFromResponsible(row.responsible), rider, type: /extravio/i.test(row.occurrence) ? "Extravio" : parseMoney(row.vale) ? "Vale" : "Ocorrência", value, code: "", reason: row.occurrence || "Desconto importado", note: row.note || "" });
    });
  });

  payments.forEach((p) => {
    const current = discounts.filter((x) => normalize(x.rider) === normalize(p.name)).reduce((s, x) => s + x.value, 0);
    const missing = Math.max(0, p.discounts - current);
    if (missing > 0) discounts.push({ id: `imp-disc-report-${p.name}`, source: "relatorio", date: "2026-06-15", partner: "", rider: p.name, type: "Outro desconto", value: missing, code: "", reason: "Diferença importada do relatório de pagamento", note: "Desconto consolidado sem detalhe completo na planilha." });
  });

  return { riders: [...riders.values()].sort((a, b) => a.name.localeCompare(b.name, "pt-BR")), daily, baseEntries, discounts, payments };
}

function paymentRecords() {
  const rows = matrix(getSheet("RELATORIO DE PAGAMENTO"));
  const header = rows.findIndex((row) => row.some((cell) => normalize(cell) === "motorista"));
  if (header < 0) return [];
  return rows.slice(header + 1).filter((row) => row[0]).map((row) => ({ name: String(row[0]).trim(), packages: parseMoney(row[1]), gross: parseMoney(row[2]), discounts: parseMoney(row[3]), net: parseMoney(row[4]), statusOriginal: String(row[5] || ""), sheet: matchingSheetName(row[0]) }));
}

function buildImportedModelV2() {
  const payments = paymentRecordsV2();
  const riders = new Map();
  const daily = [];
  const baseEntries = [];
  const discounts = [];
  const ignored = [];
  const auditBlocks = [];
  const discountOrigins = [];
  const guilhermeAuditRows = [];
  const seenDaily = new Set();
  const seenBase = new Set();
  const seenDiscount = new Set();
  let duplicateCount = 0;

  payments.forEach((p) => riders.set(normalize(p.name), { name: p.name, sheet: p.sheet, partner: isPartner(p.name, p.sheet) }));
  sheets.filter((sheet) => isRealRiderSheet(sheet.name)).forEach((sheet) => {
    if (!riders.has(normalize(sheet.name))) riders.set(normalize(sheet.name), { name: sheet.name, sheet: sheet.name, partner: false });
  });

  sheets.filter((sheet) => !isSystemSheet(sheet.name)).forEach((sheet) => {
    if (isPartnerSheetName(sheet.name)) {
      const partner = Object.keys(PARTNER_SHEETS).find((key) => PARTNER_SHEETS[key] === sheet.name);
      partnerBaseRows(sheet).forEach((row, index) => {
        const key = uniqueKey([partner, row.date, row.ml, row.shopee]);
        if (seenBase.has(key)) { duplicateCount += 1; ignored.push(audit("Duplicidade removida", `Entrada base ${partner} ${displayDate(row.date)} ML ${row.ml} SH ${row.shopee}`)); return; }
        seenBase.add(key);
        baseEntries.push({ id: `imp-base-${partner}-${index}`, source: "planilha", importKey: key, partner, ...row });
      });
      const parsed = strictPartnerDiscountRows(sheet, partner, riders);
      parsed.rows.forEach((row, index) => {
        const key = row.importKey;
        if (seenDiscount.has(key)) {
          duplicateCount += 1;
          if (partner === "GUILHERME") parsed.auditRows.filter((x) => x.importKey === key).forEach((x) => x.status = "Duplicado");
          ignored.push(audit("Duplicidade removida", `Desconto ${row.sheetOriginal} ${row.rider} ${money(row.value)} ${row.origin}`));
          return;
        }
        seenDiscount.add(key);
        if (partner === "GUILHERME") parsed.auditRows.filter((x) => x.importKey === key).forEach((x) => x.status = "Importado");
        discounts.push({ id: `imp-partner-disc-${partner}-${index}`, source: "planilha", ...row });
        discountOrigins.push(row);
      });
      ignored.push(...parsed.ignored);
      auditBlocks.push(...parsed.blocks);
      if (partner === "GUILHERME") guilhermeAuditRows.push(...parsed.auditRows);
      return;
    }

    const payment = payments.find((p) => p.sheet === sheet.name);
    const rider = payment?.name || sheet.name;
    if (isInvalidRiderName(rider)) {
      ignored.push(audit("Fechamento ignorado por motoboy invalido", sheet.name));
      return;
    }

    deliveryRows(sheet).forEach((row, index) => {
      const ml = Number(row.ml || 0);
      const shopee = Number(row.shopee || 0);
      const gross = parseMoney(row.valueMl) + parseMoney(row.valueShopee);
      if (row.date && (ml || shopee || gross)) {
        const key = uniqueKey([sheet.name, row.date, rider, ml, shopee, row.valueMl, row.valueShopee, row.gross || gross]);
        if (seenDaily.has(key)) { duplicateCount += 1; ignored.push(audit("Duplicidade removida", `Saida ${sheet.name} ${displayDate(row.date)} ${rider}`)); }
        else {
          seenDaily.add(key);
          daily.push({ id: `imp-day-${sheet.name}-${index}`, source: "planilha", importKey: key, sheetOriginal: sheet.name, lineOriginal: row.lineOriginal || index + 1, date: row.date, rider, ml, shopee, avulso: 0, rateMl: ml ? parseMoney(row.valueMl) / ml : 0, rateShopee: shopee ? parseMoney(row.valueShopee) / shopee : 0, rateAvulso: 0, gross, responsible: row.responsible || "", note: row.note || "" });
        }
      }
      const value = parseMoney(row.discountValue) + parseMoney(row.vale);
      if (value > 0) {
        const partner = partnerFromDeliveryDiscount(row.responsible);
        const type = /extravio|ocorrencia|ocorrência/i.test(row.occurrence) ? "Extravio/Ocorrência" : parseMoney(row.vale) ? "Vale" : "Outros";
        const key = uniqueKey([partner, rider, type, moneyKey(value), row.lineOriginal || index + 1, row.columnOriginal || "", sheet.name]);
        if (seenDiscount.has(key)) { duplicateCount += 1; ignored.push(audit("Duplicidade removida", `Desconto ${sheet.name} ${rider} ${money(value)}`)); }
        else {
          seenDiscount.add(key);
          const origin = `${sheet.name}!${row.columnOriginal || "?"}${row.lineOriginal || index + 1}`;
          const alias = aliasLabel(row.responsible);
          const observation = [alias ? `Alias ${alias}` : "", row.occurrence || row.note || "Desconto importado da area de entregas"].filter(Boolean).join(" | ");
          const item = { id: `imp-disc-${sheet.name}-${index}`, source: "planilha", importKey: key, date: row.date || "2026-06-15", partner, rider, closingRider: rider, riderMatched: true, type, value, code: "", reason: row.occurrence || "Desconto importado", note: row.note || "", observation, sheetOriginal: sheet.name, lineOriginal: row.lineOriginal || index + 1, columnOriginal: row.columnOriginal || "", origin, responsibleOriginal: row.responsible || "" };
          discounts.push(item);
          discountOrigins.push(item);
        }
      }
    });
  });

  payments.forEach((p) => {
    const current = discounts.filter((x) => normalize(x.rider) === normalize(p.name)).reduce((s, x) => s + x.value, 0);
    const missing = Math.max(0, p.discounts - current);
    if (missing > 0) ignored.push(audit("Conferencia relatorio x itens de desconto", `${p.name}: relatório ${money(p.discounts)} x itens detalhados ${money(current)}. Diferença ${money(missing)} usada apenas para conferência, sem criar lançamento duplicado.`));
  });

  return { riders: [...riders.values()].sort((a, b) => a.name.localeCompare(b.name, "pt-BR")), daily, baseEntries, discounts, payments, ignored, auditBlocks, discountOrigins, duplicateCount, guilhermeAuditRows };
}

function paymentRecordsV2() {
  const rows = matrix(getSheet("RELATORIO DE PAGAMENTO"));
  const header = rows.findIndex((row) => row.some((cell) => normalize(cell) === "motorista"));
  if (header < 0) return [];
  const start = rows[header].findIndex((cell) => normalize(cell) === "motorista");
  return rows.slice(header + 1)
    .map((row, offset) => ({ row, lineOriginal: header + offset + 2 }))
    .filter(({ row }) => !isInvalidRiderName(row[start]))
    .map(({ row, lineOriginal }) => ({ name: String(row[start]).trim(), packages: parseMoney(row[start + 1]), gross: parseMoney(row[start + 2]), discounts: parseMoney(row[start + 3]), net: parseMoney(row[start + 4]), statusOriginal: String(row[start + 5] || ""), sheet: matchingSheetName(row[start]), lineOriginal, columnOriginal: colName(start + 3) }));
}

function officialFinanceiro() {
  const rows = matrix(getSheet("FINANCEIRO"));
  const byLabel = (label) => {
    const rowIndex = rows.findIndex((row) => normalize(row[1]) === normalize(label));
    return rowIndex >= 0 ? { row: rows[rowIndex], line: rowIndex + 1 } : null;
  };
  const general = rows[4] || [];
  const managers = {};
  ["GUILHERME", "SALES", "GIL"].forEach((name) => {
    const found = byLabel(name);
    if (found) managers[name] = {
      packages: parseMoney(found.row[2]),
      entryValue: parseMoney(found.row[3]),
      deliveredValue: parseMoney(found.row[4]),
      discounts: parseMoney(found.row[5]),
      balance: parseMoney(found.row[6]),
      origin: `FINANCEIRO!B${found.line}:G${found.line}`
    };
  });
  const cashLabels = [
    ["totalEntryPackages", "TOTAL ENTRADA PACOTES"],
    ["payBase", "PAGAR PARA A BASE (R$)"],
    ["grossDrivers", "FOLHA BRUTA MOTORISTAS"],
    ["totalDiscounts", "TOTAL DESCONTOS"],
    ["netPay", "VALOR A PAGAR LÍQUIDO"],
    ["finalCash", "VALOR CAIXA FINAL"]
  ];
  const cash = {};
  cashLabels.forEach(([key, label]) => {
    const found = byLabel(label);
    if (found) cash[key] = { value: parseMoney(found.row[4]), origin: `FINANCEIRO!E${found.line}` };
  });
  return {
    operational: {
      outputPackages: { value: parseMoney(general[1]), origin: "FINANCEIRO!B5" },
      grossPayroll: { value: parseMoney(general[2]), origin: "FINANCEIRO!C5" },
      discounts: { value: parseMoney(general[3]), origin: "FINANCEIRO!D5" },
      netPay: { value: parseMoney(general[4]), origin: "FINANCEIRO!E5" },
      entryPackages: { value: parseMoney(general[5]), origin: "FINANCEIRO!F5" },
      packageDiff: { value: parseMoney(general[6]), origin: "FINANCEIRO!G5" }
    },
    managers,
    cash
  };
}

function strictPartnerDiscountRows(sheet, partner, riders) {
  const rows = matrix(sheet);
  const out = [];
  const ignored = [];
  const blocks = [];
  const auditRows = [];
  const isGuilherme = partner === "GUILHERME";
  const pushGuilhermeAudit = (item) => {
    if (isGuilherme) auditRows.push({ sheet: sheet.name, ...item });
  };
  const header = rows.findIndex((row) => row.some((cell) => normalize(cell) === "controle de desconto"));
  if (header < 0) return { rows: out, ignored: [audit("Controle de desconto nao encontrado", sheet.name)], blocks, auditRows };
  const columns = [];
  rows[header + 1]?.forEach((cell, index, row) => {
    if (normalize(cell) === "nome") {
      let valeColumn = -1, occurrenceColumn = -1, valueColumn = -1;
      for (let j = index + 1; j < Math.min(index + 7, row.length); j += 1) {
        const label = normalize(row[j]);
        if (label.includes("vale")) valeColumn = j;
        if (label.includes("ocorrencia")) occurrenceColumn = j;
        if (label === "valor") valueColumn = j;
      }
      if (valueColumn > -1) columns.push({ start: index, valeColumn, occurrenceColumn, valueColumn });
    }
  });
  columns.forEach(({ start, valeColumn, occurrenceColumn, valueColumn }) => {
    let itemSum = 0;
    let importedCount = 0;
    let totalFound = null;
    rows.slice(header + 2).forEach((row, offset) => {
      const lineOriginal = header + offset + 3;
      const rawName = String(row[start] || "").trim();
      if (!rawName) {
        const possibleTotal = parseMoney(row[valueColumn]);
        if (possibleTotal > 0) {
          if (totalFound == null) totalFound = { value: possibleTotal, lineOriginal, columnOriginal: colName(valueColumn) };
          pushGuilhermeAudit({ lineOriginal, columnOriginal: colName(valueColumn), rider: "", type: "TOTAL/SUBTOTAL", value: possibleTotal, observation: "Celula sem motoboy usada como total/subtotal do bloco", status: "Ignorado", reason: "Total/subtotal do bloco" });
        }
        return;
      }
      if (isHeaderText(rawName)) { pushGuilhermeAudit({ lineOriginal, columnOriginal: colName(start), rider: rawName, type: "OUTROS", value: parseMoney(row[valueColumn]), observation: rawName, status: "Ignorado", reason: "Cabecalho" }); ignored.push(audit("Celula ignorada por ser cabecalho", `${sheet.name}: ${rawName}`)); return; }
      if (isTotalText(rawName)) { pushGuilhermeAudit({ lineOriginal, columnOriginal: colName(start), rider: rawName, type: "TOTAL/SUBTOTAL", value: parseMoney(row[valueColumn]), observation: rawName, status: "Ignorado", reason: "Total/subtotal" }); ignored.push(audit("Celula ignorada por ser total", `${sheet.name}: ${rawName}`)); return; }
      if (isNumericLike(rawName) || isExcelSerialLike(rawName)) { pushGuilhermeAudit({ lineOriginal, columnOriginal: colName(start), rider: rawName, type: "OUTROS", value: parseMoney(row[valueColumn]), observation: rawName, status: "Ignorado", reason: "Nome numerico/data" }); ignored.push(audit("Desconto ignorado por nome numerico", `${sheet.name}: ${rawName}`)); return; }
      const rider = canonicalRiderName(rawName, riders);
      const riderName = rider || rawName;
      const vale = valeColumn > -1 ? parseMoney(row[valeColumn]) : 0;
      const occurrence = occurrenceColumn > -1 ? parseMoney(row[occurrenceColumn]) : 0;
      const value = parseMoney(row[valueColumn]);
      const finalValue = value || vale || occurrence;
      if (!(finalValue > 0)) { pushGuilhermeAudit({ lineOriginal, columnOriginal: colName(valueColumn), rider: rawName, type: "OUTROS", value: finalValue, observation: rawName, status: "Ignorado", reason: "Valor zerado ou invalido" }); ignored.push(audit("Desconto ignorado por valor invalido", `${sheet.name}: ${rawName}`)); return; }
      const reason = String(row[start + 4] || "").trim();
      const type = vale > 0 && !occurrence ? "Vale" : occurrence > 0 || /extravio|ocorrencia|ocorrência/i.test(reason) ? "Extravio/Ocorrência" : "Outros";
      const origin = `${sheet.name}!${colName(valueColumn)}${lineOriginal}`;
      const observation = `${rawName} | origem ${origin}${reason ? ` | obs: ${reason}` : ""}`;
      const item = { date: "2026-06-15", partner, rider: riderName, closingRider: rider || "", riderMatched: Boolean(rider), type, value: finalValue, code: "", reason: reason || "Importado do controle de desconto", note: rawName, observation, sheetOriginal: sheet.name, lineOriginal, columnOriginal: colName(valueColumn), origin };
      item.importKey = uniqueKey([partner, item.rider, item.type, moneyKey(item.value), item.lineOriginal, item.columnOriginal, item.sheetOriginal]);
      pushGuilhermeAudit({ importKey: item.importKey, lineOriginal, columnOriginal: colName(valueColumn), rider: riderName, closingRider: rider || "", type, value: finalValue, observation, status: "Pendente", reason: "" });
      out.push(item);
      itemSum += finalValue;
      importedCount += 1;
    });
    blocks.push({ sheet: sheet.name, partner, startColumn: colName(start), valueColumn: colName(valueColumn), importedCount, itemSum, total: totalFound?.value ?? null, totalOrigin: totalFound ? `${sheet.name}!${totalFound.columnOriginal}${totalFound.lineOriginal}` : "" });
    if (totalFound && Math.abs(itemSum - totalFound.value) > 0.02) ignored.push(audit("Divergencia no bloco de desconto", `${sheet.name} coluna ${colName(start)}: itens ${money(itemSum)} x total ${money(totalFound.value)} (${totalFound.columnOriginal}${totalFound.lineOriginal})`));
  });
  return { rows: out, ignored, blocks, auditRows };
}

function canonicalRiderName(rawName, riders) {
  if (isInvalidRiderName(rawName)) return "";
  let key = normalize(rawName)
    .replace(/\bbolsao\b.*$/g, "")
    .replace(/\bfatan\b.*$/g, "")
    .replace(/\bhudson\b.*$/g, "")
    .replace(/\/.*$/g, "")
    .replace(/\b\d+\s*\/\s*\d+\b/g, "")
    .replace(/\b\d{2}\/\d{2}.*$/g, "")
    .trim();
  const aliases = { isac: "isaac", walison: "wallison cotonete", cotonete: "wallison cotonete", coliin: "guilherme collin", davi: "davi rocha", vando: "carlos evandro" };
  key = aliases[key] || key;
  if (!key || isInvalidRiderName(key)) return "";
  const names = [...riders.values()].map((r) => r.name).sort((a, b) => normalize(b).length - normalize(a).length);
  return names.find((name) => normalize(name) === key || key.includes(normalize(name)) || normalize(name).includes(key)) || "";
}

function matchingSheetName(name) {
  const key = normalize(name);
  const exact = sheets.find((sheet) => normalize(sheet.name) === key);
  if (exact) return exact.name;
  return sheets.find((sheet) => key.includes(normalize(sheet.name)) || normalize(sheet.name).includes(key))?.name || "";
}

function isPartner(name, sheetName = "") {
  return Boolean(managerAlias(name) || managerAlias(sheetName) || PARTNERS.some((p) => normalize(sheetName).includes(normalize(PARTNER_SHEETS[p]))));
}

function partnerFromResponsible(text) {
  return managerAlias(text) || "BASE";
}

function partnerFromDeliveryDiscount(text) {
  const key = normalize(text).replace(/[^a-z0-9]+/g, " ").trim();
  const alias = managerAlias(text);
  if (alias === "GUILHERME" && /\b(gm|g m|guilherme m|guilherme)\b/.test(key)) return "GUILHERME";
  return "BASE";
}

function aliasLabel(text) {
  const raw = String(text || "").trim();
  if (!raw) return "";
  const alias = managerAlias(raw);
  return alias && normalize(raw) !== normalize(alias) ? `${raw} -> ${alias}` : "";
}

function deliveryRows(sheet) {
  const rows = matrix(sheet);
  const header = rows.findIndex((row) => row.some((cell, i) => normalize(cell) === "data" && normalize(row[i + 1]).includes("saida ml")));
  if (header < 0) return [];
  const start = rows[header].findIndex((cell, i) => normalize(cell) === "data" && normalize(rows[header][i + 1]).includes("saida ml"));
  const out = [];
  for (let i = header + 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (normalize(row[start]).includes("quinzena")) break;
    out.push({ lineOriginal: i + 1, columnOriginal: colName(start), date: excelDate(row[start]), ml: parseMoney(row[start + 1]), shopee: parseMoney(row[start + 2]), valueMl: row[start + 3], valueShopee: row[start + 4], gross: row[start + 5], note: row[start + 6], vale: row[start + 8], occurrence: row[start + 9], discountValue: row[start + 10], responsible: row[start + 11] });
  }
  return out.filter((row) => row.date || row.ml || row.shopee || parseMoney(row.gross));
}

function partnerBaseRows(sheet) {
  const rows = matrix(sheet);
  const header = rows.findIndex((row) => row.some((cell, i) => normalize(cell) === "data" && normalize(row[i + 1]) === "ml" && normalize(row[i + 2]) === "sh"));
  if (header < 0) return [];
  const start = rows[header].findIndex((cell, i) => normalize(cell) === "data" && normalize(rows[header][i + 1]) === "ml" && normalize(rows[header][i + 2]) === "sh");
  const out = [];
  for (let i = header + 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (normalize(row[start]).includes("total de pacotes") || normalize(row[start]).includes("valor total")) break;
    const date = excelDate(row[start]);
    const ml = parseMoney(row[start + 1]);
    const shopee = parseMoney(row[start + 2]);
    if (date && (ml || shopee)) out.push(baseEntryCalc({ date, ml, shopee, note: "", responsible: "Planilha" }));
  }
  return out;
}

function partnerDiscountRows(sheet, partner) {
  const rows = matrix(sheet);
  const out = [];
  rows.forEach((row) => {
    for (let i = 0; i < row.length - 3; i += 1) {
      const name = String(row[i] || "").trim();
      const vale = parseMoney(row[i + 1]);
      const occurrence = parseMoney(row[i + 2]);
      const value = parseMoney(row[i + 3]);
      if (name && !/nome|data|total|geral/i.test(name) && (vale || occurrence || value)) out.push({ date: "2026-06-15", partner, rider: name, type: value && /extravio/i.test(String(row[i + 4] || "")) ? "Extravio" : vale ? "Vale" : "Ocorrência", value: value || vale || occurrence, code: "", reason: String(row[i + 4] || "Importado da aba do sócio"), note: "" });
    }
  });
  return out;
}

function baseEntryCalc(row) {
  const ml = Number(row.ml || 0);
  const shopee = Number(row.shopee || 0);
  return { ...row, ml, shopee, totalPackages: ml + shopee, valueMl: ml * BASE_RATE_ML, valueShopee: shopee * BASE_RATE_SHOPEE, totalPay: ml * BASE_RATE_ML + shopee * BASE_RATE_SHOPEE };
}

function uniqueRecords(list, keyFn) {
  const seen = new Set();
  return list.filter((item) => {
    const key = item.importKey || keyFn(item);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
function isoDateOnly(value) {
  if (!value) return "";
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const br = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (br) return `${br[3]}-${br[2]}-${br[1]}`;
  const d = new Date(`${text}T00:00:00`);
  return Number.isNaN(d.getTime()) ? text : d.toISOString().slice(0, 10);
}
function dailyIdentity(row = {}) { return uniqueKey([isoDateOnly(row.date), row.rider]); }
function allDaily() { return sortByRiderName(uniqueRecords([...imported.daily, ...state.daily].filter((x) => !isInvalidRiderName(x.rider)), dailyIdentity)); }
function allBaseEntries() { return uniqueRecords([...imported.baseEntries, ...state.baseEntries], (x) => uniqueKey([x.partner, x.date, x.ml, x.shopee])); }
function isValidDiscountLabel(label) {
  const value = String(label || "").trim();
  return Boolean(value) && !isNumericLike(value) && !isExcelSerialLike(value) && !isTotalText(value) && !/^(nome|data|vale|valor|motorista)$/i.test(value);
}
function allDiscounts() { return sortByRiderName(uniqueRecords([...imported.discounts, ...state.discounts].filter((x) => isValidDiscountLabel(x.rider) && Number(x.value) > 0), (x) => uniqueKey([x.partner || "BASE", x.rider, x.type, moneyKey(x.value), x.lineOriginal, x.columnOriginal, x.sheetOriginal || x.origin || "manual"]))); }
function allRiders() { return sortByRiderName((state.riders || []).filter((x) => !isInvalidRiderName(x.name))); }
function activeRiders() { return allRiders().filter((x) => x.active !== false); }
function riderByName(name) { return allRiders().find((x) => normalize(x.name) === normalize(name)); }
function riderWorkType(rider) {
  const type = normalize(rider?.collection || "com coleta");
  if (type.includes("freelancer")) return "freelancer";
  if (type.includes("sem coleta")) return "sem coleta";
  return "com coleta";
}
function isFreelancer(rider) { return riderWorkType(rider) === "freelancer"; }
function isNoCollection(rider) { return riderWorkType(rider) === "sem coleta"; }
function isProfitRider(rider) { const type = riderWorkType(rider); return type === "sem coleta" || type === "freelancer"; }
function workTypeLabel(type) { return type === "freelancer" ? "Freelancer" : type === "sem coleta" ? "Sem coleta" : "Com coleta"; }
function workTypeClass(type) { return `type-${riderWorkType({ collection: type }).replace(/\s+/g, "-")}`; }
function workTypeInfo(type) {
  const key = riderWorkType({ collection: type });
  if (key === "freelancer") return { label: "Freelancer", description: "Motoboy que trabalha apenas em dias específicos, conforme demanda da operação.", ml: DEFAULT_NO_COLLECTION_ML, shopee: DEFAULT_NO_COLLECTION_SHOPEE };
  if (key === "sem coleta") return { label: "Sem coleta", description: "Motoboy fixo que não faz coleta, apenas entrega os pacotes.", ml: DEFAULT_NO_COLLECTION_ML, shopee: DEFAULT_NO_COLLECTION_SHOPEE };
  return { label: "Com coleta", description: "Motoboy fixo que faz coleta e entrega.", ml: BASE_RATE_ML, shopee: BASE_RATE_SHOPEE };
}
function workTypeBadge(type) {
  const info = workTypeInfo(type);
  return `<span class="type-badge ${workTypeClass(type)}">${escapeHtml(info.label)}</span>`;
}
function workTypeCardsHtml(selected = "com coleta") {
  return ["com coleta", "sem coleta", "freelancer"].map((type) => {
    const info = workTypeInfo(type);
    return `<button type="button" class="work-type-card ${workTypeClass(type)} ${riderWorkType({ collection: selected }) === type ? "active" : ""}" data-work-type="${type}"><span class="work-type-icon" aria-hidden="true"></span><strong>${escapeHtml(info.label)}</strong></button>`;
  }).join("");
}
function discountsByType(type) { return allDiscounts().filter((x) => normalize(x.type) === normalize(type)); }
function rawManagerDiscount(partner) { return sum(imported.discounts.filter((x) => (x.partner || "BASE") === partner), "value"); }
function managerFinancialDiscountItems(partner) {
  const sheetName = PARTNER_SHEETS[partner];
  return imported.discounts.filter((x) => (x.partner || "BASE") === partner && x.sheetOriginal === sheetName);
}
function effectiveManagerDiscount(partner) {
  const scoped = managerFinancialDiscountItems(partner);
  return scoped.length ? sum(scoped, "value") : rawManagerDiscount(partner);
}
function importedManagerDiscountItems(partner) { return imported.discounts.filter((x) => (x.partner || "BASE") === partner); }
function guilhermeProfileKey(item) {
  return uniqueKey([
    item.partner || "BASE",
    item.rider || "",
    moneyKey(item.value),
    item.type || "",
    item.observation || item.reason || item.note || "",
    item.sheetOriginal || "",
    item.lineOriginal || ""
  ]);
}
function guilhermeProfileAudit() {
  const source = importedManagerDiscountItems("GUILHERME");
  const officialSheet = PARTNER_SHEETS.GUILHERME;
  const seen = new Set();
  let duplicates = 0;
  let validTotal = 0;
  const rows = source.map((item) => {
    const key = guilhermeProfileKey(item);
    const isDuplicate = seen.has(key);
    if (!isDuplicate) seen.add(key);
    const isOfficial = item.sheetOriginal === officialSheet;
    let status = "Registro unico";
    if (isDuplicate) {
      duplicates += 1;
      status = "Registro duplicado removido";
    } else if (isOfficial) {
      validTotal += Number(item.value || 0);
      status = "Registro valido";
    }
    return { ...item, profileStatus: status, profileIncluded: !isDuplicate && isOfficial };
  });
  return {
    rows,
    validRows: rows.filter((x) => x.profileIncluded),
    found: rows.length,
    valid: rows.filter((x) => x.profileIncluded).length,
    duplicates,
    before: sum(source, "value"),
    after: validTotal,
    sheetTotal: effectiveManagerDiscount("GUILHERME")
  };
}
function partnerProfileDiscounts(partner) {
  if (partner !== "GUILHERME") return allDiscounts().filter((x) => (x.partner || "BASE") === partner);
  return guilhermeProfileAudit().validRows;
}
function variableExpensesByResponsible(responsible) {
  const owner = normalizeResponsible(responsible);
  return allExpenses().filter((x) => x.type === "variavel" && normalizeResponsible(x.responsible) === owner);
}
function variableExpenseTotal(responsible) {
  return sum(variableExpensesByResponsible(responsible), "value");
}
function guilhermeProfileAuditTableRows() {
  return guilhermeProfileAudit().rows.map((x) => [
    x.lineOriginal || "",
    x.sheetOriginal || "",
    x.rider || "",
    x.type || "",
    money(x.value),
    x.observation || x.reason || x.note || "",
    x.profileStatus || ""
  ]);
}
function aliasAuditRows() {
  return (imported.discountOrigins || imported.discounts)
    .filter((x) => aliasLabel(x.responsibleOriginal || ""))
    .map((x) => [aliasLabel(x.responsibleOriginal), x.partner || "BASE", x.rider, x.type, money(x.value), x.origin || `${x.sheetOriginal || ""}!${x.columnOriginal || ""}${x.lineOriginal || ""}`, x.observation || x.reason || ""]);
}

function guilhermeAuditSummary() {
  const rows = imported.guilhermeAuditRows || [];
  const importedRows = rows.filter((x) => x.status === "Importado");
  return {
    found: rows.length,
    assigned: importedRows.length,
    linked: importedRows.filter((x) => x.closingRider).length,
    withoutRider: importedRows.filter((x) => !x.closingRider).length,
    total: sum(importedRows, "value"),
    duplicates: rows.filter((x) => x.status === "Duplicado").length,
    ignored: rows.filter((x) => x.status === "Ignorado").length
  };
}

function guilhermeAuditTableRows() {
  return (imported.guilhermeAuditRows || []).map((x) => [
    x.sheet || "",
    x.lineOriginal || "",
    x.columnOriginal || "",
    x.rider || "",
    money(x.value),
    x.type || "Outros",
    x.observation || x.reason || "",
    x.status || "",
    x.reason || ""
  ]);
}

function syncReportBlocks() {
  if (!syncReport) return "";
  const summaryRows = [
    ["Arquivo anterior", syncReport.previousSourceFile || ""],
    ["Arquivo oficial atual", syncReport.sourceFile || workbook.sourceFile || ""],
    ["Abas analisadas", num(syncReport.sheetCountNew || workbook.sheetCount)],
    ["Registros existentes no app", num(syncReport.existingRecords || 0)],
    ["Registros encontrados na nova planilha", num(syncReport.foundRecords || 0)],
    ["Novos importados", num(syncReport.newRecords || 0)],
    ["Atualizados", num(syncReport.updatedRecords || 0)],
    ["Removidos da nova referencia", num(syncReport.removedRecords || 0)],
    ["Iguais", num(syncReport.unchangedRecords || 0)],
    ["Ignorados", num(syncReport.ignoredRecords || 0)],
    ["Duplicidades evitadas", num(syncReport.duplicatesAvoided || 0)]
  ];
  const changedRows = (syncReport.changesBySheet || []).map((x) => [x.sheet, num(x.added || 0), num(x.updated || 0), num(x.removed || 0)]);
  const before = syncReport.officialTotalsBefore || {};
  const after = syncReport.officialTotalsAfter || {};
  const totalRows = ["GUILHERME", "SALES", "GIL", "TOTAL"].map((name) => [
    name,
    money(before[name]?.discounts || 0),
    money(after[name]?.discounts || 0),
    money((after[name]?.discounts || 0) - (before[name]?.discounts || 0)),
    money(before[name]?.balance || 0),
    money(after[name]?.balance || 0)
  ]);
  if (before.operational || after.operational) {
    totalRows.push(["Resumo Operacional", money(before.operational?.discounts || 0), money(after.operational?.discounts || 0), money((after.operational?.discounts || 0) - (before.operational?.discounts || 0)), money(before.operational?.netPay || 0), money(after.operational?.netPay || 0)]);
  }
  return `${tableBlock("Relatorio de Sincronizacao", ["Item","Valor"], summaryRows)}${tableBlock("Abas alteradas na sincronizacao", ["Aba","Novos","Atualizados","Removidos"], changedRows)}${tableBlock("Totais oficiais antes x depois", ["Gestor/Resumo","Descontos antes","Descontos depois","Diferenca","Saldo/Liquido antes","Saldo/Liquido depois"], totalRows)}`;
}

function periodKey(date, type = "quinzenal") {
  const d = new Date(`${date}T00:00:00`);
  if (Number.isNaN(d.getTime())) return { key: "sem-periodo", label: "Sem período", start: "", end: "" };
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  if (type === "semanal") {
    const startDate = new Date(d); startDate.setDate(d.getDate() - d.getDay());
    const endDate = new Date(startDate); endDate.setDate(startDate.getDate() + 6);
    const start = startDate.toISOString().slice(0, 10), end = endDate.toISOString().slice(0, 10);
    return { key: `${start}|${end}`, label: `${displayDate(start)} a ${displayDate(end)}`, start, end };
  }
  const first = d.getDate() <= 15;
  const start = `${y}-${m}-${first ? "01" : "16"}`;
  const end = `${y}-${m}-${String(first ? 15 : new Date(y, d.getMonth() + 1, 0).getDate()).padStart(2, "0")}`;
  return { key: `${start}|${end}`, label: `${first ? "1ª" : "2ª"} quinzena ${m}/${y}`, start, end };
}

function normalizeResponsible(value) {
  const raw = normalize(value).replace(/\./g, "");
  if (raw === "gm" || raw === "guilherme" || raw === "guilherme m") return "GUILHERME";
  if (raw === "gil") return "GIL";
  if (raw === "sales") return "SALES";
  if (raw === "operador" || raw === "operator") return "OPERADOR";
  return "BASE";
}

function isoDate(y, m, d) {
  return `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

function variableExpenseDiscountPeriod(date) {
  const d = new Date(`${date}T00:00:00`);
  if (Number.isNaN(d.getTime())) return periodKey(date || new Date().toISOString().slice(0, 10));
  const y = d.getFullYear();
  const m = d.getMonth();
  return d.getDate() <= 15 ? periodKey(isoDate(y, m, 16)) : periodKey(isoDate(m === 11 ? y + 1 : y, m === 11 ? 0 : m + 1, 1));
}

function normalizeExpenseType(value) {
  return normalize(value).includes("fix") ? "fixa" : "variavel";
}

function nextQuinzenalPeriod(period) {
  if (!period?.start) return variableExpenseDiscountPeriod(new Date().toISOString().slice(0, 10));
  const d = new Date(`${period.start}T00:00:00`);
  if (Number.isNaN(d.getTime())) return variableExpenseDiscountPeriod(new Date().toISOString().slice(0, 10));
  return d.getDate() <= 1 ? periodKey(isoDate(d.getFullYear(), d.getMonth(), 16)) : periodKey(isoDate(d.getMonth() === 11 ? d.getFullYear() + 1 : d.getFullYear(), d.getMonth() === 11 ? 0 : d.getMonth() + 1, 1));
}

function variableExpensePeriodOptions(date) {
  const base = variableExpenseDiscountPeriod(date || new Date().toISOString().slice(0, 10));
  const periods = [periodKey(date || new Date().toISOString().slice(0, 10)), base, nextQuinzenalPeriod(base), nextQuinzenalPeriod(nextQuinzenalPeriod(base))];
  return uniqueRecords(periods.filter((x) => x.key !== "sem-periodo"), (x) => x.key);
}

function periodFromKey(key = "", label = "") {
  const [start = "", end = ""] = String(key || "").split("|");
  return { key, label: label || String(key || "").replace("|", " a "), start, end };
}

function enrichExpense(row) {
  const type = normalizeExpenseType(row.type || "fixa");
  const responsible = normalizeResponsible(row.responsible);
  if (type !== "variavel") return { ...row, responsible, originPeriodKey: "", originPeriodLabel: "", discountPeriodKey: "", discountPeriodLabel: "", discountPeriodManual: false };
  const origin = periodKey(row.date, "quinzenal");
  const automaticDiscount = variableExpenseDiscountPeriod(row.date);
  const manualDiscount = row.discountPeriodManual && row.discountPeriodKey ? periodFromKey(row.discountPeriodKey, row.discountPeriodLabel) : null;
  const discount = manualDiscount || automaticDiscount;
  return { ...row, type, responsible, originPeriodKey: origin.key, originPeriodLabel: origin.label, discountPeriodKey: discount.key, discountPeriodLabel: discount.label, discountPeriodManual: Boolean(manualDiscount && manualDiscount.key !== automaticDiscount.key) };
}

function allExpenses() {
  return (state.expenses || []).map(enrichExpense);
}

function currentDashboardPeriod() {
  const selected = $("periodFilter")?.value || "";
  if (selected) {
    const [start = "", end = ""] = selected.split("|");
    return { key: selected, label: selected.replace("|", " a "), start, end };
  }
  return periodKey(new Date().toISOString().slice(0, 10), "quinzenal");
}

function variableExpenseSummary(period = currentDashboardPeriod()) {
  const rows = allExpenses().filter((x) => x.type === "variavel");
  const next = nextQuinzenalPeriod(period);
  return {
    period,
    next,
    periodCosts: sum(rows.filter((x) => x.originPeriodKey === period.key), "value"),
    currentDiscount: sum(rows.filter((x) => x.discountPeriodKey === period.key), "value"),
    nextPending: sum(rows.filter((x) => x.discountPeriodKey === next.key && x.status !== "pago"), "value")
  };
}

function closingRecords(type = "quinzenal") {
  const groups = new Map();
  const ensure = (rider, date) => {
    if (isInvalidRiderName(rider)) return null;
    const p = periodKey(date || "2026-06-15", type);
    const key = `${normalize(rider)}|${p.key}`;
    if (!groups.has(key)) groups.set(key, { id: key, rider, period: p.label, start: p.start, end: p.end, ml: 0, shopee: 0, avulso: 0, gross: 0, vales: 0, losses: 0, discounts: 0, bonuses: 0, net: 0, status: "pendente" });
    return groups.get(key);
  };
  allDaily().forEach((x) => { const row = ensure(x.rider, x.date); if (!row) return; row.ml += x.ml; row.shopee += x.shopee; row.avulso += x.avulso; row.gross += x.gross; });
  allDiscounts().forEach((x) => {
    const rider = x.closingRider || (x.source === "manual" ? x.rider : "");
    if (!rider) return;
    const row = ensure(rider, x.date);
    if (!row) return;
    if (x.type === "Vale") row.vales += x.value;
    else if (normalize(x.type).includes("extravio") || normalize(x.type).includes("ocorrencia")) row.losses += x.value;
    else row.discounts += x.value;
  });
  imported.payments.forEach((p) => { if (isInvalidRiderName(p.name)) return; if (![...groups.values()].some((g) => normalize(g.rider) === normalize(p.name))) { const row = ensure(p.name, "2026-06-15"); if (!row) return; row.ml = p.packages; row.gross = p.gross; row.discounts = p.discounts; } });
  return [...groups.values()].map((row) => ({ ...row, net: row.gross - row.vales - row.losses - row.discounts + row.bonuses, status: state.paid[row.id] ? "pago" : row.status, paymentDate: state.paid[row.id]?.date || "" })).sort(compareRiderName);
}

function baseClosingRecords() {
  const groups = new Map();
  const ensure = (partner, period) => {
    const key = `${partner}|${period.key}`;
    if (!groups.has(key)) groups.set(key, { id: key, partner, period: period.label, start: period.start, end: period.end, ml: 0, shopee: 0, totalPackages: 0, valueMl: 0, valueShopee: 0, totalPay: 0, managerDiscounts: 0, variablePeriodCosts: 0, variableDiscounts: 0, netAfterVariable: 0, status: "pendente" });
    return groups.get(key);
  };
  allBaseEntries().forEach((x) => {
    const p = periodKey(x.date, "quinzenal");
    const row = ensure(x.partner, p);
    row.ml += x.ml; row.shopee += x.shopee; row.totalPackages += x.totalPackages; row.valueMl += x.valueMl; row.valueShopee += x.valueShopee; row.totalPay += x.totalPay;
  });
  allDiscounts().forEach((discount) => {
    const responsible = discount.partner || "BASE";
    const p = periodKey(discount.date, "quinzenal");
    ensure(responsible, p).managerDiscounts += Number(discount.value || 0);
  });
  allExpenses().filter((x) => x.type === "variavel").forEach((expense) => {
    const responsible = normalizeResponsible(expense.responsible);
    const origin = expense.originPeriodKey ? { key: expense.originPeriodKey, label: expense.originPeriodLabel || expense.originPeriodKey.replace("|", " a "), start: expense.originPeriodKey.split("|")[0] || "", end: expense.originPeriodKey.split("|")[1] || "" } : null;
    const discount = expense.discountPeriodKey ? { key: expense.discountPeriodKey, label: expense.discountPeriodLabel || expense.discountPeriodKey.replace("|", " a "), start: expense.discountPeriodKey.split("|")[0] || "", end: expense.discountPeriodKey.split("|")[1] || "" } : null;
    if (origin) ensure(responsible, origin).variablePeriodCosts += Number(expense.value || 0);
    if (discount) ensure(responsible, discount).variableDiscounts += Number(expense.value || 0);
  });
  return [...groups.values()].map((row) => ({ ...row, netAfterVariable: row.totalPay - row.managerDiscounts - row.variableDiscounts, status: state.basePaid[row.id] ? "pago" : row.status }));
}

function profitReportRows() {
  return allRiders().filter(isProfitRider).map((rider) => {
    const rows = allDaily().filter((x) => normalize(x.rider) === normalize(rider.name));
    const ml = sum(rows, "ml");
    const shopee = sum(rows, "shopee");
    const baseReceivedMl = ml * BASE_RATE_ML;
    const baseReceivedShopee = shopee * BASE_RATE_SHOPEE;
    const paidMl = ml * Number(rider.rateMl ?? DEFAULT_NO_COLLECTION_ML);
    const paidShopee = shopee * Number(rider.rateShopee ?? DEFAULT_NO_COLLECTION_SHOPEE);
    const closing = closingRecords().find((x) => normalize(x.rider) === normalize(rider.name));
    return {
      id: rider.id,
      name: rider.name,
      type: riderWorkType(rider),
      ml,
      shopee,
      packages: ml + shopee,
      baseReceivedMl,
      baseReceivedShopee,
      baseReceived: baseReceivedMl + baseReceivedShopee,
      paidMl,
      paidShopee,
      paid: paidMl + paidShopee,
      profitMl: baseReceivedMl - paidMl,
      profitShopee: baseReceivedShopee - paidShopee,
      profit: baseReceivedMl + baseReceivedShopee - paidMl - paidShopee,
      status: closing?.status || "pendente",
      closingId: closing?.id || ""
    };
  });
}

function profitTotals() {
  const rows = profitReportRows();
  const noCollection = rows.filter((x) => x.type === "sem coleta");
  const freelancers = rows.filter((x) => x.type === "freelancer");
  return {
    noCollectionProfit: sum(noCollection, "profit"),
    freelancerProfit: sum(freelancers, "profit"),
    totalProfit: sum(rows, "profit"),
    mlPackages: sum(rows, "ml"),
    shopeePackages: sum(rows, "shopee"),
    freelancerCount: allRiders().filter(isFreelancer).length,
    freelancerPaid: sum(freelancers, "paid"),
    freelancerPackages: sum(freelancers, "packages")
  };
}

function dashboardTotals() {
  const closings = filteredClosings();
  const bases = baseClosingRecords();
  const profits = profitTotals();
  const expenses = allExpenses();
  const fixedExpenses = sum(expenses.filter((x) => x.type === "fixa"), "value");
  const variableExpenses = sum(expenses.filter((x) => x.type === "variavel"), "value");
  const variableSummary = variableExpenseSummary();
  const totalExpenses = fixedExpenses + variableExpenses;
  const baseMl = bases.reduce((s, x) => s + x.ml, 0);
  const baseShopee = bases.reduce((s, x) => s + x.shopee, 0);
  const outMl = closings.reduce((s, x) => s + x.ml, 0);
  const outShopee = closings.reduce((s, x) => s + x.shopee, 0);
  const outAvulso = closings.reduce((s, x) => s + x.avulso, 0);
  return {
    baseMl,
    baseShopee,
    basePackages: baseMl + baseShopee,
    outMl,
    outShopee,
    outAvulso,
    outPackages: outMl + outShopee + outAvulso,
    packageDiff: (baseMl + baseShopee) - (outMl + outShopee + outAvulso),
    basePay: bases.reduce((s, x) => s + x.totalPay, 0),
    gross: closings.reduce((s, x) => s + x.gross, 0),
    vales: closings.reduce((s, x) => s + x.vales, 0),
    losses: closings.reduce((s, x) => s + x.losses, 0),
    discounts: closings.reduce((s, x) => s + x.discounts, 0),
    net: closings.reduce((s, x) => s + x.net, 0),
    paid: closings.filter((x) => x.status === "pago").reduce((s, x) => s + x.net, 0),
    pending: closings.filter((x) => x.status !== "pago").reduce((s, x) => s + x.net, 0),
    fixedExpenses,
    variableExpenses,
    variablePeriodCosts: variableSummary.periodCosts,
    variableCurrentDiscount: variableSummary.currentDiscount,
    variableNextPending: variableSummary.nextPending,
    totalExpenses,
    noCollectionProfit: profits.noCollectionProfit,
    freelancerProfit: profits.freelancerProfit,
    profitTotal: profits.totalProfit,
    profitMlPackages: profits.mlPackages,
    profitShopeePackages: profits.shopeePackages,
    freelancerCount: profits.freelancerCount,
    freelancerPaid: profits.freelancerPaid,
    freelancerPackages: profits.freelancerPackages,
    finalResult: bases.reduce((s, x) => s + x.totalPay, 0) - closings.reduce((s, x) => s + x.net, 0) - totalExpenses,
    riders: allRiders().length
  };
}

function packageComparisonRows() {
  const t = dashboardTotals();
  const rows = [
    { label: "Mercado Livre", entry: t.baseMl, exit: t.outMl },
    { label: "Shopee", entry: t.baseShopee, exit: t.outShopee },
    { label: "Total geral", entry: t.basePackages, exit: t.outPackages }
  ];
  return rows.map((row) => {
    const diff = row.entry - row.exit;
    const pct = row.entry ? (row.exit / row.entry) * 100 : 0;
    const status = diff === 0 ? "OK" : diff > 0 ? "Atenção" : "Verificar";
    return { ...row, diff, pct, status };
  });
}

function appFinancialTotals() {
  const payments = imported.payments || [];
  const base = allBaseEntries();
  const managerDiscounts = RESPONSIBLES.reduce((acc, p) => {
    acc[p] = effectiveManagerDiscount(p);
    return acc;
  }, {});
  const managerBalances = PARTNERS.reduce((acc, p) => {
    const entry = sum(base.filter((x) => x.partner === p), "totalPay");
    const delivered = managerDeliveredValue(p);
    const discounts = managerDiscounts[p] || 0;
    const variable = variableExpenseTotal(p);
    acc[p] = entry - delivered - discounts - variable;
    return acc;
  }, {});
  const payBase = Object.values(managerBalances).reduce((s, x) => s + x, 0);
  return {
    outputPackages: sum(payments, "packages"),
    grossPayroll: sum(payments, "gross"),
    discounts: sum(payments, "discounts"),
    netPay: sum(payments, "net"),
    entryPackages: sum(base, "totalPackages"),
    packageDiff: sum(base, "totalPackages") - sum(payments, "packages"),
    payBase,
    finalCash: payBase - sum(payments, "net"),
    managerDiscounts,
    managerBalances
  };
}

function managerDeliveredValue(partner) {
  const sheetName = PARTNER_SHEETS[partner];
  const sheet = getSheet(sheetName);
  return deliveryRows(sheet).reduce((s, x) => s + parseMoney(x.gross), 0);
}

function officialComparisons() {
  const official = officialFinanceiro();
  const app = appFinancialTotals();
  const comparisons = [
    { label: "Saída de Pacotes", app: app.outputPackages, sheet: official.operational.outputPackages.value, origin: official.operational.outputPackages.origin, format: "num", reason: "Soma de QTD DE PACOTES do RELATORIO DE PAGAMENTO." },
    { label: "Folha Bruta", app: app.grossPayroll, sheet: official.operational.grossPayroll.value, origin: official.operational.grossPayroll.origin, format: "money", reason: "Soma de VALOR BRUTO do RELATORIO DE PAGAMENTO." },
    { label: "Total Descontos", app: app.discounts, sheet: official.operational.discounts.value, origin: official.operational.discounts.origin, format: "money", reason: "Soma de DESCONTO do RELATORIO DE PAGAMENTO." },
    { label: "Valor Líquido", app: app.netPay, sheet: official.operational.netPay.value, origin: official.operational.netPay.origin, format: "money", reason: "Soma de VALOR LIQUIDO do RELATORIO DE PAGAMENTO." },
    { label: "Entrada de Pacotes", app: app.entryPackages, sheet: official.operational.entryPackages.value, origin: official.operational.entryPackages.origin, format: "num", reason: "Soma de entradas ML + Shopee dos gestores." },
    { label: "Diferença de Pacotes", app: app.packageDiff, sheet: official.operational.packageDiff.value, origin: official.operational.packageDiff.origin, format: "num", reason: "Entrada de pacotes menos saída de pacotes." },
    { label: "Pagar para Base", app: app.payBase, sheet: official.cash.payBase.value, origin: official.cash.payBase.origin, format: "money", reason: "Soma dos saldos dos gestores." },
    { label: "Valor Caixa Final", app: app.finalCash, sheet: official.cash.finalCash.value, origin: official.cash.finalCash.origin, format: "money", reason: "Pagar para Base menos valor líquido a pagar." }
  ];
  Object.entries(official.managers).forEach(([name, row]) => {
    const managerBase = allBaseEntries().filter((x) => x.partner === name);
    const appPackages = sum(managerBase, "totalPackages");
    const appEntry = sum(managerBase, "totalPay");
    const appDelivered = managerDeliveredValue(name);
    const appDiscounts = effectiveManagerDiscount(name);
    const appVariable = variableExpenseTotal(name);
    const appBalance = appEntry - appDelivered - appDiscounts - appVariable;
    comparisons.push({ label: `${name} pacotes`, app: appPackages, sheet: row.packages, origin: row.origin, format: "num", reason: "Soma das entradas do gestor." });
    comparisons.push({ label: `${name} valor entrada`, app: appEntry, sheet: row.entryValue, origin: row.origin, format: "money", reason: "ML x R$ 8,00 + Shopee x R$ 5,00." });
    comparisons.push({ label: `${name} valor entregue`, app: appDelivered, sheet: row.deliveredValue, origin: row.origin, format: "money", reason: "Soma de TOTAL DIA na seção ENTREGAS do gestor." });
    comparisons.push({ label: `${name} descontos`, app: appDiscounts, sheet: row.discounts, origin: row.origin, format: "money", reason: "Soma dos itens reais dos blocos CONTROLE DE DESCONTO." });
    comparisons.push({ label: `${name} saldo gestor`, app: appBalance, sheet: row.balance, origin: row.origin, format: "money", reason: "Valor entrada - valor entregue - descontos - despesas variaveis." });
  });
  return comparisons;
}

function filteredClosings() {
  const q = normalize($("quickSearch")?.value || "");
  const status = $("statusFilter")?.value || "";
  const period = $("periodFilter")?.value || "";
  return closingRecords($("closingType")?.value || "quinzenal").filter((x) => normalize(`${x.rider} ${x.period} ${x.status}`).includes(q) && (!status || x.status === status) && (!period || `${x.start}|${x.end}` === period));
}

function card(label, value) { return `<article class="card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></article>`; }
function workTypeSummaryCards() {
  const profit = profitTotals();
  const counts = {
    "com coleta": allRiders().filter((r) => riderWorkType(r) === "com coleta").length,
    "sem coleta": allRiders().filter((r) => riderWorkType(r) === "sem coleta").length,
    freelancer: allRiders().filter((r) => riderWorkType(r) === "freelancer").length
  };
  return ["com coleta", "sem coleta", "freelancer"].map((type) => {
    const info = workTypeInfo(type);
    const detail = type === "sem coleta" ? `Lucro ${money(profit.noCollectionProfit)}` : type === "freelancer" ? `Lucro ${money(profit.freelancerProfit)}` : "Lucro R$ 0,00";
    return `<article class="card type-chip ${workTypeClass(type)}"><span>${workTypeBadge(type)} ${escapeHtml(info.description)}</span><strong>${num(counts[type])} cadastro(s)</strong><small>ML ${money(info.ml)} | Shopee ${money(info.shopee)} | ${detail}</small></article>`;
  }).join("");
}
function sum(list, key) { return list.reduce((s, x) => s + Number(x[key] || 0), 0); }

function renderAll() {
  renderOptions(); renderDashboardV2(); renderMotoboys(); renderFreelancers(); renderPartners(); renderBase(); renderDaily(); renderDiscounts(); renderExpenses(); renderClosings(); renderReceipts(); renderReports(); renderBackup(); renderConfig(); renderUsers(); applyPermissions();
}

function renderOptions() {
  const currentPeriod = $("periodFilter").value || "";
  const riders = activeRiders().map((r) => `<option>${escapeHtml(r.name)}</option>`).join("");
  ["dailyRider", "discountRider", "receiptRider"].forEach((id) => $(id).innerHTML = riders);
  $("closingRider").innerHTML = `<option value="">Todos</option>${riders}`;
  $("periodFilter").innerHTML = `<option value="">Todos os períodos</option>${[...new Set(closingRecords().map((c) => `${c.start}|${c.end}`))].map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p.replace("|", " a "))}</option>`).join("")}`;
  const periodOptions = [
    ...closingRecords().map((c) => `${c.start}|${c.end}`),
    ...baseClosingRecords().map((c) => `${c.start}|${c.end}`),
    ...allExpenses().flatMap((x) => [x.originPeriodKey, x.discountPeriodKey])
  ].filter(Boolean);
  $("periodFilter").innerHTML = `<option value="">Todos os periodos</option>${[...new Set(periodOptions)].map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p.replace("|", " a "))}</option>`).join("")}`;
  if ([...$("periodFilter").options].some((option) => option.value === currentPeriod)) $("periodFilter").value = currentPeriod;
}

function renderDashboard() {
  const t = dashboardTotals();
  $("dashboardCards").innerHTML = [card("Total ML entrada", num(t.baseMl)), card("Total Shopee entrada", num(t.baseShopee)), card("Pacotes entrada", num(t.basePackages)), card("A pagar para bases", money(t.basePay)), card("Bruto motoboys", money(t.gross)), card("Vales", money(t.vales)), card("Extravios", money(t.losses)), card("Descontos", money(t.discounts)), card("Líquido motoboys", money(t.net)), card("Pago", money(t.paid)), card("Pendente", money(t.pending)), card("Motoboys", num(t.riders))].join("");
  $("rankingRows").innerHTML = filteredClosings().sort((a, b) => (b.ml + b.shopee + b.avulso) - (a.ml + a.shopee + a.avulso)).map((c, i) => `<tr><td>${i + 1}</td><td>${escapeHtml(c.rider)}</td><td>${num(c.ml + c.shopee + c.avulso)}</td><td>${money(c.gross)}</td><td>${money(c.net)}</td></tr>`).join("");
  $("alertRows").innerHTML = validations().map(alertHtml).join("") || `<div class="alert ok">Nenhuma pendência crítica encontrada.</div>`;
}

function renderDashboardV2() {
  const t = dashboardTotals();
  const official = officialFinanceiro();
  const closings = filteredClosings();
  $("dashboardCards").innerHTML = [card("Saída de Pacotes", num(official.operational.outputPackages.value)), card("Folha Bruta", money(official.operational.grossPayroll.value)), card("Descontos", money(official.operational.discounts.value)), card("Valor Líquido a Pagar", money(official.operational.netPay.value)), card("Entrada de Pacotes", num(official.operational.entryPackages.value)), card("Diferença de Pacotes", num(official.operational.packageDiff.value)), card("Pagar para Base", money(official.cash.payBase.value)), card("Valor Caixa Final", money(official.cash.finalCash.value)), card("Total ML entrada", num(t.baseMl)), card("Total Shopee entrada", num(t.baseShopee)), card("Total ML saída motoboys", num(t.outMl)), card("Total Shopee saída motoboys", num(t.outShopee))].join("");
  $("operationalSummary").innerHTML = tableBlock("Resumo Operacional", ["Indicador","Valor","Origem"], [["Saída de Pacotes", num(official.operational.outputPackages.value), official.operational.outputPackages.origin], ["Folha Bruta", money(official.operational.grossPayroll.value), official.operational.grossPayroll.origin], ["Descontos", money(official.operational.discounts.value), official.operational.discounts.origin], ["Valor Líquido a Pagar", money(official.operational.netPay.value), official.operational.netPay.origin], ["Entrada de Pacotes", num(official.operational.entryPackages.value), official.operational.entryPackages.origin], ["Diferença de Pacotes", num(official.operational.packageDiff.value), official.operational.packageDiff.origin]]);
  $("managerControl").innerHTML = tableBlock("Controle por Gestor", ["Gestor","Quantidade de Pacotes","Valor Entrada","Valor Entregue","Descontos","Saldo do Gestor","Origem"], Object.entries(official.managers).map(([name, x]) => [name, num(x.packages), money(x.entryValue), num(x.deliveredValue), money(x.discounts), money(x.balance), x.origin]));
  $("cashSummary").innerHTML = tableBlock("Resumo de Caixa", ["Descrição","Valor","Origem"], [["Total Entrada Pacotes", num(official.cash.totalEntryPackages.value), official.cash.totalEntryPackages.origin], ["Pagar para Base", money(official.cash.payBase.value), official.cash.payBase.origin], ["Folha Bruta Motoristas", money(official.cash.grossDrivers.value), official.cash.grossDrivers.origin], ["Total Descontos", money(official.cash.totalDiscounts.value), official.cash.totalDiscounts.origin], ["Valor Líquido a Pagar", money(official.cash.netPay.value), official.cash.netPay.origin], ["Valor Caixa Final", money(official.cash.finalCash.value), official.cash.finalCash.origin]]);
  if (CLEAN_OPERATIONAL_MODE) {
    $("dashboardCards").innerHTML = [
      card("Total bruto", money(t.gross)),
      card("Total descontos", money(t.discounts + t.losses)),
      card("Total vales", money(t.vales)),
      card("Total líquido", money(t.net)),
      card("Total pago", money(t.paid)),
      card("Total pendente", money(t.pending)),
      card("Quantidade de motoboys", num(activeRiders().length)),
      card("Quantidade de pacotes/saídas", num(t.outPackages)),
      card("Entrada de pacotes", num(t.basePackages)),
      card("Despesas fixas", money(t.fixedExpenses)),
      card("Despesas variáveis", money(t.variableExpenses)),
      card("Custos variáveis do período", money(t.variablePeriodCosts)),
      card("Custos variáveis a descontar", money(t.variableCurrentDiscount)),
      card("Custos pendentes próxima quinzena", money(t.variableNextPending)),
      card("Lucro motoboys sem coleta", money(t.noCollectionProfit)),
      card("Lucro freelancers", money(t.freelancerProfit)),
      card("Lucro total sem coleta + freelancers", money(t.profitTotal)),
      card("Pacotes ML sem coleta/freelancer", num(t.profitMlPackages)),
      card("Pacotes Shopee sem coleta/freelancer", num(t.profitShopeePackages)),
      card("Resultado final", money(t.finalResult))
    ].join("");
    $("dashboardCards").innerHTML += workTypeSummaryCards();
    $("operationalSummary").innerHTML = tableBlock("Resumo Operacional", ["Indicador","Valor"], [
      ["Saídas ML", num(t.outMl)],
      ["Saídas Shopee", num(t.outShopee)],
      ["Saídas Avulso", num(t.outAvulso)],
      ["Total de saídas", num(t.outPackages)],
      ["Folha bruta", money(t.gross)],
      ["Total líquido", money(t.net)]
    ]);
    $("managerControl").innerHTML = tableBlock("Controle por Gestor", ["Gestor","Entrada pacotes","Valor entrada","Descontos/vales","Despesas variáveis","Saldo gestor"], RESPONSIBLES.map((name) => {
      const bases = allBaseEntries().filter((x) => (x.partner || "BASE") === name);
      const discounts = allDiscounts().filter((x) => (x.partner || "BASE") === name);
      const variable = variableExpenseTotal(name);
      return [name, num(sum(bases, "totalPackages")), money(sum(bases, "totalPay")), money(sum(discounts, "value")), money(variable), money(sum(bases, "totalPay") - sum(discounts, "value") - variable)];
    }));
    $("cashSummary").innerHTML = tableBlock("Resumo de Caixa", ["Descrição","Valor"], [
      ["Entrada da base", money(t.basePay)],
      ["Líquido a pagar aos motoboys", money(t.net)],
      ["Despesas fixas", money(t.fixedExpenses)],
      ["Despesas variáveis", money(t.variableExpenses)],
      ["Custos variáveis do período", money(t.variablePeriodCosts)],
      ["Custos variáveis a descontar na quinzena atual", money(t.variableCurrentDiscount)],
      ["Custos variáveis pendentes para próxima quinzena", money(t.variableNextPending)],
      ["Total de despesas", money(t.totalExpenses)],
      ["Resultado final", money(t.finalResult)]
    ]);
  }
  $("packageComparison").innerHTML = `<div class="comparison-grid">${packageComparisonRows().map((x) => `<article class="comparison-card"><strong>${escapeHtml(x.label)}</strong><span>Entrada: ${num(x.entry)}</span><span>Saída: ${num(x.exit)}</span><span>Diferença: ${num(x.diff)}</span><span>Percentual de saída: ${num(x.pct)}%</span><span class="${x.status === "OK" ? "status-ok" : x.status === "Atenção" ? "status-attention" : "status-check"}">${x.status}</span></article>`).join("")}</div>`;
  $("rankingRows").innerHTML = closings.sort((a, b) => (b.ml + b.shopee + b.avulso) - (a.ml + a.shopee + a.avulso)).map((c, i) => { const r = riderByName(c.rider); return `<tr><td>${i + 1}</td><td>${escapeHtml(c.rider)} ${workTypeBadge(r?.collection)}</td><td>${num(c.ml + c.shopee + c.avulso)}</td><td>${money(c.gross)}</td><td>${money(c.net)}</td></tr>`; }).join("");
  renderCharts(closings);
  $("alertRows").innerHTML = validations().map(alertHtml).join("") || `<div class="alert ok">Nenhuma pendência crítica encontrada.</div>`;
}

function renderCharts(closings) {
  const byRider = closings.map((c) => ({ rider: c.rider, packages: c.ml + c.shopee + c.avulso, gross: c.gross, net: c.net })).sort((a, b) => b.packages - a.packages).slice(0, 15);
  $("chartPackages").innerHTML = barChart(byRider.map((x) => ({ label: x.rider, value: x.packages, display: num(x.packages) })));
  $("chartGross").innerHTML = barChart(byRider.map((x) => ({ label: x.rider, value: x.gross, display: money(x.gross) })));
  $("chartNet").innerHTML = barChart(byRider.map((x) => ({ label: x.rider, value: Math.max(0, x.net), display: money(x.net) })));
  $("chartDiscounts").innerHTML = barChart(RESPONSIBLES.map((partner) => {
    const value = effectiveManagerDiscount(partner);
    return { label: partner, value, display: money(value) };
  }));
  const t = dashboardTotals();
  $("chartEntryExit").innerHTML = barChart([{ label: "ML entrada", value: t.baseMl, display: num(t.baseMl) }, { label: "ML saída", value: t.outMl, display: num(t.outMl) }, { label: "Shopee entrada", value: t.baseShopee, display: num(t.baseShopee) }, { label: "Shopee saída", value: t.outShopee, display: num(t.outShopee) }, { label: "Total entrada", value: t.basePackages, display: num(t.basePackages) }, { label: "Total saída", value: t.outPackages, display: num(t.outPackages) }]);
  if ($("chartExpenses")) $("chartExpenses").innerHTML = barChart([{ label: "Fixas", value: t.fixedExpenses, display: money(t.fixedExpenses) }, { label: "Variáveis", value: t.variableExpenses, display: money(t.variableExpenses) }, { label: "Total", value: t.totalExpenses, display: money(t.totalExpenses) }]);
}

function barChart(rows) {
  const max = Math.max(1, ...rows.map((x) => Number(x.value || 0)));
  return rows.map((x) => `<div class="bar-row"><span title="${escapeHtml(x.label)}">${escapeHtml(x.label)}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.max(1, (Number(x.value || 0) / max) * 100)}%"></div></div><strong>${escapeHtml(x.display)}</strong></div>`).join("") || `<div class="empty">Sem dados.</div>`;
}

function renderMotoboysLegacy() {
  $("motoboyCount").textContent = `${num(activeRiders().length)} ativos / ${num(allRiders().length)} cadastros`;
  $("partnerStrip").innerHTML = RESPONSIBLES.map((p) => `<button class="partner-button" data-partner="${p}">${p}</button>`).join("");
  $("riderList").innerHTML = allRiders().map((r) => `<button class="rider-button ${normalize(r.name) === normalize(selectedRider) ? "active" : ""}" data-rider="${escapeHtml(r.name)}"><strong>${escapeHtml(r.name)}</strong><span>${r.partner ? "Sócio" : "Motoboy"}</span></button>`).join("");
  fillRiderForm(selectedRider);
  renderRiderProfile(selectedRider);
}

function fillRiderFormLegacy(name = "") {
  const rider = riderByName(name);
  if (!$("riderName")) return;
  $("riderId").value = rider?.id || "";
  $("riderName").value = rider?.name || "";
  $("riderRegion").value = rider?.region || "";
  $("riderCollection").value = rider?.collection || "com coleta";
  $("riderRateMl").value = money(rider?.rateMl ?? state.config.ml);
  $("riderRateShopee").value = money(rider?.rateShopee ?? state.config.shopee);
  $("riderRateAvulso").value = money(rider?.rateAvulso ?? state.config.avulso);
  $("riderNote").value = rider?.note || "";
}

function fillRiderForm(name = "") {
  const rider = editingRiderId ? riderById(editingRiderId) : riderByName(name);
  if (!$("riderName")) return;
  $("riderId").value = rider?.id || "";
  $("riderName").value = rider?.name || "";
  $("riderRegion").value = rider?.region || "";
  $("riderCollection").value = rider?.collection || "com coleta";
  $("riderRateMl").value = money(rider?.rateMl ?? state.config.ml);
  $("riderRateShopee").value = money(rider?.rateShopee ?? state.config.shopee);
  $("riderRateAvulso").value = money(rider?.rateAvulso ?? state.config.avulso);
  if ($("riderStatus")) $("riderStatus").value = rider?.active === false ? "inativo" : "ativo";
  $("riderNote").value = rider?.note || "";
  ensureWorkTypeCards();
  updateWorkTypeCards();
}

function riderData(name) {
  const key = normalize(name);
  return { daily: allDaily().filter((x) => normalize(x.rider) === key), discounts: allDiscounts().filter((x) => normalize(x.closingRider || x.rider) === key), closings: closingRecords().filter((x) => normalize(x.rider) === key) };
}

function renderMotoboys() {
  ensureWorkTypeCards();
  $("motoboyCount").textContent = `${num(activeRiders().length)} ativos / ${num(allRiders().length)} cadastros`;
  $("partnerStrip").innerHTML = RESPONSIBLES.map((p) => `<button class="partner-button" data-partner="${p}">${p}</button>`).join("");
  $("riderList").innerHTML = allRiders().map((r) => `<button class="rider-button ${normalize(r.name) === normalize(selectedRider) ? "active" : ""} ${r.active === false ? "inactive" : ""}" data-rider-id="${escapeHtml(r.id)}" data-rider="${escapeHtml(r.name)}"><strong>${escapeHtml(r.name)} ${workTypeBadge(r.collection)}</strong><span>${r.active === false ? "Inativo" : "Ativo"} | ${escapeHtml(r.region || "sem região")} | ${escapeHtml(workTypeInfo(r.collection).description)}</span></button>`).join("");
  fillRiderForm(selectedRider);
  renderRiderProfile(selectedRider);
}

function renderRiderProfile(name) {
  $("riderProfileTitle").textContent = name ? `Ficha individual - ${name}` : "Ficha individual";
  const rider = riderByName(name);
  const info = workTypeInfo(rider?.collection);
  const d = riderData(name);
  const vales = d.discounts.filter((x) => x.type === "Vale"), losses = d.discounts.filter((x) => normalize(x.type).includes("extravio") || normalize(x.type).includes("ocorrencia")), others = d.discounts.filter((x) => x.type !== "Vale" && !normalize(x.type).includes("extravio") && !normalize(x.type).includes("ocorrencia"));
  const gross = sum(d.daily, "gross"), net = gross - sum(vales, "value") - sum(losses, "value") - sum(others, "value");
  $("riderProfile").innerHTML = `<div class="profile"><div class="type-chip ${workTypeClass(rider?.collection)}"><strong>${workTypeBadge(rider?.collection)} ${escapeHtml(info.label)}</strong><span>${escapeHtml(info.description)}</span><small>Valores padrão: ML ${money(info.ml)} | Shopee ${money(info.shopee)}</small></div><div class="summary-grid">${card("Bruto", money(gross))}${card("Vales", money(sum(vales, "value")))}${card("Extravios", money(sum(losses, "value")))}${card("Outros descontos", money(sum(others, "value")))}${card("Líquido", money(net))}</div>${tableBlock("Produção diária", ["Data","Tipo","ML","Shopee","Avulso","Bruto"], d.daily.map((x) => [displayDate(x.date), workTypeLabel(dailyTypeFor(x)), num(x.ml), num(x.shopee), num(x.avulso), money(x.gross)]))}${tableBlock("Vales", ["Data","Valor","Motivo","Sócio"], vales.map((x) => [displayDate(x.date), money(x.value), x.reason, x.partner]))}${tableBlock("Extravios", ["Data","Código","Valor","Motivo","Sócio"], losses.map((x) => [displayDate(x.date), x.code, money(x.value), x.reason, x.partner]))}${tableBlock("Descontos", ["Data","Tipo","Valor","Motivo","Sócio"], others.map((x) => [displayDate(x.date), x.type, money(x.value), x.reason, x.partner]))}${tableBlock("Fechamento quinzenal", ["Período","Bruto","Vales","Extravios","Descontos","Líquido","Status"], d.closings.map((x) => [x.period, money(x.gross), money(x.vales), money(x.losses), money(x.discounts), money(x.net), x.status]))}</div>`;
}

function renderFreelancers() {
  if (!$("freelancerRows")) return;
  const rows = profitReportRows().filter((x) => x.type === "freelancer");
  const totals = profitTotals();
  $("freelancerMeta").textContent = `${num(totals.freelancerCount)} freelancer(s) cadastrados`;
  $("freelancerCards").innerHTML = [
    card("Freelancers", num(totals.freelancerCount)),
    card("Pacotes freelancers", num(totals.freelancerPackages)),
    card("Valor bruto pago", money(totals.freelancerPaid)),
    card("Lucro freelancers", money(totals.freelancerProfit))
  ].join("");
  $("freelancerList").innerHTML = allRiders().filter(isFreelancer).map((r) => `<button class="rider-button ${r.active === false ? "inactive" : ""}" data-rider-id="${escapeHtml(r.id)}" data-rider="${escapeHtml(r.name)}"><strong>${escapeHtml(r.name)} ${workTypeBadge(r.collection)}</strong><span>${r.active === false ? "Inativo" : "Ativo"} | ${escapeHtml(workTypeInfo(r.collection).description)} | ML ${money(r.rateMl)} | Shopee ${money(r.rateShopee)}</span></button>`).join("") || `<div class="empty">Nenhum freelancer cadastrado.</div>`;
  $("freelancerRows").innerHTML = rows.map((x) => `<tr><td>${escapeHtml(x.name)}</td><td>${workTypeBadge(x.type)}</td><td>${num(x.ml)}</td><td>${num(x.shopee)}</td><td>${num(x.packages)}</td><td>${money(x.paid)}</td><td>${money(x.profit)}</td><td>${escapeHtml(x.status)}</td><td>${x.closingId ? `<button data-freelancer-receipt="${escapeHtml(x.closingId)}">Recibo</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="9" class="empty">Sem produção de freelancers.</td></tr>`;
}

function renderPartners() {
  $("partnerCards").innerHTML = RESPONSIBLES.map((p) => `<button class="partner-button ${p === selectedPartner ? "active" : ""}" data-partner="${p}"><strong>${p}</strong><span>${money(partnerBaseTotal(p))} base</span></button>`).join("");
  renderPartnerProfile(selectedPartner);
}

function partnerBaseTotal(partner) { return allBaseEntries().filter((x) => x.partner === partner).reduce((s, x) => s + x.totalPay, 0); }
function renderPartnerProfileLegacy(partner) {
  $("partnerProfileTitle").textContent = `Ficha do sócio - ${partner}`;
  const base = allBaseEntries().filter((x) => x.partner === partner);
  const discounts = partnerProfileDiscounts(partner);
  const closings = baseClosingRecords().filter((x) => x.partner === partner);
  $("partnerProfile").innerHTML = `<div class="profile"><div class="summary-grid">${card("ML entrada", num(sum(base, "ml")))}${card("Shopee entrada", num(sum(base, "shopee")))}${card("Pacotes entrada", num(sum(base, "totalPackages")))}${card("A pagar base", money(sum(base, "totalPay")))}${card("Descontos lançados", money(sum(discounts, "value")))}</div>${tableBlock("Entrada de Pacotes da Base", ["Data","ML","Shopee","Total pacotes","Total a pagar"], base.map((x) => [displayDate(x.date), num(x.ml), num(x.shopee), num(x.totalPackages), money(x.totalPay)]))}${tableBlock("Controle de Descontos", ["Data","Motoboy","Tipo","Valor","Motivo"], discounts.map((x) => [displayDate(x.date), x.rider, x.type, money(x.value), x.reason]))}${tableBlock("Fechamento Quinzenal da Base", ["Período","ML","Shopee","Pacotes","Total a pagar","Status"], closings.map((x) => [x.period, num(x.ml), num(x.shopee), num(x.totalPackages), money(x.totalPay), x.status]))}</div>`;
}

function guilhermeProfileAuditHtml() {
  const audit = guilhermeProfileAudit();
  return `<div class="summary-grid">${card("Registros encontrados", num(audit.found))}${card("Registros validos", num(audit.valid))}${card("Duplicidades removidas", num(audit.duplicates))}${card("Total antes", money(audit.before))}${card("Total depois", money(audit.after))}${card("Total planilha", money(audit.sheetTotal))}</div>${tableBlock("Auditoria da ficha GUILHERME", ["Linha original","Aba","Motoboy","Tipo","Valor","Observacao","Status"], guilhermeProfileAuditTableRows())}`;
}

function renderPartnerProfile(partner) {
  $("partnerProfileTitle").textContent = `Ficha do socio - ${partner}`;
  const base = allBaseEntries().filter((x) => x.partner === partner);
  const discounts = partnerProfileDiscounts(partner);
  const variableExpenses = variableExpensesByResponsible(partner);
  const closings = baseClosingRecords().filter((x) => x.partner === partner);
  const auditHtml = partner === "GUILHERME" ? guilhermeProfileAuditHtml() : "";
  const partnerBalance = sum(base, "totalPay") - sum(discounts, "value") - sum(variableExpenses, "value");
  $("partnerProfile").innerHTML = `<div class="profile"><div class="summary-grid">${card("ML entrada", num(sum(base, "ml")))}${card("Shopee entrada", num(sum(base, "shopee")))}${card("Pacotes entrada", num(sum(base, "totalPackages")))}${card("A pagar base", money(sum(base, "totalPay")))}${card("Descontos lancados", money(sum(discounts, "value")))}${card("Despesas variaveis", money(sum(variableExpenses, "value")))}${card("Saldo do socio", money(partnerBalance))}</div>${tableBlock("Entrada de Pacotes da Base", ["Data","ML","Shopee","Total pacotes","Total a pagar"], base.map((x) => [displayDate(x.date), num(x.ml), num(x.shopee), num(x.totalPackages), money(x.totalPay)]))}${tableBlock("Controle de Descontos", ["Data","Motoboy","Tipo","Valor","Motivo"], discounts.map((x) => [displayDate(x.date), x.rider, x.type, money(x.value), x.reason]))}${tableBlock("Despesas variaveis", ["Data","Categoria","Descricao","Valor","Quinzena origem","Quinzena desconto","Status","Observacao"], variableExpenses.map((x) => [displayDate(x.date), x.category, x.description || "", money(x.value), x.originPeriodLabel || "", x.discountPeriodLabel || "", x.status || "pendente", x.note || ""]))}${auditHtml}${tableBlock("Fechamento Quinzenal da Base", ["Periodo","ML","Shopee","Pacotes","Total a pagar","Descontos","Desp. variaveis","Saldo","Status"], closings.map((x) => [x.period, num(x.ml), num(x.shopee), num(x.totalPackages), money(x.totalPay), money(x.managerDiscounts), money(x.variableDiscounts), money(x.netAfterVariable), x.status]))}</div>`;
}

function tableBlock(title, headers, rows) {
  return `<section class="section-block"><h3>${escapeHtml(title)}</h3><div class="table-wrap"><table><thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>${rows.map((row) => `<tr>${row.map((x) => `<td>${escapeHtml(x)}</td>`).join("")}</tr>`).join("") || `<tr><td colspan="${headers.length}" class="empty">Sem registros.</td></tr>`}</tbody></table></div></section>`;
}

function renderBase() {
  updateBaseTotals();
  $("baseRows").innerHTML = allBaseEntries().map((x) => `<tr><td>${displayDate(x.date)}</td><td>${x.partner}</td><td>${num(x.ml)}</td><td>${num(x.shopee)}</td><td>${num(x.totalPackages)}</td><td>${money(x.totalPay)}</td><td>${x.source === "manual" ? `<button data-remove-base="${escapeHtml(x.id)}">Remover</button>` : ""}</td></tr>`).join("");
  $("baseClosingRows").innerHTML = baseClosingRecords().map((x) => `<tr><td>${x.partner}</td><td>${x.period}</td><td>${num(x.ml)}</td><td>${num(x.shopee)}</td><td>${num(x.totalPackages)}</td><td>${money(x.valueMl)}</td><td>${money(x.valueShopee)}</td><td>${money(x.totalPay)}</td><td>${money(x.managerDiscounts)}</td><td>${money(x.variablePeriodCosts)}</td><td>${money(x.variableDiscounts)}</td><td>${money(x.netAfterVariable)}</td><td>${x.status}</td></tr>`).join("");
}

function renderDaily() {
  updateDailyGross();
  const selectedDate = isoDateOnly($("dailyDate")?.value || "");
  const rows = selectedDate ? allDaily().filter((x) => isoDateOnly(x.date) === selectedDate) : allDaily();
  $("dailyRows").innerHTML = rows.map((x) => { const actions = x.source === "manual" && canWriteRecord("daily", x) ? `<button data-edit-daily="${escapeHtml(x.id)}">Editar</button> <button data-remove-daily="${escapeHtml(x.id)}">Remover</button>` : ""; return `<tr><td>${displayDate(x.date)}</td><td>${escapeHtml(x.rider)}</td><td>${workTypeBadge(dailyTypeFor(x))}</td><td>${num(x.ml)}</td><td>${num(x.shopee)}</td><td>${num(x.avulso)}</td><td>${money(x.gross)}</td><td>${actions}</td></tr>`; }).join("") || `<tr><td colspan="8" class="empty">Sem lançamentos para esta data.</td></tr>`;
}
function renderDiscounts() {
  $("discountRows").innerHTML = allDiscounts().map((x) => `<tr><td>${displayDate(x.date)}</td><td>${escapeHtml(x.partner || "BASE")}</td><td>${escapeHtml(x.rider)}</td><td>${escapeHtml(x.type)}</td><td>${money(x.value)}</td><td>${escapeHtml(x.observation || x.reason || x.note || "")}</td><td>${escapeHtml(x.sheetOriginal || "")}</td><td>${escapeHtml(x.lineOriginal || "")}</td><td>${escapeHtml(x.columnOriginal || "")}</td><td>${x.source === "manual" ? `<button data-edit-discount="${escapeHtml(x.id)}">Editar</button> <button data-remove-discount="${escapeHtml(x.id)}">Remover</button>` : ""}</td></tr>`).join("");
  if ($("baseDiscountRows")) $("baseDiscountRows").innerHTML = allDiscounts().filter((x) => (x.partner || "BASE") === "BASE").map((x) => `<tr><td>${displayDate(x.date)}</td><td>${escapeHtml(x.rider || "")}</td><td>${escapeHtml(x.type)}</td><td>${money(x.value)}</td><td>${escapeHtml(x.reason || "")}</td><td>${escapeHtml(x.observation || x.note || "")}</td><td>${x.source === "manual" ? `<button data-edit-discount="${escapeHtml(x.id)}">Editar</button> <button data-remove-discount="${escapeHtml(x.id)}">Remover</button>` : ""}</td></tr>`).join("");
}

function updateExpensePeriodFields(forceCalculated = false) {
  if (!$("expenseOriginPeriod") || !$("expenseDiscountPeriod")) return;
  const isVariable = normalizeExpenseType($("expenseType").value) === "variavel";
  if (!isVariable) {
    $("expenseOriginPeriod").value = "Nao se aplica";
    $("expenseDiscountPeriod").innerHTML = `<option value="">Nao se aplica</option>`;
    $("expenseDiscountPeriod").disabled = true;
    return;
  }
  const date = $("expenseDate").value || new Date().toISOString().slice(0, 10);
  const origin = periodKey(date, "quinzenal");
  const calculated = variableExpenseDiscountPeriod(date);
  const current = forceCalculated ? calculated.key : ($("expenseDiscountPeriod").value || calculated.key);
  $("expenseOriginPeriod").value = origin.label;
  $("expenseDiscountPeriod").disabled = false;
  const options = variableExpensePeriodOptions(date);
  if (!options.some((x) => x.key === current)) options.unshift({ key: current, label: current.replace("|", " a "), start: "", end: "" });
  $("expenseDiscountPeriod").innerHTML = options.map((x) => `<option value="${escapeHtml(x.key)}">${escapeHtml(x.label)}</option>`).join("");
  $("expenseDiscountPeriod").value = current;
}

function renderExpenses() {
  if (!$("expenseRows")) return;
  updateExpensePeriodFields();
  $("expenseRows").innerHTML = allExpenses().map((x) => `<tr><td>${displayDate(x.date)}</td><td>${escapeHtml(x.type)}</td><td>${escapeHtml(x.category)}</td><td>${escapeHtml(x.description || "")}</td><td>${money(x.value)}</td><td>${escapeHtml(x.originPeriodLabel || "-")}</td><td>${escapeHtml(x.discountPeriodLabel || "-")}</td><td>${escapeHtml(x.responsible || "")}</td><td>${escapeHtml(x.status || "pendente")}</td><td><button data-edit-expense="${escapeHtml(x.id)}">Editar</button> <button data-remove-expense="${escapeHtml(x.id)}">Remover</button></td></tr>`).join("");
  if ($("variableExpenseRows")) $("variableExpenseRows").innerHTML = allExpenses().filter((x) => x.type === "variavel").map((x) => `<tr><td>${displayDate(x.date)}</td><td>${escapeHtml(x.responsible || "BASE")}</td><td>${escapeHtml(x.category)}</td><td>${escapeHtml(x.description || "")}</td><td>${money(x.value)}</td><td>${escapeHtml(x.discountPeriodLabel || "")}</td><td>${escapeHtml(x.status || "pendente")}</td><td>${escapeHtml(x.note || "")}</td></tr>`).join("") || `<tr><td colspan="8" class="empty">Sem despesas variáveis.</td></tr>`;
}

function renderClosings() {
  const rider = $("closingRider").value;
  const rows = closingRecords($("closingType").value).filter((x) => !rider || normalize(x.rider) === normalize(rider));
  $("closingMeta").textContent = `${num(rows.length)} fechamento(s) | ${money(sum(rows, "net"))}`;
  $("closingRows").innerHTML = rows.map((c) => { const r = riderByName(c.rider); return `<tr><td>${escapeHtml(c.rider)} ${workTypeBadge(r?.collection)}</td><td>${c.period}</td><td>${num(c.ml)}</td><td>${num(c.shopee)}</td><td>${num(c.avulso)}</td><td>${money(c.gross)}</td><td>${money(c.vales)}</td><td>${money(c.losses)}</td><td>${money(c.discounts)}</td><td>${money(c.bonuses)}</td><td>${money(c.net)}</td><td>${c.status}</td><td><button data-receipt="${escapeHtml(c.id)}">Recibo</button></td></tr>`; }).join("");
}

function renderReceipts(selectedId) {
  const rider = $("receiptRider").value;
  const rows = closingRecords().filter((x) => !rider || normalize(x.rider) === normalize(rider));
  $("receiptClosing").innerHTML = rows.map((c) => `<option value="${escapeHtml(c.id)}">${escapeHtml(c.rider)} - ${c.period} - ${money(c.net)}</option>`).join("");
  if (selectedId) $("receiptClosing").value = selectedId;
  const row = closingRecords().find((x) => x.id === $("receiptClosing").value) || rows[0];
  if (row && $("paymentValue")) $("paymentValue").value = money(row.net);
  $("receiptPreview").innerHTML = row ? receiptHtml(row) : `<div class="empty">Selecione um fechamento.</div>`;
}

function receiptHtml(c) {
  return `<article class="receipt"><h2>Recibo de Pagamento</h2><p><strong>Nome do motoboy:</strong> ${escapeHtml(c.rider)}</p><p><strong>Período:</strong> ${c.period} (${displayDate(c.start)} a ${displayDate(c.end)})</p><p><strong>Data do pagamento:</strong> ${displayDate(c.paymentDate || new Date().toISOString().slice(0, 10))}</p><div class="summary-grid">${card("Total ML", num(c.ml))}${card("Total Shopee", num(c.shopee))}${card("Total Avulso", num(c.avulso))}${card("Valor bruto", money(c.gross))}${card("Vales", money(c.vales))}${card("Extravios", money(c.losses))}${card("Outros descontos", money(c.discounts))}${card("Bonificações", money(c.bonuses))}${card("Total líquido pago", money(c.net))}</div><p><strong>Observações:</strong> -</p><div class="signatures"><span>Assinatura do responsável</span><span>Assinatura do motoboy</span></div></article>`;
}

function renderReports() {
  const t = dashboardTotals();
  $("riderReport").innerHTML = allRiders().map((r) => `<p><strong>${escapeHtml(r.name)} ${workTypeBadge(r.collection)}</strong>: ${money(closingRecords().filter((x) => normalize(x.rider) === normalize(r.name)).reduce((s, x) => s + x.net, 0))}</p>`).join("");
  $("partnerReport").innerHTML = RESPONSIBLES.map((p) => `<p><strong>${p}</strong>: base ${money(partnerBaseTotal(p))}, descontos ${money(sum(allDiscounts().filter((x) => (x.partner || "BASE") === p), "value"))}, despesas variáveis ${money(variableExpenseTotal(p))}</p>`).join("");
  $("baseReport").innerHTML = tableBlock("Bases", ["Sócio","Período","ML","Shopee","Total pagar","Descontos","Custos variáveis","Saldo"], baseClosingRecords().map((x) => [x.partner, x.period, num(x.ml), num(x.shopee), money(x.totalPay), money(x.managerDiscounts), money(x.variableDiscounts), money(x.netAfterVariable)]));
  $("partnerReport").innerHTML += tableBlock("Despesas variaveis por responsavel", ["Data","Responsavel","Categoria","Descricao","Valor","Quinzena desconto","Status"], allExpenses().filter((x) => x.type === "variavel").map((x) => [displayDate(x.date), x.responsible || "BASE", x.category || "", x.description || "", money(x.value), x.discountPeriodLabel || "", x.status || "pendente"]));
  $("discountReport").innerHTML = `<p>Vales: ${money(sum(discountsByType("Vale"), "value"))}</p><p>Extravios/Ocorrências: ${money(allDiscounts().filter((x) => normalize(x.type).includes("extravio") || normalize(x.type).includes("ocorrencia")).reduce((s, x) => s + x.value, 0))}</p><p>Outros: ${money(allDiscounts().filter((x) => !normalize(x.type).includes("vale") && !normalize(x.type).includes("extravio") && !normalize(x.type).includes("ocorrencia")).reduce((s, x) => s + x.value, 0))}</p>`;
  $("discountReport").innerHTML += `<p>Despesas fixas: ${money(t.fixedExpenses)}</p><p>Despesas variáveis: ${money(t.variableExpenses)}</p><p>Resultado final: ${money(t.finalResult)}</p>`;
  if ($("profitReport")) $("profitReport").innerHTML = `<div class="table-wrap"><table><thead><tr><th>Nome</th><th>Tipo</th><th>ML entregues</th><th>Shopee entregues</th><th>Valor recebido da base</th><th>Valor pago</th><th>Lucro ML</th><th>Lucro Shopee</th><th>Lucro total</th></tr></thead><tbody>${profitReportRows().map((x) => `<tr><td>${escapeHtml(x.name)}</td><td>${workTypeBadge(x.type)}</td><td>${num(x.ml)}</td><td>${num(x.shopee)}</td><td>${money(x.baseReceived)}</td><td>${money(x.paid)}</td><td>${money(x.profitMl)}</td><td>${money(x.profitShopee)}</td><td>${money(x.profit)}</td></tr>`).join("") || `<tr><td colspan="9" class="empty">Sem registros de lucro.</td></tr>`}</tbody></table></div>`;
  $("pendingReport").innerHTML = closingRecords().filter((x) => x.status !== "pago").map((x) => alertHtml({ type: "Pendente", detail: `${x.rider} - ${x.period} - ${money(x.net)}` })).join("") || `<div class="alert ok">Sem pendências.</div>`;
}

function renderAudit() {
  const cols = sheets.reduce((s, x) => s + (x.columnCount || 0), 0), importedCols = sheets.reduce((s, x) => s + (x.importedColumnCount || 0), 0);
  $("auditImport").innerHTML = `<div class="cards">${card("Abas encontradas", num(workbook.sheetCount))}${card("Motoboys importados", num(allRiders().length))}${card("Sócios importados", num(PARTNERS.length))}${card("Colunas encontradas", num(cols))}${card("Colunas importadas", num(importedCols))}${card("ML entrada importado", num(sum(imported.baseEntries, "ml")))}${card("Shopee entrada importado", num(sum(imported.baseEntries, "shopee")))}${card("A pagar bases importado", money(sum(imported.baseEntries, "totalPay")))}${card("Descontos por sócios", money(sum(imported.discounts, "value")))}${card("Divergências", num(validations().length))}</div><div class="alerts">${validations().map(alertHtml).join("") || `<div class="alert ok">Importação conferida.</div>`}</div>`;
}

function renderAuditV2() {
  const cols = sheets.reduce((s, x) => s + (x.columnCount || 0), 0), importedCols = sheets.reduce((s, x) => s + (x.importedColumnCount || 0), 0);
  const responsibleRows = RESPONSIBLES.map((p) => [p, money(effectiveManagerDiscount(p)), num(p === "BASE" ? importedManagerDiscountItems(p).length : managerFinancialDiscountItems(p).length)]);
  const riderRows = [...new Map(imported.discounts.filter((x) => x.closingRider).map((x) => [x.closingRider, x.closingRider])).values()]
    .sort((a, b) => a.localeCompare(b, "pt-BR"))
    .map((rider) => [rider, money(sum(imported.discounts.filter((x) => normalize(x.closingRider) === normalize(rider)), "value")), num(imported.discounts.filter((x) => normalize(x.closingRider) === normalize(rider)).length)]);
  const originRows = (imported.discountOrigins || []).map((x) => [x.origin || "", x.partner || "BASE", x.rider, x.type, money(x.value), x.observation || x.reason || ""]);
  const aliases = aliasAuditRows();
  const blockRows = (imported.auditBlocks || []).map((x) => [x.partner, x.sheet, x.startColumn, num(x.importedCount), money(x.itemSum), x.total == null ? "Sem total" : money(x.total), x.totalOrigin || ""]);
  const comparisonRows = CLEAN_OPERATIONAL_MODE
    ? [
      ["Modo operacional", "Limpo para uso real", "Planilha preservada como base", money(0), "Sem divergencia operacional"],
      ["Lancamentos ativos", num(allDaily().length + allBaseEntries().length + allDiscounts().length), "0 importados antigos", money(0), "App"]
    ]
    : officialComparisons().map((x) => [x.label, x.format === "money" ? money(x.app) : num(x.app), x.format === "money" ? money(x.sheet) : num(x.sheet), money(Math.abs(x.app - x.sheet)), x.origin]);
  $("auditImport").innerHTML = `${syncReportBlocks()}<div class="cards">${card("Abas lidas", num(workbook.sheetCount))}${card("Colunas encontradas", num(cols))}${card("Colunas importadas", num(importedCols))}${card("Blocos de desconto", num((imported.auditBlocks || []).length))}${card("Descontos importados", num(imported.discounts.length))}${card("Descontos ignorados", num((imported.ignored || []).length))}${card("Duplicidades removidas", num(imported.duplicateCount || 0))}${card("Valor total descontos", money(sum(imported.discounts, "value")))}</div>${tableBlock("APP x PLANILHA", ["Indicador","Valor App","Valor Planilha","Diferença","Origem Planilha"], comparisonRows)}${tableBlock("Total por responsável", ["Responsável","Total","Registros"], responsibleRows)}${tableBlock("Total por motoboy", ["Motoboy","Total","Registros"], riderRows)}${tableBlock("Conferência por bloco", ["Responsável","Aba","Coluna inicial","Itens","Soma itens","Total planilha","Origem total"], blockRows)}${tableBlock("Células usadas como origem dos descontos", ["Origem","Responsável","Motoboy","Tipo","Valor","Observação"], originRows)}<div class="alerts">${validations().map(alertHtml).join("") || `<div class="alert ok">Importação conferida.</div>`}</div>`;
  $("auditImport").innerHTML += tableBlock("Aliases de gestores unificados", ["Alias encontrado","Gestor final","Motoboy","Tipo","Valor","Origem","Observacao"], aliases);
  const guilherme = guilhermeAuditSummary();
  const guilhermeAll = importedManagerDiscountItems("GUILHERME");
  $("auditImport").innerHTML += `<div class="cards">${card("GUILHERME encontrados", num(guilherme.found))}${card("Atribuidos ao GUILHERME", num(guilherme.assigned))}${card("Vinculados a motoboys", num(guilherme.linked))}${card("Sem motoboy", num(guilherme.withoutRider))}${card("Total financeiro GUILHERME", money(effectiveManagerDiscount("GUILHERME")))}${card("Registros GUILHERME rastreados", money(sum(guilhermeAll, "value")))}${card("Registros GM separados", num(imported.discounts.filter((x) => normalize(x.partner) === "gm").length))}${card("Duplicados GUILHERME", num(guilherme.duplicates))}${card("Ignorados GUILHERME", num(guilherme.ignored))}</div>`;
  $("auditImport").innerHTML += tableBlock("Auditoria completa GUILHERME", ["Aba","Linha","Coluna","Motoboy","Valor","Tipo","Observacao","Status","Motivo"], guilhermeAuditTableRows());
}

function ensureSupabasePanel() {
  if ($("supabaseStatus")) return;
  const view = $("configuracoes");
  if (!view) return;
  view.insertAdjacentHTML("beforeend", `<section class="panel form">
    <div class="panel-head"><h2>Supabase</h2></div>
    <div id="supabaseStatus" class="cards"></div>
    <div id="supabaseFeedback" class="feedback" role="status"></div>
    <div class="actions">
      <button id="syncSupabase" type="button">Sincronizar agora</button>
      <button id="migrateLocalToSupabase" type="button">Migrar dados locais para Supabase</button>
      <button id="logoutSupabase" type="button">Sair</button>
    </div>
  </section>`);
}

function renderSupabaseStatus() {
  ensureSupabasePanel();
  if (!$("supabaseStatus")) return;
  const status = !isSupabaseConfigured ? "Nao configurado" : supabaseOnline ? "Conectado" : "Desconectado";
  const user = supabaseProfile?.username || supabaseSession?.user?.email || "Sem login";
  $("supabaseStatus").innerHTML = [
    card("Status", status),
    card("Usuario logado", user),
    card("Perfil", supabaseProfile?.role || "-"),
    card("Ultimo sync", lastSupabaseSync ? displayDate(lastSupabaseSync.slice(0, 10)) : "-")
  ].join("");
}

function renderConfig() {
  $("configMl").value = money(state.config.ml);
  $("configShopee").value = money(state.config.shopee);
  $("configAvulso").value = money(state.config.avulso);
  renderSupabaseStatus();
}

function setOptions(selectId, values) {
  const select = $(selectId);
  if (!select) return;
  select.innerHTML = values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
}

function applyRoleDefaults() {
  const role = appRole();
  const partnerLocked = PARTNERS.includes(role);
  if (partnerLocked) {
    ["basePartner", "discountPartner", "expenseResponsible", "paymentPartner"].forEach((id) => {
      const select = $(id);
      if (!select) return;
      setOptions(id, [role]);
      select.value = role;
      select.disabled = true;
    });
  } else if (isOperatorRole()) {
    ["basePartner", "discountPartner", "expenseResponsible", "paymentPartner"].forEach((id) => {
      const select = $(id);
      if (select) select.disabled = true;
    });
  } else {
    const partnerOptions = ["GIL", "SALES", "GUILHERME"];
    setOptions("basePartner", partnerOptions);
    setOptions("paymentPartner", partnerOptions);
    setOptions("discountPartner", [...partnerOptions, "BASE"]);
    setOptions("expenseResponsible", [...partnerOptions, "BASE"]);
    ["basePartner", "discountPartner", "expenseResponsible", "paymentPartner"].forEach((id) => { if ($(id)) $(id).disabled = false; });
  }
}

function setFormDisabled(formId, disabled) {
  const form = $(formId);
  if (!form) return;
  form.querySelectorAll("input, select, textarea, button").forEach((el) => {
    if (el.id === "sidebarLogout") return;
    el.disabled = Boolean(disabled);
  });
}

function applyPermissions() {
  const allowed = new Set(permittedViews());
  document.querySelectorAll(".nav-item[data-view]").forEach((button) => {
    button.hidden = !allowed.has(button.dataset.view);
  });
  if ($("sidebarLogout")) $("sidebarLogout").hidden = false;
  if (isSupabaseConfigured && !supabaseSession) return;
  document.querySelectorAll(".view").forEach((view) => {
    if (allowed.has(view.id)) return;
    view.classList.remove("active");
  });
  const active = document.querySelector(".view.active")?.id;
  if (!active || !allowed.has(active)) switchView(isOperatorRole() ? "lancamentos" : [...allowed][0] || "dashboard");
  if (document.querySelector(".top-actions")) document.querySelector(".top-actions").hidden = isOperatorRole();
  setFormDisabled("riderForm", !isAdminRole());
  setFormDisabled("baseForm", isOperatorRole());
  setFormDisabled("discountForm", isOperatorRole());
  setFormDisabled("expenseForm", isOperatorRole());
  const adminOnly = ["exportBackup", "importBackup", "backupMigrateLocal", "clearOperationalBackup", "saveConfig", "clearOperational", "saveUserProfile", "newUserProfile"];
  adminOnly.forEach((id) => { if ($(id)) $(id).disabled = !isAdminRole(); });
  applyRoleDefaults();
}

function showUserFeedback(message, type = "ok") {
  const el = $("userFeedback");
  if (!el) return;
  el.textContent = message;
  el.className = `feedback ${type}`;
}

function clearUserForm() {
  ["profileId", "profileAuthId", "profileUsername", "profileFullName"].forEach((id) => { if ($(id)) $(id).value = ""; });
  if ($("profileRole")) $("profileRole").value = "OPERADOR";
  if ($("profileActive")) $("profileActive").value = "true";
}

async function refreshProfiles() {
  if (!isAdminRole() || !supabaseOnline) {
    profiles = [];
    renderUsers();
    return;
  }
  try {
    profiles = await loadProfiles();
    renderUsers();
  } catch (error) {
    showUserFeedback(error.message || "Erro ao carregar usuarios.", "error");
  }
}

function renderUsers() {
  if (!$("usersList")) return;
  if (!isAdminRole()) {
    $("usersList").innerHTML = `<div class="empty">${PERMISSION_MESSAGE}</div>`;
    return;
  }
  $("usersList").innerHTML = `<div class="table-wrap"><table><thead><tr><th>Usuario</th><th>Perfil</th><th>Nome</th><th>Status</th><th>Acoes</th></tr></thead><tbody>${(profiles || []).map((profile) => `<tr><td>${escapeHtml(profile.username || "")}</td><td>${escapeHtml(profile.role || "")}</td><td>${escapeHtml(profile.full_name || "")}</td><td>${profile.active === false ? "inativo" : "ativo"}</td><td><button data-edit-profile="${escapeHtml(profile.id)}">Editar</button></td></tr>`).join("") || `<tr><td colspan="5" class="empty">Sem perfis cadastrados.</td></tr>`}</tbody></table></div>`;
}

async function saveUserProfile() {
  if (!isAdminRole()) return permissionDenied();
  const id = $("profileAuthId")?.value.trim() || $("profileId")?.value.trim();
  const username = $("profileUsername")?.value.trim().toLowerCase();
  if (!id || !username) {
    showUserFeedback("Informe o ID do Auth e o usuario.", "error");
    return false;
  }
  try {
    await saveProfile({
      id,
      username,
      role: $("profileRole").value,
      full_name: $("profileFullName").value.trim() || username,
      active: $("profileActive").value === "true"
    });
    showUserFeedback("Perfil salvo com sucesso.", "ok");
    clearUserForm();
    await refreshProfiles();
    return true;
  } catch (error) {
    showUserFeedback(error.message || "Erro ao salvar perfil.", "error");
    return false;
  }
}

function backupPayload() {
  return {
    app: "financeiro-motoboys",
    version: 1,
    exportedAt: new Date().toISOString(),
    partners: PARTNERS,
    responsibles: RESPONSIBLES,
    state: normalizeStateData(state)
  };
}

function showBackupFeedback(message, type = "ok") {
  const el = $("backupFeedback");
  if (!el) return;
  el.textContent = message;
  el.className = `feedback ${type}`;
}

function localMigrationCounts(source = {}) {
  return {
    motoboys: (source.riders || []).length,
    entradas: (source.baseEntries || []).length,
    lancamentos: (source.daily || []).length,
    descontos: (source.discounts || []).length,
    despesas: (source.expenses || []).length,
    pagamentos: (source.payments || []).length,
    recibos: (source.receipts || []).length
  };
}

function localMigrationScore(source = {}) {
  const counts = localMigrationCounts(source);
  return Object.values(counts).reduce((total, value) => total + Number(value || 0), 0);
}

function safeLocalStateFromText(text) {
  try {
    const parsed = JSON.parse(text || "{}");
    const scan = extractLocalStateFromPayload(parsed, "snapshot");
    return localMigrationScore(scan.state) > 0 ? scan.state : normalizeStateData(parsed.state || parsed);
  } catch {
    return null;
  }
}

function safeJsonParse(text) {
  try { return { ok: true, value: JSON.parse(text || "null") }; }
  catch (error) { return { ok: false, value: null, error }; }
}

function emptyMigrationState() {
  return normalizeStateData({});
}

function asArray(value) {
  return Array.isArray(value) ? value.filter(Boolean) : [];
}

function inferLocalBucket(name = "", sample = {}) {
  const key = normalize(name).replace(/[^a-z0-9]+/g, "");
  if (/motoboy|rider|driver|courier|motorista/.test(key)) return "riders";
  if (/baseentry|packageentry|entrada|pacote|entradaBase|packageentries/.test(key)) return "baseEntries";
  if (/daily|launch|lancamento|diaria|producao|saida/.test(key)) return "daily";
  if (/discount|desconto|vale|extravio|ocorrencia|occurrence/.test(key)) return "discounts";
  if (/expense|despesa|custo/.test(key)) return "expenses";
  if (/payment|pagamento|pago/.test(key)) return "payments";
  if (/receipt|recibo/.test(key)) return "receipts";
  if (/setting|config|configuracao/.test(key)) return "config";
  const fields = Object.keys(sample || {}).map((x) => normalize(x)).join("|");
  if (/name|nome/.test(fields) && /rateml|valorml|mlvalue|regions?/.test(fields)) return "riders";
  if (/totalpackages|totalpacotes|totalpay|entrydate|partner/.test(fields)) return "baseEntries";
  if (/rider|motoboy|ml|shopee|gross|bruto/.test(fields) && /date|data/.test(fields)) return "daily";
  if (/discount|desconto|reason|motivo|occurrence|ocorrencia|extravio|vale/.test(fields)) return "discounts";
  if (/category|categoria|expense|despesa|description|descricao/.test(fields)) return "expenses";
  if (/receipt|recibo/.test(fields)) return "receipts";
  if (/payment|pagamento|paid|pago/.test(fields)) return "payments";
  return "";
}

function numberFromAny(value) {
  if (typeof value === "number") return value;
  return parseMoney(value);
}

function textFromAny(...values) {
  const found = values.find((value) => value != null && String(value).trim() !== "");
  return found == null ? "" : String(found).trim();
}

function dateFromAny(...values) {
  const raw = textFromAny(...values);
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
    const [d, m, y] = raw.split("/");
    return `${y}-${m}-${d}`;
  }
  return excelDate(raw) || raw;
}

function normalizeLocalRow(bucket, row, key, index) {
  const baseId = textFromAny(row.id, row.supabaseId, `${key}-${bucket}-${index}`);
  if (bucket === "riders") {
    const name = textFromAny(row.name, row.nome, row.rider, row.motoboy, row.motoboyName, row.motorista);
    if (!name) return null;
    const collection = textFromAny(row.collection, row.workType, row.tipoTrabalho, row.tipo, row.coleta) || "com coleta";
    return {
      ...row,
      id: baseId,
      name,
      region: textFromAny(row.region, row.regiao),
      collection: normalize(collection).includes("freelancer") ? "freelancer" : normalize(collection).includes("sem") ? "sem coleta" : "com coleta",
      rateMl: numberFromAny(row.rateMl ?? row.valorMl ?? row.mlValue ?? row.valueMl ?? row.mlRate),
      rateShopee: numberFromAny(row.rateShopee ?? row.valorShopee ?? row.shopeeValue ?? row.valueShopee ?? row.shopeeRate),
      rateAvulso: numberFromAny(row.rateAvulso ?? row.valorAvulso ?? row.avulsoValue ?? row.valueAvulso ?? row.avulsoRate),
      note: textFromAny(row.note, row.notes, row.observacao, row.observations),
      status: textFromAny(row.status) || (row.active === false ? "inativo" : "ativo"),
      active: row.active !== false && normalize(row.status) !== "inativo",
      source: row.source || "localStorage"
    };
  }
  if (bucket === "daily") {
    const ml = Number(row.ml ?? row.mlQty ?? row.quantidadeMl ?? row.mercadoLivre ?? 0);
    const shopee = Number(row.shopee ?? row.shopeeQty ?? row.quantidadeShopee ?? 0);
    const avulso = Number(row.avulso ?? row.avulsoQty ?? row.quantidadeAvulso ?? 0);
    const rateMl = numberFromAny(row.rateMl ?? row.valorMl ?? row.valueMl ?? row.mlRate) || 8;
    const rateShopee = numberFromAny(row.rateShopee ?? row.valorShopee ?? row.valueShopee ?? row.shopeeRate) || 5;
    const rateAvulso = numberFromAny(row.rateAvulso ?? row.valorAvulso ?? row.valueAvulso ?? row.avulsoRate) || 8;
    return {
      ...row,
      id: baseId,
      date: dateFromAny(row.date, row.data, row.launchDate),
      rider: textFromAny(row.rider, row.motoboy, row.motoboyName, row.motorista),
      dailyType: textFromAny(row.dailyType, row.launchType, row.tipoLancamento, row.collection) || "com coleta",
      ml,
      shopee,
      avulso,
      rateMl,
      rateShopee,
      rateAvulso,
      gross: numberFromAny(row.gross ?? row.totalBruto ?? row.total) || ((ml * rateMl) + (shopee * rateShopee) + (avulso * rateAvulso)),
      responsible: textFromAny(row.responsible, row.responsavel),
      status: textFromAny(row.status) || "pendente",
      note: textFromAny(row.note, row.observacao, row.observation),
      source: row.source || "localStorage"
    };
  }
  if (bucket === "baseEntries") {
    const ml = Number(row.ml ?? row.mlQty ?? row.quantidadeMl ?? row.mercadoLivre ?? 0);
    const shopee = Number(row.shopee ?? row.shopeeQty ?? row.quantidadeShopee ?? 0);
    const rateMl = numberFromAny(row.rateMl ?? row.valorMl ?? row.valueMl) || 8;
    const rateShopee = numberFromAny(row.rateShopee ?? row.valorShopee ?? row.valueShopee) || 5;
    return {
      ...row,
      id: baseId,
      date: dateFromAny(row.date, row.data, row.entryDate),
      partner: normalizeResponsible(textFromAny(row.partner, row.responsible, row.responsavel, row.gestor, row.socio) || "BASE"),
      ml,
      shopee,
      totalPackages: Number(row.totalPackages ?? row.totalPacotes ?? (ml + shopee)),
      rateMl,
      rateShopee,
      valueMl: numberFromAny(row.valueMl ?? row.valorMl) || (ml * rateMl),
      valueShopee: numberFromAny(row.valueShopee ?? row.valorShopee) || (shopee * rateShopee),
      totalPay: numberFromAny(row.totalPay ?? row.totalValue ?? row.valorTotal) || ((ml * rateMl) + (shopee * rateShopee)),
      status: textFromAny(row.status) || "pendente",
      note: textFromAny(row.note, row.observacao, row.observation),
      source: row.source || "localStorage"
    };
  }
  if (bucket === "discounts") {
    const value = numberFromAny(row.value ?? row.valor ?? row.total);
    if (!value) return null;
    return {
      ...row,
      id: baseId,
      date: dateFromAny(row.date, row.data, row.discountDate),
      partner: normalizeResponsible(textFromAny(row.partner, row.responsible, row.responsavel, row.gestor, row.socio) || "BASE"),
      rider: textFromAny(row.rider, row.motoboy, row.motoboyName, row.motorista, row.closingRider),
      type: textFromAny(row.type, row.tipo, row.kind) || "OUTROS",
      value,
      reason: textFromAny(row.reason, row.motivo),
      occurrence: textFromAny(row.occurrence, row.ocorrencia),
      observation: textFromAny(row.observation, row.observacao, row.note),
      code: textFromAny(row.code, row.codigo, row.packageCode),
      status: textFromAny(row.status) || "pendente",
      source: row.source || "localStorage"
    };
  }
  if (bucket === "expenses") {
    const value = numberFromAny(row.value ?? row.valor ?? row.total);
    if (!value) return null;
    const type = normalize(textFromAny(row.type, row.tipo, row.expenseType)).includes("fix") ? "fixa" : "variavel";
    return {
      ...row,
      id: baseId,
      date: dateFromAny(row.date, row.data, row.expenseDate),
      type,
      responsible: normalizeResponsible(textFromAny(row.responsible, row.responsavel, row.partner, row.gestor) || "BASE"),
      category: textFromAny(row.category, row.categoria) || "Sem categoria",
      description: textFromAny(row.description, row.descricao),
      value,
      note: textFromAny(row.note, row.observacao, row.observation),
      status: textFromAny(row.status) || "pendente",
      source: row.source || "localStorage"
    };
  }
  if (bucket === "payments" || bucket === "receipts") {
    return { ...row, id: baseId, source: row.source || "localStorage" };
  }
  return null;
}

function mergeMigrationBucket(target, bucket, rows, key) {
  if (!Array.isArray(rows)) return 0;
  let added = 0;
  rows.forEach((row, index) => {
    const normalized = normalizeLocalRow(bucket, row, key, index);
    if (!normalized) return;
    target[bucket].push(normalized);
    added += 1;
  });
  return added;
}

function extractLocalStateFromPayload(payload, key = "") {
  const stateFound = emptyMigrationState();
  const detected = {};
  const visit = (node, label, depth = 0) => {
    if (!node || depth > 3) return;
    if (Array.isArray(node)) {
      const bucket = inferLocalBucket(label, node[0] || {});
      if (bucket && bucket !== "config") {
        const count = mergeMigrationBucket(stateFound, bucket, node, key || label);
        if (count) detected[bucket] = (detected[bucket] || 0) + count;
      }
      return;
    }
    if (typeof node !== "object") return;
    if (node.state) visit(node.state, `${label}.state`, depth + 1);
    if (node.data) visit(node.data, `${label}.data`, depth + 1);
    if (node.payload) visit(node.payload, `${label}.payload`, depth + 1);
    if (node.backup) visit(node.backup, `${label}.backup`, depth + 1);
    ["riders", "motoboys", "drivers", "couriers"].forEach((prop) => mergeMigrationBucket(stateFound, "riders", asArray(node[prop]), key || prop));
    ["daily", "dailyLaunches", "launches", "lancamentos", "lancamentosDiarios", "daily_launches"].forEach((prop) => mergeMigrationBucket(stateFound, "daily", asArray(node[prop]), key || prop));
    ["baseEntries", "packageEntries", "entradas", "entradasBase", "package_entries"].forEach((prop) => mergeMigrationBucket(stateFound, "baseEntries", asArray(node[prop]), key || prop));
    ["discounts", "descontos", "vales", "extravios", "occurrences", "ocorrencias"].forEach((prop) => mergeMigrationBucket(stateFound, "discounts", asArray(node[prop]), key || prop));
    ["expenses", "despesas"].forEach((prop) => mergeMigrationBucket(stateFound, "expenses", asArray(node[prop]), key || prop));
    ["payments", "pagamentos"].forEach((prop) => mergeMigrationBucket(stateFound, "payments", asArray(node[prop]), key || prop));
    ["receipts", "recibos"].forEach((prop) => mergeMigrationBucket(stateFound, "receipts", asArray(node[prop]), key || prop));
    if (node.config || node.settings || node.configuracoes) stateFound.config = { ...stateFound.config, ...(node.config || node.settings || node.configuracoes || {}) };
    Object.entries(node).forEach(([prop, value]) => {
      if (!Array.isArray(value)) return;
      const bucket = inferLocalBucket(prop, value[0] || {});
      if (bucket && bucket !== "config") {
        const count = mergeMigrationBucket(stateFound, bucket, value, key || prop);
        if (count) detected[bucket] = (detected[bucket] || 0) + count;
      }
    });
  };
  visit(payload, key || "localStorage", 0);
  const normalized = dedupeMigrationState(stateFound);
  return { state: normalized, detected };
}

function dedupeMigrationState(source) {
  const by = (rows, keyFn) => uniqueRecords(rows || [], keyFn);
  return {
    ...source,
    riders: by(source.riders, (x) => uniqueKey([x.id, x.name])),
    daily: by(source.daily, (x) => uniqueKey([x.id, x.date, x.rider, x.ml, x.shopee, x.avulso, x.gross])),
    baseEntries: by(source.baseEntries, (x) => uniqueKey([x.id, x.date, x.partner, x.ml, x.shopee, x.totalPay])),
    discounts: by(source.discounts, (x) => uniqueKey([x.id, x.date, x.partner, x.rider, x.type, moneyKey(x.value), x.observation || x.reason])),
    expenses: by(source.expenses, (x) => uniqueKey([x.id, x.date, x.responsible, x.type, x.category, moneyKey(x.value)])),
    payments: by(source.payments, (x) => uniqueKey([x.id, x.date, x.rider, moneyKey(x.value || x.net)])),
    receipts: by(source.receipts, (x) => uniqueKey([x.id, x.receiptNumber, x.rider, x.periodKey]))
  };
}

function localStoragePreview(value) {
  const text = String(value || "");
  return text.length > 220 ? `${text.slice(0, 220)}...` : text;
}

function localStorageDiagnostics() {
  const rows = [];
  const combined = emptyMigrationState();
  for (let index = 0; index < localStorage.length; index += 1) {
    const key = localStorage.key(index);
    const value = localStorage.getItem(key);
    const parsed = safeJsonParse(value);
    const extracted = parsed.ok ? extractLocalStateFromPayload(parsed.value, key) : { state: emptyMigrationState(), detected: {} };
    const counts = localMigrationCounts(extracted.state);
    Object.keys(combined).forEach((bucket) => {
      if (Array.isArray(combined[bucket])) combined[bucket].push(...(extracted.state[bucket] || []));
    });
    combined.config = { ...combined.config, ...(extracted.state.config || {}) };
    rows.push({
      key,
      size: String(value || "").length,
      json: parsed.ok ? "JSON" : "texto",
      recordCount: Object.values(counts).reduce((total, count) => total + Number(count || 0), 0),
      detectedType: Object.entries(counts).filter(([, count]) => count > 0).map(([name, count]) => `${name}: ${count}`).join(", ") || "sem dados migraveis",
      counts,
      preview: localStoragePreview(value)
    });
  }
  const stateFound = dedupeMigrationState(combined);
  console.group("Diagnostico localStorage - Financeiro Motoboys");
  console.table(rows.map((row) => ({ chave: row.key, tamanho: row.size, registros: row.recordCount, tipo: row.json, detectado: row.detectedType })));
  console.log("Contagens migraveis", localMigrationCounts(stateFound));
  console.groupEnd();
  return { rows, state: stateFound, score: localMigrationScore(stateFound) };
}

function renderLocalStorageDiagnostics(targetId = "backupLocalStorageDiagnostics") {
  const target = $(targetId);
  const diagnostics = localStorageDiagnostics();
  if (!target) return diagnostics;
  if (!diagnostics.rows.length) {
    target.innerHTML = `<div class="empty">Nenhuma chave encontrada no localStorage deste navegador.</div>`;
    return diagnostics;
  }
  const rows = diagnostics.rows.map((row) => [
    row.key,
    `${num(row.size)} caracteres`,
    num(row.recordCount),
    row.json,
    row.detectedType,
    row.preview
  ]);
  const counts = localMigrationCounts(diagnostics.state);
  target.innerHTML = `<div class="cards">${[
    card("Motoboys locais", num(counts.motoboys)),
    card("Entradas locais", num(counts.entradas)),
    card("Lancamentos locais", num(counts.lancamentos)),
    card("Descontos locais", num(counts.descontos))
  ].join("")}</div>${tableBlock("Chaves encontradas no localStorage", ["Chave", "Tamanho", "Registros", "Tipo", "Dados detectados", "Previa"], rows)}`;
  return diagnostics;
}

function findBestLocalMigrationSnapshot() {
  const diagnostics = localStorageDiagnostics();
  return {
    key: diagnostics.rows.map((row) => row.key).join(", "),
    state: diagnostics.state,
    score: diagnostics.score,
    rows: diagnostics.rows
  };
}

function captureLocalMigrationSnapshot() {
  const current = safeLocalStateFromText(localStorage.getItem(STORE_KEY));
  if (!current || localMigrationScore(current) === 0) return;
  const saved = safeLocalStateFromText(localStorage.getItem(MIGRATION_SNAPSHOT_KEY));
  if (saved && localMigrationScore(saved) > 0) return;
  localStorage.setItem(MIGRATION_SNAPSHOT_KEY, JSON.stringify({ capturedAt: new Date().toISOString(), state: current }));
  localStateForMigration = current;
}

function localMigrationSummaryText(source, originKey = "") {
  const counts = localMigrationCounts(source);
  return [
    `Origem localStorage: ${originKey || "snapshot atual"}`,
    `Motoboys encontrados: ${num(counts.motoboys)}`,
    `Entradas encontradas: ${num(counts.entradas)}`,
    `Lançamentos encontrados: ${num(counts.lancamentos)}`,
    `Descontos encontrados: ${num(counts.descontos)}`,
    `Despesas encontradas: ${num(counts.despesas)}`,
    `Pagamentos encontrados: ${num(counts.pagamentos)}`,
    `Recibos encontrados: ${num(counts.recibos)}`
  ].join("\n");
}

function mergeUniqueCloudState(localSource, cloudSource) {
  const mergeById = (cloudRows = [], localRows = []) => {
    const map = new Map();
    [...cloudRows, ...localRows].forEach((row) => {
      if (!row) return;
      const id = row.supabaseId || row.id || uniqueKey([row.source || "local", row.date || row.name || "", row.rider || row.partner || row.responsible || "", row.category || "", row.value || row.gross || row.totalPay || ""]);
      map.set(id, { ...map.get(id), ...row, id });
    });
    return [...map.values()];
  };
  return {
    ...cloudSource,
    riders: mergeById(cloudSource.riders, localSource.riders),
    daily: mergeById(cloudSource.daily, localSource.daily),
    baseEntries: mergeById(cloudSource.baseEntries, localSource.baseEntries),
    discounts: mergeById(cloudSource.discounts, localSource.discounts),
    expenses: mergeById(cloudSource.expenses, localSource.expenses),
    payments: mergeById(cloudSource.payments, localSource.payments),
    receipts: mergeById(cloudSource.receipts, localSource.receipts),
    paid: { ...(cloudSource.paid || {}), ...(localSource.paid || {}) },
    basePaid: { ...(cloudSource.basePaid || {}), ...(localSource.basePaid || {}) },
    config: { ...(cloudSource.config || {}), ...(localSource.config || {}) },
    audit: uniqueRecords([...(cloudSource.audit || []), ...(localSource.audit || [])], (x) => uniqueKey([x.at, x.action, x.detail])),
    lastBackupAt: cloudSource.lastBackupAt || localSource.lastBackupAt || ""
  };
}

function downloadJson(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function exportBackupData() {
  if (!isAdminRole()) return permissionDenied();
  state.lastBackupAt = new Date().toISOString();
  saveState("Backup exportado", "JSON completo dos dados do sistema.");
  const payload = backupPayload();
  if (supabaseOnline) saveBackupToCloud(payload).catch(showSupabaseError);
  downloadJson(`financeiro-motoboys-backup-${state.lastBackupAt.slice(0, 10)}.json`, payload);
  showBackupFeedback("Backup exportado com sucesso.", "ok");
  renderBackup();
}

function importBackupText(text) {
  if (!isAdminRole()) return permissionDenied();
  const parsed = JSON.parse(text);
  const nextState = normalizeStateData(parsed.state || parsed);
  if (!Array.isArray(nextState.riders)) throw new Error("Arquivo de backup invalido.");
  state = nextState;
  saveState("Backup importado", "Dados restaurados a partir de arquivo JSON.");
  selectedRider = allRiders()[0]?.name || "";
  editingRiderId = "";
  editingDiscountId = "";
  editingExpenseId = "";
  renderAll();
  fillDefaults();
  showBackupFeedback("Backup importado com sucesso.", "ok");
}

function clearOperationalData(origin = "Backup") {
  if (!isAdminRole()) return permissionDenied();
  if (!window.confirm("Limpar dados operacionais? Cadastros, valores dos motoboys e configuracoes serao preservados.")) return;
  resetOperationalBuckets();
  state.cleanOperational = true;
  state.cleanOperationalVersion = CLEAN_OPERATIONAL_VERSION;
  saveState("Limpeza operacional", `${origin}: lancamentos, descontos, despesas, recibos e pagamentos foram zerados.`);
  renderAll();
  showBackupFeedback("Dados operacionais limpos com sucesso.", "ok");
}

function renderBackup() {
  if (!$("backupInfo")) return;
  const lastBackup = state.lastBackupAt ? displayDate(state.lastBackupAt.slice(0, 10)) : "Nenhum backup exportado";
  const snapshot = findBestLocalMigrationSnapshot();
  const localCounts = localMigrationCounts(snapshot.state);
  let migratedAt = "";
  try { migratedAt = JSON.parse(localStorage.getItem(MIGRATION_MARK_KEY) || "{}").migratedAt || ""; } catch {}
  $("backupInfo").innerHTML = [
    card("Ultimo backup", lastBackup),
    card("Motoboys", num((state.riders || []).length)),
    card("Lancamentos", num((state.daily || []).length + (state.baseEntries || []).length)),
    card("Descontos", num((state.discounts || []).length)),
    card("Despesas", num((state.expenses || []).length)),
    card("Local: motoboys", num(localCounts.motoboys)),
    card("Local: entradas", num(localCounts.entradas)),
    card("Local: lancamentos", num(localCounts.lancamentos)),
    card("Local: descontos", num(localCounts.descontos)),
    card("Migracao Supabase", migratedAt ? displayDate(migratedAt.slice(0, 10)) : "Nao migrado")
  ].join("");
  const history = (state.audit || []).slice(0, 12).map((x) => [displayDate(String(x.at || "").slice(0, 10)), x.action || "", x.detail || ""]);
  $("backupHistory").innerHTML = tableBlock("Alteracoes recentes", ["Data","Acao","Detalhe"], history);
  renderLocalStorageDiagnostics("backupLocalStorageDiagnostics");
}

function isLikelyDuplicate(list, row, ignoreId = "") {
  const key = uniqueKey([row.type || row.kind || "", row.date, row.rider || row.partner || row.responsible || "", moneyKey(row.value || row.gross || row.totalPay), row.source || "manual"]);
  return list.some((x) => x.id !== ignoreId && uniqueKey([x.type || x.kind || "", x.date, x.rider || x.partner || x.responsible || "", moneyKey(x.value || x.gross || x.totalPay), x.source || "manual"]) === key);
}

function confirmDuplicate(list, row, ignoreId = "") {
  return !isLikelyDuplicate(list, row, ignoreId) || window.confirm("Esse lancamento parece duplicado. Deseja continuar?");
}

function riderById(id) { return (state.riders || []).find((x) => x.id === id); }
function riderSnapshot(rider) {
  if (!rider) return null;
  return { id: rider.id, name: rider.name, region: rider.region || "", collection: rider.collection || "com coleta", rateMl: Number(rider.rateMl || 0), rateShopee: Number(rider.rateShopee || 0), rateAvulso: Number(rider.rateAvulso || 0), note: rider.note || "", active: rider.active !== false };
}
function showRiderFeedback(message, type = "ok") {
  const el = $("riderFeedback");
  if (!el) return;
  el.textContent = message;
  el.className = `feedback ${type}`;
}
function riderHasHistory(name) {
  const key = normalize(name);
  return allDaily().some((x) => normalize(x.rider) === key)
    || allDiscounts().some((x) => normalize(x.rider) === key || normalize(x.closingRider) === key)
    || (state.payments || []).some((x) => normalize(x.rider) === key);
}
function collectRiderForm() {
  return {
    id: $("riderId").value || "",
    name: $("riderName").value.trim(),
    region: $("riderRegion").value.trim(),
    collection: $("riderCollection").value,
    rateMl: parseMoney($("riderRateMl").value),
    rateShopee: parseMoney($("riderRateShopee").value),
    rateAvulso: parseMoney($("riderRateAvulso").value),
    note: $("riderNote").value.trim(),
    active: $("riderStatus") ? $("riderStatus").value === "ativo" : true
  };
}
function defaultRatesForWorkType(type) {
  const key = normalize(type);
  if (key === "sem coleta" || key === "freelancer") return { ml: DEFAULT_NO_COLLECTION_ML, shopee: DEFAULT_NO_COLLECTION_SHOPEE, avulso: DEFAULT_NO_COLLECTION_ML };
  return { ml: BASE_RATE_ML, shopee: BASE_RATE_SHOPEE, avulso: BASE_RATE_ML };
}
function dailyTypeFromRider(rider) {
  return riderWorkType(rider) === "sem coleta" || riderWorkType(rider) === "freelancer" ? "sem coleta" : "com coleta";
}
function dailyTypeFor(row) {
  const raw = row?.dailyType || row?.launchType || row?.collection || "";
  if (raw) {
    const saved = riderWorkType({ collection: raw });
    if (saved === "sem coleta" || saved === "freelancer") return "sem coleta";
    return "com coleta";
  }
  if (Number(row?.rateMl) === DEFAULT_NO_COLLECTION_ML && Number(row?.rateShopee) === DEFAULT_NO_COLLECTION_SHOPEE) return "sem coleta";
  if (Number(row?.rateMl) === BASE_RATE_ML && Number(row?.rateShopee) === BASE_RATE_SHOPEE) return "com coleta";
  return dailyTypeFromRider(riderByName(row?.rider));
}
function applyDailyTypeDefaultRates() {
  const type = $("dailyType")?.value || "com coleta";
  const rates = defaultRatesForWorkType(type);
  $("dailyRateMl").value = money(rates.ml);
  $("dailyRateShopee").value = money(rates.shopee);
  $("dailyRateAvulso").value = money(rates.avulso);
  updateDailyGross();
}
function updateDailyTypeSwitch() {
  const btn = $("dailyTypeSwitch");
  if (!btn || !$("dailyType")) return;
  const isWithCollection = $("dailyType").value === "com coleta";
  btn.classList.toggle("is-with-collection", isWithCollection);
  btn.classList.toggle("is-no-collection", !isWithCollection);
  btn.setAttribute("aria-checked", String(isWithCollection));
  btn.setAttribute("aria-label", `Tipo do lançamento: ${isWithCollection ? "com coleta" : "sem coleta"}`);
  btn.querySelector(".daily-type-text").textContent = isWithCollection ? "COM COLETA" : "SEM COLETA";
}
function setDailyType(type, applyRates = true) {
  if (!$("dailyType")) return;
  $("dailyType").value = riderWorkType({ collection: type }) === "sem coleta" ? "sem coleta" : "com coleta";
  updateDailyTypeSwitch();
  if (applyRates) applyDailyTypeDefaultRates();
}
function applyWorkTypeDefaultRates() {
  const rates = defaultRatesForWorkType($("riderCollection").value);
  $("riderRateMl").value = money(rates.ml);
  $("riderRateShopee").value = money(rates.shopee);
  updateWorkTypeCards();
}
function ensureWorkTypeCards() {
  if ($("workTypeCards") || !$("riderCollection")) return;
  $("riderCollection").closest("label")?.insertAdjacentHTML("afterend", `<div id="workTypeCards" class="work-type-cards"></div>`);
  $("workTypeCards")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-work-type]");
    if (!btn) return;
    $("riderCollection").value = btn.dataset.workType;
    applyWorkTypeDefaultRates();
  });
}
function updateWorkTypeCards() {
  if (!$("workTypeCards")) return;
  $("workTypeCards").innerHTML = workTypeCardsHtml($("riderCollection").value);
}
function isMoneyTextValid(raw) {
  const text = String(raw || "").replace(/\s/g, "").replace("R$", "");
  if (!text) return false;
  return /^-?(\d+|\d{1,3}(\.\d{3})+)(,\d{1,2})?$/.test(text) || /^-?\d+(\.\d{1,2})?$/.test(text);
}
function validateRider(row) {
  if (!row.name) return "Informe o nome do motoboy.";
  if (!row.region) return "Informe a região do motoboy.";
  if (!row.collection) return "Selecione se o motoboy tem coleta.";
  if (!isMoneyTextValid($("riderRateMl").value) || !isMoneyTextValid($("riderRateShopee").value) || !isMoneyTextValid($("riderRateAvulso").value)) return "Informe valores válidos em Valor ML, Valor Shopee e Valor Avulso.";
  if ([row.rateMl, row.rateShopee, row.rateAvulso].some((value) => !Number.isFinite(value) || value < 0)) return "Informe valores válidos e sem números negativos.";
  return "";
}
function duplicateRiderByName(name, ignoreId = "") {
  return (state.riders || []).find((x) => x.id !== ignoreId && String(x.name || "").trim() === String(name || "").trim());
}
function logRiderChange(action, previous, next) {
  saveState(`Motoboy ${action}`, (next || previous)?.name || "", { module: "motoboys", action, previous, next });
}
function saveRiderFromForm() {
  if (!isAdminRole()) return permissionDenied();
  const form = collectRiderForm();
  const error = validateRider(form);
  if (error) {
    showRiderFeedback(`Erro ao salvar motoboy: ${error}`, "error");
    return false;
  }
  const duplicate = duplicateRiderByName(form.name, form.id);
  if (duplicate) {
    const editExisting = window.confirm("Esse motoboy já está cadastrado. Deseja editar o cadastro existente?");
    if (editExisting) {
      editingRiderId = duplicate.id;
      selectedRider = duplicate.name;
      fillRiderForm(duplicate.name);
      renderMotoboys();
    }
    showRiderFeedback("Cadastro duplicado não foi salvo.", "error");
    return false;
  }
  const id = form.id || uid("rider");
  const idx = state.riders.findIndex((x) => x.id === id);
  const previous = riderSnapshot(idx >= 0 ? state.riders[idx] : null);
  const row = { ...(idx >= 0 ? state.riders[idx] : {}), ...form, id };
  if (idx >= 0) {
    state.riders[idx] = row;
    logRiderChange("editado", previous, riderSnapshot(row));
    showRiderFeedback("Motoboy atualizado com sucesso", "ok");
  } else {
    state.riders.unshift(row);
    logRiderChange("criado", null, riderSnapshot(row));
    showRiderFeedback("Motoboy cadastrado com sucesso", "ok");
  }
  editingRiderId = id;
  selectedRider = row.name;
  persistRecord("riders", row);
  renderAll();
  return true;
}
function toggleSelectedRiderStatus() {
  if (!isAdminRole()) return permissionDenied();
  const r = riderById($("riderId").value) || riderByName(selectedRider);
  if (!r) {
    showRiderFeedback("Selecione um motoboy para ativar ou inativar.", "error");
    return false;
  }
  const previous = riderSnapshot(r);
  r.active = r.active === false;
  logRiderChange(r.active ? "reativado" : "inativado", previous, riderSnapshot(r));
  selectedRider = r.name;
  editingRiderId = r.id;
  persistRecord("riders", r);
  showRiderFeedback(r.active ? "Motoboy reativado com sucesso" : "Motoboy inativado com sucesso", "ok");
  renderAll();
  return true;
}
function removeSelectedRider() {
  if (!isAdminRole()) return permissionDenied();
  const r = riderById($("riderId").value) || riderByName(selectedRider);
  if (!r) {
    showRiderFeedback("Selecione um motoboy para remover.", "error");
    return false;
  }
  if (!window.confirm(`Remover ${r.name}?`)) return false;
  const previous = riderSnapshot(r);
  if (riderHasHistory(r.name)) {
    r.active = false;
    logRiderChange("inativado", previous, riderSnapshot(r));
    editingRiderId = r.id;
    selectedRider = r.name;
    persistRecord("riders", r);
    showRiderFeedback("Motoboy inativado com sucesso", "ok");
  } else {
    state.riders = state.riders.filter((x) => x.id !== r.id);
    logRiderChange("removido", previous, null);
    removeCloudRecord("riders", r);
    editingRiderId = "";
    selectedRider = allRiders()[0]?.name || "";
    showRiderFeedback("Motoboy removido com sucesso", "ok");
  }
  renderAll();
  return true;
}
function clearRiderForm() {
  editingRiderId = "";
  selectedRider = "";
  fillRiderForm("");
  if ($("riderCollection")) {
    $("riderCollection").value = "com coleta";
    applyWorkTypeDefaultRates();
  }
  showRiderFeedback("Formulário pronto para novo motoboy.", "ok");
}
function clearDiscountForm() { editingDiscountId = ""; ["discountDate","discountCode","discountReason","discountNote"].forEach((id) => $(id).value = ""); $("discountValue").value = money(0); }
function clearExpenseForm() {
  editingExpenseId = "";
  ["expenseId","expenseDate","expenseCategory","expenseDescription","expenseNote"].forEach((id) => $(id).value = "");
  $("expenseType").value = "fixa";
  $("expenseValue").value = money(0);
  $("expenseStatus").value = "pendente";
  updateExpensePeriodFields();
}

function validations() {
  const issues = [];
  issues.push(...(CLEAN_OPERATIONAL_MODE ? (imported.ignored || []).filter((x) => x.type !== "Versao limpa") : (imported.ignored || [])));
  if (!CLEAN_OPERATIONAL_MODE) {
    officialComparisons().forEach((x) => {
      if (Math.abs(x.app - x.sheet) > 0.02) issues.push({ type: "Divergencia APP x PLANILHA", detail: `${x.label}: app ${x.format === "money" ? money(x.app) : num(x.app)} x planilha ${x.format === "money" ? money(x.sheet) : num(x.sheet)} em ${x.origin}. Motivo: ${x.reason}` });
    });
  }
  const discountKeys = new Set();
  allDiscounts().forEach((x) => {
    if (normalize(x.partner) === "gm") issues.push({ type: "GM separado de GUILHERME", detail: `${x.sheetOriginal || ""} ${x.rider} ${money(x.value)}` });
    const key = x.importKey || uniqueKey([x.partner || "BASE", x.rider, x.type, moneyKey(x.value), x.observation || x.reason || x.note, x.lineOriginal, x.columnOriginal]);
    if (discountKeys.has(key)) issues.push({ type: "Desconto duplicado", detail: `${x.partner || "BASE"} ${x.rider} ${x.type} ${money(x.value)} ${x.origin || ""}` });
    discountKeys.add(key);
    if (isTotalText(x.rider) || isTotalText(x.observation) || isTotalText(x.reason)) issues.push({ type: "Total importado como item", detail: `${x.origin || ""} ${x.rider} ${money(x.value)}` });
    if (x.source !== "manual" && !x.origin) issues.push({ type: "Desconto sem origem rastreavel", detail: `${x.partner || "BASE"} ${x.rider} ${money(x.value)}` });
  });
  packageComparisonRows().forEach((x) => {
    if (x.diff !== 0) issues.push({ type: "Divergencia entrada x saida", detail: `${x.label}: entrada ${num(x.entry)}, saida ${num(x.exit)}, diferenca ${num(x.diff)}.` });
  });
  if ((dashboardTotals().gross || dashboardTotals().basePay) === 0 && (allDaily().length || allBaseEntries().length)) issues.push({ type: "Dashboard zerado", detail: "Existem dados importados/cadastrados, mas o dashboard calculou zero." });
  if (workbook.sheetCount !== sheets.length) issues.push({ type: "Aba não importada", detail: `${workbook.sheetCount} encontradas x ${sheets.length} importadas.` });
  const cols = sheets.reduce((s, x) => s + (x.columnCount || 0), 0), importedCols = sheets.reduce((s, x) => s + (x.importedColumnCount || 0), 0);
  if (cols !== importedCols) issues.push({ type: "Coluna não importada", detail: `${cols} encontradas x ${importedCols} importadas.` });
  allBaseEntries().forEach((x) => { if (!x.partner) issues.push({ type: "Entrada de base sem sócio", detail: displayDate(x.date) }); if (!x.date) issues.push({ type: "Entrada de base sem data", detail: x.partner }); });
  allDiscounts().forEach((x) => { if (!x.reason) issues.push({ type: "Desconto sem motivo", detail: `${x.rider} - ${money(x.value)}` }); if (x.type === "Extravio" && !x.value) issues.push({ type: "Extravio sem valor", detail: x.rider }); if (x.type === "Vale" && !x.date) issues.push({ type: "Vale sem data", detail: x.rider }); if (x.type === "Vale" && !x.rider) issues.push({ type: "Vale sem motoboy", detail: money(x.value) }); if (x.value < 0) issues.push({ type: "Valor negativo", detail: `${x.rider} - ${money(x.value)}` }); });
  allDaily().forEach((x) => { const expected = x.ml * x.rateMl + x.shopee * x.rateShopee + x.avulso * x.rateAvulso; if (Math.abs(expected - x.gross) > 0.02) issues.push({ type: "Total bruto errado", detail: `${x.rider} em ${displayDate(x.date)}: esperado ${money(expected)}, atual ${money(x.gross)}.` }); });
  closingRecords().forEach((c) => { const expected = c.gross - c.vales - c.losses - c.discounts + c.bonuses; if (Math.abs(expected - c.net) > 0.02) issues.push({ type: "Total líquido errado", detail: `${c.rider}: ${money(c.net)} x ${money(expected)}` }); });
  return issues;
}

function alertHtml(issue) { return `<div class="alert"><strong>${escapeHtml(issue.type)}</strong><br>${escapeHtml(issue.detail)}</div>`; }

function updateBaseTotals() {
  const row = baseEntryCalc({ ml: Number($("baseMl").value || 0), shopee: Number($("baseShopee").value || 0) });
  $("baseTotalPackages").value = num(row.totalPackages); $("baseValueMl").value = money(row.valueMl); $("baseValueShopee").value = money(row.valueShopee); $("baseTotalPay").value = money(row.totalPay);
}
function updateDailyGross() { $("dailyGross").value = money(Number($("dailyMl").value || 0) * parseMoney($("dailyRateMl").value) + Number($("dailyShopee").value || 0) * parseMoney($("dailyRateShopee").value) + Number($("dailyAvulso").value || 0) * parseMoney($("dailyRateAvulso").value)); }
function buildDailyRow(id = "") {
  return {
    id: id || uid("daily"),
    source: "manual",
    date: isoDateOnly($("dailyDate").value),
    rider: $("dailyRider").value,
    dailyType: $("dailyType").value,
    ml: Number($("dailyMl").value || 0),
    shopee: Number($("dailyShopee").value || 0),
    avulso: Number($("dailyAvulso").value || 0),
    rateMl: parseMoney($("dailyRateMl").value),
    rateShopee: parseMoney($("dailyRateShopee").value),
    rateAvulso: parseMoney($("dailyRateAvulso").value),
    gross: parseMoney($("dailyGross").value),
    responsible: $("dailyResponsible").value.trim(),
    note: $("dailyNote").value.trim()
  };
}
function findDailyLaunch(date, rider) {
  const key = dailyIdentity({ date, rider });
  return (state.daily || []).find((row) => dailyIdentity(row) === key);
}
function fillDailyForm(row) {
  if (!row) return;
  editingDailyId = row.id || "";
  $("dailyDate").value = isoDateOnly(row.date) || todayKey();
  $("dailyRider").value = row.rider || "";
  setDailyType(dailyTypeFor(row), false);
  $("dailyMl").value = row.ml || 0;
  $("dailyShopee").value = row.shopee || 0;
  $("dailyAvulso").value = row.avulso || 0;
  $("dailyRateMl").value = money(row.rateMl);
  $("dailyRateShopee").value = money(row.rateShopee);
  $("dailyRateAvulso").value = money(row.rateAvulso);
  $("dailyResponsible").value = row.responsible || "";
  $("dailyNote").value = row.note || "";
  updateDailyGross();
}
function clearDailyLaunchFields() {
  const r = riderByName($("dailyRider")?.value);
  editingDailyId = "";
  if (r) setDailyType(dailyTypeFromRider(r));
  $("dailyMl").value = 0;
  $("dailyShopee").value = 0;
  $("dailyAvulso").value = 0;
  $("dailyResponsible").value = "";
  $("dailyNote").value = "";
  updateDailyGross();
}
function fillDailyForSelectedDate() {
  const row = findDailyLaunch($("dailyDate")?.value, $("dailyRider")?.value);
  if (row) {
    fillDailyForm(row);
    showDailyFeedback("Lançamento carregado do Supabase.", "ok");
    return true;
  }
  clearDailyLaunchFields();
  showDailyFeedback("Nenhum lançamento salvo para esta data e motoboy.", "ok");
  return false;
}
async function saveDailyLaunchFromForm() {
  updateDailyGross();
  const existing = findDailyLaunch($("dailyDate").value, $("dailyRider").value);
  const id = editingDailyId || existing?.id || uid("daily");
  const row = buildDailyRow(id);
  if (!row.date || !row.rider) {
    showDailyFeedback("Informe data e motoboy para salvar.", "error");
    return;
  }
  if (!requireWrite("daily", row)) return;
  const button = $("dailyForm")?.querySelector('button[type="submit"]');
  if (button) button.disabled = true;
  showDailyFeedback("Salvando lançamento no Supabase...", "ok");
  try {
    const saved = await persistRecord("daily", row, { throwOnError: true });
    const index = state.daily.findIndex((item) => item.id === saved.id || dailyIdentity(item) === dailyIdentity(saved));
    if (index >= 0) state.daily[index] = saved; else state.daily.unshift(saved);
    editingDailyId = saved.id;
    await syncFromSupabase();
    const persisted = findDailyLaunch(saved.date, saved.rider) || saved;
    fillDailyForm(persisted);
    localStorage.setItem(STORE_KEY, JSON.stringify(state));
    showDailyFeedback("Lançamento salvo com sucesso.", "ok");
    renderAll();
    fillDailyForm(persisted);
    showDailyFeedback("Lançamento salvo com sucesso.", "ok");
  } catch (error) {
    console.error("Erro ao salvar lançamento diário:", error);
    const message = /row-level security|RLS/i.test(error?.message || "")
      ? "Erro de permissão no Supabase/RLS. Rode o arquivo supabase-daily-persistence-fix.sql no Supabase SQL Editor."
      : (error?.message || "Erro ao salvar lançamento diário.");
    showDailyFeedback(message, "error");
  } finally {
    if (button) button.disabled = false;
  }
}
function fillDefaults() { const today = new Date().toISOString().slice(0, 10); ["baseDate", "dailyDate", "discountDate", "expenseDate", "paymentDate"].forEach((id) => { if ($(id)) $(id).value = today; }); setDailyType("com coleta"); updateBaseTotals(); updateDailyGross(); updateExpensePeriodFields(); }

function setMobileMenu(open) {
  const sidebar = $("sidebar");
  const toggle = $("mobileMenuToggle");
  const backdrop = $("mobileMenuBackdrop");
  document.body.classList.toggle("mobile-menu-open", Boolean(open));
  sidebar?.classList.toggle("is-open", Boolean(open));
  if (toggle) toggle.setAttribute("aria-expanded", open ? "true" : "false");
  if (toggle) toggle.setAttribute("aria-label", open ? "Fechar menu" : "Abrir menu");
  if (backdrop) backdrop.hidden = !open;
}

function closeMobileMenu() {
  setMobileMenu(false);
}

function switchView(id) {
  if (!canView(id)) {
    permissionDenied();
    const fallback = isOperatorRole() ? "lancamentos" : permittedViews()[0];
    if (!fallback) return;
    id = fallback;
  }
  document.querySelectorAll(".nav-item, .view").forEach((el) => el.classList.remove("active"));
  const nav = document.querySelector(`.nav-item[data-view="${id}"]`);
  nav?.classList.add("active");
  $(id)?.classList.add("active");
  $("viewTitle").textContent = nav?.textContent || "Dashboard";
  closeMobileMenu();
  window.scrollTo({ top: 0, left: 0, behavior: "auto" });
}
function addMoneyBlur(ids) { ids.forEach((id) => $(id).addEventListener("blur", () => { $(id).value = money(parseMoney($(id).value)); updateDailyGross(); })); }

function showAuthGate(show, message = "") {
  const gate = $("authGate");
  if (!gate) return;
  gate.hidden = !show;
  if ($("loginFeedback")) $("loginFeedback").textContent = message;
}

async function syncFromSupabase() {
  captureLocalMigrationSnapshot();
  if (!isSupabaseConfigured) {
    supabaseOnline = false;
    showSupabaseError(new Error("VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY nao configuradas."));
    renderSupabaseStatus();
    return false;
  }
  const sessionInfo = await getSupabaseSession();
  supabaseSession = sessionInfo.session;
  supabaseProfile = sessionInfo.profile;
  if (!supabaseSession) {
    supabaseOnline = false;
    showAuthGate(true, sessionInfo.error || "Entre para sincronizar com Supabase.");
    renderSupabaseStatus();
    return false;
  }
  try {
    const defaults = defaultState();
    state = await loadCloudState(defaults);
    supabaseOnline = true;
    lastSupabaseSync = new Date().toISOString();
    showAuthGate(false);
    localStorage.setItem(STORE_KEY, JSON.stringify(state));
    if (isAdminRole()) await refreshProfiles();
    if (isOperatorRole()) switchView("lancamentos");
    return true;
  } catch (error) {
    supabaseOnline = false;
    showSupabaseError(error);
    return false;
  } finally {
    renderSupabaseStatus();
  }
}

async function migrateLocalToSupabase() {
  if (!isAdminRole()) return permissionDenied();
  if (!supabaseOnline) {
    const message = "Sem conexao com Supabase. Nao foi possivel salvar.";
    showSupabaseError(new Error(message));
    showBackupFeedback(message, "error");
    return;
  }
  const reportTarget = $("backupMigrationReport");
  const setMigrationReport = (message, counts = null, type = "ok", errors = []) => {
    if (reportTarget) {
      const labels = {
        motoboys: "Motoboys migrados",
        entradas: "Entradas migradas",
        lancamentos: "Lancamentos migrados",
        descontos: "Descontos migrados",
        despesas: "Despesas migradas",
        pagamentos: "Pagamentos migrados",
        recibos: "Recibos migrados",
        configuracoes: "Configuracoes migradas"
      };
      const countRows = counts ? Object.entries(labels).map(([name, label]) => {
        const failed = errors.some((error) => error.key === name || error.bucket === name);
        const marker = failed ? "Falhou" : "✓";
        const value = name === "configuracoes" ? (counts[name] ? "sim" : "pendente") : num(counts[name] || 0);
        return [`${marker} ${label}`, value];
      }) : [];
      const errorRows = errors.map((error) => [error.table || error.bucket || "", error.message || "Erro desconhecido"]);
      reportTarget.innerHTML = `<div class="feedback ${type}">${escapeHtml(message)}</div>${countRows.length ? tableBlock("Resultado da migracao", ["Etapa", "Resultado"], countRows) : ""}${errorRows.length ? tableBlock("Tabelas com erro", ["Tabela", "Erro"], errorRows) : ""}`;
    }
    showBackupFeedback(message, type);
  };
  const snapshot = findBestLocalMigrationSnapshot();
  const source = snapshot.state;
  if (!source || localMigrationScore(source) === 0) {
    const message = "Nenhum dado local antigo encontrado neste navegador.";
    showBackupFeedback(message, "error");
    renderLocalStorageDiagnostics("backupLocalStorageDiagnostics");
    renderLocalStorageDiagnostics("dashboardLocalStorageReport");
    return;
  }
  if (!window.confirm(`${localMigrationSummaryText(source, snapshot.key)}\n\nMigrar estes dados locais para o Supabase agora?\n\nOs registros com o mesmo ID serao atualizados, sem duplicar. O localStorage nao sera apagado.`)) return;
  try {
    const migratedCounts = { motoboys: 0, entradas: 0, lancamentos: 0, descontos: 0, despesas: 0, pagamentos: 0, recibos: 0, configuracoes: 0 };
    const migrationErrors = [];
    setMigrationReport("Migracao iniciada. Lendo Supabase e preparando registros locais...", localMigrationCounts(source), "ok");
    let cloudBefore = defaultState();
    try {
      cloudBefore = await loadCloudState(defaultState());
    } catch (error) {
      migrationErrors.push({ key: "leitura", bucket: "leitura", table: "loadCloudState", message: error.message || String(error) });
    }
    const merged = mergeUniqueCloudState(source, cloudBefore);
    const buckets = [
      ["riders", source.riders, "motoboys", "motoboys"],
      ["baseEntries", source.baseEntries, "entradas", "package_entries"],
      ["daily", source.daily, "lancamentos", "daily_launches"],
      ["discounts", source.discounts, "descontos", "discounts"],
      ["expenses", source.expenses, "despesas", "expenses"],
      ["payments", source.payments, "pagamentos", "payments"],
      ["receipts", source.receipts, "recibos", "receipts"]
    ];
    for (const [bucket, records, countKey, table] of buckets) {
      setMigrationReport(`Migrando ${countKey}...`, migratedCounts, "ok", migrationErrors);
      for (const record of records || []) {
        try {
          const saved = await saveCloudRecord(bucket, record);
          Object.assign(record, saved);
          migratedCounts[countKey] += 1;
        } catch (error) {
          migrationErrors.push({ key: countKey, bucket: countKey, table, message: error.message || String(error) });
        }
      }
    }
    setMigrationReport("Salvando configuracoes e recarregando dados do Supabase...", migratedCounts, "ok", migrationErrors);
    try {
      await saveCloudSettings(merged);
      migratedCounts.configuracoes = 1;
    } catch (error) {
      migrationErrors.push({ key: "configuracoes", bucket: "configuracoes", table: "settings", message: error.message || String(error) });
    }
    try {
      state = await loadCloudState(defaultState());
    } catch (error) {
      migrationErrors.push({ key: "recarregar", bucket: "recarregar", table: "loadCloudState", message: error.message || String(error) });
      state = mergeUniqueCloudState(source, state);
    }
    lastSupabaseSync = new Date().toISOString();
    const status = { migratedAt: new Date().toISOString(), sourceKey: snapshot.key, counts: migratedCounts };
    localStorage.setItem(MIGRATION_MARK_KEY, JSON.stringify(status));
    localStateForMigration = source;
    if ($("supabaseFeedback")) {
      $("supabaseFeedback").textContent = "Dados locais migrados para Supabase.";
      $("supabaseFeedback").className = "feedback ok";
    }
    const finalType = migrationErrors.length ? "error" : "ok";
    const finalMessage = migrationErrors.length
      ? "Migracao concluida parcialmente. Algumas tabelas falharam e estao listadas abaixo."
      : "Dados locais migrados para Supabase. Dashboard atualizado com os dados recarregados.";
    setMigrationReport(finalMessage, migratedCounts, finalType, migrationErrors);
    renderSupabaseStatus();
    renderAll();
    renderLocalStorageDiagnostics("dashboardLocalStorageReport");
  } catch (error) {
    showSupabaseError(error);
    setMigrationReport(`Erro ao migrar para Supabase: ${error.message}`, null, "error");
  }
}
async function initSupabaseApp() {
  ensureSupabasePanel();
  if (!isSupabaseConfigured) {
    supabaseOnline = false;
    showAuthGate(false);
    renderSupabaseStatus();
    return;
  }
  await syncFromSupabase();
  if (supabase) {
    supabase.auth.onAuthStateChange(async () => {
      await syncFromSupabase();
      renderAll();
    });
  }
}

function bindEvents() {
  $("loginForm")?.addEventListener("submit", async (e) => {
    e.preventDefault();
    try {
      await signInSupabase($("loginUser").value, $("loginPassword").value);
      await syncFromSupabase();
      renderAll();
    } catch (error) {
      $("loginFeedback").textContent = error.message || "Falha ao entrar.";
    }
  });
  const activateNav = (btn) => {
    if (!btn?.dataset?.view) return;
    switchView(btn.dataset.view);
  };
  document.querySelectorAll(".nav-item").forEach((btn) => {
    let pointerHandled = false;
    btn.addEventListener("pointerup", (event) => {
      event.preventDefault();
      pointerHandled = true;
      activateNav(btn);
      window.setTimeout(() => { pointerHandled = false; }, 350);
    });
    btn.addEventListener("click", () => {
      if (pointerHandled) return;
      activateNav(btn);
    });
  });
  $("mobileMenuToggle")?.addEventListener("click", () => setMobileMenu(!$("sidebar")?.classList.contains("is-open")));
  $("mobileMenuBackdrop")?.addEventListener("click", closeMobileMenu);
  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeMobileMenu();
  });
  ["quickSearch", "periodFilter", "statusFilter"].forEach((id) => $(id).addEventListener("input", renderAll));
  ["baseMl", "baseShopee"].forEach((id) => $(id).addEventListener("input", updateBaseTotals));
  ["dailyMl", "dailyShopee", "dailyAvulso", "dailyRateMl", "dailyRateShopee", "dailyRateAvulso"].forEach((id) => $(id).addEventListener("input", updateDailyGross));
  $("dailyTypeSwitch")?.addEventListener("click", () => setDailyType($("dailyType").value === "com coleta" ? "sem coleta" : "com coleta"));
  $("dailyTypeSwitch")?.addEventListener("keydown", (e) => { if (e.key !== " " && e.key !== "Enter") return; e.preventDefault(); setDailyType($("dailyType").value === "com coleta" ? "sem coleta" : "com coleta"); });
  $("dailyRider").addEventListener("change", () => {
    if (fillDailyForSelectedDate()) return;
    const r = riderByName($("dailyRider").value);
    if (r) setDailyType(dailyTypeFromRider(r));
    updateDailyGross();
  });
  $("dailyDate").addEventListener("change", async () => {
    $("dailyDate").value = isoDateOnly($("dailyDate").value);
    if (supabaseOnline) await syncFromSupabase();
    renderDaily();
    fillDailyForSelectedDate();
  });
  addMoneyBlur(["dailyRateMl", "dailyRateShopee", "dailyRateAvulso", "discountValue", "expenseValue", "paymentValue", "riderRateMl", "riderRateShopee", "riderRateAvulso", "configMl", "configShopee", "configAvulso"]);
  $("riderCollection").addEventListener("change", applyWorkTypeDefaultRates);
  $("riderForm").addEventListener("submit", (e) => {
    e.preventDefault();
    saveRiderFromForm();
  });
  $("newRider").addEventListener("click", clearRiderForm);
  $("toggleRider").addEventListener("click", toggleSelectedRiderStatus);
  $("removeRider").addEventListener("click", removeSelectedRider);
  const guardSubmit = (formId, bucket, rowFn) => {
    $(formId)?.addEventListener("submit", (event) => {
      const row = rowFn();
      if (requireWrite(bucket, row)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
    }, true);
  };
  guardSubmit("baseForm", "baseEntries", () => ({ partner: $("basePartner")?.value || "", responsible: $("basePartner")?.value || "" }));
  guardSubmit("dailyForm", "daily", () => ({ date: $("dailyDate")?.value || todayKey() }));
  guardSubmit("discountForm", "discounts", () => ({ partner: $("discountPartner")?.value || "" }));
  guardSubmit("expenseForm", "expenses", () => ({ responsible: $("expenseResponsible")?.value || "" }));
  $("dailyForm")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    event.stopImmediatePropagation();
    await saveDailyLaunchFromForm();
  }, true);
  $("baseForm").addEventListener("submit", (e) => { e.preventDefault(); const row = baseEntryCalc({ id: uid("base"), source: "manual", date: $("baseDate").value, partner: $("basePartner").value, ml: Number($("baseMl").value || 0), shopee: Number($("baseShopee").value || 0), note: $("baseNote").value.trim(), responsible: $("baseResponsible").value.trim() }); if (!confirmDuplicate(state.baseEntries, row)) return; state.baseEntries.unshift(row); persistRecord("baseEntries", row); saveState("Entrada de base", row.partner); renderAll(); });
  $("discountForm").addEventListener("submit", (e) => { e.preventDefault(); e.stopImmediatePropagation(); const id = editingDiscountId || uid("disc"); const row = { id, source: "manual", date: $("discountDate").value, partner: $("discountPartner").value, rider: $("discountRider").value, closingRider: $("discountRider").value, riderMatched: true, type: $("discountType").value, value: parseMoney($("discountValue").value), code: $("discountCode").value.trim(), reason: $("discountReason").value.trim(), note: $("discountNote").value.trim(), observation: $("discountNote").value.trim() || $("discountReason").value.trim(), sheetOriginal: "LANCAMENTO MANUAL", lineOriginal: "", columnOriginal: "", origin: "LANCAMENTO MANUAL" }; row.importKey = uniqueKey([row.partner, row.rider, row.type, moneyKey(row.value), row.observation, row.id]); if (!confirmDuplicate(state.discounts, row, editingDiscountId)) return; const idx = state.discounts.findIndex((x) => x.id === id); if (idx >= 0) state.discounts[idx] = row; else state.discounts.unshift(row); editingDiscountId = ""; persistRecord("discounts", row); saveState("Desconto", `${row.partner} -> ${row.rider}`); renderAll(); }, true);
  $("riderList").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-rider]");
    if (!btn) return;
    editingRiderId = btn.dataset.riderId || "";
    selectedRider = btn.dataset.rider;
    fillRiderForm(selectedRider);
    renderMotoboys();
    showRiderFeedback("Cadastro carregado para edição.", "ok");
  });
  $("freelancerList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-rider]");
    if (!btn) return;
    editingRiderId = btn.dataset.riderId || "";
    selectedRider = btn.dataset.rider;
    switchView("motoboys");
    fillRiderForm(selectedRider);
    renderMotoboys();
    showRiderFeedback("Freelancer carregado para edição.", "ok");
  });
  $("freelancerRows")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-freelancer-receipt]");
    if (!btn) return;
    switchView("recibos");
    renderReceipts(btn.dataset.freelancerReceipt);
  });
  $("baseRows")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-remove-base]");
    if (!btn) return;
    const row = state.baseEntries.find((x) => x.id === btn.dataset.removeBase);
    if (!requireWrite("baseEntries", row || {})) { e.preventDefault(); e.stopImmediatePropagation(); }
  }, true);
  $("dailyRows")?.addEventListener("click", (e) => {
    const edit = e.target.closest("button[data-edit-daily]");
    const remove = e.target.closest("button[data-remove-daily]");
    const id = edit?.dataset.editDaily || remove?.dataset.removeDaily;
    if (!id) return;
    const row = state.daily.find((x) => x.id === id);
    if (!requireWrite("daily", row || {})) { e.preventDefault(); e.stopImmediatePropagation(); }
  }, true);
  const guardDiscountClick = (e) => {
    const edit = e.target.closest("button[data-edit-discount]");
    const remove = e.target.closest("button[data-remove-discount]");
    const id = edit?.dataset.editDiscount || remove?.dataset.removeDiscount;
    if (!id) return;
    const row = state.discounts.find((x) => x.id === id);
    if (!requireWrite("discounts", row || {})) { e.preventDefault(); e.stopImmediatePropagation(); }
  };
  $("discountRows")?.addEventListener("click", guardDiscountClick, true);
  $("baseDiscountRows")?.addEventListener("click", guardDiscountClick, true);
  $("expenseRows")?.addEventListener("click", (e) => {
    const edit = e.target.closest("button[data-edit-expense]");
    const remove = e.target.closest("button[data-remove-expense]");
    const id = edit?.dataset.editExpense || remove?.dataset.removeExpense;
    if (!id) return;
    const row = state.expenses.find((x) => x.id === id);
    if (!requireWrite("expenses", row || {})) { e.preventDefault(); e.stopImmediatePropagation(); }
  }, true);
  $("dailyRows")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-edit-daily]");
    if (!btn) return;
    const row = state.daily.find((x) => x.id === btn.dataset.editDaily);
    if (!row) return;
    fillDailyForm(row);
    showDailyFeedback("Lançamento carregado para edição.", "ok");
  });
  $("baseRows").addEventListener("click", (e) => { const btn = e.target.closest("button[data-remove-base]"); if (!btn) return; const row = state.baseEntries.find((x) => x.id === btn.dataset.removeBase); if (!row || !window.confirm("Remover esta entrada de base?")) return; state.baseEntries = state.baseEntries.filter((x) => x.id !== row.id); removeCloudRecord("baseEntries", row); saveState("Entrada de base removida", row.partner); renderAll(); });
  $("dailyRows").addEventListener("click", async (e) => { const btn = e.target.closest("button[data-remove-daily]"); if (!btn) return; const row = state.daily.find((x) => x.id === btn.dataset.removeDaily); if (!row || !window.confirm("Remover este lançamento diário?")) return; try { await removeCloudRecord("daily", row); await syncFromSupabase(); renderAll(); showDailyFeedback("Lançamento removido com sucesso.", "ok"); } catch (error) { console.error("Erro ao remover lançamento diário:", error); showDailyFeedback(error?.message || "Erro ao remover lançamento diário.", "error"); } });
  const discountClick = (e) => {
    const edit = e.target.closest("button[data-edit-discount]");
    const remove = e.target.closest("button[data-remove-discount]");
    if (edit) {
      const row = state.discounts.find((x) => x.id === edit.dataset.editDiscount);
      if (!row) return;
      editingDiscountId = row.id;
      $("discountDate").value = row.date || "";
      $("discountPartner").value = row.partner || "BASE";
      $("discountRider").value = row.rider || "";
      $("discountType").value = row.type || "Outro desconto";
      $("discountValue").value = money(row.value);
      $("discountCode").value = row.code || "";
      $("discountReason").value = row.reason || "";
      $("discountNote").value = row.note || row.observation || "";
      switchView("descontos");
    }
    if (remove) {
      const row = state.discounts.find((x) => x.id === remove.dataset.removeDiscount);
      if (!row || !window.confirm("Remover este desconto?")) return;
      state.discounts = state.discounts.filter((x) => x.id !== row.id);
      removeCloudRecord("discounts", row);
      saveState("Desconto removido", `${row.partner} -> ${row.rider}`);
      renderAll();
    }
  };
  $("discountRows").addEventListener("click", discountClick);
  $("baseDiscountRows")?.addEventListener("click", discountClick);
  $("expenseForm")?.addEventListener("submit", (e) => {
    e.preventDefault();
    const id = editingExpenseId || uid("expense");
    const type = normalizeExpenseType($("expenseType").value);
    if (type === "variavel") {
      const missing = !$("expenseDate").value || !$("expenseResponsible").value || !$("expenseCategory").value.trim() || !$("expenseDescription").value.trim() || !(parseMoney($("expenseValue").value) > 0) || !$("expenseNote").value.trim() || !$("expenseStatus").value;
      if (missing) {
        window.alert("Preencha data, responsavel, categoria, descricao, valor, observacao e status da despesa variavel.");
        return;
      }
    }
    const origin = type === "variavel" ? periodKey($("expenseDate").value, "quinzenal") : { key: "", label: "" };
    const automaticDiscount = type === "variavel" ? variableExpenseDiscountPeriod($("expenseDate").value) : { key: "", label: "" };
    const discountOption = $("expenseDiscountPeriod").selectedOptions[0];
    const discountKey = type === "variavel" ? ($("expenseDiscountPeriod").value || automaticDiscount.key) : "";
    const row = { id, kind: "despesa", date: $("expenseDate").value, type, category: $("expenseCategory").value.trim(), description: $("expenseDescription").value.trim(), value: parseMoney($("expenseValue").value), responsible: normalizeResponsible($("expenseResponsible").value), originPeriodKey: origin.key, originPeriodLabel: origin.label, discountPeriodKey: discountKey, discountPeriodLabel: type === "variavel" ? (discountOption?.textContent || automaticDiscount.label) : "", discountPeriodManual: type === "variavel" && discountKey !== automaticDiscount.key, status: $("expenseStatus").value, note: $("expenseNote").value.trim(), source: "manual" };
    if (!confirmDuplicate(state.expenses, row, editingExpenseId)) return;
    const idx = state.expenses.findIndex((x) => x.id === id);
    if (idx >= 0) state.expenses[idx] = row; else state.expenses.unshift(row);
    editingExpenseId = "";
    persistRecord("expenses", row);
    saveState("Despesa", `${row.type} ${row.category}`);
    renderAll();
  });
  $("expenseType")?.addEventListener("change", () => updateExpensePeriodFields(true));
  $("expenseDate")?.addEventListener("change", () => updateExpensePeriodFields(true));
  $("newExpense")?.addEventListener("click", clearExpenseForm);
  $("expenseRows")?.addEventListener("click", (e) => {
    const edit = e.target.closest("button[data-edit-expense]");
    const remove = e.target.closest("button[data-remove-expense]");
    if (edit) {
      const rawRow = state.expenses.find((x) => x.id === edit.dataset.editExpense);
      if (!rawRow) return;
      const row = enrichExpense(rawRow);
      editingExpenseId = row.id;
      $("expenseId").value = row.id;
      $("expenseDate").value = row.date || "";
      $("expenseType").value = normalizeExpenseType(row.type || "fixa");
      $("expenseCategory").value = row.category || "";
      $("expenseDescription").value = row.description || "";
      $("expenseValue").value = money(row.value);
      $("expenseResponsible").value = row.responsible || "BASE";
      $("expenseStatus").value = row.status || "pendente";
      $("expenseNote").value = row.note || "";
      updateExpensePeriodFields();
      if (row.discountPeriodKey) {
        if (![...$("expenseDiscountPeriod").options].some((option) => option.value === row.discountPeriodKey)) {
          $("expenseDiscountPeriod").insertAdjacentHTML("afterbegin", `<option value="${escapeHtml(row.discountPeriodKey)}">${escapeHtml(row.discountPeriodLabel || row.discountPeriodKey.replace("|", " a "))}</option>`);
        }
        $("expenseDiscountPeriod").value = row.discountPeriodKey;
      }
    }
    if (remove) {
      const row = state.expenses.find((x) => x.id === remove.dataset.removeExpense);
      if (!row || !window.confirm("Remover esta despesa?")) return;
      state.expenses = state.expenses.filter((x) => x.id !== row.id);
      removeCloudRecord("expenses", row);
      saveState("Despesa removida", row.category);
      renderAll();
    }
  });
  $("partnerStrip").addEventListener("click", (e) => { const btn = e.target.closest("button[data-partner]"); if (btn) { selectedPartner = btn.dataset.partner; switchView("socios"); renderPartners(); } });
  $("partnerCards").addEventListener("click", (e) => { const btn = e.target.closest("button[data-partner]"); if (btn) { selectedPartner = btn.dataset.partner; renderPartners(); } });
  $("closingType").addEventListener("change", renderClosings); $("closingRider").addEventListener("change", renderClosings);
  $("closingRows").addEventListener("click", (e) => { const btn = e.target.closest("button[data-receipt]"); if (btn) { switchView("recibos"); renderReceipts(btn.dataset.receipt); } });
  $("receiptRider").addEventListener("change", () => renderReceipts()); $("receiptClosing").addEventListener("change", () => renderReceipts());
  $("markPaid")?.addEventListener("click", (e) => { if (requireWrite("payments", { partner: $("paymentPartner")?.value || "" })) return; e.preventDefault(); e.stopImmediatePropagation(); }, true);
  $("markPaid").addEventListener("click", () => { const c = closingRecords().find((x) => x.id === $("receiptClosing").value); if (!c) return; const partner = $("paymentPartner").value; const date = $("paymentDate").value || new Date().toISOString().slice(0, 10); const value = parseMoney($("paymentValue").value) || c.net; const note = $("paymentNote").value.trim(); state.paid[c.id] = { date, net: value, partner, note }; const payment = { id: uid("pay"), closingId: c.id, rider: c.rider, partner, date, value, note }; state.payments.unshift(payment); persistRecord("payments", payment); const receipt = { id: uid("receipt"), closingId: c.id, rider: c.rider, partner, date, value, note, html: receiptHtml(c) }; state.receipts.unshift(receipt); persistRecord("receipts", receipt); const disc = { id: uid("pay-vale"), source: "manual", date, partner, rider: c.rider, closingRider: c.rider, riderMatched: true, type: "Vale", value, code: "", reason: "Pagamento de motoboy", note, observation: note || "Vale/adiantamento gerado pelo pagamento", sheetOriginal: "PAGAMENTO", lineOriginal: "", columnOriginal: "", origin: "PAGAMENTO" }; disc.importKey = uniqueKey([disc.partner, disc.rider, disc.type, moneyKey(disc.value), disc.observation, disc.id]); state.discounts.unshift(disc); persistRecord("discounts", disc); saveState("Fechamento pago", `${c.rider} por ${partner}`); renderAll(); });
  $("printReceipt").addEventListener("click", () => window.print());
  ["saveConfig", "clearOperational", "exportBackup", "importBackup", "backupMigrateLocal", "clearOperationalBackup"].forEach((id) => {
    $(id)?.addEventListener("click", (e) => {
      if (isAdminRole()) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      permissionDenied();
    }, true);
  });
  $("saveConfig").addEventListener("click", () => { state.config = { ml: parseMoney($("configMl").value), shopee: parseMoney($("configShopee").value), avulso: parseMoney($("configAvulso").value) }; saveState("Configuração", "Valores padrão"); renderAll(); });
  $("syncSupabase")?.addEventListener("click", async () => { await syncFromSupabase(); renderAll(); });
  $("migrateLocalToSupabase")?.addEventListener("click", migrateLocalToSupabase);
  $("dashboardLocalStorageDiagnostic")?.addEventListener("click", () => renderLocalStorageDiagnostics("dashboardLocalStorageReport"));
  $("logoutSupabase")?.addEventListener("click", async () => { await signOutSupabase(); supabaseOnline = false; supabaseSession = null; supabaseProfile = null; showAuthGate(isSupabaseConfigured, "Sessao encerrada."); renderSupabaseStatus(); });
  $("sidebarLogout")?.addEventListener("click", async () => { await signOutSupabase(); supabaseOnline = false; supabaseSession = null; supabaseProfile = null; showAuthGate(isSupabaseConfigured, "Sessao encerrada."); renderSupabaseStatus(); applyPermissions(); });
  $("clearOperational")?.addEventListener("click", () => {
    if (!window.confirm("Limpar lançamentos operacionais? Cadastros de motoboys, valores e regras serão preservados.")) return;
    resetOperationalBuckets();
    state.cleanOperational = true;
    state.cleanOperationalVersion = CLEAN_OPERATIONAL_VERSION;
    saveState("Limpeza operacional", "Lancamentos, fechamentos pagos, descontos, entradas e despesas foram zerados.");
    renderAll();
  });
  $("clearOperationalBackup")?.addEventListener("click", () => clearOperationalData("Backup"));
  $("exportBackup")?.addEventListener("click", exportBackupData);
  $("importBackup")?.addEventListener("click", () => $("backupFile")?.click());
  $("backupMigrateLocal")?.addEventListener("click", migrateLocalToSupabase);
  $("backupFile")?.addEventListener("change", (e) => {
    if (!isAdminRole()) { permissionDenied(); e.target.value = ""; return; }
    const file = e.target.files?.[0];
    if (!file) return;
    if (!window.confirm("Importar este backup e substituir os dados atuais do sistema?")) {
      e.target.value = "";
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      try {
        importBackupText(String(reader.result || ""));
      } catch (error) {
        showBackupFeedback(`Erro ao importar backup: ${error.message}`, "error");
      } finally {
        e.target.value = "";
      }
    };
    reader.readAsText(file);
  });
  $("newUserProfile")?.addEventListener("click", clearUserForm);
  $("saveUserProfile")?.addEventListener("click", saveUserProfile);
  $("usersList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-edit-profile]");
    if (!btn) return;
    if (!isAdminRole()) return permissionDenied();
    const profile = profiles.find((x) => x.id === btn.dataset.editProfile);
    if (!profile) return;
    $("profileId").value = profile.id;
    $("profileAuthId").value = profile.id;
    $("profileUsername").value = profile.username || "";
    $("profileRole").value = profile.role || "OPERADOR";
    $("profileFullName").value = profile.full_name || "";
    $("profileActive").value = profile.active === false ? "false" : "true";
  });
  $("exportExcel").addEventListener("click", exportExcel); $("printPage").addEventListener("click", () => window.print());
}

function exportExcel() { const blob = new Blob([`<!doctype html><html><meta charset="utf-8"><body>${document.querySelector(".view.active").innerHTML}</body></html>`], { type: "application/vnd.ms-excel;charset=utf-8" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = "financeiro-motoboys.xls"; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }

initializeWorkbookData();
ensureSupabasePanel();
bindEvents();
await initSupabaseApp();
renderAll();
fillDefaults();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("/sw.js", { scope: "/" })
      .then((registration) => {
        console.info("Service worker ativo:", registration.scope);
        registration.update();
        if (registration.waiting) registration.waiting.postMessage({ type: "SKIP_WAITING" });
      })
      .catch((error) => {
        console.error("Falha ao registrar service worker:", error);
      });
  });
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (!window.__financeiroReloadedForSw) {
      window.__financeiroReloadedForSw = true;
      window.location.reload();
    }
  });
}
