import { createClient } from "@supabase/supabase-js";

export const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "";
export const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || "";
export const isSupabaseConfigured = Boolean(SUPABASE_URL && SUPABASE_ANON_KEY);

export const supabase = isSupabaseConfigured
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: {
      persistSession: true,
      autoRefreshToken: true,
      detectSessionInUrl: true
    }
  })
  : null;

export const TABLES = {
  riders: "motoboys",
  daily: "daily_launches",
  baseEntries: "package_entries",
  discounts: "discounts",
  expenses: "expenses",
  payments: "payments",
  receipts: "receipts"
};

export function loginToEmail(login) {
  const value = String(login || "").trim().toLowerCase();
  return value.includes("@") ? value : `${value}@financeiro.local`;
}

export function normalizeOwner(value) {
  const raw = String(value || "BASE").trim().toLowerCase().replace(/\./g, "");
  if (raw === "gil") return "GIL";
  if (raw === "sales") return "SALES";
  if (raw === "guilherme" || raw === "guilherme m" || raw === "gm") return "GUILHERME";
  if (raw === "admin") return "admin";
  return "BASE";
}

export function ownerFromRecord(record = {}) {
  return normalizeOwner(record.partner || record.responsible || record.basePartner || record.owner || "BASE");
}

function toNumber(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function toInt(value, fallback = 0) {
  const number = Number.parseInt(value, 10);
  return Number.isFinite(number) ? number : fallback;
}

function toDate(value) {
  return value || null;
}

function cleanText(value) {
  return value == null ? null : String(value);
}

function workTypeToDb(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "freelancer") return "freelancer";
  if (raw === "sem coleta" || raw === "sem_coleta" || raw === "no_collection") return "sem_coleta";
  return "com_coleta";
}

function statusToDb(value, fallback = "pendente") {
  const raw = String(value || fallback).trim().toLowerCase();
  if (raw === "conferido" || raw === "pago" || raw === "ativo" || raw === "inativo" || raw === "lancado_no_fechamento") return raw;
  if (raw === "lançado no fechamento") return "lancado_no_fechamento";
  return fallback;
}

function discountTypeToDb(value) {
  const raw = String(value || "OUTROS").trim().toUpperCase();
  if (raw.includes("VALE")) return "VALE";
  if (raw.includes("EXTRAVIO") && (raw.includes("OCORR") || raw.includes("OCORRÊNCIA"))) return "EXTRAVIO/OCORRENCIA";
  if (raw.includes("EXTRAVIO")) return "EXTRAVIO";
  if (raw.includes("OCORR")) return "OCORRENCIA";
  return "OUTROS";
}

function splitPeriodKey(key = "") {
  const [start = "", end = ""] = String(key || "").split("|");
  return { start: start || null, end: end || null };
}

export async function getSupabaseSession() {
  if (!supabase) return { session: null, profile: null, error: "Supabase nao configurado." };
  const { data, error } = await supabase.auth.getSession();
  if (error || !data.session) return { session: data?.session || null, profile: null, error: error?.message || "" };
  const { data: profile } = await supabase.from("profiles").select("*").eq("id", data.session.user.id).maybeSingle();
  return { session: data.session, profile: profile || null, error: "" };
}

export async function signInSupabase(login, password) {
  if (!supabase) throw new Error("Supabase nao configurado.");
  const { data, error } = await supabase.auth.signInWithPassword({ email: loginToEmail(login), password });
  if (error) throw error;
  return data;
}

export async function signOutSupabase() {
  if (!supabase) return;
  await supabase.auth.signOut();
}

function rowToRecord(row) {
  return { ...(row.data || {}), id: row.id, supabaseId: row.id, owner: row.owner || row.data?.owner };
}

function recordToRow(record, table) {
  const id = record.supabaseId || record.id || crypto.randomUUID();
  const data = { ...record, id, supabaseId: id };
  const base = {
    id,
    owner: table === "motoboys" ? "BASE" : ownerFromRecord(record),
    data,
    updated_at: new Date().toISOString()
  };

  if (table === "motoboys") {
    const type = workTypeToDb(record.collection || record.workType || record.type);
    return {
      ...base,
      name: record.name || record.rider || "Motoboy sem nome",
      region: cleanText(record.region),
      work_type: type,
      has_collection: type === "com_coleta",
      rate_ml: toNumber(record.rateMl ?? record.mlValue, type === "com_coleta" ? 8 : 6),
      rate_shopee: toNumber(record.rateShopee ?? record.shopeeValue, type === "com_coleta" ? 5 : 4),
      rate_avulso: toNumber(record.rateAvulso ?? record.avulsoValue, type === "com_coleta" ? 8 : 6),
      status: statusToDb(record.status, "ativo"),
      notes: cleanText(record.notes || record.note || record.observations)
    };
  }

  if (table === "daily_launches") {
    const launchType = workTypeToDb(record.dailyType || record.launchType || record.collection);
    const rateMl = toNumber(record.rateMl, launchType === "com_coleta" ? 8 : 6);
    const rateShopee = toNumber(record.rateShopee, launchType === "com_coleta" ? 5 : 4);
    const rateAvulso = toNumber(record.rateAvulso, launchType === "com_coleta" ? 8 : 6);
    const ml = toInt(record.ml);
    const shopee = toInt(record.shopee);
    const avulso = toInt(record.avulso);
    const gross = toNumber(record.gross, (ml * rateMl) + (shopee * rateShopee) + (avulso * rateAvulso));
    return {
      ...base,
      launch_date: toDate(record.date),
      motoboy_name: cleanText(record.rider || record.motoboyName),
      launch_type: launchType,
      ml_qty: ml,
      shopee_qty: shopee,
      avulso_qty: avulso,
      rate_ml: rateMl,
      rate_shopee: rateShopee,
      rate_avulso: rateAvulso,
      gross_total: gross,
      net_total: gross,
      responsible_name: cleanText(record.responsible || record.partner),
      status: statusToDb(record.status, "pendente"),
      note: cleanText(record.note)
    };
  }

  if (table === "package_entries") {
    const ml = toInt(record.ml);
    const shopee = toInt(record.shopee);
    const rateMl = toNumber(record.rateMl, 8);
    const rateShopee = toNumber(record.rateShopee, 5);
    return {
      ...base,
      entry_date: toDate(record.date),
      responsible: ownerFromRecord(record),
      ml_qty: ml,
      shopee_qty: shopee,
      total_packages: toInt(record.totalPackages, ml + shopee),
      rate_ml: rateMl,
      rate_shopee: rateShopee,
      value_ml: toNumber(record.valueMl, ml * rateMl),
      value_shopee: toNumber(record.valueShopee, shopee * rateShopee),
      total_value: toNumber(record.totalPay, (ml * rateMl) + (shopee * rateShopee)),
      status: statusToDb(record.status, "pendente"),
      note: cleanText(record.note)
    };
  }

  if (table === "discounts") {
    return {
      ...base,
      discount_date: toDate(record.date),
      responsible: ownerFromRecord(record),
      motoboy_name: cleanText(record.rider || record.closingRider),
      discount_type: discountTypeToDb(record.type),
      value: toNumber(record.value),
      reason: cleanText(record.reason),
      occurrence: cleanText(record.occurrence),
      package_code: cleanText(record.code),
      observation: cleanText(record.observation || record.note),
      original_sheet: cleanText(record.sheetOriginal || record.origin),
      original_line: record.lineOriginal ? toInt(record.lineOriginal, null) : null,
      original_column: cleanText(record.columnOriginal),
      unique_import_key: cleanText(record.importKey || `${ownerFromRecord(record)}|${record.rider || ""}|${record.type || ""}|${record.value || 0}|${id}`),
      status: statusToDb(record.status, "pendente")
    };
  }

  if (table === "expenses") {
    const origin = splitPeriodKey(record.originPeriodKey);
    const discount = splitPeriodKey(record.discountPeriodKey);
    const type = String(record.type || "variavel").toLowerCase() === "fixa" ? "fixa" : "variavel";
    return {
      ...base,
      expense_date: toDate(record.date),
      responsible: ownerFromRecord(record),
      expense_type: type,
      category: cleanText(record.category || "Sem categoria"),
      description: cleanText(record.description || record.category || "Despesa"),
      value: toNumber(record.value),
      observation: cleanText(record.note || record.observation),
      origin_period_start: origin.start,
      origin_period_end: origin.end,
      origin_period_label: cleanText(record.originPeriodLabel),
      discount_period_start: discount.start,
      discount_period_end: discount.end,
      discount_period_label: cleanText(record.discountPeriodLabel),
      status: statusToDb(record.status, "pendente")
    };
  }

  if (table === "payments") {
    const period = splitPeriodKey(record.periodKey);
    return {
      ...base,
      payment_date: toDate(record.date),
      responsible: ownerFromRecord(record),
      motoboy_name: cleanText(record.rider),
      period_start: period.start,
      period_end: period.end,
      period_label: cleanText(record.period || record.periodLabel),
      net_paid: toNumber(record.value ?? record.net),
      status: statusToDb(record.status, "pago"),
      note: cleanText(record.note)
    };
  }

  if (table === "receipts") {
    const period = splitPeriodKey(record.periodKey);
    return {
      ...base,
      receipt_number: cleanText(record.receiptNumber || record.id),
      payment_id: cleanText(record.paymentId),
      responsible: ownerFromRecord(record),
      motoboy_name: cleanText(record.rider),
      period_start: period.start,
      period_end: period.end,
      period_label: cleanText(record.period || record.periodLabel),
      payment_date: toDate(record.date),
      net_paid: toNumber(record.value ?? record.net),
      observations: cleanText(record.note),
      html: cleanText(record.html)
    };
  }

  return base;
}

export async function loadCloudState(defaultState) {
  if (!supabase) throw new Error("Supabase nao configurado.");
  const next = { ...defaultState, riders: [], daily: [], baseEntries: [], discounts: [], expenses: [], payments: [], receipts: [] };
  await Promise.all(Object.entries(TABLES).map(async ([bucket, table]) => {
    const { data, error } = await supabase.from(table).select("*").order("created_at", { ascending: false });
    if (error) throw error;
    next[bucket] = (data || []).map(rowToRecord);
  }));

  const { data: settings, error: settingsError } = await supabase
    .from("settings")
    .select("*")
    .eq("key", "app_state")
    .maybeSingle();
  if (settingsError) throw settingsError;
  const value = settings?.value || {};
  next.config = { ...next.config, ...(value.config || {}) };
  next.paid = value.paid || {};
  next.basePaid = value.basePaid || {};
  next.audit = Array.isArray(value.audit) ? value.audit : [];
  next.cleanOperational = value.cleanOperational ?? true;
  next.cleanOperationalVersion = value.cleanOperationalVersion || "";
  next.lastBackupAt = value.lastBackupAt || "";
  return next;
}

export async function saveCloudRecord(bucket, record) {
  const table = TABLES[bucket];
  if (!supabase || !table) throw new Error("Sem conexao com Supabase. Nao foi possivel salvar.");
  const row = recordToRow(record, table);
  const { error } = await supabase.from(table).upsert(row, { onConflict: "id" });
  if (error) throw error;
  return row.data;
}

export async function deleteCloudRecord(bucket, record) {
  const table = TABLES[bucket];
  if (!supabase || !table) throw new Error("Sem conexao com Supabase. Nao foi possivel salvar.");
  const id = record?.supabaseId || record?.id;
  if (!id) return;
  const { error } = await supabase.from(table).delete().eq("id", id);
  if (error) throw error;
}

export async function saveCloudSettings(state) {
  if (!supabase) throw new Error("Sem conexao com Supabase. Nao foi possivel salvar.");
  const value = {
    config: state.config || {},
    paid: state.paid || {},
    basePaid: state.basePaid || {},
    audit: state.audit || [],
    cleanOperational: state.cleanOperational,
    cleanOperationalVersion: state.cleanOperationalVersion || "",
    lastBackupAt: state.lastBackupAt || ""
  };
  const { error } = await supabase.from("settings").upsert({ key: "app_state", value, updated_at: new Date().toISOString() }, { onConflict: "key" });
  if (error) throw error;
}

export async function saveBackupToCloud(payload) {
  if (!supabase) throw new Error("Sem conexao com Supabase. Nao foi possivel salvar.");
  const { error } = await supabase.from("backups").insert({ data: payload });
  if (error) throw error;
}

export async function logCloudChange(action, detail, meta = {}) {
  if (!supabase) return;
  await supabase.from("change_logs").insert({ action, detail, metadata: meta }).throwOnError();
}
