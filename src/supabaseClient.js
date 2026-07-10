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
  if (raw === "operador" || raw === "operator") return "OPERADOR";
  if (raw === "admin") return "admin";
  return "BASE";
}

function profileFromEmail(user) {
  const email = String(user?.email || "").toLowerCase();
  const username = email.split("@")[0] || "";
  const role = normalizeOwner(username);
  return {
    id: user?.id || "",
    username,
    role: role === "BASE" ? "BASE" : role,
    full_name: username,
    active: true
  };
}

function reconcileProfileWithEmail(profile, user) {
  const fallback = profileFromEmail(user);
  if (!profile) return fallback;
  const emailRole = fallback.role;
  const profileRole = normalizeOwner(profile.role);
  if (emailRole !== "BASE" && profileRole === "BASE") {
    return {
      ...profile,
      username: profile.username || fallback.username,
      role: emailRole,
      full_name: profile.full_name || fallback.full_name,
      active: profile.active !== false
    };
  }
  return profile;
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

function isoDate(year, month, day) {
  return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function periodKey(date) {
  const d = new Date(`${date}T00:00:00`);
  if (Number.isNaN(d.getTime())) return { key: "", label: "", start: null, end: null };
  const year = d.getFullYear();
  const month = d.getMonth();
  const firstHalf = d.getDate() <= 15;
  const start = isoDate(year, month, firstHalf ? 1 : 16);
  const end = isoDate(year, month, firstHalf ? 15 : new Date(year, month + 1, 0).getDate());
  const monthLabel = String(month + 1).padStart(2, "0");
  return { key: `${start}|${end}`, label: `${firstHalf ? "1ª" : "2ª"} quinzena ${monthLabel}/${year}`, start, end };
}

function variableExpenseDiscountPeriod(date) {
  const d = new Date(`${date}T00:00:00`);
  if (Number.isNaN(d.getTime())) return periodKey(date || new Date().toISOString().slice(0, 10));
  const year = d.getFullYear();
  const month = d.getMonth();
  return d.getDate() <= 15
    ? periodKey(isoDate(year, month, 16))
    : periodKey(isoDate(month === 11 ? year + 1 : year, month === 11 ? 0 : month + 1, 1));
}

function normalizeExpenseType(value) {
  return String(value || "").trim().toLowerCase().includes("fix") ? "fixa" : "variavel";
}

export async function getSupabaseSession() {
  if (!supabase) return { session: null, profile: null, error: "Supabase nao configurado." };
  const { data, error } = await supabase.auth.getSession();
  if (error || !data.session) return { session: data?.session || null, profile: null, error: error?.message || "" };
  const { data: profile, error: profileError } = await supabase.from("profiles").select("*").eq("id", data.session.user.id).maybeSingle();
  if (profileError) return { session: data.session, profile: profileFromEmail(data.session.user), error: profileError.message || "" };
  return { session: data.session, profile: reconcileProfileWithEmail(profile, data.session.user), error: "" };
}

export async function signInSupabase(login, password) {
  if (!supabase) throw new Error("Supabase nao configurado.");
  const email = loginToEmail(login);
  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) {
    error.message = `Falha ao entrar com ${email}: ${error.message}`;
    throw error;
  }
  return data;
}

export async function signOutSupabase() {
  if (!supabase) return;
  await supabase.auth.signOut();
}

export async function loadProfiles() {
  if (!supabase) throw new Error("Supabase nao configurado.");
  const { data, error } = await supabase.from("profiles").select("*").order("username", { ascending: true });
  if (error) throw error;
  return data || [];
}

export async function loadCloudBaseEntry(date, owner) {
  if (!supabase) throw new Error("Supabase nao configurado.");
  const normalizedOwner = normalizeOwner(owner);
  const { data, error } = await supabase
    .from("package_entries")
    .select("*")
    .eq("entry_date", date)
    .eq("owner", normalizedOwner)
    .order("updated_at", { ascending: false })
    .limit(1)
    .maybeSingle();
  if (error) throw error;
  return data ? rowToRecord(data, "package_entries") : null;
}

export async function saveProfile(profile) {
  if (!supabase) throw new Error("Supabase nao configurado.");
  const row = {
    id: profile.id,
    username: profile.username,
    role: profile.role,
    full_name: profile.full_name || profile.username,
    active: profile.active !== false,
    updated_at: new Date().toISOString()
  };
  const { data, error } = await supabase.from("profiles").upsert(row, { onConflict: "id" }).select("*").single();
  if (error) throw error;
  return data;
}

function dbWorkTypeToApp(value) {
  const raw = String(value || "").trim().toLowerCase();
  return raw === "sem_coleta" ? "sem coleta" : "com coleta";
}

function rowToRecord(row, table = "") {
  if (table === "package_entries") {
    const data = row.data || {};
    return {
      ...data,
      id: data.id || row.id,
      supabaseId: row.id,
      source: data.source || "manual",
      date: data.date || row.entry_date || "",
      partner: data.partner || row.responsible || row.owner || "BASE",
      responsible: data.responsible || row.responsible || row.owner || "BASE",
      ml: Number(data.ml ?? row.ml_qty ?? 0),
      shopee: Number(data.shopee ?? row.shopee_qty ?? 0),
      totalPackages: Number(data.totalPackages ?? row.total_packages ?? 0),
      rateMl: Number(data.rateMl ?? row.rate_ml ?? 8),
      rateShopee: Number(data.rateShopee ?? row.rate_shopee ?? 5),
      valueMl: Number(data.valueMl ?? row.value_ml ?? 0),
      valueShopee: Number(data.valueShopee ?? row.value_shopee ?? 0),
      totalPay: Number(data.totalPay ?? row.total_value ?? 0),
      status: data.status || row.status || "pendente",
      note: data.note || row.note || "",
      owner: row.owner || data.owner || row.responsible || "BASE"
    };
  }
  if (table === "daily_launches") {
    const data = row.data || {};
    return {
      ...data,
      id: data.id || row.id,
      supabaseId: row.id,
      source: data.source || "manual",
      date: data.date || row.launch_date || "",
      rider: data.rider || row.motoboy_name || "",
      dailyType: data.dailyType || dbWorkTypeToApp(row.launch_type),
      ml: Number(data.ml ?? row.ml_qty ?? 0),
      shopee: Number(data.shopee ?? row.shopee_qty ?? 0),
      avulso: Number(data.avulso ?? row.avulso_qty ?? 0),
      rateMl: Number(data.rateMl ?? row.rate_ml ?? 0),
      rateShopee: Number(data.rateShopee ?? row.rate_shopee ?? 0),
      rateAvulso: Number(data.rateAvulso ?? row.rate_avulso ?? 0),
      gross: Number(data.gross ?? row.gross_total ?? 0),
      responsible: data.responsible || row.responsible_name || "",
      note: data.note || row.note || "",
      owner: row.owner || data.owner || "BASE"
    };
  }
  return { ...(row.data || {}), id: row.id, supabaseId: row.id, owner: row.owner || row.data?.owner };
}

function recordToRow(record, table) {
  const id = record.supabaseId || record.id || crypto.randomUUID();
  const data = { ...record, id, supabaseId: id };
  const base = {
    id,
    owner: table === "motoboys" || table === "daily_launches" ? "BASE" : ownerFromRecord(record),
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
    const type = normalizeExpenseType(record.type || "variavel");
    const calculatedOrigin = type === "variavel" ? periodKey(record.date) : { key: "", label: "", start: null, end: null };
    const calculatedDiscount = type === "variavel" ? variableExpenseDiscountPeriod(record.date) : { key: "", label: "", start: null, end: null };
    const origin = type === "variavel" ? splitPeriodKey(record.originPeriodKey || calculatedOrigin.key) : { start: null, end: null };
    const discount = type === "variavel" ? splitPeriodKey(record.discountPeriodKey || calculatedDiscount.key) : { start: null, end: null };
    return {
      ...base,
      expense_date: toDate(record.date),
      responsible: ownerFromRecord(record),
      expense_type: type,
      category: cleanText(record.category || "Sem categoria"),
      description: cleanText(record.description || record.category || "Despesa"),
      value: toNumber(record.value),
      observation: cleanText(record.note || record.observation),
      origin_period_start: origin.start || calculatedOrigin.start,
      origin_period_end: origin.end || calculatedOrigin.end,
      origin_period_label: cleanText(type === "variavel" ? (record.originPeriodLabel || calculatedOrigin.label) : ""),
      discount_period_start: discount.start || calculatedDiscount.start,
      discount_period_end: discount.end || calculatedDiscount.end,
      discount_period_label: cleanText(type === "variavel" ? (record.discountPeriodLabel || calculatedDiscount.label) : ""),
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
    const query = supabase.from(table).select("*");
    const orderedQuery = table === "motoboys"
      ? query.order("name", { ascending: true })
      : query.order("created_at", { ascending: false });
    const { data, error } = await orderedQuery;
    if (error) throw error;
    next[bucket] = (data || []).map((row) => rowToRecord(row, table));
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
  const { data, error } = await supabase.from(table).upsert(row, { onConflict: "id" }).select("*").single();
  if (error) throw error;
  return rowToRecord(data, table);
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
