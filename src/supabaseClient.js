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
  return {
    id,
    owner: table === "motoboys" ? "BASE" : ownerFromRecord(record),
    data: { ...record, id, supabaseId: id },
    updated_at: new Date().toISOString()
  };
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
