import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  LayoutDashboard, Bike, Users, FileText, Wallet, TrendingUp, TrendingDown,
  Wrench, Gauge, AlertTriangle, Calendar as CalendarIcon, BarChart3, Settings,
  Bell, Search, Plus, Pencil, Trash2, X, ChevronLeft, ChevronRight, LogOut,
  CheckCircle2, AlertCircle, XCircle, Download, ClipboardList, KeyRound,
  ArrowUpRight, ArrowDownRight, Fuel, ShieldCheck, ChevronDown, Filter
} from "lucide-react";
import * as XLSX from "xlsx";

/* =========================================================================
   CONSTANTES
========================================================================= */
const MOTO_STATUS = ["Disponível", "Alugada", "Em manutenção", "Reservada", "Inativa", "Vendida"];
const RENTER_STATUS = ["Ativo", "Inativo", "Bloqueado"];
const CONTRACT_STATUS = ["Ativo", "Finalizado", "Suspenso", "Cancelado"];
const PERIODICITY = ["Semanal", "Quinzenal", "Mensal", "Diário"];
const PAYMENT_METHODS = ["Pix", "Dinheiro", "Transferência", "Cartão", "Outro"];
const PAYMENT_STATUS = ["Pago", "Pendente", "Atrasado", "Parcial", "Cancelado"];
const REVENUE_CATEGORIES = ["Aluguel", "Taxa adicional", "Quilometragem excedente", "Multa", "Danos", "Outros"];
const EXPENSE_CATEGORIES = ["Combustível", "Manutenção", "Peças", "Oficina", "Pneus", "Seguro", "Licenciamento", "IPVA", "Documentação", "Lavagem", "Rastreador", "Depreciação", "Impostos", "Financiamento", "Juros", "Marketing", "Outras despesas"];
const INCIDENT_TYPES = ["Multa", "Acidente", "Dano", "Avaria", "Furto/Roubo", "Quebra", "Outra ocorrência"];
const INCIDENT_RESP = ["Locatário", "Empresa", "Seguradora"];
const INCIDENT_STATUS = ["Aberto", "Em negociação", "Resolvido"];
const DEFAULT_MAINT_TYPES = [
  { id: "mt1", name: "Troca de óleo", kmInterval: 2000, timeIntervalDays: 90 },
  { id: "mt2", name: "Filtro de óleo", kmInterval: 4000, timeIntervalDays: 180 },
  { id: "mt3", name: "Filtro de ar", kmInterval: 6000, timeIntervalDays: 180 },
  { id: "mt4", name: "Pastilhas de freio", kmInterval: 8000, timeIntervalDays: 240 },
  { id: "mt5", name: "Disco de freio", kmInterval: 20000, timeIntervalDays: 720 },
  { id: "mt6", name: "Pneus", kmInterval: 15000, timeIntervalDays: 540 },
  { id: "mt7", name: "Relação (corrente/coroa/pinhão)", kmInterval: 12000, timeIntervalDays: 365 },
  { id: "mt8", name: "Vela", kmInterval: 6000, timeIntervalDays: 365 },
  { id: "mt9", name: "Bateria", kmInterval: 20000, timeIntervalDays: 730 },
  { id: "mt10", name: "Fluido de freio", kmInterval: 10000, timeIntervalDays: 365 },
  { id: "mt11", name: "Suspensão", kmInterval: 20000, timeIntervalDays: 730 },
  { id: "mt12", name: "Embreagem", kmInterval: 20000, timeIntervalDays: 730 },
  { id: "mt13", name: "Revisão geral", kmInterval: 5000, timeIntervalDays: 180 },
];

const STORAGE_KEY = "frota_db_v1";       // usado apenas no Modo local (sem servidor)
const SHELL_STORAGE_KEY = "frota_shell_v1"; // usado no Modo conectado (settings + notificações lidas)

/* =========================================================================
   CLIENTE DE API (Modo conectado — backend real)
   Troque API_BASE pela URL do seu backend hospedado quando for para produção.
========================================================================= */
const API_BASE = "http://localhost:4000";

const SERVER_RESOURCES = {
  motorcycles: "motorcycles", renters: "renters", contracts: "contracts", payments: "payments",
  maintenanceTypes: "maintenance-types", maintenanceRecords: "maintenance-records",
  mileageRecords: "mileage-records", expenses: "expenses", revenues: "revenues",
  incidents: "incidents", users: "users",
};
// Campos aceitos pelo backend por recurso — evita mandar campos só de UI (ex: priceHistory) que o servidor rejeitaria.
const ALLOWED_FIELDS = {
  motorcycles: ["brand","model","year","color","plate","renavam","chassis","acquisitionDate","acquisitionKm","currentKm","acquisitionValue","marketValue","status","lastLicensingDate","nextLicensingDue","insurer","insuranceDue","notes"],
  renters: ["name","cpf","birthDate","phone","whatsapp","email","address","cnh","cnhCategory","cnhValidity","registrationDate","notes","status"],
  contracts: ["renterId","motorcycleId","startDate","endDate","dailyValue","weeklyValue","monthlyValue","deposit","dueDay","periodicity","kmLimit","kmExcessValue","rules","notes","status"],
  payments: ["contractId","renterId","motorcycleId","dueDate","paymentDate","expectedValue","paidValue","paymentMethod","status","notes"],
  maintenanceTypes: ["name","kmInterval","timeIntervalDays"],
  maintenanceRecords: ["motorcycleId","typeId","date","km","partsCost","laborCost","totalCost","vendor","notes","nextKm","nextDate"],
  mileageRecords: ["motorcycleId","date","km","previousKm","kmDriven","recordedBy","notes"],
  expenses: ["date","category","motorcycleId","vendor","value","paymentMethod","notes"],
  revenues: ["date","category","motorcycleId","renterId","value","paymentMethod","notes"],
  incidents: ["motorcycleId","renterId","date","type","value","responsible","status","description"],
};
const ROLE_EN_TO_PT = { ADMIN: "Administrador", MANAGER: "Gerente", EMPLOYEE: "Funcionário" };
const ROLE_PT_TO_EN = { "Administrador": "ADMIN", "Gerente": "MANAGER", "Funcionário": "EMPLOYEE" };
const fromApiUser = (u) => ({ ...u, role: ROLE_EN_TO_PT[u.role] || u.role });
const toApiUser = (u) => ({ name: u.name, username: u.username, ...(u.password ? { password: u.password } : {}), role: ROLE_PT_TO_EN[u.role] || u.role });

// Remove campos não aceitos pelo backend e troca "" por null (datas/relacionamentos opcionais vazios).
function sanitizeForApi(key, obj) {
  const allowed = ALLOWED_FIELDS[key];
  if (!allowed) return obj;
  const out = {};
  allowed.forEach((f) => { if (f in obj) out[f] = obj[f] === "" ? null : obj[f]; });
  return out;
}

async function apiFetch(path, opts = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    ...opts,
    credentials: "include",
    headers: { "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  let data = null;
  try { data = await res.json(); } catch { /* sem corpo */ }
  if (!res.ok) throw new Error(data?.error || `Erro na API (HTTP ${res.status}).`);
  return data;
}
async function apiHealthCheck() {
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 2500);
    const res = await fetch(`${API_BASE}/health`, { signal: ctrl.signal });
    clearTimeout(t);
    return res.ok;
  } catch { return false; }
}
const api = {
  login: async (username, password) => fromApiUser(await apiFetch("/auth/login", { method: "POST", body: JSON.stringify({ username, password }) })),
  logout: () => apiFetch("/auth/logout", { method: "POST" }).catch(() => {}),
  me: async () => fromApiUser(await apiFetch("/auth/me")),
  list: (endpoint) => apiFetch(`/${endpoint}`),
  create: (endpoint, body) => apiFetch(`/${endpoint}`, { method: "POST", body: JSON.stringify(body) }),
  update: (endpoint, id, body) => apiFetch(`/${endpoint}/${id}`, { method: "PUT", body: JSON.stringify(body) }),
  remove: (endpoint, id) => apiFetch(`/${endpoint}/${id}`, { method: "DELETE" }),
};

/* =========================================================================
   HELPERS
========================================================================= */
let _idc = 1;
const uid = (p) => `${p}_${Date.now().toString(36)}_${(_idc++).toString(36)}`;

const fmtBRL = (v) => (Number(v) || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtNum = (v, d = 0) => (Number(v) || 0).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtDate = (v) => {
  if (!v) return "—";
  const d = new Date(v + "T00:00:00");
  if (isNaN(d)) return "—";
  return d.toLocaleDateString("pt-BR");
};
const todayISO = () => new Date().toISOString().slice(0, 10);
const addDays = (iso, days) => {
  const d = new Date(iso + "T00:00:00");
  d.setDate(d.getDate() + Number(days || 0));
  return d.toISOString().slice(0, 10);
};
const daysDiff = (isoA, isoB) => {
  const a = new Date(isoA + "T00:00:00");
  const b = new Date(isoB + "T00:00:00");
  return Math.round((a - b) / 86400000);
};
const startOfWeek = (iso) => {
  const d = new Date(iso + "T00:00:00");
  const day = d.getDay();
  const diff = (day === 0 ? -6 : 1) - day;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
};
const startOfMonth = (iso) => iso.slice(0, 7) + "-01";
const inRange = (dateISO, startISO, endISO) => dateISO >= startISO && dateISO <= endISO;

function downloadBlob(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function exportCSV(filename, columns, rows) {
  const header = columns.map((c) => `"${c.label}"`).join(";");
  const lines = rows.map((r) => columns.map((c) => `"${String(c.value(r) ?? "").replace(/"/g, '""')}"`).join(";"));
  downloadBlob(filename, "\uFEFF" + [header, ...lines].join("\n"), "text/csv;charset=utf-8;");
}
function exportXLSX(filename, columns, rows) {
  const data = rows.map((r) => {
    const o = {};
    columns.forEach((c) => (o[c.label] = c.value(r)));
    return o;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Relatório");
  XLSX.writeFile(wb, filename);
}

/* =========================================================================
   SEED DATA
========================================================================= */
function buildSeed() {
  const motorcycles = [
    { id: "moto1", brand: "Honda", model: "CG 160 Start", year: 2023, color: "Vermelha", plate: "ABC1D23", renavam: "01234567890", chassis: "9C2KC1234RR123456", acquisitionDate: "2023-02-10", acquisitionKm: 0, currentKm: 24800, acquisitionValue: 14500, marketValue: 12800, status: "Alugada", lastLicensingDate: "2025-01-15", nextLicensingDue: "2026-01-15", insurer: "Porto Seguro", insuranceDue: "2026-10-05", notes: "Moto de exemplo." },
    { id: "moto2", brand: "Honda", model: "Biz 125", year: 2022, color: "Prata", plate: "DEF4G56", renavam: "01234567891", chassis: "9C2KC1234RR123457", acquisitionDate: "2022-08-20", acquisitionKm: 0, currentKm: 31200, acquisitionValue: 12800, marketValue: 10200, status: "Alugada", lastLicensingDate: "2025-02-01", nextLicensingDue: "2026-02-01", insurer: "Azul Seguros", insuranceDue: "2026-09-25", notes: "Moto de exemplo." },
    { id: "moto3", brand: "Yamaha", model: "Factor 150", year: 2023, color: "Azul", plate: "GHI7J89", renavam: "01234567892", chassis: "9C2KC1234RR123458", acquisitionDate: "2023-05-05", acquisitionKm: 0, currentKm: 18650, acquisitionValue: 15200, marketValue: 13600, status: "Em manutenção", lastLicensingDate: "2025-01-10", nextLicensingDue: "2026-01-10", insurer: "Porto Seguro", insuranceDue: "2026-11-12", notes: "Moto de exemplo — na oficina." },
    { id: "moto4", brand: "Honda", model: "XRE 300", year: 2021, color: "Preta", plate: "JKL0M12", renavam: "01234567893", chassis: "9C2KC1234RR123459", acquisitionDate: "2021-11-30", acquisitionKm: 500, currentKm: 42100, acquisitionValue: 21000, marketValue: 16500, status: "Disponível", lastLicensingDate: "2025-03-01", nextLicensingDue: "2026-03-01", insurer: "Sem seguro", insuranceDue: "", notes: "Moto de exemplo." },
    { id: "moto5", brand: "Yamaha", model: "Fazer 250", year: 2022, color: "Branca", plate: "NOP3Q45", renavam: "01234567894", chassis: "9C2KC1234RR123460", acquisitionDate: "2022-03-18", acquisitionKm: 0, currentKm: 27900, acquisitionValue: 18700, marketValue: 15300, status: "Alugada", lastLicensingDate: "2025-01-22", nextLicensingDue: "2025-10-01", insurer: "Bradesco Seguros", insuranceDue: "2026-09-30", notes: "Moto de exemplo." },
  ];

  const renters = [
    { id: "r1", name: "Carlos Eduardo Silva", cpf: "111.111.111-11", birthDate: "1990-04-12", phone: "(88) 99111-1111", whatsapp: "(88) 99111-1111", email: "carlos.silva@email.com", address: "Rua das Flores, 100 — Juazeiro do Norte/CE", cnh: "12345678900", cnhCategory: "A", cnhValidity: "2028-06-01", registrationDate: "2024-01-10", notes: "Locatário de exemplo.", status: "Ativo" },
    { id: "r2", name: "Marcos Vinícius Souza", cpf: "222.222.222-22", birthDate: "1988-09-23", phone: "(88) 99222-2222", whatsapp: "(88) 99222-2222", email: "marcos.souza@email.com", address: "Av. Padre Cícero, 500 — Juazeiro do Norte/CE", cnh: "22345678900", cnhCategory: "A", cnhValidity: "2027-02-15", registrationDate: "2024-02-05", notes: "Locatário de exemplo.", status: "Ativo" },
    { id: "r3", name: "Fernanda Alves Lima", cpf: "333.333.333-33", birthDate: "1995-01-30", phone: "(88) 99333-3333", whatsapp: "(88) 99333-3333", email: "fernanda.lima@email.com", address: "Rua São José, 22 — Crato/CE", cnh: "32345678900", cnhCategory: "AB", cnhValidity: "2026-11-20", registrationDate: "2024-03-18", notes: "Locatário de exemplo.", status: "Ativo" },
    { id: "r4", name: "José Ribamar Costa", cpf: "444.444.444-44", birthDate: "1985-07-08", phone: "(88) 99444-4444", whatsapp: "(88) 99444-4444", email: "jose.costa@email.com", address: "Rua do Rosário, 77 — Juazeiro do Norte/CE", cnh: "42345678900", cnhCategory: "A", cnhValidity: "2025-12-01", registrationDate: "2024-04-22", notes: "Atraso recorrente nos últimos pagamentos.", status: "Ativo" },
    { id: "r5", name: "Antônia Beatriz Nunes", cpf: "555.555.555-55", birthDate: "1998-12-02", phone: "(88) 99555-5555", whatsapp: "(88) 99555-5555", email: "beatriz.nunes@email.com", address: "Rua Santa Luzia, 340 — Barbalha/CE", cnh: "52345678900", cnhCategory: "A", cnhValidity: "2029-05-10", registrationDate: "2024-05-30", notes: "Locatária de exemplo.", status: "Inativo" },
  ];

  const today = todayISO();
  const contracts = [
    { id: "c1", renterId: "r1", motorcycleId: "moto1", startDate: "2025-06-01", endDate: addDays(today, 20), dailyValue: 45, weeklyValue: 280, monthlyValue: 1100, deposit: 300, dueDay: "Segunda-feira", periodicity: "Semanal", kmLimit: 0, kmExcessValue: 0, rules: "Combustível por conta do locatário.", notes: "Contrato de exemplo.", status: "Ativo", priceHistory: [{ date: "2025-06-01", weeklyValue: 280 }] },
    { id: "c2", renterId: "r2", motorcycleId: "moto2", startDate: "2025-05-15", endDate: addDays(today, 5), dailyValue: 42, weeklyValue: 260, monthlyValue: 1040, deposit: 300, dueDay: "Sexta-feira", periodicity: "Semanal", kmLimit: 0, kmExcessValue: 0, rules: "Combustível por conta do locatário.", notes: "Contrato de exemplo.", status: "Ativo", priceHistory: [{ date: "2025-05-15", weeklyValue: 260 }] },
    { id: "c3", renterId: "r4", motorcycleId: "moto5", startDate: "2025-04-01", endDate: addDays(today, 2), dailyValue: 48, weeklyValue: 300, monthlyValue: 1200, deposit: 300, dueDay: "Quarta-feira", periodicity: "Semanal", kmLimit: 0, kmExcessValue: 0, rules: "Combustível por conta do locatário.", notes: "Locatário com histórico de atraso.", status: "Ativo", priceHistory: [{ date: "2025-04-01", weeklyValue: 300 }] },
    { id: "c4", renterId: "r3", motorcycleId: "moto4", startDate: "2025-01-10", endDate: "2025-07-10", dailyValue: 45, weeklyValue: 280, monthlyValue: 1100, deposit: 300, dueDay: "Terça-feira", periodicity: "Semanal", kmLimit: 0, kmExcessValue: 0, rules: "", notes: "Contrato finalizado — exemplo.", status: "Finalizado", priceHistory: [{ date: "2025-01-10", weeklyValue: 280 }] },
    { id: "c5", renterId: "r5", motorcycleId: "moto3", startDate: "2025-03-01", endDate: "2025-06-01", dailyValue: 45, weeklyValue: 280, monthlyValue: 1100, deposit: 300, dueDay: "Segunda-feira", periodicity: "Semanal", kmLimit: 0, kmExcessValue: 0, rules: "", notes: "Contrato finalizado — exemplo.", status: "Finalizado", priceHistory: [{ date: "2025-03-01", weeklyValue: 280 }] },
  ];

  // payments: generate several weeks for active contracts + history for finished ones
  const payments = [];
  function genPayments(contract, weeksBack, weeksForward, lateWeeks = []) {
    for (let i = -weeksBack; i <= weeksForward; i++) {
      const due = addDays(startOfWeek(contract.startDate), i * 7 + 7);
      if (due < contract.startDate) continue;
      if (due > (contract.status === "Ativo" ? addDays(today, 7) : contract.endDate)) continue;
      const isFuture = due > today;
      const isLate = lateWeeks.includes(i);
      let status, paymentDate, paidValue;
      if (isFuture) { status = "Pendente"; paymentDate = ""; paidValue = 0; }
      else if (isLate) { status = "Atrasado"; paymentDate = ""; paidValue = 0; }
      else { status = "Pago"; paymentDate = addDays(due, Math.random() > 0.7 ? 1 : 0); paidValue = contract.weeklyValue; }
      payments.push({
        id: uid("pay"), contractId: contract.id, renterId: contract.renterId, motorcycleId: contract.motorcycleId,
        dueDate: due, paymentDate, expectedValue: contract.weeklyValue, paidValue, paymentMethod: status === "Pago" ? "Pix" : "", status, notes: "",
      });
    }
  }
  genPayments(contracts[0], 10, 0);
  genPayments(contracts[1], 8, 0);
  genPayments(contracts[2], 12, 0, [-1, 0]); // José Ribamar late
  genPayments(contracts[3], 25, 0);
  genPayments(contracts[4], 12, 0);

  const mileageRecords = [];
  function genMileage(motoId, startKm, endKm, weeks) {
    let km = startKm;
    const step = Math.round((endKm - startKm) / weeks);
    let date = addDays(today, -7 * weeks);
    let prev = startKm;
    for (let i = 0; i < weeks; i++) {
      km = i === weeks - 1 ? endKm : km + step + Math.round((Math.random() - 0.5) * 60);
      mileageRecords.push({ id: uid("km"), motorcycleId: motoId, date, km, previousKm: prev, kmDriven: km - prev, recordedBy: "Administrador", notes: "" });
      prev = km; date = addDays(date, 7);
    }
  }
  genMileage("moto1", 23400, 24800, 8);
  genMileage("moto2", 29600, 31200, 8);
  genMileage("moto3", 17800, 18650, 6);
  genMileage("moto4", 41200, 42100, 6);
  genMileage("moto5", 26500, 27900, 8);

  const maintenanceTypes = DEFAULT_MAINT_TYPES;
  const maintenanceRecords = [
    { id: uid("mnt"), motorcycleId: "moto1", typeId: "mt1", date: addDays(today, -35), km: 23200, partsCost: 45, laborCost: 25, totalCost: 70, vendor: "Oficina do Zé", notes: "", nextKm: 25200, nextDate: addDays(addDays(today, -35), 90) },
    { id: uid("mnt"), motorcycleId: "moto1", typeId: "mt4", date: addDays(today, -80), km: 21800, partsCost: 90, laborCost: 40, totalCost: 130, vendor: "Oficina do Zé", notes: "", nextKm: 29800, nextDate: addDays(addDays(today, -80), 240) },
    { id: uid("mnt"), motorcycleId: "moto2", typeId: "mt1", date: addDays(today, -20), km: 30600, partsCost: 45, laborCost: 25, totalCost: 70, vendor: "MotoPeças Central", notes: "", nextKm: 32600, nextDate: addDays(addDays(today, -20), 90) },
    { id: uid("mnt"), motorcycleId: "moto2", typeId: "mt6", date: addDays(today, -150), km: 26000, partsCost: 420, laborCost: 60, totalCost: 480, vendor: "Pneus Norte", notes: "Par de pneus novos.", nextKm: 41000, nextDate: addDays(addDays(today, -150), 540) },
    { id: uid("mnt"), motorcycleId: "moto3", typeId: "mt13", date: addDays(today, -3), km: 18650, partsCost: 180, laborCost: 120, totalCost: 300, vendor: "Oficina Central", notes: "Revisão geral — moto ainda na oficina.", nextKm: 23650, nextDate: addDays(addDays(today, -3), 180) },
    { id: uid("mnt"), motorcycleId: "moto4", typeId: "mt1", date: addDays(today, -95), km: 40100, partsCost: 45, laborCost: 25, totalCost: 70, vendor: "Oficina do Zé", notes: "", nextKm: 42100, nextDate: addDays(addDays(today, -95), 90) },
    { id: uid("mnt"), motorcycleId: "moto5", typeId: "mt7", date: addDays(today, -200), km: 22000, partsCost: 320, laborCost: 80, totalCost: 400, vendor: "MotoPeças Central", notes: "Relação completa.", nextKm: 34000, nextDate: addDays(addDays(today, -200), 365) },
  ];

  const revenues = [
    { id: uid("rev"), date: addDays(today, -3), category: "Taxa adicional", motorcycleId: "moto1", renterId: "r1", value: 30, paymentMethod: "Pix", notes: "Entrega em domicílio." },
    { id: uid("rev"), date: addDays(today, -10), category: "Multa", motorcycleId: "moto5", renterId: "r4", value: 130, paymentMethod: "Pix", notes: "Multa repassada ao locatário." },
    { id: uid("rev"), date: addDays(today, -1), category: "Quilometragem excedente", motorcycleId: "moto2", renterId: "r2", value: 45, paymentMethod: "Dinheiro", notes: "" },
  ];

  const expenses = [
    { id: uid("exp"), date: addDays(today, -2), category: "Combustível", motorcycleId: "moto4", vendor: "Posto Ipiranga", value: 60, paymentMethod: "Pix", notes: "Abastecimento antes da entrega." },
    { id: uid("exp"), date: addDays(today, -6), category: "Lavagem", motorcycleId: "moto3", vendor: "Lava-rápido São Miguel", value: 25, paymentMethod: "Dinheiro", notes: "" },
    { id: uid("exp"), date: addDays(today, -15), category: "IPVA", motorcycleId: "moto1", vendor: "Detran-CE", value: 180, paymentMethod: "Transferência", notes: "Parcela 2/3." },
    { id: uid("exp"), date: addDays(today, -22), category: "Seguro", motorcycleId: "moto5", vendor: "Bradesco Seguros", value: 210, paymentMethod: "Cartão", notes: "Parcela mensal." },
    { id: uid("exp"), date: addDays(today, -5), category: "Marketing", motorcycleId: "", vendor: "Instagram Ads", value: 50, paymentMethod: "Cartão", notes: "Divulgação de vagas." },
  ];

  const incidents = [
    { id: uid("inc"), motorcycleId: "moto5", renterId: "r4", date: addDays(today, -10), type: "Multa", value: 130, responsible: "Locatário", status: "Resolvido", description: "Avanço de sinal registrado pelo órgão de trânsito." },
    { id: uid("inc"), motorcycleId: "moto3", renterId: "r5", date: addDays(today, -4), type: "Quebra", value: 300, responsible: "Empresa", status: "Em negociação", description: "Motor apresentou ruído — moto está em revisão geral." },
  ];

  const settings = {
    companyName: "Frota Fácil Motos",
    address: "Juazeiro do Norte/CE",
    phone: "(88) 90000-0000",
    logo: "",
    currency: "BRL",
    defaultRentalWeekly: 280,
    alertKmWarning: 500,
    alertDaysWarning: 15,
    expenseCategories: EXPENSE_CATEGORIES,
    revenueCategories: REVENUE_CATEGORIES,
    paymentMethods: PAYMENT_METHODS,
  };

  const users = [{ id: "u1", name: "Administrador", username: "admin", password: "admin123", role: "Administrador" }];

  return { seeded: true, motorcycles, renters, contracts, payments, mileageRecords, maintenanceTypes, maintenanceRecords, revenues, expenses, incidents, users, readNotifications: [], settings };
}

function emptyDb() {
  const s = buildSeed();
  return { seeded: false, motorcycles: [], renters: [], contracts: [], payments: [], mileageRecords: [], maintenanceTypes: DEFAULT_MAINT_TYPES, maintenanceRecords: [], revenues: [], expenses: [], incidents: [], users: s.users, readNotifications: [], settings: s.settings };
}

/* =========================================================================
   UI PRIMITIVES
========================================================================= */
const STATUS_STYLES = {
  "Disponível": "bg-emerald-50 text-emerald-700 border-emerald-200",
  "Alugada": "bg-blue-50 text-blue-700 border-blue-200",
  "Em manutenção": "bg-amber-50 text-amber-700 border-amber-200",
  "Reservada": "bg-violet-50 text-violet-700 border-violet-200",
  "Inativa": "bg-slate-100 text-slate-600 border-slate-200",
  "Vendida": "bg-slate-100 text-slate-500 border-slate-200",
  "Ativo": "bg-emerald-50 text-emerald-700 border-emerald-200",
  "Finalizado": "bg-slate-100 text-slate-600 border-slate-200",
  "Suspenso": "bg-amber-50 text-amber-700 border-amber-200",
  "Cancelado": "bg-red-50 text-red-700 border-red-200",
  "Bloqueado": "bg-red-50 text-red-700 border-red-200",
  "Inativo": "bg-slate-100 text-slate-600 border-slate-200",
  "Pago": "bg-emerald-50 text-emerald-700 border-emerald-200",
  "Pendente": "bg-slate-100 text-slate-600 border-slate-200",
  "Atrasado": "bg-red-50 text-red-700 border-red-200",
  "Parcial": "bg-amber-50 text-amber-700 border-amber-200",
  "Em dia": "bg-emerald-50 text-emerald-700 border-emerald-200",
  "Próxima": "bg-amber-50 text-amber-700 border-amber-200",
  "Sem histórico": "bg-slate-100 text-slate-500 border-slate-200",
  "Aberto": "bg-amber-50 text-amber-700 border-amber-200",
  "Em negociação": "bg-blue-50 text-blue-700 border-blue-200",
  "Resolvido": "bg-emerald-50 text-emerald-700 border-emerald-200",
};
function Badge({ children }) {
  const cls = STATUS_STYLES[children] || "bg-slate-100 text-slate-600 border-slate-200";
  return <span className={`inline-flex items-center gap-1 rounded-full border px-2.5 py-0.5 text-xs font-medium ${cls}`}>{children}</span>;
}

function KpiCard({ label, value, sub, icon: Icon, tone = "slate", onClick }) {
  const tones = {
    slate: "bg-white text-slate-900", orange: "bg-orange-600 text-white", emerald: "bg-white text-emerald-700",
    red: "bg-white text-red-600", blue: "bg-white text-blue-700",
  };
  return (
    <button onClick={onClick} className={`text-left rounded-xl border border-slate-200 p-4 shadow-sm hover:shadow-md transition-shadow ${tones[tone] || tones.slate} ${onClick ? "cursor-pointer" : "cursor-default"}`}>
      <div className="flex items-center justify-between">
        <span className={`text-xs font-medium ${tone === "orange" ? "text-orange-100" : "text-slate-500"}`}>{label}</span>
        {Icon && <Icon size={16} className={tone === "orange" ? "text-orange-100" : "text-slate-400"} />}
      </div>
      <div className="mt-2 text-2xl font-semibold tracking-tight" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{value}</div>
      {sub && <div className={`mt-1 text-xs ${tone === "orange" ? "text-orange-100" : "text-slate-500"}`}>{sub}</div>}
    </button>
  );
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 sm:p-8" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div
        className={`w-full ${wide ? "max-w-4xl" : "max-w-lg"} rounded-xl bg-white shadow-xl`}
        style={{ maxHeight: "85vh", overflowY: "auto" }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="sticky top-0 z-10 flex items-center justify-between border-b border-slate-200 bg-white px-5 py-4">
          <h3 className="text-base font-semibold text-slate-900" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{title}</h3>
          <button type="button" onClick={onClose} className="rounded-full p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"><X size={18} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children, span }) {
  return (
    <div className={span ? "sm:col-span-2" : ""}>
      <label className="mb-1 block text-xs font-medium text-slate-600">{label}</label>
      {children}
    </div>
  );
}
const inputCls = "w-full rounded-lg border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:border-orange-500 focus:outline-none focus:ring-1 focus:ring-orange-500";

function FormModal({ title, fields, initial, onCancel, onSave, wide }) {
  const [values, setValues] = useState(() => {
    const v = {};
    fields.forEach((f) => (v[f.name] = initial?.[f.name] ?? f.default ?? ""));
    return v;
  });
  const set = (name, val) => setValues((v) => ({ ...v, [name]: val }));
  const submit = () => {
    for (const f of fields) {
      if (f.required && !String(values[f.name] ?? "").trim()) { alert(`Preencha o campo "${f.label}".`); return; }
    }
    onSave(values);
  };
  return (
    <Modal title={title} onClose={onCancel} wide={wide}>
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        {fields.map((f) => (
          <Field key={f.name} label={f.required ? `${f.label} *` : f.label} span={f.span}>
            {f.type === "select" ? (
              <select className={inputCls} value={values[f.name]} onChange={(e) => set(f.name, e.target.value)}>
                <option value="">Selecione...</option>
                {(typeof f.options === "function" ? f.options(values) : f.options).map((o) => (
                  <option key={o.value ?? o} value={o.value ?? o}>{o.label ?? o}</option>
                ))}
              </select>
            ) : f.type === "textarea" ? (
              <textarea className={inputCls} rows={3} value={values[f.name]} onChange={(e) => set(f.name, e.target.value)} />
            ) : (
              <input
                className={inputCls}
                type={f.type || "text"}
                step={f.type === "number" ? "0.01" : undefined}
                value={values[f.name]}
                onChange={(e) => set(f.name, e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter" && f.type !== "textarea") submit(); }}
              />
            )}
            {f.hint && <p className="mt-1 text-xs text-slate-400">{f.hint}</p>}
          </Field>
        ))}
        <div className="sm:col-span-2 mt-2 flex justify-end gap-2 border-t border-slate-100 pt-4">
          <button type="button" onClick={onCancel} className="rounded-lg border border-slate-300 px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cancelar</button>
          <button type="button" onClick={submit} className="rounded-lg bg-orange-600 px-4 py-2 text-sm font-medium text-white hover:bg-orange-700 active:bg-orange-800">Salvar</button>
        </div>
      </div>
    </Modal>
  );
}

function ConfirmDialog({ text, onCancel, onConfirm }) {
  return (
    <Modal title="Confirmar exclusão" onClose={onCancel}>
      <p className="text-sm text-slate-600">{text}</p>
      <div className="mt-5 flex justify-end gap-2">
        <button onClick={onCancel} className="rounded-lg border border-slate-300 px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cancelar</button>
        <button onClick={onConfirm} className="rounded-lg bg-red-600 px-4 py-2 text-sm font-medium text-white hover:bg-red-700">Excluir</button>
      </div>
    </Modal>
  );
}

function Toolbar({ search, setSearch, placeholder, onNew, newLabel, right }) {
  return (
    <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
      <div className="relative w-full sm:max-w-xs">
        <Search size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder={placeholder || "Buscar..."} className={`${inputCls} pl-8`} />
      </div>
      <div className="flex flex-wrap items-center gap-2">
        {right}
        {onNew && (
          <button onClick={onNew} className="flex items-center gap-1.5 rounded-lg bg-orange-600 px-3.5 py-2 text-sm font-medium text-white hover:bg-orange-700">
            <Plus size={15} /> {newLabel || "Novo"}
          </button>
        )}
      </div>
    </div>
  );
}

function Th({ children }) { return <th className="whitespace-nowrap px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-slate-500">{children}</th>; }
function Td({ children, mono }) { return <td className={`whitespace-nowrap px-3 py-2.5 text-sm text-slate-700 ${mono ? "font-mono" : ""}`}>{children}</td>; }

function DataTable({ columns, rows, onEdit, onDelete, onRowClick, emptyText, pageSize = 8 }) {
  const [page, setPage] = useState(0);
  useEffect(() => setPage(0), [rows.length]);
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  const pageRows = rows.slice(page * pageSize, page * pageSize + pageSize);
  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
      <div className="overflow-x-auto">
        <table className="w-full border-collapse">
          <thead className="border-b border-slate-200 bg-slate-50">
            <tr>
              {columns.map((c) => <Th key={c.key}>{c.label}</Th>)}
              {(onEdit || onDelete) && <Th>Ações</Th>}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {pageRows.length === 0 && (
              <tr><td colSpan={columns.length + 1} className="px-3 py-8 text-center text-sm text-slate-400">{emptyText || "Nenhum registro encontrado."}</td></tr>
            )}
            {pageRows.map((r, i) => (
              <tr key={r.id || i} className={onRowClick ? "cursor-pointer hover:bg-slate-50" : "hover:bg-slate-50/50"} onClick={() => onRowClick && onRowClick(r)}>
                {columns.map((c) => <Td key={c.key} mono={c.mono}>{c.render ? c.render(r) : r[c.key]}</Td>)}
                {(onEdit || onDelete) && (
                  <Td>
                    <div className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                      {onEdit && <button onClick={() => onEdit(r)} className="rounded p-1.5 text-slate-500 hover:bg-slate-100 hover:text-orange-600"><Pencil size={14} /></button>}
                      {onDelete && <button onClick={() => onDelete(r)} className="rounded p-1.5 text-slate-500 hover:bg-red-50 hover:text-red-600"><Trash2 size={14} /></button>}
                    </div>
                  </Td>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {rows.length > pageSize && (
        <div className="flex items-center justify-between border-t border-slate-200 px-3 py-2 text-xs text-slate-500">
          <span>{rows.length} registro(s) · página {page + 1} de {totalPages}</span>
          <div className="flex gap-1">
            <button disabled={page === 0} onClick={() => setPage((p) => p - 1)} className="rounded p-1 disabled:opacity-30 hover:bg-slate-100"><ChevronLeft size={16} /></button>
            <button disabled={page >= totalPages - 1} onClick={() => setPage((p) => p + 1)} className="rounded p-1 disabled:opacity-30 hover:bg-slate-100"><ChevronRight size={16} /></button>
          </div>
        </div>
      )}
    </div>
  );
}

function SectionHeader({ title, desc, actions }) {
  return (
    <div className="mb-5 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
      <div>
        <h2 className="text-xl font-semibold text-slate-900" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{title}</h2>
        {desc && <p className="text-sm text-slate-500">{desc}</p>}
      </div>
      {actions}
    </div>
  );
}

function MiniBar({ pct, tone = "orange" }) {
  const tones = { orange: "bg-orange-500", emerald: "bg-emerald-500", red: "bg-red-500", amber: "bg-amber-500", blue: "bg-blue-500" };
  return (
    <div className="h-1.5 w-full overflow-hidden rounded-full bg-slate-100">
      <div className={`h-full rounded-full ${tones[tone]}`} style={{ width: `${Math.min(100, Math.max(0, pct))}%` }} />
    </div>
  );
}

/* Simple SVG line/bar charts (no extra deps needed beyond what's already used) */
function SvgLineChart({ points, height = 160, color = "#ea580c", formatY }) {
  if (!points.length) return <div className="flex h-40 items-center justify-center text-sm text-slate-400">Sem dados suficientes.</div>;
  const w = 600, h = height, pad = 30;
  const xs = points.map((_, i) => i);
  const ys = points.map((p) => p.value);
  const minY = Math.min(...ys, 0), maxY = Math.max(...ys, 1);
  const sx = (i) => pad + (i * (w - 2 * pad)) / Math.max(1, points.length - 1);
  const sy = (v) => h - pad - ((v - minY) * (h - 2 * pad)) / Math.max(1, maxY - minY);
  const path = points.map((p, i) => `${i === 0 ? "M" : "L"} ${sx(i)} ${sy(p.value)}`).join(" ");
  return (
    <svg viewBox={`0 0 ${w} ${h}`} className="w-full" style={{ height }}>
      <line x1={pad} y1={h - pad} x2={w - pad} y2={h - pad} stroke="#e2e8f0" />
      <path d={path} fill="none" stroke={color} strokeWidth="2.5" />
      {points.map((p, i) => <circle key={i} cx={sx(i)} cy={sy(p.value)} r="3" fill={color} />)}
      {points.map((p, i) => (i % Math.ceil(points.length / 6 || 1) === 0) && (
        <text key={"t" + i} x={sx(i)} y={h - pad + 14} fontSize="9" fill="#94a3b8" textAnchor="middle">{p.label}</text>
      ))}
    </svg>
  );
}
function SvgBarChart({ bars, height = 200 }) {
  if (!bars.length) return <div className="flex h-40 items-center justify-center text-sm text-slate-400">Sem dados suficientes.</div>;
  const w = 600, h = height, pad = 30;
  const max = Math.max(...bars.flatMap((b) => [b.a, b.b]), 1);
  const bw = (w - 2 * pad) / bars.length;
  return (
    <svg viewBox={`0 0 ${w} ${h}`} className="w-full" style={{ height }}>
      <line x1={pad} y1={h - pad} x2={w - pad} y2={h - pad} stroke="#e2e8f0" />
      {bars.map((b, i) => {
        const x = pad + i * bw;
        const ah = ((b.a || 0) / max) * (h - 2 * pad);
        const bh = ((b.b || 0) / max) * (h - 2 * pad);
        return (
          <g key={i}>
            <rect x={x + bw * 0.15} y={h - pad - ah} width={bw * 0.3} height={ah} fill="#059669" rx="2" />
            <rect x={x + bw * 0.55} y={h - pad - bh} width={bw * 0.3} height={bh} fill="#dc2626" rx="2" />
            <text x={x + bw / 2} y={h - pad + 14} fontSize="9" fill="#94a3b8" textAnchor="middle">{b.label}</text>
          </g>
        );
      })}
    </svg>
  );
}

/* =========================================================================
   APP
========================================================================= */
const EMPTY_COLLECTIONS = { motorcycles: [], renters: [], contracts: [], payments: [], mileageRecords: [], maintenanceTypes: [], maintenanceRecords: [], revenues: [], expenses: [], incidents: [], users: [] };

export default function App() {
  const [db, setDb] = useState(null);
  const [loading, setLoading] = useState(true);
  const [mode, setMode] = useState("local"); // "api" (backend real) | "local" (sem servidor)
  const [currentUser, setCurrentUser] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [notifOpen, setNotifOpen] = useState(false);
  const [focusMotoId, setFocusMotoId] = useState(null);
  const [focusRenterId, setFocusRenterId] = useState(null);

  const loadAllCollections = useCallback(async () => {
    const keys = Object.keys(SERVER_RESOURCES);
    const results = await Promise.all(keys.map((k) => api.list(SERVER_RESOURCES[k]).catch(() => [])));
    const out = {};
    keys.forEach((k, i) => { out[k] = k === "users" ? results[i].map(fromApiUser) : results[i]; });
    return out;
  }, []);

  useEffect(() => {
    (async () => {
      const healthy = await apiHealthCheck();
      if (healthy) {
        setMode("api");
        let shell = null;
        try { const r = await window.storage.get(SHELL_STORAGE_KEY, false); shell = r?.value ? JSON.parse(r.value) : null; } catch { /* sem shell salvo ainda */ }
        const settings = shell?.settings || buildSeed().settings;
        const readNotifications = shell?.readNotifications || [];
        try {
          const me = await api.me();
          const collections = await loadAllCollections();
          setCurrentUser(me);
          setDb({ ...collections, settings, readNotifications });
        } catch {
          setDb({ ...EMPTY_COLLECTIONS, settings, readNotifications });
        }
      } else {
        setMode("local");
        try {
          const res = await window.storage.get(STORAGE_KEY, false);
          const loaded = res?.value ? JSON.parse(res.value) : buildSeed();
          if (!loaded.users || !loaded.users.length) loaded.users = buildSeed().users;
          if (!loaded.readNotifications) loaded.readNotifications = [];
          setDb(loaded);
        } catch { setDb(buildSeed()); }
      }
      setLoading(false);
    })();
  }, [loadAllCollections]);

  // Persistência: Modo local salva o banco inteiro; Modo conectado só salva as preferências
  // locais (configurações e notificações lidas), pois o resto já vive no servidor.
  useEffect(() => {
    if (!db || mode !== "local") return;
    const t = setTimeout(() => window.storage.set(STORAGE_KEY, JSON.stringify(db), false).catch(() => {}), 300);
    return () => clearTimeout(t);
  }, [db, mode]);
  useEffect(() => {
    if (!db || mode !== "api") return;
    const t = setTimeout(() => window.storage.set(SHELL_STORAGE_KEY, JSON.stringify({ settings: db.settings, readNotifications: db.readNotifications }), false).catch(() => {}), 300);
    return () => clearTimeout(t);
  }, [db, mode]);

  // Envia ao backend real as diferenças (criações/edições/exclusões) de uma coleção,
  // sem exigir que cada tela do sistema seja reescrita para chamar a API diretamente.
  const syncCollection = useCallback((key, oldArr, newArr) => {
    const endpoint = SERVER_RESOURCES[key];
    const oldById = new Map((oldArr || []).map((x) => [x.id, x]));
    const newById = new Map((newArr || []).map((x) => [x.id, x]));
    for (const id of oldById.keys()) {
      if (!newById.has(id)) api.remove(endpoint, id).catch((e) => console.error(`Falha ao excluir em ${key}:`, e.message));
    }
    for (const [id, item] of newById) {
      const payload = key === "users" ? toApiUser(item) : sanitizeForApi(key, item);
      if (!oldById.has(id)) {
        api.create(endpoint, payload)
          .then((created) => {
            const fixed = key === "users" ? fromApiUser(created) : created;
            setDb((d) => ({ ...d, [key]: (d[key] || []).map((x) => (x.id === id ? { ...x, ...fixed } : x)) }));
          })
          .catch((e) => { console.error(`Falha ao criar em ${key}:`, e.message); alert(`Não foi possível salvar (${key}): ${e.message}`); });
      } else if (JSON.stringify(oldById.get(id)) !== JSON.stringify(item)) {
        api.update(endpoint, id, payload).catch((e) => { console.error(`Falha ao atualizar em ${key}:`, e.message); alert(`Não foi possível atualizar (${key}): ${e.message}`); });
      }
    }
  }, []);

  const patch = useCallback((partialOrFn) => {
    setDb((d) => {
      const partial = typeof partialOrFn === "function" ? partialOrFn(d) : partialOrFn;
      if (mode === "api") {
        Object.keys(partial).forEach((key) => { if (SERVER_RESOURCES[key]) syncCollection(key, d[key], partial[key]); });
      }
      return { ...d, ...partial };
    });
  }, [mode, syncCollection]);

  const handleLoginApi = async (user) => {
    setCurrentUser(user);
    const collections = await loadAllCollections();
    setDb((d) => ({ ...d, ...collections }));
  };
  const handleLogout = async () => {
    if (mode === "api") {
      await api.logout();
      setDb((d) => ({ ...d, ...EMPTY_COLLECTIONS }));
    }
    setCurrentUser(null);
  };

  if (loading || !db) {
    return <div className="flex h-screen items-center justify-center bg-slate-50 text-slate-400">Carregando sistema...</div>;
  }
  if (!currentUser) return <LoginScreen db={db} mode={mode} onLoginLocal={setCurrentUser} onLoginApi={handleLoginApi} />;

  return (
    <AppShell
      db={db} patch={patch} tab={tab} setTab={setTab} mode={mode}
      mobileNavOpen={mobileNavOpen} setMobileNavOpen={setMobileNavOpen}
      notifOpen={notifOpen} setNotifOpen={setNotifOpen}
      currentUser={currentUser}
      onLogout={handleLogout}
      focusMotoId={focusMotoId} setFocusMotoId={setFocusMotoId}
      focusRenterId={focusRenterId} setFocusRenterId={setFocusRenterId}
    />
  );
}

const USER_ROLES = ["Administrador", "Gerente", "Funcionário"];

function LoginScreen({ db, mode, onLoginLocal, onLoginApi }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const doLogin = async () => {
    setErr("");
    if (mode === "api") {
      setBusy(true);
      try {
        const user = await api.login(username.trim(), password);
        await onLoginApi(user);
      } catch (e) {
        setErr(e.message || "Usuário ou senha inválidos.");
      } finally {
        setBusy(false);
      }
    } else {
      const u = db.users.find((x) => x.username.trim().toLowerCase() === username.trim().toLowerCase() && x.password === password);
      if (u) onLoginLocal(u);
      else setErr("Usuário ou senha inválidos.");
    }
  };
  const onKeyDownLogin = (e) => { if (e.key === "Enter") doLogin(); };

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-900 p-4">
      <div className="w-full max-w-sm rounded-2xl bg-white p-8 shadow-2xl">
        <div className="mb-6 flex items-center gap-2.5">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-orange-600 text-white"><Bike size={20} /></div>
          <div>
            <p className="text-sm font-semibold text-slate-900" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{db.settings?.companyName || "Frota Fácil Motos"}</p>
            <p className="text-xs text-slate-400">Gestão de frota e locações</p>
          </div>
        </div>
        <div className={`mb-4 flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium ${mode === "api" ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>
          <span className={`h-1.5 w-1.5 rounded-full ${mode === "api" ? "bg-emerald-500" : "bg-amber-500"}`} />
          {mode === "api" ? "Conectado ao servidor" : "Modo local — dados salvos apenas neste navegador"}
        </div>
        <h1 className="mb-1 text-lg font-semibold text-slate-900">Acessar o sistema</h1>
        <p className="mb-5 text-sm text-slate-500">Entre com seu usuário e senha.</p>
        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-600">Usuário</label>
            <input className={inputCls} value={username} onChange={(e) => setUsername(e.target.value)} onKeyDown={onKeyDownLogin} autoComplete="username" />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-600">Senha</label>
            <input className={inputCls} type="password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={onKeyDownLogin} autoComplete="current-password" />
          </div>
          {err && <p className="text-xs text-red-600">{err}</p>}
          <button type="button" disabled={busy} onClick={doLogin} className="flex w-full items-center justify-center gap-2 rounded-lg bg-orange-600 px-4 py-2.5 text-sm font-medium text-white hover:bg-orange-700 active:bg-orange-800 disabled:opacity-60">
            <KeyRound size={15} /> {busy ? "Entrando..." : "Entrar"}
          </button>
          {mode === "local" && <p className="text-center text-xs text-slate-400">Usuário de demonstração: <span className="font-mono">admin</span> / <span className="font-mono">admin123</span></p>}
          <p className="text-center text-xs text-slate-400">Novos usuários só podem ser criados por um administrador já logado, em Configurações.</p>
        </div>
      </div>
    </div>
  );
}

const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "frota", label: "Frota", icon: Bike },
  { key: "locatarios", label: "Locatários", icon: Users },
  { key: "contratos", label: "Contratos", icon: FileText },
  { key: "pagamentos", label: "Pagamentos", icon: Wallet },
  { key: "receitas", label: "Receitas", icon: TrendingUp },
  { key: "despesas", label: "Despesas", icon: TrendingDown },
  { key: "manutencao", label: "Manutenções", icon: Wrench },
  { key: "km", label: "Quilometragem", icon: Gauge },
  { key: "ocorrencias", label: "Ocorrências", icon: AlertTriangle },
  { key: "calendario", label: "Calendário", icon: CalendarIcon },
  { key: "relatorios", label: "Relatórios", icon: BarChart3 },
  { key: "config", label: "Configurações", icon: Settings },
];

function AppShell(props) {
  const { db, patch, tab, setTab, mobileNavOpen, setMobileNavOpen, notifOpen, setNotifOpen, onLogout, focusMotoId, setFocusMotoId, focusRenterId, setFocusRenterId, currentUser, mode } = props;
  const notifications = useMemo(() => buildNotifications(db), [db]);
  const readKeys = db.readNotifications || [];
  const unreadCount = notifications.filter((n) => !readKeys.includes(n.key)).length;

  const goMoto = (id) => { setFocusMotoId(id); setTab("frota"); };
  const goRenter = (id) => { setFocusRenterId(id); setTab("locatarios"); };

  const toggleNotif = () => {
    const opening = !notifOpen;
    setNotifOpen(opening);
    if (opening && notifications.length) {
      patch((d) => ({ readNotifications: Array.from(new Set([...(d.readNotifications || []), ...notifications.map((n) => n.key)])) }));
    }
  };

  const openNotification = (n) => {
    setNotifOpen(false);
    if (n.type === "moto") goMoto(n.entityId);
    else if (n.type === "renter") goRenter(n.entityId);
    else if (n.type === "contract") setTab("contratos");
    else if (n.type === "payment") setTab("pagamentos");
  };

  return (
    <div className="min-h-screen bg-slate-50" style={{ fontFamily: "'Inter', sans-serif" }}>
      <FontLoader />
      {/* Sidebar desktop */}
      <aside className="fixed inset-y-0 left-0 z-30 hidden w-60 flex-col bg-slate-900 lg:flex">
        <div className="flex items-center gap-2.5 px-5 py-5">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-orange-600 text-white"><Bike size={18} /></div>
          <div className="min-w-0">
            <p className="truncate text-sm font-semibold text-white" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{db.settings.companyName}</p>
            <p className="text-[11px] text-slate-400">Gestão de frota</p>
          </div>
        </div>
        <nav className="flex-1 space-y-0.5 overflow-y-auto px-3 py-2">
          {NAV.map((n) => (
            <button key={n.key} onClick={() => setTab(n.key)} className={`flex w-full items-center gap-2.5 rounded-lg px-3 py-2 text-sm font-medium transition-colors ${tab === n.key ? "bg-orange-600 text-white" : "text-slate-300 hover:bg-slate-800 hover:text-white"}`}>
              <n.icon size={16} /> {n.label}
            </button>
          ))}
        </nav>
        <div className="border-t border-slate-800 p-3">
          <button onClick={onLogout} className="flex w-full items-center gap-2.5 rounded-lg px-3 py-2 text-sm font-medium text-slate-300 hover:bg-slate-800 hover:text-white">
            <LogOut size={16} /> Sair
          </button>
        </div>
      </aside>

      {/* Mobile nav drawer */}
      {mobileNavOpen && (
        <div className="fixed inset-0 z-40 lg:hidden">
          <div className="absolute inset-0 bg-slate-900/50" onClick={() => setMobileNavOpen(false)} />
          <div className="absolute inset-y-0 left-0 w-64 bg-slate-900 p-3">
            <div className="mb-3 flex items-center justify-between px-2 py-2">
              <span className="text-sm font-semibold text-white">{db.settings.companyName}</span>
              <button onClick={() => setMobileNavOpen(false)}><X size={18} className="text-slate-300" /></button>
            </div>
            {NAV.map((n) => (
              <button key={n.key} onClick={() => { setTab(n.key); setMobileNavOpen(false); }} className={`flex w-full items-center gap-2.5 rounded-lg px-3 py-2.5 text-sm font-medium ${tab === n.key ? "bg-orange-600 text-white" : "text-slate-300 hover:bg-slate-800 hover:text-white"}`}>
                <n.icon size={16} /> {n.label}
              </button>
            ))}
            <button onClick={onLogout} className="mt-2 flex w-full items-center gap-2.5 rounded-lg px-3 py-2.5 text-sm font-medium text-slate-300 hover:bg-slate-800 hover:text-white">
              <LogOut size={16} /> Sair
            </button>
          </div>
        </div>
      )}

      <div className="lg:pl-60">
        {/* Topbar */}
        <header className="sticky top-0 z-20 flex items-center justify-between border-b border-slate-200 bg-white/90 px-4 py-3 backdrop-blur sm:px-6">
          <div className="flex items-center gap-3">
            <button className="rounded-lg p-1.5 hover:bg-slate-100 lg:hidden" onClick={() => setMobileNavOpen(true)}><LayoutDashboard size={18} className="text-slate-600" /></button>
            <h1 className="text-sm font-semibold text-slate-500">{NAV.find((n) => n.key === tab)?.label}</h1>
            <span className={`hidden items-center gap-1 rounded-full px-2 py-0.5 text-[10px] font-medium sm:flex ${mode === "api" ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>
              <span className={`h-1.5 w-1.5 rounded-full ${mode === "api" ? "bg-emerald-500" : "bg-amber-500"}`} /> {mode === "api" ? "Servidor" : "Local"}
            </span>
          </div>
          <div className="flex items-center gap-2">
            <div className="relative">
              <button onClick={toggleNotif} className="relative rounded-lg p-2 hover:bg-slate-100">
                <Bell size={18} className="text-slate-600" />
                {unreadCount > 0 && <span className="absolute -right-0.5 -top-0.5 flex h-4 w-4 items-center justify-center rounded-full bg-red-600 text-[10px] font-bold text-white">{unreadCount > 9 ? "9+" : unreadCount}</span>}
              </button>
              {notifOpen && (
                <div className="absolute right-0 z-30 mt-2 w-80 rounded-xl border border-slate-200 bg-white p-2 shadow-xl">
                  <div className="flex items-center justify-between px-2 py-1">
                    <span className="text-xs font-semibold text-slate-500">NOTIFICAÇÕES ({notifications.length})</span>
                  </div>
                  <div className="max-h-80 overflow-y-auto">
                    {notifications.length === 0 && <p className="px-2 py-4 text-center text-xs text-slate-400">Nenhum alerta no momento. Tudo em dia!</p>}
                    {notifications.map((n, i) => (
                      <button key={n.key || i} onClick={() => openNotification(n)} className="flex w-full items-start gap-2 rounded-lg px-2 py-2 text-left hover:bg-slate-50">
                        <AlertIcon level={n.level} />
                        <div className="min-w-0">
                          <p className="text-xs font-medium text-slate-800">{n.title}</p>
                          <p className="text-[11px] text-slate-500">{n.detail}</p>
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
            <div className="hidden items-center gap-2 rounded-full bg-slate-100 px-3 py-1.5 sm:flex">
              <div className="flex h-6 w-6 items-center justify-center rounded-full bg-orange-600 text-[11px] font-semibold text-white">{(currentUser?.name || "?").charAt(0).toUpperCase()}</div>
              <span className="text-xs font-medium text-slate-700">{currentUser?.name}</span>
              <span className="text-[11px] text-slate-400">· {currentUser?.role}</span>
            </div>
          </div>
        </header>

        <main className="p-4 sm:p-6">
          {tab === "dashboard" && <Dashboard db={db} goMoto={goMoto} goRenter={goRenter} setTab={setTab} notifications={notifications} />}
          {tab === "frota" && <FleetModule db={db} patch={patch} focusMotoId={focusMotoId} setFocusMotoId={setFocusMotoId} goRenter={goRenter} />}
          {tab === "locatarios" && <RentersModule db={db} patch={patch} focusRenterId={focusRenterId} setFocusRenterId={setFocusRenterId} goMoto={goMoto} />}
          {tab === "contratos" && <ContractsModule db={db} patch={patch} />}
          {tab === "pagamentos" && <PaymentsModule db={db} patch={patch} />}
          {tab === "receitas" && <RevenuesModule db={db} patch={patch} />}
          {tab === "despesas" && <ExpensesModule db={db} patch={patch} />}
          {tab === "manutencao" && <MaintenanceModule db={db} patch={patch} />}
          {tab === "km" && <MileageModule db={db} patch={patch} />}
          {tab === "ocorrencias" && <IncidentsModule db={db} patch={patch} />}
          {tab === "calendario" && <CalendarModule db={db} />}
          {tab === "relatorios" && <ReportsModule db={db} />}
          {tab === "config" && <SettingsModule db={db} patch={patch} currentUser={currentUser} mode={mode} />}
        </main>
      </div>
    </div>
  );
}

function FontLoader() {
  useEffect(() => {
    if (document.getElementById("frota-fonts")) return;
    const l = document.createElement("link");
    l.id = "frota-fonts"; l.rel = "stylesheet";
    l.href = "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap";
    document.head.appendChild(l);
  }, []);
  return null;
}

function AlertIcon({ level }) {
  if (level === "red") return <XCircle size={15} className="mt-0.5 shrink-0 text-red-500" />;
  if (level === "amber") return <AlertCircle size={15} className="mt-0.5 shrink-0 text-amber-500" />;
  return <CheckCircle2 size={15} className="mt-0.5 shrink-0 text-blue-500" />;
}

/* =========================================================================
   DERIVED / BUSINESS LOGIC
========================================================================= */
function maintenanceAlertsForMoto(db, motoId) {
  const moto = db.motorcycles.find((m) => m.id === motoId);
  if (!moto) return [];
  const warnKm = Number(db.settings.alertKmWarning || 500);
  return db.maintenanceTypes.map((t) => {
    const records = db.maintenanceRecords.filter((r) => r.motorcycleId === motoId && r.typeId === t.id).sort((a, b) => (a.date < b.date ? 1 : -1));
    const last = records[0];
    const baseKm = last ? last.km : moto.acquisitionKm;
    const nextKm = baseKm + t.kmInterval;
    const kmRemaining = nextKm - moto.currentKm;
    let status = "Sem histórico";
    if (last) {
      if (kmRemaining <= 0) status = "Atrasado";
      else if (kmRemaining <= warnKm) status = "Próxima";
      else status = "Em dia";
    }
    return { type: t, last, nextKm, kmRemaining, status };
  });
}

function buildNotifications(db) {
  const list = [];
  const today = todayISO();
  const warnDays = Number(db.settings.alertDaysWarning || 15);
  const warnKm = Number(db.settings.alertKmWarning || 500);

  db.motorcycles.forEach((m) => {
    if (m.status === "Vendida" || m.status === "Inativa") return;
    maintenanceAlertsForMoto(db, m.id).forEach((a) => {
      if (a.status === "Atrasado") list.push({ key: `maint-${m.id}-${a.type.id}`, type: "moto", entityId: m.id, level: "red", title: `${a.type.name} atrasada — ${m.brand} ${m.model} (${m.plate})`, detail: `Vencida há ${fmtNum(Math.abs(a.kmRemaining))} km. Toque para abrir a moto.` });
      else if (a.status === "Próxima") list.push({ key: `maint-${m.id}-${a.type.id}`, type: "moto", entityId: m.id, level: "amber", title: `${a.type.name} próxima — ${m.brand} ${m.model} (${m.plate})`, detail: `Faltam ${fmtNum(a.kmRemaining)} km. Toque para abrir a moto.` });
    });
    if (m.insuranceDue) {
      const d = daysDiff(m.insuranceDue, today);
      if (d <= warnDays) list.push({ key: `insurance-${m.id}-${m.insuranceDue}`, type: "moto", entityId: m.id, level: d < 0 ? "red" : "amber", title: `Seguro ${d < 0 ? "vencido" : "vence"} — ${m.brand} ${m.model} (${m.plate})`, detail: (d < 0 ? `Venceu em ${fmtDate(m.insuranceDue)}.` : `Vence em ${d} dia(s) — ${fmtDate(m.insuranceDue)}.`) + " Toque para abrir a moto." });
    }
    if (m.nextLicensingDue) {
      const d = daysDiff(m.nextLicensingDue, today);
      if (d <= warnDays) list.push({ key: `licensing-${m.id}-${m.nextLicensingDue}`, type: "moto", entityId: m.id, level: d < 0 ? "red" : "amber", title: `Licenciamento ${d < 0 ? "vencido" : "vence"} — ${m.brand} ${m.model} (${m.plate})`, detail: (d < 0 ? `Venceu em ${fmtDate(m.nextLicensingDue)}.` : `Vence em ${d} dia(s) — ${fmtDate(m.nextLicensingDue)}.`) + " Toque para abrir a moto." });
    }
  });

  db.contracts.forEach((c) => {
    if (c.status !== "Ativo") return;
    const d = daysDiff(c.endDate, today);
    if (d <= 7) {
      const moto = db.motorcycles.find((m) => m.id === c.motorcycleId);
      const renter = db.renters.find((r) => r.id === c.renterId);
      list.push({ key: `contract-${c.id}-${c.endDate}`, type: "contract", entityId: c.id, level: d < 0 ? "red" : "amber", title: `Contrato ${d < 0 ? "vencido" : "vencendo"} — ${renter?.name || "?"}`, detail: `${moto ? moto.plate : ""} · ${d < 0 ? "Venceu em" : "Vence em"} ${fmtDate(c.endDate)}. Toque para abrir os contratos.` });
    }
  });

  db.renters.forEach((r) => {
    if (!r.cnhValidity) return;
    const d = daysDiff(r.cnhValidity, today);
    if (d <= warnDays) list.push({ key: `cnh-${r.id}-${r.cnhValidity}`, type: "renter", entityId: r.id, level: d < 0 ? "red" : "amber", title: `CNH ${d < 0 ? "vencida" : "vencendo"} — ${r.name}`, detail: (d < 0 ? `Venceu em ${fmtDate(r.cnhValidity)}.` : `Vence em ${d} dia(s) — ${fmtDate(r.cnhValidity)}.`) + " Toque para abrir o locatário." });
  });

  db.payments.forEach((p) => {
    const st = paymentStatus(p, today);
    if (st === "Atrasado") {
      const renter = db.renters.find((r) => r.id === p.renterId);
      list.push({ key: `payment-${p.id}`, type: "payment", entityId: p.id, level: "red", title: `Pagamento atrasado — ${renter?.name || "?"}`, detail: `${fmtBRL(p.expectedValue)} venceu em ${fmtDate(p.dueDate)}. Toque para abrir os pagamentos.` });
    }
  });

  return list;
}

function paymentStatus(p, today) {
  if (p.status === "Cancelado") return "Cancelado";
  if (p.paidValue >= p.expectedValue && p.paidValue > 0) return "Pago";
  if (p.paidValue > 0 && p.paidValue < p.expectedValue) return "Parcial";
  if (p.dueDate < today) return "Atrasado";
  return "Pendente";
}

function motoLabel(m) { return m ? `${m.brand} ${m.model} — ${m.plate}` : "—"; }

function computeMotoFinancials(db, motoId, startISO, endISO) {
  const inP = (d) => (!startISO || inRange(d, startISO, endISO));
  const rentPaid = db.payments.filter((p) => p.motorcycleId === motoId && p.paidValue > 0 && inP(p.paymentDate || p.dueDate)).reduce((s, p) => s + Number(p.paidValue), 0);
  const rev = db.revenues.filter((r) => r.motorcycleId === motoId && inP(r.date)).reduce((s, r) => s + Number(r.value), 0);
  const exp = db.expenses.filter((e) => e.motorcycleId === motoId && inP(e.date)).reduce((s, e) => s + Number(e.value), 0);
  const maint = db.maintenanceRecords.filter((r) => r.motorcycleId === motoId && inP(r.date)).reduce((s, r) => s + Number(r.totalCost), 0);
  const revenueTotal = rentPaid + rev;
  const costTotal = exp + maint;
  return { rentPaid, rev, exp, maint, revenueTotal, costTotal, profit: revenueTotal - costTotal };
}

function computeRenterStats(db, renterId) {
  const pays = db.payments.filter((p) => p.renterId === renterId);
  const contracts = db.contracts.filter((c) => c.renterId === renterId);
  const today = todayISO();
  const late = pays.filter((p) => paymentStatus(p, today) === "Atrasado").length;
  const revenue = pays.reduce((s, p) => s + Number(p.paidValue || 0), 0);
  const weeks = pays.length;
  let days = 0;
  contracts.forEach((c) => { days += daysDiff(c.status === "Ativo" ? todayISO() : c.endDate, c.startDate); });
  return { weeks, late, revenue, days, contracts: contracts.length };
}

/* =========================================================================
   DASHBOARD
========================================================================= */
function Dashboard({ db, goMoto, goRenter, setTab, notifications }) {
  const today = todayISO();
  const weekStart = startOfWeek(today);
  const weekEnd = addDays(weekStart, 6);
  const monthStart = startOfMonth(today);

  const disponiveis = db.motorcycles.filter((m) => m.status === "Disponível").length;
  const alugadas = db.motorcycles.filter((m) => m.status === "Alugada").length;
  const manutencao = db.motorcycles.filter((m) => m.status === "Em manutenção").length;
  const vencendo = db.contracts.filter((c) => c.status === "Ativo" && daysDiff(c.endDate, today) <= 7).length;
  const atrasados = db.payments.filter((p) => paymentStatus(p, today) === "Atrasado").length;

  const sumRevenue = (s, e) => db.payments.filter((p) => p.paidValue > 0 && inRange(p.paymentDate || p.dueDate, s, e)).reduce((a, p) => a + Number(p.paidValue), 0)
    + db.revenues.filter((r) => inRange(r.date, s, e)).reduce((a, r) => a + Number(r.value), 0);
  const sumExpense = (s, e) => db.expenses.filter((x) => inRange(x.date, s, e)).reduce((a, x) => a + Number(x.value), 0)
    + db.maintenanceRecords.filter((m) => inRange(m.date, s, e)).reduce((a, m) => a + Number(m.totalCost), 0);

  const receitaSemana = sumRevenue(weekStart, weekEnd);
  const receitaMes = sumRevenue(monthStart, today);
  const despesaSemana = sumExpense(weekStart, weekEnd);
  const despesaMes = sumExpense(monthStart, today);
  const lucroLiquido = receitaMes - despesaMes;
  const aReceber = db.payments.filter((p) => ["Pendente", "Atrasado", "Parcial"].includes(paymentStatus(p, today))).reduce((s, p) => s + (Number(p.expectedValue) - Number(p.paidValue || 0)), 0);
  const kmTotal = db.motorcycles.reduce((s, m) => s + Number(m.currentKm || 0), 0);

  let manutPendentes = 0;
  const proximasManut = [];
  db.motorcycles.forEach((m) => {
    maintenanceAlertsForMoto(db, m.id).forEach((a) => {
      if (a.status === "Atrasado" || a.status === "Próxima") { manutPendentes++; proximasManut.push({ moto: m, ...a }); }
    });
  });
  proximasManut.sort((a, b) => a.kmRemaining - b.kmRemaining);

  const profitability = db.motorcycles.map((m) => ({ moto: m, ...computeMotoFinancials(db, m.id, monthStart, today) })).sort((a, b) => b.profit - a.profit);

  return (
    <div className="space-y-6">
      <SectionHeader title="Visão geral" desc={`Hoje é ${fmtDate(today)} · atualizado em tempo real`} />

      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
        <KpiCard label="Total de motos" value={db.motorcycles.length} icon={Bike} onClick={() => setTab("frota")} />
        <KpiCard label="Disponíveis" value={disponiveis} icon={CheckCircle2} tone="emerald" onClick={() => setTab("frota")} />
        <KpiCard label="Alugadas" value={alugadas} icon={KeyRound} onClick={() => setTab("frota")} />
        <KpiCard label="Em manutenção" value={manutencao} icon={Wrench} onClick={() => setTab("manutencao")} />
        <KpiCard label="Contratos vencendo" value={vencendo} icon={FileText} onClick={() => setTab("contratos")} />
        <KpiCard label="Pagamentos atrasados" value={atrasados} icon={AlertTriangle} tone={atrasados ? "red" : "slate"} onClick={() => setTab("pagamentos")} />
      </div>

      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
        <KpiCard label="Receita da semana" value={fmtBRL(receitaSemana)} icon={TrendingUp} tone="orange" />
        <KpiCard label="Receita do mês" value={fmtBRL(receitaMes)} icon={TrendingUp} />
        <KpiCard label="Despesas da semana" value={fmtBRL(despesaSemana)} icon={TrendingDown} />
        <KpiCard label="Despesas do mês" value={fmtBRL(despesaMes)} icon={TrendingDown} />
        <KpiCard label="Lucro líquido (mês)" value={fmtBRL(lucroLiquido)} icon={Wallet} tone={lucroLiquido >= 0 ? "emerald" : "red"} />
        <KpiCard label="Valor a receber" value={fmtBRL(aReceber)} icon={Wallet} />
        <KpiCard label="KM total da frota" value={`${fmtNum(kmTotal)} km`} icon={Gauge} />
        <KpiCard label="Manutenções pendentes" value={manutPendentes} icon={Wrench} tone={manutPendentes ? "red" : "slate"} onClick={() => setTab("manutencao")} />
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm lg:col-span-2">
          <h3 className="mb-3 text-sm font-semibold text-slate-800">Receitas x Despesas — últimos 6 meses</h3>
          <MonthlyRevenueExpenseChart db={db} />
        </div>
        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
          <h3 className="mb-3 flex items-center gap-1.5 text-sm font-semibold text-slate-800"><Bell size={14} /> Alertas importantes</h3>
          <div className="max-h-64 space-y-2 overflow-y-auto pr-1">
            {notifications.length === 0 && <p className="text-xs text-slate-400">Nenhum alerta no momento. Tudo em dia!</p>}
            {notifications.slice(0, 8).map((n, i) => (
              <button key={n.key || i} onClick={() => { if (n.type === "moto") goMoto(n.entityId); else if (n.type === "renter") goRenter(n.entityId); else if (n.type === "contract") setTab("contratos"); else if (n.type === "payment") setTab("pagamentos"); }} className="flex w-full items-start gap-2 rounded-lg bg-slate-50 px-2.5 py-2 text-left hover:bg-slate-100">
                <AlertIcon level={n.level} />
                <div className="min-w-0">
                  <p className="text-xs font-medium text-slate-800">{n.title}</p>
                  <p className="text-[11px] text-slate-500">{n.detail}</p>
                </div>
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
          <h3 className="mb-3 text-sm font-semibold text-slate-800">Próximas manutenções</h3>
          <div className="space-y-2">
            {proximasManut.length === 0 && <p className="text-xs text-slate-400">Nenhuma manutenção próxima ou atrasada.</p>}
            {proximasManut.slice(0, 6).map((a, i) => (
              <div key={i} className="flex cursor-pointer items-center justify-between rounded-lg border border-slate-100 px-3 py-2 hover:bg-slate-50" onClick={() => goMoto(a.moto.id)}>
                <div>
                  <p className="text-sm font-medium text-slate-800">{a.type.name} — {motoLabel(a.moto)}</p>
                  <p className="text-xs text-slate-500">{a.kmRemaining <= 0 ? `Atrasada em ${fmtNum(Math.abs(a.kmRemaining))} km` : `Faltam ${fmtNum(a.kmRemaining)} km`}</p>
                </div>
                <Badge>{a.status}</Badge>
              </div>
            ))}
          </div>
        </div>
        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
          <h3 className="mb-3 text-sm font-semibold text-slate-800">Ranking de rentabilidade (mês)</h3>
          <div className="space-y-2">
            {profitability.slice(0, 5).map((p, i) => (
              <div key={p.moto.id} className="flex cursor-pointer items-center justify-between rounded-lg border border-slate-100 px-3 py-2 hover:bg-slate-50" onClick={() => goMoto(p.moto.id)}>
                <div className="flex items-center gap-2">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-slate-100 text-xs font-semibold text-slate-600">{i + 1}</span>
                  <div>
                    <p className="text-sm font-medium text-slate-800">{motoLabel(p.moto)}</p>
                    <p className="text-xs text-slate-500">Receita {fmtBRL(p.revenueTotal)} · Custos {fmtBRL(p.costTotal)}</p>
                  </div>
                </div>
                <span className={`text-sm font-semibold ${p.profit >= 0 ? "text-emerald-600" : "text-red-600"}`}>{fmtBRL(p.profit)}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function MonthlyRevenueExpenseChart({ db }) {
  const months = [];
  const now = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const s = d.toISOString().slice(0, 7) + "-01";
    const e = new Date(d.getFullYear(), d.getMonth() + 1, 0).toISOString().slice(0, 10);
    months.push({ label: d.toLocaleDateString("pt-BR", { month: "short" }), s, e });
  }
  const bars = months.map((m) => {
    const rev = db.payments.filter((p) => p.paidValue > 0 && inRange(p.paymentDate || p.dueDate, m.s, m.e)).reduce((a, p) => a + Number(p.paidValue), 0)
      + db.revenues.filter((r) => inRange(r.date, m.s, m.e)).reduce((a, r) => a + Number(r.value), 0);
    const exp = db.expenses.filter((x) => inRange(x.date, m.s, m.e)).reduce((a, x) => a + Number(x.value), 0)
      + db.maintenanceRecords.filter((mm) => inRange(mm.date, m.s, m.e)).reduce((a, mm) => a + Number(mm.totalCost), 0);
    return { label: m.label, a: rev, b: exp };
  });
  return (
    <div>
      <SvgBarChart bars={bars} />
      <div className="mt-2 flex gap-4 text-xs text-slate-500">
        <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-emerald-600" /> Receita</span>
        <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-red-600" /> Despesa</span>
      </div>
    </div>
  );
}

/* =========================================================================
   FROTA
========================================================================= */
function FleetModule({ db, patch, focusMotoId, setFocusMotoId, goRenter }) {
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);
  const detail = db.motorcycles.find((m) => m.id === focusMotoId);

  const fields = [
    { name: "brand", label: "Marca", required: true }, { name: "model", label: "Modelo", required: true },
    { name: "year", label: "Ano", type: "number" }, { name: "color", label: "Cor" },
    { name: "plate", label: "Placa", required: true }, { name: "renavam", label: "Renavam" },
    { name: "chassis", label: "Chassi" }, { name: "acquisitionDate", label: "Data de aquisição", type: "date" },
    { name: "acquisitionKm", label: "KM na aquisição", type: "number" }, { name: "currentKm", label: "KM atual", type: "number" },
    { name: "acquisitionValue", label: "Valor de aquisição (R$)", type: "number" }, { name: "marketValue", label: "Valor de mercado (R$)", type: "number" },
    { name: "status", label: "Status", type: "select", options: MOTO_STATUS, required: true },
    { name: "lastLicensingDate", label: "Último licenciamento", type: "date" }, { name: "nextLicensingDue", label: "Vencimento licenciamento", type: "date" },
    { name: "insurer", label: "Seguradora" }, { name: "insuranceDue", label: "Vencimento do seguro", type: "date" },
    { name: "notes", label: "Observações", type: "textarea", span: true },
  ];

  const rows = db.motorcycles.filter((m) => {
    const q = search.toLowerCase();
    const matchQ = !q || `${m.brand} ${m.model} ${m.plate} ${m.renavam}`.toLowerCase().includes(q);
    const matchS = !statusFilter || m.status === statusFilter;
    return matchQ && matchS;
  });

  const save = (values) => {
    const v = { ...values, year: Number(values.year) || "", currentKm: Number(values.currentKm) || 0, acquisitionKm: Number(values.acquisitionKm) || 0, acquisitionValue: Number(values.acquisitionValue) || 0, marketValue: Number(values.marketValue) || 0 };
    if (editing) patch((d) => ({ motorcycles: d.motorcycles.map((m) => (m.id === editing.id ? { ...m, ...v } : m)) }));
    else patch((d) => ({ motorcycles: [...d.motorcycles, { id: uid("moto"), ...v }] }));
    setEditing(null); setCreating(false);
  };
  const remove = () => {
    patch((d) => ({ motorcycles: d.motorcycles.filter((m) => m.id !== deleting.id) }));
    setDeleting(null);
  };

  if (detail) return <FleetDetail db={db} patch={patch} moto={detail} onBack={() => setFocusMotoId(null)} onEdit={() => setEditing(detail)} goRenter={goRenter} />;

  return (
    <div>
      <SectionHeader title="Frota" desc={`${db.motorcycles.length} moto(s) cadastrada(s)`} />
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por marca, modelo, placa..." onNew={() => setCreating(true)} newLabel="Nova moto"
        right={
          <select className={`${inputCls} w-40`} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
            <option value="">Todos os status</option>
            {MOTO_STATUS.map((s) => <option key={s}>{s}</option>)}
          </select>
        } />
      <DataTable
        columns={[
          { key: "moto", label: "Moto", render: (m) => (<button className="text-left font-medium text-slate-900 hover:text-orange-600" onClick={() => setFocusMotoId(m.id)}>{m.brand} {m.model}</button>) },
          { key: "plate", label: "Placa", mono: true },
          { key: "year", label: "Ano" },
          { key: "currentKm", label: "KM atual", render: (m) => fmtNum(m.currentKm) },
          { key: "status", label: "Status", render: (m) => <Badge>{m.status}</Badge> },
          { key: "manut", label: "Manutenção", render: (m) => {
              const alerts = maintenanceAlertsForMoto(db, m.id);
              const worst = alerts.some((a) => a.status === "Atrasado") ? "Atrasado" : alerts.some((a) => a.status === "Próxima") ? "Próxima" : "Em dia";
              return <Badge>{worst}</Badge>;
            } },
        ]}
        rows={rows}
        onEdit={(m) => setEditing(m)}
        onDelete={(m) => setDeleting(m)}
        onRowClick={(m) => setFocusMotoId(m.id)}
        emptyText="Nenhuma moto encontrada."
      />
      {(editing || creating) && <FormModal title={editing ? "Editar moto" : "Nova moto"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text={`Excluir a moto ${deleting.brand} ${deleting.model} (${deleting.plate})? Esta ação não pode ser desfeita.`} onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

function Tabs({ tabs, active, onChange }) {
  return (
    <div className="mb-4 flex gap-1 overflow-x-auto border-b border-slate-200">
      {tabs.map((t) => (
        <button key={t} onClick={() => onChange(t)} className={`whitespace-nowrap border-b-2 px-3 py-2 text-sm font-medium ${active === t ? "border-orange-600 text-orange-600" : "border-transparent text-slate-500 hover:text-slate-800"}`}>{t}</button>
      ))}
    </div>
  );
}

function FleetDetail({ db, patch, moto, onBack, onEdit, goRenter }) {
  const [tab, setTab] = useState("Informações");
  const mileage = db.mileageRecords.filter((r) => r.motorcycleId === moto.id).sort((a, b) => (a.date < b.date ? -1 : 1));
  const maint = db.maintenanceRecords.filter((r) => r.motorcycleId === moto.id).sort((a, b) => (a.date < b.date ? 1 : -1));
  const contracts = db.contracts.filter((c) => c.motorcycleId === moto.id).sort((a, b) => (a.startDate < b.startDate ? 1 : -1));
  const currentContract = contracts.find((c) => c.status === "Ativo");
  const expenses = db.expenses.filter((e) => e.motorcycleId === moto.id).sort((a, b) => (a.date < b.date ? 1 : -1));
  const revenues = db.revenues.filter((r) => r.motorcycleId === moto.id).sort((a, b) => (a.date < b.date ? 1 : -1));
  const fin = computeMotoFinancials(db, moto.id, "0000-01-01", todayISO());
  const roi = moto.acquisitionValue ? (fin.profit / Number(moto.acquisitionValue)) * 100 : 0;
  const alerts = maintenanceAlertsForMoto(db, moto.id);

  const weekAgo = addDays(todayISO(), -7), monthAgo = addDays(todayISO(), -30);
  const kmWeek = mileage.filter((r) => r.date >= weekAgo).reduce((s, r) => s + Number(r.kmDriven || 0), 0);
  const kmMonth = mileage.filter((r) => r.date >= monthAgo).reduce((s, r) => s + Number(r.kmDriven || 0), 0);

  return (
    <div>
      <button onClick={onBack} className="mb-4 flex items-center gap-1 text-sm font-medium text-slate-500 hover:text-orange-600"><ChevronLeft size={16} /> Voltar para a frota</button>
      <div className="mb-5 flex flex-col justify-between gap-3 sm:flex-row sm:items-center">
        <div>
          <h2 className="text-xl font-semibold text-slate-900" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{moto.brand} {moto.model} <span className="font-mono text-base text-slate-400">({moto.plate})</span></h2>
          <div className="mt-1 flex items-center gap-2"><Badge>{moto.status}</Badge><span className="text-sm text-slate-500">{fmtNum(moto.currentKm)} km rodados</span></div>
        </div>
        <button onClick={onEdit} className="flex items-center gap-1.5 rounded-lg border border-slate-300 px-3.5 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"><Pencil size={14} /> Editar dados</button>
      </div>

      <div className="mb-5 grid grid-cols-2 gap-3 sm:grid-cols-4">
        <KpiCard label="Receita acumulada" value={fmtBRL(fin.revenueTotal)} tone="emerald" />
        <KpiCard label="Custos acumulados" value={fmtBRL(fin.costTotal)} tone="red" />
        <KpiCard label="Lucro acumulado" value={fmtBRL(fin.profit)} tone={fin.profit >= 0 ? "emerald" : "red"} />
        <KpiCard label="ROI aproximado" value={`${fmtNum(roi, 1)}%`} />
      </div>

      <Tabs tabs={["Informações", "Locação atual", "Locatários", "Quilometragem", "Manutenções", "Despesas", "Receitas", "Documentação", "Rentabilidade"]} active={tab} onChange={setTab} />

      {tab === "Informações" && (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {[["Marca", moto.brand], ["Modelo", moto.model], ["Ano", moto.year], ["Cor", moto.color], ["Placa", moto.plate], ["Renavam", moto.renavam], ["Chassi", moto.chassis],
            ["Data de aquisição", fmtDate(moto.acquisitionDate)], ["KM na aquisição", fmtNum(moto.acquisitionKm)], ["Valor de aquisição", fmtBRL(moto.acquisitionValue)],
            ["Valor de mercado", fmtBRL(moto.marketValue)], ["Observações", moto.notes || "—"]].map(([l, v]) => (
            <div key={l} className="rounded-xl border border-slate-200 bg-white p-3"><p className="text-xs text-slate-400">{l}</p><p className="text-sm font-medium text-slate-800">{v || "—"}</p></div>
          ))}
        </div>
      )}

      {tab === "Locação atual" && (
        <div className="rounded-xl border border-slate-200 bg-white p-5">
          {currentContract ? (() => {
            const renter = db.renters.find((r) => r.id === currentContract.renterId);
            return (
              <div className="flex flex-col justify-between gap-3 sm:flex-row sm:items-center">
                <div>
                  <button onClick={() => goRenter(renter.id)} className="text-base font-semibold text-slate-900 hover:text-orange-600">{renter?.name}</button>
                  <p className="text-sm text-slate-500">{renter?.phone} · Contrato desde {fmtDate(currentContract.startDate)} até {fmtDate(currentContract.endDate)}</p>
                  <p className="mt-1 text-sm text-slate-600">Valor semanal: {fmtBRL(currentContract.weeklyValue)}</p>
                </div>
                <Badge>{currentContract.status}</Badge>
              </div>
            );
          })() : <p className="text-sm text-slate-400">Esta moto não está locada no momento.</p>}
        </div>
      )}

      {tab === "Locatários" && (
        <DataTable columns={[
          { key: "renter", label: "Locatário", render: (c) => { const r = db.renters.find((x) => x.id === c.renterId); return <button className="font-medium text-slate-800 hover:text-orange-600" onClick={() => goRenter(r?.id)}>{r?.name || "—"}</button>; } },
          { key: "startDate", label: "Início", render: (c) => fmtDate(c.startDate) },
          { key: "endDate", label: "Fim", render: (c) => fmtDate(c.endDate) },
          { key: "weeklyValue", label: "Valor semanal", render: (c) => fmtBRL(c.weeklyValue) },
          { key: "status", label: "Status", render: (c) => <Badge>{c.status}</Badge> },
        ]} rows={contracts} emptyText="Nenhum locatário no histórico." />
      )}

      {tab === "Quilometragem" && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
            <KpiCard label="KM atual" value={fmtNum(moto.currentKm)} />
            <KpiCard label="KM na semana" value={fmtNum(kmWeek)} />
            <KpiCard label="KM no mês" value={fmtNum(kmMonth)} />
            <KpiCard label="Média KM/dia" value={fmtNum(mileage.length ? (moto.currentKm - moto.acquisitionKm) / Math.max(1, daysDiff(todayISO(), moto.acquisitionDate)) : 0, 1)} />
          </div>
          <div className="rounded-xl border border-slate-200 bg-white p-4">
            <h4 className="mb-2 text-sm font-semibold text-slate-700">Evolução da quilometragem</h4>
            <SvgLineChart points={mileage.map((r) => ({ label: fmtDate(r.date).slice(0, 5), value: r.km }))} />
          </div>
          <DataTable columns={[
            { key: "date", label: "Data", render: (r) => fmtDate(r.date) }, { key: "previousKm", label: "KM anterior", render: (r) => fmtNum(r.previousKm) },
            { key: "km", label: "KM registrado", render: (r) => fmtNum(r.km) }, { key: "kmDriven", label: "KM rodados", render: (r) => fmtNum(r.kmDriven) },
            { key: "recordedBy", label: "Registrado por" }, { key: "notes", label: "Observação" },
          ]} rows={[...mileage].reverse()} emptyText="Nenhum registro de quilometragem." />
        </div>
      )}

      {tab === "Manutenções" && (
        <div className="space-y-4">
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
            {alerts.filter((a) => a.status !== "Em dia").map((a) => (
              <div key={a.type.id} className="flex items-center justify-between rounded-lg border border-slate-200 bg-white px-3 py-2">
                <span className="text-sm text-slate-700">{a.type.name}</span>
                <Badge>{a.status}</Badge>
              </div>
            ))}
          </div>
          <DataTable columns={[
            { key: "date", label: "Data", render: (r) => fmtDate(r.date) },
            { key: "type", label: "Tipo", render: (r) => db.maintenanceTypes.find((t) => t.id === r.typeId)?.name || "—" },
            { key: "km", label: "KM", render: (r) => fmtNum(r.km) },
            { key: "totalCost", label: "Custo total", render: (r) => fmtBRL(r.totalCost) },
            { key: "vendor", label: "Oficina" },
            { key: "nextKm", label: "Próx. KM", render: (r) => fmtNum(r.nextKm) },
          ]} rows={maint} emptyText="Nenhuma manutenção registrada." />
        </div>
      )}

      {tab === "Despesas" && (
        <DataTable columns={[
          { key: "date", label: "Data", render: (e) => fmtDate(e.date) }, { key: "category", label: "Categoria" },
          { key: "vendor", label: "Fornecedor" }, { key: "value", label: "Valor", render: (e) => fmtBRL(e.value) },
        ]} rows={expenses} emptyText="Nenhuma despesa registrada." />
      )}

      {tab === "Receitas" && (
        <DataTable columns={[
          { key: "date", label: "Data", render: (r) => fmtDate(r.date) }, { key: "category", label: "Categoria" },
          { key: "renter", label: "Locatário", render: (r) => db.renters.find((x) => x.id === r.renterId)?.name || "—" },
          { key: "value", label: "Valor", render: (r) => fmtBRL(r.value) },
        ]} rows={revenues} emptyText="Nenhuma receita adicional registrada." />
      )}

      {tab === "Documentação" && (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
          <div className="rounded-xl border border-slate-200 bg-white p-4">
            <h4 className="mb-2 flex items-center gap-1.5 text-sm font-semibold text-slate-700"><ShieldCheck size={15} /> Licenciamento</h4>
            <p className="text-sm text-slate-600">Último: {fmtDate(moto.lastLicensingDate)}</p>
            <p className="text-sm text-slate-600">Próximo vencimento: {fmtDate(moto.nextLicensingDue)}</p>
          </div>
          <div className="rounded-xl border border-slate-200 bg-white p-4">
            <h4 className="mb-2 flex items-center gap-1.5 text-sm font-semibold text-slate-700"><ShieldCheck size={15} /> Seguro</h4>
            <p className="text-sm text-slate-600">Seguradora: {moto.insurer || "—"}</p>
            <p className="text-sm text-slate-600">Vencimento: {fmtDate(moto.insuranceDue)}</p>
          </div>
        </div>
      )}

      {tab === "Rentabilidade" && (
        <div className="space-y-3">
          <div className="rounded-xl border border-slate-200 bg-white p-4 text-sm">
            <div className="flex justify-between border-b border-slate-100 py-1.5"><span className="text-slate-500">Receita de aluguel</span><span className="font-medium">{fmtBRL(fin.rentPaid)}</span></div>
            <div className="flex justify-between border-b border-slate-100 py-1.5"><span className="text-slate-500">Outras receitas</span><span className="font-medium">{fmtBRL(fin.rev)}</span></div>
            <div className="flex justify-between border-b border-slate-100 py-1.5"><span className="text-slate-500">Manutenção</span><span className="font-medium text-red-600">- {fmtBRL(fin.maint)}</span></div>
            <div className="flex justify-between border-b border-slate-100 py-1.5"><span className="text-slate-500">Outras despesas</span><span className="font-medium text-red-600">- {fmtBRL(fin.exp)}</span></div>
            <div className="flex justify-between py-2 text-base font-semibold"><span>Lucro acumulado</span><span className={fin.profit >= 0 ? "text-emerald-600" : "text-red-600"}>{fmtBRL(fin.profit)}</span></div>
          </div>
          <p className="text-xs text-slate-400">ROI aproximado = lucro acumulado ÷ valor investido na moto = {fmtNum(roi, 1)}%</p>
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   LOCATÁRIOS
========================================================================= */
function RentersModule({ db, patch, focusRenterId, setFocusRenterId, goMoto }) {
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);
  const detail = db.renters.find((r) => r.id === focusRenterId);

  const fields = [
    { name: "name", label: "Nome completo", required: true, span: true }, { name: "cpf", label: "CPF" },
    { name: "birthDate", label: "Data de nascimento", type: "date" }, { name: "phone", label: "Telefone" },
    { name: "whatsapp", label: "WhatsApp" }, { name: "email", label: "E-mail", type: "email" },
    { name: "address", label: "Endereço", span: true }, { name: "cnh", label: "CNH" },
    { name: "cnhCategory", label: "Categoria da CNH" }, { name: "cnhValidity", label: "Validade da CNH", type: "date" },
    { name: "status", label: "Status", type: "select", options: RENTER_STATUS, required: true },
    { name: "notes", label: "Observações", type: "textarea", span: true },
  ];

  const rows = db.renters.filter((r) => !search || `${r.name} ${r.cpf}`.toLowerCase().includes(search.toLowerCase()));

  const save = (values) => {
    if (editing) patch((d) => ({ renters: d.renters.map((r) => (r.id === editing.id ? { ...r, ...values } : r)) }));
    else patch((d) => ({ renters: [...d.renters, { id: uid("r"), registrationDate: todayISO(), ...values }] }));
    setEditing(null); setCreating(false);
  };
  const remove = () => { patch((d) => ({ renters: d.renters.filter((r) => r.id !== deleting.id) })); setDeleting(null); };

  if (detail) return <RenterDetail db={db} renter={detail} onBack={() => setFocusRenterId(null)} onEdit={() => setEditing(detail)} goMoto={goMoto} />;

  return (
    <div>
      <SectionHeader title="Locatários" desc={`${db.renters.length} locatário(s) cadastrado(s)`} />
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por nome ou CPF..." onNew={() => setCreating(true)} newLabel="Novo locatário" />
      <DataTable columns={[
        { key: "name", label: "Nome", render: (r) => <button className="font-medium text-slate-900 hover:text-orange-600" onClick={() => setFocusRenterId(r.id)}>{r.name}</button> },
        { key: "cpf", label: "CPF", mono: true }, { key: "phone", label: "Telefone" },
        { key: "cnhValidity", label: "Validade CNH", render: (r) => fmtDate(r.cnhValidity) },
        { key: "status", label: "Status", render: (r) => <Badge>{r.status}</Badge> },
      ]} rows={rows} onEdit={setEditing} onDelete={setDeleting} onRowClick={(r) => setFocusRenterId(r.id)} emptyText="Nenhum locatário encontrado." />
      {(editing || creating) && <FormModal title={editing ? "Editar locatário" : "Novo locatário"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text={`Excluir o locatário ${deleting.name}?`} onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

function RenterDetail({ db, renter, onBack, onEdit, goMoto }) {
  const [tab, setTab] = useState("Dados");
  const contracts = db.contracts.filter((c) => c.renterId === renter.id).sort((a, b) => (a.startDate < b.startDate ? 1 : -1));
  const payments = db.payments.filter((p) => p.renterId === renter.id).sort((a, b) => (a.dueDate < b.dueDate ? 1 : -1));
  const incidents = db.incidents.filter((i) => i.renterId === renter.id);
  const stats = computeRenterStats(db, renter.id);

  return (
    <div>
      <button onClick={onBack} className="mb-4 flex items-center gap-1 text-sm font-medium text-slate-500 hover:text-orange-600"><ChevronLeft size={16} /> Voltar para locatários</button>
      <div className="mb-5 flex flex-col justify-between gap-3 sm:flex-row sm:items-center">
        <div>
          <h2 className="text-xl font-semibold text-slate-900" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{renter.name}</h2>
          <div className="mt-1 flex items-center gap-2"><Badge>{renter.status}</Badge><span className="text-sm text-slate-500">{renter.phone}</span></div>
        </div>
        <button onClick={onEdit} className="flex items-center gap-1.5 rounded-lg border border-slate-300 px-3.5 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"><Pencil size={14} /> Editar dados</button>
      </div>

      <div className="mb-5 grid grid-cols-2 gap-3 sm:grid-cols-4">
        <KpiCard label="Receita gerada" value={fmtBRL(stats.revenue)} tone="emerald" />
        <KpiCard label="Semanas alugadas" value={stats.weeks} />
        <KpiCard label="Atrasos" value={stats.late} tone={stats.late ? "red" : "slate"} />
        <KpiCard label="Contratos" value={stats.contracts} />
      </div>

      <Tabs tabs={["Dados", "Motos alugadas", "Pagamentos", "Ocorrências"]} active={tab} onChange={setTab} />

      {tab === "Dados" && (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {[["CPF", renter.cpf], ["Data de nascimento", fmtDate(renter.birthDate)], ["Telefone", renter.phone], ["WhatsApp", renter.whatsapp], ["E-mail", renter.email],
            ["Endereço", renter.address], ["CNH", renter.cnh], ["Categoria CNH", renter.cnhCategory], ["Validade CNH", fmtDate(renter.cnhValidity)],
            ["Cadastrado em", fmtDate(renter.registrationDate)], ["Observações", renter.notes || "—"]].map(([l, v]) => (
            <div key={l} className="rounded-xl border border-slate-200 bg-white p-3"><p className="text-xs text-slate-400">{l}</p><p className="text-sm font-medium text-slate-800">{v || "—"}</p></div>
          ))}
        </div>
      )}

      {tab === "Motos alugadas" && (
        <DataTable columns={[
          { key: "moto", label: "Moto", render: (c) => { const m = db.motorcycles.find((x) => x.id === c.motorcycleId); return <button className="font-medium text-slate-800 hover:text-orange-600" onClick={() => goMoto(m?.id)}>{motoLabel(m)}</button>; } },
          { key: "startDate", label: "Início", render: (c) => fmtDate(c.startDate) }, { key: "endDate", label: "Fim", render: (c) => fmtDate(c.endDate) },
          { key: "weeklyValue", label: "Valor semanal", render: (c) => fmtBRL(c.weeklyValue) }, { key: "status", label: "Status", render: (c) => <Badge>{c.status}</Badge> },
        ]} rows={contracts} emptyText="Nenhuma moto alugada." />
      )}

      {tab === "Pagamentos" && (
        <DataTable columns={[
          { key: "dueDate", label: "Vencimento", render: (p) => fmtDate(p.dueDate) }, { key: "expectedValue", label: "Valor previsto", render: (p) => fmtBRL(p.expectedValue) },
          { key: "paidValue", label: "Valor pago", render: (p) => fmtBRL(p.paidValue) }, { key: "status", label: "Status", render: (p) => <Badge>{paymentStatus(p, todayISO())}</Badge> },
        ]} rows={payments} emptyText="Nenhum pagamento registrado." />
      )}

      {tab === "Ocorrências" && (
        <DataTable columns={[
          { key: "date", label: "Data", render: (i) => fmtDate(i.date) }, { key: "type", label: "Tipo" },
          { key: "value", label: "Valor", render: (i) => fmtBRL(i.value) }, { key: "status", label: "Status", render: (i) => <Badge>{i.status}</Badge> },
        ]} rows={incidents} emptyText="Nenhuma ocorrência registrada." />
      )}
    </div>
  );
}

/* =========================================================================
   CONTRATOS
========================================================================= */
function ContractsModule({ db, patch }) {
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);

  const availableMotos = (currentId) => db.motorcycles.filter((m) => m.status === "Disponível" || m.id === currentId);

  const fields = [
    { name: "renterId", label: "Locatário", type: "select", required: true, options: () => db.renters.map((r) => ({ value: r.id, label: r.name })) },
    { name: "motorcycleId", label: "Moto", type: "select", required: true, options: (v) => availableMotos(v.motorcycleId).map((m) => ({ value: m.id, label: motoLabel(m) })), hint: "Somente motos disponíveis (ou a atual, em edição) aparecem aqui." },
    { name: "startDate", label: "Data de início", type: "date", required: true }, { name: "endDate", label: "Data de término", type: "date", required: true },
    { name: "dailyValue", label: "Valor da diária (R$)", type: "number" }, { name: "weeklyValue", label: "Valor semanal (R$)", type: "number", required: true },
    { name: "monthlyValue", label: "Valor mensal (R$)", type: "number" }, { name: "deposit", label: "Caução (R$)", type: "number" },
    { name: "periodicity", label: "Periodicidade do pagamento", type: "select", options: PERIODICITY },
    { name: "dueDay", label: "Dia de vencimento" }, { name: "kmLimit", label: "Limite de KM (0 = sem limite)", type: "number" },
    { name: "kmExcessValue", label: "Valor por KM excedente (R$)", type: "number" },
    { name: "status", label: "Status", type: "select", options: CONTRACT_STATUS, required: true },
    { name: "rules", label: "Regras específicas", type: "textarea", span: true }, { name: "notes", label: "Observações", type: "textarea", span: true },
  ];

  const rows = db.contracts.filter((c) => {
    const moto = db.motorcycles.find((m) => m.id === c.motorcycleId);
    const renter = db.renters.find((r) => r.id === c.renterId);
    const q = search.toLowerCase();
    const matchQ = !q || `${renter?.name} ${moto?.plate}`.toLowerCase().includes(q);
    const matchS = !statusFilter || c.status === statusFilter;
    return matchQ && matchS;
  }).sort((a, b) => (a.startDate < b.startDate ? 1 : -1));

  const save = (values) => {
    const v = { ...values, dailyValue: Number(values.dailyValue) || 0, weeklyValue: Number(values.weeklyValue) || 0, monthlyValue: Number(values.monthlyValue) || 0, deposit: Number(values.deposit) || 0, kmLimit: Number(values.kmLimit) || 0, kmExcessValue: Number(values.kmExcessValue) || 0 };
    const prevMoto = editing ? db.motorcycles.find((m) => m.id === editing.motorcycleId) : null;

    if (editing) {
      const priceChanged = Number(editing.weeklyValue) !== v.weeklyValue;
      patch((d) => ({
        contracts: d.contracts.map((c) => c.id === editing.id ? { ...c, ...v, priceHistory: priceChanged ? [...(c.priceHistory || []), { date: todayISO(), weeklyValue: v.weeklyValue }] : c.priceHistory } : c),
        motorcycles: d.motorcycles.map((m) => {
          if (m.id === v.motorcycleId) return { ...m, status: v.status === "Ativo" ? "Alugada" : m.status === "Alugada" ? "Disponível" : m.status };
          if (prevMoto && m.id === prevMoto.id && prevMoto.id !== v.motorcycleId) return { ...m, status: "Disponível" };
          return m;
        }),
      }));
    } else {
      const moto = db.motorcycles.find((m) => m.id === v.motorcycleId);
      if (moto && moto.status === "Em manutenção") { alert("Esta moto está em manutenção e não pode ser atribuída a um novo contrato."); return; }
      patch((d) => ({
        contracts: [...d.contracts, { id: uid("c"), priceHistory: [{ date: todayISO(), weeklyValue: v.weeklyValue }], ...v }],
        motorcycles: d.motorcycles.map((m) => (m.id === v.motorcycleId && v.status === "Ativo" ? { ...m, status: "Alugada" } : m)),
      }));
    }
    setEditing(null); setCreating(false);
  };

  const remove = () => {
    patch((d) => ({
      contracts: d.contracts.filter((c) => c.id !== deleting.id),
      motorcycles: d.motorcycles.map((m) => (m.id === deleting.motorcycleId && m.status === "Alugada" ? { ...m, status: "Disponível" } : m)),
    }));
    setDeleting(null);
  };

  const finalize = (c, status) => {
    patch((d) => ({
      contracts: d.contracts.map((x) => (x.id === c.id ? { ...x, status } : x)),
      motorcycles: d.motorcycles.map((m) => (m.id === c.motorcycleId && !["Em manutenção", "Reservada"].includes(m.status) ? { ...m, status: "Disponível" } : m)),
    }));
  };

  return (
    <div>
      <SectionHeader title="Contratos de locação" desc={`${db.contracts.length} contrato(s)`} />
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por locatário ou placa..." onNew={() => setCreating(true)} newLabel="Novo contrato"
        right={<select className={`${inputCls} w-40`} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}><option value="">Todos os status</option>{CONTRACT_STATUS.map((s) => <option key={s}>{s}</option>)}</select>} />
      <DataTable columns={[
        { key: "renter", label: "Locatário", render: (c) => db.renters.find((r) => r.id === c.renterId)?.name || "—" },
        { key: "moto", label: "Moto", render: (c) => motoLabel(db.motorcycles.find((m) => m.id === c.motorcycleId)) },
        { key: "period", label: "Período", render: (c) => `${fmtDate(c.startDate)} — ${fmtDate(c.endDate)}` },
        { key: "weeklyValue", label: "Valor semanal", render: (c) => fmtBRL(c.weeklyValue) },
        { key: "status", label: "Status", render: (c) => <Badge>{c.status}</Badge> },
        { key: "actions2", label: "Encerrar", render: (c) => c.status === "Ativo" ? (
            <div className="flex gap-1">
              <button onClick={() => finalize(c, "Finalizado")} className="rounded border border-slate-300 px-2 py-1 text-xs hover:bg-slate-50">Finalizar</button>
              <button onClick={() => finalize(c, "Cancelado")} className="rounded border border-slate-300 px-2 py-1 text-xs hover:bg-slate-50">Cancelar</button>
            </div>
          ) : "—" },
      ]} rows={rows} onEdit={setEditing} onDelete={setDeleting} emptyText="Nenhum contrato encontrado." />
      {(editing || creating) && <FormModal title={editing ? "Editar contrato" : "Novo contrato"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text="Excluir este contrato? A moto voltará a ficar disponível." onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

/* =========================================================================
   PAGAMENTOS + CONTROLE SEMANAL
========================================================================= */
function PaymentsModule({ db, patch }) {
  const [view, setView] = useState("Todos os pagamentos");
  return (
    <div>
      <SectionHeader title="Pagamentos" desc="Controle financeiro das locações" />
      <Tabs tabs={["Todos os pagamentos", "Controle semanal"]} active={view} onChange={setView} />
      {view === "Todos os pagamentos" ? <AllPayments db={db} patch={patch} /> : <WeeklyControl db={db} patch={patch} />}
    </div>
  );
}

function AllPayments({ db, patch }) {
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);
  const today = todayISO();

  const fields = [
    { name: "contractId", label: "Contrato", type: "select", required: true, options: () => db.contracts.map((c) => { const r = db.renters.find((x) => x.id === c.renterId); const m = db.motorcycles.find((x) => x.id === c.motorcycleId); return { value: c.id, label: `${r?.name} — ${m?.plate}` }; }) },
    { name: "dueDate", label: "Data de vencimento", type: "date", required: true }, { name: "paymentDate", label: "Data do pagamento", type: "date" },
    { name: "expectedValue", label: "Valor previsto (R$)", type: "number", required: true }, { name: "paidValue", label: "Valor pago (R$)", type: "number" },
    { name: "paymentMethod", label: "Forma de pagamento", type: "select", options: PAYMENT_METHODS },
    { name: "status", label: "Status (manual, opcional)", type: "select", options: PAYMENT_STATUS, hint: "Deixe em branco para cálculo automático." },
    { name: "notes", label: "Observação", type: "textarea", span: true },
  ];

  const rows = db.payments.filter((p) => {
    const c = db.contracts.find((x) => x.id === p.contractId);
    const renter = db.renters.find((r) => r.id === (p.renterId || c?.renterId));
    const moto = db.motorcycles.find((m) => m.id === (p.motorcycleId || c?.motorcycleId));
    const st = p.status === "Cancelado" ? "Cancelado" : paymentStatus(p, today);
    const q = search.toLowerCase();
    const matchQ = !q || `${renter?.name} ${moto?.plate}`.toLowerCase().includes(q);
    const matchS = !statusFilter || st === statusFilter;
    return matchQ && matchS;
  }).sort((a, b) => (a.dueDate < b.dueDate ? 1 : -1));

  const totals = {
    previsto: rows.reduce((s, p) => s + Number(p.expectedValue), 0),
    pago: rows.reduce((s, p) => s + Number(p.paidValue || 0), 0),
  };

  const save = (values) => {
    const c = db.contracts.find((x) => x.id === values.contractId);
    const v = { ...values, expectedValue: Number(values.expectedValue) || 0, paidValue: Number(values.paidValue) || 0, renterId: c?.renterId, motorcycleId: c?.motorcycleId, status: values.status || "" };
    if (editing) patch((d) => ({ payments: d.payments.map((p) => (p.id === editing.id ? { ...p, ...v } : p)) }));
    else patch((d) => ({ payments: [...d.payments, { id: uid("pay"), ...v }] }));
    setEditing(null); setCreating(false);
  };
  const remove = () => { patch((d) => ({ payments: d.payments.filter((p) => p.id !== deleting.id) })); setDeleting(null); };

  return (
    <div>
      <div className="mb-4 grid grid-cols-2 gap-3 sm:grid-cols-4">
        <KpiCard label="Total previsto" value={fmtBRL(totals.previsto)} />
        <KpiCard label="Total recebido" value={fmtBRL(totals.pago)} tone="emerald" />
        <KpiCard label="Total em aberto" value={fmtBRL(totals.previsto - totals.pago)} />
        <KpiCard label="Taxa de inadimplência" value={`${fmtNum(totals.previsto ? ((totals.previsto - totals.pago) / totals.previsto) * 100 : 0, 1)}%`} />
      </div>
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por locatário ou placa..." onNew={() => setCreating(true)} newLabel="Novo pagamento"
        right={<select className={`${inputCls} w-40`} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}><option value="">Todos os status</option>{PAYMENT_STATUS.map((s) => <option key={s}>{s}</option>)}</select>} />
      <DataTable columns={[
        { key: "renter", label: "Locatário", render: (p) => db.renters.find((r) => r.id === p.renterId)?.name || "—" },
        { key: "moto", label: "Moto", render: (p) => db.motorcycles.find((m) => m.id === p.motorcycleId)?.plate || "—" },
        { key: "dueDate", label: "Vencimento", render: (p) => fmtDate(p.dueDate) },
        { key: "expectedValue", label: "Previsto", render: (p) => fmtBRL(p.expectedValue) },
        { key: "paidValue", label: "Pago", render: (p) => fmtBRL(p.paidValue) },
        { key: "diff", label: "Diferença", render: (p) => fmtBRL(Number(p.paidValue || 0) - Number(p.expectedValue)) },
        { key: "status", label: "Status", render: (p) => <Badge>{p.status === "Cancelado" ? "Cancelado" : paymentStatus(p, today)}</Badge> },
      ]} rows={rows} onEdit={setEditing} onDelete={setDeleting} emptyText="Nenhum pagamento encontrado." />
      {(editing || creating) && <FormModal title={editing ? "Editar pagamento" : "Novo pagamento"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text="Excluir este pagamento?" onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

function WeeklyControl({ db, patch }) {
  const [offset, setOffset] = useState(0);
  const weekStart = addDays(startOfWeek(todayISO()), offset * 7);
  const weekEnd = addDays(weekStart, 6);
  const [quickPay, setQuickPay] = useState(null);

  const rows = db.payments.filter((p) => inRange(p.dueDate, weekStart, weekEnd)).map((p) => {
    const c = db.contracts.find((x) => x.id === p.contractId);
    return { ...p, renter: db.renters.find((r) => r.id === p.renterId)?.name, moto: db.motorcycles.find((m) => m.id === p.motorcycleId)?.plate, contract: c };
  });

  const totals = {
    previsto: rows.reduce((s, p) => s + Number(p.expectedValue), 0),
    recebido: rows.reduce((s, p) => s + Number(p.paidValue || 0), 0),
    pendente: rows.filter((p) => paymentStatus(p, todayISO()) === "Pendente").reduce((s, p) => s + Number(p.expectedValue) - Number(p.paidValue || 0), 0),
    atrasado: rows.filter((p) => paymentStatus(p, todayISO()) === "Atrasado").reduce((s, p) => s + Number(p.expectedValue) - Number(p.paidValue || 0), 0),
  };

  const markPaid = (p, method) => {
    patch((d) => ({ payments: d.payments.map((x) => (x.id === p.id ? { ...x, paidValue: x.expectedValue, paymentDate: todayISO(), paymentMethod: method, status: "" } : x)) }));
    setQuickPay(null);
  };

  return (
    <div>
      <div className="mb-4 flex items-center justify-between rounded-xl border border-slate-200 bg-white p-3">
        <button onClick={() => setOffset((o) => o - 1)} className="flex items-center gap-1 rounded-lg px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50"><ChevronLeft size={16} /> Semana anterior</button>
        <span className="text-sm font-semibold text-slate-800">{fmtDate(weekStart)} — {fmtDate(weekEnd)} {offset === 0 && <span className="ml-1 text-orange-600">(atual)</span>}</span>
        <button onClick={() => setOffset((o) => o + 1)} className="flex items-center gap-1 rounded-lg px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50">Próxima semana <ChevronRight size={16} /></button>
      </div>
      <div className="mb-4 grid grid-cols-2 gap-3 sm:grid-cols-4">
        <KpiCard label="Total previsto" value={fmtBRL(totals.previsto)} />
        <KpiCard label="Total recebido" value={fmtBRL(totals.recebido)} tone="emerald" />
        <KpiCard label="Total pendente" value={fmtBRL(totals.pendente)} />
        <KpiCard label="Total atrasado" value={fmtBRL(totals.atrasado)} tone={totals.atrasado ? "red" : "slate"} />
      </div>
      <DataTable columns={[
        { key: "moto", label: "Moto" }, { key: "renter", label: "Locatário" },
        { key: "semana", label: "Semana", render: () => `${fmtDate(weekStart)}` },
        { key: "expectedValue", label: "Valor previsto", render: (p) => fmtBRL(p.expectedValue) },
        { key: "paidValue", label: "Valor pago", render: (p) => fmtBRL(p.paidValue) },
        { key: "paymentDate", label: "Data pagamento", render: (p) => fmtDate(p.paymentDate) },
        { key: "status", label: "Status", render: (p) => <Badge>{paymentStatus(p, todayISO())}</Badge> },
        { key: "quick", label: "Ação", render: (p) => paymentStatus(p, todayISO()) !== "Pago" ? (
            <button onClick={() => setQuickPay(p)} className="rounded-lg bg-emerald-600 px-2.5 py-1 text-xs font-medium text-white hover:bg-emerald-700">Registrar pagamento</button>
          ) : <span className="text-xs text-slate-400">—</span> },
        { key: "notes", label: "Observação", render: (p) => p.notes || "—" },
      ]} rows={rows} pageSize={12} emptyText="Nenhum pagamento previsto para esta semana." />

      {quickPay && (
        <Modal title="Registrar pagamento" onClose={() => setQuickPay(null)}>
          <p className="mb-3 text-sm text-slate-600">{quickPay.renter} · {quickPay.moto} · {fmtBRL(quickPay.expectedValue)}</p>
          <div className="grid grid-cols-2 gap-2">
            {PAYMENT_METHODS.map((m) => <button key={m} onClick={() => markPaid(quickPay, m)} className="rounded-lg border border-slate-300 px-3 py-2 text-sm font-medium hover:border-orange-500 hover:text-orange-600">{m}</button>)}
          </div>
        </Modal>
      )}
    </div>
  );
}

/* =========================================================================
   RECEITAS
========================================================================= */
function RevenuesModule({ db, patch }) {
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);

  const fields = [
    { name: "date", label: "Data", type: "date", required: true },
    { name: "category", label: "Categoria", type: "select", options: db.settings.revenueCategories || REVENUE_CATEGORIES, required: true },
    { name: "motorcycleId", label: "Moto", type: "select", options: () => db.motorcycles.map((m) => ({ value: m.id, label: motoLabel(m) })) },
    { name: "renterId", label: "Locatário", type: "select", options: () => db.renters.map((r) => ({ value: r.id, label: r.name })) },
    { name: "value", label: "Valor (R$)", type: "number", required: true },
    { name: "paymentMethod", label: "Forma de pagamento", type: "select", options: PAYMENT_METHODS },
    { name: "notes", label: "Observação", type: "textarea", span: true },
  ];

  const rows = db.revenues.filter((r) => {
    const moto = db.motorcycles.find((m) => m.id === r.motorcycleId);
    const q = search.toLowerCase();
    return !q || `${r.category} ${moto?.plate}`.toLowerCase().includes(q);
  }).sort((a, b) => (a.date < b.date ? 1 : -1));
  const total = rows.reduce((s, r) => s + Number(r.value), 0);

  const save = (values) => {
    const v = { ...values, value: Number(values.value) || 0 };
    if (editing) patch((d) => ({ revenues: d.revenues.map((r) => (r.id === editing.id ? { ...r, ...v } : r)) }));
    else patch((d) => ({ revenues: [...d.revenues, { id: uid("rev"), ...v }] }));
    setEditing(null); setCreating(false);
  };
  const remove = () => { patch((d) => ({ revenues: d.revenues.filter((r) => r.id !== deleting.id) })); setDeleting(null); };

  return (
    <div>
      <SectionHeader title="Receitas" desc="Receitas extras: taxas, km excedente, multas repassadas, danos e outras entradas" actions={<KpiCard label="Total no filtro" value={fmtBRL(total)} tone="orange" />} />
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por categoria ou placa..." onNew={() => setCreating(true)} newLabel="Nova receita" />
      <DataTable columns={[
        { key: "date", label: "Data", render: (r) => fmtDate(r.date) }, { key: "category", label: "Categoria" },
        { key: "moto", label: "Moto", render: (r) => db.motorcycles.find((m) => m.id === r.motorcycleId)?.plate || "—" },
        { key: "renter", label: "Locatário", render: (r) => db.renters.find((x) => x.id === r.renterId)?.name || "—" },
        { key: "value", label: "Valor", render: (r) => fmtBRL(r.value) }, { key: "paymentMethod", label: "Forma de pagto" },
      ]} rows={rows} onEdit={setEditing} onDelete={setDeleting} emptyText="Nenhuma receita registrada." />
      {(editing || creating) && <FormModal title={editing ? "Editar receita" : "Nova receita"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text="Excluir esta receita?" onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

/* =========================================================================
   DESPESAS
========================================================================= */
function ExpensesModule({ db, patch }) {
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);

  const fields = [
    { name: "date", label: "Data", type: "date", required: true },
    { name: "category", label: "Categoria", type: "select", options: db.settings.expenseCategories || EXPENSE_CATEGORIES, required: true },
    { name: "motorcycleId", label: "Moto (opcional)", type: "select", options: () => db.motorcycles.map((m) => ({ value: m.id, label: motoLabel(m) })) },
    { name: "vendor", label: "Fornecedor" }, { name: "value", label: "Valor (R$)", type: "number", required: true },
    { name: "paymentMethod", label: "Forma de pagamento", type: "select", options: PAYMENT_METHODS },
    { name: "notes", label: "Observação", type: "textarea", span: true },
  ];

  const rows = db.expenses.filter((e) => {
    const moto = db.motorcycles.find((m) => m.id === e.motorcycleId);
    const q = search.toLowerCase();
    return !q || `${e.category} ${e.vendor} ${moto?.plate}`.toLowerCase().includes(q);
  }).sort((a, b) => (a.date < b.date ? 1 : -1));
  const total = rows.reduce((s, e) => s + Number(e.value), 0);

  const save = (values) => {
    const v = { ...values, value: Number(values.value) || 0 };
    if (editing) patch((d) => ({ expenses: d.expenses.map((e) => (e.id === editing.id ? { ...e, ...v } : e)) }));
    else patch((d) => ({ expenses: [...d.expenses, { id: uid("exp"), ...v }] }));
    setEditing(null); setCreating(false);
  };
  const remove = () => { patch((d) => ({ expenses: d.expenses.filter((e) => e.id !== deleting.id) })); setDeleting(null); };

  return (
    <div>
      <SectionHeader title="Despesas" desc="Combustível, peças, documentação e demais custos" actions={<KpiCard label="Total no filtro" value={fmtBRL(total)} />} />
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por categoria, fornecedor ou placa..." onNew={() => setCreating(true)} newLabel="Nova despesa" />
      <DataTable columns={[
        { key: "date", label: "Data", render: (e) => fmtDate(e.date) }, { key: "category", label: "Categoria" },
        { key: "moto", label: "Moto", render: (e) => db.motorcycles.find((m) => m.id === e.motorcycleId)?.plate || "—" },
        { key: "vendor", label: "Fornecedor" }, { key: "value", label: "Valor", render: (e) => fmtBRL(e.value) },
      ]} rows={rows} onEdit={setEditing} onDelete={setDeleting} emptyText="Nenhuma despesa registrada." />
      {(editing || creating) && <FormModal title={editing ? "Editar despesa" : "Nova despesa"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text="Excluir esta despesa?" onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

/* =========================================================================
   MANUTENÇÃO
========================================================================= */
function MaintenanceModule({ db, patch }) {
  const [view, setView] = useState("Registros");
  return (
    <div>
      <SectionHeader title="Manutenções" desc="Preventivas, corretivas e alertas de vencimento" />
      <Tabs tabs={["Registros", "Central de alertas", "Tipos e intervalos"]} active={view} onChange={setView} />
      {view === "Registros" && <MaintenanceRecords db={db} patch={patch} />}
      {view === "Central de alertas" && <MaintenanceAlertsCenter db={db} />}
      {view === "Tipos e intervalos" && <MaintenanceTypesConfig db={db} patch={patch} />}
    </div>
  );
}

function MaintenanceRecords({ db, patch }) {
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);

  const fields = [
    { name: "motorcycleId", label: "Moto", type: "select", required: true, options: () => db.motorcycles.map((m) => ({ value: m.id, label: motoLabel(m) })) },
    { name: "typeId", label: "Tipo de manutenção", type: "select", required: true, options: () => db.maintenanceTypes.map((t) => ({ value: t.id, label: t.name })) },
    { name: "date", label: "Data", type: "date", required: true }, { name: "km", label: "Quilometragem no momento", type: "number", required: true },
    { name: "partsCost", label: "Custo da peça (R$)", type: "number" }, { name: "laborCost", label: "Custo da mão de obra (R$)", type: "number" },
    { name: "vendor", label: "Oficina / fornecedor" }, { name: "notes", label: "Observações", type: "textarea", span: true },
  ];

  const rows = db.maintenanceRecords.filter((r) => {
    const moto = db.motorcycles.find((m) => m.id === r.motorcycleId);
    const type = db.maintenanceTypes.find((t) => t.id === r.typeId);
    const q = search.toLowerCase();
    return !q || `${moto?.plate} ${type?.name}`.toLowerCase().includes(q);
  }).sort((a, b) => (a.date < b.date ? 1 : -1));
  const total = rows.reduce((s, r) => s + Number(r.totalCost), 0);

  const save = (values) => {
    const type = db.maintenanceTypes.find((t) => t.id === values.typeId);
    const partsCost = Number(values.partsCost) || 0, laborCost = Number(values.laborCost) || 0;
    const km = Number(values.km) || 0;
    const v = { ...values, partsCost, laborCost, totalCost: partsCost + laborCost, km, nextKm: km + (type?.kmInterval || 0), nextDate: addDays(values.date, type?.timeIntervalDays || 0) };
    if (editing) patch((d) => ({ maintenanceRecords: d.maintenanceRecords.map((r) => (r.id === editing.id ? { ...r, ...v } : r)) }));
    else patch((d) => ({ maintenanceRecords: [...d.maintenanceRecords, { id: uid("mnt"), ...v }] }));
    setEditing(null); setCreating(false);
  };
  const remove = () => { patch((d) => ({ maintenanceRecords: d.maintenanceRecords.filter((r) => r.id !== deleting.id) })); setDeleting(null); };

  return (
    <div>
      <div className="mb-4"><KpiCard label="Custo total de manutenção (filtro)" value={fmtBRL(total)} /></div>
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por placa ou tipo..." onNew={() => setCreating(true)} newLabel="Nova manutenção" />
      <DataTable columns={[
        { key: "moto", label: "Moto", render: (r) => db.motorcycles.find((m) => m.id === r.motorcycleId)?.plate || "—" },
        { key: "type", label: "Tipo", render: (r) => db.maintenanceTypes.find((t) => t.id === r.typeId)?.name || "—" },
        { key: "date", label: "Data", render: (r) => fmtDate(r.date) }, { key: "km", label: "KM", render: (r) => fmtNum(r.km) },
        { key: "totalCost", label: "Custo total", render: (r) => fmtBRL(r.totalCost) }, { key: "vendor", label: "Oficina" },
        { key: "next", label: "Próxima", render: (r) => `${fmtNum(r.nextKm)} km / ${fmtDate(r.nextDate)}` },
      ]} rows={rows} onEdit={setEditing} onDelete={setDeleting} emptyText="Nenhuma manutenção registrada." />
      {(editing || creating) && <FormModal title={editing ? "Editar manutenção" : "Nova manutenção"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text="Excluir este registro de manutenção?" onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

function MaintenanceAlertsCenter({ db }) {
  const rows = [];
  db.motorcycles.forEach((m) => {
    maintenanceAlertsForMoto(db, m.id).forEach((a) => rows.push({ moto: m, ...a }));
  });
  rows.sort((a, b) => (a.kmRemaining ?? 1e9) - (b.kmRemaining ?? 1e9));
  return (
    <DataTable columns={[
      { key: "moto", label: "Moto", render: (r) => motoLabel(r.moto) }, { key: "type", label: "Tipo", render: (r) => r.type.name },
      { key: "last", label: "Última execução", render: (r) => r.last ? `${fmtDate(r.last.date)} (${fmtNum(r.last.km)} km)` : "Sem histórico" },
      { key: "nextKm", label: "Próxima em (KM)", render: (r) => fmtNum(r.nextKm) },
      { key: "remaining", label: "KM restantes", render: (r) => r.status === "Sem histórico" ? "—" : fmtNum(r.kmRemaining) },
      { key: "status", label: "Status", render: (r) => <Badge>{r.status}</Badge> },
    ]} rows={rows} pageSize={15} emptyText="Nenhuma moto cadastrada." />
  );
}

function MaintenanceTypesConfig({ db, patch }) {
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const fields = [
    { name: "name", label: "Nome da manutenção", required: true },
    { name: "kmInterval", label: "Intervalo (KM)", type: "number", required: true },
    { name: "timeIntervalDays", label: "Intervalo (dias)", type: "number" },
  ];
  const save = (values) => {
    const v = { ...values, kmInterval: Number(values.kmInterval) || 0, timeIntervalDays: Number(values.timeIntervalDays) || 0 };
    if (editing) patch((d) => ({ maintenanceTypes: d.maintenanceTypes.map((t) => (t.id === editing.id ? { ...t, ...v } : t)) }));
    else patch((d) => ({ maintenanceTypes: [...d.maintenanceTypes, { id: uid("mt"), ...v }] }));
    setEditing(null); setCreating(false);
  };
  return (
    <div>
      <Toolbar search="" setSearch={() => {}} placeholder="" onNew={() => setCreating(true)} newLabel="Novo tipo" right={null} />
      <DataTable columns={[
        { key: "name", label: "Tipo" }, { key: "kmInterval", label: "Intervalo (KM)", render: (t) => fmtNum(t.kmInterval) },
        { key: "timeIntervalDays", label: "Intervalo (dias)", render: (t) => fmtNum(t.timeIntervalDays) },
      ]} rows={db.maintenanceTypes} onEdit={setEditing} onDelete={(t) => patch((d) => ({ maintenanceTypes: d.maintenanceTypes.filter((x) => x.id !== t.id) }))} emptyText="Nenhum tipo cadastrado." />
      {(editing || creating) && <FormModal title={editing ? "Editar tipo" : "Novo tipo de manutenção"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} />}
    </div>
  );
}

/* =========================================================================
   QUILOMETRAGEM
========================================================================= */
function MileageModule({ db, patch }) {
  const [motoId, setMotoId] = useState(db.motorcycles[0]?.id || "");
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);
  const moto = db.motorcycles.find((m) => m.id === motoId);
  const records = db.mileageRecords.filter((r) => r.motorcycleId === motoId).sort((a, b) => (a.date < b.date ? -1 : 1));
  const today = todayISO(), weekAgo = addDays(today, -7), monthAgo = addDays(today, -30), sinceLastServiceKm = (() => {
    const last = db.maintenanceRecords.filter((r) => r.motorcycleId === motoId).sort((a, b) => (a.date < b.date ? 1 : -1))[0];
    return last ? (moto?.currentKm || 0) - last.km : null;
  })();
  const kmWeek = records.filter((r) => r.date >= weekAgo).reduce((s, r) => s + Number(r.kmDriven || 0), 0);
  const kmMonth = records.filter((r) => r.date >= monthAgo).reduce((s, r) => s + Number(r.kmDriven || 0), 0);
  const daysSinceAcq = moto ? Math.max(1, daysDiff(today, moto.acquisitionDate)) : 1;
  const avgDay = moto ? (moto.currentKm - moto.acquisitionKm) / daysSinceAcq : 0;

  const fields = [
    { name: "date", label: "Data", type: "date", required: true },
    { name: "km", label: "Quilometragem registrada", type: "number", required: true },
    { name: "recordedBy", label: "Registrado por", default: "Administrador" },
    { name: "notes", label: "Observação", type: "textarea", span: true },
  ];

  const save = (values) => {
    const km = Number(values.km) || 0;
    const prevRecords = db.mileageRecords.filter((r) => r.motorcycleId === motoId).sort((a, b) => (a.date < b.date ? 1 : -1));
    const previousKm = prevRecords[0]?.km ?? moto.acquisitionKm;
    if (km < previousKm) { alert("A quilometragem informada é menor que o último registro. Verifique o valor."); return; }
    patch((d) => ({
      mileageRecords: [...d.mileageRecords, { id: uid("km"), motorcycleId: motoId, date: values.date, km, previousKm, kmDriven: km - previousKm, recordedBy: values.recordedBy, notes: values.notes }],
      motorcycles: d.motorcycles.map((m) => (m.id === motoId ? { ...m, currentKm: Math.max(m.currentKm, km) } : m)),
    }));
    setCreating(false);
  };
  const remove = () => { patch((d) => ({ mileageRecords: d.mileageRecords.filter((r) => r.id !== deleting.id) })); setDeleting(null); };

  return (
    <div>
      <SectionHeader title="Controle de quilometragem" desc="Selecione uma moto para ver e lançar registros" />
      <div className="mb-4"><select className={`${inputCls} max-w-sm`} value={motoId} onChange={(e) => setMotoId(e.target.value)}>{db.motorcycles.map((m) => <option key={m.id} value={m.id}>{motoLabel(m)}</option>)}</select></div>
      {moto && (
        <>
          <div className="mb-4 grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
            <KpiCard label="KM atual" value={fmtNum(moto.currentKm)} />
            <KpiCard label="KM na semana" value={fmtNum(kmWeek)} />
            <KpiCard label="KM no mês" value={fmtNum(kmMonth)} />
            <KpiCard label="Desde último serviço" value={sinceLastServiceKm !== null ? fmtNum(sinceLastServiceKm) : "—"} />
            <KpiCard label="Média por dia" value={fmtNum(avgDay, 1)} />
            <KpiCard label="Média por semana" value={fmtNum(avgDay * 7, 1)} />
          </div>
          <div className="mb-4 rounded-xl border border-slate-200 bg-white p-4">
            <h4 className="mb-2 text-sm font-semibold text-slate-700">Evolução da quilometragem</h4>
            <SvgLineChart points={records.map((r) => ({ label: fmtDate(r.date).slice(0, 5), value: r.km }))} />
          </div>
          <Toolbar search="" setSearch={() => {}} placeholder="" onNew={() => setCreating(true)} newLabel="Lançar quilometragem" right={null} />
          <DataTable columns={[
            { key: "date", label: "Data", render: (r) => fmtDate(r.date) }, { key: "previousKm", label: "KM anterior", render: (r) => fmtNum(r.previousKm) },
            { key: "km", label: "KM registrado", render: (r) => fmtNum(r.km) }, { key: "kmDriven", label: "KM rodados", render: (r) => fmtNum(r.kmDriven) },
            { key: "recordedBy", label: "Registrado por" }, { key: "notes", label: "Observação" },
          ]} rows={[...records].reverse()} onDelete={setDeleting} emptyText="Nenhum registro de quilometragem." />
        </>
      )}
      {creating && <FormModal title="Lançar quilometragem" fields={fields} initial={{ date: todayISO(), recordedBy: "Administrador" }} onCancel={() => setCreating(false)} onSave={save} />}
      {deleting && <ConfirmDialog text="Excluir este registro de quilometragem?" onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

/* =========================================================================
   OCORRÊNCIAS
========================================================================= */
function IncidentsModule({ db, patch }) {
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState(null);

  const fields = [
    { name: "motorcycleId", label: "Moto", type: "select", required: true, options: () => db.motorcycles.map((m) => ({ value: m.id, label: motoLabel(m) })) },
    { name: "renterId", label: "Locatário (opcional)", type: "select", options: () => db.renters.map((r) => ({ value: r.id, label: r.name })) },
    { name: "date", label: "Data", type: "date", required: true }, { name: "type", label: "Tipo", type: "select", options: INCIDENT_TYPES, required: true },
    { name: "value", label: "Valor (R$)", type: "number" }, { name: "responsible", label: "Responsável pelo pagamento", type: "select", options: INCIDENT_RESP },
    { name: "status", label: "Status", type: "select", options: INCIDENT_STATUS, required: true },
    { name: "description", label: "Descrição", type: "textarea", span: true },
  ];

  const rows = db.incidents.filter((i) => {
    const moto = db.motorcycles.find((m) => m.id === i.motorcycleId);
    const q = search.toLowerCase();
    return !q || `${i.type} ${moto?.plate}`.toLowerCase().includes(q);
  }).sort((a, b) => (a.date < b.date ? 1 : -1));

  const save = (values) => {
    const v = { ...values, value: Number(values.value) || 0 };
    if (editing) patch((d) => ({ incidents: d.incidents.map((i) => (i.id === editing.id ? { ...i, ...v } : i)) }));
    else patch((d) => ({ incidents: [...d.incidents, { id: uid("inc"), ...v }] }));
    setEditing(null); setCreating(false);
  };
  const remove = () => { patch((d) => ({ incidents: d.incidents.filter((i) => i.id !== deleting.id) })); setDeleting(null); };

  return (
    <div>
      <SectionHeader title="Multas e ocorrências" desc="Acidentes, danos, avarias, furto/roubo e outras ocorrências" />
      <Toolbar search={search} setSearch={setSearch} placeholder="Buscar por tipo ou placa..." onNew={() => setCreating(true)} newLabel="Nova ocorrência" />
      <DataTable columns={[
        { key: "date", label: "Data", render: (i) => fmtDate(i.date) },
        { key: "moto", label: "Moto", render: (i) => db.motorcycles.find((m) => m.id === i.motorcycleId)?.plate || "—" },
        { key: "renter", label: "Locatário", render: (i) => db.renters.find((r) => r.id === i.renterId)?.name || "—" },
        { key: "type", label: "Tipo" }, { key: "value", label: "Valor", render: (i) => fmtBRL(i.value) },
        { key: "responsible", label: "Responsável" }, { key: "status", label: "Status", render: (i) => <Badge>{i.status}</Badge> },
      ]} rows={rows} onEdit={setEditing} onDelete={setDeleting} emptyText="Nenhuma ocorrência registrada." />
      {(editing || creating) && <FormModal title={editing ? "Editar ocorrência" : "Nova ocorrência"} fields={fields} initial={editing} onCancel={() => { setEditing(null); setCreating(false); }} onSave={save} wide />}
      {deleting && <ConfirmDialog text="Excluir esta ocorrência?" onCancel={() => setDeleting(null)} onConfirm={remove} />}
    </div>
  );
}

/* =========================================================================
   CALENDÁRIO
========================================================================= */
function CalendarModule({ db }) {
  const [monthOffset, setMonthOffset] = useState(0);
  const now = new Date();
  const base = new Date(now.getFullYear(), now.getMonth() + monthOffset, 1);
  const year = base.getFullYear(), month = base.getMonth();
  const firstDay = new Date(year, month, 1);
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startWeekday = (firstDay.getDay() + 6) % 7; // Monday=0

  const events = {};
  const push = (dateISO, label, tone) => { if (!dateISO) return; (events[dateISO] = events[dateISO] || []).push({ label, tone }); };
  db.contracts.forEach((c) => { if (c.status === "Ativo") push(c.endDate, `Fim de contrato — ${db.renters.find((r) => r.id === c.renterId)?.name || ""}`, "blue"); });
  db.payments.forEach((p) => { if (paymentStatus(p, todayISO()) !== "Pago") push(p.dueDate, `Pagamento — ${db.renters.find((r) => r.id === p.renterId)?.name || ""}`, "amber"); });
  db.motorcycles.forEach((m) => { push(m.insuranceDue, `Seguro — ${m.plate}`, "red"); push(m.nextLicensingDue, `Licenciamento — ${m.plate}`, "red"); });
  db.renters.forEach((r) => push(r.cnhValidity, `CNH — ${r.name}`, "red"));
  db.motorcycles.forEach((m) => { maintenanceAlertsForMoto(db, m.id).forEach((a) => { if (a.last && a.status !== "Em dia") push(a.nextDate, `${a.type.name} — ${m.plate}`, "amber"); }); });

  const cells = [];
  for (let i = 0; i < startWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  return (
    <div>
      <SectionHeader title="Calendário" desc="Vencimentos, contratos, manutenções e pagamentos" />
      <div className="mb-4 flex items-center justify-between rounded-xl border border-slate-200 bg-white p-3">
        <button onClick={() => setMonthOffset((o) => o - 1)} className="rounded-lg p-1.5 hover:bg-slate-100"><ChevronLeft size={18} /></button>
        <span className="text-sm font-semibold text-slate-800" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{base.toLocaleDateString("pt-BR", { month: "long", year: "numeric" })}</span>
        <button onClick={() => setMonthOffset((o) => o + 1)} className="rounded-lg p-1.5 hover:bg-slate-100"><ChevronRight size={18} /></button>
      </div>
      <div className="grid grid-cols-7 gap-1.5 text-center text-xs font-semibold text-slate-400">
        {["SEG", "TER", "QUA", "QUI", "SEX", "SÁB", "DOM"].map((d) => <div key={d} className="py-1">{d}</div>)}
      </div>
      <div className="grid grid-cols-7 gap-1.5">
        {cells.map((d, i) => {
          const iso = d ? `${year}-${String(month + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}` : null;
          const evs = iso ? events[iso] || [] : [];
          const isToday = iso === todayISO();
          return (
            <div key={i} className={`min-h-[84px] rounded-lg border p-1.5 text-left ${d ? "bg-white border-slate-200" : "border-transparent"} ${isToday ? "ring-2 ring-orange-400" : ""}`}>
              {d && <span className="text-xs font-medium text-slate-500">{d}</span>}
              <div className="mt-1 space-y-0.5">
                {evs.slice(0, 3).map((e, j) => (
                  <div key={j} className={`truncate rounded px-1 py-0.5 text-[10px] font-medium ${e.tone === "red" ? "bg-red-50 text-red-700" : e.tone === "blue" ? "bg-blue-50 text-blue-700" : "bg-amber-50 text-amber-700"}`}>{e.label}</div>
                ))}
                {evs.length > 3 && <div className="text-[10px] text-slate-400">+{evs.length - 3} mais</div>}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* =========================================================================
   RELATÓRIOS
========================================================================= */
function ReportsModule({ db }) {
  const [report, setReport] = useState("Rentabilidade por moto");
  const reports = ["Rentabilidade por moto", "Rentabilidade por locatário", "Receita x Despesa", "Manutenção por moto", "Quilometragem", "Inadimplência", "Ocupação da frota"];
  const today = todayISO(), monthStart = startOfMonth(today);

  return (
    <div>
      <SectionHeader title="Relatórios" desc="Análises detalhadas para tomada de decisão" />
      <Tabs tabs={reports} active={report} onChange={setReport} />

      {report === "Rentabilidade por moto" && <ProfitabilityByMoto db={db} />}
      {report === "Rentabilidade por locatário" && <ProfitabilityByRenter db={db} />}
      {report === "Receita x Despesa" && (
        <div className="rounded-xl border border-slate-200 bg-white p-4">
          <MonthlyRevenueExpenseChart db={db} />
        </div>
      )}
      {report === "Manutenção por moto" && <MaintenanceCostReport db={db} />}
      {report === "Quilometragem" && <MileageReport db={db} />}
      {report === "Inadimplência" && <DefaultReport db={db} />}
      {report === "Ocupação da frota" && <OccupancyReport db={db} />}
    </div>
  );
}

function ExportBar({ columns, rows, filename }) {
  return (
    <div className="mb-3 flex justify-end gap-2">
      <button onClick={() => exportCSV(`${filename}.csv`, columns, rows)} className="flex items-center gap-1.5 rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"><Download size={13} /> CSV</button>
      <button onClick={() => exportXLSX(`${filename}.xlsx`, columns, rows)} className="flex items-center gap-1.5 rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"><Download size={13} /> Excel</button>
      <button onClick={() => window.print()} className="flex items-center gap-1.5 rounded-lg border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"><Download size={13} /> PDF (imprimir)</button>
    </div>
  );
}

function ProfitabilityByMoto({ db }) {
  const data = db.motorcycles.map((m) => ({ moto: m, ...computeMotoFinancials(db, m.id, "0000-01-01", todayISO()) })).sort((a, b) => b.profit - a.profit);
  const columns = [
    { label: "Moto", value: (r) => motoLabel(r.moto) }, { label: "Receita aluguel", value: (r) => r.rentPaid }, { label: "Outras receitas", value: (r) => r.rev },
    { label: "Manutenção", value: (r) => r.maint }, { label: "Outras despesas", value: (r) => r.exp }, { label: "Lucro", value: (r) => r.profit },
  ];
  return (
    <div>
      <ExportBar columns={columns} rows={data} filename="rentabilidade_por_moto" />
      <div className="mb-4 grid grid-cols-1 gap-3 sm:grid-cols-3">
        {data.slice(0, 3).map((d, i) => (
          <div key={d.moto.id} className="rounded-xl border border-slate-200 bg-white p-4">
            <p className="text-xs font-semibold text-orange-600">{i + 1}º MAIS LUCRATIVA</p>
            <p className="mt-1 text-sm font-medium text-slate-800">{motoLabel(d.moto)}</p>
            <p className="text-lg font-semibold text-emerald-600">{fmtBRL(d.profit)}</p>
          </div>
        ))}
      </div>
      <DataTable columns={[
        { key: "moto", label: "Moto", render: (r) => motoLabel(r.moto) }, { key: "revenueTotal", label: "Receita total", render: (r) => fmtBRL(r.revenueTotal) },
        { key: "costTotal", label: "Custos totais", render: (r) => fmtBRL(r.costTotal) }, { key: "profit", label: "Lucro", render: (r) => <span className={r.profit >= 0 ? "text-emerald-600 font-medium" : "text-red-600 font-medium"}>{fmtBRL(r.profit)}</span> },
      ]} rows={data} pageSize={10} />
      {data.some((d) => d.profit < 0) && (
        <div className="mt-4 rounded-xl border border-red-200 bg-red-50 p-4">
          <p className="text-sm font-semibold text-red-700">Motos com prejuízo</p>
          <ul className="mt-1 list-inside list-disc text-sm text-red-600">{data.filter((d) => d.profit < 0).map((d) => <li key={d.moto.id}>{motoLabel(d.moto)} — {fmtBRL(d.profit)}</li>)}</ul>
        </div>
      )}
    </div>
  );
}

function ProfitabilityByRenter({ db }) {
  const data = db.renters.map((r) => ({ renter: r, ...computeRenterStats(db, r.id) })).sort((a, b) => b.revenue - a.revenue);
  const columns = [
    { label: "Locatário", value: (r) => r.renter.name }, { label: "Receita gerada", value: (r) => r.revenue },
    { label: "Semanas alugadas", value: (r) => r.weeks }, { label: "Atrasos", value: (r) => r.late }, { label: "Dias com moto", value: (r) => r.days },
  ];
  return (
    <div>
      <ExportBar columns={columns} rows={data} filename="rentabilidade_por_locatario" />
      <DataTable columns={[
        { key: "renter", label: "Locatário", render: (r) => r.renter.name }, { key: "revenue", label: "Receita gerada", render: (r) => fmtBRL(r.revenue) },
        { key: "weeks", label: "Semanas alugadas" }, { key: "late", label: "Atrasos", render: (r) => <span className={r.late ? "text-red-600 font-medium" : ""}>{r.late}</span> },
        { key: "days", label: "Dias com moto" }, { key: "contracts", label: "Contratos" },
      ]} rows={data} pageSize={10} />
    </div>
  );
}

function MaintenanceCostReport({ db }) {
  const data = db.motorcycles.map((m) => {
    const recs = db.maintenanceRecords.filter((r) => r.motorcycleId === m.id);
    const total = recs.reduce((s, r) => s + Number(r.totalCost), 0);
    const kmRodado = m.currentKm - m.acquisitionKm;
    return { moto: m, count: recs.length, total, costPerKm: kmRodado ? total / kmRodado : 0 };
  }).sort((a, b) => b.total - a.total);
  const columns = [{ label: "Moto", value: (r) => motoLabel(r.moto) }, { label: "Nº manutenções", value: (r) => r.count }, { label: "Custo total", value: (r) => r.total }, { label: "Custo por KM", value: (r) => r.costPerKm.toFixed(2) }];
  return (
    <div>
      <ExportBar columns={columns} rows={data} filename="manutencao_por_moto" />
      <DataTable columns={[
        { key: "moto", label: "Moto", render: (r) => motoLabel(r.moto) }, { key: "count", label: "Nº manutenções" },
        { key: "total", label: "Custo total", render: (r) => fmtBRL(r.total) }, { key: "costPerKm", label: "Custo por KM", render: (r) => fmtBRL(r.costPerKm) },
      ]} rows={data} pageSize={10} />
    </div>
  );
}

function MileageReport({ db }) {
  const data = db.motorcycles.map((m) => {
    const recs = db.mileageRecords.filter((r) => r.motorcycleId === m.id);
    const total = m.currentKm - m.acquisitionKm;
    const days = Math.max(1, daysDiff(todayISO(), m.acquisitionDate));
    return { moto: m, total, avgDay: total / days, records: recs.length };
  });
  const columns = [{ label: "Moto", value: (r) => motoLabel(r.moto) }, { label: "KM total rodado", value: (r) => r.total }, { label: "Média diária", value: (r) => r.avgDay.toFixed(1) }];
  return (
    <div>
      <ExportBar columns={columns} rows={data} filename="km_rodados" />
      <DataTable columns={[
        { key: "moto", label: "Moto", render: (r) => motoLabel(r.moto) }, { key: "total", label: "KM total rodado", render: (r) => fmtNum(r.total) },
        { key: "avgDay", label: "Média diária", render: (r) => fmtNum(r.avgDay, 1) }, { key: "records", label: "Registros" },
      ]} rows={data} pageSize={10} />
    </div>
  );
}

function DefaultReport({ db }) {
  const today = todayISO();
  const late = db.payments.filter((p) => paymentStatus(p, today) === "Atrasado").map((p) => ({ ...p, renter: db.renters.find((r) => r.id === p.renterId)?.name, moto: db.motorcycles.find((m) => m.id === p.motorcycleId)?.plate, daysLate: daysDiff(today, p.dueDate) }));
  const totalExpected = db.payments.reduce((s, p) => s + Number(p.expectedValue), 0);
  const totalLate = late.reduce((s, p) => s + (Number(p.expectedValue) - Number(p.paidValue || 0)), 0);
  const columns = [{ label: "Locatário", value: (r) => r.renter }, { label: "Moto", value: (r) => r.moto }, { label: "Vencimento", value: (r) => fmtDate(r.dueDate) }, { label: "Valor", value: (r) => r.expectedValue }, { label: "Dias de atraso", value: (r) => r.daysLate }];
  return (
    <div>
      <div className="mb-4"><KpiCard label="Taxa de inadimplência geral" value={`${fmtNum(totalExpected ? (totalLate / totalExpected) * 100 : 0, 1)}%`} tone={totalLate ? "red" : "slate"} /></div>
      <ExportBar columns={columns} rows={late} filename="inadimplencia" />
      <DataTable columns={[
        { key: "renter", label: "Locatário" }, { key: "moto", label: "Moto" }, { key: "dueDate", label: "Vencimento", render: (p) => fmtDate(p.dueDate) },
        { key: "expectedValue", label: "Valor", render: (p) => fmtBRL(p.expectedValue) }, { key: "daysLate", label: "Dias de atraso", render: (p) => <span className="font-medium text-red-600">{p.daysLate}</span> },
      ]} rows={late} pageSize={10} emptyText="Nenhum pagamento em atraso." />
    </div>
  );
}

function OccupancyReport({ db }) {
  const total = db.motorcycles.filter((m) => m.status !== "Vendida").length;
  const alugadas = db.motorcycles.filter((m) => m.status === "Alugada").length;
  const pct = total ? (alugadas / total) * 100 : 0;
  return (
    <div className="rounded-xl border border-slate-200 bg-white p-6">
      <p className="text-sm text-slate-500">Taxa de ocupação da frota</p>
      <p className="mt-1 text-4xl font-semibold text-slate-900" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{fmtNum(pct, 1)}%</p>
      <div className="mt-3 max-w-md"><MiniBar pct={pct} tone="orange" /></div>
      <p className="mt-2 text-xs text-slate-400">{alugadas} de {total} motos disponíveis para locação estão alugadas atualmente.</p>
      <div className="mt-5 grid grid-cols-2 gap-3 sm:grid-cols-4">
        {MOTO_STATUS.map((s) => <KpiCard key={s} label={s} value={db.motorcycles.filter((m) => m.status === s).length} />)}
      </div>
    </div>
  );
}

/* =========================================================================
   CONFIGURAÇÕES
========================================================================= */
function SettingsModule({ db, patch, currentUser, mode }) {
  const [form, setForm] = useState(db.settings);
  const [newExpenseCat, setNewExpenseCat] = useState("");
  const [newRevenueCat, setNewRevenueCat] = useState("");
  const [confirmReset, setConfirmReset] = useState(false);
  const [editingUser, setEditingUser] = useState(null);
  const [creatingUser, setCreatingUser] = useState(false);
  const [deletingUser, setDeletingUser] = useState(null);

  const userFields = [
    { name: "name", label: "Nome completo", required: true },
    { name: "username", label: "Usuário (login)", required: true },
    { name: "password", label: "Senha", type: "password", hint: editingUser ? "Deixe em branco para manter a senha atual." : "Mínimo de 8 caracteres." },
    { name: "role", label: "Nível de acesso", type: "select", options: USER_ROLES, required: true },
  ];
  const saveUser = (values) => {
    const uname = values.username.trim();
    const dup = db.users.some((u) => u.username.toLowerCase() === uname.toLowerCase() && u.id !== editingUser?.id);
    if (dup) { alert("Já existe um usuário com esse nome de usuário."); return; }
    if (editingUser) {
      patch((d) => ({ users: d.users.map((u) => u.id === editingUser.id ? { ...u, name: values.name, username: uname, role: values.role, password: values.password?.trim() ? values.password : u.password } : u) }));
    } else {
      if (!values.password || values.password.length < 8) { alert("A senha deve ter pelo menos 8 caracteres."); return; }
      patch((d) => ({ users: [...d.users, { id: uid("user"), name: values.name, username: uname, password: values.password, role: values.role }] }));
    }
    setEditingUser(null); setCreatingUser(false);
  };
  const removeUser = () => {
    if (db.users.length <= 1) { alert("Não é possível excluir o único usuário do sistema."); setDeletingUser(null); return; }
    if (deletingUser.id === currentUser?.id) { alert("Você não pode excluir o usuário com o qual está logado."); setDeletingUser(null); return; }
    patch((d) => ({ users: d.users.filter((u) => u.id !== deletingUser.id) }));
    setDeletingUser(null);
  };

  const save = () => patch((d) => ({ settings: { ...d.settings, ...form, companyName: form.companyName, defaultRentalWeekly: Number(form.defaultRentalWeekly) || 0, alertKmWarning: Number(form.alertKmWarning) || 0, alertDaysWarning: Number(form.alertDaysWarning) || 0 } }));

  const addExpenseCat = () => {
    if (!newExpenseCat.trim()) return;
    const cats = [...(db.settings.expenseCategories || []), newExpenseCat.trim()];
    patch((d) => ({ settings: { ...d.settings, expenseCategories: cats } }));
    setNewExpenseCat("");
  };
  const addRevenueCat = () => {
    if (!newRevenueCat.trim()) return;
    const cats = [...(db.settings.revenueCategories || []), newRevenueCat.trim()];
    patch((d) => ({ settings: { ...d.settings, revenueCategories: cats } }));
    setNewRevenueCat("");
  };
  const removeCat = (list, val) => patch((d) => ({ settings: { ...d.settings, [list]: d.settings[list].filter((c) => c !== val) } }));

  const resetExamples = () => {
    patch((d) => ({ ...emptyDb(), settings: d.settings, users: d.users }));
    setConfirmReset(false);
  };

  return (
    <div className="max-w-3xl space-y-6">
      <SectionHeader title="Configurações" desc="Dados da empresa, regras de alerta e categorias" />

      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <h3 className="mb-3 text-sm font-semibold text-slate-800">Dados da empresa</h3>
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
          <Field label="Nome da empresa"><input className={inputCls} value={form.companyName} onChange={(e) => setForm({ ...form, companyName: e.target.value })} /></Field>
          <Field label="Telefone"><input className={inputCls} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label="Endereço" span><input className={inputCls} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></Field>
          <Field label="Valor padrão de aluguel semanal (R$)"><input className={inputCls} type="number" value={form.defaultRentalWeekly} onChange={(e) => setForm({ ...form, defaultRentalWeekly: e.target.value })} /></Field>
          <Field label="Moeda"><input className={inputCls} value="BRL (Real brasileiro)" disabled /></Field>
        </div>
      </div>

      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <h3 className="mb-3 text-sm font-semibold text-slate-800">Regras de alerta</h3>
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
          <Field label="Alertar manutenção com quantos KM de antecedência"><input className={inputCls} type="number" value={form.alertKmWarning} onChange={(e) => setForm({ ...form, alertKmWarning: e.target.value })} /></Field>
          <Field label="Alertar vencimentos com quantos dias de antecedência"><input className={inputCls} type="number" value={form.alertDaysWarning} onChange={(e) => setForm({ ...form, alertDaysWarning: e.target.value })} /></Field>
        </div>
      </div>

      <div className="flex justify-end"><button onClick={save} className="rounded-lg bg-orange-600 px-4 py-2 text-sm font-medium text-white hover:bg-orange-700">Salvar configurações</button></div>

      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <h3 className="mb-3 text-sm font-semibold text-slate-800">Categorias de despesas</h3>
        <div className="mb-3 flex flex-wrap gap-2">{(db.settings.expenseCategories || []).map((c) => (
          <span key={c} className="flex items-center gap-1.5 rounded-full bg-slate-100 px-3 py-1 text-xs font-medium text-slate-700">{c}<button onClick={() => removeCat("expenseCategories", c)}><X size={11} /></button></span>
        ))}</div>
        <div className="flex gap-2"><input className={`${inputCls} max-w-xs`} value={newExpenseCat} onChange={(e) => setNewExpenseCat(e.target.value)} placeholder="Nova categoria" /><button onClick={addExpenseCat} className="rounded-lg border border-slate-300 px-3 py-2 text-sm font-medium hover:bg-slate-50">Adicionar</button></div>
      </div>

      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <h3 className="mb-3 text-sm font-semibold text-slate-800">Categorias de receitas</h3>
        <div className="mb-3 flex flex-wrap gap-2">{(db.settings.revenueCategories || []).map((c) => (
          <span key={c} className="flex items-center gap-1.5 rounded-full bg-slate-100 px-3 py-1 text-xs font-medium text-slate-700">{c}<button onClick={() => removeCat("revenueCategories", c)}><X size={11} /></button></span>
        ))}</div>
        <div className="flex gap-2"><input className={`${inputCls} max-w-xs`} value={newRevenueCat} onChange={(e) => setNewRevenueCat(e.target.value)} placeholder="Nova categoria" /><button onClick={addRevenueCat} className="rounded-lg border border-slate-300 px-3 py-2 text-sm font-medium hover:bg-slate-50">Adicionar</button></div>
      </div>

      {currentUser?.role === "Administrador" && (
      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <div className="mb-3 flex items-center justify-between">
          <h3 className="text-sm font-semibold text-slate-800">Usuários do sistema</h3>
          <button onClick={() => setCreatingUser(true)} className="flex items-center gap-1.5 rounded-lg bg-orange-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-orange-700"><Plus size={13} /> Novo usuário</button>
        </div>
        <DataTable
          columns={[
            { key: "name", label: "Nome", render: (u) => <span className="font-medium text-slate-800">{u.name}{u.id === currentUser?.id ? " (você)" : ""}</span> },
            { key: "username", label: "Usuário", mono: true },
            { key: "role", label: "Nível", render: (u) => <Badge>{u.role}</Badge> },
          ]}
          rows={db.users}
          onEdit={(u) => setEditingUser(u)}
          onDelete={(u) => setDeletingUser(u)}
          emptyText="Nenhum usuário cadastrado."
        />
        {(creatingUser || editingUser) && (
          <FormModal
            title={editingUser ? "Editar usuário" : "Novo usuário"}
            fields={userFields}
            initial={editingUser}
            onCancel={() => { setEditingUser(null); setCreatingUser(false); }}
            onSave={saveUser}
          />
        )}
        {deletingUser && <ConfirmDialog text={`Excluir o usuário ${deletingUser.name} (${deletingUser.username})?`} onCancel={() => setDeletingUser(null)} onConfirm={removeUser} />}
      </div>
      )}

      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <h3 className="mb-1 text-sm font-semibold text-slate-800">Conexão</h3>
        <p className="text-sm text-slate-500">{mode === "api" ? "Conectado a um servidor real — os dados abaixo vêm do banco de dados, não deste navegador." : "Modo local — os dados estão salvos apenas neste navegador (sem servidor conectado)."}</p>
      </div>

      {mode !== "api" && (
      <div className="rounded-xl border border-amber-200 bg-amber-50 p-5">
        <h3 className="mb-1 text-sm font-semibold text-amber-800">Dados de exemplo</h3>
        <p className="mb-3 text-sm text-amber-700">Este sistema foi iniciado com motos, locatários, contratos e lançamentos fictícios para demonstração. Você pode apagar todos os dados de exemplo quando quiser começar a usar o sistema com informações reais.</p>
        <button onClick={() => setConfirmReset(true)} className="rounded-lg border border-amber-400 bg-white px-4 py-2 text-sm font-medium text-amber-800 hover:bg-amber-100">Apagar dados de exemplo</button>
      </div>
      )}

      {confirmReset && <ConfirmDialog text="Isso apagará TODOS os dados atuais (motos, locatários, contratos, pagamentos, receitas, despesas, manutenções e ocorrências), mantendo apenas as configurações. Deseja continuar?" onCancel={() => setConfirmReset(false)} onConfirm={resetExamples} />}
    </div>
  );
}
