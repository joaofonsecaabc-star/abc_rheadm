import { useEffect, useMemo, useRef, useState } from "react";
import type { FormEvent, ReactNode } from "react";
import { createPortal } from "react-dom";
import {
  ArrowLeft,
  Banknote,
  BadgeCheck,
  Bell,
  CalendarDays,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CircleHelp,
  Clock3,
  CreditCard,
  Download,
  DollarSign,
  Eye,
  Gift,
  HandCoins,
  Hourglass,
  FileSpreadsheet,
  FileText,
  LayoutDashboard,
  Menu,
  MoreHorizontal,
  Pencil,
  Plus,
  ReceiptText,
  Search,
  Settings,
  Sparkles,
  Store,
  Ticket,
  TriangleAlert,
  Umbrella,
  UserRound,
  Users,
  X,
} from "lucide-react";
import {
  formatDate,
  initialRecharges,
  type Recharge,
  type Status,
} from "./data";
import HRReports from "./HRReports";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { loadDatabase, monthlyRows, type RechargeEvent } from "./domain";
import {
  cloudCreateAdmin,
  cloudCurrentUser,
  cloudEnabled,
  cloudLogin,
  cloudLogout,
  cloudSetupRequired,
  loadCloudState,
  saveCloudState,
  type ModuleAccess,
  type SessionUser,
} from "./cloud";
import UserProfiles from "./UserProfiles";

function ActionToast() {
  const [message, setMessage] = useState("");
  useEffect(() => {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const notify = (event: Event) => {
        const text = (event as CustomEvent<string>).detail;
        if (!text) return;
        setMessage(text);
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => setMessage(""), 2800);
      };
    window.addEventListener("abc:toast", notify);
    return () => {
      window.removeEventListener("abc:toast", notify);
      if (timer) clearTimeout(timer);
    };
  }, []);
  return (
    <div
      aria-live="polite"
      className={`pointer-events-none fixed bottom-6 right-6 z-[250] flex items-center gap-3 rounded-xl bg-slate-900 px-5 py-3 text-sm font-bold text-white shadow-2xl transition-all duration-200 ${message ? "translate-y-0 opacity-100" : "translate-y-3 opacity-0"}`}
    >
      <span className="grid h-6 w-6 place-items-center rounded-full bg-emerald-500">
        <Check size={15} strokeWidth={3} />
      </span>
      {message || "Ação concluída"}
    </div>
  );
}

function SearchableEmployeeSelect({
  employees,
  value,
  onChange,
  placeholder = "Pesquisar funcionário...",
  required = false,
  className = "",
}: {
  employees: Recharge[];
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  className?: string;
}) {
  const rootRef = useRef<HTMLDivElement>(null);
  const preserveTypedQueryRef = useRef(false);
  const selectedEmployee = employees.find((employee) => String(employee.id) === value);
  const [query, setQuery] = useState(selectedEmployee?.employee || "");
  const [open, setOpen] = useState(false);

  useEffect(() => {
    if (preserveTypedQueryRef.current) {
      preserveTypedQueryRef.current = false;
      return;
    }
    setQuery(selectedEmployee?.employee || "");
  }, [value, selectedEmployee?.employee]);

  useEffect(() => {
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, []);

  const normalizedQuery = query.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  const matches = normalizedQuery
    ? employees.filter((employee) => employee.employee
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .includes(normalizedQuery))
    : [];

  return (
    <div ref={rootRef} className={`relative ${className}`}>
      <Search className="pointer-events-none absolute left-4 top-1/2 z-10 -translate-y-1/2 text-slate-400" size={18} />
      <input
        required={required}
        value={query}
        onFocus={(event) => {
          setOpen(true);
          event.currentTarget.select();
        }}
        onChange={(event) => {
          setQuery(event.target.value);
          preserveTypedQueryRef.current = true;
          onChange("");
          setOpen(true);
        }}
        placeholder={placeholder}
        autoComplete="off"
        className="h-12 w-full rounded-xl border border-slate-300 bg-white py-2 pl-11 pr-10 font-semibold text-slate-900 shadow-sm outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200"
      />
      {query && (
        <button
          type="button"
          aria-label="Limpar funcionário"
          onClick={() => { setQuery(""); preserveTypedQueryRef.current = true; onChange(""); setOpen(true); }}
          className="absolute right-3 top-1/2 -translate-y-1/2 rounded-md p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
        >
          <X size={16} />
        </button>
      )}
      {open && normalizedQuery && (
        <div className="absolute left-0 right-0 top-[calc(100%+6px)] z-[120] max-h-64 overflow-y-auto rounded-xl border border-slate-200 bg-white p-1.5 shadow-2xl">
          {matches.length ? matches.slice(0, 30).map((employee) => (
            <button
              type="button"
              key={employee.id}
              onClick={() => {
                onChange(String(employee.id));
                setQuery(employee.employee);
                setOpen(false);
              }}
              className="block w-full rounded-lg px-3 py-2.5 text-left text-sm font-semibold text-slate-800 hover:bg-slate-100"
            >
              {employee.employee}
            </button>
          )) : (
            <p className="px-3 py-3 text-sm text-slate-500">Nenhum funcionário encontrado.</p>
          )}
        </div>
      )}
    </div>
  );
}

function usePopupBackgroundLock() {
  useEffect(() => {
    const body = document.body,
      html = document.documentElement,
      previousBodyOverflow = body.style.overflow,
      previousHtmlOverflow = html.style.overflow;
    const updateLock = () => {
      const popupOpen = !!document.querySelector(
        ".fixed.inset-0.backdrop-blur-sm",
      );
      body.style.overflow = popupOpen ? "hidden" : previousBodyOverflow;
      html.style.overflow = popupOpen ? "hidden" : previousHtmlOverflow;
    };
    const observer = new MutationObserver(updateLock);
    observer.observe(body, { childList: true, subtree: true });
    updateLock();
    return () => {
      observer.disconnect();
      body.style.overflow = previousBodyOverflow;
      html.style.overflow = previousHtmlOverflow;
    };
  }, []);
}

function FirstAdminScreen({ done }: { done: () => void }) {
  const [fullName, setFullName] = useState(""),
    [username, setUsername] = useState(""),
    [password, setPassword] = useState(""),
    [confirmPassword, setConfirmPassword] = useState(""),
    [error, setError] = useState(""),
    [loading, setLoading] = useState(false);
  return (
    <div className="grid min-h-screen place-items-center bg-gradient-to-br from-forest-950 via-forest-900 to-forest-700 p-4">
      <div className="w-full max-w-md rounded-3xl bg-white p-7 shadow-2xl sm:p-9">
        <div className="text-center">
          <img
            src="/sacolao-abc-logo.png?v=4"
            alt="Sacolão ABC"
            className="mx-auto h-24 w-56 object-contain"
          />
          <h1 className="mt-3 text-2xl font-bold text-slate-900">
            Criar primeiro administrador
          </h1>
          <p className="mt-1 text-sm text-slate-500">
            Este será o perfil principal do sistema.
          </p>
        </div>
        <form
          onSubmit={async (e) => {
            e.preventDefault();
            setError("");
            if (password !== confirmPassword) {
              setError("As senhas não coincidem.");
              return;
            }
            setLoading(true);
            try {
              await cloudCreateAdmin(
                fullName.trim(),
                username.trim(),
                password,
              );
              done();
            } catch (reason) {
              setError(
                reason instanceof Error
                  ? reason.message
                  : "Não foi possível criar o administrador.",
              );
            } finally {
              setLoading(false);
            }
          }}
          className="mt-7 space-y-4"
        >
          <label className="block text-sm font-semibold text-slate-700">
            Nome completo
            <input
              autoFocus
              required
              value={fullName}
              onChange={(e) => setFullName(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3"
            />
          </label>
          <label className="block text-sm font-semibold text-slate-700">
            Usuário
            <input
              required
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3"
            />
          </label>
          <label className="block text-sm font-semibold text-slate-700">
            Senha
            <input
              required
              minLength={8}
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3"
            />
          </label>
          <label className="block text-sm font-semibold text-slate-700">
            Confirmar senha
            <input
              required
              minLength={8}
              type="password"
              value={confirmPassword}
              onChange={(e) => setConfirmPassword(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3"
            />
          </label>
          {error && (
            <p className="rounded-xl bg-red-50 p-3 text-sm font-semibold text-red-600">
              {error}
            </p>
          )}
          <button
            disabled={loading}
            className="w-full rounded-xl bg-slate-800 py-3.5 font-bold text-white disabled:opacity-60"
          >
            {loading ? "Criando..." : "Criar administrador"}
          </button>
        </form>
      </div>
    </div>
  );
}

const statusStyle: Record<Status, string> = {
  Pendente: "bg-amber-50 text-amber-700 border-amber-200",
  Recarregado: "bg-emerald-50 text-emerald-700 border-emerald-200",
  "Recarregado atrasado": "bg-orange-50 text-orange-700 border-orange-200",
  Atrasado: "bg-red-50 text-red-700 border-red-200",
  Próximo: "bg-blue-50 text-blue-700 border-blue-200",
};
type Module = "people" | "transit" | "finance" | "administrative";
type FinancialEntry = {
  id: number;
  employeeId: number;
  period: string;
  salary: number;
  advance: number;
  vacation?: number;
  severance?: number;
  salaryPaidAt?: string;
  advancePaidAt?: string;
  vacationPaidAt?: string;
  severancePaidAt?: string;
  noPayments?: boolean;
  noPaymentsFrom?: string;
};
type TaxEntry = {
  id: number;
  period: string;
  description: string;
  category: string;
  dueDate: string;
  amount: number;
  paid: boolean;
  paidAt?: string;
  note?: string;
};
function DismissedEmployeesPage({
  rows,
  restore,
}: {
  rows: Recharge[];
  restore: (record: Recharge) => void;
}) {
  const dismissed = rows
    .filter(
      (record) =>
        !!record.terminationDate ||
        record.active === false ||
        record.employmentStatus === "Desligado",
    )
    .sort((a, b) =>
      (b.terminationDate || "").localeCompare(a.terminationDate || ""),
    );
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Funcionários desligados"
        sub={`${dismissed.length} funcionário(s) no histórico`}
      />
      <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white shadow-soft">
        <table className="w-full min-w-[760px] text-left text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-400">
            <tr>
              <th className="px-5 py-3">Funcionário</th>
              <th className="px-5 py-3">Loja</th>
              <th className="px-5 py-3">Cargo</th>
              <th className="px-5 py-3">Desligamento</th>
              <th className="px-5 py-3">Situação</th>
              <th className="px-5 py-3">Ação</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {dismissed.map((record) => (
              <tr key={record.id}>
                <td className="px-5 py-4 font-semibold">
                  {record.employee}
                  <div className="text-xs font-normal text-slate-400">
                    {record.cpf || "CPF não informado"}
                  </div>
                </td>
                <td className="px-5 py-4">{record.store}</td>
                <td className="px-5 py-4">{record.role}</td>
                <td className="px-5 py-4">
                  {record.terminationDate
                    ? new Date(
                        record.terminationDate + "T12:00:00",
                      ).toLocaleDateString("pt-BR")
                    : "Não informada"}
                </td>
                <td className="px-5 py-4">
                  {isDismissalPending(record) ? (
                    <span className="rounded-full bg-amber-100 px-2.5 py-1 text-xs font-bold text-amber-700">
                      Desligamento em andamento
                    </span>
                  ) : (
                    <span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs font-bold text-slate-600">
                      Desligado
                    </span>
                  )}
                </td>
                <td className="px-5 py-4">
                  <button
                    onClick={() => {
                      if (
                        confirm(
                          `Reverter o desligamento de ${record.employee}?`,
                        )
                      )
                        restore(record);
                    }}
                    className="rounded-xl border border-slate-800 px-4 py-2 text-xs font-bold text-slate-800"
                  >
                    Reverter desligamento
                  </button>
                </td>
              </tr>
            ))}
            {!dismissed.length && (
              <tr>
                <td colSpan={6} className="py-12 text-center text-slate-400">
                  Nenhum funcionário desligado.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </main>
  );
}
const peopleNav = [
  ["Visão geral", LayoutDashboard],
  ["Ocorrências", TriangleAlert],
  ["Funcionários", Users],
  ["Relatórios", FileSpreadsheet],
] as const;
const transitNav = [
  ["Visão geral", LayoutDashboard],
  ["Funcionários", Users],
  ["Recargas", Ticket],
  ["Calendário", CalendarDays],
  ["Relatórios", FileSpreadsheet],
] as const;
const financeNav = [
  ["Dashboard", LayoutDashboard],
  ["Cadastros", Plus],
  ["Bonificação", Gift],
  ["Impostos", ReceiptText],
  ["Relatórios", FileSpreadsheet],
] as const;
const administrativeNav = [
  ["Salário", Banknote],
  ["Adiantamento", HandCoins],
  ["Advertência", TriangleAlert],
  ["Recibo", ReceiptText],
] as const;
let groupLogoDataPromise: Promise<string> | null = null;
let hlmLogoDataPromise: Promise<string> | null = null;
const loadLogo = (path: string) => fetch(path)
  .then((response) => {
    if (!response.ok) throw new Error("Logo não encontrada");
    return response.blob();
  })
  .then((blob) => new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  }));
const loadGroupLogo = () => {
  if (!groupLogoDataPromise) {
    groupLogoDataPromise = loadLogo("/grupo-abc-vera-cruz.png");
  }
  return groupLogoDataPromise;
};
const loadHlmLogo = () => {
  if (!hlmLogoDataPromise) hlmLogoDataPromise = loadLogo("/hlm-gestao-logo.png");
  return hlmLogoDataPromise;
};
const addCompanyLogo = async (doc: jsPDF, company: string, x: number, y: number, width = 52) => {
  try {
    const isHlm = /\bHLM\b/i.test(company);
    const data = isHlm ? await loadHlmLogo() : await loadGroupLogo();
    doc.addImage(data, "PNG", x, y, width, width / (isHlm ? 1.56 : 1.78));
    return true;
  } catch (error) {
    console.warn("Não foi possível adicionar a logo da empresa ao documento", error);
    return false;
  }
};
const formatCpf = (value: string) =>
  value
    .replace(/\D/g, "")
    .slice(0, 11)
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d{1,2})$/, "$1-$2");
const formatCnpj = (value: string) =>
  value
    .replace(/\D/g, "")
    .slice(0, 14)
    .replace(/(\d{2})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1/$2")
    .replace(/(\d{4})(\d{1,2})$/, "$1-$2");
const genderFromRole = (role: string): Recharge["gender"] | undefined => {
  const normalized = String(role || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  return /aux(?:iliar)?\.?\s*(?:de\s*)?servicos?\s*gerais|frente\s*(?:de\s*)?caixa|operador\(?a?\)?\s*(?:de\s*)?caixa|\bcaixa\b|\bsuporte\b|\batendente\b/.test(normalized)
    ? "Feminino"
    : undefined;
};
const formatMoneyInput = (value: string) => {
  const digits = value.replace(/\D/g, "").slice(0, 10);
  if (!digits) return "";
  return (Number(digits) / 100).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
};
const parseMoney = (value: string) =>
  Number(value.replace(/\./g, "").replace(",", ".")) || 0;
const numberUnderThousandInWords = (value: number) => {
  const units = ["zero", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"];
  const teens = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"];
  const tens = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"];
  const hundreds = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"];
  if (value === 0) return "zero";
  if (value === 100) return "cem";
  const parts: string[] = [];
  const hundred = Math.floor(value / 100), remainder = value % 100;
  if (hundred) parts.push(hundreds[hundred]);
  if (remainder) {
    if (remainder < 10) parts.push(units[remainder]);
    else if (remainder < 20) parts.push(teens[remainder - 10]);
    else {
      const ten = Math.floor(remainder / 10), unit = remainder % 10;
      parts.push(unit ? `${tens[ten]} e ${units[unit]}` : tens[ten]);
    }
  }
  return parts.join(" e ");
};
const integerInWords = (value: number) => {
  if (value === 0) return "zero";
  const groups = [
    { size: 1_000_000_000, singular: "bilhão", plural: "bilhões" },
    { size: 1_000_000, singular: "milhão", plural: "milhões" },
    { size: 1_000, singular: "mil", plural: "mil" },
  ];
  let remaining = Math.floor(value);
  const parts: string[] = [];
  groups.forEach((group) => {
    const count = Math.floor(remaining / group.size);
    if (!count) return;
    if (group.size === 1_000 && count === 1) parts.push("mil");
    else parts.push(`${numberUnderThousandInWords(count)} ${count === 1 ? group.singular : group.plural}`);
    remaining %= group.size;
  });
  if (remaining) parts.push(numberUnderThousandInWords(remaining));
  return parts.length > 1 ? `${parts.slice(0, -1).join(", ")} e ${parts[parts.length - 1]}` : parts[0];
};
const moneyInWords = (value: number) => {
  const totalCents = Math.round(Math.max(0, value) * 100);
  const reais = Math.floor(totalCents / 100), cents = totalCents % 100;
  const parts: string[] = [];
  if (reais) parts.push(`${integerInWords(reais)} ${reais === 1 ? "real" : "reais"}`);
  if (cents) parts.push(`${integerInWords(cents)} ${cents === 1 ? "centavo" : "centavos"}`);
  return parts.length ? parts.join(" e ") : "zero reais";
};
const capitalizeMonth = (value: string) =>
  value.charAt(0).toUpperCase() + value.slice(1);
const isWorkDay = (employee: Recharge, date: Date) => {
  if (employee.scheduleType === "12x36") {
    const anchor = new Date(
        (employee.scheduleStartDate ||
          employee.hiredAt ||
          date.toISOString().slice(0, 10)) + "T12:00:00",
      ),
      difference = Math.floor((date.getTime() - anchor.getTime()) / 86400000);
    return difference >= 0 && difference % 2 === 0;
  }
  return (employee.workDays || [1, 2, 3, 4, 5]).includes(date.getDay());
};
const isEmployeeDismissed = (employee: Recharge, reference = new Date()) => {
  if (employee.terminationDate) {
    const termination = new Date(employee.terminationDate + "T12:00:00"),
      day = new Date(reference);
    day.setHours(12, 0, 0, 0);
    return termination <= day;
  }
  return employee.active === false || employee.employmentStatus === "Desligado";
};
const isDismissalPending = (employee: Recharge, reference = new Date()) => {
  if (!employee.terminationDate) return false;
  const termination = new Date(employee.terminationDate + "T12:00:00"),
    day = new Date(reference);
  day.setHours(12, 0, 0, 0);
  return termination > day;
};
type HROccurrence = {
  id: number;
  employeeId: number;
  date: string;
  endDate?: string;
  type: "Falta" | "Atestado" | "Atraso" | "Aviso" | "Férias";
  hours?: number;
  minutes?: number;
  days?: number;
  note?: string;
};
const projectRowsToPeriod = (
  allRows: Recharge[],
  period: string,
  referenceDate: string,
): Recharge[] => {
  const [year, month] = period.split("-").map(Number),
    identity = (r: Recharge) =>
      r.cpf?.replace(/\D/g, "") ||
      `${r.employee.trim().toLowerCase()}|${r.store.trim().toLowerCase()}`;
  return allRows.map((r) => {
    if (
      (r.status === "Recarregado" || r.status === "Recarregado atrasado") &&
      allRows.some((n) => n.sourceRechargeId === r.id)
    )
      return r;
    const original = new Date(r.rechargeDate + "T12:00:00"),
      lastDay = new Date(year, month, 0).getDate(),
      recharge = new Date(
        year,
        month - 1,
        Math.min(original.getDate(), lastDay),
        12,
      );
    const credit = new Date(recharge);
    credit.setDate(credit.getDate() + r.advance);
    const iso = (d: Date) =>
      `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    const rechargeDate = iso(recharge),
      creditDate = iso(credit);
    let days = 0;
    const end = new Date(credit);
    end.setMonth(end.getMonth() + 1);
    for (const d = new Date(credit); d < end; d.setDate(d.getDate() + 1))
      if (isWorkDay(r, d)) days++;
    const cardAmount = days * (r.cardDailyFare ?? r.dailyFare ?? 0),
      secondCardAmount = days * (r.secondCardDailyFare ?? 0);
    const completion = allRows
      .filter(
        (x) =>
          identity(x) === identity(r) &&
          x.completedDate?.startsWith(period) &&
          x.completedDate <= referenceDate,
      )
      .sort((a, b) =>
        (b.completedDate || "").localeCompare(a.completedDate || ""),
      )[0];
    const status: Status = completion
      ? completion.completedDate! > rechargeDate
        ? "Recarregado atrasado"
        : "Recarregado"
      : rechargeDate < referenceDate
        ? "Atrasado"
        : rechargeDate === referenceDate
          ? "Pendente"
          : "Próximo";
    return {
      ...r,
      rechargeDate,
      creditDate,
      status,
      completedDate: completion?.completedDate,
      periodCompletionId: completion?.id,
      cardAmount,
      secondCardAmount: r.secondCardType ? secondCardAmount : undefined,
      amount: cardAmount + secondCardAmount,
    };
  });
};
const currentEmployeeRows = (rows: Recharge[]) => {
  const leaves = rows.filter(
    (r) => !rows.some((next) => next.sourceRechargeId === r.id),
  );
  const unique = new Map<string, Recharge>();
  for (const row of leaves) {
    const key =
      row.cpf?.replace(/\D/g, "") ||
      `${row.employee.trim().toLowerCase()}|${row.store.trim().toLowerCase()}`;
    const saved = unique.get(key);
    if (!saved || row.id > saved.id) unique.set(key, row);
  }
  return [...unique.values()];
};
const remainingTransit = (r: Recharge, referenceDate: string) => {
  if (!r.completedDate) return { days: 0, balance: 0, card1: 0, card2: 0 };
  const start = new Date(r.completedDate + "T12:00:00");
  start.setDate(start.getDate() + r.advance);
  const reference = new Date(referenceDate + "T12:00:00");
  const workDays = r.workDays || [1, 2, 3, 4, 5];
  let charged = r.chargedDays ?? 0;
  if (!charged) {
    const fallbackEnd = new Date(start);
    fallbackEnd.setMonth(fallbackEnd.getMonth() + 1);
    for (const d = new Date(start); d < fallbackEnd; d.setDate(d.getDate() + 1))
      if (isWorkDay(r, d)) charged++;
  }
  let consumed = 0;
  for (const d = new Date(start); d < reference; d.setDate(d.getDate() + 1))
    if (isWorkDay(r, d)) consumed++;
  const days = Math.max(0, charged - consumed),
    card1 = days * (r.cardDailyFare ?? r.dailyFare ?? 0),
    card2 = days * (r.secondCardDailyFare ?? 0);
  return { days, balance: card1 + card2, card1, card2 };
};

function Sidebar({
  open,
  close,
  page,
  setPage,
  module,
  rechargeAlertCount,
  onChangeModule,
  onLogout,
}: {
  open: boolean;
  close: () => void;
  page: string;
  setPage: (v: string) => void;
  module: Module;
  rechargeAlertCount: number;
  onChangeModule: () => void;
  onLogout: () => void;
}) {
  const [userMenu, setUserMenu] = useState(false),
    [signedUser, setSignedUser] = useState<SessionUser | null>(null),
    nav =
      module === "people"
        ? peopleNav
        : module === "transit"
          ? transitNav
          : module === "finance"
            ? financeNav
            : administrativeNav;
  useEffect(() => {
    void cloudCurrentUser().then(setSignedUser);
  }, []);
  const nameParts = (signedUser?.fullName || "Usuário").trim().split(/\s+/),
    shortName =
      nameParts.length > 1
        ? `${nameParts[0]} ${nameParts[nameParts.length - 1]}`
        : nameParts[0],
    initials =
      (nameParts[0]?.[0] || "U") +
      (nameParts.length > 1 ? nameParts[nameParts.length - 1]?.[0] || "" : "");
  return (
    <>
      <div
        onClick={close}
        className={`fixed inset-0 z-30 bg-black/30 lg:hidden ${open ? "block" : "hidden"}`}
      />
      <aside
        className={`fixed inset-y-0 left-0 z-40 flex w-[246px] flex-col bg-forest-900 text-white transition-transform lg:translate-x-0 ${open ? "translate-x-0" : "-translate-x-full"}`}
      >
        <div className="relative flex h-32 flex-col items-center justify-center border-b border-white/10 px-5">
          <img
            src="/sacolao-abc-logo.png?v=4"
            alt="Sacolão ABC"
            className="h-[82px] w-[208px] object-contain"
          />
          <span className="text-[9px] font-semibold uppercase tracking-[.18em] text-white/60">
            {module === "people"
              ? "Gestão de pessoas"
              : module === "transit"
                ? "Cartões de passagem"
                : module === "finance"
                  ? "Gestão financeira"
                  : "Gestão administrativa"}
          </span>
          <button onClick={close} className="absolute right-4 top-4 lg:hidden">
            <X size={20} />
          </button>
        </div>
        <div className="px-4 pt-7 text-[10px] font-semibold uppercase tracking-[.18em] text-emerald-200/60">
          Menu principal
        </div>
        <nav className="mt-3 space-y-1 px-3">
          {nav.map(([label, Icon]) => (
            <button
              key={label}
              onClick={() => {
                setPage(label);
                close();
              }}
              className={`flex w-full items-center gap-3 rounded-xl px-3 py-3 text-sm transition ${page === label ? "bg-white text-forest-900 shadow-lg" : "text-emerald-50/75 hover:bg-white/10 hover:text-white"}`}
            >
              <Icon size={19} />
              {label}
              {label === "Recargas" && rechargeAlertCount > 0 && (
                <span className="ml-auto rounded-full bg-red-500 px-2 py-0.5 text-[10px] text-white">
                  {rechargeAlertCount}
                </span>
              )}
            </button>
          ))}
        </nav>
        <div className="relative mt-auto border-t border-white/10 p-3">
          <button
            onClick={onChangeModule}
            className="flex w-full items-center gap-3 rounded-xl px-3 py-3 text-sm text-white/75 hover:bg-white/10"
          >
            <ArrowLeft size={19} />
            Trocar de módulo
          </button>
          {signedUser?.role !== "operator" && <button
            onClick={() => setPage("Configurações")}
            className="flex w-full items-center gap-3 rounded-xl px-3 py-3 text-sm text-white/75 hover:bg-white/10"
          >
            <Settings size={19} />
            Configurações
          </button>}
          {userMenu && (
            <div className="absolute bottom-[82px] left-3 right-3 overflow-hidden rounded-xl border border-slate-200 bg-white p-1 text-slate-700 shadow-2xl">
              <button
                onClick={() => {
                  setUserMenu(false);
                  onLogout();
                }}
                className="flex w-full items-center gap-2 rounded-lg px-3 py-2.5 text-sm font-semibold text-red-600 hover:bg-red-50"
              >
                <X size={17} />
                Sair do sistema
              </button>
            </div>
          )}
          <button
            onClick={() => setUserMenu(!userMenu)}
            className="mt-2 flex w-full items-center gap-3 rounded-xl bg-white/5 p-3 text-left hover:bg-white/10"
          >
            <div className="grid h-9 w-9 place-items-center rounded-full border border-white bg-white font-bold text-black grayscale">
              {initials.toUpperCase()}
            </div>
            <div className="min-w-0">
              <div className="truncate text-sm font-semibold">{shortName}</div>
              <div className="text-[11px] text-white/50">
                {signedUser?.role === "operator" ? "Operador" : "Administrador"}
              </div>
            </div>
            <MoreHorizontal className="ml-auto" size={18} />
          </button>
        </div>
      </aside>
    </>
  );
}

function Header({
  menu,
  page,
  module,
  referenceDate,
  setReferenceDate,
  alertCount,
  stores,
  selectedStore,
  setSelectedStore,
  roles,
  selectedRole,
  setSelectedRole,
  dark,
  toggleTheme,
}: {
  menu: () => void;
  page: string;
  module: Module;
  referenceDate: string;
  setReferenceDate: (v: string) => void;
  alertCount: number;
  stores: string[];
  selectedStore: string;
  setSelectedStore: (v: string) => void;
  roles: string[];
  selectedRole: string;
  setSelectedRole: (v: string) => void;
  dark: boolean;
  toggleTheme: () => void;
}) {
  return (
    <header className="flex h-20 items-center border-b border-slate-200 bg-white px-4 sm:px-7">
      <button
        onClick={menu}
        className="mr-3 rounded-lg p-2 hover:bg-slate-100 lg:hidden"
      >
        <Menu />
      </button>
      <div>
        <div className="text-xs text-slate-400">
          {module === "people"
            ? "Gestão de Pessoas"
            : module === "transit"
              ? "Gestão de Cartões de Passagem"
              : module === "finance"
                ? "Gestão Financeira"
                : "Gestão Administrativa"}
        </div>
        <h1 className="text-xl font-bold text-slate-900">{page}</h1>
      </div>
      <div className="ml-auto flex items-center gap-2">
        <button
          onClick={toggleTheme}
          title={dark ? "Usar tema claro" : "Usar tema escuro"}
          aria-label="Alternar tema"
          className="grid h-10 w-10 place-items-center rounded-xl border border-slate-200 bg-white text-lg text-slate-600 hover:bg-slate-50"
        >
          {dark ? "☀" : "◐"}
        </button>
        {module === "transit" && (
          <button
            onClick={() =>
              alert(
                `Você possui ${alertCount} recarga(s) que precisam de atenção.`,
              )
            }
            className="relative rounded-xl border border-slate-200 p-2.5 text-slate-500 hover:bg-slate-50"
          >
            <Bell size={19} />
            {alertCount > 0 && (
              <i className="absolute right-2 top-2 h-2 w-2 rounded-full bg-red-500 ring-2 ring-white" />
            )}
          </button>
        )}
        <div className="relative hidden md:block">
          <Store
            className="pointer-events-none absolute left-3 top-2.5 text-forest-700"
            size={17}
          />
          <select
            aria-label="Filtrar por loja"
            value={selectedStore}
            onChange={(e) => setSelectedStore(e.target.value)}
            title={selectedStore === "Todas" ? "Todas as lojas" : selectedStore}
            className="w-48 cursor-pointer rounded-xl border border-slate-200 bg-white py-2.5 pl-9 pr-8 text-xs font-semibold text-slate-600 outline-none hover:border-forest-300"
          >
            <option value="Todas">Todas as lojas</option>
            {stores.map((s) => (
              <option key={s}>{s}</option>
            ))}
          </select>
        </div>
        <div className="relative hidden md:block">
          <UserRound
            className="pointer-events-none absolute left-3 top-2.5 text-forest-700"
            size={17}
          />
          <select
            aria-label="Filtrar por função"
            value={selectedRole}
            onChange={(e) => setSelectedRole(e.target.value)}
            title={selectedRole === "Todas" ? "Todas as funções" : selectedRole}
            className="w-48 cursor-pointer rounded-xl border border-slate-200 bg-white py-2.5 pl-9 pr-8 text-xs font-semibold text-slate-600 outline-none hover:border-forest-300"
          >
            <option value="Todas">Todas as funções</option>
            {roles.map((role) => (
              <option key={role}>{role}</option>
            ))}
          </select>
        </div>
        {module === "transit" && (
          <div className="relative hidden sm:block">
            <CalendarDays
              className="pointer-events-none absolute left-3 top-2.5 text-forest-700"
              size={17}
            />
            <input
              aria-label="Selecionar dia, mês e ano"
              title="Escolher data de referência"
              type="date"
              value={referenceDate}
              onChange={(e) => setReferenceDate(e.target.value)}
              className="cursor-pointer rounded-xl border border-slate-200 bg-white py-2.5 pl-9 pr-3 text-xs font-semibold text-slate-600 outline-none hover:border-forest-300 focus:border-forest-500"
            />
          </div>
        )}
      </div>
    </header>
  );
}

function ReferenceDayBar({
  period,
  day,
  setDay,
}: {
  period: string;
  day: number;
  setDay: (d: number) => void;
}) {
  const [year, month] = period.split("-").map(Number),
    count = new Date(year, month, 0).getDate();
  return (
    <div className="border-b border-slate-200 bg-white px-4 py-3 sm:px-7">
      <div className="flex items-center gap-3">
        <div className="hidden shrink-0 sm:block">
          <div className="text-xs font-bold text-slate-700">Saldo na data</div>
          <div className="text-[10px] text-slate-400">Escolha o dia</div>
        </div>
        <div className="flex flex-1 gap-1.5 overflow-x-auto pb-1">
          {Array.from({ length: count }, (_, i) => i + 1).map((d) => (
            <button
              key={d}
              onClick={() => setDay(d)}
              className={`min-w-9 rounded-lg px-2 py-2 text-xs font-bold transition ${day === d ? "bg-forest-700 text-white shadow-md" : "bg-slate-50 text-slate-500 hover:bg-forest-50 hover:text-forest-700"}`}
            >
              {String(d).padStart(2, "0")}
            </button>
          ))}
        </div>
        <div className="hidden shrink-0 rounded-xl bg-forest-50 px-3 py-2 text-right md:block">
          <div className="text-[10px] text-slate-400">Data de referência</div>
          <b className="text-sm text-forest-800">
            {String(day).padStart(2, "0")}/{String(month).padStart(2, "0")}/
            {year}
          </b>
        </div>
      </div>
    </div>
  );
}

function StatCard({
  title,
  value,
  note,
  icon: Icon,
  tone,
  onClick,
}: {
  title: string;
  value: string;
  note: string;
  icon: any;
  tone: string;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="w-full rounded-2xl border border-slate-200 bg-white p-5 text-left shadow-soft transition hover:-translate-y-1 hover:border-slate-400 hover:shadow-lg"
    >
      <div className="flex items-start justify-between">
        <div className={`grid h-10 w-10 place-items-center rounded-xl ${tone}`}>
          <Icon size={20} />
        </div>
        <ChevronRight size={18} className="text-slate-400" />
      </div>
      <div className="mt-5 text-3xl font-bold tracking-tight text-slate-900">
        {value}
      </div>
      <div className="mt-1 text-sm font-semibold text-slate-700">{title}</div>
      <div className="mt-2 text-xs text-slate-400">{note}</div>
    </button>
  );
}

function StatusBadge({ status }: { status: Status }) {
  return (
    <span
      className={`inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-[11px] font-semibold ${statusStyle[status]}`}
    >
      <i className="h-1.5 w-1.5 rounded-full bg-current" />
      {status}
    </span>
  );
}

function IndicatorDetails({
  title,
  rows,
  close,
  occurrences = [],
}: {
  title: string;
  rows: Recharge[];
  close: () => void;
  occurrences?: HROccurrence[];
}) {
  const transit = /recarga|valor|vale-transporte/i.test(title),
    experience = /experiência/i.test(title),
    criticalExperience = false,
    criticalPeople = /funcionários críticos/i.test(title),
    notice = /cumprindo aviso/i.test(title),
    vacation = /de férias/i.test(title),
    termination = /rescisões/i.test(title),
    unregistered = /sem carteira/i.test(title),
    today = new Date();
  today.setHours(12, 0, 0, 0);
  const experienceDeadline = (r: Recharge) => {
    if (!r.hiredAt || r.formalEmployment === false) return null;
    const admission = new Date(r.hiredAt + "T12:00:00"),
      first = new Date(admission),
      final = new Date(admission);
    first.setDate(first.getDate() + 29);
    final.setDate(final.getDate() + 89);
    const inCurrentMonth = (d: Date) =>
      d.getFullYear() === today.getFullYear() &&
      d.getMonth() === today.getMonth();
    const date = criticalExperience
      ? final
      : inCurrentMonth(first)
        ? first
        : inCurrentMonth(final)
          ? final
          : null;
    if (!date) return null;
    const days = Math.ceil((date.getTime() - today.getTime()) / 86400000);
    return {
      date,
      days,
      stage: criticalExperience
        ? "Fim do período de experiência (90 dias)"
        : date === first
          ? "Fim dos primeiros 30 dias"
          : "Fim da prorrogação de 60 dias",
    };
  };
  const noticeDetails = (r: Recharge) => {
    if (!r.noticeStart || !r.noticeEnd) return null;
    const start = new Date(r.noticeStart + "T12:00:00"),
      end = new Date(r.noticeEnd + "T12:00:00"),
      days = Math.ceil((end.getTime() - today.getTime()) / 86400000);
    return { start, end, days };
  };
  const vacationDetails = (r: Recharge) => {
    const item = occurrences.find((occurrence) => {
      if (
        occurrence.employeeId !== r.id ||
        occurrence.type !== "Férias" ||
        !occurrence.endDate
      )
        return false;
      const start = new Date(occurrence.date + "T12:00:00"),
        end = new Date(occurrence.endDate + "T12:00:00");
      return start <= today && end >= today;
    });
    if (!item?.endDate) return null;
    return {
      start: new Date(item.date + "T12:00:00"),
      end: new Date(item.endDate + "T12:00:00"),
    };
  };
  const tenureDetails = (r: Recharge) => {
    if (!r.hiredAt) return null;
    const admission = new Date(r.hiredAt + "T12:00:00"),
      days = Math.max(
        0,
        Math.floor((today.getTime() - admission.getTime()) / 86400000),
      );
    if (days < 90 && r.formalEmployment !== false)
      return { admission, text: `${days} dia(s) · Em experiência` };
    if (days < 365) {
      const months = Math.floor(days / 30),
        remaining = days % 30;
      return {
        admission,
        text: `${months} ${months === 1 ? "mês" : "meses"} e ${remaining} dia(s)`,
      };
    }
    const years = Math.floor(days / 365),
      remaining = days % 365;
    return {
      admission,
      text: `${years} ${years === 1 ? "ano" : "anos"} e ${remaining} dia(s)`,
    };
  };
  const terminationDetails = (r: Recharge) => {
    if (!r.terminationDate) return null;
    const date = new Date(r.terminationDate + "T12:00:00"),
      payment = new Date(date);
    payment.setDate(payment.getDate() + 10);
    const days = Math.ceil((payment.getTime() - today.getTime()) / 86400000);
    return { date, payment, days };
  };
  return createPortal(
    <div className="fixed inset-0 z-[80] grid place-items-center bg-slate-950/50 p-4 backdrop-blur-sm">
      <div className="max-h-[85vh] w-full max-w-2xl overflow-hidden rounded-2xl bg-white shadow-2xl">
        <div className="flex items-center border-b border-slate-200 px-5 py-4">
          <div>
            <h3 className="text-lg font-bold text-slate-900">{title}</h3>
            <p className="text-xs text-slate-400">
              {rows.length} funcionário(s)
            </p>
          </div>
          <button
            onClick={close}
            className="ml-auto rounded-lg p-2 text-slate-500 hover:bg-slate-100"
          >
            <X size={20} />
          </button>
        </div>
        <div className="max-h-[65vh] overflow-y-auto p-4">
          <div className="space-y-2">
            {rows.map((r) => {
              const deadline = experience ? experienceDeadline(r) : null,
                noticeInfo = notice ? noticeDetails(r) : null,
                vacationInfo = vacation ? vacationDetails(r) : null,
                terminationInfo = termination ? terminationDetails(r) : null,
                unregisteredInfo = unregistered
                  ? {
                      start: r.unregisteredStartDate || r.hiredAt,
                      reason: r.unregisteredReason || "Motivo não informado",
                    }
                  : null,
                tenureInfo = criticalPeople ? tenureDetails(r) : null;
              return (
                <div
                  key={r.id}
                  className="flex items-center gap-3 rounded-xl border border-slate-200 p-4"
                >
                  <div className="grid h-10 w-10 shrink-0 place-items-center rounded-full bg-slate-100 text-xs font-bold text-slate-700">
                    {r.employee
                      .split(" ")
                      .map((x) => x[0])
                      .slice(0, 2)
                      .join("")}
                  </div>
                  <div className="min-w-0">
                    <div className="truncate font-semibold text-slate-900">
                      {r.employee}
                    </div>
                    <div className="text-xs text-slate-500">
                      {r.store} · {r.role}
                    </div>
                    {deadline && (
                      <div className="mt-1 text-xs font-medium text-slate-600">
                        {deadline.stage}
                      </div>
                    )}
                  </div>
                  <div className="ml-auto shrink-0 text-right text-xs text-slate-500">
                    {deadline ? (
                      <>
                        <div className="font-bold text-slate-800">
                          Vencimento:{" "}
                          {deadline.date.toLocaleDateString("pt-BR")}
                        </div>
                        <div
                          className={
                            deadline.days <= 3
                              ? "mt-1 font-bold text-red-600"
                              : "mt-1 font-semibold text-amber-600"
                          }
                        >
                          {deadline.days > 1
                            ? `Faltam ${deadline.days} dias`
                            : deadline.days === 1
                              ? "Falta 1 dia"
                              : deadline.days === 0
                                ? "Vence hoje"
                                : `Venceu há ${Math.abs(deadline.days)} dia(s)`}
                        </div>
                      </>
                    ) : noticeInfo ? (
                      <>
                        <div>
                          Início:{" "}
                          <b>{noticeInfo.start.toLocaleDateString("pt-BR")}</b>
                        </div>
                        <div>
                          Término:{" "}
                          <b>{noticeInfo.end.toLocaleDateString("pt-BR")}</b>
                        </div>
                        <div
                          className={
                            noticeInfo.days <= 3
                              ? "mt-1 font-bold text-red-600"
                              : "mt-1 font-semibold text-amber-600"
                          }
                        >
                          {noticeInfo.days > 1
                            ? `Faltam ${noticeInfo.days} dias`
                            : noticeInfo.days === 1
                              ? "Falta 1 dia"
                              : noticeInfo.days === 0
                                ? "Termina hoje"
                                : "Aviso encerrado"}
                        </div>
                      </>
                    ) : vacationInfo ? (
                      <>
                        <div>
                          Início: <b>{vacationInfo.start.toLocaleDateString("pt-BR")}</b>
                        </div>
                        <div>
                          Término: <b>{vacationInfo.end.toLocaleDateString("pt-BR")}</b>
                        </div>
                      </>
                    ) : terminationInfo ? (
                      <>
                        <div>
                          Desligamento:{" "}
                          <b>
                            {terminationInfo.date.toLocaleDateString("pt-BR")}
                          </b>
                        </div>
                        <div>
                          Acerto até:{" "}
                          <b>
                            {terminationInfo.payment.toLocaleDateString(
                              "pt-BR",
                            )}
                          </b>
                        </div>
                        <div
                          className={
                            terminationInfo.days < 0
                              ? "mt-1 font-bold text-red-600"
                              : "mt-1 font-semibold text-amber-600"
                          }
                        >
                          {terminationInfo.days > 1
                            ? `Faltam ${terminationInfo.days} dias para o acerto`
                            : terminationInfo.days === 1
                              ? "Falta 1 dia para o acerto"
                              : terminationInfo.days === 0
                                ? "Acerto vence hoje"
                                : `Acerto atrasado há ${Math.abs(terminationInfo.days)} dia(s)`}
                        </div>
                      </>
                    ) : tenureInfo ? (
                      <>
                        <div>
                          Admissão:{" "}
                          <b>
                            {tenureInfo.admission.toLocaleDateString("pt-BR")}
                          </b>
                        </div>
                        <div className="mt-1">
                          Tempo de casa: <b>{tenureInfo.text}</b>
                        </div>
                      </>
                    ) : unregisteredInfo ? (
                      <>
                        <div>
                          Início:{" "}
                          <b>
                            {unregisteredInfo.start
                              ? formatDate(unregisteredInfo.start)
                              : "Não informado"}
                          </b>
                        </div>
                        <div className="mt-1 max-w-[240px]">
                          Motivo: <b>{unregisteredInfo.reason}</b>
                        </div>
                      </>
                    ) : (
                      <>
                        {transit ? r.status : r.employmentStatus || "Ativo"}
                        {transit && r.rechargeDate ? (
                          <div>Recarga: {formatDate(r.rechargeDate)}</div>
                        ) : r.hiredAt ? (
                          <div>Admissão: {formatDate(r.hiredAt)}</div>
                        ) : null}
                      </>
                    )}
                  </div>
                </div>
              );
            })}
            {!rows.length && (
              <div className="py-12 text-center text-sm text-slate-400">
                Nenhum funcionário nesta categoria.
              </div>
            )}
          </div>
        </div>
      </div>
    </div>,
    document.body,
  );
}

function RechargeDayPicker({
  selected,
  onSelect,
  occupancy,
  store,
}: {
  selected: number;
  onSelect: (day: number) => void;
  occupancy: (day: number) => number;
  store: string;
}) {
  const now = new Date(),
    year = now.getFullYear(),
    month = now.getMonth(),
    first = new Date(year, month, 1).getDay(),
    count = new Date(year, month + 1, 0).getDate(),
    cells = Array(first)
      .fill(null)
      .concat(Array.from({ length: count }, (_, i) => i + 1));
  return (
    <div>
      <div className="mb-2 flex items-center justify-between">
        <div>
          <div className="text-xs font-semibold text-slate-600">
            Dia mensal da recarga
          </div>
          <div className="text-[11px] capitalize text-slate-400">
            {now.toLocaleDateString("pt-BR", {
              month: "long",
              year: "numeric",
            })}
          </div>
        </div>
        <CalendarDays size={18} className="text-forest-700" />
      </div>
      <div className="rounded-2xl border border-slate-200 bg-white p-3">
        <div className="grid grid-cols-7 text-center text-[10px] font-bold uppercase text-slate-400">
          {["D", "S", "T", "Q", "Q", "S", "S"].map((d, i) => (
            <div className="py-1" key={i}>
              {d}
            </div>
          ))}
        </div>
        <div className="mt-1 grid grid-cols-7 gap-1">
          {cells.map((day, i) => {
            if (!day) return <div key={i} />;
            const used = store ? occupancy(day) : 0,
              full = used >= 2 && day !== selected;
            return (
              <button
                type="button"
                key={i}
                disabled={full}
                onClick={() => onSelect(day)}
                title={
                  !store
                    ? "Selecione primeiro uma loja"
                    : full
                      ? "Dia lotado"
                      : `${2 - used} vaga(s) disponível(is)`
                }
                className={`relative aspect-square rounded-lg text-xs font-semibold transition ${day === selected ? "bg-forest-700 text-white shadow-md" : full ? "cursor-not-allowed bg-red-50 text-red-300 line-through" : "text-slate-600 hover:bg-forest-50 hover:text-forest-700"}`}
              >
                {day}
                {store && used > 0 && (
                  <span
                    className={`absolute bottom-0.5 left-1/2 h-1 w-1 -translate-x-1/2 rounded-full ${full ? "bg-red-400" : day === selected ? "bg-white" : "bg-amber-500"}`}
                  />
                )}
              </button>
            );
          })}
        </div>
        <div className="mt-3 flex items-center gap-4 border-t border-slate-100 pt-3 text-[10px] text-slate-500">
          <span className="flex items-center gap-1">
            <i className="h-2 w-2 rounded-full bg-amber-500" />1 pessoa
          </span>
          <span className="flex items-center gap-1">
            <i className="h-2 w-2 rounded-full bg-red-400" />
            Lotado
          </span>
          <span className="ml-auto font-semibold text-forest-700">
            Selecionado: dia {String(selected).padStart(2, "0")}
          </span>
        </div>
      </div>
    </div>
  );
}

function Dashboard({
  rows,
  setRows,
  openForm,
  onMark,
  referenceDate,
}: {
  rows: Recharge[];
  setRows: (r: Recharge[]) => void;
  openForm: () => void;
  onMark: (r: Recharge) => void;
  referenceDate: string;
}) {
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState("Todos");
  const [detail, setDetail] = useState<{
    title: string;
    rows: Recharge[];
  } | null>(null);
  const currentRows = currentEmployeeRows(rows);
  const visible = currentRows.filter(
    (r) =>
      (filter === "Todos" || r.status === filter) &&
      `${r.employee} ${r.store} ${r.cardType}`
        .toLowerCase()
        .includes(query.toLowerCase()),
  );
  const today = new Date().toISOString().slice(0, 10),
    tomorrow = new Date(Date.now() + 86400000).toISOString().slice(0, 10);
  const todayRows = currentRows.filter((r) => r.rechargeDate === today),
    late = currentRows.filter((r) => r.status === "Atrasado"),
    next = currentRows.filter((r) => r.status === "Próximo"),
    tomorrowRows = currentRows.filter((r) => r.rechargeDate === tomorrow);
  const total = (list: Recharge[]) =>
    list
      .reduce((s, r) => s + (r.amount || 0), 0)
      .toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  return (
    <main className="fade-in p-4 sm:p-7">
      <div className="flex flex-col justify-between gap-4 sm:flex-row sm:items-center">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">
            Bom dia, João! <span className="inline-block">👋</span>
          </h2>
          <p className="mt-1 text-sm text-slate-500">
            Aqui está o resumo das recargas da sua equipe.
          </p>
        </div>
        <button
          onClick={openForm}
          className="flex items-center justify-center gap-2 rounded-xl bg-forest-700 px-4 py-3 text-sm font-semibold text-white shadow-lg shadow-forest-700/20 hover:bg-forest-800"
        >
          <Plus size={18} />
          Novo funcionário
        </button>
      </div>
      <section className="mt-6 grid grid-cols-2 gap-3 lg:grid-cols-5">
        <StatCard
          onClick={() =>
            setDetail({
              title: "Funcionários com vale-transporte",
              rows: currentRows,
            })
          }
          title="Funcionários"
          value={String(currentRows.length)}
          note="No filtro selecionado"
          icon={Users}
          tone="bg-forest-50 text-forest-700"
        />
        <StatCard
          onClick={() =>
            setDetail({ title: "Recargas previstas no mês", rows: currentRows })
          }
          title="Recargas no mês"
          value={String(currentRows.length)}
          note={`${total(currentRows)} previstos`}
          icon={Clock3}
          tone="bg-amber-50 text-amber-600"
        />
        <StatCard
          onClick={() => setDetail({ title: "Recargas em atraso", rows: late })}
          title="Em atraso"
          value={String(late.length)}
          note="No período selecionado"
          icon={TriangleAlert}
          tone="bg-red-50 text-red-600"
        />
        <StatCard
          onClick={() =>
            setDetail({
              title: "Composição do valor previsto",
              rows: currentRows,
            })
          }
          title="Valor previsto"
          value={total(currentRows)}
          note="Total do mês"
          icon={CalendarDays}
          tone="bg-blue-50 text-blue-600"
        />
        <StatCard
          onClick={() => setDetail({ title: "Próximas recargas", rows: next })}
          title="Próximas"
          value={String(next.length)}
          note="A realizar no período"
          icon={ChevronRight}
          tone="bg-violet-50 text-violet-600"
        />
      </section>
      <section className="mt-6">
        <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
          <div className="flex flex-col gap-3 border-b border-slate-100 p-5 sm:flex-row sm:items-center">
            <div>
              <h3 className="font-bold text-slate-900">
                Recargas prioritárias
              </h3>
              <p className="text-xs text-slate-400">
                Situação no mês selecionado
              </p>
            </div>
            <div className="relative sm:ml-auto">
              <Search
                className="absolute left-3 top-2.5 text-slate-400"
                size={17}
              />
              <input
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Buscar funcionário..."
                className="w-full rounded-xl border border-slate-200 py-2 pl-9 pr-3 text-sm outline-none focus:border-forest-500 sm:w-52"
              />
            </div>
            <select
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              className="rounded-xl border border-slate-200 px-3 py-2 text-sm text-slate-600 outline-none"
            >
              <option>Todos</option>
              <option>Pendente</option>
              <option>Atrasado</option>
              <option>Próximo</option>
              <option>Recarregado</option>
              <option>Recarregado atrasado</option>
            </select>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full min-w-[1100px] text-left">
              <thead className="bg-slate-50 text-[10px] uppercase tracking-wider text-slate-400">
                <tr>
                  {[
                    "Funcionário",
                    "Loja / cartão",
                    "Última recarga",
                    "Saldo estimado",
                    "Crédito disponível",
                    "Data da recarga",
                    "Status",
                    "Ação",
                  ].map((x) => (
                    <th key={x} className="px-5 py-3 font-semibold">
                      {x}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {visible.map((r) => {
                  const remaining = remainingTransit(r, referenceDate);
                  return (
                    <tr key={r.id} className="hover:bg-slate-50/70">
                      <td className="px-5 py-4">
                        <div className="flex items-center gap-3">
                          <div className="grid h-9 w-9 place-items-center rounded-full bg-forest-100 text-xs font-bold text-forest-700">
                            {r.employee
                              .split(" ")
                              .map((x) => x[0])
                              .slice(0, 2)
                              .join("")}
                          </div>
                          <div>
                            <div className="text-sm font-semibold text-slate-800">
                              {r.employee}
                            </div>
                            <div className="text-[11px] text-slate-400">
                              {r.role}
                            </div>
                          </div>
                        </div>
                      </td>
                      <td className="px-5 py-4">
                        <div className="text-sm text-slate-700">{r.store}</div>
                        <div className="text-[11px] text-slate-400">
                          {r.cardType}
                          {r.secondCardType ? ` + ${r.secondCardType}` : ""}
                        </div>
                      </td>
                      <td className="px-5 py-4">
                        <div className="text-sm font-semibold text-slate-700">
                          {r.completedDate ? formatDate(r.completedDate) : "—"}
                        </div>
                        <div className="text-[11px] text-slate-400">
                          Data realizada
                        </div>
                      </td>
                      <td className="px-5 py-4">
                        <div className="text-sm font-bold text-forest-700">
                          {remaining.days} dias
                        </div>
                        <div className="text-[11px] text-slate-500">
                          {remaining.balance.toLocaleString("pt-BR", {
                            style: "currency",
                            currency: "BRL",
                          })}
                        </div>
                        {r.secondCardType && (
                          <div className="mt-1 text-[10px] text-slate-400">
                            {r.cardType}:{" "}
                            {remaining.card1.toLocaleString("pt-BR", {
                              style: "currency",
                              currency: "BRL",
                            })}
                            <br />
                            {r.secondCardType}:{" "}
                            {remaining.card2.toLocaleString("pt-BR", {
                              style: "currency",
                              currency: "BRL",
                            })}
                          </div>
                        )}
                      </td>
                      <td className="px-5 py-4 text-sm font-medium text-slate-700">
                        {formatDate(r.creditDate)}
                      </td>
                      <td className="px-5 py-4">
                        <div className="text-sm font-bold text-forest-700">
                          {formatDate(r.rechargeDate)}
                        </div>
                        <div className="text-[11px] text-slate-400">
                          {r.advance} dias depois
                        </div>
                      </td>
                      <td className="px-5 py-4">
                        <StatusBadge status={r.status} />
                      </td>
                      <td className="px-5 py-4">
                        {r.status !== "Recarregado" &&
                        r.status !== "Recarregado atrasado" ? (
                          <button
                            onClick={() => onMark(r)}
                            className="rounded-lg border border-forest-200 px-3 py-2 text-xs font-semibold text-forest-700 hover:bg-forest-50"
                          >
                            Marcar recarga
                          </button>
                        ) : (
                          <span className="flex items-center gap-1 text-xs text-slate-400">
                            <Check size={14} />
                            Feita em{" "}
                            {r.completedDate ? formatDate(r.completedDate) : ""}
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {visible.length === 0 && (
              <div className="px-5 py-14 text-center text-sm text-slate-400">
                Nenhuma recarga cadastrada. Cadastre a primeira loja e depois um
                funcionário.
              </div>
            )}
          </div>
          <div className="flex items-center justify-between border-t border-slate-100 px-5 py-4 text-xs text-slate-400">
            <span>Mostrando {visible.length} recargas</span>
          </div>
        </div>
      </section>
      {detail && (
        <IndicatorDetails
          title={detail.title}
          rows={detail.rows}
          close={() => setDetail(null)}
        />
      )}
    </main>
  );
}

function MiniCalendar({ rows }: { rows: Recharge[] }) {
  const base = rows[0]
      ? new Date(rows[0].rechargeDate + "T12:00:00")
      : new Date(),
    year = base.getFullYear(),
    month = base.getMonth(),
    first = new Date(year, month, 1).getDay(),
    count = new Date(year, month + 1, 0).getDate(),
    days = Array(first)
      .fill(null)
      .concat(Array.from({ length: count }, (_, i) => i + 1));
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
      <div>
        <h3 className="font-bold capitalize">
          {base.toLocaleDateString("pt-BR", { month: "long", year: "numeric" })}
        </h3>
        <span className="text-xs text-slate-400">Calendário de recargas</span>
      </div>
      <div className="mt-4 grid grid-cols-7 text-center text-[10px] font-semibold text-slate-400">
        {["D", "S", "T", "Q", "Q", "S", "S"].map((d, i) => (
          <div key={i}>{d}</div>
        ))}
      </div>
      <div className="mt-2 grid grid-cols-7 gap-y-1 text-center text-xs">
        {days.map((d, i) => {
          const has =
            d &&
            rows.some(
              (r) => new Date(r.rechargeDate + "T12:00:00").getDate() === d,
            );
          return (
            <div
              key={i}
              className={`relative grid aspect-square place-items-center rounded-lg ${has ? "bg-forest-700 font-bold text-white" : d ? "text-slate-600" : "text-slate-200"}`}
            >
              {d || ""}
              {has && (
                <i className="absolute bottom-0.5 h-1 w-1 rounded-full bg-white" />
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function EmployeeModal({
  close,
  add,
  stores,
  positions,
  unregisteredReasons,
  initial,
  rows,
}: {
  close: () => void;
  add: (r: Recharge) => void;
  stores: string[];
  positions: string[];
  unregisteredReasons: string[];
  initial?: Recharge | null;
  rows: Recharge[];
}) {
  const today = new Date().toISOString().slice(0, 10),
    todayDate = new Date();
  const emptyForm = {
    name: "",
    cpf: "",
    role: "",
    store: "",
    hiredAt: today,
    cardType: "BHBus",
    cardFare: "",
    hasSecondCard: false,
    secondCardType: "Ótimo",
    secondCardFare: "",
    rechargeDay: String(todayDate.getDate()),
    advance: Number(localStorage.getItem("valefluxo_advance") || 3),
    active: true,
    birthDate: "",
    gender: "" as Recharge["gender"] | "",
    experienceDays: Number(
      localStorage.getItem("valefluxo_experience_days") || 90,
    ),
    experienceCritical: false,
    priorityTerminationDate: "",
    employmentStatus: "Ativo" as NonNullable<Recharge["employmentStatus"]>,
    noticeStart: "",
    noticeEnd: "",
    terminationDate: "",
    formalEmployment: true,
    unregisteredStartDate: today,
    unregisteredReason: "",
    receivesTransit: true,
    receivesCostAssistance: false,
    costAssistanceAmount: "",
  };
  const initialForm = initial
    ? {
        ...emptyForm,
        name: initial.employee,
        cpf: initial.cpf || "",
        role: initial.role,
        store: initial.store,
        hiredAt: initial.hiredAt || today,
        cardType: initial.cardType,
        cardFare: String(initial.cardDailyFare ?? initial.dailyFare ?? ""),
        hasSecondCard: !!initial.secondCardType,
        secondCardType: initial.secondCardType || "Ótimo",
        secondCardFare: String(initial.secondCardDailyFare ?? ""),
        rechargeDay: String(
          new Date(initial.rechargeDate + "T12:00:00").getDate(),
        ),
        advance: initial.advance,
        active: initial.active !== false,
        birthDate: initial.birthDate || "",
        gender: initial.gender || genderFromRole(initial.role) || "",
        experienceDays:
          initial.experienceDays ||
          Number(localStorage.getItem("valefluxo_experience_days") || 90),
        experienceCritical: initial.experienceCritical === true,
        priorityTerminationDate: initial.priorityTerminationDate || "",
        employmentStatus:
          initial.employmentStatus ||
          (initial.active === false ? "Desligado" : "Ativo"),
        noticeStart: initial.noticeStart || "",
        noticeEnd: initial.noticeEnd || "",
        terminationDate: initial.terminationDate || "",
        formalEmployment: initial.formalEmployment !== false,
        unregisteredStartDate:
          initial.unregisteredStartDate || initial.hiredAt || today,
        unregisteredReason: initial.unregisteredReason || "",
        receivesTransit: initial.receivesTransit !== false,
        receivesCostAssistance: initial.receivesCostAssistance === true,
        costAssistanceAmount: initial.costAssistanceAmount
          ? formatMoneyInput(
              String(Math.round(initial.costAssistanceAmount * 100)),
            )
          : "",
      }
    : null;
  const [form, setForm] = useState(() => {
    if (initialForm) return initialForm;
    try {
      return {
        ...emptyForm,
        ...JSON.parse(localStorage.getItem("valefluxo_employee_draft") || "{}"),
      };
    } catch {
      return emptyForm;
    }
  });
  const [scheduleType, setScheduleType] = useState<
    "Personalizada" | "6x1" | "12x36"
  >(initial?.scheduleType || "Personalizada");
  const [scheduleStartDate, setScheduleStartDate] = useState(
    initial?.scheduleStartDate || initial?.hiredAt || today,
  );
  const [workDays, setWorkDays] = useState<number[]>(
    () =>
      initial?.workDays ||
      (() => {
        try {
          return JSON.parse(
            localStorage.getItem("valefluxo_workdays_draft") || "[1,2,3,4,5]",
          );
        } catch {
          return [1, 2, 3, 4, 5];
        }
      })(),
  );
  useEffect(() => {
    if (!initial)
      localStorage.setItem("valefluxo_employee_draft", JSON.stringify(form));
  }, [form, initial]);
  useEffect(() => {
    if (!initial)
      localStorage.setItem(
        "valefluxo_workdays_draft",
        JSON.stringify(workDays),
      );
  }, [workDays, initial]);
  const set = (key: string, value: any) => setForm({ ...form, [key]: value });
  const recharge = useMemo(() => {
    const day = Math.min(
        Number(form.rechargeDay) || 1,
        new Date(
          todayDate.getFullYear(),
          todayDate.getMonth() + 1,
          0,
        ).getDate(),
      ),
      d = new Date(todayDate.getFullYear(), todayDate.getMonth(), day, 12);
    return d.toISOString().slice(0, 10);
  }, [form.rechargeDay]);
  const credit = useMemo(() => {
    const d = new Date(recharge + "T12:00:00");
    d.setDate(d.getDate() + form.advance);
    return d.toISOString().slice(0, 10);
  }, [recharge, form.advance]);
  const calculation = useMemo(() => {
    const start = new Date(credit + "T12:00:00"),
      end = new Date(start);
    end.setMonth(end.getMonth() + 1);
    let days = 0;
    for (const d = new Date(start); d < end; d.setDate(d.getDate() + 1))
      if (
        scheduleType === "12x36"
          ? Math.abs(
              Math.floor(
                (d.getTime() -
                  new Date(scheduleStartDate + "T12:00:00").getTime()) /
                  86400000,
              ),
            ) %
              2 ===
            0
          : workDays.includes(d.getDay())
      )
        days++;
    const fare1 = parseMoney(form.cardFare),
      fare2 = form.hasSecondCard ? parseMoney(form.secondCardFare) : 0;
    return {
      days,
      fare1,
      fare2,
      amount1: days * fare1,
      amount2: days * fare2,
      total: days * (fare1 + fare2),
      next: end.toISOString().slice(0, 10),
    };
  }, [
    credit,
    form.cardFare,
    form.secondCardFare,
    form.hasSecondCard,
    workDays,
    scheduleType,
    scheduleStartDate,
  ]);
  const dayOccupancy = (day: number) => {
    const names = new Set(
      rows
        .filter(
          (r) =>
            r.id !== initial?.id &&
            r.store === form.store &&
            new Date(r.rechargeDate + "T12:00:00").getDate() === day,
        )
        .map((r) => r.employee),
    );
    return names.size;
  };
  const selectedDayFull =
    !!form.store && dayOccupancy(Number(form.rechargeDay)) >= 2;
  const weekdays = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-slate-950/40 p-4 backdrop-blur-sm">
      <div className="fade-in max-h-[92vh] w-full max-w-3xl overflow-y-auto rounded-2xl bg-white shadow-2xl">
        <div className="sticky top-0 z-10 flex items-center border-b border-slate-100 bg-white px-6 py-5">
          <div>
            <h3 className="text-lg font-bold">
              {initial ? "Editar funcionário" : "Novo funcionário"}
            </h3>
            <p className="text-xs text-slate-400">
              Cadastro, escala de trabalho e cálculo do benefício
            </p>
          </div>
          <button
            onClick={close}
            className="ml-auto rounded-lg p-2 text-slate-400 hover:bg-slate-100"
          >
            <X size={20} />
          </button>
        </div>
        <form
          onSubmit={(e) => {
            e.preventDefault();
            if (form.receivesTransit && selectedDayFull) return;
            add({
              id: initial?.id ?? Date.now(),
              employee: form.name,
              cpf: form.cpf,
              role: form.role,
              store: form.store,
              phone: initial?.phone || "",
              hiredAt: form.hiredAt,
              cardType: form.cardType,
              card: "",
              cardDailyFare: calculation.fare1,
              cardAmount: calculation.amount1,
              secondCardType: form.hasSecondCard
                ? form.secondCardType
                : undefined,
              secondCard: undefined,
              secondCardDailyFare: form.hasSecondCard
                ? calculation.fare2
                : undefined,
              secondCardAmount: form.hasSecondCard
                ? calculation.amount2
                : undefined,
              creditDate: credit,
              advance: form.advance,
              rechargeDate: recharge,
              status:
                initial?.status === "Recarregado"
                  ? "Recarregado"
                  : new Date(recharge) < new Date(new Date().toDateString())
                    ? "Atrasado"
                    : recharge === today
                      ? "Pendente"
                      : "Próximo",
              dailyFare: calculation.fare1 + calculation.fare2,
              scheduleType,
              scheduleStartDate,
              workDays,
              amount: calculation.total,
              active:
                form.employmentStatus === "Desligado" ? false : form.active,
              birthDate: form.birthDate,
              gender: form.gender || undefined,
              experienceDays: 90,
              experienceCritical: form.experienceCritical,
              priorityTerminationDate: undefined,
              employmentStatus: form.employmentStatus,
              noticeStart: form.noticeStart,
              noticeEnd: form.noticeEnd,
              terminationDate: form.terminationDate,
              formalEmployment: form.formalEmployment,
              unregisteredStartDate: form.formalEmployment
                ? undefined
                : form.unregisteredStartDate,
              unregisteredReason: form.formalEmployment
                ? undefined
                : form.unregisteredReason,
              receivesTransit: form.receivesCostAssistance
                ? false
                : form.receivesTransit,
              receivesCostAssistance: form.receivesTransit
                ? false
                : form.receivesCostAssistance,
              costAssistanceAmount: form.receivesCostAssistance
                ? parseMoney(form.costAssistanceAmount)
                : 0,
            });
            localStorage.removeItem("valefluxo_employee_draft");
            localStorage.removeItem("valefluxo_workdays_draft");
            close();
          }}
          className="p-6"
        >
          <section className="rounded-2xl border border-slate-200 bg-slate-50/70 p-5">
          <div className="mb-1 text-xs font-bold uppercase tracking-wider text-forest-700">
            Identificação
          </div>
          <p className="mb-4 text-xs text-slate-400">Dados pessoais do funcionário.</p>
          <div className="grid gap-4 sm:grid-cols-2">
            <Field label="Nome completo">
              <input
                required
                value={form.name}
                onChange={(e) => set("name", e.target.value)}
                placeholder="Nome do funcionário"
              />
            </Field>
            <Field label="CPF">
              <input
                required
                inputMode="numeric"
                maxLength={14}
                value={form.cpf}
                onChange={(e) => set("cpf", formatCpf(e.target.value))}
                placeholder="000.000.000-00"
              />
            </Field>
            <Field label="Data de nascimento">
              <input
                type="date"
                value={form.birthDate}
                onChange={(e) => set("birthDate", e.target.value)}
              />
            </Field>
            <Field label="Sexo">
              <select
                required
                value={form.gender}
                onChange={(e) => set("gender", e.target.value as Recharge["gender"])}
              >
                <option value="">Selecione</option>
                <option value="Masculino">Masculino</option>
                <option value="Feminino">Feminino</option>
              </select>
            </Field>
          </div>
          </section>
          <section className="mt-5 rounded-2xl border border-slate-200 bg-slate-50/70 p-5">
          <div className="mb-1 text-xs font-bold uppercase tracking-wider text-forest-700">
            Vínculo profissional
          </div>
          <p className="mb-4 text-xs text-slate-400">Cargo, unidade e informações da contratação.</p>
          <div className="grid gap-4 sm:grid-cols-2">
            <Field label="Cargo">
              <select
                required
                value={form.role}
                onChange={(e) => set("role", e.target.value)}
              >
                <option value="">Selecione um cargo</option>
                {positions.map((p) => (
                  <option key={p}>{p}</option>
                ))}
              </select>
            </Field>
            <Field label="Loja">
              <select
                required
                value={form.store}
                onChange={(e) => set("store", e.target.value)}
              >
                <option value="">Selecione uma loja</option>
                {stores.map((s) => (
                  <option key={s}>{s}</option>
                ))}
              </select>
            </Field>
            <Field label="Data de admissão">
              <input
                required
                type="date"
                value={form.hiredAt}
                onChange={(e) => set("hiredAt", e.target.value)}
              />
            </Field>
          </div>
          </section>
          <section className="mt-5 rounded-2xl border border-slate-200 bg-slate-50/70 p-5">
          <div className="mb-1 text-xs font-bold uppercase tracking-wider text-forest-700">
            Situação trabalhista
          </div>
          <p className="mb-4 text-xs text-slate-400">Carteira assinada e período de experiência.</p>
          <div className="grid gap-4 sm:grid-cols-2">
            {form.formalEmployment && <div className="sm:col-span-2 rounded-xl border border-slate-200 bg-slate-50 p-4">
              <div className="text-xs font-bold uppercase tracking-wider text-slate-500">
                Período de experiência
              </div>
              {(() => {
                const admission = new Date(form.hiredAt + "T12:00:00"),
                  first = new Date(admission),
                  final = new Date(admission);
                first.setDate(first.getDate() + 29);
                final.setDate(final.getDate() + 89);
                return (
                  <div className="mt-3 grid gap-3 sm:grid-cols-2">
                    <div className="rounded-lg bg-white p-3">
                      <div className="text-xs text-slate-500">
                        1º período — 30 dias
                      </div>
                      <b className="text-sm text-slate-900">
                        Até {first.toLocaleDateString("pt-BR")}
                      </b>
                    </div>
                    <div className="rounded-lg bg-white p-3">
                      <div className="text-xs text-slate-500">
                        Prorrogação — mais 60 dias
                      </div>
                      <b className="text-sm text-slate-900">
                        Final em {final.toLocaleDateString("pt-BR")}
                      </b>
                    </div>
                  </div>
                );
              })()}
            </div>}
            <Field label="Carteira assinada">
              <select
                value={form.formalEmployment ? "Sim" : "Não"}
                onChange={(e) =>
                  set("formalEmployment", e.target.value === "Sim")
                }
              >
                <option>Sim</option>
                <option>Não</option>
              </select>
            </Field>
            {!form.formalEmployment && (
              <>
                <Field label="Início sem carteira">
                  <input
                    required
                    type="date"
                    value={form.unregisteredStartDate}
                    onChange={(e) =>
                      set("unregisteredStartDate", e.target.value)
                    }
                  />
                </Field>
                <Field label="Motivo de estar sem carteira">
                  <select
                    required
                    value={form.unregisteredReason}
                    onChange={(e) => set("unregisteredReason", e.target.value)}
                  >
                    <option value="">Selecione o motivo</option>
                    {unregisteredReasons.map((reason) => (
                      <option key={reason}>{reason}</option>
                    ))}
                  </select>
                </Field>
              </>
            )}
          </div>
          </section>
          <section className="mt-5 rounded-2xl border border-slate-200 bg-slate-50/70 p-5">
            <div className="mb-1 text-xs font-bold uppercase tracking-wider text-forest-700">
              Benefícios
            </div>
            <p className="mb-4 text-xs text-slate-400">Ajuda de custo ou vale-transporte.</p>
            <div className="grid gap-4 sm:grid-cols-2">
              <Field label="Recebe ajuda de custo">
                <select
                  value={form.receivesCostAssistance ? "Sim" : "Não"}
                  onChange={(e) => {
                    const yes = e.target.value === "Sim";
                    setForm({
                      ...form,
                      receivesCostAssistance: yes,
                      receivesTransit: yes ? false : form.receivesTransit,
                      hasSecondCard: yes ? false : form.hasSecondCard,
                    });
                  }}
                >
                  <option>Sim</option>
                  <option>Não</option>
                </select>
              </Field>
              {form.receivesCostAssistance && (
                <Field label="Valor mensal da ajuda de custo (R$)">
                  <input
                    inputMode="numeric"
                    value={form.costAssistanceAmount}
                    onChange={(e) =>
                      set("costAssistanceAmount", formatMoneyInput(e.target.value))
                    }
                    placeholder="Ex.: 300,00"
                  />
                </Field>
              )}
              {!form.receivesCostAssistance && (
                <Field label="Recebe vale-transporte">
                  <select
                    value={form.receivesTransit ? "Sim" : "Não"}
                    onChange={(e) => {
                      const yes = e.target.value === "Sim";
                      setForm({
                        ...form,
                        receivesTransit: yes,
                        receivesCostAssistance: yes ? false : form.receivesCostAssistance,
                        costAssistanceAmount: yes ? "" : form.costAssistanceAmount,
                      });
                    }}
                  >
                    <option>Sim</option>
                    <option>Não</option>
                  </select>
                </Field>
              )}
            </div>
          </section>
          {form.receivesTransit && !form.receivesCostAssistance && (
            <>
              <div className="mb-1 mt-5 rounded-t-2xl border border-b-0 border-slate-200 bg-slate-50/70 px-5 pt-5 text-xs font-bold uppercase tracking-wider text-forest-700">
                Cartões e recarga
              </div>
              <div className="grid gap-4 rounded-b-2xl border border-t-0 border-slate-200 bg-slate-50/70 px-5 pb-5 sm:grid-cols-2">
                <Field label="Tipo do cartão 1">
                  <select
                    value={form.cardType}
                    onChange={(e) => set("cardType", e.target.value)}
                  >
                    <option>BHBus</option>
                    <option>Ótimo</option>
                    <option>Outro</option>
                  </select>
                </Field>
                <Field label={`Valor diário — ${form.cardType} (R$)`}>
                  <input
                    required={form.receivesTransit}
                    inputMode="numeric"
                    value={form.cardFare}
                    onChange={(e) =>
                      set("cardFare", formatMoneyInput(e.target.value))
                    }
                    placeholder="Ex.: 12,40"
                  />
                </Field>
              </div>
              <label className="mt-4 flex items-center gap-3 rounded-xl border border-slate-200 p-3 text-sm font-semibold text-slate-700">
                <input
                  type="checkbox"
                  checked={form.hasSecondCard}
                  onChange={(e) => set("hasSecondCard", e.target.checked)}
                  className="h-4 w-4 accent-forest-700"
                />
                Este funcionário utiliza um segundo cartão
              </label>
              {form.hasSecondCard && (
                <div className="mt-4 grid gap-4 rounded-xl bg-slate-50 p-4 sm:grid-cols-2">
                  <Field label="Tipo do cartão 2">
                    <select
                      value={form.secondCardType}
                      onChange={(e) => set("secondCardType", e.target.value)}
                    >
                      <option>BHBus</option>
                      <option>Ótimo</option>
                      <option>Outro</option>
                    </select>
                  </Field>
                  <Field label={`Valor diário — ${form.secondCardType} (R$)`}>
                    <input
                      required={form.receivesTransit}
                      inputMode="numeric"
                      value={form.secondCardFare}
                      onChange={(e) =>
                        set("secondCardFare", formatMoneyInput(e.target.value))
                      }
                      placeholder="Ex.: 10,50"
                    />
                  </Field>
                </div>
              )}
              <div className="mt-5 grid items-start gap-5 sm:grid-cols-[1.4fr_1fr]">
                <RechargeDayPicker
                  selected={Number(form.rechargeDay)}
                  onSelect={(day) => set("rechargeDay", String(day))}
                  occupancy={dayOccupancy}
                  store={form.store}
                />
                <div className="rounded-2xl bg-slate-50 p-4">
                  <Field label="Prazo para liberar o crédito">
                    <select
                      value={form.advance}
                      onChange={(e) => set("advance", Number(e.target.value))}
                    >
                      {[2, 3, 4, 5].map((n) => (
                        <option key={n} value={n}>
                          {n} dias depois
                        </option>
                      ))}
                    </select>
                  </Field>
                  <div className="mt-4 rounded-xl bg-white p-3 text-xs text-slate-500">
                    A recarga será feita todo{" "}
                    <b className="text-forest-800">
                      dia {String(form.rechargeDay).padStart(2, "0")}
                    </b>
                    . O crédito ficará disponível aproximadamente{" "}
                    <b className="text-forest-800">
                      {form.advance} dias depois
                    </b>
                    .
                  </div>
                </div>
              </div>
              {selectedDayFull && (
                <p className="mt-2 text-sm font-semibold text-red-600">
                  Este dia já possui dois funcionários para a loja {form.store}.
                  Escolha outro dia.
                </p>
              )}
            </>
          )}
          <div className="mt-6">
            <div className="text-xs font-bold uppercase tracking-wider text-forest-700">
              Calendário semanal de trabalho
            </div>
            <p className="mt-1 text-xs text-slate-400">
              Escolha uma escala pronta ou selecione os dias manualmente.
            </p>
            <div className="mt-3 flex flex-wrap gap-2">
              <button
                type="button"
                onClick={() => {
                  setScheduleType("6x1");
                  setWorkDays([1, 2, 3, 4, 5, 6]);
                }}
                className={`rounded-xl border px-4 py-2 text-sm font-bold ${scheduleType === "6x1" ? "border-slate-800 bg-slate-800 text-white" : "border-slate-300"}`}
              >
                Escala 6x1
              </button>
              <button
                type="button"
                onClick={() => {
                  setScheduleType("12x36");
                  setWorkDays([0, 1, 2, 3, 4, 5, 6]);
                }}
                className={`rounded-xl border px-4 py-2 text-sm font-bold ${scheduleType === "12x36" ? "border-slate-800 bg-slate-800 text-white" : "border-slate-300"}`}
              >
                Escala 12x36
              </button>
              <button
                type="button"
                onClick={() => setScheduleType("Personalizada")}
                className={`rounded-xl border px-4 py-2 text-sm font-bold ${scheduleType === "Personalizada" ? "border-slate-800 bg-slate-800 text-white" : "border-slate-300"}`}
              >
                Personalizada
              </button>
            </div>
            {scheduleType === "12x36" ? (
              <div className="mt-3 rounded-xl border border-slate-200 bg-slate-50 p-4">
                <label className="text-sm font-semibold text-slate-700">
                  Data do primeiro plantão
                  <input
                    required
                    type="date"
                    value={scheduleStartDate}
                    onChange={(event) =>
                      setScheduleStartDate(event.target.value)
                    }
                    className="mt-2 w-full rounded-xl border border-slate-200 bg-white px-4 py-3"
                  />
                </label>
                <p className="mt-2 text-xs text-slate-500">
                  A partir desta data, o sistema alternará 12 horas de trabalho
                  por 36 horas de descanso.
                </p>
              </div>
            ) : (
              <div className="mt-3 grid grid-cols-4 gap-2 sm:grid-cols-7">
                {weekdays.map((day, i) => (
                  <button
                    type="button"
                    key={day}
                    onClick={() => {
                      setScheduleType("Personalizada");
                      setWorkDays(
                        workDays.includes(i)
                          ? workDays.filter((x) => x !== i)
                          : [...workDays, i],
                      );
                    }}
                    className={`rounded-xl border px-2 py-3 text-xs font-bold transition ${workDays.includes(i) ? "border-forest-700 bg-forest-700 text-white" : "border-slate-200 text-slate-500 hover:border-forest-300"}`}
                  >
                    <CalendarDays className="mx-auto mb-1" size={16} />
                    {day}
                  </button>
                ))}
              </div>
            )}
          </div>
          {form.receivesTransit && !form.receivesCostAssistance ? (
            <>
              <div className="mt-6 rounded-2xl border border-forest-100 bg-forest-50 p-4">
                <div className="grid gap-4 sm:grid-cols-3">
                  <div>
                    <div className="text-xs text-slate-500">
                      Dia mensal da recarga
                    </div>
                    <b className="text-lg text-forest-800">
                      Dia {String(form.rechargeDay).padStart(2, "0")}
                    </b>
                    <div className="text-[11px] text-slate-500">
                      Crédito previsto: {formatDate(credit)}
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500">
                      Dias trabalhados
                    </div>
                    <b className="text-lg text-forest-800">
                      {calculation.days} dias
                    </b>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500">
                      Gasto diário total
                    </div>
                    <b className="text-lg text-forest-800">
                      {(calculation.fare1 + calculation.fare2).toLocaleString(
                        "pt-BR",
                        { style: "currency", currency: "BRL" },
                      )}
                    </b>
                  </div>
                </div>
                <div className="mt-4 grid gap-2 border-t border-forest-200 pt-4 sm:grid-cols-2">
                  {" "}
                  <div className="rounded-xl bg-white/70 p-3">
                    <div className="text-xs font-semibold text-slate-500">
                      {form.cardType}
                    </div>
                    <div className="mt-1 text-sm">
                      Por dia:{" "}
                      <b>
                        {calculation.fare1.toLocaleString("pt-BR", {
                          style: "currency",
                          currency: "BRL",
                        })}
                      </b>
                    </div>
                    <div className="text-sm">
                      Total mensal:{" "}
                      <b className="text-forest-800">
                        {calculation.amount1.toLocaleString("pt-BR", {
                          style: "currency",
                          currency: "BRL",
                        })}
                      </b>
                    </div>
                  </div>
                  {form.hasSecondCard && (
                    <div className="rounded-xl bg-white/70 p-3">
                      <div className="text-xs font-semibold text-slate-500">
                        {form.secondCardType}
                      </div>
                      <div className="mt-1 text-sm">
                        Por dia:{" "}
                        <b>
                          {calculation.fare2.toLocaleString("pt-BR", {
                            style: "currency",
                            currency: "BRL",
                          })}
                        </b>
                      </div>
                      <div className="text-sm">
                        Total mensal:{" "}
                        <b className="text-forest-800">
                          {calculation.amount2.toLocaleString("pt-BR", {
                            style: "currency",
                            currency: "BRL",
                          })}
                        </b>
                      </div>
                    </div>
                  )}
                </div>
                <div className="mt-4 flex items-center justify-between rounded-xl bg-forest-800 px-4 py-3 text-white">
                  <span className="text-sm font-semibold">
                    Total geral do mês
                  </span>
                  <b className="text-xl">
                    {calculation.total.toLocaleString("pt-BR", {
                      style: "currency",
                      currency: "BRL",
                    })}
                  </b>
                </div>
                <div className="mt-3 text-[11px] text-slate-500">
                  Ciclo de{" "}
                  {new Date(credit + "T12:00:00").toLocaleDateString("pt-BR")}{" "}
                  até{" "}
                  {new Date(calculation.next + "T12:00:00").toLocaleDateString(
                    "pt-BR",
                  )}
                  , considerando somente os dias selecionados.
                </div>
              </div>
            </>
          ) : form.receivesCostAssistance ? (
            <div className="mt-6 rounded-2xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex items-center justify-between gap-4">
                <div>
                  <div className="text-xs font-bold uppercase tracking-wider text-slate-500">
                    Benefício selecionado
                  </div>
                  <h4 className="mt-1 text-lg font-bold text-slate-900">
                    Ajuda de custo
                  </h4>
                  <p className="mt-1 text-xs text-slate-500">
                    Este funcionário não utiliza cartão de passagem.
                  </p>
                </div>
                <div className="rounded-xl bg-[#262626] px-5 py-4 text-right text-white">
                  <div className="text-xs text-white/70">Valor mensal</div>
                  <b className="text-xl">
                    {parseMoney(form.costAssistanceAmount).toLocaleString(
                      "pt-BR",
                      { style: "currency", currency: "BRL" },
                    )}
                  </b>
                </div>
              </div>
            </div>
          ) : null}
          <div className="mt-6 flex justify-end gap-3">
            <button
              type="button"
              onClick={close}
              className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600"
            >
              Cancelar
            </button>
            <button
              disabled={
                !stores.length ||
                !positions.length ||
                !workDays.length ||
                (form.receivesTransit && selectedDayFull)
              }
              className="rounded-xl bg-forest-700 px-5 py-2.5 text-sm font-semibold text-white disabled:bg-slate-300"
            >
              {initial ? "Salvar alterações" : "Salvar funcionário"}
            </button>
          </div>
          {(!stores.length || !positions.length) && (
            <p className="mt-2 text-right text-xs text-red-600">
              Cadastre pelo menos uma loja e um cargo antes de adicionar o
              funcionário.
            </p>
          )}
        </form>
      </div>
    </div>
  );
}
function Field({ label, children }: { label: string; children: any }) {
  return (
    <label className="text-xs font-semibold text-slate-600">
      {label}
      <div className="mt-1.5 [&>*]:w-full [&>*]:rounded-xl [&>*]:border [&>*]:border-slate-200 [&>*]:bg-white [&>*]:px-3 [&>*]:py-2.5 [&>*]:text-sm [&>*]:font-normal [&>*]:outline-none focus-within:[&>*]:border-forest-500">
        {children}
      </div>
    </label>
  );
}

function RechargeDateModal({
  record,
  close,
  confirm,
}: {
  record: Recharge;
  close: () => void;
  confirm: (date: string, days: number, total: number) => void;
}) {
  const today = new Date().toISOString().slice(0, 10);
  const [date, setDate] = useState(record.completedDate || today);
  const actual = new Date(date + "T12:00:00"),
    available = new Date(actual);
  available.setDate(available.getDate() + record.advance);
  const nextRecharge = new Date(record.rechargeDate + "T12:00:00");
  nextRecharge.setMonth(nextRecharge.getMonth() + 1);
  const nextCredit = new Date(nextRecharge);
  nextCredit.setDate(nextCredit.getDate() + record.advance);
  let suggestedDays = 0;
  for (
    const d = new Date(available);
    d < nextCredit;
    d.setDate(d.getDate() + 1)
  )
    if (isWorkDay(record, d)) suggestedDays++;
  const [days, setDays] = useState(record.chargedDays ?? suggestedDays),
    [mode, setMode] = useState<"days" | "amount">("days"),
    [manualAmount, setManualAmount] = useState(String(record.amount || ""));
  const fare1 = record.cardDailyFare ?? record.dailyFare ?? 0,
    fare2 = record.secondCardDailyFare ?? 0,
    dailyTotal = fare1 + fare2,
    enteredAmount = Number(manualAmount.replace(",", ".")) || 0,
    effectiveDays =
      mode === "amount" && dailyTotal > 0
        ? Math.floor(enteredAmount / dailyTotal)
        : days,
    total = mode === "amount" ? enteredAmount : effectiveDays * dailyTotal,
    amount1 = dailyTotal ? total * (fare1 / dailyTotal) : 0,
    amount2 = dailyTotal ? total * (fare2 / dailyTotal) : 0;
  const lastsUntil = new Date(available);
  let counted = 0;
  while (counted < effectiveDays) {
    if (isWorkDay(record, lastsUntil)) counted++;
    if (counted < effectiveDays) lastsUntil.setDate(lastsUntil.getDate() + 1);
  }
  const money = (v: number) =>
    v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  return (
    <div className="fixed inset-0 z-[60] grid place-items-center bg-slate-950/45 p-4 backdrop-blur-sm">
      <div className="fade-in max-h-[92vh] w-full max-w-md overflow-y-auto rounded-2xl bg-white p-6 shadow-2xl">
        <div className="flex items-start">
          <div>
            <h3 className="text-lg font-bold">
              {record.completedDate ? "Editar recarga" : "Confirmar recarga"}
            </h3>
            <p className="mt-1 text-sm text-slate-500">{record.employee}</p>
          </div>
          <button
            onClick={close}
            className="ml-auto rounded-lg p-2 text-slate-400 hover:bg-slate-100"
          >
            <X size={19} />
          </button>
        </div>
        <label className="mt-6 block text-sm font-semibold text-slate-700">
          Em qual dia a recarga foi realizada?
          <input
            autoFocus
            required
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 outline-none focus:border-forest-500"
          />
        </label>
        <div className="mt-4 grid grid-cols-2 rounded-xl bg-slate-100 p-1">
          <button
            onClick={() => setMode("days")}
            className={`rounded-lg py-2 text-xs font-bold ${mode === "days" ? "bg-white text-forest-700 shadow-sm" : "text-slate-500"}`}
          >
            Informar dias
          </button>
          <button
            onClick={() => setMode("amount")}
            className={`rounded-lg py-2 text-xs font-bold ${mode === "amount" ? "bg-white text-forest-700 shadow-sm" : "text-slate-500"}`}
          >
            Informar valor
          </button>
        </div>
        {mode === "days" ? (
          <label className="mt-4 block text-sm font-semibold text-slate-700">
            Número de dias que serão recarregados
            <input
              required
              min="0"
              max="62"
              type="number"
              value={days}
              onChange={(e) => setDays(Math.max(0, Number(e.target.value)))}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-lg font-bold outline-none focus:border-forest-500"
            />
          </label>
        ) : (
          <label className="mt-4 block text-sm font-semibold text-slate-700">
            Valor total da recarga (R$)
            <input
              required
              min="0"
              step="0.01"
              inputMode="decimal"
              value={manualAmount}
              onChange={(e) => setManualAmount(e.target.value)}
              placeholder="Ex.: 200,00"
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-lg font-bold outline-none focus:border-forest-500"
            />
          </label>
        )}
        <div className="mt-5 rounded-xl bg-forest-50 p-4 text-sm">
          <div className="flex justify-between">
            <span className="text-slate-500">Crédito disponível</span>
            <b className="text-forest-800">
              {available.toLocaleDateString("pt-BR")}
            </b>
          </div>
          <div className="mt-2 flex justify-between">
            <span className="text-slate-500">Duração estimada</span>
            <b className="text-forest-800">
              {effectiveDays} dias · até{" "}
              {effectiveDays ? lastsUntil.toLocaleDateString("pt-BR") : "—"}
            </b>
          </div>
          <div className="my-3 border-t border-forest-100" />
          <div className="flex justify-between">
            <span className="text-slate-500">
              {record.cardType}: {money(fare1)}/dia
            </span>
            <b>{money(amount1)}</b>
          </div>
          {record.secondCardType && (
            <div className="mt-2 flex justify-between">
              <span className="text-slate-500">
                {record.secondCardType}: {money(fare2)}/dia
              </span>
              <b>{money(amount2)}</b>
            </div>
          )}
          <div className="mt-3 flex justify-between rounded-lg bg-forest-800 px-3 py-2.5 text-white">
            <span className="font-semibold">Total da recarga</span>
            <b className="text-lg">{money(total)}</b>
          </div>
          <p className="mt-3 text-xs text-slate-500">
            Sugestão pela escala: {suggestedDays} dias. Você pode ajustar antes
            de confirmar.
          </p>
        </div>
        <div className="mt-6 flex justify-end gap-2">
          <button
            onClick={close}
            className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600"
          >
            Cancelar
          </button>
          <button
            disabled={total <= 0}
            onClick={() => confirm(date, effectiveDays, total)}
            className="rounded-xl bg-forest-700 px-4 py-2.5 text-sm font-semibold text-white disabled:bg-slate-300"
          >
            {record.completedDate ? "Salvar alterações" : "Confirmar recarga"}
          </button>
        </div>
      </div>
    </div>
  );
}

function TerminationModal({
  employee,
  close,
  confirm,
}: {
  employee: Recharge;
  close: () => void;
  confirm: (
    terminationDate: string,
    notice: boolean,
    noticeStart: string,
    noticeEnd: string,
  ) => void;
}) {
  const today = new Date().toISOString().slice(0, 10),
    [date, setDate] = useState(today),
    [notice, setNotice] = useState(false),
    noticeEnd = (() => {
      const value = new Date(date + "T12:00:00");
      value.setDate(value.getDate() + 30);
      return value;
    })(),
    noticeEndIso = noticeEnd.toISOString().slice(0, 10),
    firstExperienceEnd = employee.hiredAt
      ? (() => {
          const value = new Date(employee.hiredAt + "T12:00:00");
          value.setDate(value.getDate() + 29);
          return value;
        })()
      : null,
    experienceEnd = employee.hiredAt
      ? (() => {
          const value = new Date(employee.hiredAt + "T12:00:00");
          value.setDate(value.getDate() + 89);
          return value;
        })()
      : null,
    inExperience =
      employee.formalEmployment !== false &&
      !!experienceEnd &&
      new Date(date + "T12:00:00") <= experienceEnd;
  return createPortal(
    <div className="fixed inset-0 z-[90] grid place-items-center bg-slate-950/50 p-4 backdrop-blur-sm">
      <form
        onSubmit={(event) => {
          event.preventDefault();
          confirm(
            date,
            inExperience ? false : notice,
            inExperience ? "" : date,
            inExperience ? "" : noticeEndIso,
          );
        }}
        className="w-full max-w-lg rounded-2xl bg-white p-6 shadow-2xl"
      >
        <div className="flex items-start">
          <div>
            <h3 className="text-xl font-bold text-slate-900">
              Demitir funcionário
            </h3>
            <p className="mt-1 text-sm text-slate-500">{employee.employee}</p>
          </div>
          <button
            type="button"
            onClick={close}
            className="ml-auto rounded-lg p-2 text-slate-500 hover:bg-slate-100"
          >
            <X size={20} />
          </button>
        </div>
        <label className="mt-6 block text-sm font-semibold text-slate-700">
          Data do desligamento
          <input
            required
            type="date"
            value={date}
            onChange={(event) => setDate(event.target.value)}
            className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3"
          />
        </label>
        {inExperience ? (
          <div className="mt-4 rounded-xl bg-amber-50 p-4 text-sm text-amber-800">
            <b>Funcionário em período de experiência.</b>
            <div className="mt-1 text-xs">
              Não há aviso prévio. Informe somente a data do desligamento.
            </div>
            <div className="mt-3 grid gap-2 rounded-lg bg-white/70 p-3 text-xs">
              <div className="flex justify-between">
                <span>Primeiros 30 dias:</span>
                <b>{firstExperienceEnd?.toLocaleDateString("pt-BR")}</b>
              </div>
              <div className="flex justify-between">
                <span>Prorrogação de 60 dias:</span>
                <b>{experienceEnd?.toLocaleDateString("pt-BR")}</b>
              </div>
              <div className="border-t border-amber-200 pt-2 font-bold">
                {firstExperienceEnd &&
                new Date(date + "T12:00:00") <= firstExperienceEnd
                  ? "Desligamento no período inicial de 30 dias"
                  : "Desligamento durante a prorrogação de 60 dias"}
              </div>
            </div>
          </div>
        ) : (
          <label className="mt-4 block text-sm font-semibold text-slate-700">
            Vai cumprir aviso?
            <select
              value={notice ? "Sim" : "Não"}
              onChange={(event) => setNotice(event.target.value === "Sim")}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3"
            >
              <option>Não</option>
              <option>Sim</option>
            </select>
          </label>
        )}
        {!inExperience && notice && (
          <div className="mt-4 grid gap-3 rounded-xl bg-slate-50 p-4 sm:grid-cols-2">
            <div>
              <div className="text-xs text-slate-500">Início do aviso</div>
              <b className="mt-1 block text-sm text-slate-900">
                {new Date(date + "T12:00:00").toLocaleDateString("pt-BR")}
              </b>
            </div>
            <div>
              <div className="text-xs text-slate-500">Término do aviso</div>
              <b className="mt-1 block text-sm text-slate-900">
                {noticeEnd.toLocaleDateString("pt-BR")}
              </b>
            </div>
          </div>
        )}
        <div className="mt-6 flex justify-end gap-2">
          <button
            type="button"
            onClick={close}
            className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold"
          >
            Cancelar
          </button>
          <button className="rounded-xl bg-red-600 px-4 py-2.5 text-sm font-semibold text-white">
            Confirmar desligamento
          </button>
        </div>
      </form>
    </div>,
    document.body,
  );
}
function SectionHead({
  title,
  sub,
  action,
}: {
  title: string;
  sub: string;
  action?: any;
}) {
  return (
    <div className="mb-6 flex flex-col gap-3 sm:flex-row sm:items-center">
      <div>
        <h2 className="text-2xl font-bold text-slate-900">{title}</h2>
        <p className="text-sm text-slate-500">{sub}</p>
      </div>
      {action && <div className="sm:ml-auto">{action}</div>}
    </div>
  );
}
function HREmployeesPage({
  rows,
  openForm,
  edit,
  toggleCritical,
  dismiss,
  readOnly = false,
  manage,
}: {
  rows: Recharge[];
  openForm?: () => void;
  edit?: (r: Recharge) => void;
  toggleCritical?: (r: Recharge) => void;
  dismiss?: (
    r: Recharge,
    date: string,
    notice: boolean,
    start: string,
    end: string,
  ) => void;
  readOnly?: boolean;
  manage?: (employee: Recharge) => void;
}) {
  const [query, setQuery] = useState(""),
    [storeFilter, setStoreFilter] = useState("Todas"),
    [dismissing, setDismissing] = useState<Recharge | null>(null),
    [viewing, setViewing] = useState<Recharge | null>(null);
  const stores = [...new Set(rows.map((r) => r.store))].sort(),
    list = rows
      .filter((r) => !isEmployeeDismissed(r))
      .filter(
        (r) =>
          (storeFilter === "Todas" || r.store === storeFilter) &&
          `${r.employee} ${r.role} ${r.store}`
            .toLowerCase()
            .includes(query.toLowerCase()),
      );
  const experiencePeriods = (r: Recharge) => {
    if (!r.hiredAt || r.formalEmployment === false) return null;
    const admission = new Date(r.hiredAt + "T12:00:00"),
      first = new Date(admission),
      final = new Date(admission);
    first.setDate(first.getDate() + 29);
    final.setDate(final.getDate() + 89);
    return {
      first: first.toLocaleDateString("pt-BR"),
      final: final.toLocaleDateString("pt-BR"),
    };
  };
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Funcionários"
        sub={`${list.length} de ${rows.length} funcionário(s)`}
        action={!readOnly && openForm ? (
          <button
            onClick={openForm}
            className="flex items-center gap-2 rounded-xl bg-forest-700 px-5 py-3 text-sm font-bold text-white shadow-lg"
          >
            <Plus size={18} /> Novo funcionário
          </button>
        ) : undefined}
      />
      <div className="rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="grid gap-3 border-b border-slate-100 p-5 sm:grid-cols-[1fr_260px]">
          <div className="relative">
            <Search
              className="absolute left-3 top-3 text-slate-400"
              size={17}
            />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Pesquisar funcionário, cargo ou loja..."
              className="w-full rounded-xl border border-slate-200 py-2.5 pl-9 pr-3 text-sm outline-none focus:border-slate-500"
            />
          </div>
          <div className="relative">
            <Store
              className="pointer-events-none absolute left-3 top-3 text-slate-500"
              size={17}
            />
            <select
              value={storeFilter}
              onChange={(e) => setStoreFilter(e.target.value)}
              className="w-full rounded-xl border border-slate-200 py-2.5 pl-9 pr-3 text-sm"
            >
              <option>Todas</option>
              {stores.map((s) => (
                <option key={s}>{s}</option>
              ))}
            </select>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[920px] text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase text-slate-400">
              <tr>
                {[
                  "Funcionário",
                  "Loja",
                  "Cargo",
                  "Admissão",
                  "Situação",
                  "Prioridade",
                  "Ações",
                ].map((h) => (
                  <th key={h} className="px-5 py-3">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {list.map((r) => (
                <tr
                  key={r.id}
                  className={r.experienceCritical ? "bg-red-50/60" : ""}
                >
                  <td className="px-5 py-4 font-semibold">
                    {r.employee}
                    <div className="text-xs font-normal text-slate-400">
                      {r.cpf || "CPF não informado"}
                    </div>
                  </td>
                  <td className="px-5 py-4">{r.store}</td>
                  <td className="px-5 py-4">{r.role}</td>
                  <td className="px-5 py-4">
                    {r.hiredAt
                      ? new Date(r.hiredAt + "T12:00:00").toLocaleDateString(
                          "pt-BR",
                        )
                      : "-"}
                  </td>
                  <td className="px-5 py-4">
                    {isDismissalPending(r) ? (
                      <span className="font-semibold text-amber-600">
                        Desligamento em andamento
                      </span>
                    ) : (
                      r.employmentStatus || "Ativo"
                    )}
                  </td>
                  <td className="px-5 py-4">
                    {r.experienceCritical ? (
                      <span className="rounded-full bg-red-100 px-2.5 py-1 text-xs font-bold text-red-700">
                        Crítico
                      </span>
                    ) : (
                      <span className="text-xs text-slate-400">Normal</span>
                    )}
                  </td>
                  <td className="px-5 py-4">
                    {readOnly ? (
                      <div className="flex gap-2">
                        <button onClick={() => setViewing(r)} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 px-3 py-2 text-xs font-bold text-slate-700 hover:bg-slate-50">
                          <Eye size={16} /> Visualizar
                        </button>
                        {toggleCritical && (
                          <button
                            onClick={() => toggleCritical(r)}
                            className={`rounded-lg border px-3 py-2 text-xs font-bold ${r.experienceCritical ? "border-slate-300 text-slate-600 hover:bg-slate-50" : "border-red-200 text-red-600 hover:bg-red-50"}`}
                          >
                            {r.experienceCritical ? "Retirar crítico" : "Marcar crítico"}
                          </button>
                        )}
                        {manage && <button onClick={() => manage(r)} className="inline-flex items-center gap-2 rounded-lg bg-slate-800 px-3 py-2 text-xs font-bold text-white hover:bg-slate-700">
                          <Pencil size={15} /> Editar
                        </button>}
                      </div>
                    ) : <div className="flex gap-2">
                      <button
                        onClick={() => toggleCritical?.(r)}
                        className={`rounded-lg border px-3 py-2 text-xs font-semibold ${r.experienceCritical ? "border-slate-300 text-slate-600" : "border-red-200 text-red-600"}`}
                      >
                        {r.experienceCritical
                          ? "Retirar crítico"
                          : "Marcar crítico"}
                      </button>
                      <button
                        onClick={() => edit?.(r)}
                        className="rounded-lg border border-slate-300 px-3 py-2 text-xs font-semibold text-slate-700"
                      >
                        Editar
                      </button>
                      {r.active !== false &&
                        r.employmentStatus !== "Desligado" && (
                          <button
                            onClick={() => setDismissing(r)}
                            className="rounded-lg border border-red-300 px-3 py-2 text-xs font-semibold text-red-600"
                          >
                            Demitir funcionário
                          </button>
                        )}
                    </div>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {!list.length && (
            <div className="py-14 text-center text-sm text-slate-400">
              Nenhum funcionário encontrado.
            </div>
          )}
        </div>
      </div>
      {!readOnly && dismissing && dismiss && (
        <TerminationModal
          employee={dismissing}
          close={() => setDismissing(null)}
          confirm={(date, notice, start, end) => {
            dismiss(dismissing, date, notice, start, end);
            setDismissing(null);
          }}
        />
      )}
      {viewing && <EmployeeDetailsModal employee={viewing} close={() => setViewing(null)} />}
    </main>
  );
}

function EmployeeDetailsModal({ employee, close }: { employee: Recharge; close: () => void }) {
  const money = (value = 0) => value.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  const benefit = employee.receivesCostAssistance
    ? `Ajuda de custo${employee.costAssistanceAmount ? ` · ${money(employee.costAssistanceAmount)}` : ""}`
    : employee.receivesTransit === false ? "Sem vale-transporte" : "Vale-transporte";
  const details = [
    ["Nome completo", employee.employee], ["CPF", employee.cpf || "Não informado"],
    ["Loja", employee.store], ["Função", employee.role],
    ["Data de admissão", employee.hiredAt ? formatDate(employee.hiredAt) : "Não informada"],
    ["Data de nascimento", employee.birthDate ? formatDate(employee.birthDate) : "Não informada"],
    ["Sexo", employee.gender || "Não informado"],
    ["Situação", isDismissalPending(employee) ? "Desligamento em andamento" : employee.employmentStatus || "Ativo"],
    ["Carteira", employee.formalEmployment === false ? "Sem carteira assinada" : "Carteira assinada"],
    ["Benefício", benefit], ["Escala", employee.scheduleType || "Personalizada"],
    ["Prioridade", employee.experienceCritical ? "Crítico" : "Normal"],
    ["Desligamento", employee.terminationDate ? formatDate(employee.terminationDate) : "Não informado"],
  ];
  return createPortal(
    <div className="fixed inset-0 z-[100] grid place-items-center bg-slate-950/50 p-4 backdrop-blur-sm" onMouseDown={close}>
      <div className="max-h-[88vh] w-full max-w-3xl overflow-y-auto rounded-2xl bg-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
        <div className="flex items-start border-b border-slate-200 px-6 py-5"><div><h3 className="text-xl font-bold text-slate-900">{employee.employee}</h3><p className="text-sm text-slate-500">Ficha completa do funcionário</p></div><button onClick={close} className="ml-auto rounded-lg p-2 text-slate-500 hover:bg-slate-100"><X size={20} /></button></div>
        <div className="grid gap-3 p-6 sm:grid-cols-2">{details.map(([label, value]) => <div key={label} className="rounded-xl border border-slate-200 bg-slate-50 p-4"><div className="text-xs font-bold uppercase tracking-wide text-slate-400">{label}</div><div className="mt-1 font-semibold text-slate-800">{value}</div></div>)}</div>
        {(employee.receivesTransit !== false && !employee.receivesCostAssistance) && <div className="mx-6 mb-6 rounded-xl border border-slate-200 p-4"><h4 className="font-bold text-slate-900">Cartões de passagem</h4><div className="mt-3 grid gap-3 sm:grid-cols-2"><div className="rounded-lg bg-slate-50 p-3"><b>{employee.cardType || "Cartão principal"}</b><div className="text-sm text-slate-500">Custo diário: {money(employee.cardDailyFare ?? employee.dailyFare ?? 0)}</div></div>{employee.secondCardType && <div className="rounded-lg bg-slate-50 p-3"><b>{employee.secondCardType}</b><div className="text-sm text-slate-500">Custo diário: {money(employee.secondCardDailyFare || 0)}</div></div>}</div></div>}
      </div>
    </div>, document.body,
  );
}

function EmployeesPage({
  rows,
  openForm,
  remove,
  edit,
}: {
  rows: Recharge[];
  openForm: () => void;
  remove: (id: number) => void;
  edit: (r: Recharge) => void;
}) {
  const [q, setQ] = useState("");
  const list = rows.filter(
    (r, i) =>
      rows.findIndex((x) => x.employee === r.employee) === i &&
      r.employee.toLowerCase().includes(q.toLowerCase()),
  );
  const money = (v = 0) =>
    v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Funcionários"
        sub={`${list.length} funcionários cadastrados`}
        action={
          <button
            onClick={openForm}
            className="flex items-center gap-2 rounded-xl bg-forest-700 px-4 py-2.5 text-sm font-semibold text-white"
          >
            <Plus size={17} />
            Novo funcionário
          </button>
        }
      />
      <div className="rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="relative m-5 max-w-sm">
          <Search
            className="absolute left-3 top-2.5 text-slate-400"
            size={17}
          />
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Pesquisar por nome..."
            className="w-full rounded-xl border border-slate-200 py-2 pl-9 pr-3 text-sm outline-none focus:border-forest-500"
          />
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[950px] text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase text-slate-400">
              <tr>
                {[
                  "Funcionário",
                  "Loja",
                  "Cartões",
                  "Custo diário",
                  "Totais mensais",
                  "Situação",
                  "Ações",
                ].map((h) => (
                  <th className="px-5 py-3" key={h}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {list.map((r) => (
                <tr key={r.id}>
                  <td className="px-5 py-4">
                    <b>{r.employee}</b>
                    <div className="text-xs text-slate-400">{r.role}</div>
                  </td>
                  <td className="px-5 py-4">{r.store}</td>
                  <td className="px-5 py-4">
                    <div>
                      <b>{r.cardType}</b>
                    </div>
                    {r.secondCardType && (
                      <div className="mt-1">
                        <b>{r.secondCardType}</b>
                      </div>
                    )}
                  </td>
                  <td className="px-5 py-4">
                    <div>
                      {r.cardType}:{" "}
                      <b>{money(r.cardDailyFare ?? r.dailyFare)}</b>
                    </div>
                    {r.secondCardType && (
                      <div>
                        {r.secondCardType}:{" "}
                        <b>{money(r.secondCardDailyFare)}</b>
                      </div>
                    )}
                    <div className="mt-1 border-t pt-1 text-xs">
                      Total: <b>{money(r.dailyFare)}</b>
                    </div>
                  </td>
                  <td className="px-5 py-4">
                    <div>
                      {r.cardType}: <b>{money(r.cardAmount ?? r.amount)}</b>
                    </div>
                    {r.secondCardType && (
                      <div>
                        {r.secondCardType}: <b>{money(r.secondCardAmount)}</b>
                      </div>
                    )}
                    <div className="mt-1 border-t pt-1 text-forest-700">
                      Geral: <b>{money(r.amount)}</b>
                    </div>
                  </td>
                  <td className="px-5 py-4">
                    <span
                      className={`rounded-full px-2 py-1 text-xs font-semibold ${r.active === false ? "bg-slate-100 text-slate-600" : "bg-emerald-50 text-emerald-700"}`}
                    >
                      {r.active === false ? "Inativo" : "Ativo"}
                    </span>
                  </td>
                  <td className="px-5 py-4">
                    <div className="flex gap-3">
                      <button
                        onClick={() => edit(r)}
                        className="text-xs font-semibold text-forest-700 hover:underline"
                      >
                        Editar
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {list.length === 0 && (
            <div className="py-12 text-center text-sm text-slate-400">
              Nenhum funcionário cadastrado.
            </div>
          )}
        </div>
      </div>
    </main>
  );
}
function RechargesPage({
  rows,
  onMark,
  onUndo,
}: {
  rows: Recharge[];
  onMark: (r: Recharge) => void;
  onUndo: (r: Recharge) => void;
}) {
  const [q, setQ] = useState("");
  const [st, setSt] = useState("Todos");
  const currentRows = currentEmployeeRows(rows);
  const list = currentRows.filter(
    (r) =>
      (st === "Todos" || r.status === st) &&
      `${r.employee} ${r.store} ${r.cardType}`
        .toLowerCase()
        .includes(q.toLowerCase()),
  );
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Recargas"
        sub="Acompanhe e conclua todas as solicitações"
      />
      <div className="rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="flex flex-col gap-3 p-5 sm:flex-row">
          <div className="relative flex-1">
            <Search
              className="absolute left-3 top-2.5 text-slate-400"
              size={17}
            />
            <input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder="Nome, loja ou cartão..."
              className="w-full rounded-xl border border-slate-200 py-2 pl-9 pr-3 text-sm"
            />
          </div>
          <select
            value={st}
            onChange={(e) => setSt(e.target.value)}
            className="rounded-xl border border-slate-200 px-3 py-2 text-sm"
          >
            {[
              "Todos",
              "Pendente",
              "Atrasado",
              "Próximo",
              "Recarregado",
              "Recarregado atrasado",
            ].map((x) => (
              <option key={x}>{x}</option>
            ))}
          </select>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[900px] text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase text-slate-400">
              <tr>
                {[
                  "Funcionário",
                  "Loja / cartão",
                  "Data do crédito",
                  "Antecedência",
                  "Data da recarga",
                  "Status",
                  "Ação",
                ].map((h) => (
                  <th className="px-5 py-3" key={h}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {list.map((r) => {
                const previous = r.periodCompletionId
                  ? rows.find((x) => x.id === r.periodCompletionId)
                  : r.sourceRechargeId
                    ? rows.find((x) => x.id === r.sourceRechargeId)
                    : undefined;
                const done =
                  r.status === "Recarregado" ||
                  r.status === "Recarregado atrasado";
                return (
                  <tr key={r.id}>
                    <td className="px-5 py-4 font-semibold">{r.employee}</td>
                    <td className="px-5 py-4">
                      {r.store} · {r.cardType}
                      {r.secondCardType ? ` + ${r.secondCardType}` : ""}
                    </td>
                    <td className="px-5 py-4">{formatDate(r.creditDate)}</td>
                    <td className="px-5 py-4">{r.advance} dias</td>
                    <td className="px-5 py-4 font-bold text-forest-700">
                      {formatDate(r.rechargeDate)}
                    </td>
                    <td className="px-5 py-4">
                      <StatusBadge status={r.status} />
                      {r.completedDate && (
                        <div className="mt-1 text-[11px] text-slate-400">
                          Feita em {formatDate(r.completedDate)}
                        </div>
                      )}
                    </td>
                    <td className="px-5 py-4">
                      {done ? (
                        <div className="flex gap-2">
                          <button
                            onClick={() => onMark(previous || r)}
                            className="rounded-lg border border-forest-200 px-3 py-2 text-xs font-semibold text-forest-700"
                          >
                            Editar data
                          </button>
                          <button
                            onClick={() => onUndo(previous || r)}
                            className="rounded-lg border border-red-200 px-3 py-2 text-xs font-semibold text-red-600"
                          >
                            Desfazer
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => onMark(r)}
                          className="rounded-lg border border-forest-200 px-3 py-2 text-xs font-semibold text-forest-700"
                        >
                          Marcar como recarregado
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </main>
  );
}
function CalendarPage({
  rows,
  period,
  setPeriod,
}: {
  rows: Recharge[];
  period: string;
  setPeriod: (v: string) => void;
}) {
  const calendarRows = currentEmployeeRows(rows);
  const [yearText, monthText] = period.split("-");
  const month = new Date(Number(yearText), Number(monthText) - 1, 1);
  const y = month.getFullYear(),
    m = month.getMonth(),
    first = new Date(y, m, 1).getDay(),
    count = new Date(y, m + 1, 0).getDate();
  const cells = Array(first)
    .fill(null)
    .concat(Array.from({ length: count }, (_, i) => i + 1));
  const move = (n: number) => {
    const d = new Date(y, m + n, 1);
    setPeriod(
      `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`,
    );
  };
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Calendário de recargas"
        sub="Planejamento mensal por data de execução"
        action={
          <div className="flex items-center gap-2">
            <button
              onClick={() => move(-1)}
              className="rounded-lg border bg-white p-2"
            >
              <ChevronLeft />
            </button>
            <b className="min-w-36 text-center capitalize">
              {month.toLocaleDateString("pt-BR", {
                month: "long",
                year: "numeric",
              })}
            </b>
            <button
              onClick={() => move(1)}
              className="rounded-lg border bg-white p-2"
            >
              <ChevronRight />
            </button>
          </div>
        }
      />
      <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="grid grid-cols-7 bg-slate-50 text-center text-xs font-bold uppercase text-slate-400">
          {["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"].map((d) => (
            <div className="p-3" key={d}>
              {d}
            </div>
          ))}
        </div>
        <div className="grid grid-cols-7">
          {cells.map((d, i) => {
            const matches = d
              ? calendarRows.filter((r) => {
                  const x = new Date(r.rechargeDate + "T12:00:00");
                  return (
                    x.getFullYear() === y &&
                    x.getMonth() === m &&
                    x.getDate() === d
                  );
                })
              : [];
            return (
              <div
                key={i}
                className="min-h-24 border-r border-t border-slate-100 p-2 sm:min-h-32"
              >
                <span className="text-xs font-semibold text-slate-500">
                  {d}
                </span>
                {matches.slice(0, 3).map((r) => (
                  <div
                    key={r.id}
                    title={r.employee}
                    className={`mt-1 truncate rounded px-1.5 py-1 text-[10px] font-semibold ${statusStyle[r.status]}`}
                  >
                    {r.employee}
                  </div>
                ))}
                {matches.length > 3 && (
                  <div className="mt-1 text-[10px] text-slate-400">
                    +{matches.length - 3} recargas
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </main>
  );
}
function ReportsPage({ rows }: { rows: Recharge[] }) {
  const exportXlsx = async () => {
    const wb = new ExcelJS.Workbook(),
      ws = wb.addWorksheet("Recargas");
    ws.columns = [
      { header: "Funcionário", key: "employee", width: 32 },
      { header: "Loja", key: "store", width: 24 },
      { header: "Cartão", key: "card", width: 16 },
      { header: "Data crédito", key: "credit", width: 16 },
      { header: "Data recarga", key: "recharge", width: 16 },
      { header: "Antecedência", key: "advance", width: 14 },
      { header: "Status", key: "status", width: 16 },
    ];
    rows.forEach((r) => ws.addRow({
      employee: r.employee,
      store: r.store,
      card: r.cardType,
      credit: r.creditDate,
      recharge: r.rechargeDate,
      advance: r.advance,
      status: r.status,
    }));
    ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
    const bytes = await wb.xlsx.writeBuffer(),
      link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    link.download = "relatorio-recargas.xlsx";
    link.click();
    URL.revokeObjectURL(link.href);
  };
  const exportPdf = () => {
    const doc = new jsPDF();
    doc.setFontSize(18);
    doc.text("Relatório de recargas", 14, 18);
    doc.setFontSize(9);
    doc.text(`Gerado em ${new Date().toLocaleString("pt-BR")}`, 14, 25);
    autoTable(doc, {
      startY: 31,
      head: [["Funcionário", "Loja", "Crédito", "Recarga", "Status"]],
      body: rows.map((r) => [
        r.employee,
        r.store,
        r.creditDate,
        r.rechargeDate,
        r.status,
      ]),
      styles: { fontSize: 8 },
      headStyles: { fillColor: [18, 90, 72] },
    });
    doc.save("relatorio-recargas.pdf");
  };
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Relatórios"
        sub="Exporte os dados de recarga por período"
      />
      <div className="grid gap-4 sm:grid-cols-2">
        <button
          onClick={exportPdf}
          className="rounded-2xl border border-slate-200 bg-white p-7 text-left shadow-soft hover:border-forest-300"
        >
          <Download className="text-red-500" />
          <h3 className="mt-5 font-bold">Relatório em PDF</h3>
          <p className="mt-1 text-sm text-slate-500">
            Documento pronto para impressão com todas as recargas.
          </p>
        </button>
        <button
          onClick={exportXlsx}
          className="rounded-2xl border border-slate-200 bg-white p-7 text-left shadow-soft hover:border-forest-300"
        >
          <FileSpreadsheet className="text-emerald-600" />
          <h3 className="mt-5 font-bold">Planilha Excel</h3>
          <p className="mt-1 text-sm text-slate-500">
            Dados completos em formato XLSX para análise.
          </p>
        </button>
      </div>
    </main>
  );
}
function StoresPage({
  stores,
  setStores,
  companyCnpjs,
  setCompanyCnpjs,
}: {
  stores: string[];
  setStores: (s: string[]) => void;
  companyCnpjs: Record<string, string>;
  setCompanyCnpjs: (values: Record<string, string>) => void;
}) {
  const [name, setName] = useState("");
  const [cnpj, setCnpj] = useState("");
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead title="Lojas e empresas" sub="Gerencie cada unidade e seu CNPJ em um único cadastro" />
      <form
        onSubmit={(e) => {
          e.preventDefault();
          const value = name.trim();
          if (value && !stores.includes(value)) setStores([...stores, value]);
          if (value) setCompanyCnpjs({ ...companyCnpjs, [value]: cnpj });
          setName("");
          setCnpj("");
        }}
        className="mb-5 grid max-w-3xl gap-2 sm:grid-cols-[1fr_230px_auto]"
      >
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Nome da nova loja"
          className="flex-1 rounded-xl border border-slate-200 px-4 py-2.5 text-sm"
        />
        <input
          required
          value={cnpj}
          onChange={(e) => setCnpj(formatCnpj(e.target.value))}
          placeholder="00.000.000/0000-00"
          className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm"
        />
        <button className="rounded-xl bg-forest-700 px-4 text-sm font-semibold text-white">
          Adicionar
        </button>
      </form>
      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
        {stores.map((s) => (
          <div
            key={s}
            className="flex items-center rounded-2xl border border-slate-200 bg-white p-5 shadow-soft"
          >
            <div className="grid h-10 w-10 place-items-center rounded-xl bg-forest-50 text-forest-700">
              <Store size={19} />
            </div>
            <div className="ml-3 min-w-0"><b className="block truncate">{s}</b><span className="text-xs text-slate-500">CNPJ: {companyCnpjs[s] || "não informado"}</span></div>
            <button type="button" onClick={() => { const next = prompt(`CNPJ de ${s}`, companyCnpjs[s] || ""); if (next !== null) setCompanyCnpjs({ ...companyCnpjs, [s]: formatCnpj(next) }); }} className="ml-auto mr-3 text-xs font-semibold text-slate-600">Editar</button>
            <button
              onClick={() => { setStores(stores.filter((x) => x !== s)); const next = { ...companyCnpjs }; delete next[s]; setCompanyCnpjs(next); }}
              className="text-xs font-semibold text-red-500"
            >
              Excluir
            </button>
          </div>
        ))}
      </div>
    </main>
  );
}
function SettingsPage({
  positions,
  setPositions,
  unregisteredReasons,
  setUnregisteredReasons,
}: {
  positions: string[];
  setPositions: (p: string[]) => void;
  unregisteredReasons: string[];
  setUnregisteredReasons: (p: string[]) => void;
}) {
  const [advance, setAdvance] = useState(
    () => localStorage.getItem("valefluxo_advance") || "3",
  );
  const [position, setPosition] = useState("");
  const [reason, setReason] = useState("");
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Configurações"
        sub="Defina o comportamento padrão e os cargos disponíveis"
      />
      <div className="grid gap-5 lg:grid-cols-2">
        <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-soft">
            <h3 className="font-bold">Antecedência padrão</h3>
            <p className="mt-1 text-sm text-slate-500">
              Quantidade de dias antes do crédito em que a recarga deve ser
              feita.
            </p>
            <select
              value={advance}
              onChange={(e) => setAdvance(e.target.value)}
              className="mt-5 w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm"
            >
              {[2, 3, 4, 5].map((n) => (
                <option value={n} key={n}>
                  {n} dias
                </option>
              ))}
            </select>
            <button
              onClick={() => {
                localStorage.setItem("valefluxo_advance", advance);
                alert("Configurações salvas com sucesso.");
              }}
              className="mt-5 rounded-xl bg-forest-700 px-4 py-2.5 text-sm font-semibold text-white"
            >
              Salvar configurações
            </button>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-soft">
            <h3 className="font-bold">Funções e cargos</h3>
            <p className="mt-1 text-sm text-slate-500">
              Os cargos aparecem no cadastro dos funcionários.
            </p>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                const v = position.trim();
                if (v && !positions.includes(v))
                  setPositions([...positions, v]);
                setPosition("");
              }}
              className="mt-5 flex gap-2"
            >
              <input
                required
                value={position}
                onChange={(e) => setPosition(e.target.value)}
                placeholder="Novo cargo"
                className="min-w-0 flex-1 rounded-xl border border-slate-200 px-3 py-2.5 text-sm"
              />
              <button className="rounded-xl bg-forest-700 px-4 text-sm font-semibold text-white">
                Adicionar
              </button>
            </form>
            <div className="mt-4 space-y-2">
              {positions.map((p) => (
                <div
                  key={p}
                  className="flex items-center rounded-xl bg-slate-50 px-4 py-3 text-sm font-semibold text-slate-700"
                >
                  <UserRound className="mr-2 text-forest-700" size={17} />
                  {p}
                  <button
                    onClick={() =>
                      setPositions(positions.filter((x) => x !== p))
                    }
                    className="ml-auto text-xs text-red-500"
                  >
                    Excluir
                  </button>
                </div>
              ))}
              {positions.length === 0 && (
                <p className="py-4 text-center text-sm text-slate-400">
                  Nenhum cargo cadastrado.
                </p>
              )}
            </div>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-soft">
            <h3 className="font-bold">Motivos para funcionário sem carteira</h3>
            <p className="mt-1 text-sm text-slate-500">
              Cadastre os motivos disponíveis no cadastro do funcionário.
            </p>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                const value = reason.trim();
                if (value && !unregisteredReasons.includes(value))
                  setUnregisteredReasons([...unregisteredReasons, value]);
                setReason("");
              }}
              className="mt-5 flex gap-2"
            >
              <input
                required
                value={reason}
                onChange={(e) => setReason(e.target.value)}
                placeholder="Novo motivo"
                className="min-w-0 flex-1 rounded-xl border border-slate-200 px-3 py-2.5 text-sm"
              />
              <button className="rounded-xl bg-forest-700 px-4 text-sm font-semibold text-white">
                Adicionar
              </button>
            </form>
            <div className="mt-4 space-y-2">
              {unregisteredReasons.map((item) => (
                <div
                  key={item}
                  className="flex items-center rounded-xl bg-slate-50 px-4 py-3 text-sm font-semibold text-slate-700"
                >
                  <TriangleAlert className="mr-2 text-slate-700" size={17} />
                  {item}
                  <button
                    type="button"
                    onClick={() =>
                      setUnregisteredReasons(
                        unregisteredReasons.filter((x) => x !== item),
                      )
                    }
                    className="ml-auto text-xs text-red-500"
                  >
                    Excluir
                  </button>
                </div>
              ))}
              {unregisteredReasons.length === 0 && (
                <p className="py-4 text-center text-sm text-slate-400">
                  Nenhum motivo cadastrado.
                </p>
              )}
            </div>
        </div>
      </div>
    </main>
  );
}

function ConfigurationsPage({
  positions,
  setPositions,
  unregisteredReasons,
  setUnregisteredReasons,
  stores,
  setStores,
  rows,
  setRows,
  financialEntries,
  setFinancialEntries,
  financialPeriod,
  setFinancialPeriod,
  openEmployeeForm,
  editEmployee,
  companyCnpjs,
  setCompanyCnpjs,
}: {
  positions: string[];
  setPositions: (p: string[]) => void;
  unregisteredReasons: string[];
  setUnregisteredReasons: (p: string[]) => void;
  stores: string[];
  setStores: (s: string[]) => void;
  rows: Recharge[];
  setRows: (rows: Recharge[]) => void;
  financialEntries: FinancialEntry[];
  setFinancialEntries: (entries: FinancialEntry[]) => void;
  financialPeriod: string;
  setFinancialPeriod: (period: string) => void;
  openEmployeeForm: () => void;
  editEmployee: (employee: Recharge) => void;
  companyCnpjs: Record<string, string>;
  setCompanyCnpjs: (values: Record<string, string>) => void;
}) {
  const [tab, setTab] = useState<"geral" | "funcionarios" | "lojas" | "desligados" | "financeiro" | "perfis">(() => {
    const saved = localStorage.getItem("abc_settings_tab");
    return saved === "funcionarios" || saved === "lojas" || saved === "desligados" || saved === "financeiro" || saved === "perfis" ? saved : "geral";
  });
  const button = (key: typeof tab, label: string, icon: ReactNode) => (
    <button
      onClick={() => { setTab(key); localStorage.setItem("abc_settings_tab", key); }}
      className={`rounded-lg px-4 py-2 text-sm font-semibold ${tab === key ? "bg-forest-700 text-white" : "text-slate-500 hover:bg-slate-50"}`}
    >
      {icon}
      {label}
    </button>
  );
  return (
    <>
      <div className="px-4 pt-5 sm:px-7">
        <div className="inline-flex flex-wrap rounded-xl border border-slate-200 bg-white p-1 shadow-sm">
          {button(
            "geral",
            "Geral e funções",
            <Settings className="mr-2 inline" size={16} />,
          )}
          {button(
            "funcionarios",
            "Funcionários",
            <Users className="mr-2 inline" size={16} />,
          )}
          {button(
            "lojas",
            "Lojas e empresas",
            <Store className="mr-2 inline" size={16} />,
          )}
          {button(
              "desligados",
              "Desligados",
              <UserRound className="mr-2 inline" size={16} />,
            )}
          {button(
              "financeiro",
              "Ocultados do financeiro",
              <DollarSign className="mr-2 inline" size={16} />,
            )}
          {button(
            "perfis",
            "Perfis e usuários",
            <Users className="mr-2 inline" size={16} />,
          )}
        </div>
      </div>
      {tab === "geral" ? (
        <SettingsPage
          positions={positions}
          setPositions={setPositions}
          unregisteredReasons={unregisteredReasons}
          setUnregisteredReasons={setUnregisteredReasons}
        />
      ) : tab === "funcionarios" ? (
        <HREmployeesPage
          rows={rows}
          openForm={openEmployeeForm}
          edit={editEmployee}
          toggleCritical={(record) => setRows(rows.map((item) => item.id === record.id ? { ...item, experienceCritical: !item.experienceCritical } : item))}
          dismiss={(record, date, notice, start, end) => setRows(rows.map((item) => item.id === record.id ? {
            ...item,
            active: new Date(date + "T12:00:00") > new Date(),
            employmentStatus: new Date(date + "T12:00:00") > new Date() ? "Ativo" : "Desligado",
            terminationDate: date,
            noticeStart: notice ? start : "",
            noticeEnd: notice ? end : "",
            experienceCritical: false,
          } : item))}
        />
      ) : tab === "lojas" ? (
        <StoresPage stores={stores} setStores={setStores} companyCnpjs={companyCnpjs} setCompanyCnpjs={setCompanyCnpjs} />
      ) : tab === "perfis" ? (
        <UserProfiles stores={stores} />
      ) : tab === "financeiro" ? (
        <FinancialRegistrations
          employees={rows.filter((employee) =>
            financialEntries.some(
              (entry) => entry.employeeId === employee.id && !!entry.noPaymentsFrom,
            ),
          )}
          entries={financialEntries}
          setEntries={setFinancialEntries}
          period={financialPeriod}
          setPeriod={setFinancialPeriod}
          showHidden
        />
      ) : (
        <DismissedEmployeesPage
          rows={rows}
          restore={(record) =>
            setRows(
              rows.map((item) =>
                item.id === record.id
                  ? {
                      ...item,
                      active: true,
                      employmentStatus: "Ativo",
                      terminationDate: "",
                      noticeStart: "",
                      noticeEnd: "",
                    }
                  : item,
              ),
            )
          }
        />
      )}
    </>
  );
}

function SmartRechargeModal({
  record,
  close,
  confirm,
}: {
  record: Recharge;
  close: () => void;
  confirm: (date: string, days: number, total: number) => void;
}) {
  const today = new Date().toISOString().slice(0, 10),
    [date, setDate] = useState(record.completedDate || today),
    [mode, setMode] = useState<"days" | "amount" | "until">("days");
  const actual = new Date(date + "T12:00:00"),
    available = new Date(actual);
  available.setDate(available.getDate() + record.advance);
  const nextRecharge = new Date(record.rechargeDate + "T12:00:00");
  nextRecharge.setMonth(nextRecharge.getMonth() + 1);
  const nextCredit = new Date(nextRecharge);
  nextCredit.setDate(nextCredit.getDate() + record.advance);
  const workDays = record.workDays || [1, 2, 3, 4, 5];
  let suggestedDays = 0;
  for (
    const d = new Date(available);
    d < nextCredit;
    d.setDate(d.getDate() + 1)
  )
    if (isWorkDay(record, d)) suggestedDays++;
  const [days, setDays] = useState(record.chargedDays ?? suggestedDays),
    [manualAmount, setManualAmount] = useState(
      record.amount
        ? formatMoneyInput(String(Math.round(record.amount * 100)))
        : "",
    ),
    [until, setUntil] = useState(() => {
      const d = new Date(record.rechargeDate + "T12:00:00");
      d.setMonth(d.getMonth() + 1);
      return d.toISOString().slice(0, 10);
    });
  const fare1 = record.cardDailyFare ?? record.dailyFare ?? 0,
    fare2 = record.secondCardDailyFare ?? 0,
    dailyTotal = fare1 + fare2,
    enteredAmount = parseMoney(manualAmount);
  let untilDays = 0;
  if (until) {
    const limit = new Date(until + "T12:00:00");
    limit.setDate(limit.getDate() + record.advance);
    for (const d = new Date(available); d < limit; d.setDate(d.getDate() + 1))
      if (isWorkDay(record, d)) untilDays++;
  }
  const effectiveDays =
      mode === "amount" && dailyTotal > 0
        ? Math.floor(enteredAmount / dailyTotal)
        : mode === "until"
          ? untilDays
          : days,
    total = mode === "amount" ? enteredAmount : effectiveDays * dailyTotal;
  const amount1 = dailyTotal ? total * (fare1 / dailyTotal) : 0,
    amount2 = dailyTotal ? total * (fare2 / dailyTotal) : 0,
    lastsUntil = new Date(available);
  let counted = 0;
  while (counted < effectiveDays) {
    if (isWorkDay(record, lastsUntil)) counted++;
    if (counted < effectiveDays) lastsUntil.setDate(lastsUntil.getDate() + 1);
  }
  const money = (v: number) =>
      v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }),
    tab = (value: "days" | "amount" | "until", label: string) => (
      <button
        type="button"
        onClick={() => setMode(value)}
        className={`rounded-lg px-2 py-2 text-xs font-bold ${mode === value ? "bg-white text-forest-700 shadow-sm" : "text-slate-500"}`}
      >
        {label}
      </button>
    );
  return (
    <div className="fixed inset-0 z-[60] grid place-items-center bg-slate-950/45 p-4 backdrop-blur-sm">
      <div className="fade-in max-h-[92vh] w-full max-w-md overflow-y-auto rounded-2xl bg-white p-6 shadow-2xl">
        <div className="flex items-start">
          <div>
            <h3 className="text-lg font-bold">
              {record.completedDate ? "Editar recarga" : "Confirmar recarga"}
            </h3>
            <p className="mt-1 text-sm text-slate-500">{record.employee}</p>
          </div>
          <button
            type="button"
            onClick={close}
            className="ml-auto rounded-lg p-2 text-slate-400 hover:bg-slate-100"
          >
            <X size={19} />
          </button>
        </div>
        <label className="mt-6 block text-sm font-semibold text-slate-700">
          Em qual dia a recarga foi realizada?
          <input
            autoFocus
            required
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 outline-none focus:border-forest-500"
          />
        </label>
        <div className="mt-4 grid grid-cols-3 rounded-xl bg-slate-100 p-1">
          {tab("days", "Informar dias")}
          {tab("amount", "Informar valor")}
          {tab("until", "Até a data")}
        </div>
        {mode === "days" ? (
          <label className="mt-4 block text-sm font-semibold text-slate-700">
            Número de dias que serão recarregados
            <input
              min="0"
              max="366"
              type="number"
              value={days}
              onChange={(e) => setDays(Math.max(0, Number(e.target.value)))}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-lg font-bold outline-none focus:border-forest-500"
            />
          </label>
        ) : mode === "amount" ? (
          <label className="mt-4 block text-sm font-semibold text-slate-700">
            Valor total da recarga (R$)
            <input
              inputMode="decimal"
              value={manualAmount}
              onChange={(e) =>
                setManualAmount(formatMoneyInput(e.target.value))
              }
              placeholder="Ex.: 500,00"
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-lg font-bold outline-none focus:border-forest-500"
            />
          </label>
        ) : (
          <label className="mt-4 block text-sm font-semibold text-slate-700">
            Até qual dia o cartão deve ter passagem?
            <input
              min={date}
              type="date"
              value={until}
              onChange={(e) => setUntil(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-lg font-bold outline-none focus:border-forest-500"
            />
            <span className="mt-2 block text-xs font-normal text-slate-400">
              Serão contados somente os dias da escala de trabalho.
            </span>
          </label>
        )}
        <div className="mt-5 rounded-xl bg-forest-50 p-4 text-sm">
          <div className="flex justify-between">
            <span className="text-slate-500">Crédito disponível</span>
            <b className="text-forest-800">
              {available.toLocaleDateString("pt-BR")}
            </b>
          </div>
          <div className="mt-2 flex justify-between">
            <span className="text-slate-500">Duração estimada</span>
            <b className="text-right text-forest-800">
              {effectiveDays} dias · até{" "}
              {effectiveDays ? lastsUntil.toLocaleDateString("pt-BR") : "—"}
            </b>
          </div>
          <div className="my-3 border-t border-forest-100" />
          <div className="flex justify-between">
            <span className="text-slate-500">
              {record.cardType}: {money(fare1)}/dia
            </span>
            <b>{money(amount1)}</b>
          </div>
          {record.secondCardType && (
            <div className="mt-2 flex justify-between">
              <span className="text-slate-500">
                {record.secondCardType}: {money(fare2)}/dia
              </span>
              <b>{money(amount2)}</b>
            </div>
          )}
          <div className="mt-3 flex justify-between rounded-lg bg-forest-800 px-3 py-2.5 text-white">
            <span className="font-semibold">Total da recarga</span>
            <b className="text-lg">{money(total)}</b>
          </div>
          <p className="mt-3 text-xs text-slate-500">
            Cálculo baseado na escala cadastrada e no valor diário de cada
            cartão.
          </p>
        </div>
        <div className="mt-6 flex justify-end gap-2">
          <button
            type="button"
            onClick={close}
            className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600"
          >
            Cancelar
          </button>
          <button
            type="button"
            disabled={total <= 0}
            onClick={() => confirm(date, effectiveDays, total)}
            className="rounded-xl bg-forest-700 px-4 py-2.5 text-sm font-semibold text-white disabled:bg-slate-300"
          >
            {record.completedDate ? "Salvar alterações" : "Confirmar recarga"}
          </button>
        </div>
      </div>
    </div>
  );
}

function RechargeHistoryPage({
  employees,
  events,
  planned,
  onMark,
  onManual,
  onEdit,
  onDelete,
}: {
  employees: Recharge[];
  events: RechargeEvent[];
  planned: Recharge[];
  onMark: (r: Recharge) => void;
  onManual: (r: Recharge) => void;
  onEdit: (r: Recharge, e: RechargeEvent) => void;
  onDelete: (e: RechargeEvent) => void;
}) {
  const [q, setQ] = useState(""),
    [employeeId, setEmployeeId] = useState("");
  const byId = new Map(employees.map((e) => [e.id, e])),
    money = (v: number | undefined) =>
      (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  const history = events
    .map((event) => ({ event, employee: byId.get(event.employeeId) }))
    .filter(
      (x) =>
        x.employee &&
        `${x.employee.employee} ${x.employee.store} ${x.employee.cardType}`
          .toLowerCase()
          .includes(q.toLowerCase()),
    )
    .sort(
      (a, b) =>
        b.event.completedDate.localeCompare(a.event.completedDate) ||
        b.event.id - a.event.id,
    );
  const pending = planned.filter(
    (r) =>
      r.status !== "Recarregado" &&
      r.status !== "Recarregado atrasado" &&
      `${r.employee} ${r.store} ${r.cardType}`
        .toLowerCase()
        .includes(q.toLowerCase()),
  );
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Histórico de recargas"
        sub="Todas as recargas realizadas e os próximos pagamentos"
      />
      <div className="mb-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <div className="flex flex-col gap-3 md:flex-row">
          <SearchableEmployeeSelect
            employees={employees.filter((employee) => employee.active !== false)}
            value={employeeId}
            onChange={setEmployeeId}
            placeholder="Pesquisar funcionário para recarga manual..."
            className="min-w-0 flex-1"
          />
          <button
            disabled={!employeeId}
            onClick={() => {
              const employee = employees.find(
                (e) => e.id === Number(employeeId),
              );
              if (employee) onManual(employee);
            }}
            className="rounded-xl bg-forest-700 px-5 py-3 text-sm font-semibold text-white disabled:bg-slate-300"
          >
            <Plus className="mr-2 inline" size={17} />
            Cadastrar recarga manual
          </button>
        </div>
        <div className="relative mt-3">
          <Search className="absolute left-3 top-3 text-slate-400" size={17} />
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Buscar por funcionário, loja ou cartão..."
            className="w-full rounded-xl border border-slate-200 py-3 pl-10 pr-3 text-sm"
          />
        </div>
      </div>
      <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="border-b border-slate-100 px-5 py-4">
          <h3 className="font-bold text-slate-800">Recargas realizadas</h3>
          <p className="text-xs text-slate-400">
            Cada baixa aparece individualmente neste histórico.
          </p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[1050px] text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase text-slate-400">
              <tr>
                {[
                  "Data",
                  "Funcionário",
                  "Loja / cartão",
                  "Dias",
                  "BHBus / 1º cartão",
                  "Ótimo / 2º cartão",
                  "Total",
                  "Ações",
                ].map((h) => (
                  <th key={h} className="px-5 py-3">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {history.map(({ event, employee }) => (
                <tr key={event.id}>
                  <td className="px-5 py-4 font-bold text-forest-700">
                    {formatDate(event.completedDate)}
                  </td>
                  <td className="px-5 py-4 font-semibold">
                    {employee!.employee}
                  </td>
                  <td className="px-5 py-4">
                    {employee!.store}
                    <div className="text-xs text-slate-400">
                      {employee!.cardType}
                      {employee!.secondCardType
                        ? ` + ${employee!.secondCardType}`
                        : ""}
                    </div>
                  </td>
                  <td className="px-5 py-4">{event.chargedDays ?? "—"}</td>
                  <td className="px-5 py-4">{money(event.cardAmount)}</td>
                  <td className="px-5 py-4">
                    {employee!.secondCardType
                      ? money(event.secondCardAmount)
                      : "—"}
                  </td>
                  <td className="px-5 py-4 font-bold">
                    {money(event.totalAmount)}
                  </td>
                  <td className="px-5 py-4">
                    <div className="flex gap-2">
                      <button
                        onClick={() => onEdit(employee!, event)}
                        className="rounded-lg border border-forest-200 px-3 py-2 text-xs font-semibold text-forest-700"
                      >
                        Editar
                      </button>
                      <button
                        onClick={() => onDelete(event)}
                        className="rounded-lg border border-red-200 px-3 py-2 text-xs font-semibold text-red-600"
                      >
                        Cancelar
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {history.length === 0 && (
                <tr>
                  <td
                    colSpan={8}
                    className="px-5 py-12 text-center text-slate-400"
                  >
                    Nenhuma recarga realizada.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
      {pending.length > 0 && (
        <div className="mt-5 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
          <div className="border-b border-slate-100 px-5 py-4">
            <h3 className="font-bold text-slate-800">Aguardando baixa</h3>
          </div>
          {pending.map((r) => (
            <div
              key={r.id}
              className="flex flex-col gap-3 border-b border-slate-100 px-5 py-4 sm:flex-row sm:items-center"
            >
              <div className="flex-1">
                <b>{r.employee}</b>
                <div className="text-xs text-slate-400">
                  {r.store} · prevista para {formatDate(r.rechargeDate)}
                </div>
              </div>
              <StatusBadge status={r.status} />
              <button
                onClick={() => onMark(r)}
                className="rounded-lg border border-forest-200 px-4 py-2 text-xs font-semibold text-forest-700"
              >
                Marcar como recarregado
              </button>
            </div>
          ))}
        </div>
      )}
    </main>
  );
}

function AutoRechargeModal({
  record,
  close,
  confirm,
}: {
  record: Recharge;
  close: () => void;
  confirm: (date: string, days: number, total: number) => void;
}) {
  const today = new Date().toISOString().slice(0, 10),
    [date, setDate] = useState(record.completedDate || today),
    workDays = record.workDays || [1, 2, 3, 4, 5];
  const paidAt = new Date(date + "T12:00:00"),
    available = new Date(paidAt);
  available.setDate(available.getDate() + record.advance);
  const nextRecharge = new Date(record.rechargeDate + "T12:00:00");
  nextRecharge.setMonth(nextRecharge.getMonth() + 1);
  const nextCredit = new Date(nextRecharge);
  nextCredit.setDate(nextCredit.getDate() + record.advance);
  let days = 0;
  for (
    const d = new Date(available);
    d < nextCredit;
    d.setDate(d.getDate() + 1)
  )
    if (isWorkDay(record, d)) days++;
  const fare1 = record.cardDailyFare ?? record.dailyFare ?? 0,
    fare2 = record.secondCardDailyFare ?? 0,
    amount1 = days * fare1,
    amount2 = days * fare2,
    total = amount1 + amount2,
    money = (v: number) =>
      v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }),
    invalid = available >= nextCredit;
  return (
    <div className="fixed inset-0 z-[60] grid place-items-center bg-slate-950/45 p-4 backdrop-blur-sm">
      <div className="fade-in w-full max-w-md rounded-2xl bg-white p-6 shadow-2xl">
        <div className="flex items-start">
          <div>
            <h3 className="text-lg font-bold">
              {record.completedDate
                ? "Editar data da recarga"
                : "Confirmar recarga"}
            </h3>
            <p className="mt-1 text-sm text-slate-500">{record.employee}</p>
          </div>
          <button
            type="button"
            onClick={close}
            className="ml-auto rounded-lg p-2 text-slate-400 hover:bg-slate-100"
          >
            <X size={19} />
          </button>
        </div>
        <label className="mt-6 block text-sm font-semibold text-slate-700">
          Em qual dia a recarga foi realizada?
          <input
            autoFocus
            required
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-lg font-bold outline-none focus:border-forest-500"
          />
        </label>
        <div className="mt-5 rounded-xl bg-forest-50 p-4 text-sm">
          <div className="flex justify-between">
            <span className="text-slate-500">Crédito desta recarga</span>
            <b className="text-forest-800">
              {available.toLocaleDateString("pt-BR")}
            </b>
          </div>
          <div className="mt-2 flex justify-between">
            <span className="text-slate-500">Próxima recarga fixa</span>
            <b className="text-forest-800">
              {nextRecharge.toLocaleDateString("pt-BR")}
            </b>
          </div>
          <div className="mt-2 flex justify-between">
            <span className="text-slate-500">Próximo crédito</span>
            <b className="text-forest-800">
              {nextCredit.toLocaleDateString("pt-BR")}
            </b>
          </div>
          <div className="my-3 border-t border-forest-100" />
          <div className="flex justify-between">
            <span className="text-slate-500">Dias trabalhados a cobrir</span>
            <b className="text-forest-800">{days} dias</b>
          </div>
          <div className="mt-3 flex justify-between">
            <span className="text-slate-500">
              {record.cardType}: {money(fare1)}/dia
            </span>
            <b>{money(amount1)}</b>
          </div>
          {record.secondCardType && (
            <div className="mt-2 flex justify-between">
              <span className="text-slate-500">
                {record.secondCardType}: {money(fare2)}/dia
              </span>
              <b>{money(amount2)}</b>
            </div>
          )}
          <div className="mt-3 flex justify-between rounded-lg bg-forest-800 px-3 py-2.5 text-white">
            <span className="font-semibold">Valor total a recarregar</span>
            <b className="text-lg">{money(total)}</b>
          </div>
          <p className="mt-3 text-xs leading-relaxed text-slate-500">
            Calculado automaticamente da liberação desta carga até a liberação
            da próxima, considerando somente a escala de trabalho.
          </p>
          {invalid && (
            <p className="mt-3 rounded-lg bg-red-50 p-3 text-xs font-semibold text-red-600">
              A data informada ultrapassa a liberação da próxima recarga.
            </p>
          )}
        </div>
        <div className="mt-6 flex justify-end gap-2">
          <button
            type="button"
            onClick={close}
            className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600"
          >
            Cancelar
          </button>
          <button
            type="button"
            disabled={invalid || total <= 0}
            onClick={() => confirm(date, days, total)}
            className="rounded-xl bg-forest-700 px-4 py-2.5 text-sm font-semibold text-white disabled:bg-slate-300"
          >
            {record.completedDate ? "Salvar alteração" : "Confirmar recarga"}
          </button>
        </div>
      </div>
    </div>
  );
}

function ReportsPageFiltered({ rows }: { rows: Recharge[] }) {
  const [store, setStore] = useState("Todas"),
    [employee, setEmployee] = useState("Todos"),
    [start, setStart] = useState(""),
    [end, setEnd] = useState("");
  const stores = [...new Set(rows.map((r) => r.store))].sort(),
    employees = [
      ...new Set(
        rows
          .filter((r) => store === "Todas" || r.store === store)
          .map((r) => r.employee),
      ),
    ].sort();
  const filtered = rows.filter(
    (r) =>
      (store === "Todas" || r.store === store) &&
      (employee === "Todos" || r.employee === employee) &&
      (!start || r.rechargeDate >= start) &&
      (!end || r.rechargeDate <= end),
  );
  const exportXlsx = async () => {
    const wb = new ExcelJS.Workbook();
    wb.creator = "Sacolão ABC";
    wb.created = new Date();
    const ws = wb.addWorksheet("Recargas", {
      views: [{ state: "frozen", ySplit: 7, showGridLines: false }],
    });
    const logoBytes = await fetch("/sacolao-abc-logo.png?v=4").then(
      (response) => response.arrayBuffer(),
    );
    const logoId = wb.addImage({
      buffer: logoBytes as never,
      extension: "png",
    });
    ws.addImage(logoId, {
      tl: { col: 0.2, row: 0.15 },
      ext: { width: 170, height: 76 },
    });
    ws.mergeCells("C1:H2");
    const title = ws.getCell("C1");
    title.value = "RELATÓRIO DE RECARGAS";
    title.font = {
      name: "Arial",
      size: 20,
      bold: true,
      color: { argb: "FF0E4E3E" },
    };
    title.alignment = { vertical: "middle" };
    ws.mergeCells("C3:H3");
    ws.getCell("C3").value =
      `Sacolão ABC - Gerado em ${new Date().toLocaleString("pt-BR")}`;
    ws.getCell("C3").font = {
      name: "Arial",
      size: 10,
      color: { argb: "FF718096" },
    };
    const total = filtered.reduce((sum, r) => sum + (r.amount || 0), 0);
    ws.getCell("A5").value = "Registros";
    ws.getCell("B5").value = filtered.length;
    ws.getCell("D5").value = "Valor total";
    ws.getCell("E5").value = total;
    ws.getCell("E5").numFmt = "R$ #,##0.00";
    ws.getCell("G5").value = "Atrasadas";
    ws.getCell("H5").value = filtered.filter(
      (r) => r.status === "Atrasado" || r.status === "Recarregado atrasado",
    ).length;
    ["A5", "D5", "G5"].forEach((ref) => {
      ws.getCell(ref).font = { bold: true, color: { argb: "FFFFFFFF" } };
      ws.getCell(ref).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0E4E3E" },
      };
    });
    ["B5", "E5", "H5"].forEach((ref) => {
      ws.getCell(ref).font = { bold: true, color: { argb: "FF0E4E3E" } };
      ws.getCell(ref).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFEAF4F0" },
      };
    });
    const header = [
      "Funcionário",
      "Loja",
      "Cartão",
      "Data do crédito",
      "Data da recarga",
      "Antecedência",
      "Status",
      "Valor",
    ];
    ws.addRow([]);
    ws.addRow(header);
    filtered.forEach((r) =>
      ws.addRow([
        r.employee,
        r.store,
        `${r.cardType}${r.secondCardType ? ` + ${r.secondCardType}` : ""}`,
        new Date(r.creditDate + "T12:00:00"),
        new Date(r.rechargeDate + "T12:00:00"),
        r.advance,
        r.status,
        r.amount || 0,
      ]),
    );
    const headerRow = ws.getRow(7);
    headerRow.height = 24;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0E4E3E" },
      };
      cell.alignment = { vertical: "middle" };
    });
    for (let row = 8; row <= ws.rowCount; row++) {
      ws.getRow(row).height = 21;
      if (row % 2 === 0)
        ws.getRow(row).eachCell(
          (cell) =>
            (cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFF4F8F6" },
            }),
        );
    }
    ws.getColumn(4).numFmt = "dd/mm/yyyy";
    ws.getColumn(5).numFmt = "dd/mm/yyyy";
    ws.getColumn(8).numFmt = "R$ #,##0.00";
    [28, 20, 22, 18, 18, 15, 23, 16].forEach(
      (width, index) => (ws.getColumn(index + 1).width = width),
    );
    ws.autoFilter = { from: "A7", to: `H${Math.max(7, ws.rowCount)}` };
    const bytes = await wb.xlsx.writeBuffer();
    const url = URL.createObjectURL(
      new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
    const link = document.createElement("a");
    link.href = url;
    link.download = "relatorio-recargas-filtrado.xlsx";
    link.click();
    URL.revokeObjectURL(url);
  };
  const exportPdf = async () => {
    const logoData = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = () => reject(reader.error);
      fetch("/sacolao-abc-logo.png?v=4")
        .then((response) => response.blob())
        .then((blob) => reader.readAsDataURL(blob))
        .catch(reject);
    });
    const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
      }),
      pageWidth = doc.internal.pageSize.getWidth(),
      total = filtered.reduce((sum, r) => sum + (r.amount || 0), 0),
      late = filtered.filter(
        (r) => r.status === "Atrasado" || r.status === "Recarregado atrasado",
      ).length,
      recharged = filtered.filter(
        (r) =>
          r.status === "Recarregado" || r.status === "Recarregado atrasado",
      ).length,
      moneyPdf = (v: number) =>
        v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }),
      filterText = [
        store !== "Todas" ? `Loja: ${store}` : "Todas as lojas",
        employee !== "Todos"
          ? `Funcionário: ${employee}`
          : "Todos os funcionários",
        start ? `De: ${formatDate(start)}` : "Sem data inicial",
        end ? `Até: ${formatDate(end)}` : "Sem data final",
      ].join("   |   ");
    doc.setFillColor(14, 78, 62);
    doc.rect(0, 0, pageWidth, 39, "F");
    doc.addImage(logoData, "PNG", 14, 7, 27, 25);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(22);
    doc.text("Relatório de recargas", 43, 18);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text("Sacolão ABC - Gestão de benefícios de transporte", 43, 25);
    doc.setTextColor(83, 102, 113);
    doc.setFontSize(8.5);
    doc.text(filterText, 14, 48);
    const cards = [
      ["REGISTROS", String(filtered.length)],
      ["VALOR TOTAL", moneyPdf(total)],
      ["RECARREGADAS", String(recharged)],
      ["COM ATRASO", String(late)],
    ];
    cards.forEach(([label, value], index) => {
      const x = 14 + index * 68;
      doc.setFillColor(
        index === 3 && late > 0 ? 255 : 242,
        index === 3 && late > 0 ? 244 : 248,
        index === 3 && late > 0 ? 244 : 246,
      );
      doc.setDrawColor(
        index === 3 && late > 0 ? 244 : 220,
        index === 3 && late > 0 ? 190 : 229,
        index === 3 && late > 0 ? 190 : 225,
      );
      doc.roundedRect(x, 54, 61, 24, 3, 3, "FD");
      doc.setTextColor(120, 139, 150);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(7);
      doc.text(label, x + 5, 62);
      doc.setTextColor(
        index === 3 && late > 0 ? 190 : 14,
        index === 3 && late > 0 ? 45 : 78,
        index === 3 && late > 0 ? 45 : 62,
      );
      doc.setFontSize(13);
      doc.text(value, x + 5, 72);
    });
    autoTable(doc, {
      startY: 85,
      margin: { left: 14, right: 14, bottom: 18 },
      head: [
        [
          "Funcionário",
          "Loja",
          "Cartão",
          "Crédito",
          "Recarga",
          "Dias",
          "Status",
          "Valor",
        ],
      ],
      body: filtered.map((r) => [
        r.employee,
        r.store,
        `${r.cardType}${r.secondCardType ? ` + ${r.secondCardType}` : ""}`,
        formatDate(r.creditDate),
        formatDate(r.rechargeDate),
        String(r.chargedDays ?? "-"),
        r.status,
        moneyPdf(r.amount || 0),
      ]),
      theme: "grid",
      styles: {
        fontSize: 8,
        cellPadding: 3.2,
        textColor: [45, 58, 64],
        lineColor: [226, 232, 229],
        lineWidth: 0.2,
        valign: "middle",
      },
      headStyles: {
        fillColor: [14, 78, 62],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        lineColor: [14, 78, 62],
      },
      alternateRowStyles: { fillColor: [246, 249, 248] },
      columnStyles: {
        0: { cellWidth: 48, fontStyle: "bold" },
        1: { cellWidth: 34 },
        2: { cellWidth: 38 },
        3: { cellWidth: 27 },
        4: { cellWidth: 27 },
        5: { cellWidth: 15, halign: "center" },
        6: { cellWidth: 34 },
        7: { cellWidth: 29, halign: "right", fontStyle: "bold" },
      },
    });
    // A API existe no jsPDF em tempo de execução, mas não consta na tipagem desta versão.
    // @ts-expect-error getNumberOfPages faz parte da API interna do jsPDF
    const pages = doc.internal.getNumberOfPages();
    for (let page = 1; page <= pages; page++) {
      doc.setPage(page);
      doc.setDrawColor(220, 229, 225);
      doc.line(14, 194, pageWidth - 14, 194);
      doc.setTextColor(130, 145, 152);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7.5);
      doc.text(`Gerado em ${new Date().toLocaleString("pt-BR")}`, 14, 200);
      doc.text(`Página ${page} de ${pages}`, pageWidth - 14, 200, {
        align: "right",
      });
    }
    doc.save("relatorio-recargas-filtrado.pdf");
  };
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Relatórios"
        sub="Filtre os dados antes de gerar os arquivos"
      />
      <div className="mb-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          <label className="text-xs font-semibold text-slate-600">
            Loja
            <select
              value={store}
              onChange={(e) => {
                setStore(e.target.value);
                setEmployee("Todos");
              }}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm"
            >
              <option>Todas</option>
              {stores.map((x) => (
                <option key={x}>{x}</option>
              ))}
            </select>
          </label>
          <label className="text-xs font-semibold text-slate-600">
            Funcionário
            <select
              value={employee}
              onChange={(e) => setEmployee(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm"
            >
              <option>Todos</option>
              {employees.map((x) => (
                <option key={x}>{x}</option>
              ))}
            </select>
          </label>
          <label className="text-xs font-semibold text-slate-600">
            Data inicial
            <input
              type="date"
              value={start}
              onChange={(e) => setStart(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm"
            />
          </label>
          <label className="text-xs font-semibold text-slate-600">
            Data final
            <input
              type="date"
              value={end}
              min={start}
              onChange={(e) => setEnd(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm"
            />
          </label>
        </div>
        <div className="mt-4 flex items-center justify-between border-t border-slate-100 pt-4">
          <span className="text-sm text-slate-500">
            <b className="text-forest-700">{filtered.length}</b> registro(s)
            encontrado(s)
          </span>
          <button
            onClick={() => {
              setStore("Todas");
              setEmployee("Todos");
              setStart("");
              setEnd("");
            }}
            className="text-xs font-semibold text-forest-700"
          >
            Limpar filtros
          </button>
        </div>
      </div>
      <div className="grid gap-4 sm:grid-cols-2">
        <button
          disabled={!filtered.length}
          onClick={exportPdf}
          className="rounded-2xl border border-slate-200 bg-white p-7 text-left shadow-soft hover:border-forest-300 disabled:opacity-50"
        >
          <Download className="text-red-500" />
          <h3 className="mt-5 font-bold">Relatório em PDF</h3>
          <p className="mt-1 text-sm text-slate-500">
            Exportar os {filtered.length} registros filtrados.
          </p>
        </button>
        <button
          disabled={!filtered.length}
          onClick={exportXlsx}
          className="rounded-2xl border border-slate-200 bg-white p-7 text-left shadow-soft hover:border-forest-300 disabled:opacity-50"
        >
          <FileSpreadsheet className="text-emerald-600" />
          <h3 className="mt-5 font-bold">Planilha Excel</h3>
          <p className="mt-1 text-sm text-slate-500">
            Exportar os {filtered.length} registros filtrados.
          </p>
        </button>
      </div>
    </main>
  );
}

function LoginScreen({ onLogin }: { onLogin: () => void }) {
  const [user, setUser] = useState(""),
    [password, setPassword] = useState(""),
    [show, setShow] = useState(false),
    [error, setError] = useState(""),
    [loading, setLoading] = useState(false);
  return (
    <div className="auth-screen flex min-h-screen items-center justify-center bg-[#eef0f2] p-4 sm:p-8">
      <div className="grid w-full max-w-[980px] overflow-hidden rounded-[32px] border border-white/80 bg-white shadow-[0_32px_90px_rgba(15,23,42,.14)] md:grid-cols-[.9fr_1.1fr]">
        <section className="relative hidden min-h-[650px] flex-col justify-between overflow-hidden bg-[#242424] p-12 text-white md:flex">
          <div className="absolute -right-24 -top-24 h-72 w-72 rounded-full border-[52px] border-white/[.035]" />
          <div className="absolute -bottom-28 -left-20 h-80 w-80 rounded-full border-[60px] border-white/[.03]" />
          <img src="/sacolao-abc-logo.png?v=4" alt="Sacolão ABC" className="relative h-28 w-60 object-contain object-left" />
          <div className="relative">
            <span className="text-[11px] font-bold uppercase tracking-[.28em] text-white/45">Portal administrativo</span>
            <h1 className="mt-5 text-4xl font-bold leading-tight">Pessoas e benefícios,<br/>em um só lugar.</h1>
            <p className="mt-5 max-w-sm text-sm leading-7 text-white/55">Acompanhe sua equipe, ocorrências de RH e cartões de passagem com segurança e organização.</p>
          </div>
          <p className="relative text-xs text-white/30">Sacolão ABC · Ambiente seguro</p>
        </section>
        <section className="flex min-h-[650px] flex-col justify-center px-7 py-12 sm:px-14 lg:px-16">
          <img src="/sacolao-abc-logo.png?v=4" alt="Sacolão ABC" className="mb-8 h-20 w-40 object-contain object-left md:hidden" />
          <div>
            <span className="text-xs font-bold uppercase tracking-[.2em] text-slate-400">Bem-vindo de volta</span>
            <h2 className="mt-3 text-3xl font-bold tracking-tight text-slate-900">Acesse sua conta</h2>
            <p className="mt-2 text-sm text-slate-500">Informe seus dados para entrar no sistema.</p>
          </div>
        <form
          onSubmit={async (e) => {
            e.preventDefault();
            setError("");
            setLoading(true);
            try {
              await cloudLogin(user.trim(), password);
              localStorage.setItem("valefluxo_session", "1");
              onLogin();
            } catch (reason) {
              setError(
                reason instanceof Error
                  ? reason.message
                  : "Não foi possível entrar.",
              );
            } finally {
              setLoading(false);
            }
          }}
          className="mt-9 space-y-5"
        >
          <label className="block text-sm font-semibold text-slate-700">Usuário
            <input
              autoFocus
              required
              value={user}
              onChange={(e) => setUser(e.target.value)}
              placeholder="Digite seu usuário"
              className="mt-2 h-14 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 outline-none transition focus:border-slate-500 focus:bg-white focus:ring-4 focus:ring-slate-100"
            />
          </label>
          <label className="block text-sm font-semibold text-slate-700">Senha
            <div className="relative mt-2">
              <input
                required
                type={show ? "text" : "password"}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Digite sua senha"
                className="h-14 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 pr-20 outline-none transition focus:border-slate-500 focus:bg-white focus:ring-4 focus:ring-slate-100"
              />
              <button
                type="button"
                onClick={() => setShow(!show)}
                className="absolute right-4 top-[18px] text-xs font-bold text-slate-500 hover:text-slate-900"
              >
                {show ? "Ocultar" : "Mostrar"}
              </button>
            </div>
          </label>
          {error && (
            <p className="rounded-2xl border border-red-100 bg-red-50 p-4 text-sm font-semibold text-red-600">
              {error}
            </p>
          )}
          <button
            disabled={loading}
            className="h-14 w-full rounded-2xl bg-[#262626] font-bold text-white shadow-lg shadow-black/10 transition hover:-translate-y-0.5 hover:bg-black disabled:opacity-60"
          >
            {loading ? "Entrando..." : "Entrar"}
          </button>
        </form>
          <div className="mt-8 flex items-center gap-3 text-xs text-slate-400"><span className="h-px flex-1 bg-slate-200"/>Acesso restrito e protegido<span className="h-px flex-1 bg-slate-200"/></div>
        </section>
      </div>
    </div>
  );
}

function InitialLoadingScreen() {
  return (
    <div className="initial-loading-screen" role="status" aria-label="Carregando">
      <span className="initial-loading-spinner" />
    </div>
  );
}

function ModuleMenu({
  select,
  onLogout,
  dark,
  toggleTheme,
  allowedModules,
}: {
  select: (module: Module) => void;
  onLogout: () => void;
  dark: boolean;
  toggleTheme: () => void;
  allowedModules: ModuleAccess[];
}) {
  return (
    <div className="module-menu min-h-screen bg-slate-100 px-4 py-8 dark:bg-[#111317] sm:px-8">
      <div className="mx-auto flex max-w-7xl items-center justify-between">
        <img
          src="/sacolao-abc-logo.png?v=4"
          alt="Sacolão ABC"
          className="h-16 w-40 object-contain"
        />
        <div className="flex gap-2">
          <button
            onClick={toggleTheme}
            className="module-control grid h-11 w-11 place-items-center rounded-xl border border-slate-200 bg-white text-slate-600 shadow-sm dark:border-slate-700 dark:bg-[#25272b] dark:text-white"
            title="Alternar tema"
          >
            {dark ? "☀" : "◐"}
          </button>
          <button
            onClick={onLogout}
            className="module-control rounded-xl border border-slate-200 bg-white px-4 text-sm font-semibold text-slate-600 shadow-sm hover:bg-slate-50 dark:border-slate-700 dark:bg-[#25272b] dark:text-white"
          >
            Sair
          </button>
        </div>
      </div>
      <main className="mx-auto flex min-h-[calc(100vh-130px)] max-w-7xl flex-col justify-center py-10">
        <div className="text-center">
          <span className="text-xs font-bold uppercase tracking-[.22em] text-slate-400">
            Menu principal
          </span>
          <h1 className="mt-3 text-3xl font-bold text-slate-900 dark:text-white sm:text-4xl">
            O que você deseja acessar?
          </h1>
          <p className="mx-auto mt-3 max-w-xl text-slate-500 dark:text-slate-400">
            Escolha uma área para continuar. Os funcionários cadastrados são
            compartilhados entre os módulos autorizados.
          </p>
        </div>
        <div
          className={`mx-auto mt-10 grid w-full gap-5 ${
            allowedModules.length === 1
              ? "max-w-lg grid-cols-1"
              : allowedModules.length === 2
                ? "max-w-4xl md:grid-cols-2"
                : allowedModules.length === 3
                  ? "max-w-7xl md:grid-cols-3"
                  : "max-w-7xl md:grid-cols-2 xl:grid-cols-4"
          }`}
        >
          {allowedModules.includes("people") && <button
            onClick={() => select("people")}
            className="module-card group rounded-3xl border border-slate-200 bg-white p-7 text-left shadow-soft transition hover:-translate-y-1 hover:border-slate-400 hover:shadow-xl dark:border-slate-700 dark:bg-[#25272b] dark:hover:border-slate-500 sm:p-9"
          >
            <div className="module-icon grid h-14 w-14 place-items-center rounded-2xl bg-[#262626] text-white">
              <Users size={28} />
            </div>
            <h2 className="mt-7 text-2xl font-bold text-slate-900 dark:text-white">
              Gestão de Pessoas
            </h2>
            <p className="mt-2 text-sm leading-6 text-slate-500 dark:text-slate-400">
              Funcionários, quadro por loja, experiências, aniversários, avisos,
              rescisões e ocorrências mensais.
            </p>
            <span className="mt-7 flex items-center gap-2 text-sm font-bold text-slate-800 dark:text-white">
              Entrar no RH{" "}
              <ChevronRight
                className="transition group-hover:translate-x-1"
                size={18}
              />
            </span>
          </button>}
          {allowedModules.includes("transit") && <button
            onClick={() => select("transit")}
            className="module-card group rounded-3xl border border-slate-200 bg-white p-7 text-left shadow-soft transition hover:-translate-y-1 hover:border-slate-400 hover:shadow-xl dark:border-slate-700 dark:bg-[#25272b] dark:hover:border-slate-500 sm:p-9"
          >
            <div className="module-icon grid h-14 w-14 place-items-center rounded-2xl bg-[#262626] text-white">
              <CreditCard size={28} />
            </div>
            <h2 className="mt-7 text-2xl font-bold text-slate-900 dark:text-white">
              Gestão de Cartões de Passagem
            </h2>
            <p className="mt-2 text-sm leading-6 text-slate-500 dark:text-slate-400">
              Planejamento das recargas, cartões BHBus e Ótimo, calendário,
              histórico e relatórios de transporte.
            </p>
            <span className="mt-7 flex items-center gap-2 text-sm font-bold text-slate-800 dark:text-white">
              Entrar nas passagens{" "}
              <ChevronRight
                className="transition group-hover:translate-x-1"
                size={18}
              />
            </span>
          </button>}
          {allowedModules.includes("finance") && <button
            onClick={() => select("finance")}
            className="module-card group rounded-3xl border border-slate-200 bg-white p-7 text-left shadow-soft transition hover:-translate-y-1 hover:border-slate-400 hover:shadow-xl dark:border-slate-700 dark:bg-[#25272b] dark:hover:border-slate-500 sm:p-9"
          >
            <div className="module-icon grid h-14 w-14 place-items-center rounded-2xl bg-[#262626] text-white">
              <DollarSign size={28} />
            </div>
            <h2 className="mt-7 text-2xl font-bold text-slate-900 dark:text-white">
              Gestão Financeira
            </h2>
            <p className="mt-2 text-sm leading-6 text-slate-500 dark:text-slate-400">
              Salários, adiantamentos do dia 20 e controle mensal das saídas de caixa.
            </p>
            <span className="mt-7 flex items-center gap-2 text-sm font-bold text-slate-800 dark:text-white">
              Entrar no financeiro
              <ChevronRight className="transition group-hover:translate-x-1" size={18} />
            </span>
          </button>}
          {allowedModules.includes("administrative") && <button
            onClick={() => select("administrative")}
            className="module-card group rounded-3xl border border-slate-200 bg-white p-7 text-left shadow-soft transition hover:-translate-y-1 hover:border-slate-400 hover:shadow-xl dark:border-slate-700 dark:bg-[#25272b] dark:hover:border-slate-500 sm:p-9"
          >
            <div className="module-icon grid h-14 w-14 place-items-center rounded-2xl bg-[#262626] text-white"><FileText size={28} /></div>
            <h2 className="mt-7 text-2xl font-bold text-slate-900 dark:text-white">Administrativo</h2>
            <p className="mt-2 text-sm leading-6 text-slate-500 dark:text-slate-400">Geração de recibos, advertências e documentos administrativos.</p>
            <span className="mt-7 flex items-center gap-2 text-sm font-bold text-slate-800 dark:text-white">Entrar no administrativo<ChevronRight className="transition group-hover:translate-x-1" size={18} /></span>
          </button>}
        </div>
      </main>
    </div>
  );
}

function HRModal({
  employee,
  close,
  save,
}: {
  employee: Recharge;
  close: () => void;
  save: (r: Recharge) => void;
}) {
  const defaultDays = Number(
      localStorage.getItem("valefluxo_experience_days") || 90,
    ),
    [form, setForm] = useState({
      birthDate: employee.birthDate || "",
      experienceDays: employee.experienceDays || defaultDays,
      employmentStatus: (employee.employmentStatus ||
        (employee.active === false ? "Desligado" : "Ativo")) as NonNullable<
        Recharge["employmentStatus"]
      >,
      noticeStart: employee.noticeStart || "",
      noticeEnd: employee.noticeEnd || "",
      terminationDate: employee.terminationDate || "",
      formalEmployment: employee.formalEmployment !== false,
      receivesTransit: employee.receivesTransit !== false,
    });
  return (
    <div className="fixed inset-0 z-[70] grid place-items-center bg-slate-950/45 p-4 backdrop-blur-sm">
      <div className="max-h-[92vh] w-full max-w-xl overflow-y-auto rounded-2xl bg-white shadow-2xl">
        <div className="flex items-center border-b border-slate-100 px-6 py-5">
          <div>
            <h3 className="text-lg font-bold">Dados de RH</h3>
            <p className="text-sm text-slate-500">
              {employee.employee} - {employee.store}
            </p>
          </div>
          <button
            onClick={close}
            className="ml-auto rounded-lg p-2 text-slate-400 hover:bg-slate-100"
          >
            <X size={19} />
          </button>
        </div>
        <form
          onSubmit={(e) => {
            e.preventDefault();
            save({
              ...employee,
              ...form,
              employmentStatus:
                form.employmentStatus as Recharge["employmentStatus"],
              active: form.employmentStatus !== "Desligado",
            });
          }}
          className="grid gap-4 p-6 sm:grid-cols-2"
        >
          <label className="text-sm font-semibold text-slate-700">
            Data de nascimento
            <input
              type="date"
              value={form.birthDate}
              onChange={(e) => setForm({ ...form, birthDate: e.target.value })}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            />
          </label>
          <label className="text-sm font-semibold text-slate-700">
            Dias de experiência
            <input
              min="1"
              max="365"
              type="number"
              value={form.experienceDays}
              onChange={(e) =>
                setForm({ ...form, experienceDays: Number(e.target.value) })
              }
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            />
          </label>
          <label className="text-sm font-semibold text-slate-700 sm:col-span-2">
            Situação
            <select
              value={form.employmentStatus}
              onChange={(e) =>
                setForm({
                  ...form,
                  employmentStatus: e.target.value as NonNullable<
                    Recharge["employmentStatus"]
                  >,
                })
              }
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            >
              {["Ativo", "Experiência", "Aviso prévio", "Desligado"].map(
                (x) => (
                  <option key={x}>{x}</option>
                ),
              )}
            </select>
          </label>
          <label className="text-sm font-semibold text-slate-700">
            Início do aviso
            <input
              type="date"
              value={form.noticeStart}
              onChange={(e) =>
                setForm({ ...form, noticeStart: e.target.value })
              }
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            />
          </label>
          <label className="text-sm font-semibold text-slate-700">
            Fim do aviso
            <input
              type="date"
              min={form.noticeStart}
              value={form.noticeEnd}
              onChange={(e) => setForm({ ...form, noticeEnd: e.target.value })}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            />
          </label>
          <label className="text-sm font-semibold text-slate-700 sm:col-span-2">
            Data da rescisão
            <input
              type="date"
              value={form.terminationDate}
              onChange={(e) =>
                setForm({ ...form, terminationDate: e.target.value })
              }
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            />
          </label>
          <label className="text-sm font-semibold text-slate-700">
            Carteira assinada
            <select
              value={form.formalEmployment ? "Sim" : "Não"}
              onChange={(e) =>
                setForm({ ...form, formalEmployment: e.target.value === "Sim" })
              }
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            >
              <option>Sim</option>
              <option>Não</option>
            </select>
          </label>
          <label className="text-sm font-semibold text-slate-700">
            Recebe passagem
            <select
              value={form.receivesTransit ? "Sim" : "Não"}
              onChange={(e) =>
                setForm({ ...form, receivesTransit: e.target.value === "Sim" })
              }
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-3"
            >
              <option>Sim</option>
              <option>Não</option>
            </select>
          </label>
          <div className="flex justify-end gap-2 border-t border-slate-100 pt-5 sm:col-span-2">
            <button
              type="button"
              onClick={close}
              className="rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-semibold text-slate-600"
            >
              Cancelar
            </button>
            <button className="rounded-xl bg-forest-700 px-5 py-2.5 text-sm font-semibold text-white">
              Salvar dados de RH
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

function FinancePage({
  employees,
  entries,
  setEntries,
  period,
  setPeriod,
}: {
  employees: Recharge[];
  entries: FinancialEntry[];
  setEntries: (entries: FinancialEntry[]) => void;
  period: string;
  setPeriod: (period: string) => void;
}) {
  const [generalSalaryDate, setGeneralSalaryDate] = useState("");
  const [generalAdvanceDate, setGeneralAdvanceDate] = useState("");
  const [financialDetail, setFinancialDetail] = useState<{
    key: "salary" | "advance" | "vacation" | "severance" | "pending" | "paid" | "total";
    title: string;
  } | null>(null);
  const visibleEmployeeIds = new Set(employees.map((employee) => employee.id));
  const current = entries.filter(
      (entry) => entry.period === period && visibleEmployeeIds.has(entry.employeeId),
    ),
    byEmployee = new Map(current.map((entry) => [entry.employeeId, entry])),
    todayIso = new Date().toISOString().slice(0, 10),
    paymentReached = (date?: string) => !!date && date <= todayIso,
    money = (value: number) =>
      (value || 0).toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL",
      }),
    plannedSalary = current.reduce((sum, entry) => sum + entry.salary, 0),
    plannedAdvance = current.reduce((sum, entry) => sum + entry.advance, 0),
    plannedVacation = current.reduce((sum, entry) => sum + (entry.vacation || 0), 0),
    plannedSeverance = current.reduce((sum, entry) => sum + (entry.severance || 0), 0),
    paid = current.reduce(
      (sum, entry) =>
        sum +
        (paymentReached(entry.salaryPaidAt) ? entry.salary : 0) +
        (paymentReached(entry.advancePaidAt) ? entry.advance : 0) +
        (paymentReached(entry.vacationPaidAt) ? entry.vacation || 0 : 0) +
        (paymentReached(entry.severancePaidAt) ? entry.severance || 0 : 0),
      0,
    ),
    planned = plannedSalary + plannedAdvance + plannedVacation + plannedSeverance,
    storeBreakdown = [...new Set(employees.map((employee) => employee.store))]
      .sort((a, b) => a.localeCompare(b, "pt-BR"))
      .map((store) => {
        const ids = new Set(
          employees
            .filter((employee) => employee.store === store)
            .map((employee) => employee.id),
        );
        return {
          store,
          salary: current
            .filter((entry) => ids.has(entry.employeeId))
            .reduce((sum, entry) => sum + entry.salary, 0),
          advance: current
            .filter((entry) => ids.has(entry.employeeId))
            .reduce((sum, entry) => sum + entry.advance, 0),
        };
      }),
    roleBreakdown = [...new Set(employees.map((employee) => employee.role || "Sem função"))]
      .sort((a, b) => a.localeCompare(b, "pt-BR"))
      .map((role) => {
        const roleEmployees = employees.filter(
            (employee) => (employee.role || "Sem função") === role,
          ),
          ids = new Set(roleEmployees.map((employee) => employee.id)),
          roleEntries = current.filter((entry) => ids.has(entry.employeeId));
        return {
          role,
          employees: roleEmployees.length,
          salary: roleEntries.reduce((sum, entry) => sum + entry.salary, 0),
          advance: roleEntries.reduce((sum, entry) => sum + entry.advance, 0),
          total: roleEntries.reduce(
            (sum, entry) =>
              sum +
              entry.salary +
              entry.advance +
              (entry.vacation || 0) +
              (entry.severance || 0),
            0,
          ),
        };
      }),
    update = (employeeId: number, patch: Partial<FinancialEntry>) => {
      const existing = byEmployee.get(employeeId),
        next: FinancialEntry = {
          id: existing?.id || Date.now() + employeeId,
          employeeId,
          period,
          salary: existing?.salary || 0,
          advance: existing?.advance || 0,
          vacation: existing?.vacation || 0,
          severance: existing?.severance || 0,
          salaryPaidAt: existing?.salaryPaidAt,
          advancePaidAt: existing?.advancePaidAt,
          vacationPaidAt: existing?.vacationPaidAt,
          severancePaidAt: existing?.severancePaidAt,
          noPayments: existing?.noPayments,
          noPaymentsFrom: existing?.noPaymentsFrom,
          ...patch,
        };
      setEntries([
        ...entries.filter(
          (entry) => !(entry.employeeId === employeeId && entry.period === period),
        ),
        next,
      ]);
    },
    applyDateToAll = (field: "salaryPaidAt" | "advancePaidAt", date: string) => {
      const employeeIds = new Set(employees.map((employee) => employee.id)),
        currentEntries = new Map(
          entries
            .filter((entry) => entry.period === period && employeeIds.has(entry.employeeId))
            .map((entry) => [entry.employeeId, entry]),
        ),
        updated = employees.map((employee) => {
          const existing = currentEntries.get(employee.id);
          return {
            id: existing?.id || Date.now() + employee.id,
            employeeId: employee.id,
            period,
            salary: existing?.salary || 0,
            advance: existing?.advance || 0,
            vacation: existing?.vacation || 0,
            severance: existing?.severance || 0,
            salaryPaidAt: existing?.salaryPaidAt,
            advancePaidAt: existing?.advancePaidAt,
            vacationPaidAt: existing?.vacationPaidAt,
            severancePaidAt: existing?.severancePaidAt,
            noPayments: existing?.noPayments,
            noPaymentsFrom: existing?.noPaymentsFrom,
            [field]: date || undefined,
          } as FinancialEntry;
        });
      setEntries([
        ...entries.filter(
          (entry) => !(entry.period === period && employeeIds.has(entry.employeeId)),
        ),
        ...updated,
      ]);
      alert(
        `Data aplicada para ${employees.length} funcionário(s) exibido(s).`,
      );
    };
  useEffect(() => {
    setGeneralSalaryDate("");
    setGeneralAdvanceDate("");
  }, [period]);
  const employeeById = new Map(employees.map((employee) => [employee.id, employee]));
  const detailRows = financialDetail
    ? current
        .map((entry) => {
          const employee = employeeById.get(entry.employeeId),
            categories = [
              { key: "salary", label: "Salário", value: entry.salary, date: entry.salaryPaidAt },
              { key: "advance", label: "Adiantamento", value: entry.advance, date: entry.advancePaidAt },
              { key: "vacation", label: "Férias", value: entry.vacation || 0, date: entry.vacationPaidAt },
              { key: "severance", label: "Verbas rescisórias", value: entry.severance || 0, date: entry.severancePaidAt },
            ],
            selected =
              financialDetail.key === "paid"
                ? categories.filter((item) => item.value > 0 && paymentReached(item.date))
                : financialDetail.key === "pending"
                  ? categories.filter((item) => item.value > 0 && !paymentReached(item.date))
                  : financialDetail.key === "total"
                    ? categories.filter((item) => item.value > 0)
                  : categories.filter((item) => item.key === financialDetail.key && item.value > 0);
          return {
            employee,
            amount: selected.reduce((sum, item) => sum + item.value, 0),
            items: selected,
          };
        })
        .filter((row) => row.employee && row.amount > 0)
        .sort((a, b) => b.amount - a.amount)
    : [];
  return (
    <main className="fade-in p-4 sm:p-7">
      <div className="mb-7 flex flex-wrap items-end justify-between gap-4">
        <SectionHead
          title="Gestão Financeira"
          sub="Salários, adiantamentos e controle de saída de caixa"
        />
        <label className="text-xs font-semibold text-slate-500">
          Mês de referência
          <input
            type="month"
            value={period}
            onChange={(event) => setPeriod(event.target.value)}
            className="mt-1 block rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold outline-none"
          />
        </label>
      </div>
      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-3">
        {[
          { key: "salary" as const, title: "Pagamento de salário", value: plannedSalary, sub: "Folha cadastrada", icon: Banknote, card: "border-t-blue-500 bg-blue-50/60", iconStyle: "bg-blue-100 text-blue-700", valueStyle: "text-blue-700" },
          { key: "advance" as const, title: "Pagamento de adiantamento", value: plannedAdvance, sub: "Valores cadastrados", icon: HandCoins, card: "border-t-violet-500 bg-violet-50/60", iconStyle: "bg-violet-100 text-violet-700", valueStyle: "text-violet-700" },
          { key: "vacation" as const, title: "Férias", value: plannedVacation, sub: "Valores cadastrados", icon: Umbrella, card: "border-t-amber-500 bg-amber-50/70", iconStyle: "bg-amber-100 text-amber-700", valueStyle: "text-amber-700" },
          { key: "severance" as const, title: "Verbas rescisórias", value: plannedSeverance, sub: "Valores cadastrados", icon: ReceiptText, card: "border-t-red-500 bg-red-50/60", iconStyle: "bg-red-100 text-red-700", valueStyle: "text-red-700" },
          { key: "pending" as const, title: "Pendente", value: Math.max(0, planned - paid), sub: "Saída prevista", icon: Hourglass, card: "border-t-orange-500 bg-orange-50/70", iconStyle: "bg-orange-100 text-orange-700", valueStyle: "text-orange-700" },
          { key: "paid" as const, title: "Total já pago", value: paid, sub: "Saída realizada", icon: BadgeCheck, card: "border-t-green-500 bg-green-50/60", iconStyle: "bg-green-100 text-green-700", valueStyle: "text-green-700" },
        ].map(({ key, title, value, sub, icon: Icon, card, iconStyle, valueStyle }) => (
          <button type="button" onClick={() => setFinancialDetail({ key, title })} key={title} className={`rounded-2xl border border-t-4 border-slate-200 bg-white p-5 text-left shadow-soft transition hover:-translate-y-0.5 hover:shadow-lg ${card}`}>
            <div className={`grid h-11 w-11 place-items-center rounded-xl ${iconStyle}`}>
              <Icon size={21} />
            </div>
            <div className={`mt-4 text-2xl font-bold ${valueStyle}`}>{money(Number(value))}</div>
            <div className="mt-1 text-sm font-semibold text-slate-700">{title}</div>
            <div className="mt-1 text-xs text-slate-400">{sub}</div>
          </button>
        ))}
      </div>
      <button
        type="button"
        onClick={() => setFinancialDetail({ key: "total", title: "Total geral" })}
        className="mt-4 flex w-full items-center justify-between rounded-2xl border border-t-4 border-cyan-200 border-t-cyan-600 bg-gradient-to-r from-cyan-50 to-sky-100 p-5 text-left shadow-soft transition hover:-translate-y-0.5 hover:shadow-lg"
      >
        <div className="flex items-center gap-4">
          <div className="grid h-12 w-12 place-items-center rounded-xl bg-cyan-100 text-cyan-800">
            <DollarSign size={23} />
          </div>
          <div>
            <div className="text-sm font-bold text-slate-800">Total geral</div>
            <div className="mt-1 text-xs text-slate-400">Pendente + total já pago</div>
          </div>
        </div>
        <div className="text-2xl font-black text-cyan-900">{money(Math.max(0, planned - paid) + paid)}</div>
      </button>
      <div className="mt-5 grid gap-4 lg:grid-cols-2">
        {[
          {
            title: "Pagamento de salário por loja",
            field: "salary" as const,
          },
          {
            title: "Pagamento de adiantamento por loja",
            field: "advance" as const,
          },
        ].map((indicator) => (
          <div
            key={indicator.title}
            className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft"
          >
            <div className="flex items-center gap-3">
              <div className="grid h-10 w-10 place-items-center rounded-xl bg-slate-100 text-slate-700">
                <Store size={19} />
              </div>
              <div>
                <h3 className="font-bold">{indicator.title}</h3>
                <p className="text-xs text-slate-400">Valores do mês selecionado</p>
              </div>
            </div>
            <div className="mt-4 divide-y divide-slate-100">
              {storeBreakdown.map((item) => (
                <div key={item.store} className="flex items-center justify-between gap-4 py-3">
                  <span className="text-sm font-semibold text-slate-600">{item.store}</span>
                  <b className="text-sm">{money(item[indicator.field])}</b>
                </div>
              ))}
              {!storeBreakdown.length && (
                <p className="py-5 text-center text-sm text-slate-400">
                  Nenhuma loja encontrada para os filtros selecionados.
                </p>
              )}
            </div>
          </div>
        ))}
      </div>
      <section className="mt-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <div className="flex items-center gap-3">
          <div className="grid h-10 w-10 place-items-center rounded-xl bg-slate-900 text-white">
            <Users size={19} />
          </div>
          <div>
            <h3 className="font-bold text-slate-900">Visão financeira por função</h3>
            <p className="text-xs font-medium text-slate-500">
              Valores separados por Repositor, Operador, Vigia e demais funções
            </p>
          </div>
        </div>
        <div className="mt-5 grid gap-4 sm:grid-cols-2 xl:grid-cols-3">
          {roleBreakdown.map((item) => (
            <div
              key={item.role}
              className="rounded-xl border border-slate-200 bg-slate-50 p-4"
            >
              <div className="flex items-start justify-between gap-3">
                <div>
                  <h4 className="font-bold text-slate-900">{item.role}</h4>
                  <p className="mt-0.5 text-xs font-semibold text-slate-500">
                    {item.employees} funcionário(s)
                  </p>
                </div>
                <span className="rounded-lg bg-white px-2.5 py-1 text-xs font-bold text-slate-600 shadow-sm">
                  {item.employees}
                </span>
              </div>
              <div className="mt-4 space-y-2 border-t border-slate-200 pt-3 text-sm">
                <div className="flex justify-between gap-3 text-slate-600">
                  <span>Salários</span><b className="text-slate-900">{money(item.salary)}</b>
                </div>
                <div className="flex justify-between gap-3 text-slate-600">
                  <span>Adiantamentos</span><b className="text-slate-900">{money(item.advance)}</b>
                </div>
                <div className="flex justify-between gap-3 rounded-lg bg-slate-900 px-3 py-2 text-white">
                  <span className="font-semibold">Total da função</span><b>{money(item.total)}</b>
                </div>
              </div>
            </div>
          ))}
          {!roleBreakdown.length && (
            <p className="py-5 text-sm text-slate-400">
              Nenhuma função encontrada para os filtros selecionados.
            </p>
          )}
        </div>
      </section>
      {false && <>
      <div className="mt-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <div>
          <h3 className="font-bold">Datas gerais de pagamento</h3>
          <p className="text-xs text-slate-400">
            Escolha a data e clique em Aplicar a todos. Depois, você poderá alterar exceções diretamente na tabela.
          </p>
        </div>
        <div className="mt-4 grid gap-4 md:grid-cols-2">
          <label className="text-sm font-semibold text-slate-600">
            Pagamento de salário
            <div className="mt-2 flex gap-2">
              <input
                type="date"
                value={generalSalaryDate}
                onChange={(event) => setGeneralSalaryDate(event.target.value)}
                className="min-w-0 flex-1 rounded-xl border border-slate-200 bg-white px-4 py-3 outline-none focus:border-slate-400"
              />
              <button
                type="button"
                disabled={!generalSalaryDate}
                onClick={() => applyDateToAll("salaryPaidAt", generalSalaryDate)}
                className="rounded-xl bg-forest-700 px-4 py-3 text-xs font-bold text-white disabled:cursor-not-allowed disabled:opacity-40"
              >
                Aplicar a todos
              </button>
            </div>
          </label>
          <label className="text-sm font-semibold text-slate-600">
            Pagamento de adiantamento
            <div className="mt-2 flex gap-2">
              <input
                type="date"
                value={generalAdvanceDate}
                onChange={(event) => setGeneralAdvanceDate(event.target.value)}
                className="min-w-0 flex-1 rounded-xl border border-slate-200 bg-white px-4 py-3 outline-none focus:border-slate-400"
              />
              <button
                type="button"
                disabled={!generalAdvanceDate}
                onClick={() => applyDateToAll("advancePaidAt", generalAdvanceDate)}
                className="rounded-xl bg-forest-700 px-4 py-3 text-xs font-bold text-white disabled:cursor-not-allowed disabled:opacity-40"
              >
                Aplicar a todos
              </button>
            </div>
          </label>
        </div>
      </div>
      <div className="mt-5 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="border-b border-slate-100 px-5 py-4">
          <h3 className="font-bold">Folha mensal por funcionário</h3>
          <p className="text-xs text-slate-400">
            Informe os valores e registre a data em que cada pagamento saiu do caixa.
          </p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[1450px] text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase text-slate-400">
              <tr>
                <th className="px-5 py-3">Funcionário</th>
                <th className="px-5 py-3">Loja / Função</th>
                <th className="px-5 py-3">Salário do mês</th>
                <th className="px-5 py-3">Pagamento do salário</th>
                <th className="px-5 py-3">Adiantamento dia 20</th>
                <th className="px-5 py-3">Pagamento do adiantamento</th>
                <th className="px-5 py-3">Férias</th>
                <th className="px-5 py-3">Verbas rescisórias</th>
                <th className="px-5 py-3">Total</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {employees.map((employee) => {
                const entry = byEmployee.get(employee.id);
                return (
                  <tr key={employee.id}>
                    <td className="px-5 py-4 font-semibold">
                      {employee.employee}
                      {employee.terminationDate && (
                        <div className="mt-1 text-xs font-semibold text-red-600">
                          Desligado em {formatDate(employee.terminationDate)}
                        </div>
                      )}
                    </td>
                    <td className="px-5 py-4">
                      {employee.store}<div className="text-xs text-slate-400">{employee.role}</div>
                    </td>
                    <td className="px-5 py-4">
                      <input type="number" min="0" step="0.01" value={entry?.salary || ""} onChange={(event) => update(employee.id, { salary: Number(event.target.value) })} placeholder="R$ 0,00" className="w-36 rounded-lg border border-slate-200 px-3 py-2" />
                    </td>
                    <td className="px-5 py-4">
                      <input type="date" value={entry?.salaryPaidAt || ""} onChange={(event) => update(employee.id, { salaryPaidAt: event.target.value || undefined })} className="rounded-lg border border-slate-200 px-3 py-2" />
                    </td>
                    <td className="px-5 py-4">
                      <input type="number" min="0" step="0.01" value={entry?.advance || ""} onChange={(event) => update(employee.id, { advance: Number(event.target.value) })} placeholder="R$ 0,00" className="w-36 rounded-lg border border-slate-200 px-3 py-2" />
                    </td>
                    <td className="px-5 py-4">
                      <input type="date" value={entry?.advancePaidAt || ""} onChange={(event) => update(employee.id, { advancePaidAt: event.target.value || undefined })} className="rounded-lg border border-slate-200 px-3 py-2" />
                    </td>
                    <td className="px-5 py-4">{money(entry?.vacation || 0)}</td>
                    <td className="px-5 py-4">{money(entry?.severance || 0)}</td>
                    <td className="px-5 py-4 font-bold">{money((entry?.salary || 0) + (entry?.advance || 0) + (entry?.vacation || 0) + (entry?.severance || 0))}</td>
                  </tr>
                );
              })}
              {!employees.length && (
                <tr><td colSpan={9} className="py-12 text-center text-slate-400">Nenhum funcionário encontrado para os filtros selecionados.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
      </>}
      {financialDetail && createPortal(
        <div className="fixed inset-0 z-[100] grid place-items-center bg-slate-950/45 p-4 backdrop-blur-sm" onMouseDown={() => setFinancialDetail(null)}>
          <div className="max-h-[82vh] w-full max-w-3xl overflow-hidden rounded-2xl bg-white shadow-2xl" onMouseDown={(event) => event.stopPropagation()}>
            <div className="flex items-center border-b border-slate-200 px-6 py-5">
              <div><h3 className="text-xl font-bold">{financialDetail.title}</h3><p className="text-sm text-slate-400">{capitalizeMonth(new Date(`${period}-01T12:00:00`).toLocaleDateString("pt-BR", { month: "long", year: "numeric" }))} · {detailRows.length} funcionário(s)</p></div>
              <button type="button" onClick={() => setFinancialDetail(null)} className="ml-auto rounded-lg p-2 text-slate-500 hover:bg-slate-100"><X size={20} /></button>
            </div>
            <div className="max-h-[65vh] overflow-y-auto p-5">
              <div className="space-y-3">
                {detailRows.map(({ employee, amount, items }) => (
                  <div key={employee!.id} className="rounded-xl border border-slate-200 p-4">
                    <div className="flex flex-wrap items-start justify-between gap-3"><div><b>{employee!.employee}</b><div className="text-xs text-slate-400">{employee!.store} · {employee!.role}</div></div><b className="text-lg">{money(amount)}</b></div>
                    <div className="mt-3 flex flex-wrap gap-2">{items.map((item) => <span key={item.key} className="rounded-lg bg-slate-100 px-3 py-2 text-xs text-slate-600"><b>{item.label}:</b> {money(item.value)} · {item.date ? formatDate(item.date) : "Sem data"}</span>)}</div>
                  </div>
                ))}
                {!detailRows.length && <p className="py-12 text-center text-sm text-slate-400">Nenhum lançamento nesta categoria.</p>}
              </div>
            </div>
          </div>
        </div>,
        document.body,
      )}
    </main>
  );
}

function FinancialRegistrations({
  employees,
  entries,
  setEntries,
  period,
  setPeriod,
  showHidden = false,
}: {
  employees: Recharge[];
  entries: FinancialEntry[];
  setEntries: (entries: FinancialEntry[]) => void;
  period: string;
  setPeriod: (period: string) => void;
  showHidden?: boolean;
}) {
  const [generalSalaryDate, setGeneralSalaryDate] = useState("");
  const [generalAdvanceDate, setGeneralAdvanceDate] = useState("");
  const [employeeSearch, setEmployeeSearch] = useState("");
  const excludedEmployees = new Set(
      entries
        .filter((entry) => !!entry.noPaymentsFrom && entry.noPaymentsFrom <= period)
        .map((entry) => entry.employeeId),
    ),
    searchedEmployees = employees
      .filter((employee) => showHidden || !excludedEmployees.has(employee.id))
      .filter((employee) =>
        employee.employee
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .toLowerCase()
          .includes(
            employeeSearch
              .normalize("NFD")
              .replace(/[\u0300-\u036f]/g, "")
              .toLowerCase()
              .trim(),
          ),
      );
  const byEmployee = new Map(
      entries.filter((entry) => entry.period === period).map((entry) => [entry.employeeId, entry]),
    ),
    update = (employeeId: number, patch: Partial<FinancialEntry>) => {
      const existing = byEmployee.get(employeeId),
        next: FinancialEntry = {
          id: existing?.id || Date.now() + employeeId,
          employeeId,
          period,
          salary: existing?.salary || 0,
          advance: existing?.advance || 0,
          vacation: existing?.vacation || 0,
          severance: existing?.severance || 0,
          salaryPaidAt: existing?.salaryPaidAt,
          advancePaidAt: existing?.advancePaidAt,
          vacationPaidAt: existing?.vacationPaidAt,
          severancePaidAt: existing?.severancePaidAt,
          noPayments: existing?.noPayments,
          noPaymentsFrom: existing?.noPaymentsFrom,
          ...patch,
        };
      setEntries([
        ...entries.filter((entry) => !(entry.employeeId === employeeId && entry.period === period)),
        next,
      ]);
    },
    applyToAll = (field: "salaryPaidAt" | "advancePaidAt", date: string) => {
      if (!date) return;
      const employeeIds = new Set(searchedEmployees.map((employee) => employee.id)),
        updated = searchedEmployees.map((employee) => {
          const existing = byEmployee.get(employee.id);
          return {
            id: existing?.id || Date.now() + employee.id,
            employeeId: employee.id,
            period,
            salary: existing?.salary || 0,
            advance: existing?.advance || 0,
            vacation: existing?.vacation || 0,
            severance: existing?.severance || 0,
            salaryPaidAt: existing?.salaryPaidAt,
            advancePaidAt: existing?.advancePaidAt,
            vacationPaidAt: existing?.vacationPaidAt,
            severancePaidAt: existing?.severancePaidAt,
            noPayments: existing?.noPayments,
            noPaymentsFrom: existing?.noPaymentsFrom,
            [field]: date,
          } as FinancialEntry;
        });
      setEntries([
        ...entries.filter((entry) => !(entry.period === period && employeeIds.has(entry.employeeId))),
        ...updated,
      ]);
      alert(`Data aplicada para ${searchedEmployees.length} funcionário(s).`);
    };
  useEffect(() => {
    setGeneralSalaryDate("");
    setGeneralAdvanceDate("");
  }, [period]);
  return (
    <main className="fade-in p-4 sm:p-7">
      <div className="mb-7 flex flex-wrap items-end justify-between gap-4">
        <SectionHead title="Cadastros financeiros" sub="Cadastre salário, adiantamento, férias e verbas rescisórias" />
        <label className="text-xs font-semibold text-slate-500">Mês de referência<input type="month" value={period} onChange={(event) => setPeriod(event.target.value)} className="mt-1 block rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold outline-none" /></label>
      </div>
      <div className="mb-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <h3 className="font-bold">Datas gerais</h3>
        <p className="text-xs text-slate-400">Defina as datas padrão do salário e do adiantamento para todos os funcionários exibidos.</p>
        <div className="mt-4 grid gap-4 md:grid-cols-2">
          <label className="text-sm font-semibold text-slate-600">Data do salário<div className="mt-2 flex gap-2"><input type="date" value={generalSalaryDate} onChange={(event) => setGeneralSalaryDate(event.target.value)} className="min-w-0 flex-1 rounded-xl border border-slate-200 px-4 py-3" /><button type="button" disabled={!generalSalaryDate} onClick={() => applyToAll("salaryPaidAt", generalSalaryDate)} className="rounded-xl bg-forest-700 px-4 py-3 text-xs font-bold text-white disabled:opacity-40">Aplicar a todos</button></div></label>
          <label className="text-sm font-semibold text-slate-600">Data do adiantamento<div className="mt-2 flex gap-2"><input type="date" value={generalAdvanceDate} onChange={(event) => setGeneralAdvanceDate(event.target.value)} className="min-w-0 flex-1 rounded-xl border border-slate-200 px-4 py-3" /><button type="button" disabled={!generalAdvanceDate} onClick={() => applyToAll("advancePaidAt", generalAdvanceDate)} className="rounded-xl bg-forest-700 px-4 py-3 text-xs font-bold text-white disabled:opacity-40">Aplicar a todos</button></div></label>
        </div>
      </div>
      <div className="mb-5 rounded-2xl border border-slate-200 bg-white p-4 shadow-soft">
        <div className="relative">
          <Search className="pointer-events-none absolute left-4 top-3.5 text-slate-400" size={19} />
          <input
            type="search"
            value={employeeSearch}
            onChange={(event) => setEmployeeSearch(event.target.value)}
            placeholder="Pesquisar funcionário pelo nome..."
            aria-label="Pesquisar funcionário pelo nome"
            className="w-full rounded-xl border border-slate-200 bg-white py-3 pl-11 pr-4 outline-none focus:border-slate-400"
          />
        </div>
        <p className="mt-2 text-xs text-slate-400">
          Mostrando {searchedEmployees.length} de {employees.length} funcionário(s)
        </p>
      </div>
      <div className="space-y-4">
        {searchedEmployees.map((employee) => {
          const entry = byEmployee.get(employee.id);
          return (
            <div key={employee.id} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div><h3 className="text-lg font-bold text-slate-900">{employee.employee}</h3><p className="mt-0.5 text-sm font-semibold text-slate-500">{employee.store} · {employee.role}</p></div>
                {(entry?.noPayments || excludedEmployees.has(employee.id)) && <span className="rounded-full bg-red-50 px-3 py-1 text-xs font-bold text-red-600">Salvo e ocultado</span>}
              </div>
              <div className="mt-4 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                {([
                  ["Salário mensal", "salary", "salaryPaidAt"],
                  ["Adiantamento", "advance", "advancePaidAt"],
                  ["Férias", "vacation", "vacationPaidAt"],
                  ["Verbas rescisórias", "severance", "severancePaidAt"],
                ] as Array<[string, "salary" | "advance" | "vacation" | "severance", "salaryPaidAt" | "advancePaidAt" | "vacationPaidAt" | "severancePaidAt"]>).map(([label, field, dateField]) => (
                  <div key={field} className="rounded-xl border border-slate-300 bg-slate-50 p-4 shadow-sm"><label className="text-sm font-bold text-slate-700">{label}<input type="number" min="0" step="0.01" value={entry?.[field] || ""} onChange={(event) => update(employee.id, { [field]: Number(event.target.value), noPayments: false })} placeholder="R$ 0,00" className="mt-2 w-full rounded-lg border border-slate-300 bg-white px-3 py-3 font-bold text-slate-900 outline-none focus:border-slate-500" /></label><label className="mt-4 block text-sm font-bold text-slate-700">Data do pagamento<input type="date" value={entry?.[dateField] || ""} onChange={(event) => update(employee.id, { [dateField]: event.target.value || undefined, noPayments: false })} className="mt-2 w-full rounded-lg border border-slate-300 bg-white px-3 py-3 font-bold text-slate-900 outline-none focus:border-slate-500" /></label></div>
                ))}
              </div>
              <div className="mt-4 flex gap-2">{showHidden && excludedEmployees.has(employee.id) ? <button type="button" onClick={() => setEntries(entries.map((item) => item.employeeId === employee.id ? { ...item, noPayments: false, noPaymentsFrom: undefined } : item))} className="rounded-xl bg-forest-700 px-4 py-2.5 text-xs font-bold text-white">Voltar para cadastros</button> : employee.terminationDate ? <button type="button" onClick={() => update(employee.id, { noPayments: true, noPaymentsFrom: period })} className="rounded-xl border border-red-200 px-4 py-2.5 text-xs font-bold text-red-600">Salvar e ocultar</button> : null}</div>
            </div>
          );
        })}
        {!searchedEmployees.length && (
          <div className="rounded-2xl border border-slate-200 bg-white py-12 text-center text-sm text-slate-400 shadow-soft">
            Nenhum funcionário encontrado.
          </div>
        )}
      </div>
    </main>
  );
}

function TaxesPage({
  entries,
  setEntries,
  companies,
  period,
  setPeriod,
}: {
  entries: TaxEntry[];
  setEntries: (entries: TaxEntry[]) => void;
  companies: string[];
  period: string;
  setPeriod: (period: string) => void;
}) {
  const blank = { description: "", category: "", dueDate: `${period}-10`, amount: "", paid: false, paidAt: "", note: "" };
  const [form, setForm] = useState(blank),
    [editingId, setEditingId] = useState<number | null>(null),
    [query, setQuery] = useState(""),
    [exportStatus, setExportStatus] = useState<"Todos" | "Pagos" | "Não pagos">("Todos");
  useEffect(() => {
    if (!editingId) setForm((current) => ({ ...current, dueDate: `${period}-10` }));
  }, [period]);
  const periodEntries = entries.filter((entry) => entry.period === period),
    monthEntries = periodEntries
      .filter((entry) => `${entry.description} ${entry.category} ${entry.note || ""}`.toLowerCase().includes(query.toLowerCase()))
      .sort((a, b) => a.dueDate.localeCompare(b.dueDate) || a.description.localeCompare(b.description, "pt-BR")),
    total = periodEntries.reduce((sum, entry) => sum + entry.amount, 0),
    paidTotal = periodEntries.filter((entry) => entry.paid).reduce((sum, entry) => sum + entry.amount, 0),
    pendingTotal = total - paidTotal,
    companyNames = [...new Set(periodEntries.map((entry) => entry.category || "Sem empresa"))].sort((a, b) => a.localeCompare(b, "pt-BR")),
    companyGroups = companyNames.map((company) => ({ company, rows: monthEntries.filter((entry) => (entry.category || "Sem empresa") === company), allRows: periodEntries.filter((entry) => (entry.category || "Sem empresa") === company) })),
    money = (value: number) => value.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }),
    formatDateBr = (date: string) => date ? new Date(`${date}T12:00:00`).toLocaleDateString("pt-BR") : "—";
  const reset = () => {
    setEditingId(null);
    setForm({ ...blank, dueDate: `${period}-10` });
  };
  const submit = (event: FormEvent) => {
    event.preventDefault();
    const entry: TaxEntry = {
      id: editingId || Date.now(), period, description: form.description.trim(), category: form.category.trim(),
      dueDate: form.dueDate, amount: parseMoney(form.amount), paid: form.paid,
      paidAt: form.paid ? form.paidAt || undefined : undefined, note: form.note.trim() || undefined,
    };
    setEntries(editingId ? entries.map((item) => item.id === editingId ? entry : item) : [...entries, entry]);
    window.dispatchEvent(new CustomEvent("abc:toast", { detail: editingId ? "Guia atualizada" : "Guia cadastrada" }));
    reset();
  };
  const startEdit = (entry: TaxEntry) => {
    setEditingId(entry.id);
    setForm({ description: entry.description, category: entry.category, dueDate: entry.dueDate, amount: entry.amount.toLocaleString("pt-BR", { minimumFractionDigits: 2 }), paid: entry.paid, paidAt: entry.paidAt || "", note: entry.note || "" });
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const exportExcel = async () => {
    const exportEntries = periodEntries.filter((entry) => exportStatus === "Todos" || (exportStatus === "Pagos" ? entry.paid : !entry.paid)),
      exportCompanies = [...new Set(exportEntries.map((entry) => entry.category || "Sem empresa"))].sort((a, b) => a.localeCompare(b, "pt-BR")),
      exportGroups = exportCompanies.map((company) => ({ company, rows: exportEntries.filter((entry) => (entry.category || "Sem empresa") === company) })),
      exportTotal = exportEntries.reduce((sum, entry) => sum + entry.amount, 0),
      workbook = new ExcelJS.Workbook(),
      sheet = workbook.addWorksheet("Impostos", { views: [{ showGridLines: false }], pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 } }),
      competence = capitalizeMonth(new Date(`${period}-01T12:00:00`).toLocaleDateString("pt-BR", { month: "long", year: "numeric" })),
      border = { top: { style: "thin" as const, color: { argb: "FF94A3B8" } }, left: { style: "thin" as const, color: { argb: "FF94A3B8" } }, bottom: { style: "thin" as const, color: { argb: "FF94A3B8" } }, right: { style: "thin" as const, color: { argb: "FF94A3B8" } } };
    workbook.creator = "Sacolão ABC";
    sheet.columns = [{ width: 42 }, { width: 34 }, { width: 20 }];
    sheet.mergeCells("A1:C1");
    const title = sheet.getCell("A1");
    title.value = `CONTROLE DE IMPOSTOS — ${competence.toUpperCase()}`;
    title.font = { name: "Arial", size: 15, bold: true, color: { argb: "FFFFFFFF" } };
    title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
    title.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(1).height = 28;
    sheet.mergeCells("A2:C2");
    sheet.getCell("A2").value = `FILTRO: ${exportStatus.toUpperCase()}`;
    sheet.getCell("A2").font = { name: "Arial", size: 10, bold: true, color: { argb: "FF475569" } };
    sheet.getCell("A2").alignment = { horizontal: "center" };
    sheet.addRow([]);
    const palette = ["FFE31B23", "FF92D050", "FF0E83C7", "FFF59E0B", "FF7C3AED"];
    exportGroups.forEach((group, groupIndex) => {
      const titleRowNumber = sheet.rowCount + 1;
      sheet.addRow([group.company.toUpperCase()]);
      sheet.mergeCells(titleRowNumber, 1, titleRowNumber, 3);
      const companyTitle = sheet.getCell(titleRowNumber, 1);
      companyTitle.font = { name: "Arial", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
      companyTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: palette[groupIndex % palette.length] } };
      companyTitle.alignment = { horizontal: "center", vertical: "middle" };
      companyTitle.border = border;
      sheet.getRow(titleRowNumber).height = 24;
      const header = sheet.addRow(["NOME DO IMPOSTO", "EMPRESA", "VALOR"]);
      header.height = 23;
      header.eachCell((cell) => { cell.font = { name: "Arial", bold: true }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } }; cell.border = border; cell.alignment = { horizontal: "center", vertical: "middle" }; });
      group.rows.sort((a, b) => a.description.localeCompare(b.description, "pt-BR")).forEach((entry) => {
        const row = sheet.addRow([entry.description, entry.category, entry.amount]);
        row.height = 21;
        row.eachCell((cell, column) => { cell.border = border; cell.font = { name: "Arial", size: 10, bold: column === 1 || column === 3 }; cell.alignment = { horizontal: column === 3 ? "right" : "left", vertical: "middle" }; });
        row.getCell(3).numFmt = '"R$" #,##0.00';
      });
      const companyTotal = group.rows.reduce((sum, entry) => sum + entry.amount, 0), subtotal = sheet.addRow(["", "TOTAL DA EMPRESA", companyTotal]);
      subtotal.height = 23;
      subtotal.eachCell((cell) => { cell.font = { name: "Arial", bold: true }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD1D5DB" } }; cell.border = border; });
      subtotal.getCell(3).numFmt = '"R$" #,##0.00';
      sheet.addRow([]);
    });
    const totalRow = sheet.addRow(["", "TOTAL GERAL", exportTotal]);
    totalRow.height = 24;
    totalRow.eachCell((cell) => { cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } }; cell.border = border; });
    totalRow.getCell(3).numFmt = '"R$" #,##0.00';
    sheet.views = [{ state: "frozen", ySplit: 2, showGridLines: false }];
    const bytes = await workbook.xlsx.writeBuffer(), url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })), link = document.createElement("a");
    link.href = url; link.download = `impostos-${period}.xlsx`; link.click(); URL.revokeObjectURL(url);
  };
  return (
    <main className="fade-in p-4 sm:p-7">
      <div className="mb-7 flex flex-wrap items-end justify-between gap-4"><SectionHead title="Impostos" sub="Controle mensal de guias, vencimentos e pagamentos" /><label className="text-xs font-semibold text-slate-500">Mês de referência<input type="month" value={period} onChange={(event) => setPeriod(event.target.value)} className="mt-1 block rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold" /></label></div>
      <div className="mb-5 grid gap-4 sm:grid-cols-3">{[["Total do mês", total, "text-slate-900", <ReceiptText size={21} />], ["Total pago", paidTotal, "text-emerald-700", <BadgeCheck size={21} />], ["Pendente", pendingTotal, "text-amber-700", <Clock3 size={21} />]].map(([label, value, color, icon]) => <div key={String(label)} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft"><div className="flex items-start justify-between"><div className="grid h-11 w-11 place-items-center rounded-xl bg-slate-50 text-slate-700">{icon as ReactNode}</div><ChevronRight size={20} className="text-slate-400" /></div><div className={`mt-5 text-3xl font-black ${color}`}>{money(Number(value))}</div><div className="mt-1 text-sm font-bold text-slate-700">{label}</div></div>)}</div>
      {!!companyGroups.length && <div className="mb-5"><h3 className="mb-3 text-sm font-black text-slate-700">Resumo por empresa</h3><div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">{companyGroups.map((group) => { const companyTotal = group.allRows.reduce((sum, entry) => sum + entry.amount, 0), companyPaid = group.allRows.filter((entry) => entry.paid).reduce((sum, entry) => sum + entry.amount, 0); return <div key={group.company} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft"><div className="flex items-center justify-between"><div className="grid h-11 w-11 place-items-center rounded-xl bg-slate-50 text-slate-700"><Store size={20} /></div><ChevronRight size={20} className="text-slate-400" /></div><h4 className="mt-4 text-lg font-black text-slate-900">{group.company}</h4><div className="mt-4 grid grid-cols-3 divide-x divide-slate-100 text-center"><div><div className="text-[10px] font-bold uppercase text-slate-400">Total</div><div className="mt-1 font-black">{money(companyTotal)}</div></div><div><div className="text-[10px] font-bold uppercase text-slate-400">Pago</div><div className="mt-1 font-black text-emerald-700">{money(companyPaid)}</div></div><div><div className="text-[10px] font-bold uppercase text-slate-400">Pendente</div><div className="mt-1 font-black text-amber-700">{money(companyTotal - companyPaid)}</div></div></div></div>; })}</div></div>}
      <form onSubmit={submit} className="mb-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <div className="flex items-center justify-between"><div><h3 className="font-bold">{editingId ? "Editar guia" : "Cadastrar nova guia"}</h3><p className="text-xs text-slate-400">Os valores cadastrados entram automaticamente no total do mês.</p></div>{editingId && <button type="button" onClick={reset} className="text-sm font-bold text-slate-500">Cancelar edição</button>}</div>
        <div className="mt-4 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          <label className="text-sm font-semibold text-slate-600">Guia / imposto<input required value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} placeholder="Ex.: DARF, FGTS, ICMS" className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3" /></label>
          <label className="text-sm font-semibold text-slate-600">Loja / empresa<select required value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className="mt-2 w-full rounded-xl border border-slate-200 bg-white px-4 py-3"><option value="">Selecione a loja</option>{[...new Set([...companies, ...entries.map((entry) => entry.category).filter(Boolean)])].sort((a, b) => a.localeCompare(b, "pt-BR")).map((company) => <option key={company} value={company}>{company}</option>)}</select>{!companies.length && <span className="mt-1 block text-xs font-normal text-amber-600">Cadastre as lojas em Configurações.</span>}</label>
          <label className="text-sm font-semibold text-slate-600">Vencimento<input required type="date" value={form.dueDate} onChange={(e) => setForm({ ...form, dueDate: e.target.value })} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3" /></label>
          <label className="text-sm font-semibold text-slate-600">Valor<input required inputMode="decimal" value={form.amount} onChange={(e) => setForm({ ...form, amount: formatMoneyInput(e.target.value) })} placeholder="0,00" className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 font-bold" /></label>
          <label className="text-sm font-semibold text-slate-600 md:col-span-2">Observação<input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="Informação opcional" className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3" /></label>
          <label className="flex items-center gap-3 self-end rounded-xl border border-slate-200 px-4 py-3 text-sm font-bold"><input type="checkbox" checked={form.paid} onChange={(e) => setForm({ ...form, paid: e.target.checked, paidAt: e.target.checked ? (form.paidAt || new Date().toISOString().slice(0, 10)) : "" })} />Guia já paga</label>
          {form.paid && <label className="text-sm font-semibold text-slate-600">Data do pagamento<input required type="date" value={form.paidAt} onChange={(e) => setForm({ ...form, paidAt: e.target.value })} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3" /></label>}
        </div>
        <button className="mt-4 rounded-xl bg-slate-800 px-5 py-3 text-sm font-bold text-white">{editingId ? "Salvar alterações" : "Cadastrar guia"}</button>
      </form>
      <div className="mb-5 flex flex-wrap gap-3"><div className="relative min-w-[240px] flex-1"><Search className="absolute left-4 top-3 text-slate-400" size={19} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Pesquisar guia..." className="w-full rounded-xl border border-slate-200 bg-white py-2.5 pl-11 pr-4" /></div><label className="flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 text-xs font-bold text-slate-500">Exportar<select value={exportStatus} onChange={(event) => setExportStatus(event.target.value as "Todos" | "Pagos" | "Não pagos")} className="bg-transparent py-2.5 text-sm font-bold text-slate-700 outline-none"><option>Todos</option><option>Pagos</option><option>Não pagos</option></select></label><button onClick={exportExcel} disabled={!periodEntries.some((entry) => exportStatus === "Todos" || (exportStatus === "Pagos" ? entry.paid : !entry.paid))} className="flex items-center gap-2 rounded-xl bg-forest-700 px-5 py-2.5 text-sm font-bold text-white disabled:opacity-40"><FileSpreadsheet size={18} />Gerar planilha Excel</button></div>
      <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white shadow-soft"><table className="w-full min-w-[980px] text-sm"><thead className="bg-slate-100 text-xs uppercase text-slate-500"><tr><th className="px-4 py-3 text-left">Guia / imposto</th><th className="px-4 py-3 text-left">Empresa</th><th className="px-4 py-3 text-left">Vencimento</th><th className="px-4 py-3 text-right">Valor</th><th className="px-4 py-3 text-left">Situação</th><th className="px-4 py-3 text-right">Ações</th></tr></thead><tbody className="divide-y divide-slate-100">{monthEntries.map((entry) => <tr key={entry.id}><td className="px-4 py-3 font-bold">{entry.description}<div className="text-xs font-normal text-slate-400">{entry.note}</div></td><td className="px-4 py-3">{entry.category}</td><td className="px-4 py-3">{formatDateBr(entry.dueDate)}</td><td className="px-4 py-3 text-right font-bold">{money(entry.amount)}</td><td className="px-4 py-3"><span className={`rounded-full px-3 py-1 text-xs font-bold ${entry.paid ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"}`}>{entry.paid ? "Pago" : "Pendente"}</span></td><td className="whitespace-nowrap px-4 py-3 text-right"><button onClick={() => setEntries(entries.map((item) => item.id === entry.id ? { ...item, paid: !item.paid, paidAt: !item.paid ? new Date().toISOString().slice(0, 10) : undefined } : item))} className={`mr-2 rounded-lg px-3 py-2 text-xs font-bold ${entry.paid ? "border border-amber-300 text-amber-700" : "bg-emerald-600 text-white"}`}>{entry.paid ? "Desfazer pagamento" : "Marcar como pago"}</button><button onClick={() => startEdit(entry)} className="mr-2 rounded-lg border border-slate-200 p-2 text-slate-600" aria-label="Editar guia"><Pencil size={16} /></button><button onClick={() => { if (confirm(`Excluir a guia ${entry.description}?`)) setEntries(entries.filter((item) => item.id !== entry.id)); }} className="rounded-lg border border-red-200 p-2 text-red-600" aria-label="Excluir guia"><X size={16} /></button></td></tr>)}{!monthEntries.length && <tr><td colSpan={6} className="py-12 text-center text-slate-400">Nenhuma guia cadastrada neste mês.</td></tr>}</tbody><tfoot className="bg-slate-900 font-black text-white"><tr><td colSpan={3} className="px-4 py-4 text-right">TOTAL DO MÊS</td><td className="px-4 py-4 text-right">{money(total)}</td><td colSpan={2}></td></tr></tfoot></table></div>
    </main>
  );
}

function BonusPage({
  employees,
  occurrences,
  period,
  setPeriod,
}: {
  employees: Recharge[];
  occurrences: HROccurrence[];
  period: string;
  setPeriod: (period: string) => void;
}) {
  const [query, setQuery] = useState(""),
    [statuses, setStatuses] = useState<Array<"Elegível" | "Não elegível">>([
      "Elegível",
      "Não elegível",
    ]),
    [year, month] = period.split("-").map(Number),
    cycleStart = new Date(year, month - 2, 14, 12),
    cycleEnd = new Date(year, month - 1, 14, 12),
    paymentDate = new Date(year, month - 1, 15, 12),
    money = (value: number) =>
      value.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }),
    format = (date: Date) => date.toLocaleDateString("pt-BR"),
    eligibleEmployees = employees.filter((employee) => {
      if (
        !employee.terminationDate &&
        (employee.active === false || employee.employmentStatus === "Desligado")
      )
        return false;
      const hired = employee.hiredAt
          ? new Date(employee.hiredAt + "T12:00:00")
          : null,
        terminated = employee.terminationDate
          ? new Date(employee.terminationDate + "T12:00:00")
          : null;
      return (!hired || hired <= cycleEnd) && (!terminated || terminated >= cycleStart);
    }),
    rows = eligibleEmployees.map((employee) => {
      const hired = employee.hiredAt
          ? new Date(employee.hiredAt + "T12:00:00")
          : null,
        terminated = employee.terminationDate
          ? new Date(employee.terminationDate + "T12:00:00")
          : null,
        employmentReasons: string[] = [];
      if (!hired)
        employmentReasons.push("Data de admissão não informada");
      else if (hired > cycleStart)
        employmentReasons.push(
          `Admitido em ${hired.toLocaleDateString("pt-BR")}, após o início do ciclo`,
        );
      if (terminated && terminated < cycleEnd)
        employmentReasons.push(
          `Desligado em ${terminated.toLocaleDateString("pt-BR")}, antes do fim do ciclo`,
        );
      const issues = occurrences.filter((item) => {
        if (
          item.employeeId !== employee.id ||
          (item.type !== "Falta" &&
            item.type !== "Atestado" &&
            item.type !== "Férias")
        )
          return false;
        const occurrenceStart = new Date(item.date + "T12:00:00"),
          occurrenceEnd = new Date(occurrenceStart);
        if (item.type === "Férias" && item.endDate) {
          occurrenceEnd.setTime(
            new Date(item.endDate + "T12:00:00").getTime(),
          );
        } else {
          occurrenceEnd.setDate(
            occurrenceEnd.getDate() +
              (item.type === "Atestado"
                ? Math.max(1, item.days || 1) - 1
                : 0),
          );
        }
        return occurrenceStart <= cycleEnd && occurrenceEnd >= cycleStart;
      }),
        occurrenceReasons = issues.map((item) =>
          item.type === "Férias"
            ? `Férias de ${new Date(item.date + "T12:00:00").toLocaleDateString("pt-BR")} até ${item.endDate ? new Date(item.endDate + "T12:00:00").toLocaleDateString("pt-BR") : "data não informada"}`
            : `${item.type} em ${new Date(item.date + "T12:00:00").toLocaleDateString("pt-BR")}`,
        ),
        reasons = [...employmentReasons, ...occurrenceReasons],
        eligible = reasons.length === 0;
      return {
        employee,
        issues,
        status: (eligible ? "Elegível" : "Não elegível") as "Elegível" | "Não elegível",
        amount: eligible ? 80 : 0,
        reason: eligible
          ? "Trabalhou o ciclo completo, sem faltas ou atestados"
          : reasons.join("; "),
      };
    }),
    visibleRows = rows.filter(
      (row) =>
        statuses.includes(row.status) &&
        `${row.employee.employee} ${row.employee.store} ${row.employee.role}`
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .toLowerCase()
          .includes(query.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()),
    ),
    stores = [...new Set(visibleRows.map((row) => row.employee.store))].sort(),
    grouped = stores.map((store) => ({
      store,
      rows: visibleRows.filter((row) => row.employee.store === store),
    })),
    eligibleCount = rows.filter((row) => row.status === "Elegível").length,
    ineligibleCount = rows.length - eligibleCount,
    total = eligibleCount * 80;

  const exportPdf = () => {
    const pdf = new jsPDF({ orientation: "landscape" });
    pdf.setFillColor(30, 41, 59);
    pdf.rect(0, 0, 297, 32, "F");
    pdf.setTextColor(255, 255, 255);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(17);
    pdf.text("RELATÓRIO DE BONIFICAÇÃO", 14, 14);
    pdf.setFontSize(9);
    pdf.text(`Ciclo: ${format(cycleStart)} a ${format(cycleEnd)} | Pagamento: ${format(paymentDate)} | Valor individual: ${money(80)}`, 14, 23);
    let startY = 39;
    grouped.forEach((group) => {
      if (startY > 174) { pdf.addPage(); startY = 18; }
      pdf.setFillColor(30, 41, 59);
      pdf.roundedRect(14, startY, 269, 10, 2, 2, "F");
      pdf.setTextColor(255, 255, 255);
      pdf.setFontSize(11);
      pdf.text(group.store.toUpperCase(), 148.5, startY + 6.8, { align: "center" });
      autoTable(pdf, {
        startY: startY + 10,
        head: [["NOME", "FUNÇÃO", "STATUS", "MOTIVO", "VALOR"]],
        body: [
          ...group.rows.map((row) => [row.employee.employee, row.employee.role, row.status, row.reason, money(row.amount)]),
          ["", "", "", "TOTAL DA LOJA", money(group.rows.reduce((sum, row) => sum + row.amount, 0))],
        ],
        theme: "grid",
        styles: { fontSize: 8, textColor: [15, 23, 42], lineColor: [218, 223, 230], lineWidth: 0.1 },
        headStyles: { fillColor: [229, 231, 235], textColor: [15, 23, 42], fontStyle: "bold" },
        columnStyles: { 0: { cellWidth: 65, fontStyle: "bold" }, 1: { cellWidth: 46 }, 2: { cellWidth: 30 }, 3: { cellWidth: 90 }, 4: { cellWidth: 38, halign: "right", fontStyle: "bold" } },
        didParseCell: (data) => {
          if (data.section === "body" && data.row.index === group.rows.length) {
            data.cell.styles.fillColor = [229, 231, 235];
            data.cell.styles.fontStyle = "bold";
          }
        },
        margin: { left: 14, right: 14 },
      });
      startY = ((pdf as any).lastAutoTable?.finalY || startY) + 6;
    });
    if (startY > 186) { pdf.addPage(); startY = 18; }
    pdf.setFillColor(22, 101, 52);
    pdf.roundedRect(178, startY, 105, 12, 2, 2, "F");
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(11);
    pdf.text(`TOTAL GERAL: ${money(visibleRows.reduce((sum, row) => sum + row.amount, 0))}`, 279, startY + 7.7, { align: "right" });
    pdf.save(`bonificacao-${period}.pdf`);
  };

  const exportExcel = async () => {
    const workbook = new ExcelJS.Workbook(),
      sheet = workbook.addWorksheet("Bonificação", { views: [{ showGridLines: false }] }),
      palette = ["FFE31B23", "FF92D050", "FF0E83C7", "FFF59E0B", "FF7C3AED"];
    sheet.columns = [{ width: 42 }, { width: 30 }, { width: 18 }, { width: 48 }, { width: 16 }];
    sheet.mergeCells("A1:E1");
    const title = sheet.getCell("A1");
    title.value = `RELATÓRIO DE BONIFICAÇÃO — ${capitalizeMonth(new Date(`${period}-01T12:00:00`).toLocaleDateString("pt-BR", { month: "long", year: "numeric" })).toUpperCase()}`;
    title.font = { name: "Arial", size: 15, bold: true, color: { argb: "FFFFFFFF" } };
    title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF111111" } };
    title.alignment = { horizontal: "center" };
    sheet.mergeCells("A2:E2");
    sheet.getCell("A2").value = `Ciclo: ${format(cycleStart)} a ${format(cycleEnd)} | Pagamento: ${format(paymentDate)} | Valor individual: R$ 80,00`;
    sheet.getCell("A2").alignment = { horizontal: "center" };
    sheet.addRow([]);
    grouped.forEach((group, index) => {
      const storeRow = sheet.addRow([group.store.toUpperCase()]);
      sheet.mergeCells(storeRow.number, 1, storeRow.number, 5);
      storeRow.getCell(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
      storeRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: palette[index % palette.length] } };
      storeRow.getCell(1).alignment = { horizontal: "center" };
      const header = sheet.addRow(["NOME", "FUNÇÃO", "STATUS", "MOTIVO", "VALOR"]);
      header.eachCell((cell) => { cell.font = { bold: true }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE5E7EB" } }; });
      group.rows.forEach((row) => {
        const excelRow = sheet.addRow([row.employee.employee.toUpperCase(), row.employee.role.toUpperCase(), row.status, row.reason, row.amount]);
        excelRow.getCell(5).numFmt = '"R$" #,##0.00';
        excelRow.getCell(1).font = { bold: true };
        if (row.status === "Não elegível") excelRow.getCell(3).font = { bold: true, color: { argb: "FFDC2626" } };
      });
      const subtotal = sheet.addRow(["", "", "", "TOTAL DA LOJA", group.rows.reduce((sum, row) => sum + row.amount, 0)]);
      subtotal.eachCell((cell) => { cell.font = { bold: true }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD1D5DB" } }; });
      subtotal.getCell(5).numFmt = '"R$" #,##0.00';
      sheet.addRow([]);
    });
    const grandTotal = sheet.addRow(["", "", "", "TOTAL GERAL", visibleRows.reduce((sum, row) => sum + row.amount, 0)]);
    grandTotal.eachCell((cell) => { cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF166534" } }; });
    grandTotal.getCell(5).numFmt = '"R$" #,##0.00';
    const bytes = await workbook.xlsx.writeBuffer(),
      url = URL.createObjectURL(new Blob([bytes])),
      link = document.createElement("a");
    link.href = url;
    link.download = `bonificacao-${period}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);
  };

  return (
    <main className="fade-in p-4 sm:p-7">
      <div className="mb-7 flex flex-wrap items-end justify-between gap-4">
        <SectionHead title="Bonificação" sub="Controle mensal de assiduidade e bonificação de R$ 80,00" />
        <label className="text-xs font-semibold text-slate-500">Mês de pagamento<input type="month" value={period} onChange={(event) => setPeriod(event.target.value)} className="mt-1 block rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold" /></label>
      </div>
      <div className="mb-5 rounded-2xl border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-900">
        <strong>Ciclo avaliado:</strong> {format(cycleStart)} até {format(cycleEnd)} · <strong>Pagamento:</strong> {format(paymentDate)}
        <div className="mt-1 text-xs text-emerald-800">
          Para ser elegível, o funcionário precisa ter trabalhado durante todo o ciclo e não possuir falta nem atestado no período.
        </div>
      </div>
      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
        {[
          ["Funcionários avaliados", rows.length, "bg-slate-50 border-slate-300 text-slate-800"],
          ["Elegíveis", eligibleCount, "bg-emerald-50 border-emerald-300 text-emerald-700"],
          ["Não elegíveis", ineligibleCount, "bg-red-50 border-red-300 text-red-700"],
          ["Total da bonificação", money(total), "bg-blue-50 border-blue-300 text-blue-700"],
        ].map(([label, value, style]) => <div key={String(label)} className={`rounded-2xl border border-t-4 p-5 shadow-soft ${style}`}><Gift size={21}/><div className="mt-4 text-2xl font-black">{value}</div><div className="mt-1 text-sm font-bold">{label}</div></div>)}
      </div>
      <div className="mt-5 grid gap-3 rounded-2xl border border-slate-200 bg-white p-4 shadow-soft md:grid-cols-[1fr_260px_auto_auto]">
        <div className="relative"><Search className="absolute left-3 top-3 text-slate-400" size={18}/><input value={query} onChange={(event)=>setQuery(event.target.value)} placeholder="Pesquisar funcionário, loja ou função..." className="w-full rounded-xl border border-slate-200 py-2.5 pl-10 pr-3"/></div>
        <details className="relative"><summary className="cursor-pointer list-none rounded-xl border border-slate-200 px-4 py-2.5 text-sm font-bold">Status: {statuses.length===2?"Todos":statuses[0]||"Nenhum"}</summary><div className="absolute z-20 mt-2 w-full rounded-xl border bg-white p-3 shadow-xl">{(["Elegível","Não elegível"] as const).map(status=><label key={status} className="flex gap-2 py-2 text-sm"><input type="checkbox" checked={statuses.includes(status)} onChange={()=>setStatuses(current=>current.includes(status)?current.filter(item=>item!==status):[...current,status])}/>{status}</label>)}</div></details>
        <button onClick={exportPdf} className="rounded-xl border border-red-300 px-4 py-2.5 text-sm font-bold text-red-600">Gerar PDF</button>
        <button onClick={exportExcel} className="rounded-xl bg-slate-800 px-4 py-2.5 text-sm font-bold text-white">Gerar Excel</button>
      </div>
      <div className="mt-5 space-y-5">
        {grouped.map((group)=><div key={group.store} className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft"><div className="bg-slate-800 px-5 py-3 text-center font-black uppercase text-white">{group.store}</div><div className="overflow-x-auto"><table className="w-full min-w-[900px] text-sm"><thead className="bg-slate-100 text-xs uppercase"><tr><th className="px-4 py-3 text-left">Funcionário</th><th className="px-4 py-3 text-left">Função</th><th className="px-4 py-3 text-left">Status</th><th className="px-4 py-3 text-left">Motivo</th><th className="px-4 py-3 text-right">Valor</th></tr></thead><tbody className="divide-y">{group.rows.map(row=><tr key={row.employee.id}><td className="px-4 py-3 font-bold">{row.employee.employee}</td><td className="px-4 py-3">{row.employee.role}</td><td className="px-4 py-3"><span className={`rounded-full px-2.5 py-1 text-xs font-bold ${row.status==="Elegível"?"bg-emerald-100 text-emerald-700":"bg-red-100 text-red-700"}`}>{row.status}</span></td><td className="px-4 py-3 text-slate-600">{row.reason}</td><td className="px-4 py-3 text-right font-bold">{money(row.amount)}</td></tr>)}</tbody><tfoot className="bg-slate-200 font-black"><tr><td colSpan={4} className="px-4 py-3 text-right">TOTAL DA LOJA</td><td className="px-4 py-3 text-right">{money(group.rows.reduce((sum,row)=>sum+row.amount,0))}</td></tr></tfoot></table></div></div>)}
        {!visibleRows.length&&<div className="rounded-2xl border bg-white py-12 text-center text-slate-400">Nenhum funcionário para os filtros selecionados.</div>}
      </div>
    </main>
  );
}

function FinancialReports({
  employees,
  entries,
  period,
  setPeriod,
}: {
  employees: Recharge[];
  entries: FinancialEntry[];
  period: string;
  setPeriod: (period: string) => void;
}) {
  type ReportKind = "salary" | "advance" | "vacation" | "severance";
  const reportKinds: { key: ReportKind; label: string }[] = [
      { key: "salary", label: "Pagamento de salário" },
      { key: "advance", label: "Pagamento de adiantamento" },
      { key: "vacation", label: "Férias" },
      { key: "severance", label: "Verbas rescisórias" },
    ],
    allStores = [...new Set(employees.map((employee) => employee.store))].sort(),
    [selectedStores, setSelectedStores] = useState<string[]>(allStores),
    [selectedKinds, setSelectedKinds] = useState<ReportKind[]>([
      "salary",
      "advance",
      "vacation",
      "severance",
    ]);
  useEffect(() => {
    setSelectedStores((current) => {
      const valid = current.filter((store) => allStores.includes(store));
      return valid.length ? valid : allStores;
    });
  }, [employees.map((employee) => employee.store).join("|")]);
  const byEmployee = new Map(
      entries
        .filter((entry) => entry.period === period)
        .map((entry) => [entry.employeeId, entry]),
    ),
    rows = employees
      .filter((employee) => selectedStores.includes(employee.store))
      .map((employee) => ({
        employee,
        entry: byEmployee.get(employee.id),
      })),
    currency = (value: number) =>
      (value || 0).toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL",
      }),
    todayIso = new Date().toISOString().slice(0, 10),
    paymentReached = (date?: string) => !!date && date <= todayIso,
    totalSalary = rows.reduce((sum, row) => sum + (row.entry?.salary || 0), 0),
    totalAdvance = rows.reduce((sum, row) => sum + (row.entry?.advance || 0), 0),
    totalVacation = rows.reduce((sum, row) => sum + (row.entry?.vacation || 0), 0),
    totalSeverance = rows.reduce((sum, row) => sum + (row.entry?.severance || 0), 0),
    totalPaid = rows.reduce(
      (sum, row) =>
        sum +
        (paymentReached(row.entry?.salaryPaidAt) ? row.entry?.salary || 0 : 0) +
        (paymentReached(row.entry?.advancePaidAt) ? row.entry?.advance || 0 : 0) +
        (paymentReached(row.entry?.vacationPaidAt) ? row.entry?.vacation || 0 : 0) +
        (paymentReached(row.entry?.severancePaidAt) ? row.entry?.severance || 0 : 0),
      0,
    ),
    selectedValue = (entry?: FinancialEntry) =>
      selectedKinds.reduce((sum, kind) => sum + (entry?.[kind] || 0), 0),
    groupedRows = allStores
      .filter((store) => selectedStores.includes(store))
      .map((store) => ({
        store,
        rows: rows.filter(
          ({ employee, entry }) =>
            employee.store === store && selectedValue(entry) > 0,
        ),
      })),
    filteredTotal = rows.reduce(
      (sum, { entry }) => sum + selectedValue(entry),
      0,
    );
  const exportExcel = async () => {
    const workbook = new ExcelJS.Workbook(),
      sheet = workbook.addWorksheet("Financeiro", {
        views: [{ showGridLines: false }],
        pageSetup: {
          orientation: "landscape",
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          margins: {
            left: 0.25,
            right: 0.25,
            top: 0.4,
            bottom: 0.4,
            header: 0.15,
            footer: 0.15,
          },
        },
      }),
      palette = ["FFE31B23", "FF92D050", "FF0E83C7", "FFF59E0B", "FF7C3AED"],
      border = {
        top: { style: "thin" as const, color: { argb: "FF555555" } },
        left: { style: "thin" as const, color: { argb: "FF555555" } },
        bottom: { style: "thin" as const, color: { argb: "FF555555" } },
        right: { style: "thin" as const, color: { argb: "FF555555" } },
      },
      competence = capitalizeMonth(
        new Date(`${period}-01T12:00:00`).toLocaleDateString("pt-BR", {
          month: "long",
          year: "numeric",
        }),
      );
    workbook.creator = "Sacolão ABC";
    sheet.columns = [{ width: 48 }, { width: 36 }, { width: 20 }];
    sheet.mergeCells("A1:C1");
    const mainTitle = sheet.getCell("A1");
    mainTitle.value = `RELATÓRIO FINANCEIRO — COMPETÊNCIA ${competence.toUpperCase()}`;
    mainTitle.font = { name: "Arial", size: 15, bold: true, color: { argb: "FFFFFFFF" } };
    mainTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF111111" } };
    mainTitle.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(1).height = 28;
    sheet.addRow([]);
    groupedRows.forEach((group, groupIndex) => {
      const color = palette[groupIndex % palette.length],
        titleRowNumber = sheet.rowCount + 1;
      sheet.addRow([group.store.toUpperCase()]);
      sheet.mergeCells(titleRowNumber, 1, titleRowNumber, 3);
      const storeTitle = sheet.getCell(titleRowNumber, 1);
      storeTitle.font = { name: "Arial", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
      storeTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
      storeTitle.alignment = { horizontal: "center", vertical: "middle" };
      storeTitle.border = border;
      sheet.getRow(titleRowNumber).height = 24;
      const header = sheet.addRow(["NOME", "FUNÇÃO", "VALOR"]);
      header.height = 22;
      header.eachCell((cell, column) => {
        cell.font = { name: "Arial", bold: true, color: { argb: "FF111827" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE5E7EB" } };
        cell.border = border;
        cell.alignment = { horizontal: column === 3 ? "right" : "center", vertical: "middle" };
      });
      group.rows.forEach(({ employee, entry }) => {
        const row = sheet.addRow([
          employee.employee.toUpperCase(),
          employee.role.toUpperCase(),
          selectedValue(entry),
        ]);
        row.height = 21;
        row.eachCell((cell, column) => {
          cell.font = { name: "Arial", size: 10, bold: column === 1 || column === 3 };
          cell.border = border;
          cell.alignment = { horizontal: column === 3 ? "right" : "left", vertical: "middle" };
        });
        row.getCell(3).numFmt = '"R$" #,##0.00';
      });
      const subtotal = group.rows.reduce((sum, { entry }) => sum + selectedValue(entry), 0),
        totalRow = sheet.addRow(["", "TOTAL", subtotal]);
      totalRow.height = 22;
      totalRow.eachCell((cell, column) => {
        cell.font = { name: "Arial", bold: true, color: { argb: "FF111827" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD1D5DB" } };
        cell.border = border;
        cell.alignment = { horizontal: column === 3 ? "right" : "center", vertical: "middle" };
      });
      totalRow.getCell(3).numFmt = '"R$" #,##0.00';
      sheet.addRow([]);
    });
    const grandTotal = sheet.addRow(["", "TOTAL GERAL", filteredTotal]);
    grandTotal.height = 25;
    grandTotal.eachCell((cell, column) => {
      cell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      cell.border = border;
      cell.alignment = { horizontal: column === 3 ? "right" : "center", vertical: "middle" };
    });
    grandTotal.getCell(3).numFmt = '"R$" #,##0.00';
    sheet.pageSetup.printArea = `A1:C${sheet.rowCount}`;
    const bytes = await workbook.xlsx.writeBuffer(),
      url = URL.createObjectURL(
        new Blob([bytes], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
      ),
      link = document.createElement("a");
    link.href = url;
    link.download = `financeiro-${period}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);
  };
  const exportPdf = () => {
    const pdf = new jsPDF({ orientation: "landscape" });
    pdf.setFontSize(18);
    pdf.text("Relatório financeiro", 14, 17);
    pdf.setFontSize(9);
    pdf.text(`Mês: ${period.split("-").reverse().join("/")} | Gerado em ${new Date().toLocaleString("pt-BR")}`, 14, 24);
    pdf.text(`Tipos: ${reportKinds.filter((kind) => selectedKinds.includes(kind.key)).map((kind) => kind.label).join(", ")}`, 14, 30);
    let startY = 36;
    groupedRows.forEach((group) => {
      const groupTotal = group.rows.reduce(
          (sum, { entry }) => sum + selectedValue(entry),
          0,
        ),
        body = [
          ...group.rows.map(({ employee, entry }) => [
            employee.employee.toUpperCase(),
            employee.role.toUpperCase(),
            currency(selectedValue(entry)),
          ]),
          ["", "TOTAL", currency(groupTotal)],
        ];
      if (startY > 175) {
        pdf.addPage();
        startY = 18;
      }
      pdf.setFillColor(30, 41, 59);
      pdf.roundedRect(14, startY, 269, 11, 2, 2, "F");
      pdf.setTextColor(255, 255, 255);
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(12);
      pdf.text(group.store.toUpperCase(), 148.5, startY + 7.2, {
        align: "center",
      });
      pdf.setTextColor(15, 23, 42);
      autoTable(pdf, {
        startY: startY + 11,
        head: [["NOME", "FUNÇÃO", "VALOR"]],
        body,
        theme: "grid",
        styles: {
          font: "helvetica",
          fontSize: 9,
          textColor: [15, 23, 42],
          fillColor: [255, 255, 255],
          lineColor: [218, 223, 230],
          lineWidth: 0.12,
          cellPadding: { top: 3, right: 4, bottom: 3, left: 4 },
        },
        headStyles: {
          fillColor: [229, 231, 235],
          textColor: [15, 23, 42],
          fontStyle: "bold",
          halign: "left",
        },
        columnStyles: {
          0: { cellWidth: 132, fontStyle: "bold" },
          1: { cellWidth: 92 },
          2: { cellWidth: 45, halign: "right", fontStyle: "bold" },
        },
        didParseCell: (data) => {
          if (data.section === "body" && data.row.index === body.length - 1) {
            data.cell.styles.fillColor = [229, 231, 235];
            data.cell.styles.fontStyle = "bold";
          }
        },
        margin: { left: 14, right: 14 },
      });
      startY = ((pdf as any).lastAutoTable?.finalY || startY) + 6;
    });
    if (startY > 188) {
      pdf.addPage();
      startY = 18;
    }
    pdf.setFillColor(30, 41, 59);
    pdf.roundedRect(185, startY, 98, 12, 2, 2, "F");
    pdf.setTextColor(255, 255, 255);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(11);
    pdf.text(`TOTAL GERAL: ${currency(filteredTotal)}`, 279, startY + 7.7, {
      align: "right",
    });
    pdf.save(`financeiro-${period}.pdf`);
  };
  return (
    <main className="fade-in p-4 sm:p-7">
      <div className="mb-7 flex flex-wrap items-end justify-between gap-4">
        <SectionHead title="Relatórios Financeiros" sub="Folha, adiantamentos e saídas de caixa por mês" />
        <label className="text-xs font-semibold text-slate-500">
          Mês de referência
          <input type="month" value={period} onChange={(event) => setPeriod(event.target.value)} className="mt-1 block rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold outline-none" />
        </label>
      </div>
      <div className="mb-5 grid gap-4 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft sm:grid-cols-2">
        <details className="relative">
          <summary className="cursor-pointer list-none rounded-xl border border-slate-200 px-4 py-3 text-sm font-bold text-slate-700">
            Lojas: {selectedStores.length === allStores.length ? "Todas" : `${selectedStores.length} selecionada(s)`}
          </summary>
          <div className="absolute z-20 mt-2 max-h-72 w-full overflow-auto rounded-xl border border-slate-200 bg-white p-3 shadow-xl">
            <label className="flex cursor-pointer items-center gap-3 border-b border-slate-100 px-2 py-2 font-bold"><input type="checkbox" checked={selectedStores.length === allStores.length} onChange={(event) => setSelectedStores(event.target.checked ? allStores : [])} /> Todas as lojas</label>
            {allStores.map((store) => <label key={store} className="flex cursor-pointer items-center gap-3 px-2 py-2 text-sm"><input type="checkbox" checked={selectedStores.includes(store)} onChange={() => setSelectedStores((current) => current.includes(store) ? current.filter((item) => item !== store) : [...current, store])} /> {store}</label>)}
          </div>
        </details>
        <details className="relative">
          <summary className="cursor-pointer list-none rounded-xl border border-slate-200 px-4 py-3 text-sm font-bold text-slate-700">
            Tipos: {selectedKinds.length === reportKinds.length ? "Todos" : `${selectedKinds.length} selecionado(s)`}
          </summary>
          <div className="absolute z-20 mt-2 w-full rounded-xl border border-slate-200 bg-white p-3 shadow-xl">
            <label className="flex cursor-pointer items-center gap-3 border-b border-slate-100 px-2 py-2 font-bold"><input type="checkbox" checked={selectedKinds.length === reportKinds.length} onChange={(event) => setSelectedKinds(event.target.checked ? reportKinds.map((kind) => kind.key) : [])} /> Todos os tipos</label>
            {reportKinds.map((kind) => <label key={kind.key} className="flex cursor-pointer items-center gap-3 px-2 py-2 text-sm"><input type="checkbox" checked={selectedKinds.includes(kind.key)} onChange={() => setSelectedKinds((current) => current.includes(kind.key) ? current.filter((item) => item !== kind.key) : [...current, kind.key])} /> {kind.label}</label>)}
          </div>
        </details>
      </div>
      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
        {[
          ["Pagamento de salário", totalSalary],
          ["Pagamento de adiantamento", totalAdvance],
          ["Total pago", totalPaid],
          ["Pendente", Math.max(0, totalSalary + totalAdvance + totalVacation + totalSeverance - totalPaid)],
        ].map(([label, value]) => (
          <div key={String(label)} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
            <div className="text-xs font-semibold uppercase text-slate-400">{label}</div>
            <div className="mt-2 text-2xl font-bold">{currency(Number(value))}</div>
          </div>
        ))}
      </div>
      <div className="mt-5 grid gap-4 sm:grid-cols-2">
        <button onClick={exportPdf} className="rounded-2xl border border-slate-200 bg-white p-6 text-left shadow-soft transition hover:border-slate-400">
          <Download className="text-red-500" />
          <h3 className="mt-4 font-bold">Gerar relatório em PDF</h3>
          <p className="mt-1 text-sm text-slate-500">Resumo pronto para impressão e envio ao financeiro.</p>
        </button>
        <button onClick={exportExcel} className="rounded-2xl border border-slate-200 bg-white p-6 text-left shadow-soft transition hover:border-slate-400">
          <FileSpreadsheet className="text-slate-700" />
          <h3 className="mt-4 font-bold">Gerar planilha Excel</h3>
          <p className="mt-1 text-sm text-slate-500">Dados mensais completos para conferência e análise.</p>
        </button>
      </div>
      <div className="mt-5 space-y-5">
        {groupedRows.map((group) => (
          <div key={group.store} className="overflow-hidden rounded-2xl border border-slate-300 bg-white shadow-soft">
            <div className="bg-slate-800 px-5 py-3 text-center text-lg font-black uppercase text-white">{group.store}</div>
            <div className="overflow-x-auto">
              <table className="w-full min-w-[680px] text-sm">
                <thead className="border-b border-slate-300 bg-slate-100 text-slate-900"><tr><th className="px-5 py-2 text-left">NOME</th><th className="px-5 py-2 text-left">FUNÇÃO</th><th className="px-5 py-2 text-right">VALOR</th></tr></thead>
                <tbody className="divide-y divide-slate-200">
                  {group.rows.map(({ employee, entry }) => (
                    <tr key={employee.id}><td className="px-5 py-2 font-semibold uppercase">{employee.employee}</td><td className="px-5 py-2 uppercase">{employee.role}</td><td className="px-5 py-2 text-right font-semibold">{currency(selectedValue(entry))}</td></tr>
                  ))}
                  {!group.rows.length && <tr><td colSpan={3} className="px-5 py-7 text-center text-slate-400">Nenhum valor cadastrado para os filtros escolhidos.</td></tr>}
                </tbody>
                <tfoot className="border-t-2 border-slate-400 bg-slate-200 font-black"><tr><td></td><td className="px-5 py-3 text-right">TOTAL</td><td className="px-5 py-3 text-right">{currency(group.rows.reduce((sum, { entry }) => sum + selectedValue(entry), 0))}</td></tr></tfoot>
              </table>
            </div>
          </div>
        ))}
        <div className="flex justify-end rounded-2xl bg-slate-900 px-6 py-4 text-lg font-black text-white"><span className="mr-8">TOTAL GERAL</span>{currency(filteredTotal)}</div>
      </div>
    </main>
  );
}

function HRPage({
  rows,
  occurrences,
  onEdit,
}: {
  rows: Recharge[];
  occurrences: HROccurrence[];
  onEdit: (r: Recharge) => void;
}) {
  const [detail, setDetail] = useState<{
    title: string;
    rows: Recharge[];
  } | null>(null);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const currentMonthLabel = capitalizeMonth(
    today.toLocaleDateString("pt-BR", { month: "long", year: "numeric" }),
  );
  const active = rows.filter((r) => !isEmployeeDismissed(r)),
    stores = [...new Set(active.map((r) => r.store))],
    addDays = (date: string, days: number) => {
      const d = new Date(date + "T12:00:00");
      d.setDate(d.getDate() + days);
      return d;
    },
    daysUntil = (d: Date) =>
      Math.ceil((d.getTime() - today.getTime()) / 86400000);
  const experiences = active
    .filter((r) => r.hiredAt && r.formalEmployment !== false)
    .map((r) => {
      const first = addDays(r.hiredAt!, 29),
        end = addDays(r.hiredAt!, 89),
        firstInMonth =
          first.getFullYear() === today.getFullYear() &&
          first.getMonth() === today.getMonth(),
        endInMonth =
          end.getFullYear() === today.getFullYear() &&
          end.getMonth() === today.getMonth(),
        firstUpcoming = firstInMonth && first >= today,
        endUpcoming = endInMonth && end >= today,
        deadline = firstUpcoming ? first : end,
        stage = firstUpcoming
          ? "Fim dos primeiros 30 dias"
          : "Fim da prorrogação de 60 dias",
        endsThisMonth = firstUpcoming || endUpcoming;
      return { r, first, end, deadline, stage, endsThisMonth };
    })
    .filter((x) => x.endsThisMonth);
  const birthdays = active
    .filter((r) => r.birthDate)
    .map((r) => {
      const original = new Date(r.birthDate + "T12:00:00"),
        next = new Date(
          today.getFullYear(),
          original.getMonth(),
          original.getDate(),
          12,
        );
      if (next < today) next.setFullYear(next.getFullYear() + 1);
      return { r, date: next };
    })
    .filter((x) => daysUntil(x.date) <= 30)
    .sort((a, b) => a.date.getTime() - b.date.getTime());
  const employeeNotices = rows.filter((r) => {
      if (!r.noticeStart || !r.noticeEnd) return false;
      const start = new Date(r.noticeStart + "T12:00:00"),
        end = new Date(r.noticeEnd + "T12:00:00");
      return start <= today && end >= today;
    }),
    rowById = new Map(rows.map((r) => [r.id, r])),
    occurrenceNotices = occurrences
      .filter((item) => {
        if (item.type !== "Aviso" || !item.endDate) return false;
        const start = new Date(item.date + "T12:00:00"),
          end = new Date(item.endDate + "T12:00:00");
        return start <= today && end >= today && rowById.has(item.employeeId);
      })
      .map((item) => ({
        ...rowById.get(item.employeeId)!,
        noticeStart: item.date,
        noticeEnd: item.endDate,
      })),
    notices = [
      ...employeeNotices,
      ...occurrenceNotices.filter(
        (record) => !employeeNotices.some((item) => item.id === record.id),
      ),
    ],
    vacationToday = new Date(today),
    vacationEmployeeIds = new Set(
      occurrences
        .filter((item) => {
          if (item.type !== "Férias" || !item.endDate) return false;
          vacationToday.setHours(12, 0, 0, 0);
          const start = new Date(item.date + "T12:00:00"),
            end = new Date(item.endDate + "T12:00:00");
          return start <= vacationToday && end >= vacationToday;
        })
        .map((item) => item.employeeId),
    ),
    vacations = active.filter((employee) => vacationEmployeeIds.has(employee.id)),
    terminations = rows.filter((r) => {
      if (!r.terminationDate) return false;
      const date = new Date(r.terminationDate + "T12:00:00");
      return (
        date.getFullYear() === today.getFullYear() &&
        date.getMonth() === today.getMonth()
      );
    }),
    withoutFormal = active.filter((r) => r.formalEmployment === false),
    critical = active.filter((r) => r.experienceCritical),
    monthlyOccurrences = occurrences.filter((item) => {
      const date = new Date(item.date + "T12:00:00");
      return (
        rowById.has(item.employeeId) &&
        date.getFullYear() === today.getFullYear() &&
        date.getMonth() === today.getMonth()
      );
    }),
    occurrenceRanking = (type: "Falta" | "Atestado") => {
      const totals = new Map<number, number>();
      monthlyOccurrences
        .filter((item) => item.type === type)
        .forEach((item) =>
          totals.set(item.employeeId, (totals.get(item.employeeId) || 0) + 1),
        );
      return [...totals.entries()]
        .map(([employeeId, count]) => ({
          employee: rowById.get(employeeId)!,
          count,
        }))
        .filter((item) => item.employee && !isEmployeeDismissed(item.employee))
        .sort((a, b) => b.count - a.count || a.employee.employee.localeCompare(b.employee.employee))
        .slice(0, 5);
    },
    absenceRanking = occurrenceRanking("Falta"),
    certificateRanking = occurrenceRanking("Atestado");
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title="Recursos Humanos"
        sub="Funcionários por loja, experiências, aniversários e desligamentos"
      />
      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
        <Metric
          onClick={() =>
            setDetail({ title: "Funcionários ativos", rows: active })
          }
          title="Funcionários ativos"
          value={active.length}
          sub={`${stores.length} lojas`}
          icon={Users}
        />
        <Metric
          onClick={() =>
            setDetail({
              title: `Experiências que encerram em ${currentMonthLabel}`,
              rows: experiences.map((x) => x.r),
            })
          }
          title="Experiências vencendo"
          value={experiences.length}
          sub={`Encerram em ${currentMonthLabel}`}
          icon={Clock3}
        />
        <Metric
          onClick={() =>
            setDetail({ title: "Funcionários cumprindo aviso", rows: notices })
          }
          title="Cumprindo aviso"
          value={notices.length}
          sub="Avisos em andamento"
          icon={TriangleAlert}
        />
        <Metric
          onClick={() =>
            setDetail({
              title: `Rescisões de ${currentMonthLabel}`,
              rows: terminations,
            })
          }
          title="Rescisões cadastradas"
          value={terminations.length}
          sub={`Desligados em ${currentMonthLabel}`}
          icon={UserRound}
        />
        <Metric
          onClick={() =>
            setDetail({
              title: "Funcionários sem carteira assinada",
              rows: withoutFormal,
            })
          }
          title="Sem carteira assinada"
          value={withoutFormal.length}
          sub="Funcionários ativos"
          icon={TriangleAlert}
        />
        <Metric
          onClick={() =>
            setDetail({ title: "Funcionários críticos", rows: critical })
          }
          title="Funcionários críticos"
          value={critical.length}
          sub="Acompanhamento prioritário"
          icon={Bell}
        />
        <Metric
          onClick={() =>
            setDetail({ title: "Funcionários de férias", rows: vacations })
          }
          title="Funcionários de férias"
          value={vacations.length}
          sub="Férias em andamento"
          icon={Umbrella}
        />
      </div>
      <div className="mt-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <div>
          <h3 className="font-bold">Indicadores de frequência</h3>
          <p className="text-xs text-slate-400">
            Funcionários com mais ocorrências em {currentMonthLabel}
          </p>
        </div>
        <div className="mt-4 grid gap-4 lg:grid-cols-2">
          {[
            {
              title: "Mais faltas",
              singular: "falta",
              plural: "faltas",
              icon: TriangleAlert,
              ranking: absenceRanking,
              tone: "text-red-600 bg-red-50",
            },
            {
              title: "Mais atestados",
              singular: "atestado",
              plural: "atestados",
              icon: FileSpreadsheet,
              ranking: certificateRanking,
              tone: "text-amber-700 bg-amber-50",
            },
          ].map((indicator) => {
            const Icon = indicator.icon;
            return (
              <div key={indicator.title} className="rounded-xl border border-slate-200 p-4">
                <div className="flex items-center gap-3">
                  <div className={`grid h-10 w-10 place-items-center rounded-xl ${indicator.tone}`}>
                    <Icon size={19} />
                  </div>
                  <div>
                    <b className="text-sm">{indicator.title}</b>
                    <div className="text-xs text-slate-400">Ranking do mês</div>
                  </div>
                </div>
                <div className="mt-3 divide-y divide-slate-100">
                  {indicator.ranking.map(({ employee, count }, index) => (
                    <div key={employee.id} className="flex items-center gap-3 py-3">
                      <span className="grid h-7 w-7 shrink-0 place-items-center rounded-full bg-slate-100 text-xs font-bold text-slate-600">
                        {index + 1}
                      </span>
                      <div className="min-w-0 flex-1">
                        <div className="truncate text-sm font-semibold">{employee.employee}</div>
                        <div className="truncate text-xs text-slate-400">{employee.store}</div>
                      </div>
                      <span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs font-bold text-slate-700">
                        {count} {count === 1 ? indicator.singular : indicator.plural}
                      </span>
                    </div>
                  ))}
                  {!indicator.ranking.length && (
                    <p className="py-6 text-center text-sm text-slate-400">
                      Nenhum registro neste mês.
                    </p>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>
      <div className="mt-5 rounded-2xl border border-slate-300 bg-white p-5 shadow-soft">
        <div className="flex items-center gap-3">
          <div className="grid h-10 w-10 place-items-center rounded-xl bg-red-50 text-red-600">
            <Bell size={20} />
          </div>
          <div>
            <h3 className="font-bold">Mural de avisos</h3>
            <p className="text-xs text-slate-400">
              Críticos e experiências que encerram neste mês
            </p>
          </div>
        </div>
        <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
          {critical.map((r) => {
            const first = r.hiredAt && r.formalEmployment !== false ? addDays(r.hiredAt, 29) : null,
              end = r.hiredAt && r.formalEmployment !== false ? addDays(r.hiredAt, 89) : null;
            return (
              <div
                key={r.id}
                className="rounded-xl border border-red-200 bg-red-50 p-4"
              >
                <div className="flex items-center justify-between gap-2">
                  <b className="text-sm text-red-800">{r.employee}</b>
                  <span className="rounded-full bg-red-600 px-2 py-1 text-[10px] font-bold uppercase text-white">
                    Crítico
                  </span>
                </div>
                <div className="mt-2 text-xs text-red-700">
                  {r.store}
                  {r.formalEmployment === false ? (
                    <><br />Sem carteira assinada</>
                  ) : (
                    <>
                      <br />
                      30 dias: {first ? first.toLocaleDateString("pt-BR") : "-"}
                      <br />
                      Final 60 dias: {end ? end.toLocaleDateString("pt-BR") : "-"}
                    </>
                  )}
                </div>
              </div>
            );
          })}
          {experiences
            .filter((x) => !x.r.experienceCritical)
            .map(({ r, deadline, stage }) => (
              <div
                key={r.id}
                className="rounded-xl border border-amber-200 bg-amber-50 p-4"
              >
                <b className="text-sm text-amber-800">{r.employee}</b>
                <div className="mt-2 text-xs text-amber-700">
                  {r.store} · {stage}: {deadline.toLocaleDateString("pt-BR")}
                </div>
              </div>
            ))}
          {!critical.length && !experiences.length && (
            <p className="py-4 text-sm text-slate-400">
              Nenhum aviso prioritário no momento.
            </p>
          )}
        </div>
      </div>
      <div className="mt-5 grid gap-5 xl:grid-cols-2">
        <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
          <h3 className="font-bold">Funcionários por loja</h3>
          <div className="mt-4 space-y-3">
            {stores.map((store) => {
              const count = active.filter((r) => r.store === store).length,
                percent = active.length ? (count / active.length) * 100 : 0;
              return (
                <div key={store}>
                  <div className="flex justify-between text-sm">
                    <span>{store}</span>
                    <b>{count}</b>
                  </div>
                  <div className="mt-1.5 h-2 overflow-hidden rounded-full bg-slate-100">
                    <div
                      className="h-full rounded-full bg-forest-600"
                      style={{ width: `${percent}%` }}
                    />
                  </div>
                </div>
              );
            })}
            {!stores.length && (
              <p className="py-5 text-center text-sm text-slate-400">
                Nenhum funcionário ativo.
              </p>
            )}
          </div>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
          <h3 className="font-bold">Aniversários nos próximos 30 dias</h3>
          <div className="mt-3 divide-y divide-slate-100">
            {birthdays.map(({ r, date }) => (
              <div key={r.id} className="flex items-center py-3">
                <div className="grid h-9 w-9 place-items-center rounded-full bg-amber-50 text-sm font-bold text-amber-600">
                  {date.getDate()}
                </div>
                <div className="ml-3">
                  <b className="text-sm">{r.employee}</b>
                  <div className="text-xs text-slate-400">{r.store}</div>
                </div>
                <span className="ml-auto text-xs font-semibold text-slate-500">
                  {date.toLocaleDateString("pt-BR")}
                </span>
              </div>
            ))}
            {!birthdays.length && (
              <p className="py-5 text-center text-sm text-slate-400">
                Nenhum aniversário próximo cadastrado.
              </p>
            )}
          </div>
        </div>
      </div>
      <div className="mt-5 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="border-b border-slate-100 px-5 py-4">
          <h3 className="font-bold">Controle do quadro</h3>
          <p className="text-xs text-slate-400">
            Edite a situação de RH de cada funcionário.
          </p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[950px] text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase text-slate-400">
              <tr>
                {[
                  "Funcionário",
                  "Loja",
                  "Admissão",
                  "Períodos de experiência",
                  "Aniversário",
                  "Situação",
                  "Aviso / rescisão",
                  "Carteira",
                  "Ação",
                ].map((h) => (
                  <th key={h} className="px-4 py-3">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {rows
                .filter((r) => !isEmployeeDismissed(r))
                .map((r) => {
                  const firstExp = r.hiredAt && r.formalEmployment !== false ? addDays(r.hiredAt, 29) : null,
                    exp = r.hiredAt && r.formalEmployment !== false ? addDays(r.hiredAt, 89) : null,
                    status = isDismissalPending(r)
                      ? "Desligamento em andamento"
                      : r.employmentStatus || "Ativo";
                  return (
                    <tr key={r.id}>
                      <td className="px-4 py-4 font-semibold">
                        {r.employee}
                        <div className="text-xs font-normal text-slate-400">
                          {r.role}
                        </div>
                      </td>
                      <td className="px-4 py-4">{r.store}</td>
                      <td className="px-4 py-4">
                        {r.hiredAt ? formatDate(r.hiredAt) : "-"}
                      </td>
                      <td className="px-4 py-4">
                        {firstExp && exp ? (
                          <>
                            <div className="text-xs">
                              30 dias:{" "}
                              <b>{firstExp.toLocaleDateString("pt-BR")}</b>
                            </div>
                            <div className="mt-1 text-xs">
                              60 dias: <b>{exp.toLocaleDateString("pt-BR")}</b>
                            </div>
                          </>
                        ) : (
                          "-"
                        )}
                        {r.experienceCritical && (
                          <div className="mt-1 font-bold text-red-600">
                            CRÍTICO
                          </div>
                        )}
                      </td>
                      <td className="px-4 py-4">
                        {r.birthDate ? formatDate(r.birthDate) : "-"}
                      </td>
                      <td className="px-4 py-4">
                        <span
                          className={`rounded-full px-2.5 py-1 text-xs font-semibold ${status === "Aviso prévio" ? "bg-amber-50 text-amber-700" : status === "Desligado" ? "bg-slate-100 text-slate-500" : "bg-emerald-50 text-emerald-700"}`}
                        >
                          {status}
                        </span>
                      </td>
                      <td className="px-4 py-4 text-xs">
                        {r.noticeEnd && (
                          <div>Aviso até {formatDate(r.noticeEnd)}</div>
                        )}
                        {r.terminationDate && (
                          <div className="font-semibold text-red-600">
                            Rescisão: {formatDate(r.terminationDate)}
                          </div>
                        )}
                        {!r.noticeEnd && !r.terminationDate && "-"}
                      </td>
                      <td className="px-4 py-4 text-xs">
                        <div
                          className={
                            r.formalEmployment === false
                              ? "font-semibold text-amber-600"
                              : "text-emerald-700"
                          }
                        >
                          {r.formalEmployment === false
                            ? "Sem carteira"
                            : "Carteira assinada"}
                        </div>
                      </td>
                      <td className="px-4 py-4">
                        <button
                          onClick={() => onEdit(r)}
                          className="rounded-lg border border-forest-200 px-3 py-2 text-xs font-semibold text-forest-700"
                        >
                          Editar RH
                        </button>
                      </td>
                    </tr>
                  );
                })}
            </tbody>
          </table>
        </div>
      </div>
      {detail && (
        <IndicatorDetails
          title={detail.title}
          rows={detail.rows}
          occurrences={occurrences}
          close={() => setDetail(null)}
        />
      )}
    </main>
  );
}

function Metric({
  title,
  value,
  sub,
  icon: Icon,
  onClick,
}: {
  title: string;
  value: number;
  sub: string;
  icon: any;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="w-full rounded-2xl border border-slate-200 bg-white p-5 text-left shadow-soft transition hover:-translate-y-1 hover:border-slate-400 hover:shadow-lg"
    >
      <div className="flex items-start justify-between">
        <div className="grid h-10 w-10 place-items-center rounded-xl bg-forest-50 text-forest-700">
          <Icon size={20} />
        </div>
        <ChevronRight size={18} className="text-slate-400" />
      </div>
      <div className="mt-4 text-3xl font-bold text-slate-900">{value}</div>
      <div className="mt-1 text-sm font-semibold text-slate-700">{title}</div>
      <div className="mt-1 text-xs text-slate-400">{sub}</div>
    </button>
  );
}

function OccurrencesPage({
  employees,
  items,
  setItems,
  reportOnly = false,
  readOnly = false,
}: {
  employees: Recharge[];
  items: HROccurrence[];
  setItems: (v: HROccurrence[]) => void;
  reportOnly?: boolean;
  readOnly?: boolean;
}) {
  const now = new Date(),
    [detailType, setDetailType] = useState<HROccurrence["type"] | null>(null),
    [month, setMonth] = useState(
      `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`,
    ),
    [store, setStore] = useState("Todas"),
    [form, setForm] = useState({
      employeeId: "",
      date: new Date().toISOString().slice(0, 10),
      type: "Falta" as HROccurrence["type"],
      endDate: "",
      hours: "",
      minutes: "",
      days: "",
      note: "",
    });
  const stores = [...new Set(employees.map((e) => e.store))].sort(),
    eligible = employees.filter((e) => store === "Todas" || e.store === store),
    map = new Map(employees.map((e) => [e.id, e])),
    list = items
      .filter(
        (i) =>
          i.date.startsWith(month) &&
          map.has(i.employeeId) &&
          (store === "Todas" || map.get(i.employeeId)?.store === store),
      )
      .sort((a, b) => b.date.localeCompare(a.date)),
    counts = {
      Falta: list.filter((i) => i.type === "Falta").length,
      Atestado: list.filter((i) => i.type === "Atestado").length,
      Atraso: list.filter((i) => i.type === "Atraso").length,
      Aviso: list.filter((i) => i.type === "Aviso").length,
      Férias: list.filter((i) => i.type === "Férias").length,
    };
  const rows = eligible.map((employee) => {
    const own = list.filter((i) => i.employeeId === employee.id);
    return {
      employee,
      faults: own.filter((i) => i.type === "Falta").length,
      certificates: own.filter((i) => i.type === "Atestado").length,
      delays: own.filter((i) => i.type === "Atraso").length,
      warnings: own.filter((i) => i.type === "Aviso").length,
      vacations: own.filter((i) => i.type === "Férias").length,
      certificateDays: own
        .filter((i) => i.type === "Atestado")
        .reduce((s, i) => s + (i.days || 1), 0),
      minutes: own
        .filter((i) => i.type === "Atraso")
        .reduce((s, i) => s + (i.minutes || 0), 0),
    };
  });
  const add = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.employeeId) return;
    setItems([
      ...items,
      {
        id: Date.now(),
        employeeId: Number(form.employeeId),
        date: form.date,
        endDate:
          form.type === "Aviso" || form.type === "Férias"
            ? form.endDate
            : undefined,
        type: form.type,
        hours: form.type === "Atraso" ? Number(form.hours) || 0 : undefined,
        minutes: form.type === "Atraso" ? Number(form.minutes) || 0 : undefined,
        days: form.type === "Atestado" ? Number(form.days) || 1 : undefined,
        note: form.note,
      },
    ]);
    setForm({
      ...form,
      endDate: "",
      hours: "",
      minutes: "",
      days: "",
      note: "",
    });
  };
  const exportPdf = async () => {
    const logo = await fetch("/sacolao-abc-logo.png?v=4")
        .then((r) => r.blob())
        .then(
          (blob) =>
            new Promise<string>((resolve) => {
              const reader = new FileReader();
              reader.onload = () => resolve(String(reader.result));
              reader.readAsDataURL(blob);
            }),
        ),
      doc = new jsPDF({ orientation: "landscape" });
    doc.setFillColor(14, 78, 62);
    doc.rect(0, 0, 297, 34, "F");
    doc.addImage(logo, "PNG", 12, 5, 27, 23);
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(19);
    doc.text("Relatório mensal de RH", 45, 15);
    doc.setFontSize(9);
    doc.text(
      `${month.split("-").reverse().join("/")} - ${store === "Todas" ? "Todas as lojas" : store}`,
      45,
      23,
    );
    autoTable(doc, {
      startY: 42,
      head: [
        [
          "Funcionário",
          "Loja",
          "Faltas",
          "Atestados",
          "Dias atestado",
          "Atrasos",
          "Avisos",
          "Férias",
          "Min. atraso",
        ],
      ],
      body: rows.map((x) => [
        x.employee.employee,
        x.employee.store,
        x.faults,
        x.certificates,
        x.certificateDays,
        x.delays,
        x.warnings,
        x.vacations,
        x.minutes,
      ]),
      headStyles: { fillColor: [14, 78, 62] },
      alternateRowStyles: { fillColor: [244, 248, 246] },
      styles: { fontSize: 9, cellPadding: 3 },
    });
    doc.save(`relatorio-rh-${month}.pdf`);
  };
  const exportExcel = async () => {
    const wb = new ExcelJS.Workbook(),
      ws = wb.addWorksheet("Relatório RH", {
        views: [{ state: "frozen", ySplit: 5, showGridLines: false }],
      });
    const logo = await fetch("/sacolao-abc-logo.png?v=4").then((r) =>
        r.arrayBuffer(),
      ),
      id = wb.addImage({ buffer: logo as never, extension: "png" });
    ws.addImage(id, {
      tl: { col: 0.2, row: 0.1 },
      ext: { width: 145, height: 65 },
    });
    ws.mergeCells("C1:G2");
    ws.getCell("C1").value = "RELATÓRIO MENSAL DE RH";
    ws.getCell("C1").font = {
      size: 18,
      bold: true,
      color: { argb: "FF0E4E3E" },
    };
    ws.getCell("C3").value =
      `Mês: ${month.split("-").reverse().join("/")} | Loja: ${store}`;
    ws.mergeCells("C3:G3");
    ws.addRow([]);
    ws.addRow([
      "Funcionário",
      "Loja",
      "Faltas",
      "Atestados",
      "Dias de atestado",
      "Atrasos",
      "Avisos",
      "Férias",
      "Minutos de atraso",
    ]);
    rows.forEach((x) =>
      ws.addRow([
        x.employee.employee,
        x.employee.store,
        x.faults,
        x.certificates,
        x.certificateDays,
        x.delays,
        x.warnings,
        x.vacations,
        x.minutes,
      ]),
    );
    ws.getRow(5).eachCell((c) => {
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0E4E3E" },
      };
    });
    [28, 20, 10, 12, 16, 10, 10, 10, 18].forEach(
      (w, i) => (ws.getColumn(i + 1).width = w),
    );
    const bytes = await wb.xlsx.writeBuffer(),
      url = URL.createObjectURL(new Blob([bytes]));
    const a = document.createElement("a");
    a.href = url;
    a.download = `relatorio-rh-${month}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };
  return (
    <main className="fade-in p-4 sm:p-7">
      <SectionHead
        title={
          reportOnly ? "Relatórios de Recursos Humanos" : "Ocorrências mensais"
        }
        sub={
          reportOnly
            ? "Resumo mensal de faltas, atestados e atrasos"
            : "Registre faltas, atestados e atrasos dos funcionários"
        }
      />
      <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-5">
        <Metric
          onClick={() => setDetailType("Falta")}
          title="Faltas"
          value={counts.Falta}
          sub="No mês selecionado"
          icon={TriangleAlert}
        />
        <Metric
          onClick={() => setDetailType("Atestado")}
          title="Atestados"
          value={counts.Atestado}
          sub="No mês selecionado"
          icon={FileSpreadsheet}
        />
        <Metric
          onClick={() => setDetailType("Atraso")}
          title="Atrasos"
          value={counts.Atraso}
          sub="No mês selecionado"
          icon={Clock3}
        />
        <Metric
          onClick={() => setDetailType("Aviso")}
          title="Avisos"
          value={counts.Aviso}
          sub="No mês selecionado"
          icon={Bell}
        />
        <Metric
          onClick={() => setDetailType("Férias")}
          title="Férias"
          value={counts.Férias}
          sub="No mês selecionado"
          icon={Umbrella}
        />
      </div>
      <div className="mt-5 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft">
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <label className="text-xs font-semibold text-slate-600">
            Mês
            <input
              type="month"
              value={month}
              onChange={(e) => setMonth(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-2.5"
            />
          </label>
          <label className="text-xs font-semibold text-slate-600">
            Loja
            <select
              value={store}
              onChange={(e) => setStore(e.target.value)}
              className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-2.5"
            >
              <option>Todas</option>
              {stores.map((s) => (
                <option key={s}>{s}</option>
              ))}
            </select>
          </label>
          <button
            onClick={exportPdf}
            className="self-end rounded-xl border border-red-200 px-4 py-2.5 text-sm font-semibold text-red-600"
          >
            Gerar PDF
          </button>
          <button
            onClick={exportExcel}
            className="self-end rounded-xl bg-forest-700 px-4 py-2.5 text-sm font-semibold text-white"
          >
            Gerar Excel
          </button>
        </div>
      </div>
      {!reportOnly && !readOnly && (
        <form
          onSubmit={add}
          className="mt-5 grid items-end gap-3 rounded-2xl border border-slate-200 bg-white p-5 shadow-soft md:grid-cols-2 xl:grid-cols-6"
        >
          <SearchableEmployeeSelect
            required
            employees={eligible}
            value={form.employeeId}
            onChange={(employeeId) => setForm({ ...form, employeeId })}
            placeholder="Pesquisar funcionário..."
            className="md:col-span-2"
          />
          <label className="text-xs font-semibold text-slate-600">
            {form.type === "Férias" ? "Início das férias" : "Data"}
            <input
              required
              type="date"
              value={form.date}
              onChange={(e) => setForm({ ...form, date: e.target.value })}
              className="mt-1 h-12 w-full rounded-xl border border-slate-200 px-3 text-sm font-normal text-slate-800"
            />
          </label>
          <select
            value={form.type}
            onChange={(e) =>
              setForm({ ...form, type: e.target.value as HROccurrence["type"] })
            }
            className="h-12 w-full rounded-xl border border-slate-200 px-3"
          >
            {["Falta", "Atestado", "Atraso", "Aviso", "Férias"].map((t) => (
              <option key={t}>{t}</option>
            ))}
          </select>
          {form.type === "Atraso" ? (
            <div className="grid h-12 grid-cols-2 gap-2">
              <input
                required
                type="number"
                min="0"
                value={form.hours}
                onChange={(e) => setForm({ ...form, hours: e.target.value })}
                placeholder="Horas"
                title="Horas de atraso"
                className="h-12 min-w-0 rounded-xl border border-slate-200 px-3"
              />
              <input
                required
                type="number"
                min="0"
                max="59"
                value={form.minutes}
                onChange={(e) => setForm({ ...form, minutes: e.target.value })}
                placeholder="Minutos"
                title="Minutos de atraso"
                className="h-12 min-w-0 rounded-xl border border-slate-200 px-3"
              />
            </div>
          ) : form.type === "Atestado" ? (
            <input
              required
              type="number"
              min="1"
              value={form.days}
              onChange={(e) => setForm({ ...form, days: e.target.value })}
              placeholder="Número de dias"
              className="h-12 w-full rounded-xl border border-slate-200 px-3"
            />
          ) : form.type === "Aviso" || form.type === "Férias" ? (
            <label className="text-xs font-semibold text-slate-600">
              {form.type === "Férias" ? "Fim das férias" : "Término do aviso"}
              <input
                required
                type="date"
                min={form.date}
                value={form.endDate}
                onChange={(e) => setForm({ ...form, endDate: e.target.value })}
                className="mt-1 h-12 w-full rounded-xl border border-slate-200 px-3 text-sm font-normal text-slate-800"
              />
            </label>
          ) : <div className="hidden md:block" />}
          <button className="h-12 w-full rounded-xl bg-forest-700 px-4 font-semibold text-white">
            Registrar
          </button>
        </form>
      )}
      <div className="mt-5 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <table className="w-full min-w-[760px] text-left text-sm">
          <thead className="bg-slate-50 text-xs uppercase text-slate-400">
            <tr>
              {(reportOnly || readOnly
                ? ["Data", "Funcionário", "Loja", "Tipo", "Detalhe"]
                : ["Data", "Funcionário", "Loja", "Tipo", "Detalhe", "Ação"]
              ).map((h) => (
                <th key={h} className="px-4 py-3">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {list.map((i) => {
              const e = map.get(i.employeeId);
              return (
                <tr key={i.id}>
                  <td className="px-4 py-3">{formatDate(i.date)}</td>
                  <td className="px-4 py-3 font-semibold">{e?.employee}</td>
                  <td className="px-4 py-3">{e?.store}</td>
                  <td className="px-4 py-3">{i.type}</td>
                  <td className="px-4 py-3">
                    {i.type === "Atraso"
                      ? `${i.hours || 0}h ${i.minutes || 0}min de atraso`
                      : i.type === "Atestado"
                        ? `${i.days || 1} dia(s) de atestado`
                        : i.type === "Aviso"
                          ? `Até ${i.endDate ? formatDate(i.endDate) : "data não informada"}`
                          : i.type === "Férias"
                            ? `De ${formatDate(i.date)} até ${i.endDate ? formatDate(i.endDate) : "data não informada"}`
                          : "Falta registrada nesta data"}
                  </td>
                  {!reportOnly && !readOnly && (
                    <td className="px-4 py-3">
                      <button
                        onClick={() =>
                          setItems(items.filter((x) => x.id !== i.id))
                        }
                        className="text-xs font-semibold text-red-600"
                      >
                        Excluir
                      </button>
                    </td>
                  )}
                </tr>
              );
            })}
            {!list.length && (
              <tr>
                <td
                  colSpan={reportOnly || readOnly ? 5 : 6}
                  className="py-10 text-center text-slate-400"
                >
                  Nenhuma ocorrência neste mês.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {detailType &&
        createPortal(
          <div className="fixed inset-0 z-[80] grid place-items-center bg-slate-950/50 p-4 backdrop-blur-sm">
            <div className="max-h-[85vh] w-full max-w-3xl overflow-hidden rounded-2xl bg-white shadow-2xl">
              <div className="flex items-center border-b border-slate-200 px-5 py-4">
                <div>
                  <h3 className="text-lg font-bold text-slate-900">
                    {detailType === "Aviso"
                      ? "Avisos"
                      : detailType === "Férias"
                        ? "Férias"
                        : `${detailType}s`} do mês
                  </h3>
                  <p className="text-xs text-slate-400">
                    {list.filter((item) => item.type === detailType).length} registro(s)
                  </p>
                </div>
                <button
                  onClick={() => setDetailType(null)}
                  className="ml-auto rounded-lg p-2 text-slate-500 hover:bg-slate-100"
                >
                  <X size={20} />
                </button>
              </div>
              <div className="max-h-[65vh] space-y-2 overflow-y-auto p-4">
                {list
                  .filter((item) => item.type === detailType)
                  .map((item) => {
                    const employee = map.get(item.employeeId);
                    const detail =
                      item.type === "Atestado"
                        ? `${item.days || 1} dia(s) de atestado`
                        : item.type === "Aviso"
                          ? `De ${formatDate(item.date)} até ${item.endDate ? formatDate(item.endDate) : "data não informada"}`
                          : item.type === "Férias"
                            ? `De ${formatDate(item.date)} até ${item.endDate ? formatDate(item.endDate) : "data não informada"}`
                          : item.type === "Falta"
                            ? "Falta registrada nesta data"
                            : `${item.hours || 0}h ${item.minutes || 0}min`;
                    return (
                      <div key={item.id} className="rounded-xl border border-slate-200 p-4">
                        <div className="flex flex-wrap items-start justify-between gap-3">
                          <div>
                            <b className="text-sm text-slate-900">{employee?.employee || "Funcionário não encontrado"}</b>
                            <div className="mt-1 text-xs text-slate-500">{employee?.store || "Loja não informada"}</div>
                          </div>
                          <div className="text-right text-xs">
                            <b className="text-slate-700">{formatDate(item.date)}</b>
                            <div className="mt-1 text-slate-500">{detail}</div>
                          </div>
                        </div>
                        {item.note && <p className="mt-3 text-xs text-slate-500">{item.note}</p>}
                      </div>
                    );
                  })}
                {!list.some((item) => item.type === detailType) && (
                  <p className="py-12 text-center text-sm text-slate-400">
                    Nenhum registro nesta categoria.
                  </p>
                )}
              </div>
            </div>
          </div>,
          document.body,
        )}
    </main>
  );
}

function AdministrativePage({ page, employees, companies, companyCnpjs, financialEntries }: { page: string; employees: Recharge[]; companies: string[]; companyCnpjs: Record<string, string>; financialEntries: FinancialEntry[] }) {
  const today = new Date().toISOString().slice(0, 10);
  const [employeeId, setEmployeeId] = useState("");
  const [documentDate, setDocumentDate] = useState(today);
  const [occurredDates, setOccurredDates] = useState<string[]>([today]);
  const [reason, setReason] = useState(
    "falta injustificada, sem apresentação de justificativa válida",
  );
  const [warningSummary, setWarningSummary] = useState("");
  const [generatingReason, setGeneratingReason] = useState(false);
  const [warningPersonMode, setWarningPersonMode] = useState<"registered" | "manual">("registered");
  const [warningManualName, setWarningManualName] = useState("");
  const [warningManualCpf, setWarningManualCpf] = useState("");
  const [warningManualGender, setWarningManualGender] = useState<Recharge["gender"] | "">("");
  const [receiptEmployeeId, setReceiptEmployeeId] = useState("");
  const [receiptPersonMode, setReceiptPersonMode] = useState<"registered" | "manual">("registered");
  const [receiptManualName, setReceiptManualName] = useState("");
  const [receiptManualCpf, setReceiptManualCpf] = useState("");
  const [receiptCompany, setReceiptCompany] = useState("");
  const [receiptDate, setReceiptDate] = useState(today);
  const [receiptKind, setReceiptKind] = useState<"salary" | "advance">("salary");
  const [receiptPeriod, setReceiptPeriod] = useState(today.slice(0, 7));
  const [receiptSalaryStart, setReceiptSalaryStart] = useState(`${today.slice(0, 7)}-01`);
  const [receiptSalaryEnd, setReceiptSalaryEnd] = useState(() => {
    const [year, month] = today.slice(0, 7).split("-").map(Number);
    return new Date(year, month, 0, 12).toISOString().slice(0, 10);
  });
  const [receiptGross, setReceiptGross] = useState("");
  const [receiptAdvance, setReceiptAdvance] = useState("");
  const [receiptAdditions, setReceiptAdditions] = useState<Array<{ id: number; name: string; mode: "value" | "percent"; value: string }>>([]);
  const [receiptDiscounts, setReceiptDiscounts] = useState<Array<{ id: number; name: string; mode: "value" | "percent"; value: string }>>([]);
  const [genericAmount, setGenericAmount] = useState("");
  const [genericReference, setGenericReference] = useState("");
  const [improvingGenericReference, setImprovingGenericReference] = useState(false);
  const [improvingWarningReason, setImprovingWarningReason] = useState(false);
  useEffect(() => {
    if (page === "Salário") setReceiptKind("salary");
    if (page === "Adiantamento") setReceiptKind("advance");
  }, [page]);
  const selected = employees.find((employee) => String(employee.id) === employeeId);
  const activeEmployees = employees
    .filter((employee) => employee.active !== false)
    .sort((a, b) => a.employee.localeCompare(b.employee, "pt-BR"));
  const receiptEmployee = employees.find((employee) => String(employee.id) === receiptEmployeeId);
  const warningPerson = warningPersonMode === "registered"
    ? selected
    : warningManualName.trim()
      ? { employee: warningManualName.trim(), cpf: warningManualCpf.trim(), gender: warningManualGender || undefined }
      : undefined;
  const receiptPerson = receiptPersonMode === "registered"
    ? receiptEmployee
    : receiptManualName.trim()
      ? { employee: receiptManualName.trim(), cpf: receiptManualCpf.trim() }
      : undefined;
  const salaryPeriodDays = (() => {
    if (!receiptSalaryStart || !receiptSalaryEnd) return 0;
    const start = new Date(`${receiptSalaryStart}T12:00:00`);
    const end = new Date(`${receiptSalaryEnd}T12:00:00`);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end < start) return 0;
    // Para fins de salário mensal, a competência integral equivale a 30 dias,
    // inclusive nos meses civis com 31 dias.
    return Math.min(30, Math.floor((end.getTime() - start.getTime()) / 86_400_000) + 1);
  })();
  const proportionalSalary = (parseMoney(receiptGross) / 30) * salaryPeriodDays;
  const receiptAdditionTotal = receiptAdditions.reduce((total, item) => {
    const value = parseMoney(item.value);
    return total + (item.mode === "percent" ? (proportionalSalary * value) / 100 : value);
  }, 0);
  const receiptDiscountTotal = receiptDiscounts.reduce((total, item) => {
    const value = parseMoney(item.value);
    return total + (item.mode === "percent" ? (proportionalSalary * value) / 100 : value);
  }, 0);
  useEffect(() => {
    const [year, month] = receiptPeriod.split("-").map(Number);
    if (!year || !month) return;
    setReceiptSalaryStart(`${receiptPeriod}-01`);
    setReceiptSalaryEnd(new Date(year, month, 0, 12).toISOString().slice(0, 10));
  }, [receiptPeriod]);
  useEffect(() => {
    const entry = financialEntries.find((item) => item.employeeId === Number(receiptEmployeeId) && item.period === receiptPeriod);
    if (!entry) {
      if (receiptKind === "salary") setReceiptAdvance("");
      return;
    }
    setReceiptGross(formatMoneyInput(String(Math.round(Number(entry.salary || 0) * 100))));
    setReceiptAdvance(receiptKind === "advance" ? formatMoneyInput(String(Math.round(Number(entry.advance || 0) * 100))) : "");
  }, [receiptEmployeeId, receiptPeriod, financialEntries, receiptKind]);
  const longDate = (value: string) => {
    if (!value) return "";
    return new Date(`${value}T12:00:00`).toLocaleDateString("pt-BR", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });
  };
  const administrativeFileName = (type: string, fullName: string, date: string) => {
    const nameParts = fullName.trim().replace(/\s+/g, " ").split(" ");
    const shortName = nameParts.length > 1
      ? `${nameParts[0]} ${nameParts[nameParts.length - 1]}`
      : nameParts[0] || "Funcionário";
    const [year, month, day] = date.split("-");
    const shortDate = day && month && year ? `${day}.${month}` : "sem data";
    const safeType = type.replace(/[\\/:*?"<>|]/g, "").trim();
    const safeName = shortName.replace(/[\\/:*?"<>|]/g, "").trim();
    return `${safeType} - ${safeName} ${shortDate}.pdf`;
  };
  const generateReasonWithAI = async () => {
    if (warningSummary.trim().length < 5) {
      alert("Conte resumidamente o que aconteceu para a IA criar o motivo.");
      return;
    }
    setGeneratingReason(true);
    try {
      const response = await fetch("/api/warning-reason", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ description: warningSummary.trim(), gender: warningPerson?.gender }),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.reason) throw new Error(data.error || "Não foi possível gerar o texto.");
      setReason(String(data.reason));
      window.dispatchEvent(new CustomEvent("abc:toast", { detail: "Motivo criado com IA. Revise antes de gerar." }));
    } catch (error) {
      alert(error instanceof Error ? error.message : "Não foi possível gerar o texto com IA.");
    } finally {
      setGeneratingReason(false);
    }
  };
  const improveAdministrativeText = async (
    text: string,
    context: "receipt_reference" | "warning_reason",
    update: (value: string) => void,
    setLoading: (value: boolean) => void,
  ) => {
    if (text.trim().length < 2) {
      alert("Digite primeiro um texto para a IA melhorar.");
      return;
    }
    setLoading(true);
    try {
      const response = await fetch("/api/improve-text", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: text.trim(), context }),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.text) throw new Error(data.error || "Não foi possível melhorar o texto.");
      update(String(data.text));
      window.dispatchEvent(new CustomEvent("abc:toast", { detail: "Texto melhorado com IA. Revise antes de gerar." }));
    } catch (error) {
      alert(error instanceof Error ? error.message : "Não foi possível melhorar o texto com IA agora.");
    } finally {
      setLoading(false);
    }
  };
  const generateWarning = async () => {
    const validOccurredDates = [...new Set(occurredDates.filter(Boolean))].sort();
    if (!warningPerson || !warningPerson.cpf || !documentDate || !validOccurredDates.length || !reason.trim()) {
      alert("Selecione um funcionário ou informe nome e CPF, além das datas e do motivo.");
      return;
    }
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    const left = 23,
      right = 187,
      width = right - left;
    await addCompanyLogo(doc, "HLM GESTÃO LTDA", 79, 8, 52);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(15);
    doc.setFontSize(10.5);
    doc.text("55.566.792/0001-58", 105, 47, { align: "center" });
    doc.setFontSize(15);
    doc.text("ADVERTÊNCIA DISCIPLINAR", 105, 61, { align: "center" });
    doc.setFontSize(10.5);
    doc.setFont("helvetica", "normal");
    doc.text(`Pessoa advertida: ${warningPerson.employee}`, left, 75);
    doc.text(`CPF: ${warningPerson.cpf}`, left, 83);
    const collaboratorReference = warningPerson.gender === "Feminino"
      ? "a colaboradora mencionada"
      : warningPerson.gender === "Masculino"
        ? "o colaborador mencionado"
        : "a pessoa mencionada";
    const cleanedReason = reason
      .trim()
      .replace(/\s+/g, " ")
      .replace(/[.]+$/, "")
      .replace(/^(?:foi\s+(?:constatado|apurado|verificado)\s+que|constatou-se\s+que|ocorreu(?:\s+que)?)[,:\s]*/i, "")
      .replace(/\bna data informada[,]?\s*/gi, "")
      .replace(/\b(?:um|uma|o|a)\s+(?:funcionári[oa]|colaborador[oa])(?:\s+mencionad[oa])?\b/gi, collaboratorReference)
      .replace(/\b(?:funcionári[oa]|colaborador[oa])\s+mencionad[oa]\b/gi, collaboratorReference)
      .replace(/\ba pessoa mencionada\b/gi, collaboratorReference);
    const reasonClause = cleanedReason
      ? cleanedReason.charAt(0).toLowerCase() + cleanedReason.slice(1)
      : "houve a ocorrência informada";
    const occurredDatesText = validOccurredDates.map((date) => longDate(date));
    const occurredWhen = occurredDatesText.length === 1
      ? `no dia ${occurredDatesText[0]}`
      : `nos dias ${occurredDatesText.slice(0, -1).join(", ")} e ${occurredDatesText.at(-1)}`;
    const body =
      `Vimos, pelo presente, aplicar-lhe advertência disciplinar, pois foi apurado que, ${occurredWhen}, ${reasonClause}. Tal conduta representa descumprimento das obrigações inerentes ao contrato de trabalho e poderá caracterizar desídia no desempenho das funções, na forma do art. 482, letra “e”, da Consolidação das Leis do Trabalho (CLT).`;
    const warning =
      "Esclarecemos, ainda, que a repetição de procedimentos como este poderá ser considerada ato faltoso, passível de suspensão e, consequentemente, de dispensa por justa causa, conforme o artigo 482 da CLT.";
    doc.setFontSize(11);
    const bodyLines = doc.splitTextToSize(body, width);
    doc.text(bodyLines, left, 99, { align: "justify", maxWidth: width, lineHeightFactor: 1.55 });
    const warningY = 99 + bodyLines.length * 6.2 + 8;
    const warningLines = doc.splitTextToSize(warning, width);
    doc.text(warningLines, left, warningY, {
      align: "justify",
      maxWidth: width,
      lineHeightFactor: 1.55,
    });
    const dateY = Math.max(158, warningY + warningLines.length * 6.2 + 16);
    doc.text(`Belo Horizonte, ${longDate(documentDate)}.`, left, dateY);
    const signatureY = 215;
    doc.line(23, signatureY, 88, signatureY);
    doc.line(119, signatureY, 184, signatureY);
    doc.setFontSize(9.5);
    doc.text("Pessoa advertida", 55.5, signatureY + 6, { align: "center" });
    doc.text("Responsável legal (quando menor)", 151.5, signatureY + 6, {
      align: "center",
    });
    doc.line(23, 239, 88, 239);
    doc.text("Empresa", 55.5, 245, { align: "center" });
    doc.setFontSize(9);
    doc.text("Testemunhas", left, 260);
    doc.line(23, 270, 88, 270);
    doc.line(119, 270, 184, 270);
    doc.text("Nome", 55.5, 276, { align: "center" });
    doc.text("CPF", 151.5, 276, { align: "center" });
    doc.save(administrativeFileName("Advertência", warningPerson.employee, documentDate));
    window.dispatchEvent(
      new CustomEvent("abc:toast", { detail: "Advertência gerada com sucesso" }),
    );
  };
  const generateReceipt = async () => {
    const monthlyGross = parseMoney(receiptGross);
    const gross = receiptKind === "salary" ? proportionalSalary : monthlyGross;
    const discount = receiptDiscountTotal, advance = parseMoney(receiptAdvance);
    const amount = receiptKind === "salary" ? Math.max(0, gross + receiptAdditionTotal - discount - advance) : advance;
    if (!receiptPerson || !receiptPerson.cpf || !receiptCompany || !receiptDate || !receiptPeriod || !amount) {
      alert("Selecione um funcionário ou informe nome e CPF, além da empresa, competência, data e valores.");
      return;
    }
    if (receiptKind === "salary" && (!receiptSalaryStart || !receiptSalaryEnd || salaryPeriodDays < 1)) {
      alert("Informe uma data inicial e uma data final válidas para calcular os dias trabalhados.");
      return;
    }
    const cnpj = companyCnpjs[receiptCompany] || "CNPJ não informado";
    const [year, month] = receiptPeriod.split("-").map(Number);
    const periodStart = new Date(`${receiptSalaryStart}T12:00:00`).toLocaleDateString("pt-BR");
    const periodEnd = new Date(`${receiptSalaryEnd}T12:00:00`).toLocaleDateString("pt-BR");
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    const isHlmCompany = /\bHLM\b/i.test(receiptCompany);
    const hasCompanyLogo = await addCompanyLogo(doc, receiptCompany, 79, 8, 52);
    const companyY = hasCompanyLogo ? 45 : 22;
    const cnpjY = hasCompanyLogo ? 53 : 30;
    const dividerY = hasCompanyLogo ? 61 : 40;
    const titleY = hasCompanyLogo ? 74 : 55;
    const employeeY = hasCompanyLogo ? 87 : 68;
    const cpfY = hasCompanyLogo ? 95 : 76;
    const receivedY = hasCompanyLogo ? 108 : 89;
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    if (!isHlmCompany) doc.text(receiptCompany.toUpperCase(), 105, companyY, { align: "center" });
    doc.text(cnpj, 105, cnpjY, { align: "center" });
    doc.line(30, dividerY, 180, dividerY);
    doc.setFontSize(16);
    doc.text(receiptKind === "salary" ? "RECIBO DE PAGAMENTO" : "RECIBO DE ADIANTAMENTO", 30, titleY);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text(`Empregado(a): ${receiptPerson.employee}`, 30, employeeY);
    doc.text(`CPF: ${receiptPerson.cpf}`, 30, cpfY);
    const receivedText = `Recebi da empresa ${receiptCompany.toUpperCase()} a importância de ${amount.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })} (${moneyInWords(amount)}),`;
    const receivedLines = doc.splitTextToSize(receivedText, 150);
    doc.text(receivedLines, 30, receivedY, { lineHeightFactor: 1.35 });
    const referenceText = receiptKind === "salary"
      ? `referente ao pagamento de salário do período de ${periodStart} até ${periodEnd}.`
      : `referente ao adiantamento salarial da competência ${String(month).padStart(2, "0")}/${year}.`;
    const referenceY = receivedY + receivedLines.length * 5.5 + 3;
    const referenceLines = doc.splitTextToSize(referenceText, 150);
    doc.text(referenceLines, 30, referenceY, { lineHeightFactor: 1.35 });
    const tableLineY = referenceY + referenceLines.length * 5.5 + 8;
    doc.line(30, tableLineY, 180, tableLineY);
    doc.setFont("helvetica", "bold");
    const tableHeaderY = tableLineY + 12;
    doc.text("Descrição", 30, tableHeaderY);
    doc.text("Valor (R$)", 155, tableHeaderY);
    doc.setFont("helvetica", "normal");
    const currency = (value: number) => value.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    let detailY = tableHeaderY + 12;
    if (receiptKind === "salary") {
      doc.text(`+ Saldo de salário - ${salaryPeriodDays} dias`, 30, detailY); doc.text(currency(gross), 155, detailY);
      receiptAdditions.forEach((item) => {
        const informedValue = parseMoney(item.value);
        const value = item.mode === "percent" ? (gross * informedValue) / 100 : informedValue;
        if (!value) return;
        detailY += 7;
        doc.text(`+ ${item.name.trim() || "Valor adicional"}${item.mode === "percent" ? ` (${informedValue.toLocaleString("pt-BR")}% do bruto)` : ""}`, 30, detailY); doc.text(currency(value), 155, detailY);
      });
      receiptDiscounts.forEach((item) => {
        detailY += 7;
        const value = item.mode === "percent" ? (gross * parseMoney(item.value)) / 100 : parseMoney(item.value);
        const suffix = item.mode === "percent" ? ` (${parseMoney(item.value).toLocaleString("pt-BR")}% do bruto)` : "";
        doc.text(`- ${item.name.trim() || "Desconto"}${suffix}`, 30, detailY); doc.text(currency(value), 155, detailY);
      });
      if (advance > 0) {
        detailY += 7;
        doc.text(`- Adiantamento 20/${String(month).padStart(2, "0")}`, 30, detailY); doc.text(currency(advance), 155, detailY);
      }
      detailY += 10;
      doc.setFont("helvetica", "bold"); doc.text("Total recebido", 30, detailY); doc.text(currency(amount), 155, detailY);
    } else {
      doc.text(`+ Adiantamento salarial 20/${String(month).padStart(2, "0")}`, 30, detailY); doc.text(currency(advance), 155, detailY);
      detailY += 10;
      doc.setFont("helvetica", "bold"); doc.text("Total recebido", 30, detailY); doc.text(currency(amount), 155, detailY);
    }
    const declarationLineY = Math.max(184, detailY + 8);
    doc.line(30, declarationLineY, 180, declarationLineY);
    doc.setFont("helvetica", "normal");
    doc.text("Declaro que recebi a quantia acima discriminada, dando plena, geral e irrevogável", 30, declarationLineY + 10);
    doc.text("quitação dos valores mencionados.", 30, declarationLineY + 17);
    const receiptDateText = new Date(`${receiptDate}T12:00:00`).toLocaleDateString("pt-BR", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });
    doc.line(30, declarationLineY + 24, 180, declarationLineY + 24);
    doc.setFont("helvetica", "bold");
    const dateY = declarationLineY + 34;
    doc.text(`Belo Horizonte, ${receiptDateText}.`, 30, dateY);
    const signatureY = dateY + 20;
    doc.line(30, signatureY, 180, signatureY);
    doc.setFont("helvetica", "bold");
    doc.text("Assinatura do(a) empregado(a)", 30, signatureY + 8);
    doc.setFont("helvetica", "normal");
    doc.text(receiptPerson.employee, 30, signatureY + 16);
    doc.text(`CPF: ${receiptPerson.cpf}`, 30, signatureY + 24);
    doc.save(administrativeFileName(receiptKind === "salary" ? "Salário" : "Adiantamento", receiptPerson.employee, receiptDate));
    window.dispatchEvent(new CustomEvent("abc:toast", { detail: "Recibo gerado com sucesso" }));
  };
  const generateGenericReceipt = async () => {
    const amount = parseMoney(genericAmount);
    if (!receiptPerson || !receiptPerson.cpf || !receiptCompany || !receiptDate || !amount || !genericReference.trim()) {
      alert("Preencha nome, CPF, loja, data, valor e a referência do recibo.");
      return;
    }
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    const isHlmCompany = /\bHLM\b/i.test(receiptCompany);
    const hasCompanyLogo = await addCompanyLogo(doc, receiptCompany, 79, 8, 52);
    const titleY = hasCompanyLogo ? 54 : 28;
    const companyY = hasCompanyLogo ? 70 : 44;
    const cnpjY = hasCompanyLogo ? 78 : 52;
    const employeeY = hasCompanyLogo ? 98 : 72;
    const declarationY = hasCompanyLogo ? 114 : 88;
    const dateY = hasCompanyLogo ? 165 : 128;
    const signatureY = hasCompanyLogo ? 210 : 168;
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(18);
    doc.text("RECIBO", 105, titleY, { align: "center" });
    doc.setFontSize(11);
    if (!isHlmCompany) doc.text(receiptCompany.toUpperCase(), 105, companyY, { align: "center" });
    doc.text(companyCnpjs[receiptCompany] || "CNPJ não informado", 105, cnpjY, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.text(`EU, ${receiptPerson.employee.toUpperCase()}, CPF ${receiptPerson.cpf},`, 105, employeeY, { align: "center" });
    const declaration = `DECLARO QUE RECEBI da empresa ${receiptCompany.toUpperCase()} a importância de ${amount.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })} (${moneyInWords(amount)}), referente a ${genericReference.trim()}.`;
    const declarationLines = doc.splitTextToSize(declaration, 150);
    doc.text(declarationLines, 105, declarationY, { align: "center", lineHeightFactor: 1.5 });
    const dateText = new Date(`${receiptDate}T12:00:00`).toLocaleDateString("pt-BR", { day: "2-digit", month: "long", year: "numeric" });
    doc.text(`Belo Horizonte, ${dateText}.`, 105, dateY, { align: "center" });
    doc.line(35, signatureY, 175, signatureY);
    doc.setFont("helvetica", "bold");
    doc.text("ASSINATURA DO(A) RECEBEDOR(A)", 105, signatureY + 9, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.text(receiptPerson.employee, 105, signatureY + 19, { align: "center" });
    doc.text(`CPF: ${receiptPerson.cpf}`, 105, signatureY + 28, { align: "center" });
    doc.save(administrativeFileName("Recibo", receiptPerson.employee, receiptDate));
    window.dispatchEvent(new CustomEvent("abc:toast", { detail: "Recibo gerado com sucesso" }));
  };
  return (
    <main className="p-4 sm:p-7">
      <div className="mb-6">
        <h1 className="text-2xl font-black text-slate-900">Administrativo</h1>
        <p className="text-sm text-slate-500">
          Gere documentos padronizados usando os dados dos funcionários.
        </p>
      </div>
      {(page === "Salário" || page === "Adiantamento") && <section className="administrative-document mx-auto mb-6 max-w-6xl overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="border-b border-slate-200 bg-gradient-to-r from-slate-950 to-slate-800 px-6 py-5 text-white">
          <div className="flex items-center gap-3">
            <span className="grid h-11 w-11 place-items-center rounded-xl bg-white/10 text-white"><ReceiptText size={21} /></span>
            <div><h2 className="text-xl font-black">{page === "Salário" ? "Recibo de Salário" : "Recibo de Adiantamento"}</h2><p className="text-sm text-slate-300">{page === "Salário" ? "Informe o período, confira o cálculo e gere o documento." : "Preencha os dados do adiantamento e gere o documento."}</p></div>
          </div>
        </div>
        <div className="grid gap-x-5 gap-y-4 p-6 sm:grid-cols-2">
          <div className="sm:col-span-2 flex flex-col gap-3 border-b border-slate-200 pb-5 sm:flex-row sm:items-center sm:justify-between">
            <div><b className="text-sm text-slate-900">1. Identificação</b><p className="text-xs text-slate-500">Escolha um funcionário ou preencha os dados manualmente.</p></div>
            <div className="administrative-segmented grid min-w-0 grid-cols-2 rounded-xl bg-slate-100 p-1 sm:min-w-[390px]">
              <button type="button" onClick={() => setReceiptPersonMode("registered")} className={`rounded-lg px-4 py-2.5 text-sm font-bold transition ${receiptPersonMode === "registered" ? "administrative-tab-active bg-white text-slate-950 shadow-sm" : "administrative-tab-inactive text-slate-500 hover:text-slate-800"}`}>Funcionário cadastrado</button>
              <button type="button" onClick={() => setReceiptPersonMode("manual")} className={`rounded-lg px-4 py-2.5 text-sm font-bold transition ${receiptPersonMode === "manual" ? "administrative-tab-active bg-white text-slate-950 shadow-sm" : "administrative-tab-inactive text-slate-500 hover:text-slate-800"}`}>Preencher manualmente</button>
            </div>
          </div>
          {receiptPersonMode === "registered" ? (
            <label className="sm:col-span-2"><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Funcionário</span><SearchableEmployeeSelect employees={activeEmployees} value={receiptEmployeeId} onChange={setReceiptEmployeeId} /></label>
          ) : (
            <>
              <label><span className="mb-2 block text-sm font-bold text-slate-700">Nome completo</span><input value={receiptManualName} onChange={(event) => setReceiptManualName(event.target.value)} placeholder="Digite o nome completo" className="h-12 w-full rounded-xl border border-slate-200 px-4" /></label>
              <label><span className="mb-2 block text-sm font-bold text-slate-700">CPF</span><input inputMode="numeric" maxLength={14} value={receiptManualCpf} onChange={(event) => setReceiptManualCpf(formatCpf(event.target.value))} placeholder="000.000.000-00" className="h-12 w-full rounded-xl border border-slate-200 px-4" /></label>
            </>
          )}
          <div className="sm:col-span-2 mt-2 border-t border-slate-200 pt-5"><b className="text-sm text-slate-900">2. Documento e período</b><p className="text-xs text-slate-500">Defina a empresa, a competência e as datas do recibo.</p></div>
          <label><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Loja / empresa</span><select value={receiptCompany} onChange={(event) => setReceiptCompany(event.target.value)} className="h-12 w-full rounded-xl border border-slate-300 bg-white px-4 font-semibold shadow-sm"><option value="">Selecione a loja</option>{companies.map((company) => <option key={company} value={company}>{company}{companyCnpjs[company] ? ` - ${companyCnpjs[company]}` : ""}</option>)}</select>{!companies.length && <span className="mt-1 block text-xs text-amber-600">Cadastre a loja e o CNPJ em Configurações → Lojas e empresas.</span>}</label>
          <label><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Competência</span><input type="month" value={receiptPeriod} onChange={(event) => setReceiptPeriod(event.target.value)} className="h-12 w-full rounded-xl border border-slate-300 bg-white px-4 font-semibold shadow-sm" /></label>
          <label><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Data do recibo</span><input type="date" value={receiptDate} onChange={(event) => setReceiptDate(event.target.value)} className="h-12 w-full rounded-xl border border-slate-300 bg-white px-4 font-semibold shadow-sm" /></label>
          {receiptKind === "salary" && <>
            <label><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Início do período trabalhado</span><input type="date" value={receiptSalaryStart} onChange={(event) => setReceiptSalaryStart(event.target.value)} className="h-12 w-full rounded-xl border border-slate-300 bg-white px-4 font-semibold shadow-sm" /></label>
            <label><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Fim do período trabalhado</span><input type="date" min={receiptSalaryStart} value={receiptSalaryEnd} onChange={(event) => setReceiptSalaryEnd(event.target.value)} className="h-12 w-full rounded-xl border border-slate-300 bg-white px-4 font-semibold shadow-sm" /></label>
            <div className="sm:col-span-2 mt-2 border-t border-slate-200 pt-5"><b className="text-sm text-slate-900">3. Valores</b><p className="text-xs text-slate-500">Informe o salário e confira o cálculo automático.</p></div>
            <label className="sm:col-span-2"><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Salário bruto mensal</span><input inputMode="decimal" value={receiptGross} onChange={(event) => setReceiptGross(formatMoneyInput(event.target.value))} placeholder="0,00" className="h-14 w-full rounded-xl border border-slate-300 bg-white px-4 text-lg font-black shadow-sm" /></label>
            <div className="sm:col-span-2 grid gap-4 rounded-3xl border border-blue-200 bg-gradient-to-br from-blue-50 to-indigo-50 p-5 shadow-sm sm:grid-cols-3">
              <div><span className="text-xs font-bold uppercase tracking-wide text-blue-600">Dias considerados</span><b className="mt-1 block text-xl text-slate-900">{salaryPeriodDays} dias</b></div>
              <div><span className="text-xs font-bold uppercase tracking-wide text-blue-600">Valor por dia</span><b className="mt-1 block text-xl text-slate-900">{(parseMoney(receiptGross) / 30).toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</b></div>
              <div><span className="text-xs font-bold uppercase tracking-wide text-blue-600">Salário proporcional</span><b className="mt-1 block text-xl text-blue-700">{proportionalSalary.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</b></div>
              <p className="text-xs text-blue-700 sm:col-span-3">Cálculo: salário bruto ÷ 30 × {salaryPeriodDays} dias. Um mês completo considera no máximo 30 dias.</p>
            </div>
            <div className="sm:col-span-2 rounded-2xl border border-slate-200 bg-slate-50 p-4">
              <div className="mb-3 flex items-center justify-between gap-3"><div><b className="text-sm text-slate-800">Valores adicionais</b><p className="text-xs text-slate-500">Inclua um valor fixo ou uma porcentagem sobre o salário proporcional.</p></div><button type="button" onClick={() => setReceiptAdditions([...receiptAdditions, { id: Date.now(), name: "", mode: "value", value: "" }])} className="flex shrink-0 items-center gap-1 rounded-xl bg-slate-900 px-3 py-2 text-xs font-bold text-white"><Plus size={15} />Incluir valor</button></div>
              <div className="space-y-3">
                {receiptAdditions.map((item) => <div key={item.id} className="grid gap-2 rounded-xl border border-slate-200 bg-white p-3 sm:grid-cols-[1fr_150px_180px_auto]">
                  <input value={item.name} onChange={(event) => setReceiptAdditions(receiptAdditions.map((addition) => addition.id === item.id ? { ...addition, name: event.target.value } : addition))} placeholder="Nome do valor adicional" className="h-11 rounded-lg border border-slate-200 px-3 text-sm font-semibold" />
                  <select value={item.mode} onChange={(event) => setReceiptAdditions(receiptAdditions.map((addition) => addition.id === item.id ? { ...addition, mode: event.target.value as "value" | "percent", value: "" } : addition))} className="h-11 rounded-lg border border-slate-200 px-3 text-sm"><option value="value">Valor em R$</option><option value="percent">% do bruto</option></select>
                  <input inputMode="decimal" value={item.value} onChange={(event) => setReceiptAdditions(receiptAdditions.map((addition) => addition.id === item.id ? { ...addition, value: item.mode === "value" ? formatMoneyInput(event.target.value) : event.target.value.replace(/[^0-9,.]/g, "").slice(0, 6) } : addition))} placeholder={item.mode === "value" ? "R$ 0,00" : "0,00%"} className="h-11 rounded-lg border border-slate-200 px-3 text-sm font-bold" />
                  <button type="button" onClick={() => setReceiptAdditions(receiptAdditions.filter((addition) => addition.id !== item.id))} className="h-11 rounded-lg px-3 text-xs font-bold text-red-500">Remover</button>
                </div>)}
                {!receiptAdditions.length && <p className="py-2 text-center text-xs text-slate-400">Nenhum valor adicional incluído.</p>}
              </div>
            </div>
            <div className="sm:col-span-2 rounded-2xl border border-slate-200 bg-slate-50 p-4">
              <div className="mb-3 flex items-center justify-between gap-3"><div><b className="text-sm text-slate-800">Descontos</b><p className="text-xs text-slate-500">Informe um valor fixo ou uma porcentagem sobre o salário bruto.</p></div><button type="button" onClick={() => setReceiptDiscounts([...receiptDiscounts, { id: Date.now(), name: "", mode: "value", value: "" }])} className="flex shrink-0 items-center gap-1 rounded-xl bg-slate-900 px-3 py-2 text-xs font-bold text-white"><Plus size={15} />Incluir desconto</button></div>
              <div className="space-y-3">
                {receiptDiscounts.map((item) => <div key={item.id} className="grid gap-2 rounded-xl border border-slate-200 bg-white p-3 sm:grid-cols-[1fr_150px_180px_auto]">
                  <input value={item.name} onChange={(event) => setReceiptDiscounts(receiptDiscounts.map((discount) => discount.id === item.id ? { ...discount, name: event.target.value } : discount))} placeholder="Nome do desconto" className="h-11 rounded-lg border border-slate-200 px-3 text-sm font-semibold" />
                  <select value={item.mode} onChange={(event) => setReceiptDiscounts(receiptDiscounts.map((discount) => discount.id === item.id ? { ...discount, mode: event.target.value as "value" | "percent", value: "" } : discount))} className="h-11 rounded-lg border border-slate-200 px-3 text-sm"><option value="value">Valor em R$</option><option value="percent">% do bruto</option></select>
                  <input inputMode="decimal" value={item.value} onChange={(event) => setReceiptDiscounts(receiptDiscounts.map((discount) => discount.id === item.id ? { ...discount, value: item.mode === "value" ? formatMoneyInput(event.target.value) : event.target.value.replace(/[^0-9,.]/g, "").slice(0, 6) } : discount))} placeholder={item.mode === "value" ? "R$ 0,00" : "0,00%"} className="h-11 rounded-lg border border-slate-200 px-3 text-sm font-bold" />
                  <button type="button" onClick={() => setReceiptDiscounts(receiptDiscounts.filter((discount) => discount.id !== item.id))} className="h-11 rounded-lg px-3 text-xs font-bold text-red-500">Remover</button>
                </div>)}
                {!receiptDiscounts.length && <p className="py-2 text-center text-xs text-slate-400">Nenhum desconto incluído.</p>}
                <div className="grid gap-3 rounded-xl border border-violet-200 bg-violet-50 p-3 sm:grid-cols-[1fr_220px] sm:items-center">
                  <div><b className="text-sm text-violet-900">Adiantamento já pago</b><p className="text-xs text-violet-700">Será descontado do salário e só aparecerá no recibo quando houver valor.</p></div>
                  <input inputMode="decimal" value={receiptAdvance} onChange={(event) => setReceiptAdvance(formatMoneyInput(event.target.value))} placeholder="R$ 0,00" className="h-11 rounded-lg border border-violet-200 bg-white px-3 text-sm font-black" />
                </div>
              </div>
            </div>
          </>}
          {receiptKind === "advance" && <><div className="sm:col-span-2 mt-2 border-t border-slate-200 pt-5"><b className="text-sm text-slate-900">3. Valores</b><p className="text-xs text-slate-500">Informe o valor do adiantamento.</p></div><label className="sm:col-span-2"><span className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">Valor do adiantamento</span><input inputMode="decimal" value={receiptAdvance} onChange={(event) => setReceiptAdvance(formatMoneyInput(event.target.value))} placeholder="R$ 0,00" className="h-14 w-full rounded-xl border border-slate-300 bg-white px-4 text-lg font-black shadow-sm" /></label></>}
          {receiptKind === "salary" && <div className="sm:col-span-2 rounded-xl border border-slate-200 bg-slate-50 p-4"><span className="text-sm text-slate-500">Valor líquido do recibo</span><b className="mt-1 block text-xl">{Math.max(0, proportionalSalary + receiptAdditionTotal - receiptDiscountTotal - parseMoney(receiptAdvance)).toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</b><span className="mt-1 block text-xs text-slate-500">Adicionais: {receiptAdditionTotal.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })} · Descontos: {receiptDiscountTotal.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</span></div>}
          <div className="sm:col-span-2 flex justify-end border-t border-slate-200 pt-5"><button type="button" onClick={generateReceipt} className="flex w-full items-center justify-center gap-2 rounded-xl bg-slate-950 px-6 py-3.5 font-black text-white shadow-lg transition hover:-translate-y-0.5 hover:bg-slate-800 sm:w-auto"><Download size={19} />Gerar recibo de {receiptKind === "salary" ? "salário" : "adiantamento"}</button></div>
        </div>
      </section>}
      {page === "Recibo" && <section className="administrative-document mx-auto max-w-6xl overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="border-b border-slate-200 bg-gradient-to-r from-slate-950 to-slate-800 px-6 py-5 text-white"><div className="flex items-center gap-3"><span className="grid h-11 w-11 place-items-center rounded-xl bg-white/10"><ReceiptText size={21} /></span><div><h2 className="text-xl font-black">Recibo simples</h2><p className="text-sm text-slate-300">Informe recebedor, pagamento e referência.</p></div></div></div>
        <div className="grid gap-x-5 gap-y-4 p-6 sm:grid-cols-2">
          <div className="sm:col-span-2 flex flex-col gap-3 border-b border-slate-200 pb-5 sm:flex-row sm:items-center sm:justify-between"><div><b className="text-sm text-slate-900">1. Identificação</b><p className="text-xs text-slate-500">Escolha um funcionário ou preencha os dados manualmente.</p></div><div className="administrative-segmented grid min-w-0 grid-cols-2 rounded-xl bg-slate-100 p-1 sm:min-w-[390px]"><button type="button" onClick={() => setReceiptPersonMode("registered")} className={`rounded-lg px-4 py-2.5 text-sm font-bold transition ${receiptPersonMode === "registered" ? "administrative-tab-active bg-white text-slate-950 shadow-sm" : "administrative-tab-inactive text-slate-500"}`}>Funcionário cadastrado</button><button type="button" onClick={() => setReceiptPersonMode("manual")} className={`rounded-lg px-4 py-2.5 text-sm font-bold transition ${receiptPersonMode === "manual" ? "administrative-tab-active bg-white text-slate-950 shadow-sm" : "administrative-tab-inactive text-slate-500"}`}>Preencher manualmente</button></div></div>
          {receiptPersonMode === "registered" ? <label className="sm:col-span-2"><span className="mb-2 block text-sm font-bold text-slate-700">Funcionário</span><SearchableEmployeeSelect employees={activeEmployees} value={receiptEmployeeId} onChange={setReceiptEmployeeId} /></label> : <><label><span className="mb-2 block text-sm font-bold text-slate-700">Nome completo</span><input value={receiptManualName} onChange={(event) => setReceiptManualName(event.target.value)} placeholder="Digite o nome completo" className="h-12 w-full rounded-xl border border-slate-200 px-4" /></label><label><span className="mb-2 block text-sm font-bold text-slate-700">CPF</span><input inputMode="numeric" maxLength={14} value={receiptManualCpf} onChange={(event) => setReceiptManualCpf(formatCpf(event.target.value))} placeholder="000.000.000-00" className="h-12 w-full rounded-xl border border-slate-200 px-4" /></label></>}
          <div className="sm:col-span-2 mt-2 border-t border-slate-200 pt-5"><b className="text-sm text-slate-900">2. Documento</b><p className="text-xs text-slate-500">Defina a empresa e a data do recibo.</p></div>
          <label><span className="mb-2 block text-sm font-bold text-slate-700">Loja / empresa</span><select value={receiptCompany} onChange={(event) => setReceiptCompany(event.target.value)} className="h-12 w-full rounded-xl border border-slate-300 bg-white px-4 shadow-sm"><option value="">Selecione a loja</option>{companies.map((company) => <option key={company} value={company}>{company}{companyCnpjs[company] ? ` - ${companyCnpjs[company]}` : ""}</option>)}</select></label>
          <label><span className="mb-2 block text-sm font-bold text-slate-700">Data do recibo</span><input type="date" value={receiptDate} onChange={(event) => setReceiptDate(event.target.value)} className="h-12 w-full rounded-xl border border-slate-200 px-4" /></label>
          <div className="sm:col-span-2 mt-2 border-t border-slate-200 pt-5"><b className="text-sm text-slate-900">3. Pagamento</b><p className="text-xs text-slate-500">Informe o valor e o motivo do pagamento.</p></div>
          <label><span className="mb-2 block text-sm font-bold text-slate-700">Valor</span><input inputMode="decimal" value={genericAmount} onChange={(event) => setGenericAmount(formatMoneyInput(event.target.value))} placeholder="0,00" className="h-12 w-full rounded-xl border border-slate-300 px-4 font-bold shadow-sm" /></label>
          <label className="sm:col-span-2"><span className="mb-2 block text-sm font-bold text-slate-700">Referente a</span><div className="flex flex-col gap-3 sm:flex-row"><input value={genericReference} onChange={(event) => setGenericReference(event.target.value)} placeholder="Ex.: vale, prestação de serviço ou reembolso" className="h-12 flex-1 rounded-xl border border-slate-200 px-4" /><button type="button" disabled={improvingGenericReference} onClick={() => void improveAdministrativeText(genericReference, "receipt_reference", setGenericReference, setImprovingGenericReference)} className="flex h-12 items-center justify-center gap-2 rounded-xl bg-slate-900 px-5 font-bold text-white hover:bg-slate-700 disabled:cursor-wait disabled:opacity-60"><Sparkles size={18} />{improvingGenericReference ? "Melhorando..." : "Melhorar com IA"}</button></div><span className="mt-2 block text-xs text-slate-500">A IA transforma a descrição em uma referência mais natural para o recibo.</span></label>
          <div className="sm:col-span-2 flex justify-end border-t border-slate-200 pt-5"><button type="button" onClick={generateGenericReceipt} className="flex w-full items-center justify-center gap-2 rounded-xl bg-slate-950 px-6 py-3.5 font-black text-white shadow-lg transition hover:-translate-y-0.5 hover:bg-slate-800 sm:w-auto"><Download size={19} />Gerar recibo em PDF</button></div>
        </div>
      </section>}
      {page === "Advertência" && <section className="administrative-document mx-auto max-w-6xl overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
        <div className="border-b border-slate-200 bg-gradient-to-r from-slate-950 to-slate-800 px-6 py-5 text-white">
          <div className="flex items-center gap-3">
            <span className="grid h-11 w-11 place-items-center rounded-xl bg-white/10 text-white">
              <TriangleAlert size={21} />
            </span>
            <div>
              <h2 className="text-xl font-black">Advertência disciplinar</h2>
              <p className="text-sm text-slate-300">
                Selecione o funcionário, informe as datas e descreva o ocorrido.
              </p>
            </div>
          </div>
        </div>
        <div className="grid gap-x-5 gap-y-4 p-6 sm:grid-cols-2">
          <div className="sm:col-span-2 flex flex-col gap-3 border-b border-slate-200 pb-5 sm:flex-row sm:items-center sm:justify-between">
            <div><b className="text-sm text-slate-900">1. Identificação</b><p className="text-xs text-slate-500">Escolha um funcionário ou preencha os dados manualmente.</p></div>
            <div className="administrative-segmented grid min-w-0 grid-cols-2 rounded-xl bg-slate-100 p-1 sm:min-w-[390px]">
            <button type="button" onClick={() => setWarningPersonMode("registered")} className={`rounded-lg px-4 py-2.5 text-sm font-bold transition ${warningPersonMode === "registered" ? "administrative-tab-active bg-white text-slate-950 shadow-sm" : "administrative-tab-inactive text-slate-500"}`}>Funcionário cadastrado</button>
            <button type="button" onClick={() => setWarningPersonMode("manual")} className={`rounded-lg px-4 py-2.5 text-sm font-bold transition ${warningPersonMode === "manual" ? "administrative-tab-active bg-white text-slate-950 shadow-sm" : "administrative-tab-inactive text-slate-500"}`}>Preencher manualmente</button>
            </div>
          </div>
          {warningPersonMode === "registered" ? (
            <label className="sm:col-span-2">
              <span className="mb-2 block text-sm font-bold text-slate-700">Funcionário</span>
              <SearchableEmployeeSelect employees={activeEmployees} value={employeeId} onChange={setEmployeeId} />
            </label>
          ) : (
            <>
              <label><span className="mb-2 block text-sm font-bold text-slate-700">Nome completo</span><input value={warningManualName} onChange={(event) => setWarningManualName(event.target.value)} placeholder="Digite o nome completo" className="h-12 w-full rounded-xl border border-slate-200 px-4" /></label>
              <label><span className="mb-2 block text-sm font-bold text-slate-700">CPF</span><input inputMode="numeric" maxLength={14} value={warningManualCpf} onChange={(event) => setWarningManualCpf(formatCpf(event.target.value))} placeholder="000.000.000-00" className="h-12 w-full rounded-xl border border-slate-200 px-4" /></label>
              <label className="sm:col-span-2"><span className="mb-2 block text-sm font-bold text-slate-700">Sexo</span><select value={warningManualGender} onChange={(event) => setWarningManualGender(event.target.value as Recharge["gender"] | "")} className="h-12 w-full rounded-xl border border-slate-200 bg-white px-4"><option value="">Selecione</option><option value="Masculino">Masculino</option><option value="Feminino">Feminino</option></select></label>
            </>
          )}
          <div className="sm:col-span-2 mt-2 border-t border-slate-200 pt-5"><b className="text-sm text-slate-900">2. Ocorrência</b><p className="text-xs text-slate-500">Informe as datas e descreva objetivamente o que aconteceu.</p></div>
          <label>
            <span className="mb-2 block text-sm font-bold text-slate-700">Data da ocorrência (aplicação)</span>
            <input
              type="date"
              value={documentDate}
              onChange={(event) => setDocumentDate(event.target.value)}
              className="h-12 w-full rounded-xl border border-slate-300 bg-white px-4 outline-none shadow-sm focus:border-slate-500"
            />
            <span className="mt-1.5 block text-xs text-slate-500">Dia em que a advertência foi aplicada.</span>
          </label>
          <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
            <div className="mb-2 flex items-center justify-between gap-3"><div><b className="text-sm text-slate-700">Data(s) do ocorrido</b><p className="text-xs text-slate-500">Dias em que o fato aconteceu.</p></div><button type="button" onClick={() => setOccurredDates([...occurredDates, ""])} className="inline-flex h-9 shrink-0 items-center gap-1 rounded-lg bg-slate-900 px-3 text-xs font-bold text-white"><Plus size={14} />Adicionar dia</button></div>
            <div className="space-y-2">{occurredDates.map((date, index) => <div key={index} className="flex gap-2"><input type="date" value={date} onChange={(event) => setOccurredDates(occurredDates.map((item, itemIndex) => itemIndex === index ? event.target.value : item))} className="h-11 min-w-0 flex-1 rounded-lg border border-slate-300 bg-white px-3 font-semibold outline-none focus:border-slate-500" /><button type="button" disabled={occurredDates.length === 1} onClick={() => setOccurredDates(occurredDates.filter((_, itemIndex) => itemIndex !== index))} aria-label="Remover data" className="grid h-11 w-11 place-items-center rounded-lg border border-slate-200 bg-white text-red-500 disabled:cursor-not-allowed disabled:opacity-30"><X size={17} /></button></div>)}</div>
          </div>
          <label className="sm:col-span-2">
            <span className="mb-2 block text-sm font-bold text-slate-700">Conte resumidamente o que aconteceu</span>
            <div>
              <textarea
                rows={3}
                value={warningSummary}
                onChange={(event) => setWarningSummary(event.target.value)}
                placeholder="Ex.: o funcionário faltou ao trabalho no dia informado e não apresentou justificativa"
                className="min-h-24 w-full rounded-xl border border-slate-200 px-4 py-3 outline-none focus:border-slate-500"
              />
              <div className="mt-2 flex justify-end"><button
                  type="button"
                  onClick={generateReasonWithAI}
                  disabled={generatingReason}
                  className="inline-flex h-9 items-center justify-center gap-2 rounded-lg bg-slate-900 px-3 text-xs font-bold text-white shadow-sm hover:bg-slate-700 disabled:cursor-wait disabled:opacity-60"
                >
                  <Sparkles size={15} />
                  {generatingReason ? "Criando..." : "Criar com IA"}
                </button></div>
            </div>
            <span className="mt-2 block text-xs text-slate-500">A IA apenas organiza o texto informado. Revise o conteúdo antes de gerar o documento.</span>
          </label>
          <div className="sm:col-span-2 mt-2 border-t border-slate-200 pt-5"><b className="text-sm text-slate-900">3. Redação final</b><p className="text-xs text-slate-500">Confira e ajuste o texto que será impresso.</p></div>
          <label className="sm:col-span-2">
            <span className="mb-2 block text-sm font-bold text-slate-700">Motivo da advertência</span>
            <div>
              <textarea
                rows={4}
                value={reason}
                onChange={(event) => setReason(event.target.value)}
                placeholder="Descreva de forma objetiva o motivo da advertência"
                className="min-h-28 w-full rounded-xl border border-slate-200 px-4 py-3 outline-none focus:border-slate-500"
              />
              <div className="mt-2 flex justify-end"><button type="button" disabled={improvingWarningReason} onClick={() => void improveAdministrativeText(reason, "warning_reason", setReason, setImprovingWarningReason)} className="inline-flex h-9 items-center justify-center gap-2 rounded-lg bg-slate-900 px-3 text-xs font-bold text-white shadow-sm hover:bg-slate-700 disabled:cursor-wait disabled:opacity-60"><Sparkles size={15} />{improvingWarningReason ? "Melhorando..." : "Melhorar com IA"}</button></div>
            </div>
          </label>
          {warningPerson && (
            <div className="sm:col-span-2 rounded-xl border border-slate-200 bg-slate-50 p-4 text-sm">
              <b>{warningPerson.employee}</b>
              <span className="ml-2 text-slate-500">CPF: {warningPerson.cpf || "não informado"}</span>
            </div>
          )}
          <div className="sm:col-span-2 flex justify-end border-t border-slate-200 pt-5">
            <button
              type="button"
              onClick={generateWarning}
              className="flex w-full items-center justify-center gap-2 rounded-xl bg-slate-950 px-6 py-3.5 font-black text-white shadow-lg transition hover:-translate-y-0.5 hover:bg-slate-800 sm:w-auto"
            >
              <Download size={18} />
              Gerar advertência em PDF
            </button>
          </div>
        </div>
      </section>}
    </main>
  );
}

export default function App() {
  usePopupBackgroundLock();
  const [needsSetup, setNeedsSetup] = useState(false),
    [authChecked, setAuthChecked] = useState(() => !cloudEnabled()),
    [sessionChecked, setSessionChecked] = useState(() => !cloudEnabled());
  const now = new Date();
  const [dark, setDark] = useState(
    () => localStorage.getItem("valefluxo_theme") === "dark",
  );
  const [loggedIn, setLoggedIn] = useState(() =>
    cloudEnabled() ? false : localStorage.getItem("valefluxo_session") !== "0",
  );
  const [sessionUser, setSessionUser] = useState<SessionUser | null>(() =>
    cloudEnabled()
      ? null
      : { id: 0, username: "local", fullName: "João Fonseca", role: "admin", modules: ["people", "finance", "transit", "administrative"], storeAccess: ["*"] },
  );
  const [module, setModule] = useState<Module | null>(() => {
    const saved = localStorage.getItem("abc_current_module");
    return saved === "people" || saved === "transit" || saved === "finance" || saved === "administrative"
      ? saved
      : null;
  });
  const [side, setSide] = useState(false);
  const [page, setPage] = useState(
    () => localStorage.getItem("abc_current_page") || "Visão geral",
  );
  const [modal, setModal] = useState(false);
  const [editing, setEditing] = useState<Recharge | null>(null);
  const [marking, setMarking] = useState<Recharge | null>(null);
  const [period, setPeriod] = useState(
    `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`,
  );
  const [referenceDay, setReferenceDay] = useState(now.getDate());
  const [selectedStore, setSelectedStore] = useState("Todas");
  const [selectedRole, setSelectedRole] = useState("Todas");
  const initialDb = useMemo(() => loadDatabase(), []);
  const [rows, setRows] = useState<Recharge[]>(initialDb.employees);
  const [events, setEvents] = useState<RechargeEvent[]>(initialDb.events);
  const [stores, setStores] = useState<string[]>(() => {
    try {
      const savedStores = JSON.parse(localStorage.getItem("valefluxo_stores_v2") || "[]");
      const legacyCompanies = JSON.parse(localStorage.getItem("abc_tax_companies") || "[]");
      return [...new Set([...(Array.isArray(savedStores) ? savedStores : []), ...(Array.isArray(legacyCompanies) ? legacyCompanies : [])])];
    } catch {
      return [];
    }
  });
  const [positions, setPositions] = useState<string[]>(() => {
    try {
      return JSON.parse(localStorage.getItem("valefluxo_positions") || "[]");
    } catch {
      return [];
    }
  });
  const [unregisteredReasons, setUnregisteredReasons] = useState<string[]>(
    () => {
      try {
        return JSON.parse(
          localStorage.getItem("valefluxo_unregistered_reasons") || "[]",
        );
      } catch {
        return [];
      }
    },
  );
  const [occurrences, setOccurrences] = useState<HROccurrence[]>(() => {
    try {
      return JSON.parse(
        localStorage.getItem("valefluxo_hr_occurrences") || "[]",
      );
    } catch {
      return [];
    }
  });
  const [financialEntries, setFinancialEntries] = useState<FinancialEntry[]>(() => {
    try {
      return JSON.parse(localStorage.getItem("abc_financial_entries") || "[]");
    } catch {
      return [];
    }
  });
  const [taxEntries, setTaxEntries] = useState<TaxEntry[]>(() => {
    try {
      return JSON.parse(localStorage.getItem("abc_tax_entries") || "[]");
    } catch {
      return [];
    }
  });
  const [companyCnpjs, setCompanyCnpjs] = useState<Record<string, string>>(() => {
    try {
      return JSON.parse(localStorage.getItem("abc_company_cnpjs") || "{}");
    } catch {
      return {};
    }
  });
  const [manualMode, setManualMode] = useState(false);
  const [cloudReady, setCloudReady] = useState(false);
  const [cloudSaveError, setCloudSaveError] = useState("");
  const cloudRevision = useRef(0);
  const cloudSaveQueue = useRef<Promise<void>>(Promise.resolve());
  const genderInferenceStarted = useRef(false);
  useEffect(() => {
    localStorage.setItem(
      "valefluxo_db_v3",
      JSON.stringify({ employees: rows, events }),
    );
  }, [rows, events]);
  useEffect(() => {
    localStorage.setItem("valefluxo_stores_v2", JSON.stringify(stores));
    localStorage.setItem("abc_tax_companies", JSON.stringify(stores));
  }, [stores]);
  useEffect(() => {
    localStorage.setItem("valefluxo_positions", JSON.stringify(positions));
  }, [positions]);
  useEffect(() => {
    localStorage.setItem(
      "valefluxo_unregistered_reasons",
      JSON.stringify(unregisteredReasons),
    );
  }, [unregisteredReasons]);
  useEffect(() => {
    localStorage.setItem(
      "valefluxo_hr_occurrences",
      JSON.stringify(occurrences),
    );
  }, [occurrences]);
  useEffect(() => {
    localStorage.setItem("abc_financial_entries", JSON.stringify(financialEntries));
  }, [financialEntries]);
  useEffect(() => {
    localStorage.setItem("abc_tax_entries", JSON.stringify(taxEntries));
  }, [taxEntries]);
  useEffect(() => {
    localStorage.setItem("abc_company_cnpjs", JSON.stringify(companyCnpjs));
  }, [companyCnpjs]);
  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
    localStorage.setItem("valefluxo_theme", dark ? "dark" : "light");
  }, [dark]);
  useEffect(() => {
    if (module) {
      localStorage.setItem("abc_current_module", module);
      localStorage.setItem("abc_current_page", page);
    }
  }, [module, page]);
  useEffect(() => {
    let active = true;
    if (!cloudEnabled()) {
      setCloudReady(true);
      return () => {
        active = false;
      };
    }
    loadCloudState()
      .then((state) => {
        if (!active || !state) return;
        cloudRevision.current = Number(state.revision || 0);
        const hasRemoteData =
          state.employees.length ||
          state.events.length ||
          state.occurrences.length ||
          state.stores.length ||
          state.positions.length;
        if (hasRemoteData) {
          setRows(state.employees);
          setEvents(state.events);
          setOccurrences(state.occurrences as HROccurrence[]);
          const remoteCompanies = Array.isArray(state.settings?.taxCompanies)
            ? (state.settings.taxCompanies as string[])
            : [];
          const unifiedStores = [...new Set([...state.stores, ...remoteCompanies])];
          setStores(unifiedStores);
          setPositions(state.positions);
          setUnregisteredReasons(state.unregisteredReasons);
          setFinancialEntries(
            Array.isArray(state.settings?.financialEntries)
              ? (state.settings.financialEntries as FinancialEntry[])
              : [],
          );
          setTaxEntries(
            Array.isArray(state.settings?.taxEntries)
              ? (state.settings.taxEntries as TaxEntry[])
              : [],
          );
          setCompanyCnpjs(
            state.settings?.companyCnpjs && typeof state.settings.companyCnpjs === "object"
              ? (state.settings.companyCnpjs as Record<string, string>)
              : {},
          );
        }
        setCloudReady(true);
      })
      .catch((error) => {
        console.error("Falha ao conectar ao D1", error);
        setCloudReady(false);
      });
    return () => {
      active = false;
    };
  }, []);
  useEffect(() => {
    if (!cloudReady || !cloudEnabled()) return;
    const timer = setTimeout(() => {
      const snapshot = {
        employees: rows,
        events,
        occurrences,
        stores,
        positions,
        unregisteredReasons,
        settings: {
          advanceDays: Number(localStorage.getItem("valefluxo_advance") || 3),
          financialEntries,
          taxEntries,
          taxCompanies: stores,
          companyCnpjs,
        },
      };
      cloudSaveQueue.current = cloudSaveQueue.current.then(async () => {
        try {
          const result = await saveCloudState({
            ...snapshot,
            revision: cloudRevision.current,
          });
          cloudRevision.current = result.revision;
          setCloudSaveError("");
        } catch (error) {
          console.error("Falha ao salvar no D1", error);
          const message =
            error instanceof Error
              ? error.message
              : "Não foi possível salvar no banco.";
          setCloudSaveError(message);
        }
      });
    }, 700);
    return () => clearTimeout(timer);
  }, [
    cloudReady,
    rows,
    events,
    occurrences,
    stores,
    positions,
    unregisteredReasons,
    financialEntries,
    taxEntries,
    companyCnpjs,
  ]);
  useEffect(() => {
    if (!cloudReady || !cloudEnabled() || sessionUser?.role !== "admin" || genderInferenceStarted.current) return;
    const missingGender = rows.filter((employee) => !employee.gender);
    if (!missingGender.length) return;
    genderInferenceStarted.current = true;
    fetch("/api/infer-genders", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ people: missingGender.map((employee) => ({ id: employee.id, name: employee.employee, role: employee.role })) }),
    })
      .then(async (response) => {
        const data = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(data.error || "Não foi possível completar os sexos dos cadastros.");
        const inferred = new Map<number, Recharge["gender"]>(
          (Array.isArray(data.results) ? data.results : []).map((item: { id: number; gender: Recharge["gender"] }) => [Number(item.id), item.gender]),
        );
        if (!inferred.size) return;
        setRows((current) => current.map((employee) => employee.gender ? employee : { ...employee, gender: inferred.get(employee.id) || employee.gender }));
        window.dispatchEvent(new CustomEvent("abc:toast", { detail: `Sexo preenchido automaticamente em ${inferred.size} cadastro(s)` }));
      })
      .catch((error) => console.error("Falha ao completar o sexo dos cadastros com IA", error));
  }, [cloudReady, rows, sessionUser?.role]);
  useEffect(() => {
    if (!cloudReady || sessionUser?.role !== "admin") return;
    if (!rows.some((employee) => genderFromRole(employee.role) === "Feminino" && employee.gender !== "Feminino")) return;
    setRows((current) => current.map((employee) => genderFromRole(employee.role) === "Feminino" ? { ...employee, gender: "Feminino" } : employee));
    window.dispatchEvent(new CustomEvent("abc:toast", { detail: "Sexo corrigido conforme a função cadastrada" }));
  }, [cloudReady, rows, sessionUser?.role]);
  useEffect(() => {
    if (cloudEnabled())
      cloudCurrentUser()
        .then((user) => {
          setSessionUser(user);
          setLoggedIn(!!user);
          if (user?.storeAccess.length === 1 && user.storeAccess[0] !== "*") setSelectedStore(user.storeAccess[0]);
          if (user?.role === "operator" && page === "Configurações") {
            setPage(module === "finance" ? "Dashboard" : "Visão geral");
          }
          if (user && module && !user.modules.includes(module)) {
            localStorage.removeItem("abc_current_module");
            localStorage.removeItem("abc_current_page");
            setModule(null);
          }
        })
        .catch(() => {
          setSessionUser(null);
          setLoggedIn(false);
        })
        .finally(() => setSessionChecked(true));
  }, []);
  useEffect(() => {
    if (!cloudEnabled()) return;
    cloudSetupRequired()
      .then(setNeedsSetup)
      .catch(() => setNeedsSetup(false))
      .finally(() => setAuthChecked(true));
  }, []);
  const saveEmployee = (record: Recharge) =>
    setRows(
      editing
        ? rows.map((r) =>
            r.id === record.id
              ? {
                  ...record,
                  completedDate: undefined,
                  periodCompletionId: undefined,
                }
              : r,
          )
        : [
            {
              ...record,
              completedDate: undefined,
              periodCompletionId: undefined,
            },
            ...rows,
          ],
    );
  const completeRecharge = (
    record: Recharge,
    date: string,
    days: number,
    total: number,
  ) => {
    const eventId = record.periodCompletionId,
      daily1 = record.cardDailyFare ?? record.dailyFare ?? 0,
      daily2 = record.secondCardDailyFare ?? 0,
      dailyTotal = daily1 + daily2;
    const normalizedTotal = Math.round(total * 100) / 100,
      cardAmount = dailyTotal
        ? Math.round(normalizedTotal * (daily1 / dailyTotal) * 100) / 100
        : 0,
      secondCardAmount = record.secondCardType
        ? Math.round((normalizedTotal - cardAmount) * 100) / 100
        : undefined,
      eventPeriod = (record.completedDate || date).slice(0, 7);
    const event: RechargeEvent = {
      id: eventId || Date.now(),
      employeeId: record.id,
      period: eventPeriod,
      scheduledDate: record.rechargeDate,
      completedDate: date,
      chargedDays: days,
      cardAmount,
      secondCardAmount,
      totalAmount: normalizedTotal,
    };
    setEvents(
      eventId
        ? events.map((item) => (item.id === eventId ? event : item))
        : [...events, event],
    );
    setMarking(null);
    setManualMode(false);
  };
  const undoRecharge = (record: Recharge) => {
    if (!confirm("Desfazer a recarga deste mês?")) return;
    const eventId = record.periodCompletionId;
    setEvents(
      events.filter(
        (event) =>
          event.id !== eventId &&
          !(event.employeeId === record.id && event.period === period),
      ),
    );
  };
  const [refYear, refMonth] = period.split("-").map(Number),
    maxReferenceDay = new Date(refYear, refMonth, 0).getDate(),
    safeReferenceDay = Math.min(referenceDay, maxReferenceDay),
    referenceDate = `${period}-${String(safeReferenceDay).padStart(2, "0")}`;
  const storeScope = sessionUser?.storeAccess || ["*"],
    canAccessStore = (store: string) => storeScope.includes("*") || storeScope.includes(store),
    accessibleRows = rows.filter((record) => canAccessStore(record.store)),
    accessibleStores = stores.filter(canAccessStore);
  const availableRoles = useMemo(
    () =>
      [...new Set(accessibleRows.map((record) => record.role).filter(Boolean))].sort((a, b) =>
        a.localeCompare(b, "pt-BR"),
      ),
    [accessibleRows],
  );
  const viewRows = useMemo(
    () =>
      monthlyRows(accessibleRows, events, period, referenceDate).filter(
        (r) =>
          (selectedStore === "Todas" || r.store === selectedStore) &&
          (selectedRole === "Todas" || r.role === selectedRole),
      ),
    [accessibleRows, events, period, referenceDate, selectedStore, selectedRole],
  );
  const rechargeAlertCount = viewRows.filter(
    (r) => r.status === "Atrasado" || r.status === "Pendente",
  ).length;
  const openNew = () => {
      setEditing(null);
      setModal(true);
    },
    openEdit = (r: Recharge) => {
      setEditing(rows.find((x) => x.id === r.id) || r);
      setModal(true);
    };
  const removeEmployee = (id: number) => {
    setRows(rows.filter((r) => r.id !== id));
    setEvents(events.filter((event) => event.employeeId !== id));
  };
  const operationalRows = accessibleRows.filter((r) => !isEmployeeDismissed(r));
  const filteredEmployees = operationalRows.filter(
    (record) =>
      (selectedStore === "Todas" || record.store === selectedStore) &&
      (selectedRole === "Todas" || record.role === selectedRole),
  );
  const [financialYear, financialMonth] = period.split("-").map(Number),
    financialReference = new Date(financialYear, financialMonth - 1, 1, 12),
    financeEligibleEmployees = accessibleRows.filter((record) => {
      const matchesFilters =
        (selectedStore === "Todas" || record.store === selectedStore) &&
        (selectedRole === "Todas" || record.role === selectedRole);
      if (!matchesFilters) return false;
      if (!isEmployeeDismissed(record, financialReference)) return true;
      if (!record.terminationDate) return false;
      const termination = new Date(record.terminationDate + "T12:00:00"),
        terminationMonth = new Date(
          termination.getFullYear(),
          termination.getMonth(),
          1,
          12,
        ),
        visibleUntil = new Date(
          termination.getFullYear(),
          termination.getMonth() + 3,
          1,
          12,
        );
      return financialReference >= terminationMonth && financialReference < visibleUntil;
    }),
    financeEmployees = financeEligibleEmployees,
    financeReportEmployees = financeEligibleEmployees;
  const content =
    page === "Configurações" && sessionUser?.role === "admin" ? (
      <ConfigurationsPage
        positions={positions}
        setPositions={setPositions}
        unregisteredReasons={unregisteredReasons}
        setUnregisteredReasons={setUnregisteredReasons}
        stores={stores}
        setStores={setStores}
        rows={rows}
        setRows={setRows}
        financialEntries={financialEntries}
        setFinancialEntries={setFinancialEntries}
        financialPeriod={period}
        setFinancialPeriod={setPeriod}
        openEmployeeForm={openNew}
        editEmployee={openEdit}
        companyCnpjs={companyCnpjs}
        setCompanyCnpjs={setCompanyCnpjs}
      />
    ) : module === "administrative" ? (
      <AdministrativePage page={administrativeNav.some(([label]) => label === page) ? page : "Salário"} employees={filteredEmployees} companies={stores} companyCnpjs={companyCnpjs} financialEntries={financialEntries} />
    ) : module === "finance" ? (
      page === "Relatórios" ? (
        <FinancialReports
          employees={financeReportEmployees}
          entries={financialEntries}
          period={period}
          setPeriod={setPeriod}
        />
      ) : page === "Cadastros" ? (
        <FinancialRegistrations
          employees={financeEligibleEmployees}
          entries={financialEntries}
          setEntries={setFinancialEntries}
          period={period}
          setPeriod={setPeriod}
        />
      ) : page === "Bonificação" ? (
        <BonusPage
          employees={accessibleRows.filter(
            (record) =>
              (selectedStore === "Todas" || record.store === selectedStore) &&
              (selectedRole === "Todas" || record.role === selectedRole),
          )}
          occurrences={occurrences}
          period={period}
          setPeriod={setPeriod}
        />
      ) : page === "Impostos" ? (
        <TaxesPage
          entries={taxEntries}
          setEntries={setTaxEntries}
          companies={stores}
          period={period}
          setPeriod={setPeriod}
        />
      ) : (
        <FinancePage
          employees={financeEmployees}
          entries={financialEntries}
          setEntries={setFinancialEntries}
          period={period}
          setPeriod={setPeriod}
        />
      )
    ) : module === "people" ? (
      page === "Visão geral" ? (
        <HRPage
          rows={accessibleRows.filter(
            (record) =>
              (selectedStore === "Todas" || record.store === selectedStore) &&
              (selectedRole === "Todas" || record.role === selectedRole),
          )}
          occurrences={occurrences}
          onEdit={openEdit}
        />
      ) : page === "Funcionários" ? (
        <HREmployeesPage
          rows={filteredEmployees}
          readOnly
          toggleCritical={
            sessionUser?.role === "admin"
              ? (record) =>
                  setRows(
                    rows.map((item) =>
                      item.id === record.id
                        ? {
                            ...item,
                            experienceCritical: !item.experienceCritical,
                          }
                        : item,
                    ),
                  )
              : undefined
          }
          manage={
            sessionUser?.role === "admin"
              ? (employee) => {
                  localStorage.setItem("abc_settings_tab", "funcionarios");
                  setPage("Configurações");
                  openEdit(employee);
                }
              : undefined
          }
        />
      ) : page === "Ocorrências" ? (
        <OccurrencesPage
          employees={filteredEmployees}
          items={occurrences}
          setItems={setOccurrences}
          readOnly={sessionUser?.role === "operator"}
        />
      ) : (
        <HRReports
          employees={accessibleRows.filter(
            (record) =>
              (selectedStore === "Todas" || record.store === selectedStore) &&
              (selectedRole === "Todas" || record.role === selectedRole),
          )}
          occurrences={occurrences}
        />
      )
    ) : page === "Funcionários" ? (
      <EmployeesPage
        rows={filteredEmployees}
        openForm={openNew}
        edit={openEdit}
        remove={removeEmployee}
      />
    ) : page === "Visão geral" ? (
      <Dashboard
        rows={viewRows}
        setRows={setRows}
        openForm={openNew}
        onMark={(r) => {
          setManualMode(false);
          setMarking(r);
        }}
        referenceDate={referenceDate}
      />
    ) : page === "Recargas" ? (
      <RechargeHistoryPage
        employees={filteredEmployees}
        events={events}
        planned={viewRows}
        onMark={(r) => {
          setManualMode(false);
          setMarking(r);
        }}
        onManual={(r) => {
          setManualMode(true);
          setMarking(r);
        }}
        onEdit={(employee, event) => {
          setManualMode(true);
          setMarking({
            ...employee,
            completedDate: event.completedDate,
            chargedDays: event.chargedDays,
            cardAmount: event.cardAmount,
            secondCardAmount: event.secondCardAmount,
            amount: event.totalAmount,
            periodCompletionId: event.id,
            rechargeDate: event.scheduledDate,
            status: "Recarregado",
          });
        }}
        onDelete={(event) => {
          if (confirm("Cancelar esta recarga do histórico?"))
            setEvents(events.filter((item) => item.id !== event.id));
        }}
      />
    ) : page === "Calendário" ? (
      <CalendarPage rows={viewRows} period={period} setPeriod={setPeriod} />
    ) : (
      <ReportsPageFiltered rows={viewRows} />
    );
  const logout = () => {
    void cloudLogout();
    localStorage.setItem("valefluxo_session", "0");
    localStorage.removeItem("abc_current_module");
    localStorage.removeItem("abc_current_page");
    setLoggedIn(false);
    setSessionUser(null);
    setModule(null);
    setSide(false);
  };
  if (!authChecked || !sessionChecked) return <InitialLoadingScreen />;
  if (!loggedIn)
    return (
      <LoginScreen
        onLogin={() => {
          if (cloudEnabled()) location.reload();
          else {
            setLoggedIn(true);
            setModule(null);
          }
        }}
      />
    );
  if (!module)
    return (
      <ModuleMenu
        select={(choice) => {
          if (!sessionUser?.modules.includes(choice)) return;
          setModule(choice);
          setPage(choice === "finance" ? "Dashboard" : choice === "administrative" ? "Salário" : "Visão geral");
        }}
        onLogout={logout}
        dark={dark}
        toggleTheme={() => setDark(!dark)}
        allowedModules={sessionUser?.modules || []}
      />
    );
  return (
    <div className="min-h-screen bg-[#f5f7f6]">
      <Sidebar
        open={side}
        close={() => setSide(false)}
        page={page}
        setPage={setPage}
        module={module}
        rechargeAlertCount={rechargeAlertCount}
        onChangeModule={() => {
          localStorage.removeItem("abc_current_module");
          localStorage.removeItem("abc_current_page");
          setModule(null);
          setSide(false);
          setPage("Visão geral");
        }}
        onLogout={logout}
      />
      <div className="lg:pl-[246px]">
        <Header
          menu={() => setSide(true)}
          page={page}
          module={module}
          referenceDate={referenceDate}
          setReferenceDate={(value) => {
            if (value) {
              setPeriod(value.slice(0, 7));
              setReferenceDay(Number(value.slice(8, 10)));
            }
          }}
          alertCount={rechargeAlertCount}
          stores={accessibleStores}
          selectedStore={selectedStore}
          setSelectedStore={setSelectedStore}
          roles={availableRoles}
          selectedRole={selectedRole}
          setSelectedRole={setSelectedRole}
          dark={dark}
          toggleTheme={() => setDark(!dark)}
        />
        {cloudSaveError && (
          <div className="mx-4 mt-4 flex flex-wrap items-center gap-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800 sm:mx-7">
            <TriangleAlert size={18} className="shrink-0" />
            <b>Os dados ainda não foram salvos:</b>
            <span>{cloudSaveError}</span>
            <button
              type="button"
              onClick={() => location.reload()}
              className="ml-auto rounded-lg bg-red-700 px-3 py-2 text-xs font-bold text-white"
            >
              Recarregar dados
            </button>
          </div>
        )}
        {content}
      </div>
      {modal && (
        <EmployeeModal
          close={() => {
            setModal(false);
            setEditing(null);
          }}
          add={saveEmployee}
          stores={stores}
          positions={positions}
          unregisteredReasons={unregisteredReasons}
          initial={editing}
          rows={rows}
        />
      )}{" "}
      {marking &&
        (manualMode ? (
          <SmartRechargeModal
            record={marking}
            close={() => {
              setMarking(null);
              setManualMode(false);
            }}
            confirm={(date, days, total) =>
              completeRecharge(marking, date, days, total)
            }
          />
        ) : (
          <AutoRechargeModal
            record={marking}
            close={() => setMarking(null)}
            confirm={(date, days, total) =>
              completeRecharge(marking, date, days, total)
            }
          />
        ))}
      <ActionToast />
    </div>
  );
}
