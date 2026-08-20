import { useEffect, useState } from "react";
import { CreditCard, DollarSign, FileText, Pencil, ShieldCheck, UserPlus, Users } from "lucide-react";
import {
  createCloudUser,
  listCloudUsers,
  updateCloudUser,
  type CloudUser,
  type ModuleAccess,
} from "./cloud";

type UserForm = {
  id: number;
  fullName: string;
  username: string;
  password: string;
  role: "admin" | "operator";
  active: boolean;
  modules: ModuleAccess[];
  storeAccess: string[];
};

const empty: UserForm = {
  id: 0,
  fullName: "",
  username: "",
  password: "",
  role: "operator",
  active: true,
  modules: ["people"],
  storeAccess: ["*"],
};

const moduleOptions: Array<{
  value: ModuleAccess;
  label: string;
  description: string;
  icon: typeof Users;
}> = [
  { value: "people", label: "RH", description: "Pessoas, ocorrências e relatórios", icon: Users },
  { value: "finance", label: "Financeiro", description: "Folha, pagamentos e relatórios", icon: DollarSign },
  { value: "transit", label: "Cartões de passagem", description: "Recargas, calendário e transporte", icon: CreditCard },
  { value: "administrative", label: "Administrativo", description: "Recibos, advertências e documentos", icon: FileText },
];

export default function UserProfiles({ stores }: { stores: string[] }) {
  const [users, setUsers] = useState<CloudUser[]>([]);
  const [form, setForm] = useState<UserForm>(empty);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const load = () =>
    listCloudUsers()
      .then(setUsers)
      .catch((reason) => setError(reason.message))
      .finally(() => setLoading(false));
  useEffect(() => {
    void load();
  }, []);

  const toggleModule = (module: ModuleAccess) =>
    setForm((current) => ({
      ...current,
      modules: current.modules.includes(module)
        ? current.modules.filter((item) => item !== module)
        : [...current.modules, module],
    }));
  const toggleStore = (store: string) =>
    setForm((current) => {
      if (store === "*") return { ...current, storeAccess: ["*"] };
      const selected = current.storeAccess.filter((item) => item !== "*");
      return {
        ...current,
        storeAccess: selected.includes(store)
          ? selected.filter((item) => item !== store)
          : [...selected, store],
      };
    });

  const submit = async (event: React.FormEvent) => {
    event.preventDefault();
    setError("");
    if (!form.modules.length) {
      setError("Selecione pelo menos uma área para o usuário.");
      return;
    }
    if (!form.storeAccess.length) {
      setError("Selecione pelo menos uma loja para o usuário.");
      return;
    }
    setSaving(true);
    try {
      if (form.id) await updateCloudUser(form);
      else await createCloudUser(form);
      setForm(empty);
      await load();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "Não foi possível salvar.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <main className="fade-in p-4 sm:p-7">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-slate-900">Perfis e usuários</h2>
        <p className="text-sm text-slate-500">Controle o perfil e as áreas que cada usuário pode acessar.</p>
      </div>
      <div className="grid gap-5 xl:grid-cols-[460px_1fr]">
        <form onSubmit={submit} className="rounded-2xl border border-slate-200 bg-white p-6 shadow-soft">
          <div className="flex items-center gap-3">
            <div className="grid h-11 w-11 place-items-center rounded-xl bg-slate-100"><UserPlus size={20} /></div>
            <div><h3 className="font-bold">{form.id ? "Editar perfil" : "Novo perfil"}</h3><p className="text-xs text-slate-500">Senha mínima de 8 caracteres</p></div>
          </div>
          <div className="mt-5 space-y-4">
            <label className="block text-sm font-semibold">Nome completo<input required value={form.fullName} onChange={(e) => setForm({ ...form, fullName: e.target.value })} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5" /></label>
            <label className="block text-sm font-semibold">Usuário<input required value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5" /></label>
            <label className="block text-sm font-semibold">{form.id ? "Nova senha (opcional)" : "Senha"}<input required={!form.id} type="password" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5" /></label>
            <label className="block text-sm font-semibold">Permissão<select value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value as "admin" | "operator" })} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5"><option value="admin">Administrador</option><option value="operator">Operador</option></select></label>
            <fieldset>
              <legend className="text-sm font-bold text-slate-800">Lojas permitidas</legend>
              <p className="mt-1 text-xs text-slate-500">Marque uma, várias ou todas as lojas.</p>
              <div className="mt-3 grid gap-2 sm:grid-cols-2">
                {["*", ...stores].map((store) => {
                  const checked = form.storeAccess.includes(store);
                  return <label key={store} className={`flex cursor-pointer items-center gap-2 rounded-xl border px-3 py-2.5 text-sm font-semibold transition ${checked ? "border-blue-500 bg-blue-50 text-blue-800" : "border-slate-200 text-slate-600 hover:border-slate-300"}`}><input type="checkbox" checked={checked} onChange={() => toggleStore(store)} className="h-4 w-4 accent-blue-600" /><span>{store === "*" ? "Todas as lojas" : store}</span></label>;
                })}
              </div>
            </fieldset>
            <fieldset>
              <legend className="text-sm font-bold text-slate-800">Áreas permitidas</legend>
              <p className="mt-1 text-xs text-slate-500">Marque tudo que este usuário poderá acessar.</p>
              <div className="mt-3 space-y-2">
                {moduleOptions.map(({ value, label, description, icon: Icon }) => {
                  const checked = form.modules.includes(value);
                  return (
                    <label key={value} className={`flex cursor-pointer items-center gap-3 rounded-xl border p-3 transition ${checked ? "border-slate-700 bg-slate-50" : "border-slate-200 hover:border-slate-300"}`}>
                      <input type="checkbox" checked={checked} onChange={() => toggleModule(value)} className="h-4 w-4 accent-slate-800" />
                      <span className={`grid h-9 w-9 place-items-center rounded-lg ${checked ? "bg-slate-800 text-white" : "bg-slate-100 text-slate-500"}`}><Icon size={17} /></span>
                      <span><b className="block text-sm">{label}</b><span className="text-xs text-slate-500">{description}</span></span>
                    </label>
                  );
                })}
              </div>
            </fieldset>
          </div>
          {error && <p className="mt-4 rounded-xl bg-red-50 p-3 text-sm font-semibold text-red-600">{error}</p>}
          <div className="mt-5 flex gap-2"><button disabled={saving} className="flex-1 rounded-xl bg-slate-800 px-4 py-3 font-bold text-white disabled:opacity-50">{saving ? "Salvando..." : form.id ? "Salvar alterações" : "Cadastrar perfil"}</button>{!!form.id && <button type="button" onClick={() => setForm(empty)} className="rounded-xl border border-slate-200 px-4 font-semibold">Cancelar</button>}</div>
        </form>
        <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft">
          <div className="border-b border-slate-100 px-5 py-4"><h3 className="font-bold">Usuários cadastrados</h3></div>
          {loading ? <p className="p-8 text-center text-slate-400">Carregando...</p> : <div className="divide-y divide-slate-100">{users.map((user) => (
            <div key={user.id} className="flex flex-wrap items-center gap-3 px-5 py-4">
              <div className="grid h-10 w-10 place-items-center rounded-full bg-slate-100 font-bold">{user.fullName.split(" ").map((x) => x[0]).slice(0, 2).join("").toUpperCase()}</div>
              <div className="min-w-[180px] flex-1"><b>{user.fullName}</b><div className="text-xs text-slate-500">@{user.username}</div><div className="mt-2 flex flex-wrap gap-1">{moduleOptions.filter((item) => user.modules.includes(item.value)).map((item) => <span key={item.value} className="rounded-full bg-slate-100 px-2 py-1 text-[10px] font-bold text-slate-600">{item.label}</span>)}{user.storeAccess.map((store) => <span key={store} className="rounded-full bg-blue-50 px-2 py-1 text-[10px] font-bold text-blue-700">{store === "*" ? "Todas as lojas" : store}</span>)}</div></div>
              <span className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-3 py-1 text-xs font-bold"><ShieldCheck size={14} />{user.role === "admin" ? "Administrador" : "Operador"}</span>
              <span className={`rounded-full px-3 py-1 text-xs font-bold ${user.active ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-600"}`}>{user.active ? "Ativo" : "Inativo"}</span>
              <button onClick={() => setForm({ id: user.id, fullName: user.fullName, username: user.username, password: "", role: user.role, active: user.active, modules: user.modules, storeAccess: user.storeAccess })} className="rounded-lg border border-slate-200 p-2" title="Editar"><Pencil size={16} /></button>
              <button onClick={async () => { await updateCloudUser({ ...user, password: "", active: !user.active }); load(); }} className="rounded-lg border border-slate-200 px-3 py-2 text-xs font-bold">{user.active ? "Desativar" : "Ativar"}</button>
            </div>
          ))}</div>}
        </div>
      </div>
    </main>
  );
}
