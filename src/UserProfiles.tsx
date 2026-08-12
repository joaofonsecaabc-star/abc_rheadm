import { useEffect, useState } from 'react'
import { Pencil, ShieldCheck, UserPlus } from 'lucide-react'
import { createCloudUser, listCloudUsers, updateCloudUser, type CloudUser } from './cloud'

type UserForm = { id:number; fullName:string; username:string; password:string; role:'admin'|'operator'; active:boolean }
const empty:UserForm = { id: 0, fullName: '', username: '', password: '', role: 'operator', active: true }

export default function UserProfiles() {
  const [users, setUsers] = useState<CloudUser[]>([])
  const [form, setForm] = useState(empty)
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  const load = () => listCloudUsers().then(setUsers).catch(reason => setError(reason.message)).finally(() => setLoading(false))
  useEffect(()=>{void load()}, [])
  const submit = async (event: React.FormEvent) => {
    event.preventDefault(); setSaving(true); setError('')
    try {
      if (form.id) await updateCloudUser(form)
      else await createCloudUser(form)
      setForm(empty); await load()
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Não foi possível salvar.') }
    finally { setSaving(false) }
  }
  return <main className="fade-in p-4 sm:p-7"><div className="mb-6"><h2 className="text-2xl font-bold text-slate-900">Perfis e usuários</h2><p className="text-sm text-slate-500">Controle quem pode acessar o sistema e o nível de permissão.</p></div>
    <div className="grid gap-5 xl:grid-cols-[420px_1fr]"><form onSubmit={submit} className="rounded-2xl border border-slate-200 bg-white p-6 shadow-soft"><div className="flex items-center gap-3"><div className="grid h-11 w-11 place-items-center rounded-xl bg-slate-100"><UserPlus size={20}/></div><div><h3 className="font-bold">{form.id?'Editar perfil':'Novo perfil'}</h3><p className="text-xs text-slate-500">Senha mínima de 8 caracteres</p></div></div>
      <div className="mt-5 space-y-4"><label className="block text-sm font-semibold">Nome completo<input required value={form.fullName} onChange={e=>setForm({...form,fullName:e.target.value})} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5"/></label><label className="block text-sm font-semibold">Usuário<input required value={form.username} onChange={e=>setForm({...form,username:e.target.value})} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5"/></label><label className="block text-sm font-semibold">{form.id?'Nova senha (opcional)':'Senha'}<input required={!form.id} type="password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5"/></label><label className="block text-sm font-semibold">Permissão<select value={form.role} onChange={e=>setForm({...form,role:e.target.value as 'admin'|'operator'})} className="mt-1.5 w-full rounded-xl border border-slate-200 px-3 py-2.5"><option value="admin">Administrador</option><option value="operator">Operador</option></select></label></div>
      {error&&<p className="mt-4 rounded-xl bg-red-50 p-3 text-sm font-semibold text-red-600">{error}</p>}<div className="mt-5 flex gap-2"><button disabled={saving} className="flex-1 rounded-xl bg-slate-800 px-4 py-3 font-bold text-white disabled:opacity-50">{saving?'Salvando...':form.id?'Salvar alterações':'Cadastrar perfil'}</button>{!!form.id&&<button type="button" onClick={()=>setForm(empty)} className="rounded-xl border border-slate-200 px-4 font-semibold">Cancelar</button>}</div></form>
      <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-soft"><div className="border-b border-slate-100 px-5 py-4"><h3 className="font-bold">Usuários cadastrados</h3></div>{loading?<p className="p-8 text-center text-slate-400">Carregando...</p>:<div className="divide-y divide-slate-100">{users.map(user=><div key={user.id} className="flex flex-wrap items-center gap-3 px-5 py-4"><div className="grid h-10 w-10 place-items-center rounded-full bg-slate-100 font-bold">{user.fullName.split(' ').map(x=>x[0]).slice(0,2).join('').toUpperCase()}</div><div className="min-w-[180px] flex-1"><b>{user.fullName}</b><div className="text-xs text-slate-500">@{user.username}</div></div><span className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-3 py-1 text-xs font-bold"><ShieldCheck size={14}/>{user.role==='admin'?'Administrador':'Operador'}</span><span className={`rounded-full px-3 py-1 text-xs font-bold ${user.active?'bg-emerald-50 text-emerald-700':'bg-red-50 text-red-600'}`}>{user.active?'Ativo':'Inativo'}</span><button onClick={()=>setForm({id:user.id,fullName:user.fullName,username:user.username,password:'',role:user.role,active:user.active})} className="rounded-lg border border-slate-200 p-2" title="Editar"><Pencil size={16}/></button><button onClick={async()=>{await updateCloudUser({...user,password:'',active:!user.active});load()}} className="rounded-lg border border-slate-200 px-3 py-2 text-xs font-bold">{user.active?'Desativar':'Ativar'}</button></div>)}</div>}</div></div></main>
}
