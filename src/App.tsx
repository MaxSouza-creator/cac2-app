import React, { useEffect, useMemo, useState } from "react";

// =============================================
// PALETA ADESTE
// =============================================
const C = {
  bg: "#0A1F2C",
  sidebar: "#007B93",
  card: "#0F2E3D",
  primary: "#0097A9",
  secondary: "#2DC5C4",
  danger: "#EF3340",
  gray: "#4D4D4F",
  border: "rgba(255,255,255,.12)",
  cB: "#0097A9",
  cG: "#2DC5C4",
  cR: "#EF3340",
  cY: "#F2A900",
  cP: "#7B61FF",
};

// =============================================
// TEAMS WEBHOOK
// =============================================
const WH =
  "https://default156b91ccb0c942d7937e54e4869724.59.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/03bf29936ab04702aa229404e488beb7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=2oOrp3KE6276N3A9osajTG_i-bpySme_TksCdSY0pZI";

async function nt(t: string, m: string) {
  try {
    await fetch(WH, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            contentUrl: null,
            content: {
              type: "AdaptiveCard",
              $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
              version: "1.4",
              body: [
                { type: "TextBlock", text: t, weight: "Bolder", size: "Medium", color: "Accent", wrap: true },
                { type: "TextBlock", text: m, wrap: true },
                { type: "TextBlock", text: "CAC 2.0 | " + new Date().toLocaleString("pt-BR"), size: "Small", isSubtle: true },
              ],
            },
          },
        ],
      }),
    });
  } catch (e) {
    console.log("Teams:", e);
  }
}

// =============================================
// STATUS / ROLES
// =============================================
const ST: string[] = ["Solicitacao criada", "Aguardando aprovacao", "Aprovado", "NF Anexada", "Concluido"];
const RL = { AD: "Administrador", MG: "Gerenciador", RQ: "Solicitante" } as const;
type Ro = (typeof RL)[keyof typeof RL];

function can(r: Ro, p: string): boolean {
  const m: any = {
    [RL.AD]: { BE: 1, WT: 1, WC: 1, AU: 1, CF: 1 },
    [RL.MG]: { BE: 1, WT: 1, WC: 1, AU: 1, CF: 1 },
    [RL.RQ]: { WC: 1 },
  };
  return !!(m[r] && m[r][p]);
}

// =============================================
// HELPERS
// =============================================
function sf(v: any, f = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : f;
}
function br(v: any) {
  return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(sf(v));
}
function ui(p = "ID") {
  return p + "-" + Math.random().toString(16).slice(2, 10).toUpperCase();
}
function nw() {
  return new Date().toISOString();
}

// =============================================
// TIPOS
// =============================================
type Si = { id: string; nome: string };
type Ct = { id: string; codigo: string; descricao: string };
type Cc = { id: string; codigo: string };
type Us = { id: string; nome: string; email: string; senha: string; role: Ro };
type Bu = { id: string; ano: number; siteId: string; cc: string; contaId: string; tipo: string; orcado: number; reservado: number; debitado: number };
type Wf = {
  id: string; criadoEm: string; solicitante: string; solicitanteEmail: string;
  siteId: string; cc: string; contaId: string; tipo: string; descricao: string;
  valor: number; status: string; valorReservado: number; reservadoAno: number | null;
  anoDebito: number | null; fiscalNome: string | null; fiscalData: string | null;
  fiscalAnexadoEm: string | null;
  historico: { at: string; by: string; action: string; to: string }[];
};
type Lg = { id: string; at: string; user: string; action: string; note: string };

// =============================================
// LOCALSTORAGE
// =============================================
function saveLS(key: string, data: any) {
  try { localStorage.setItem("cac2_" + key, JSON.stringify(data)); } catch {}
}
function loadLS<T>(key: string, fallback: T): T {
  try { const d = localStorage.getItem("cac2_" + key); return d ? JSON.parse(d) : fallback; } catch { return fallback; }
}

// =============================================
// EXPORT CSV / TXT
// =============================================
function downloadCSV(filename: string, headers: string[], rows: string[][]) {
  const bom = "\uFEFF";
  const csv = bom + [headers.join(";"), ...rows.map((r) => r.join(";"))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 500);
}
function downloadTXT(filename: string, text: string) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

// =============================================
// CHARTS
// =============================================
function Bar({ data, mx }: { data: { l: string; v: number; c: string }[]; mx?: number }) {
  const m = mx || Math.max(...data.map((d) => d.v), 1);
  return (
    <div style={{ display: "grid", gap: 8 }}>
      {data.map((d, i) => (
        <div key={i}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
            <span>{d.l}</span>
            <span style={{ fontWeight: 900 }}>{br(d.v)}</span>
          </div>
          <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 6, height: 20, overflow: "hidden" }}>
            <div style={{ width: `${Math.max((d.v / m) * 100, 2)}%`, height: "100%", background: d.c, borderRadius: 6 }} />
          </div>
        </div>
      ))}
    </div>
  );
}

function Don({ data }: { data: { l: string; v: number; c: string }[] }) {
  const t = data.reduce((a, d) => a + d.v, 0) || 1;
  let ac = 0;
  const sg = data.map((d) => { const s = ac; const p = (d.v / t) * 100; ac += p; return { ...d, s, p }; });
  const g = sg.map((s) => `${s.c} ${s.s}% ${s.s + s.p}%`).join(", ");
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
      <div style={{ width: 120, height: 120, borderRadius: "50%", background: `conic-gradient(${g})`, position: "relative", flexShrink: 0 }}>
        <div style={{ position: "absolute", top: "50%", left: "50%", transform: "translate(-50%,-50%)", width: 60, height: 60, borderRadius: "50%", background: C.card }} />
      </div>
      <div style={{ display: "grid", gap: 6 }}>
        {data.map((d, i) => (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12 }}>
            <div style={{ width: 12, height: 12, borderRadius: 3, background: d.c, flexShrink: 0 }} />
            <span>{d.l}: <b>{d.v}</b></span>
          </div>
        ))}
      </div>
    </div>
  );
}

// =============================================
// ESTILOS
// =============================================
const ff = "Trebuchet MS, Segoe UI, Arial, sans-serif";
const sP: React.CSSProperties = { display: "flex", minHeight: "100vh", background: C.bg, color: "#fff", fontFamily: ff };
const sS: React.CSSProperties = { width: 270, background: C.sidebar, padding: 18, boxSizing: "border-box", overflowY: "auto" };
const sM: React.CSSProperties = { flex: 1, padding: 18, boxSizing: "border-box", overflowY: "auto" };
const sC: React.CSSProperties = { background: C.card, padding: 16, borderRadius: 12, marginBottom: 12, border: "1px solid " + C.border };
const sB: React.CSSProperties = { background: C.primary, border: "none", padding: "10px 14px", borderRadius: 8, color: "#fff", cursor: "pointer", fontWeight: 800, marginRight: 8, marginBottom: 6 };
const sBD: React.CSSProperties = { ...sB, background: C.danger };
const sBG: React.CSSProperties = { ...sB, background: "rgba(255,255,255,.10)", border: "1px solid rgba(255,255,255,.18)" };
const sMn: React.CSSProperties = { display: "block", width: "100%", padding: "10px 12px", marginBottom: 8, borderRadius: 10, border: "none", background: C.primary, color: "#fff", cursor: "pointer", fontWeight: 800, textAlign: "left" };
const sI: React.CSSProperties = { width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid rgba(255,255,255,.18)", background: "rgba(0,0,0,.18)", color: "#fff", outline: "none", boxSizing: "border-box", marginBottom: 10 };
const sL: React.CSSProperties = { display: "block", fontSize: 12, opacity: 0.85, fontWeight: 800, marginBottom: 4 };
const sTh: React.CSSProperties = { textAlign: "left", padding: 10, fontSize: 12, opacity: 0.8, borderBottom: "1px solid " + C.border };
const sTd: React.CSSProperties = { padding: 10, borderBottom: "1px solid rgba(255,255,255,.06)", verticalAlign: "top" };
const sSm: React.CSSProperties = { fontSize: 12, opacity: 0.85 };
const sBa: React.CSSProperties = { display: "inline-block", padding: "4px 10px", borderRadius: 999, background: "rgba(255,255,255,.10)", border: "1px solid rgba(255,255,255,.18)", fontWeight: 800, fontSize: 12, marginRight: 8, marginBottom: 6 };
const sG2: React.CSSProperties = { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 };
const sG3: React.CSSProperties = { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 };
const sAl: React.CSSProperties = { background: "#EF3340", color: "#fff", padding: 12, borderRadius: 10, marginBottom: 12, fontWeight: 800 };

// =============================================
// SEEDS
// =============================================
const yr = new Date().getFullYear();

const seedUsers: Us[] = [
  { id: "U1", nome: "Admin", email: "admin@adeste.com.br", senha: "123456", role: RL.AD },
  { id: "U2", nome: "Gerenciador", email: "gerente@adeste.com.br", senha: "123456", role: RL.MG },
  { id: "U3", nome: "Solicitante", email: "solicitante@adeste.com.br", senha: "123456", role: RL.RQ },
];

const seedSites: Si[] = [
  { id: "S1", nome: "Site Santo Andre" },
  { id: "S2", nome: "Site Sao Bernardo" },
];

const seedCcs: Cc[] = [
  { id: "CC1", codigo: "CC-001" },
  { id: "CC2", codigo: "CC-002" },
];

const seedContas: Ct[] = [
  { id: "C1", codigo: "5100", descricao: "Manutencao predial" },
  { id: "C2", codigo: "5200", descricao: "Servicos terceirizados" },
  { id: "C3", codigo: "6100", descricao: "Investimento em maquinas" },
];

const seedBudgets: Bu[] = [
  { id: "B1", ano: yr, siteId: "S1", cc: "CC-001", contaId: "C1", tipo: "OPEX", orcado: 250000, reservado: 30000, debitado: 45000 },
  { id: "B2", ano: yr, siteId: "S1", cc: "CC-001", contaId: "C3", tipo: "CAPEX", orcado: 400000, reservado: 80000, debitado: 0 },
  { id: "B3", ano: yr, siteId: "S2", cc: "CC-002", contaId: "C2", tipo: "OPEX", orcado: 180000, reservado: 10000, debitado: 25000 },
];

const seedWfs: Wf[] = [
  {
    id: "WF-0001", criadoEm: nw(), solicitante: "Admin", solicitanteEmail: "admin@adeste.com.br",
    siteId: "S1", cc: "CC-001", contaId: "C1", tipo: "OPEX", descricao: "Manutencao predial",
    valor: 15000, status: ST[0], valorReservado: 0, reservadoAno: null, anoDebito: null,
    fiscalNome: null, fiscalData: null, fiscalAnexadoEm: null,
    historico: [{ at: nw(), by: "system", action: "Seed", to: ST[0] }],
  },
  {
    id: "WF-0002", criadoEm: nw(), solicitante: "Admin", solicitanteEmail: "admin@adeste.com.br",
    siteId: "S2", cc: "CC-002", contaId: "C2", tipo: "OPEX", descricao: "Servicos de limpeza",
    valor: 8000, status: "Aprovado", valorReservado: 8000, reservadoAno: yr, anoDebito: null,
    fiscalNome: null, fiscalData: null, fiscalAnexadoEm: null,
    historico: [{ at: nw(), by: "system", action: "Seed", to: "Aprovado" }],
  },
];

// =============================================
// APP
// =============================================
export default function App() {
  // === AUTH ===
  const [logged, setLogged] = useState(false);
  const [loginEmail, setLoginEmail] = useState("");
  const [loginSenha, setLoginSenha] = useState("");
  const [loginError, setLoginError] = useState("");

  const [us, setUs] = useState<Us[]>(() => loadLS("users", seedUsers));
  const [user, setUser] = useState<Us | null>(null);
  const [tab, setTab] = useState("dashboard");
  const [tL, setTL] = useState<string[]>([]);

  const [sites, setSites] = useState<Si[]>(() => loadLS("sites", seedSites));
  const [ccs, setCcs] = useState<Cc[]>(() => loadLS("ccs", seedCcs));
  const [contas, setContas] = useState<Ct[]>(() => loadLS("contas", seedContas));
  const [budgets, setBudgets] = useState<Bu[]>(() => loadLS("budgets", seedBudgets));
  const [bf, setBf] = useState({ ano: yr, siteId: "S1", cc: "CC-001", contaId: "C1", tipo: "OPEX", orcado: 0 });

  const [wfs, setWfs] = useState<Wf[]>(() => loadLS("wfs", seedWfs));
  const [wfF, setWfF] = useState({ descricao: "", valor: "", tipo: "OPEX", siteId: "S1", cc: "CC-001", contaId: "C1" });
  const [alertas, setAlertas] = useState<string[]>([]);
  const [audit, setAudit] = useState<Lg[]>(() => loadLS("audit", [{ id: ui("LOG"), at: nw(), user: "system", action: "Seed", note: "Dados iniciais" }]));
  const [sfilt, setSfilt] = useState("ALL");
  const [cfU, setCfU] = useState({ nome: "", email: "", senha: "123456", role: RL.RQ as Ro });
  const [cfS, setCfS] = useState("");
  const [cfC, setCfC] = useState("");
  const [cfCt, setCfCt] = useState({ codigo: "", descricao: "" });

  // === PERSIST ===
  useEffect(() => { saveLS("users", us); }, [us]);
  useEffect(() => { saveLS("sites", sites); }, [sites]);
  useEffect(() => { saveLS("ccs", ccs); }, [ccs]);
  useEffect(() => { saveLS("contas", contas); }, [contas]);
  useEffect(() => { saveLS("budgets", budgets); }, [budgets]);
  useEffect(() => { saveLS("wfs", wfs); }, [wfs]);
  useEffect(() => { saveLS("audit", audit); }, [audit]);

  // === AUTO LOGIN ===
  useEffect(() => {
    const saved = loadLS<Us | null>("session", null);
    if (saved) { setUser(saved); setLogged(true); }
  }, []);

  // === LOGIN ===
  function handleLogin() {
    const found = us.find((u) => u.email === loginEmail.trim().toLowerCase() && u.senha === loginSenha);
    if (!found) { setLoginError("Email ou senha incorretos."); return; }
    setUser(found);
    setLogged(true);
    saveLS("session", found);
    setLoginEmail("");
    setLoginSenha("");
    setLoginError("");
  }

  function handleLogout() {
    setUser(null);
    setLogged(false);
    localStorage.removeItem("cac2_session");
    setTab("dashboard");
  }

  // === HELPERS ===
  const sn = (id: string) => sites.find((s) => s.id === id)?.nome || id;
  const cl = (id: string) => { const c = contas.find((x) => x.id === id); return c ? c.codigo + " - " + c.descricao : id; };

  function lg(a: string, n: string) {
    setAudit((p) => [{ id: ui("LOG"), at: nw(), user: user?.email || "system", action: a, note: n }, ...p]);
  }

  function notify(t: string, m: string) {
    nt(t, m);
    setTL((p) => [new Date().toLocaleString("pt-BR") + " | " + t, ...p]);
  }

  // === KPIs ===
  const wff = useMemo(() => wfs.filter((w) => sfilt === "ALL" || w.siteId === sfilt), [wfs, sfilt]);
  const kO = useMemo(() => budgets.reduce((a, b) => a + sf(b.orcado), 0), [budgets]);
  const kR = useMemo(() => budgets.reduce((a, b) => a + sf(b.reservado), 0), [budgets]);
  const kD = useMemo(() => budgets.reduce((a, b) => a + sf(b.debitado), 0), [budgets]);
  const kDi = kO - kR - kD;
  const kT = useMemo(() => wff.reduce((a, b) => a + sf(b.valor), 0), [wff]);

  const cBS = useMemo(() => {
    const f = sfilt === "ALL" ? budgets : budgets.filter((b) => b.siteId === sfilt);
    const m: Record<string, { o: number; r: number; d: number }> = {};
    f.forEach((b) => { const n = sn(b.siteId); if (!m[n]) m[n] = { o: 0, r: 0, d: 0 }; m[n].o += sf(b.orcado); m[n].r += sf(b.reservado); m[n].d += sf(b.debitado); });
    return m;
  }, [budgets, sfilt, sites]);

  const cSC = useMemo(() => {
    const c: Record<string, number> = {};
    ST.forEach((s) => (c[s] = 0));
    wff.forEach((w) => { c[w.status] = (c[w.status] || 0) + 1; });
    return c;
  }, [wff]);

  const cTp = useMemo(() => {
    let o = 0, ca = 0;
    wff.forEach((w) => { if (w.tipo === "OPEX") o += sf(w.valor); else ca += sf(w.valor); });
    return { o, c: ca };
  }, [wff]);

  function fl(ano: number, si: string, cc: string, ci: string, tp: string) {
    return budgets.find((b) => b.ano === ano && b.siteId === si && b.cc === cc && b.contaId === ci && b.tipo === tp) || null;
  }

  // === CRIAR WF ===
  function criarWf() {
    if (!user) return;
    if (!can(user.role, "WC") && !can(user.role, "WT")) { alert("Sem permissao."); return; }
    if (!wfF.descricao || !wfF.valor) { alert("Preencha todos os campos."); return; }
    const n: Wf = {
      id: ui("WF"), criadoEm: nw(), solicitante: user.nome, solicitanteEmail: user.email,
      siteId: wfF.siteId, cc: wfF.cc.trim(), contaId: wfF.contaId, tipo: wfF.tipo,
      descricao: wfF.descricao.trim(), valor: sf(wfF.valor), status: ST[0],
      valorReservado: 0, reservadoAno: null, anoDebito: null,
      fiscalNome: null, fiscalData: null, fiscalAnexadoEm: null,
      historico: [{ at: nw(), by: user.email, action: "Criou", to: ST[0] }],
    };
    setWfs((p) => [n, ...p]);
    lg("Workflow", "Criou " + n.id);
    notify("Novo Workflow", n.id + " - " + n.descricao + " | " + br(n.valor) + " | " + n.solicitante);
    setWfF({ descricao: "", valor: "", tipo: "OPEX", siteId: "S1", cc: "CC-001", contaId: "C1" });
  }

  // === ANEXAR NF (com arquivo real) ===
  function anexar(id: string, file: File) {
    if (!user) return;
    const reader = new FileReader();
    reader.onload = () => {
      setWfs((p) => p.map((w) =>
        w.id !== id ? w : {
          ...w,
          fiscalNome: file.name,
          fiscalData: reader.result as string,
          fiscalAnexadoEm: nw(),
          status: "NF Anexada",
          historico: [{ at: nw(), by: user.email, action: "Anexou NF: " + file.name, to: "NF Anexada" }, ...w.historico],
        }
      ));
      lg("Fiscal", "NF em " + id);
      setAlertas((p) => ["NF anexada no " + id + " por " + user.nome + " (" + file.name + ")", ...p]);
      notify("NF Anexada", id + " | NF: " + file.name + " | Por: " + user.nome);
    };
    reader.readAsDataURL(file);
  }

  // === BAIXAR NF ===
  function baixarNf(w: Wf) {
    if (!w.fiscalData) return;
    const a = document.createElement("a");
    a.href = w.fiscalData;
    a.download = w.fiscalNome || "nota_fiscal";
    a.click();
  }

  // === AVANCAR STATUS ===
  function av(id: string) {
    if (!user) return;
    const w = wfs.find((x) => x.id === id);
    if (!w) return;
    if (!can(user.role, "WT")) { alert("Sem permissao."); return; }
    const i = ST.indexOf(w.status);
    if (i < 0 || i >= ST.length - 1) return;
    const nx = ST[i + 1];

    if (nx === "NF Anexada") { alert("Aguardando solicitante anexar NF."); return; }
    if (nx === "Concluido" && !w.fiscalNome) { alert("NF nao anexada."); return; }

    if (nx === "Aguardando aprovacao" && sf(w.valorReservado) === 0) {
      const li = fl(yr, w.siteId, w.cc, w.contaId, w.tipo);
      if (!li) { alert("Linha orcamento inexistente."); return; }
      const dp = sf(li.orcado) - sf(li.reservado) - sf(li.debitado);
      if (sf(w.valor) > dp) { alert("Saldo insuficiente. Disponivel: " + br(dp)); return; }
      setBudgets((p) => p.map((b) => b.id === li.id ? { ...b, reservado: sf(b.reservado) + sf(w.valor) } : b));
      setWfs((p) => p.map((x) => x.id !== id ? x : { ...x, status: nx, valorReservado: sf(w.valor), reservadoAno: yr, historico: [{ at: nw(), by: user.email, action: "Avancou", to: nx }, ...x.historico] }));
      lg("Reserva", br(w.valor) + " p/ " + id);
      notify("Status: " + nx, id + " | Reservado: " + br(w.valor) + " | " + user.nome);
      return;
    }

    if (nx === "Concluido") {
      const ld = fl(yr, w.siteId, w.cc, w.contaId, w.tipo);
      if (!ld) { alert("Linha inexistente."); return; }
      if (w.reservadoAno) {
        const ol = fl(w.reservadoAno, w.siteId, w.cc, w.contaId, w.tipo);
        if (ol) setBudgets((p) => p.map((b) => b.id === ol.id ? { ...b, reservado: Math.max(0, sf(b.reservado) - sf(w.valorReservado)) } : b));
      }
      setBudgets((p) => p.map((b) => b.id === ld.id ? { ...b, debitado: sf(b.debitado) + sf(w.valor) } : b));
      setWfs((p) => p.map((x) => x.id !== id ? x : { ...x, status: nx, anoDebito: yr, valorReservado: 0, historico: [{ at: nw(), by: user.email, action: "Concluiu+Debitou", to: nx }, ...x.historico] }));
      lg("SAP", "Debitou " + br(w.valor) + " p/ " + id);
      notify("Concluido", id + " | Debitado: " + br(w.valor) + " | " + user.nome);
      return;
    }

    setWfs((p) => p.map((x) => x.id !== id ? x : { ...x, status: nx, historico: [{ at: nw(), by: user.email, action: "Avancou", to: nx }, ...x.historico] }));
    lg("Workflow", id + " -> " + nx);
    notify("Status: " + nx, id + " | " + user.nome);
  }

  function exWf(id: string) { if (!confirm("Excluir?")) return; setWfs((p) => p.filter((w) => w.id !== id)); lg("Workflow", "Excluiu " + id); }

  // === ORCAMENTO ===
  function addB() {
    if (!user) return;
    if (!can(user.role, "BE")) { alert("Sem permissao."); return; }
    const d = fl(sf(bf.ano, yr), bf.siteId, bf.cc.trim(), bf.contaId, bf.tipo);
    if (d) { alert("Ja existe."); return; }
    const n: Bu = { id: ui("BUD"), ano: sf(bf.ano, yr), siteId: bf.siteId, cc: bf.cc.trim(), contaId: bf.contaId, tipo: bf.tipo, orcado: sf(bf.orcado), reservado: 0, debitado: 0 };
    setBudgets((p) => [n, ...p]);
    lg("Orcamento", "Criou " + n.id);
    setBf({ ano: yr, siteId: "S1", cc: "CC-001", contaId: "C1", tipo: "OPEX", orcado: 0 });
  }

  function delB(id: string) { if (!user || !can(user.role, "BE")) return; if (!confirm("Excluir?")) return; setBudgets((p) => p.filter((b) => b.id !== id)); lg("Orcamento", "Excluiu " + id); }

  // === CONFIG ===
  function addU() {
    if (!cfU.nome || !cfU.email) { alert("Preencha nome e email."); return; }
    const dup = us.find((u) => u.email === cfU.email.trim().toLowerCase());
    if (dup) { alert("Email ja cadastrado."); return; }
    setUs((p) => [...p, { id: ui("U"), nome: cfU.nome.trim(), email: cfU.email.trim().toLowerCase(), senha: cfU.senha || "123456", role: cfU.role }]);
    lg("Config", "Usuario " + cfU.email);
    setCfU({ nome: "", email: "", senha: "123456", role: RL.RQ });
  }
  function delU(id: string) { if (!confirm("Excluir?")) return; setUs((p) => p.filter((u) => u.id !== id)); }
  function addSi() { if (!cfS.trim()) return; setSites((p) => [...p, { id: ui("S"), nome: cfS.trim() }]); lg("Config", "Site"); setCfS(""); }
  function delSi(id: string) { if (!confirm("Excluir?")) return; setSites((p) => p.filter((s) => s.id !== id)); }
  function addCc() { if (!cfC.trim()) return; setCcs((p) => [...p, { id: ui("CC"), codigo: cfC.trim() }]); lg("Config", "CC"); setCfC(""); }
  function delCc(id: string) { if (!confirm("Excluir?")) return; setCcs((p) => p.filter((c) => c.id !== id)); }
  function addCt() { if (!cfCt.codigo) return; setContas((p) => [...p, { id: ui("CT"), codigo: cfCt.codigo.trim(), descricao: cfCt.descricao.trim() }]); lg("Config", "Conta"); setCfCt({ codigo: "", descricao: "" }); }
  function delCt(id: string) { if (!confirm("Excluir?")) return; setContas((p) => p.filter((c) => c.id !== id)); }

  // === EXPORTS ===
  function exportWfCSV() {
    const h = ["ID", "Descricao", "Solicitante", "Site", "CC", "Conta", "Tipo", "Valor", "Status", "NF", "Criado"];
    const r = wfs.map((w) => [w.id, w.descricao, w.solicitante, sn(w.siteId), w.cc, cl(w.contaId), w.tipo, String(w.valor), w.status, w.fiscalNome || "", w.criadoEm]);
    downloadCSV("workflows_cac2.csv", h, r);
  }
  function exportBudCSV() {
    const h = ["ID", "Ano", "Site", "CC", "Conta", "Tipo", "Orcado", "Reservado", "Debitado", "Disponivel"];
    const r = budgets.map((b) => [b.id, String(b.ano), sn(b.siteId), b.cc, cl(b.contaId), b.tipo, String(b.orcado), String(b.reservado), String(b.debitado), String(sf(b.orcado) - sf(b.reservado) - sf(b.debitado))]);
    downloadCSV("orcamento_cac2.csv", h, r);
  }
  function exportAuditTXT() {
    const lines = audit.map((a) => a.at + " | " + a.user + " | " + a.action + " | " + a.note);
    downloadTXT("auditoria_cac2.txt", "AUDITORIA CAC 2.0\n" + "=".repeat(50) + "\n\n" + lines.join("\n"));
  }

  function resetData() { if (!confirm("ATENCAO: Isso vai apagar TODOS os dados. Continuar?")) return; localStorage.clear(); window.location.reload(); }
  function testTeams() { notify("Teste CAC 2.0", "Integracao funcionando!"); alert("Enviado!"); }

  // =============================================
  // TELA DE LOGIN
  // =============================================
  if (!logged || !user) {
    return (
      <div style={{
        minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
        background: "linear-gradient(135deg, #0A1F2C 0%, #007B93 100%)", fontFamily: ff, color: "#fff",
      }}>
        <div style={{
          width: 400, padding: 36, borderRadius: 20,
          background: "rgba(15,46,61,.9)", border: "1px solid rgba(255,255,255,.15)",
          boxShadow: "0 20px 60px rgba(0,0,0,.5)",
        }}>
          <div style={{ textAlign: "center", marginBottom: 28 }}>
            <div style={{ fontSize: 36, fontWeight: 900, color: C.secondary }}>Adeste</div>
            <div style={{ fontSize: 15, opacity: 0.8, marginTop: 4 }}>CAC 2.0 — Orcamento & Workflow</div>
          </div>

          {loginError && (
            <div style={{
              background: "rgba(239,51,64,.2)", border: "1px solid " + C.danger,
              borderRadius: 10, padding: 10, marginBottom: 14, fontSize: 13, textAlign: "center",
            }}>{loginError}</div>
          )}

          <div>
            <label style={sL}>Email</label>
            <input style={sI} placeholder="seu.email@adeste.com.br" value={loginEmail}
              onChange={(e) => setLoginEmail(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
          </div>
          <div>
            <label style={sL}>Senha</label>
            <input style={sI} type="password" placeholder="Digite sua senha" value={loginSenha}
              onChange={(e) => setLoginSenha(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
          </div>

          <button onClick={handleLogin} style={{
            ...sB, width: "100%", padding: 14, fontSize: 16, marginTop: 8, borderRadius: 12, marginRight: 0,
          }}>Entrar</button>

         
        </div>
      </div>
    );
  }

  // =============================================
  // LAYOUT PRINCIPAL
  // =============================================
  return (
    <div style={sP}>
      <aside style={sS}>
        <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 12, padding: 12, marginBottom: 14, border: "1px solid rgba(255,255,255,.15)" }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 900 }}>Adeste CAC 2.0</h3>
          <p style={{ margin: "4px 0 0", fontSize: 11, opacity: 0.85 }}>OPEX / CAPEX / Workflow</p>
          <p style={{ margin: "4px 0 0", fontSize: 10, opacity: 0.6 }}>Dados salvos automaticamente</p>
        </div>

        <button style={sMn} onClick={() => setTab("dashboard")}>Dashboard</button>
        <button style={sMn} onClick={() => setTab("cadastro")}>Nova Solicitacao</button>
        <button style={sMn} onClick={() => setTab("workflow")}>Workflows</button>
        <button style={sMn} onClick={() => setTab("fiscal")}>Anexar NF</button>
        <button style={sMn} onClick={() => setTab("orcamento")}>Orcamento</button>
        {can(user.role, "AU") && <button style={sMn} onClick={() => setTab("auditoria")}>Auditoria</button>}
        {can(user.role, "CF") && <button style={{ ...sMn, background: C.gray }} onClick={() => setTab("config")}>Configuracoes</button>}

        <div style={{ height: 1, background: "rgba(255,255,255,.18)", margin: "14px 0" }} />

        <div style={{ background: "rgba(0,0,0,.2)", borderRadius: 12, padding: 12, marginBottom: 10 }}>
          <div style={{ fontWeight: 900, fontSize: 13 }}>{user.nome}</div>
          <div style={{ fontSize: 11, opacity: 0.7 }}>{user.email}</div>
          <div style={sBa}>{user.role}</div>
          <button onClick={handleLogout} style={{ ...sBG, width: "100%", marginTop: 8, marginRight: 0, fontSize: 12, padding: 8 }}>Sair</button>
        </div>

        <div style={{ height: 1, background: "rgba(255,255,255,.18)", margin: "14px 0" }} />
        <button style={{ ...sMn, background: "#F2A900", color: "#000" }} onClick={testTeams}>Testar Teams</button>
        <button style={{ ...sMn, background: "rgba(255,255,255,.08)", fontSize: 10 }} onClick={resetData}>Resetar dados</button>

        {tL.length > 0 && (
          <div style={{ marginTop: 8 }}>
            <div style={{ fontSize: 10, opacity: 0.7 }}>Teams:</div>
            {tL.slice(0, 3).map((t, i) => (<div key={i} style={{ fontSize: 10, opacity: 0.5, marginTop: 4 }}>{t}</div>))}
          </div>
        )}
      </aside>

      <main style={sM}>
        {alertas.length > 0 && can(user.role, "WT") && (
          <div>{alertas.map((a, i) => (
            <div key={i} style={sAl}>{a}
              <button style={{ ...sB, background: "#fff", color: C.danger, marginLeft: 10 }} onClick={() => setAlertas((p) => p.filter((_, idx) => idx !== i))}>OK</button>
            </div>
          ))}</div>
        )}

        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
          <span style={sBa}>WFs: {wff.length}</span>
          <span style={sBa}>Total: {br(kT)}</span>
          <span style={sBa}>Orcado: {br(kO)}</span>
          <span style={sBa}>Reservado: {br(kR)}</span>
          <span style={sBa}>Debitado: {br(kD)}</span>
          <span style={sBa}>Disponivel: {br(kDi)}</span>
        </div>

        {/* === DASHBOARD === */}
        {tab === "dashboard" && (
          <div>
            <h1 style={{ marginTop: 0 }}>Dashboard</h1>
            <div style={sC}><label style={sL}>Filtrar por site</label><select style={{ ...sI, maxWidth: 300 }} value={sfilt} onChange={(e) => setSfilt(e.target.value)}><option value="ALL">Todos</option>{sites.map((s) => (<option key={s.id} value={s.id}>{s.nome}</option>))}</select></div>
            <div style={sG2}>
              <div style={sC}><div style={sSm}>Workflows</div><div style={{ fontSize: 28, fontWeight: 950 }}>{wff.length}</div></div>
              <div style={sC}><div style={sSm}>Valor total</div><div style={{ fontSize: 28, fontWeight: 950 }}>{br(kT)}</div></div>
              <div style={sC}><div style={sSm}>Orcado</div><div style={{ fontSize: 28, fontWeight: 950 }}>{br(kO)}</div></div>
              <div style={sC}><div style={sSm}>Disponivel</div><div style={{ fontSize: 28, fontWeight: 950, color: kDi < 0 ? "#ffd1d1" : "#d9fff7" }}>{br(kDi)}</div></div>
            </div>
            <div style={sC}><h3 style={{ marginTop: 0 }}>Orcamento por Site</h3>{Object.entries(cBS).map(([n, v]) => (<div key={n} style={{ marginBottom: 16 }}><div style={{ fontWeight: 900, marginBottom: 6 }}>{n}</div><Bar mx={Math.max(...Object.values(cBS).map((x) => x.o), 1)} data={[{ l: "Orcado", v: v.o, c: C.cB }, { l: "Reservado", v: v.r, c: C.cY }, { l: "Debitado", v: v.d, c: C.cR }]} /></div>))}</div>
            <div style={sG2}>
              <div style={sC}><h3 style={{ marginTop: 0 }}>WFs por Status</h3><Don data={ST.map((s, i) => ({ l: s, v: cSC[s] || 0, c: [C.cB, C.cY, C.cG, C.cP, C.cR][i] || C.gray }))} /></div>
              <div style={sC}><h3 style={{ marginTop: 0 }}>OPEX vs CAPEX</h3><Bar data={[{ l: "OPEX", v: cTp.o, c: C.cB }, { l: "CAPEX", v: cTp.c, c: C.cG }]} /></div>
            </div>
            <div style={{ display: "flex", gap: 8, marginBottom: 12 }}><button style={sB} onClick={exportWfCSV}>Exportar Workflows (Excel)</button><button style={sB} onClick={exportBudCSV}>Exportar Orcamento (Excel)</button></div>
            <div style={sC}><div style={{ fontWeight: 950, marginBottom: 8 }}>Ultimos workflows</div><table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr><th style={sTh}>ID</th><th style={sTh}>Descricao</th><th style={sTh}>Status</th><th style={sTh}>Valor</th><th style={sTh}>NF</th></tr></thead><tbody>{wff.slice(0, 8).map((w) => (<tr key={w.id}><td style={sTd}>{w.id}</td><td style={sTd}>{w.descricao}</td><td style={sTd}>{w.status}</td><td style={sTd}>{br(w.valor)}</td><td style={sTd}>{w.fiscalNome || "---"}</td></tr>))}</tbody></table></div>
          </div>
        )}

        {/* === CADASTRO === */}
        {tab === "cadastro" && (
          <div><h1 style={{ marginTop: 0 }}>Nova Solicitacao</h1><div style={sC}><div style={sG2}>
            <div><label style={sL}>Descricao</label><input style={sI} value={wfF.descricao} onChange={(e) => setWfF({ ...wfF, descricao: e.target.value })} placeholder="Ex.: Manutencao" /></div>
            <div><label style={sL}>Valor</label><input style={sI} type="number" value={wfF.valor} onChange={(e) => setWfF({ ...wfF, valor: e.target.value })} placeholder="15000" /></div>
            <div><label style={sL}>Tipo</label><select style={sI} value={wfF.tipo} onChange={(e) => setWfF({ ...wfF, tipo: e.target.value })}><option value="OPEX">OPEX</option><option value="CAPEX">CAPEX</option></select></div>
            <div><label style={sL}>Solicitante</label><input style={sI} value={user.nome} disabled /></div>
            <div><label style={sL}>Site</label><select style={sI} value={wfF.siteId} onChange={(e) => setWfF({ ...wfF, siteId: e.target.value })}>{sites.map((s) => (<option key={s.id} value={s.id}>{s.nome}</option>))}</select></div>
            <div><label style={sL}>CC</label><select style={sI} value={wfF.cc} onChange={(e) => setWfF({ ...wfF, cc: e.target.value })}>{ccs.map((c) => (<option key={c.id} value={c.codigo}>{c.codigo}</option>))}</select></div>
          </div><div><label style={sL}>Conta</label><select style={sI} value={wfF.contaId} onChange={(e) => setWfF({ ...wfF, contaId: e.target.value })}>{contas.map((c) => (<option key={c.id} value={c.id}>{c.codigo} - {c.descricao}</option>))}</select></div>
          <div style={{ marginTop: 10 }}><button style={sB} onClick={criarWf}>Criar workflow</button><button style={sBG} onClick={() => setWfF({ descricao: "", valor: "", tipo: "OPEX", siteId: "S1", cc: "CC-001", contaId: "C1" })}>Limpar</button></div></div></div>
        )}

        {/* === WORKFLOWS === */}
        {tab === "workflow" && (
          <div><h1 style={{ marginTop: 0 }}>Workflows</h1>
          <div style={{ marginBottom: 12 }}><button style={sB} onClick={exportWfCSV}>Exportar Excel</button></div>
          <div style={sC}><div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr><th style={sTh}>ID</th><th style={sTh}>Descricao</th><th style={sTh}>Site/CC</th><th style={sTh}>Tipo</th><th style={sTh}>Valor</th><th style={sTh}>Status</th><th style={sTh}>NF</th><th style={sTh}>Acoes</th></tr></thead>
          <tbody>{wfs.map((w) => (
            <tr key={w.id}>
              <td style={sTd}>{w.id}</td>
              <td style={sTd}><b>{w.descricao}</b><div style={sSm}>{w.solicitante}</div></td>
              <td style={sTd}>{sn(w.siteId)} / {w.cc}</td>
              <td style={sTd}>{w.tipo}</td>
              <td style={sTd}>{br(w.valor)}</td>
              <td style={sTd}><b>{w.status}</b><div style={sSm}>Res: {br(w.valorReservado)}</div></td>
              <td style={sTd}>
                {w.fiscalNome ? (<span>{w.fiscalNome} <button style={{ ...sBG, padding: "4px 8px", fontSize: 11 }} onClick={() => baixarNf(w)}>Baixar</button></span>) : "---"}
              </td>
              <td style={sTd}>
                {can(user.role, "WT") && w.status !== "Concluido" && (<button style={sB} onClick={() => av(w.id)}>Avancar</button>)}
                <button style={sBD} onClick={() => exWf(w.id)}>Excluir</button>
              </td>
            </tr>
          ))}{wfs.length === 0 && (<tr><td style={sTd} colSpan={8}>Nenhum.</td></tr>)}</tbody></table></div></div>
          {wfs.length > 0 && (<div style={sC}><div style={{ fontWeight: 950, marginBottom: 8 }}>Historico: {wfs[0].id}</div>{wfs[0].historico.map((h, i) => (<div key={i} style={{ padding: 8, background: "rgba(255,255,255,.04)", borderRadius: 8, marginBottom: 6 }}><b>{h.action}</b> {"->"} {h.to}<div style={sSm}>{h.by} | {h.at}</div></div>))}</div>)}
          </div>
        )}

        {/* === ANEXAR NF === */}
        {tab === "fiscal" && (
          <div><h1 style={{ marginTop: 0 }}>Anexar NF</h1>
          <div style={sC}><div style={sSm}>Selecione um workflow <b>Aprovado</b> e anexe o documento fiscal.</div></div>
          {wfs.filter((w) => w.status === "Aprovado").length === 0 && (<div style={sC}>Nenhum workflow com status "Aprovado".</div>)}
          {wfs.filter((w) => w.status === "Aprovado").map((w) => (
            <div key={w.id} style={sC}>
              <b>{w.id}</b> - {w.descricao} - {br(w.valor)}
              <div style={sSm}>{w.solicitante} | {sn(w.siteId)}</div>
              <div style={{ marginTop: 10 }}>
                <input type="file" onChange={(e) => { const f = e.target.files?.[0]; if (f) anexar(w.id, f); }} />
              </div>
            </div>
          ))}
          <div style={{ marginTop: 20 }}><h3>Workflows com NF anexada</h3>
          {wfs.filter((w) => w.fiscalNome).map((w) => (
            <div key={w.id} style={{ ...sC, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div><b>{w.id}</b> - {w.descricao}<div style={sSm}>NF: {w.fiscalNome} | {w.fiscalAnexadoEm}</div><div style={sSm}>Status: {w.status}</div></div>
              <button style={sBG} onClick={() => baixarNf(w)}>Baixar NF</button>
            </div>
          ))}</div></div>
        )}

        {/* === ORCAMENTO === */}
        {tab === "orcamento" && (
          <div><h1 style={{ marginTop: 0 }}>Orcamento</h1>
          <div style={{ marginBottom: 12 }}><button style={sB} onClick={exportBudCSV}>Exportar Excel</button></div>
          {!can(user.role, "BE") && (<div style={sC}>Somente Admin/Gerenciador.</div>)}
          {can(user.role, "BE") && (<div style={sC}><div style={{ fontWeight: 950, marginBottom: 10 }}>Nova linha</div><div style={sG3}>
            <div><label style={sL}>Ano</label><input style={sI} type="number" value={bf.ano} onChange={(e) => setBf({ ...bf, ano: sf(e.target.value, yr) })} /></div>
            <div><label style={sL}>Site</label><select style={sI} value={bf.siteId} onChange={(e) => setBf({ ...bf, siteId: e.target.value })}>{sites.map((s) => (<option key={s.id} value={s.id}>{s.nome}</option>))}</select></div>
            <div><label style={sL}>Tipo</label><select style={sI} value={bf.tipo} onChange={(e) => setBf({ ...bf, tipo: e.target.value })}><option value="OPEX">OPEX</option><option value="CAPEX">CAPEX</option></select></div>
            <div><label style={sL}>CC</label><select style={sI} value={bf.cc} onChange={(e) => setBf({ ...bf, cc: e.target.value })}>{ccs.map((c) => (<option key={c.id} value={c.codigo}>{c.codigo}</option>))}</select></div>
            <div><label style={sL}>Conta</label><select style={sI} value={bf.contaId} onChange={(e) => setBf({ ...bf, contaId: e.target.value })}>{contas.map((c) => (<option key={c.id} value={c.id}>{c.codigo} - {c.descricao}</option>))}</select></div>
            <div><label style={sL}>Orcado</label><input style={sI} type="number" value={bf.orcado} onChange={(e) => setBf({ ...bf, orcado: sf(e.target.value) })} /></div>
          </div><button style={sB} onClick={addB}>Adicionar</button></div>)}
          <div style={sC}><div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr><th style={sTh}>Ano</th><th style={sTh}>Site</th><th style={sTh}>CC</th><th style={sTh}>Conta</th><th style={sTh}>Tipo</th><th style={sTh}>Orcado</th><th style={sTh}>Reservado</th><th style={sTh}>Debitado</th><th style={sTh}>Disponivel</th><th style={sTh}>Acoes</th></tr></thead>
          <tbody>{budgets.map((b) => { const d = sf(b.orcado) - sf(b.reservado) - sf(b.debitado); return (
            <tr key={b.id}><td style={sTd}>{b.ano}</td><td style={sTd}>{sn(b.siteId)}</td><td style={sTd}>{b.cc}</td><td style={sTd}>{cl(b.contaId)}</td><td style={sTd}><b>{b.tipo}</b></td><td style={sTd}>{br(b.orcado)}</td><td style={sTd}>{br(b.reservado)}</td><td style={sTd}>{br(b.debitado)}</td><td style={{ ...sTd, fontWeight: 900, color: d < 0 ? "#ffd1d1" : "#d9fff7" }}>{br(d)}</td><td style={sTd}>{can(user.role, "BE") ? (<button style={sBD} onClick={() => delB(b.id)}>Excluir</button>) : "---"}</td></tr>
          ); })}{budgets.length === 0 && (<tr><td style={sTd} colSpan={10}>Nenhuma.</td></tr>)}</tbody></table></div></div></div>
        )}

        {/* === AUDITORIA === */}
        {tab === "auditoria" && (
          <div><h1 style={{ marginTop: 0 }}>Auditoria</h1>
          <div style={{ marginBottom: 12 }}><button style={sB} onClick={exportAuditTXT}>Exportar Auditoria (TXT)</button></div>
          {!can(user.role, "AU") && (<div style={sC}>Sem permissao.</div>)}
          {can(user.role, "AU") && (<div style={sC}><div style={{ fontWeight: 950, marginBottom: 10 }}>Logs ({audit.length})</div>
          {audit.map((a) => (<div key={a.id} style={{ padding: 12, borderRadius: 10, background: "rgba(255,255,255,.06)", border: "1px solid rgba(255,255,255,.10)", marginBottom: 8 }}>
            <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}><b>{a.action}</b><span style={sSm}>{a.at}</span></div>
            <div style={sSm}>Usuario: {a.user}</div><div style={{ marginTop: 4 }}>{a.note}</div>
          </div>))}{audit.length === 0 && <div>Sem registros.</div>}</div>)}</div>
        )}

        {/* === CONFIG === */}
        {tab === "config" && (
          <div><h1 style={{ marginTop: 0 }}>Configuracoes</h1>
          {!can(user.role, "CF") && (<div style={sC}>Sem permissao.</div>)}
          {can(user.role, "CF") && (<>
            {/* USUARIOS */}
            <div style={sC}><h3 style={{ marginTop: 0 }}>Usuarios</h3>
            {us.map((u) => (<div key={u.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: 8, background: "rgba(255,255,255,.04)", borderRadius: 8, marginBottom: 6 }}><div><b>{u.nome}</b> - {u.email} <span style={sBa}>{u.role}</span> <span style={sSm}>(senha: {u.senha})</span></div><button style={sBD} onClick={() => delU(u.id)}>X</button></div>))}
            <div style={{ ...sG3, marginTop: 10, alignItems: "end" }}>
              <div><label style={sL}>Nome</label><input style={sI} value={cfU.nome} onChange={(e) => setCfU({ ...cfU, nome: e.target.value })} placeholder="Nome completo" /></div>
              <div><label style={sL}>Email</label><input style={sI} value={cfU.email} onChange={(e) => setCfU({ ...cfU, email: e.target.value })} placeholder="email@adeste.com.br" /></div>
              <div><label style={sL}>Perfil</label><select style={sI} value={cfU.role} onChange={(e) => setCfU({ ...cfU, role: e.target.value as Ro })}><option value={RL.AD}>{RL.AD}</option><option value={RL.MG}>{RL.MG}</option><option value={RL.RQ}>{RL.RQ}</option></select></div>
            </div>
            <div style={sG2}>
              <div><label style={sL}>Senha</label><input style={sI} value={cfU.senha} onChange={(e) => setCfU({ ...cfU, senha: e.target.value })} placeholder="123456" /></div>
              <div style={{ display: "flex", alignItems: "end" }}><button style={sB} onClick={addU}>Adicionar usuario</button></div>
            </div></div>

            {/* SITES */}
            <div style={sC}><h3 style={{ marginTop: 0 }}>Sites</h3>
            {sites.map((s) => (<div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: 8, background: "rgba(255,255,255,.04)", borderRadius: 8, marginBottom: 6 }}><div>{s.nome}</div><button style={sBD} onClick={() => delSi(s.id)}>X</button></div>))}
            <div style={{ display: "flex", gap: 10, marginTop: 10 }}><input style={sI} value={cfS} onChange={(e) => setCfS(e.target.value)} placeholder="Nome do site" /><button style={sB} onClick={addSi}>Adicionar</button></div></div>

            {/* CC */}
            <div style={sC}><h3 style={{ marginTop: 0 }}>Centros de Custo</h3>
            {ccs.map((c) => (<div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: 8, background: "rgba(255,255,255,.04)", borderRadius: 8, marginBottom: 6 }}><div>{c.codigo}</div><button style={sBD} onClick={() => delCc(c.id)}>X</button></div>))}
            <div style={{ display: "flex", gap: 10, marginTop: 10 }}><input style={sI} value={cfC} onChange={(e) => setCfC(e.target.value)} placeholder="CC-003" /><button style={sB} onClick={addCc}>Adicionar</button></div></div>

            {/* CONTAS */}
            <div style={sC}><h3 style={{ marginTop: 0 }}>Contas</h3>
            {contas.map((c) => (<div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: 8, background: "rgba(255,255,255,.04)", borderRadius: 8, marginBottom: 6 }}><div>{c.codigo} - {c.descricao}</div><button style={sBD} onClick={() => delCt(c.id)}>X</button></div>))}
            <div style={sG2}><div><label style={sL}>Codigo</label><input style={sI} value={cfCt.codigo} onChange={(e) => setCfCt({ ...cfCt, codigo: e.target.value })} placeholder="7100" /></div><div><label style={sL}>Descricao</label><input style={sI} value={cfCt.descricao} onChange={(e) => setCfCt({ ...cfCt, descricao: e.target.value })} placeholder="Consultoria" /></div></div>
            <button style={sB} onClick={addCt}>Adicionar</button></div>
          </>)}</div>
        )}
      </main>
    </div>
  );
}
