import { useState, useMemo } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell, Area, AreaChart
} from "recharts";

// ════════════════════════════════════════════════════════════════════════════
// TEMAS DE CORES — 3 perfis selecionáveis no canto superior direito
// ════════════════════════════════════════════════════════════════════════════
const TEMAS = {
  padrao: {
    nome: "Padrão",
    emoji: "🌙",
    bg: "#0a0a14",
    bgHeader: "linear-gradient(135deg,#1a1040 0%,#0a0a14 100%)",
    card: "#131325",
    cardAlt: "#0f0f1a",
    border: "#1e293b",
    borderSoft: "#334155",
    text: "#e2e8f0",
    textDim: "#94a3b8",
    textMute: "#64748b",
    textFaint: "#475569",
    accent: "#6366f1",
    accentSoft: "#a5b4fc",
    accentBg: "#1a1040",
    accentBorder: "#4338ca",
    green: "#34d399",
    red: "#f87171",
    amber: "#fbbf24",
    cyan: "#22d3ee",
    // paletas de ativos
    corFII: { MFII11:"#06b6d4",VGHF11:"#0891b2",RBRY11:"#0e7490",OIAG11:"#22d3ee",CPTS11:"#67e8f9",GARE11:"#2dd4bf",TRXF11:"#0d9488",RBRX11:"#14b8a6",IRIM11:"#5eead4",KNCR11:"#99f6e4",VISC11:"#34d399",XPML11:"#6ee7b7",MXRF11:"#a7f3d0" },
    corAcao: { BBAS3:"#6366f1",PETR4:"#8b5cf6",BBDC3:"#a78bfa",ITSA4:"#c084fc",BBSE3:"#e879f9",TAEE11:"#f472b6",BRSR6:"#818cf8",KLBN4:"#7c3aed",KLBN11:"#db2777",CXSE3:"#ec4899",WEGE3:"#be185d",CPLE3:"#9d174d",SANB3:"#fb7185" },
  },
  office: {
    nome: "Office",
    emoji: "💼",
    bg: "#1a1d21",
    bgHeader: "linear-gradient(135deg,#2d3748 0%,#1a1d21 100%)",
    card: "#22262b",
    cardAlt: "#1d2125",
    border: "#374151",
    borderSoft: "#4b5563",
    text: "#e5e7eb",
    textDim: "#9ca3af",
    textMute: "#6b7280",
    textFaint: "#4b5563",
    accent: "#10b981",
    accentSoft: "#6ee7b7",
    accentBg: "#1f3a32",
    accentBorder: "#047857",
    green: "#10b981",
    red: "#ef4444",
    amber: "#d97706",
    cyan: "#0ea5e9",
    corFII: { MFII11:"#0ea5e9",VGHF11:"#0284c7",RBRY11:"#0369a1",OIAG11:"#38bdf8",CPTS11:"#7dd3fc",GARE11:"#0891b2",TRXF11:"#075985",RBRX11:"#0e7490",IRIM11:"#22d3ee",KNCR11:"#67e8f9",VISC11:"#06b6d4",XPML11:"#a5f3fc",MXRF11:"#cffafe" },
    corAcao: { BBAS3:"#10b981",PETR4:"#059669",BBDC3:"#34d399",ITSA4:"#6ee7b7",BBSE3:"#047857",TAEE11:"#065f46",BRSR6:"#22c55e",KLBN4:"#16a34a",KLBN11:"#15803d",CXSE3:"#4ade80",WEGE3:"#84cc16",CPLE3:"#65a30d",SANB3:"#a3e635" },
  },
  conforto: {
    nome: "Conforto",
    emoji: "🎨",
    bg: "#1c1b22",
    bgHeader: "linear-gradient(135deg,#3b2f4a 0%,#1c1b22 100%)",
    card: "#26242e",
    cardAlt: "#201f27",
    border: "#3a3744",
    borderSoft: "#524d5e",
    text: "#ede9f0",
    textDim: "#b8b2c4",
    textMute: "#857d92",
    textFaint: "#5c5568",
    accent: "#c084fc",
    accentSoft: "#e9d5ff",
    accentBg: "#3b2f4a",
    accentBorder: "#9333ea",
    green: "#86efac",
    red: "#fca5a5",
    amber: "#fcd34d",
    cyan: "#7dd3fc",
    corFII: { MFII11:"#7dd3fc",VGHF11:"#93c5fd",RBRY11:"#a5b4fc",OIAG11:"#a7f3d0",CPTS11:"#bbf7d0",GARE11:"#99f6e4",TRXF11:"#5eead4",RBRX11:"#6ee7b7",IRIM11:"#86efac",KNCR11:"#bef264",VISC11:"#7dd3fc",XPML11:"#c4b5fd",MXRF11:"#ddd6fe" },
    corAcao: { BBAS3:"#c084fc",PETR4:"#d8b4fe",BBDC3:"#f0abfc",ITSA4:"#f5d0fe",BBSE3:"#fbcfe8",TAEE11:"#fda4af",BRSR6:"#fecaca",KLBN4:"#fed7aa",KLBN11:"#fde68a",CXSE3:"#fef08a",WEGE3:"#d9f99d",CPLE3:"#bef264",SANB3:"#f9a8d4" },
  },
};

// ════════════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════════════
function mesesIdx(mesesCal) {
  return mesesCal.map(m => m >= 7 ? m - 7 : m + 5).filter(i => i >= 0 && i < 12);
}
function labelMes(offset) {
  const meses = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  const totalMes = 6 + offset;
  return `${meses[totalMes % 12]}/${String(2026 + Math.floor(totalMes / 12)).slice(2)}`;
}
const fmt  = v => v.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
const fmtK = v => v >= 1000 ? `R$${(v/1000).toFixed(1)}k` : fmt(v);

const MESES_BASE = ["Jul/26","Ago/26","Set/26","Out/26","Nov/26","Dez/26","Jan/27","Fev/27","Mar/27","Abr/27","Mai/27","Jun/27"];

// ════════════════════════════════════════════════════════════════════════════
// DADOS INICIAIS DOS ATIVOS (editáveis na aba "Editar")
// ════════════════════════════════════════════════════════════════════════════
const ATIVOS_INICIAIS = [
  { ticker:"MFII11",nome:"Mérito Desenvolvimento",  cat:"FII", freq:"Mensal",    qtd:132,prov:1.06,precoMedio:52.69,cotacao:50.71,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"VGHF11",nome:"Valora Hedge Fund",        cat:"FII", freq:"Mensal",    qtd:333,prov:0.07,precoMedio:6.05, cotacao:5.99, meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"RBRY11",nome:"BR Credit",                cat:"FII", freq:"Mensal",    qtd:21, prov:1.00,precoMedio:91.23,cotacao:89.49,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"OIAG11",nome:"Ourinvest Agro",           cat:"FII", freq:"Mensal",    qtd:146,prov:0.08,precoMedio:8.57, cotacao:8.15, meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"CPTS11",nome:"Capitânia Securities",     cat:"FII", freq:"Mensal",    qtd:152,prov:0.09,precoMedio:7.42, cotacao:7.39, meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"GARE11",nome:"Guardian",                 cat:"FII", freq:"Mensal",    qtd:114,prov:0.08,precoMedio:8.38, cotacao:8.14, meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"TRXF11",nome:"TRX Real Estate",          cat:"FII", freq:"Mensal",    qtd:10, prov:0.85,precoMedio:91.46,cotacao:91.53,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"RBRX11",nome:"RBR Properties",           cat:"FII", freq:"Mensal",    qtd:112,prov:0.09,precoMedio:8.62, cotacao:8.11, meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"IRIM11",nome:"Iridium Recebíveis",       cat:"FII", freq:"Mensal",    qtd:12, prov:0.90,precoMedio:65.79,cotacao:65.95,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"KNCR11",nome:"Kinea Crédito Real",       cat:"FII", freq:"Mensal",    qtd:5,  prov:1.10,precoMedio:106.83,cotacao:107.28,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"VISC11",nome:"Vinci Shopping Centers",   cat:"FII", freq:"Mensal",    qtd:3,  prov:0.72,precoMedio:106.14,cotacao:103.86,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"XPML11",nome:"XP Malls",                 cat:"FII", freq:"Mensal",    qtd:2,  prov:0.72,precoMedio:107.44,cotacao:103.98,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"MXRF11",nome:"Maxi Renda",               cat:"FII", freq:"Mensal",    qtd:16, prov:0.10,precoMedio:9.71, cotacao:9.68, meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"BBAS3", nome:"Banco do Brasil",  cat:"Ação",freq:"Trimestral", qtd:291,prov:0.14,precoMedio:20.55,cotacao:19.86,meses:[3,6,9,12] },
  { ticker:"PETR4", nome:"Petrobras",        cat:"Ação",freq:"Trimestral", qtd:145,prov:0.47,precoMedio:40.75,cotacao:39.33,meses:[2,5,8,11] },
  { ticker:"BBDC3", nome:"Bradesco",         cat:"Ação",freq:"Mensal",     qtd:50, prov:0.02,precoMedio:15.73,cotacao:15.54,meses:[1,2,3,4,5,6,7,8,9,10,11,12] },
  { ticker:"ITSA4", nome:"Itaúsa",           cat:"Ação",freq:"Trimestral", qtd:116,prov:0.08,precoMedio:12.88,cotacao:13.03,meses:[1,4,7,10] },
  { ticker:"BBSE3", nome:"BB Seguridade",    cat:"Ação",freq:"Semestral",  qtd:9,  prov:0.40,precoMedio:37.47,cotacao:38.27,meses:[5,11] },
  { ticker:"TAEE11",nome:"Taesa",            cat:"Ação",freq:"Semestral",  qtd:6,  prov:0.90,precoMedio:41.68,cotacao:39.79,meses:[6,12] },
  { ticker:"BRSR6", nome:"Banrisul",         cat:"Ação",freq:"Trimestral", qtd:44, prov:0.32,precoMedio:15.76,cotacao:13.55,meses:[3,6,9,12] },
  { ticker:"KLBN4", nome:"Klabin PN",        cat:"Ação",freq:"Semestral",  qtd:332,prov:0.10,precoMedio:3.41, cotacao:3.37, meses:[4,10] },
  { ticker:"KLBN11",nome:"Klabin UNT",       cat:"Ação",freq:"Semestral",  qtd:14, prov:0.25,precoMedio:17.03,cotacao:16.81,meses:[4,10] },
  { ticker:"CXSE3", nome:"Caixa Seguridade", cat:"Ação",freq:"Semestral",  qtd:13, prov:0.35,precoMedio:17.66,cotacao:19.43,meses:[5,11] },
  { ticker:"WEGE3", nome:"Weg",              cat:"Ação",freq:"Semestral",  qtd:2,  prov:0.12,precoMedio:44.96,cotacao:45.71,meses:[4,10] },
  { ticker:"CPLE3", nome:"Copel",            cat:"Ação",freq:"Semestral",  qtd:3,  prov:0.22,precoMedio:15.63,cotacao:14.99,meses:[5,11] },
  { ticker:"SANB3", nome:"Santander",        cat:"Ação",freq:"Trimestral", qtd:102,prov:0.06,precoMedio:12.90,cotacao:12.86,meses:[3,6,9,12] },
];

// ════════════════════════════════════════════════════════════════════════════
// FUNÇÕES DE CÁLCULO (recebem a lista de ativos como parâmetro)
// ════════════════════════════════════════════════════════════════════════════
function corDe(ticker, cat, T) {
  return cat === "FII" ? (T.corFII[ticker] || T.cyan) : (T.corAcao[ticker] || T.accent);
}

function buildChart(ativos, filtro) {
  const lista = filtro==="TUDO" ? ativos : filtro==="FII" ? ativos.filter(a=>a.cat==="FII") : ativos.filter(a=>a.cat==="Ação");
  return MESES_BASE.map((mes,i) => {
    const e = { mes };
    let t = 0;
    lista.forEach(a => { const v = mesesIdx(a.meses).includes(i) ? +(a.prov*a.qtd).toFixed(2) : 0; e[a.ticker]=v; t+=v; });
    e._total = +t.toFixed(2);
    return e;
  });
}

function simular(ativos, regras, horizonte, aporte) {
  const estado = ativos.map(a => ({ ...a }));
  let acum = 0;
  return Array.from({ length: horizonte }, (_, m) => {
    let provBrutoMes = 0;
    const detalhes = [];
    estado.forEach(a => {
      if (mesesIdx(a.meses).includes(m % 12)) {
        const v = a.prov * a.qtd;
        provBrutoMes += v;
        detalhes.push({ ticker: a.ticker, cat: a.cat, val: +v.toFixed(2) });
      }
    });
    const caixaInicial = acum + provBrutoMes + aporte;
    let caixa = caixaInicial;
    const compras = [];
    regras.forEach(r => {
      if (r.pct <= 0) return;
      const ativo = estado.find(a => a.ticker === r.ticker);
      if (!ativo) return;
      const cot = r.cotacaoAlvo || ativo.cotacao;
      const cotas = Math.floor((caixaInicial * (r.pct / 100)) / cot);
      if (cotas > 0) {
        ativo.qtd += cotas;
        caixa -= cotas * cot;
        compras.push({ ticker: r.ticker, cotas, gasto: +(cotas*cot).toFixed(2) });
      }
    });
    acum = Math.max(caixa, 0);
    const patri = estado.reduce((s,a) => s + a.qtd*a.cotacao, 0);
    const provMedio = estado.reduce((s,a) => s + a.prov*a.qtd*a.meses.length/12, 0);
    return {
      mes: labelMes(m),
      provento: +provBrutoMes.toFixed(2),
      provMedio: +provMedio.toFixed(2),
      patrimonio: +patri.toFixed(2),
      caixa: +acum.toFixed(2),
      compras,
      detalhes: detalhes.sort((a,b)=>b.val-a.val),
    };
  });
}

// ════════════════════════════════════════════════════════════════════════════
// COMPONENTES VISUAIS
// ════════════════════════════════════════════════════════════════════════════
function TipBar({ active, payload, label, ativos, T }) {
  if (!active || !payload?.length) return null;
  const items = payload.filter(p=>p.value>0).sort((a,b)=>b.value-a.value);
  const findCat = tk => ativos.find(a=>a.ticker===tk)?.cat;
  return (
    <div style={{ background:T.card,border:`1px solid ${T.accentBorder}`,borderRadius:10,padding:"10px 14px",minWidth:200,maxHeight:300,overflowY:"auto",boxShadow:"0 8px 24px #0008" }}>
      <div style={{ fontSize:11,fontWeight:800,color:T.accentSoft,marginBottom:6,borderBottom:`1px solid ${T.border}`,paddingBottom:4 }}>{label}</div>
      {items.map(p=>{
        const c = corDe(p.dataKey, findCat(p.dataKey), T);
        return (
          <div key={p.dataKey} style={{ display:"flex",justifyContent:"space-between",gap:10,marginBottom:3,alignItems:"center" }}>
            <div style={{ display:"flex",alignItems:"center",gap:5 }}>
              <div style={{ width:8,height:8,borderRadius:2,background:c,flexShrink:0 }}/>
              <span style={{ fontSize:10,color:T.textDim }}>{p.dataKey}</span>
            </div>
            <span style={{ fontSize:11,fontWeight:700,color:c }}>{fmt(p.value)}</span>
          </div>
        );
      })}
      <div style={{ borderTop:`1px solid ${T.accentBorder}`,marginTop:6,paddingTop:6,fontSize:12,fontWeight:800,color:T.accentSoft,display:"flex",justifyContent:"space-between" }}>
        <span>Total</span><span>{fmt(items.reduce((s,p)=>s+p.value,0))}</span>
      </div>
    </div>
  );
}

function TipSim({ active, payload, label, T }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background:T.card,border:`1px solid ${T.accentBorder}`,borderRadius:10,padding:"10px 14px",minWidth:190,boxShadow:"0 8px 24px #0008" }}>
      <div style={{ fontSize:11,fontWeight:800,color:T.accentSoft,marginBottom:6,borderBottom:`1px solid ${T.border}`,paddingBottom:4 }}>{label}</div>
      {payload.map(p=>(
        <div key={p.name} style={{ display:"flex",justifyContent:"space-between",gap:14,marginBottom:3 }}>
          <div style={{ display:"flex",alignItems:"center",gap:5 }}>
            <div style={{ width:8,height:8,borderRadius:2,background:p.color }}/>
            <span style={{ fontSize:10,color:T.textDim }}>{p.name}</span>
          </div>
          <span style={{ fontSize:11,fontWeight:700,color:p.color }}>{fmt(p.value)}</span>
        </div>
      ))}
    </div>
  );
}

function Legenda({ ativos, T }) {
  return (
    <div style={{ display:"flex",flexWrap:"wrap",gap:"4px 10px",padding:"6px 8px 2px" }}>
      {ativos.map(a=>(
        <div key={a.ticker} style={{ display:"flex",alignItems:"center",gap:4 }}>
          <div style={{ width:8,height:8,borderRadius:2,background:corDe(a.ticker,a.cat,T) }}/>
          <span style={{ fontSize:9,color:T.textMute }}>{a.ticker}</span>
        </div>
      ))}
    </div>
  );
}

function DetalheMes({ ativos, idx, filtro, T }) {
  const lista = filtro==="TUDO" ? ativos : filtro==="FII" ? ativos.filter(a=>a.cat==="FII") : ativos.filter(a=>a.cat==="Ação");
  const pag = lista.filter(a=>mesesIdx(a.meses).includes(idx)).map(a=>({...a,total:+(a.prov*a.qtd).toFixed(2)})).sort((a,b)=>b.total-a.total);
  const total = pag.reduce((s,a)=>s+a.total,0);
  return (
    <div>
      <div style={{ display:"flex",alignItems:"baseline",gap:8,marginBottom:10 }}>
        <span style={{ fontSize:12,fontWeight:700,color:T.accentSoft }}>{MESES_BASE[idx]}</span>
        <span style={{ fontSize:20,fontWeight:800,color:T.text }}>{fmt(total)}</span>
      </div>
      {pag.length===0
        ? <div style={{ fontSize:12,color:T.textFaint,textAlign:"center",padding:"16px 0" }}>Nenhum provento neste mês</div>
        : pag.map(a=>{
          const c = corDe(a.ticker,a.cat,T);
          return (
            <div key={a.ticker} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",background:`${c}0d`,border:`1px solid ${c}33`,borderLeft:`3px solid ${c}`,borderRadius:8,padding:"8px 12px",marginBottom:5 }}>
              <div style={{ display:"flex",alignItems:"center",gap:7 }}>
                <div style={{ width:8,height:8,borderRadius:2,background:c }}/>
                <div>
                  <span style={{ fontSize:12,fontWeight:700,color:T.text }}>{a.ticker}</span>
                  <span style={{ fontSize:9,color:T.textFaint,background:T.border,padding:"1px 5px",borderRadius:4,marginLeft:6 }}>{a.freq}</span>
                  <div style={{ fontSize:10,color:T.textFaint,marginTop:1 }}>{a.nome} · {a.qtd}× R${a.prov.toFixed(2)}</div>
                </div>
              </div>
              <div style={{ fontSize:13,fontWeight:800,color:c }}>{fmt(a.total)}</div>
            </div>
          );
        })
      }
    </div>
  );
}

function Ranking({ ativos, T }) {
  const ranked = ativos.map(a=>({...a,ano:+(mesesIdx(a.meses).length*a.prov*a.qtd).toFixed(2)})).filter(a=>a.ano>0).sort((a,b)=>b.ano-a.ano);
  const total = ranked.reduce((s,a)=>s+a.ano,0); const max = ranked[0]?.ano||1;
  return (
    <div>
      <div style={{ fontSize:10,color:T.textFaint,textTransform:"uppercase",letterSpacing:1,marginBottom:12 }}>Ranking anual · {fmt(total)}</div>
      {ranked.map((a,i)=>{
        const c = corDe(a.ticker,a.cat,T);
        return (
          <div key={a.ticker} style={{ marginBottom:9 }}>
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3 }}>
              <div style={{ display:"flex",alignItems:"center",gap:7 }}>
                <span style={{ fontSize:10,color:T.textFaint,fontWeight:700,minWidth:18 }}>#{i+1}</span>
                <div style={{ width:9,height:9,borderRadius:2,background:c }}/>
                <span style={{ fontSize:12,fontWeight:700,color:T.text }}>{a.ticker}</span>
                <span style={{ fontSize:9,color:T.textFaint,background:T.border,padding:"1px 5px",borderRadius:4 }}>{a.freq}</span>
              </div>
              <span style={{ fontSize:12,fontWeight:800,color:c }}>{fmt(a.ano)}</span>
            </div>
            <div style={{ display:"flex",alignItems:"center",gap:6 }}>
              <div style={{ flex:1,height:5,background:T.border,borderRadius:4,overflow:"hidden" }}>
                <div style={{ height:"100%",width:`${(a.ano/max)*100}%`,background:c,borderRadius:4 }}/>
              </div>
              <span style={{ fontSize:9,color:T.textFaint,minWidth:32,textAlign:"right" }}>{((a.ano/total)*100).toFixed(1)}%</span>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ════════════════════════════════════════════════════════════════════════════
// ABA EDITAR ATIVOS
// ════════════════════════════════════════════════════════════════════════════
function EditarAtivos({ ativos, setAtivos, T }) {
  const [filtroCat, setFiltroCat] = useState("TODOS");
  const [busca, setBusca] = useState("");

  const lista = ativos.filter(a => {
    if (filtroCat === "FII" && a.cat !== "FII") return false;
    if (filtroCat === "Ação" && a.cat !== "Ação") return false;
    if (busca && !a.ticker.toLowerCase().includes(busca.toLowerCase()) && !a.nome.toLowerCase().includes(busca.toLowerCase())) return false;
    return true;
  });

  function atualizar(ticker, campo, valor) {
    setAtivos(prev => prev.map(a => a.ticker === ticker ? { ...a, [campo]: valor } : a));
  }

  function resetar() {
    if (window.confirm("Restaurar todos os valores originais? Suas edições serão perdidas.")) {
      setAtivos(ATIVOS_INICIAIS.map(a => ({ ...a })));
    }
  }

  const totalInvestido = ativos.reduce((s,a)=>s + a.qtd*a.precoMedio, 0);
  const totalAtual = ativos.reduce((s,a)=>s + a.qtd*a.cotacao, 0);
  const lucro = totalAtual - totalInvestido;

  const FREQS = ["Mensal","Trimestral","Semestral"];

  return (
    <div>
      {/* Resumo topo */}
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14 }}>
        {[
          {l:"Investido", v:fmtK(totalInvestido), c:T.textDim},
          {l:"Atual", v:fmtK(totalAtual), c:T.accentSoft},
          {l:"Resultado", v:`${lucro>=0?"+":""}${fmtK(lucro)}`, c:lucro>=0?T.green:T.red},
        ].map(k=>(
          <div key={k.l} style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:10,padding:"10px 8px" }}>
            <div style={{ fontSize:9,color:T.textFaint,marginBottom:2 }}>{k.l}</div>
            <div style={{ fontSize:13,fontWeight:800,color:k.c }}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* Busca + filtro */}
      <div style={{ display:"flex",gap:8,marginBottom:6,flexWrap:"wrap" }}>
        <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Buscar ativo..."
          style={{ flex:"1 1 140px",background:T.cardAlt,border:`1px solid ${T.border}`,borderRadius:8,color:T.text,padding:"8px 12px",fontSize:13 }}/>
      </div>
      <div style={{ display:"flex",gap:6,marginBottom:12 }}>
        {[["TODOS","Todos"],["FII","FIIs"],["Ação","Ações"]].map(([k,l])=>(
          <button key={k} onClick={()=>setFiltroCat(k)} style={{ flex:"1 1 0",padding:"7px 4px",borderRadius:8,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:filtroCat===k?T.accent:T.border,color:filtroCat===k?"#fff":T.textMute }}>{l}</button>
        ))}
      </div>

      {/* Lista editável */}
      {lista.map(a=>{
        const c = corDe(a.ticker,a.cat,T);
        const valorTotal = a.qtd*a.cotacao;
        const result = (a.cotacao-a.precoMedio)*a.qtd;
        return (
          <div key={a.ticker} style={{ background:T.card,border:`1px solid ${T.border}`,borderLeft:`3px solid ${c}`,borderRadius:10,padding:"12px",marginBottom:8 }}>
            {/* cabeçalho */}
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
              <div style={{ display:"flex",alignItems:"center",gap:7 }}>
                <div style={{ width:9,height:9,borderRadius:2,background:c }}/>
                <div>
                  <span style={{ fontSize:13,fontWeight:800,color:T.text }}>{a.ticker}</span>
                  <span style={{ fontSize:9,color:c,background:`${c}1a`,padding:"1px 6px",borderRadius:4,marginLeft:6 }}>{a.cat}</span>
                  <div style={{ fontSize:10,color:T.textFaint,marginTop:1 }}>{a.nome}</div>
                </div>
              </div>
              <div style={{ textAlign:"right" }}>
                <div style={{ fontSize:13,fontWeight:800,color:T.text }}>{fmt(valorTotal)}</div>
                <div style={{ fontSize:10,color:result>=0?T.green:T.red }}>{result>=0?"+":""}{fmt(result)}</div>
              </div>
            </div>

            {/* campos editáveis */}
            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
              <CampoEdit label="Quantidade" valor={a.qtd} step="1" onChange={v=>atualizar(a.ticker,"qtd",Math.max(0,Math.round(v)))} T={T}/>
              <CampoEdit label="Preço médio (R$)" valor={a.precoMedio} step="0.01" onChange={v=>atualizar(a.ticker,"precoMedio",Math.max(0,v))} T={T}/>
              <CampoEdit label="Cotação atual (R$)" valor={a.cotacao} step="0.01" onChange={v=>atualizar(a.ticker,"cotacao",Math.max(0,v))} T={T}/>
              <CampoEdit label="Provento/cota (R$)" valor={a.prov} step="0.01" onChange={v=>atualizar(a.ticker,"prov",Math.max(0,v))} T={T}/>
            </div>

            {/* frequência */}
            <div style={{ marginTop:8 }}>
              <div style={{ fontSize:9,color:T.textFaint,marginBottom:4 }}>Frequência de pagamento</div>
              <div style={{ display:"flex",gap:4 }}>
                {FREQS.map(f=>(
                  <button key={f} onClick={()=>{
                    atualizar(a.ticker,"freq",f);
                    // ajusta meses conforme frequência
                    const novosMeses = f==="Mensal" ? [1,2,3,4,5,6,7,8,9,10,11,12]
                      : f==="Trimestral" ? [3,6,9,12]
                      : [6,12];
                    atualizar(a.ticker,"meses",novosMeses);
                  }} style={{ flex:"1 1 0",padding:"5px 4px",borderRadius:6,border:"none",cursor:"pointer",fontSize:10,fontWeight:600,background:a.freq===f?c:T.border,color:a.freq===f?"#fff":T.textMute }}>{f}</button>
                ))}
              </div>
            </div>
          </div>
        );
      })}

      {/* Reset */}
      <button onClick={resetar} style={{ width:"100%",padding:"10px",borderRadius:8,border:`1px solid ${T.red}44`,background:`${T.red}11`,color:T.red,cursor:"pointer",fontSize:12,fontWeight:600,marginTop:6 }}>
        ↺ Restaurar valores originais
      </button>

      <div style={{ background:T.cardAlt,border:`1px dashed ${T.borderSoft}`,borderRadius:10,padding:"10px 12px",marginTop:12 }}>
        <div style={{ fontSize:10,color:T.textFaint,lineHeight:1.7 }}>
          ✏️ As alterações refletem na hora em todas as abas (Gráfico, Ranking e Cenário Futuro). Mudar a frequência ajusta automaticamente os meses de pagamento. As edições valem durante o uso do app.
        </div>
      </div>
    </div>
  );
}

function CampoEdit({ label, valor, step, onChange, T }) {
  return (
    <div>
      <div style={{ fontSize:9,color:T.textFaint,marginBottom:3 }}>{label}</div>
      <input type="number" step={step} value={valor}
        onChange={e=>onChange(parseFloat(e.target.value)||0)}
        style={{ width:"100%",background:T.cardAlt,border:`1px solid ${T.border}`,borderRadius:6,color:T.text,padding:"7px 8px",fontSize:13,fontWeight:600 }}/>
    </div>
  );
}

// ════════════════════════════════════════════════════════════════════════════
// ABA CENÁRIO FUTURO
// ════════════════════════════════════════════════════════════════════════════
function CenarioFuturo({ ativos, T }) {
  const PRESETS = [
    { id:"mfii",  emoji:"🏗️", label:"100% MFII11",   desc:"Mérito — paga todo mês",     regras:[{ticker:"MFII11",pct:100,cotacaoAlvo:50.71}] },
    { id:"kncr",  emoji:"💼", label:"100% KNCR11",   desc:"Kinea — CRI seguro",         regras:[{ticker:"KNCR11",pct:100,cotacaoAlvo:107.28}] },
    { id:"vghf",  emoji:"📦", label:"100% VGHF11",   desc:"Valora — mais cotas/R$",     regras:[{ticker:"VGHF11",pct:100,cotacaoAlvo:5.99}] },
    { id:"mix",   emoji:"⚖️", label:"Mix FIIs",       desc:"MFII 50%·VGHF 30%·KNCR 20%",regras:[{ticker:"MFII11",pct:50,cotacaoAlvo:50.71},{ticker:"VGHF11",pct:30,cotacaoAlvo:5.99},{ticker:"KNCR11",pct:20,cotacaoAlvo:107.28}] },
    { id:"custom",emoji:"⚙️", label:"Personalizado",  desc:"Configure você mesmo",       regras:[] },
  ];

  const [presetId,  setPresetId]  = useState("mfii");
  const [horizonte, setHorizonte] = useState(60);
  const [aporte,    setAporte]    = useState(0);
  const [modo,      setModo]      = useState("provento");
  const [mesSelSim, setMesSelSim] = useState(0);
  const [customR,   setCustomR]   = useState([
    { ticker:"MFII11",pct:60,cotacaoAlvo:50.71 },
    { ticker:"VGHF11",pct:40,cotacaoAlvo:5.99 },
  ]);

  const preset = PRESETS.find(p=>p.id===presetId);
  const regras = presetId==="custom" ? customR : preset.regras;
  const pctTotal = regras.reduce((s,r)=>s+r.pct,0);
  const PATRI_INICIAL = ativos.reduce((s,a)=>s+a.qtd*a.cotacao,0);

  const dados = useMemo(
    () => simular(ativos, regras, horizonte, aporte),
    [presetId, horizonte, aporte, JSON.stringify(customR), JSON.stringify(ativos)]
  );

  const provInicial = dados[0]?.provento || 0;
  const provMedioFinal = dados[dados.length-1]?.provMedio || 0;
  const patriFinal  = dados[dados.length-1]?.patrimonio || 0;
  const crescProv   = (dados[0]?.provMedio||0)>0 ? +((provMedioFinal/(dados[0]?.provMedio||1)-1)*100).toFixed(1) : 0;
  const crescPatri  = +((patriFinal/PATRI_INICIAL-1)*100).toFixed(1);

  function updR(i,f,v) { setCustomR(prev=>prev.map((r,j)=>j===i?{...r,[f]:v}:r)); }

  const HORIZ = [{v:12,l:"1a"},{v:24,l:"2a"},{v:36,l:"3a"},{v:60,l:"5a"},{v:120,l:"10a"}];
  const APORTES = [0,100,200,500,1000];
  const mSel = dados[Math.min(mesSelSim, dados.length-1)];

  return (
    <div>
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14 }}>
        {[
          {l:"Provento 1º mês", v:fmt(provInicial), c:T.textDim},
          {l:`Provento médio em ${horizonte>=12?`${horizonte/12}a`:horizonte+"m"}`, v:fmt(provMedioFinal), c:T.green},
          {l:"Crescimento proventos", v:`+${crescProv}%`, c:T.amber},
          {l:"Patrimônio projetado", v:fmtK(patriFinal), c:T.accentSoft},
        ].map(k=>(
          <div key={k.l} style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:10,padding:"10px 12px" }}>
            <div style={{ fontSize:9,color:T.textFaint,marginBottom:2 }}>{k.l}</div>
            <div style={{ fontSize:16,fontWeight:800,color:k.c }}>{k.v}</div>
            {k.l.includes("Patrimônio")&&<div style={{ fontSize:9,color:T.textFaint }}>+{crescPatri}% vs hoje</div>}
          </div>
        ))}
      </div>

      <div style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"14px",marginBottom:14 }}>
        <div style={{ fontSize:10,color:T.textFaint,textTransform:"uppercase",letterSpacing:1,marginBottom:12 }}>⚙️ Configuração</div>

        <div style={{ marginBottom:12 }}>
          <div style={{ fontSize:11,color:T.textMute,marginBottom:6 }}>Horizonte</div>
          <div style={{ display:"flex",gap:6 }}>
            {HORIZ.map(o=>(
              <button key={o.v} onClick={()=>{setHorizonte(o.v); setMesSelSim(0);}} style={{ flex:"1 1 0",padding:"6px 4px",borderRadius:8,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:horizonte===o.v?T.accent:T.border,color:horizonte===o.v?"#fff":T.textMute }}>{o.l}</button>
            ))}
          </div>
        </div>

        <div style={{ marginBottom:12 }}>
          <div style={{ fontSize:11,color:T.textMute,marginBottom:6 }}>Aporte mensal extra (além dos proventos)</div>
          <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
            {APORTES.map(v=>(
              <button key={v} onClick={()=>setAporte(v)} style={{ padding:"6px 10px",borderRadius:8,border:"none",cursor:"pointer",fontSize:10,fontWeight:600,background:aporte===v?T.green:T.border,color:aporte===v?"#fff":T.textMute }}>{v===0?"Sem aporte":`+${fmt(v)}`}</button>
            ))}
          </div>
        </div>

        <div>
          <div style={{ fontSize:11,color:T.textMute,marginBottom:6 }}>Estratégia de reinvestimento</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:presetId==="custom"?12:0 }}>
            {PRESETS.map(p=>{
              const sel=presetId===p.id;
              return (
                <button key={p.id} onClick={()=>{setPresetId(p.id); setMesSelSim(0);}} style={{ padding:"9px 10px",borderRadius:10,border:`2px solid ${sel?T.accent:T.border}`,background:sel?T.accentBg:T.cardAlt,cursor:"pointer",textAlign:"left" }}>
                  <div style={{ fontSize:11,fontWeight:700,color:sel?T.accentSoft:T.textMute }}>{p.emoji} {p.label}</div>
                  <div style={{ fontSize:9,color:T.textFaint,marginTop:2 }}>{p.desc}</div>
                </button>
              );
            })}
          </div>

          {presetId==="custom" && (
            <div>
              <div style={{ fontSize:10,color:T.textMute,marginBottom:8 }}>
                Regras · <span style={{ color:pctTotal===100?T.green:pctTotal>100?T.red:T.amber,fontWeight:700 }}>{pctTotal}% alocado {pctTotal!==100?"⚠️":"✓"}</span>
              </div>
              {customR.map((r,i)=>{
                const base = ativos.find(a=>a.ticker===r.ticker);
                return (
                  <div key={i} style={{ background:T.cardAlt,border:`1px solid ${T.border}`,borderRadius:8,padding:"10px",marginBottom:6 }}>
                    <div style={{ display:"flex",gap:6,alignItems:"center",flexWrap:"wrap" }}>
                      <select value={r.ticker} onChange={e=>{ const a=ativos.find(x=>x.ticker===e.target.value); updR(i,"ticker",e.target.value); if(a)updR(i,"cotacaoAlvo",a.cotacao); }} style={{ flex:"1 1 140px",background:T.border,border:`1px solid ${T.borderSoft}`,borderRadius:6,color:T.text,padding:"5px 8px",fontSize:11 }}>
                        {ativos.map(a=><option key={a.ticker} value={a.ticker}>{a.ticker} — {a.nome}</option>)}
                      </select>
                      <div style={{ display:"flex",alignItems:"center",gap:4 }}>
                        <input type="number" min="0" max="100" value={r.pct} onChange={e=>updR(i,"pct",+e.target.value)} style={{ width:50,background:T.border,border:`1px solid ${T.borderSoft}`,borderRadius:6,color:T.text,padding:"5px 6px",fontSize:11,textAlign:"center" }} />
                        <span style={{ fontSize:10,color:T.textFaint }}>%</span>
                      </div>
                      <div style={{ display:"flex",alignItems:"center",gap:4 }}>
                        <span style={{ fontSize:10,color:T.textFaint }}>R$</span>
                        <input type="number" min="0" step="0.01" value={r.cotacaoAlvo} onChange={e=>updR(i,"cotacaoAlvo",+e.target.value)} style={{ width:60,background:T.border,border:`1px solid ${T.borderSoft}`,borderRadius:6,color:T.text,padding:"5px 6px",fontSize:11,textAlign:"center" }} />
                      </div>
                      <button onClick={()=>setCustomR(p=>p.filter((_,j)=>j!==i))} style={{ background:`${T.red}33`,border:"none",borderRadius:6,color:T.red,padding:"5px 9px",cursor:"pointer",fontSize:11 }}>✕</button>
                    </div>
                    {base&&<div style={{ fontSize:9,color:T.textFaint,marginTop:4 }}>Prov: R${base.prov.toFixed(2)}/cota · DY {((base.prov/base.cotacao)*100).toFixed(2)}%/mês · {base.freq}</div>}
                  </div>
                );
              })}
              <button onClick={()=>setCustomR(p=>[...p,{ticker:"MXRF11",pct:0,cotacaoAlvo:9.68}])} style={{ width:"100%",padding:"7px",borderRadius:8,border:`1px dashed ${T.borderSoft}`,background:"transparent",color:T.textFaint,cursor:"pointer",fontSize:11,marginTop:4 }}>+ Adicionar ativo</button>
            </div>
          )}
        </div>
      </div>

      <div style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"14px 6px 10px",marginBottom:14 }}>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",paddingLeft:8,marginBottom:10 }}>
          <div style={{ fontSize:10,color:T.textFaint,textTransform:"uppercase",letterSpacing:1 }}>Evolução mês a mês</div>
          <div style={{ display:"flex",gap:4 }}>
            {[["provento","💰 Proventos"],["patrimonio","💼 Patrimônio"]].map(([id,l])=>(
              <button key={id} onClick={()=>setModo(id)} style={{ padding:"4px 8px",borderRadius:6,border:"none",cursor:"pointer",fontSize:9,fontWeight:600,background:modo===id?T.accent:T.border,color:modo===id?"#fff":T.textMute }}>{l}</button>
            ))}
          </div>
        </div>
        <div style={{ width:"100%", height:240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={dados} margin={{ top:8,right:12,left:0,bottom:0 }}
              onClick={(e)=>{ if(e && e.activeTooltipIndex!=null) setMesSelSim(e.activeTooltipIndex); }}>
              <defs>
                <linearGradient id="gradProv" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={T.green} stopOpacity={0.4}/>
                  <stop offset="100%" stopColor={T.green} stopOpacity={0.02}/>
                </linearGradient>
                <linearGradient id="gradPatri" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={T.accent} stopOpacity={0.4}/>
                  <stop offset="100%" stopColor={T.accent} stopOpacity={0.02}/>
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
              <XAxis dataKey="mes" tick={{ fontSize:8,fill:T.textMute }} axisLine={false} tickLine={false}
                interval={Math.max(Math.floor(dados.length/7)-1,0)} minTickGap={10}/>
              <YAxis tick={{ fontSize:8,fill:T.textMute }} axisLine={false} tickLine={false}
                width={42} domain={[0,'auto']}
                tickFormatter={v=>v>=1000?`${(v/1000).toFixed(0)}k`:`${Math.round(v)}`}/>
              <Tooltip content={<TipSim T={T} />}/>
              {modo==="provento" ? (
                <Area type="monotone" dataKey="provMedio" name="Provento médio/mês"
                  stroke={T.green} strokeWidth={2.5} fill="url(#gradProv)" dot={false} activeDot={{ r:4, fill:T.green }}/>
              ) : (
                <Area type="monotone" dataKey="patrimonio" name="Patrimônio total"
                  stroke={T.accent} strokeWidth={2.5} fill="url(#gradPatri)" dot={false} activeDot={{ r:4, fill:T.accent }}/>
              )}
            </AreaChart>
          </ResponsiveContainer>
        </div>

        <div style={{ display:"flex",gap:6,marginTop:10,paddingLeft:8,overflowX:"auto" }}>
          {[12,24,36,60,120].filter(m=>m<=horizonte).map(m=>{
            const d = dados[m-1]; if(!d) return null;
            return (
              <div key={m} style={{ flex:"0 0 auto",background:T.cardAlt,border:`1px solid ${T.border}`,borderRadius:8,padding:"5px 10px",textAlign:"center" }}>
                <div style={{ fontSize:8,color:T.textFaint,marginBottom:2 }}>{m>=12?`${m/12} ano${m/12>1?"s":""}`:m+"m"}</div>
                <div style={{ fontSize:11,fontWeight:800,color:T.green }}>{fmtK(d.provMedio)}</div>
                <div style={{ fontSize:8,color:T.textMute }}>/mês</div>
                <div style={{ fontSize:10,fontWeight:700,color:T.accent,marginTop:2 }}>{fmtK(d.patrimonio)}</div>
              </div>
            );
          })}
        </div>
      </div>

      <div style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"14px",marginBottom:14 }}>
        <div style={{ fontSize:10,color:T.textFaint,textTransform:"uppercase",letterSpacing:1,marginBottom:10 }}>
          📋 Fluxo de reinvestimento — toque numa barra do gráfico
        </div>
        <div style={{ display:"flex",gap:2,marginBottom:12,overflowX:"auto" }}>
          {dados.slice(0,Math.min(horizonte,24)).map((d,i)=>(
            <button key={i} onClick={()=>setMesSelSim(i)} style={{ flex:"0 0 auto",minWidth:54,padding:"5px 4px",borderRadius:6,border:`1px solid ${i===mesSelSim?T.accent:T.border}`,background:i===mesSelSim?T.accentBg:T.cardAlt,cursor:"pointer" }}>
              <div style={{ fontSize:8,color:i===mesSelSim?T.accentSoft:T.textFaint }}>{d.mes}</div>
              <div style={{ fontSize:9,fontWeight:700,color:i===mesSelSim?T.green:T.textMute }}>{fmtK(d.provento)}</div>
            </button>
          ))}
        </div>

        {mSel && (
          <div>
            <div style={{ display:"flex",gap:8,marginBottom:12,flexWrap:"wrap" }}>
              <div style={{ background:`${T.cyan}1a`,border:`1px solid ${T.cyan}55`,borderRadius:8,padding:"8px 12px",flex:"1 1 auto" }}>
                <div style={{ fontSize:9,color:T.cyan }}>Provento recebido</div>
                <div style={{ fontSize:15,fontWeight:800,color:T.cyan }}>{fmt(mSel.provento)}</div>
                {aporte>0 && <div style={{ fontSize:9,color:T.textFaint }}>+ {fmt(aporte)} aporte</div>}
              </div>
              <div style={{ background:T.accentBg,border:`1px solid ${T.accentBorder}`,borderRadius:8,padding:"8px 12px",flex:"1 1 auto" }}>
                <div style={{ fontSize:9,color:T.accentSoft }}>Patrimônio acumulado</div>
                <div style={{ fontSize:15,fontWeight:800,color:T.accentSoft }}>{fmtK(mSel.patrimonio)}</div>
                {mSel.caixa>0 && <div style={{ fontSize:9,color:T.textFaint }}>Caixa: {fmt(mSel.caixa)}</div>}
              </div>
            </div>
            <div style={{ fontSize:10,color:T.textMute,marginBottom:6 }}>🛒 Cotas compradas com o reinvestimento:</div>
            {mSel.compras.length===0
              ? <div style={{ fontSize:11,color:T.textFaint,padding:"8px 0",textAlign:"center" }}>Saldo insuficiente para cota inteira (acumulou {fmt(mSel.caixa)})</div>
              : mSel.compras.map(co=>{
                const c = corDe(co.ticker, ativos.find(a=>a.ticker===co.ticker)?.cat, T);
                return (
                  <div key={co.ticker} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",background:`${c}0d`,border:`1px solid ${c}33`,borderLeft:`3px solid ${c}`,borderRadius:8,padding:"8px 12px",marginBottom:5 }}>
                    <div style={{ display:"flex",alignItems:"center",gap:7 }}>
                      <div style={{ width:8,height:8,borderRadius:2,background:c }}/>
                      <div>
                        <span style={{ fontSize:12,fontWeight:700,color:T.text }}>{co.ticker}</span>
                        <div style={{ fontSize:10,color:T.textFaint }}>+{co.cotas} cota{co.cotas>1?"s":""}</div>
                      </div>
                    </div>
                    <div style={{ fontSize:12,fontWeight:700,color:c }}>{fmt(co.gasto)}</div>
                  </div>
                );
              })
            }
            {mSel.detalhes.length>0 && (
              <div style={{ marginTop:10 }}>
                <div style={{ fontSize:10,color:T.textMute,marginBottom:6 }}>💵 De onde veio o provento deste mês:</div>
                <div style={{ display:"flex",flexWrap:"wrap",gap:5 }}>
                  {mSel.detalhes.slice(0,12).map(d=>{
                    const c = corDe(d.ticker, d.cat, T);
                    return (
                      <div key={d.ticker} style={{ background:T.cardAlt,border:`1px solid ${c}33`,borderRadius:6,padding:"3px 8px",display:"flex",alignItems:"center",gap:5 }}>
                        <div style={{ width:6,height:6,borderRadius:1,background:c }}/>
                        <span style={{ fontSize:9,color:T.textDim }}>{d.ticker}</span>
                        <span style={{ fontSize:9,fontWeight:700,color:c }}>{fmt(d.val)}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      <div style={{ background:T.cardAlt,border:`1px dashed ${T.borderSoft}`,borderRadius:10,padding:"10px 12px" }}>
        <div style={{ fontSize:10,color:T.textFaint,lineHeight:1.7 }}>
          🤖 <strong style={{ color:T.textMute }}>Como funciona:</strong> Cada mês recebe os proventos das cotas que você já tem (incluindo as compradas antes — efeito bola de neve). Esse valor + aporte compra novas cotas inteiras. A sobra acumula. Cotações e proventos/cota mantidos constantes. Sem IR/taxas.
        </div>
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════════════════════
// APP PRINCIPAL
// ════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [temaId, setTemaId] = useState("padrao");
  const [ativos, setAtivos] = useState(ATIVOS_INICIAIS.map(a=>({...a})));
  const [filtro, setFiltro] = useState("TUDO");
  const [mesSel, setMesSel] = useState(0);
  const [aba,    setAba]    = useState("grafico");

  const T = TEMAS[temaId];

  const listaFiltrada = filtro==="TUDO"?ativos:filtro==="FII"?ativos.filter(a=>a.cat==="FII"):ativos.filter(a=>a.cat==="Ação");
  const chartData = useMemo(()=>buildChart(ativos, filtro),[ativos, filtro]);
  const totalAnual = chartData.reduce((s,d)=>s+d._total,0);
  const mediaMes   = totalAnual/12;
  const maxMes     = Math.max(...chartData.map(d=>d._total), 0);
  const positivos  = chartData.filter(d=>d._total>0);
  const minMes     = positivos.length ? Math.min(...positivos.map(d=>d._total)) : 0;
  const totFII  = useMemo(()=>buildChart(ativos,"FII").reduce((s,d)=>s+d._total,0),[ativos]);
  const totAcao = useMemo(()=>buildChart(ativos,"Ação").reduce((s,d)=>s+d._total,0),[ativos]);

  const FILTROS = [
    {id:"TUDO",emoji:"📊",label:"Panorama", sub:"Ações + FIIs",cor:T.accentSoft},
    {id:"FII", emoji:"🏢",label:"Só FIIs",  sub:`${ativos.filter(a=>a.cat==="FII").length} fundos`, cor:T.cyan},
    {id:"Ação",emoji:"📈",label:"Só Ações", sub:`${ativos.filter(a=>a.cat==="Ação").length} ações`, cor:T.accent},
  ];
  const ABAS = [
    {id:"grafico", label:"📈 Gráfico"},
    {id:"ranking", label:"🏆 Ranking"},
    {id:"cenario", label:"🤖 Cenário"},
    {id:"editar",  label:"✏️ Editar"},
  ];

  return (
    <div style={{ background:T.bg,minHeight:"100vh",color:T.text,fontFamily:"'Inter',system-ui,sans-serif",paddingBottom:48,transition:"background 0.3s" }}>

      {/* HEADER */}
      <div style={{ background:T.bgHeader,padding:"16px 16px 16px",borderBottom:`1px solid ${T.border}` }}>
        {/* linha superior: título + seletor de tema */}
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10 }}>
          <div style={{ fontSize:10,color:T.accent,letterSpacing:2,textTransform:"uppercase",paddingTop:6 }}>💰 Carteira Proventos</div>
          {/* SELETOR DE TEMAS */}
          <div style={{ display:"flex",gap:3,background:T.cardAlt,padding:3,borderRadius:10,border:`1px solid ${T.border}` }}>
            {Object.entries(TEMAS).map(([id,tema])=>(
              <button key={id} onClick={()=>setTemaId(id)} title={tema.nome} style={{
                width:34,height:30,borderRadius:7,border:temaId===id?`2px solid ${T.accent}`:"2px solid transparent",
                background:temaId===id?T.accentBg:"transparent",cursor:"pointer",fontSize:15,padding:0,
                display:"flex",alignItems:"center",justifyContent:"center",transition:"all 0.2s"
              }}>{tema.emoji}</button>
            ))}
          </div>
        </div>
        <div style={{ fontSize:26,fontWeight:800,color:T.text,letterSpacing:-1 }}>{fmt(totalAnual)}<span style={{ fontSize:12,color:T.textFaint,fontWeight:400,marginLeft:8 }}>/ ano base</span></div>
        <div style={{ display:"flex",gap:8,marginTop:10,flexWrap:"wrap" }}>
          {[
            {l:"Média/mês",v:fmt(mediaMes),c:T.accentSoft},
            {l:"Maior mês",v:fmt(maxMes),  c:T.green},
            {l:"Menor mês",v:fmt(minMes),  c:T.red},
            {l:"FIIs/ano", v:fmt(totFII),  c:T.cyan},
            {l:"Ações/ano",v:fmt(totAcao), c:T.accent},
          ].map(x=>(
            <div key={x.l} style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:8,padding:"5px 10px" }}>
              <div style={{ fontSize:9,color:T.textFaint }}>{x.l}</div>
              <div style={{ fontSize:11,fontWeight:700,color:x.c }}>{x.v}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ padding:"14px" }}>
        {/* FILTROS — ocultos no cenário e editar */}
        {(aba==="grafico"||aba==="ranking") && (
          <div style={{ marginBottom:12 }}>
            <div style={{ fontSize:10,color:T.textFaint,textTransform:"uppercase",letterSpacing:1,marginBottom:8 }}>Modo de visualização</div>
            <div style={{ display:"flex",gap:8 }}>
              {FILTROS.map(f=>{
                const sel=filtro===f.id;
                return (
                  <button key={f.id} onClick={()=>{setFiltro(f.id);setMesSel(0);}} style={{ flex:"1 1 0",padding:"8px 4px",borderRadius:12,border:`2px solid ${sel?f.cor:T.border}`,background:sel?`${f.cor}14`:T.card,cursor:"pointer",textAlign:"center" }}>
                    <div style={{ fontSize:16,marginBottom:2 }}>{f.emoji}</div>
                    <div style={{ fontSize:11,fontWeight:700,color:sel?f.cor:T.textMute }}>{f.label}</div>
                    <div style={{ fontSize:9,color:T.textFaint }}>{f.sub}</div>
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* ABAS */}
        <div style={{ display:"flex",gap:4,marginBottom:12,overflowX:"auto" }}>
          {ABAS.map(a=>(
            <button key={a.id} onClick={()=>setAba(a.id)} style={{ flex:"1 1 0",padding:"8px 8px",borderRadius:8,border:"none",cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap",background:aba===a.id?(a.id==="cenario"?T.green:a.id==="editar"?T.amber:T.accent):T.card,color:aba===a.id?(a.id==="editar"?"#1a1d21":"#fff"):T.textMute }}>{a.label}</button>
          ))}
        </div>

        {/* ABA GRÁFICO */}
        {aba==="grafico" && (
          <>
            <div style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:14,padding:"14px 6px 8px",marginBottom:14 }}>
              <div style={{ paddingLeft:8,marginBottom:6,fontSize:10,color:T.textFaint,textTransform:"uppercase",letterSpacing:1 }}>
                {filtro==="TUDO"?"Todos os ativos — cada cor = um ativo":filtro==="FII"?"FIIs":"Ações"}
              </div>
              <ResponsiveContainer width="100%" height={230}>
                <BarChart data={chartData} margin={{ top:4,right:6,left:0,bottom:0 }} onClick={({activeTooltipIndex})=>{ if(activeTooltipIndex!=null)setMesSel(activeTooltipIndex); }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
                  <XAxis dataKey="mes" tick={{ fontSize:9,fill:T.textMute }} axisLine={false} tickLine={false}/>
                  <YAxis tick={{ fontSize:9,fill:T.textMute }} axisLine={false} tickLine={false} tickFormatter={v=>v===0?"":v>=1000?`${(v/1000).toFixed(1)}k`:`${v}`} width={32}/>
                  <Tooltip content={<TipBar ativos={ativos} T={T} />}/>
                  {listaFiltrada.map((a,ai)=>(
                    <Bar key={a.ticker} dataKey={a.ticker} stackId="s" fill={corDe(a.ticker,a.cat,T)} radius={ai===listaFiltrada.length-1?[4,4,0,0]:[0,0,0,0]}>
                      {chartData.map((_,ci)=><Cell key={ci} fill={corDe(a.ticker,a.cat,T)} opacity={ci===mesSel?1:0.72} stroke={ci===mesSel?T.text:"none"} strokeWidth={ci===mesSel?0.5:0}/>)}
                    </Bar>
                  ))}
                </BarChart>
              </ResponsiveContainer>
              <Legenda ativos={listaFiltrada} T={T}/>
              <div style={{ display:"flex",gap:2,marginTop:8,paddingLeft:4,paddingRight:4,overflowX:"auto" }}>
                {chartData.map((d,i)=>{
                  const pct=maxMes>0?d._total/maxMes:0; const sel=i===mesSel;
                  return (
                    <div key={i} onClick={()=>setMesSel(i)} style={{ flex:"0 0 auto",minWidth:40,textAlign:"center",cursor:"pointer",background:sel?T.accentBg:T.cardAlt,border:`1px solid ${sel?T.accent:T.border}`,borderRadius:8,padding:"4px" }}>
                      <div style={{ fontSize:8,color:sel?T.accentSoft:T.textFaint,marginBottom:2 }}>{d.mes.split("/")[0]}</div>
                      <div style={{ width:22,height:22,margin:"0 auto 2px",display:"flex",alignItems:"flex-end",justifyContent:"center" }}>
                        <div style={{ width:14,height:`${Math.max(pct*22,2)}px`,background:sel?T.accent:T.borderSoft,borderRadius:"2px 2px 0 0" }}/>
                      </div>
                      <div style={{ fontSize:9,fontWeight:700,color:sel?T.text:T.textFaint }}>{d._total>=1000?`${(d._total/1000).toFixed(1)}k`:d._total>0?`${Math.round(d._total)}`:"—"}</div>
                    </div>
                  );
                })}
              </div>
            </div>
            <div style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:14,padding:"14px" }}>
              <div style={{ fontSize:10,color:T.textFaint,textTransform:"uppercase",letterSpacing:1,marginBottom:10 }}>Detalhes do mês · toque em uma barra</div>
              <DetalheMes ativos={ativos} idx={mesSel} filtro={filtro} T={T}/>
            </div>
          </>
        )}

        {/* ABA RANKING */}
        {aba==="ranking" && (
          <div style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:14,padding:"14px" }}>
            <Ranking ativos={listaFiltrada} T={T}/>
          </div>
        )}

        {/* ABA CENÁRIO */}
        {aba==="cenario" && <CenarioFuturo ativos={ativos} T={T}/>}

        {/* ABA EDITAR */}
        {aba==="editar" && <EditarAtivos ativos={ativos} setAtivos={setAtivos} T={T}/>}

        {(aba==="grafico"||aba==="ranking") && (
          <div style={{ background:T.cardAlt,border:`1px dashed ${T.borderSoft}`,borderRadius:10,padding:"10px 12px",marginTop:14 }}>
            <div style={{ fontSize:10,color:T.textFaint,lineHeight:1.7 }}>📌 FIIs pagam mensalmente. Ações seguem calendário histórico. Valores brutos — JCP têm IR 15%; FIIs isentos para PF.</div>
          </div>
        )}
      </div>
    </div>
  );
}
