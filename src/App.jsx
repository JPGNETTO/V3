import React, { useState, useMemo, useRef, useLayoutEffect, useEffect } from "react";
// Motion (ex-Framer Motion): animações aceleradas por hardware, fluidas no APK e no desktop.
// Respeita a preferência de "movimento reduzido" do sistema por padrão.
import { motion, AnimatePresence } from "framer-motion";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, Cell, Area, AreaChart, ReferenceLine, ReferenceDot,
  PieChart, Pie
} from "recharts";

// ════════════════════════════════════════════════════════════════════════════
// LEITOR DA PLANILHA DA B3 (Movimentação) — calcula posição e preço médio
// ════════════════════════════════════════════════════════════════════════════
// Tipos de movimento que NÃO mexem na quantidade (são proventos/atualizações)
const B3_MOV_IGNORAR = new Set(["Rendimento","Dividendo","Juros Sobre Capital Próprio","Atualização"]);
// Tipos de movimento que SÃO proventos recebidos (entram como "realizado")
const B3_MOV_PROVENTO = new Set(["Rendimento","Dividendo","Juros Sobre Capital Próprio","Juros","Amortização","Rendimento Tributado","Leilão de Fração"]);
function b3Num(v) {
  if (v == null || v === "") return NaN;
  if (typeof v === "number") return v;
  let s = String(v).trim();
  // formato de cercadura da B3: "-R$ 0.15-" (hífens dos dois lados = visual, não sinal)
  if (s.startsWith("-") && s.endsWith("-")) s = s.slice(1, -1);
  s = s.replace(/[R$\s]/g,"");
  if (s.endsWith("-")) s = s.slice(0, -1);
  if (s.includes(",")) s = s.replace(/\./g,"").replace(",",".");
  return parseFloat(s);
}
function b3TickerDeProduto(produto) {
  const p = String(produto||"").trim();
  if (p.includes(" - ")) {
    const t = p.split(" - ")[0].trim().toUpperCase();
    if (/^[A-Z0-9]{4,6}$/.test(t)) return t;
  }
  if (/^tesouro/i.test(p)) return p.split(" - ")[0].trim(); // Tesouro: usa o nome
  return p.split(" - ")[0].trim().toUpperCase();
}
// Converte data da B3 ("dd/mm/aaaa") para ISO "aaaa-mm-dd"
function b3Data(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  const s = String(v).trim();
  const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  const m2 = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return m2[0];
  return null;
}

// Recebe as linhas (array de objetos) da aba Movimentação e devolve posições
function parseB3Movimentacao(linhas) {
  const pos = {};
  let lidas = 0, ignoradas = 0;
  for (const r of linhas) {
    const mov = String(r["Movimentação"] ?? r["Movimentacao"] ?? "").trim();
    const es  = String(r["Entrada/Saída"] ?? r["Entrada/Saida"] ?? "").trim();
    const prod = r["Produto"];
    const data = b3Data(r["Data"] ?? r["Data do Negócio"] ?? r["Data Negócio"]);
    if (!prod) continue;
    if (B3_MOV_IGNORAR.has(mov)) { ignoradas++; continue; }
    const qtd = b3Num(r["Quantidade"]);
    if (!isFinite(qtd) || qtd === 0) continue;
    let preco = b3Num(r["Preço unitário"] ?? r["Preco unitario"]);
    if (!isFinite(preco)) preco = null;
    const tk = b3TickerDeProduto(prod);
    if (!tk) continue;
    if (!pos[tk]) pos[tk] = { qtd:0, custo:0, qc:0, primeiraCompra:null, ultimaCompra:null, movs:[] };
    if (es === "Credito" || es === "Crédito") {
      pos[tk].qtd += qtd;
      if (preco) { pos[tk].custo += qtd*preco; pos[tk].qc += qtd; }
      if (data) {
        if (!pos[tk].primeiraCompra || data < pos[tk].primeiraCompra) pos[tk].primeiraCompra = data;
        if (!pos[tk].ultimaCompra || data > pos[tk].ultimaCompra) pos[tk].ultimaCompra = data;
        pos[tk].movs.push({ data, qtd:+qtd, preco:preco||0, tipo:"compra" });
      }
    } else if (es === "Debito" || es === "Débito") {
      pos[tk].qtd -= qtd;
      if (data) pos[tk].movs.push({ data, qtd:-qtd, preco:preco||0, tipo:"venda" });
    }
    lidas++;
  }
  const itens = [];
  for (const tk in pos) {
    const p = pos[tk];
    if (Math.abs(p.qtd) < 0.0001) continue; // posição zerada (vendeu tudo)
    const ehTesouro = /^tesouro/i.test(tk);
    const qtdFinal = ehTesouro ? +p.qtd.toFixed(4) : Math.round(p.qtd); // Tesouro tem frações
    if (qtdFinal === 0) continue;
    const pm = p.qc > 0 ? +(p.custo/p.qc).toFixed(2) : null;
    p.movs.sort((a,b)=>a.data<b.data?-1:1);
    itens.push({ ticker:tk, qtd:qtdFinal, precoMedio:pm, cotacao:null, dataCompra:p.primeiraCompra, ultimaCompra:p.ultimaCompra, movimentacoes:p.movs });
  }
  itens.sort((a,b)=>a.ticker.localeCompare(b.ticker));
  return { itens, lidas, ignoradas };
}

// ════════════════════════════════════════════════════════════════════════════
// REGRA FISCAL DOS PROVENTOS (Brasil) — calcula o LÍQUIDO real automaticamente
// A planilha da B3 traz o valor creditado, mas não discrimina o imposto.
// Aplicamos a regra oficial por tipo de provento:
//  • Dividendo................ isento de IR para pessoa física (líquido = bruto)
//  • Rendimento (FII)......... isento para PF que cumpre os requisitos (líquido = bruto)
//  • JCP...................... IR retido na fonte: 15% até 2025, 17,5% a partir de 2026
//  • Amortização.............. devolução de capital, não é renda (líquido = bruto)
// Obs.: a B3 credita o JCP já líquido em muitos casos; por isso guardamos os dois
// valores (bruto estimado e líquido creditado) e sinalizamos o imposto retido.
// ════════════════════════════════════════════════════════════════════════════
function aliquotaJCP(dataISO) {
  // A partir de 01/01/2026 a alíquota do JCP passou de 15% para 17,5%
  const ano = parseInt(String(dataISO||"").slice(0,4), 10) || new Date().getFullYear();
  return ano >= 2026 ? 0.175 : 0.15;
}
function fiscalProvento(tipo, valorCreditado, dataISO) {
  const t = String(tipo||"").toLowerCase();
  const ehJCP = t.includes("juros"); // "Juros Sobre Capital Próprio" / "Juros"
  if (!ehJCP) {
    // Dividendo, Rendimento de FII e Amortização: sem IR na fonte para PF
    return { bruto:+valorCreditado.toFixed(2), imposto:0, liquido:+valorCreditado.toFixed(2), tributavel:false, aliquota:0 };
  }
  // JCP: o valor creditado na conta já vem LÍQUIDO (IR retido na fonte pela empresa).
  // Reconstituímos o bruto para transparência: bruto = liquido / (1 - aliquota)
  const aliq = aliquotaJCP(dataISO);
  const liquido = +valorCreditado.toFixed(2);
  const bruto = +(liquido / (1 - aliq)).toFixed(2);
  const imposto = +(bruto - liquido).toFixed(2);
  return { bruto, imposto, liquido, tributavel:true, aliquota:aliq };
}

// Extrai os PROVENTOS RECEBIDOS (Rendimento/Dividendo/JCP/Amortização) da planilha B3.
// Retorna { porMes, total (líquido), totalBruto, totalImposto, registros: [{data, ticker, valor, bruto, imposto, tipo}] }
function parseProventosRecebidos(linhas) {
  const porMes = {};       // líquido por mês (o que realmente entrou na conta)
  const porMesBruto = {};  // bruto por mês (antes do IR)
  const registros = [];
  let total = 0, totalBruto = 0, totalImposto = 0;
  for (const r of linhas) {
    const mov = String(r["Movimentação"] ?? r["Movimentacao"] ?? "").trim();
    if (!B3_MOV_PROVENTO.has(mov)) continue;
    const es = String(r["Entrada/Saída"] ?? r["Entrada/Saida"] ?? "").trim();
    if (es && !(es === "Credito" || es === "Crédito")) continue; // só entradas
    const valor = b3Num(r["Valor da Operação"] ?? r["Valor"] ?? r["Valor da Operacao"]);
    if (!isFinite(valor) || valor <= 0) continue;
    const data = b3Data(r["Data"] ?? r["Data do Negócio"]);
    if (!data) continue;
    // quantidade e valor unitário do provento (colunas da planilha)
    const qtdReg = b3Num(r["Quantidade"]);
    const unitReg = b3Num(r["Preço unitário"] ?? r["Preço Unitário"] ?? r["Preco unitario"] ?? r["Preco Unitario"]);
    // aplica a regra fiscal real por tipo de provento
    const f = fiscalProvento(mov, valor, data);
    const mesKey = data.slice(0,7); // YYYY-MM
    porMes[mesKey]      = +((porMes[mesKey] || 0) + f.liquido).toFixed(2);
    porMesBruto[mesKey] = +((porMesBruto[mesKey] || 0) + f.bruto).toFixed(2);
    total += f.liquido; totalBruto += f.bruto; totalImposto += f.imposto;
    registros.push({
      data, ticker: b3TickerDeProduto(r["Produto"]) || "—", tipo: mov,
      valor: f.liquido,           // líquido (o que caiu na conta)
      bruto: f.bruto, imposto: f.imposto, aliquota: f.aliquota, tributavel: f.tributavel,
      qtd: (isFinite(qtdReg) && qtdReg > 0) ? qtdReg : null,
      unit: (isFinite(unitReg) && unitReg > 0) ? +unitReg.toFixed(4) : null,
    });
  }
  registros.sort((a,b)=> a.data < b.data ? 1 : -1); // mais recentes primeiro
  return {
    porMes, porMesBruto, registros,
    total: +total.toFixed(2),
    totalBruto: +totalBruto.toFixed(2),
    totalImposto: +totalImposto.toFixed(2),
  };
}

// ════════════════════════════════════════════════════════════════════════════
// CONEXÃO COM O SERVIDOR — auto-detect (Tailscale → Local → localhost)
// ════════════════════════════════════════════════════════════════════════════
const SERVIDORES = [
  { nome:"Tailscale", url:"http://100.100.195.84:4000" },
  { nome:"Local",     url:"http://192.168.1.17:4000" },
  { nome:"localhost", url:"http://localhost:4000" },
];
async function detectarServidor(aoTestar) {
  for (const s of SERVIDORES) {
    try {
      const ctrl = new AbortController();
      const t = setTimeout(()=>ctrl.abort(), 3000);
      const r = await fetch(`${s.url}/health`, { signal: ctrl.signal });
      clearTimeout(t);
      if (r.ok) {
        const d = await r.json().catch(()=>({}));
        aoTestar && aoTestar(s, true);
        if (d.status === "ok" || r.ok) return s;
      } else { aoTestar && aoTestar(s, false); }
    } catch(e) { aoTestar && aoTestar(s, false); }
  }
  return null;
}

// ════════════════════════════════════════════════════════════════════════════
// AÇÕES DA IA (Camada 1) — comandos SEGUROS que não quebram a estrutura base.
// A IA emite [[ACAO]]{"tipo":"tema","valor":"padrao"}[[/ACAO]] e o app aplica
// numa camada acima (preview → aplicar/cancelar). A base fica sempre intacta.
// ════════════════════════════════════════════════════════════════════════════
// Elementos que a IA pode estilizar (whitelist — protege a base do app)
const ELEMENTOS_ESTILO = {
  // ─ Painel · cabeçalho ─
  saudacao:            { nome:"Saudação (Boa noite!)",        grupo:"Cabeçalho", base:24 },
  dataHoje:            { nome:"Data de hoje",                 grupo:"Cabeçalho", base:11 },
  destaqueDia:         { nome:"Card de destaque do dia",      grupo:"Cabeçalho", base:11, caixa:true },
  headerApp:           { nome:"Barra do topo do app",         grupo:"Cabeçalho", base:13, caixa:true },
  logoApp:             { nome:"Logo/nome no topo",            grupo:"Cabeçalho", base:13 },
  // ─ Valores principais ─
  dividendosHeader:    { nome:"Rótulo 'Dividendos do ano'",   grupo:"Valores", base:11 },
  dividendosValor:     { nome:"Valor dos dividendos",         grupo:"Valores", base:26 },
  dividendosCard:      { nome:"Card dos dividendos",          grupo:"Valores", base:13, caixa:true },
  patrimonioLabel:     { nome:"Rótulo 'Patrimônio total'",    grupo:"Valores", base:10 },
  patrimonioValor:     { nome:"Valor do patrimônio",          grupo:"Valores", base:26 },
  patrimonioCard:      { nome:"Card do patrimônio",           grupo:"Valores", base:13, caixa:true },
  // ─ KPIs ─
  kpiCarrossel:        { nome:"Carrossel de KPIs",            grupo:"KPIs", base:13, caixa:true },
  kpiCard:             { nome:"Cards dos KPIs",               grupo:"KPIs", base:12, caixa:true },
  kpiValor:            { nome:"Números dos KPIs",             grupo:"KPIs", base:16 },
  kpiLabel:            { nome:"Rótulos dos KPIs",             grupo:"KPIs", base:9 },
  // ─ Metas ─
  metaCard:            { nome:"Card da meta",                 grupo:"Metas", base:13, caixa:true },
  tituloMetaProventos: { nome:"Título 'Meta de proventos'",   grupo:"Metas", base:13 },
  metaValorAtual:      { nome:"Valor atual da meta",          grupo:"Metas", base:22 },
  metaAlvo:            { nome:"Valor alvo da meta",           grupo:"Metas", base:11 },
  metaBarra:           { nome:"Barra de progresso da meta",   grupo:"Metas", base:13, caixa:true },
  // ─ Blocos do painel ─
  tituloProx3Meses:    { nome:"Título 'Proventos 3 meses'",   grupo:"Blocos", base:11 },
  cardProx3Meses:      { nome:"Cards dos 3 meses",            grupo:"Blocos", base:13, caixa:true },
  tituloGraficoMensal: { nome:"Título 'Proventos mês a mês'", grupo:"Blocos", base:11 },
  cardGraficoMensal:   { nome:"Card do gráfico mensal",       grupo:"Blocos", base:13, caixa:true },
  tituloProjecao:      { nome:"Título da Projeção",           grupo:"Blocos", base:13 },
  cardProjecao:        { nome:"Card da Projeção",             grupo:"Blocos", base:13, caixa:true },
  tituloPrevistoRealizado:{ nome:"Título 'Previsto vs Realizado'", grupo:"Blocos", base:13 },
  cardPrevistoRealizado:{ nome:"Card do Previsto vs Realizado", grupo:"Blocos", base:13, caixa:true },
  // ─ Listas e ativos ─
  listaAtivoItem:      { nome:"Itens da lista de ativos",     grupo:"Listas", base:13, caixa:true },
  listaAtivoTicker:    { nome:"Ticker do ativo",              grupo:"Listas", base:13 },
  listaAtivoValor:     { nome:"Valor do ativo",               grupo:"Listas", base:13 },
  // ─ Navegação ─
  menuFlutuante:       { nome:"Menu flutuante inferior",      grupo:"Navegação", base:13, caixa:true },
  menuBotao:           { nome:"Botões do menu",               grupo:"Navegação", base:8 },
  abasInternas:        { nome:"Mini-abas do painel",          grupo:"Navegação", base:11, caixa:true },
  // ─ Chat ─
  chatBalaoIA:         { nome:"Balão de resposta da IA",      grupo:"Chat", base:13, caixa:true },
  chatBalaoUsuario:    { nome:"Balão da sua mensagem",        grupo:"Chat", base:13, caixa:true },
  chatCaixaEnvio:      { nome:"Caixa de digitar do chat",     grupo:"Chat", base:14, caixa:true },
  chatStatus:          { nome:"Barra de status da IA",        grupo:"Chat", base:12, caixa:true },
  // ─ Geral ─
  fundoApp:            { nome:"Fundo geral do app",           grupo:"Geral", base:13, caixa:true },
  todosCards:          { nome:"Todos os cards (global)",      grupo:"Geral", base:13, caixa:true },
};
const ANIMACOES_IA = ["nenhuma","pulsar","brilhar","flutuar","balancar","girar","tremer","surgir"];

const ACOES_VALIDAS = {
  tema:      { label:"Trocar tema" },
  layout:    { label:"Trocar layout (celular/TV)" },
  densidade: { label:"Trocar densidade" },
  fonte:     { label:"Ajustar tamanho da fonte" },
  meta:      { label:"Definir meta de proventos" },
  aporte:    { label:"Definir meta de aporte" },
  navegar:   { label:"Ir para uma tela" },
  blocoMover:   { label:"Mover bloco de página" },
  blocoVisivel: { label:"Mostrar/ocultar bloco" },
  estilo:       { label:"Estilizar elemento (fonte/cor/efeitos)" },
};
function detectarAcoes(txt) {
  const acoes = [];
  const blocos = String(txt||"").match(/\[\[ACAO\]\][\s\S]*?\[\[\/ACAO\]\]/g) || [];
  blocos.forEach(b=>{
    const corpo = b.replace("[[ACAO]]","").replace("[[/ACAO]]","").trim();
    try { const o = JSON.parse(corpo); if (o && o.tipo && ACOES_VALIDAS[o.tipo]) acoes.push(o); } catch(e){}
  });
  return acoes;
}
// remove os blocos [[ACAO]] do texto exibido ao usuário (fica limpo no chat)
function limparAcoesDoTexto(txt) {
  return String(txt||"").replace(/\[\[ACAO\]\][\s\S]*?\[\[\/ACAO\]\]/g, "").trim();
}


// ════════════════════════════════════════════════════════════════════════════
function gerarRelatorioPDF({ ativos, metaMensal, custoVida, totalAnual, provEsteMes, mediaMes, patrimonioTotal }) {
  const doc = new jsPDF({ unit:"mm", format:"a4" });
  const W = 210; const M = 16; let y = 0;
  const VERDE = [79,97,71], CINZA=[120,128,118], ESCURO=[51,64,47], CLARO=[236,238,231], CIANO=[109,150,144];
  const real = (v)=>"R$ "+(v||0).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2});
  const linha = (yy)=>{ doc.setDrawColor(220,224,214); doc.line(M,yy,W-M,yy); };
  const checkPage = (precisa=10)=>{ if (y+precisa>285){ doc.addPage(); y=20; } };

  // cabeçalho
  doc.setFillColor(...VERDE); doc.rect(0,0,W,30,"F");
  doc.setTextColor(255,255,255); doc.setFont("helvetica","bold"); doc.setFontSize(18);
  doc.text("Carteira de Proventos", M, 13);
  doc.setFont("helvetica","normal"); doc.setFontSize(10);
  const hoje = new Date().toLocaleDateString("pt-BR",{ day:"2-digit", month:"long", year:"numeric" });
  doc.text("Relatório gerado em "+hoje, M, 21);
  y = 40;

  // números principais
  const ativosVivos = ativos.filter(a=>a.qtd>0);
  doc.setTextColor(...CINZA); doc.setFontSize(9); doc.setFont("helvetica","normal");
  doc.text("PATRIMÔNIO TOTAL", M, y);
  doc.text("DIVIDENDOS / ANO", M+62, y);
  doc.text("RECEBIDO ESTE MÊS", M+124, y);
  y += 7;
  doc.setTextColor(...ESCURO); doc.setFont("helvetica","bold"); doc.setFontSize(15);
  doc.text(real(patrimonioTotal), M, y);
  doc.setTextColor(...VERDE); doc.text(real(totalAnual), M+62, y);
  doc.setTextColor(...CIANO); doc.text(real(provEsteMes), M+124, y);
  y += 6;
  doc.setTextColor(...CINZA); doc.setFont("helvetica","normal"); doc.setFontSize(8);
  doc.text(`${ativosVivos.length} ativos`, M, y);
  doc.text(`média ${real(mediaMes)}/mês`, M+62, y);
  y += 8; linha(y); y += 8;

  // composição por classe
  doc.setTextColor(...ESCURO); doc.setFont("helvetica","bold"); doc.setFontSize(11);
  doc.text("Composição por classe", M, y); y += 7;
  const total = ativosVivos.reduce((s,a)=>s+a.qtd*a.cotacao,0);
  const grupos = {};
  ativosVivos.forEach(a=>{ const c=classeDe(a); grupos[c]=(grupos[c]||0)+a.qtd*a.cotacao; });
  const ordenadas = Object.entries(grupos).sort((a,b)=>b[1]-a[1]);
  const corClasse = (c)=>{ const h=(CLASSE_COR[c]||"#4f6147").replace("#",""); return [parseInt(h.slice(0,2),16),parseInt(h.slice(2,4),16),parseInt(h.slice(4,6),16)]; };
  ordenadas.forEach(([cl,val])=>{
    checkPage(8);
    const pct = total>0?val/total*100:0;
    doc.setFont("helvetica","normal"); doc.setFontSize(9); doc.setTextColor(...ESCURO);
    doc.text(cl, M, y);
    doc.setTextColor(...CINZA); doc.text(real(val)+`  (${pct.toFixed(1)}%)`, W-M, y, { align:"right" });
    y += 2.5;
    doc.setFillColor(...CLARO); doc.rect(M, y, W-2*M, 2.5, "F");
    doc.setFillColor(...corClasse(cl)); doc.rect(M, y, (W-2*M)*pct/100, 2.5, "F");
    y += 7;
  });
  y += 4; linha(y); y += 8;

  // lista de ativos
  checkPage(20);
  doc.setTextColor(...ESCURO); doc.setFont("helvetica","bold"); doc.setFontSize(11);
  doc.text("Ativos", M, y); y += 7;
  doc.setFontSize(8); doc.setTextColor(...CINZA); doc.setFont("helvetica","bold");
  doc.text("ATIVO", M, y); doc.text("QTD", M+70, y, {align:"right"}); doc.text("P.MÉDIO", M+100, y, {align:"right"});
  doc.text("COTAÇÃO", M+130, y, {align:"right"}); doc.text("VALOR", W-M, y, {align:"right"}); y += 2;
  linha(y); y += 5;
  ativosVivos.sort((a,b)=>b.qtd*b.cotacao-a.qtd*a.cotacao).forEach(a=>{
    checkPage(7);
    doc.setFont("helvetica","bold"); doc.setFontSize(8.5); doc.setTextColor(...ESCURO);
    doc.text(String(a.ticker).slice(0,28), M, y);
    doc.setFont("helvetica","normal"); doc.setTextColor(...CINZA);
    doc.text(String(a.qtd<1?a.qtd.toFixed(2):a.qtd), M+70, y, {align:"right"});
    doc.text(real(a.precoMedio).replace("R$ ",""), M+100, y, {align:"right"});
    doc.text(real(a.cotacao).replace("R$ ",""), M+130, y, {align:"right"});
    doc.setTextColor(...ESCURO); doc.setFont("helvetica","bold");
    doc.text(real(a.qtd*a.cotacao), W-M, y, {align:"right"});
    y += 5.5;
  });
  y += 3; linha(y); y += 8;

  // metas resumo
  checkPage(30);
  doc.setTextColor(...ESCURO); doc.setFont("helvetica","bold"); doc.setFontSize(11);
  doc.text("Metas", M, y); y += 7;
  doc.setFont("helvetica","normal"); doc.setFontSize(9);
  const metaPct = metaMensal>0?Math.min(mediaMes/metaMensal*100,100):0;
  doc.setTextColor(...CINZA);
  doc.text(`Meta de proventos: ${real(mediaMes)} / ${real(metaMensal)} por mês (${metaPct.toFixed(0)}%)`, M, y); y += 6;
  const custoTotal = Object.values(custoVida||{}).reduce((s,v)=>s+(+v||0),0);
  if (custoTotal>0) { const cp=Math.min(mediaMes/custoTotal*100,100); doc.text(`Contas fixas cobertas: ${cp.toFixed(0)}% de ${real(custoTotal)}/mês`, M, y); y += 6; }
  const reserva = ativosVivos.filter(a=>a.cat==="Tesouro"||ETFS_GLOBAIS.includes(a.ticker)||/cdb|lci|lca/i.test(String(a.nome))).reduce((s,a)=>s+a.qtd*a.cotacao,0);
  doc.text(`Reserva Plus (Tesouro + ETFs globais): ${real(reserva)}`, M, y); y += 10;

  // rodapé
  const paginas = doc.internal.getNumberOfPages();
  for (let p=1;p<=paginas;p++){ doc.setPage(p); doc.setFontSize(7); doc.setTextColor(...CINZA);
    doc.text("Gerado pelo app Carteira de Proventos · estimativas informativas, não é recomendação de investimento", W/2, 292, {align:"center"});
    doc.text(`${p}/${paginas}`, W-M, 292, {align:"right"});
  }

  const nomeArq = `carteira-proventos-${new Date().toISOString().slice(0,10)}.pdf`;
  // No Android (app empacotado): salva e abre a tela de compartilhar. No navegador: baixa direto.
  (async () => {
    if (typeof window!=="undefined" && window.Capacitor?.isNativePlatform?.()) {
      try {
        const fsMod = await import("@capacitor/filesystem");
        const shareMod = await import("@capacitor/share");
        const base64 = doc.output("datauristring").split(",")[1];
        const res = await fsMod.Filesystem.writeFile({ path:nomeArq, data:base64, directory:fsMod.Directory.Cache });
        await shareMod.Share.share({ title:"Relatório da carteira", text:"Relatório da minha carteira de proventos", url:res.uri });
        return;
      } catch(e) { /* se falhar, tenta o download normal abaixo */ }
    }
    try { doc.save(nomeArq); }
    catch(e) {
      try { const url = doc.output("bloburl"); window.open(url, "_blank"); }
      catch(e2) { alert("Não consegui gerar o PDF neste dispositivo."); }
    }
  })();
  return nomeArq;
}

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
  minimalista: {
    nome: "Minimalista",
    emoji: "🌿",
    claro: true, // tema claro
    bg: "#eceee7",            // verde-acinzentado bem claro (parede)
    bgHeader: "linear-gradient(160deg,#dfe4da 0%,#eceee7 100%)",
    card: "#f6f7f2",          // quase branco, levemente quente
    cardAlt: "#e7e9e1",       // cinza do sofá
    border: "#d6dacf",        // borda suave
    borderSoft: "#c3c9b9",
    text: "#33402f",          // verde escuro (almofada) p/ texto — delicado, não preto
    textDim: "#5a6553",
    textMute: "#8a9382",
    textFaint: "#a9b0a0",
    accent: "#4f6147",        // verde escuro da almofada (cor principal)
    accentSoft: "#6f8166",
    accentBg: "#dde3d7",
    accentBorder: "#8a9b80",
    green: "#5e8a5a",         // verde suave p/ positivo
    red: "#b07c6e",           // terracota suave p/ negativo (sem vermelho gritante)
    amber: "#b69a5e",         // ocre suave
    cyan: "#6d9690",          // verde-água suave
    // paletas de ativos: tons terrosos/sálvia, pouca cor (conceito minimalista)
    corFII: { MFII11:"#6d9690",VGHF11:"#7ba39c",RBRY11:"#89a59f",OIAG11:"#9bb0a0",CPTS11:"#a9bbac",GARE11:"#7f9b86",TRXF11:"#6e8a74",RBRX11:"#83a08a",IRIM11:"#96ad99",KNCR11:"#aabba0",VISC11:"#7d978a",XPML11:"#92a995",MXRF11:"#b3c0b0" },
    corAcao: { BBAS3:"#4f6147",PETR4:"#5e7155",BBDC3:"#6f8166",ITSA4:"#7d8f73",BBSE3:"#8a9b80",TAEE11:"#97a78d",BRSR6:"#a4b29a",KLBN4:"#69745e",KLBN11:"#76836a",CXSE3:"#838f76",WEGE3:"#909b82",CPLE3:"#9da78e",SANB3:"#aab39b" },
  },
  minimalista2: {
    nome: "Minimal+",
    emoji: "🍃",
    claro: true,
    espacoso: true, // espaçamentos generosos estilo B3/Nubank
    bg: "#eef0ea",
    bgHeader: "linear-gradient(160deg,#e3e7df 0%,#eef0ea 100%)",
    card: "#f8f9f5",
    cardAlt: "#e9ebe3",
    border: "#dde1d7",
    borderSoft: "#cad0c1",
    text: "#33402f",
    textDim: "#5a6553",
    textMute: "#8a9382",
    textFaint: "#aab0a0",
    accent: "#4f6147",
    accentSoft: "#6f8166",
    accentBg: "#e1e7db",
    accentBorder: "#8a9b80",
    green: "#5e8a5a",
    red: "#b07c6e",
    amber: "#b69a5e",
    cyan: "#6d9690",
    corFII: { MFII11:"#6d9690",VGHF11:"#7ba39c",RBRY11:"#89a59f",OIAG11:"#9bb0a0",CPTS11:"#a9bbac",GARE11:"#7f9b86",TRXF11:"#6e8a74",RBRX11:"#83a08a",IRIM11:"#96ad99",KNCR11:"#aabba0",VISC11:"#7d978a",XPML11:"#92a995",MXRF11:"#b3c0b0" },
    corAcao: { BBAS3:"#4f6147",PETR4:"#5e7155",BBDC3:"#6f8166",ITSA4:"#7d8f73",BBSE3:"#8a9b80",TAEE11:"#97a78d",BRSR6:"#a4b29a",KLBN4:"#69745e",KLBN11:"#76836a",CXSE3:"#838f76",WEGE3:"#909b82",CPLE3:"#9da78e",SANB3:"#aab39b" },
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

// ── IMPORTAÇÃO EM MASSA — lê texto colado tipo "BBAS3 291 20,55" ──────────
// Aceita: TICKER QTD [PREÇO_MÉDIO] [COTAÇÃO], com vírgula ou ponto decimal.
function parseNumBR(s) {
  if (s == null) return NaN;
  s = String(s).replace(/[R$\s]/g, "");
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", "."); // formato BR: 1.234,56
  return parseFloat(s);
}
function parseImportacao(txt) {
  const linhas = String(txt||"").split(/\n+/).map(l=>l.trim()).filter(Boolean);
  const itens = [];
  for (const l of linhas) {
    const m = l.match(/([A-Za-z]{4}\d{1,2})/); // ticker tipo BBAS3, MFII11
    if (!m) { itens.push({ raw:l, erro:true }); continue; }
    const ticker = m[1].toUpperCase();
    const resto = l.slice(l.indexOf(m[1]) + m[1].length);
    const nums = (resto.match(/[\d.,]+/g) || []).map(parseNumBR).filter(n=>!isNaN(n));
    if (nums.length === 0) { itens.push({ raw:l, ticker, erro:true }); continue; }
    itens.push({
      ticker,
      qtd:        nums[0] != null ? Math.round(nums[0]) : null,
      precoMedio: nums[1] != null ? nums[1] : null,
      cotacao:    nums[2] != null ? nums[2] : null,
    });
  }
  return itens;
}


// Grupo de exibição (para listas agrupadas e expansíveis). NÃO altera o campo
// 'cat' do ativo (que continua alimentando gráficos e filtros) — é só para exibir.
function grupoDe(a) {
  if (a.cat === "FII") return "FIIs";
  if (a.cat === "Cripto") return "Cripto";
  if (/^(BOVA11|BOVX11|IVVB11|SMAL11|HASH11|XINA11|NASD11|GOLD11|VWRA11)$/.test(a.ticker)) return "ETFs";
  if (/tesouro|tesour/i.test(a.nome) || /^(LFT|LTN|NTN|TD)/.test(a.ticker)) return "Tesouro Direto";
  return "Ações";
}
const ORDEM_GRUPOS = ["FIIs","Ações","ETFs","Cripto","Tesouro Direto","Outros"];
const EMOJI_GRUPO = { "FIIs":"🏢","Ações":"📈","ETFs":"📊","Cripto":"🪙","Tesouro Direto":"🏛️","Outros":"📦" };
const COR_GRUPO = (nome,T) => ({ "FIIs":T.cyan,"Ações":T.accent,"ETFs":"#f472b6","Cripto":T.amber,"Tesouro Direto":T.green,"Outros":T.textMute }[nome] || T.accent);


const MESES_BASE = ["Jul/26","Ago/26","Set/26","Out/26","Nov/26","Dez/26","Jan/27","Fev/27","Mar/27","Abr/27","Mai/27","Jun/27"];

// ════════════════════════════════════════════════════════════════════════════
// DADOS INICIAIS DOS ATIVOS (editáveis na aba "Editar")
// ════════════════════════════════════════════════════════════════════════════
// Setor (subcategoria) de cada ativo — poucos grupos, fácil de ler
const SETOR_TICKER = {
  // Ações
  BBAS3:"Bancos", BBDC3:"Bancos", SANB3:"Bancos", BRSR6:"Bancos", ITSA4:"Bancos", B3SA3:"Bancos",
  BBSE3:"Seguradoras", CXSE3:"Seguradoras",
  TAEE11:"Energia & Infra", CPLE3:"Energia & Infra", SBSP3:"Energia & Infra", WEGE3:"Energia & Infra", EMBJ3:"Energia & Infra",
  PETR4:"Commodities", KLBN4:"Commodities", KLBN11:"Commodities",
  // FIIs
  VGHF11:"FII Papel", RBRY11:"FII Papel", CPTS11:"FII Papel", IRIM11:"FII Papel", KNCR11:"FII Papel", MXRF11:"FII Papel", RBRX11:"FII Papel",
  MFII11:"FII Tijolo", TRXF11:"FII Tijolo", GARE11:"FII Tijolo", VISC11:"FII Tijolo", XPML11:"FII Tijolo", OIAG11:"FII Tijolo",
  // Índice & Cripto
  BOVA11:"Índice & Cripto", BOVX11:"Índice & Cripto", COIN11:"Índice & Cripto",
};
// Cor de cada setor (para gráficos do dashboard)
const COR_SETOR = {
  "Bancos":"#6366f1", "Seguradoras":"#a78bfa", "Energia & Infra":"#22d3ee",
  "Commodities":"#fbbf24", "FII Papel":"#34d399", "FII Tijolo":"#2dd4bf",
  "Índice & Cripto":"#f472b6", "Outros":"#64748b",
};

const ATIVOS_INICIAIS = []; // conta nova começa do zero — ativos são adicionados pelo usuário

// ════════════════════════════════════════════════════════════════════════════
// CATÁLOGO DE ATIVOS — sugestões ao adicionar (t=ticker, n=nome, c=cat, s=setor, f=freq)
// ════════════════════════════════════════════════════════════════════════════
const CATALOGO_ATIVOS = [
  {t:"VALE3",n:"Vale",c:"Ação",s:"Mineração",f:"Semestral"},{t:"PETR4",n:"Petrobras PN",c:"Ação",s:"Petróleo",f:"Trimestral"},{t:"PETR3",n:"Petrobras ON",c:"Ação",s:"Petróleo",f:"Trimestral"},
  {t:"ITUB4",n:"Itaú Unibanco",c:"Ação",s:"Bancos",f:"Mensal"},{t:"BBDC4",n:"Bradesco",c:"Ação",s:"Bancos",f:"Mensal"},{t:"BBAS3",n:"Banco do Brasil",c:"Ação",s:"Bancos",f:"Trimestral"},
  {t:"SANB11",n:"Santander",c:"Ação",s:"Bancos",f:"Semestral"},{t:"ITSA4",n:"Itaúsa",c:"Ação",s:"Holdings",f:"Trimestral"},{t:"B3SA3",n:"B3",c:"Ação",s:"Financeiro",f:"Trimestral"},
  {t:"BBSE3",n:"BB Seguridade",c:"Ação",s:"Seguros",f:"Semestral"},{t:"CXSE3",n:"Caixa Seguridade",c:"Ação",s:"Seguros",f:"Semestral"},{t:"PSSA3",n:"Porto Seguro",c:"Ação",s:"Seguros",f:"Trimestral"},
  {t:"WEGE3",n:"WEG",c:"Ação",s:"Industrial",f:"Trimestral"},{t:"ABEV3",n:"Ambev",c:"Ação",s:"Bebidas",f:"Anual"},{t:"SUZB3",n:"Suzano",c:"Ação",s:"Papel e Celulose",f:"Anual"},
  {t:"KLBN11",n:"Klabin",c:"Ação",s:"Papel e Celulose",f:"Trimestral"},{t:"GGBR4",n:"Gerdau",c:"Ação",s:"Siderurgia",f:"Trimestral"},{t:"CSNA3",n:"CSN",c:"Ação",s:"Siderurgia",f:"Anual"},
  {t:"USIM5",n:"Usiminas",c:"Ação",s:"Siderurgia",f:"Anual"},{t:"CMIG4",n:"Cemig",c:"Ação",s:"Energia",f:"Semestral"},{t:"CPLE6",n:"Copel",c:"Ação",s:"Energia",f:"Semestral"},
  {t:"ELET3",n:"Eletrobras",c:"Ação",s:"Energia",f:"Anual"},{t:"EGIE3",n:"Engie Brasil",c:"Ação",s:"Energia",f:"Semestral"},{t:"TAEE11",n:"Taesa",c:"Ação",s:"Energia",f:"Trimestral"},
  {t:"TRPL4",n:"ISA Energia",c:"Ação",s:"Energia",f:"Anual"},{t:"CPFE3",n:"CPFL Energia",c:"Ação",s:"Energia",f:"Anual"},{t:"AURE3",n:"Auren",c:"Ação",s:"Energia",f:"Anual"},
  {t:"NEOE3",n:"Neoenergia",c:"Ação",s:"Energia",f:"Anual"},{t:"SBSP3",n:"Sabesp",c:"Ação",s:"Saneamento",f:"Anual"},{t:"CSMG3",n:"Copasa",c:"Ação",s:"Saneamento",f:"Trimestral"},
  {t:"SAPR11",n:"Sanepar",c:"Ação",s:"Saneamento",f:"Trimestral"},{t:"VIVT3",n:"Vivo",c:"Ação",s:"Telecom",f:"Trimestral"},{t:"TIMS3",n:"TIM",c:"Ação",s:"Telecom",f:"Trimestral"},
  {t:"RADL3",n:"RaiaDrogasil",c:"Ação",s:"Saúde",f:"Trimestral"},{t:"HYPE3",n:"Hypera",c:"Ação",s:"Saúde",f:"Trimestral"},{t:"FLRY3",n:"Fleury",c:"Ação",s:"Saúde",f:"Trimestral"},
  {t:"RENT3",n:"Localiza",c:"Ação",s:"Locação",f:"Trimestral"},{t:"RAIL3",n:"Rumo",c:"Ação",s:"Logística",f:"Anual"},{t:"MOTV3",n:"Motiva (CCR)",c:"Ação",s:"Concessões",f:"Anual"},
  {t:"EMBR3",n:"Embraer",c:"Ação",s:"Aeroespacial",f:"—"},{t:"PRIO3",n:"PRIO",c:"Ação",s:"Petróleo",f:"—"},{t:"RECV3",n:"PetroRecôncavo",c:"Ação",s:"Petróleo",f:"Trimestral"},
  {t:"CSAN3",n:"Cosan",c:"Ação",s:"Energia",f:"Anual"},{t:"UGPA3",n:"Ultrapar",c:"Ação",s:"Distribuição",f:"Semestral"},{t:"VBBR3",n:"Vibra",c:"Ação",s:"Distribuição",f:"Semestral"},
  {t:"LREN3",n:"Lojas Renner",c:"Ação",s:"Varejo",f:"Trimestral"},{t:"MGLU3",n:"Magazine Luiza",c:"Ação",s:"Varejo",f:"—"},{t:"JBSS3",n:"JBS",c:"Ação",s:"Alimentos",f:"Anual"},
  {t:"BRFS3",n:"BRF",c:"Ação",s:"Alimentos",f:"—"},{t:"BEEF3",n:"Minerva",c:"Ação",s:"Alimentos",f:"Anual"},
  {t:"MXRF11",n:"Maxi Renda",c:"FII",s:"Papel",f:"Mensal"},{t:"KNCR11",n:"Kinea Rendimentos",c:"FII",s:"Papel",f:"Mensal"},{t:"KNIP11",n:"Kinea Índice de Preços",c:"FII",s:"Papel",f:"Mensal"},
  {t:"KNRI11",n:"Kinea Renda Imobiliária",c:"FII",s:"Híbrido",f:"Mensal"},{t:"HGLG11",n:"CSHG Logística",c:"FII",s:"Logística",f:"Mensal"},{t:"BTLG11",n:"BTG Logística",c:"FII",s:"Logística",f:"Mensal"},
  {t:"VILG11",n:"Vinci Logística",c:"FII",s:"Logística",f:"Mensal"},{t:"XPLG11",n:"XP Log",c:"FII",s:"Logística",f:"Mensal"},{t:"LVBI11",n:"VBI Logístico",c:"FII",s:"Logística",f:"Mensal"},
  {t:"HGRU11",n:"CSHG Renda Urbana",c:"FII",s:"Renda Urbana",f:"Mensal"},{t:"TRXF11",n:"TRX Real Estate",c:"FII",s:"Renda Urbana",f:"Mensal"},{t:"GARE11",n:"Guardian Real Estate",c:"FII",s:"Renda Urbana",f:"Mensal"},
  {t:"HGRE11",n:"CSHG Real Estate",c:"FII",s:"Lajes",f:"Mensal"},{t:"PVBI11",n:"VBI Prime Properties",c:"FII",s:"Lajes",f:"Mensal"},{t:"JSRE11",n:"JS Real Estate",c:"FII",s:"Lajes",f:"Mensal"},
  {t:"XPML11",n:"XP Malls",c:"FII",s:"Shoppings",f:"Mensal"},{t:"VISC11",n:"Vinci Shopping Centers",c:"FII",s:"Shoppings",f:"Mensal"},{t:"HGBS11",n:"Hedge Brasil Shopping",c:"FII",s:"Shoppings",f:"Mensal"},
  {t:"MALL11",n:"Malls Brasil Plural",c:"FII",s:"Shoppings",f:"Mensal"},{t:"HSML11",n:"HSI Malls",c:"FII",s:"Shoppings",f:"Mensal"},{t:"MFII11",n:"Mérito Desenvolvimento",c:"FII",s:"Desenvolvimento",f:"Mensal"},
  {t:"CPTS11",n:"Capitânia Securities",c:"FII",s:"Papel",f:"Mensal"},{t:"VGHF11",n:"Valora Hedge Fund",c:"FII",s:"Multiestratégia",f:"Mensal"},{t:"VGIR11",n:"Valora CRI",c:"FII",s:"Papel",f:"Mensal"},
  {t:"RBRY11",n:"RBR Crédito Estruturado",c:"FII",s:"Papel",f:"Mensal"},{t:"RBRR11",n:"RBR High Grade",c:"FII",s:"Papel",f:"Mensal"},{t:"IRDM11",n:"Iridium Recebíveis",c:"FII",s:"Papel",f:"Mensal"},
  {t:"VRTA11",n:"Fator Verità",c:"FII",s:"Papel",f:"Mensal"},{t:"RECR11",n:"REC Recebíveis",c:"FII",s:"Papel",f:"Mensal"},{t:"BCFF11",n:"BTG Fundo de Fundos",c:"FII",s:"FoF",f:"Mensal"},
  {t:"RBRF11",n:"RBR Alpha",c:"FII",s:"FoF",f:"Mensal"},{t:"HFOF11",n:"Hedge Top FoF",c:"FII",s:"FoF",f:"Mensal"},{t:"VINO11",n:"Vinci Offices",c:"FII",s:"Lajes",f:"Mensal"},
  {t:"RZTR11",n:"Riza Terrax",c:"FII",s:"Agro",f:"Mensal"},{t:"OIAG11",n:"Ourinvest Agro",c:"FII",s:"Agro",f:"Mensal"},{t:"KNSC11",n:"Kinea Securities",c:"FII",s:"Papel",f:"Mensal"},
  {t:"BOVA11",n:"iShares Ibovespa",c:"ETF",s:"Índice BR",f:"—"},{t:"IVVB11",n:"iShares S&P 500",c:"ETF",s:"Internacional",f:"—"},{t:"SMAL11",n:"iShares Small Caps",c:"ETF",s:"Índice BR",f:"—"},
  {t:"DIVO11",n:"It Now Dividendos",c:"ETF",s:"Dividendos",f:"—"},{t:"HASH11",n:"Hashdex Cripto",c:"ETF",s:"Cripto",f:"—"},{t:"GOLD11",n:"Trend Ouro",c:"ETF",s:"Ouro",f:"—"},
  {t:"XFIX11",n:"Trend IFIX",c:"ETF",s:"FIIs",f:"Mensal"},{t:"IMAB11",n:"It Now IMA-B",c:"ETF",s:"Renda Fixa",f:"—"},{t:"B5P211",n:"It Now IMA-B5",c:"ETF",s:"Renda Fixa",f:"—"},{t:"WRLD11",n:"Trend Mundo",c:"ETF",s:"Internacional",f:"—"},
  {t:"AAPL34",n:"Apple",c:"Ação",s:"BDR · Tecnologia",f:"Trimestral"},{t:"MSFT34",n:"Microsoft",c:"Ação",s:"BDR · Tecnologia",f:"Trimestral"},{t:"GOGL34",n:"Alphabet (Google)",c:"Ação",s:"BDR · Tecnologia",f:"—"},
  {t:"AMZO34",n:"Amazon",c:"Ação",s:"BDR · Tecnologia",f:"—"},{t:"TSLA34",n:"Tesla",c:"Ação",s:"BDR · Automotivo",f:"—"},{t:"NVDC34",n:"NVIDIA",c:"Ação",s:"BDR · Tecnologia",f:"Trimestral"},
  {t:"META34",n:"Meta (Facebook)",c:"Ação",s:"BDR · Tecnologia",f:"Trimestral"},{t:"NFLX34",n:"Netflix",c:"Ação",s:"BDR · Mídia",f:"—"},{t:"DISB34",n:"Disney",c:"Ação",s:"BDR · Mídia",f:"Anual"},{t:"COCA34",n:"Coca-Cola",c:"Ação",s:"BDR · Consumo",f:"Trimestral"},
  {t:"PAGS34",n:"PagSeguro",c:"Ação",s:"BDR · Pagamentos",f:"—"},{t:"VWRA11",n:"Investo FTSE All-World",c:"ETF",s:"Internacional",f:"—"},
  {t:"TESOURO SELIC 2029",n:"Tesouro Selic 2029",c:"Tesouro",s:"Renda Fixa",f:"—"},{t:"TESOURO IPCA+ 2035",n:"Tesouro IPCA+ 2035",c:"Tesouro",s:"Renda Fixa",f:"—"},
  {t:"TESOURO IPCA+ 2045",n:"Tesouro IPCA+ 2045",c:"Tesouro",s:"Renda Fixa",f:"—"},{t:"TESOURO PREFIXADO 2029",n:"Tesouro Prefixado 2029",c:"Tesouro",s:"Renda Fixa",f:"—"},
  {t:"TESOURO RENDA+ 2065",n:"Tesouro Renda+ 2065",c:"Tesouro",s:"Renda Fixa",f:"—"},
  {t:"BTC",n:"Bitcoin",c:"Cripto",s:"Cripto",f:"—"},{t:"ETH",n:"Ethereum",c:"Cripto",s:"Cripto",f:"—"},
];

// ════════════════════════════════════════════════════════════════════════════
// PERSISTÊNCIA — salva tudo na memória permanente do app (localStorage)
// Funciona no APK. No preview do Claude o localStorage é bloqueado, então
// envolvemos em try/catch: lá não salva, mas também não quebra.
// ════════════════════════════════════════════════════════════════════════════
let PREFIXO = "carteiraProventos_"; // muda por usuário após o login (dados separados por conta)
const CONTAS_KEY = "carteiraProventos_CONTAS";
const SESSAO_KEY = "carteiraProventos_SESSAO";
function definirUsuarioAtivo(email) {
  const slug = String(email||"").toLowerCase().replace(/[^a-z0-9]/g, "_").slice(0, 40);
  PREFIXO = "carteiraProventos_u_" + slug + "_";
}
function hashSenha(sx) { let h=5381; const str=String(sx); for (let i=0;i<str.length;i++){ h=((h<<5)+h+str.charCodeAt(i))|0; } return "h"+(h>>>0).toString(36); }
function lerContas(){ try { return JSON.parse(localStorage.getItem(CONTAS_KEY)||"{}")||{}; } catch { return {}; } }
function gravarContas(c){ try { localStorage.setItem(CONTAS_KEY, JSON.stringify(c)); } catch {} }
function gerarFraseRec(){ const P=["verde","ouro","lua","rio","trigo","norte","pinho","sol","mar","serra","campo","aurora"]; const r=()=>P[Math.floor(Math.random()*P.length)]; return `${r()}-${r()}-${r()}-${Math.floor(100+Math.random()*900)}`; }

function lerSalvo(chave, padrao) {
  try {
    const v = localStorage.getItem(PREFIXO + chave);
    return v !== null ? JSON.parse(v) : padrao;
  } catch { return padrao; }
}
function gravarSalvo(chave, valor) {
  try { localStorage.setItem(PREFIXO + chave, JSON.stringify(valor)); } catch { /* preview: ignora */ }
}

// ════════════════════════════════════════════════════════════════════════════
// MOTOR DE LOGS — registro central de tudo que acontece no app
// Cada log: { id, ts, tipo, direcao, origem, msg, detalhe }
//   tipo:    "cotacao" | "edicao" | "chat" | "erro" | "sistema" | "import"
//   direcao: "ida" (app->fora) | "volta" (fora->app) | "interno"
//   origem:  "app" | "api" (brapi) | "servidor" (bridge/IA)
// ════════════════════════════════════════════════════════════════════════════
const LOG_CHAVE = "logs";
const LOG_MAX = 500; // mantém os últimos 500 registros
let _logListeners = [];

function registrarLog(tipo, msg, opts = {}) {
  const entrada = {
    id: Date.now() + "_" + Math.random().toString(36).slice(2, 7),
    ts: new Date().toISOString(),
    tipo,
    direcao: opts.direcao || "interno",
    origem: opts.origem || "app",
    msg: String(msg),
    detalhe: opts.detalhe != null ? (typeof opts.detalhe === "string" ? opts.detalhe : JSON.stringify(opts.detalhe)) : null,
  };
  try {
    const atuais = lerSalvo(LOG_CHAVE, []);
    const novos = [entrada, ...atuais].slice(0, LOG_MAX);
    gravarSalvo(LOG_CHAVE, novos);
    _logListeners.forEach(fn => { try { fn(novos); } catch {} });
  } catch {}
  return entrada;
}
function lerLogs() { return lerSalvo(LOG_CHAVE, []); }
function limparLogs() { gravarSalvo(LOG_CHAVE, []); _logListeners.forEach(fn=>{try{fn([]);}catch{}}); }
function inscreverLogs(fn) { _logListeners.push(fn); return () => { _logListeners = _logListeners.filter(f=>f!==fn); }; }


// Hook: igual ao useState, mas grava automaticamente toda mudança na memória.
function useEstadoSalvo(chave, padrao) {
  const [valor, setValor] = useState(() => lerSalvo(chave, padrao));
  useEffect(() => { gravarSalvo(chave, valor); }, [chave, valor]);
  return [valor, setValor];
}

// Carrega os ativos salvos, mas SEMPRE inclui ativos novos que eu adicionar no
// código no futuro (faz merge por ticker): suas edições são preservadas e os
// ativos novos aparecem mesmo que você já tenha dados salvos.
function carregarAtivos() {
  const salvos = lerSalvo("ativos", null);
  if (!Array.isArray(salvos)) return ATIVOS_INICIAIS.map(a => ({ ...a }));
  const porTicker = {};
  salvos.forEach(a => { if (a && a.ticker) porTicker[a.ticker] = a; });
  // base = lista do código; aplica edições salvas por cima; adiciona tickers novos
  const merge = ATIVOS_INICIAIS.map(base =>
    porTicker[base.ticker]
      ? { ...base, qtd:porTicker[base.ticker].qtd, prov:porTicker[base.ticker].prov,
          precoMedio:porTicker[base.ticker].precoMedio, cotacao:porTicker[base.ticker].cotacao,
          freq:porTicker[base.ticker].freq ?? base.freq, meses:porTicker[base.ticker].meses ?? base.meses }
      : { ...base }
  );
  return merge;
}

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

function simular(ativos, regras, horizonte, aporte, aportesExtras = [], creditoCartao = 0) {
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
    // aportes esporádicos que caem neste mês (13º, férias, venda de item, etc.)
    const extrasMes = aportesExtras.filter(e => e.mes === m);
    const extraValor = extrasMes.reduce((s,e)=>s+(+e.valor||0), 0);
    // crédito do cartão garantido por ativos: injeção única no 1º mês (acelera a bola de neve)
    const creditoMes = m === 0 ? creditoCartao : 0;
    const caixaInicial = acum + provBrutoMes + aporte + extraValor + creditoMes;
    let caixa = caixaInicial;
    const compras = [];
    let reinvestido = 0;
    regras.forEach(r => {
      if (r.pct <= 0) return;
      const ativo = estado.find(a => a.ticker === r.ticker);
      if (!ativo) return;
      const cot = r.cotacaoAlvo || ativo.cotacao;
      const cotas = Math.floor((caixaInicial * (r.pct / 100)) / cot);
      if (cotas > 0) {
        ativo.qtd += cotas;
        caixa -= cotas * cot;
        reinvestido += cotas * cot;
        compras.push({ ticker: r.ticker, cotas, gasto: +(cotas*cot).toFixed(2) });
      }
    });
    acum = Math.max(caixa, 0);
    const patri = estado.reduce((s,a) => s + a.qtd*a.cotacao, 0);
    const provMedio = estado.reduce((s,a) => s + a.prov*a.qtd*a.meses.length/12, 0);
    // limite de cartão (PLACEHOLDER): projeção de crédito usando ativos como garantia.
    // Será refinado quando a aba de cartões for construída (ex: Tesouro como garantia).
    const limiteCartao = patri * 0.40;
    return {
      mes: labelMes(m),
      provento: +provBrutoMes.toFixed(2),
      provMedio: +provMedio.toFixed(2),
      patrimonio: +patri.toFixed(2),
      caixa: +acum.toFixed(2),
      limiteCartao: +limiteCartao.toFixed(2),
      reinvestido: +reinvestido.toFixed(2),
      aporteExtra: +extraValor.toFixed(2),
      extrasMes,
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
  const [abertos, setAbertos] = useState({});
  const lista = filtro==="TUDO" ? ativos : filtro==="FII" ? ativos.filter(a=>a.cat==="FII") : ativos.filter(a=>a.cat==="Ação");
  const pag = lista.filter(a=>mesesIdx(a.meses).includes(idx)).map(a=>({...a,total:+(a.prov*a.qtd).toFixed(2)})).sort((a,b)=>b.total-a.total);
  const total = pag.reduce((s,a)=>s+a.total,0);

  // agrupa por categoria de exibição
  const grupos = {};
  pag.forEach(a => { const g = grupoDe(a); (grupos[g] = grupos[g] || []).push(a); });
  const ordenados = ORDEM_GRUPOS.filter(g => grupos[g]);

  return (
    <div>
      <div style={{ display:"flex",alignItems:"baseline",gap:8,marginBottom:10 }}>
        <span style={{ fontSize:12,fontWeight:700,color:T.accentSoft }}>{MESES_BASE[idx]}</span>
        <span style={{ fontSize:20,fontWeight:800,color:T.text }}>{fmt(total)}</span>
      </div>
      {pag.length===0
        ? <div style={{ fontSize:12,color:T.textFaint,textAlign:"center",padding:"16px 0" }}>Nenhum provento neste mês</div>
        : ordenados.map(g=>{
          const itens = grupos[g];
          const cg = COR_GRUPO(g,T);
          const aberto = abertos[g];
          const subtotal = itens.reduce((s,a)=>s+a.total,0);
          const patri = itens.reduce((s,a)=>s+a.qtd*a.cotacao,0);
          return (
            <div key={g} style={{ marginBottom:6 }}>
              {/* linha do grupo (clicável) */}
              <div onClick={()=>setAbertos(p=>({ ...p, [g]:!p[g] }))} style={{
                display:"flex", alignItems:"center", justifyContent:"space-between", cursor:"pointer",
                background:`${cg}10`, border:`1px solid ${cg}33`, borderLeft:`3px solid ${cg}`,
                borderRadius:8, padding:"9px 12px"
              }}>
                <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                  <span style={{ fontSize:10, color:T.textFaint, transform:aberto?"rotate(90deg)":"none", transition:"transform 0.2s", display:"inline-block" }}>▶</span>
                  <span style={{ fontSize:13 }}>{EMOJI_GRUPO[g]}</span>
                  <div>
                    <span style={{ fontSize:12, fontWeight:800, color:T.text }}>{g}</span>
                    <div style={{ fontSize:9, color:T.textFaint }}>{itens.length} ativo{itens.length>1?"s":""} · {fmt(patri)} em carteira</div>
                  </div>
                </div>
                <div style={{ fontSize:13, fontWeight:800, color:cg }}>{fmt(subtotal)}</div>
              </div>
              {/* itens do grupo */}
              {aberto && (
                <div style={{ paddingLeft:6, marginTop:5 }}>
                  {itens.map(a=>{
                    const c = corDe(a.ticker,a.cat,T);
                    return (
                      <div key={a.ticker} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",background:`${c}0d`,border:`1px solid ${c}33`,borderLeft:`3px solid ${c}`,borderRadius:8,padding:"8px 12px",marginBottom:5 }}>
                        <div style={{ display:"flex",alignItems:"center",gap:7 }}>
                          <div style={{ width:8,height:8,borderRadius:2,background:c }}/>
                          <div>
                            <span onClick={()=>abrirEditorAtivo(a.ticker)} style={{ fontSize:12,fontWeight:700,color:T.text, cursor:"pointer" }}>{a.ticker}</span>
                            <span style={{ fontSize:9,color:T.textFaint,background:T.border,padding:"1px 5px",borderRadius:4,marginLeft:6 }}>{a.freq}</span>
                            <div style={{ fontSize:10,color:T.textFaint,marginTop:1 }}>{a.nome} · {a.qtd}× R${a.prov.toFixed(2)}</div>
                          </div>
                        </div>
                        <div style={{ fontSize:13,fontWeight:800,color:c }}>{fmt(a.total)}</div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })
      }
    </div>
  );
}

// ════════════════════════════════════════════════════════════════════════════
// DASHBOARD DA CARTEIRA — composição (pizza) + dinheiro (barras), por grupo
// ════════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════════
// BLOCOS CONFIGURÁVEIS (Camada 2) — componentes que podem mudar de página/ordem
// ════════════════════════════════════════════════════════════════════════════
function BlocoProx3Meses({ ativos, estiloDe = () => ({ style:{}, cls:"" }), T }) {
  const NOMES_MES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
  const hoje = new Date();
  const mAtual = hoje.getMonth()+1;
  const prox = (n) => ((mAtual-1+n)%12)+1;
  const provDoMes = (m) => ativos
    .filter(a=>a.qtd>0 && a.prov>0 && a.meses.includes(m))
    .map(a=>({ ticker:a.ticker, valor:+(a.prov*a.qtd).toFixed(2) }))
    .sort((x,y)=>y.valor-x.valor);
  const meses = [
    { rotulo:"📥 Este mês",     m:mAtual,  cor:T.green },
    { rotulo:"📅 Próximo mês",  m:prox(1), cor:T.accentSoft },
    { rotulo:"🗓️ Mês seguinte", m:prox(2), cor:T.cyan },
  ].map(x=>{ const lista = provDoMes(x.m); return { ...x, total:lista.reduce((s,p)=>s+p.valor,0), n:lista.length }; });
  return (
    <div style={{ marginBottom:18 }}>
      <div className={estiloDe("tituloProx3Meses",11).cls} style={{ fontSize:11, color:T.textMute, fontWeight:600, marginBottom:10, ...estiloDe("tituloProx3Meses",11).style }}>📆 Proventos a receber · próximos 3 meses</div>
      <div style={{ display:"flex", gap:10, overflowX:"auto", paddingBottom:4, scrollSnapType:"x mandatory" }}>
        {meses.map((mes,i)=>(
          <div key={i} style={{ flex:"0 0 auto", width:"46%", minWidth:150, scrollSnapAlign:"start",
            background:`linear-gradient(135deg, ${mes.cor}22, ${T.card})`, border:`1px solid ${mes.cor}44`, borderRadius:14, padding:"15px 15px 16px" }}>
            <div style={{ fontSize:9, color:mes.cor, fontWeight:700, textTransform:"uppercase", letterSpacing:0.5 }}>{mes.rotulo}</div>
            <div style={{ fontSize:10, color:T.textFaint, marginBottom:8 }}>{NOMES_MES[mes.m-1]}</div>
            <div style={{ fontSize:22, fontWeight:800, color:mes.cor, letterSpacing:-0.5 }}>{fmt(mes.total)}</div>
            <div style={{ fontSize:9, color:T.textMute, marginTop:5 }}>{mes.n} ativo{mes.n!==1?"s":""} pagando</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function BlocoGraficoMensal({ ativos, estiloDe = () => ({ style:{}, cls:"" }), T }) {
  const dataMes = MESES_BASE.map((label,idx)=>({
    mes: label.slice(0,3),
    total: +ativos.filter(a=>a.qtd>0 && mesesIdx(a.meses).includes(idx)).reduce((s,a)=>s+a.prov*a.qtd,0).toFixed(2)
  }));
  const temDados = dataMes.some(d=>d.total>0);
  if (!temDados) return null;
  return (
    <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:"14px 6px 8px", marginBottom:16 }}>
      <div className={estiloDe("tituloGraficoMensal",11).cls} style={{ paddingLeft:8, marginBottom:8, fontSize:11, color:T.textMute, textTransform:"uppercase", letterSpacing:1, ...estiloDe("tituloGraficoMensal",11).style }}>💵 Proventos mês a mês</div>
      <div style={{ width:"100%", height:150 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={dataMes} margin={{ top:4,right:8,left:0,bottom:0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
            <XAxis dataKey="mes" tick={{ fontSize:8,fill:T.textMute }} axisLine={false} tickLine={false} interval={0}/>
            <YAxis tick={{ fontSize:8,fill:T.textMute }} axisLine={false} tickLine={false} width={38} tickFormatter={v=>v>=1000?`${(v/1000).toFixed(0)}k`:`${Math.round(v)}`}/>
            <Tooltip formatter={(v)=>fmt(v)} contentStyle={{ background:T.bg, border:`1px solid ${T.borderSoft}`, borderRadius:8, fontSize:12 }} cursor={{ fill:`${T.accent}11` }}/>
            <Bar dataKey="total" radius={[4,4,0,0]} fill={T.cyan}/>
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div style={{ fontSize:9, color:T.textFaint, padding:"4px 10px 0" }}>Visão completa e detalhada na tela 🔬 Análises.</div>
    </div>
  );
}

// Registro dos blocos móveis: id, nome, página padrão e ordem padrão.
// A base do app usa SEMPRE este padrão se a configuração estiver vazia/quebrada.
const BLOCOS_PAGINAS = ["painel","analises","calendario","ranking","custovida"];
const BLOCOS_DEF = [
  { id:"prox3meses",       nome:"Proventos · 3 meses",    emoji:"📆", paginaPadrao:"painel", ordemPadrao:1 },
  { id:"graficoMensal",    nome:"Proventos mês a mês",    emoji:"💵", paginaPadrao:"painel", ordemPadrao:2 },
  { id:"projecao",         nome:"Projeção + contribuição",emoji:"🔮", paginaPadrao:"painel", ordemPadrao:3 },
  { id:"previstoRealizado",nome:"Previsto vs Realizado",  emoji:"⚖️", paginaPadrao:"painel", ordemPadrao:4 },
];

// ════════════════════════════════════════════════════════════════════════════
// MODO DE EDIÇÃO VISUAL — envolve elementos e mostra o ✏️ lapisinho ao editar
// ════════════════════════════════════════════════════════════════════════════
function Editavel({ id, modoEdicao, onEditar, children, inline=false, T }) {
  if (!modoEdicao || !ELEMENTOS_ESTILO[id]) return children;
  const def = ELEMENTOS_ESTILO[id];
  return (
    <span style={{ position:"relative", display: inline ? "inline-block" : "block" }}>
      {/* contorno tracejado indicando que é editável */}
      <span style={{ position:"absolute", inset:-3, border:`1.5px dashed ${T.accent}88`, borderRadius:8, pointerEvents:"none", zIndex:1 }}/>
      {children}
      <button
        onClick={(e)=>{ e.stopPropagation(); e.preventDefault(); onEditar(id); }}
        title={`Editar: ${def.nome}`}
        style={{
          position:"absolute", top:-9, right:-9, zIndex:3,
          width:22, height:22, borderRadius:"50%", border:"none",
          background:T.accent, color:"#fff", cursor:"pointer",
          fontSize:11, lineHeight:1, display:"flex", alignItems:"center", justifyContent:"center",
          boxShadow:`0 2px 8px ${T.accent}88`, padding:0,
        }}
      >✏️</button>
    </span>
  );
}

// Popup com TODAS as customizações possíveis do elemento escolhido
function PopupCustomizar({ id, estilosCustom, setEstilosCustom, onClose, T }) {
  const def = ELEMENTOS_ESTILO[id] || {};
  const atual = (estilosCustom||{})[id] || {};
  const set = (patch) => setEstilosCustom(prev=>({ ...(prev||{}), [id]: { ...((prev||{})[id]||{}), ...patch } }));
  const resetar = () => setEstilosCustom(prev=>{ const c={ ...(prev||{}) }; delete c[id]; return Object.keys(c).length?c:null; });

  const CORES = ["#e74c3c","#e67e22","#f1c40f","#2ecc71","#1abc9c","#3498db","#9b59b6","#e91e63","#607d8b","#111827","#ffffff"];
  const FUNDOS = [
    { lb:"Nenhum", v:null },
    { lb:"Roxo", v:"linear-gradient(135deg,#667eea,#764ba2)" },
    { lb:"Oceano", v:"linear-gradient(135deg,#2193b0,#6dd5ed)" },
    { lb:"Pôr do sol", v:"linear-gradient(135deg,#ff6a00,#ee0979)" },
    { lb:"Verde", v:"linear-gradient(135deg,#11998e,#38ef7d)" },
    { lb:"Ouro", v:"linear-gradient(135deg,#f7971e,#ffd200)" },
    { lb:"Escuro", v:"#1a1a2e" },
  ];
  const FONTES = [
    { lb:"Padrão", v:null },
    { lb:"Serifada", v:"Georgia, serif" },
    { lb:"Mono", v:"monospace" },
    { lb:"Sistema", v:"system-ui, sans-serif" },
  ];
  const rot = (t) => <div style={{ fontSize:10, color:T.textMute, fontWeight:700, textTransform:"uppercase", letterSpacing:0.5, margin:"14px 0 7px" }}>{t}</div>;
  const chip = (ativo, onClick, children, extra={}) => (
    <button onClick={onClick} style={{ padding:"7px 11px", borderRadius:8, cursor:"pointer", fontSize:11, fontWeight:700,
      border:`1px solid ${ativo?T.accent:T.border}`, background: ativo?`${T.accent}18`:T.cardAlt, color: ativo?T.accentSoft:T.textMute, ...extra }}>{children}</button>
  );
  const slider = (label, prop, min, max, step, padrao) => (
    <div style={{ marginBottom:10 }}>
      <div style={{ display:"flex", justifyContent:"space-between", fontSize:10, color:T.textMute, marginBottom:4 }}>
        <span>{label}</span><span style={{ fontWeight:700, color:T.text }}>{atual[prop] ?? padrao}</span>
      </div>
      <input type="range" min={min} max={max} step={step} value={atual[prop] ?? padrao}
        onChange={e=>set({ [prop]: parseFloat(e.target.value) })}
        style={{ width:"100%", accentColor:T.accent }}/>
    </div>
  );

  return (
    <div onClick={onClose} style={{ position:"fixed", inset:0, background:"#000b", zIndex:1600, display:"flex", alignItems:"flex-end", justifyContent:"center", padding:0 }}>
      <motion.div
        onClick={e=>e.stopPropagation()}
        initial={{ y:"100%" }} animate={{ y:0 }} exit={{ y:"100%" }}
        transition={{ type:"spring", stiffness:340, damping:32 }}
        style={{ background:T.bg, borderRadius:"20px 20px 0 0", width:"100%", maxWidth:460, maxHeight:"85vh", overflowY:"auto", padding:"18px 18px 26px", boxShadow:"0 -10px 40px #000a" }}
      >
        <div style={{ width:38, height:4, borderRadius:2, background:T.border, margin:"0 auto 14px" }}/>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:4 }}>
          <div>
            <div style={{ fontSize:16, fontWeight:800, color:T.text }}>🎨 {def.nome}</div>
            <div style={{ fontSize:9, color:T.textFaint }}>{def.grupo}</div>
          </div>
          <button onClick={onClose} style={{ width:32, height:32, borderRadius:8, border:`1px solid ${T.border}`, background:T.cardAlt, color:T.text, cursor:"pointer", fontSize:15 }}>✕</button>
        </div>

        {rot("Tamanho e peso")}
        {slider("Escala da fonte", "escala", 0.5, 3, 0.05, 1)}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {chip(atual.negrito, ()=>set({ negrito:!atual.negrito }), "B", { fontWeight:900 })}
          {chip(atual.italico, ()=>set({ italico:!atual.italico }), "I", { fontStyle:"italic" })}
          {chip(atual.sublinhado, ()=>set({ sublinhado:!atual.sublinhado }), "U", { textDecoration:"underline" })}
          {chip(atual.maiuscula, ()=>set({ maiuscula:!atual.maiuscula }), "MAIÚSC")}
        </div>

        {rot("Cor do texto")}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {chip(!atual.cor, ()=>set({ cor:null }), "Padrão")}
          {CORES.map(c=>(
            <button key={c} onClick={()=>set({ cor:c })} style={{ width:30, height:30, borderRadius:"50%", cursor:"pointer",
              border: atual.cor===c ? `3px solid ${T.accent}` : `1px solid ${T.border}`, background:c }}/>
          ))}
        </div>

        {rot("Alinhamento")}
        <div style={{ display:"flex", gap:6 }}>
          {[["left","⬅️ Esq"],["center","↔️ Centro"],["right","➡️ Dir"]].map(([v,lb])=>chip(atual.alinhamento===v, ()=>set({ alinhamento: atual.alinhamento===v?null:v }), lb))}
        </div>

        {rot("Fonte")}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {FONTES.map(f=>chip(atual.fonte===f.v, ()=>set({ fonte:f.v }), f.lb, f.v?{ fontFamily:f.v }:{}))}
        </div>

        {rot("Fundo")}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {FUNDOS.map(f=>(
            <button key={f.lb} onClick={()=>set({ fundo:f.v })} style={{ padding:"9px 12px", borderRadius:8, cursor:"pointer", fontSize:10, fontWeight:700,
              border: atual.fundo===f.v ? `2px solid ${T.accent}` : `1px solid ${T.border}`,
              background: f.v || T.cardAlt, color: f.v ? "#fff" : T.textMute }}>{f.lb}</button>
          ))}
        </div>

        {rot("Caixa")}
        {slider("Cantos arredondados", "raio", 0, 40, 1, 0)}
        {slider("Espaço interno", "padding", 0, 40, 1, 0)}
        {slider("Opacidade", "opacidade", 0.1, 1, 0.05, 1)}
        {slider("Rotação (graus)", "rotacao", -30, 30, 1, 0)}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginTop:6 }}>
          {chip(!!atual.borda, ()=>set({ borda: atual.borda ? null : `2px solid ${T.accent}` }), "Borda")}
          {chip(!!atual.sombra, ()=>set({ sombra: atual.sombra ? null : "0 8px 24px #0005" }), "Sombra")}
          {chip(!!atual.oculto, ()=>set({ oculto:!atual.oculto }), atual.oculto ? "🚫 Oculto" : "👁 Visível")}
        </div>

        {rot("Animação")}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {ANIMACOES_IA.map(a=>chip((atual.animacao||"nenhuma")===a, ()=>set({ animacao:a }), a))}
        </div>

        <div style={{ display:"flex", gap:8, marginTop:20 }}>
          <button onClick={resetar} style={{ flex:1, padding:"13px", borderRadius:10, border:`1px solid ${T.border}`, background:T.cardAlt, color:T.textMute, cursor:"pointer", fontSize:12, fontWeight:700 }}>♻️ Restaurar padrão</button>
          <button onClick={onClose} style={{ flex:1, padding:"13px", borderRadius:10, border:"none", background:T.accent, color:"#fff", cursor:"pointer", fontSize:12, fontWeight:700 }}>✓ Pronto</button>
        </div>
      </motion.div>
    </div>
  );
}

function PainelCarteira({ ativos, historico = [], proventosRecebidos, blocosRender, estiloDe = () => ({ style:{}, cls:"" }), modoEdicao=false, onEditar=()=>{}, T }) {
  const [agrupar, setAgrupar] = useState("cat");   // "cat" | "setor"
  const [metrica, setMetrica] = useState("atual"); // "atual" | "investido"
  const [filtroCat, setFiltroCat] = useState("TUDO"); // "TUDO" | "FII" | "Ação" | "Cripto"

  // cores por categoria
  const corCat = { "FII":T.cyan, "Ação":T.accent, "Cripto":T.amber };

  // aplica filtro de categoria
  const lista = filtroCat==="TUDO" ? ativos : ativos.filter(a=>a.cat===filtroCat);

  // agrupa e soma
  const grupos = {};
  lista.forEach(a => {
    const chave = agrupar==="cat" ? a.cat : a.setor;
    const atual = a.qtd * a.cotacao;
    const investido = a.qtd * a.precoMedio;
    if (!grupos[chave]) grupos[chave] = { nome:chave, atual:0, investido:0, qtdAtivos:0 };
    grupos[chave].atual += atual;
    grupos[chave].investido += investido;
    grupos[chave].qtdAtivos += 1;
  });
  const arr = Object.values(grupos)
    .map(g => ({ ...g, valor: metrica==="atual"?g.atual:g.investido, resultado:g.atual-g.investido }))
    .sort((a,b)=>b.valor-a.valor);

  const totalAtual = lista.reduce((s,a)=>s+a.qtd*a.cotacao,0);
  const totalInvest = lista.reduce((s,a)=>s+a.qtd*a.precoMedio,0);
  const totalMetrica = metrica==="atual"?totalAtual:totalInvest;
  const resultadoGeral = totalAtual-totalInvest;
  const mediaMesPainel = lista.reduce((s,a)=>s+a.prov*a.qtd*a.meses.length,0)/12; // proventos médios por mês

  const corGrupo = (nome) => agrupar==="cat" ? (corCat[nome]||T.accent) : (COR_SETOR[nome]||T.textMute);

  const pieData = arr.map(g => ({ name:g.nome, value:+g.valor.toFixed(2), cor:corGrupo(g.nome) }));
  const maxBar = Math.max(...arr.map(g=>g.valor), 1);

  const FILTROS_CAT = [
    { id:"TUDO", label:"Tudo" },
    { id:"FII",  label:"FIIs" },
    { id:"Ação", label:"Ações" },
    { id:"Cripto",label:"Cripto" },
  ];

  const [vista, setVista] = useState("proventos"); // proventos | resumo | composicao | historico
  const MINI = [
    { id:"proventos",  label:"Proventos",  emoji:"💰", destaque:true },
    { id:"resumo",     label:"Resumo",     emoji:"👁️" },
    { id:"composicao", label:"Composição", emoji:"🥧" },
    { id:"historico",  label:"Histórico",  emoji:"🕒" },
  ];

  return (
    <div>
      {/* SAUDAÇÃO + DATA + DESTAQUE DO DIA */}
      {(() => {
        const hoje = new Date();
        const h = hoje.getHours();
        const saud = h<12 ? "Bom dia" : h<18 ? "Boa tarde" : "Boa noite";
        const NOMES_MES = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
        const dataFmt = `${hoje.getDate()} de ${NOMES_MES[hoje.getMonth()]}`;
        const mAtual = hoje.getMonth()+1;

        // candidatos a destaque do dia
        const candidatos = [];
        const pagadores = ativos.filter(a=>a.qtd>0 && a.prov>0 && a.meses.includes(mAtual)).map(a=>({t:a.ticker, v:a.prov*a.qtd})).sort((x,y)=>y.v-x.v);
        if (pagadores[0]) candidatos.push({ icon:"💰", txt:`Seu maior pagador este mês é o ${pagadores[0].t} (${fmt(pagadores[0].v)})` });
        let magicProx=null;
        ativos.forEach(a=>{
          const pm=a.prov*(a.meses.length/12);
          if(pm>0 && a.cotacao>0){ const magic=Math.ceil(a.cotacao/pm); const faltam=magic-a.qtd;
            if(faltam>0 && (!magicProx||faltam<magicProx.faltam)) magicProx={t:a.ticker, faltam}; }
        });
        if (magicProx && magicProx.faltam<=15) candidatos.push({ icon:"✨", txt:`Você está a ${magicProx.faltam} cota${magicProx.faltam>1?"s":""} do Magic Number do ${magicProx.t}!` });
        const patri = ativos.reduce((s,a)=>s+a.qtd*a.cotacao,0);
        const porA = ativos.map(a=>({t:a.ticker, v:a.qtd*a.cotacao})).sort((x,y)=>y.v-x.v);
        if (porA[0] && patri>0) { const cc=porA[0].v/patri*100;
          if (cc>30) candidatos.push({ icon:"⚖️", txt:`${porA[0].t} já é ${cc.toFixed(0)}% da carteira — talvez diversificar.` });
        }
        const totProx = ativos.filter(a=>a.qtd>0 && a.prov>0 && a.meses.includes(mAtual===12?1:mAtual+1)).reduce((s,a)=>s+a.prov*a.qtd,0);
        if (totProx>0) candidatos.push({ icon:"📅", txt:`No próximo mês você deve receber cerca de ${fmt(totProx)}.` });
        if (candidatos.length===0) candidatos.push({ icon:"🌱", txt:"Cadastre seus ativos para ver destaques personalizados." });
        // varia o destaque conforme o dia
        const destaque = candidatos[hoje.getDate() % candidatos.length];

        return (
          <div style={{ marginBottom:20, marginTop:8 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"baseline", marginBottom:14 }}>
              <Editavel id="saudacao" modoEdicao={modoEdicao} onEditar={onEditar} inline T={T}><span className={estiloDe("saudacao",24).cls} style={{ fontSize:24, fontWeight:800, color:T.text, letterSpacing:-0.5, ...estiloDe("saudacao",24).style }}>{saud}! 👋</span></Editavel>
              <Editavel id="dataHoje" modoEdicao={modoEdicao} onEditar={onEditar} inline T={T}><span className={estiloDe("dataHoje",12).cls} style={{ fontSize:12, color:T.textFaint, ...estiloDe("dataHoje",12).style }}>{dataFmt}</span></Editavel>
            </div>
            <Editavel id="destaqueDia" modoEdicao={modoEdicao} onEditar={onEditar} T={T}><div className={estiloDe("destaqueDia",11).cls} style={{ background:`linear-gradient(135deg, ${T.amber}1c, ${T.card})`, border:`1px solid ${T.amber}44`, borderRadius:12, padding:"11px 13px", display:"flex", alignItems:"center", gap:10, ...estiloDe("destaqueDia",11).style }}>
              <span style={{ fontSize:20 }}>{destaque.icon}</span>
              <div style={{ flex:1 }}>
                <div style={{ fontSize:8, color:T.amber, fontWeight:700, textTransform:"uppercase", letterSpacing:1 }}>Destaque do dia</div>
                <div style={{ fontSize:12, color:T.text, fontWeight:600, lineHeight:1.3 }}>{destaque.txt}</div>
              </div>
            </div></Editavel>
          </div>
        );
      })()}

      {/* mini-abas internas do Painel — carrossel deslizável, Proventos em destaque */}
      <div style={{ display:"flex", gap:7, overflowX:"auto", paddingBottom:6, marginBottom:18, scrollbarWidth:"none" }}>
        {MINI.map(m=>{
          const sel = vista===m.id;
          if (m.destaque) {
            // botão Proventos: maior, chapado, verde, sempre primeiro
            return (
              <button key={m.id} onClick={()=>setVista(m.id)} style={{
                flexShrink:0, padding:"11px 20px", borderRadius:12, border:"none", cursor:"pointer",
                fontSize:14, fontWeight:800, whiteSpace:"nowrap",
                background: sel ? T.green : `${T.green}1f`,
                color: sel ? "#06281b" : T.green,
                boxShadow: sel ? `0 4px 14px ${T.green}55` : "none",
              }}>{m.emoji} {m.label}</button>
            );
          }
          return (
            <button key={m.id} onClick={()=>setVista(m.id)} style={{
              flexShrink:0, padding:"11px 16px", borderRadius:12, border:`1px solid ${sel?T.accent:T.border}`, cursor:"pointer",
              fontSize:12, fontWeight:700, whiteSpace:"nowrap",
              background: sel?T.accentBg:T.card, color: sel?T.accentSoft:T.textMute
            }}>{m.emoji} {m.label}</button>
          );
        })}
      </div>

      {/* ═══ VISTA: PROVENTOS — blocos configuráveis (Camada 2) ═══ */}
      {vista==="proventos" && (<>
      {ativos.filter(a=>a.qtd>0).length===0 && (
        <div style={{ background:T.card, border:`2px dashed ${T.accentBorder}`, borderRadius:16, padding:"26px 18px", textAlign:"center", marginBottom:16 }}>
          <div style={{ fontSize:38, marginBottom:8 }}>🌱</div>
          <div style={{ fontSize:15, fontWeight:800, color:T.text, marginBottom:6 }}>Sua carteira está vazia</div>
          <div style={{ fontSize:11, color:T.textMute, lineHeight:1.6, marginBottom:14 }}>Adicione seu primeiro ativo (temos um catálogo com sugestões) ou importe a planilha da B3 na tela ✏️ Editar.</div>
          <button onClick={()=>abrirEditorAtivo(null)} style={{ padding:"12px 20px", borderRadius:10, border:"none", background:T.accent, color:"#fff", cursor:"pointer", fontSize:13, fontWeight:700 }}>➕ Adicionar primeiro ativo</button>
        </div>
      )}
      {blocosRender
        ? blocosRender
        : (<>{/* fallback: ordem padrão da instalação (base intacta) */}
          <BlocoProx3Meses ativos={ativos} estiloDe={estiloDe} T={T}/>
          <BlocoGraficoMensal ativos={ativos} estiloDe={estiloDe} T={T}/>
          <ProjecaoProventos ativos={ativos} T={T}/>
          <PrevistoVsRealizado ativos={ativos} proventosRecebidos={proventosRecebidos} estiloDe={estiloDe} T={T}/>
        </>)}
      </>)}

      {/* ═══ VISTA: RESUMO (patrimônio total + médias) ═══ */}
      {vista==="resumo" && (<>
      {/* resumo geral */}
      <div style={{ background:`linear-gradient(135deg, ${T.accent}22, ${T.card})`, border:`1px solid ${T.border}`, borderRadius:14, padding:"18px", marginBottom:16 }}>
        <Editavel id="patrimonioLabel" modoEdicao={modoEdicao} onEditar={onEditar} T={T}><div className={estiloDe("patrimonioLabel",10).cls} style={{ fontSize:10, color:T.textFaint, textTransform:"uppercase", letterSpacing:1, ...estiloDe("patrimonioLabel",10).style }}>💼 Patrimônio total</div></Editavel>
        <div style={{ fontSize:30, fontWeight:800, color:T.text, letterSpacing:-1, lineHeight:1.1 }}>{fmt(totalAtual)}</div>
        <div style={{ fontSize:11, color:T.green, marginTop:3 }}>Proventos médios: <strong>{fmt(mediaMesPainel)}/mês</strong></div>
        <div style={{ display:"flex", gap:8, marginTop:12, flexWrap:"wrap" }}>
          <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:8, padding:"6px 11px" }}>
            <div style={{ fontSize:9, color:T.textFaint }}>Investido</div>
            <div style={{ fontSize:12, fontWeight:700, color:T.textDim }}>{fmt(totalInvest)}</div>
          </div>
          <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:8, padding:"6px 11px" }}>
            <div style={{ fontSize:9, color:T.textFaint }}>Resultado</div>
            <div style={{ fontSize:12, fontWeight:700, color:resultadoGeral>=0?T.green:T.red }}>{resultadoGeral>=0?"+":""}{fmt(resultadoGeral)}</div>
          </div>
          <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:8, padding:"6px 11px" }}>
            <div style={{ fontSize:9, color:T.textFaint }}>Rentab.</div>
            <div style={{ fontSize:12, fontWeight:700, color:resultadoGeral>=0?T.green:T.red }}>{totalInvest>0?`${resultadoGeral>=0?"+":""}${((resultadoGeral/totalInvest)*100).toFixed(1)}%`:"—"}</div>
          </div>
        </div>
      </div>

      </>)}

      {/* ═══ VISTA: HISTÓRICO ═══ */}
      {vista==="historico" && (
        historico.length>0 ? (
          <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:"14px 6px 12px", marginBottom:16 }}>
          <div style={{ paddingLeft:8, marginBottom:8, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <div style={{ fontSize:11, color:T.textMute, textTransform:"uppercase", letterSpacing:1 }}>🕒 Histórico real</div>
            <span style={{ fontSize:9, color:T.textFaint }}>{historico.length} mês{historico.length>1?"es":""} registrado{historico.length>1?"s":""}</span>
          </div>
          {historico.length===1 ? (
            <div style={{ fontSize:10, color:T.textFaint, padding:"8px 12px", lineHeight:1.6 }}>
              📸 Primeiro retrato guardado! O app grava a evolução do seu patrimônio automaticamente a cada mês. Volte nos próximos meses para ver a linha crescer.
            </div>
          ) : (
            <div style={{ width:"100%", height:180 }}>
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={historico.map(h=>({ mes:h.mes.slice(5)+"/"+h.mes.slice(2,4), patrimonio:h.patrimonio }))} margin={{ top:8,right:12,left:0,bottom:0 }}>
                  <defs>
                    <linearGradient id="gradHist" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={T.accent} stopOpacity={0.4}/>
                      <stop offset="100%" stopColor={T.accent} stopOpacity={0.02}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
                  <XAxis dataKey="mes" tick={{ fontSize:8,fill:T.textMute }} axisLine={false} tickLine={false}/>
                  <YAxis tick={{ fontSize:8,fill:T.textMute }} axisLine={false} tickLine={false} width={42}
                    tickFormatter={v=>v>=1000?`${(v/1000).toFixed(0)}k`:`${Math.round(v)}`}/>
                  <Tooltip formatter={(v)=>fmt(v)} contentStyle={{ background:T.bg, border:`1px solid ${T.borderSoft}`, borderRadius:8, fontSize:12 }}/>
                  <Area type="monotone" dataKey="patrimonio" name="Patrimônio" stroke={T.accent} strokeWidth={2.5} fill="url(#gradHist)" dot={{ r:3, fill:T.accent }}/>
                </AreaChart>
              </ResponsiveContainer>
            </div>
          )}
          <div style={{ fontSize:9, color:T.textFaint, padding:"4px 12px 0", lineHeight:1.5 }}>
            Diferente das projeções: aqui é o valor real registrado a cada mês conforme você usa o app e atualiza as cotações.
          </div>
        </div>
        ) : (
          <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:"24px 16px", textAlign:"center" }}>
            <div style={{ fontSize:28, marginBottom:6 }}>🕒</div>
            <div style={{ fontSize:12, color:T.textMute, lineHeight:1.6 }}>O histórico começa a ser gravado automaticamente. Volte nos próximos meses para ver a evolução real do seu patrimônio.</div>
          </div>
        )
      )}

      {/* ═══ VISTA: COMPOSIÇÃO ═══ */}
      {vista==="composicao" && (<>
      {/* controles */}
      <div style={{ display:"flex", gap:6, marginBottom:8, flexWrap:"wrap" }}>
        {FILTROS_CAT.map(f=>(
          <button key={f.id} onClick={()=>setFiltroCat(f.id)} style={{
            padding:"6px 12px", borderRadius:8, border:"none", cursor:"pointer", fontSize:11, fontWeight:700,
            background:filtroCat===f.id?T.accent:T.border, color:filtroCat===f.id?"#fff":T.textMute
          }}>{f.label}</button>
        ))}
      </div>
      <div style={{ display:"flex", gap:8, marginBottom:20 }}>
        <div style={{ flex:1, display:"flex", gap:4, background:T.card, borderRadius:8, padding:3 }}>
          {[{id:"cat",l:"Por categoria"},{id:"setor",l:"Por setor"}].map(o=>(
            <button key={o.id} onClick={()=>setAgrupar(o.id)} style={{ flex:1, padding:"6px 4px", borderRadius:6, border:"none", cursor:"pointer", fontSize:10, fontWeight:700, background:agrupar===o.id?T.accentBg:"transparent", color:agrupar===o.id?T.accentSoft:T.textMute }}>{o.l}</button>
          ))}
        </div>
        <div style={{ flex:1, display:"flex", gap:4, background:T.card, borderRadius:8, padding:3 }}>
          {[{id:"atual",l:"Valor atual"},{id:"investido",l:"Investido"}].map(o=>(
            <button key={o.id} onClick={()=>setMetrica(o.id)} style={{ flex:1, padding:"6px 4px", borderRadius:6, border:"none", cursor:"pointer", fontSize:10, fontWeight:700, background:metrica===o.id?T.accentBg:"transparent", color:metrica===o.id?T.accentSoft:T.textMute }}>{o.l}</button>
          ))}
        </div>
      </div>

      {/* PIZZA — composição */}
      <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:"14px", marginBottom:16 }}>
        <div style={{ fontSize:11, color:T.textMute, marginBottom:8, textTransform:"uppercase", letterSpacing:1 }}>
          Composição {agrupar==="cat"?"por categoria":"por setor"}
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
          <div style={{ width:160, height:160, flexShrink:0, margin:"0 auto" }}>
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={42} outerRadius={72} paddingAngle={2} stroke="none">
                  {pieData.map((e,i)=><Cell key={i} fill={e.cor}/>)}
                </Pie>
                <Tooltip formatter={(v)=>fmt(v)} contentStyle={{ background:T.bg, border:`1px solid ${T.borderSoft}`, borderRadius:8, fontSize:12 }}/>
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div style={{ flex:1, minWidth:140 }}>
            {arr.map(g=>{
              const pct = totalMetrica>0 ? (g.valor/totalMetrica)*100 : 0;
              return (
                <div key={g.nome} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
                  <div style={{ width:10, height:10, borderRadius:3, background:corGrupo(g.nome), flexShrink:0 }}/>
                  <div style={{ flex:1 }}>
                    <div style={{ display:"flex", justifyContent:"space-between" }}>
                      <span style={{ fontSize:11, color:T.textDim, fontWeight:600 }}>{g.nome}</span>
                      <span style={{ fontSize:11, color:T.text, fontWeight:700 }}>{pct.toFixed(1)}%</span>
                    </div>
                    <div style={{ fontSize:9, color:T.textFaint }}>{fmt(g.valor)}</div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* BARRAS — dinheiro por grupo */}
      <div style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:"14px", marginBottom:16 }}>
        <div style={{ fontSize:11, color:T.textMute, marginBottom:16, textTransform:"uppercase", letterSpacing:1 }}>
          {metrica==="atual"?"Valor atual":"Investido"} por {agrupar==="cat"?"categoria":"setor"}
        </div>
        {arr.map(g=>{
          const pct = (g.valor/maxBar)*100;
          return (
            <div key={g.nome} style={{ marginBottom:10 }}>
              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                <span style={{ fontSize:11, color:T.textDim, fontWeight:600 }}>{g.nome} <span style={{ fontSize:9, color:T.textFaint }}>· {g.qtdAtivos} ativo{g.qtdAtivos>1?"s":""}</span></span>
                <div style={{ textAlign:"right" }}>
                  <span style={{ fontSize:11, color:T.text, fontWeight:700 }}>{fmt(g.valor)}</span>
                  <span style={{ fontSize:9, color:g.resultado>=0?T.green:T.red, marginLeft:6 }}>{g.resultado>=0?"+":""}{fmtK(g.resultado)}</span>
                </div>
              </div>
              <div style={{ height:8, background:T.cardAlt, borderRadius:5, overflow:"hidden" }}>
                <div style={{ height:"100%", width:`${pct}%`, background:corGrupo(g.nome), borderRadius:5, transition:"width 0.4s" }}/>
              </div>
            </div>
          );
        })}
        {arr.length===0 && <div style={{ fontSize:11, color:T.textFaint, textAlign:"center", padding:"16px 0" }}>Nenhum ativo neste filtro</div>}
      </div>
      </>)}
    </div>
  );
}

// ════════════════════════════════════════════════════════════════════════════
// ANÁLISE DA CARTEIRA — yield, saúde, benchmarks (usada na aba Análises)
// ════════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════════
// PROJEÇÃO DE PROVENTOS — estimativa do próximo mês (card verde), dados ao vivo
// ════════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════════
// TELA "MINHA CARTEIRA" — página de distribuição por classe (estilo dedicado)
// Acessível pelo ícone 💼 no cabeçalho. Usa os dados reais dos ativos.
// ════════════════════════════════════════════════════════════════════════════
function classeDe(a) {
  if (a.cat === "FII") return "Imobiliário";
  if (a.cat === "Cripto" || /^(COIN11|HASH11|QBTC|BITH11|QETH)/.test(a.ticker)) return "Criptoativos";
  if (/tesouro/i.test(a.nome) || /^(LFT|LTN|NTN|TD)/.test(a.ticker)) {
    return /prefix|ltn|ntn-f/i.test(a.nome+a.ticker) ? "Renda Fixa Prefixada" : "Renda Fixa Inflação";
  }
  if (/^(IVVB11|NASD11|XINA11|PAGS34|ROXO34|AAPL34|MSFT34|GOGL34|AMZO34)/.test(a.ticker) || /s&p|nasdaq|internacional|\(bdr\)/i.test(a.nome)) return "Internacional";
  if (/^(GOLD11|OZ1|AURA|GOLD)/.test(a.ticker) || /ouro|gold/i.test(a.nome)) return "Ouro";
  return "Ações";
}
const CLASSE_COR = {
  "Ações":"#6366f1", "Imobiliário":"#22d3ee", "Criptoativos":"#f59e0b",
  "Renda Fixa Inflação":"#34d399", "Renda Fixa Prefixada":"#a78bfa",
  "Internacional":"#60a5fa", "Ouro":"#fbbf24",
};
const ORDEM_CLASSE = ["Ações","Imobiliário","Criptoativos","Renda Fixa Inflação","Renda Fixa Prefixada","Internacional","Ouro"];

function TelaCarteira({ ativos, onClose, onEditar, T }) {
  const [tab, setTab] = useState("resumo");
  const [sortA, setSortA] = useState("valor");
  const [sortF, setSortF] = useState("valor");

  const C = {
    bg: T.bg, header: T.bgHeader, card: T.card, borda: T.border,
    txt: T.text, txtDim: T.textDim, txtFaint: T.textFaint,
    roxo: T.accent, verde: T.green, vermelho: T.red,
  };
  const pctFmt = v => v!=null ? `${v>0?"+":""}${v.toFixed(2)}%` : "—";

  // calcula por ativo
  const comCalc = ativos.filter(a=>a.qtd>0).map(a=>({
    ...a, classe:classeDe(a),
    valor:+(a.qtd*a.cotacao).toFixed(2),
    custo:+(a.qtd*a.precoMedio).toFixed(2),
    variacao:+((a.qtd*a.cotacao)-(a.qtd*a.precoMedio)).toFixed(2),
    varPct:a.precoMedio>0 ? ((a.cotacao-a.precoMedio)/a.precoMedio*100) : 0,
  }));
  const total = comCalc.reduce((s,a)=>s+a.valor,0);

  // agrupa por classe
  const porClasse = ORDEM_CLASSE.map(cl=>{
    const itens = comCalc.filter(a=>a.classe===cl);
    const v = itens.reduce((s,a)=>s+a.valor,0);
    const variacao = itens.reduce((s,a)=>s+a.variacao,0);
    return { classe:cl, valor:v, variacao, percentual: total>0?v/total*100:0, n:itens.length };
  }).filter(c=>c.valor>0);

  const acoes = comCalc.filter(a=>a.classe==="Ações" || a.classe==="Internacional" || a.classe==="Ouro");
  const fiis = comCalc.filter(a=>a.classe==="Imobiliário");
  const totAcoes = acoes.reduce((s,a)=>s+a.valor,0);
  const totFiis = fiis.reduce((s,a)=>s+a.valor,0);
  const provMes = ativos.reduce((s,a)=>{
    const m=new Date().getMonth()+1;
    return s + (a.meses?.includes(m) ? a.prov*a.qtd : 0);
  },0);

  const ordenar = (arr,k) => [...arr].sort((x,y)=> k==="valor"?y.valor-x.valor : k==="var"?(x.varPct-y.varPct) : x.ticker.localeCompare(y.ticker));

  const Linha = ({ a }) => (
    <div onClick={()=>abrirEditorAtivo(a.ticker)} style={{ display:"grid", gridTemplateColumns:"1fr 44px 70px 70px 58px", gap:6, background:C.card, border:`1px solid ${C.borda}`, borderRadius:10, padding:"11px 12px", marginBottom:6, alignItems:"center", cursor:"pointer" }}>
      <div>
        <div style={{ fontWeight:700, fontSize:13, color:C.txt }}>{a.ticker}</div>
        <div style={{ fontSize:10, color:C.txtFaint, marginTop:1 }}>{a.nome}</div>
        <div style={{ fontSize:11, color:C.roxo, marginTop:1, fontWeight:600 }}>{fmt(a.valor)}</div>
      </div>
      <div style={{ fontSize:12, color:C.txtDim, textAlign:"center" }}>{a.qtd}</div>
      <div style={{ fontSize:11, color:C.txtDim, textAlign:"right" }}>R${a.precoMedio.toFixed(2)}</div>
      <div style={{ fontSize:11, color:C.txt, textAlign:"right" }}>R${a.cotacao.toFixed(2)}</div>
      <div style={{ fontSize:12, fontWeight:700, color:a.varPct>=0?C.verde:C.vermelho, textAlign:"right" }}>{pctFmt(a.varPct)}</div>
    </div>
  );
  const Cab = () => (
    <div style={{ display:"grid", gridTemplateColumns:"1fr 44px 70px 70px 58px", gap:6, padding:"6px 12px", marginBottom:4 }}>
      {["Ativo","Qtd","P.Médio","Cotação","Var%"].map(h=><div key={h} style={{ fontSize:10, color:C.txtFaint, fontWeight:600, textTransform:"uppercase" }}>{h}</div>)}
    </div>
  );
  const Ordenacao = ({ val, set }) => (
    <div style={{ display:"flex", gap:6, marginBottom:14, alignItems:"center", flexWrap:"wrap" }}>
      <span style={{ fontSize:11, color:C.txtFaint, marginRight:4 }}>Ordenar:</span>
      {[["valor","Valor"],["var","Variação"],["ticker","Ticker"]].map(([k,l])=>(
        <button key={k} onClick={()=>set(k)} style={{ padding:"5px 12px", borderRadius:6, border:"none", cursor:"pointer", fontSize:11, fontWeight:600, background:val===k?C.roxo:C.borda, color:val===k?"#fff":C.txtDim }}>{l}</button>
      ))}
    </div>
  );

  const MES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][new Date().getMonth()];

  return (
    <div className="modal-fullscreen" style={{ position:"fixed", inset:0, background:C.bg, zIndex:1300, overflowY:"auto", color:C.txt, fontFamily:"'Inter',system-ui,sans-serif", paddingBottom:48 }}>
      {/* Header */}
      <div style={{ background:C.header, padding:"22px 18px 18px", borderBottom:`1px solid ${C.borda}`, position:"relative" }}>
        <div style={{ position:"absolute", top:16, right:16, display:"flex", gap:8 }}>
          <button onClick={()=>{ onClose(); onEditar && onEditar(); }} title="Editar ativos" style={{ width:34, height:34, borderRadius:8, border:`1px solid ${C.borda}`, background:C.card, color:C.txt, cursor:"pointer", fontSize:15 }}>✏️</button>
          <button onClick={onClose} style={{ width:34, height:34, borderRadius:8, border:`1px solid ${C.borda}`, background:C.card, color:C.txt, cursor:"pointer", fontSize:16 }}>✕</button>
        </div>
        <div style={{ fontSize:11, color:C.txtFaint, letterSpacing:2, textTransform:"uppercase", marginBottom:4 }}>Minha Carteira · {MES} {new Date().getFullYear()}</div>
        <div style={{ fontSize:30, fontWeight:700, color:C.txt, letterSpacing:-1 }}>{fmt(total)}</div>
        <div style={{ display:"flex", gap:8, marginTop:12, flexWrap:"wrap" }}>
          {[{l:"Ações",v:fmt(totAcoes),c:C.roxo},{l:"FIIs",v:fmt(totFiis),c:T.cyan},{l:"Proventos/mês",v:fmt(provMes),c:C.verde}].map(x=>(
            <div key={x.l} style={{ background:T.cardAlt, borderRadius:8, padding:"6px 12px" }}>
              <div style={{ fontSize:10, color:C.txtFaint }}>{x.l}</div>
              <div style={{ fontSize:13, fontWeight:700, color:x.c }}>{x.v}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Tabs */}
      <div style={{ display:"flex", gap:2, padding:"10px 12px 0", background:C.bg, borderBottom:`1px solid ${C.borda}`, overflowX:"auto" }}>
        {[{id:"resumo",l:"Resumo"},{id:"acoes",l:"Ações"},{id:"fiis",l:"FIIs"},{id:"outros",l:"Outros"},{id:"proventos",l:"💰 Proventos"}].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{ padding:"8px 12px", borderRadius:8, border:"none", cursor:"pointer", fontSize:12, fontWeight:600, whiteSpace:"nowrap", background:tab===t.id?(t.id==="proventos"?T.green:C.roxo):"transparent", color:tab===t.id?"#fff":C.txtFaint }}>{t.l}</button>
        ))}
      </div>

      <div style={{ padding:"14px" }}>
        {/* RESUMO */}
        {tab==="resumo" && (
          <>
            <div style={{ background:C.card, border:`1px solid ${C.borda}`, borderRadius:12, padding:"14px", marginBottom:12 }}>
              <div style={{ fontSize:11, color:C.txtFaint, marginBottom:10, textTransform:"uppercase", letterSpacing:1 }}>Distribuição</div>
              <div style={{ display:"flex", height:10, borderRadius:8, overflow:"hidden", gap:2, marginBottom:12 }}>
                {porClasse.map(c=><div key={c.classe} style={{ flex:c.percentual, background:CLASSE_COR[c.classe], borderRadius:4 }}/>)}
              </div>
              <div style={{ display:"flex", flexWrap:"wrap", gap:"6px 14px" }}>
                {porClasse.map(c=>(
                  <div key={c.classe} style={{ display:"flex", alignItems:"center", gap:5 }}>
                    <div style={{ width:8, height:8, borderRadius:"50%", background:CLASSE_COR[c.classe] }}/>
                    <span style={{ fontSize:11, color:C.txtDim }}>{c.classe} <span style={{ color:C.txt, fontWeight:600 }}>{c.percentual.toFixed(1)}%</span></span>
                  </div>
                ))}
              </div>
            </div>
            {porClasse.map(c=>{
              const pos = c.variacao>=0;
              return (
                <div key={c.classe} style={{ background:C.card, border:`1px solid ${C.borda}`, borderRadius:12, padding:"13px 14px", marginBottom:8, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                  <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                    <div style={{ width:8, height:36, borderRadius:4, background:CLASSE_COR[c.classe] }}/>
                    <div>
                      <div style={{ fontWeight:600, fontSize:14, color:C.txt }}>{c.classe}</div>
                      <div style={{ fontSize:12, color:C.txtFaint, marginTop:2 }}>{c.percentual.toFixed(1)}% · {c.n} ativo{c.n!==1?"s":""}</div>
                    </div>
                  </div>
                  <div style={{ textAlign:"right" }}>
                    <div style={{ fontWeight:700, fontSize:15, color:C.txt }}>{fmt(c.valor)}</div>
                    <div style={{ fontSize:12, color:pos?C.verde:C.vermelho, marginTop:2 }}>{pos?"▲":"▼"} {fmt(Math.abs(c.variacao))}</div>
                  </div>
                </div>
              );
            })}
            <div style={{ background:`linear-gradient(135deg,${T.accent},${T.accentSoft})`, borderRadius:12, padding:"14px", marginTop:8, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <div style={{ fontSize:13, color:"#fff", fontWeight:600 }}>TOTAL GERAL</div>
              <div style={{ fontSize:20, fontWeight:800, color:"#fff" }}>{fmt(total)}</div>
            </div>
          </>
        )}
        {/* AÇÕES */}
        {tab==="acoes" && (
          <>
            <Ordenacao val={sortA} set={setSortA}/>
            <div style={{ background:C.card, border:`1px solid ${C.borda}`, borderRadius:10, padding:"11px 14px", marginBottom:12, display:"flex", justifyContent:"space-between" }}>
              <span style={{ color:C.txtFaint, fontSize:12 }}>Total · {acoes.length} ativos</span>
              <span style={{ color:C.roxo, fontWeight:700, fontSize:14 }}>{fmt(totAcoes)}</span>
            </div>
            <Cab/>
            {ordenar(acoes,sortA).map(a=><Linha key={a.ticker} a={a}/>)}
          </>
        )}
        {/* FIIs */}
        {tab==="fiis" && (
          <>
            <Ordenacao val={sortF} set={setSortF}/>
            <div style={{ background:C.card, border:`1px solid ${C.borda}`, borderRadius:10, padding:"11px 14px", marginBottom:12, display:"flex", justifyContent:"space-between" }}>
              <span style={{ color:C.txtFaint, fontSize:12 }}>Total · {fiis.length} fundos</span>
              <span style={{ color:T.cyan, fontWeight:700, fontSize:14 }}>{fmt(totFiis)}</span>
            </div>
            <Cab/>
            {ordenar(fiis,sortF).map(a=><Linha key={a.ticker} a={a}/>)}
          </>
        )}
        {/* OUTROS */}
        {tab==="outros" && (
          <>
            {["Criptoativos","Renda Fixa Inflação","Renda Fixa Prefixada","Internacional","Ouro"].map(cl=>{
              const itens = comCalc.filter(a=>a.classe===cl);
              if (itens.length===0) return null;
              const tot = itens.reduce((s,a)=>s+a.valor,0);
              return (
                <div key={cl} style={{ marginBottom:20 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
                    <div style={{ width:8, height:8, borderRadius:"50%", background:CLASSE_COR[cl] }}/>
                    <span style={{ fontSize:11, color:C.txtDim, fontWeight:700, textTransform:"uppercase", letterSpacing:1 }}>{cl}</span>
                    <span style={{ fontSize:13, color:CLASSE_COR[cl], fontWeight:700, marginLeft:"auto" }}>{fmt(tot)}</span>
                  </div>
                  {itens.map(a=>(
                    <div key={a.ticker} style={{ background:C.card, border:`1px solid ${C.borda}`, borderRadius:10, padding:"12px 14px", marginBottom:6 }}>
                     
