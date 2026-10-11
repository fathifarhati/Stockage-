<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gestion de Stock — Premium</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&family=Cairo:wght@500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<style>
  :root{
    --bg: #f2f4f8; --panel: #ffffff; --panel-alt: #f1f3f7; --border: #d5dbe5;
    --text: #1d2433; --text-muted: #5b6575; --accent: #d68c45; --blue: #1d5fb8;
    --danger: #d32f3f; --ok: #1f9d4a; --radius: 3px;
  }
  *{ box-sizing:border-box; }
  html,body{ margin:0; padding:0; }
  body{
    background:
      repeating-linear-gradient(0deg, #0000000a 0px, #0000000a 1px, transparent 1px, transparent 40px),
      repeating-linear-gradient(90deg, #0000000a 0px, #0000000a 1px, transparent 1px, transparent 40px),
      var(--bg);
    color: var(--text); font-family: 'IBM Plex Sans', sans-serif; min-height: 100vh; -webkit-font-smoothing: antialiased;
  }
  [dir="rtl"] body{ font-family:'Cairo','IBM Plex Sans', sans-serif; }
  h1,h2,h3,.brand{ font-family:'Space Grotesk', sans-serif; }
  [dir="rtl"] h1,[dir="rtl"] h2,[dir="rtl"] h3,[dir="rtl"] .brand{ font-family:'Cairo', sans-serif; }
  .num, .ref-cell, input[type=number], input[type=date]{ font-family:'IBM Plex Mono', monospace; }
  #app{ min-height:100vh; display:flex; flex-direction:column; }
  .topbar{ display:flex; align-items:center; justify-content:space-between; padding:14px 20px; border-bottom:1px solid var(--border); background:var(--panel); }
  .brand{ font-size:17px; font-weight:700; display:flex; align-items:center; gap:10px; }
  .brand .dot{ width:8px; height:8px; background:var(--accent); transform:rotate(45deg); flex-shrink:0; }
  .who{ font-size:13px; color:var(--text-muted); display:flex; align-items:center; gap:10px; }
  .who button, .lang-btn{ background:transparent; border:1px solid var(--border); color:var(--text-muted); padding:6px 10px; border-radius:var(--radius); font-size:12px; cursor:pointer; font-family:inherit; }
  .who button:hover{ border-color:var(--danger); color:var(--danger); }
  .lang-btn{ font-weight:600; }
  .lang-btn:hover{ border-color:var(--accent); color:var(--accent); }
  .login-lang{ position:absolute; top:18px; inset-inline-end:18px; }
  main{ flex:1; padding:24px 20px 60px; max-width:1080px; width:100%; margin:0 auto; }
  .login-wrap{ flex:1; display:flex; align-items:center; justify-content:center; padding:20px; position:relative; }
  .login-card{ background:var(--panel); border:1px solid var(--border); border-radius:var(--radius); padding:36px 32px; width:100%; max-width:360px; }
  .login-card .dot{ width:10px; height:10px; background:var(--accent); transform:rotate(45deg); margin-bottom:18px; }
  .login-card h1{ font-size:20px; margin:0 0 4px; }
  .login-card p.sub{ color:var(--text-muted); font-size:13px; margin:0 0 26px; }
  label{ display:block; font-size:12px; color:var(--text-muted); margin-bottom:6px; }
  input, select{ width:100%; background:var(--panel-alt); border:1px solid var(--border); color:var(--text); padding:10px 12px; border-radius:var(--radius); font-size:14px; font-family:inherit; margin-bottom:16px; outline:none; }
  input:focus, select:focus{ border-color:var(--accent); }
  input:disabled{ opacity:.65; }
  .btn{ display:inline-flex; align-items:center; justify-content:center; gap:8px; background:var(--accent); color:#161311; border:none; font-weight:600; padding:11px 18px; border-radius:var(--radius); cursor:pointer; font-size:14px; font-family:inherit; width:100%; }
  .btn:hover{ filter:brightness(1.08); }
  .btn:disabled{ opacity:.6; cursor:default; }
  .btn.secondary{ background:transparent; border:1px solid var(--border); color:var(--text); width:auto; }
  .btn.blue{ background:var(--blue); color:#fff; }
  .login-hint{ margin-top:18px; font-size:11px; color:var(--text-muted); line-height:1.6; border-top:1px dashed var(--border); padding-top:14px; }
  .error{ color:var(--danger); font-size:13px; margin:-8px 0 14px; }
  .menu-head{ margin-bottom:26px; }
  .menu-head h1{ font-size:22px; margin:0 0 6px; }
  .menu-head p{ color:var(--text-muted); font-size:13px; margin:0; }
  .menu-grid{ display:grid; grid-template-columns:repeat(auto-fit, minmax(230px,1fr)); gap:14px; }
  @media (max-width:560px){ .menu-grid{ grid-template-columns:1fr; } }
  .menu-card{ background:var(--panel); border:1px solid var(--border); border-inline-start:3px solid var(--accent); border-radius:var(--radius); padding:22px 20px; cursor:pointer; text-align:start; transition:background .15s; }
  .menu-card:nth-child(2){ border-inline-start-color:var(--blue); }
  .menu-card:nth-child(3){ border-inline-start-color:var(--ok); }
  .menu-card:nth-child(4){ border-inline-start-color:#a687c9; }
  .menu-card:nth-child(5){ border-inline-start-color:#6fb0c9; }
  .menu-card:hover{ background:var(--panel-alt); }
  .menu-card svg{ width:26px; height:26px; margin-bottom:14px; opacity:.9; }
  .menu-card h3{ font-size:15px; margin:0 0 6px; }
  .menu-card p{ font-size:12.5px; color:var(--text-muted); margin:0; line-height:1.5; }
  .page-head{ display:flex; align-items:center; justify-content:space-between; margin-bottom:20px; flex-wrap:wrap; gap:10px; }
  .page-head h2{ font-size:18px; margin:0; display:flex; align-items:center; gap:10px; }
  .back-link{ color:var(--text-muted); font-size:13px; cursor:pointer; display:flex; align-items:center; gap:6px; }
  .back-link:hover{ color:var(--text); }
  [dir="rtl"] .back-link svg{ transform:scaleX(-1); }
  .form-panel{ background:var(--panel); border:1px solid var(--border); border-radius:var(--radius); padding:24px; }
  .form-grid{ display:grid; grid-template-columns:1fr 1fr; gap:0 16px; }
  @media (max-width:600px){ .form-grid{ grid-template-columns:1fr; } }
  .form-actions{ display:flex; gap:10px; margin-top:6px; }
  .form-actions .btn{ width:auto; }
  .toast{ position:fixed; bottom:24px; left:50%; transform:translateX(-50%); background:var(--ok); color:#fff; padding:10px 18px; border-radius:var(--radius); font-size:13px; font-weight:600; opacity:0; transition:opacity .2s; pointer-events:none; max-width:90vw; text-align:center; }
  .toast.show{ opacity:1; }
  .toolbar{ display:flex; gap:10px; margin-bottom:16px; flex-wrap:wrap; }
  .table-wrap{ background:var(--panel); border:1px solid var(--border); border-radius:var(--radius); overflow-x:auto; }
  table{ width:100%; border-collapse:collapse; font-size:13px; min-width:920px; }
  th{ text-align:start; padding:10px 12px; color:var(--text-muted); font-weight:500; font-size:11px; text-transform:uppercase; letter-spacing:.4px; border-bottom:1px solid var(--border); background:var(--panel-alt); white-space:nowrap; }
  tr.filter-row th{ padding:6px 8px; text-transform:none; letter-spacing:0; background:var(--panel); }
  tr.filter-row input{ margin:0; padding:6px 8px; font-size:12px; font-weight:400; }
  td{ padding:9px 12px; border-bottom:1px solid var(--border); vertical-align:middle; }
  tr:last-child td{ border-bottom:none; }
  td input, td select{ margin:0; padding:6px 8px; font-size:13px; }
  .row-actions{ display:flex; gap:6px; }
  .icon-btn{ background:transparent; border:1px solid var(--border); color:var(--text-muted); width:28px; height:28px; border-radius:var(--radius); cursor:pointer; display:flex; align-items:center; justify-content:center; }
  .icon-btn:hover{ color:var(--accent); border-color:var(--accent); }
  .icon-btn.del:hover{ color:var(--danger); border-color:var(--danger); }
  .icon-btn svg{ width:14px; height:14px; }
  .empty-state{ padding:50px 20px; text-align:center; color:var(--text-muted); font-size:13px; }
  .stat-grid{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:20px; }
  @media (max-width:720px){ .stat-grid{ grid-template-columns:1fr 1fr; } }
  .stat-card{ background:var(--panel); border:1px solid var(--border); border-radius:var(--radius); padding:16px 18px; }
  .stat-card .label{ font-size:11px; color:var(--text-muted); text-transform:uppercase; letter-spacing:.4px; margin-bottom:8px; }
  .stat-card .value{ font-size:26px; font-weight:700; font-family:'Space Grotesk',sans-serif; }
  .stat-card .value.accent{ color:var(--accent); }
  .stat-card .value.blue{ color:var(--blue); }
  .charts-grid{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }
  @media (max-width:780px){ .charts-grid{ grid-template-columns:1fr; } }
  .chart-panel{ background:var(--panel); border:1px solid var(--border); border-radius:var(--radius); padding:18px; }
  .chart-panel.wide{ grid-column:1/-1; }
  .chart-panel h3{ font-size:13px; margin:0 0 14px; color:var(--text-muted); text-transform:uppercase; letter-spacing:.4px; font-weight:500; font-family:inherit; }
  .chart-panel canvas{ max-height:260px; }
  .banner{ background:var(--panel-alt); border:1px solid var(--border); border-radius:var(--radius); padding:14px 16px; font-size:12.5px; color:var(--text-muted); line-height:1.6; margin-bottom:16px; }
  .user-form{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:0 12px; align-items:end; }
  @media (max-width:700px){ .user-form{ grid-template-columns:1fr 1fr; } }
  .user-form .btn{ width:auto; margin-bottom:16px; padding:10px 16px; }
  .config-error{ display:flex; flex:1; align-items:center; justify-content:center; padding:30px; text-align:center; }
  .config-error .box{ max-width:480px; }
  .config-error h2{ margin:0 0 10px; }
  .config-error p{ color:var(--text-muted); font-size:13.5px; line-height:1.7; }
  .config-error code{ display:block; background:var(--panel); border:1px solid var(--border); padding:12px; border-radius:var(--radius); margin:14px 0; font-family:'IBM Plex Mono',monospace; font-size:12px; text-align:start; overflow-x:auto; }
  .tile-grid{ display:grid; grid-template-columns:repeat(auto-fill, minmax(130px,1fr)); gap:10px; }
  .tile{ position:relative; background:var(--panel); border:1px solid var(--border); border-radius:var(--radius); padding:16px 10px; text-align:center; cursor:pointer; font-size:13px; font-weight:600; transition:background .15s; word-break:break-word; }
  .tile:hover{ background:var(--panel-alt); border-color:var(--accent); }
  .tile .tile-icon{ font-size:22px; display:block; margin-bottom:8px; }
  .tile .tile-del{ position:absolute; top:4px; inset-inline-end:4px; background:transparent; border:none; color:var(--text-muted); width:22px; height:22px; border-radius:var(--radius); cursor:pointer; display:flex; align-items:center; justify-content:center; }
  .tile .tile-del:hover{ color:var(--danger); }
  .tile .tile-del svg{ width:12px; height:12px; }

/* ============================================================
   PREMIUM EXPERIENCE LAYER — visual refinement only
   Keeps the existing application logic, database calls and forms intact.
   ============================================================ */
:root{
  --bg:#f4f7fb;
  --panel:#ffffff;
  --panel-alt:#f6f8fc;
  --border:#e1e7f0;
  --text:#172337;
  --text-muted:#66758a;
  --accent:#0f9f8f;
  --blue:#315fd5;
  --danger:#d9475c;
  --ok:#15966a;
  --radius:14px;
  --premium-shadow:0 10px 28px rgba(25,43,74,.055);
}
html{scroll-behavior:smooth;}
body{
  background:
    radial-gradient(ellipse at 7% 0%, rgba(49,95,213,.075), transparent 32%),
    radial-gradient(ellipse at 100% 10%, rgba(15,159,143,.075), transparent 28%),
    var(--bg);
  color:var(--text);
  line-height:1.5;
}
button,input,select,textarea{transition:border-color .18s ease,box-shadow .18s ease,background-color .18s ease,transform .18s ease;}
.topbar{
  position:sticky;top:0;z-index:100;
  padding:13px clamp(14px,3vw,34px);
  background:rgba(255,255,255,.88);
  -webkit-backdrop-filter:blur(18px);backdrop-filter:blur(18px);
  border-bottom:1px solid rgba(218,226,238,.9);
  box-shadow:0 4px 20px rgba(28,45,75,.035);
}
.brand{font-size:18px;letter-spacing:-.35px;gap:12px;}
.brand .dot{width:12px;height:12px;border-radius:4px;transform:rotate(45deg);box-shadow:0 0 0 5px rgba(15,159,143,.11);}
.who{gap:12px;flex-wrap:wrap;}
main{max-width:1440px;padding:clamp(18px,3vw,34px) clamp(14px,3vw,34px) 64px;}
.login-card,.form-panel,.table-wrap,.stat-card,.chart-panel,.menu-card,.tile,.config-error .box{
  border-radius:var(--radius);
  box-shadow:var(--premium-shadow);
}
.login-card{padding:clamp(25px,4vw,38px);border-color:rgba(218,226,238,.95);}
.login-card .dot{border-radius:4px;box-shadow:0 0 0 6px rgba(15,159,143,.10);}
.login-card h1{letter-spacing:-.6px;font-size:24px;}
.menu-head{margin-bottom:28px;}
.menu-head h1{font-size:clamp(23px,3vw,30px);letter-spacing:-.8px;}
.menu-head p{font-size:14px;max-width:760px;line-height:1.75;}
.menu-grid{gap:16px;}
.menu-card{
  position:relative;overflow:hidden;min-height:142px;padding:24px;
  border:1px solid rgba(220,228,239,.95);border-inline-start:4px solid var(--accent);
  background:linear-gradient(145deg,#fff 0%,#fbfcff 100%);
  transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease;
}
.menu-card:after{content:"";position:absolute;width:120px;height:120px;inset-inline-end:-48px;bottom:-64px;border-radius:50%;background:rgba(15,159,143,.055);pointer-events:none;}
.menu-card:hover{transform:translateY(-3px);background:#fff;box-shadow:0 16px 34px rgba(25,43,74,.10);border-color:#cbd8e8;}
.menu-card svg{width:29px;height:29px;margin-bottom:15px;color:var(--accent);}
.menu-card h3{font-size:16px;letter-spacing:-.2px;}
.menu-card p{font-size:13px;line-height:1.65;}
.page-head{margin-bottom:24px;}
.page-head h2{font-size:clamp(19px,2.4vw,24px);letter-spacing:-.45px;}
.back-link{padding:8px 11px;border-radius:10px;transition:background .15s ease;}
.back-link:hover{background:rgba(49,95,213,.07);color:var(--blue);}
.form-panel{padding:clamp(16px,2.5vw,28px);border-color:rgba(220,228,239,.95);}
label{font-size:12px;font-weight:600;letter-spacing:.1px;color:#52637a;}
input,select,textarea{
  background:#f8fafd;border-color:#dce4ef;border-radius:10px;
  padding:11px 13px;color:var(--text);font-size:14px;
}
input:hover,select:hover,textarea:hover{border-color:#b9c9dd;}
input:focus,select:focus,textarea:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(49,95,213,.12);background:#fff;}
.btn{min-height:42px;border-radius:10px;padding:11px 18px;font-weight:700;letter-spacing:.05px;box-shadow:0 3px 8px rgba(15,159,143,.12);}
.btn:hover{filter:none;transform:translateY(-1px);box-shadow:0 7px 16px rgba(15,159,143,.18);}
.btn:active{transform:translateY(0);}
.btn.secondary{box-shadow:none;background:#fff;border-color:#dce4ef;}
.btn.secondary:hover{background:#f3f7fc;border-color:#b9c9dd;}
.btn.blue{background:linear-gradient(135deg,#315fd5,#244db9);box-shadow:0 4px 10px rgba(49,95,213,.16);}
.btn.blue:hover{box-shadow:0 8px 18px rgba(49,95,213,.22);}
.who button,.lang-btn,.icon-btn{border-radius:9px;background:#fff;}
.who button:hover,.lang-btn:hover{background:#f6f8fc;}
.toolbar{align-items:center;gap:9px;}
.table-wrap{border-color:#dfe6f0;}
table{font-size:13px;}
th{background:#f3f6fb;color:#52637a;font-weight:700;font-size:10.5px;letter-spacing:.55px;padding:12px;}
td{padding:11px 12px;border-bottom-color:#edf1f6;}
tbody tr{transition:background .14s ease;}
tbody tr:hover{background:#f8faff;}
tr.filter-row th{background:#fff;}
tr.filter-row input{border-radius:7px;}
td input,td select{border-radius:7px;}
.icon-btn{width:32px;height:32px;}
.icon-btn:hover{background:rgba(15,159,143,.07);}
.stat-grid{gap:14px;}
.stat-card{padding:20px;border-color:#e0e7f1;position:relative;overflow:hidden;}
.stat-card:after{content:"";position:absolute;inset-inline-end:-25px;top:-35px;width:95px;height:95px;border-radius:50%;background:rgba(49,95,213,.045);pointer-events:none;}
.stat-card .label{font-size:10.5px;font-weight:700;letter-spacing:.7px;}
.stat-card .value{font-size:clamp(24px,3vw,32px);letter-spacing:-.8px;}
.chart-panel{padding:22px;border-color:#e0e7f1;}
.chart-panel h3{font-size:12px;font-weight:700;}
.banner{border-radius:11px;background:linear-gradient(135deg,#f0f6ff,#f1fbf8);border-color:#dce8f3;}
.tile-grid{gap:12px;}
.tile{padding:18px 12px;border-color:#dfe6f0;background:linear-gradient(145deg,#fff,#fbfcff);transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease;}
.tile:hover{transform:translateY(-2px);background:#fff;box-shadow:0 10px 22px rgba(25,43,74,.08);border-color:#b7c9e2;}
.tile .tile-icon{font-size:25px;}
.toast{z-index:9999;bottom:22px;border-radius:12px;box-shadow:0 12px 30px rgba(16,33,57,.2);padding:12px 20px;}
.error{font-weight:600;}
.config-error code{border-radius:10px;}
:focus-visible{outline:3px solid rgba(49,95,213,.28);outline-offset:2px;}
@media(max-width:720px){
  .topbar{padding:12px 14px;align-items:flex-start;gap:10px;}
  .brand{font-size:15px;}
  .who{justify-content:flex-end;gap:6px;font-size:12px;}
  main{padding-top:20px;}
  .form-panel{padding:16px;}
  .stat-grid{gap:10px;}
  .stat-card{padding:15px;}
  .form-actions{flex-wrap:wrap;}
  .form-actions .btn{flex:1 1 auto;}
}
@media(max-width:420px){
  .topbar{flex-direction:column;}
  .who{width:100%;justify-content:space-between;}
  .menu-card{padding:20px;}
  .stat-card .value{font-size:23px;}
}
@media(prefers-reduced-motion:reduce){
  *,*::before,*::after{scroll-behavior:auto!important;transition:none!important;animation:none!important;}
}


/* ===== INDUSTRIAL ANALYTICS DASHBOARD (separate page; existing screens untouched) ===== */
.analytics-entry{border-inline-start-color:#3b82f6!important;background:linear-gradient(135deg,#111827,#172554)!important;color:#f8fafc!important}
.analytics-entry p{color:#aab7cf!important}
.analytics-entry svg{color:#60a5fa}
.analytics-shell{background:#080d18;color:#e8eef9;border:1px solid #263247;border-radius:18px;padding:22px;max-width:1500px;margin:0 auto;box-shadow:0 18px 55px #0b122022}
.analytics-shell *{box-sizing:border-box}
.analytics-shell .an-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.analytics-shell .an-eyebrow{font-size:11px;text-transform:uppercase;letter-spacing:1.5px;color:#6ea8ff;font-weight:700}
.analytics-shell h2{color:#f8fafc;font-size:24px;margin:5px 0 4px}
.analytics-shell .an-sub{color:#9aa9c1;font-size:13px;margin:0}
.analytics-shell .an-back{background:#121b2c;border:1px solid #2a3850;color:#dbe7fa;border-radius:10px;padding:9px 13px;cursor:pointer;font:inherit}
.analytics-shell .an-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:14px}
.analytics-shell .an-card,.analytics-shell .an-panel{background:linear-gradient(145deg,#101827,#0c1320);border:1px solid #202c40;border-radius:15px;padding:16px;min-width:0}
.analytics-shell .an-label{font-size:12px;color:#9aa9c1;margin-bottom:13px;display:flex;gap:8px;align-items:center}
.analytics-shell .an-value{font-size:clamp(22px,3vw,31px);font-weight:700;letter-spacing:-.8px;color:#f4f7ff;overflow-wrap:anywhere}
.analytics-shell .an-note{font-size:11px;color:#8293ad;margin-top:9px}
.analytics-shell .an-blue{color:#60a5fa}.analytics-shell .an-cyan{color:#4dd9e8}.analytics-shell .an-purple{color:#a78bfa}.analytics-shell .an-green{color:#34d399}
.analytics-shell .an-layout{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(260px,.9fr);gap:14px;margin-bottom:14px}
.analytics-shell .an-panel h3{font-size:14px;color:#f1f5fc;margin:0 0 5px;font-weight:600}
.analytics-shell .an-panel .an-desc{font-size:11px;color:#8293ad;margin:0 0 15px}
.analytics-shell .an-bars{display:flex;flex-direction:column;gap:14px}
.analytics-shell .an-bar-head{display:flex;justify-content:space-between;gap:10px;font-size:12px;color:#dce5f5;margin-bottom:7px}
.analytics-shell .an-track{height:7px;border-radius:20px;background:#202b3e;overflow:hidden}
.analytics-shell .an-fill{height:100%;border-radius:20px;background:linear-gradient(90deg,#3b82f6,#22d3ee)}
.analytics-shell .an-bottom{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}
.analytics-shell .an-mini{border:1px solid #202c40;background:#0e1726;border-radius:12px;padding:14px}
.analytics-shell .an-mini strong{font-size:22px;color:#f8fafc;display:block;margin:9px 0 5px}
.analytics-shell .an-mini span{font-size:11px;color:#8ea0bb}
.analytics-shell .an-foot{color:#71829d;font-size:10px;margin-top:15px}
@media(max-width:900px){.analytics-shell .an-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.analytics-shell .an-layout{grid-template-columns:1fr}.analytics-shell .an-bottom{grid-template-columns:1fr 1fr}}
@media(max-width:520px){.analytics-shell{padding:13px;border-radius:12px}.analytics-shell .an-grid{gap:8px}.analytics-shell .an-card,.analytics-shell .an-panel{padding:12px}.analytics-shell .an-bottom{grid-template-columns:1fr}.analytics-shell h2{font-size:20px}}


/* ===== APP WORKSPACE LAYOUT: sidebar + central workspace ===== */
body:has(.workspace-shell){background:#070c16;color:#e7edf8;}
.workspace-shell{min-height:100vh;display:grid;grid-template-columns:238px minmax(0,1fr);background:#080d18;color:#e8eef9;}
.workspace-sidebar{background:#101722;border-right:1px solid #202b3c;display:flex;flex-direction:column;min-height:100vh;padding:18px 12px 14px;}
.workspace-brand{display:flex;align-items:center;gap:10px;padding:3px 8px 22px;border-bottom:1px solid #202b3c;margin-bottom:17px;color:#f8fafc;font-weight:700;font-size:17px;}
.workspace-brand-mark{width:34px;height:34px;display:grid;place-items:center;border-radius:10px;background:#d8edff;color:#2474c8;font-size:18px;}
.workspace-brand small{display:block;color:#8e9cb1;font-size:11px;font-weight:400;margin-top:2px;}
.workspace-nav{display:flex;flex-direction:column;gap:5px;}
.workspace-nav button,.workspace-logout{width:100%;border:0;background:transparent;color:#9ba9bd;display:flex;align-items:center;gap:11px;text-align:start;padding:11px 12px;border-radius:11px;cursor:pointer;font:500 13px 'IBM Plex Sans',sans-serif;}
.workspace-nav button:hover,.workspace-logout:hover{background:#192438;color:#f8fafc;}
.workspace-nav button.active{background:#1c3b6c;color:#fff;box-shadow:inset 3px 0 #438bff;}
.workspace-nav .nav-symbol{width:18px;text-align:center;font-size:16px;flex-shrink:0;}
.workspace-logout-wrap{margin-top:auto;padding-top:16px;border-top:1px solid #202b3c;}
.workspace-logout{color:#9ba9bd;}
.workspace-logout .nav-symbol{color:#8fa3bd;}
.workspace-main{min-width:0;display:flex;flex-direction:column;min-height:100vh;background:#080d18;}
.workspace-topbar{height:72px;display:flex;align-items:center;justify-content:space-between;gap:14px;padding:12px 26px;border-bottom:1px solid #202b3c;background:#090e18;}
.workspace-search{max-width:390px;width:100%;display:flex;align-items:center;gap:9px;border:1px solid #202b3c;background:#111824;border-radius:22px;padding:9px 13px;color:#8190a6;font-size:13px;}
.workspace-user{display:flex;align-items:center;gap:10px;color:#e7edf8;font-size:13px;}
.workspace-avatar{width:37px;height:37px;border-radius:50%;display:grid;place-items:center;background:#d8edff;color:#245a9b;font-weight:700;border:2px solid #2b3a51;}
.workspace-user-name{max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.workspace-content{width:100%;max-width:none;margin:0;padding:25px 26px 36px;flex:1;min-width:0;}
.workspace-content .analytics-shell{max-width:none;}
.workspace-content .page-head,.workspace-content .menu-head{color:#f0f5ff;}
.workspace-content .menu-head p{color:#9ba9bd;}
.workspace-content .menu-card,.workspace-content .stat-card,.workspace-content .chart-panel,.workspace-content .form-panel,.workspace-content .table-wrap{background:#101722;color:#e8eef9;border-color:#263247;}
.workspace-content .menu-card p,.workspace-content .label,.workspace-content label,.workspace-content .back-link{color:#9ba9bd;}
.workspace-content input,.workspace-content select{background:#0c1320;color:#e8eef9;border-color:#2a3850;}
.workspace-content th{background:#141e2d;color:#aab8cd;border-color:#263247;}
.workspace-content td{border-color:#263247;}
.workspace-content .menu-card:hover{background:#172235;}
@media(max-width:850px){.workspace-shell{grid-template-columns:190px minmax(0,1fr)}.workspace-content{padding:18px 15px}.workspace-topbar{padding:10px 15px}}
@media(max-width:620px){.workspace-shell{grid-template-columns:1fr}.workspace-sidebar{min-height:auto;padding:10px;border-right:0;border-bottom:1px solid #202b3c}.workspace-brand{padding:4px 6px 12px;margin-bottom:10px}.workspace-nav{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:4px}.workspace-nav button{padding:9px 8px;font-size:12px}.workspace-logout-wrap{margin-top:10px;padding-top:8px}.workspace-topbar{height:auto;min-height:62px}.workspace-search{max-width:180px}.workspace-user-name{max-width:100px}.workspace-content{padding:15px 10px}.workspace-content .analytics-shell{padding:14px}.analytics-shell .an-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.analytics-shell .an-layout{grid-template-columns:1fr}}

</style>

</head>
<body>
<div id="app"></div>
<div class="toast" id="toast"></div>

<script>
/* ============================================================
   CONFIGURATION SUPABASE
   ============================================================ */
const SUPABASE_URL = 'https://ykbfeywmuemuwdiqcjow.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlrYmZleXdtdWVtdXdkaXFjam93Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTEyOTcwNjQsImV4cCI6MjEwNjg3MzA2NH0.MtuHMsJepuEcSRZrW5lOtSLAw8smosTfScw6HG9zeaw';

/* ================= I18N ================= */
const I18N = {
  appTitle:{fr:'Gestion de Stock', ar:'إدارة المخزون'},
  loginSubtitle:{fr:'Connectez-vous pour accéder au système', ar:'سجّل الدخول للوصول إلى النظام'},
  loginUser:{fr:'Identifiant', ar:'اسم المستخدم'},
  loginPass:{fr:'Mot de passe', ar:'كلمة المرور'},
  loginBtn:{fr:'Se connecter', ar:'تسجيل الدخول'},
  loginBtnLoading:{fr:'Connexion...', ar:'جارٍ الدخول...'},
  loginError:{fr:'Identifiant ou mot de passe incorrect.', ar:'اسم المستخدم أو كلمة المرور غير صحيحة.'},
  logout:{fr:'Déconnexion', ar:'تسجيل الخروج'},
  menuTitle:{fr:'Menu principal', ar:'القائمة الرئيسية'},
  menuSubtitle:{fr:'Choisissez une opération', ar:'اختر عملية'},
  moduleStockTitle:{fr:'Gestion de Stock', ar:'إدارة المخزون'},
  moduleStockDesc:{fr:'Saisie, données et tableau de bord du stock.', ar:'إدخال، بيانات ولوحة تحكم المخزون.'},
  moduleQcTitle:{fr:'Contrôle Qualité', ar:'مراقبة الجودة'},
  moduleQcDesc:{fr:'Saisie, données et tableau de bord du contrôle qualité.', ar:'إدخال، بيانات ولوحة تحكم مراقبة الجودة.'},
  cardEntryTitle:{fr:'Saisie des données', ar:'إدخال البيانات'},
  cardEntryDesc:{fr:'Enregistrer une nouvelle pièce.', ar:'تسجيل قطعة جديدة.'},
  cardDataTitle:{fr:'Données enregistrées', ar:'البيانات المسجلة'},
  cardDataDesc:{fr:'Consulter, modifier ou supprimer les enregistrements.', ar:'عرض أو تعديل أو حذف السجلات.'},
  cardDashTitle:{fr:'Tableau de bord', ar:'لوحة التحكم'},
  cardDashDesc:{fr:'Statistiques et répartition des données.', ar:'إحصائيات وتوزيع البيانات.'},
  cardUsersTitle:{fr:'Utilisateurs', ar:'المستخدمون'},
  cardUsersDesc:{fr:'Gérer les rôles des comptes existants.', ar:'إدارة صلاحيات الحسابات الموجودة.'},
  cardRecipientsTitle:{fr:'Destinataires du rapport', ar:'مستلمو التقرير'},
  cardRecipientsDesc:{fr:'Gérer la liste des e-mails qui reçoivent le rapport automatique.', ar:'إدارة قائمة البريد الإلكتروني الذي يستلم التقرير التلقائي.'},
  cardPdfTitle:{fr:'Bibliothèque PDF', ar:'مكتبة PDF'},
  cardPdfDesc:{fr:'Stocker, consulter et télécharger des fichiers PDF.', ar:'تخزين وعرض وتحميل ملفات PDF.'},
  pdfLibraryTitle:{fr:'Bibliothèque PDF', ar:'مكتبة PDF'},
  cardBqTitle:{fr:'Bonne qualité', ar:'Bonne qualité'},
  cardBqDesc:{fr:'Certificats et documents d\'approbation qualité.', ar:'شهادات ووثائق موافقة الجودة.'},
  bqLibraryTitle:{fr:'Bonne qualité', ar:'Bonne qualité'},
  cardDebitTitle:{fr:'Note de Débit', ar:'مذكرة خصم'},
  cardDebitDesc:{fr:'Créer et consulter les notes de débit fournisseur.', ar:'إنشاء ومراجعة مذكرات الخصم للموردين.'},
  sendTestReportBtn:{fr:'Envoyer un test maintenant', ar:'إرسال اختبار الآن'},
  debitSaveBtn:{fr:'Enregistrer (N° + Historique)', ar:'حفظ (رقم + السجل)'},
  debitPdfBtn:{fr:'Exporter PDF', ar:'تصدير PDF'},
  debitSavedToast:{fr:'Note enregistrée', ar:'تم حفظ المذكرة'},
  debitSaveFirst:{fr:'Enregistrez d’abord la feuille pour obtenir son numéro', ar:'احفظ الورقة أولاً للحصول على رقمها'},
  debitNewTitle:{fr:'Nouvelle note de débit', ar:'مذكرة خصم جديدة'},
  debitNewDesc:{fr:'Créer une nouvelle note avec calcul automatique et export PDF.', ar:'إنشاء مذكرة جديدة مع حساب تلقائي وتصدير PDF.'},
  debitListTitle:{fr:'Historique', ar:'السجل'},
  debitListDesc:{fr:'Consulter toutes les notes de débit créées.', ar:'مراجعة كل مذكرات الخصم المُنشأة.'},
  debitNoteNumber:{fr:'Débit note N°', ar:'رقم المذكرة'},
  debitAswt:{fr:'ASWT Department', ar:'ASWT Department'},
  debitSupplierInfo:{fr:'Supplier Contact Information', ar:'معلومات الاتصال بالمورد'},
  debitSupplierName:{fr:'Name', ar:'الاسم'},
  debitSupplierMail:{fr:'Mail', ar:'البريد'},
  debitSupplierPhone:{fr:'Phone', ar:'الهاتف'},
  debitDepartment:{fr:'Department', ar:'القسم'},
  debitPartNumber:{fr:'Part number', ar:'رقم القطعة'},
  debitPartName:{fr:'Part Name', ar:'اسم القطعة'},
  debitCause:{fr:'Cause', ar:'السبب'},
  debitReportBy:{fr:'Report raised by', ar:'أُعدّ بواسطة'},
  debitProductionCost:{fr:'PRODUCTION COST', ar:'تكلفة الإنتاج'},
  debitLogisticCost:{fr:'LOGISTIC COST', ar:'تكلفة اللوجستيك'},
  debitQualityCost:{fr:'QUALITY COST', ar:'تكلفة الجودة'},
  debitAdminCost:{fr:'ADMINISTRATION COST', ar:'التكلفة الإدارية'},
  debitColDate:{fr:'Date', ar:'التاريخ'},
  debitColDesc:{fr:'Désignation', ar:'الوصف'},
  debitColQty:{fr:'Quantité', ar:'الكمية'},
  debitColPrice:{fr:'Prix unit.', ar:'سعر الوحدة'},
  debitColTotal:{fr:'Total', ar:'الإجمالي'},
  debitAddRow:{fr:'+ Ajouter une ligne', ar:'+ إضافة سطر'},
  debitSectionTotal:{fr:'Total', ar:'الإجمالي'},
  debitGrandTotal:{fr:'TOTAL GÉNÉRAL', ar:'الإجمالي الكلي'},
  debitExportBtn:{fr:'Exporter en PDF', ar:'تصدير PDF'},
  debitExportedToast:{fr:'PDF généré et téléchargé sur l\'appareil', ar:'تم إنشاء PDF وتنزيله على الجهاز'},
  debitExportErr:{fr:'Échec de la génération du PDF', ar:'فشل إنشاء PDF'},
  debitRequiredErr:{fr:'Fournisseur et cause obligatoires', ar:'المورد والسبب إلزاميان'},
  debitNoNotes:{fr:'Aucune note de débit', ar:'لا توجد مذكرات خصم'},
  debitConfirmDelete:{fr:'Supprimer cette note ?', ar:'حذف هذه المذكرة؟'},
  debitColSite:{fr:'Lieu essai/atelier', ar:'مكان الاختبار/الورشة'},
  debitColCarrier:{fr:'Transporteur', ar:'الناقل'},
  debitColPerson:{fr:'Personne', ar:'الشخص'},
  debitColHeures:{fr:'Heures', ar:'الساعات'},
  debitColCoeff:{fr:'Coeff.', ar:'المعامل'},
  debitColCost:{fr:'Coût', ar:'التكلفة'},
  debitColUnitCost:{fr:'Coût unit.', ar:'تكلفة الوحدة'},
  debitNcm:{fr:'NCM', ar:'NCM'},
  debitCart:{fr:'Cart', ar:'Cart'},
  debitInvoiceNumber:{fr:'Invoice N°', ar:'رقم الفاتورة'},
  debitApprovalTitle:{fr:'APPROVAL', ar:'الموافقات'},
  debitColRole:{fr:'Approval', ar:'الجهة'},
  debitColWhen:{fr:'When (Date)', ar:'التاريخ'},
  debitColComment:{fr:'Comments if refused', ar:'ملاحظات في حال الرفض'},
  debitRolePlant:{fr:'Plant Manager / Quality Manager', ar:'مدير المصنع / مدير الجودة'},
  debitRoleLogistic:{fr:'Logistic Manager', ar:'مدير اللوجستيك'},
  debitRoleBuyer:{fr:'Commodity Buyer', ar:'مسؤول المشتريات'},
  debitRoleFinance:{fr:'Finances Controller', ar:'المراقب المالي'},
  debitRoleGeneral:{fr:'General Manager (si Total > 20 000)', ar:'المدير العام (إذا الإجمالي > 20000)'},
  debitPdfTitle:{fr:'PDF', ar:'PDF'},
  debitPdfDesc:{fr:'Ajouter et consulter manuellement des fichiers PDF.', ar:'إدخال ومراجعة ملفات PDF يدويًا.'},
  debitStatusCol:{fr:'Statut', ar:'الحالة'},
  debitStatusDone:{fr:'Fait', ar:'تم'},
  debitStatusNotDone:{fr:'Pas fait', ar:'لم يتم'},
  uploadPdfBtn:{fr:'Téléverser un PDF', ar:'رفع ملف PDF'},
  noPdf:{fr:'Aucun fichier PDF pour le moment', ar:'لا توجد ملفات PDF حالياً'},
  confirmDeletePdf:{fr:'Supprimer ce fichier ?', ar:'حذف هذا الملف؟'},
  pdfUploadedToast:{fr:'Fichier téléversé', ar:'تم رفع الملف'},
  pdfDeletedToast:{fr:'Fichier supprimé', ar:'تم حذف الملف'},
  pdfUploadErr:{fr:'Échec du téléversement (le fichier doit être un PDF)', ar:'فشل الرفع (يجب أن يكون الملف بصيغة PDF)'},
  pdfSelectErr:{fr:'Veuillez sélectionner un fichier PDF', ar:'يرجى اختيار ملف PDF'},
  viewBtn:{fr:'Ouvrir', ar:'فتح'},
  recipientsTitle:{fr:'Destinataires du rapport', ar:'مستلمو التقرير'},
  recipientsBanner:{fr:'Ces adresses e-mail recevront automatiquement le rapport Contrôle Qualité, 2 fois par jour (13:10 et 21:10).', ar:'هذه العناوين تستلم تلقائياً تقرير مراقبة الجودة، مرتين يومياً (13:10 و 21:10).'},
  addRecipientBtn:{fr:'Ajouter', ar:'إضافة'},
  newRecipientEmail:{fr:'Adresse e-mail', ar:'البريد الإلكتروني'},
  recipientExistsErr:{fr:'Cet e-mail est déjà dans la liste', ar:'هذا البريد موجود مسبقاً بالقائمة'},
  recipientFieldErr:{fr:'Veuillez saisir un e-mail valide', ar:'يرجى إدخال بريد إلكتروني صحيح'},
  recipientAddedToast:{fr:'Destinataire ajouté', ar:'تمت إضافة المستلم'},
  recipientDeletedToast:{fr:'Destinataire supprimé', ar:'تم حذف المستلم'},
  confirmDeleteRecipient:{fr:'Supprimer ce destinataire ?', ar:'حذف هذا المستلم؟'},
  noRecipients:{fr:'Aucun destinataire — le rapport ne sera envoyé à personne', ar:'لا يوجد مستلمون — لن يُرسل التقرير لأحد'},
  cardCatalogTitle:{fr:'Catalogue des références', ar:'كتالوج المراجع'},
  cardCatalogDesc:{fr:'Base Ref → Désignation → Fournisseur, pour le remplissage automatique.', ar:'قاعدة المرجع ← التسمية ← المورد، للتعبئة التلقائية.'},
  cardQcEntryTitle:{fr:'Saisie contrôle', ar:'إدخال المراقبة'},
  cardQcEntryDesc:{fr:'Enregistrer un contrôle : quantité contrôlée, non OK, et types de défauts.', ar:'تسجيل عملية مراقبة: الكمية المراقبة، غير المطابقة، وأنواع العيوب.'},
  cardQcDataTitle:{fr:'Données contrôle', ar:'بيانات المراقبة'},
  cardQcDataDesc:{fr:'Consulter, modifier ou supprimer les contrôles enregistrés.', ar:'عرض أو تعديل أو حذف عمليات المراقبة.'},
  cardQcDashTitle:{fr:'Tableau de bord contrôle', ar:'لوحة تحكم المراقبة'},
  cardQcDashDesc:{fr:'Taux de non-conformité et répartition des défauts.', ar:'نسبة عدم المطابقة وتوزيع العيوب.'},
  back:{fr:'Menu', ar:'القائمة'},
  fieldLot:{fr:'Lot', ar:'الدفعة (Lot)'},
  fieldRef:{fr:'Référence (Ref)', ar:'المرجع (Ref)'},
  fieldDesig:{fr:'Désignation', ar:'التسمية'},
  fieldDefaut:{fr:'Défaut', ar:'العيب'},
  fieldQte:{fr:'Quantité', ar:'الكمية'},
  fieldLoc:{fr:'Emplacement (Location)', ar:'الموقع'},
  fieldFourn:{fr:'Fournisseur', ar:'المورد'},
  fieldPrix:{fr:'Prix', ar:'السعر'},
  fieldUser:{fr:'Utilisateur', ar:'المستخدم'},
  fieldDate:{fr:'Date', ar:'التاريخ'},
  fieldQteControle:{fr:'Qtité contrôle', ar:'الكمية المراقَبة'},
  fieldQteNonOk:{fr:'Quantité non OK', ar:'الكمية غير المطابقة'},
  fieldPourcentage:{fr:'Pourcentage', ar:'النسبة المئوية'},
  fieldGrani:{fr:'Grani', ar:'Grani'},
  fieldRayure:{fr:'Rayure', ar:'خدش (Rayure)'},
  fieldTrace:{fr:'Trace', ar:'أثر (Trace)'},
  fieldPiqure:{fr:'Piqûre', ar:'ثقب (Piqûre)'},
  fieldCoup:{fr:'Coup', ar:'صدمة (Coup)'},
  fieldCosse:{fr:'Cosse', ar:'Cosse'},
  fieldAutre:{fr:'Autre', ar:'أخرى (Autre)'},
  saveBtn:{fr:'Enregistrer', ar:'حفظ'},
  clearBtn:{fr:'Effacer', ar:'مسح'},
  requiredError:{fr:'Référence et désignation obligatoires', ar:'المرجع والتسمية إلزاميان'},
  qcRequiredError:{fr:'Référence et quantité contrôlée obligatoires', ar:'المرجع والكمية المراقبة إلزاميان'},
  savedToast:{fr:'Enregistrement sauvegardé', ar:'تم الحفظ بنجاح'},
  qcSavedToast:{fr:'Contrôle enregistré', ar:'تم حفظ عملية المراقبة'},
  exportBtn:{fr:'Exporter en Excel', ar:'تصدير إلى Excel'},
  noRecords:{fr:'Aucun enregistrement', ar:'لا توجد بيانات'},
  qcNoRecords:{fr:'Aucun contrôle enregistré', ar:'لا توجد عمليات مراقبة'},
  noRecordsSearch:{fr:' pour cette recherche', ar:' لهذا البحث'},
  confirmDelete:{fr:'Supprimer cet enregistrement ?', ar:'هل تريد حذف هذا السجل؟'},
  qcConfirmDelete:{fr:'Supprimer ce contrôle ?', ar:'هل تريد حذف عملية المراقبة هذه؟'},
  deletedToast:{fr:'Enregistrement supprimé', ar:'تم الحذف'},
  editedToast:{fr:'Modifications enregistrées', ar:'تم حفظ التعديلات'},
  statRecords:{fr:'Enregistrements', ar:'السجلات'},
  statQty:{fr:'Quantité totale', ar:'الكمية الإجمالية'},
  statFourn:{fr:'Fournisseurs', ar:'الموردون'},
  statLoc:{fr:'Emplacements', ar:'المواقع'},
  noData:{fr:'Aucune donnée à afficher pour le moment.', ar:'لا توجد بيانات لعرضها حالياً.'},
  chartFournTitle:{fr:'Quantité par fournisseur', ar:'الكمية حسب المورد'},
  chartDefautTitle:{fr:'Répartition par défaut', ar:'التوزيع حسب العيب'},
  chartLocTitle:{fr:'Quantité par emplacement', ar:'الكمية حسب الموقع'},
  chartUserTitle:{fr:'Enregistrements par utilisateur', ar:'السجلات حسب المستخدم'},
  chartRefFournTitle:{fr:'Quantité par référence, pour chaque fournisseur', ar:'الكمية حسب المرجع، لكل مورد'},
  noExportData:{fr:'Aucune donnée à exporter', ar:'لا توجد بيانات للتصدير'},
  exportedToast:{fr:'Fichier Excel téléchargé', ar:'تم تحميل ملف Excel'},
  undefinedLabel:{fr:'Non défini', ar:'غير محدد'},
  usersTitle:{fr:'Gestion des utilisateurs', ar:'إدارة المستخدمين'},
  usersBanner:{fr:'Pour créer un nouveau compte, utilisez le tableau de bord Supabase : Authentication → Users → Add user (email au format identifiant@stock.local). Il apparaîtra automatiquement ici — vous pourrez ensuite lui attribuer un rôle.', ar:'لإنشاء حساب جديد، استخدم لوحة تحكم Supabase: Authentication → Users → Add user (البريد بصيغة username@stock.local). سيظهر تلقائياً هنا وتقدر تحدد صلاحيته.'},
  addUserBtn:{fr:'Ajouter un utilisateur', ar:'إضافة مستخدم'},
  newUsername:{fr:'Identifiant', ar:'اسم المستخدم'},
  newName:{fr:'Nom complet', ar:'الاسم الكامل'},
  newPass:{fr:'Mot de passe', ar:'كلمة المرور'},
  userFieldsErr:{fr:'Veuillez remplir tous les champs', ar:'يرجى تعبئة جميع الحقول'},
  userPassShortErr:{fr:'Le mot de passe doit contenir au moins 9 caractères', ar:'يجب أن تحتوي كلمة المرور على 9 أحرف على الأقل'},
  userExistsErr:{fr:'Cet identifiant existe déjà', ar:'اسم المستخدم موجود مسبقاً'},
  userAddedToast:{fr:'Utilisateur ajouté', ar:'تمت إضافة المستخدم'},
  userAddErrGeneric:{fr:"Échec de la création du compte", ar:'فشل إنشاء الحساب'},
  colUsername:{fr:'Identifiant', ar:'اسم المستخدم'},
  colName:{fr:'Nom', ar:'الاسم'},
  colRole:{fr:'Rôle', ar:'الصلاحية'},
  roleAdmin:{fr:'Administrateur', ar:'مسؤول'},
  roleUser:{fr:'Opérateur', ar:'مشغل'},
  roleEmployee:{fr:'Employé(e)', ar:'موظف'},
  roleVisitor:{fr:'Visiteur', ar:'زائر'},
  roleEntrepreneur:{fr:'Entrepreneur', ar:'مقاول'},
  cardOpTitle:{fr:'Portail opérateur', ar:'بوابة المشغّل'},
  cardOpDesc:{fr:'8 tables de contrôle : saisie horaire, rejets et commentaires', ar:'8 طاولات مراقبة: إدخال كل ساعة والمرفوضات والتعليقات'},
  roleUpdatedToast:{fr:'Rôle mis à jour', ar:'تم تحديث الصلاحية'},
  connError:{fr:"Erreur de connexion au serveur. Vérifiez votre connexion internet.", ar:'خطأ في الاتصال بالخادم. تحقق من اتصالك بالإنترنت.'},
  configErrorTitle:{fr:'Configuration requise', ar:'الإعداد مطلوب'},
  configErrorBody:{fr:"Ce fichier n'est pas encore connecté à une base de données. Ouvrez le fichier dans un éditeur de texte et remplacez SUPABASE_URL et SUPABASE_ANON_KEY en haut du fichier.", ar:'هذا الملف غير متصل بعد بقاعدة بيانات. افتح الملف بمحرر نصوص واستبدل SUPABASE_URL وSUPABASE_ANON_KEY أعلى الملف.'},
  catalogBanner:{fr:'Ajoutez vos références ici : lors de la saisie, taper une référence connue remplira automatiquement la désignation et le fournisseur.', ar:'أضف مراجعك هنا: عند كتابة مرجع معروف، سيتم تعبئة التسمية والمورد تلقائياً.'},
  addCatalogBtn:{fr:'Ajouter au catalogue', ar:'إضافة للكتالوج'},
  catalogExistsErr:{fr:'Cette référence existe déjà dans le catalogue', ar:'هذا المرجع موجود مسبقاً في الكتالوج'},
  catalogFieldsErr:{fr:'Veuillez remplir tous les champs', ar:'يرجى تعبئة جميع الحقول'},
  catalogAddedToast:{fr:'Référence ajoutée au catalogue', ar:'تمت إضافة المرجع للكتالوج'},
  catalogDeletedToast:{fr:'Référence supprimée du catalogue', ar:'تم حذف المرجع من الكتالوج'},
  confirmDeleteCatalog:{fr:'Supprimer cette référence du catalogue ?', ar:'حذف هذا المرجع من الكتالوج؟'},
  noCatalog:{fr:'Aucune référence dans le catalogue', ar:'لا توجد مراجع في الكتالوج'},
  qcStatControls:{fr:'Contrôles', ar:'عمليات المراقبة'},
  qcStatControlled:{fr:'Quantité contrôlée', ar:'الكمية المراقَبة'},
  qcStatNonOk:{fr:'Quantité non OK', ar:'الكمية غير المطابقة'},
  qcStatRate:{fr:'Taux de non-conformité', ar:'نسبة عدم المطابقة'},
  qcChartDefectTitle:{fr:'Répartition des types de défauts', ar:'توزيع أنواع العيوب'},
  qcChartRateFournTitle:{fr:'Taux de non-conformité par fournisseur (%)', ar:'نسبة عدم المطابقة حسب المورد (%)'},
  qcChartRefTitle:{fr:'Quantité contrôlée vs non OK par référence', ar:'الكمية المراقَبة مقابل غير المطابقة حسب المرجع'},
};
function t(key){ return (I18N[key] && I18N[key][state.lang]) || key; }

/* ================= STATE ================= */
let state = {
  view: 'login', currentUser: null, records: [], profiles: [], catalog: [], qcRecords: [], recipients: [], pdfFiles: [], pdfFolder: '', currentLibrary: 'pdf', debitNotes: [], debitForm: null,
  editingId: null, editingQcId: null,
  search: { lot:'', ref:'', designation:'', defaut:'', qtite:'', location:'', fournisseur:'', user:'', date:'' },
  qcSearch: { date:'', fournisseur:'', ref:'', designation:'' },
  lastSearchFocus: null, lastQcSearchFocus: null, lang: 'fr',
};
const DATA_COLS = [
  {key:'lot', field:'lot', label:'fieldLot'}, {key:'ref', field:'ref', label:'fieldRef'},
  {key:'designation', field:'designation', label:'fieldDesig'}, {key:'defaut', field:'defaut', label:'fieldDefaut'},
  {key:'qtite', field:'qtite', label:'fieldQte'}, {key:'location', field:'location', label:'fieldLoc'},
  {key:'fournisseur', field:'fournisseur', label:'fieldFourn'}, {key:'user', field:'user', label:'fieldUser'},
  {key:'date', field:'date', label:'fieldDate'},
];
const DEFECT_COLS = [
  {key:'grani', label:'fieldGrani'}, {key:'rayure', label:'fieldRayure'}, {key:'trace', label:'fieldTrace'},
  {key:'piqure', label:'fieldPiqure'}, {key:'coup', label:'fieldCoup'}, {key:'cosse', label:'fieldCosse'}, {key:'autre', label:'fieldAutre'},
];
const QC_DISPLAY_COLS = [
  {key:'date', label:'fieldDate', filterable:true}, {key:'fournisseur', label:'fieldFourn', filterable:true},
  {key:'ref', label:'fieldRef', filterable:true}, {key:'designation', label:'fieldDesig', filterable:true},
  {key:'qteControle', label:'fieldQteControle', filterable:false}, {key:'qteNonOk', label:'fieldQteNonOk', filterable:false},
  {key:'pourcentage', label:'fieldPourcentage', filterable:false}, ...DEFECT_COLS, {key:'user', label:'fieldUser', filterable:false},
];
function qcPct(r){ return r.qteControle > 0 ? (r.qteNonOk / r.qteControle * 100) : 0; }
function calcPrixTotal(r){
  const cat = state.catalog.find(c => c.ref.toLowerCase() === (r.ref||'').toLowerCase());
  const unit = cat ? Number(cat.prix)||0 : 0;
  return r.qtite * unit;
}

/* ================= SUPABASE CLIENT ================= */
const configOk = SUPABASE_URL.startsWith('https://') && (SUPABASE_ANON_KEY.startsWith('eyJ') || SUPABASE_ANON_KEY.startsWith('sb_publishable_'));
const sb = configOk ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;
function usernameToEmail(u){ return u.trim().toLowerCase().replace(/\s+/g,'') + '@stock.local'; }
function uid(){ return 'c' + Date.now() + Math.floor(Math.random()*1000); }
function todayISO(){ return new Date().toISOString().slice(0,10); }

/* ================= DATA LAYER (Supabase) — STOCK ================= */
function mapRecord(row){
  return { id: row.id, lot: row.lot||'', ref: row.ref, designation: row.designation, defaut: row.defaut||'', qtite: row.qtite||0, location: row.location||'', fournisseur: row.fournisseur||'', user: row.user_name||'', date: row.record_date||'' };
}
async function loadRecords(){
  const { data, error } = await sb.from('stock_records').select('*').order('created_at', { ascending:false });
  if(error){ showToast(t('connError'), true); state.records = []; return; }
  state.records = data.map(mapRecord);
}
async function insertRecord(rec){
  const { error } = await sb.from('stock_records').insert({ lot: rec.lot, ref: rec.ref, designation: rec.designation, defaut: rec.defaut, qtite: rec.qtite, location: rec.location, fournisseur: rec.fournisseur, user_name: rec.user, record_date: rec.date });
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
async function updateRecord(r){
  const { error } = await sb.from('stock_records').update({ lot: r.lot, ref: r.ref, designation: r.designation, defaut: r.defaut, qtite: r.qtite, location: r.location, fournisseur: r.fournisseur, record_date: r.date }).eq('id', r.id);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
async function deleteRecord(id){
  const { error } = await sb.from('stock_records').delete().eq('id', id);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
async function loadProfiles(){
  const { data, error } = await sb.from('profiles').select('*').order('created_at', { ascending:true });
  state.profiles = error ? [] : data;
}
async function updateRole(id, role){
  const { error } = await sb.from('profiles').update({ role }).eq('id', id);
  return !error;
}
async function loadRecipients(){
  const { data, error } = await sb.from('report_recipients').select('*').order('created_at', { ascending:true });
  state.recipients = error ? [] : data;
}
async function insertRecipient(email){
  const { error } = await sb.from('report_recipients').insert({ email });
  if(error){ showToast(error.code === '23505' ? t('recipientExistsErr') : t('connError'), true); return false; }
  return true;
}
async function deleteRecipient(id){
  const { error } = await sb.from('report_recipients').delete().eq('id', id);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
const LIBRARY_BUCKETS = { pdf: 'pdf-library', bonnequalite: 'bonne-qualite' };
function currentBucket(){ return LIBRARY_BUCKETS[state.currentLibrary] || 'pdf-library'; }
function slugFolder(name){
  return name.trim().replace(/[\/\\]/g, '-').replace(/\s+/g, '_');
}
async function listPdfEntries(folder){
  const { data, error } = await sb.storage.from(currentBucket()).list(folder, { sortBy: { column:'name', order:'asc' } });
  if(error) showToast(`Stockage « ${currentBucket()} » : ${error.message}`, true);
  state.pdfFiles = error ? [] : (data || []).filter(f => f.name !== '.emptyFolderPlaceholder');
}
async function uploadPdfFile(folder, file){
  const path = `${folder}/${Date.now()}_${file.name}`;
  const { error } = await sb.storage.from(currentBucket()).upload(path, file, { contentType: 'application/pdf' });
  if(error){ showToast(t('pdfUploadErr'), true); return false; }
  return true;
}
async function getPdfSignedUrl(path){
  const { data, error } = await sb.storage.from(currentBucket()).createSignedUrl(path, 3600);
  return error ? null : data.signedUrl;
}
async function deletePdfFile(path){
  const { error } = await sb.storage.from(currentBucket()).remove([path]);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
async function loadDebitNotes(){
  const { data, error } = await sb.from('debit_notes').select('*').order('note_number', { ascending:false });
  state.debitNotes = error ? [] : data;
}
async function insertDebitNote(rec){
  const { data, error } = await sb.from('debit_notes').insert(rec).select().single();
  if(error){ showToast(t('connError'), true); return null; }
  return data;
}
async function deleteDebitNote(id, pdfPath){
  if(pdfPath) await sb.storage.from('debit-notes').remove([pdfPath]);
  const { error } = await sb.from('debit_notes').delete().eq('id', id);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
async function uploadDebitPdf(folder, blob, filename){
  const path = `${folder}/${filename}`;
  const { error } = await sb.storage.from('debit-notes').upload(path, blob, { contentType: 'application/pdf' });
  if(error) return null;
  return path;
}
async function getDebitPdfUrl(path){
  const { data, error } = await sb.storage.from('debit-notes').createSignedUrl(path, 3600);
  return error ? null : data.signedUrl;
}
function downloadBlobToDevice(blob, filename){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.style.display = 'none';
  document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 60000);
}
function debitIsDone(n){ return !!(n && n.form_data && n.form_data.done); }
async function setDebitDone(n, done){
  const fd = Object.assign({}, n.form_data || {}, { done: !!done });
  const { error } = await sb.from('debit_notes').update({ form_data: fd }).eq('id', n.id);
  if(error){ showToast(t('connError'), true); return false; }
  n.form_data = fd;
  return true;
}
const DEBIT_MANUAL_FOLDER = 'manual';
async function listDebitManualPdfs(){
  const { data, error } = await sb.storage.from('debit-notes').list(DEBIT_MANUAL_FOLDER, { sortBy: { column:'name', order:'desc' } });
  state.debitPdfFiles = error ? [] : (data || []).filter(f => f.name !== '.emptyFolderPlaceholder' && f.id !== null);
}
async function uploadDebitManualPdf(file){
  const path = `${DEBIT_MANUAL_FOLDER}/${Date.now()}_${file.name}`;
  const { error } = await sb.storage.from('debit-notes').upload(path, file, { contentType: 'application/pdf' });
  if(error){ showToast(t('pdfUploadErr'), true); return false; }
  return true;
}
async function updateDebitNotePdfPath(id, pdfPath){
  const { error } = await sb.from('debit_notes').update({ pdf_path: pdfPath }).eq('id', id);
  return !error;
}
async function createUserApi({ username, name, password, role }){
  try{
    const { data, error } = await sb.rpc('admin_create_user', {
      p_username: username, p_name: name, p_password: password, p_role: role
    });
    if(error) return { ok:false, error: error.message };
    return data || { ok:false, error:'Réponse vide' };
  }catch(e){
    return { ok:false, error: String(e) };
  }
}
async function loadCatalog(){
  const { data, error } = await sb.from('ref_catalog').select('*').order('created_at', { ascending:true });
  state.catalog = error ? [] : data;
}
async function insertCatalog(c){
  const { error } = await sb.from('ref_catalog').insert({ ref:c.ref, designation:c.designation, fournisseur:c.fournisseur, prix:c.prix });
  if(error){ showToast(error.code === '23505' ? t('catalogExistsErr') : t('connError'), true); return false; }
  return true;
}
async function deleteCatalog(id){
  const { error } = await sb.from('ref_catalog').delete().eq('id', id);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}

/* ================= DATA LAYER (Supabase) — QC ================= */
function mapQc(row){
  return { id: row.id, date: row.record_date||'', fournisseur: row.fournisseur||'', ref: row.ref, designation: row.designation||'',
    qteControle: row.qte_controle||0, qteNonOk: row.qte_non_ok||0, grani: row.grani||0, rayure: row.rayure||0, trace: row.trace||0,
    piqure: row.piqure||0, coup: row.coup||0, cosse: row.cosse||0, autre: row.autre||0, user: row.user_name||'' };
}
async function loadQcRecords(){
  const { data, error } = await sb.from('qc_records').select('*').order('created_at', { ascending:false });
  if(error){ showToast(t('connError'), true); state.qcRecords = []; return; }
  state.qcRecords = data.map(mapQc);
}
async function insertQc(r){
  const { error } = await sb.from('qc_records').insert({
    record_date: r.date, fournisseur: r.fournisseur, ref: r.ref, designation: r.designation,
    qte_controle: r.qteControle, qte_non_ok: r.qteNonOk, grani: r.grani, rayure: r.rayure, trace: r.trace,
    piqure: r.piqure, coup: r.coup, cosse: r.cosse, autre: r.autre, user_name: r.user,
  });
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
async function updateQc(r){
  const { error } = await sb.from('qc_records').update({
    record_date: r.date, fournisseur: r.fournisseur, ref: r.ref, designation: r.designation,
    qte_controle: r.qteControle, qte_non_ok: r.qteNonOk, grani: r.grani, rayure: r.rayure, trace: r.trace,
    piqure: r.piqure, coup: r.coup, cosse: r.cosse, autre: r.autre,
  }).eq('id', r.id);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}
async function deleteQc(id){
  const { error } = await sb.from('qc_records').delete().eq('id', id);
  if(error){ showToast(t('connError'), true); return false; }
  return true;
}

function loadLang(){
  try{ const l = localStorage.getItem('stock:lang'); if(l==='ar'||l==='fr') state.lang = l; }catch(e){}
  applyDir();
}
function setLang(l){ state.lang = l; try{ localStorage.setItem('stock:lang', l); }catch(e){} applyDir(); render(); }
function applyDir(){
  document.documentElement.setAttribute('lang', state.lang);
  document.documentElement.setAttribute('dir', state.lang==='ar' ? 'rtl' : 'ltr');
  document.title = t('appTitle');
}

/* ================= ICONS ================= */
const ICONS = {
  pdf: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M6 2h9l5 5v15H6z"/><path d="M15 2v5h5"/><path d="M8.5 17v-4h1.3a1.4 1.4 0 010 2.8H8.5m4-2.8h1.6a1.2 1.2 0 011.2 1.2v.4a1.2 1.2 0 01-1.2 1.2H12.5m0-2.8V17m4-4v4m0-2h1.5"/></svg>`,
  debit: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="3" y="3" width="18" height="18" rx="1"/><path d="M7 8h10M7 12h10M7 16h6"/></svg>`,
  mail: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="3" y="5" width="18" height="14" rx="1"/><path d="M4 6l8 7 8-7"/></svg>`,
  entry: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 5v14M5 12h14"/></svg>`,
  data: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="3" y="4" width="18" height="16" rx="1"/><path d="M3 10h18M9 10v10"/></svg>`,
  dash: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="3" y="12" width="4" height="8"/><rect x="10" y="7" width="4" height="13"/><rect x="17" y="3" width="4" height="17"/></svg>`,
  users: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="9" cy="8" r="3.2"/><path d="M2.5 20c1-4 4-6 6.5-6s5.5 2 6.5 6"/><circle cx="17" cy="8" r="2.6"/><path d="M16 14.3c2 .4 4 2 4.8 5.7"/></svg>`,
  catalog: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><ellipse cx="12" cy="5.5" rx="8" ry="3"/><path d="M4 5.5v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6M4 11.5v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"/></svg>`,
  qc: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3M8.5 11l1.8 1.8L14 9.2"/></svg>`,
  back: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="14" height="14"><path d="M15 18l-6-6 6-6"/></svg>`,
  edit: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 20h9M16.5 3.5a2.1 2.1 0 013 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>`,
  del: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg>`,
  check: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M20 6L9 17l-5-5"/></svg>`,
  download: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" width="15" height="15" style="vertical-align:-2px;margin-inline-end:4px;"><path d="M12 3v12m0 0l-4-4m4 4l4-4M4 19h16"/></svg>`,
};

/* ================= RENDER ================= */
const app = document.getElementById('app');

function renderConfigError(){
  applyDir();
  app.innerHTML = `<div class="config-error"><div class="box"><h2>${t('configErrorTitle')}</h2><p>${t('configErrorBody')}</p></div></div>`;
}

/* ===== Entrepreneur : voit toute l'application en lecture seule (aucun ajout, modification ni suppression) ===== */
let entObserver = null;
function entCleanup(){
  const m = document.getElementById('main');
  if(!m) return;
  m.querySelectorAll('#card-entry, #card-qc-entry, #card-debit-new, #card-ccs-new, .row-actions button, .tile-del').forEach(el => el.remove());
  if(['catalog', 'pdf-library', 'debit-pdf', 'ccs-pdf'].includes(state.view)) m.querySelectorAll('.form-panel').forEach(el => el.remove());
  m.querySelectorAll('select[id^="cstatus-"], select[id^="dstatus-"]').forEach(el => { el.disabled = true; });
}
function entSetupReadOnly(main){
  if(entObserver){ entObserver.disconnect(); entObserver = null; }
  if(!main || !state.currentUser || state.currentUser.role !== 'entrepreneur') return;
  entCleanup();
  entObserver = new MutationObserver(() => {
    entObserver.disconnect();
    entCleanup();
    entObserver.observe(main, { childList: true, subtree: true });
  });
  entObserver.observe(main, { childList: true, subtree: true });
}

function render(){
  if(!configOk) return renderConfigError();
  if(state.view === 'login') return renderLogin();
  {
    const rl = state.currentUser && state.currentUser.role;
    if(rl === 'visitor' && state.view !== 'pdf-library') state.view = 'pdf-library';
    else if(rl === 'employee' && !['qc-entry','op-portal','op-table'].includes(state.view)) state.view = 'qc-entry';
    else if(rl === 'entrepreneur' && !['op-portal','op-table','qc-data'].includes(state.view)) state.view = 'op-portal';
    else if(rl !== 'admin' && (state.view === 'users' || state.view === 'recipients')) state.view = 'menu';
  }
  const otherLang = state.lang==='fr' ? 'ar' : 'fr';
  const otherLabel = state.lang==='fr' ? 'العربية' : 'Français';
  const isAdmin = state.currentUser && state.currentUser.role === 'admin';
  const nav = [
    ['analytics-hub','▦','Analytics Dashboard'],['stock-menu','▤',t('moduleStockTitle')],['qc-menu','◇',t('moduleQcTitle')],
    ['pdf-library','▧',t('cardPdfTitle')],['pdf-bq','▧',t('cardBqTitle')],['debit-menu','⇄',t('cardDebitTitle')],
    ['ccs-menu','▣',tc('cardTitle')],['op-portal','▥',t('cardOpTitle')]
  ];
  if(isAdmin){nav.push(['users','♙',t('cardUsersTitle')],['recipients','✉',t('cardRecipientsTitle')]);}
  const roleNow = state.currentUser && state.currentUser.role;
  if(roleNow === 'entrepreneur') nav.splice(0, nav.length, ['op-portal','▥',t('cardOpTitle')],['qc-data','◇',t('cardQcDataTitle')]);
  if(roleNow === 'employee') nav.splice(0, nav.length, ['qc-entry','◇',t('moduleQcTitle')],['op-portal','▥',t('cardOpTitle')]);
  const activeView = state.view === 'dashboard' ? 'analytics-hub' : (state.view === 'pdf-library' && state.currentLibrary === 'bonnequalite') ? 'pdf-bq' : state.view;
  app.innerHTML = `<div class="workspace-shell">
    <aside class="workspace-sidebar">
      <div class="workspace-brand"><span class="workspace-brand-mark">⌁</span><span>Gestion de Stock<small>Industrial Management</small></span></div>
      <nav class="workspace-nav" aria-label="Navigation principale">
        ${nav.map(([view,icon,label])=>`<button type="button" data-view="${view}" class="${activeView===view?'active':''}"><span class="nav-symbol">${icon}</span><span>${label}</span></button>`).join('')}
      </nav>
      <div class="workspace-logout-wrap"><button class="workspace-logout" id="logout-btn"><span class="nav-symbol">↪</span><span>${t('logout')}</span></button></div>
    </aside>
    <section class="workspace-main">
      <header class="workspace-topbar"><div class="workspace-search"><span>⌕</span><span>Search...</span></div><div class="workspace-user"><button class="lang-btn" id="lang-btn">${otherLabel}</button><span class="workspace-user-name">${escHtml(state.currentUser.name)}</span><span class="workspace-avatar">${escHtml((state.currentUser.name||'U').trim().charAt(0).toUpperCase())}</span></div></header>
      <main id="main" class="workspace-content"></main>
    </section>
  </div>`;
  document.getElementById('logout-btn').onclick = async () => { await sb.auth.signOut(); state.currentUser = null; state.view='login'; render(); };
  document.getElementById('lang-btn').onclick = () => setLang(otherLang);
  app.querySelectorAll('.workspace-nav [data-view]').forEach(btn=>btn.onclick=async()=>{
    const view=btn.dataset.view;
    if(view==='analytics-hub'){await loadRecords();await loadQcRecords();}
    if(view==='stock-menu'){await loadRecords();}
    if(view==='qc-menu'){await loadCatalog();await loadQcRecords();}
    if(view==='pdf-library'){state.currentLibrary='pdf';state.pdfFolder='';await listPdfEntries('');}
    if(view==='pdf-bq'){state.currentLibrary='bonnequalite';state.pdfFolder='';await listPdfEntries('');state.view='pdf-library';render();return;}
    if(view==='recipients') await loadRecipients();
    if(view==='qc-entry') await loadCatalog();
    if(view==='qc-data') await loadQcRecords();
    state.view=view;render();
  });
  const main = document.getElementById('main');
  if(state.view === 'menu') renderMenu(main);
  if(state.view === 'stock-menu') renderStockMenu(main);
  if(state.view === 'qc-menu') renderQcMenu(main);
  if(state.view === 'entry') renderEntry(main);
  if(state.view === 'data') renderData(main);
  if(state.view === 'dashboard') renderAnalyticsHub(main);
  if(state.view === 'analytics-hub') renderAnalyticsHub(main);
  if(state.view === 'users') renderUsers(main);
  if(state.view === 'recipients') renderRecipients(main);
  if(state.view === 'pdf-library') renderPdfLibrary(main);
  if(state.view === 'debit-menu') renderDebitMenu(main);
  if(state.view === 'debit-new') renderDebitEntry(main);
  if(state.view === 'debit-list') renderDebitList(main);
  if(state.view === 'debit-pdf') renderDebitPdfWindow(main);
  if(state.view === 'ccs-menu') renderCcsMenu(main);
  if(state.view === 'ccs-new') renderCcsNew(main);
  if(state.view === 'ccs-list') renderCcsList(main);
  if(state.view === 'ccs-pdf') renderCcsPdf(main);
  if(state.view === 'op-portal') renderOpPortal(main);
  if(state.view === 'op-table') renderOpTable(main);
  if(state.view === 'op-history') renderOpHistory(main);
  if(state.view === 'catalog') renderCatalog(main);
  if(state.view === 'qc-entry') renderQcEntry(main);
  if(state.view === 'qc-data') renderQcData(main);
  if(state.view === 'qc-dashboard') renderQcDashboard(main);
  entSetupReadOnly(main);
}

function renderLogin(){
  const otherLang = state.lang==='fr' ? 'ar' : 'fr';
  const otherLabel = state.lang==='fr' ? 'العربية' : 'Français';
  app.innerHTML = `
    <div class="login-wrap">
      <button class="lang-btn login-lang" id="lang-btn">${otherLabel}</button>
      <div class="login-card">
        <div class="dot"></div>
        <h1>${t('appTitle')}</h1>
        <p class="sub">${t('loginSubtitle')}</p>
        <div id="login-error"></div>
        <label>${t('loginUser')}</label>
        <input id="login-user" type="text" autocomplete="username">
        <label>${t('loginPass')}</label>
        <input id="login-pass" type="password" autocomplete="current-password">
        <button class="btn" id="login-btn">${t('loginBtn')}</button>
      </div>
    </div>`;
  document.getElementById('lang-btn').onclick = () => setLang(otherLang);
  const tryLogin = async () => {
    const u = document.getElementById('login-user').value.trim();
    const p = document.getElementById('login-pass').value;
    const errBox = document.getElementById('login-error');
    const btn = document.getElementById('login-btn');
    if(!u || !p) return;
    btn.disabled = true; btn.textContent = t('loginBtnLoading');
    const { data, error } = await sb.auth.signInWithPassword({ email: usernameToEmail(u), password: p });
    if(error || !data.user){
      errBox.innerHTML = `<div class="error">${t('loginError')}</div>`;
      btn.disabled = false; btn.textContent = t('loginBtn');
      return;
    }
    const { data: profile } = await sb.from('profiles').select('*').eq('id', data.user.id).single();
    state.currentUser = { id: data.user.id, username: u, name: profile ? profile.name : u, role: profile ? profile.role : 'user' };
    await enterByRole();
    render();
  };
  document.getElementById('login-btn').onclick = tryLogin;
  document.getElementById('login-pass').addEventListener('keydown', e => { if(e.key==='Enter') tryLogin(); });
}

async function enterByRole(){
  const role = state.currentUser.role;
  if(role === 'employee'){
    await loadCatalog();
    state.view = 'qc-entry';
  }else if(role === 'entrepreneur'){
    state.view = 'op-portal';
  }else if(role === 'visitor'){
    state.currentLibrary = 'pdf'; state.pdfFolder = '';
    await listPdfEntries('');
    state.view = 'pdf-library';
  }else{
    state.view = 'analytics-hub';
    await loadRecords();
    await loadQcRecords();
  }
}

function renderMenu(main){
  const isAdmin = state.currentUser.role === 'admin';
  main.innerHTML = `
    <div class="menu-head"><h1>${t('menuTitle')}</h1><p>${t('menuSubtitle')}</p></div>
    <div class="menu-grid">
      <div class="menu-card analytics-entry" id="card-analytics">${ICONS.dash}<h3>Analytics Dashboard</h3><p>Vue industrielle : stock, qualité, activité et indicateurs de gestion.</p></div>
      <div class="menu-card" id="card-stock">${ICONS.data}<h3>${t('moduleStockTitle')}</h3><p>${t('moduleStockDesc')}</p></div>
      <div class="menu-card" id="card-qc">${ICONS.qc}<h3>${t('moduleQcTitle')}</h3><p>${t('moduleQcDesc')}</p></div>
      <div class="menu-card" id="card-pdf">${ICONS.pdf}<h3>${t('cardPdfTitle')}</h3><p>${t('cardPdfDesc')}</p></div>
      <div class="menu-card" id="card-bq">${ICONS.pdf}<h3>${t('cardBqTitle')}</h3><p>${t('cardBqDesc')}</p></div>
      <div class="menu-card" id="card-debit">${ICONS.debit}<h3>${t('cardDebitTitle')}</h3><p>${t('cardDebitDesc')}</p></div>
      <div class="menu-card" id="card-ccs">${ICONS.entry}<h3>${tc('cardTitle')}</h3><p>${tc('cardDesc')}</p></div>
      <div class="menu-card" id="card-op">${ICONS.qc}<h3>${t('cardOpTitle')}</h3><p>${t('cardOpDesc')}</p></div>
      ${isAdmin ? `<div class="menu-card" id="card-users">${ICONS.users}<h3>${t('cardUsersTitle')}</h3><p>${t('cardUsersDesc')}</p></div>` : ''}
      ${isAdmin ? `<div class="menu-card" id="card-recipients">${ICONS.mail}<h3>${t('cardRecipientsTitle')}</h3><p>${t('cardRecipientsDesc')}</p></div>` : ''}
    </div>`;
  document.getElementById('card-analytics').onclick = async () => { await loadRecords(); await loadQcRecords(); state.view='analytics-hub'; render(); };
  document.getElementById('card-stock').onclick = () => { state.view='stock-menu'; render(); };
  document.getElementById('card-qc').onclick = async () => { await loadCatalog(); await loadQcRecords(); state.view='qc-menu'; render(); };
  document.getElementById('card-pdf').onclick = async () => { state.currentLibrary='pdf'; state.pdfFolder=''; await listPdfEntries(''); state.view='pdf-library'; render(); };
  document.getElementById('card-bq').onclick = async () => { state.currentLibrary='bonnequalite'; state.pdfFolder=''; await listPdfEntries(''); state.view='pdf-library'; render(); };
  document.getElementById('card-debit').onclick = () => { state.view='debit-menu'; render(); };
  document.getElementById('card-ccs').onclick = () => { state.view='ccs-menu'; render(); };
  document.getElementById('card-op').onclick = () => { state.view='op-portal'; render(); };
  if(isAdmin) document.getElementById('card-users').onclick = () => { state.view='users'; render(); };
  if(isAdmin) document.getElementById('card-recipients').onclick = async () => { await loadRecipients(); state.view='recipients'; render(); };
}

function renderDebitMenu(main){
  main.innerHTML = `
    <div class="menu-head"><h1>${t('cardDebitTitle')}</h1><p>${t('menuSubtitle')}</p></div>
    <div class="menu-grid">
      <div class="menu-card" id="card-debit-new">${ICONS.entry}<h3>${t('debitNewTitle')}</h3><p>${t('debitNewDesc')}</p></div>
      <div class="menu-card" id="card-debit-list">${ICONS.data}<h3>${t('debitListTitle')}</h3><p>${t('debitListDesc')}</p></div>
      <div class="menu-card" id="card-debit-pdf">${ICONS.pdf}<h3>${t('debitPdfTitle')}</h3><p>${t('debitPdfDesc')}</p></div>
    </div>
    <div class="back-link" id="back-top" style="margin-top:22px;">${ICONS.back} ${t('menuTitle')}</div>`;
  document.getElementById('card-debit-new').onclick = async () => { await loadCatalog(); state.view='debit-new'; render(); };
  document.getElementById('card-debit-list').onclick = async () => { await loadDebitNotes(); state.view='debit-list'; render(); };
  document.getElementById('card-debit-pdf').onclick = async () => { await listDebitManualPdfs(); state.view='debit-pdf'; render(); };
  document.getElementById('back-top').onclick = () => { state.view='menu'; render(); };
}

function renderStockMenu(main){
  main.innerHTML = `
    <div class="menu-head"><h1>${t('moduleStockTitle')}</h1><p>${t('menuSubtitle')}</p></div>
    <div class="menu-grid">
      <div class="menu-card" id="card-entry">${ICONS.entry}<h3>${t('cardEntryTitle')}</h3><p>${t('cardEntryDesc')}</p></div>
      <div class="menu-card" id="card-data">${ICONS.data}<h3>${t('cardDataTitle')}</h3><p>${t('cardDataDesc')}</p></div>
      <div class="menu-card" id="card-dash">${ICONS.dash}<h3>${t('cardDashTitle')}</h3><p>${t('cardDashDesc')}</p></div>
      <div class="menu-card" id="card-catalog">${ICONS.catalog}<h3>${t('cardCatalogTitle')}</h3><p>${t('cardCatalogDesc')}</p></div>
    </div>
    <div class="back-link" id="back-top" style="margin-top:22px;">${ICONS.back} ${t('menuTitle')}</div>`;
  document.getElementById('card-entry').onclick = () => { state.view='entry'; render(); };
  document.getElementById('card-data').onclick  = async () => { await loadCatalog(); state.view='data'; render(); };
  document.getElementById('card-dash').onclick  = () => { state.view='dashboard'; render(); };
  document.getElementById('card-catalog').onclick = async () => { await loadCatalog(); state.view='catalog'; render(); };
  document.getElementById('back-top').onclick = () => { state.view='menu'; render(); };
}

function renderQcMenu(main){
  main.innerHTML = `
    <div class="menu-head"><h1>${t('moduleQcTitle')}</h1><p>${t('menuSubtitle')}</p></div>
    <div class="menu-grid">
      <div class="menu-card" id="card-qc-entry">${ICONS.entry}<h3>${t('cardQcEntryTitle')}</h3><p>${t('cardQcEntryDesc')}</p></div>
      <div class="menu-card" id="card-qc-data">${ICONS.data}<h3>${t('cardQcDataTitle')}</h3><p>${t('cardQcDataDesc')}</p></div>
      <div class="menu-card" id="card-qc-dash">${ICONS.dash}<h3>${t('cardQcDashTitle')}</h3><p>${t('cardQcDashDesc')}</p></div>
    </div>
    <div class="back-link" id="back-top" style="margin-top:22px;">${ICONS.back} ${t('menuTitle')}</div>`;
  document.getElementById('card-qc-entry').onclick = () => { state.view='qc-entry'; render(); };
  document.getElementById('card-qc-data').onclick  = () => { state.view='qc-data'; render(); };
  document.getElementById('card-qc-dash').onclick  = () => { state.view='qc-dashboard'; render(); };
  document.getElementById('back-top').onclick = () => { state.view='menu'; render(); };
}

function renderEntry(main){
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.entry.replace('<svg','<svg width="18" height="18"')} ${t('cardEntryTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="form-panel">
      <div class="form-grid">
        <div><label>${t('fieldLot')}</label><input id="f-lot" type="text" placeholder="LOT-0001"></div>
        <div><label>${t('fieldRef')}</label><input id="f-ref" type="text" placeholder="REF-0001" list="ref-catalog-list"></div>
        <datalist id="ref-catalog-list">${state.catalog.map(c=>`<option value="${escAttr(c.ref)}">`).join('')}</datalist>
        <div><label>${t('fieldDesig')}</label><input id="f-desig" type="text"></div>
        <div><label>${t('fieldDefaut')}</label><input id="f-defaut" type="text"></div>
        <div><label>${t('fieldQte')}</label><input id="f-qte" type="number" min="0" placeholder="0"></div>
        <div><label>${t('fieldLoc')}</label><input id="f-loc" type="text"></div>
        <div><label>${t('fieldFourn')}</label><input id="f-fourn" type="text"></div>
        <div><label>${t('fieldUser')}</label><input id="f-user" type="text" value="${escAttr(state.currentUser.name)}" disabled></div>
        <div><label>${t('fieldDate')}</label><input id="f-date" type="date" value="${todayISO()}"></div>
      </div>
      <div class="form-actions">
        <button class="btn" id="save-btn" style="max-width:200px;">${t('saveBtn')}</button>
        <button class="btn secondary" id="clear-btn">${t('clearBtn')}</button>
      </div>
    </div>`;
  document.getElementById('back').onclick = () => { state.view='stock-menu'; render(); };
  document.getElementById('clear-btn').onclick = () => renderEntry(main);
  const refInput = document.getElementById('f-ref');
  refInput.addEventListener('input', () => {
    const match = state.catalog.find(c => c.ref.toLowerCase() === refInput.value.trim().toLowerCase());
    if(match){ document.getElementById('f-desig').value = match.designation; document.getElementById('f-fourn').value = match.fournisseur; }
  });
  document.getElementById('save-btn').onclick = async () => {
    const ref = document.getElementById('f-ref').value.trim();
    const desig = document.getElementById('f-desig').value.trim();
    if(!ref || !desig){ showToast(t('requiredError'), true); return; }
    const rec = {
      lot: document.getElementById('f-lot').value.trim(), ref, designation: desig,
      defaut: document.getElementById('f-defaut').value.trim(), qtite: Number(document.getElementById('f-qte').value)||0,
      location: document.getElementById('f-loc').value.trim(), fournisseur: document.getElementById('f-fourn').value.trim(),
      user: state.currentUser.name, date: document.getElementById('f-date').value || todayISO(),
    };
    const btn = document.getElementById('save-btn'); btn.disabled = true;
    const ok = await insertRecord(rec);
    btn.disabled = false;
    if(ok){ await loadRecords(); showToast(t('savedToast')); renderEntry(main); }
  };
}

function renderData(main){
  const filtered = state.records.filter(r => DATA_COLS.every(c => {
    const q = (state.search[c.key]||'').toLowerCase().trim();
    if(!q) return true;
    return (r[c.field] ?? '').toString().toLowerCase().includes(q);
  }));
  const anyFilter = DATA_COLS.some(c => (state.search[c.key]||'').trim());
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.data.replace('<svg','<svg width="18" height="18"')} ${t('cardDataTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="toolbar"><button class="btn blue" id="export-btn" style="width:auto;padding:10px 16px;">${ICONS.download} ${t('exportBtn')}</button></div>
    <div class="table-wrap">
      ${state.records.length===0 ? `<div class="empty-state">${t('noRecords')}.</div>` : `
      <table>
        <thead>
          <tr>${DATA_COLS.map(c=>`<th>${t(c.label)}</th>`).join('')}<th>${t('fieldPrix')}</th><th></th></tr>
          <tr class="filter-row">${DATA_COLS.map(c=>`<th><input class="col-search" data-key="${c.key}" placeholder="${t(c.label)}" value="${escAttr(state.search[c.key]||'')}"></th>`).join('')}<th></th><th></th></tr>
        </thead>
        <tbody id="tbody">
          ${filtered.length===0 ? `<tr><td colspan="${DATA_COLS.length+2}" class="empty-state">${t('noRecords')}${anyFilter?t('noRecordsSearch'):''}.</td></tr>` : filtered.map(r=>renderRow(r)).join('')}
        </tbody>
      </table>`}
    </div>`;
  document.getElementById('back').onclick = () => { state.view='stock-menu'; render(); };
  document.getElementById('export-btn').onclick = () => exportToExcel(filtered);
  main.querySelectorAll('.col-search').forEach(inp => {
    inp.oninput = () => { state.search[inp.dataset.key] = inp.value; state.lastSearchFocus = inp.dataset.key; renderData(main); };
  });
  if(state.lastSearchFocus){
    const el = main.querySelector(`.col-search[data-key="${state.lastSearchFocus}"]`);
    if(el){ el.focus(); el.selectionStart = el.selectionEnd = el.value.length; }
  }
  filtered.forEach(r => {
    const editBtn = document.getElementById('edit-'+r.id);
    const delBtn = document.getElementById('del-'+r.id);
    if(editBtn) editBtn.onclick = () => { state.editingId = (state.editingId===r.id?null:r.id); renderData(main); };
    if(delBtn) delBtn.onclick = async () => {
      if(confirm(t('confirmDelete'))){
        const ok = await deleteRecord(r.id);
        if(ok){ await loadRecords(); showToast(t('deletedToast')); renderData(main); }
      }
    };
    if(state.editingId === r.id){
      const saveBtn = document.getElementById('rowsave-'+r.id);
      if(saveBtn) saveBtn.onclick = async () => {
        r.lot = document.getElementById('e-lot-'+r.id).value.trim();
        r.ref = document.getElementById('e-ref-'+r.id).value.trim() || r.ref;
        r.designation = document.getElementById('e-desig-'+r.id).value.trim() || r.designation;
        r.defaut = document.getElementById('e-defaut-'+r.id).value.trim();
        r.qtite = Number(document.getElementById('e-qte-'+r.id).value) || 0;
        r.location = document.getElementById('e-loc-'+r.id).value.trim();
        r.fournisseur = document.getElementById('e-fourn-'+r.id).value.trim();
        r.date = document.getElementById('e-date-'+r.id).value || r.date;
        const ok = await updateRecord(r);
        if(ok){ state.editingId = null; showToast(t('editedToast')); renderData(main); }
      };
    }
  });
}

function renderRow(r){
  if(state.editingId === r.id){
    return `<tr>
      <td><input id="e-lot-${r.id}" value="${escAttr(r.lot)}"></td>
      <td><input id="e-ref-${r.id}" value="${escAttr(r.ref)}"></td>
      <td><input id="e-desig-${r.id}" value="${escAttr(r.designation)}"></td>
      <td><input id="e-defaut-${r.id}" value="${escAttr(r.defaut)}"></td>
      <td><input id="e-qte-${r.id}" type="number" value="${r.qtite}" style="width:70px"></td>
      <td><input id="e-loc-${r.id}" value="${escAttr(r.location)}"></td>
      <td><input id="e-fourn-${r.id}" value="${escAttr(r.fournisseur)}"></td>
      <td>${escHtml(r.user)}</td>
      <td><input id="e-date-${r.id}" type="date" value="${r.date}"></td>
      <td class="num">${calcPrixTotal(r).toFixed(2)}</td>
      <td class="row-actions">
        <button class="icon-btn" id="rowsave-${r.id}" title="${t('saveBtn')}">${ICONS.check}</button>
        <button class="icon-btn del" id="del-${r.id}" title="${t('deletedToast')}">${ICONS.del}</button>
      </td>
    </tr>`;
  }
  return `<tr>
    <td class="ref-cell">${escHtml(r.lot)||'—'}</td>
    <td class="ref-cell">${escHtml(r.ref)}</td>
    <td>${escHtml(r.designation)}</td>
    <td>${escHtml(r.defaut)||'—'}</td>
    <td class="num">${r.qtite}</td>
    <td>${escHtml(r.location)||'—'}</td>
    <td>${escHtml(r.fournisseur)||'—'}</td>
    <td>${escHtml(r.user)}</td>
    <td class="num">${r.date}</td>
    <td class="num">${calcPrixTotal(r).toFixed(2)}</td>
    <td class="row-actions">
      <button class="icon-btn" id="edit-${r.id}" title="${t('editedToast')}">${ICONS.edit}</button>
      <button class="icon-btn del" id="del-${r.id}" title="${t('deletedToast')}">${ICONS.del}</button>
    </td>
  </tr>`;
}

/* ===== Filtre par date commun aux 3 tableaux de bord (Analytics, Stock, Contrôle qualité) ===== */
function dashInRange(r){
  const d = ((r && r.date) || '').toString().slice(0, 10);
  if(state.dashFrom && d < state.dashFrom) return false;
  if(state.dashTo && d > state.dashTo) return false;
  return true;
}
function dashBarHtml(){
  const L = state.lang === 'ar'
    ? { from: 'من', to: 'إلى', today: 'اليوم', w: '7 أيام', m: 'هذا الشهر', all: 'الكل' }
    : { from: 'Du', to: 'Au', today: "Aujourd'hui", w: '7 jours', m: 'Ce mois', all: 'Tout' };
  return `<style>.dash-bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 16px;padding:10px 12px;background:var(--panel,#161b26);border:1px solid var(--border,#2a3850);border-radius:8px}
    .dash-bar label{font-size:13px;font-weight:600;margin:0}
    .dash-bar input[type=date]{width:auto;margin:0;padding:7px 9px;font-size:14px}
    .dash-bar button{width:auto;margin:0;padding:7px 12px;font-size:13px;cursor:pointer;font-family:inherit;border-radius:4px;border:1px solid var(--border,#2a3850);background:var(--panel-alt,#1d2431);color:var(--text,#e6eaf2)}
    .dash-bar button.on{background:var(--accent,#d68c45);color:#161311;border-color:var(--accent,#d68c45)}</style>
    <div class="dash-bar">
      <label for="dash-from">${L.from}</label><input type="date" id="dash-from" value="${escAttr(state.dashFrom || '')}">
      <label for="dash-to">${L.to}</label><input type="date" id="dash-to" value="${escAttr(state.dashTo || '')}">
      <button type="button" id="dash-today">${L.today}</button><button type="button" id="dash-7">${L.w}</button><button type="button" id="dash-month">${L.m}</button><button type="button" id="dash-all">${L.all}</button>
    </div>`;
}
function dashBarWire(rerender){
  const iso = (y, m, d) => `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  const n = opNowTunis();
  const today = iso(n.y, n.m, n.d);
  const set = (f, t2) => { state.dashFrom = f; state.dashTo = t2; rerender(); };
  const $ = (id) => document.getElementById(id);
  $('dash-from').onchange = (e) => { state.dashFrom = e.target.value; rerender(); };
  $('dash-to').onchange = (e) => { state.dashTo = e.target.value; rerender(); };
  $('dash-today').onclick = () => set(today, today);
  $('dash-7').onclick = () => { const d = new Date(Date.UTC(n.y, n.m - 1, n.d - 6)); set(d.toISOString().slice(0, 10), today); };
  $('dash-month').onclick = () => set(iso(n.y, n.m, 1), today);
  $('dash-all').onclick = () => set('', '');
  /* met en évidence le bouton actif */
  const f = state.dashFrom || '', t2 = state.dashTo || '';
  const d7 = new Date(Date.UTC(n.y, n.m - 1, n.d - 6)).toISOString().slice(0, 10);
  const active = (!f && !t2) ? 'dash-all' : (f === today && t2 === today) ? 'dash-today' : (f === d7 && t2 === today) ? 'dash-7' : (f === iso(n.y, n.m, 1) && t2 === today) ? 'dash-month' : '';
  if(active) $(active).classList.add('on');
}

let chartRefs = [];
function renderDashboard(main){
  chartRefs.forEach(c=>c.destroy()); chartRefs = [];
  const recs = state.records.filter(dashInRange);
  const totalQte = recs.reduce((s,r)=>s+r.qtite,0);
  const fournisseurs = [...new Set(recs.map(r=>r.fournisseur).filter(Boolean))];
  const locations = [...new Set(recs.map(r=>r.location).filter(Boolean))];
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.dash.replace('<svg','<svg width="18" height="18"')} ${t('cardDashTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    ${dashBarHtml()}
    <div class="stat-grid">
      <div class="stat-card"><div class="label">${t('statRecords')}</div><div class="value">${recs.length}</div></div>
      <div class="stat-card"><div class="label">${t('statQty')}</div><div class="value accent">${totalQte}</div></div>
      <div class="stat-card"><div class="label">${t('statFourn')}</div><div class="value blue">${fournisseurs.length}</div></div>
      <div class="stat-card"><div class="label">${t('statLoc')}</div><div class="value blue">${locations.length}</div></div>
    </div>
    ${recs.length===0 ? `<div class="empty-state" style="background:var(--panel);border:1px solid var(--border);border-radius:var(--radius);">${t('noData')}</div>` : `
    <div class="charts-grid">
      <div class="chart-panel"><h3>${t('chartFournTitle')}</h3><canvas id="chart-fourn"></canvas></div>
      <div class="chart-panel"><h3>${t('chartDefautTitle')}</h3><canvas id="chart-defaut"></canvas></div>
      <div class="chart-panel"><h3>${t('chartLocTitle')}</h3><canvas id="chart-loc"></canvas></div>
      <div class="chart-panel"><h3>${t('chartUserTitle')}</h3><canvas id="chart-user"></canvas></div>
      <div class="chart-panel wide"><h3>${t('chartRefFournTitle')}</h3><canvas id="chart-ref-fourn"></canvas></div>
    </div>`}`;
  document.getElementById('back').onclick = () => { state.view='stock-menu'; render(); };
  dashBarWire(() => renderDashboard(main));
  if(recs.length===0) return;
  const palette = ['#d68c45','#5b8fb0','#5aad8c','#e2574c','#8a93a3','#a687c9','#c9a05b','#6fb0c9','#c97878'];
  const undef = t('undefinedLabel');
  const groupSum = (key)=>{ const m={}; recs.forEach(r=>{ const k=r[key]||undef; m[k]=(m[k]||0)+r.qtite; }); return m; };
  const groupCount = (key)=>{ const m={}; recs.forEach(r=>{ const k=r[key]||undef; m[k]=(m[k]||0)+1; }); return m; };
  const mkBar = (id,dataMap,label)=>{ const labels=Object.keys(dataMap); chartRefs.push(new Chart(document.getElementById(id),{type:'bar',data:{labels,datasets:[{label,data:Object.values(dataMap),backgroundColor:labels.map((_,i)=>palette[i%palette.length]),borderRadius:2}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#8a93a3'},grid:{display:false}},y:{ticks:{color:'#8a93a3'},grid:{color:'#313949'}}}}})); };
  const mkPie = (id,dataMap)=>{ const labels=Object.keys(dataMap); chartRefs.push(new Chart(document.getElementById(id),{type:'doughnut',data:{labels,datasets:[{data:Object.values(dataMap),backgroundColor:labels.map((_,i)=>palette[i%palette.length])}]},options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:'#8a93a3',boxWidth:12,font:{size:11}}}}}})); };
  mkBar('chart-fourn', groupSum('fournisseur'), t('statQty'));
  mkPie('chart-defaut', groupCount('defaut'));
  mkBar('chart-loc', groupSum('location'), t('statQty'));
  mkBar('chart-user', groupCount('user'), t('statRecords'));
  const refs = [...new Set(recs.map(r=>r.ref||undef))];
  const fournList = [...new Set(recs.map(r=>r.fournisseur||undef))];
  const datasets = fournList.map((f,i)=>({ label:f, data: refs.map(ref=>recs.filter(r=>(r.ref||undef)===ref && (r.fournisseur||undef)===f).reduce((s,r)=>s+r.qtite,0)), backgroundColor: palette[i%palette.length], borderRadius:2 }));
  chartRefs.push(new Chart(document.getElementById('chart-ref-fourn'),{type:'bar',data:{labels:refs,datasets},options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:'#8a93a3',boxWidth:12,font:{size:11}}}},scales:{x:{stacked:true,ticks:{color:'#8a93a3'},grid:{display:false}},y:{stacked:true,ticks:{color:'#8a93a3'},grid:{color:'#313949'}}}}}));
}

function renderAnalyticsHub(main){
  const recs = (Array.isArray(state.records) ? state.records : []).filter(dashInRange);
  const qcs = (Array.isArray(state.qcRecords) ? state.qcRecords : []).filter(dashInRange);
  const totalQty = recs.reduce((n,r)=>n+(Number(r.qtite)||0),0);
  const supplierCount = new Set(recs.map(r=>(r.fournisseur||'').trim()).filter(Boolean)).size;
  const locCount = new Set(recs.map(r=>(r.location||'').trim()).filter(Boolean)).size;
  const inspected = qcs.reduce((n,r)=>n+(Number(r.qteControle)||0),0);
  const rejects = qcs.reduce((n,r)=>n+(Number(r.qteNonOk)||0),0);
  const rejectPct = inspected ? rejects/inspected*100 : 0;
  const qtyBySupplier = {};
  recs.forEach(r=>{const k=(r.fournisseur||'Non renseigné').trim()||'Non renseigné';qtyBySupplier[k]=(qtyBySupplier[k]||0)+(Number(r.qtite)||0);});
  const topSuppliers = Object.entries(qtyBySupplier).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const maxSupplier = Math.max(1,...topSuppliers.map(x=>x[1]));
  const fmt = n => Number(n||0).toLocaleString(state.lang==='ar'?'ar-TN':'fr-FR');
  main.innerHTML = `<section class="analytics-shell">
    <header class="an-head"><div><div class="an-eyebrow">IndustryOS · AI Monitoring</div><h2>Analytics Dashboard</h2><p class="an-sub">Vue opérationnelle fondée sur les données actuellement enregistrées dans votre application.</p></div><button class="an-back" id="an-back">← Retour au menu</button></header>
    ${dashBarHtml()}
    <div class="an-grid">
      <div class="an-card"><div class="an-label">▦ Enregistrements de stock</div><div class="an-value an-blue">${fmt(recs.length)}</div><div class="an-note">Nombre de fiches stock chargées</div></div>
      <div class="an-card"><div class="an-label">⇅ Quantité totale en stock</div><div class="an-value an-cyan">${fmt(totalQty)}</div><div class="an-note">Somme des quantités enregistrées</div></div>
      <div class="an-card"><div class="an-label">◇ Fournisseurs</div><div class="an-value an-purple">${fmt(supplierCount)}</div><div class="an-note">Fournisseurs distincts dans les fiches</div></div>
      <div class="an-card"><div class="an-label">⌖ Emplacements</div><div class="an-value an-green">${fmt(locCount)}</div><div class="an-note">Emplacements distincts renseignés</div></div>
    </div>
    <div class="an-layout">
      <section class="an-panel"><h3>Répartition des quantités par fournisseur</h3><p class="an-desc">Top 5 — calculé à partir des fiches de stock chargées</p>
        ${topSuppliers.length ? `<div class="an-bars">${topSuppliers.map(([name,val],i)=>`<div><div class="an-bar-head"><span>${escHtml(name)}</span><strong>${fmt(val)}</strong></div><div class="an-track"><div class="an-fill" style="width:${Math.max(2,val/maxSupplier*100)}%;opacity:${1-i*.1}"></div></div></div>`).join('')}</div>` : `<div class="empty-state" style="color:#91a0b8">Aucune donnée de stock disponible pour le moment.</div>`}
      </section>
      <section class="an-panel"><h3>Qualité · Contrôle</h3><p class="an-desc">Indicateurs calculés à partir des enregistrements de contrôle qualité</p>
        <div class="an-value">${fmt(inspected)}</div><div class="an-note">Pièces contrôlées</div><div style="height:1px;background:#202c40;margin:15px 0"></div>
        <div class="an-bar-head"><span>Pièces non conformes</span><strong>${fmt(rejects)}</strong></div><div class="an-track"><div class="an-fill" style="width:${Math.min(100,rejectPct)}%;background:linear-gradient(90deg,#f59e0b,#f87171)"></div></div><div class="an-note">Taux de non-conformité : ${rejectPct.toLocaleString('fr-FR',{maximumFractionDigits:1})}% · ${qcs.length} fiches de contrôle</div>
      </section>
    </div>
    <div class="an-bottom">
      <div class="an-mini"><span>Activité stock</span><strong>${fmt(recs.length)}</strong><span>fiches disponibles</span></div>
      <div class="an-mini"><span>Contrôle qualité</span><strong>${fmt(qcs.length)}</strong><span>fiches de contrôle disponibles</span></div>
      <div class="an-mini"><span>Couverture des données</span><strong>${recs.length || qcs.length ? 'Active' : 'À alimenter'}</strong><span>Les indicateurs évoluent avec les données enregistrées</span></div>
    </div>
    <div class="an-foot">Les valeurs affichées proviennent des données chargées dans l’application. Aucun chiffre de démonstration n’est ajouté. Les coûts, arrêts machine et économies ne sont pas affichés tant qu’une source réelle correspondante n’est pas disponible.</div>
  </section>`;
  document.getElementById('an-back').onclick=async()=>{await loadRecords();await loadQcRecords();state.view='analytics-hub';render();};
  dashBarWire(() => renderAnalyticsHub(main));
}

/* ================= CATALOGUE ================= */
function renderCatalog(main){
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.catalog.replace('<svg','<svg width="18" height="18"')} ${t('cardCatalogTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="banner">${t('catalogBanner')}</div>
    <div class="form-panel" style="margin-bottom:16px;">
      <div class="user-form">
        <div><label>${t('fieldRef')}</label><input id="c-ref" type="text" placeholder="REF-0001"></div>
        <div><label>${t('fieldDesig')}</label><input id="c-desig" type="text"></div>
        <div><label>${t('fieldFourn')}</label><input id="c-fourn" type="text"></div>
        <div><label>${t('fieldPrix')}</label><input id="c-prix" type="number" min="0" step="0.01" placeholder="0.00"></div>
      </div>
      <button class="btn" id="add-catalog-btn" style="max-width:220px;">${t('addCatalogBtn')}</button>
    </div>
    <div class="table-wrap">
      ${state.catalog.length===0 ? `<div class="empty-state">${t('noCatalog')}.</div>` : `
      <table style="min-width:0;">
        <thead><tr><th>${t('fieldRef')}</th><th>${t('fieldDesig')}</th><th>${t('fieldFourn')}</th><th>${t('fieldPrix')}</th><th></th></tr></thead>
        <tbody>
          ${state.catalog.map(c => `
            <tr>
              <td class="ref-cell">${escHtml(c.ref)}</td>
              <td>${escHtml(c.designation)}</td>
              <td>${escHtml(c.fournisseur)}</td>
              <td class="num">${Number(c.prix||0).toFixed(2)}</td>
              <td class="row-actions"><button class="icon-btn del" id="cdel-${c.id}" title="${t('deletedToast')}">${ICONS.del}</button></td>
            </tr>`).join('')}
        </tbody>
      </table>`}
    </div>`;
  document.getElementById('back').onclick = () => { state.view='stock-menu'; render(); };
  document.getElementById('add-catalog-btn').onclick = async () => {
    const ref = document.getElementById('c-ref').value.trim();
    const designation = document.getElementById('c-desig').value.trim();
    const fournisseur = document.getElementById('c-fourn').value.trim();
    const prix = Number(document.getElementById('c-prix').value) || 0;
    if(!ref || !designation || !fournisseur){ showToast(t('catalogFieldsErr'), true); return; }
    const ok = await insertCatalog({ ref, designation, fournisseur, prix });
    if(ok){ await loadCatalog(); showToast(t('catalogAddedToast')); renderCatalog(main); }
  };
  state.catalog.forEach(c => {
    const btn = document.getElementById('cdel-'+c.id);
    if(!btn) return;
    btn.onclick = async () => {
      if(confirm(t('confirmDeleteCatalog'))){
        const ok = await deleteCatalog(c.id);
        if(ok){ await loadCatalog(); showToast(t('catalogDeletedToast')); renderCatalog(main); }
      }
    };
  });
}

/* ================= QC MODULE ================= */
function renderQcEntry(main){
  const isEmployee = state.currentUser.role === 'employee';
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.entry.replace('<svg','<svg width="18" height="18"')} ${t('cardQcEntryTitle')}</h2>${isEmployee ? '' : `<div class="back-link" id="back">${ICONS.back} ${t('back')}</div>`}</div>
    <div class="form-panel">
      <div class="form-grid">
        <div><label>${t('fieldDate')}</label><input id="q-date" type="date" value="${todayISO()}"></div>
        <div><label>${t('fieldRef')}</label><input id="q-ref" type="text" placeholder="REF-0001" list="ref-catalog-list"></div>
        <datalist id="ref-catalog-list">${state.catalog.map(c=>`<option value="${escAttr(c.ref)}">`).join('')}</datalist>
        <div><label>${t('fieldDesig')}</label><input id="q-desig" type="text"></div>
        <div><label>${t('fieldFourn')}</label><input id="q-fourn" type="text"></div>
        <div><label>${t('fieldQteControle')}</label><input id="q-qtec" type="number" min="0" placeholder="0"></div>
        <div><label>${t('fieldQteNonOk')}</label><input id="q-qtenok" type="number" min="0" placeholder="0" disabled></div>
        <div><label>${t('fieldPourcentage')}</label><input id="q-pct" type="text" value="0%" disabled></div>
      </div>
      <div class="form-grid" style="margin-top:4px;">
        ${DEFECT_COLS.map(d => `<div><label>${t(d.label)}</label><input id="q-${d.key}" type="number" min="0" placeholder="0"></div>`).join('')}
      </div>
      <div class="form-actions">
        <button class="btn" id="save-btn" style="max-width:200px;">${t('saveBtn')}</button>
        <button class="btn secondary" id="clear-btn">${t('clearBtn')}</button>
      </div>
    </div>`;
  if(!isEmployee) document.getElementById('back').onclick = () => { state.view='qc-menu'; render(); };
  document.getElementById('clear-btn').onclick = () => renderQcEntry(main);
  const refInput = document.getElementById('q-ref');
  refInput.addEventListener('input', () => {
    const match = state.catalog.find(c => c.ref.toLowerCase() === refInput.value.trim().toLowerCase());
    if(match){ document.getElementById('q-desig').value = match.designation; document.getElementById('q-fourn').value = match.fournisseur; }
  });
  const qtecInput = document.getElementById('q-qtec');
  const qtenokInput = document.getElementById('q-qtenok');
  const pctInput = document.getElementById('q-pct');
  const updatePct = () => {
    const qc = Number(qtecInput.value) || 0;
    const nok = Number(qtenokInput.value) || 0;
    pctInput.value = (qc > 0 ? (nok/qc*100) : 0).toFixed(1) + '%';
  };
  const updateNonOk = () => {
    const sum = DEFECT_COLS.reduce((s,d) => s + (Number(document.getElementById('q-'+d.key).value) || 0), 0);
    qtenokInput.value = sum;
    updatePct();
  };
  DEFECT_COLS.forEach(d => { document.getElementById('q-'+d.key).addEventListener('input', updateNonOk); });
  qtecInput.addEventListener('input', updatePct);
  document.getElementById('save-btn').onclick = async () => {
    const ref = document.getElementById('q-ref').value.trim();
    const qteControle = Number(document.getElementById('q-qtec').value) || 0;
    if(!ref || !qteControle){ showToast(t('qcRequiredError'), true); return; }
    const rec = {
      date: document.getElementById('q-date').value || todayISO(),
      fournisseur: document.getElementById('q-fourn').value.trim(),
      ref, designation: document.getElementById('q-desig').value.trim(),
      qteControle, qteNonOk: Number(document.getElementById('q-qtenok').value) || 0,
      user: state.currentUser.name,
    };
    DEFECT_COLS.forEach(d => { rec[d.key] = Number(document.getElementById('q-'+d.key).value) || 0; });
    const btn = document.getElementById('save-btn'); btn.disabled = true;
    const ok = await insertQc(rec);
    btn.disabled = false;
    if(ok){ await loadQcRecords(); showToast(t('qcSavedToast')); renderQcEntry(main); }
  };
}

function renderQcData(main){
  const filtered = state.qcRecords.filter(r => QC_DISPLAY_COLS.filter(c=>c.filterable).every(c => {
    const q = (state.qcSearch[c.key]||'').toLowerCase().trim();
    if(!q) return true;
    return (r[c.key] ?? '').toString().toLowerCase().includes(q);
  }));
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.data.replace('<svg','<svg width="18" height="18"')} ${t('cardQcDataTitle')}</h2>${state.currentUser.role === 'entrepreneur' ? '' : `<div class="back-link" id="back">${ICONS.back} ${t('back')}</div>`}</div>
    <div class="toolbar"><button class="btn blue" id="export-btn" style="width:auto;padding:10px 16px;">${ICONS.download} ${t('exportBtn')}</button></div>
    <div class="table-wrap">
      ${state.qcRecords.length===0 ? `<div class="empty-state">${t('qcNoRecords')}.</div>` : `
      <table>
        <thead>
          <tr>${QC_DISPLAY_COLS.map(c=>`<th>${t(c.label)}</th>`).join('')}<th></th></tr>
          <tr class="filter-row">${QC_DISPLAY_COLS.map(c=>`<th>${c.filterable ? `<input class="qc-col-search" data-key="${c.key}" placeholder="${t(c.label)}" value="${escAttr(state.qcSearch[c.key]||'')}">` : ''}</th>`).join('')}<th></th></tr>
        </thead>
        <tbody>
          ${filtered.length===0 ? `<tr><td colspan="${QC_DISPLAY_COLS.length+1}" class="empty-state">${t('qcNoRecords')}.</td></tr>` : filtered.map(r=>renderQcRow(r)).join('')}
        </tbody>
      </table>`}
    </div>`;
  const qcBackEl = document.getElementById('back');
  if(qcBackEl) qcBackEl.onclick = () => { state.view='qc-menu'; render(); };
  document.getElementById('export-btn').onclick = () => exportQcToExcel(filtered);
  main.querySelectorAll('.qc-col-search').forEach(inp => {
    inp.oninput = () => { state.qcSearch[inp.dataset.key] = inp.value; state.lastQcSearchFocus = inp.dataset.key; renderQcData(main); };
  });
  if(state.lastQcSearchFocus){
    const el = main.querySelector(`.qc-col-search[data-key="${state.lastQcSearchFocus}"]`);
    if(el){ el.focus(); el.selectionStart = el.selectionEnd = el.value.length; }
  }
  filtered.forEach(r => {
    const editBtn = document.getElementById('qedit-'+r.id);
    const delBtn = document.getElementById('qdel-'+r.id);
    if(editBtn) editBtn.onclick = () => { state.editingQcId = (state.editingQcId===r.id?null:r.id); renderQcData(main); };
    if(delBtn) delBtn.onclick = async () => {
      if(confirm(t('qcConfirmDelete'))){
        const ok = await deleteQc(r.id);
        if(ok){ await loadQcRecords(); showToast(t('deletedToast')); renderQcData(main); }
      }
    };
    if(state.editingQcId === r.id){
      const nonOkField = document.getElementById('qe-qteNonOk-'+r.id);
      const updateEditNonOk = () => {
        const sum = DEFECT_COLS.reduce((s,d) => s + (Number(document.getElementById('qe-'+d.key+'-'+r.id).value) || 0), 0);
        if(nonOkField) nonOkField.value = sum;
      };
      DEFECT_COLS.forEach(d => {
        const el = document.getElementById('qe-'+d.key+'-'+r.id);
        if(el) el.addEventListener('input', updateEditNonOk);
      });
      const saveBtn = document.getElementById('qrowsave-'+r.id);
      if(saveBtn) saveBtn.onclick = async () => {
        r.date = document.getElementById('qe-date-'+r.id).value || r.date;
        r.fournisseur = document.getElementById('qe-fournisseur-'+r.id).value.trim();
        r.ref = document.getElementById('qe-ref-'+r.id).value.trim() || r.ref;
        r.designation = document.getElementById('qe-designation-'+r.id).value.trim();
        r.qteControle = Number(document.getElementById('qe-qteControle-'+r.id).value) || 0;
        r.qteNonOk = Number(document.getElementById('qe-qteNonOk-'+r.id).value) || 0;
        DEFECT_COLS.forEach(d => { r[d.key] = Number(document.getElementById('qe-'+d.key+'-'+r.id).value) || 0; });
        const ok = await updateQc(r);
        if(ok){ state.editingQcId = null; showToast(t('editedToast')); renderQcData(main); }
      };
    }
  });
}

function renderQcRow(r){
  const pct = qcPct(r).toFixed(1) + '%';
  if(state.editingQcId === r.id){
    return `<tr>
      <td><input id="qe-date-${r.id}" type="date" value="${r.date}"></td>
      <td><input id="qe-fournisseur-${r.id}" value="${escAttr(r.fournisseur)}"></td>
      <td><input id="qe-ref-${r.id}" value="${escAttr(r.ref)}"></td>
      <td><input id="qe-designation-${r.id}" value="${escAttr(r.designation)}"></td>
      <td><input id="qe-qteControle-${r.id}" type="number" value="${r.qteControle}" style="width:70px"></td>
      <td><input id="qe-qteNonOk-${r.id}" type="number" value="${r.qteNonOk}" style="width:70px" disabled></td>
      <td class="num">${pct}</td>
      ${DEFECT_COLS.map(d => `<td><input id="qe-${d.key}-${r.id}" type="number" value="${r[d.key]||0}" style="width:60px"></td>`).join('')}
      <td>${escHtml(r.user)}</td>
      <td class="row-actions">
        <button class="icon-btn" id="qrowsave-${r.id}" title="${t('saveBtn')}">${ICONS.check}</button>
        <button class="icon-btn del" id="qdel-${r.id}" title="${t('deletedToast')}">${ICONS.del}</button>
      </td>
    </tr>`;
  }
  return `<tr>
    <td class="num">${r.date}</td>
    <td>${escHtml(r.fournisseur)||'—'}</td>
    <td class="ref-cell">${escHtml(r.ref)}</td>
    <td>${escHtml(r.designation)||'—'}</td>
    <td class="num">${r.qteControle}</td>
    <td class="num">${r.qteNonOk}</td>
    <td class="num">${pct}</td>
    ${DEFECT_COLS.map(d => `<td class="num">${r[d.key]||0}</td>`).join('')}
    <td>${escHtml(r.user)}</td>
    <td class="row-actions">
      ${state.currentUser.role === 'entrepreneur' ? '' : `<button class="icon-btn" id="qedit-${r.id}" title="${t('editedToast')}">${ICONS.edit}</button>
      <button class="icon-btn del" id="qdel-${r.id}" title="${t('deletedToast')}">${ICONS.del}</button>`}
    </td>
  </tr>`;
}

let qcChartRefs = [];
function renderQcDashboard(main){
  qcChartRefs.forEach(c => c.destroy()); qcChartRefs = [];
  const recs = state.qcRecords.filter(dashInRange);
  const totalControle = recs.reduce((s,r)=>s+r.qteControle,0);
  const totalNonOk = recs.reduce((s,r)=>s+r.qteNonOk,0);
  const globalRate = totalControle > 0 ? (totalNonOk/totalControle*100) : 0;
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.dash.replace('<svg','<svg width="18" height="18"')} ${t('cardQcDashTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    ${dashBarHtml()}
    <div class="stat-grid">
      <div class="stat-card"><div class="label">${t('qcStatControls')}</div><div class="value">${recs.length}</div></div>
      <div class="stat-card"><div class="label">${t('qcStatControlled')}</div><div class="value blue">${totalControle}</div></div>
      <div class="stat-card"><div class="label">${t('qcStatNonOk')}</div><div class="value">${totalNonOk}</div></div>
      <div class="stat-card"><div class="label">${t('qcStatRate')}</div><div class="value accent">${globalRate.toFixed(1)}%</div></div>
    </div>
    ${recs.length===0 ? `<div class="empty-state" style="background:var(--panel);border:1px solid var(--border);border-radius:var(--radius);">${t('noData')}</div>` : `
    <div class="charts-grid">
      <div class="chart-panel"><h3>${t('qcChartDefectTitle')}</h3><canvas id="qc-chart-defect"></canvas></div>
      <div class="chart-panel"><h3>${t('qcChartRateFournTitle')}</h3><canvas id="qc-chart-rate-fourn"></canvas></div>
      <div class="chart-panel wide"><h3>${t('qcChartRefTitle')}</h3><canvas id="qc-chart-ref"></canvas></div>
    </div>`}`;
  document.getElementById('back').onclick = () => { state.view='qc-menu'; render(); };
  dashBarWire(() => renderQcDashboard(main));
  if(recs.length===0) return;
  const palette = ['#d68c45','#5b8fb0','#5aad8c','#e2574c','#8a93a3','#a687c9','#c9a05b','#6fb0c9'];
  const defectTotals = {};
  DEFECT_COLS.forEach(d => { defectTotals[t(d.label)] = recs.reduce((s,r)=>s+(r[d.key]||0),0); });
  qcChartRefs.push(new Chart(document.getElementById('qc-chart-defect'), { type:'bar', data:{ labels:Object.keys(defectTotals), datasets:[{ data:Object.values(defectTotals), backgroundColor:Object.keys(defectTotals).map((_,i)=>palette[i%palette.length]), borderRadius:2 }] }, options:{ responsive:true, plugins:{legend:{display:false}}, scales:{ x:{ticks:{color:'#8a93a3'},grid:{display:false}}, y:{ticks:{color:'#8a93a3'},grid:{color:'#313949'}} } } }));
  const fournList = [...new Set(recs.map(r=>r.fournisseur || t('undefinedLabel')))];
  const rateByFourn = fournList.map(f => { const rs = recs.filter(r => (r.fournisseur||t('undefinedLabel')) === f); const c = rs.reduce((s,r)=>s+r.qteControle,0); const n = rs.reduce((s,r)=>s+r.qteNonOk,0); return c > 0 ? +(n/c*100).toFixed(1) : 0; });
  qcChartRefs.push(new Chart(document.getElementById('qc-chart-rate-fourn'), { type:'bar', data:{ labels:fournList, datasets:[{ data:rateByFourn, backgroundColor:fournList.map((_,i)=>palette[i%palette.length]), borderRadius:2 }] }, options:{ responsive:true, plugins:{legend:{display:false}}, scales:{ x:{ticks:{color:'#8a93a3'},grid:{display:false}}, y:{ticks:{color:'#8a93a3',callback:v=>v+'%'},grid:{color:'#313949'}} } } }));
  const refList = [...new Set(recs.map(r=>r.ref || t('undefinedLabel')))];
  const controleByRef = refList.map(ref => recs.filter(r=>(r.ref||t('undefinedLabel'))===ref).reduce((s,r)=>s+r.qteControle,0));
  const nonOkByRef = refList.map(ref => recs.filter(r=>(r.ref||t('undefinedLabel'))===ref).reduce((s,r)=>s+r.qteNonOk,0));
  qcChartRefs.push(new Chart(document.getElementById('qc-chart-ref'), { type:'bar', data:{ labels:refList, datasets:[ { label:t('qcStatControlled'), data:controleByRef, backgroundColor:'#5b8fb0', borderRadius:2 }, { label:t('qcStatNonOk'), data:nonOkByRef, backgroundColor:'#e2574c', borderRadius:2 } ]}, options:{ responsive:true, plugins:{legend:{position:'bottom',labels:{color:'#8a93a3',boxWidth:12,font:{size:11}}}}, scales:{ x:{ticks:{color:'#8a93a3'},grid:{display:false}}, y:{ticks:{color:'#8a93a3'},grid:{color:'#313949'}} } } }));
}

function exportQcToExcel(records){
  if(!records.length){ showToast(t('noExportData'), true); return; }
  const rows = records.map(r => {
    const row = { 'Date': r.date, 'Fournisseur': r.fournisseur, 'Ref': r.ref, 'Désignation': r.designation, 'Qtité contrôle': r.qteControle, 'Quantité non OK': r.qteNonOk, 'Pourcentage': qcPct(r).toFixed(1)+'%' };
    DEFECT_COLS.forEach(d => { row[t(d.label)] = r[d.key]||0; });
    row['User'] = r.user;
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Controle');
  XLSX.writeFile(wb, `controle_qualite_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast(t('exportedToast'));
}

/* ================= USERS ================= */
async function renderUsers(main){
  await loadProfiles();
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.users.replace('<svg','<svg width="18" height="18"')} ${t('usersTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="form-panel" style="margin-bottom:16px;">
      <div id="add-user-error"></div>
      <div class="user-form">
        <div><label>${t('newUsername')}</label><input id="u-username" type="text"></div>
        <div><label>${t('newName')}</label><input id="u-name" type="text"></div>
        <div><label>${t('newPass')}</label><input id="u-pass" type="text"></div>
        <div>
          <label>${t('colRole')}</label>
          <select id="u-role" style="margin-bottom:16px;">
            <option value="user">${t('roleUser')}</option>
            <option value="employee">${t('roleEmployee')}</option>
            <option value="visitor">${t('roleVisitor')}</option>
            <option value="entrepreneur">${t('roleEntrepreneur')}</option>
            <option value="admin">${t('roleAdmin')}</option>
          </select>
        </div>
      </div>
      <button class="btn" id="add-user-btn" style="max-width:220px;">${t('addUserBtn')}</button>
    </div>
    <div class="banner">${t('usersBanner')}</div>
    <div class="table-wrap">
      <table style="min-width:0;">
        <thead><tr><th>${t('colUsername')}</th><th>${t('colName')}</th><th>${t('colRole')}</th></tr></thead>
        <tbody>
          ${state.profiles.map(p => `
            <tr>
              <td class="ref-cell">${escHtml(p.username)}</td>
              <td>${escHtml(p.name)}</td>
              <td>
                <select class="role-select" data-id="${p.id}" ${p.id===state.currentUser.id?'disabled':''} style="margin:0;padding:6px 8px;width:auto;">
                  <option value="user" ${p.role==='user'?'selected':''}>${t('roleUser')}</option>
                  <option value="employee" ${p.role==='employee'?'selected':''}>${t('roleEmployee')}</option>
                  <option value="visitor" ${p.role==='visitor'?'selected':''}>${t('roleVisitor')}</option>
                  <option value="entrepreneur" ${p.role==='entrepreneur'?'selected':''}>${t('roleEntrepreneur')}</option>
                  <option value="admin" ${p.role==='admin'?'selected':''}>${t('roleAdmin')}</option>
                </select>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
  document.getElementById('back').onclick = () => { state.view='menu'; render(); };
  main.querySelectorAll('.role-select').forEach(sel => {
    sel.onchange = async () => { const ok = await updateRole(sel.dataset.id, sel.value); if(ok) showToast(t('roleUpdatedToast')); };
  });
  document.getElementById('add-user-btn').onclick = async () => {
    const username = document.getElementById('u-username').value.trim();
    const name = document.getElementById('u-name').value.trim();
    const password = document.getElementById('u-pass').value.trim();
    const role = document.getElementById('u-role').value;
    const errBox = document.getElementById('add-user-error');
    errBox.innerHTML = '';
    if(!username || !name || !password){ showToast(t('userFieldsErr'), true); return; }
    if(password.length < 9){ showToast(t('userPassShortErr'), true); return; }
    const btn = document.getElementById('add-user-btn'); btn.disabled = true;
    const result = await createUserApi({ username, name, password, role });
    btn.disabled = false;
    if(result.ok){
      showToast(t('userAddedToast'));
      renderUsers(main);
    }else if(result.error && (result.error.includes('already been registered') || result.error.includes('duplicate'))){
      showToast(t('userExistsErr'), true);
    }else{
      showToast((result.error || t('userAddErrGeneric')) + '', true);
    }
  };
}

/* ================= RECIPIENTS (rapport auto par e-mail) ================= */
function renderRecipients(main){
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.mail.replace('<svg','<svg width="18" height="18"')} ${t('recipientsTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="banner">${t('recipientsBanner')}</div>
    <div class="form-panel" style="margin-bottom:16px;">
      <div class="user-form" style="grid-template-columns:1fr auto;">
        <div><label>${t('newRecipientEmail')}</label><input id="r-email" type="email" placeholder="nom@exemple.com"></div>
      </div>
      <button class="btn" id="add-recipient-btn" style="max-width:180px;">${t('addRecipientBtn')}</button>
      <button class="btn" id="send-test-report-btn" style="max-width:280px;margin-inline-start:8px;">${t('sendTestReportBtn')}</button>
    </div>
    <div class="table-wrap">
      ${state.recipients.length===0 ? `<div class="empty-state">${t('noRecipients')}.</div>` : `
      <table style="min-width:0;">
        <thead><tr><th>${t('newRecipientEmail')}</th><th></th></tr></thead>
        <tbody>
          ${state.recipients.map(r => `
            <tr>
              <td>${escHtml(r.email)}</td>
              <td class="row-actions"><button class="icon-btn del" id="rdel-${r.id}" title="${t('deletedToast')}">${ICONS.del}</button></td>
            </tr>`).join('')}
        </tbody>
      </table>`}
    </div>`;
  document.getElementById('back').onclick = () => { state.view='menu'; render(); };
  document.getElementById('add-recipient-btn').onclick = async () => {
    const email = document.getElementById('r-email').value.trim();
    if(!email || !email.includes('@')){ showToast(t('recipientFieldErr'), true); return; }
    const ok = await insertRecipient(email);
    if(ok){ await loadRecipients(); showToast(t('recipientAddedToast')); renderRecipients(main); }
  };
  document.getElementById('send-test-report-btn').onclick = () => sendQualityReportTest();
  state.recipients.forEach(r => {
    const btn = document.getElementById('rdel-'+r.id);
    if(!btn) return;
    btn.onclick = async () => {
      if(confirm(t('confirmDeleteRecipient'))){
        const ok = await deleteRecipient(r.id);
        if(ok){ await loadRecipients(); showToast(t('recipientDeletedToast')); renderRecipients(main); }
      }
    };
  });
}

/* ================= PDF LIBRARY ================= */
function fmtSize(bytes){
  if(!bytes) return '';
  if(bytes < 1024*1024) return (bytes/1024).toFixed(0)+' KB';
  return (bytes/(1024*1024)).toFixed(1)+' MB';
}
function renderPdfLibrary(main){
  const readOnly = !!(state.currentUser && state.currentUser.role === 'visitor');
  if(readOnly) state.currentLibrary = 'pdf';
  const $id = (id) => document.getElementById(id) || { set onclick(v){} };
  const libTitle = state.currentLibrary === 'bonnequalite' ? t('bqLibraryTitle') : t('pdfLibraryTitle');
  const atRoot = !state.pdfFolder;
  if(atRoot){
    const folders = state.pdfFiles.filter(f => f.id === null);
    main.innerHTML = `
      <div class="page-head"><h2>${ICONS.pdf.replace('<svg','<svg width="18" height="18"')} ${libTitle}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
      <div class="form-panel" style="margin-bottom:16px;">
        <label>${t('fieldFourn')}</label>
        <input id="pdf-fourn" type="text" list="fourn-suggest" placeholder="${t('fieldFourn')}">
        <datalist id="fourn-suggest">${folders.map(f=>`<option value="${escAttr(f.name)}">`).join('')}${state.catalog.map(c=>`<option value="${escAttr(c.fournisseur)}">`).join('')}</datalist>
        <label>${t('uploadPdfBtn')}</label>
        <input id="pdf-file" type="file" accept="application/pdf,.pdf">
        <button class="btn" id="upload-pdf-btn" style="max-width:220px;">${t('uploadPdfBtn')}</button>
      </div>
      ${folders.length===0 ? `<div class="empty-state">${t('noPdf')}.</div>` : `
      <div class="tile-grid">
        ${folders.map(f => `<div class="tile" id="folder-${escAttr(f.name)}"><span class="tile-icon">📁</span>${escHtml(f.name.replace(/_/g,' '))}</div>`).join('')}
      </div>`}`;
    if(readOnly) main.querySelectorAll('.form-panel, #back').forEach(el => el.remove());
    $id('back').onclick = () => { state.view='menu'; render(); };
    $id('upload-pdf-btn').onclick = async () => {
      const fourn = document.getElementById('pdf-fourn').value.trim();
      const input = document.getElementById('pdf-file');
      const file = input.files && input.files[0];
      if(!fourn){ showToast(t('recipientFieldErr'), true); return; }
      if(!file || file.type !== 'application/pdf'){ showToast(t('pdfSelectErr'), true); return; }
      const btn = document.getElementById('upload-pdf-btn'); btn.disabled = true;
      const ok = await uploadPdfFile(slugFolder(fourn), file);
      btn.disabled = false;
      if(ok){
        state.pdfFolder = slugFolder(fourn);
        await listPdfEntries(state.pdfFolder);
        showToast(t('pdfUploadedToast'));
        renderPdfLibrary(main);
      }
    };
    folders.forEach(f => {
      const el = document.getElementById('folder-'+f.name);
      if(el) el.onclick = async () => { state.pdfFolder = f.name; await listPdfEntries(f.name); renderPdfLibrary(main); };
    });
    return;
  }

  const files = state.pdfFiles.filter(f => f.id !== null);
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.pdf.replace('<svg','<svg width="18" height="18"')} ${escHtml(state.pdfFolder.replace(/_/g,' '))}</h2><div class="back-link" id="back">${ICONS.back} ${libTitle}</div></div>
    <div class="form-panel" style="margin-bottom:16px;">
      <label>${t('uploadPdfBtn')}</label>
      <input id="pdf-file" type="file" accept="application/pdf,.pdf">
      <button class="btn" id="upload-pdf-btn" style="max-width:220px;">${t('uploadPdfBtn')}</button>
    </div>
    ${files.length===0 ? `<div class="empty-state">${t('noPdf')}.</div>` : `
    <div class="tile-grid">
      ${files.map(f => `
        <div class="tile" id="pview-${escAttr(f.name)}">
          <button class="tile-del" id="pdel-${escAttr(f.name)}" title="${t('deletedToast')}">${ICONS.del}</button>
          <span class="tile-icon">📄</span>${escHtml(f.name.replace(/^\d+_/,'').replace(/\.pdf$/i,''))}
          ${f.metadata ? `<div style="color:var(--text-muted);font-weight:400;font-size:11px;margin-top:4px;">${fmtSize(f.metadata.size)}</div>` : ''}
        </div>`).join('')}
    </div>`}`;
  if(readOnly) main.querySelectorAll('.form-panel, .tile-del').forEach(el => el.remove());
  $id('back').onclick = async () => { state.pdfFolder=''; await listPdfEntries(''); renderPdfLibrary(main); };
  $id('upload-pdf-btn').onclick = async () => {
    const input = document.getElementById('pdf-file');
    const file = input.files && input.files[0];
    if(!file || file.type !== 'application/pdf'){ showToast(t('pdfSelectErr'), true); return; }
    const btn = document.getElementById('upload-pdf-btn'); btn.disabled = true;
    const ok = await uploadPdfFile(state.pdfFolder, file);
    btn.disabled = false;
    if(ok){ await listPdfEntries(state.pdfFolder); showToast(t('pdfUploadedToast')); renderPdfLibrary(main); }
  };
  files.forEach(f => {
    const fullPath = `${state.pdfFolder}/${f.name}`;
    const tile = document.getElementById('pview-'+f.name);
    const delBtn = document.getElementById('pdel-'+f.name);
    if(tile) tile.onclick = async (e) => {
      if(e.target.closest('.tile-del')) return;
      const url = await getPdfSignedUrl(fullPath);
      if(url) window.open(url, '_blank');
    };
    if(delBtn) delBtn.onclick = async (e) => {
      e.stopPropagation();
      if(confirm(t('confirmDeletePdf'))){
        const ok = await deletePdfFile(fullPath);
        if(ok){ await listPdfEntries(state.pdfFolder); showToast(t('pdfDeletedToast')); renderPdfLibrary(main); }
      }
    };
  });
}

/* ================= DÉBIT NOTE MODULE ================= */
/* ================= AUTOLIV-STYLE DÉBIT NOTE (exact template) ================= */
const AUTOLIV_DOC_CSS = `
.autoliv-doc-scroll{ overflow-x:auto; -webkit-overflow-scrolling:touch; max-width:100%; }
.autoliv-doc{background:#d8d8d8;font-family:Arial,Helvetica,sans-serif;color:#111;font-size:7px;padding:10px 0;}
.autoliv-doc *{box-sizing:border-box}
.autoliv-doc button{font:inherit;cursor:pointer}
.autoliv-doc .exportBtn{margin:0;width:100%;height:6.2mm;padding:0 1mm;border:1px solid #173f91;background:#173f91;color:#fff;font-weight:bold;font-size:6.5px;white-space:nowrap;}
.autoliv-doc .toolbar{position:sticky;top:0;z-index:20;display:flex;gap:4px;justify-content:center;background:#eee;padding:6px;margin-bottom:8px;}
.autoliv-doc .toolbar button{padding:6px 10px;border:1px solid #777;background:#fff;font-size:11px}
.autoliv-doc .toolbar .transfer{background:#eee04a;border-color:#b9aa18;font-weight:bold}
.autoliv-doc .page{width:210mm;min-height:297mm;margin:10px auto;background:#fff;padding:8mm 9mm 5mm;box-shadow:0 1px 8px #888;overflow:hidden;position:relative;}
.autoliv-doc .logo{width:34mm;margin-top:1mm}
.autoliv-doc .logoText{font-size:21px;font-weight:bold;color:#234f91;letter-spacing:-1.5px;line-height:20px}
.autoliv-doc .logoLine{height:2.8mm;background:#234f91;margin-top:1.5mm}
.autoliv-doc .topArea{height:57mm;position:relative}
.autoliv-doc .owner{position:absolute;left:0;top:30mm;width:62mm;display:grid;grid-template-columns:28mm 34mm}
.autoliv-doc .owner div{height:5mm;border:1px solid #777;padding:1mm;font-size:7px}
.autoliv-doc .owner .label{background:#173f91;color:#fff;font-weight:bold}
.autoliv-doc .docbox{position:absolute;right:0;top:0;width:62mm}
.autoliv-doc .dochead{display:grid;grid-template-columns:22mm 16mm 24mm;align-items:center}
.autoliv-doc .doclabel{height:6.2mm;background:#173f91;color:#fff;border:1px solid #173f91;padding:1.4mm 1.5mm;font-size:7px;font-weight:bold}
.autoliv-doc .docno{height:6.2mm;border:1px solid #777;text-align:center;font-weight:bold;font-size:10px;background:#fafafa}
.autoliv-doc .docinfo{position:absolute;left:66mm;top:0;width:60mm}
.autoliv-doc .docinfo table{width:100%;border-collapse:collapse}
.autoliv-doc .docinfo td{height:4.15mm;border:1px solid #777;padding:0.5mm 1mm;font-size:6.3px}
.autoliv-doc .docinfo td:first-child{border:0;font-weight:bold;width:28mm}
.autoliv-doc .docinfo td:last-child{text-align:center;width:32mm}
.autoliv-doc input,.autoliv-doc textarea,.autoliv-doc select{width:100%;min-width:0;height:100%;border:0;outline:none;background:transparent;padding:0;font:inherit;color:#111;}
.autoliv-doc input:focus,.autoliv-doc textarea:focus{background:#fff8a8}
.autoliv-doc input:disabled{color:#173f91;font-weight:bold;background:transparent;}
.autoliv-doc .bar{height:5mm;background:#173f91;color:#fff;font-weight:bold;font-size:7.2px;padding:1.2mm 1.5mm;display:flex;align-items:center}
.autoliv-doc .bar .obj{margin-left:75mm}
.autoliv-doc .section{border:1px solid #777;border-top:0}
.autoliv-doc table{border-collapse:collapse;width:100%;table-layout:fixed}
.autoliv-doc th,.autoliv-doc td{border:1px solid #777;padding:0.45mm 0.8mm;height:4.3mm;font-size:6.1px;vertical-align:middle}
.autoliv-doc th{font-weight:normal}
.autoliv-doc .num{text-align:right}
.autoliv-doc .totalCell{text-align:right;font-weight:bold}
.autoliv-doc .prodArea{height:40mm;position:relative}
.autoliv-doc .prodTable{position:absolute;right:0;top:9mm;width:101mm}
.autoliv-doc .prodTable .lineTitle{position:absolute;left:0;top:-6mm;width:25mm;height:5mm;border:1px solid #777;text-align:center}
.autoliv-doc .prodTable .comment{position:absolute;left:27mm;top:-5.4mm;font-size:6px}
.autoliv-doc .logArea{height:32mm;position:relative}
.autoliv-doc .logTable{position:absolute;left:0;top:2mm;width:82mm}
.autoliv-doc .logRightLabel{position:absolute;left:84mm;top:1mm;font-size:6px}
.autoliv-doc .qualityArea{height:41mm;position:relative}
.autoliv-doc .qualityLeft{position:absolute;left:0;top:2mm;width:88mm}
.autoliv-doc .qualityRight{position:absolute;right:0;top:2mm;width:101mm}
.autoliv-doc .yellowBar{position:absolute;right:0;top:-5mm;width:80mm;height:5mm;background:#eee04a;border:1px solid #777;text-align:center;font-size:6px;padding:0 1mm}
.autoliv-doc .qualitySub{height:5mm;display:flex;align-items:center}
.autoliv-doc .adminArea{height:40mm;position:relative}
.autoliv-doc .ncmBox{position:absolute;left:0;top:3mm;width:50mm}
.autoliv-doc .ncmRow{display:flex;align-items:center;gap:4mm;margin-bottom:8mm}
.autoliv-doc .ncmRow label{font-weight:bold;width:10mm}
.autoliv-doc .ncmInput{width:18mm;height:5mm;border:1px solid #777;text-align:center}
.autoliv-doc .costLabel{margin-top:1mm}
.autoliv-doc .travel{position:absolute;right:0;top:2mm;width:83mm}
.autoliv-doc .travelTitle{text-align:center;height:5mm}
.autoliv-doc .materialNote{position:absolute;right:30mm;bottom:2mm;font-size:6px}
.autoliv-doc .smallBox{position:absolute;right:0;bottom:1mm;width:13mm;height:5mm;border:1px solid #777}
.autoliv-doc .approval td,.autoliv-doc .approval th{height:5.2mm}
.autoliv-doc .approval .role{width:33mm}
.autoliv-doc .approval .when{width:42mm}
.autoliv-doc .approval .comments{width:auto}
.autoliv-doc .approval .amount{width:25mm;text-align:center}
.autoliv-doc .finalAmount{font-size:9px;font-weight:bold;text-align:center}
.autoliv-doc .footer{font-size:5.7px;text-align:center;margin-top:2mm}

/* --- reset: empêche les styles globaux de l'application de déformer la feuille --- */
.autoliv-doc table{min-width:0;font-size:6.1px;table-layout:fixed;width:100%;border-collapse:collapse}
.autoliv-doc th{background:transparent;color:#111;text-transform:none;letter-spacing:0;text-align:left;font-weight:normal;white-space:normal;font-size:6.1px;font-family:inherit}
.autoliv-doc td{white-space:normal;font-family:inherit}
.autoliv-doc tr:last-child td{border-bottom:1px solid #777}
.autoliv-doc input,.autoliv-doc select,.autoliv-doc textarea{margin:0;border-radius:0;opacity:1;font-family:Arial,Helvetica,sans-serif}
.autoliv-doc input:disabled{opacity:1}
.autoliv-doc label{display:inline;margin:0;font-size:7px;color:#111}
.autoliv-doc .num{font-family:Arial,Helvetica,sans-serif}
.autoliv-doc input[type=date]:invalid:not(:focus){color:transparent}
.autoliv-doc input[type=date]::-webkit-calendar-picker-indicator{opacity:.35;width:6px;height:6px;padding:0;margin:0}
.autoliv-doc .ncmInput{border:1px solid #777}
`;

/* The sheet is designed in tiny units (6-7px fonts). Browsers with a larger "minimum font size" (many PCs)
   inflate those fonts and the sheet overlaps itself. So the sheet is laid out 3x bigger (fonts >= 17px, which no
   browser setting inflates) and shrunk back with CSS zoom: same look, immune to the minimum font size. */
const DEBIT_SCALE = 3;
function debitScaleLen(str){
  return str.replace(/(-?\d*\.?\d+)(mm|px)/g, (m, n, u) => (Math.round(parseFloat(n) * DEBIT_SCALE * 1000) / 1000) + u);
}
function debitCssScaled(){
  return debitScaleLen(AUTOLIV_DOC_CSS)
    .replace('padding:' + (10 * DEBIT_SCALE) + 'px 0;', 'padding:10px 0;')
    + `\n.autoliv-doc .page{zoom:${(1 / DEBIT_SCALE).toFixed(6)};}`;
}

function debitMoney(n){ return (Number(n)||0).toLocaleString('fr-FR',{minimumFractionDigits:2,maximumFractionDigits:2})+' \u20ac'; }
function debitV(e){ return parseFloat(e && e.value) || 0; }

function autolivDocMarkup(){
  return `
<div class="page" id="debit-form-capture">
    <div class="logo"><div class="logoText">Autoliv</div><div class="logoLine"></div></div>
    <div class="topArea">
      <div class="owner">
        <div class="label">${t('debitReportBy')}:</div><div><input id="a-reportby" disabled></div>
        <div class="label">${t('debitAswt')}:</div><div><input id="a-aswt"></div>
      </div>
      <div class="docbox">
        <div class="dochead">
          <div class="doclabel">DOCUMENT N\u00b0</div>
          <input id="a-docno" class="docno" value="\u2014" disabled>
        </div>
      </div>
      <div class="docinfo">
        <table>
          <tr><td>Date</td><td><input id="a-date" type="date"></td></tr>
          <tr><td>SUPPLIER</td><td><input id="a-supplier" list="debit-fourn-list"></td></tr>
          <tr><td>Supplier Contact Information</td><td></td></tr>
          <tr><td>Name</td><td><input id="a-contact-name"></td></tr>
          <tr><td>Mail</td><td><input id="a-contact-mail"></td></tr>
          <tr><td>Phone</td><td><input id="a-contact-phone"></td></tr>
          <tr><td>Department</td><td><input id="a-contact-dept"></td></tr>
          <tr><td>Part number</td><td><input id="a-hdr-partnum"></td></tr>
          <tr><td>Debit note N\u00b0</td><td><input id="a-hdr-debitno"></td></tr>
          <tr><td>Part Name</td><td><input id="a-hdr-partname"></td></tr>
          <tr><td>Cause</td><td><input id="a-cause"></td></tr>
        </table>
      </div>
    </div>

    <div class="bar">PRODUCTION COST <span class="obj">CB OBJECT:</span></div>
    <div class="prodArea">
      <div class="prodTable section">
        <div class="lineTitle">001</div>
        <div class="comment">Line concerned</div>
        <table>
          <tr><th style="width:14mm">Date</th><th>Line concerned</th><th style="width:10mm">Hours</th><th style="width:10mm">Rate</th><th style="width:12mm">Total</th></tr>
          <tbody id="a-prod"></tbody>
          <tr><td colspan="4" class="totalCell">Total machine cost</td><td id="a-prodTotal" class="num">0.00 \u20ac</td></tr>
        </table>
      </div>
    </div>

    <div class="bar">LOGISTIC COST <span class="obj">CB OBJECT:</span></div>
    <div class="logArea">
      <div class="logTable section">
        <table>
          <tr><th style="width:14mm">Date</th><th style="width:30mm">Carrier</th><th>Freight</th><th style="width:22mm">Cost</th></tr>
          <tbody id="a-log"></tbody>
          <tr><td colspan="3" class="totalCell">Total</td><td id="a-logTotal" class="num">0.00 \u20ac</td></tr>
        </table>
      </div>
      <div class="logRightLabel">EUROS/Hour</div>
    </div>

    <div class="bar">QUALITY COST <span class="obj">CB OBJECT:</span></div>
    <div class="qualityArea">
      <div class="qualityLeft section">
        <div class="qualitySub"><b style="margin-left:10mm">10) Scrap/rework</b><span style="margin-left:auto;margin-right:5mm">EUROS/Hour</span></div>
        <table>
          <tr><th style="width:27mm">Part number</th><th>Part name</th><th style="width:13mm">Hours</th><th style="width:13mm">Total</th></tr>
          <tbody id="a-quality"></tbody>
          <tr><td colspan="3" class="totalCell">Total</td><td id="a-qualityTotal" class="num">0.00 \u20ac</td></tr>
        </table>
      </div>
      <div class="yellowBar"><input id="a-yellow" placeholder=""></div>
      <div class="qualityRight section">
        <table>
          <tr><th style="width:20mm">Part nbr</th><th>Part name</th><th style="width:12mm">Quantity</th><th style="width:13mm">Unit cost</th><th style="width:13mm">Total</th></tr>
          <tbody id="a-material"></tbody>
          <tr><td colspan="4" class="totalCell">Total material</td><td id="a-materialTotal" class="num">0.00 \u20ac</td></tr>
        </table>
      </div>
    </div>

    <div class="bar">ADMINISTRATION COST <span class="obj">CB OBJECT:</span></div>
    <div class="adminArea">
      <div class="ncmBox">
        <div class="ncmRow"><label>NCM</label><input id="a-ncm" class="ncmInput" type="number" value="0"></div>
        <div class="costLabel">Cost</div>
        <div style="margin-left:23mm;margin-top:-2mm;font-size:7px" id="a-adminDisplay">0</div>
      </div>
      <div class="travel section">
        <div class="travelTitle">Travel expenses</div>
        <table>
          <tr><th style="width:14mm">Date</th><th>Person</th><th style="width:17mm">Cost</th></tr>
          <tbody id="a-travel"></tbody>
          <tr><td colspan="2" class="totalCell">Total</td><td id="a-travelTotal" class="num">0.00 \u20ac</td></tr>
        </table>
      </div>
      <div class="materialNote">Material / rework / SCR from customer</div>
      <div class="smallBox"></div>
    </div>

    <div class="bar">APPROVAL</div>
    <div class="approvalArea section">
      <table class="approval">
        <tr><th class="role">APPROVAL</th><th class="when">WHEN</th><th class="comments">COMMENTS IF REFUSED</th><th class="amount">TOTAL</th></tr>
        <tr><td>PLANT MANAGER / QUALITY MANAGER</td><td><input type="date"></td><td><input></td><td rowspan="1" class="finalAmount" id="a-grandTotal">0.00 \u20ac</td></tr>
        <tr><td>LOGISTIC MANAGER</td><td><input type="date"></td><td><input></td><td rowspan="4" style="vertical-align:top;padding-top:2mm;font-weight:bold">INVOICE N\u00b0</td></tr>
        <tr><td>COMMODITY BUYER</td><td><input type="date"></td><td><input></td></tr>
        <tr><td>FINANCES CONTROLLER</td><td><input type="date"></td><td><input></td></tr>
        <tr><td>GENERAL MANAGER</td><td><input type="date"></td><td><input></td></tr>
      </table>
    </div>
    <div class="footer">In case of refusal, please add the justification of the refusal on the present document to the issuer.</div>
  </div>`;
}

function debitMarkDates(){ document.querySelectorAll('.autoliv-doc input[type=date]').forEach(i=>{ i.required = true; }); }
function debitAddRows(id, n, type, main){
  const b = document.getElementById(id);
  for(let i=0;i<n;i++){
    const r = document.createElement('tr');
    if(type==='p') r.innerHTML = '<td><input type="date"></td><td><input></td><td><input type="number" step=".001" class="a-recalc"></td><td><input type="number" step=".001" class="a-recalc"></td><td class="num rowTotal">0.00 \u20ac</td>';
    if(type==='l') r.innerHTML = '<td><input type="date"></td><td><input></td><td><input></td><td><input type="number" step=".001" class="a-recalc"></td>';
    if(type==='q') r.innerHTML = '<td><input class="a-quality-ref" list="debit-ref-list"></td><td><input></td><td><input type="number" step=".001" class="a-recalc"></td><td class="num qTotal">0.00 \u20ac</td>';
    if(type==='t') r.innerHTML = '<td><input type="date"></td><td><input></td><td><input type="number" step=".001" class="a-recalc"></td>';
    if(type==='m') r.innerHTML = '<td><input class="a-material-ref" list="debit-ref-list"></td><td><input></td><td><input type="number" step=".001" class="a-recalc"></td><td><input type="number" step=".001" class="a-recalc"></td><td class="num">0.00 \u20ac</td>';
    b.appendChild(r);
  }
}

function debitCalc(){
  let p=0;
  document.querySelectorAll('#a-prod tr').forEach(r=>{
    const x = debitV(r.cells[2] && r.cells[2].querySelector('input')) * debitV(r.cells[3] && r.cells[3].querySelector('input'));
    p+=x; const c=r.querySelector('.rowTotal'); if(c) c.textContent=debitMoney(x);
  });
  let l=0; document.querySelectorAll('#a-log input[type=number]').forEach(x=>l+=debitV(x));
  let q=0;
  document.querySelectorAll('#a-quality tr').forEach(r=>{
    const x = debitV(r.cells[2] && r.cells[2].querySelector('input'));
    q+=x; const c=r.querySelector('.qTotal'); if(c) c.textContent=debitMoney(x);
  });
  let material=0;
  document.querySelectorAll('#a-material tr').forEach(r=>{
    const qty = debitV(r.cells[2] && r.cells[2].querySelector('input'));
    const cost = debitV(r.cells[3] && r.cells[3].querySelector('input'));
    const x = qty*cost; material+=x;
    const cell = r.cells[4]; if(cell) cell.textContent=debitMoney(x);
  });
  let tr=0; document.querySelectorAll('#a-travel input[type=number]').forEach(x=>tr+=debitV(x));
  const a=0;
  const total = p+l+q+material+tr+a;
  document.getElementById('a-prodTotal').textContent = debitMoney(p);
  document.getElementById('a-logTotal').textContent = debitMoney(l);
  document.getElementById('a-qualityTotal').textContent = debitMoney(q);
  document.getElementById('a-materialTotal').textContent = debitMoney(material);
  document.getElementById('a-travelTotal').textContent = debitMoney(tr);
  document.getElementById('a-adminDisplay').textContent = a.toFixed(0);
  document.getElementById('a-grandTotal').textContent = debitMoney(total);
  return { p, l, q, material, travel:tr, admin:a, total };
}

function collectDebitFormData(root){
  if(!root) return null;
  const fields = Array.from(root.querySelectorAll('input, textarea, select')).map((el, index) => ({
    index,
    tag: el.tagName,
    type: el.type || '',
    value: el.value ?? '',
    checked: !!el.checked,
  }));
  const editable = Array.from(root.querySelectorAll('[contenteditable="true"]')).map((el, index) => ({ index, html: el.innerHTML }));
  return { version: 1, fields, editable };
}

function applyDebitFormData(root, data){
  if(!root || !data) return;
  const fields = Array.from(root.querySelectorAll('input, textarea, select'));
  (data.fields || []).forEach(item => {
    const el = fields[item.index];
    if(!el) return;
    if(el.type === 'checkbox' || el.type === 'radio') el.checked = !!item.checked;
    else el.value = item.value ?? '';
  });
  const editable = Array.from(root.querySelectorAll('[contenteditable="true"]'));
  (data.editable || []).forEach(item => {
    if(editable[item.index]) editable[item.index].innerHTML = item.html || '';
  });
}

function collectDebitFormDataFromDocument(){
  return collectDebitFormData(document.getElementById('debit-form-capture'));
}

function renderDebitEntry(main, savedFormData=null){
  state.debitPending = null;
  main.innerHTML = `
    <div class="page-head">
      <h2>${ICONS.debit.replace('<svg','<svg width="18" height="18"')} ${t('debitNewTitle')}</h2>
      <div class="back-link" id="back">${ICONS.back} ${t('back')}</div>
    </div>
    <style>${debitCssScaled()}</style>
    <datalist id="debit-fourn-list">${state.catalog.map(c=>`<option value="${escAttr(c.fournisseur)}">`).join('')}</datalist>
    <datalist id="debit-ref-list">${state.catalog.map(c=>`<option value="${escAttr(c.ref)}">`).join('')}</datalist>
    <div style="display:flex;flex-wrap:wrap;gap:8px;margin:0 0 12px;">
      <button class="btn" id="a-save-btn" style="width:auto;">${t('debitSaveBtn')}</button>
      <button class="btn" id="a-pdf-btn" style="width:auto;">&#128196; ${t('debitPdfBtn')}</button>
    </div>
    <div class="autoliv-doc-scroll">
      <div class="autoliv-doc">${debitScaleLen(autolivDocMarkup())}</div>
    </div>
  `;
  document.getElementById('back').onclick = () => { state.view='debit-menu'; render(); };

  debitAddRows('a-prod', 5, 'p', main);
  debitAddRows('a-log', 4, 'l', main);
  debitAddRows('a-quality', 5, 'q', main);
  debitAddRows('a-material', 5, 'm', main);
  debitAddRows('a-travel', 4, 't', main);

  debitMarkDates();
  document.getElementById('a-date').value = todayISO();
  document.getElementById('a-reportby').value = state.currentUser.name;
  document.getElementById('a-aswt').value = '';

  main.addEventListener('input', (e) => {
    if(e.target.classList.contains('a-recalc')) debitCalc();
  });

  main.querySelectorAll('.a-quality-ref').forEach(inp => {
    inp.addEventListener('input', () => {
      const match = state.catalog.find(c => c.ref.toLowerCase() === inp.value.trim().toLowerCase());
      if(match){
        const nameInput = inp.closest('tr').cells[1].querySelector('input');
        if(nameInput) nameInput.value = match.designation;
      }
    });
  });
  main.querySelectorAll('.a-material-ref').forEach(inp => {
    inp.addEventListener('input', () => {
      const match = state.catalog.find(c => c.ref.toLowerCase() === inp.value.trim().toLowerCase());
      if(match){
        const tr = inp.closest('tr');
        const nameInput = tr.cells[1].querySelector('input');
        const costInput = tr.cells[3].querySelector('input');
        if(nameInput) nameInput.value = match.designation;
        if(costInput){ costInput.value = Number(match.prix)||0; }
        debitCalc();
      }
    });
  });

  document.getElementById('a-supplier').addEventListener('input', (e) => {
    const match = state.catalog.find(c => c.fournisseur.toLowerCase() === e.target.value.trim().toLowerCase());
    if(match) document.getElementById('a-hdr-partnum').value = document.getElementById('a-hdr-partnum').value || '';
  });

  if(savedFormData) applyDebitFormData(document.getElementById('debit-form-capture'), savedFormData);
  debitCalc();

  document.getElementById('a-save-btn').onclick = () => debitSaveToHistory(main);
  document.getElementById('a-pdf-btn').onclick = () => debitExportPdf(main);
}

function pdfBusyOverlay(text){
  const busy = document.createElement('div');
  busy.style.cssText = 'position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;color:#fff;font:600 15px sans-serif;text-align:center;padding:20px;';
  busy.textContent = text;
  document.body.appendChild(busy);
  return busy;
}

function debitSheetRecord(){
  const supplier = document.getElementById('a-supplier').value.trim();
  const cause = document.getElementById('a-cause').value.trim();
  const totals = debitCalc();
  const rec = {
    note_date: document.getElementById('a-date').value,
    aswt_department: document.getElementById('a-aswt').value.trim(),
    fournisseur: supplier,
    supplier_name: document.getElementById('a-contact-name').value.trim(),
    supplier_mail: document.getElementById('a-contact-mail').value.trim(),
    supplier_phone: document.getElementById('a-contact-phone').value.trim(),
    supplier_department: document.getElementById('a-contact-dept').value.trim(),
    ref: document.getElementById('a-hdr-partnum').value.trim(),
    part_name: document.getElementById('a-hdr-partname').value.trim(),
    cause,
    report_raised_by: state.currentUser.name,
    production_total: totals.p,
    logistic_total: totals.l,
    quality_total: totals.q + totals.material,
    admin_total: totals.travel + totals.admin,
    grand_total: totals.total,
    user_name: state.currentUser.name,
    form_data: collectDebitFormDataFromDocument(),
  };
  return { supplier, cause, totals, rec };
}

function debitShowNumber(no){
  ['a-docno','a-hdr-debitno'].forEach(id => {
    const f = document.getElementById(id);
    if(!f) return;
    f.value = no;
    f.setAttribute('value', no);
    f.dispatchEvent(new Event('input', { bubbles:true }));
  });
}

/* Button 1: gives the sheet its number and saves the data in Historique. The sheet stays on screen. */
async function debitSaveToHistory(main){
  const btn = document.getElementById('a-save-btn');
  const supplier = document.getElementById('a-supplier').value.trim();
  const cause = document.getElementById('a-cause').value.trim();
  if(!supplier || !cause){ showToast(t('debitRequiredErr'), true); return; }
  if(btn) btn.disabled = true;
  try{
    const { rec } = debitSheetRecord();
    let saved = null;
    if(state.debitPending){
      const { data: upd, error: updErr } = await sb.from('debit_notes').update(rec).eq('id', state.debitPending.id).select().single();
      if(!updErr) saved = upd;
    }
    if(!saved) saved = await insertDebitNote(rec);
    if(!saved){ showToast(t('debitExportErr'), true); return; }
    const no = String(saved.note_number ?? '').trim();
    if(!no) throw new Error('Le numéro du document n’a pas été généré par la base de données.');
    state.debitPending = { id: saved.id, number: no };
    debitShowNumber(no);
    await new Promise(requestAnimationFrame);
    await new Promise(requestAnimationFrame);
    const { error: fdErr } = await sb.from('debit_notes').update({ form_data: collectDebitFormDataFromDocument() }).eq('id', saved.id);
    if(fdErr) throw fdErr;
    showToast(`${t('debitSavedToast')} (N\u00b0 ${no})`);
  }catch(e){
    console.error('debitSaveToHistory error:', e);
    showToast(`${t('debitExportErr')}: ${(e && e.message) || e}`, true);
  }finally{
    if(btn) btn.disabled = false;
  }
}

/* Replaces every input/select of the cloned sheet by plain text boxes, so html2canvas
   prints the filled values exactly inside their cells (inputs are often drawn blank or clipped). */
function debitPrintifyControls(original, clone, clonedDoc){
  const src = Array.from(original.querySelectorAll('input, textarea, select'));
  const dst = Array.from(clone.querySelectorAll('input, textarea, select'));
  clone.style.zoom = '1';                      // capture at full size (no CSS zoom in the picture)
  const win = clonedDoc.defaultView || window;
  const jobs = [];
  src.forEach((s, i) => {                      // 1) read everything first
    const d = dst[i];
    if(!d || !d.parentNode || s.type === 'hidden') return;
    const cs = win.getComputedStyle(d);
    let text = '';
    if(s.tagName === 'SELECT'){
      const opt = s.options[s.selectedIndex];
      text = opt ? opt.text : '';
    }else if(s.type === 'date'){
      text = ccsIsoToFr(s.value);
    }else{
      text = s.value || '';
    }
    jobs.push({
      d, text, tag: s.tagName, h: d.offsetHeight,
      align: cs.textAlign, color: cs.color, bg: cs.backgroundColor,
      bStyle: cs.borderTopStyle, bWidth: cs.borderTopWidth, bColor: cs.borderTopColor,
      pad: `${cs.paddingTop} ${cs.paddingRight} ${cs.paddingBottom} ${cs.paddingLeft}`,
      ff: cs.fontFamily, fs: cs.fontSize, fw: cs.fontWeight, fst: cs.fontStyle, ls: cs.letterSpacing
    });
  });
  jobs.forEach(j => {                          // 2) then replace each control by a plain text box
    const justify = j.align === 'center' ? 'center' : (j.align === 'right' || j.align === 'end') ? 'flex-end' : 'flex-start';
    const color = (!j.color || j.color === 'transparent' || j.color === 'rgba(0, 0, 0, 0)') ? '#111' : j.color;
    const border = (j.bStyle && j.bStyle !== 'none' && parseFloat(j.bWidth) > 0) ? `${j.bWidth} ${j.bStyle} ${j.bColor}` : '0';
    const box = clonedDoc.createElement('div');
    box.textContent = j.text;
    box.style.cssText = [
      'display:flex', 'align-items:' + (j.tag === 'TEXTAREA' ? 'flex-start' : 'center'),
      'justify-content:' + justify, 'box-sizing:border-box', 'width:100%',
      j.h ? 'height:' + j.h + 'px' : 'min-height:6px',
      'overflow:hidden', 'white-space:' + (j.tag === 'TEXTAREA' ? 'pre-wrap' : 'nowrap'),
      'margin:0', 'padding:' + j.pad,
      `font-family:${j.ff}`, `font-size:${j.fs}`, `font-weight:${j.fw}`,
      `font-style:${j.fst}`, `letter-spacing:${j.ls}`,
      `color:${color}`, `background:${j.bg}`, `border:${border}`
    ].join(';');
    j.d.replaceWith(box);
  });
}

async function debitMakePdfBlob(el){
  const canvas = await html2canvas(el, {
    scale: 2 / DEBIT_SCALE, backgroundColor: '#ffffff', useCORS: true,
    windowWidth: 2800,
    onclone: (clonedDoc) => {
      const original = document.getElementById('debit-form-capture');
      const clone = clonedDoc.getElementById('debit-form-capture');
      if(!original || !clone) return;
      debitPrintifyControls(original, clone, clonedDoc);
    }
  });
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF('p', 'mm', 'a4');
  const pw = pdf.internal.pageSize.getWidth();
  const ph = pdf.internal.pageSize.getHeight();
  let w = pw, h = canvas.height * w / canvas.width;
  if(h > ph){ h = ph; w = canvas.width * h / canvas.height; }   // one A4 page, never a blank 2nd page
  pdf.addImage(canvas.toDataURL('image/png'), 'PNG', (pw - w) / 2, 0, w, h);
  return pdf.output('blob');
}

/* Button 2: exports the sheet exactly as it is on screen (filled) to the device. */
async function debitExportPdf(main){
  if(!state.debitPending){ showToast(t('debitSaveFirst'), true); return; }
  if(typeof html2canvas === 'undefined' || !window.jspdf){
    showToast('html2canvas/jsPDF non chargés — vérifiez votre connexion internet et réessayez', true);
    return;
  }
  const btn = document.getElementById('a-pdf-btn');
  if(btn) btn.disabled = true;
  const busy = pdfBusyOverlay(state.lang === 'ar' ? 'جارٍ إنشاء PDF… الرجاء الانتظار' : 'Génération du PDF… veuillez patienter');
  try{
    const no = state.debitPending.number;
    debitCalc();
    debitShowNumber(no);
    await new Promise(requestAnimationFrame);
    await new Promise(requestAnimationFrame);
    // keep Historique identical to the sheet that is exported
    const { rec } = debitSheetRecord();
    const { error: hErr } = await sb.from('debit_notes').update(rec).eq('id', state.debitPending.id);
    if(hErr) console.warn('history update failed', hErr);
    const el = document.getElementById('debit-form-capture');
    if(!el) throw new Error(state.lang === 'ar' ? 'الوثيقة لم تعد معروضة على الشاشة' : 'Le document n’est plus affiché à l’écran');
    const blob = await debitMakePdfBlob(el);
    if(!blob || blob.size < 1000) throw new Error('PDF vide');
    downloadBlobToDevice(blob, `debit_note_${no}.pdf`);
    showToast(`${t('debitExportedToast')} (N\u00b0 ${no})`);
  }catch(e){
    console.error('debitExportPdf error:', e);
    showToast(`${t('debitExportErr')}: ${(e && e.message) || e}`, true);
  }finally{
    if(busy.parentNode) busy.parentNode.removeChild(busy);
    if(btn) btn.disabled = false;
  }
}

async function debitExportExisting(main, note){
  const btn = document.getElementById('a-transfer-btn');
  if(btn) btn.disabled = true;
  try{
    debitCalc();
    const formData = collectDebitFormDataFromDocument();
    const el = document.getElementById('debit-form-capture');
    if(typeof html2canvas === 'undefined' || !window.jspdf || !el) throw new Error('Export libraries/document unavailable');
    const canvas = await html2canvas(el, {
      scale: 2, backgroundColor: '#ffffff', useCORS: true,
      onclone: (clonedDoc) => {
        const original = document.getElementById('debit-form-capture');
        const clone = clonedDoc.getElementById('debit-form-capture');
        if(!original || !clone) return;
        const origFields = original.querySelectorAll('input, textarea, select');
        const cloneFields = clone.querySelectorAll('input, textarea, select');
        origFields.forEach((origEl, i) => {
          const cloneEl = cloneFields[i]; if(!cloneEl) return;
          if(origEl.tagName === 'SELECT'){
            Array.from(cloneEl.options).forEach(o=>{o.removeAttribute('selected');o.selected=false;});
            const selected = Array.from(cloneEl.options).find(o=>o.value===origEl.value);
            if(selected){selected.setAttribute('selected','selected');selected.selected=true;}
            cloneEl.value = origEl.value;
          }else if(origEl.type==='checkbox'||origEl.type==='radio'){
            cloneEl.checked=origEl.checked;
            if(origEl.checked) cloneEl.setAttribute('checked','checked'); else cloneEl.removeAttribute('checked');
          }else if(origEl.tagName==='TEXTAREA'){
            cloneEl.value=origEl.value; cloneEl.textContent=origEl.value;
          }else{
            cloneEl.value=origEl.value; cloneEl.setAttribute('value',origEl.value);
          }
        });
      }
    });
    const imgData = canvas.toDataURL('image/png');
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p','mm','a4');
    const pageWidth=210, pageHeight=297;
    const imgWidth=pageWidth, imgHeight=canvas.height*imgWidth/canvas.width;
    let heightLeft=imgHeight, position=0;
    pdf.addImage(imgData,'PNG',0,position,imgWidth,imgHeight);
    heightLeft-=pageHeight;
    while(heightLeft>0){ position=heightLeft-imgHeight; pdf.addPage(); pdf.addImage(imgData,'PNG',0,position,imgWidth,imgHeight); heightLeft-=pageHeight; }
    const blob=pdf.output('blob');
    const supplier=(document.getElementById('a-supplier')?.value||note.fournisseur||'supplier').trim();
    const folder=slugFolder(supplier);
    const filename=`debit_${note.note_number}_${Date.now()}.pdf`;
    const pdfPath=await uploadDebitPdf(folder,blob,filename);
    if(!pdfPath) throw new Error('PDF upload failed');
    await sb.from('debit_notes').update({form_data:formData,pdf_path:pdfPath}).eq('id',note.id);
    note.form_data=formData; note.pdf_path=pdfPath;
    showToast(`PDF exporté (N° ${note.note_number})`);
    renderDebitList(main);
  }catch(e){
    console.error('debitExportExisting error:',e);
    showToast(`${t('debitExportErr')}: ${(e&&e.message)||e}`,true);
  }finally{ if(btn) btn.disabled=false; }
}

function renderDebitPdfWindow(main){
  const files = state.debitPdfFiles || [];
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.pdf.replace('<svg','<svg width="18" height="18"')} ${t('debitPdfTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="form-panel" style="margin-bottom:16px;">
      <label>${t('uploadPdfBtn')}</label>
      <input id="dpdf-file" type="file" accept="application/pdf,.pdf">
      <button class="btn" id="dpdf-upload" style="max-width:220px;">${t('uploadPdfBtn')}</button>
    </div>
    ${files.length===0 ? `<div class="empty-state">${t('noPdf')}.</div>` : `
    <div class="tile-grid">
      ${files.map((f,i) => `
        <div class="tile" id="dpv-${i}">
          <button class="tile-del" id="dpd-${i}" title="${t('deletedToast')}">${ICONS.del}</button>
          <span class="tile-icon">📄</span>${escHtml(f.name.replace(/^\d+_/,'').replace(/\.pdf$/i,''))}
          ${f.metadata ? `<div style="color:var(--text-muted);font-weight:400;font-size:11px;margin-top:4px;">${fmtSize(f.metadata.size)}</div>` : ''}
        </div>`).join('')}
    </div>`}`;
  document.getElementById('back').onclick = () => { state.view='debit-menu'; render(); };
  document.getElementById('dpdf-upload').onclick = async () => {
    const file = document.getElementById('dpdf-file').files[0];
    if(!file || !(file.type === 'application/pdf' || /\.pdf$/i.test(file.name))){ showToast(t('pdfSelectErr'), true); return; }
    const btn = document.getElementById('dpdf-upload'); btn.disabled = true;
    const ok = await uploadDebitManualPdf(file);
    btn.disabled = false;
    if(ok){ await listDebitManualPdfs(); showToast(t('pdfUploadedToast')); renderDebitPdfWindow(main); }
  };
  files.forEach((f,i) => {
    const fullPath = `${DEBIT_MANUAL_FOLDER}/${f.name}`;
    const tile = document.getElementById('dpv-'+i);
    const del = document.getElementById('dpd-'+i);
    if(tile) tile.onclick = async (e) => {
      if(e.target.closest('.tile-del')) return;
      const url = await getDebitPdfUrl(fullPath);
      if(url) window.open(url, '_blank');
    };
    if(del) del.onclick = async (e) => {
      e.stopPropagation();
      if(!confirm(t('confirmDeletePdf'))) return;
      const { error } = await sb.storage.from('debit-notes').remove([fullPath]);
      if(error){ showToast(t('connError'), true); return; }
      await listDebitManualPdfs(); showToast(t('pdfDeletedToast')); renderDebitPdfWindow(main);
    };
  });
}

/* ================= COST COLLECTION SHEET AEU ================= */
const CCS_TYPES = [
  'OTHER \u2013 SEE DETAILS',
  'OVERTIME (ACTUAL DIRECT LABOUR COSTS)',
  'REWORK (CUSTOMER INVOICES, ALV HARDWARE)',
  'REWORK AT ALV (LABOUR)',
  'SCRAP (FINISHED GOODS, WITHOUT PROFIT)',
  'SORTING COSTS EXTERNAL',
  'SORTING COSTS INTERNAL (LABOUR)',
  'TRAVEL COSTS ALV (W/O HOURLY COSTS)',
  'LINE STOPPAGE CUSTOMER (INVOICED TO ALV)',
  'LINE STOPPAGE ALV (LOST DIRECT LABOUR)',
  'NCM ADMINISTRATION',
  'PREMIUM FREIGHT OUTBOUND (INVOICES)',
  'PREMIUM FREIGHT INBOUND (INVOICES)'
];
const CCS_BUCKET = 'debit-notes';
const CCS_FOLDER = 'ccs-manual';

const CCS_I18N = {
  cardTitle:{fr:'Cost Collection Sheet', ar:'Cost Collection Sheet'},
  cardDesc:{fr:'Feuille de collecte des co\u00fbts AEU \u2014 demande de confirmation.', ar:'\u0648\u0631\u0642\u0629 \u062c\u0645\u0639 \u0627\u0644\u062a\u0643\u0627\u0644\u064a\u0641 AEU \u2014 \u0637\u0644\u0628 \u062a\u0623\u0643\u064a\u062f \u062a\u0639\u0648\u064a\u0636 \u0627\u0644\u062a\u0643\u0627\u0644\u064a\u0641.'},
  newTitle:{fr:'Nouvelle feuille', ar:'\u0648\u0631\u0642\u0629 \u062c\u062f\u064a\u062f\u0629'},
  newDesc:{fr:'Remplir la feuille et l\u2019exporter en PDF.', ar:'\u062a\u0639\u0628\u0626\u0629 \u0627\u0644\u0648\u0631\u0642\u0629 \u0648\u062a\u0635\u062f\u064a\u0631\u0647\u0627 \u0625\u0644\u0649 PDF.'},
  listTitle:{fr:'Historique', ar:'\u0627\u0644\u0633\u062c\u0644'},
  listDesc:{fr:'Feuilles enregistr\u00e9es et leur statut.', ar:'\u0627\u0644\u0623\u0648\u0631\u0627\u0642 \u0627\u0644\u0645\u062d\u0641\u0648\u0638\u0629 \u0648\u062d\u0627\u0644\u062a\u0647\u0627.'},
  pdfTitle:{fr:'PDF', ar:'PDF'},
  pdfDesc:{fr:'Ajouter et consulter des PDF manuellement.', ar:'\u0625\u062f\u062e\u0627\u0644 \u0648\u0645\u0631\u0627\u062c\u0639\u0629 \u0645\u0644\u0641\u0627\u062a PDF \u064a\u062f\u0648\u064a\u064b\u0627.'},
  exportBtn:{fr:'Export PDF', ar:'Export PDF'},
  noSheets:{fr:'Aucune feuille enregistr\u00e9e', ar:'\u0644\u0627 \u062a\u0648\u062c\u062f \u0623\u0648\u0631\u0627\u0642 \u0645\u062d\u0641\u0648\u0638\u0629'},
  colNo:{fr:'N\u00b0', ar:'\u0627\u0644\u0631\u0642\u0645'},
  colDate:{fr:'Date', ar:'\u0627\u0644\u062a\u0627\u0631\u064a\u062e'},
  colSupplier:{fr:'Fournisseur', ar:'\u0627\u0644\u0645\u0648\u0631\u062f'},
  colPart:{fr:'Pi\u00e8ce', ar:'\u0627\u0644\u0642\u0637\u0639\u0629'},
  colBy:{fr:'Cr\u00e9\u00e9 par', ar:'\u0623\u0646\u0634\u0626\u062a \u0628\u0648\u0627\u0633\u0637\u0629'},
  errSupplier:{fr:'Renseignez le nom du fournisseur (Supplier Name)', ar:'\u0623\u062f\u062e\u0644 \u0627\u0633\u0645 \u0627\u0644\u0645\u0648\u0631\u062f (Supplier Name)'},
  errLine:{fr:'Choisissez au moins un Cost Type', ar:'\u0627\u062e\u062a\u0631 Cost Type \u0648\u0627\u062d\u062f\u064b\u0627 \u0639\u0644\u0649 \u0627\u0644\u0623\u0642\u0644'},
  errLibs:{fr:'Biblioth\u00e8ques PDF indisponibles', ar:'\u0645\u0643\u062a\u0628\u0627\u062a PDF \u063a\u064a\u0631 \u0645\u062a\u0627\u062d\u0629'},
  errNoTable:{fr:'Table cost_sheets introuvable : ex\u00e9cutez le script SQL', ar:'\u0627\u0644\u062c\u062f\u0648\u0644 cost_sheets \u063a\u064a\u0631 \u0645\u0648\u062c\u0648\u062f: \u0634\u063a\u0651\u0644 \u0633\u0643\u0631\u0628\u062a SQL \u0623\u0648\u0644\u064b\u0627'},
  busy:{fr:'G\u00e9n\u00e9ration du PDF\u2026 veuillez patienter', ar:'\u062c\u0627\u0631\u064d \u0625\u0646\u0634\u0627\u0621 PDF\u2026 \u0627\u0644\u0631\u062c\u0627\u0621 \u0627\u0644\u0627\u0646\u062a\u0638\u0627\u0631'},
  exported:{fr:'PDF g\u00e9n\u00e9r\u00e9 et t\u00e9l\u00e9charg\u00e9', ar:'\u062a\u0645 \u0625\u0646\u0634\u0627\u0621 PDF \u0648\u062a\u0646\u0632\u064a\u0644\u0647 \u0639\u0644\u0649 \u0627\u0644\u062c\u0647\u0627\u0632'},
  exportErr:{fr:'\u00c9chec de la g\u00e9n\u00e9ration du PDF', ar:'\u0641\u0634\u0644 \u0625\u0646\u0634\u0627\u0621 PDF'},
  kept:{fr:' \u2014 donn\u00e9es enregistr\u00e9es dans Historique', ar:' \u2014 \u0627\u0644\u0628\u064a\u0627\u0646\u0627\u062a \u0645\u062d\u0641\u0648\u0638\u0629 \u0641\u064a \u0627\u0644\u0633\u062c\u0644'},
  confirmDelete:{fr:'Supprimer cette feuille ?', ar:'\u062d\u0630\u0641 \u0647\u0630\u0647 \u0627\u0644\u0648\u0631\u0642\u0629\u061f'},
  saveBtn:{fr:'Enregistrer (N° + Historique)', ar:'حفظ (رقم + السجل)'},
  saved:{fr:'Feuille enregistrée', ar:'تم حفظ الورقة'},
  saveFirst:{fr:'Enregistrez d’abord la feuille pour obtenir son numéro', ar:'احفظ الورقة أولاً للحصول على رقمها'},
  saveErr:{fr:'Échec de l’enregistrement', ar:'فشل الحفظ'},
  autoNo:{fr:'AUTO', ar:'AUTO'}
};
function tc(k){ const e = CCS_I18N[k]; return e ? (e[state.lang] || e.fr) : k; }

const CCS_CSS = `
.ccs-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;max-width:100%;background:#d8d8d8;padding:10px 0;border-radius:6px}
.ccs-tools{display:flex;gap:8px;margin:0 0 12px;max-width:260px}
.ccs-page{box-sizing:border-box;width:210mm;height:297mm;margin:0 auto;background:#fff;padding:8mm 8mm 6mm;color:#111;font-family:Arial,Helvetica,sans-serif;font-size:11.5px;line-height:1.2;position:relative;overflow:hidden;box-shadow:0 1px 8px #888;direction:ltr;text-align:left}
.ccs-page *{box-sizing:border-box}
.ccs-head{display:grid;grid-template-columns:40mm 1fr 48mm;height:25mm;border-bottom:1px solid #777}
.ccs-logo{font-weight:bold;font-size:21px;color:#1d5aa6;letter-spacing:-.5px;padding-top:5mm}
.ccs-title{text-align:center;padding-top:4mm}
.ccs-title b{display:block;font-size:19px}
.ccs-title span{display:block;font-size:13px;margin-top:2mm}
.ccs-ver{border:1px solid #777;margin-top:1mm;height:21mm}
.ccs-ver div{height:50%;display:flex;align-items:center;justify-content:center;font-size:16px}
.ccs-ver div:first-child{border-bottom:1px solid #777}
.ccs-sec{font-weight:bold;font-size:13px;margin:4mm 0 1mm}
.ccs-grid{display:grid;grid-template-columns:35mm 73mm 16mm 16mm 54mm;row-gap:1px;align-items:stretch}
.ccs-lbl{display:flex;align-items:center;justify-content:flex-end;padding-right:1.5mm;font-size:11.5px;text-align:right}
.ccs-cell{position:relative;display:flex;border:1px solid #8a8a8a;background:#fff;min-width:0}
.ccs-cell.y{background:#fff}
.ccs-cell.g{background:#fff;color:#111;font-weight:bold}
.ccs-cell.box{background:#fff}
.ccs-cell>input,.ccs-cell>textarea,.ccs-cell>.ccs-val,.ccs-cell>select{flex:1 1 auto;min-width:0;width:100%}
.ccs-cell input,.ccs-cell textarea,.ccs-val{border:0;outline:0;background:transparent;color:inherit;font:inherit;padding:0 1.2mm;margin:0;border-radius:0;text-align:center;-webkit-appearance:none;appearance:none}
.ccs-cell textarea{resize:none;font-size:8.5px;font-weight:bold;padding:1mm 1.2mm;line-height:1.25;overflow:hidden}
.ccs-cell input::placeholder{color:#111;opacity:1}
.ccs-val{display:flex;align-items:center;justify-content:center;white-space:pre-wrap;word-break:break-word;min-width:0;width:100%}
.ccs-l{justify-content:flex-start;text-align:left}
.ccs-r{justify-content:flex-end;text-align:right}
.ccs-top{align-items:flex-start}
.ccs-ta{font-size:8.5px;font-weight:bold;line-height:1.25;padding:1mm 1.2mm}
.ccs-cell textarea.ccs-top-ta{font-size:8.5px}
.h5{height:5.4mm}.h6{height:6.4mm}.h8{height:8.4mm}.h11{height:11mm}.h12{height:12mm}.h13{height:13mm}
.ccs-conf{font-size:11.5px;line-height:1.38}
.ccs-conf b{font-size:13px}
.ccs-cd{display:grid;grid-template-columns:35mm 73mm 16mm 16.5mm 18mm 35.5mm;row-gap:0}
.ccs-th{font-weight:bold;font-size:11.5px;color:#333;padding:1mm 1.2mm;border-bottom:1px solid #777}
.ccs-fix{align-items:center;justify-content:center;font-size:10.5px}
.ccs-cost input,.ccs-cost .ccs-val{padding-right:7mm}
.ccs-eur{position:absolute;right:1.6mm;top:0;bottom:0;display:flex;align-items:center;font-size:11.5px;pointer-events:none}
.ccs-selwrap .ccs-val{font-size:7.6px;text-transform:uppercase;line-height:1.15;padding:0 5mm 0 1.2mm}
.ccs-selwrap select{position:absolute;inset:0;width:100%;height:100%;opacity:0;font-size:16px;cursor:pointer;margin:0;border:0}
.ccs-arrow{position:absolute;right:1mm;top:0;bottom:0;display:flex;align-items:center;font-size:9px;color:#333;pointer-events:none}
.ccs-tot{font-weight:bold;align-items:center;justify-content:center;font-size:11.5px}
.ccs-foot{margin-top:5mm}
.ccs-fbox{border:1px solid #999;padding:1.4mm 2mm;font-size:11.5px}
.ccs-fbox+.ccs-fbox{margin-top:2mm}
#ccs-print-host{position:absolute;left:0;top:0;width:210mm;height:297mm;z-index:-1;background:#fff;overflow:hidden;pointer-events:none}
#ccs-print-host .ccs-page{margin:0;box-shadow:none}
`;
(function(){ const st = document.createElement('style'); st.id = 'ccs-style'; st.textContent = CCS_CSS; document.head.appendChild(st); })();

function ccsTodayISO(){ const d = new Date(); const p = n => String(n).padStart(2,'0'); return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`; }
function ccsIsoToFr(iso){ const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? `${m[3]}/${m[2]}/${m[1]}` : (iso || ''); }
function ccsNum(v){ const n = parseFloat(String(v == null ? '' : v).replace(/[^\d,.\-]/g,'').replace(',', '.')); return Number.isFinite(n) ? n : 0; }
function ccsFmt(n){ return Number(n).toLocaleString('fr-FR', { minimumFractionDigits:2, maximumFractionDigits:2 }); }
function ccsDbErr(e){
  const m = (e && e.message) || String(e || '');
  if(/cost_sheets/i.test(m) && /(schema cache|does not exist|relation|not find)/i.test(m)) return tc('errNoTable');
  return `${t('connError')}: ${m}`;
}
function ccsDefaults(){ try{ return JSON.parse(localStorage.getItem('ccs:defaults') || '{}') || {}; }catch(e){ return {}; } }
function ccsSaveDefaults(d){ try{ localStorage.setItem('ccs:defaults', JSON.stringify(d)); }catch(e){} }

function ccsPageHTML(){
  const opts = '<option value=""></option>' + CCS_TYPES.map(x => `<option value="${escAttr(x)}">${escHtml(x)}</option>`).join('');
  const lineRow = i => `
    <div class="ccs-cell y h8 ccs-selwrap"><div class="ccs-val ccs-l"></div><select data-f="type${i}">${opts}</select><span class="ccs-arrow ccs-noprint">&#9662;</span></div>
    <div class="ccs-cell y h8"><input data-f="detail${i}" data-al="ccs-l" autocomplete="off"></div>
    <div class="ccs-cell y h8"><input data-f="units${i}" inputmode="decimal" autocomplete="off"></div>
    <div class="ccs-cell h8 ccs-fix">UNIT</div>
    <div class="ccs-cell h8 ccs-fix">EUR</div>
    <div class="ccs-cell y h8 ccs-cost"><input data-f="cost${i}" data-al="ccs-r" data-dash="1" inputmode="decimal" placeholder="-" autocomplete="off"><span class="ccs-eur">&euro;</span></div>`;
  return `
  <div class="ccs-page" id="ccs-page">
    <div class="ccs-head">
      <div class="ccs-logo">Autoliv</div>
      <div class="ccs-title"><b>Cost Collection Sheet AEU</b><span>Request for confirmation of cost compensation</span></div>
      <div class="ccs-ver"><div>Version: 1.1</div><div id="ccs-verdate"></div></div>
    </div>

    <div class="ccs-sec">Issued by</div>
    <div class="ccs-grid">
      <div class="ccs-lbl">Plant:</div><div class="ccs-cell y h5"><input data-f="plant" autocomplete="off"></div>
      <div class="ccs-lbl">Name:</div><div class="ccs-cell y h5" style="grid-column:4 / 6"><input data-f="name" autocomplete="off"></div>
      <div class="ccs-lbl">Date:</div><div class="ccs-cell y h5"><input data-f="date" readonly></div>
      <div class="ccs-lbl">Phone:</div><div class="ccs-cell y h5" style="grid-column:4 / 6"><input data-f="phone" autocomplete="off"></div>
      <div class="ccs-lbl">No. Of Issue:</div><div class="ccs-cell g h5"><input data-f="issue" readonly placeholder="${escAttr(tc('autoNo'))}"></div>
      <div class="ccs-lbl">eMail:</div><div class="ccs-cell y h5" style="grid-column:4 / 6"><input data-f="email" autocomplete="off"></div>
    </div>

    <div class="ccs-sec" style="margin-top:6mm">Issued to</div>
    <div class="ccs-grid">
      <div class="ccs-lbl">Supplier Name:</div><div class="ccs-cell y h5"><input data-f="supplier" list="ccs-supplier-list" autocomplete="off"></div>
      <div class="ccs-lbl" style="grid-column:3 / 5">Supplier ID</div><div class="ccs-cell y h5"><input data-f="supplierId" autocomplete="off"></div>
      <div class="ccs-lbl" style="grid-column:1">Supplier Contact:</div><div class="ccs-cell y h5"><input data-f="contact" autocomplete="off"></div>
    </div>

    <div class="ccs-sec" style="margin-top:6mm">Reference</div>
    <div class="ccs-grid">
      <div class="ccs-lbl">NCM No.:</div><div class="ccs-cell y h11"><input data-f="ncmNo" autocomplete="off"></div>
      <div></div><div class="ccs-cell box h11" style="align-items:center;justify-content:center">Area:</div><div class="ccs-cell y h11"><input data-f="area" autocomplete="off"></div>
      <div class="ccs-lbl" style="grid-column:1">Problem Description:</div><div class="ccs-cell y h12"><textarea data-f="problem" data-al="ccs-l ccs-top ccs-ta" spellcheck="false"></textarea></div>
      <div class="ccs-lbl" style="grid-column:1">NCM creation date:</div><div class="ccs-cell h5"><input data-f="ncmDate" autocomplete="off"></div>
      <div class="ccs-lbl" style="grid-column:1">NCM Type:</div><div class="ccs-cell y h5"><input data-f="ncmType" autocomplete="off"></div>
      <div class="ccs-lbl" style="grid-column:1">Part No./Name:</div><div class="ccs-cell y h13"><textarea data-f="part" data-al="ccs-l ccs-top ccs-ta" spellcheck="false"></textarea></div>
    </div>

    <div class="ccs-conf" style="margin-top:5mm">
      <b>Supplier confirmation</b><br>
      Dear supplier,<br>
      you have been informed in advance that above described incident is related to costs within the Autoliv organisation.<br>
      We request you to confirm acceptance of debit note as below.<br>
      Please respond to the sender within 1 week. No response in this period is considered as acceptance.<br>
      Please note that there might be other cost charged related to above incident at later time until NCM is completed.<br>
      In case of questions contact the sender respectively the responsible for NCM.
    </div>

    <div class="ccs-sec" style="margin-top:4mm">Cost Details</div>
    <div class="ccs-cd">
      <div class="ccs-th">Cost Type</div><div class="ccs-th">Cost Details</div><div class="ccs-th">Units</div><div class="ccs-th">UoM</div><div class="ccs-th">Currency</div><div class="ccs-th">Cost Supplier</div>
      ${[0,1,2,3].map(lineRow).join('')}
      <div class="ccs-cell h6"></div><div class="ccs-cell h6"></div><div class="ccs-cell h6"></div><div class="ccs-cell h6"></div>
      <div class="ccs-cell h6 ccs-tot">TOTAL</div>
      <div class="ccs-cell y h6 ccs-cost"><div class="ccs-val ccs-r" id="ccs-total">-</div><span class="ccs-eur">&euro;</span></div>
    </div>

    <div class="ccs-foot">
      <div class="ccs-fbox" style="height:15mm">Supplier Confirmation that invoice will be accepted</div>
      <div class="ccs-fbox" style="height:24mm">Supplier representative: Name, Function, Signature, Date</div>
    </div>
    <datalist id="ccs-supplier-list">${[...new Set((state.catalog || []).map(c => c.fournisseur).filter(Boolean))].map(f => `<option value="${escAttr(f)}">`).join('')}</datalist>
  </div>`;
}

function ccsRecalc(page){
  let sum = 0;
  for(let i = 0; i < 4; i++){ const el = page.querySelector(`[data-f="cost${i}"]`); if(el) sum += ccsNum(el.value); }
  const tot = page.querySelector('#ccs-total');
  if(tot) tot.textContent = sum ? ccsFmt(sum) : '-';
  return sum;
}

function ccsWire(page){
  page.querySelectorAll('[data-al]').forEach(el => el.classList.add(...el.dataset.al.split(' ').filter(Boolean)));
  const d = ccsDefaults();
  const set = (f, v) => { const el = page.querySelector(`[data-f="${f}"]`); if(el) el.value = v; };
  set('plant', d.plant || 'SWT'); set('name', d.name || ''); set('phone', d.phone || ''); set('email', d.email || '');
  set('area', d.area || 'Quality');
  set('date', ccsIsoToFr(ccsTodayISO()));
  const vd = page.querySelector('#ccs-verdate'); if(vd) vd.textContent = ccsIsoToFr(ccsTodayISO());
  page.querySelectorAll('.ccs-selwrap').forEach(w => {
    const sel = w.querySelector('select'), val = w.querySelector('.ccs-val');
    sel.onchange = () => { val.textContent = sel.value; sel.title = sel.value; };
  });
  for(let i = 0; i < 4; i++){
    const c = page.querySelector(`[data-f="cost${i}"]`);
    c.addEventListener('input', () => ccsRecalc(page));
    c.addEventListener('blur', () => { if(c.value.trim() !== ''){ c.value = ccsFmt(ccsNum(c.value)); } ccsRecalc(page); });
    c.addEventListener('focus', () => { if(c.value.trim() !== '') c.value = String(ccsNum(c.value)).replace('.', ','); });
  }
  ccsRecalc(page);
}

function ccsCollect(page){
  const g = f => { const el = page.querySelector(`[data-f="${f}"]`); return el ? String(el.value || '').trim() : ''; };
  const lines = [0,1,2,3].map(i => ({ type: g('type'+i), detail: g('detail'+i), units: g('units'+i), cost: ccsNum(g('cost'+i)) }));
  return {
    plant: g('plant'), name: g('name'), phone: g('phone'), email: g('email'),
    supplier: g('supplier'), supplierId: g('supplierId'), contact: g('contact'),
    ncmNo: g('ncmNo'), area: g('area'), problem: g('problem'), ncmDate: g('ncmDate'), ncmType: g('ncmType'), part: g('part'),
    lines, total: lines.reduce((s, l) => s + l.cost, 0)
  };
}

function ccsBuildPrintNode(page){
  const clone = page.cloneNode(true);
  const orig = page.querySelectorAll('input, textarea, select');
  const copy = clone.querySelectorAll('input, textarea, select');
  orig.forEach((el, i) => {
    const c = copy[i];
    if(!c || !c.parentNode) return;
    if(el.tagName === 'SELECT'){ c.remove(); return; }
    const box = document.createElement('div');
    box.className = 'ccs-val ' + (el.dataset.al || '');
    box.textContent = el.value || (el.dataset.dash ? '-' : '');
    c.parentNode.replaceChild(box, c);
  });
  clone.querySelectorAll('.ccs-noprint, datalist').forEach(n => n.remove());
  clone.removeAttribute('id');
  return clone;
}

async function ccsMakePdfBlob(page){
  const host = document.createElement('div');
  host.id = 'ccs-print-host';
  host.appendChild(ccsBuildPrintNode(page));
  document.body.appendChild(host);
  try{
    window.scrollTo(0, 0);
    await new Promise(r => setTimeout(r, 150));
    if(document.fonts && document.fonts.ready){ try{ await document.fonts.ready; }catch(e){} }
    const canvas = await html2canvas(host.firstElementChild, {
      scale: 2, backgroundColor: '#ffffff', useCORS: true,
      scrollX: 0, scrollY: 0, windowWidth: 1100, windowHeight: 1500
    });
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p', 'mm', 'a4');
    pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, 210, 297);
    return pdf.output('blob');
  } finally {
    if(host.parentNode) host.parentNode.removeChild(host);
  }
}

function ccsRecordFrom(data){
  return {
    sheet_date: ccsTodayISO(), plant: data.plant, issuer_name: data.name, issuer_phone: data.phone, issuer_email: data.email,
    supplier_name: data.supplier, supplier_contact: data.contact, supplier_id: data.supplierId,
    ncm_no: data.ncmNo, area: data.area, problem: data.problem, ncm_date: data.ncmDate, ncm_type: data.ncmType, part: data.part,
    lines: data.lines, total: data.total, created_by: state.currentUser ? state.currentUser.name : ''
  };
}

/* Button 1: gives the sheet its number and saves the data in Historique. The sheet stays on screen. */
async function ccsSave(main){
  const page = document.getElementById('ccs-page');
  const btn = document.getElementById('ccs-save');
  if(!page || !btn) return;
  const data = ccsCollect(page);
  if(!data.supplier){ showToast(tc('errSupplier'), true); return; }
  if(!data.lines.some(l => l.type)){ showToast(tc('errLine'), true); return; }
  btn.disabled = true;
  try{
    const rec = ccsRecordFrom(data);
    let saved = null;
    if(state.ccsPending){
      const { data: upd, error: ue } = await sb.from('cost_sheets').update(rec).eq('id', state.ccsPending.id).select().single();
      if(!ue) saved = upd;
    }
    if(!saved){
      const { data: ins, error: ie } = await sb.from('cost_sheets').insert(rec).select().single();
      if(ie) throw ie;
      saved = ins;
    }
    state.ccsPending = { id: saved.id, issue_no: saved.issue_no };
    page.querySelector('[data-f="issue"]').value = String(saved.issue_no);
    ccsSaveDefaults({ plant: data.plant, name: data.name, phone: data.phone, email: data.email, area: data.area });
    showToast(`${tc('saved')} (N\u00b0 ${saved.issue_no})`);
  }catch(e){
    console.error('ccsSave error:', e);
    const msg = (e && e.message) ? (/cost_sheets/i.test(e.message) ? tc('errNoTable') : e.message) : e;
    showToast(`${tc('saveErr')}: ${msg}`, true);
  }finally{
    btn.disabled = false;
  }
}

/* Button 2: exports the sheet exactly as it is on screen (filled) to the device. */
async function ccsExportPdf(main){
  const page = document.getElementById('ccs-page');
  const btn = document.getElementById('ccs-pdf-btn');
  if(!page || !btn) return;
  if(!state.ccsPending){ showToast(tc('saveFirst'), true); return; }
  if(typeof html2canvas === 'undefined' || !window.jspdf){ showToast(tc('errLibs'), true); return; }
  btn.disabled = true;
  const busy = pdfBusyOverlay(tc('busy'));
  try{
    const no = state.ccsPending.issue_no;
    page.querySelector('[data-f="issue"]').value = String(no);
    ccsRecalc(page);
    // keep Historique identical to the sheet that is exported
    const { error: hErr } = await sb.from('cost_sheets').update(ccsRecordFrom(ccsCollect(page))).eq('id', state.ccsPending.id);
    if(hErr) console.warn('history update failed', hErr);
    const blob = await ccsMakePdfBlob(page);
    if(!blob || blob.size < 5000) throw new Error('PDF vide');
    downloadBlobToDevice(blob, `cost_sheet_${no}.pdf`);
    showToast(`${tc('exported')} (N\u00b0 ${no})`);
  }catch(e){
    console.error('ccsExportPdf error:', e);
    showToast(`${tc('exportErr')}: ${(e && e.message) ? e.message : e}`, true);
  }finally{
    if(busy.parentNode) busy.parentNode.removeChild(busy);
    btn.disabled = false;
  }
}

function renderCcsMenu(main){
  main.innerHTML = `
    <div class="menu-head"><h1>${tc('cardTitle')}</h1><p>${t('menuSubtitle')}</p></div>
    <div class="menu-grid">
      <div class="menu-card" id="card-ccs-new">${ICONS.entry}<h3>${tc('newTitle')}</h3><p>${tc('newDesc')}</p></div>
      <div class="menu-card" id="card-ccs-list">${ICONS.data}<h3>${tc('listTitle')}</h3><p>${tc('listDesc')}</p></div>
      <div class="menu-card" id="card-ccs-pdf">${ICONS.pdf}<h3>${tc('pdfTitle')}</h3><p>${tc('pdfDesc')}</p></div>
    </div>
    <div class="back-link" id="back-top" style="margin-top:22px;">${ICONS.back} ${t('menuTitle')}</div>`;
  document.getElementById('card-ccs-new').onclick = async () => { await loadCatalog(); state.view='ccs-new'; render(); };
  document.getElementById('card-ccs-list').onclick = async () => { await ccsLoad(); state.view='ccs-list'; render(); };
  document.getElementById('card-ccs-pdf').onclick = async () => { await ccsListPdfs(); state.view='ccs-pdf'; render(); };
  document.getElementById('back-top').onclick = () => { state.view='menu'; render(); };
}

function renderCcsNew(main){
  state.ccsPending = null;
  main.innerHTML = `
    <div class="page-head"><h2>${tc('newTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div style="display:flex;flex-wrap:wrap;gap:8px;margin:0 0 12px;">
      <button class="btn" id="ccs-save" style="width:auto;">${tc('saveBtn')}</button>
      <button class="btn" id="ccs-pdf-btn" style="width:auto;">&#128196; ${tc('exportBtn')}</button>
    </div>
    <div class="ccs-wrap">${ccsPageHTML()}</div>`;
  document.getElementById('back').onclick = () => { state.view='ccs-menu'; render(); };
  ccsWire(document.getElementById('ccs-page'));
  document.getElementById('ccs-save').onclick = () => ccsSave(main);
  document.getElementById('ccs-pdf-btn').onclick = () => ccsExportPdf(main);
}

async function ccsLoad(){
  const { data, error } = await sb.from('cost_sheets').select('*').order('issue_no', { ascending:false });
  if(error){ showToast(ccsDbErr(error), true); state.ccsSheets = []; return; }
  state.ccsSheets = data || [];
}

function renderCcsList(main){
  const rows = state.ccsSheets || [];
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.data.replace('<svg','<svg width="18" height="18"')} ${tc('listTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="table-wrap">
      ${rows.length === 0 ? `<div class="empty-state">${tc('noSheets')}.</div>` : `
      <table>
        <thead><tr>
          <th>No. Of Issue</th><th>Plant</th><th>Supplier Name</th><th>${tc('colDate')}</th><th>Problem Description</th><th>NCM No.</th><th>Area</th><th>Name</th><th>Total</th><th>${t('debitStatusCol')}</th><th></th>
        </tr></thead>
        <tbody>
          ${rows.map(r => `
            <tr>
              <td class="num">${r.issue_no}</td>
              <td>${escHtml(r.plant) || '\u2014'}</td>
              <td>${escHtml(r.supplier_name)}</td>
              <td class="num">${ccsIsoToFr(r.sheet_date)}</td>
              <td style="min-width:200px;max-width:320px;white-space:normal;">${escHtml(r.problem) || '\u2014'}</td>
              <td>${escHtml(r.ncm_no) || '\u2014'}</td>
              <td>${escHtml(r.area) || '\u2014'}</td>
              <td>${escHtml(r.issuer_name) || '\u2014'}</td>
              <td class="num">${debitMoney(r.total)}</td>
              <td>
                <select id="cstatus-${r.id}" style="min-width:110px;border-color:${r.done?'var(--ok)':'var(--border)'};">
                  <option value="0"${r.done?'':' selected'}>${t('debitStatusNotDone')}</option>
                  <option value="1"${r.done?' selected':''}>${t('debitStatusDone')}</option>
                </select>
              </td>
              <td class="row-actions">
                ${state.currentUser.role==='admin' ? `<button class="icon-btn del" id="cdel-${r.id}" title="${t('deletedToast')}">${ICONS.del}</button>` : ''}
              </td>
            </tr>`).join('')}
        </tbody>
      </table>`}
    </div>`;
  document.getElementById('back').onclick = () => { state.view='ccs-menu'; render(); };
  rows.forEach(r => {
    const sel = document.getElementById('cstatus-'+r.id);
    if(sel) sel.onchange = async () => {
      const wanted = sel.value === '1';
      const { error } = await sb.from('cost_sheets').update({ done: wanted }).eq('id', r.id);
      if(error){ showToast(ccsDbErr(error), true); sel.value = r.done ? '1' : '0'; return; }
      r.done = wanted;
      sel.style.borderColor = wanted ? 'var(--ok)' : 'var(--border)';
      showToast(wanted ? t('debitStatusDone') : t('debitStatusNotDone'));
    };
    const del = document.getElementById('cdel-'+r.id);
    if(del) del.onclick = async () => {
      if(!confirm(tc('confirmDelete'))) return;
      const { error } = await sb.from('cost_sheets').delete().eq('id', r.id);
      if(error){ showToast(ccsDbErr(error), true); return; }
      await ccsLoad(); showToast(t('deletedToast')); renderCcsList(main);
    };
  });
}

async function ccsListPdfs(){
  const { data, error } = await sb.storage.from(CCS_BUCKET).list(CCS_FOLDER, { sortBy: { column:'name', order:'desc' } });
  state.ccsPdfFiles = error ? [] : (data || []).filter(f => f.name !== '.emptyFolderPlaceholder' && f.id !== null);
}

function renderCcsPdf(main){
  const files = state.ccsPdfFiles || [];
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.pdf.replace('<svg','<svg width="18" height="18"')} ${tc('pdfTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="form-panel" style="margin-bottom:16px;">
      <label>${t('uploadPdfBtn')}</label>
      <input id="cpdf-file" type="file" accept="application/pdf,.pdf">
      <button class="btn" id="cpdf-upload" style="max-width:220px;">${t('uploadPdfBtn')}</button>
    </div>
    ${files.length === 0 ? `<div class="empty-state">${t('noPdf')}.</div>` : `
    <div class="tile-grid">
      ${files.map((f,i) => `
        <div class="tile" id="cpv-${i}">
          <button class="tile-del" id="cpd-${i}" title="${t('deletedToast')}">${ICONS.del}</button>
          <span class="tile-icon">\ud83d\udcc4</span>${escHtml(f.name.replace(/^\d+_/,'').replace(/\.pdf$/i,''))}
          ${f.metadata ? `<div style="color:var(--text-muted);font-weight:400;font-size:11px;margin-top:4px;">${fmtSize(f.metadata.size)}</div>` : ''}
        </div>`).join('')}
    </div>`}`;
  document.getElementById('back').onclick = () => { state.view='ccs-menu'; render(); };
  document.getElementById('cpdf-upload').onclick = async () => {
    const file = document.getElementById('cpdf-file').files[0];
    if(!file || !(file.type === 'application/pdf' || /\.pdf$/i.test(file.name))){ showToast(t('pdfSelectErr'), true); return; }
    const btn = document.getElementById('cpdf-upload'); btn.disabled = true;
    const { error } = await sb.storage.from(CCS_BUCKET).upload(`${CCS_FOLDER}/${Date.now()}_${file.name}`, file, { contentType:'application/pdf' });
    btn.disabled = false;
    if(error){ showToast(t('pdfUploadErr'), true); return; }
    await ccsListPdfs(); showToast(t('pdfUploadedToast')); renderCcsPdf(main);
  };
  files.forEach((f,i) => {
    const full = `${CCS_FOLDER}/${f.name}`;
    const tile = document.getElementById('cpv-'+i), del = document.getElementById('cpd-'+i);
    if(tile) tile.onclick = async (e) => {
      if(e.target.closest('.tile-del')) return;
      const url = await getDebitPdfUrl(full);
      if(url) window.open(url, '_blank');
    };
    if(del) del.onclick = async (e) => {
      e.stopPropagation();
      if(!confirm(t('confirmDeletePdf'))) return;
      const { error } = await sb.storage.from(CCS_BUCKET).remove([full]);
      if(error){ showToast(t('connError'), true); return; }
      await ccsListPdfs(); showToast(t('pdfDeletedToast')); renderCcsPdf(main);
    };
  });
}

function renderDebitList(main){
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.data.replace('<svg','<svg width="18" height="18"')} ${t('debitListTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
    <div class="table-wrap">
      ${state.debitNotes.length===0 ? `<div class="empty-state">${t('debitNoNotes')}.</div>` : `
      <table>
        <thead><tr>
          <th>${t('debitNoteNumber')}</th><th>${t('fieldDate')}</th><th>${t('debitAswt')}</th><th>${t('fieldFourn')}</th>
          <th>${t('debitDepartment')}</th><th>${t('debitCause')}</th><th>TOTAL</th><th>${t('debitReportBy')}</th><th>${t('debitStatusCol')}</th><th></th>
        </tr></thead>
        <tbody>
          ${state.debitNotes.map(n => `
            <tr>
              <td class="num">${n.note_number}</td>
              <td class="num">${n.note_date||''}</td>
              <td>${escHtml(n.aswt_department)||'—'}</td>
              <td>${escHtml(n.fournisseur)}</td>
              <td>${escHtml(n.supplier_department)||'—'}</td>
              <td>${escHtml(n.cause)}</td>
              <td class="num">${debitMoney(n.grand_total)}</td>
              <td>${escHtml(n.report_raised_by)}</td>
              <td>
                <select id="dstatus-${n.id}" style="min-width:110px;border-color:${debitIsDone(n)?'var(--ok)':'var(--border)'};">
                  <option value="0"${debitIsDone(n)?'':' selected'}>${t('debitStatusNotDone')}</option>
                  <option value="1"${debitIsDone(n)?' selected':''}>${t('debitStatusDone')}</option>
                </select>
              </td>
              <td class="row-actions">
                ${state.currentUser.role==='admin' ? `<button class="icon-btn del" id="ddel-${n.id}" title="${t('deletedToast')}">${ICONS.del}</button>` : ''}
              </td>
            </tr>`).join('')}
        </tbody>
      </table>`}
    </div>`;
  document.getElementById('back').onclick = () => { state.view='debit-menu'; render(); };
  state.debitNotes.forEach(n => {
    const statusSel = document.getElementById('dstatus-'+n.id);
    if(statusSel) statusSel.onchange = async () => {
      const wanted = statusSel.value === '1';
      const ok = await setDebitDone(n, wanted);
      if(!ok){ statusSel.value = debitIsDone(n) ? '1' : '0'; return; }
      statusSel.style.borderColor = wanted ? 'var(--ok)' : 'var(--border)';
      showToast(wanted ? t('debitStatusDone') : t('debitStatusNotDone'));
    };
    const delBtn = document.getElementById('ddel-'+n.id);
    if(delBtn) delBtn.onclick = async () => {
      if(confirm(t('debitConfirmDelete'))){
        const ok = await deleteDebitNote(n.id, n.pdf_path);
        if(ok){ await loadDebitNotes(); showToast(t('deletedToast')); renderDebitList(main); }
      }
    };
  });
}


/* ================= TEST E-MAIL (recipients page) ================= */
async function sendQualityReportTest(){
  const btn = document.getElementById('send-test-report-btn');
  if(btn){ btn.disabled = true; btn.textContent = 'Envoi...'; }
  try{
    const { data, error } = await sb.rpc('send_quality_report', { p_test: true });
    if(error) throw new Error(error.message);
    if(!data || data.ok !== true) throw new Error((data && data.error) || 'Échec de l’envoi');
    let status = null;
    for(let i = 0; i < 6; i++){
      await new Promise(r => setTimeout(r, 1500));
      const { data: st, error: stErr } = await sb.rpc('report_request_status', { p_id: data.request_id });
      if(stErr) break;
      if(st && st.done){ status = st; break; }
    }
    if(status){
      if(status.status_code >= 200 && status.status_code < 300){
        showToast(`Rapport envoyé à ${data.recipients} destinataire(s)`);
      }else{
        throw new Error(`[${status.status_code}] ` + (status.message || 'Échec Resend'));
      }
    }else{
      showToast(`Envoi lancé (${data.recipients} destinataire(s)) — vérifiez votre boîte mail`);
    }
  }catch(e){
    console.error('sendQualityReportTest:', e);
    showToast('Erreur envoi : ' + (e && e.message ? e.message : 'Erreur inconnue'), true);
  }finally{
    if(btn){ btn.disabled = false; btn.textContent = t('sendTestReportBtn'); }
  }
}

/* ================= PORTAIL OPÉRATEUR (8 tables de contrôle) ================= */
const OP_TYPES = ['Grani', 'Rayure', 'Trace', 'Piqûre', 'Coup', 'Cosse', 'Autre'];
/* Horaires des équipes (minutes depuis minuit) : matin 5h30->13h15, après-midi 13h15->21h00, nuit 21h00->5h30 */
const OP_SHIFTS = {
  morning:   { start: 5 * 60 + 30,  end: 13 * 60 + 15, label: '5h30 \u2192 13h15' },
  afternoon: { start: 13 * 60 + 15, end: 21 * 60,      label: '13h15 \u2192 21h00' },
  night:     { start: 21 * 60,      end: 29 * 60 + 30, label: '21h00 \u2192 5h30' }
};
/* Pauses : 20 minutes déduites des heures 10:00-11:00 et 15:00-16:00 */
const OP_BREAKS = [{ hour: 10, minutes: 20 }, { hour: 15, minutes: 20 }];
function opFmtMin(m){ const x = ((m % 1440) + 1440) % 1440; return `${Math.floor(x / 60)}:${String(x % 60).padStart(2, '0')}`; }
function opSlots(key){
  const s = OP_SHIFTS[key] || OP_SHIFTS.morning;
  const slots = [];
  let t = s.start;
  while(t < s.end){
    const e = Math.min((Math.floor(t / 60) + 1) * 60, s.end);
    let work = e - t, brk = 0;
    const clockHour = Math.floor(t / 60) % 24;
    OP_BREAKS.forEach(b => { if(b.hour === clockHour && t % 60 === 0 && e - t === 60){ brk = b.minutes; work -= b.minutes; } });
    slots.push({ time: `${opFmtMin(t)} \u2192 ${opFmtMin(e)}`, start: t, end: e, work, brk });
    t = e;
  }
  return slots;
}
const OP_I18N = {
  portal:  { fr: 'Portail opérateur', ar: 'بوابة المشغّل' },
  team:    { fr: 'Équipe actuelle', ar: 'الفريق الحالي' },
  saved:   { fr: 'Enregistré', ar: 'تم الحفظ' },
  err:     { fr: 'Erreur', ar: 'خطأ' },
  noTable: { fr: 'Table op_tables introuvable : exécutez le script SQL 13_op_tables.sql dans Supabase', ar: 'جدول op_tables غير موجود: شغّل ملف SQL رقم 13_op_tables.sql في Supabase' },
  add:     { fr: '+ Ajouter un type de rejet', ar: '+ إضافة نوع مرفوض' },
  hint:    { fr: 'Touchez une ligne pour saisir le réel, les rejets et les commentaires.', ar: 'المس أي سطر لإدخال الكمية الفعلية والمرفوضات والتعليقات.' },
  history: { fr: 'Historique', ar: 'السجل' },
  needSql: { fr: "Colonne employee_name absente : exécutez 14_op_employee.sql (le nom de l'opérateur n'est pas enregistré)", ar: 'عمود employee_name غير موجود: شغّل ملف 14_op_employee.sql (اسم المشغّل لن يُحفظ)' },
  needComment: { fr: 'Commentaire obligatoire : le réel est inférieur à la demande', ar: 'التعليق إلزامي: الكمية الفعلية أقل من الطلب' },
  req:     { fr: '(obligatoire)', ar: '(إلزامي)' },
  histSub: { fr: 'Données enregistrées par jour, équipe et heure', ar: 'البيانات المحفوظة حسب اليوم والفريق والساعة' },
  today:   { fr: "Aujourd'hui", ar: 'اليوم' },
  noData:  { fr: 'Aucune donnée pour cette table à cette période', ar: 'لا توجد بيانات لهذه الطاولة في هذه الفترة' },
  readonly:{ fr: 'Lecture seule', ar: 'للقراءة فقط' },
  close:   { fr: 'Fermer', ar: 'إغلاق' },
  loading: { fr: 'Chargement…', ar: 'جارٍ التحميل…' },
  last:    { fr: 'Dernière saisie', ar: 'آخر إدخال' },
  teamTot: { fr: 'Total équipe', ar: 'مجموع الفريق' },
  sMorning:{ fr: 'Matin', ar: 'الصباح' },
  sAfter:  { fr: 'Après-midi', ar: 'المساء' },
  sNight:  { fr: 'Nuit', ar: 'الليل' },
  refDem:  { fr: 'Ref / Demande', ar: 'المرجع / الطلب' },
  refDemSub:{ fr: 'Liste Ref → Demande : en tapant une Ref dans une table de contrôle, la Qtite (demande) se remplit automatiquement', ar: 'قائمة المرجع ← الطلب: عند كتابة المرجع في أي طاولة تُملأ الكمية (Qtite) تلقائياً' },
  addBtn:  { fr: 'Ajouter', ar: 'إضافة' },
  noRefDem:{ fr: 'Aucune référence enregistrée', ar: 'لا توجد مراجع مسجلة' },
  needRefSql:{ fr: 'Table op_ref_demand introuvable : exécutez le script SQL 15_op_ref_demand.sql dans Supabase', ar: 'جدول op_ref_demand غير موجود: شغّل ملف SQL رقم 15_op_ref_demand.sql في Supabase' },
  delConfirm:{ fr: 'Supprimer cette référence ?', ar: 'حذف هذا المرجع؟' },
  tableLbl:{ fr: 'Table', ar: 'الطاولة' },
  refLbl:  { fr: 'Ref', ar: 'المرجع' },
  demLbl:  { fr: 'Demande', ar: 'الطلب' },
  emptyDay:{ fr: 'Aucune donnée enregistrée pour cette table à cette date / équipe', ar: 'لا توجد بيانات مسجلة لهذه الطاولة في هذا اليوم / الفريق' },
  totalLbl:{ fr: 'Total', ar: 'المجموع' },
  exportQc:{ fr: 'Exporter vers Données contrôle', ar: 'تصدير إلى بيانات المراقبة' },
  exportedQc:{ fr: 'fiche(s) exportée(s) vers Données contrôle', ar: 'سجل(ات) تم تصديرها إلى بيانات المراقبة' },
  exportNone:{ fr: 'Rien à exporter pour cette équipe (aucune Ref avec Réel ou Rejet)', ar: 'لا يوجد ما يُصدَّر لهذا الفريق (لا مرجع مع Réel أو Rejet)' },
  needQcSql:{ fr: 'Colonne op_key absente : exécutez le script SQL 18_qc_op_key.sql dans Supabase', ar: 'عمود op_key غير موجود: شغّل ملف SQL رقم 18_qc_op_key.sql في Supabase' }
};
function opT(k){ const o = OP_I18N[k]; return o ? (o[state.lang] || o.fr) : k; }
function opPad(n){ return String(n).padStart(2, '0'); }
function opErrMsg(e){
  const m = (e && e.message) || String(e);
  /* droits refusés (RLS) : ce n'est PAS une table manquante */
  if(/permission denied|row-level security|violates|policy|not allowed|42501/i.test(m) || (e && e.code === '42501')) return 'Accès refusé par Supabase (' + m + ') : exécutez 17_op_tables_policies.sql';
  /* table réellement absente */
  if(/op_tables/i.test(m) && /does not exist|could not find|schema cache|42P01|PGRST205/i.test(m + ' ' + ((e && e.code) || ''))) return opT('noTable');
  return m;
}

function opNowTunis(){
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Africa/Tunis', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
  }).formatToParts(new Date());
  const g = (t) => parts.find(p => p.type === t).value;
  return { y: +g('year'), m: +g('month'), d: +g('day'), hour: (+g('hour')) % 24, minute: +g('minute') };
}
/* Équipe selon l'heure d'ouverture (date = jour où l'équipe a commencé) */
function opCurrentShift(){
  const n = opNowTunis();
  const m = n.hour * 60 + n.minute;
  let key, date = `${n.y}-${opPad(n.m)}-${opPad(n.d)}`;
  if(m >= 21 * 60){ key = 'night'; }
  else if(m >= 13 * 60 + 15){ key = 'afternoon'; }
  else if(m >= 5 * 60 + 30){ key = 'morning'; }
  else{ key = 'night'; date = new Date(Date.UTC(n.y, n.m - 1, n.d - 1)).toISOString().slice(0, 10); }
  return { key, date };
}
function opShiftLabel(key){ return (OP_SHIFTS[key] || OP_SHIFTS.morning).label; }
/* minute courante dans l'échelle de l'équipe (la nuit après minuit = +24h) */
function opShiftNowMin(key){
  const n = opNowTunis();
  let m = n.hour * 60 + n.minute;
  if(key === 'night' && m < OP_SHIFTS.night.start) m += 1440;
  return m;
}
function opBlankRows(key){
  return opSlots(key).map(s => ({ time: s.time, work: s.work, brk: s.brk, demand: 0, real: null, reject: 0, rejects: [], comment: '' }));
}
function opNormalize(rec){
  const base = opBlankRows(rec.shift);
  const src = Array.isArray(rec.hours_data) ? rec.hours_data : [];
  const byTime = {};
  src.forEach(r => { if(r && r.time) byTime[r.time] = r; });
  rec.hours_data = base.map(b => {
    const o = byTime[b.time];
    if(o){ b.real = (o.real === undefined ? null : o.real); b.reject = Number(o.reject) || 0; b.rejects = Array.isArray(o.rejects) ? o.rejects : []; b.comment = o.comment || ''; }
    return b;
  });
  rec.ref = rec.ref || '';
  rec.employee_name = rec.employee_name || '';
  rec.qty_total = Number(rec.qty_total) || 0;
  const parts = opDistribute(rec.qty_total, rec.hours_data.map(r => r.work));
  rec.hours_data.forEach((r, i) => { r.demand = parts[i]; });
  return rec;
}
/* Quantité répartie selon les minutes de travail de chaque ligne (méthode du plus fort reste : la somme reste exacte) */
function opDistribute(total, works){
  const q = Math.max(0, Math.floor(Number(total) || 0));
  const w = (works && works.length) ? works : Array(8).fill(60);
  const sum = w.reduce((s, x) => s + x, 0) || 1;
  const exact = w.map(x => q * x / sum);
  const out = exact.map(Math.floor);
  let rest = q - out.reduce((s, x) => s + x, 0);
  exact.map((v, i) => [v - Math.floor(v), i]).sort((a, b) => b[0] - a[0] || a[1] - b[1]).forEach(([, i]) => { if(rest > 0){ out[i]++; rest--; } });
  return out;
}
function opTotals(rec){
  let demand = 0, real = 0, rej = 0;
  rec.hours_data.forEach(r => { demand += Number(r.demand) || 0; real += Number(r.real) || 0; rej += Number(r.reject) || 0; });
  return { demand, real, rej };
}
function opTypeTotals(rec){
  const m = {};
  OP_TYPES.forEach(tp => { m[tp] = 0; });
  rec.hours_data.forEach(r => (r.rejects || []).forEach(x => { if(m[x.type] !== undefined) m[x.type] += Number(x.qty) || 0; }));
  return m;
}

function opIsCurrent(rec){ const sh = opCurrentShift(); return rec.sheet_date === sh.date && rec.shift === sh.key; }
function opRole(){ return state.currentUser ? state.currentUser.role : ''; }
/* Employé(e) et Entrepreneur : pas de boutons Historique / Ref-Demande ni de retour au menu général */
function opIsRestricted(){ return opRole() === 'employee' || opRole() === 'entrepreneur'; }
function opIsReadOnly(rec){ return opRole() === 'entrepreneur' || (!opIsCurrent(rec) && opRole() !== 'admin'); }
async function opLoadOrCreate(no, date, shiftKey, create){
  const key = { table_no: no, sheet_date: date, shift: shiftKey };
  let res = await sb.from('op_tables').select('*').match(key).maybeSingle();
  if(res.error) throw res.error;
  if(!res.data){
    if(!create) return null;
    const who = state.currentUser ? state.currentUser.name : '';
    const row = Object.assign({}, key, { ref: '', qty_total: 0, hours_data: opBlankRows(shiftKey), created_by: who, updated_by: who, employee_name: who });
    let up = await sb.from('op_tables').upsert(row, { onConflict: 'table_no,sheet_date,shift', ignoreDuplicates: true });
    if(up.error && /employee_name/i.test(up.error.message || '')){
      delete row.employee_name;
      up = await sb.from('op_tables').upsert(row, { onConflict: 'table_no,sheet_date,shift', ignoreDuplicates: true });
    }
    if(up.error) throw up.error;
    res = await sb.from('op_tables').select('*').match(key).single();
    if(res.error) throw res.error;
  }
  return opNormalize(res.data);
}
/* Crée la ligne en base si elle n'existe pas encore (jour sans données ouvert en consultation) */
async function opEnsureRow(r){
  if(r.id) return;
  const who = state.currentUser ? state.currentUser.name : '';
  const key = { table_no: r.table_no, sheet_date: r.sheet_date, shift: r.shift };
  const row = Object.assign({}, key, { ref: r.ref || '', qty_total: r.qty_total || 0, hours_data: r.hours_data, created_by: who, updated_by: who, employee_name: r.employee_name || who });
  let up = await sb.from('op_tables').upsert(row, { onConflict: 'table_no,sheet_date,shift', ignoreDuplicates: true });
  if(up.error && /employee_name/i.test(up.error.message || '')){
    delete row.employee_name;
    up = await sb.from('op_tables').upsert(row, { onConflict: 'table_no,sheet_date,shift', ignoreDuplicates: true });
  }
  if(up.error) throw up.error;
  const res = await sb.from('op_tables').select('id').match(key).single();
  if(res.error) throw res.error;
  r.id = res.data.id;
  r._virtual = false;
}
async function opSaveRec(r, showOk){
  if(!r) return;
  /* jour sans données et rien saisi : on ne crée pas de ligne vide */
  if(r._virtual && !r.ref && !r.qty_total && !r.hours_data.some(h => (h.real !== null && h.real !== undefined) || h.reject || h.comment)) return;
  try{
    await opEnsureRow(r);
    const payload = {
      ref: r.ref || '', qty_total: r.qty_total || 0, hours_data: r.hours_data, employee_name: r.employee_name || '',
      updated_by: state.currentUser ? state.currentUser.name : '', updated_at: new Date().toISOString()
    };
    let { error } = await sb.from('op_tables').update(payload).eq('id', r.id);
    if(error && /employee_name/i.test(error.message || '')){
      delete payload.employee_name;
      ({ error } = await sb.from('op_tables').update(payload).eq('id', r.id));
      if(!error && !state.opEmpWarned){ state.opEmpWarned = true; showToast(opT('needSql'), true); return; }
    }
    if(error) throw error;
    if(showOk) showToast(opT('saved'));
  }catch(e){
    console.error('opPersist error:', e);
    showToast(`${opT('err')}: ${opErrMsg(e)}`, true);
  }
}
async function opPersist(showOk){ return opSaveRec(state.opRec, showOk); }
async function opOpenTable(no, o){
  o = o || {};
  try{
    const sh = opCurrentShift();
    const date = o.date || sh.date, shiftKey = o.shift || sh.key;
    const isCur = date === sh.date && shiftKey === sh.key;
    let rec = await opLoadOrCreate(no, date, shiftKey, isCur && opRole() !== 'entrepreneur');
    if(!rec){
      /* jour / équipe sans données : on affiche quand même la table vide (rien n'est créé tant que rien n'est saisi) */
      rec = opNormalize({ id: null, _virtual: true, table_no: no, sheet_date: date, shift: shiftKey, ref: '', qty_total: 0, hours_data: opBlankRows(shiftKey), employee_name: '' });
    }
    await opLoadRefDem();
    state.opRec = rec;
    state.opFrom = o.from || 'portal';
    state.view = 'op-table';
    render();
    return true;
  }catch(e){
    console.error('opOpenTable error:', e);
    showToast(`${opT('err')}: ${opErrMsg(e)}`, true);
    return false;
  }
}
async function opSwitchTable(n){
  const r = state.opRec;
  if(!r || n === r.table_no) return;
  if(!opIsReadOnly(r)) await opPersist(false);
  await opOpenTable(n, { date: r.sheet_date, shift: r.shift, from: state.opFrom });
}

const OP_CSS = `
  .op-tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:14px;margin-top:6px}
  .op-tile{background:#2f6fc4;color:#fff;border-radius:4px;padding:18px 16px;cursor:pointer;border:1px solid #4a85d6;min-height:96px;display:flex;flex-direction:column;justify-content:space-between;gap:10px}
  .op-tile:hover{background:#3a7bd3}
  .op-tile-name{font-weight:700;font-size:17px;line-height:1.25}
  .op-tile-stat{font-size:12px;opacity:.9}
  .op-head{display:grid;grid-template-columns:repeat(auto-fit,minmax(112px,1fr));gap:8px;margin:0 0 14px}
  .op-hl{font-size:12px;text-align:center;background:var(--panel-alt);border:1px solid var(--border);padding:5px 4px;border-radius:3px;color:var(--text)}
  .op-hv{display:flex;align-items:center;justify-content:center;min-height:58px;font-weight:700;font-size:24px;border-radius:3px;margin-top:4px;overflow:hidden;text-align:center;padding:2px}
  .op-light{background:#fff;color:#1a2230;font-size:15px}
  .op-num{font-size:24px}
  .op-green{background:#fff;color:#1a2230}
  .op-blue{background:#fff;color:#1a2230}
  .op-red{background:#fff;color:#1a2230}
  .op-yellow{background:#fff;color:#1a2230}
  .op-gray{background:#fff;color:#1a2230}
  .op-hv input{width:100%;height:100%;min-height:54px;background:transparent;border:0;outline:0;text-align:center;font:inherit;color:inherit;padding:0 4px}
  .op-hv input::placeholder{color:#8a93a3;font-weight:400}
  .op-scroll{overflow-x:auto}
  .op-table{width:100%;min-width:640px;border-collapse:separate;border-spacing:4px}
  .op-table th{background:#fff;color:#1a2230;font-size:13px;padding:8px 6px;border-radius:2px}
  .op-table td{height:46px;text-align:center;font-weight:700;font-size:17px;border-radius:2px;padding:4px 8px}
  .op-row{cursor:pointer}
  .op-row:hover td{filter:brightness(1.12)}
  .op-table .c-time{background:#fff;color:#1a2230;font-size:14px;white-space:nowrap}
  .op-table .c-prod{background:#fff;color:#1a2230;font-size:14px}
  .op-table .c-dem{background:#fff;color:#1a2230}
  .op-table .c-real{background:#fff;color:#1a2230}
  .op-table .c-rej{background:#fff;color:#1a2230}
  .op-table .c-com{background:var(--panel);border:1px solid var(--border);color:var(--text);font-weight:400;font-size:13px;text-align:start;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .op-total td{background:transparent;color:var(--text);font-size:22px;border-top:2px solid var(--blue);border-radius:0}
  .op-total td:first-child{text-align:start;font-size:20px}
  .op-chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:14px}
  .op-chip{background:#fff;color:#1a2230;border-radius:3px;padding:7px 12px;font-size:13px;display:flex;gap:8px;align-items:center}
  .op-chip b{background:#e6ebf2;border-radius:10px;padding:1px 8px;font-size:13px}
  .op-hint{font-size:12px;color:var(--text-muted);margin:10px 0 0}
  .op-ov{position:fixed;inset:0;background:rgba(0,0,0,.62);z-index:9999;display:flex;align-items:center;justify-content:center;padding:14px}
  .op-modal{background:var(--panel);border:1px solid var(--border);border-radius:6px;padding:18px;width:min(440px,100%);max-height:92vh;overflow:auto;color:var(--text)}
  .op-m-title{text-align:center;font-weight:700;font-size:16px;margin-bottom:14px;padding:8px;border:1px solid var(--border);border-radius:3px;background:var(--panel-alt)}
  .op-m-row{display:grid;grid-template-columns:110px 1fr;gap:10px;align-items:center;margin-bottom:12px}
  .op-m-row label,.op-m-blk label{font-weight:600;font-size:14px;margin:0}
  .op-m-time{background:#fff;color:#1a2230;border-radius:3px;padding:8px;text-align:center;font-weight:700}
  .op-modal input,.op-modal select,.op-modal textarea{width:100%;background:var(--panel-alt);color:var(--text);border:1px solid var(--border);border-radius:3px;padding:9px;font-size:16px;font-family:inherit}
  .op-m-blk{margin-bottom:12px}
  .op-m-blk label{display:block;margin-bottom:6px}
  .op-rj{display:grid;grid-template-columns:1fr 92px 36px;gap:6px;margin-bottom:6px}
  .op-x{background:var(--danger);color:#fff;border:0;border-radius:3px;font-size:18px;cursor:pointer}
  .op-add{background:transparent;color:var(--blue);border:1px dashed var(--blue);border-radius:3px;padding:7px 10px;cursor:pointer;font-size:13px;font-family:inherit}
  .op-m-actions{display:flex;gap:10px;margin-top:6px}
  .op-m-actions .btn{flex:1}
  .op-cancel{background:var(--panel-alt)!important;color:var(--text)!important;border:1px solid var(--border)!important}
  .op-hv{height:58px}
  .op-hv input{min-height:0;height:100%!important;padding:0 4px!important;margin:0!important;line-height:normal}
  #op-ref{font-size:18px}
  .op-modal input,.op-modal select,.op-modal textarea{margin:0}
  @media (max-width:640px){
    .op-head{grid-template-columns:repeat(2,1fr);gap:6px}
    .op-c-qty{grid-column:span 2}
    .op-kpi{grid-template-columns:repeat(3,1fr);gap:6px}
    .op-hl{font-size:11px}
    .op-hv{height:50px;font-size:20px}
    #op-ref{font-size:14px}
    .op-light{font-size:13px}
    .op-table{min-width:0;table-layout:fixed;border-spacing:2px}
    .op-table th{font-size:9px;padding:5px 1px;text-transform:none!important;letter-spacing:0!important;overflow:hidden}
    .op-table td{height:42px;padding:2px;font-size:14px}
    .op-table th:nth-child(1){width:23%} .op-table th:nth-child(2){width:22%} .op-table th:nth-child(3){width:14%}
    .op-table th:nth-child(4){width:14%} .op-table th:nth-child(5){width:12%} .op-table th:nth-child(6){width:15%}
    .op-table .c-time{font-size:11px;white-space:normal}
    .op-table .c-prod{font-size:10px;word-break:break-all}
    .op-table .c-com{font-size:11px;padding:2px 4px}
    .op-total td{font-size:16px}
    .op-total td:first-child{font-size:15px}
  }
  .op-nav{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 12px}
  .op-nb{min-width:46px;padding:9px 12px;background:#2f6fc4;color:#fff;border:1px solid #4a85d6;border-radius:3px;font-weight:700;cursor:pointer;font-family:inherit;font-size:14px}
  .op-nb.on{background:var(--accent);color:#161311;border-color:var(--accent)}
  .op-ro{display:inline-block;background:var(--danger);color:#fff;border-radius:3px;padding:2px 8px;font-size:12px;margin-inline-start:8px}
  .op-hist-bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 18px}
  .op-hist-bar input{background:var(--panel-alt);color:var(--text);border:1px solid var(--border);border-radius:3px;padding:9px;font-size:16px;margin:0;width:auto}
  .op-hist-bar button{background:var(--panel-alt);color:var(--text);border:1px solid var(--border);border-radius:3px;padding:9px 14px;cursor:pointer;font-family:inherit;font-size:15px}
  .op-sec{margin:0 0 24px}
  .op-sec h3{margin:0 0 4px;font-size:16px}
  .op-sec-tot{font-size:13px;color:var(--text-muted);margin:0 0 10px}
  .op-tile.off{background:var(--panel-alt);border-color:var(--border);color:var(--text-muted);cursor:default;opacity:.55}
  .op-tile.off:hover{background:var(--panel-alt)}
  .op-table .c-real.c-short,.op-hv.c-short{background:#e3162c!important;color:#fff!important}
  .op-modal #op-m-dem{background:#fff;color:#1a2230;font-weight:700;text-align:center;opacity:1}
  .op-modal textarea.op-need{border-color:#e3162c;box-shadow:0 0 0 2px rgba(227,22,44,.45)}
  .op-req{color:#ff6b6b;font-weight:700;font-size:12px;margin-inline-start:6px}
  .op-hl{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .op-table td:not(.c-com){border:1px solid #cfd6e2}
  .op-kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:8px;margin:0 0 14px}
  @media (min-width:900px){
    .op-head{grid-template-columns:minmax(210px,1.7fr) minmax(160px,1.3fr) repeat(4,minmax(96px,1fr))}
    .op-kpi{grid-template-columns:repeat(5,1fr)}
  }
`;

const OP2_CSS = `
  .op2{background:#f4f6f9;color:#1d2433;border-radius:6px;overflow:hidden;border:1px solid #d5dbe5;font-size:14px}
  .op2 *{box-sizing:border-box}
  .op2-crumbs{padding:10px 14px;font-size:13px;color:#4b5565;background:#fff;border-bottom:1px solid #e3e8ef;display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  .op2-crumbs a{color:#1d5fb8;cursor:pointer;text-decoration:underline}
  .op2-title{background:#2e333b;color:#fff;font-weight:700;padding:11px 14px;font-size:15px;display:flex;justify-content:space-between;align-items:center;gap:10px}
  .op2-ro{background:#d32f3f;color:#fff;border-radius:3px;padding:2px 9px;font-size:12px;font-weight:600}
  .op2-body{padding:12px}
  .op2-sumwrap{overflow-x:auto;margin-bottom:12px}
  .op2-sum{width:100%;border-collapse:collapse;background:#fff}
  .op2-sum th,.op2-sum td{border:1px solid #d9dee7;padding:8px 10px}
  .op2-sum thead th{background:#eef1f6;font-size:12.5px;text-align:center;white-space:nowrap}
  .op2-sum tbody th{text-align:left;font-weight:600;font-size:15px;background:#fff;white-space:nowrap}
  .op2-sum td{text-align:right;font-size:18px;font-weight:600}
  .op2-sum td.good{background:#1f9d4a;color:#fff}
  .op2-sum td.bad{background:#d32f3f;color:#fff}
  .op2-sum td.warn{background:#f2a516;color:#1c1c1c}
  .op2-bar{display:flex;flex-wrap:wrap;gap:10px 14px;align-items:center;margin:0 0 10px}
  .op2-bar label{font-size:13px;font-weight:600;margin:0}
  .op2-datebar{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
  .op2-bar select,.op2-bar input[type=date]{border:1px solid #b9c1ce;border-radius:3px;padding:7px 9px;background:#fff;color:#1d2433;font-size:14px;margin:0;width:auto;font-family:inherit}
  .op2-btn{border:1px solid #b9c1ce;background:#fff;color:#1d2433;border-radius:3px;padding:7px 12px;cursor:pointer;font-size:14px;font-family:inherit;margin:0;width:auto}
  .op2-btn:hover{background:#eef1f6}
  .op2-btn.primary{background:#1d5fb8;border-color:#1d5fb8;color:#fff;font-weight:600}
  .op2-btn.primary:hover{background:#2a70cf}
  .op2-pills{display:flex;gap:5px;flex-wrap:wrap;margin:0 0 12px}
  .op2-pill{border:1px solid #b9c1ce;background:#fff;color:#1d2433;border-radius:14px;padding:5px 12px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit;margin:0;width:auto}
  .op2-pill:hover{background:#eef1f6}
  .op2-pill.on{background:#1d5fb8;border-color:#1d5fb8;color:#fff}
  .op2-params{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin:0 0 14px;background:#fff;border:1px solid #d9dee7;padding:12px;border-radius:4px}
  .op2-params label{display:block;font-size:12px;font-weight:700;color:#4b5565;margin:0 0 4px}
  .op2-params input{width:100%;margin:0;padding:9px 10px;border:1px solid #b9c1ce;border-radius:3px;background:#fff;color:#1d2433;font-size:15px;font-family:inherit}
  .op2-params input[readonly]{background:#f1f3f7;font-weight:700}
  .op2-params input:disabled{background:#f1f3f7;color:#1d2433;opacity:1}
  .op2-scroll{overflow-x:auto;background:#fff;border:1px solid #d9dee7}
  .op2-table{width:100%;border-collapse:collapse}
  .op2-table th{background:#eef1f6;padding:10px 8px;font-size:13px;border:1px solid #d9dee7;text-align:center;color:#1d2433;text-transform:none;letter-spacing:0}
  .op2-table td{border:1px solid #e1e5ec;padding:6px 8px;text-align:center;font-size:16px;background:#fff;height:46px;color:#1d2433}
  .op2-row{cursor:pointer}
  .op2-row:hover td{background:#f3f7fd}
  .op2-table .c-exp{width:40px;padding:0}
  .op2-table .c-act{width:52px;padding:0}
  .op2-table .c-time{font-weight:700;text-align:left;white-space:nowrap}
  .op2-table .c-time small{display:block;font-weight:600;font-size:11px;color:#1d5fb8}
  .op2-table .c-prod{font-size:14px}
  .op2-table .c-dem{font-weight:600}
  .op2-table .c-real{font-weight:700}
  .op2-table .c-real.c-short{background:#d32f3f;color:#fff}
  .op2-table .c-real.c-miss{color:#b26a00;font-style:italic;font-size:13px;font-weight:400;background:#fff8e6}
  .op2-table .c-com{text-align:left;font-size:13px;max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .op2-table tr.op2-now td{background:#fffbe0;box-shadow:inset 0 2px 0 #f2c200,inset 0 -2px 0 #f2c200}
  .op2-table tr.op2-now td.c-real.c-short{background:#d32f3f}
  .op2-arrow{width:30px;height:30px;border:0;background:transparent;font-size:22px;cursor:pointer;transition:transform .15s;color:#4b5565;margin:0;padding:0;width:auto}
  .op2-arrow.open{transform:rotate(90deg)}
  .op2-edit{border:1px solid #b9c1ce;background:#fff;color:#1d2433;border-radius:3px;width:36px;height:32px;cursor:pointer;font-size:16px;margin:0;padding:0}
  .op2-edit:hover{background:#e8effa}
  .op2-det td{background:#f7f9fc;text-align:left;font-size:13px;padding:10px 14px;height:auto}
  .op2-total td{background:#eef1f6;font-weight:700;font-size:17px;height:44px}
  .op2-total td.c-short{background:#d32f3f;color:#fff}
  .op2-types{display:flex;flex-wrap:wrap;gap:6px;margin-top:12px}
  .op2-tchip{background:#fff;border:1px solid #c6cdd9;border-radius:14px;padding:4px 11px;font-size:13px;display:inline-flex;gap:7px;align-items:center;color:#1d2433;margin:0 4px 4px 0}
  .op2-tchip b{background:#e6ebf2;border-radius:10px;padding:0 7px}
  .op2-table,.op2-sum{min-width:0!important}
  .op-hf{align-items:center}
  .op-hf label{margin:0;font-size:13px;font-weight:600}
  .op-hf input,.op-hf select{width:auto;margin:0}
  .op-ht{min-width:860px}
  .op-ht td.num{text-align:right}
  .op-ht td.bad{color:#d32f3f;font-weight:700}
  .op-ht td.good{color:#1f9d4a;font-weight:700}
  .op-ht td.warn{color:#b26a00;font-weight:700}
  .op-hrow{cursor:pointer}
  .op-hrow:hover td{background:var(--panel-alt)}
  .op-ht tfoot td{font-weight:700;background:var(--panel-alt);border-top:2px solid var(--border);border-bottom:none}
  .op2-params select{width:100%;margin:0;padding:9px 10px;border:1px solid #b9c1ce;border-radius:3px;background:#fff;color:#1d2433;font-size:15px;font-weight:700;font-family:inherit}
  .op2-params #op-ref{font-size:15px}
  .op2-sum thead th,.op2-sum tbody th{text-transform:none;letter-spacing:0}
  .op2-sum tbody th{color:#1d2433}
  .op-modal{background:#fff;color:#1d2433;border-color:#c6cdd9}
  .op-m-title{background:#eef1f6;border-color:#d0d6e0;color:#1d2433}
  .op-modal input,.op-modal select,.op-modal textarea{background:#fff;color:#1d2433;border:1px solid #b9c1ce}
  .op-modal #op-m-dem{background:#f1f3f7;color:#1d2433}
  .op-m-time{background:#f1f3f7}
  .op-cancel{background:#fff!important;color:#1d2433!important;border:1px solid #b9c1ce!important}
  @media (max-width:640px){
    .op2-body{padding:8px}
    .op2 .col-com{display:none}
    .op2-table td{padding:4px 3px;font-size:14px;height:44px}
    .op2-table th{font-size:11px;padding:7px 3px}
    .op2-table .c-time{font-size:12px;white-space:normal}
    .op2-table .c-prod{font-size:11px;word-break:break-all}
    .op2-table .c-exp{width:28px}
    .op2-table .c-act{width:40px}
    .op2-edit{width:32px}
    .op2-sum th,.op2-sum td{padding:5px 3px}
    .op2-sum td{font-size:14px}
    .op2-sum thead th{font-size:10px;white-space:normal}
    .op2-sum tbody th{font-size:12px;white-space:normal}
    .op2-table th{white-space:normal}
    .op2-total td{font-size:15px}
  }
`;

function renderOpPortal(main){
  const sh = opCurrentShift();
  main.innerHTML = `<style>${OP_CSS}</style>
    ${opIsRestricted() ? '' : `<div class="back-link" id="op-back" style="margin-bottom:14px;">${ICONS.back} ${t('menuTitle')}</div>`}
    <div class="menu-head"><h1>${opT('portal')}</h1><p>${opT('team')} : <strong>${opShiftLabel(sh.key)}</strong> &mdash; ${sh.date}</p></div>
    ${opIsRestricted() ? '' : `<div class="op-nav" style="margin-bottom:16px;"><button type="button" class="op-nb" id="op-hist" style="padding:11px 18px;">&#128197; ${opT('history')}</button> <button type="button" class="op-nb" id="op-refdem" style="padding:11px 18px;">&#9636; ${opT('refDem')}</button> <button type="button" class="op-nb" id="op-export" style="padding:11px 18px;">&#8680; ${opT('exportQc')}</button></div>`}
    <div class="op-tiles">${Array.from({ length: 8 }, (_, i) => `
      <div class="op-tile" data-no="${i + 1}"><div class="op-tile-name">Table de contr\u00f4le ${opPad(i + 1)}</div><div class="op-tile-stat" id="op-stat-${i + 1}">&nbsp;</div></div>`).join('')}
    </div>`;
  const opBackEl = document.getElementById('op-back'), opHistEl = document.getElementById('op-hist');
  if(opBackEl) opBackEl.onclick = () => { state.view = 'menu'; render(); };
  if(opHistEl) opHistEl.onclick = () => { state.opHistDate = null; state.view = 'op-history'; render(); };
  main.querySelectorAll('.op-tile').forEach(el => { el.onclick = () => opOpenTable(+el.dataset.no, { from: 'portal' }); });
  const opRdEl = document.getElementById('op-refdem');
  if(opRdEl) opRdEl.onclick = () => opOpenRefDemand();
  const opExEl = document.getElementById('op-export');
  if(opExEl){ opExEl.onclick = () => opExportCurrentShift(); opAutoExport(); }
  (async () => {
    try{
      const { data, error } = await sb.from('op_tables').select('table_no, qty_total, hours_data').eq('sheet_date', sh.date).eq('shift', sh.key);
      if(error) throw error;
      (data || []).forEach(r => {
        const el = document.getElementById('op-stat-' + r.table_no);
        if(!el) return;
        let real = 0, rej = 0;
        (r.hours_data || []).forEach(x => { real += Number(x.real) || 0; rej += Number(x.reject) || 0; });
        el.textContent = `R\u00e9el ${real} / Qtite ${Number(r.qty_total) || 0} \u00b7 Rejet ${rej}`;
      });
    }catch(e){
      console.error('op overview error:', e);
      showToast(`${opT('err')}: ${opErrMsg(e)}`, true);
    }
  })();
}


/* ===== Liste Ref -> Demande : sert au remplissage automatique de Qtite dans les tables de contrôle ===== */
function opRefKey(r){ return (r || '').toString().trim().toLowerCase(); }
async function opLoadRefDem(){
  try{
    const { data, error } = await sb.from('op_ref_demand').select('ref, demande');
    if(error) throw error;
    state.opRefDem = {};
    (data || []).forEach(x => { state.opRefDem[opRefKey(x.ref)] = { ref: x.ref, demande: Number(x.demande) || 0 }; });
    return true;
  }catch(e){
    console.error('op ref_demand error:', e);
    state.opRefDem = state.opRefDem || {};
    return false;
  }
}
async function opOpenRefDemand(){
  const okLoad = await opLoadRefDem();
  if(!okLoad){ showToast(opT('needRefSql'), true); return; }
  const ov = document.createElement('div');
  ov.className = 'op-ov';
  const th = 'padding:6px;border:1px solid #d9dee7;background:#eef1f6;font-size:13px';
  const td = 'padding:3px;border:1px solid #e1e5ec';
  ov.innerHTML = `<div class="op-modal" style="width:min(520px,100%)">
    <div class="op-m-title">${opT('refDem')}</div>
    <p style="font-size:12px;color:#5b6575;margin:0 0 10px">${opT('refDemSub')}</p>
    <table style="width:100%;border-collapse:collapse">
      <thead><tr><th style="${th}">${opT('refLbl')}</th><th style="${th}">${opT('demLbl')}</th><th style="${th};width:44px"></th></tr></thead>
      <tbody id="rd-body"></tbody>
      <tfoot><tr>
        <td style="${td}"><input type="text" id="rd-new-ref" autocomplete="off" placeholder="ref" style="text-align:center"></td>
        <td style="${td}"><input type="number" id="rd-new-dem" min="0" step="1" inputmode="numeric" placeholder="0" style="text-align:center;font-weight:700"></td>
        <td style="${td}"><button type="button" class="btn" id="rd-add" style="width:100%;padding:8px 4px">+</button></td>
      </tr></tfoot>
    </table>
    <div class="op-m-actions" style="margin-top:12px"><button type="button" class="btn op-cancel" id="rd-close">${opT('close')}</button></div>
  </div>`;
  document.body.appendChild(ov);
  const body = ov.querySelector('#rd-body');
  const toInt = (v) => Math.max(0, Math.floor(Number(v) || 0));
  function draw(){
    const list = Object.values(state.opRefDem).sort((a, b) => a.ref.localeCompare(b.ref));
    body.innerHTML = list.length ? list.map(x => `
      <tr><td style="${td};text-align:center;font-weight:700">${escHtml(x.ref)}</td>
      <td style="${td}"><input type="number" class="rd-dem" data-ref="${escHtml(x.ref)}" min="0" step="1" inputmode="numeric" value="${x.demande}" style="text-align:center;font-weight:700"></td>
      <td style="${td};text-align:center"><button type="button" class="op-x rd-del" data-ref="${escHtml(x.ref)}" style="width:34px;height:34px">&times;</button></td></tr>`).join('')
      : `<tr><td colspan="3" style="padding:12px;text-align:center;color:#5b6575;font-size:13px">${opT('noRefDem')}</td></tr>`;
    body.querySelectorAll('.rd-dem').forEach(el => {
      el.onchange = async () => {
        const ref = el.dataset.ref, v = toInt(el.value);
        const { error } = await sb.from('op_ref_demand').update({ demande: v, updated_at: new Date().toISOString() }).eq('ref', ref);
        if(error){ showToast(`${opT('err')}: ${error.message}`, true); return; }
        state.opRefDem[opRefKey(ref)].demande = v;
        showToast(opT('saved'));
      };
    });
    body.querySelectorAll('.rd-del').forEach(el => {
      el.onclick = async () => {
        if(!confirm(opT('delConfirm'))) return;
        const ref = el.dataset.ref;
        const { error } = await sb.from('op_ref_demand').delete().eq('ref', ref);
        if(error){ showToast(`${opT('err')}: ${error.message}`, true); return; }
        delete state.opRefDem[opRefKey(ref)];
        draw();
      };
    });
  }
  draw();
  ov.querySelector('#rd-add').onclick = async () => {
    const refEl = ov.querySelector('#rd-new-ref'), demEl = ov.querySelector('#rd-new-dem');
    const ref = refEl.value.trim(), v = toInt(demEl.value);
    if(!ref){ refEl.focus(); return; }
    const key = opRefKey(ref);
    const exists = state.opRefDem[key];
    const row = { ref: exists ? exists.ref : ref, demande: v, updated_at: new Date().toISOString() };
    const { error } = await sb.from('op_ref_demand').upsert(row, { onConflict: 'ref' });
    if(error){ showToast(`${opT('err')}: ${error.message}`, true); return; }
    state.opRefDem[key] = { ref: row.ref, demande: v };
    refEl.value = ''; demEl.value = '';
    showToast(opT('saved'));
    draw();
    refEl.focus();
  };
  const onKey = (e) => { if(e.key === 'Escape') close(); };
  function close(){ document.removeEventListener('keydown', onKey); ov.remove(); }
  document.addEventListener('keydown', onKey);
  ov.querySelector('#rd-close').onclick = close;
  ov.addEventListener('mousedown', (e) => { if(e.target === ov) close(); });
}

/* ===== Export d'une équipe du portail opérateur vers Données contrôle (qc_records) ===== */
function opNorm(x){ return (x || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim(); }
function opShiftEnded(date, key){
  const n = opNowTunis();
  const [y, m, d] = date.split('-').map(Number);
  const endMs = Date.UTC(y, m - 1, d) + OP_SHIFTS[key].end * 60000;
  const nowMs = Date.UTC(n.y, n.m - 1, n.d) + (n.hour * 60 + n.minute) * 60000;
  return nowMs >= endMs;
}
/* une fiche par table (Ref, Qtité contrôle = Réel, Non OK = Rejet, défauts par type) ; ré-export = mise à jour, pas de doublon */
async function opExportShiftToQc(date, key){
  const { data, error } = await sb.from('op_tables').select('*').eq('sheet_date', date).eq('shift', key);
  if(error) throw error;
  if(!state.catalog || !state.catalog.length) await loadCatalog();
  const rows = [];
  (data || []).forEach(r => {
    const rec = opNormalize(Object.assign({}, r));
    const ref = (rec.ref || '').trim();
    const tot = opTotals(rec);
    if(!ref || (tot.real <= 0 && tot.rej <= 0)) return;
    const cat = (state.catalog || []).find(c => opNorm(c.ref) === opNorm(ref));
    const ty = opTypeTotals(rec);
    const row = {
      record_date: date, fournisseur: cat ? (cat.fournisseur || '') : '', ref: ref, designation: cat ? (cat.designation || '') : '',
      qte_controle: tot.real, qte_non_ok: tot.rej, grani: 0, rayure: 0, trace: 0, piqure: 0, coup: 0, cosse: 0, autre: 0,
      user_name: rec.employee_name || 'Portail opérateur', op_key: `${date}|${key}|${r.table_no}`
    };
    OP_TYPES.forEach(tp => { const k = opNorm(tp); if(k in row) row[k] = ty[tp] || 0; });
    rows.push(row);
  });
  if(!rows.length) return 0;
  const up = await sb.from('qc_records').upsert(rows, { onConflict: 'op_key' });
  if(up.error) throw up.error;
  return rows.length;
}
async function opExportCurrentShift(){
  const sh = opCurrentShift();
  try{
    const n = await opExportShiftToQc(sh.date, sh.key);
    showToast(n ? `${n} ${opT('exportedQc')}` : opT('exportNone'), !n);
  }catch(e){
    console.error('op export error:', e);
    showToast(/op_key|PGRST204|42703/i.test((e && e.message) || '') ? opT('needQcSql') : `${opT('err')}: ${opErrMsg(e)}`, true);
  }
}
/* export automatique des équipes terminées (aujourd'hui et hier) qui n'ont pas encore été exportées ; se lance à l'ouverture du portail */
async function opAutoExport(){
  if(state.opAutoDone) return;
  state.opAutoDone = true;
  try{
    const n = opNowTunis();
    const today = `${n.y}-${opPad(n.m)}-${opPad(n.d)}`;
    const yest = new Date(Date.UTC(n.y, n.m - 1, n.d - 1)).toISOString().slice(0, 10);
    const cands = [];
    [yest, today].forEach(d => ['morning', 'afternoon', 'night'].forEach(k => { if(opShiftEnded(d, k)) cands.push({ d, k }); }));
    if(!cands.length) return;
    const ex = await sb.from('qc_records').select('op_key').in('record_date', [yest, today]).not('op_key', 'is', null);
    if(ex.error) return;   /* colonne absente : pas d'export automatique */
    const done = new Set((ex.data || []).map(x => x.op_key.split('|').slice(0, 2).join('|')));
    for(const c of cands){ if(!done.has(`${c.d}|${c.k}`)) await opExportShiftToQc(c.d, c.k); }
  }catch(e){ console.error('op auto export error:', e); }
}

const OP_REJ_OK = 2, OP_REJ_WARN = 5;   /* taux de rejet : <=2% vert, <=5% orange, au-dela rouge */
function opPct(a, b){ return b > 0 ? (Math.round(a / b * 1000) / 10) + ' %' : '\u2014'; }
function opAgg(lists){
  let dem = 0, real = 0, demE = 0, rej = 0, n = 0;
  lists.forEach(hd => (hd || []).forEach(r => {
    const d = Number(r.demand) || 0;
    dem += d; rej += Number(r.reject) || 0;
    if(r.real !== null && r.real !== undefined){ real += Number(r.real) || 0; demE += d; n++; }
  }));
  return { dem, real, av: real - demE, rej, n };
}
function opWeekRange(iso){
  const [y, m, d] = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  const dow = (dt.getUTCDay() + 6) % 7;
  const mon = new Date(Date.UTC(y, m - 1, d - dow));
  const sun = new Date(Date.UTC(y, m - 1, d - dow + 6));
  return [mon.toISOString().slice(0, 10), sun.toISOString().slice(0, 10)];
}
async function opLoadWeek(rec){
  const [mon, sun] = opWeekRange(rec.sheet_date);
  try{
    const { data, error } = await sb.from('op_tables').select('id, sheet_date, shift, hours_data')
      .eq('table_no', rec.table_no).gte('sheet_date', mon).lte('sheet_date', sun);
    if(error) throw error;
    if(state.opRec !== rec) return;
    state.opWeek = data || [];
  }catch(e){
    console.error('op week error:', e);
    if(state.opRec !== rec) return;
    state.opWeek = [];
  }
  opRefreshCells();
}
async function opNavigate(date, shift){
  const r = state.opRec;
  if(!r || !date || !shift) return;
  if(!opIsReadOnly(r)) await opPersist(false);
  const ok = await opOpenTable(r.table_no, { date, shift, from: state.opFrom });
  if(!ok) render();
}

function opFillSum(row, a, has){
  const set = (col, txt, cls) => {
    const el = document.getElementById(`op-s-${row}-${col}`);
    if(!el) return;
    el.textContent = txt;
    el.className = cls || '';
  };
  if(!a){ ['d', 'r', 'p', 'a', 'j', 't'].forEach(c => set(c, '\u2026', '')); return; }
  set('d', a.dem, '');
  set('r', a.real, '');
  set('p', opPct(a.real, a.dem), '');
  set('a', a.n ? (a.av > 0 ? '+' + a.av : String(a.av)) : '\u2014', a.n ? (a.av < 0 ? 'bad' : 'good') : '');
  set('j', a.rej, '');
  const rate = a.real > 0 ? a.rej / a.real * 100 : null;
  set('t', rate === null ? '\u2014' : (Math.round(rate * 100) / 100).toFixed(2) + '%', rate === null ? '' : (rate <= OP_REJ_OK ? 'good' : rate <= OP_REJ_WARN ? 'warn' : 'bad'));
}

function opRefreshCells(){
  const rec = state.opRec;
  if(!rec) return;
  const set = (id, v) => { const el = document.getElementById(id); if(el) el.textContent = (v === null || v === undefined) ? '' : String(v); };
  const cur = opIsCurrent(rec);
  const slots = opSlots(rec.shift);
  const nowMin = cur ? opShiftNowMin(rec.shift) : -1;
  let nowIdx = -1;
  if(cur) slots.forEach((s, i) => { if(nowMin >= s.start && nowMin < s.end) nowIdx = i; });
  const tot = opTotals(rec);
  rec.hours_data.forEach((r, i) => {
    set('op-p' + i, rec.ref || '');
    set('op-d' + i, r.demand);
    const empty = (r.real === null || r.real === undefined);
    const missing = empty && cur && !!slots[i] && slots[i].end <= nowMin;
    const ra = document.getElementById('op-a' + i);
    if(ra){
      ra.textContent = empty ? (missing ? '\u00e0 saisir' : '') : String(r.real);
      ra.classList.toggle('c-short', !empty && (Number(r.real) || 0) < (Number(r.demand) || 0));
      ra.classList.toggle('c-miss', missing);
    }
    set('op-j' + i, r.reject ? r.reject : '');
    set('op-c' + i, r.comment || '');
    const rj = (r.rejects || []).map(x => `<span class="op2-tchip">${escHtml(x.type)} <b>${Number(x.qty) || 0}</b></span>`).join('');
    const det = document.getElementById('op-x' + i);
    if(det) det.innerHTML = `<div><b>Rejets :</b> ${rj || '\u2014'}</div><div style="margin-top:6px"><b>Commentaire :</b> ${r.comment ? escHtml(r.comment) : '\u2014'}</div>`;
    const tr = document.querySelector(`.op2-row[data-i="${i}"]`);
    if(tr) tr.classList.toggle('op2-now', i === nowIdx);
    const tag = document.getElementById('op-tag' + i);
    if(tag){
      const note = r.brk ? `${r.work} min (pause ${r.brk})` : (r.work < 60 ? `${r.work} min` : '');
      tag.textContent = [note, i === nowIdx ? 'heure en cours' : ''].filter(Boolean).join(' \u00b7 ');
    }
  });
  set('op-t-dem', tot.demand); set('op-t-real', tot.real); set('op-t-rej', tot.rej);
  let demEntered = 0;
  rec.hours_data.forEach(r => { if(r.real !== null && r.real !== undefined) demEntered += Number(r.demand) || 0; });
  const tr = document.getElementById('op-t-real');
  if(tr) tr.classList.toggle('c-short', tot.real < demEntered);
  const byType = opTypeTotals(rec);
  OP_TYPES.forEach((tp, k) => set('op-chip-' + k, byType[tp]));
  const tt = document.getElementById('op-title-ref');
  if(tt) tt.textContent = rec.ref ? ' \u2014 ' + rec.ref : '';
  /* Line total */
  opFillSum('eq', opAgg([rec.hours_data]));
  const wk = state.opWeek;
  if(wk){
    const others = wk.filter(x => !(x.sheet_date === rec.sheet_date && x.shift === rec.shift));
    opFillSum('day', opAgg([rec.hours_data].concat(others.filter(x => x.sheet_date === rec.sheet_date).map(x => x.hours_data))));
    opFillSum('wk', opAgg([rec.hours_data].concat(others.map(x => x.hours_data))));
  }else{
    opFillSum('day', null); opFillSum('wk', null);
  }
}

const OP_TEAMS = [
  { key: 'morning',   label: 'F1. \u00c9quipe Matin (5h30 \u2192 13h15)' },
  { key: 'afternoon', label: 'F2. \u00c9quipe Apr\u00e8s-midi (13h15 \u2192 21h00)' },
  { key: 'night',     label: 'F3. \u00c9quipe Nuit (21h00 \u2192 5h30)' }
];
function opCurrentSlotIdx(rec){
  const slots = opSlots(rec.shift), m = opShiftNowMin(rec.shift);
  let idx = -1;
  slots.forEach((s, i) => { if(m >= s.start && m < s.end) idx = i; });
  return idx;
}

function renderOpTable(main){
  const rec = state.opRec;
  if(!rec){ state.view = 'op-portal'; return render(); }
  const ro = opIsReadOnly(rec);
  const cur = opIsCurrent(rec);
  const fromHist = state.opFrom === 'history';
  const nameTxt = `Table de contr\u00f4le ${opPad(rec.table_no)}`;
  state.opWeek = null;
  state.opOpen = state.opOpen || {};
  main.innerHTML = `<style>${OP_CSS}${OP2_CSS}</style>
  <div class="op2">
    <div class="op2-crumbs"><a id="op-home">Accueil</a><span>\u203a</span><a id="op-back">${fromHist ? opT('history') : opT('portal')}</a><span>\u203a</span><b>${nameTxt}</b></div>
    <div class="op2-title"><span>Portail op\u00e9rateur : ${nameTxt}<span id="op-title-ref"></span></span>${ro ? `<span class="op2-ro">${opT('readonly')}</span>` : ''}</div>
    <div class="op2-body">
      ${rec._virtual ? `<div style="background:#fff8e6;border:1px solid #f2c200;color:#6b4e00;border-radius:4px;padding:8px 12px;margin-bottom:10px;font-size:13px">${opT('emptyDay')}</div>` : ''}
      <div class="op2-sumwrap"><table class="op2-sum">
        <thead><tr><th>Line total</th><th>Demande</th><th>R\u00e9el</th><th>% R\u00e9el / Demande</th><th title="Avance (+) / Retard (\u2212) sur les heures saisies">AV / RE</th><th>Rejet</th><th>Taux de rejet %</th></tr></thead>
        <tbody>
          ${[['eq', cur ? '\u00c9quipe actuelle' : '\u00c9quipe affich\u00e9e'], ['day', 'Cumul jour'], ['wk', 'Cumul semaine']].map(([k, lab]) => `
          <tr><th>${lab}</th>${['d', 'r', 'p', 'a', 'j', 't'].map(c => `<td id="op-s-${k}-${c}">\u2026</td>`).join('')}</tr>`).join('')}
        </tbody>
      </table></div>

      <div class="op2-bar">
        <label for="op-team">\u00c9quipe</label>
        <select id="op-team">${OP_TEAMS.map(t => `<option value="${t.key}" ${t.key === rec.shift ? 'selected' : ''}>${t.label}</option>`).join('')}</select>
        <div class="op2-datebar">
          <input type="date" id="op-ddate" value="${escHtml(rec.sheet_date)}">
          <button type="button" class="op2-btn" id="op-dtoday">Aujourd'hui</button>
          <button type="button" class="op2-btn" id="op-dprev" aria-label="Jour pr\u00e9c\u00e9dent">&lsaquo;</button>
          <button type="button" class="op2-btn" id="op-dnext" aria-label="Jour suivant">&rsaquo;</button>
        </div>
        ${(!ro && cur) ? '<button type="button" class="op2-btn primary" id="op-now-btn">\u270e Saisir l\'heure en cours</button>' : ''}
      </div>

      <div class="op2-params">
        <div><label for="op-name">Table</label><select id="op-name">${Array.from({ length: 8 }, (_, i) => `<option value="${i + 1}" ${rec.table_no === i + 1 ? 'selected' : ''}>Table de contr\u00f4le ${opPad(i + 1)}</option>`).join('')}</select></div>
        <div><label for="op-emp">Op\u00e9rateur</label><input type="text" id="op-emp" autocomplete="off" placeholder="Nom de l'op\u00e9rateur"></div>
        <div><label for="op-ref">R\u00e9f\u00e9rence (ref)</label><input type="text" id="op-ref" autocomplete="off" placeholder="ref" list="op-ref-list"><datalist id="op-ref-list">${Object.values(state.opRefDem || {}).map(x => `<option value="${escHtml(x.ref)}"></option>`).join('')}</datalist></div>
        <div><label for="op-qty">Qtite (demande)</label><input type="number" id="op-qty" min="0" step="1" inputmode="numeric" placeholder="0"></div>
      </div>

      <div class="op2-scroll"><table class="op2-table">
        <thead><tr><th class="c-exp"></th><th>Temps</th><th>Produit actuel</th><th>Demande</th><th>R\u00e9el</th><th>Rejet</th><th class="col-com">Commentaires / Contre-mesures</th><th class="c-act"></th></tr></thead>
        <tbody>${rec.hours_data.map((r, i) => `
          <tr class="op2-row" data-i="${i}">
            <td class="c-exp"><button type="button" class="op2-arrow${state.opOpen[i] ? ' open' : ''}" data-i="${i}" aria-label="D\u00e9tails">&rsaquo;</button></td>
            <td class="c-time">${escHtml(r.time)}<small id="op-tag${i}"></small></td>
            <td class="c-prod" id="op-p${i}"></td><td class="c-dem" id="op-d${i}"></td>
            <td class="c-real" id="op-a${i}"></td><td class="c-rej" id="op-j${i}"></td>
            <td class="c-com col-com" id="op-c${i}"></td>
            <td class="c-act">${ro ? '' : `<button type="button" class="op2-edit" data-i="${i}" aria-label="Saisir">\u270e</button>`}</td>
          </tr>
          <tr class="op2-det" id="op-xr${i}" style="display:${state.opOpen[i] ? 'table-row' : 'none'}"><td colspan="8"><div id="op-x${i}"></div></td></tr>`).join('')}
        </tbody>
        <tfoot><tr class="op2-total"><td></td><td style="text-align:left">Total</td><td>\u2014</td><td id="op-t-dem">0</td><td id="op-t-real">0</td><td id="op-t-rej">0</td><td class="col-com"></td><td></td></tr></tfoot>
      </table></div>
      <div class="op2-types">${OP_TYPES.map((tp, k) => `<span class="op2-tchip">${tp} <b id="op-chip-${k}">0</b></span>`).join('')}</div>
      <p class="op-hint" style="color:#5b6575;margin:10px 0 0">${ro ? '' : opT('hint')}</p>
    </div>
  </div>`;

  const refEl = document.getElementById('op-ref'), qtyEl = document.getElementById('op-qty'), empEl = document.getElementById('op-emp');
  refEl.value = rec.ref || '';
  qtyEl.value = rec.qty_total ? String(rec.qty_total) : '';
  let needSave = false;
  if(!rec.employee_name && !ro){ rec.employee_name = state.currentUser ? state.currentUser.name : ''; needSave = !!rec.employee_name; }
  empEl.value = rec.employee_name || '';
  if(ro){ refEl.disabled = true; qtyEl.disabled = true; empEl.disabled = true; }
  empEl.oninput = () => { if(!ro) rec.employee_name = empEl.value.trim(); };
  empEl.onchange = () => { if(!ro) opPersist(false); };
  refEl.oninput = () => {
    if(ro) return;
    rec.ref = refEl.value.trim();
    /* Ref connue dans la liste Ref / Demande : la Qtite (demande) se remplit automatiquement */
    const hit = (state.opRefDem || {})[opRefKey(rec.ref)];
    if(hit){
      rec.qty_total = hit.demande;
      qtyEl.value = hit.demande ? String(hit.demande) : '';
      const parts = opDistribute(rec.qty_total, rec.hours_data.map(r => r.work));
      rec.hours_data.forEach((r, i) => { r.demand = parts[i]; });
    }
    opRefreshCells();
  };
  refEl.onchange = () => { if(!ro) opPersist(false); };
  qtyEl.oninput = () => {
    if(ro) return;
    rec.qty_total = Math.max(0, Math.floor(Number(qtyEl.value) || 0));
    const parts = opDistribute(rec.qty_total, rec.hours_data.map(r => r.work));
    rec.hours_data.forEach((r, i) => { r.demand = parts[i]; });
    opRefreshCells();
  };
  qtyEl.onchange = () => { if(!ro) opPersist(false); };
  if(needSave) opPersist(false);

  /* lignes : arrow = d\u00e9tails, crayon / ligne = saisie */
  main.querySelectorAll('.op2-row').forEach(tr => { tr.onclick = () => opOpenPopup(+tr.dataset.i, ro); });
  main.querySelectorAll('.op2-arrow').forEach(b => {
    b.onclick = (e) => {
      e.stopPropagation();
      const i = +b.dataset.i;
      state.opOpen[i] = !state.opOpen[i];
      b.classList.toggle('open', !!state.opOpen[i]);
      document.getElementById('op-xr' + i).style.display = state.opOpen[i] ? 'table-row' : 'none';
    };
  });
  main.querySelectorAll('.op2-edit').forEach(b => { b.onclick = (e) => { e.stopPropagation(); opOpenPopup(+b.dataset.i, ro); }; });
  const nowBtn = document.getElementById('op-now-btn');
  if(nowBtn) nowBtn.onclick = () => opOpenPopup(Math.max(0, opCurrentSlotIdx(rec)), ro);

  /* navigation : \u00e9quipe / date / table */
  document.getElementById('op-team').onchange = (e) => opNavigate(rec.sheet_date, e.target.value);
  document.getElementById('op-ddate').onchange = (e) => opNavigate(e.target.value, rec.shift);
  document.getElementById('op-dprev').onclick = () => opNavigate(opShiftDate(rec.sheet_date, -1), rec.shift);
  document.getElementById('op-dnext').onclick = () => opNavigate(opShiftDate(rec.sheet_date, 1), rec.shift);
  document.getElementById('op-dtoday').onclick = () => { const sh = opCurrentShift(); opNavigate(sh.date, sh.key); };
  document.getElementById('op-name').onchange = (e) => opSwitchTable(+e.target.value);
  document.getElementById('op-home').onclick = async () => { if(!ro) await opPersist(false); state.opRec = null; state.view = 'menu'; render(); };
  document.getElementById('op-back').onclick = async () => {
    if(!ro) await opPersist(false);
    state.opRec = null;
    state.view = fromHist ? 'op-history' : 'op-portal';
    render();
  };
  opRefreshCells();
  opLoadWeek(rec);
}

/* Tableau de remplissage (fen\u00eatre surgissante) */
function opOpenPopup(i, ro){
  const rec = state.opRec, row = rec.hours_data[i];
  const ov = document.createElement('div');
  ov.className = 'op-ov';
  ov.innerHTML = `<div class="op-modal">
    <div class="op-m-title">Tableau de remplissage</div>
    <div class="op-m-row"><label>Temps :</label><div class="op-m-time" id="op-m-time"></div></div>
    <div class="op-m-row"><label>Demande :</label><input id="op-m-dem" type="number" readonly tabindex="-1"></div>
    <div class="op-m-row"><label>R\u00e9el :</label><input id="op-m-real" type="number" min="0" step="1" inputmode="numeric"></div>
    <div class="op-m-blk"><label>Type de rejet :</label><div id="op-m-lines"></div><button type="button" class="op-add" id="op-m-add">${opT('add')}</button></div>
    <div class="op-m-blk"><label>Commentaires :<span class="op-req" id="op-m-comreq" style="display:none">${opT('req')}</span></label><textarea id="op-m-com" rows="3"></textarea></div>
    <div class="op-m-actions"><button type="button" class="btn" id="op-m-save">Enregistrer</button><button type="button" class="btn op-cancel" id="op-m-cancel">Annuler</button></div>
  </div>`;
  document.body.appendChild(ov);
  const $ = (id) => ov.querySelector('#' + id);
  $('op-m-time').textContent = row.time;
  const demand = Number(row.demand) || 0;
  $('op-m-dem').value = String(demand);
  $('op-m-real').value = (row.real === null || row.real === undefined) ? '' : String(row.real);
  $('op-m-com').value = row.comment || '';
  /* r\u00e9el < demande  =>  commentaire obligatoire */
  const needComment = () => { const v = $('op-m-real').value; return v !== '' && (Number(v) || 0) < demand; };
  const refreshNeed = () => {
    const need = needComment();
    $('op-m-comreq').style.display = need ? 'inline' : 'none';
    $('op-m-com').classList.toggle('op-need', need && !$('op-m-com').value.trim());
  };
  $('op-m-real').oninput = refreshNeed;
  $('op-m-com').oninput = refreshNeed;
  refreshNeed();
  const lines = [], box = $('op-m-lines');
  function addLine(type, qty){
    const wrap = document.createElement('div'); wrap.className = 'op-rj';
    const sel = document.createElement('select');
    OP_TYPES.forEach(tp => { const o = document.createElement('option'); o.value = tp; o.textContent = tp; sel.appendChild(o); });
    sel.value = OP_TYPES.indexOf(type) >= 0 ? type : OP_TYPES[0];
    const q = document.createElement('input'); q.type = 'number'; q.min = '0'; q.step = '1'; q.inputMode = 'numeric'; q.placeholder = '0';
    q.value = qty ? String(qty) : '';
    const x = document.createElement('button'); x.type = 'button'; x.className = 'op-x'; x.textContent = '\u00d7';
    const item = { wrap, sel, q };
    x.onclick = () => { wrap.remove(); lines.splice(lines.indexOf(item), 1); };
    wrap.append(sel, q, x); box.appendChild(wrap); lines.push(item);
  }
  (row.rejects && row.rejects.length ? row.rejects : [{}]).forEach(r => addLine(r.type, r.qty));
  $('op-m-add').onclick = () => addLine('', 0);
  const onKey = (e) => { if(e.key === 'Escape') close(); };
  function close(){ document.removeEventListener('keydown', onKey); ov.remove(); }
  document.addEventListener('keydown', onKey);
  $('op-m-cancel').onclick = close;
  ov.addEventListener('mousedown', (e) => { if(e.target === ov) close(); });
  $('op-m-save').onclick = async () => {
    const realRaw = $('op-m-real').value;
    if(needComment() && !$('op-m-com').value.trim()){
      showToast(opT('needComment'), true);
      refreshNeed();
      $('op-m-com').focus();
      return;
    }
    row.real = realRaw === '' ? null : Math.max(0, Math.floor(Number(realRaw) || 0));
    const rejects = [];
    lines.forEach(l => { const qn = Math.max(0, Math.floor(Number(l.q.value) || 0)); if(qn > 0) rejects.push({ type: l.sel.value, qty: qn }); });
    row.rejects = rejects;
    row.reject = rejects.reduce((s, x) => s + x.qty, 0);
    row.comment = $('op-m-com').value.trim();
    close();
    opRefreshCells();
    await opPersist(true);
  };
  if(ro){
    ov.querySelectorAll('input, select, textarea').forEach(e => { e.disabled = true; });
    ['op-m-save', 'op-m-add'].forEach(id => { const e = $(id); if(e) e.style.display = 'none'; });
    ov.querySelectorAll('.op-x').forEach(e => { e.style.display = 'none'; });
    $('op-m-cancel').textContent = opT('close');
  }else{
    setTimeout(() => { const r = $('op-m-real'); if(r) r.focus(); }, 50);
  }
}

function opShiftName(key){ return opT(key === 'morning' ? 'sMorning' : key === 'afternoon' ? 'sAfter' : 'sNight'); }
function opShiftDate(iso, delta){
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + delta)).toISOString().slice(0, 10);
}
/* ===== Historique : totaux de fin d'équipe, les 8 tables dans une seule liste ===== */
const OP_SHIFT_ORDER = { morning: 1, afternoon: 2, night: 3 };
async function opHistLoad(){
  const f = state.opHF;
  const { data, error } = await sb.from('op_tables').select('*').gte('sheet_date', f.from).lte('sheet_date', f.to).order('sheet_date', { ascending: false });
  if(error) throw error;
  state.opHRows = (data || []).map(r => {
    const rec = opNormalize(Object.assign({}, r));
    const a = opAgg([rec.hours_data]);
    return {
      rec, date: r.sheet_date, shift: r.shift, table: r.table_no, emp: r.employee_name || '', ref: r.ref || '',
      dem: a.dem, real: a.real, av: a.av, n: a.n, rej: a.rej,
      pr: a.dem > 0 ? a.real / a.dem * 100 : null, pj: a.real > 0 ? a.rej / a.real * 100 : null
    };
  });
}
function opHistFiltered(){
  const f = state.opHF, q = (s) => (s || '').toLowerCase().trim();
  return (state.opHRows || []).filter(r =>
    (!f.shift || r.shift === f.shift) && (!f.table || String(r.table) === String(f.table)) &&
    (!q(f.ref) || r.ref.toLowerCase().includes(q(f.ref))) && (!q(f.emp) || r.emp.toLowerCase().includes(q(f.emp)))
  ).sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : (OP_SHIFT_ORDER[b.shift] - OP_SHIFT_ORDER[a.shift]) || (a.table - b.table)));
}
function opHistExport(rows){
  if(!rows.length){ showToast(t('noExportData'), true); return; }
  if(typeof XLSX === 'undefined'){ showToast('XLSX non charg\u00e9 \u2014 v\u00e9rifiez votre connexion internet', true); return; }
  const out = rows.map(r => ({
    'Date': r.date, '\u00c9quipe': `${opShiftName(r.shift)} (${opShiftLabel(r.shift)})`, 'Table': `Table de contr\u00f4le ${opPad(r.table)}`,
    'Op\u00e9rateur': r.emp, 'Ref': r.ref, 'Demande': r.dem, 'R\u00e9el': r.real,
    '% R\u00e9el / Demande': r.pr === null ? '' : Math.round(r.pr * 10) / 10, 'AV / RE': r.n ? r.av : '',
    'Rejet': r.rej, 'Taux de rejet %': r.pj === null ? '' : Math.round(r.pj * 100) / 100
  }));
  const ws = XLSX.utils.json_to_sheet(out);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Historique');
  XLSX.writeFile(wb, `historique_equipes_${new Date().toISOString().slice(0, 10)}.xlsx`);
  showToast(t('exportedToast'));
}
function renderOpHistory(main){
  if(!state.opHF){
    const td = opCurrentShift().date;
    state.opHF = { from: opShiftDate(td, -13), to: td, shift: '', table: '', ref: '', emp: '' };
  }
  const f = state.opHF;
  main.innerHTML = `<style>${OP_CSS}${OP2_CSS}</style>
    <div class="page-head"><h2>${opT('history')} &mdash; totaux par \u00e9quipe</h2><div class="back-link" id="op-back">${ICONS.back} ${opT('portal')}</div></div>
    <div class="toolbar op-hf">
      <label>Du</label><input type="date" id="op-hf-from" value="${escHtml(f.from)}">
      <label>Au</label><input type="date" id="op-hf-to" value="${escHtml(f.to)}">
      <select id="op-hf-shift"><option value="">Toutes les \u00e9quipes</option>${OP_TEAMS.map(t2 => `<option value="${t2.key}" ${f.shift === t2.key ? 'selected' : ''}>${t2.label}</option>`).join('')}</select>
      <select id="op-hf-table"><option value="">Toutes les tables</option>${Array.from({ length: 8 }, (_, i) => `<option value="${i + 1}" ${String(f.table) === String(i + 1) ? 'selected' : ''}>Table de contr\u00f4le ${opPad(i + 1)}</option>`).join('')}</select>
      <input type="text" id="op-hf-ref" placeholder="Ref" value="${escAttr(f.ref)}" style="max-width:140px">
      <input type="text" id="op-hf-emp" placeholder="Op\u00e9rateur" value="${escAttr(f.emp)}" style="max-width:160px">
      <button class="btn blue" id="op-hf-export" style="width:auto;padding:9px 14px;">${ICONS.download} ${t('exportBtn')}</button>
    </div>
    <div class="table-wrap"><table class="op-ht">
      <thead><tr><th>Date</th><th>\u00c9quipe</th><th>Table</th><th>Op\u00e9rateur</th><th>Ref</th><th>Demande</th><th>R\u00e9el</th><th>% R\u00e9el / Demande</th><th>AV / RE</th><th>Rejet</th><th>Taux de rejet %</th></tr></thead>
      <tbody id="op-hbody"><tr><td colspan="11" class="empty-state">${opT('loading')}</td></tr></tbody>
      <tfoot id="op-hfoot"></tfoot>
    </table></div>
    <p class="op-hint">Touchez une ligne pour ouvrir la feuille de l'\u00e9quipe.</p>`;

  const fill = () => {
    const rows = opHistFiltered();
    const pct = (v) => v === null ? '\u2014' : (Math.round(v * 10) / 10) + ' %';
    const rateCls = (v) => v === null ? '' : (v <= OP_REJ_OK ? 'good' : v <= OP_REJ_WARN ? 'warn' : 'bad');
    const body = document.getElementById('op-hbody');
    if(!rows.length){
      body.innerHTML = `<tr><td colspan="11" class="empty-state">${opT('noData')}</td></tr>`;
      document.getElementById('op-hfoot').innerHTML = '';
      return rows;
    }
    body.innerHTML = rows.map((r, i) => `<tr class="op-hrow" data-i="${i}">
      <td>${escHtml(r.date)}</td><td>${opShiftName(r.shift)}</td><td>T${opPad(r.table)}</td><td>${escHtml(r.emp)}</td><td>${escHtml(r.ref)}</td>
      <td class="num">${r.dem}</td><td class="num">${r.real}</td><td class="num">${pct(r.pr)}</td>
      <td class="num ${r.n ? (r.av < 0 ? 'bad' : 'good') : ''}">${r.n ? (r.av > 0 ? '+' + r.av : r.av) : '\u2014'}</td>
      <td class="num">${r.rej}</td><td class="num ${rateCls(r.pj)}">${r.pj === null ? '\u2014' : (Math.round(r.pj * 100) / 100).toFixed(2) + '%'}</td></tr>`).join('');
    let dem = 0, real = 0, rej = 0, av = 0, n = 0;
    rows.forEach(r => { dem += r.dem; real += r.real; rej += r.rej; av += r.av; n += r.n; });
    const pr = dem > 0 ? real / dem * 100 : null, pj = real > 0 ? rej / real * 100 : null;
    document.getElementById('op-hfoot').innerHTML = `<tr><td colspan="5">Total (${rows.length} feuille${rows.length > 1 ? 's' : ''})</td>
      <td class="num">${dem}</td><td class="num">${real}</td><td class="num">${pct(pr)}</td>
      <td class="num ${n ? (av < 0 ? 'bad' : 'good') : ''}">${n ? (av > 0 ? '+' + av : av) : '\u2014'}</td><td class="num">${rej}</td>
      <td class="num ${rateCls(pj)}">${pj === null ? '\u2014' : (Math.round(pj * 100) / 100).toFixed(2) + '%'}</td></tr>`;
    body.querySelectorAll('.op-hrow').forEach(tr => {
      tr.onclick = () => { const r = rows[+tr.dataset.i]; opOpenTable(r.table, { date: r.date, shift: r.shift, from: 'history' }); };
    });
    return rows;
  };
  const reload = async () => {
    try{ await opHistLoad(); }
    catch(e){ console.error('op history error:', e); showToast(`${opT('err')}: ${opErrMsg(e)}`, true); state.opHRows = []; }
    fill();
  };
  document.getElementById('op-back').onclick = () => { state.view = 'op-portal'; render(); };
  document.getElementById('op-hf-from').onchange = (e) => { f.from = e.target.value || f.from; reload(); };
  document.getElementById('op-hf-to').onchange = (e) => { f.to = e.target.value || f.to; reload(); };
  document.getElementById('op-hf-shift').onchange = (e) => { f.shift = e.target.value; fill(); };
  document.getElementById('op-hf-table').onchange = (e) => { f.table = e.target.value; fill(); };
  document.getElementById('op-hf-ref').oninput = (e) => { f.ref = e.target.value; fill(); };
  document.getElementById('op-hf-emp').oninput = (e) => { f.emp = e.target.value; fill(); };
  document.getElementById('op-hf-export').onclick = () => opHistExport(opHistFiltered());
  reload();
}

/* ================= EXPORT EXCEL (Stock) ================= */
function exportToExcel(records){
  if(!records.length){ showToast(t('noExportData'), true); return; }
  const rows = records.map(r => ({ 'Lot': r.lot, 'Ref': r.ref, 'Désignation': r.designation, 'Défaut': r.defaut, 'Qtite': r.qtite, 'Locations': r.location, 'Fournisseur': r.fournisseur, 'User': r.user, 'Date': r.date, 'Prix': calcPrixTotal(r).toFixed(2) }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{wch:14},{wch:14},{wch:24},{wch:18},{wch:8},{wch:14},{wch:20},{wch:16},{wch:12},{wch:12}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Stock');
  XLSX.writeFile(wb, `stock_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast(t('exportedToast'));
}

/* ================= HELPERS ================= */
function escHtml(s){ return (s||'').toString().replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function escAttr(s){ return escHtml(s); }
function showToast(msg, isError){
  const el = document.getElementById('toast');
  el.textContent = msg; el.style.background = isError ? 'var(--danger)' : 'var(--ok)';
  el.classList.add('show'); setTimeout(()=>el.classList.remove('show'), 2400);
}

/* ================= INIT ================= */
(async function init(){
  loadLang();
  if(!configOk){ render(); return; }
  const { data } = await sb.auth.getSession();
  if(data.session && data.session.user){
    const { data: profile } = await sb.from('profiles').select('*').eq('id', data.session.user.id).single();
    state.currentUser = { id: data.session.user.id, username: profile ? profile.username : '', name: profile ? profile.name : '', role: profile ? profile.role : 'user' };
    await enterByRole();
  }
  render();
})();
</script>
</body>
</html>






