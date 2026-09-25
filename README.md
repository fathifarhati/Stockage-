<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gestion de Stock</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&family=Cairo:wght@500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<style>
  :root{
    --bg: #12161d; --panel: #1b212b; --panel-alt: #232a36; --border: #313a49;
    --text: #e7eaee; --text-muted: #8a93a3; --accent: #d68c45; --blue: #5b8fb0;
    --danger: #e2574c; --ok: #5aad8c; --radius: 3px;
  }
  *{ box-sizing:border-box; }
  html,body{ margin:0; padding:0; }
  body{
    background:
      repeating-linear-gradient(0deg, #ffffff05 0px, #ffffff05 1px, transparent 1px, transparent 40px),
      repeating-linear-gradient(90deg, #ffffff05 0px, #ffffff05 1px, transparent 1px, transparent 40px),
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
  .btn.blue{ background:var(--blue); color:#0e1a22; }
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
  .toast{ position:fixed; bottom:24px; left:50%; transform:translateX(-50%); background:var(--ok); color:#0e1a15; padding:10px 18px; border-radius:var(--radius); font-size:13px; font-weight:600; opacity:0; transition:opacity .2s; pointer-events:none; max-width:90vw; text-align:center; }
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
  .debit-header{ display:flex; align-items:center; justify-content:space-between; background:var(--panel); border:1px solid var(--border); padding:14px 18px; border-radius:var(--radius); margin-bottom:14px; }
  .debit-header .brand-mark{ font-size:20px; font-weight:700; font-family:'Space Grotesk',sans-serif; }
  .debit-header .doc-num{ background:var(--panel-alt); border:1px solid var(--border); padding:8px 14px; border-radius:var(--radius); font-family:'IBM Plex Mono',monospace; font-size:13px; }
  .debit-section{ border:1px solid var(--border); border-radius:var(--radius); margin-bottom:14px; overflow:hidden; }
  .debit-section-head{ background:var(--blue); color:#0e1a22; font-weight:700; padding:9px 14px; font-size:13px; letter-spacing:.3px; }
  .debit-table-wrap{ overflow-x:auto; }
  .debit-table{ width:100%; border-collapse:collapse; font-size:12.5px; min-width:560px; }
  .debit-table th{ background:var(--panel-alt); color:var(--text-muted); font-weight:500; font-size:10.5px; text-transform:uppercase; padding:7px 8px; text-align:start; border-bottom:1px solid var(--border); white-space:nowrap; }
  .debit-table td{ padding:5px 6px; border-bottom:1px solid var(--border); }
  .debit-table input{ margin:0; padding:6px 7px; font-size:12.5px; width:100%; min-width:70px; }
  .debit-table td.total-cell{ font-family:'IBM Plex Mono',monospace; text-align:end; padding-inline-end:10px; color:var(--accent); font-weight:600; white-space:nowrap; }
  .debit-section-foot{ display:flex; justify-content:space-between; align-items:center; padding:8px 14px; background:var(--panel-alt); }
  .debit-add-row{ background:transparent; border:none; color:var(--blue); font-size:12px; cursor:pointer; padding:6px 14px; }
  .debit-row-del{ background:transparent; border:none; color:var(--text-muted); cursor:pointer; padding:2px; display:flex; }
  .debit-row-del:hover{ color:var(--danger); }
  .debit-row-del svg{ width:13px; height:13px; }
  .debit-grand-total{ display:flex; justify-content:space-between; align-items:center; background:var(--panel); border:2px solid var(--danger); border-radius:var(--radius); padding:14px 18px; font-size:16px; font-weight:700; font-family:'Space Grotesk',sans-serif; margin-bottom:18px; }
  .debit-grand-total .amount{ color:var(--danger); font-family:'IBM Plex Mono',monospace; }
</style>
</head>
<body>
<div id="app"></div>
<div class="toast" id="toast"></div>

<script>
/* ============================================================
   CONFIGURATION SUPABASE
   ============================================================ */
const SUPABASE_URL = 'https://wajjhfgcybipnplmjoeq.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndhampoZmdjeWJpcG5wbG1qb2VxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODg1OTkyMzgsImV4cCI6MjEwNDE3NTIzOH0.paDNHxDzWqSdLM5CiOF25roD-l0neafwrqfbh8g27dY';

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
  debitExportedToast:{fr:'PDF généré et enregistré', ar:'تم إنشاء وحفظ PDF'},
  debitExportErr:{fr:'Échec de la génération du PDF', ar:'فشل إنشاء PDF'},
  debitRequiredErr:{fr:'Fournisseur et cause obligatoires', ar:'المورد والسبب إلزاميان'},
  debitNoNotes:{fr:'Aucune note de débit', ar:'لا توجد مذكرات خصم'},
  debitConfirmDelete:{fr:'Supprimer cette note ?', ar:'حذف هذه المذكرة؟'},
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
  recipientDeletedToast:{fr:'Destinata
