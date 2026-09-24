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
  cardBqTitle:{fr:'Panel qualité', ar:'bonne qualité'},
  cardBqDesc:{fr:'Certificats et documents d\'approbation qualité.', ar:'شهادات ووثائق موافقة الجودة.'},
  bqLibraryTitle:{fr:'Bonne qualité', ar:'Bonne qualité'},
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
  view: 'login', currentUser: null, records: [], profiles: [], catalog: [], qcRecords: [], recipients: [], pdfFiles: [], pdfFolder: '', currentLibrary: 'pdf',
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
const configOk = SUPABASE_URL.startsWith('https://') && SUPABASE_ANON_KEY.startsWith('eyJ');
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
async function createUserApi({ username, name, password, role }){
  const { data: sessionData } = await sb.auth.getSession();
  const token = sessionData.session ? sessionData.session.access_token : '';
  try{
    const res = await fetch(`${SUPABASE_URL}/functions/v1/Create_user-ts`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${token}` },
      body: JSON.stringify({ username, name, password, role }),
    });
    const json = await res.json();
    return json;
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

function render(){
  if(!configOk) return renderConfigError();
  if(state.view === 'login') return renderLogin();
  const otherLang = state.lang==='fr' ? 'ar' : 'fr';
  const otherLabel = state.lang==='fr' ? 'العربية' : 'Français';
  app.innerHTML = `
    <div class="topbar">
      <div class="brand"><span class="dot"></span>${t('appTitle')}</div>
      <div class="who">
        <button class="lang-btn" id="lang-btn">${otherLabel}</button>
        <span>${escHtml(state.currentUser.name)}</span>
        <button id="logout-btn">${t('logout')}</button>
      </div>
    </div>
    <main id="main"></main>`;
  document.getElementById('logout-btn').onclick = async () => { await sb.auth.signOut(); state.currentUser = null; state.view='login'; render(); };
  document.getElementById('lang-btn').onclick = () => setLang(otherLang);
  const main = document.getElementById('main');
  if(state.view === 'menu') renderMenu(main);
  if(state.view === 'stock-menu') renderStockMenu(main);
  if(state.view === 'qc-menu') renderQcMenu(main);
  if(state.view === 'entry') renderEntry(main);
  if(state.view === 'data') renderData(main);
  if(state.view === 'dashboard') renderDashboard(main);
  if(state.view === 'users') renderUsers(main);
  if(state.view === 'recipients') renderRecipients(main);
  if(state.view === 'pdf-library') renderPdfLibrary(main);
  if(state.view === 'catalog') renderCatalog(main);
  if(state.view === 'qc-entry') renderQcEntry(main);
  if(state.view === 'qc-data') renderQcData(main);
  if(state.view === 'qc-dashboard') renderQcDashboard(main);
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
    if(state.currentUser.role === 'employee'){
      await loadCatalog();
      state.view = 'qc-entry';
    }else{
      state.view = 'menu';
      await loadRecords();
    }
    render();
  };
  document.getElementById('login-btn').onclick = tryLogin;
  document.getElementById('login-pass').addEventListener('keydown', e => { if(e.key==='Enter') tryLogin(); });
}

function renderMenu(main){
  const isAdmin = state.currentUser.role === 'admin';
  main.innerHTML = `
    <div class="menu-head"><h1>${t('menuTitle')}</h1><p>${t('menuSubtitle')}</p></div>
    <div class="menu-grid">
      <div class="menu-card" id="card-stock">${ICONS.data}<h3>${t('moduleStockTitle')}</h3><p>${t('moduleStockDesc')}</p></div>
      <div class="menu-card" id="card-qc">${ICONS.qc}<h3>${t('moduleQcTitle')}</h3><p>${t('moduleQcDesc')}</p></div>
      <div class="menu-card" id="card-pdf">${ICONS.pdf}<h3>${t('cardPdfTitle')}</h3><p>${t('cardPdfDesc')}</p></div>
      <div class="menu-card" id="card-bq">${ICONS.pdf}<h3>${t('cardBqTitle')}</h3><p>${t('cardBqDesc')}</p></div>
      ${isAdmin ? `<div class="menu-card" id="card-users">${ICONS.users}<h3>${t('cardUsersTitle')}</h3><p>${t('cardUsersDesc')}</p></div>` : ''}
      ${isAdmin ? `<div class="menu-card" id="card-recipients">${ICONS.mail}<h3>${t('cardRecipientsTitle')}</h3><p>${t('cardRecipientsDesc')}</p></div>` : ''}
    </div>`;
  document.getElementById('card-stock').onclick = () => { state.view='stock-menu'; render(); };
  document.getElementById('card-qc').onclick = async () => { await loadCatalog(); await loadQcRecords(); state.view='qc-menu'; render(); };
  document.getElementById('card-pdf').onclick = async () => { state.currentLibrary='pdf'; state.pdfFolder=''; await listPdfEntries(''); state.view='pdf-library'; render(); };
  document.getElementById('card-bq').onclick = async () => { state.currentLibrary='bonnequalite'; state.pdfFolder=''; await listPdfEntries(''); state.view='pdf-library'; render(); };
  if(isAdmin) document.getElementById('card-users').onclick = () => { state.view='users'; render(); };
  if(isAdmin) document.getElementById('card-recipients').onclick = async () => { await loadRecipients(); state.view='recipients'; render(); };
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

let chartRefs = [];
function renderDashboard(main){
  chartRefs.forEach(c=>c.destroy()); chartRefs = [];
  const recs = state.records;
  const totalQte = recs.reduce((s,r)=>s+r.qtite,0);
  const fournisseurs = [...new Set(recs.map(r=>r.fournisseur).filter(Boolean))];
  const locations = [...new Set(recs.map(r=>r.location).filter(Boolean))];
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.dash.replace('<svg','<svg width="18" height="18"')} ${t('cardDashTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
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
    <div class="page-head"><h2>${ICONS.data.replace('<svg','<svg width="18" height="18"')} ${t('cardQcDataTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
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
  document.getElementById('back').onclick = () => { state.view='qc-menu'; render(); };
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
      <button class="icon-btn" id="qedit-${r.id}" title="${t('editedToast')}">${ICONS.edit}</button>
      <button class="icon-btn del" id="qdel-${r.id}" title="${t('deletedToast')}">${ICONS.del}</button>
    </td>
  </tr>`;
}

let qcChartRefs = [];
function renderQcDashboard(main){
  qcChartRefs.forEach(c => c.destroy()); qcChartRefs = [];
  const recs = state.qcRecords;
  const totalControle = recs.reduce((s,r)=>s+r.qteControle,0);
  const totalNonOk = recs.reduce((s,r)=>s+r.qteNonOk,0);
  const globalRate = totalControle > 0 ? (totalNonOk/totalControle*100) : 0;
  main.innerHTML = `
    <div class="page-head"><h2>${ICONS.dash.replace('<svg','<svg width="18" height="18"')} ${t('cardQcDashTitle')}</h2><div class="back-link" id="back">${ICONS.back} ${t('back')}</div></div>
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
    document.getElementById('back').onclick = () => { state.view='menu'; render(); };
    document.getElementById('upload-pdf-btn').onclick = async () => {
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
  document.getElementById('back').onclick = async () => { state.pdfFolder=''; await listPdfEntries(''); renderPdfLibrary(main); };
  document.getElementById('upload-pdf-btn').onclick = async () => {
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
    if(state.currentUser.role === 'employee'){
      await loadCatalog();
      state.view = 'qc-entry';
    }else{
      state.view = 'menu';
      await loadRecords();
    }
  }
  render();
})();
</script>
</body>
</html>




















































































        
