/**
 * مشروعك – Espace client des études NESDA
 * Back-end gratuit : Google Apps Script + Google Sheets + Google Drive.
 * À coller dans un projet Apps Script (voir LISEZ-MOI.txt).
 *
 * Le site réalise UNIQUEMENT les études. Il ne fournit ni facture proforma ni assurance :
 * le porteur de projet les apporte lui-même (fournisseur, assureur).
 */
const ADMIN_PASSWORD = '';   // À REMPLIR uniquement dans script.google.com, jamais sur GitHub                 // mot de passe de la page admin
const DRIVE_FOLDER   = 'مشروعك - وثائق الزبائن';          // dossier Drive des documents
const MAX_MB = 5;                                                 // taille max d'un fichier client (Mo)
const QUOTA_MB = 100;                                             // espace total par client (Mo)
const ENVOIS_HEURE = 15;                                          // envois / suppressions max par heure et par client
const PRIX = 5000;                                        // montant de l'étude (DA)
const TEL_APPEL = '0556417326';                           // numéro à appeler (chiffres seulement)
const HORAIRES = {jours: [6, 0, 1, 2, 3, 4], debut: '08:30', fin: '17:00'};   // heures d'appel : samedi(6) … jeudi(4), heure d'Algérie
const SITE_URL  = 'https://machrouak.pages.dev';         // adresse publique de l'espace client (après mise en ligne)
const RELANCE_JOURS = 3;                                  // relance email tous les X jours tant que l'étude n'est pas payée
const RELANCE_MAX   = 3;
const SESSION_H = 6;                                      // durée d'une connexion client (heures)

// Seuls documents reçus : la facture proforma et le devis d'assurance, que le client se procure lui-même
const DOCS = [
  ['proforma', 'final', 'Facture proforma des équipements (fournie par le fournisseur du client)'],
  ['assurance', 'final', 'Devis d\'assurance (fourni par l\'assureur du client)'],
  ['autre', 'final', 'Autre document (fiche technique, photo…)'],
];
const MAX_FICHIERS = 20;   // fichiers max par type de document
// Fiche de coordonnées remplie par le client lui-même : [clé, libellé admin, obligatoire]
const FICHE = [
  ['nom_ar', 'Nom et prénom (en arabe)', 1], ['nom_lat', 'Nom et prénom (lettres latines)', 1], ['nin', 'NIN (numéro d\'identification nationale, 18 chiffres)', 1], ['sexe', 'Sexe', 1], ['lieu_naissance', 'Lieu de naissance', 1],
  ['adresse', 'Adresse personnelle', 1], ['commune', 'Commune', 1], ['wilaya_res', 'Wilaya', 1], ['situation', 'Situation familiale', 0],
  ['diplome', 'Niveau d\'études / diplôme', 1], ['experience', 'Expérience professionnelle', 1], ['activite', 'Activité du projet', 1],
  ['local_adresse', 'Adresse du local', 1], ['local_statut', 'Le local', 1], ['remarques', 'Remarques', 0],
];
const STEPS = ['Réception du dossier', 'Étude préliminaire d\'inscription', 'Préparation de l\'étude finale', 'Révision avec le porteur de projet', 'Étude prête'];
const STAT_DOC = ['مطلوب', 'تم الاستلام', 'مقبول', 'يجب إعادة إرسالها'];   // statuts stockés (traduits côté pages)

const H = {
  Dossiers:  ['code', 'nom', 'tel', 'ddn', 'email', 'lang', 'projet', 'wilaya', 'langue', 'etape', 'note', 'cree', 'maj',
              'mdp', 'sel', 'paiement', 'recu', 'recu_url', 'coord_envoi', 'etude', 'etude_url', 'resume', 'envoi', 'envoi_date', 'relances', 'relance_date', 'bloque', 'quiz', 'prep', 'accepte'],
  Documents: ['code', 'doc', 'statut', 'fichier', 'url', 'remarque', 'date', 'id', 'taille'],
  Journal:   ['date', 'code', 'action', 'detail'],
  Fiches:    ['code'].concat(FICHE.map(f => f[0]), ['maj']),
  Demandes:  ['id', 'code', 'date', 'type', 'message', 'statut', 'reponse', 'date_reponse'],
  Prospects: ['date', 'nom', 'tel', 'wilaya', 'activite', 'type', 'statut', 'zone', 'formule', 'age', 'cout', 'apport', 'nesda', 'banque', 'lang', 'source', 'suivi', 'loyer', 'poche'],
};
const TEXTE = ['tel', 'ddn', 'nin'];   // colonnes gardées en texte (le 0 du téléphone, la date de naissance)

/* ---------- feuille ---------- */
function book() {
  const P = PropertiesService.getScriptProperties(); const id = P.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create('مشروعك - Dossiers'); P.setProperty('SHEET_ID', ss.getId()); return ss;
}
function sh(name) {
  const ss = book(); let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(H[name]); s.setFrozenRows(1); s.getRange(1, 1, 1, H[name].length).setFontWeight('bold'); }
  else if (s.getLastColumn() < H[name].length) s.getRange(1, 1, 1, H[name].length).setValues([H[name]]).setFontWeight('bold');   // nouvelles colonnes ajoutées à la fin
  return s;
}
function rows(name) {
  const v = sh(name).getDataRange().getValues(); const h = v.shift();
  return v.map((r, i) => { const o = {_row: i + 2}; h.forEach((k, j) => o[k] = r[j] instanceof Date ? (k === 'ddn' ? Utilities.formatDate(r[j], 'Africa/Algiers', 'yyyy-MM-dd') : fmt(r[j])) : r[j]); return o; });
}
function put(name, o) {
  const s = sh(name), r = H[name].map(k => { let x = o[k] === undefined ? '' : o[k]; if (TEXTE.includes(k) && x !== '') x = "'" + String(x).replace(/^'/, ''); return x; });
  if (o._row) s.getRange(o._row, 1, 1, r.length).setValues([r]); else s.appendRow(r);
}
const fmt = d => Utilities.formatDate(d, 'Africa/Algiers', 'dd/MM/yyyy HH:mm');
const now = () => fmt(new Date());
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const norm = s => String(s || '').trim().toUpperCase();
const telN = t => { let n = String(t || '').replace(/\D/g, ''); if (n.startsWith('00213')) n = '0' + n.slice(5); else if (n.startsWith('213')) n = '0' + n.slice(3); return n; };
const dossier = code => { const d = rows('Dossiers').find(x => norm(x.code) === norm(code)); if (!d) throw new Error('Dossier introuvable'); return d; };

/* ---------- connexion client : téléphone + date de naissance -> mot de passe choisi par le client ---------- */
const E = {   // messages d'erreur dans la langue du client
  ar: {id: 'رقم الهاتف أو كلمة السر غير صحيحة', ddn: 'رقم الهاتف أو تاريخ الميلاد غير صحيح', court: 'كلمة السر قصيرة (6 أحرف على الأقل)', bloque: 'محاولات كثيرة، أعد المحاولة بعد 30 دقيقة', suspendu: 'فضاؤك موقوف، اتصل بنا', limite: 'عدد كبير من الإرسالات، أعد المحاولة بعد ساعة', quota: 'امتلأت المساحة المخصصة لك (100 ميغا)، احذف بعض الملفات أو اتصل بنا', session: 'انتهت الجلسة، ادخل من جديد', premier: 'أنشئ حسابك أولاً (حساب جديد)', existe: 'هذا الرقم لديه حساب، ادخل بكلمة السر', telInv: 'رقم الهاتف غير صحيح (10 أرقام، مثل 0550123456)', nom: 'اكتب اسمك ولقبك', plein: 'تسجيلات كثيرة الآن، أعد المحاولة بعد قليل'},
  fr: {id: 'Téléphone ou mot de passe incorrect', ddn: 'Téléphone ou date de naissance incorrect', court: 'Mot de passe trop court (6 caractères minimum)', bloque: 'Trop d\'essais, réessayez dans 30 minutes', suspendu: 'Votre espace est suspendu, contactez-nous', limite: 'Trop d\'envois, réessayez dans une heure', quota: 'Votre espace est plein (100 Mo) : supprimez des fichiers ou contactez-nous', session: 'Session expirée, reconnectez-vous', premier: 'Créez d\'abord votre compte (onglet Nouveau compte)', existe: 'Ce numéro a déjà un compte : connectez-vous avec votre mot de passe', telInv: 'Numéro de téléphone invalide (10 chiffres, ex. 0550123456)', nom: 'Indiquez votre nom et prénom', plein: 'Trop d\'inscriptions en ce moment, réessayez dans un instant'},
  en: {id: 'Wrong phone number or password', ddn: 'Wrong phone number or date of birth', court: 'Password too short (at least 6 characters)', bloque: 'Too many attempts, try again in 30 minutes', suspendu: 'Your space is suspended, please contact us', limite: 'Too many uploads, try again in an hour', quota: 'Your space is full (100 MB): delete some files or contact us', session: 'Session expired, please sign in again', premier: 'Create your account first (New account tab)', existe: 'This number already has an account: sign in with your password', telInv: 'Invalid phone number (10 digits, e.g. 0550123456)', nom: 'Enter your full name', plein: 'Too many sign-ups right now, try again shortly'},
};
const err = (q, k) => new Error((E[q.lang] || E.ar)[k]);
function hash(sel, mdp) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sel + '|' + mdp, Utilities.Charset.UTF_8)); }
function garde(q) {   // 5 essais ratés -> blocage 30 min (par téléphone)
  const c = CacheService.getScriptCache(), k = 'e_' + telN(q.tel), n = Number(c.get(k) || 0);
  if (n >= 5) throw err(q, 'bloque');
  return {rate: () => c.put(k, String(n + 1), 1800), ok: () => c.remove(k)};
}
function ouvrir(d) { const t = Utilities.getUuid(); CacheService.getScriptCache().put('t_' + t, d.code, SESSION_H * 3600); return t; }
function parToken(q) { const c = q.token && CacheService.getScriptCache().get('t_' + q.token); if (!c) throw err(q, 'session'); const d = dossier(c); if (d.bloque) throw err(q, 'suspendu'); return d; }
function journal(code, action, detail) { try { sh('Journal').appendRow([now(), code || '', action, String(detail || '').slice(0, 300)]); } catch (e) {} }
function cadence(q, d) {   // limite les envois/suppressions par heure et par client
  const c = CacheService.getScriptCache(), k = 'u_' + norm(d.code), n = Number(c.get(k) || 0);
  if (n >= ENVOIS_HEURE) { journal(d.code, 'LIMITE', 'trop d\'envois en une heure'); throw err(q, 'limite'); }
  c.put(k, String(n + 1), 3600);
}
function quota(q, d, octets) {
  const tot = rows('Documents').filter(y => norm(y.code) === norm(d.code)).reduce((s, y) => s + (Number(y.taille) || 0), 0);
  if (tot + octets > QUOTA_MB * 1024 * 1024) { journal(d.code, 'QUOTA', Math.round(tot / 1048576) + ' Mo'); throw err(q, 'quota'); }
}
function gardeAdmin() {   // 5 mauvais mots de passe admin -> blocage 30 min
  const c = CacheService.getScriptCache(), n = Number(c.get('adm') || 0);
  if (n >= 5) throw new Error('Trop d\'essais : page admin bloquée 30 minutes');
  return {rate: () => { c.put('adm', String(n + 1), 1800); journal('', 'ADMIN_ECHEC', 'mauvais mot de passe admin (' + (n + 1) + '/5)'); }, ok: () => { if (n) c.remove('adm'); }};
}
const parTel = t => rows('Dossiers').find(x => telN(x.tel) === telN(t) && telN(t));

/* ---------- format des nombres : en arabe un point tous les 3 chiffres (4.200.000) ---------- */
const AR = /[؀-ۿ]/;
function nombre(n, lang) { const s = String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, lang === 'en' ? ',' : lang === 'fr' ? ' ' : '.'); return s; }
function chiffres(txt, lang) {   // « 4 200 000 » / « 4200000 » écrits par l'admin -> séparateur adapté
  const l = AR.test(txt || '') ? 'ar' : lang;
  return String(txt || '').replace(/\d{1,3}(?:[   .,]\d{3})+(?!\d)|\d{5,}(?![\d.,]\d)/g, m => /^0\d{8,9}$/.test(m) ? m : nombre(m.replace(/\D/g, ''), l));
}
const telAff = lang => { const t = TEL_APPEL; const g = [t.slice(0, 4), t.slice(4, 6), t.slice(6, 8), t.slice(8)]; return lang === 'ar' ? g.join('.') : g.join(' '); };
const JOURS = {ar: ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'], fr: ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'], en: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']};
function horairesTxt(lang) {
  const j = JOURS[lang] || JOURS.ar, a = HORAIRES.jours[0], b = HORAIRES.jours[HORAIRES.jours.length - 1];
  return lang === 'fr' ? `du ${j[a]} au ${j[b]}, de ${HORAIRES.debut.replace(':', 'h')} à ${HORAIRES.fin.replace(':', 'h')}`
    : lang === 'en' ? `${j[a]} to ${j[b]}, ${HORAIRES.debut} – ${HORAIRES.fin}` : `من ${j[a]} إلى ${j[b]}، من ${HORAIRES.debut} إلى ${HORAIRES.fin}`;
}

/* ---------- vue envoyée aux pages ---------- */
const ficheDe = code => rows('Fiches').find(x => norm(x.code) === norm(code)) || {code};
const ficheOk = f => FICHE.every(([k, , o]) => !o || String(f[k] || '').trim());
function view(d, admin) {
  const docs = rows('Documents').filter(x => norm(x.code) === norm(d.code)), fi = ficheDe(d.code);
  return {
    dossier: {accepte: d.accepte || '', bloque: admin ? (d.bloque || '') : undefined, code: d.code, nom: d.nom, tel: d.tel, email: d.email || '', lang: d.lang || 'ar', projet: d.projet, wilaya: d.wilaya, langue: d.langue, etape: Number(d.etape) || 1, note: d.note, maj: d.maj,
      ddn: admin ? d.ddn : undefined, acces: admin ? (d.mdp ? 'oui' : 'non') : undefined},
    nsteps: STEPS.length, steps: admin ? STEPS : undefined,
    contact: {tel: TEL_APPEL, horaires: HORAIRES},
    docs: DOCS.map(([k, g, fr]) => {
      const fs = docs.filter(y => y.doc === k && y.fichier).map(y => ({id: docId(y), fichier: y.fichier, statut: y.statut || 'تم الاستلام', url: admin ? (y.url || '') : '', remarque: y.remarque || '', date: y.date || ''}));
      const st = !fs.length ? (k === 'autre' ? '' : 'مطلوب') : fs.some(f => f.statut === 'يجب إعادة إرسالها') ? 'يجب إعادة إرسالها' : fs.some(f => f.statut !== 'مقبول') ? 'تم الاستلام' : 'مقبول';
      return {key: k, groupe: g, nom: admin ? fr : undefined, optionnel: k === 'autre', statut: st, fichiers: fs};
    }),
    fiche: {valeurs: Object.fromEntries(FICHE.map(([k]) => [k, fi[k] || ''])), complete: ficheOk(fi), maj: fi.maj || '', champs: admin ? FICHE : undefined},
    paiement: {montant: PRIX, statut: d.paiement || 'غير مدفوع', recu: d.recu || '', recu_url: admin ? (d.recu_url || '') : '', coord_envoi: admin ? (d.coord_envoi || '') : undefined},
    etude: {prete: !!d.etude_url, nom: d.etude || '', url: (admin || d.paiement === 'مؤكد') ? (d.etude_url || '') : '', resume: d.etude_url ? (d.resume || '') : '',
      envoi: admin ? etatEnvoi(d) : undefined, message: admin && d.etude_url ? messageEtude(d, false) : undefined},
    quiz: quizDe(d),
    prep: (() => { const p = prepDe(d), T = p.texte || {}, ok = d.paiement === 'مؤكد'; if (admin || ok) return p;   // avant paiement : seulement facture + projet
      return {materiel: p.materiel || [], texte: {activite: T.activite, lieu: T.lieu, fournisseur: T.fournisseur}, chiffres: {}, produits: [], postes: []}; })(),
    invitation: admin ? invitation(d) : undefined,
    journal: admin ? rows('Journal').filter(x => norm(x.code) === norm(d.code)).slice(-25).reverse().map(({_row, ...x}) => x) : undefined,
    demandes: rows('Demandes').filter(x => norm(x.code) === norm(d.code)).map(({_row, ...x}) => x).reverse(),
  };
}

/* ---------- point d'entrée ---------- */
function doGet() { return out({ok: true, service: 'مشروعك'}); }
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let q = {};
  try {
    q = JSON.parse(e.postData.contents || '{}'); const a = q.action;
    /* ----- client : connexion ----- */
    if (a === 'login') {
      const g = garde(q), d = parTel(q.tel);
      if (d && d.bloque) throw err(q, 'suspendu');
      if (d && !d.mdp) throw err(q, 'premier');
      if (!d || hash(d.sel, String(q.mdp || '')) !== d.mdp) { g.rate(); throw err(q, 'id'); }
      g.ok(); journal(d.code, 'CONNEXION', ''); if (q.lang && q.lang !== d.lang) { d.lang = q.lang; put('Dossiers', d); }
      return out({ok: true, token: ouvrir(d), data: view(d)});
    }
    if (a === 'simConfig') return out({ok: true, cfg: simCfg()});   // frais de démarrage du simulateur (public)
    if (a === 'lead') {   // simulateur public : on garde le contact du visiteur (Prospects)
      const tel = telN(q.tel), nom = String(q.nom || '').trim().slice(0, 120);
      if (!/^0[5-7]\d{8}$/.test(tel)) throw err(q, 'telInv');
      if (nom.length < 3) throw err(q, 'nom');
      const c = CacheService.getScriptCache(), n = Number(c.get('leads') || 0), kt = 'l_' + tel;
      if (n >= 200 || c.get(kt)) return out({ok: true});   // anti-robot : 200 / heure, 1 envoi par numéro et par 10 min
      c.put('leads', String(n + 1), 3600); c.put(kt, '1', 600);
      const num = v => Math.max(0, Math.round(Number(v) || 0)), txt = (v, m) => String(v || '').trim().slice(0, m);
      put('Prospects', {date: now(), nom, tel, wilaya: txt(q.wilaya, 40), activite: txt(q.activite, 80), type: txt(q.type, 20), statut: txt(q.statut, 20), zone: txt(q.zone, 20),
        formule: txt(q.formule, 10), age: num(q.age) || '', cout: num(q.cout), apport: num(q.apport), nesda: num(q.nesda), banque: num(q.banque), lang: txt(q.lang, 2), source: txt(q.source, 40), suivi: '', loyer: num(q.loyer), poche: num(q.poche)});
      return out({ok: true});
    }
    if (a === 'inscription') {   // nouveau client : nom + téléphone + mot de passe, le dossier est créé tout de suite
      const g = garde(q), tel = telN(q.tel), nom = String(q.nom || '').trim().slice(0, 120), mdp = String(q.mdp || '');
      if (!/^0[5-7]\d{8}$/.test(tel)) throw err(q, 'telInv');
      if (nom.length < 3) throw err(q, 'nom');
      if (mdp.length < 6) throw err(q, 'court');
      const c = CacheService.getScriptCache(), n = Number(c.get('inscr') || 0); if (n >= 40) throw err(q, 'plein');   // anti-robot : 40 / heure
      let d = parTel(tel);
      if (d && d.mdp) { g.rate(); throw err(q, 'existe'); }
      if (d && d.bloque) throw err(q, 'suspendu');
      if (!d) {
        const all = rows('Dossiers'), an = new Date().getFullYear().toString().slice(2);
        let k = all.length + 1, code; do { code = 'M' + an + '-' + String(k++).padStart(3, '0'); } while (all.some(x => norm(x.code) === code));
        d = {code, nom, tel, ddn: '', email: '', lang: q.lang || 'ar', projet: String(q.projet || '').trim().slice(0, 200), wilaya: '', langue: '', etape: 1, note: '', cree: now(), maj: now()};
        c.put('inscr', String(n + 1), 3600); journal(code, 'INSCRIPTION', 'compte créé par le client');
      } else if (!d.nom) d.nom = nom;   // dossier déjà ouvert par l'admin, ou accès réinitialisé
      g.ok(); d.sel = Utilities.getUuid(); d.mdp = hash(d.sel, mdp); if (q.lang) d.lang = q.lang; d.maj = now(); put('Dossiers', d);
      journal(d.code, 'MOT_DE_PASSE', 'créé'); return out({ok: true, token: ouvrir(d), data: view(d)});
    }
    if (a === 'creerMdp') {   // première connexion ou mot de passe oublié : téléphone + date de naissance
      const g = garde(q), d = parTel(q.tel);
      if (!d || !d.ddn || String(d.ddn).replace(/^'/, '') !== String(q.ddn || '')) { g.rate(); throw err(q, 'ddn'); }
      if (d.bloque) throw err(q, 'suspendu');
      if (String(q.mdp || '').length < 6) throw err(q, 'court');
      g.ok(); d.sel = Utilities.getUuid(); d.mdp = hash(d.sel, String(q.mdp)); if (q.lang) d.lang = q.lang; d.maj = now(); put('Dossiers', d); journal(d.code, 'MOT_DE_PASSE', 'créé / changé');
      return out({ok: true, token: ouvrir(d), data: view(d)});
    }
    /* ----- client connecté ----- */
    if (['get', 'upload', 'supprimerDoc', 'recu', 'request', 'email', 'lang', 'fiche', 'quiz'].includes(a)) {
      const d = parToken(q);
      if (a === 'fiche') { sauverFiche(d, q.fiche); return out({ok: true, data: view(d)}); }
      if (a === 'quiz') {   // préparation à la commission : on garde la meilleure note par niveau
        const n = String(q.niveau); if (!['1', '2', '3'].includes(n)) throw new Error('?');
        const s = Math.max(0, Math.min(100, Math.round(Number(q.score) || 0))), o = quizDe(d);
        if (s > (o[n] || 0)) { o[n] = s; d.quiz = JSON.stringify(o); put('Dossiers', d); }
        journal(d.code, 'QUIZ', 'niveau ' + n + ' : ' + s + '%'); return out({ok: true, data: view(d)});
      }
      if (a === 'get') return out({ok: true, data: view(d)});
      if (a === 'lang') { if (['ar', 'fr', 'en'].includes(q.lang)) { d.lang = q.lang; put('Dossiers', d); } return out({ok: true, data: view(d)}); }
      if (a === 'upload') {
        if (!DOCS.some(([k]) => k === q.doc)) throw new Error('?');
        if (rows('Documents').filter(y => norm(y.code) === norm(d.code) && y.doc === q.doc && y.fichier).length >= MAX_FICHIERS) throw new Error('Max ' + MAX_FICHIERS);
        cadence(q, d); const taille = Math.floor(String(q.b64 || '').length * 3 / 4); quota(q, d, taille);
        const f = saveFile(d, q.doc + '_' + q.name, q);
        put('Documents', {code: d.code, doc: q.doc, statut: 'تم الاستلام', fichier: q.name, url: f.getUrl(), remarque: '', date: now(), id: Utilities.getUuid().slice(0, 8), taille});
        journal(d.code, 'ENVOI', q.doc + ' : ' + q.name + ' (' + Math.round(taille / 1024) + ' Ko)');
        d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d)});
      }
      if (a === 'supprimerDoc') {   // le client retire un fichier envoyé par erreur (sauf s'il est déjà accepté)
        const x = docDe(d.code, q.id); if (x.statut === 'مقبول') throw new Error('✔');
        cadence(q, d); supprimerLigne(x); journal(d.code, 'SUPPRESSION', x.doc + ' : ' + x.fichier); d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d)});
      }
      if (a === 'recu') {
        if (d.paiement === 'مؤكد') return out({ok: true, data: view(d)});
        cadence(q, d); const f = saveFile(d, 'RECU_PAIEMENT_' + q.name, q); journal(d.code, 'RECU', q.name);
        d.recu = q.name; d.recu_url = f.getUrl(); d.paiement = 'قيد التحقق'; d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d)});
      }
      if (a === 'request') {
        const msg = String(q.message || '').trim().slice(0, 2000); if (!msg) throw new Error('…');
        put('Demandes', {id: Utilities.getUuid().slice(0, 8), code: d.code, date: now(), type: q.type || 'modif', message: msg, statut: 'قيد المعالجة', reponse: '', date_reponse: ''});
        return out({ok: true, data: view(d)});
      }
      if (a === 'email') {
        const m = String(q.email || '').trim(); if (m && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(m)) throw new Error('@ ?');
        d.email = m; d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d)});
      }
    }
    /* ----- admin ----- */
    const ga = gardeAdmin();
    if (q.password !== ADMIN_PASSWORD) { ga.rate(); throw new Error('Mot de passe admin incorrect'); }
    ga.ok();
    if (a === 'adminList') {
      const dem = rows('Demandes');
      return out({ok: true, steps: STEPS, data: rows('Dossiers').map(d => ({code: d.code, nom: d.nom, tel: d.tel, projet: d.projet, etape: Number(d.etape) || 1, maj: d.maj, paiement: d.paiement || 'غير مدفوع',
        acces: d.mdp ? 'oui' : 'non', bloque: !!d.bloque, fiche: ficheOk(ficheDe(d.code)), etude: etatEnvoi(d).code, demandes: dem.filter(x => norm(x.code) === norm(d.code) && x.statut !== 'تمت المعالجة').length}))});
    }
    if (a === 'adminGet') return out({ok: true, data: view(dossier(q.code), true)});
    if (a === 'adminSimConfig') {   // l'admin règle les frais du simulateur
      const c = q.cfg || {}, n = (v, max) => Math.max(0, Math.min(max, Number(v) || 0));
      const cfg = {notaireTaux: n(c.notaireTaux, 20), notaireMin: Math.round(n(c.notaireMin, 1e7)), loyerFinance: !!c.loyerFinance, loyerPlafond: Math.round(n(c.loyerPlafond, 1e8)),
        frais: (c.frais || []).slice(0, 15).map(f => ({ar: String(f.ar || '').slice(0, 60), fr: String(f.fr || '').slice(0, 60), m: Math.round(n(f.m, 1e7)), etape: ['loc', 'fisc', 'compte'].includes(f.etape) ? f.etape : 'fisc'})).filter(f => f.ar || f.fr)};
      PropertiesService.getScriptProperties().setProperty('SIM_CFG', JSON.stringify(cfg)); return out({ok: true, cfg});
    }
    if (a === 'adminGetSimConfig') return out({ok: true, cfg: simCfg()});
    if (a === 'adminProspects') return out({ok: true, data: rows('Prospects').slice(-300).reverse().map(({_row, ...x}) => Object.assign(x, {row: _row}))});
    if (a === 'adminProspectSuivi') { const r = rows('Prospects').find(x => x._row === Number(q.row)); if (!r) throw new Error('Prospect introuvable'); r.suivi = String(q.suivi || '').slice(0, 200); put('Prospects', r); return out({ok: true}); }
    if (a === 'adminCreate') {
      if (!telN(q.tel)) throw new Error('Téléphone obligatoire (il sert à la connexion du client)');
      const all = rows('Dossiers'); if (parTel(q.tel)) throw new Error('Ce numéro de téléphone a déjà un dossier');
      const code = 'M' + new Date().getFullYear().toString().slice(2) + '-' + String(all.length + 1).padStart(3, '0');
      const d = {code, nom: q.nom || '', tel: q.tel, ddn: q.ddn, email: q.email || '', lang: q.lang || 'ar', projet: q.projet || '', wilaya: q.wilaya || '', langue: q.langue || 'العربية', etape: 1, note: '', cree: now(), maj: now()};
      put('Dossiers', d); return out({ok: true, code, invitation: invitation(d)});
    }
    if (a === 'adminUpdate') {
      const d = dossier(q.code);
      if (q.tel !== undefined && telN(q.tel) !== telN(d.tel) && parTel(q.tel)) throw new Error('Ce numéro de téléphone a déjà un dossier');
      ['nom', 'tel', 'ddn', 'email', 'lang', 'projet', 'wilaya', 'langue', 'etape', 'note', 'resume'].forEach(k => { if (q[k] !== undefined) d[k] = q[k]; });
      d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)});
    }
    if (a === 'adminFiche') { const d = dossier(q.code); sauverFiche(d, q.fiche); return out({ok: true, data: view(d, true)}); }
    if (a === 'adminAccepte') { const d = dossier(q.code); d.accepte = q.accepte ? now() : ''; d.maj = now(); put('Dossiers', d); journal(d.code, q.accepte ? 'ACCEPTE' : 'ACCEPTE_ANNULE', 'dossier accepté par l\'accompagnateur'); return out({ok: true, data: view(d, true)}); }
    if (a === 'adminBloquer') { const d = dossier(q.code); d.bloque = q.bloque ? now() : ''; d.maj = now(); put('Dossiers', d); journal(d.code, q.bloque ? 'BLOQUE' : 'DEBLOQUE', 'par l\'admin'); return out({ok: true, data: view(d, true)}); }
    if (a === 'adminPrep') {   // chiffres de l'étude + matériel de la facture, pour le quiz du client
      const d = dossier(q.code), p = q.prep || {}, C = {};
      ['cout', 'apport', 'pret_nesda', 'pret_banque', 'ca', 'benefice', 'emplois', 'seuil', 'loyer', 'prime'].forEach(k => { const v = Number(String((p.chiffres || {})[k] || '').replace(/[^\d]/g, '')); if (v) C[k] = v; });
      if ((p.chiffres || {}).retour) C.retour = String(p.chiffres.retour).slice(0, 80);
      const M = (p.materiel || []).slice(0, 60).map(m => ({nom: String(m.nom || '').slice(0, 120), prix: Number(String(m.prix || '').replace(/[^\d]/g, '')) || '', role: String(m.role || '').slice(0, 300)})).filter(m => m.nom);
      const T = {}; ['activite', 'lieu', 'clients', 'fournisseur', 'assureur', 'local'].forEach(k => { const v = String((p.texte || {})[k] || '').trim(); if (v) T[k] = v.slice(0, 300); });
      const PR = (p.produits || []).slice(0, 40).map(x => ({nom: String(x.nom || '').slice(0, 120), prix: Number(String(x.prix || '').replace(/[^\d]/g, '')) || '', unite: String(x.unite || '').slice(0, 40)})).filter(x => x.nom);
      const PO = (p.postes || []).slice(0, 30).map(x => ({poste: String(x.poste || '').slice(0, 120), salaire: Number(String(x.salaire || '').replace(/[^\d]/g, '')) || ''})).filter(x => x.poste);
      d.prep = JSON.stringify({chiffres: C, materiel: M, texte: T, produits: PR, postes: PO}); d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)});
    }
    if (a === 'adminDocSuppr') { const d = dossier(q.code); supprimerLigne(docDe(d.code, q.id)); d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)}); }
    if (a === 'adminAcces') { const d = dossier(q.code); d.mdp = ''; d.sel = ''; d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)}); }   // le client recrée son mot de passe
    if (a === 'adminDoc') {
      const d = dossier(q.code);
      const x = docDe(d.code, q.id);
      if (!STAT_DOC.includes(q.statut)) throw new Error('Statut inconnu');
      x.statut = q.statut; x.remarque = q.remarque || ''; x.date = now(); put('Documents', x);
      d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)});
    }
    if (a === 'adminCoord') { const d = dossier(q.code); d.coord_envoi = now() + ' (' + (q.canal === 'sms' ? 'SMS' : 'WhatsApp') + ')'; d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)}); }
    if (a === 'adminPay') {
      const d = dossier(q.code);
      if (q.statut === 'مؤكد' && d.paiement !== 'مؤكد' && d.etude_url) { d.envoi = ''; d.envoi_date = ''; }   // à renvoyer avec le lien de téléchargement
      d.paiement = q.statut; d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)});
    }
    if (a === 'adminEtude') {   // étude finale : PDF uniquement
      const d = dossier(q.code);
      if (!/\.pdf$/i.test(q.name || '') || (q.mime && q.mime !== 'application/pdf')) throw new Error('L\'étude doit être un fichier PDF');
      q.mime = 'application/pdf';
      const f = saveFile(d, 'ETUDE_' + q.name, q, 40); f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      d.etude = q.name; d.etude_url = f.getUrl(); if (q.resume !== undefined) d.resume = q.resume;
      d.envoi = ''; d.envoi_date = ''; d.relances = 0; d.relance_date = ''; d.etape = STEPS.length;   // « Étude prête » + « prête à l'envoi »
      d.maj = now(); put('Dossiers', d); return out({ok: true, data: view(d, true)});
    }
    if (a === 'adminEnvoi') {
      const d = dossier(q.code); if (!d.etude_url) throw new Error('Déposez d\'abord l\'étude (PDF)');
      if (q.resume !== undefined) d.resume = q.resume;
      const canaux = [].concat(q.canaux || []), faits = [];
      if (canaux.includes('email')) { if (!d.email) throw new Error('Pas d\'email pour ce client'); envoyerEmail(d, false); faits.push('email'); }
      if (canaux.includes('whatsapp')) faits.push('WhatsApp');
      if (!faits.length) throw new Error('Choisissez email et/ou WhatsApp');
      d.envoi = faits.join(' + '); d.envoi_date = now(); d.relance_date = now(); d.maj = now(); put('Dossiers', d);
      return out({ok: true, data: view(d, true)});
    }
    if (a === 'adminAnswer') {
      const r = rows('Demandes').find(x => x.id === q.id); if (!r) throw new Error('Demande introuvable');
      r.reponse = q.reponse || ''; r.statut = 'تمت المعالجة'; r.date_reponse = now(); put('Demandes', r);
      return out({ok: true, data: view(dossier(r.code), true)});
    }
    throw new Error('Action inconnue');
  } catch (er) { return out({ok: false, error: er.message}); }
  finally { lock.releaseLock(); }
}

const SIM_DEF = {notaireTaux: 1, notaireMin: 15000, loyerFinance: true, loyerPlafond: 500000, frais: [
  {ar: 'فتح الحساب البنكي', fr: 'Ouverture du compte bancaire', m: 5000, etape: 'compte'}, {ar: 'السجل التجاري أو بطاقة الحرفي', fr: 'Registre de commerce ou carte d\'artisan', m: 10000, etape: 'fisc'},
  {ar: 'تسجيل عقد الإيجار', fr: 'Enregistrement du bail', m: 10000, etape: 'loc'}, {ar: 'طوابع ووثائق إدارية', fr: 'Timbres et documents', m: 5000, etape: 'fisc'}]};
function simCfg() { try { return Object.assign({}, SIM_DEF, JSON.parse(PropertiesService.getScriptProperties().getProperty('SIM_CFG') || '{}')); } catch (e) { return SIM_DEF; } }
function sauverFiche(d, v) {
  const f = ficheDe(d.code); v = v || {};
  if (v.nin !== undefined) { v.nin = String(v.nin).replace(/\D/g, ''); if (v.nin && v.nin.length !== 18) throw new Error({ar: 'رقم التعريف الوطني (NIN) يتكون من 18 رقماً', fr: 'Le NIN doit comporter 18 chiffres', en: 'The NIN must have 18 digits'}[d.lang || 'ar'] || 'NIN : 18 chiffres'); }
  FICHE.forEach(([k]) => { if (v[k] !== undefined) f[k] = String(v[k]).slice(0, 500); });
  f.code = d.code; f.maj = now(); put('Fiches', f); d.maj = now(); put('Dossiers', d);
}

/* ---------- messages (WhatsApp / email) dans la langue du client ---------- */
const lien = d => SITE_URL.replace(/\/$/, '') + '/';
function invitation(d) {   // message d'accès à l'espace client
  const l = d.lang || 'ar', u = lien(d);
  if (l === 'fr') return `Bonjour ${d.nom},\nSuivez l'étude de votre projet « ${d.projet} » remplissez vos coordonnées et envoyez votre facture proforma et votre devis d'assurance sur votre espace :\n${u}\nPremière fois : cliquez sur « Nouveau compte », entrez votre nom, ce numéro de téléphone et choisissez un mot de passe.\nمشروعك – Études NESDA · ${telAff('fr')}`;
  if (l === 'en') return `Hello ${d.nom},\nFollow the study of your project "${d.projet}" fill in your details and send your pro-forma invoice and insurance quote on your space:\n${u}\nFirst time: click “New account”, enter your name, this phone number and choose a password.\nمشروعك – NESDA studies · ${telAff('en')}`;
  return `السلام عليكم ${d.nom}،\nتابع دراسة مشروعك « ${d.projet} » واملأ بياناتك وأرسل الفاتورة الشكلية وعرض التأمين عبر فضائك :\n${u}\nأول مرة : اضغط على «حساب جديد»، اكتب اسمك وهذا الرقم واختر كلمة سر.\nمشروعك – دراسات وملفات NESDA · ${telAff('ar')}`;
}
function messageEtude(d, relance) {
  const l = d.lang || 'ar', paye = d.paiement === 'مؤكد', res = d.resume ? chiffres(String(d.resume).trim(), l) : '', p = nombre(PRIX, l), u = lien(d), t = telAff(l), h = horairesTxt(l);
  const T = {
    ar: {s: 'السلام عليكم ' + d.nom + '،', a: relance ? `🔔 تذكير : دراسة مشروعكم « ${d.projet} » جاهزة وفي انتظاركم.` : `🎉 يسعدنا إعلامكم أن دراسة مشروعكم « ${d.projet} » جاهزة.`, r: '📋 ملخص الدراسة :',
      ok: '✅ تم تأكيد الدفع. حمّلوا الدراسة (PDF) من هنا :', g: '📥 لاستلام الدراسة (PDF) :', p1: `1. ادفعوا مبلغ ${p} دج (تصلكم معلومات الدفع عبر SMS أو واتساب)`, p2: `2. أرسلوا صورة الوصل عبر فضائكم : ${u}`,
      c: `📞 أو اتصلوا بنا للحصول على الدراسة : ${t}`, h: `ساعات الاتصال : ${h}. إذا لم نرد فوراً، اتركوا رسالة وسنعيد الاتصال بكم.`, f: 'مشروعك – دراسات وملفات NESDA'},
    fr: {s: 'Bonjour ' + d.nom + ',', a: relance ? `🔔 Rappel : l'étude de votre projet « ${d.projet} » est prête et vous attend.` : `🎉 L'étude de votre projet « ${d.projet} » est prête.`, r: '📋 Résumé de l\'étude :',
      ok: '✅ Paiement confirmé. Téléchargez l\'étude (PDF) ici :', g: '📥 Pour recevoir l\'étude (PDF) :', p1: `1. Payez ${p} DA (les coordonnées de paiement vous sont envoyées par SMS ou WhatsApp)`, p2: `2. Envoyez la photo du reçu sur votre espace : ${u}`,
      c: `📞 Ou appelez-nous pour obtenir l'étude : ${t}`, h: `Heures d'appel : ${h}. Si nous ne répondons pas tout de suite, laissez un message, nous vous rappellerons.`, f: 'مشروعك – Études NESDA'},
    en: {s: 'Hello ' + d.nom + ',', a: relance ? `🔔 Reminder: the study of your project "${d.projet}" is ready and waiting for you.` : `🎉 The study of your project "${d.projet}" is ready.`, r: '📋 Study summary:',
      ok: '✅ Payment confirmed. Download the study (PDF) here:', g: '📥 To receive the study (PDF):', p1: `1. Pay ${p} DA (payment details are sent to you by SMS or WhatsApp)`, p2: `2. Send a photo of the receipt on your space: ${u}`,
      c: `📞 Or call us to get the study: ${t}`, h: `Call hours: ${h}. If we don't answer right away, leave a message and we will call you back.`, f: 'مشروعك – NESDA studies'},
  }[l] || {};
  const L = [T.s, T.a];
  if (res) L.push('', T.r, res);
  L.push('');
  if (paye) L.push(T.ok, d.etude_url); else L.push(T.g, T.p1, T.p2);
  L.push('', T.c, T.h, T.f);
  return L.join('\n');
}
function envoyerEmail(d, relance) {
  const l = d.lang || 'ar', txt = messageEtude(d, relance);
  const html = '<div dir="' + (l === 'ar' ? 'rtl' : 'ltr') + '" style="font-family:Tahoma,Arial,sans-serif;font-size:16px;line-height:1.8;color:#12302f">'
    + txt.split('\n').map(x => /^https?:\/\//.test(x.trim()) ? '<a href="' + x.trim() + '">' + x.trim() + '</a>' : x.replace(/[&<>]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;'}[c]))).join('<br>')
    + '<br><br><a href="tel:' + TEL_APPEL + '" style="background:#0e4d52;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none">📞 ' + telAff(l) + '</a></div>';
  const sujet = {ar: (relance ? 'تذكير – ' : '') + 'دراسة مشروعكم جاهزة – مشروعك', fr: (relance ? 'Rappel – ' : '') + 'Votre étude est prête – مشروعك', en: (relance ? 'Reminder – ' : '') + 'Your study is ready – مشروعك'}[l];
  MailApp.sendEmail({to: d.email, subject: sujet, body: txt, htmlBody: html, name: 'مشروعك – NESDA'});
}
function jours(s) { const m = String(s || '').match(/(\d+)\/(\d+)\/(\d+)/); if (!m) return 999; return (Date.now() - new Date(+m[3], m[2] - 1, +m[1]).getTime()) / 864e5; }
function etatEnvoi(d) {
  if (!d.etude_url) return {code: 'aucune', texte: 'Aucune étude'};
  if (d.paiement === 'مؤكد' && d.envoi) return {code: 'livree', texte: 'Payée – envoyée le ' + d.envoi_date + ' (' + d.envoi + ')'};
  if (!d.envoi) return {code: 'prete', texte: 'Prête à l\'envoi'};
  const r = Number(d.relances) || 0, late = jours(d.relance_date || d.envoi_date) >= RELANCE_JOURS;
  return {code: late ? 'relancer' : 'envoyee', texte: 'Envoyée le ' + d.envoi_date + ' (' + d.envoi + ')' + (r ? ' · ' + r + ' relance(s) email' : '') + (late ? ' · à relancer' : '')};
}
function relancesAuto() {   // chaque jour à 9h (déclencheur créé par setup)
  rows('Dossiers').forEach(d => {
    if (!d.etude_url || !d.envoi || d.paiement === 'مؤكد' || !d.email) return;
    if ((Number(d.relances) || 0) >= RELANCE_MAX || jours(d.relance_date || d.envoi_date) < RELANCE_JOURS) return;
    try { envoyerEmail(d, true); d.relances = (Number(d.relances) || 0) + 1; d.relance_date = now(); put('Dossiers', d); } catch (e) { Logger.log(d.code + ' : ' + e.message); }
  });
}
function installerRelances() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'relancesAuto').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('relancesAuto').timeBased().everyDays(1).atHour(9).inTimezone('Africa/Algiers').create();
}

/* ---------- fichiers ---------- */
function saveFile(d, name, q, maxMb) {
  const bytes = Utilities.base64Decode(q.b64), lim = (maxMb || MAX_MB);
  if (bytes.length > lim * 1024 * 1024) throw new Error('> ' + lim + ' Mo');
  const dir = folder(d.code + ' - ' + d.nom, folder(DRIVE_FOLDER, DriveApp.getRootFolder()));
  return dir.createFile(Utilities.newBlob(bytes, q.mime || 'application/octet-stream', name));
}
const prepDe = d => { try { return JSON.parse(d.prep || '{}') || {}; } catch (e) { return {}; } };
const quizDe = d => { try { return JSON.parse(d.quiz || '{}') || {}; } catch (e) { return {}; } };
const docId = y => String(y.id || 'r' + y._row);
function docDe(code, id) { const x = rows('Documents').find(y => norm(y.code) === norm(code) && docId(y) === String(id || '')); if (!x) throw new Error('Document introuvable'); return x; }
function supprimerLigne(x) {
  const m = String(x.url || '').match(/[-\w]{25,}/); if (m) try { DriveApp.getFileById(m[0]).setTrashed(true); } catch (e) {}
  sh('Documents').deleteRow(x._row);
}
function folder(name, parent) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function setup() {
  Object.keys(H).forEach(sh); const ss = book(); const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Feuille 1'); if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  folder(DRIVE_FOLDER, DriveApp.getRootFolder()); installerRelances(); Logger.log('Feuille : ' + ss.getUrl());
}
