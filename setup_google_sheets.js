// ═══════════════════════════════════════════════════════════════════════
//  T.A INVEST DASHBOARD — Google Apps Script de configuration
//  À coller dans : Extensions > Apps Script > Nouveau script
//  Puis cliquer : Exécuter > setupDashboard
// ═══════════════════════════════════════════════════════════════════════

function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  setupProjets(ss);
  setupEngagements(ss);
  setupParticipations(ss);
  setupRoadmap(ss);
  setupCommercial(ss);

  ui.alert('✅ Structure T.A Invest créée avec succès !\n\n' +
    'Onglets créés :\n' +
    '• PROJETS (avec colonne cf_gid)\n' +
    '• ENGAGEMENTS\n' +
    '• PARTICIPATIONS\n' +
    '• ROADMAP\n' +
    '• COMMERCIAL\n\n' +
    'Publie chaque onglet sur le web :\n' +
    'Fichier > Partager > Publier sur le web > sélectionner onglet > CSV');
}

// ── Couleurs ──────────────────────────────────────────────────────────
const TEAL   = '#1E4A44';
const TEAL2  = '#E8F2F1';
const ROSE   = '#C49878';
const GREY   = '#F0EDE8';
const WHITE  = '#FFFFFF';
const RED    = '#FCEAEA';
const GREEN  = '#E8F5EE';
const AMBER  = '#FDF3DC';

// ── Helper : crée ou vide un onglet ──────────────────────────────────
function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    sh.clearContents();
    sh.clearFormats();
  }
  return sh;
}

// ── Formate une ligne d'en-tête ───────────────────────────────────────
function styleHeader(range) {
  range.setBackground(TEAL)
       .setFontColor(WHITE)
       .setFontWeight('bold')
       .setFontSize(10)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(true);
  range.getSheet().setRowHeight(range.getRow(), 36);
}

// ── Formate une ligne de sous-en-tête ────────────────────────────────
function styleSubHeader(range) {
  range.setBackground(TEAL2)
       .setFontColor(TEAL)
       .setFontWeight('bold')
       .setFontSize(9)
       .setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════════════════════════════════
//  1. ONGLET PROJETS
//  Parser dashboard : parseRows() — colonnes A=r[0] … K=r[10]
// ═══════════════════════════════════════════════════════════════════════
function setupProjets(ss) {
  const sh = getOrCreateSheet(ss, 'PROJETS');

  // ── Ligne 1 : Titre ────────────────────────────────────────────────
  sh.getRange('A1:K1').merge()
    .setValue('T.A INVEST — PORTEFEUILLE PROJETS')
    .setBackground(TEAL).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);

  // ── Ligne 2 : En-têtes colonnes ───────────────────────────────────
  const headers = [
    'Projet',       // A = r[0]  — Nom complet du projet
    'Secteur',      // B = r[1]  — Immobilier / Énergie / Services…
    'Budget (MAD)', // C = r[2]  — Valeur brute en MAD (ex: 500000000)
    'Engagé (MAD)', // D = r[3]  — Montant déjà engagé en MAD
    'Reste (MAD)',  // E = r[4]  — Budget - Engagé
    'Besoin FP (MAD)', // F = r[5] — Fonds propres nécessaires
    'TRI',          // G = r[6]  — Ex: 25% ou 0.25
    'Payback',      // H = r[7]  — Ex: Fin 2028
    'Marge nette',  // I = r[8]  — Ex: 30% ou 0.30
    'Statut',       // J = r[9]  — En cours / À définir / Terminé
    'cf_gid'        // K = r[10] — GID de la feuille Cash Flow projet
  ];
  const hRange = sh.getRange(2, 1, 1, headers.length);
  hRange.setValues([headers]);
  styleHeader(hRange);

  // ── Ligne 3 : Instructions (grisée) ───────────────────────────────
  const instHeaders = [
    'Nom complet',
    'Ex: Immobilier',
    'Entier ex: 500000000',
    'Entier MAD',
    '=C3-D3',
    'Entier MAD',
    'Ex: 25% ou 0.25',
    'Ex: Fin 2028',
    'Ex: 30% ou 0.30',
    'En cours | À définir | Terminé',
    'Copier #gid=XXXXX de l\'URL'
  ];
  sh.getRange(3, 1, 1, instHeaders.length).setValues([instHeaders])
    .setBackground('#D9D9D9').setFontColor('#666666')
    .setFontStyle('italic').setFontSize(8)
    .setHorizontalAlignment('center');

  // ── Données exemple (3 projets) ───────────────────────────────────
  const data = [
    ['SARAYA VILLAS DESIGNED BY BENTLEY HOME', 'Immobilier', 750000000, 110000000, 640000000, 225000000, '22%', 'T4 2028', '30%', 'En cours',   '1488747750'],
    ['AZ SIGNATURE',                            'Immobilier', 200000000, 0,         200000000, 60000000,  '18%', 'Fin 2029', '25%', 'À définir', ''],
    ['JARDINS DE MESNANA',                      'Immobilier', 150000000, 0,         150000000, 45000000,  '17%', 'Fin 2029', '22%', 'À définir', ''],
  ];
  sh.getRange(4, 1, data.length, data[0].length).setValues(data);

  // ── Alternance couleurs lignes data ───────────────────────────────
  for (let i = 0; i < data.length; i++) {
    sh.getRange(4 + i, 1, 1, headers.length)
      .setBackground(i % 2 === 0 ? WHITE : GREY)
      .setFontSize(10).setVerticalAlignment('middle');
  }
  sh.setRowHeights(4, data.length, 28);

  // ── Validation Statut (colonne J) ────────────────────────────────
  const statutRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['En cours', 'À définir', 'Terminé', 'Suspendu'], true)
    .setAllowInvalid(false).build();
  sh.getRange('J4:J50').setDataValidation(statutRule);

  // ── Largeurs colonnes ─────────────────────────────────────────────
  const colWidths = [280, 100, 130, 130, 130, 140, 70, 90, 90, 100, 110];
  colWidths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ── Note explicative ─────────────────────────────────────────────
  sh.getRange('A1').setNote(
    '⚠️ IMPORTANT :\n' +
    '• Budget / Engagé / Reste / Besoin : valeurs en MAD bruts (pas en millions)\n' +
    '• TRI et Marge : écrire 25% ou 0.25 (les deux formats sont acceptés)\n' +
    '• cf_gid : coller le numéro après #gid= dans l\'URL de la feuille Cash Flow du projet\n' +
    '• Ne pas changer les noms d\'en-têtes (ligne 2)\n' +
    '• Ligne 3 (grisée) = aide-mémoire, ne pas supprimer'
  );

  Logger.log('✅ Onglet PROJETS configuré');
}

// ═══════════════════════════════════════════════════════════════════════
//  2. ONGLET ENGAGEMENTS
//  Parser : fetchEngagements() — colonnes A:E
// ═══════════════════════════════════════════════════════════════════════
function setupEngagements(ss) {
  const sh = getOrCreateSheet(ss, 'ENGAGEMENTS');

  sh.getRange('A1:E1').merge()
    .setValue('T.A INVEST — ENGAGEMENTS FINANCIERS')
    .setBackground(TEAL).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);

  const headers = [
    'Projet',       // A = r[0]  — Doit correspondre exactement au nom dans PROJETS
    'Nature',       // B = r[1]  — Ex: Achat foncier / Contrat entreprise
    'Montant (MAD)',// C = r[2]  — En MAD (ex: 30000000) OU en M (ex: 30)
    'Date',         // D = r[3]  — Ex: Juin 2026
    'Statut'        // E = r[4]  — Engagé / En cours / Planifié
  ];
  const hRange = sh.getRange(2, 1, 1, headers.length);
  hRange.setValues([headers]);
  styleHeader(hRange);

  const data = [
    ['SARAYA VILLAS DESIGNED BY BENTLEY HOME', 'Achat foncier',          30000000,  'Juin 2026',  'Engagé'],
    ['SARAYA VILLAS DESIGNED BY BENTLEY HOME', 'Contrat entreprise',      80000000,  'Sept. 2026', 'En cours'],
    ['SARAYA VILLAS DESIGNED BY BENTLEY HOME', 'Études & conception',      8000000,  'Mars 2026',  'Engagé'],
    ['AZ SIGNATURE',                            'Acquisition terrain',     40000000,  'T3 2026',    'Planifié'],
    ['JARDINS DE MESNANA',                      'Études préliminaires',     2000000,  'T4 2026',    'Planifié'],
  ];
  sh.getRange(3, 1, data.length, data[0].length).setValues(data);

  for (let i = 0; i < data.length; i++) {
    sh.getRange(3 + i, 1, 1, headers.length)
      .setBackground(i % 2 === 0 ? WHITE : GREY).setFontSize(10);
  }
  sh.setRowHeights(3, data.length, 26);

  // Validation statut
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Engagé', 'En cours', 'Planifié', 'Annulé'], true)
    .setAllowInvalid(false).build();
  sh.getRange('E3:E100').setDataValidation(rule);

  sh.setColumnWidths(1, 1, 300);
  sh.setColumnWidths(2, 1, 180);
  sh.setColumnWidths(3, 1, 130);
  sh.setColumnWidths(4, 1, 110);
  sh.setColumnWidths(5, 1, 100);

  sh.getRange('A1').setNote(
    '⚠️ IMPORTANT :\n' +
    '• La colonne Projet doit correspondre EXACTEMENT (ou partiellement) au nom dans PROJETS\n' +
    '• Montant : peut être en MAD bruts (30000000) ou en M MAD (30) — le dashboard gère les deux\n' +
    '• Statut : Engagé | En cours | Planifié | Annulé'
  );

  Logger.log('✅ Onglet ENGAGEMENTS configuré');
}

// ═══════════════════════════════════════════════════════════════════════
//  3. ONGLET PARTICIPATIONS
//  Parser : fetchParticipations() — commence ligne 3 (index 2), colonnes A:L
// ═══════════════════════════════════════════════════════════════════════
function setupParticipations(ss) {
  const sh = getOrCreateSheet(ss, 'PARTICIPATIONS');

  // Ligne 1 : titre
  sh.getRange('A1:L1').merge()
    .setValue('T.A INVEST — PORTEFEUILLE PARTICIPATIONS')
    .setBackground(TEAL).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);

  // Ligne 2 : sous-titre / légende (ignorée par le parser — commence ligne 3)
  sh.getRange('A2:L2').merge()
    .setValue('↓ Données à partir de la ligne 3 (ligne 2 = libre pour commentaires)')
    .setBackground(TEAL2).setFontColor(TEAL)
    .setFontStyle('italic').setFontSize(9)
    .setHorizontalAlignment('center');

  // Ligne 3 : en-têtes (index 1 = ignoré, index 2+ = lu)
  // Le parser démarre à i=2, donc ligne 3 = header ignoré si pas de chiffres,
  // ligne 4+ = données
  const headers = [
    'Nom participation', // A = r[0]
    'Type',              // B = r[1] — Participation opérationnelle / Fonds PE / Club Deal…
    'Secteur',           // C = r[2] — Immobilier / Énergie / Industrie…
    '% Détenu',          // D = r[3] — Ex: 25% ou 0.25
    'Coût acq. (MAD)',   // E = r[4] — En MAD bruts (ex: 15000000)
    'Val. actuelle (MAD)',// F = r[5]
    'Plus-value (MAD)',  // G = r[6] — Laisser vide = calculé auto (Val - Coût)
    'Dividendes (MAD)',  // H = r[7]
    'MOIC',              // I = r[8] — Ex: 1.55 — laisser vide = calculé auto
    'TRI',               // J = r[9] — Ex: 18% ou 0.18
    'Horizon sortie',    // K = r[10]— Ex: 2028
    'Statut',            // L = r[11]— Actif / En cours de cession / Sorti
    'bp_gid'             // M = r[12]— GID de la feuille BP_PART_* (détail participation)
  ];
  const hRange = sh.getRange(3, 1, 1, headers.length);
  hRange.setValues([headers]);
  styleHeader(hRange);

  const data = [
    ['ABC Holding SA',           'Participation opérationnelle', 'Immobilier', '25%', 15000000, 22000000, 7000000,  1200000, 1.55, '18%', '2028', 'Actif',                ''],
    ['XYZ Renewable Fund',       'Fonds PE / OPCI',              'Énergie',    '10%', 8000000,  10500000, 2500000,  500000,  1.38, '15%', '2027', 'Actif',                ''],
    ['Immo Club Deal Marrakech', 'Club Deal immobilier',         'Immobilier', '33%', 5000000,  6200000,  1200000,  200000,  1.28, '14%', '2026', 'En cours de cession',  ''],
    ['DEF Industrie SARL',       'Participation opérationnelle', 'Industrie',  '40%', 12000000, 14000000, 2000000,  800000,  1.23, '12%', '2029', 'Actif',                ''],
    ['Obligations Privées GHI',  'Dette privée / Obligations',  'Services',   '100%',6000000,  6300000,  300000,   420000,  1.12, '7%',  '2027', 'Actif',                ''],
  ];
  sh.getRange(4, 1, data.length, data[0].length).setValues(data);

  for (let i = 0; i < data.length; i++) {
    sh.getRange(4 + i, 1, 1, headers.length)
      .setBackground(i % 2 === 0 ? WHITE : GREY).setFontSize(10);
  }
  sh.setRowHeights(4, data.length, 28);

  // Validation Statut
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Actif', 'En cours de cession', 'Sorti', 'Suspendu'], true)
    .setAllowInvalid(false).build();
  sh.getRange('L4:L50').setDataValidation(rule);

  const colW = [200, 180, 100, 80, 140, 150, 130, 130, 70, 70, 100, 140];
  colW.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('A1').setNote(
    '⚠️ IMPORTANT :\n' +
    '• Les données commencent ligne 4 (ligne 3 = en-têtes ignorés par le parser)\n' +
    '• Coût / Val / PV / Dividendes : en MAD bruts\n' +
    '• % Détenu : 25% ou 0.25 (les deux formats acceptés)\n' +
    '• Plus-value et MOIC : laissez vide = calculé automatiquement\n' +
    '• TRI : 18% ou 0.18'
  );

  Logger.log('✅ Onglet PARTICIPATIONS configuré');
}

// ═══════════════════════════════════════════════════════════════════════
//  4. ONGLET ROADMAP
//  Parser : fetchRoadmap() — type BAR ou ETAPE
// ═══════════════════════════════════════════════════════════════════════
function setupRoadmap(ss) {
  const sh = getOrCreateSheet(ss, 'ROADMAP');

  sh.getRange('A1:F1').merge()
    .setValue('T.A INVEST — ROADMAP STRATÉGIQUE 2026-2030')
    .setBackground(TEAL).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);

  const headers = ['Type', 'Axe / Date', 'Label / Description', 'Début (1-10)', 'Fin (1-10)', 'Couleur HEX'];
  const hRange = sh.getRange(2, 1, 1, 6);
  hRange.setValues([headers]);
  styleHeader(hRange);

  // Ligne instruction
  sh.getRange('A3:F3').setValues([[
    'BAR ou ETAPE',
    'BAR=axe thématique, ETAPE=date',
    'Texte affiché sur la barre',
    'Semestre départ (1=S1 2026)',
    'Semestre fin (10=S2 2030)',
    'Sans # ex: 1E4A44'
  ]]).setBackground('#D9D9D9').setFontColor('#666666').setFontStyle('italic').setFontSize(8);

  // Données exemple
  const data = [
    // BARRES (axe + semestres)
    ['BAR', 'Acquisitions Foncieres',       'Saraya Villas',          1, 2,  '1E4A44'],
    ['BAR', 'Acquisitions Foncieres',       'AZ Signature',           3, 4,  '2B5F58'],
    ['BAR', 'Etudes et Autorisations',      'Permis Saraya',          1, 3,  '4A8C84'],
    ['BAR', 'Etudes et Autorisations',      'Études AZ Sig.',         4, 5,  '4A8C84'],
    ['BAR', 'Developpement Construction',   'Construction Saraya',    3, 6,  'C49878'],
    ['BAR', 'Developpement Construction',   'Construction AZ',        5, 8,  'D4AA8A'],
    ['BAR', 'Commercialisation',            'Ventes Saraya',          4, 7,  '2E7D55'],
    ['BAR', 'Commercialisation',            'Ventes AZ Sig.',         6, 9,  '2E7D55'],
    ['BAR', 'Financements',                 'Crédit Saraya',          2, 4,  'C4920A'],
    ['BAR', 'Optimisation Valeur',          'Asset management',       7, 10, '1E4A44'],
    // ÉTAPES CLÉS
    ['ETAPE', 'T2 2026', 'Lancement commercial Saraya', '', '', ''],
    ['ETAPE', 'T4 2026', 'Démarrage construction Saraya', '', '', ''],
    ['ETAPE', 'T2 2028', 'Livraison Phase 1 Saraya', '', '', ''],
    ['ETAPE', 'T1 2029', 'Livraison finale Saraya', '', '', ''],
  ];
  sh.getRange(4, 1, data.length, 6).setValues(data);

  for (let i = 0; i < data.length; i++) {
    const type = data[i][0];
    const bg = type === 'BAR' ? (i % 2 === 0 ? WHITE : GREY) : AMBER;
    sh.getRange(4 + i, 1, 1, 6).setBackground(bg).setFontSize(10);
  }
  sh.setRowHeights(4, data.length, 26);

  // Validation Type
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['BAR', 'ETAPE'], true).setAllowInvalid(false).build();
  sh.getRange('A4:A50').setDataValidation(rule);

  sh.setColumnWidths(1, 1, 80);
  sh.setColumnWidths(2, 1, 200);
  sh.setColumnWidths(3, 1, 220);
  sh.setColumnWidths(4, 1, 100);
  sh.setColumnWidths(5, 1, 100);
  sh.setColumnWidths(6, 1, 110);

  sh.getRange('A1').setNote(
    '⚠️ GUIDE ROADMAP :\n\n' +
    'Semestres (colonne Début/Fin) :\n' +
    '  1 = S1 2026 | 2 = S2 2026\n' +
    '  3 = S1 2027 | 4 = S2 2027\n' +
    '  5 = S1 2028 | 6 = S2 2028\n' +
    '  7 = S1 2029 | 8 = S2 2029\n' +
    '  9 = S1 2030 | 10 = S2 2030\n\n' +
    'Couleur HEX : sans le # (ex: 1E4A44)\n\n' +
    'Axes reconnus automatiquement :\n' +
    '  Acquisitions Foncieres\n' +
    '  Etudes et Autorisations\n' +
    '  Developpement Construction\n' +
    '  Commercialisation\n' +
    '  Financements\n' +
    '  Optimisation Valeur'
  );

  Logger.log('✅ Onglet ROADMAP configuré');
}

// ═══════════════════════════════════════════════════════════════════════
//  5. ONGLET COMMERCIAL
//  Parser : fetchCommercial() — type U (Unité) / P (Prospect) / E (Échéancier)
// ═══════════════════════════════════════════════════════════════════════
function setupCommercial(ss) {
  const sh = getOrCreateSheet(ss, 'COMMERCIAL');

  sh.getRange('A1:I1').merge()
    .setValue('T.A INVEST — SUIVI COMMERCIAL')
    .setBackground(TEAL).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);

  // Légende des types
  sh.getRange('A2:I2').merge()
    .setValue('Type U = Unité vendue/disponible  |  Type P = Prospect  |  Type E = Échéancier paiement')
    .setBackground(TEAL2).setFontColor(TEAL)
    .setFontStyle('italic').setFontSize(9)
    .setHorizontalAlignment('center');

  // En-têtes (ligne 3)
  const headers = ['Type', 'Col 1', 'Col 2', 'Col 3', 'Col 4', 'Col 5', 'Col 6', 'Col 7', 'Col 8'];
  const hRange = sh.getRange(3, 1, 1, 9);
  hRange.setValues([headers]);
  styleHeader(hRange);

  // Sous-en-têtes par type (ligne 4 — commentaire)
  sh.getRange('A4').setValue('— UNITÉS (U) →');
  sh.getRange('B4').setValue('Projet');
  sh.getRange('C4').setValue('Réf');
  sh.getRange('D4').setValue('Type logement');
  sh.getRange('E4').setValue('Surface m²');
  sh.getRange('F4').setValue('Prix (M MAD)');
  sh.getRange('G4').setValue('Client');
  sh.getRange('H4').setValue('Date');
  sh.getRange('I4').setValue('Statut');
  sh.getRange('A4:I4').setBackground(GREEN).setFontColor('#2E7D55').setFontWeight('bold').setFontSize(9);

  // Données Unités
  const unites = [
    ['U', 'Saraya Villas', 'SV-01', 'Villa Premium',    450, 8.5,  'M. Al-Rashidi',    'Jan 2026', 'Vendu'],
    ['U', 'Saraya Villas', 'SV-02', 'Villa Premium',    420, 8.0,  'Famille Benali',   'Fév 2026', 'Vendu'],
    ['U', 'Saraya Villas', 'SV-03', 'Villa Signature',  520, 10.5, 'En cours',         'Mar 2026', 'Compromis'],
    ['U', 'Saraya Villas', 'SV-04', 'Villa Premium',    410, 7.8,  '',                 '',         'Disponible'],
    ['U', 'Saraya Villas', 'SV-05', 'Villa Signature',  580, 11.2, '',                 '',         'Disponible'],
  ];
  sh.getRange(5, 1, unites.length, 9).setValues(unites)
    .setBackground(WHITE).setFontSize(10);
  sh.setRowHeights(5, unites.length, 26);

  // Séparateur Prospects
  const rowP = 5 + unites.length;
  sh.getRange(rowP, 1).setValue('— PROSPECTS (P) →');
  sh.getRange(rowP, 2).setValue('Nom prospect');
  sh.getRange(rowP, 3).setValue('Projet ciblé');
  sh.getRange(rowP, 4).setValue('Budget (M MAD)');
  sh.getRange(rowP, 5).setValue('Statut');
  sh.getRange(rowP, 6).setValue('Action suivante');
  sh.getRange(rowP, 1, 1, 9).setBackground(AMBER).setFontColor('#C4920A').setFontWeight('bold').setFontSize(9);

  const prospects = [
    ['P', 'Al-Mansouri Family', 'Saraya Villas', 9.5,  'Intéressé',    'Visite planifiée'],
    ['P', 'Invest Dubai LLC',    'Saraya Villas', 25.0, 'En négociation','Envoi dossier'],
    ['P', 'M. Benkirane',        'AZ Signature',  5.0,  'Contact initial','Rappel T3 2026'],
  ];
  sh.getRange(rowP + 1, 1, prospects.length, 6).setValues(prospects)
    .setBackground(WHITE).setFontSize(10);
  sh.setRowHeights(rowP + 1, prospects.length, 26);

  // Séparateur Échéancier
  const rowE = rowP + 1 + prospects.length;
  sh.getRange(rowE, 1).setValue('— ÉCHÉANCIER (E) →');
  sh.getRange(rowE, 2).setValue('Réf');
  sh.getRange(rowE, 3).setValue('Client');
  sh.getRange(rowE, 4).setValue('Montant (M MAD)');
  sh.getRange(rowE, 5).setValue('Échéance');
  sh.getRange(rowE, 6).setValue('Statut');
  sh.getRange(rowE, 1, 1, 9).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold').setFontSize(9);

  const echeancier = [
    ['E', 'SV-01', 'M. Al-Rashidi',  2.5,  'Mars 2026',  'Payé'],
    ['E', 'SV-01', 'M. Al-Rashidi',  3.0,  'Juin 2026',  'À venir'],
    ['E', 'SV-02', 'Famille Benali', 2.0,  'Avr. 2026',  'Payé'],
    ['E', 'SV-03', 'En cours',       2.625,'Juil. 2026',  'En attente'],
  ];
  sh.getRange(rowE + 1, 1, echeancier.length, 6).setValues(echeancier)
    .setBackground(WHITE).setFontSize(10);
  sh.setRowHeights(rowE + 1, echeancier.length, 26);

  const colW = [80, 180, 160, 150, 100, 160, 140, 90, 110];
  colW.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('A1').setNote(
    '⚠️ GUIDE COMMERCIAL :\n\n' +
    'Type U (Unité) :\n' +
    '  B=Projet | C=Réf | D=Type | E=Surface m²\n' +
    '  F=Prix M MAD | G=Client | H=Date | I=Statut\n' +
    '  Statut : Vendu | Compromis | Disponible | Réservé\n\n' +
    'Type P (Prospect) :\n' +
    '  B=Nom | C=Projet | D=Budget M MAD\n' +
    '  E=Statut | F=Action suivante\n\n' +
    'Type E (Échéancier) :\n' +
    '  B=Réf | C=Client | D=Montant M MAD\n' +
    '  E=Échéance | F=Statut (Payé / À venir / En attente)'
  );

  Logger.log('✅ Onglet COMMERCIAL configuré');
}

// ═══════════════════════════════════════════════════════════════════════
//  6. CRÉER UN ONGLET CASH FLOW POUR UN PROJET
//  À appeler manuellement ou adapter par projet
//  Structure lue par fetchCashFlowForProject(gid)
// ═══════════════════════════════════════════════════════════════════════
function createCashFlowSheet(ss, nomProjet, co2026, co2027, co2028, co2029, co2030,
                                              ci2026, ci2027, ci2028, ci2029, ci2030) {
  const shName = nomProjet.substring(0, 28) + '_CF';
  const sh = getOrCreateSheet(ss, shName);

  sh.getRange('A1:G1').merge()
    .setValue('C. CASH FLOW PRÉVISIONNEL (M MAD) — Saisie directe')
    .setBackground(TEAL).setFontColor(WHITE).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 38);

  // En-têtes (ligne 1 lue par le parser comme row[0])
  const headers = ['Flux', '2026', '2027', '2028', '2029', '2030', 'TOTAL'];
  sh.getRange(1, 1, 1, 7).setValues([headers]); // row index 0 du CSV

  // Ligne 2 : Décaissements (cashOut) — valeurs en MAD (ex: -100000000)
  const coTotal = co2026 + co2027 + co2028 + co2029 + co2030;
  sh.getRange(2, 1, 1, 7).setValues([[
    'Décaissements invest. (M)',
    co2026, co2027, co2028, co2029, co2030, coTotal
  ]]).setBackground(RED).setFontSize(10);

  // Ligne 3 : Encaissements (cashIn) — valeurs en MAD
  const ciTotal = ci2026 + ci2027 + ci2028 + ci2029 + ci2030;
  sh.getRange(3, 1, 1, 7).setValues([[
    'Encaissements ventes',
    ci2026, ci2027, ci2028, ci2029, ci2030, ciTotal
  ]]).setBackground(GREEN).setFontSize(10);

  // Ligne 4 : Flux net — formule
  sh.getRange(4, 1).setValue('FLUX NET ANNUEL');
  for (let c = 2; c <= 7; c++) {
    sh.getRange(4, c).setFormula(`=B3+B2`.replace(/B/g, String.fromCharCode(64 + c)));
  }
  sh.getRange(4, 1, 1, 7).setFontWeight('bold').setBackground(TEAL2).setFontSize(10);

  sh.setColumnWidths(1, 1, 220);
  sh.setColumnWidths(2, 6, 110);

  const gid = sh.getSheetId();
  Logger.log(`✅ Feuille CF créée : ${shName} — GID = ${gid}`);
  Logger.log(`   → Copier ce GID dans la colonne K (cf_gid) de PROJETS`);
  return gid;
}

// ═══════════════════════════════════════════════════════════════════════
//  CRÉER LES FEUILLES CASH FLOW DES PROJETS RÉELS
//  Adapter les valeurs selon vos données
// ═══════════════════════════════════════════════════════════════════════
function createAllCashFlowSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // SARAYA VILLAS — valeurs en MAD (négatif pour décaissements)
  const gidSaraya = createCashFlowSheet(ss, 'SARAYA VILLAS',
    -100000000, -300000000, -350000000, 0, 0,   // Cash-out 2026-2030
     25000000,  150000000,  500000000, 300000000, 0   // Cash-in 2026-2030
  );

  // AZ SIGNATURE — à adapter
  const gidAZ = createCashFlowSheet(ss, 'AZ SIGNATURE',
    -70000000, -80000000, -50000000, 0, 0,
     0, 0, 100000000, 160000000, 60000000
  );

  // JARDINS DE MESNANA — à adapter
  const gidJardins = createCashFlowSheet(ss, 'JARDINS DE MESNANA',
    -50000000, -60000000, -40000000, 0, 0,
     0, 0, 80000000, 100000000, 60000000
  );

  SpreadsheetApp.getUi().alert(
    '✅ Feuilles Cash Flow créées !\n\n' +
    'GIDs à copier dans colonne K (cf_gid) de l\'onglet PROJETS :\n\n' +
    '• SARAYA VILLAS  → ' + gidSaraya + '\n' +
    '• AZ SIGNATURE   → ' + gidAZ + '\n' +
    '• JARDINS MESNANA→ ' + gidJardins + '\n\n' +
    'Puis publie ces onglets :\nFichier > Partager > Publier sur le web > CSV'
  );
}

// ═══════════════════════════════════════════════════════════════════════
//  7. BP DÉTAILLÉ PAR PROJET (sections A→F)
//  Structure lue par fetchProjectBP(gid) dans le dashboard
//  Sections identifiées par marqueur ##SECTION en colonne A
// ═══════════════════════════════════════════════════════════════════════
function createProjectBPSheet(ss, nomProjet, d) {
  const shName = ('BP_' + nomProjet).substring(0, 31);
  const sh = getOrCreateSheet(ss, shName);

  // ── Titre ─────────────────────────────────────────────────────────
  sh.getRange('A1:H1').merge()
    .setValue('BP — ' + nomProjet.toUpperCase())
    .setBackground(TEAL).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  let row = 2;

  // ── §A — KPIs Synthèse ───────────────────────────────────────────
  sh.getRange(row, 1).setValue('##A_KPIS');
  sh.getRange(row, 1, 1, 8).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('A. KPIs Synthèse');
  row++;

  const kpis = [
    ['Avancement',        d.avancement||0,   '%'],
    ['Budget Total',      d.budget||0,        'MAD'],
    ['CA Prévisionnel',   d.ca_prev||0,       'MAD'],
    ['Fonds Propres',     d.fp||0,            'MAD'],
    ['MOIC Cible',        d.moic||1.5,        'x'],
    ['Date livraison',    d.livraison||'—',   ''],
  ];
  sh.getRange(row, 1, 1, 3).setValues([['KPI','Valeur','Unité']]);
  sh.getRange(row, 1, 1, 3).setBackground(GREY).setFontWeight('bold').setFontSize(9);
  row++;
  kpis.forEach(k => { sh.getRange(row++, 1, 1, 3).setValues([k]).setFontSize(10); });
  row++;

  // ── §B — Plan d'Investissement ───────────────────────────────────
  sh.getRange(row, 1).setValue('##B_INVEST');
  sh.getRange(row, 1, 1, 8).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('B. Plan d\'Investissement');
  row++;

  sh.getRange(row, 1, 1, 3).setValues([['Poste','Montant (MAD)','% Budget']]);
  sh.getRange(row, 1, 1, 3).setBackground(GREY).setFontWeight('bold').setFontSize(9);
  row++;
  const invest = d.invest || [
    ['Acquisition foncière',  d.budget*0.30, '30%'],
    ['Études & conception',   d.budget*0.05, '5%'],
    ['Travaux construction',  d.budget*0.45, '45%'],
    ['Commercialisation',     d.budget*0.08, '8%'],
    ['Frais financiers',      d.budget*0.07, '7%'],
    ['Imprévus',              d.budget*0.05, '5%'],
  ];
  invest.forEach((r, i) => {
    sh.getRange(row, 1, 1, 3).setValues([r])
      .setBackground(i%2===0?WHITE:GREY).setFontSize(10);
    row++;
  });
  // Total invest
  const totInvest = 'TOTAL INVESTISSEMENT';
  sh.getRange(row, 1).setValue(totInvest);
  sh.getRange(row, 2).setFormula('=SUM(B'+(row-invest.length)+':B'+(row-1)+')');
  sh.getRange(row, 3).setValue('100%');
  sh.getRange(row, 1, 1, 3).setFontWeight('bold').setBackground(TEAL2).setFontColor(TEAL);
  row += 2;

  // ── §C — Cash Flow Prévisionnel ──────────────────────────────────
  sh.getRange(row, 1).setValue('##C_CASHFLOW');
  sh.getRange(row, 1, 1, 8).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('C. Cash Flow Prévisionnel (M MAD) — Saisie directe');
  row++;

  sh.getRange(row, 1, 1, 7).setValues([['Flux','2026','2027','2028','2029','2030','TOTAL']]);
  sh.getRange(row, 1, 1, 7).setBackground(GREY).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  row++;

  const cf = d.cf || { co:[-0,-0,-0,-0,-0], ci:[0,0,0,0,0] };
  const coRow2 = row;
  sh.getRange(row, 1, 1, 7).setValues([[
    'Décaissements invest. (M)',
    cf.co[0]||0, cf.co[1]||0, cf.co[2]||0, cf.co[3]||0, cf.co[4]||0,
    (cf.co[0]||0)+(cf.co[1]||0)+(cf.co[2]||0)+(cf.co[3]||0)+(cf.co[4]||0)
  ]]).setBackground(RED).setFontSize(10);
  row++;

  const ciRow2 = row;
  sh.getRange(row, 1, 1, 7).setValues([[
    'Encaissements ventes',
    cf.ci[0]||0, cf.ci[1]||0, cf.ci[2]||0, cf.ci[3]||0, cf.ci[4]||0,
    (cf.ci[0]||0)+(cf.ci[1]||0)+(cf.ci[2]||0)+(cf.ci[3]||0)+(cf.ci[4]||0)
  ]]).setBackground(GREEN).setFontSize(10);
  row++;

  sh.getRange(row, 1).setValue('FLUX NET ANNUEL');
  for (let c = 2; c <= 7; c++) {
    const col = String.fromCharCode(64+c);
    sh.getRange(row, c).setFormula('='+col+ciRow2+'+'+col+coRow2);
  }
  sh.getRange(row, 1, 1, 7).setFontWeight('bold').setBackground(TEAL2).setFontColor(TEAL);
  row += 2;

  // ── §D — Plan de Financement ─────────────────────────────────────
  sh.getRange(row, 1).setValue('##D_FINANCEMENT');
  sh.getRange(row, 1, 1, 8).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('D. Plan de Financement');
  row++;

  sh.getRange(row, 1, 1, 4).setValues([['Source','Montant (MAD)','%','Conditions']]);
  sh.getRange(row, 1, 1, 4).setBackground(GREY).setFontWeight('bold').setFontSize(9);
  row++;
  const fin = d.financement || [
    ['Fonds propres',   d.budget*0.30, '30%', 'Apport actionnaire'],
    ['Crédit bancaire', d.budget*0.53, '53%', 'Taux 5% / 7 ans'],
    ['Préventes',       d.budget*0.17, '17%', 'Compromis signés'],
  ];
  fin.forEach((f, i) => {
    sh.getRange(row, 1, 1, 4).setValues([f])
      .setBackground(i%2===0?WHITE:GREY).setFontSize(10);
    row++;
  });
  sh.getRange(row, 1).setValue('TOTAL');
  sh.getRange(row, 2).setFormula('=SUM(B'+(row-fin.length)+':B'+(row-1)+')');
  sh.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground(TEAL2).setFontColor(TEAL);
  row += 2;

  // ── §E — Compte de Résultat ──────────────────────────────────────
  sh.getRange(row, 1).setValue('##E_RESULTAT');
  sh.getRange(row, 1, 1, 8).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('E. Compte de Résultat Prévisionnel (MAD)');
  row++;

  sh.getRange(row, 1, 1, 5).setValues([['Exercice','CA','Charges','EBITDA','Résultat Net']]);
  sh.getRange(row, 1, 1, 5).setBackground(GREY).setFontWeight('bold').setFontSize(9);
  row++;
  const pnl = d.pnl || [
    [2026,  cf.ci[0]||0,  Math.abs(cf.co[0]||0)*1.1,  0, 0],
    [2027,  cf.ci[1]||0,  Math.abs(cf.co[1]||0)*0.4,  0, 0],
    [2028,  cf.ci[2]||0,  Math.abs(cf.co[2]||0)*0.3,  0, 0],
    [2029,  cf.ci[3]||0,  Math.abs(cf.co[3]||0)*0.2,  0, 0],
    [2030,  cf.ci[4]||0,  Math.abs(cf.co[4]||0)*0.5,  0, 0],
  ];
  pnl.forEach((r, i) => {
    const pnlRow = row;
    sh.getRange(row, 1, 1, 3).setValues([[r[0], r[1], r[2]]]);
    // EBITDA = CA - Charges (formula)
    sh.getRange(row, 4).setFormula('=B'+pnlRow+'-C'+pnlRow);
    // RN = EBITDA * 0.93 (estimation IS)
    sh.getRange(row, 5).setFormula('=D'+pnlRow+'*0.93');
    sh.getRange(row, 1, 1, 5).setBackground(i%2===0?WHITE:GREY).setFontSize(10);
    row++;
  });
  row++;

  // ── §F — Analyse Rentabilité ─────────────────────────────────────
  sh.getRange(row, 1).setValue('##F_RENTABILITE');
  sh.getRange(row, 1, 1, 8).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('F. Analyse de Rentabilité');
  row++;

  sh.getRange(row, 1, 1, 4).setValues([['Indicateur','Valeur','Objectif','Statut']]);
  sh.getRange(row, 1, 1, 4).setBackground(GREY).setFontWeight('bold').setFontSize(9);
  row++;
  const rent = d.rentabilite || [
    ['TRI',       (d.tri||'22%'),      '≥ 15%',    'BON'],
    ['VAN (10%)', (d.van||'95M MAD'),  '> 0',      'BON'],
    ['MOIC',      (d.moic||'1.50x'),   '≥ 1.30x',  'BON'],
    ['Payback',   (d.payback||'T4 2028'),'< 4 ans', 'BON'],
    ['Marge nette',(d.marge||'30%'),   '≥ 20%',    'BON'],
  ];
  rent.forEach((r, i) => {
    sh.getRange(row, 1, 1, 4).setValues([r])
      .setBackground(i%2===0?WHITE:GREY).setFontSize(10);
    row++;
  });

  // ── Mise en forme globale ─────────────────────────────────────────
  sh.setColumnWidths(1, 1, 230);
  sh.setColumnWidths(2, 1, 150);
  sh.setColumnWidths(3, 1, 130);
  sh.setColumnWidths(4, 1, 130);
  sh.setColumnWidths(5, 1, 130);
  sh.setColumnWidths(6, 1, 130);
  sh.setColumnWidths(7, 1, 130);
  sh.setColumnWidths(8, 1, 130);

  const gid = sh.getSheetId();
  Logger.log('✅ BP Projet créé : ' + shName + ' — GID = ' + gid);
  return gid;
}

// ═══════════════════════════════════════════════════════════════════════
//  8. BP DÉTAILLÉ PAR PARTICIPATION (sections A→D)
//  Structure lue par fetchParticipationBP(gid) dans le dashboard
// ═══════════════════════════════════════════════════════════════════════
function createParticipationBPSheet(ss, nomPart, d) {
  const shName = ('BP_PART_' + nomPart).substring(0, 31);
  const sh = getOrCreateSheet(ss, shName);

  sh.getRange('A1:G1').merge()
    .setValue('BP PARTICIPATION — ' + nomPart.toUpperCase())
    .setBackground(TEAL).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  let row = 2;

  // ── §A — Fiche & KPIs ────────────────────────────────────────────
  sh.getRange(row, 1).setValue('##A_FICHE');
  sh.getRange(row, 1, 1, 7).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('A. Fiche Participation');
  row++;

  const ficheData = [
    ['Type',          d.type||'—'],
    ['Secteur',       d.secteur||'—'],
    ['% Détenu',      d.pct||'—'],
    ['Coût acq.',     d.cout||0],
    ['Valeur actuelle',d.val||0],
    ['Plus-value',    d.pv||0],
    ['Dividendes reçus',d.div||0],
    ['MOIC',          d.moic||'—'],
    ['TRI',           d.tri||'—'],
    ['Horizon sortie',d.horizon||'—'],
    ['Statut',        d.statut||'Actif'],
  ];
  ficheData.forEach((r, i) => {
    sh.getRange(row, 1, 1, 2).setValues([r])
      .setBackground(i%2===0?WHITE:GREY).setFontSize(10);
    row++;
  });
  row++;

  // ── §B — Flux Historiques ─────────────────────────────────────────
  sh.getRange(row, 1).setValue('##B_FLUX_HIST');
  sh.getRange(row, 1, 1, 7).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('B. Flux Historiques');
  row++;

  sh.getRange(row, 1, 1, 4).setValues([['Année','Dividendes (MAD)','Variation NAV (MAD)','Retour Total']]);
  sh.getRange(row, 1, 1, 4).setBackground(GREY).setFontWeight('bold').setFontSize(9);
  row++;
  const hist = d.fluxHist || [
    [2023, d.div*0.3||0, d.pv*0.4||0, 0],
    [2024, d.div*0.4||0, d.pv*0.3||0, 0],
    [2025, d.div*0.3||0, d.pv*0.3||0, 0],
  ];
  hist.forEach((h, i) => {
    const hRow = row;
    sh.getRange(row, 1, 1, 3).setValues([[h[0], h[1], h[2]]]);
    sh.getRange(row, 4).setFormula('=B'+hRow+'+C'+hRow);
    sh.getRange(row, 1, 1, 4).setBackground(i%2===0?WHITE:GREY).setFontSize(10);
    row++;
  });
  row++;

  // ── §C — Prévision Dividendes ─────────────────────────────────────
  sh.getRange(row, 1).setValue('##C_DIVIDENDES');
  sh.getRange(row, 1, 1, 7).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('C. Prévision Dividendes 2026-2030 (MAD)');
  row++;

  sh.getRange(row, 1, 1, 7).setValues([['Flux','2026','2027','2028','2029','2030','TOTAL']]);
  sh.getRange(row, 1, 1, 7).setBackground(GREY).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  row++;

  const divPrev = d.dividendes || [d.div*0.3||0, d.div*0.35||0, d.div*0.35||0, 0, 0];
  sh.getRange(row, 1, 1, 7).setValues([[
    'Dividendes prévus',
    divPrev[0], divPrev[1], divPrev[2], divPrev[3], divPrev[4],
    divPrev.reduce((a,v)=>a+v,0)
  ]]).setBackground(GREEN).setFontSize(10);
  row++;

  sh.getRange(row, 1, 1, 7).setValues([[
    'Retour de capital prévu',
    0, 0, d.val*0.5||0, d.val*0.5||0, 0,
    d.val||0
  ]]).setBackground(WHITE).setFontSize(10);
  row += 2;

  // ── §D — Stratégie de Sortie ──────────────────────────────────────
  sh.getRange(row, 1).setValue('##D_SORTIE');
  sh.getRange(row, 1, 1, 7).setBackground(TEAL2).setFontColor(TEAL).setFontWeight('bold');
  sh.getRange(row, 2).setValue('D. Stratégie de Sortie');
  row++;

  const sortie = d.sortie || [
    ['Type de sortie',       'Cession de parts'],
    ['Horizon cible',        d.horizon||'2028'],
    ['Prix cible (MAD)',     d.val*1.1||0],
    ['Acquéreur potentiel',  'À identifier'],
    ['Conditions de marché', 'Favorable'],
    ['Clause de préemption', 'Oui — statuts art. 12'],
    ['Valorisation méthode', 'DCF + comparables'],
    ['TRI sortie attendu',   d.tri||'—'],
  ];
  sortie.forEach((s, i) => {
    sh.getRange(row, 1, 1, 2).setValues([s])
      .setBackground(i%2===0?WHITE:GREY).setFontSize(10);
    row++;
  });

  // ── Mise en forme ─────────────────────────────────────────────────
  sh.setColumnWidths(1, 1, 220);
  sh.setColumnWidths(2, 6, 130);

  const gid = sh.getSheetId();
  Logger.log('✅ BP Participation créé : ' + shName + ' — GID = ' + gid);
  return gid;
}

// ═══════════════════════════════════════════════════════════════════════
//  9. CRÉER TOUS LES BPs (projets + participations)
//  Exécuter : createAllBPs()
// ═══════════════════════════════════════════════════════════════════════
function createAllBPs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Données 11 Projets ────────────────────────────────────────────
  const PROJETS = [
    {
      nom: 'SARAYA VILLAS', avancement:45, budget:750000000, ca_prev:975000000,
      fp:225000000, moic:'1.50x', livraison:'T1 2029', tri:'22%', van:'95 M MAD', payback:'T4 2028', marge:'30%',
      cf:{ co:[-100000000,-300000000,-350000000,0,0], ci:[25000000,150000000,500000000,300000000,0] },
      invest:[['Acquisition foncière',225000000,'30%'],['Études & conception',37500000,'5%'],['Travaux construction',337500000,'45%'],['Commercialisation',60000000,'8%'],['Frais financiers',52500000,'7%'],['Imprévus',37500000,'5%']],
      financement:[['Fonds propres',225000000,'30%','Apport actionnaire'],['Crédit bancaire',400000000,'53%','Taux 5% / 7 ans'],['Préventes',125000000,'17%','Compromis signés']],
      rentabilite:[['TRI','22%','≥ 15%','BON'],['VAN (10%)','95 M MAD','> 0','BON'],['MOIC','1.50x','≥ 1.30x','BON'],['Payback','T4 2028','< 4 ans','BON'],['Marge nette','30%','≥ 20%','BON']]
    },
    {
      nom: 'AZ SIGNATURE', avancement:5, budget:200000000, ca_prev:290000000,
      fp:60000000, moic:'1.45x', livraison:'Fin 2029', tri:'18%', van:'À compléter', payback:'T2 2030', marge:'25%',
      cf:{ co:[-70000000,-80000000,-50000000,0,0], ci:[0,0,100000000,160000000,30000000] }
    },
    {
      nom: 'JARDINS DE MESNANA', avancement:5, budget:150000000, ca_prev:213000000,
      fp:45000000, moic:'1.42x', livraison:'Fin 2029', tri:'17%', van:'À compléter', payback:'T3 2030', marge:'22%',
      cf:{ co:[-50000000,-60000000,-40000000,0,0], ci:[0,0,80000000,100000000,33000000] }
    },
    {
      nom: 'PROJET D', avancement:0, budget:120000000, ca_prev:156000000,
      fp:36000000, moic:'1.30x', livraison:'Fin 2029', tri:'16%', van:'À compléter', payback:'Fin 2029', marge:'20%',
      cf:{ co:[0,-40000000,-50000000,-30000000,0], ci:[0,0,0,80000000,76000000] }
    },
    {
      nom: 'PROJET E', avancement:0, budget:80000000, ca_prev:112000000,
      fp:24000000, moic:'1.40x', livraison:'Fin 2028', tri:'20%', van:'À compléter', payback:'Fin 2028', marge:'28%',
      cf:{ co:[0,-30000000,-35000000,-15000000,0], ci:[0,0,50000000,50000000,12000000] }
    },
    {
      nom: 'PROJET F', avancement:2, budget:200000000, ca_prev:290000000,
      fp:60000000, moic:'1.45x', livraison:'Fin 2029', tri:'19%', van:'À compléter', payback:'Fin 2029', marge:'25%',
      cf:{ co:[0,-60000000,-80000000,-60000000,0], ci:[0,0,0,150000000,140000000] }
    },
    {
      nom: 'PROJET G', avancement:0, budget:45000000, ca_prev:63000000,
      fp:13500000, moic:'1.40x', livraison:'Fin 2030', tri:'14%', van:'À compléter', payback:'Fin 2030', marge:'18%',
      cf:{ co:[0,-15000000,-18000000,-12000000,0], ci:[0,0,0,30000000,33000000] }
    },
    {
      nom: 'PROJET H', avancement:0, budget:30000000, ca_prev:40800000,
      fp:9000000, moic:'1.36x', livraison:'Fin 2030', tri:'13%', van:'À compléter', payback:'Fin 2030', marge:'16%',
      cf:{ co:[0,-10000000,-12000000,-8000000,0], ci:[0,0,0,20000000,20800000] }
    },
    {
      nom: 'PROJET I', avancement:0, budget:25000000, ca_prev:34750000,
      fp:7500000, moic:'1.39x', livraison:'Fin 2030', tri:'15%', van:'À compléter', payback:'Fin 2030', marge:'19%',
      cf:{ co:[0,-8000000,-10000000,-7000000,0], ci:[0,0,0,17000000,17750000] }
    },
    {
      nom: 'PROJET J', avancement:0, budget:70000000, ca_prev:98700000,
      fp:21000000, moic:'1.41x', livraison:'Fin 2030', tri:'17%', van:'À compléter', payback:'Fin 2030', marge:'22%',
      cf:{ co:[0,-25000000,-28000000,-17000000,0], ci:[0,0,0,50000000,48700000] }
    },
    {
      nom: 'PROJET K', avancement:0, budget:50000000, ca_prev:72000000,
      fp:15000000, moic:'1.44x', livraison:'À définir', tri:'À compléter', van:'À compléter', payback:'À définir', marge:'À compléter',
      cf:{ co:[0,0,-20000000,-20000000,-10000000], ci:[0,0,0,35000000,37000000] }
    }
  ];

  // ── Données 5 Participations ──────────────────────────────────────
  const PARTICIPATIONS = [
    {
      nom:'ABC HOLDING SA', type:'Participation opérationnelle', secteur:'Immobilier',
      pct:'25%', cout:15000000, val:22000000, pv:7000000, div:1200000,
      moic:'1.55x', tri:'18%', horizon:'2028', statut:'Actif',
      dividendes:[300000,400000,500000,0,0],
      sortie:[['Type de sortie','Cession de parts'],['Horizon cible','2028'],['Prix cible','25 M MAD'],['Acquéreur potentiel','Partenaire stratégique identifié'],['Valorisation méthode','DCF + ANR'],['TRI sortie attendu','18%']]
    },
    {
      nom:'XYZ RENEWABLE FUND', type:'Fonds PE / OPCI', secteur:'Énergie',
      pct:'10%', cout:8000000, val:10500000, pv:2500000, div:500000,
      moic:'1.38x', tri:'15%', horizon:'2027', statut:'Actif',
      dividendes:[100000,200000,200000,0,0],
      sortie:[['Type de sortie','Rachat par le fonds'],['Horizon cible','2027'],['Prix cible','12 M MAD'],['Valorisation méthode','NAV fonds'],['TRI sortie attendu','15%']]
    },
    {
      nom:'IMMO CLUB DEAL MARRAKECH', type:'Club Deal immobilier', secteur:'Immobilier',
      pct:'33%', cout:5000000, val:6200000, pv:1200000, div:200000,
      moic:'1.28x', tri:'14%', horizon:'2026', statut:'En cours de cession',
      dividendes:[200000,0,0,0,0],
      sortie:[['Type de sortie','Cession club deal'],['Horizon cible','2026'],['Prix cible','7 M MAD'],['Acquéreur potentiel','Promoteur local'],['TRI sortie attendu','14%']]
    },
    {
      nom:'DEF INDUSTRIE SARL', type:'Participation opérationnelle', secteur:'Industrie',
      pct:'40%', cout:12000000, val:14000000, pv:2000000, div:800000,
      moic:'1.23x', tri:'12%', horizon:'2029', statut:'Actif',
      dividendes:[200000,250000,300000,50000,0],
      sortie:[['Type de sortie','Cession à partenaire'],['Horizon cible','2029'],['Prix cible','18 M MAD'],['Valorisation méthode','EBE × 5'],['TRI sortie attendu','12%']]
    },
    {
      nom:'OBLIGATIONS PRIVEES GHI', type:'Dette privée / Obligations', secteur:'Services',
      pct:'100%', cout:6000000, val:6300000, pv:300000, div:420000,
      moic:'1.12x', tri:'7%', horizon:'2027', statut:'Actif',
      dividendes:[420000,420000,420000,0,0],
      sortie:[['Type de sortie','Remboursement obligataire'],['Horizon cible','2027'],['Prix cible','6.3 M MAD'],['Conditions','Remboursement à l\'échéance'],['TRI sortie attendu','7%']]
    }
  ];

  // ── Créer tous les BPs Projets ────────────────────────────────────
  const gidsProj = {};
  PROJETS.forEach(p => {
    try {
      gidsProj[p.nom] = createProjectBPSheet(ss, p.nom, p);
      Utilities.sleep(500); // éviter les limites API
    } catch(e) {
      Logger.log('Erreur ' + p.nom + ': ' + e.message);
      gidsProj[p.nom] = 'ERREUR';
    }
  });

  // ── Créer tous les BPs Participations ─────────────────────────────
  const gidsPart = {};
  PARTICIPATIONS.forEach(p => {
    try {
      gidsPart[p.nom] = createParticipationBPSheet(ss, p.nom, p);
      Utilities.sleep(500);
    } catch(e) {
      Logger.log('Erreur ' + p.nom + ': ' + e.message);
      gidsPart[p.nom] = 'ERREUR';
    }
  });

  // ── Afficher le résumé des GIDs ───────────────────────────────────
  let msg = '✅ ' + PROJETS.length + ' BPs Projets + ' + PARTICIPATIONS.length + ' BPs Participations créés !\n\n';
  msg += '══ PROJETS → col K (cf_gid) de l\'onglet PROJETS ══\n';
  Object.entries(gidsProj).forEach(([nom, gid]) => { msg += '• ' + nom.substring(0,22).padEnd(23) + ' → ' + gid + '\n'; });
  msg += '\n══ PARTICIPATIONS → col M (bp_gid) de PARTICIPATIONS ══\n';
  Object.entries(gidsPart).forEach(([nom, gid]) => { msg += '• ' + nom.substring(0,22).padEnd(23) + ' → ' + gid + '\n'; });
  msg += '\n⚠️ Publie tous les onglets BP_* :\nFichier > Partager > Publier sur le web > CSV';

  SpreadsheetApp.getUi().alert(msg);

  // ── Log complet dans la console ───────────────────────────────────
  Logger.log('=== GIDS PROJETS ===');
  Object.entries(gidsProj).forEach(([n,g]) => Logger.log(n + ' → ' + g));
  Logger.log('=== GIDS PARTICIPATIONS ===');
  Object.entries(gidsPart).forEach(([n,g]) => Logger.log(n + ' → ' + g));
}
