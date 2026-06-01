/**
 * T.A INVEST HOLDING — Dashboard API
 * Google Apps Script Web App
 * Déployer : Déployer > Nouveau déploiement > Application Web
 * Accès : Tout le monde (anonyme)
 */

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'all';

  try {
    let data = {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'all' || action === 'projets') {
      data.projets = getProjets(ss);
    }
    if (action === 'all' || action === 'engagements') {
      data.engagements = getEngagements(ss);
    }
    if (action === 'all' || action === 'participations') {
      data.participations = getParticipations(ss);
    }
    if (action === 'all' || action === 'roadmap') {
      data.roadmap = getRoadmap(ss);
    }
    if (action === 'all' || action === 'commercial') {
      data.commercial = getCommercial(ss);
    }

    data.timestamp = new Date().toISOString();
    data.status = 'ok';

    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── PROJETS ──────────────────────────────────────────
function getProjets(ss) {
  const sheet = ss.getSheetByName('Projets');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || String(r[0]).toUpperCase().includes('TOTAL') || r[0] === '') continue;
    const budget = parseFloat(String(r[2]).replace(/[\s,]/g, '')) || 0;
    if (budget === 0) continue;
    out.push({
      name:    String(r[0]).trim(),
      sector:  String(r[1] || '—').trim(),
      budget:  budget / 1e6,
      engage:  parseNum(r[3]) / 1e6,
      reste:   parseNum(r[4]) / 1e6,
      besoin:  parseNum(r[5]) / 1e6,
      tri:     parsePct(r[6]),
      payback: String(r[7] || '—'),
      marge:   parsePct(r[8]),
      statut:  String(r[9] || 'À définir')
    });
  }
  return out;
}

// ── ENGAGEMENTS ──────────────────────────────────────
function getEngagements(ss) {
  const sheet = ss.getSheetByName('Engagements');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || r[0] === '') continue;
    const rawM = parseFloat(String(r[2]).replace(/[\s,]/g, '')) || 0;
    out.push({
      projet:  String(r[0]).trim(),
      nature:  String(r[1] || '—'),
      montant: rawM > 1000 ? rawM / 1e6 : rawM,
      date:    String(r[3] || '—'),
      statut:  String(r[4] || '—'),
      decision:String(r[5] || '—')
    });
  }
  return out;
}

// ── PARTICIPATIONS ───────────────────────────────────
function getParticipations(ss) {
  const sheet = ss.getSheetByName('Participations');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  // Chercher la ligne de début (après les en-têtes)
  let start = 1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] && !String(rows[i][0]).toUpperCase().includes('NOM') && !String(rows[i][0]).toUpperCase().includes('SOCIÉTÉ')) {
      start = i; break;
    }
  }
  for (let i = start; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || String(r[0]).toUpperCase().includes('TOTAL') || r[0] === '') continue;
    const cout = parseNum(r[4]);
    if (cout === 0) continue;
    const coutM = cout / 1e6;
    const valM  = parseNum(r[5]) / 1e6;
    const pvM   = parseNum(r[6]) / 1e6;
    const moicRaw = parseFloat(String(r[8]).replace(/[\s,]/g, '').replace(',', '.')) || 0;
    out.push({
      nom:     String(r[0]).trim(),
      type:    String(r[1] || '—'),
      sect:    String(r[2] || '—'),
      pct:     parsePct(r[3]),
      cout:    coutM,
      val:     valM,
      pv:      pvM !== 0 ? pvM : (valM - coutM),
      div:     parseNum(r[7]) / 1e6,
      moic:    moicRaw > 0 ? moicRaw : (coutM > 0 ? valM / coutM : 0),
      tri:     parsePct(r[9]),
      horizon: String(r[10] || '—'),
      statut:  String(r[11] || '—')
    });
  }
  return out;
}

// ── ROADMAP ──────────────────────────────────────────
function getRoadmap(ss) {
  const sheet = ss.getSheetByName('Roadmap');
  if (!sheet) return { bars: [], etapes: [] };
  const rows = sheet.getDataRange().getValues();
  const bars = [], etapes = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    if (String(r[0]).toUpperCase() === 'BAR') {
      bars.push({ axe: String(r[1]), label: String(r[2]), start: parseInt(r[3]) || 1, end: parseInt(r[4]) || 1, color: '#' + String(r[5]).replace('#', '') });
    }
    if (String(r[0]).toUpperCase() === 'ETAPE') {
      etapes.push({ date: String(r[1]), desc: String(r[2]) });
    }
  }
  return { bars, etapes };
}

// ── COMMERCIAL ───────────────────────────────────────
function getCommercial(ss) {
  const sheet = ss.getSheetByName('Commercial');
  if (!sheet) return { unites: [], prospects: [], echeancier: [] };
  const rows = sheet.getDataRange().getValues();
  const unites = [], prospects = [], echeancier = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const type = String(r[0]).trim().toUpperCase();
    if (!type) continue;

    if (type === 'U') {
      unites.push({
        projet:  String(r[1] || '—').trim(),
        ref:     String(r[2] || '—').trim(),
        type:    String(r[3] || '—').trim(),
        surf:    parseFloat(r[4]) || 0,
        prix:    parseFloat(String(r[5]).replace(/[\s,]/g, '').replace(',', '.')) || 0,
        client:  String(r[6] || '—').trim(),
        date:    String(r[7] || '—').trim(),
        statut:  String(r[8] || 'Disponible').trim()
      });
    }
    if (type === 'P') {
      prospects.push({
        nom:    String(r[1] || '—').trim(),
        projet: String(r[2] || '—').trim(),
        budget: parseFloat(String(r[3]).replace(/[\s,]/g, '').replace(',', '.')) || 0,
        statut: String(r[4] || '—').trim(),
        action: String(r[5] || '—').trim()
      });
    }
    if (type === 'E') {
      echeancier.push({
        ref:      String(r[1] || '—').trim(),
        client:   String(r[2] || '—').trim(),
        montant:  parseFloat(String(r[3]).replace(/[\s,]/g, '').replace(',', '.')) || 0,
        echeance: String(r[4] || '—').trim(),
        statut:   String(r[5] || 'À venir').trim()
      });
    }
  }
  return { unites, prospects, echeancier };
}

// ── HELPERS ──────────────────────────────────────────
function parseNum(v) {
  return parseFloat(String(v).replace(/[\s]/g, '').replace(',', '.')) || 0;
}
function parsePct(v) {
  const s = String(v).replace(/[%\s]/g, '').replace(',', '.');
  const n = parseFloat(s) || 0;
  return n > 1 ? n / 100 : n;
}
