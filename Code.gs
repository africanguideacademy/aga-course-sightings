// ════════════════════════════════════════════════════════════════
// AFRICAN GUIDE ACADEMY — GAME DRIVE SIGHTINGS
// Google Apps Script backend (bound to the Sightings Google Sheet)
// ════════════════════════════════════════════════════════════════
//
// This is the database side of the AGA Sightings app. The app itself
// lives on GitHub Pages and talks to this script.
//
// SETUP (once):
// 1. In the "AGA Sightings" Google Sheet: Extensions > Apps Script.
// 2. Replace everything in Code.gs with this file. (You can delete Index.html.)
// 3. Project Settings (cog) > Time zone: (GMT+02:00) Central Africa Time - Gaborone.
// 4. Choose the function "setup" and press Run. Allow access.
// 5. Deploy > New deployment > Web app.   Execute as: Me   Who has access: Anyone
//    Copy the web app URL (ends in /exec) into config.js in the app.
//
// After editing this code later: Deploy > Manage deployments > pencil
// > Version: New version > Deploy. The URL stays the same.
// ════════════════════════════════════════════════════════════════

const SH_SIGHTINGS = 'Sightings';
const SH_SPECIES   = 'Species';
const SH_COURSES   = 'Courses';

const SIGHT_HEADERS   = ['Sighting ID', 'Logged At', 'Date', 'Course Code', 'Group', 'Activity', 'Session', 'Category', 'Species', 'Count', 'Notes', 'Observer'];
const SPECIES_HEADERS = ['Category', 'Species', 'Points'];
const COURSE_HEADERS  = ['Course Code', 'Start Date', 'End Date', 'Groups (comma separated)', 'Active (Y/N)'];

// Must match APP_KEY in the app's config.js. Stops strangers writing to your sheet.
const APP_KEY = 'aga-sightings-2026';

const CATEGORIES = ['Mammals', 'Birds', 'Reptiles', 'Amphibians', 'Insects', 'Other invertebrates'];

// ── API FOR THE APP ───────────────────────────────────────────────
function doGet(e) {
  // Open the /exec URL in a browser to check it's working.
  return json_({ ok: true, app: 'AGA Sightings', time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'Bad request' }); }
  if (req.key !== APP_KEY) return json_({ error: 'Wrong app key. Check config.js matches APP_KEY in Code.gs.' });
  try {
    switch (req.action) {
      case 'config':  return json_(getConfig());
      case 'results': return json_(getResults(req.course));
      case 'save':    return json_(saveSightings(req.items));
      case 'delete':  return json_(deleteSighting(req.id));
      default:        return json_({ error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('AGA Sightings')
    .addItem('Set up / repair sheets', 'setup')
    .addToUi();
}

// ── SETUP ─────────────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const s = ensureSheet_(ss, SH_SIGHTINGS, SIGHT_HEADERS);
  s.getRange('C:C').setNumberFormat('dd mmm yyyy');
  s.getRange('D:D').setNumberFormat('@');
  s.getRange('B:B').setNumberFormat('dd mmm yyyy hh:mm');

  const sp = ensureSheet_(ss, SH_SPECIES, SPECIES_HEADERS);
  if (sp.getLastRow() < 2) {
    const rows = seedSpecies_();
    sp.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  const c = ensureSheet_(ss, SH_COURSES, COURSE_HEADERS);
  c.getRange('A:A').setNumberFormat('@');
  if (c.getLastRow() < 2) {
    c.getRange(2, 1, 1, 5).setValues([[
      '2606 KWP NG', new Date(), '', 'Buffalo, Kudu, Hornbill, Mopane', 'Y'
    ]]);
  }

  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#5C3D1E').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}

// ── CONFIG FOR THE APP ────────────────────────────────────────────
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cRows = ss.getSheetByName(SH_COURSES).getDataRange().getValues().slice(1);
  const courses = cRows
    .filter(r => String(r[0]).trim() && String(r[4]).trim().toUpperCase() !== 'N')
    .map(r => ({
      code: String(r[0]).trim(),
      groups: String(r[3]).split(',').map(g => g.trim()).filter(Boolean)
    }));

  const spRows = ss.getSheetByName(SH_SPECIES).getDataRange().getValues().slice(1);
  const species = spRows
    .filter(r => String(r[1]).trim())
    .map(r => ({ category: String(r[0]).trim(), name: String(r[1]).trim(), points: Number(r[2]) || 1 }));

  return { courses: courses, species: species, categories: CATEGORIES };
}

// ── SAVE A BATCH OF SIGHTINGS ─────────────────────────────────────
// Each item: {id, date:'yyyy-mm-dd', course, group, activity, session, category, species, count, notes, observer}
function saveSightings(items) {
  if (!items || !items.length) return { saved: 0, ids: [] };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SH_SIGHTINGS);
    if (sh.getRange(1, SIGHT_HEADERS.length).getValue() !== 'Observer') {
      sh.getRange(1, 1, 1, SIGHT_HEADERS.length).setValues([SIGHT_HEADERS])
        .setFontWeight('bold').setBackground('#5C3D1E').setFontColor('#FFFFFF');
    }

    // Skip anything already saved (protects against double-sends from weak signal)
    const last = sh.getLastRow();
    const existing = new Set(last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : []);
    const fresh = items.filter(it => it.id && !existing.has(String(it.id)));

    if (fresh.length) {
      const now = new Date();
      const rows = fresh.map(it => [
        String(it.id), now, parseDate_(it.date), String(it.course || ''), String(it.group || ''),
        String(it.activity || ''), String(it.session || ''), String(it.category || ''),
        String(it.species || '').trim(), Number(it.count) || 1, String(it.notes || ''),
        String(it.observer || '').trim()
      ]);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, SIGHT_HEADERS.length).setValues(rows);
      addNewSpecies_(ss, fresh);
    }
    return { saved: fresh.length, ids: items.map(it => String(it.id)) };
  } finally {
    lock.releaseLock();
  }
}

// Species typed in by guides that aren't on the master list get added (1 point)
function addNewSpecies_(ss, items) {
  const sp = ss.getSheetByName(SH_SPECIES);
  const known = new Set(sp.getDataRange().getValues().slice(1)
    .map(r => (String(r[0]).trim() + '|' + String(r[1]).trim()).toLowerCase()));
  const add = [];
  items.forEach(it => {
    const k = (String(it.category).trim() + '|' + String(it.species).trim()).toLowerCase();
    if (it.species && !known.has(k)) { known.add(k); add.push([it.category, String(it.species).trim(), 1]); }
  });
  if (add.length) sp.getRange(sp.getLastRow() + 1, 1, add.length, 3).setValues(add);
}

function parseDate_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date();
}

// ── RESULTS FOR ONE COURSE ────────────────────────────────────────
function getResults(course) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const code = String(course || '').trim();

  const data = ss.getSheetByName(SH_SIGHTINGS).getDataRange().getValues().slice(1);
  const rows = data
    .filter(r => String(r[3]).trim() === code && String(r[8]).trim())
    .map(r => ({
      id: String(r[0]),
      loggedAt: r[1] instanceof Date ? r[1].getTime() : 0,
      date: r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'yyyy-MM-dd') : String(r[2]),
      group: String(r[4]).trim(),
      activity: String(r[5]),
      session: String(r[6]),
      category: String(r[7]).trim(),
      species: String(r[8]).trim(),
      count: Number(r[9]) || 1,
      notes: String(r[10] || ''),
      observer: String(r[11] || '').trim()
    }));

  const points = {};
  ss.getSheetByName(SH_SPECIES).getDataRange().getValues().slice(1).forEach(r => {
    points[(String(r[0]).trim() + '|' + String(r[1]).trim()).toLowerCase()] = Number(r[2]) || 1;
  });

  return { course: code, rows: rows, points: points };
}

// ── REMOVE A MISTAKE ──────────────────────────────────────────────
function deleteSighting(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SIGHTINGS);
    const last = sh.getLastRow();
    if (last < 2) return { deleted: false };
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); return { deleted: true }; }
    }
    return { deleted: false };
  } finally {
    lock.releaseLock();
  }
}

// ── STARTER SPECIES LIST (northern Botswana) ──────────────────────
// Points reward harder sightings. Edit freely in the Species sheet.
function seedSpecies_() {
  const list = {
    'Mammals': [
      ['African Elephant',1],['Cape Buffalo',1],['Giraffe',1],['Plains Zebra',1],['Blue Wildebeest',1],
      ['Impala',1],['Greater Kudu',1],['Red Lechwe',1],['Tsessebe',1],['Waterbuck',1],['Common Warthog',1],
      ['Hippopotamus',1],['Chacma Baboon',1],['Vervet Monkey',1],['Steenbok',1],['Common Duiker',1],
      ['Bushbuck',1],['Tree Squirrel',1],['Scrub Hare',1],['Slender Mongoose',1],['Banded Mongoose',1],
      ['Dwarf Mongoose',1],['Black-backed Jackal',1],['Eland',2],['Southern Reedbuck',2],['Gemsbok',2],
      ['Springbok',2],['Red Hartebeest',2],['Spotted Hyaena',2],['Side-striped Jackal',2],['Bat-eared Fox',2],
      ['Large-spotted Genet',2],['Springhare',2],['Lesser Galago',2],['Yellow Mongoose',2],['Meerkat',2],
      ['Lion',3],['Sable Antelope',3],['Puku',3],['African Civet',3],['Cape Porcupine',3],['African Wild Cat',3],
      ['Water Mongoose',3],['Bushpig',3],['Klipspringer',3],['Leopard',4],['Cheetah',4],['Roan Antelope',4],
      ['Sitatunga',4],['Honey Badger',4],['Caracal',4],['Serval',4],['Brown Hyaena',4],['White Rhinoceros',4],
      ['Cape Clawless Otter',4],['African Wild Dog',5],['Aardvark',5],['Ground Pangolin',5],['Black Rhinoceros',5],['Aardwolf',5]
    ],
    'Birds': [
      ['Lilac-breasted Roller',1],['Southern Yellow-billed Hornbill',1],['Red-billed Hornbill',1],['African Grey Hornbill',1],
      ['Grey Go-away-bird',1],['Cape Turtle Dove',1],['Laughing Dove',1],['Emerald-spotted Wood Dove',1],['Helmeted Guineafowl',1],
      ['Crested Francolin',1],['Red-billed Spurfowl',1],['Swainson\'s Spurfowl',1],['African Fish Eagle',1],['Bateleur',1],
      ['Tawny Eagle',1],['White-backed Vulture',1],['Marabou Stork',1],['Yellow-billed Stork',1],['African Openbill',1],
      ['Hamerkop',1],['Grey Heron',1],['Great Egret',1],['Cattle Egret',1],['Sacred Ibis',1],['Hadada Ibis',1],
      ['Egyptian Goose',1],['Spur-winged Goose',1],['White-faced Whistling Duck',1],['African Jacana',1],['Blacksmith Lapwing',1],
      ['Crowned Lapwing',1],['Water Thick-knee',1],['Pied Kingfisher',1],['Malachite Kingfisher',1],['Woodland Kingfisher',1],
      ['Little Bee-eater',1],['Fork-tailed Drongo',1],['Magpie Shrike',1],['Burchell\'s Starling',1],['Meves\'s Starling',1],
      ['Cape Glossy Starling',1],['Arrow-marked Babbler',1],['Red-billed Oxpecker',1],['Yellow-billed Oxpecker',1],
      ['Blue Waxbill',1],['Red-billed Firefinch',1],['African Hoopoe',1],['Green Wood-hoopoe',1],['Meyer\'s Parrot',1],
      ['Black-collared Barbet',1],['Coppery-tailed Coucal',1],['Crimson-breasted Shrike',1],['Pearl-spotted Owlet',1],
      ['Red-crested Korhaan',1],['Kori Bustard',2],['Ostrich',2],['Saddle-billed Stork',2],['Goliath Heron',2],['Black Heron',2],
      ['Giant Kingfisher',2],['Southern Carmine Bee-eater',2],['Swallow-tailed Bee-eater',2],['Lappet-faced Vulture',2],
      ['Hooded Vulture',2],['Brown Snake Eagle',2],['African Scops Owl',2],['Southern White-faced Owl',2],
      ['Southern Ground Hornbill',3],['Secretarybird',3],['Wattled Crane',3],['Martial Eagle',3],['Slaty Egret',3],
      ['Verreaux\'s Eagle-Owl',3],['African Skimmer',3],['Pel\'s Fishing Owl',5]
    ],
    'Reptiles': [
      ['Nile Crocodile',1],['Water Monitor',1],['Tree Agama',1],['Striped Skink',1],['Variable Skink',1],
      ['Tropical House Gecko',1],['Serrated Hinged Terrapin',1],['Marsh Terrapin',1],['Leopard Tortoise',2],
      ['Rock Monitor',2],['Ground Agama',2],['Flap-necked Chameleon',2],['Puff Adder',2],['Spotted Bush Snake',2],
      ['Black Mamba',3],['Mozambique Spitting Cobra',3],['Snouted Cobra',3],['Boomslang',3],['Southern African Python',4]
    ],
    'Amphibians': [
      ['Guttural Toad',1],['Painted Reed Frog',1],['Angolan Reed Frog',1],['Common River Frog',1],['Red Toad',1],
      ['Bubbling Kassina',2],['Southern Foam Nest Frog',2],['African Bullfrog',3]
    ],
    'Insects': [
      ['Dung Beetle',1],['Termites (mound building)',1],['Matabele Ant',1],['Honey Bee',1],['Carpenter Bee',1],
      ['Dragonfly',1],['African Monarch',1],['Grasshopper',1],['Cicada',1],['Antlion',1],['Armoured Ground Cricket',2],
      ['Praying Mantis',2],['Blister Beetle',2],['Tok-tokkie Beetle',2],['Mopane Worm',2],['Stick Insect',3]
    ],
    'Other invertebrates': [
      ['Golden Orb-web Spider',1],['Giant Millipede',1],['Freshwater Crab',1],['Solifuge',2],
      ['Thick-tailed Scorpion',2],['Baboon Spider',3]
    ]
  };
  const rows = [];
  Object.keys(list).forEach(cat => list[cat].forEach(s => rows.push([cat, s[0], s[1]])));
  return rows;
}
