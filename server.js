require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const multer = require('multer');
const ExcelJS = require('exceljs');

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET,
};
const ADMIN_PASSCODE = process.env.ADMIN_PASSCODE || 'changeme';
const UPLOAD_PASSCODE = process.env.UPLOAD_PASSCODE || 'changeme123';
// ใช้สำหรับฟีเจอร์ "แท็กชื่อบอทแล้วถามคำถามทั่วไป" - ใช้ Gemini API ของ Google ซึ่งมี free tier
// (ไม่ต้องผูกบัตรเครดิต) ขอ API key ฟรีได้ที่ https://aistudio.google.com/apikey แล้วเอามาตั้งเป็น
// ตัวแปรนี้ใน Railway ถ้าไม่ตั้งค่า ฟีเจอร์นี้จะปิดอยู่เฉยๆ (ฟีเจอร์อื่นในบอทยังใช้ได้ปกติ)
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
const GEMINI_MODEL = process.env.GEMINI_MODEL || 'gemini-2.5-flash';

// โดเมนสาธารณะของบอทเอง (URL ที่คนทั่วไปเข้าเว็บนี้ได้จริง เช่น https://your-app.up.railway.app)
// ใช้สร้างลิงก์รูปสัตว์เลี้ยงที่ host เอง (โฟลเดอร์ public/pet-images) เพื่อส่งเป็น LINE image message
// หา URL จริงได้จากหน้า Settings > Networking ของ service ใน Railway (หรือ URL ของ service บน Render)
// ถ้าไม่ตั้งค่า ฟีเจอร์ส่งรูปสัตว์เลี้ยงจะปิดอยู่เฉยๆ (บอทยังใช้งานได้ปกติทุกอย่าง แค่ไม่มีรูปแนบ)
const PUBLIC_BASE_URL = (process.env.PUBLIC_BASE_URL || '').trim().replace(/\/+$/, '');

// บอทตัวเดียวกันนี้ใช้ได้หลายกลุ่ม แยกบทบาทกันด้วย groupId
// ถ้าไม่ตั้งค่า (เว้นว่างไว้) = เปิดใช้งานฟีเจอร์นั้นได้ทุกกลุ่ม (พฤติกรรมเดิม)
// ถ้าตั้งค่าแล้ว = ฟีเจอร์นั้นจะทำงานเฉพาะกลุ่มที่ระบุ groupId ตรงกันเท่านั้น
// หา groupId ได้โดยพิมพ์ /groupid ในกลุ่มนั้นๆ
function isQuizGroup(groupId) {
  const configured = process.env.QUIZ_GROUP_ID;
  return !configured || configured === groupId;
}
function isStatusGroup(groupId) {
  const configured = process.env.STATUS_GROUP_ID;
  return !configured || configured === groupId;
}

if (!config.channelAccessToken || !config.channelSecret) {
  console.warn('!!! CHANNEL_ACCESS_TOKEN / CHANNEL_SECRET ยังไม่ได้ตั้งค่า (ดูไฟล์ .env.example)');
}

const client = new line.Client(config);
const app = express();

// log ทุก request ที่เข้ามา ไว้ช่วย debug (ดูได้จากแท็บ Logs)
app.use((req, res, next) => {
  console.log(`[incoming] ${req.method} ${req.originalUrl}`);
  next();
});

const DB_PATH = path.join(__dirname, 'data', 'scores.json');

function loadDB() {
  try {
    return JSON.parse(fs.readFileSync(DB_PATH, 'utf8'));
  } catch (e) {
    return {
      groupId: null,
      active: false,
      currentQuestion: 0,
      admins: [],
      players: {},        // userId -> { name, score, answeredCount, correctCount }
      pendingAnswers: {},  // userId -> "A" | "B" | ...  (for the question in progress)
      log: [],             // history of scored/skipped questions
    };
  }
}

function saveDB(db) {
  fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2));
  gistDirty = true; // มีการเปลี่ยนแปลงข้อมูล รอบ backup ถัดไปจะส่งขึ้น Gist ให้
}

// ==== สำรองข้อมูลอัตโนมัติไป GitHub Gist (กันคะแนนหายตอน deploy ใหม่/รีสตาร์ท) ==================
// โฮสต์ฟรีอย่าง Render มีดิสก์แบบไม่ถาวร (ephemeral) พอ deploy โค้ดใหม่หรือรีสตาร์ทเซิร์ฟเวอร์ ไฟล์
// data/scores.json ในเครื่องจะถูกล้าง ฟีเจอร์นี้จะคอยสำรอง db ทั้งก้อน (คะแนนเกม, สัตว์เลี้ยง, ทุกอย่าง)
// ขึ้น GitHub Gist ส่วนตัวเป็นระยะๆ แล้วดึงกลับมาคืนอัตโนมัติตอนบอทเริ่มทำงานใหม่ ถ้าเครื่องข้อมูลว่างเปล่า
// (หรือของใน Gist ใหม่กว่า) — ต้องตั้งค่า GITHUB_GIST_TOKEN (Personal Access Token สโคป "gist") เป็น
// Environment Variable ก่อนถึงจะเปิดใช้งาน ถ้าไม่ตั้งค่า ฟีเจอร์นี้จะปิดอยู่เฉยๆ บอทยังทำงานได้ปกติทุกอย่าง
// (ดูวิธีสร้าง Token ในไฟล์ README.md)
const GITHUB_GIST_TOKEN = process.env.GITHUB_GIST_TOKEN || '';
const GIST_DESCRIPTION = 'line-quiz-bot database backup (auto-managed by the bot — do not delete)';
const GIST_FILENAME = 'line-quiz-bot-db-backup.json';
const GIST_BACKUP_INTERVAL_MS = 30 * 1000; // ส่งขึ้น Gist อย่างมากทุก 30 วินาทีเมื่อมีข้อมูลเปลี่ยนแปลง
let gistId = null;
let gistDirty = false;
let gistSyncing = false;

async function githubGistRequest(pathSuffix, options = {}) {
  const res = await fetch(`https://api.github.com${pathSuffix}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${GITHUB_GIST_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'User-Agent': 'line-quiz-bot',
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`GitHub API ${res.status}: ${body.slice(0, 300)}`);
  }
  return res.json();
}

// หา Gist สำรองข้อมูลที่เคยสร้างไว้ (เทียบจาก description) ถ้ายังไม่เคยมีก็สร้างใหม่ให้เลย
// ทำแบบนี้แทนการจำ Gist ID ไว้ใน Environment Variable เพื่อไม่ต้องให้ผู้ใช้ไปคัดลอก ID มาตั้งเองอีกขั้นตอน
async function findOrCreateBackupGist() {
  const gists = await githubGistRequest('/gists?per_page=100');
  const existing = gists.find((g) => g.description === GIST_DESCRIPTION);
  if (existing) return existing.id;
  const created = await githubGistRequest('/gists', {
    method: 'POST',
    body: JSON.stringify({
      description: GIST_DESCRIPTION,
      public: false,
      files: { [GIST_FILENAME]: { content: JSON.stringify({ _bootstrap: true }) } },
    }),
  });
  return created.id;
}

async function fetchBackupFromGist(id) {
  const gist = await githubGistRequest(`/gists/${id}`);
  const file = gist.files && gist.files[GIST_FILENAME];
  if (!file || !file.content) return null;
  try {
    return JSON.parse(file.content);
  } catch (e) {
    return null;
  }
}

async function pushBackupToGist(id, dbSnapshot) {
  await githubGistRequest(`/gists/${id}`, {
    method: 'PATCH',
    body: JSON.stringify({ files: { [GIST_FILENAME]: { content: JSON.stringify(dbSnapshot, null, 2) } } }),
  });
}

let db = loadDB();

// ตอนเริ่มบอท: ถ้าตั้งค่า Token ไว้ ให้เช็ค/สร้าง Gist สำรอง แล้วกู้คืนข้อมูลกลับมาถ้าดิสก์ในเครื่องว่างเปล่า
// (เช่นเพิ่ง deploy ใหม่) หรือถ้าของใน Gist ใหม่กว่าที่มีในเครื่อง — server จะเริ่มรับ request หลังจากขั้นตอนนี้
// เสร็จเท่านั้น กันไม่ให้มีคนเล่นเกมแทรกเข้ามาก่อนกู้ข้อมูลเสร็จ
let dbBackupReady = Promise.resolve();
if (GITHUB_GIST_TOKEN) {
  dbBackupReady = (async () => {
    try {
      const localHadData = fs.existsSync(DB_PATH);
      gistId = await findOrCreateBackupGist();
      const remote = await fetchBackupFromGist(gistId);
      if (remote && remote._bootstrap !== true) {
        const localTs = db._backupSavedAt || 0;
        const remoteTs = remote._backupSavedAt || 0;
        if (!localHadData || remoteTs > localTs) {
          db = remote;
          delete db._bootstrap;
          saveDB(db);
          gistDirty = false; // เพิ่งซิงค์มาตรงกันแล้ว ยังไม่ต้อง push กลับทันที
          console.log('[gist-backup] restored database from GitHub Gist (local disk was empty or Gist copy was newer)');
        }
      }
      console.log(`[gist-backup] enabled, using gist ${gistId} (auto-syncs every ~${GIST_BACKUP_INTERVAL_MS / 1000}s when data changes)`);
    } catch (e) {
      console.error('[gist-backup] setup failed, continuing WITHOUT auto-backup (bot still works normally):', e.message);
      gistId = null;
    }
  })();
}

setInterval(async () => {
  if (!gistId || !gistDirty || gistSyncing) return;
  gistSyncing = true;
  gistDirty = false;
  try {
    db._backupSavedAt = Date.now();
    await pushBackupToGist(gistId, db);
  } catch (e) {
    console.error('[gist-backup] periodic push failed, will retry next round:', e.message);
    gistDirty = true;
  } finally {
    gistSyncing = false;
  }
}, GIST_BACKUP_INTERVAL_MS);

async function flushGistBackupBeforeExit(signal) {
  if (!gistId) return;
  try {
    db._backupSavedAt = Date.now();
    await pushBackupToGist(gistId, db);
    console.log(`[gist-backup] flushed latest data to Gist before ${signal}`);
  } catch (e) {
    console.error(`[gist-backup] flush before ${signal} failed:`, e.message);
  }
}
// Render (และโฮสต์อื่นๆ ส่วนใหญ่) จะส่งสัญญาณนี้มาก่อนจะปิด/สลับไปรันโค้ดเวอร์ชันใหม่ตอน deploy เสมอ
// ดักไว้ให้ส่งข้อมูลล่าสุดขึ้น Gist ก่อนปิดตัวจริง จะได้ไม่มีช่วงเวลาสูญหายแม้แต่วินาทีเดียว
['SIGTERM', 'SIGINT'].forEach((sig) => {
  process.on(sig, async () => {
    await flushGistBackupBeforeExit(sig);
    process.exit(0);
  });
});

async function getDisplayName(groupId, userId) {
  if (db.players[userId] && db.players[userId].name) return db.players[userId].name;
  try {
    const profile = await client.getGroupMemberProfile(groupId, userId);
    return profile.displayName;
  } catch (e) {
    return 'ผู้เล่น' + userId.slice(-4);
  }
}

function ensurePlayer(userId, name) {
  if (!db.players[userId]) {
    db.players[userId] = { name: name || userId, score: 0, answeredCount: 0, correctCount: 0 };
  } else if (name) {
    db.players[userId].name = name;
  }
}

function isAdmin(userId) {
  return db.admins.includes(userId);
}

function topScores(n) {
  return Object.entries(db.players)
    .map(([uid, p]) => ({ uid, ...p }))
    .sort((a, b) => b.score - a.score)
    .slice(0, n);
}

function formatBoard(list) {
  if (!list.length) return '(ยังไม่มีคะแนน)';
  return list.map((p, i) => `${i + 1}. ${p.name} - ${p.score} คะแนน`).join('\n');
}

// รองรับคำตอบเป็นตัวอักษรไทย ก/ข/ค/ง/จ/ฉ เทียบเท่า A/B/C/D/E/F ของอังกฤษ
// แปลงให้เป็นตัวอักษรอังกฤษตัวพิมพ์ใหญ่เสมอ เพื่อให้เทียบคำตอบถูก/ผิดกันได้ไม่ว่าใครจะพิมพ์แบบไหน
const THAI_TO_ENG_ANSWER = { ก: 'A', ข: 'B', ค: 'C', ง: 'D', จ: 'E', ฉ: 'F' };
function normalizeAnswerLetter(text) {
  const t = (text || '').trim();
  if (/^[a-fA-F]$/.test(t)) return t.toUpperCase();
  if (THAI_TO_ENG_ANSWER[t]) return THAI_TO_ENG_ANSWER[t];
  return null;
}

// ==== เช็คสถานะงานลูกค้า (อ่านจากไฟล์ Excel ที่แอดมินอัปโหลด) ====================
const XLSX_PATH = path.join(__dirname, 'data', 'customers.xlsx');

let DB = {
  updatedAt: null, // เวลาที่อัปโหลดไฟล์ล่าสุด
  initial: [],      // ขอใบรับรองใหม่ (sheet "Initial")
  su: [],           // ตรวจติดตาม SU1/SU2 (sheet "SU")
  recer: [],        // ต่ออายุใบรับรอง (sheet "Recer")
};

function cellStr(row, col) {
  const v = row.getCell(col).value;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return formatDate(v);
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('').trim() || null;
    if (v.result !== undefined) return v.result instanceof Date ? formatDate(v.result) : String(v.result).trim() || null;
    if (v.text) return String(v.text).trim() || null;
    return null;
  }
  const s = String(v).trim();
  return s === '' ? null : s;
}

function formatDate(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// ดึงค่าวันที่แบบ raw (ไม่ format) ไว้เทียบว่ารอบไหนใหม่กว่ากัน - คืนค่าเป็น timestamp (ms) หรือ null
function cellDate(row, col) {
  const v = row.getCell(col).value;
  if (v instanceof Date) return v.getTime();
  if (v && typeof v === 'object' && v.result instanceof Date) return v.result.getTime();
  return null;
}

// เช็คว่าข้อความสถานะถือว่า "เสร็จแล้ว" หรือยัง - นับทั้งคำว่า "done" (ไม่สนตัวพิมพ์เล็ก/ใหญ่) และ "ผ่าน"
// (ชีต SU ใช้คำว่า "ผ่าน" แทนคำว่า Done เช่น "5.ผ่าน SU1", "ผ่าน SU2(2)")
function isDoneStatus(status) {
  if (!status) return false;
  const s = status.toLowerCase();
  return s.includes('done') || s.includes('ผ่าน');
}

// โครงสร้างคอลัมน์ยึดตามไฟล์ "Customers Status" ต้นฉบับ (header แถวที่ 2, ข้อมูลเริ่มแถวที่ 3)
async function loadWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const initial = [];
  const wsInitial = workbook.getWorksheet('Initial');
  if (wsInitial) {
    for (let r = 3; r <= wsInitial.rowCount; r++) {
      const row = wsInitial.getRow(r);
      const companyName = cellStr(row, 2);
      if (!companyName) continue;
      initial.push({
        companyName,
        system: cellStr(row, 3),
        contractNo: cellStr(row, 6),
        submittedAt: cellDate(row, 4), // วันที่ยื่นเอกสารขึ้นทะเบียน
        status: cellStr(row, 26),
        cerIssueDate: cellStr(row, 27),
        cerExpireDate: cellStr(row, 28),
        cerNo: cellStr(row, 29),
      });
    }
  }

  const su = [];
  const wsSU = workbook.getWorksheet('SU');
  if (wsSU) {
    for (let r = 3; r <= wsSU.rowCount; r++) {
      const row = wsSU.getRow(r);
      const companyName = cellStr(row, 2);
      if (!companyName) continue;
      su.push({
        companyName,
        system: cellStr(row, 3),
        su1Status: cellStr(row, 15),
        su1CertDate: cellStr(row, 13), // Cer. SU1 Date
        su1SentPlan: cellDate(row, 7), // Sent Plan SU1 Date
        su2Status: cellStr(row, 27),
        su2CertDate: cellStr(row, 25), // Cer. SU2 Date
        su2SentPlan: cellDate(row, 17), // Sent Plan SU2 Date
      });
    }
  }

  const recer = [];
  const wsRecer = workbook.getWorksheet('Recer');
  if (wsRecer) {
    for (let r = 3; r <= wsRecer.rowCount; r++) {
      const row = wsRecer.getRow(r);
      const companyName = cellStr(row, 2);
      if (!companyName) continue;
      recer.push({
        companyName,
        system: cellStr(row, 3),
        certExpireDate: cellStr(row, 4),
        submittedAt: cellDate(row, 7), // วันที่ยื่นขึ้นทะเบียนกับจีน
        status: cellStr(row, 18),
        cerIssueDate: cellStr(row, 19),
        cerNo: cellStr(row, 20),
      });
    }
  }

  return { initial, su, recer };
}

async function reloadCustomerDB() {
  const { initial, su, recer } = await loadWorkbook(XLSX_PATH);
  DB = { updatedAt: new Date(), initial, su, recer };
  console.log(`[customer-db] โหลดข้อมูลใหม่: Initial ${initial.length}, SU ${su.length}, Recer ${recer.length}`);
}

if (fs.existsSync(XLSX_PATH)) {
  reloadCustomerDB().catch((e) => console.error('โหลดไฟล์ Excel เดิมไม่สำเร็จ:', e.message));
}

// คำทั่วไปที่บ่งบอกแค่ "ประเภทนิติบุคคล" ไม่ใช่ชื่อเฉพาะของบริษัท - ตัดออกจากการค้นหาเสมอ (ทั้งจากคำค้นหาและชื่อบริษัทที่เทียบด้วย)
// เพราะบริษัทเกือบทุกชื่อมีคำพวกนี้ต่อท้าย ถ้าไม่ตัดออก พิมพ์แค่ "Co., Ltd." คำเดียวก็จะเจอทุกบริษัทเลย
const GENERIC_SUFFIX_WORDS = new Set([
  'co', 'ltd', 'limited', 'company', 'corp', 'corporation', 'inc', 'plc',
  'จำกัด', 'บริษัท', 'มหาชน',
]);

// ตัดจุด/คอมมา/โคลอน/ขีดกลางออก เพื่อให้ค้นหาแบบไม่สนใจเครื่องหมายวรรคตอนได้
// เช่น พิมพ์ "QTT" ก็เจอ "Q.T.T. Co., Ltd." เพราะทั้งคู่ normalize เหลือแค่ "qtt" (ตัดคำว่า co/ltd ออกไปแล้ว)
function norm(s) {
  const base = (s || '')
    .toLowerCase()
    .replace(/[.,:\-]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  const words = base.split(' ').filter((w) => w && !GENERIC_SUFFIX_WORDS.has(w));
  return words.join(' ').trim();
}

// ค้นหาชื่อบริษัทแบบ substring (ไม่สนตัวพิมพ์เล็ก/ใหญ่ ไม่สนเครื่องหมายวรรคตอน) จากทั้ง 3 ชีตพร้อมกัน
// ไม่ต้องพิมพ์ชื่อเต็ม แค่พิมพ์บางส่วน (คำแรก, ไม่กี่ตัวอักษร ฯลฯ) ที่ตรงกับส่วนใดส่วนหนึ่งของชื่อก็เจอ
function searchCompanies(query) {
  const q = norm(query);
  if (!q) return []; // คำค้นหาเป็นแค่คำทั่วไป (เช่น "Co., Ltd." ล้วนๆ) ไม่มีส่วนที่เจาะจงเหลืออยู่เลย ไม่ถือว่าเจอ
  const results = [];
  for (const rec of DB.initial) if (norm(rec.companyName).includes(q)) results.push({ type: 'initial', rec });
  for (const rec of DB.su) if (norm(rec.companyName).includes(q)) results.push({ type: 'su', rec });
  for (const rec of DB.recer) if (norm(rec.companyName).includes(q)) results.push({ type: 'recer', rec });
  return results;
}

// นับจำนวนวันตั้งแต่วันที่ที่กำหนด จนถึงวันนี้ (นับเป็นจำนวนเต็มวัน ไม่สนเวลาในวัน) - คืน null ถ้าไม่มีวันที่
function daysSince(fromMs) {
  if (fromMs === null || fromMs === undefined) return null;
  const from = new Date(fromMs);
  const fromMidnight = new Date(from.getFullYear(), from.getMonth(), from.getDate()).getTime();
  const now = new Date();
  const nowMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  return Math.round((nowMidnight - fromMidnight) / 86400000);
}

// คืนข้อความสถานะ "ที่กำลังดำเนินการอยู่" ของ record เดียว - ถ้าเสร็จ(Done)หมดแล้วคืน null (ไม่ต้องโชว์)
// ไม่บอกชื่อชีต (Initial/SU/Recer) ที่หัวข้อ แต่จะแปะ (Initial)/(Recer) ต่อท้ายสถานะจริงแทน เพื่อบอกว่าสถานะนี้มาจากรอบไหน
// latestRound (ถ้ามี) คือรอบล่าสุดที่มีวันที่บันทึกไว้ของบริษัทนี้ (นับรวม SU ด้วย) - { dateStr, label: 'Initial'|'SU1'|'SU2'|'Recer' }
//   ใช้สำหรับข้อความสถานะ "เสร็จสิ้นรอบ(X)" เท่านั้น
// certStart (ถ้ามี) คือวันที่ "ใบรับรองฉบับปัจจุบันเริ่มมีผล" - ดูจาก Initial/Recer เท่านั้น (ไม่รวม SU เพราะ SU ไม่ได้ออกใบรับรองใหม่)
// certExpire (ถ้ามี) คือวันที่ "ใบรับรองฉบับปัจจุบันหมดอายุ" - คู่กับ certStart เดียวกัน (ดูจาก Initial/Recer เท่านั้น)
//   ทั้งสองค่านี้โชว์ทุก sheet ยกเว้น Initial (เพราะรอบ Initial ยังไม่เคยได้ใบ cer มาก่อน จึงไม่มีข้อมูลใบเดิม)
// แปลง timestamp (ms) เป็นวันที่ dd/mm/yyyy สำหรับโชว์ควบคู่กับบรรทัด "ระยะเวลาดำเนินการ" - คืน null ถ้าไม่มีข้อมูลวันที่
function formatEpochDate(ms) {
  if (ms === null || ms === undefined) return null;
  return formatDate(new Date(ms));
}

function formatInProgress({ type, rec }, latestRound, certStart, certExpire) {
  if (type === 'initial') {
    if (isDoneStatus(rec.status)) return null;
    const statusText = rec.status ? `${rec.status}(Initial)` : '(ยังไม่ระบุ)';
    const lines = [`🆕 ${rec.companyName}`, `ระบบ: ${rec.system || '-'}`, `สถานะ: ${statusText}`];
    // ระยะเวลาดำเนินการ: นับเฉพาะตอนที่ Status มีการลงข้อมูลแล้ว (ไม่ใช่ช่องว่าง) - นับจากวันที่ยื่นเอกสารขึ้นทะเบียน ถึงวันนี้
    // ถ้าช่องวันที่ตั้งต้นว่าง (ข้อมูลต้นทางไม่มี) ให้นับเป็น 0 วัน แทนที่จะซ่อนบรรทัดไปเลย
    if (rec.status) {
      const d = daysSince(rec.submittedAt) ?? 0;
      const startText = formatEpochDate(rec.submittedAt);
      lines.push(`ระยะเวลาดำเนินการ: ${d} วัน${startText ? ` (เริ่ม ${startText})` : ''}`);
    }
    return lines.join('\n');
  }
  if (type === 'su') {
    const su1Done = isDoneStatus(rec.su1Status);
    const su2Done = isDoneStatus(rec.su2Status);
    if (su1Done && su2Done) return null; // ผ่านทั้ง SU1 และ SU2 แล้ว ถือว่าจบรอบนี้
    const lines = [`🔄 ${rec.companyName}`, `ระบบ: ${rec.system || '-'}`];
    // ถ้า SU1 ผ่านแล้ว แต่ SU2 ยังไม่มีข้อมูลอะไรเลยในไฟล์ (ยังไม่เริ่มตรวจ ไม่ใช่แค่ "ยังไม่เสร็จ")
    // สรุปเป็นสถานะเดียวให้อ่านง่ายชัดเจนไปเลยว่า "ตรวจ SU1 เสร็จแล้ว (วันที่...) รอตรวจ SU2" แทนที่จะโชว์ช่อง SU2 เป็น placeholder เฉยๆ
    // ใส่วันที่ตรวจ SU1 (Cer. SU1 Date) ต่อท้ายด้วยถ้ามีข้อมูล
    if (su1Done && !su2Done && !rec.su2Status) {
      const su1DateText = rec.su1CertDate ? ` (${rec.su1CertDate})` : '';
      lines.push(`สถานะ: ตรวจ SU1 เสร็จแล้ว${su1DateText} รอตรวจ SU2`);
    } else {
      // ถ้า SU1 ผ่านแล้วแต่ SU2 ยังไม่เสร็จ (มีข้อมูลบ้างแล้ว) โชว์เฉพาะ SU2 (ไม่โชว์ SU1 ที่ผ่านแล้วซ้ำ)
      // ระยะเวลาดำเนินการของแต่ละรอบ: นับจากวันที่ Sent Plan SU1/SU2 (เฉพาะรอบที่ Status มีข้อมูลแล้ว) ถึงวันนี้
      if (!su1Done) {
        lines.push(`SU1: ${rec.su1Status || '(ยังไม่ระบุ)'}`);
        if (rec.su1Status) {
          const d = daysSince(rec.su1SentPlan) ?? 0;
          const startText = formatEpochDate(rec.su1SentPlan);
          lines.push(`ระยะเวลาดำเนินการ SU1: ${d} วัน${startText ? ` (เริ่ม ${startText})` : ''}`);
        }
      }
      if (!su2Done) {
        lines.push(`SU2: ${rec.su2Status || '(ยังไม่ถึงรอบ/ยังไม่ระบุ)'}`);
        if (rec.su2Status) {
          const d = daysSince(rec.su2SentPlan) ?? 0;
          const startText = formatEpochDate(rec.su2SentPlan);
          lines.push(`ระยะเวลาดำเนินการ SU2: ${d} วัน${startText ? ` (เริ่ม ${startText})` : ''}`);
        }
      }
    }
    if (certStart) lines.push(`ใบรับรองเดิมเริ่ม: ${certStart}`);
    if (certExpire) lines.push(`ใบรับรองเดิมหมดอายุ: ${certExpire}`);
    return lines.join('\n');
  }
  // recer
  if (isDoneStatus(rec.status)) return null;
  const lines = [`♻️ ${rec.companyName}`, `ระบบ: ${rec.system || '-'}`];
  if (rec.status) {
    lines.push(`สถานะ: ${rec.status}(Recer)`);
    // ระยะเวลาดำเนินการ: นับจากวันที่ยื่นขึ้นทะเบียนกับจีน ถึงวันนี้ (เฉพาะตอนที่ Status มีการลงข้อมูลแล้ว)
    const d = daysSince(rec.submittedAt) ?? 0;
    const startText = formatEpochDate(rec.submittedAt);
    lines.push(`ระยะเวลาดำเนินการ: ${d} วัน${startText ? ` (เริ่ม ${startText})` : ''}`);
  } else if (latestRound) {
    // ช่อง Status ว่าง แต่มีรอบล่าสุดที่เสร็จแล้วอยู่ -> อนุมานว่ารอบล่าสุดเสร็จสิ้นแล้ว โชว์รอบ+วันที่แทนคำว่า "ยังไม่ระบุ"
    lines.push(`สถานะ: เสร็จสิ้นรอบ(${latestRound.label}) ${latestRound.dateStr}`);
  } else {
    lines.push(`สถานะ: (ยังไม่ระบุ)`);
  }
  if (certStart) lines.push(`ใบรับรองเดิมเริ่ม: ${certStart}`);
  if (certExpire) lines.push(`ใบรับรองเดิมหมดอายุ: ${certExpire}`);
  return lines.join('\n');
}

// แปลงวันที่แบบ "dd/mm/yyyy" (string ที่ format ไว้แล้ว) กลับเป็นตัวเลขไว้เทียบว่าอันไหนใหม่กว่า
function parseThaiDate(str) {
  if (!str) return null;
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(str.trim());
  if (!m) return null;
  const [, d, mo, y] = m;
  const t = new Date(Number(y), Number(mo) - 1, Number(d)).getTime();
  return Number.isNaN(t) ? null : t;
}

// รหัสมาตรฐานที่บริษัทนี้ตรวจ (พบในข้อมูลจริงของไฟล์ทั้งหมด) - ใช้เพื่อรู้จำ "แท็กระบบ" ที่แอดมินแปะต่อ/แทรกไว้ในชื่อบริษัท
// เช่น "ABC Co., Ltd. (9001)", "ABC Co., Ltd. [14001]", "ABC Co., Ltd. 9001" ซึ่งไม่ใช่ส่วนหนึ่งของชื่อบริษัทจริงๆ
// แต่เป็นโน้ตบอกว่าแถวนี้เกี่ยวกับระบบ/มาตรฐานไหน (บริษัทเดียวกันบางเจ้าตรวจหลายระบบแยกกันคนละรอบ เช่น 9001 กับ 14001)
const KNOWN_STANDARD_CODES = ['9001', '14001', '45001'];
// (A) แท็กระบบที่มีวงเล็บ(เหลี่ยม/กลม)เปิด-ปิดครบ และในวงเล็บมี "รหัสเดียวล้วนๆ" เท่านั้น เช่น "(9001)", "[14001]", "(9001:2015)"
// อยู่ตรงไหนของชื่อก็ตัดได้ (ไม่ต้องอยู่ท้ายสุด เช่น "ABC Co., Ltd. [14001] (2)") - ต้องมีวงเล็บเปิดจริงๆ ไม่ใช่ optional
// เพื่อไม่ให้ไปแมตช์ท้ายๆ ของวงเล็บที่ระบุหลายมาตรฐานพร้อมกัน เช่น "(9001 & 14001)" หรือ "(9001, 14001, 45001)"
// (พวกนี้ไม่มีวงเล็บเปิดอยู่ติดกับรหัสตัวท้ายเลย เลยไม่ตรง pattern นี้ - ปลอดภัย)
const BRACKETED_SINGLE_TAG_RE = new RegExp(
  `[\\[\\(]\\s*(?:${KNOWN_STANDARD_CODES.join('|')})(?::\\d{4})?\\s*[\\]\\)]`,
  'gi'
);
// (B) แท็กระบบท้ายชื่อแบบไม่มีวงเล็บเลย เช่น "... 9001" - ต้องอยู่ท้ายสุดจริงๆ (ไม่มีวงเล็บปิดตามหลัง กันซ้อนกับ (A))
const BARE_TRAILING_TAG_RE = new RegExp(`(?:^|\\s)(?:${KNOWN_STANDARD_CODES.join('|')})(?::\\d{4})?\\s*$`, 'i');
function stripSystemTag(name) {
  let s = (name || '').replace(BRACKETED_SINGLE_TAG_RE, ' ').replace(/\s+/g, ' ').trim();
  s = s.replace(BARE_TRAILING_TAG_RE, '').trim();
  return s || (name || '').trim(); // กันเผื่อ regex กินชื่อทั้งหมดจนว่าง (ไม่ควรเกิด แต่กันไว้)
}

// (C) แท็กระบบท้ายชื่อที่ระบุ "หลายมาตรฐานรวมกัน" เช่น "(9001 & 45001)" หรือ "(9001, 14001, 45001)"
// ตัดออกได้ก็ต่อเมื่อชุดมาตรฐานในวงเล็บ "ตรงกับช่องระบบ ของแถวนั้นเป๊ะๆ" เท่านั้น (แปลว่าเป็นแค่โน้ตย้ำขอบข่ายของรอบเดียวกัน)
// ถ้าไม่ตรง (เช่นวงเล็บบอก 3 มาตรฐาน แต่ช่องระบบมีแค่มาตรฐานเดียว) ถือว่าข้อมูลน่าสงสัย ไม่แตะ ปล่อยให้แยกกลุ่มไว้เพื่อให้เห็นความผิดปกติ
// (เจอจากเคสจริง เช่น C.E.G. Engineering ที่ตัดแล้วรวมกลุ่มถูก ตรงข้ามกับ BEST IN GROUND บางแถวที่ตัดไม่ได้เพราะช่องระบบขัดแย้งกับวงเล็บ)
function extractSystemCodes(text) {
  return new Set((text || '').match(new RegExp(KNOWN_STANDARD_CODES.join('|'), 'g')) || []);
}
function sameCodeSet(a, b) {
  return a.size > 0 && a.size === b.size && [...a].every((c) => b.has(c));
}
const COMPOUND_BRACKET_RE = /\(([^()]+)\)\s*$/;
function stripCompoundSystemTagIfSafe(name, system) {
  const trimmed = (name || '').trim();
  const m = COMPOUND_BRACKET_RE.exec(trimmed);
  if (!m) return trimmed;
  const inside = m[1];
  const codesInBracket = extractSystemCodes(inside);
  if (codesInBracket.size < 2) return trimmed; // วงเล็บโค้ดเดียวจัดการโดย (A)/(B) ไปแล้ว ไม่เกี่ยวกับฟังก์ชันนี้
  // เนื้อหาในวงเล็บต้องมีแค่รหัสมาตรฐาน + เครื่องหมายคั่น (comma, &, colon, ปี ค.ศ., เว้นวรรค) เท่านั้น ห้ามมีคำอื่นปนเลย
  const residue = inside
    .replace(new RegExp(KNOWN_STANDARD_CODES.join('|'), 'g'), '')
    .replace(/\d{4}/g, '')
    .replace(/[\s,:&]/g, '');
  if (residue !== '') return trimmed;
  if (!sameCodeSet(codesInBracket, extractSystemCodes(system))) return trimmed;
  return trimmed.slice(0, m.index).trim();
}

// รวมชื่อระบบที่มีได้หลายบรรทัด (เช่น "9001:2015\n14001:2015") ให้เทียบกันได้โดยไม่สนลำดับ/ตัวพิมพ์เล็กใหญ่
// ใช้เป็นส่วนหนึ่งของ key ตอนจัดกลุ่มบริษัท กันไม่ให้ข้อมูลของคนละระบบ (เช่น SU ของ 9001) ไปปนกับอีกระบบ (เช่น Initial/Recer ของ 14001)
function normalizeSystem(system) {
  if (!system) return '';
  return system
    .split('\n')
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean)
    .sort()
    .join('+');
}

// บางบริษัทในไฟล์มีชื่อซ้ำแบบมีเลขวงเล็บต่อท้าย เช่น "ABC Co., Ltd." กับ "ABC Co., Ltd. (2)"
// ซึ่งหมายถึงบริษัทเดียวกันแต่เป็นรอบ/สัญญาใหม่ - ตัดวงเล็บออกเพื่อจัดกลุ่มเป็นบริษัทเดียวกัน
// และถือว่าเลขวงเล็บที่สูงกว่า = รอบที่ใหม่กว่าเสมอ (ไม่มีวงเล็บ = รอบแรก/เก่าที่สุด)
function companyNameParts(name, system) {
  const trimmed = stripSystemTag(stripCompoundSystemTagIfSafe((name || '').trim(), system));

  // กรณีวงเล็บเป็นเลขล้วนๆ เช่น "ABC Co., Ltd. (2)"
  let m = /^(.*?)\s*\((\d+)\)\s*$/.exec(trimmed);
  if (m) return { base: m[1].trim(), suffix: Number(m[2]) };

  // กรณีวงเล็บระบุรอบงานเป็นคำ+เลข เช่น "ABC(Recer2)", "ABC Co., Ltd.(Recer3)", "ABC (New Initial)", "ABC (2nd New Initial)"
  // (พบในไฟล์จริงว่าแอดมินบางแถวใช้ป้ายกำกับแบบนี้แทนเลขวงเล็บล้วนๆ ต้องรองรับด้วย ไม่งั้นแถวพวกนี้จะไม่ถูกรวมเป็นบริษัทเดียวกัน)
  // ถ้าไม่มีเลขกำกับเลย (เช่น "(New Initial)" เฉยๆ) ถือเป็นรอบที่ 1
  // ไม่แตะวงเล็บอื่นๆ ที่ไม่ใช่รูปแบบนี้ (เช่น "(9001 & 14001)" ที่บอกขอบข่ายที่ตรวจ หรือ "(ยกเลิก)" หรือโน้ตอื่นๆ) เพราะไม่ใช่ตัวบ่งบอกรอบงานซ้ำ
  m = /^(.*?)\s*\(\s*(?:(\d+)\s*(?:st|nd|rd|th)?\s*)?(?:recer|su|initial|new\s*initial)\s*(\d+)?\s*\)\s*$/i.exec(trimmed);
  if (m) {
    const leading = m[2] !== undefined ? Number(m[2]) : null;
    const trailing = m[3] !== undefined ? Number(m[3]) : null;
    const suffix = trailing !== null ? trailing : leading !== null ? leading : 1;
    return { base: m[1].trim(), suffix };
  }

  return { base: trimmed, suffix: 0 };
}

// เลือก record ที่ "ใหม่ที่สุด" จากลิสต์เดียวกัน - ดูเลขวงเล็บต่อท้ายชื่อบริษัทก่อน (สูงกว่า = ใหม่กว่า)
// ถ้าเลขวงเล็บเท่ากัน ค่อยเทียบจากวันออกใบรับรอง แล้ว fallback ไปวันที่ยื่นขึ้นทะเบียน
function pickLatestRecord(records) {
  if (records.length === 0) return null;
  return records
    .slice()
    .sort((a, b) => {
      const sa = companyNameParts(a.companyName, a.system).suffix;
      const sb = companyNameParts(b.companyName, b.system).suffix;
      if (sb !== sa) return sb - sa;
      const da = parseThaiDate(a.cerIssueDate) ?? a.submittedAt ?? 0;
      const db = parseThaiDate(b.cerIssueDate) ?? b.submittedAt ?? 0;
      return db - da;
    })[0];
}

// เมื่อทุกอย่าง (ทุก record ที่พบของบริษัทนี้) Done หมดแล้ว - สรุปรอบล่าสุดแทน
// Recer ถือว่าเป็นรอบที่ใหม่กว่า Initial เสมอ (เกิดขึ้นทีหลัง Initial ในไทม์ไลน์จริง)
// ถ้ามี record ของ Recer อยู่ ให้ใช้ Recer ล่าสุดเป็นตัวสรุป ถ้าไม่มีเลยค่อย fallback ไปที่ Initial ล่าสุด
function buildLatestRoundSummary(companyName, records) {
  const initials = records.filter((r) => r.type === 'initial').map((r) => r.rec);
  const recers = records.filter((r) => r.type === 'recer').map((r) => r.rec);
  const rec = pickLatestRecord(recers) || pickLatestRecord(initials);
  if (!rec) {
    return `✅ ${companyName} — ทุกขั้นตอนเสร็จสมบูรณ์แล้วครับ ไม่มีสถานะที่กำลังดำเนินการอยู่`;
  }
  const lines = [
    `✅ ${companyName} — ทุกขั้นตอนเสร็จสมบูรณ์แล้ว (ไม่มีสถานะที่กำลังดำเนินการอยู่) สรุปรอบล่าสุด:`,
    `ระบบ/ขอบข่ายที่ตรวจ: ${rec.system || '-'}`,
  ];
  if (rec.cerNo) lines.push(`เลขที่ใบรับรอง: ${rec.cerNo}`);
  if (rec.cerIssueDate) lines.push(`วันที่ออกใบรับรอง: ${rec.cerIssueDate}`);
  // ช่อง "วันหมดอายุ" ชื่อคอลัมน์ในไฟล์ไม่เหมือนกันระหว่าง Initial (cerExpireDate) กับ Recer (certExpireDate มี t เพิ่ม)
  // ต้องเช็คทั้ง 2 ชื่อ ไม่งั้นถ้า rec ที่เลือกมาเป็น Recer จะดึงค่าไม่เจอ (ได้ undefined) แล้วเงียบๆ ไม่โชว์บรรทัดนี้เลย
  const expireDate = rec.cerExpireDate || rec.certExpireDate;
  if (expireDate) lines.push(`วันที่หมดอายุ: ${expireDate}`);
  return lines.join('\n');
}

// หารอบที่มีวันที่บันทึกไว้ "ใกล้ปัจจุบันที่สุด" ของบริษัทนี้ - เทียบทั้ง 4 แหล่ง:
// Initial (Cer Issue Date), SU1 (Cer. SU1 Date), SU2 (Cer. SU2 Date), Recer (Cer Issue Date)
// ไม่สนว่า record นั้น Done หรือยัง (record ที่ยังไม่เสร็จมักไม่มีวันที่อยู่แล้ว จึงไม่ถูกเลือกโดยธรรมชาติ)
// คืนค่า { dateStr, label } โดย label บอกว่าวันที่นั้นมาจากรอบไหน (Initial/SU1/SU2/Recer)
function latestCompletedRound(records) {
  const candidates = [];
  for (const { type, rec } of records) {
    if (type === 'initial' || type === 'recer') {
      const t = parseThaiDate(rec.cerIssueDate);
      if (t !== null) candidates.push({ t, dateStr: rec.cerIssueDate, label: type === 'initial' ? 'Initial' : 'Recer' });
    } else if (type === 'su') {
      const t1 = parseThaiDate(rec.su1CertDate);
      if (t1 !== null) candidates.push({ t: t1, dateStr: rec.su1CertDate, label: 'SU1' });
      const t2 = parseThaiDate(rec.su2CertDate);
      if (t2 !== null) candidates.push({ t: t2, dateStr: rec.su2CertDate, label: 'SU2' });
    }
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.t - a.t);
  return candidates[0];
}

// หา record ที่ถือว่าเป็น "รอบปัจจุบัน" ของใบรับรอง เพื่อเอาวันที่เริ่ม/หมดอายุมาโชว์ - ใช้ตรรกะเดียวกับ buildLatestRoundSummary
// คือ Recer ถือว่าใหม่กว่า Initial เสมอ (ถ้ามี Recer ให้ใช้ Recer รอบล่าสุดเป็นตัวอ้างอิง ไม่ใช้ Initial เก่าแทนแม้ Recer จะยังไม่กรอกวันที่ก็ตาม)
// สำคัญ: ต้องยึดตาม "รอบที่ใหม่ที่สุดจริงๆ" ไม่ใช่ "รอบไหนก็ได้ที่บังเอิญมีวันที่กรอกไว้" ไม่งั้นถ้า Recer ยังไม่กรอกวันที่
// จะเผลอไปดึงวันที่ของ Initial เก่ามาโชว์เป็นวันที่ "ปัจจุบัน" ซึ่งผิด (เจอจากเคสจริง เช่น DEE Q OLO ASSET)
function pickCurrentCertRecord(records) {
  const initials = records.filter((r) => r.type === 'initial').map((r) => r.rec);
  const recers = records.filter((r) => r.type === 'recer').map((r) => r.rec);
  const recerPick = pickLatestRecord(recers);
  if (recerPick) return { rec: recerPick, type: 'recer' };
  const initialPick = pickLatestRecord(initials);
  if (initialPick) return { rec: initialPick, type: 'initial' };
  return null; // ไม่มีทั้ง Initial และ Recer เลย (เช่น กลุ่มที่มีแต่ SU) - ไม่มี "รอบปัจจุบัน" ให้อ้างอิง
}

// วันที่ "ใบรับรองฉบับปัจจุบันเริ่มมีผล" - ดูจาก Cer Issue Date ของรอบปัจจุบัน (Initial หรือ Recer ล่าสุด) เท่านั้น
// ไม่ใช้วันที่ของ SU เลย เพราะ SU เป็นแค่การตรวจติดตามใบรับรองเดิม ไม่ได้ออกใบรับรองใหม่
// ถ้ามีรอบปัจจุบันแต่ยังไม่ได้กรอกวันที่ไว้ ให้บอกตรงๆ ว่า "ไม่ได้บันทึก" (ไม่ใช่ไปเงียบๆ ดึงวันที่ของรอบเก่ากว่ามาแทน)
// ถ้าไม่มีรอบปัจจุบันเลย (ไม่มีทั้ง Initial/Recer) คืน null เพื่อให้ฝั่งเรียกใช้ซ่อนบรรทัดนี้ไปเลย
function latestCertStart(records) {
  const current = pickCurrentCertRecord(records);
  if (!current) return null;
  return current.rec.cerIssueDate || 'ไม่ได้บันทึก';
}

// วันที่ "ใบรับรองฉบับปัจจุบันหมดอายุ" - คู่กับ latestCertStart ใช้รอบปัจจุบันเดียวกัน (Initial ใช้ Cer Expire Date,
// Recer ใช้วันที่ cer จีนหมดอายุ/certExpireDate) เพื่อให้สอดคล้องกับใบรับรองฉบับเดียวกับที่ certStart อ้างถึงเสมอ
function latestCertExpire(records) {
  const current = pickCurrentCertRecord(records);
  if (!current) return null;
  const val = current.type === 'recer' ? current.rec.certExpireDate : current.rec.cerExpireDate;
  return val || 'ไม่ได้บันทึก';
}

// คีย์สำหรับ "จัดกลุ่ม" ชื่อบริษัทเข้าด้วยกัน (คนละอันกับ companyNameParts().base ที่ใช้โชว์ผล)
// ตัดจุด/คอมมา/เว้นวรรคซ้ำ/ตัวพิมพ์เล็กใหญ่ออกก่อนเทียบ เพราะบางแถวในไฟล์จริงพิมพ์ชื่อบริษัทเดียวกันไม่ตรงกันเป๊ะ
// (เช่น "...CO., LTD (ยกเลิก)" ไม่มีจุด กับ "...CO., LTD. (ยกเลิก)" มีจุด) ถ้าไม่ตัดจะถูกมองว่าเป็นคนละบริษัท
function groupingKey(base) {
  return (base || '')
    .toLowerCase()
    .replace(/[.,]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// เช็คว่าชื่อบริษัท (record ใดก็ตามของกลุ่มนี้) มีคำว่า "ยกเลิก" กำกับไว้ไหม - แปลว่ารายการนี้ถูกยกเลิกไปแล้ว
function isCancelledName(name) {
  return /ยกเลิก/.test(name || '');
}

// รวม record ทั้งหมดของบริษัทเดียวกันเข้าด้วยกัน แล้วตัดสินใจว่าจะโชว์สถานะที่กำลังทำอยู่ หรือสรุปรอบล่าสุด
function buildCompanyBlock(companyName, records) {
  // ถ้ามีคำว่า "ยกเลิก" กำกับอยู่ในชื่อบริษัทของ record ไหนก็ตามในกลุ่มนี้ ถือว่ารายการนี้ถูกยกเลิกไปแล้ว
  // ไม่ต้องไปคำนวณสถานะ/วันที่ตามปกติ เพราะข้อมูลอาจไม่สมบูรณ์และไม่มีความหมายอีกต่อไป
  const cancelled = records.some((r) => isCancelledName(r.rec.companyName)) || isCancelledName(companyName);
  if (cancelled) {
    const cleanName = companyName
      .replace(/\s*\(\s*ยกเลิก\s*\)\s*/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    return `🚫 ${cleanName}\nรายการนี้ถูกยกเลิกแล้วครับ (มีการระบุ "ยกเลิก" ไว้ในข้อมูล) กรุณาตรวจสอบกับแอดมินอีกครั้งหากต้องการรายละเอียดเพิ่มเติม`;
  }
  const latestRound = latestCompletedRound(records);
  const certStart = latestCertStart(records);
  const certExpire = latestCertExpire(records);
  const inProgress = records.map((r) => formatInProgress(r, latestRound, certStart, certExpire)).filter(Boolean);
  if (inProgress.length > 0) return inProgress.join('\n\n');
  return buildLatestRoundSummary(companyName, records);
}

const MAX_RESULTS = 8;

function buildStatusReply(query) {
  if (!DB.updatedAt) {
    return 'ยังไม่มีข้อมูลในระบบครับ รบกวนให้แอดมินอัปโหลดไฟล์ Excel ที่หน้า /upload ก่อนนะครับ';
  }
  if (!norm(query)) {
    return `"${query}" เป็นคำทั่วไปเกินไปครับ (เช่น Co., Ltd. / บริษัท / จำกัด) รบกวนพิมพ์ชื่อเฉพาะของบริษัทด้วยครับ`;
  }
  const results = searchCompanies(query);
  if (results.length === 0) {
    return `ไม่พบชื่อบริษัทที่ตรงกับ "${query}" ครับ ลองพิมพ์บางส่วนของชื่อบริษัทดูใหม่`;
  }

  // จัดกลุ่ม record ตามชื่อบริษัท (query แบบ substring อาจเจอได้หลายบริษัท)
  // ใช้ชื่อบริษัทแบบตัดเลขวงเล็บ/แท็กระบบต่อท้ายออก (companyNameParts().base) เป็นชื่อที่โชว์
  // แต่ใช้ groupingKey() (ตัดจุด/คอมมา/เว้นวรรคซ้ำ/ตัวพิมพ์เล็กใหญ่) + ระบบมาตรฐานที่ตรวจ (normalizeSystem)
  // เป็น key ในการรวมกลุ่มจริงๆ - ต้องรวมระบบเข้าไปด้วย เพราะบางบริษัทตรวจหลายระบบ (เช่น 9001 กับ 14001)
  // แยกกันคนละรอบ/คนละไทม์ไลน์ ถ้ารวมกลุ่มแค่ตามชื่อเฉยๆ ข้อมูล SU/Recer ของระบบหนึ่งจะไปปนกับอีกระบบ
  // (เจอจากเคสจริง เช่น VBS. SERVICE COMPANY LIMITED และ Centralize Power Industry Co., Ltd.)
  // ทำ 2 รอบ: รอบแรกรวม record ที่ "มี" ระบบระบุไว้ก่อน (แยกกลุ่มตามระบบจริงๆ)
  // รอบสองค่อยจัดการ record ที่ "ไม่มี" ระบบระบุไว้ (ช่องว่าง/null เช่นแถว Recer ที่ถูกยกเลิกจนข้อมูลว่างหมด)
  // - ถ้าบริษัทนี้มีกลุ่มที่มีระบบอยู่แล้ว ให้รวมเข้ากลุ่มแรกที่เจอไปเลย (ดีกว่าแยกเป็นกลุ่มลอยๆ ไม่มีข้อมูลอะไรเลย)
  // - ถ้ายังไม่มีกลุ่มไหนของบริษัทนี้เลย ค่อยตั้งกลุ่มใหม่แบบไม่ระบุระบบ
  const byCompany = new Map(); // "groupingKey::system" -> { displayName, baseKey, records }
  const withoutSystem = [];
  for (const r of results) {
    const base = companyNameParts(r.rec.companyName, r.rec.system).base;
    const baseKey = groupingKey(base);
    const sysKey = normalizeSystem(r.rec.system);
    if (!sysKey) {
      withoutSystem.push({ r, base, baseKey });
      continue;
    }
    const key = baseKey + '::' + sysKey;
    if (!byCompany.has(key)) byCompany.set(key, { displayName: base, baseKey, records: [] });
    byCompany.get(key).records.push(r);
  }
  for (const { r, base, baseKey } of withoutSystem) {
    const existing = [...byCompany.values()].find((v) => v.baseKey === baseKey);
    if (existing) {
      existing.records.push(r);
    } else {
      const key = baseKey + '::';
      if (!byCompany.has(key)) byCompany.set(key, { displayName: base, baseKey, records: [] });
      byCompany.get(key).records.push(r);
    }
  }

  // ถ้าบริษัทเดียวกัน (ชื่อฐานเดียวกัน) ถูกแยกเป็นหลายกลุ่มเพราะคนละระบบ ให้ต่อท้ายชื่อที่โชว์ด้วยระบบนั้นๆ
  // กันสับสนว่าบล็อกไหนเป็นของระบบไหน (ถ้ามีระบบเดียวกลุ่มเดียวอยู่แล้ว ไม่ต้องต่อท้ายให้รกโดยไม่จำเป็น)
  const baseKeyCount = new Map();
  for (const v of byCompany.values()) baseKeyCount.set(v.baseKey, (baseKeyCount.get(v.baseKey) || 0) + 1);
  for (const v of byCompany.values()) {
    if (baseKeyCount.get(v.baseKey) > 1) {
      const sysLabel = v.records.map((r) => r.rec.system).find(Boolean);
      if (sysLabel) v.displayName = `${v.displayName} [${sysLabel.replace(/\n/g, ' & ')}]`;
    }
  }

  if (byCompany.size > MAX_RESULTS) {
    const names = [...byCompany.values()].map((v) => v.displayName).slice(0, MAX_RESULTS);
    return (
      `พบ ${byCompany.size} บริษัทที่ตรงกับ "${query}" เยอะเกินไป กรุณาพิมพ์ชื่อให้เจาะจงมากขึ้นครับ\n\n` +
      `ตัวอย่างที่พบ:\n` +
      names.map((n) => '- ' + n).join('\n')
    );
  }

  const blocks = [...byCompany.values()].map(({ displayName, records }) => buildCompanyBlock(displayName, records));
  return blocks.join('\n\n---\n\n');
}
// ==== จบส่วนเช็คสถานะงานลูกค้า ====================================================

// ---- สุ่มเมนูอาหารไทย (200 เมนู พร้อมแคลอรี่โดยประมาณ) ----
// FOOD_MENU: ชื่อเมนู -> แคลอรี่โดยประมาณต่อ 1 ที่ (kcal) - เป็นค่าประมาณคร่าวๆ ไม่ใช่ค่าทางโภชนาการที่แม่นยำ
const FOOD_MENU = {
  // เมนูจานเดียว จบในตัวเอง (รวมข้าว/เส้นแล้ว) - สุ่มแล้วเสนอเดี่ยวๆ ได้
  'ผัดไทย': 486, 'ข้าวผัดกระเพราหมูสับ': 600, 'ข้าวผัดกระเพราไก่': 550,
  'ข้าวผัดกระเพราหมูกรอบ': 610, 'ข้าวมันไก่': 600, 'ข้าวหมูแดง': 650,
  'ข้าวขาหมู': 590, 'ข้าวคลุกกะปิ': 550, 'ข้าวหน้าเป็ด': 580,
  'ก๋วยเตี๋ยวเรือ': 350, 'ก๋วยเตี๋ยวต้มยำ': 530, 'ก๋วยเตี๋ยวหมูตุ๋น': 420,
  'บะหมี่เกี๊ยวหมูแดง': 400, 'เย็นตาโฟ': 400, 'ก๋วยจั๊บน้ำข้น': 650,
  'ข้าวซอยไก่': 600, 'ผัดซีอิ๊วหมู': 580, 'ผัดหมี่โคราช': 510, 'หมี่กะทิ': 460,
  'ขนมจีนน้ำยา': 400, 'ขนมจีนน้ำพริก': 540, 'สุกี้น้ำ': 350, 'สุกี้แห้ง': 550,
  'ชาบู': 400, 'ข้าวต้มกุ้ง': 300, 'ข้าวต้มปลา': 400, 'โจ๊กหมู': 300,
  'ข้าวไข่เจียวหมูสับ': 400, 'ข้าวกะเพราทะเล': 360, 'ผัดไทยกุ้งสด': 500,
  'ก๋วยเตี๋ยวคั่วไก่': 650, 'ข้าวหน้าไก่ทอด': 520, 'ข้าวมันส้มตำไก่ทอด': 400,
  'ราดหน้าหมู': 380, 'ราดหน้าทะเล': 530, 'บะหมี่แห้งหมูแดง': 630,
  'ข้าวหมกไก่': 450, 'ข้าวผัดปู': 430, 'ข้าวผัดอเมริกัน': 560, 'ผัดมาม่าไข่เจียว': 550,
  // แกงกะทิ/แกงน้ำ
  'แกงเขียวหวานไก่': 450, 'แกงเขียวหวานลูกชิ้น': 340, 'แกงมัสมั่นไก่': 420,
  'แกงมัสมั่นเนื้อ': 400, 'แกงพะแนงหมู': 280, 'แกงพะแนงไก่': 230,
  'แกงป่าหมู': 330, 'แกงป่าไก่': 280, 'แกงส้มผักรวม': 280, 'แกงส้มกุ้ง': 440,
  'แกงเหลืองปลา': 430, 'แกงไตปลา': 350, 'แกงกะหรี่ไก่': 300, 'แกงคั่วสับปะรด': 310,
  'แกงบวดฟักทอง': 300, 'แกงจืดเต้าหู้หมูสับ': 300, 'แกงจืดวุ้นเส้น': 330,
  'แกงจืดมะระยัดไส้': 370, 'แกงเลียงผักรวม': 240, 'แกงอ่อมไก่': 350,
  // ต้ม/ต้มยำ/ซุป
  'ต้มยำกุ้งน้ำข้น': 250, 'ต้มยำกุ้งน้ำใส': 150, 'ต้มยำไก่': 250, 'ต้มยำปลา': 130,
  'ต้มยำทะเล': 190, 'ต้มข่าไก่': 210, 'ต้มข่ากุ้ง': 170, 'ต้มแซ่บกระดูกหมู': 270,
  'ต้มแซ่บซี่โครงอ่อน': 120, 'ต้มจับฉ่าย': 170, 'ต้มส้มปลา': 170,
  'ต้มโคล้งปลาแห้ง': 190, 'แกงจืดฟักเขียว': 120, 'ซุปหางวัว': 140, 'ซุปเห็ด': 260,
  'ข้าวต้มกระดูกหมู': 250, 'ต้มยำเห็ดรวม': 270, 'ต้มยำปลากระป๋อง': 170,
  'ต้มไก่บ้านใส่ข่า': 160, 'ต้มแซ่บหมูสามชั้น': 130,
  // ผัด
  'ผัดกะเพราหมูสับ': 340, 'ผัดกะเพราไก่': 200, 'ผัดกะเพราเนื้อ': 310,
  'ผัดกะเพราทะเล': 220, 'ผัดกะเพราหมูกรอบ': 370, 'หมูกระเทียมพริกไทย': 300,
  'ไก่กระเทียมพริกไทย': 340, 'ผัดผักบุ้งไฟแดง': 280, 'ผัดคะน้าหมูกรอบ': 370,
  'ผัดคะน้าปลาเค็ม': 370, 'ผัดถั่วงอกตับ': 320, 'ผัดมะเขือยาวหมูสับ': 200,
  'ผัดฟักทองไข่': 280, 'ผัดวุ้นเส้นหมูสับ': 360, 'ผัดผักรวมมิตร': 380,
  'ไก่ผัดเม็ดมะม่วง': 320, 'ไก่ผัดขิง': 370, 'เนื้อผัดน้ำมันหอย': 280,
  'หมูผัดน้ำมันหอย': 200, 'ผัดถั่วฝักยาวหมูสับ': 380, 'ผัดเปรี้ยวหวานหมู': 230,
  'ผัดเปรี้ยวหวานไก่': 370, 'ผัดพริกแกงหมูป่า': 330, 'ผัดฉ่าทะเล': 240,
  'ผัดฉ่าหมูป่า': 250, 'กุ้งผัดพริกเกลือ': 320, 'ปลาหมึกผัดไข่เค็ม': 240,
  'ผัดกะหล่ำน้ำปลา': 320, 'ผัดบวบใส่ไข่': 340, 'คะน้าน้ำมันหอย': 360,
  // ของทอด
  'ไก่ทอดหาดใหญ่': 320, 'ไก่ทอดกระเทียม': 270, 'ปีกไก่ทอด': 380,
  'หมูทอดกระเทียม': 210, 'ปลาทอดน้ำปลา': 320, 'ปลาทอดขมิ้น': 310,
  'ทอดมันปลา': 330, 'ทอดมันข้าวโพด': 240, 'หอยทอด': 380, 'กุ้งชุบแป้งทอด': 320,
  'ปลาหมึกทอดกระเทียม': 180, 'ไข่เจียวหมูสับ': 320, 'ไข่เจียวปู': 240,
  'ไข่ลูกเขย': 210, 'ไข่พะโล้': 190, 'หมูพะโล้': 360, 'ไก่พะโล้': 360,
  'หมูสามชั้นทอด': 270, 'เต้าหู้ทอด': 360, 'ปอเปี๊ยะทอด': 190, 'ลูกชิ้นทอด': 320,
  'กล้วยแขก': 170, 'มันทอด': 340, 'ข้าวเกรียบกุ้ง': 340, 'แคบหมู': 250,
  // ปิ้งย่าง
  'หมูปิ้งข้าวเหนียว': 250, 'ไก่ย่าง': 280, 'คอหมูย่าง': 300, 'หมูย่างเกาหลี': 270,
  'ปลาเผา': 340, 'ปลาหมึกย่าง': 230, 'กุ้งเผา': 130, 'ไก่ย่างวิเชียรบุรี': 290,
  'เนื้อย่างจิ้มแจ่ว': 180, 'ตับหวาน': 240, 'ไส้กรอกอีสาน': 260, 'ลูกชิ้นปิ้ง': 300,
  'ปีกไก่ย่าง': 130, 'เป็ดย่าง': 130, 'ปลาช่อนเผา': 320, 'หอยแมลงภู่เผา': 200,
  'ข้าวเหนียวหมูย่าง': 230, 'สันคอหมูย่างนมสด': 140, 'ปลาดุกย่าง': 140, 'ไข่ปิ้ง': 320,
  // ยำ/ลาบ/ตำ
  'ส้มตำไทย': 120, 'ส้มตำปูปลาร้า': 150, 'ส้มตำทะเล': 220, 'ลาบหมู': 300,
  'ลาบไก่': 250, 'ลาบเนื้อ': 150, 'น้ำตกหมู': 210, 'น้ำตกเนื้อ': 210,
  'ยำวุ้นเส้น': 220, 'ยำวุ้นเส้นทะเล': 280, 'ยำมาม่า': 180, 'ยำปลาดุกฟู': 140,
  'ยำไข่ดาว': 160, 'ยำหมูยอ': 100, 'ยำแหนม': 280, 'ยำสามกรอบ': 100,
  'ยำเห็ดรวม': 130, 'ยำถั่วพู': 160, 'ยำกุ้งสด': 180, 'ยำปลากระป๋อง': 290,
  'พล่ากุ้ง': 130, 'พล่าทะเล': 290, 'ตำแตง': 240, 'ตำถั่ว': 180, 'ซกเล็ก': 190,
  // น้ำพริก/เครื่องเคียง
  'น้ำพริกกะปิผักต้ม': 150, 'น้ำพริกปลาทู': 220, 'น้ำพริกหนุ่ม': 90,
  'น้ำพริกอ่อง': 200, 'น้ำพริกลงเรือ': 170, 'น้ำพริกมะขาม': 140,
  'จิ้มจุ่มหมู': 140, 'จิ้มจุ่มทะเล': 160, 'หมูสะเต๊ะ': 270, 'ไข่ต้มยางมะตูม': 270,
  'ผักบุ้งไฟแดง(เครื่องเคียง)': 270, 'แหนมทอด': 110, 'ปูนิ่มทอดกระเทียม': 180,
  'สะเต๊ะไก่': 100, 'ข้าวเหนียวมะม่วง': 300, 'กุ้งแช่น้ำปลา': 210,
  'ยำถั่วฝักยาว': 230, 'ผักชุบแป้งทอด': 170, 'ทอดมันกุ้ง': 120, 'ปีกไก่ยัดไส้': 190,
};

// เมนูจานเดียวจบในตัว (ใช้ตอนสุ่มเป็นเมนูเดี่ยว)
const ONEPLATE_NAMES = [
  'ผัดไทย', 'ข้าวผัดกระเพราหมูสับ', 'ข้าวผัดกระเพราไก่', 'ข้าวผัดกระเพราหมูกรอบ',
  'ข้าวมันไก่', 'ข้าวหมูแดง', 'ข้าวขาหมู', 'ข้าวคลุกกะปิ', 'ข้าวหน้าเป็ด',
  'ก๋วยเตี๋ยวเรือ', 'ก๋วยเตี๋ยวต้มยำ', 'ก๋วยเตี๋ยวหมูตุ๋น', 'บะหมี่เกี๊ยวหมูแดง',
  'เย็นตาโฟ', 'ก๋วยจั๊บน้ำข้น', 'ข้าวซอยไก่', 'ผัดซีอิ๊วหมู', 'ผัดหมี่โคราช',
  'หมี่กะทิ', 'ขนมจีนน้ำยา', 'ขนมจีนน้ำพริก', 'สุกี้น้ำ', 'สุกี้แห้ง', 'ชาบู',
  'ข้าวต้มกุ้ง', 'ข้าวต้มปลา', 'โจ๊กหมู', 'ข้าวไข่เจียวหมูสับ', 'ข้าวกะเพราทะเล',
  'ผัดไทยกุ้งสด', 'ก๋วยเตี๋ยวคั่วไก่', 'ข้าวหน้าไก่ทอด', 'ข้าวมันส้มตำไก่ทอด',
  'ราดหน้าหมู', 'ราดหน้าทะเล', 'บะหมี่แห้งหมูแดง', 'ข้าวหมกไก่', 'ข้าวผัดปู',
  'ข้าวผัดอเมริกัน', 'ผัดมาม่าไข่เจียว',
];

// เมนูกับข้าว/กับแกล้ม - สุ่มผสมกันได้อิสระข้ามหมวด (แกง/ต้ม/ผัด/ทอด/ย่าง/ยำ/น้ำพริก) ไม่จำกัดว่าต้องคนละหมวด
const PAIRABLE_NAMES = Object.keys(FOOD_MENU).filter((name) => !ONEPLATE_NAMES.includes(name));

function pickRandomUnique(arr, count) {
  const shuffled = [...arr].sort(() => Math.random() - 0.5);
  return shuffled.slice(0, count);
}

// แคลอรี่ข้าวสวย 1 จาน (ทัพพีมาตรฐาน ~1.5-2 ทัพพี) ที่กินคู่กับกับข้าวหมวด curry/soup/stirfry/fried/grilled/salad/other
const RICE_KCAL = 250;

// สุ่มเมนู: ~30% สุ่มจานเดียวจบ (คำนวณแคลอรี่ของจานนั้น เมนูกลุ่มนี้รวมข้าว/เส้นอยู่แล้ว ไม่บวกข้าวสวยเพิ่ม)
// ~70% สุ่มผสม 2-3 เมนูจากกับข้าวทั้งหมด (ข้ามหมวดได้อิสระ) แล้วรวมแคลอรี่พร้อมข้าวสวยให้ด้วย
function randomFoodSuggestion() {
  const useOnePlate = Math.random() < 0.3;
  if (useOnePlate) {
    const name = pickRandomUnique(ONEPLATE_NAMES, 1)[0];
    return `${name} (~${FOOD_MENU[name]} kcal)`;
  }
  const setSize = Math.random() < 0.5 ? 2 : 3;
  const picks = pickRandomUnique(PAIRABLE_NAMES, setSize);
  const dishTotal = picks.reduce((sum, name) => sum + FOOD_MENU[name], 0);
  const total = dishTotal + RICE_KCAL;
  const lines = picks.map((name) => `${name} (~${FOOD_MENU[name]} kcal)`).join(' + ');
  return `${lines} + ข้าวสวย (~${RICE_KCAL} kcal)\nรวมประมาณ ${total} kcal`;
}

// ==== ดูดวงรายวัน (เปิดไพ่ทาโร่) =====================================================
// ไพ่ทาโร่จริง 22 ใบ (ไพ่ชุดใหญ่ / Major Arcana ทั้งสำรับ) จากสำรับ Rider-Waite-Smith
// (วาดปี 1909 พ้นลิขสิทธิ์แล้ว ใช้ได้ฟรี) รูปโหลดสดจาก Wikimedia Commons ผ่าน
// Special:FilePath (ลิงก์ทางการของ Commons สำหรับฝัง/ลิงก์ตรงไปที่รูป ไม่ต้องรู้ URL เต็ม)
// หมายเหตุ: จากสำรับจริงมี 78 ใบ (22 ใบใหญ่ + 56 ใบเล็ก) แต่ใบเล็กมีชื่อไฟล์ไม่เป็นมาตรฐาน
// เท่าไพ่ใหญ่ จึงใช้แค่ไพ่ใหญ่ทั้ง 22 ใบก่อน (ของจริงครบชุด ไม่ต้องแต่งชื่อไพ่เอง)
// 1 ใบ = 1 ชุดคำตอบ (รูปไพ่ + คำทำนาย 4 บรรทัด + สีมงคล + เลขเด่น 2 ชุด)
// ใช้ "วันที่ (เวลาไทย) + userId" เป็น seed เพื่อให้ดวงของคนคนเดิมคงที่ตลอดทั้งวัน
// (สอดคล้องกับคำว่า "รายวัน") แต่จะเปลี่ยนเป็นใบใหม่เมื่อขึ้นวันถัดไป
const FORTUNE_CARDS = [
  { name: '0. ไพ่คนโง่ (The Fool)', file: 'RWS_Tarot_00_Fool.jpg', message: 'เป็นวันของการเริ่มต้นใหม่และการผจญภัย เปิดใจรับสิ่งที่ยังไม่รู้จัก\nความรัก: มีโอกาสเจอสิ่งใหม่หรือเริ่มความสัมพันธ์แบบไม่มีสคริปต์ ลองเปิดใจดู\nการงาน/เงิน: เหมาะกับการลองอะไรใหม่ๆ แม้จะยังไม่มั่นใจร้อยเปอร์เซ็นต์\nคำแนะนำ: เดินหน้าลุยเลย ไม่ต้องรู้ทุกอย่างก่อนก็ได้ ขนาด GPS ยังพาไปผิดทางบ่อยๆ แล้วก็ยังไปถึงจนได้ 😅', caution: 'บุ่มบ่ามเกินไปอาจทำพลาดเรื่องเล็กๆ เช่นลืมเช็กก่อนส่งงาน ตรวจทานอีกนิดก่อนมั่นใจเต็มร้อย' },
  { name: '1. ไพ่นักมายากล (The Magician)', file: 'RWS_Tarot_01_Magician.jpg', message: 'พลังความคิดสร้างสรรค์พุ่งแรงในวันนี้\nความรัก: กล้าเปิดใจแสดงความรู้สึกออกไป ผลลัพธ์จะดีกว่าที่คิด\nการงาน/เงิน: ไอเดียที่คิดค้างไว้ วันนี้ลงมือทำได้เลย\nคำแนะนำ: ไอเดียมาแล้วอย่าดองไว้ในดราฟต์ รีบเทลงมือทำก่อนมันจะเงียบหายไปเหมือนแชทที่พิมพ์ค้างไว้ 😏', caution: 'พูดจาโอ้อวดเกินจริงอาจทำให้คนรอบข้างไม่เชื่อใจ เก็บไอเดียไว้พิสูจน์ด้วยผลงานดีกว่า' },
  { name: '2. ไพ่นักบวชหญิง (The High Priestess)', file: 'RWS_Tarot_02_High_Priestess.jpg', message: 'วันนี้สัญชาตญาณของคุณแม่นกว่าปกติ ลองเชื่อเสียงในใจดูบ้าง\nความรัก: มีบางอย่างที่ยังไม่พูดออกมาตรงๆ ลองฟังสิ่งที่ไม่ได้พูดด้วยใจ\nการงาน/เงิน: อย่าเพิ่งตัดสินใจจากข้อมูลผิวเผิน ลองศึกษาให้ลึกอีกนิด\nคำแนะนำ: ปิดมือถือ นั่งเงียบๆ สัก 5 นาที (ไม่ใช่แกล้งหลับหนีงานนะ) เดี๋ยวคำตอบจะลอยมาเอง 🔮', caution: 'เก็บความรู้สึกไว้คนเดียวมากไป อาจกลายเป็นเข้าใจผิดกันโดยไม่รู้ตัว' },
  { name: '3. ไพ่จักรพรรดินี (The Empress)', file: 'RWS_Tarot_03_Empress.jpg', message: 'พลังความอุดมสมบูรณ์เต็มที่ในวันนี้\nความรัก: บรรยากาศอบอุ่นละมุน คนมีคู่จะรู้สึกได้รับการดูแลดี\nการงาน/เงิน: งานและเงินมีแนวโน้มไหลลื่นกว่าปกติ\nคำแนะนำ: กินให้อิ่ม นอนให้พอ แล้วค่อยไปรวยทีหลัง ร่างกายดีคือทุนที่ถูกที่สุดแล้วครับ 🍚', caution: 'ตามใจตัวเองเรื่องกินเรื่องใช้จ่ายมากไปหน่อย วันนี้ควบคุมสติไว้บ้าง' },
  { name: '4. ไพ่จักรพรรดิ (The Emperor)', file: 'RWS_Tarot_04_Emperor.jpg', message: 'วันนี้เหมาะกับการใช้เหตุผลและความมั่นคงนำทาง\nความรัก: ควรคุยกันด้วยเหตุผล อย่าตัดสินใจตามอารมณ์ชั่ววูบ\nการงาน/เงิน: เป็นวันที่เหมาะกับการวางแผนระยะยาวและจัดระเบียบ\nคำแนะนำ: วันนี้ขอเป็นเจ้านายตัวเองสักวัน สั่งตัวเองให้ทำตามแผน แล้วอย่าเพิ่งเถียงตัวเองเลย 😤', caution: 'ยึดความคิดตัวเองมากไปอาจทำให้คนรอบข้างอึดอัด ลองประนีประนอมดูบ้าง' },
  { name: '5. ไพ่นักบวช (The Hierophant)', file: 'RWS_Tarot_05_Hierophant.jpg', message: 'วันนี้เหมาะกับการทำตามแบบแผนหรือขอคำแนะนำจากคนที่มีประสบการณ์มากกว่า\nความรัก: การพูดคุยกับผู้ใหญ่หรือคนที่ไว้ใจอาจช่วยให้เห็นภาพชัดขึ้น\nการงาน/เงิน: ทำตามขั้นตอนที่มีอยู่แล้วจะปลอดภัยกว่าการลองของใหม่วันนี้\nคำแนะนำ: ถามรุ่นพี่หรือคนที่เก๋ากว่าดีกว่าไปงมเอง เชื่อเถอะ เขาล้มมาก่อนคุณแล้ว 😂', caution: 'ทำตามคนอื่นจนไม่กล้าคิดต่าง อาจพลาดโอกาสดีๆ ที่ต่างจากกรอบเดิม' },
  { name: '6. ไพ่คู่รัก (The Lovers)', file: 'RWS_Tarot_06_Lovers.jpg', message: 'เรื่องความสัมพันธ์กำลังไปได้สวยในวันนี้\nความรัก: คนโสดมีโอกาสเจอคนถูกใจ คนมีคู่ความหวานจะเพิ่มขึ้น\nการงาน/เงิน: การร่วมงานกับผู้อื่นราบรื่น เจรจาอะไรลงตัวง่าย\nคำแนะนำ: อยากได้ใจใคร ให้ฟังเขาพูดจบก่อนพิมพ์ตอบ อย่ารีบส่ง "ครับ/ค่ะ" ทิ้งไว้ห้วนๆ 💕', caution: 'ตัดสินใจด้วยอารมณ์ล้วนๆ อาจนำไปสู่ความผิดพลาดที่ต้องเสียใจทีหลัง' },
  { name: '7. ไพ่ราชรถ (The Chariot)', file: 'RWS_Tarot_07_Chariot.jpg', message: 'วันนี้พลังใจแข็งแกร่ง มุ่งไปข้างหน้าได้อย่างมั่นใจ\nความรัก: ความสัมพันธ์ที่ตั้งใจฝ่าฟันมาด้วยกัน จะเริ่มเห็นผลลัพธ์ที่ดี\nการงาน/เงิน: เป้าหมายที่วางไว้มีโอกาสสำเร็จได้ถ้าโฟกัสไม่วอกแวก\nคำแนะนำ: ปิดแจ้งเตือนไลน์สัก 1 ชั่วโมง แล้วซัดงานให้จบ รถศึกไม่รอคนเช็กมือถือ 🏎️', caution: 'เร่งรีบเกินไปจนพลาดรายละเอียดสำคัญ ช้าลงอีกนิดจะดีกว่า' },
  { name: '8. ไพ่พลัง (Strength)', file: 'RWS_Tarot_08_Strength.jpg', message: 'ใจแข็งแกร่งพอจะผ่านเรื่องยากไปได้ในวันนี้\nความรัก: กล้าเผชิญหน้ากับปัญหาที่ค้างคาในความสัมพันธ์ได้แล้ว\nการงาน/เงิน: อุปสรรคที่เจอวันนี้ เอาชนะได้ด้วยความอดทน\nคำแนะนำ: ถ้ายังไหวให้สู้ต่ออีกยกนึง ถ้าไม่ไหวจริงๆ ก็กินของหวานสักคำแล้วค่อยสู้ใหม่พรุ่งนี้ 🍰', caution: 'ฝืนตัวเองมากไปจนเหนื่อยเกินพอดี บางทีการพักก็คือความแข็งแกร่งแบบหนึ่ง' },
  { name: '9. ไพ่ฤๅษี (The Hermit)', file: 'RWS_Tarot_09_Hermit.jpg', message: 'เป็นวันที่เหมาะกับการพักและทบทวนตัวเอง\nความรัก: อาจต้องการเวลาส่วนตัวมากกว่าปกติ บอกคนรักให้เข้าใจ\nการงาน/เงิน: ไม่ใช่วันที่เหมาะกับการตัดสินใจเรื่องใหญ่หรือรีบร้อน\nคำแนะนำ: วันนี้ขออนุญาตตัวเองหายตัวไปงีบสัก 20 นาที ไม่ต้องรู้สึกผิด ฤๅษีท่านก็พักเหมือนกัน 😴', caution: 'แยกตัวมากไปอาจทำให้พลาดข่าวสารหรือความช่วยเหลือที่คนอื่นอยากให้' },
  { name: '10. ไพ่กงล้อแห่งโชคชะตา (Wheel of Fortune)', file: 'RWS_Tarot_10_Wheel_of_Fortune.jpg', message: 'จังหวะชีวิตกำลังหมุนเปลี่ยน โชคเข้าข้างในเรื่องที่ไม่คาดคิด\nความรัก: อาจมีจุดเปลี่ยนเล็กๆ ที่ทำให้ความสัมพันธ์พลิกไปในทางดีขึ้น\nการงาน/เงิน: มีโอกาสหรือข่าวดีแทรกเข้ามาแบบไม่ทันตั้งตัว\nคำแนะนำ: แผนพลิกก็อย่าตกใจ ชีวิตเหมือนวงล้อ หมุนไปเดี๋ยวก็เจอเลขสวยเข้าสักวัน 🎡', caution: 'ประมาทเรื่องเงินหรือของมีค่าวันนี้ เผลอนิดเดียวอาจเสียของ' },
  { name: '11. ไพ่ความยุติธรรม (Justice)', file: 'RWS_Tarot_11_Justice.jpg', message: 'วันนี้เกี่ยวข้องกับเรื่องความยุติธรรมและการตัดสินใจที่เป็นธรรม\nความรัก: หากมีเรื่องค้างคาใจกับคนรัก คุยกันตรงๆ จะช่วยได้มาก\nการงาน/เงิน: เรื่องสัญญา เอกสาร มีแนวโน้มคลี่คลายไปในทางที่ดี\nคำแนะนำ: อ่านให้จบก่อนกด "ยอมรับเงื่อนไข" นะครับ อย่าทำเหมือนกดโหลดแอปใหม่ 😅', caution: 'ตัดสินใครเร็วเกินไปโดยยังไม่ฟังทุกด้าน อาจทำให้เข้าใจผิดกัน' },
  { name: '12. ไพ่ชายแขวน (The Hanged Man)', file: 'RWS_Tarot_12_Hanged_Man.jpg', message: 'วันนี้บางอย่างอาจดูเหมือนหยุดชะงัก แต่จริงๆ กำลังให้เวลาคุณมองมุมใหม่\nความรัก: ลองมองสถานการณ์จากมุมของอีกฝ่ายดูบ้าง อาจเข้าใจกันมากขึ้น\nการงาน/เงิน: ถ้าตอนนี้ยังไม่มีคำตอบ การรอดูก่อนอาจดีกว่าการรีบตัดสินใจ\nคำแนะนำ: บางเรื่องปล่อยให้มันค้างไว้ก่อนก็ได้ เหมือนไฟล์โหลดช้าๆ อย่าไปกดรีเฟรชรัวๆ ให้มันงอแง 🙃', caution: 'ผัดวันประกันพรุ่งจนงานค้างสะสม อย่าปล่อยไว้นานเกินไป' },
  { name: '13. ไพ่การเปลี่ยนผ่าน (Death)', file: 'RWS_Tarot_13_Death.jpg', message: 'วันนี้อาจมีบางอย่างที่ต้องปิดฉากลง เพื่อเปิดทางให้สิ่งใหม่เข้ามาแทนที่\nความรัก: การปล่อยสิ่งที่ไม่เหมาะกับคุณ อาจเป็นจุดเริ่มต้นของสิ่งที่ดีกว่า\nการงาน/เงิน: การเปลี่ยนแปลงที่เกิดขึ้นอาจดูน่ากลัวตอนแรก แต่จะพาไปในทางที่ดีขึ้น\nคำแนะนำ: ของเก่าที่ไม่ใช้แล้วก็โละทิ้งได้ ทั้งของในตู้และดราม่าในใจ เคลียร์ที่ว่างไว้รับของใหม่ 🗑️✨', caution: 'ต่อต้านการเปลี่ยนแปลงมากไปอาจทำให้ตัวเองเครียดโดยไม่จำเป็น' },
  { name: '14. ไพ่ความสมดุล (Temperance)', file: 'RWS_Tarot_14_Temperance.jpg', message: 'ความพอดีคือกุญแจสำคัญของวันนี้\nความรัก: อย่าทุ่มเทหรือเรียกร้องมากเกินไป สมดุลคือคำตอบ\nการงาน/เงิน: บริหารเวลาและเงินให้พอดี ไม่ตึงหรือหย่อนเกินไป\nคำแนะนำ: กินหวานได้แต่อย่าลืมกินผัก ทำงานหนักได้แต่อย่าลืมพัก สูตรลับความพอดีมีแค่นี้แหละ 🧉', caution: 'หักโหมทำหลายอย่างพร้อมกันจนสมดุลเสีย เลือกทำทีละเรื่องดีกว่า' },
  { name: '15. ไพ่สิ่งยึดติด (The Devil)', file: 'RWS_Tarot_15_Devil.jpg', message: 'วันนี้อาจมีสิ่งล่อใจหรือพฤติกรรมเดิมๆ ที่ทำให้รู้สึกติดอยู่กับที่ ลองสังเกตดู\nความรัก: ระวังความสัมพันธ์ที่ทำให้รู้สึกอึดอัดหรือพึ่งพากันมากเกินไป\nการงาน/เงิน: ระวังการใช้จ่ายตามอารมณ์หรือการผูกมัดที่ไม่จำเป็นวันนี้\nคำแนะนำ: มือถือ ของหวาน หรือดราม่าในกลุ่มไลน์ วันนี้ลองวางสักอย่างดูก่อน แล้วจะรู้สึกโล่งขึ้นเยอะ 📵', caution: 'หลงกับความสบายชั่วคราว (ของหวาน มือถือ ดราม่า) จนลืมเรื่องสำคัญที่ต้องทำ' },
  { name: '16. ไพ่หอคอย (The Tower)', file: 'RWS_Tarot_16_Tower.jpg', message: 'อาจมีเรื่องที่เปลี่ยนแปลงกะทันหันเข้ามาสั่นสะเทือนแผนเดิมบ้าง\nความรัก: ความจริงบางอย่างอาจถูกเปิดเผยแบบไม่ทันตั้งตัว แต่จะช่วยให้ทุกอย่างชัดเจนขึ้น\nการงาน/เงิน: แผนที่วางไว้อาจต้องพลิกกะทันหัน เตรียมใจปรับตัวไว้บ้าง\nคำแนะนำ: ถ้าวันนี้แผนพัง อย่าเพิ่งดราม่า สูดหายใจลึกๆ แล้วเริ่มก่อร่างใหม่ ตึกถล่มยังสร้างใหม่ได้เลย 🏗️', caution: 'อารมณ์วูบวาบเวลาเจอเรื่องกะทันหัน อย่าตัดสินใจอะไรใหญ่ตอนหัวร้อน' },
  { name: '17. ไพ่ดวงดาว (The Star)', file: 'RWS_Tarot_17_Star.jpg', message: 'ความหวังที่รอคอยมานานกำลังใกล้เป็นจริงแล้ว\nความรัก: ความสัมพันธ์มีแนวโน้มสดใสขึ้น สิ่งที่หวังไว้ค่อยๆ ชัดเจน\nการงาน/เงิน: โอกาสดีๆ เริ่มเผยตัว แม้ยังไม่เห็นผลทันที\nคำแนะนำ: อดทนอีกนิดเดียวจริงๆ เหมือนรอมาม่าสุก อีกแป๊บเดียวก็ได้กินแล้ว 🍜✨', caution: 'หวังมากไปจนผิดหวังง่าย ตั้งความหวังพอดีๆ จะสบายใจกว่า' },
  { name: '18. ไพ่ดวงจันทร์ (The Moon)', file: 'RWS_Tarot_18_Moon.jpg', message: 'วันนี้อารมณ์อาจแกว่งไปมา ควรระวังเรื่องเข้าใจผิดจากการสื่อสาร\nความรัก: อาจมีเรื่องเข้าใจคลาดเคลื่อนกับคนใกล้ตัว พูดให้ชัดเจน\nการงาน/เงิน: ตัวเลขหรือรายละเอียดเล็กๆ อาจผิดพลาดได้ง่าย\nคำแนะนำ: พิมพ์เสร็จอย่าเพิ่งกดส่ง อ่านทวนอีกรอบ กันพลาดแบบ "ครับ" เพี้ยนเป็น "ขอบ" 😂', caution: 'คิดมากไปกับเรื่องที่ยังไม่ชัดเจน อาจทำให้เครียดเกินจริง' },
  { name: '19. ไพ่พระอาทิตย์ (The Sun)', file: 'RWS_Tarot_19_Sun.jpg', message: 'วันนี้พลังบวกเต็มเปี่ยม เรื่องดีๆ มีโอกาสเข้ามาแบบไม่ทันตั้งตัว\nความรัก: คนโสดมีเสน่ห์เพิ่มขึ้น คนมีคู่บรรยากาศจะอบอุ่นสดใส\nการงาน/เงิน: งานที่ทำค้างไว้มีโอกาสคืบหน้าเร็วกว่าที่คิด\nคำแนะนำ: วันนี้ฟ้าเป็นใจ ยิ้มเยอะๆ เข้าไว้ เผื่อแดดจะได้ยิ้มตอบ ☀️😄', caution: 'มั่นใจมากไปจนลืมเช็กรายละเอียด ความประมาทมาพร้อมความสำเร็จได้เหมือนกัน' },
  { name: '20. ไพ่การพิพากษา (Judgement)', file: 'RWS_Tarot_20_Judgement.jpg', message: 'วันนี้เหมาะกับการทบทวนสิ่งที่ผ่านมา แล้วตัดสินใจก้าวต่อไปอย่างมั่นใจ\nความรัก: การให้อภัยตัวเองหรือคนอื่น อาจช่วยเปิดทางให้ความสัมพันธ์ดีขึ้น\nการงาน/เงิน: ผลตอบรับหรือข่าวสำคัญบางอย่างมีแนวโน้มชัดเจนขึ้นในวันนี้\nคำแนะนำ: เลิกเลื่อนไปพรุ่งนี้ได้แล้ว วันนี้แหละฤกษ์ดี กล้าตัดสินใจสักที 🔔', caution: 'ตัดสินใจเร็วเกินไปโดยยังคิดไม่รอบคอบ อาจต้องย้อนกลับมาแก้ทีหลัง' },
  { name: '21. ไพ่โลก (The World)', file: 'RWS_Tarot_21_World.jpg', message: 'งานที่ทำค้างไว้นานมีแนวโน้มสำเร็จลุล่วงในเร็วๆ นี้\nความรัก: ความสัมพันธ์ที่ผ่านอุปสรรคมาด้วยกันจะแน่นแฟ้นขึ้น\nการงาน/เงิน: โปรเจกต์หรือเป้าหมายที่ตั้งไว้ใกล้จะปิดจบสวยๆ\nคำแนะนำ: ใกล้จบแล้ว อย่าเพิ่งปล่อยมือตอนใกล้เส้นชัย เหมือนดูซีรีส์ใกล้จบแล้วมาห้ามดูตอนจบไม่ได้ 🎬', caution: 'ใจร้อนอยากรีบปิดงานจนพลาดรายละเอียดตอนท้าย' },
];

// สีมงคลกับคนที่ควรเลี่ยงคุยด้วยวันนี้ สุ่มแยกอิสระจากไพ่ที่จับได้ (ไม่ผูกติดกับไพ่ใบไหนใบหนึ่งตายตัว)
const COLOR_POOL = [
  'สีทอง', 'สีฟ้าอ่อน', 'สีฟ้าใส', 'สีชมพู', 'สีเหลืองทอง', 'สีแดงเข้ม', 'สีเขียวมรกต', 'สีม่วง',
  'สีขาว', 'สีน้ำตาล', 'สีแดงสด', 'สีฟ้าเทอร์ควอยซ์', 'สีส้มสด', 'สีที่ชอบที่สุด', 'สีเหลืองสด',
  'สีเขียวเข้ม', 'สีเทาเข้ม', 'สีฟ้าพาสเทล', 'สีเขียวอ่อน', 'สีฟ้าคราม', 'สีเงิน', 'สีฟ้าเข้ม',
  'สีม่วงเข้ม', 'สีขาวนวล', 'สีเทาอมฟ้า', 'สีชมพูพีช', 'สีน้ำเงินเข้ม', 'สีน้ำตาลทอง', 'สีเทาเงิน',
  'สีส้มแดง', 'สีขาวทอง',
];

// ลิสต์คนในออฟฟิศที่บอทจะสุ่มบอกว่า "วันนี้เลี่ยงคุยด้วยดีกว่า" (มุกฮาๆ ในกลุ่ม ไม่ใช่เรื่องจริงจัง)
const OFFICE_AVOID_PEOPLE = [
  'AB + NT', 'ทราย', 'Auditor ผญ.', 'Auditor ผช.', 'นุ่น', 'เนย์', 'น้ำ', 'เบญ', 'อ๊อฟ', 'นีน',
  'กิ๊ก', 'ปลา', 'จอย', 'น้ำตาล', 'PK', 'พี่หนิง',
];

// สร้าง URL รูปไพ่จาก Wikimedia Commons ผ่าน Special:FilePath (รองรับพารามิเตอร์ ?width=
// เพื่อ redirect ไปที่รูปย่อขนาดนั้นโดยอัตโนมัติ ไม่ต้องรู้ URL เต็มที่มี hash โฟลเดอร์)
function tarotImageUrl(filename, width) {
  return `https://commons.wikimedia.org/wiki/Special:FilePath/${encodeURIComponent(filename)}?width=${width}`;
}

// hash เลข seed จากข้อความ -> ตัวเลข 0-1 (deterministic ไม่ใช้ Math.random เพื่อให้ผลลัพธ์คงที่ตลอดวัน
// ใช้ md5 เพื่อให้การกระจายผลลัพธ์แต่ละแบบสม่ำเสมอ ไม่เอนไปทางใดทางหนึ่งบ่อยผิดปกติ)
function seededRandom01(seedStr) {
  const hash = crypto.createHash('md5').update(seedStr).digest();
  const n = hash.readUInt32BE(0);
  return n / 4294967296;
}

function seededPick(seedStr, arr) {
  const idx = Math.floor(seededRandom01(seedStr) * arr.length) % arr.length;
  return arr[idx];
}

function seededNumber3(seedStr) {
  const n = Math.floor(seededRandom01(seedStr) * 1000);
  return String(n).padStart(3, '0');
}

// "วันนี้" อิงเวลาไทย (UTC+7) ไม่ใช่เวลาของ server เพื่อให้ดวงเปลี่ยนตอนเที่ยงคืนเมืองไทยจริงๆ
function bangkokNow() {
  return new Date(Date.now() + 7 * 60 * 60 * 1000);
}
function todayKeyBangkok() {
  const bkk = bangkokNow();
  return `${bkk.getUTCFullYear()}-${bkk.getUTCMonth() + 1}-${bkk.getUTCDate()}`;
}
// เลขวัน (นับจาก epoch) ตามเวลาไทย ใช้ทำเลขคณิตเรื่องวัน (เช่น ข้ามไปกี่วันแล้ว) ได้ง่ายกว่า string
function bangkokDayIndex() {
  return Math.floor(bangkokNow().getTime() / 86400000);
}

// เลือกไพ่ของ "วันนี้" ให้ userId คนนี้ (deterministic ตามวันที่+userId เหมือนเดิม)
// แยกออกมาเป็นฟังก์ชันของตัวเอง เพราะต้องใช้ทั้งตอนเปิดไพ่จริง และตอนแจ้งว่าเปิดไปแล้ววันนี้
function pickTodayCard(userId) {
  const seed = `${todayKeyBangkok()}:${userId}`;
  const idx = Math.floor(seededRandom01(seed) * FORTUNE_CARDS.length) % FORTUNE_CARDS.length;
  return FORTUNE_CARDS[idx];
}

// คืนค่าเป็น "อาเรย์ของ LINE message object" (รูปไพ่ + ข้อความคำทำนาย) ส่งพร้อมกันได้เลยใน
// replyMessage เดียว (LINE รองรับส่งได้สูงสุด 5 ข้อความต่อการ reply ครั้งเดียว)
function buildDailyFortuneMessages(userId) {
  const card = pickTodayCard(userId);
  const baseSeed = `${todayKeyBangkok()}:${userId}`;
  // สุ่มสี/เลข/คนที่ควรเลี่ยง แยกอิสระจากไพ่ที่จับได้ ไม่ใช่ค่าตายตัวของไพ่ใบนั้นอีกต่อไป
  const color = seededPick(`${baseSeed}:color`, COLOR_POOL);
  const num1 = seededNumber3(`${baseSeed}:num1`);
  let num2 = seededNumber3(`${baseSeed}:num2`);
  if (num2 === num1) num2 = seededNumber3(`${baseSeed}:num2b`);
  const avoidPerson = seededPick(`${baseSeed}:avoid`, OFFICE_AVOID_PEOPLE);
  const text = [
    '🔮 ดวงประจำวันนี้ของคุณ',
    `${card.name}`,
    '',
    card.message,
    '',
    `⚠️ ข้อควรระวังวันนี้: ${card.caution}`,
    '',
    `🎨 สีมงคลวันนี้: ${color}`,
    `🔢 เลขเด่นวันนี้: ${num1} , ${num2}`,
    `🙊 วันนี้เลี่ยงคุยเรื่องเครียดกับ: ${avoidPerson}`,
  ].join('\n');
  return [
    {
      type: 'image',
      originalContentUrl: tarotImageUrl(card.file, 1024),
      previewImageUrl: tarotImageUrl(card.file, 240),
    },
    { type: 'text', text },
  ];
}

// --- LINE webhook -----------------------------------------------------
// line.middleware verifies the signature and parses the body itself,
// so no extra body-parser is used on this route.
app.post(
  '/webhook',
  line.middleware(config),
  (req, res) => {
    res.status(200).end(); // ack LINE immediately
    const events = req.body.events || [];
    events.forEach((event) => {
      handleEvent(event).catch((err) => console.error('handleEvent error:', err));
    });
  },
  // error handler เฉพาะเส้นทางนี้ - จับ error จาก line.middleware เช่น signature ไม่ตรง
  (err, req, res, next) => {
    console.error('[webhook error]', err.name, err.message);
    res.status(err.statusCode || 500).end();
  }
);

async function reply(event, text) {
  try {
    await client.replyMessage(event.replyToken, { type: 'text', text });
  } catch (e) {
    console.error('reply failed:', e.originalError?.response?.data || e.message);
  }
}

// ใช้ตอนต้องส่งหลายข้อความในครั้งเดียว (เช่น รูปไพ่ + ข้อความคำทำนาย) LINE รองรับสูงสุด 5 ข้อความ/ครั้ง
async function replyMessages(event, messages) {
  try {
    await client.replyMessage(event.replyToken, messages);
  } catch (e) {
    console.error('reply (multi) failed:', e.originalError?.response?.data || e.message);
  }
}

// ==== สุ่มคำตอบคำถาม "/ใคร..." ====================================================
// เก็บรายชื่อคนที่เคยพิมพ์ในแต่ละกลุ่มไว้ (แยกตามกลุ่ม) เพื่อใช้สุ่มตอบคำถาม /ใคร
// ไม่ได้ใช้ LINE API ดึงสมาชิกทั้งกลุ่ม เพราะ API นั้น (getGroupMembersIds) ใช้ได้เฉพาะ
// Official Account ที่ผ่านการ verify/premium แล้วเท่านั้น บอททั่วไปเรียกจะโดน 403
async function ensureGroupMemberCached(groupId, userId) {
  if (!db.groupMembers) db.groupMembers = {};
  if (!db.groupMembers[groupId]) db.groupMembers[groupId] = {};
  if (db.groupMembers[groupId][userId]) return db.groupMembers[groupId][userId];
  let name;
  try {
    const profile = await client.getGroupMemberProfile(groupId, userId);
    name = profile.displayName;
  } catch (e) {
    name = 'สมาชิก' + userId.slice(-4);
  }
  db.groupMembers[groupId][userId] = name;
  saveDB(db);
  return name;
}

function pickRandomFrom(arr, n) {
  const shuffled = [...arr].sort(() => Math.random() - 0.5);
  return shuffled.slice(0, Math.max(0, n));
}

async function handleWhoQuestion(event, groupId, userId, questionText) {
  const members = (db.groupMembers && db.groupMembers[groupId]) || {};
  const entries = Object.entries(members); // [ [userId, name], ... ]

  if (entries.length === 0) {
    await reply(
      event,
      'ยังไม่รู้จักใครในกลุ่มนี้พอจะสุ่มได้ครับ ให้สมาชิกลองพิมพ์อะไรในกลุ่มสักครั้งก่อน แล้วค่อยถามใหม่'
    );
    return;
  }

  const askerName = await ensureGroupMemberCached(groupId, userId);
  const modes = ['one', 'two', 'few', 'all', 'asker'];
  const mode = modes[Math.floor(Math.random() * modes.length)];

  let answerText;
  if (mode === 'asker') {
    answerText = `${askerName} ไง (ถามเอง โดนเองสิ 😏)`;
  } else if (mode === 'all') {
    answerText = 'ทุกคนในกลุ่มเลย! 🎉';
  } else {
    let n;
    if (mode === 'one') n = 1;
    else if (mode === 'two') n = 2;
    else n = 3; // คน 3 คน
    n = Math.min(n, entries.length);
    const picks = pickRandomFrom(entries, n).map(([, name]) => name);
    answerText = picks.join(', ');
  }

  await reply(event, `🔮 ${questionText}\nคำตอบคือ: ${answerText}`);
}
// ==== จบส่วนสุ่มคำตอบคำถาม /ใคร ====================================================

// ==== แท็กชื่อบอทแล้วถามคำถามทั่วไป (คุยกับ Gemini API) ============================
// เช็คว่าข้อความนี้ "แท็ก" บอทตัวเองหรือไม่ โดยดูจาก event.message.mention ที่ LINE ส่งมาให้
// (ฟีเจอร์ mention ของ LINE เอง ไม่ใช่การเทียบชื่อบอทกับข้อความ จึงแม่นยำกว่า)
// คืนค่าเป็นคำถามที่ตัดส่วน "@ชื่อบอท " ออกแล้ว หรือ null ถ้าข้อความนี้ไม่ได้แท็กบอท
function extractQuestionIfMentioned(event) {
  const mention = event.message.mention;
  if (!mention || !Array.isArray(mention.mentionees)) return null;
  const mentionsSelf = mention.mentionees.some((m) => m.isSelf);
  if (!mentionsSelf) return null;

  const text = event.message.text;
  // ตัดข้อความส่วนที่เป็น mention (เช่น "@ชื่อบอท ") ออกทั้งหมด เหลือแต่คำถามจริงๆ
  let question = '';
  let lastIndex = 0;
  const sorted = [...mention.mentionees].sort((a, b) => a.index - b.index);
  for (const m of sorted) {
    question += text.slice(lastIndex, m.index);
    lastIndex = m.index + m.length;
  }
  question += text.slice(lastIndex);
  return question.trim();
}

// เรียก Gemini API (free tier ของ Google) ให้ตอบคำถามทั่วไปแบบกระชับ เป็นภาษาไทย
async function askGemini(question) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent?key=${GEMINI_API_KEY}`;
  const systemInstruction =
    'คุณเป็นผู้ช่วยตอบคำถามในกลุ่มไลน์ของออฟฟิศ ตอบให้กระชับ ตรงประเด็น ไม่ต้องยาวเกินจำเป็น ' +
    'ใช้ภาษาไทยเป็นหลัก เว้นแต่ผู้ถามพิมพ์ถามเป็นภาษาอื่น ให้ตอบเป็นภาษานั้นแทน';
  const body = {
    system_instruction: { parts: [{ text: systemInstruction }] },
    contents: [{ role: 'user', parts: [{ text: question }] }],
  };
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const errText = await res.text().catch(() => '');
    throw new Error(`Gemini API ${res.status}: ${errText.slice(0, 300)}`);
  }
  const data = await res.json();
  const answer = data?.candidates?.[0]?.content?.parts?.map((p) => p.text).join('') || '';
  return answer.trim();
}
// ==== จบส่วนถามคำถามทั่วไป ==========================================================

// ==== เกมเลี้ยงสัตว์เลี้ยง (Tamagotchi) ==============================================
// แต่ละคนมีสัตว์เลี้ยงของตัวเอง (เก็บใน db.pets[userId]) ต้องดูแล (ให้อาหาร) วันละ 2 ช่วงตามเวลาไทยจริง
// เช้า 06:00-12:00 น. / บ่าย 12:01-18:00 น. (นอกช่วงนี้ป้อนอาหารไม่ได้) สะสม "พลังชีวิต" (lifeForce)
// เป็นแต้มสะสมไม่มีเพดาน ใช้เป็นตัวตัดสินระยะการเติบโต — กด /ให้อาหาร
// แต่ละรอบ (เช้าหรือบ่าย) ได้ +25 แต้มทันที ดูแลครบทั้งวันเท่ากับ +50 แต้ม ไม่มีโทษ ดูแลครึ่งวัน
// (แค่เช้าหรือแค่บ่าย) ก็ได้ +25 เฉยๆ ไม่มีโทษเช่นกัน แต่ถ้าขาดดูแลไปเลยทั้งวัน (ไม่ป้อนอาหารแม้แต่รอบเดียว)
// จะโดนหักท้ายวัน -20 แต้ม (พื้นที่ 0 ไม่ติดลบ) ขาดดูแลทั้งวันติดกัน
// 2 วันขึ้นไปจะป่วยทันที ป้อนยาแล้วหายได้เสมอ (ไม่มีตายถาวร) ระยะการเติบโตไม่มีวันถอยกลับแม้แต้มจะลดลง
// ภายหลังจากถูกปล่อยละเลย นอกจากนี้บอทจะสุ่ม "เหตุการณ์พิเศษ" push เข้ากลุ่มวันละ 4 รอบ (เวลาสุ่มระหว่าง
// 08:30-17:00 น.) แบ่งครึ่งๆ สุ่มลำดับ เป็น 2 รอบเกม "โจทย์เลขบวกลบ" (เลข 2-3 หลัก พิมพ์คำตอบให้ถูก)
// กับ 2 รอบเกม "พิมพ์ตามให้ตรงเป๊ะ" เปิดให้ทุกคนในกลุ่มแข่งกัน แต่ละรอบจำกัดเวลา 1 นาที
// 5 คนแรกที่พิมพ์ถูกได้ +30 แต้ม คนที่เหลือได้ +10 แต้ม แล้วบอทจะแจ้งหมดเวลาเมื่อครบ 1 นาที
// รวมใช้ประมาณ 240 ข้อความ push/เดือน (แพ็กเกจฟรีของไทยมี 300 ข้อความ/เดือน ยังพอมีเผื่อฟีเจอร์อื่น)

const PET_SPECIES = [
  '🐶 หมาน้อย', '🐱 แมวเหมียว', '🐰 กระต่าย', '🐹 แฮมสเตอร์', '🐼 แพนด้า',
  '🐧 เพนกวิน', '🦊 จิ้งจอก', '🐢 เต่า', '🐥 ลูกเจี๊ยบ', '🐸 กบ',
];

// จับคู่ชนิดสัตว์ (ข้อความเต็มใน PET_SPECIES) กับชื่อโฟลเดอร์รูปภาพใน public/pet-images/<slug>/<stage>.png
const PET_SPECIES_SLUG = {
  '🐶 หมาน้อย': 'dog',
  '🐱 แมวเหมียว': 'cat',
  '🐰 กระต่าย': 'rabbit',
  '🐹 แฮมสเตอร์': 'hamster',
  '🐼 แพนด้า': 'panda',
  '🐧 เพนกวิน': 'penguin',
  '🦊 จิ้งจอก': 'fox',
  '🐢 เต่า': 'turtle',
  '🐥 ลูกเจี๊ยบ': 'chick',
  '🐸 กบ': 'frog',
};

// สร้าง URL รูปสัตว์เลี้ยงตามชนิด+ระยะการโตปัจจุบัน (คืนค่า null ถ้ายังไม่ได้ตั้งค่า PUBLIC_BASE_URL)
function petImageUrl(pet) {
  if (!PUBLIC_BASE_URL) return null;
  const slug = PET_SPECIES_SLUG[pet.species];
  if (!slug) return null;
  return `${PUBLIC_BASE_URL}/pet-images/${slug}/${pet.stage}.png`;
}

// สร้าง LINE image message จากสัตว์เลี้ยง (คืนค่า null ถ้าไม่มีรูปให้ส่ง เพื่อให้ผู้เรียกข้ามได้ง่ายๆ)
function petImageMessage(pet) {
  const url = petImageUrl(pet);
  if (!url) return null;
  return { type: 'image', originalContentUrl: url, previewImageUrl: url };
}

// ลำดับระยะการเติบโต (เรียงจากน้อยไปมาก) ใช้เทียบอันดับเพื่อการันตีว่าสเตจไม่มีวันถอยกลับ
const PET_STAGE_ORDER = ['egg', 'baby', 'child', 'teen', 'adult', 'breed', 'old', 'legend'];

const PET_STAGE_LABEL = {
  egg: 'ไข่ลึกลับ 🥚',
  baby: 'ลูก (Baby) 🐣',
  child: 'เด็ก (Child) 🌱',
  teen: 'วัยรุ่น (Teen) 🌿',
  adult: 'โตเต็มวัย (Adult) 🎊',
  breed: 'พร้อมสืบพันธุ์ (Breeding) 💕',
  old: 'แก่จัด (Old) 👴',
  legend: 'ตำนาน (Legend) 👑',
};

// เกณฑ์แต้มพลังชีวิตสะสม (lifeForce) ต่อระยะการเติบโต
function petStageFromLifeForce(points) {
  if (points >= 15000) return 'legend';
  if (points >= 7500) return 'old';
  if (points >= 4000) return 'breed';
  if (points >= 2000) return 'adult';
  if (points >= 1000) return 'teen';
  if (points >= 400) return 'child';
  if (points >= 50) return 'baby';
  return 'egg';
}

// อัปเดต pet.stage เฉพาะตอนที่ระยะใหม่ "สูงกว่า" ระยะเดิมเท่านั้น (การันตีไม่มีวันถอยกลับ
// แม้ lifeForce จะลดลงภายหลังจากถูกปล่อยละเลย) คืนค่า true ถ้ามีการอัปเกรดระยะจริง เพื่อให้ผู้เรียก
// รู้ว่าควรส่งรูป/ประกาศฉลองการโตขึ้นหรือไม่
function advancePetStageIfHigher(pet) {
  const candidate = petStageFromLifeForce(pet.lifeForce);
  if (PET_STAGE_ORDER.indexOf(candidate) > PET_STAGE_ORDER.indexOf(pet.stage)) {
    pet.stage = candidate;
    return true;
  }
  return false;
}

function adoptPet(name) {
  return {
    name,
    species: PET_SPECIES[Math.floor(Math.random() * PET_SPECIES.length)],
    stage: 'egg',
    lifeForce: 0,
    sick: false,
    missedDaysInRow: 0,
    caredToday: { morning: false, afternoon: false },
    currentDayIndex: bangkokDayIndex(),
    adoptedAt: Date.now(),
  };
}

// ปิดวันของสัตว์เลี้ยง 1 วัน ตามผลการดูแลของวันนั้น (caredToday ก่อนถูกรีเซ็ต)
// หมายเหตุ: แต้ม +25 ต่อรอบ (เช้า/บ่าย) ให้ทันทีตอนกด /ให้อาหาร แล้ว (ดู feedPet ด้านล่าง) ตรงนี้จัดการ
// เฉพาะ "โทษ" ของการดูแลไม่ครบวันเท่านั้น (ดูแลครบทั้งวันจะไม่มีโทษ เพราะได้แต้มครบ 25+25=50 ไปแล้ว)
function applyPetDailyRollover(pet) {
  const { morning, afternoon } = pet.caredToday;
  if (morning && afternoon) {
    pet.missedDaysInRow = 0;
  } else if (morning || afternoon) {
    // ดูแลแค่ครึ่งวัน ยังได้ +25 จากตอนกด /ให้อาหาร ไปแล้ว ไม่มีโทษเพิ่ม แค่รีเซ็ตตัวนับวันขาดเพราะทำอะไรบ้างแล้ว
    pet.missedDaysInRow = 0;
  } else {
    // ขาดดูแลไปเลยทั้งวัน (ไม่ป้อนอาหารแม้แต่รอบเดียว) ถึงจะโดนหักพลังชีวิต
    pet.lifeForce = Math.max(0, pet.lifeForce - 20);
    pet.missedDaysInRow += 1;
  }
  if (pet.missedDaysInRow >= 2) pet.sick = true;
  advancePetStageIfHigher(pet);
  pet.currentDayIndex += 1;
  pet.caredToday = { morning: false, afternoon: false };
}

// ไล่ปิดวันที่ค้างอยู่ทั้งหมดให้ทันวันปัจจุบัน (เผื่อ server ปิดไปหลายวัน หรือไม่มีใครพิมพ์อะไรเลย)
function advancePetDays(pet) {
  const today = bangkokDayIndex();
  let guard = 0;
  while (pet.currentDayIndex < today && guard < 3650) {
    applyPetDailyRollover(pet);
    guard += 1;
  }
}

// รอบ "เช้า" คือ 06:00-12:00 น. รอบ "บ่าย" คือ 12:01-18:00 น. (เวลาไทย) นอกช่วงนี้ (18:01-05:59) ป้อนอาหารไม่ได้
function petSessionNow() {
  const bkk = bangkokNow();
  const hour = bkk.getUTCHours() + bkk.getUTCMinutes() / 60;
  if (hour >= 6 && hour < 12) return 'morning';
  if (hour >= 12 && hour < 18) return 'afternoon';
  return null;
}

function feedPet(pet) {
  advancePetDays(pet);
  const session = petSessionNow();
  if (!session) {
    return { ok: false, reason: 'outside_hours' };
  }
  if (pet.caredToday.morning && pet.caredToday.afternoon) {
    return { ok: false, reason: 'already_fed_today' };
  }
  if (pet.caredToday[session]) {
    return { ok: false, reason: 'already_fed_this_session', session };
  }
  pet.caredToday[session] = true;
  pet.lifeForce = Math.max(0, pet.lifeForce + 25); // กดป้อนอาหาร ได้ +25 คะแนนทันที ง่ายๆ/รอบ (เช้าหรือบ่าย)
  const leveledUp = advancePetStageIfHigher(pet);
  return { ok: true, session, leveledUp };
}

function givePetMedicine(pet) {
  advancePetDays(pet);
  if (!pet.sick) return { ok: false, reason: 'not_sick' };
  pet.sick = false;
  pet.missedDaysInRow = 0;
  return { ok: true };
}

function formatPetStatus(pet) {
  const careTodayText =
    pet.caredToday.morning || pet.caredToday.afternoon
      ? `วันนี้: ${pet.caredToday.morning ? '✅ เช้า' : '⬜ เช้า'} ${pet.caredToday.afternoon ? '✅ บ่าย' : '⬜ บ่าย'}`
      : 'วันนี้ยังไม่ได้ดูแลเลย พิมพ์ /ให้อาหาร ได้เลยครับ';
  return [
    `${pet.species} "${pet.name}"`,
    `ระยะการเติบโต: ${PET_STAGE_LABEL[pet.stage]}`,
    `พลังชีวิตสะสม: ${pet.lifeForce} คะแนน`,
    pet.sick ? '🤒 กำลังป่วยอยู่ พิมพ์ /ป้อนยา ด่วน!' : '💚 สุขภาพแข็งแรงดี',
    careTodayText,
  ].join('\n');
}

// สุ่มเวลาช่วงหนึ่งของวัน (เวลาไทย) สำหรับ dayIndex ที่กำหนด ระหว่าง startHour-endHour (เป็นทศนิยมชั่วโมงได้)
function pickRandomTimeInWindow(dayIndex, startHour, endHour) {
  const bangkokMidnightUtcMs = dayIndex * 86400000 - 7 * 60 * 60 * 1000;
  const startMs = bangkokMidnightUtcMs + startHour * 60 * 60 * 1000;
  const endMs = bangkokMidnightUtcMs + endHour * 60 * 60 * 1000;
  return startMs + Math.random() * (endMs - startMs);
}

// ---- เหตุการณ์พิเศษ (สุ่มวันละ 4 รอบ ครึ่งหนึ่งเป็นโจทย์เลขบวกลบ ครึ่งหนึ่งเป็น "พิมพ์ตามให้ตรงเป๊ะ" เปิดให้ทุกคนในกลุ่มแข่งกัน) ----
const PET_EVENT_WINDOW_START_HOUR = 8.5; // 08:30 น.
const PET_EVENT_WINDOW_END_HOUR = 17; // 17:00 น.
const PET_EVENT_ROUNDS_PER_DAY = 4;
const PET_CHALLENGE_ROUND_MS = 60 * 1000; // แต่ละรอบเปิดให้แข่ง 1 นาที พิมพ์หลังจากนี้ไม่ได้คะแนน
const PET_CHALLENGE_TOP_SLOTS = 5; // 5 คนแรกที่พิมพ์ถูก
const PET_CHALLENGE_TOP_POINTS = 30; // ...ได้คนละ 30 คะแนน
const PET_CHALLENGE_REST_POINTS = 10; // คนที่เหลือ (ถ้าพิมพ์ถูกทันภายใน 1 นาที) ได้คนละ 10 คะแนน

// เก็บเฉพาะข้อความ (ไม่ผูกกับคำสั่งใดๆ แล้ว) แบ่งหมวดไว้แค่ให้เลือกโทนข้อความตอนสุ่มเฉยๆ
const PET_EVENT_HOOKS = {
  feed_extra: [
    'หิวจัง ขอเพิ่มอีกคำ', 'ท้องร้องจ๊อกๆ ป้อนหน่อยสิ', 'ขอขนมเพิ่ม อีกนิดนะ',
    'แอบหิว มาป้อนหน่อยได้ไหม', 'ท้องยังไม่อิ่มเลย ป้อนอีกที', 'อยากกินของอร่อย เพิ่มอีก',
    'ขอข้าวเพิ่ม อีกจานนะ', 'หิวอีกแล้ว มาป้อนหน่อย', 'ขอของว่าง เพิ่มหน่อยจ้า',
    'พลังงานหมด ขอเติมหน่อย', 'อยากกินอีกคำ ด่วนเลย', 'ท้องกิ๋วๆ ป้อนหน่อยนะ',
    'หิวสุดๆ ช่วยป้อนที',
  ],
  walk: [
    'อยากออกไปเดินเล่น จัง', 'เบื่อจัง พาไปเดินหน่อย', 'ขาคัน อยากวิ่งเล่น',
    'อยากสูดอากาศ ข้างนอกบ้าง', 'พาไปเที่ยว หน่อยได้ไหม', 'อยากยืดเส้น ยืดสาย',
    'นั่งเฉยมานาน พาไปเดิน', 'อยากไปดู โลกกว้าง', 'ขอออกกำลังกาย หน่อยนะ',
    'เหงาแล้ว พาไปเที่ยวสิ', 'อยากไปเล่น ข้างนอกบ้าง', 'พาไปดมกลิ่นหญ้า หน่อย',
    'คิดถึงสวน พาไปเดินที',
  ],
  cuddle: [
    'อยากให้อุ้ม จังเลย', 'ขอกอดหน่อย ได้ไหม', 'คิดถึงจัง มาอุ้มหน่อย',
    'อยากอ้อน เจ้าของหน่อย', 'ขอความรัก หน่อยนะ', 'อยากนอนตัก เจ้าของ',
    'มาหอมแก้ม หน่อยสิ', 'อยากให้ลูบหัว หน่อย', 'เหงาๆ อยากให้กอด',
    'ขออ้อน สักครู่นะ', 'อยากให้เอาใจ หน่อย', 'มากอดกัน หน่อยไหม',
  ],
  sick: [
    'รู้สึกไม่ค่อยสบาย เลย', 'เวียนหัว ขอยาหน่อย', 'ท้องไส้ปั่นป่วน จัง',
    'ตัวร้อนๆ เหมือนจะไม่สบาย', 'ปวดหัวนิดๆ ขอยาที', 'รู้สึกอ่อนเพลีย มาก',
    'ไม่ค่อยมีแรง เลย', 'แอบป่วย นิดหน่อยนะ', 'รู้สึกซึมๆ ไม่สดใส',
    'ขอยาหน่อย ไม่ค่อยไหว', 'ตัวโย้เย้ ไม่ค่อยดี', 'เหมือนจะไข้ขึ้น นะ',
  ],
};

function pickPetEventTimes(dayIndex) {
  const sliceLen = (PET_EVENT_WINDOW_END_HOUR - PET_EVENT_WINDOW_START_HOUR) / PET_EVENT_ROUNDS_PER_DAY;
  const times = [];
  for (let i = 0; i < PET_EVENT_ROUNDS_PER_DAY; i++) {
    const startHour = PET_EVENT_WINDOW_START_HOUR + i * sliceLen;
    times.push(pickRandomTimeInWindow(dayIndex, startHour, startHour + sliceLen));
  }
  return times;
}

function randInt(min, max) {
  // สุ่มจำนวนเต็ม รวมค่า min และ max ทั้งคู่
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// สุ่มว่าแต่ละรอบใน 1 วัน (มี PET_EVENT_ROUNDS_PER_DAY รอบ) จะเป็นเกมชนิดไหน แบ่งครึ่งๆ ระหว่าง
// "โจทย์เลขบวกลบ" กับ "พิมพ์ตามให้ตรงเป๊ะ" (ถ้าจำนวนรอบเป็นเลขคี่ ที่เหลือจะเป็นพิมพ์ตามให้ตรงเป๊ะ)
// แล้วสลับลำดับให้สุ่มว่าจะเจอโจทย์เลขหรือพิมพ์ตามก่อน ไม่ตายตัว
function pickDailyRoundKinds(n) {
  const mathCount = Math.floor(n / 2);
  const kinds = [];
  for (let i = 0; i < n; i++) kinds.push(i < mathCount ? 'math' : 'phrase');
  for (let i = kinds.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [kinds[i], kinds[j]] = [kinds[j], kinds[i]];
  }
  return kinds;
}

// สุ่มโจทย์เลขบวก/ลบ เลข 2-3 หลักทั้งคู่ (10-999) ลบแล้วรับประกันไม่ติดลบ (สลับให้เลขตัวใหญ่ลบก่อนเสมอ)
function generateMathChallenge() {
  const digitsA = randInt(2, 3);
  const digitsB = randInt(2, 3);
  const a = randInt(Math.pow(10, digitsA - 1), Math.pow(10, digitsA) - 1);
  const b = randInt(Math.pow(10, digitsB - 1), Math.pow(10, digitsB) - 1);
  const isAdd = Math.random() < 0.5;
  if (isAdd) {
    return { promptText: `${a} + ${b} = เท่าไหร่?`, answerText: String(a + b) };
  }
  const big = Math.max(a, b);
  const small = Math.min(a, b);
  return { promptText: `${big} - ${small} = เท่าไหร่?`, answerText: String(big - small) };
}

// เริ่มเหตุการณ์พิเศษรอบหนึ่ง (roundIndex = รอบที่เท่าไหร่ของวันนี้ ใช้เลือกชนิดเกมจาก db.petSchedule.kinds)
// เป็นได้ 2 แบบ: "math" (โจทย์เลขบวกลบ พิมพ์คำตอบให้ถูก) หรือ "phrase" (สุ่มข้อความ พิมพ์ตามให้ตรงเป๊ะ)
// เงื่อนไขรับรางวัลเหมือนกันทั้ง 2 แบบ (จำกัดเวลา 1 นาที เปิดให้ได้หลายคน)
async function triggerPetEvent(roundIndex) {
  if (!db.groupId || !db.pets) return;
  if (Object.keys(db.pets).length === 0) return; // ยังไม่มีใครเลี้ยงสัตว์เลี้ยงเลยในกลุ่มนี้ ไม่ต้องส่ง
  const roundKind =
    (db.petSchedule && Array.isArray(db.petSchedule.kinds) && db.petSchedule.kinds[roundIndex]) || 'phrase';

  let text;
  if (roundKind === 'math') {
    const { promptText, answerText } = generateMathChallenge();
    db.petChallenge = { hookText: answerText, kind: 'math', promptText, startedAt: Date.now(), winners: [], timeUpAnnounced: false };
    text = [
      '🔢 โจทย์เลขด่วน! ใครคิดเลขไวที่สุด...',
      promptText,
      `พิมพ์คำตอบ (ตัวเลขล้วนๆ) ให้ถูกเป็นคนแรกๆ ภายใน 1 นาที! ${PET_CHALLENGE_TOP_SLOTS} คนแรกได้ +${PET_CHALLENGE_TOP_POINTS} คะแนน คนที่เหลือได้ +${PET_CHALLENGE_REST_POINTS} คะแนน`,
      '(ต้องมีสัตว์เลี้ยงของตัวเองก่อนถึงจะรับรางวัลได้ — หมดเวลา 1 นาทีแล้วพิมพ์ถูกจะไม่ได้คะแนนนะ)',
    ].join('\n');
  } else {
    const kinds = Object.keys(PET_EVENT_HOOKS);
    const kind = kinds[Math.floor(Math.random() * kinds.length)];
    const hooks = PET_EVENT_HOOKS[kind];
    const hookText = hooks[Math.floor(Math.random() * hooks.length)];
    db.petChallenge = { hookText, kind, promptText: hookText, startedAt: Date.now(), winners: [], timeUpAnnounced: false };
    text = [
      '🎲 เหตุการณ์พิเศษ! มีเสียงจากสัตว์เลี้ยงในกลุ่มดังขึ้นว่า...',
      `"${hookText}"`,
      `ใครพิมพ์ข้อความนี้ตามให้ตรงเป๊ะได้ก่อน ภายใน 1 นาที! ${PET_CHALLENGE_TOP_SLOTS} คนแรกได้ +${PET_CHALLENGE_TOP_POINTS} คะแนน คนที่เหลือได้ +${PET_CHALLENGE_REST_POINTS} คะแนน`,
      '(ต้องมีสัตว์เลี้ยงของตัวเองก่อนถึงจะรับรางวัลได้ — หมดเวลา 1 นาทีแล้วพิมพ์ถูกจะไม่ได้คะแนนนะ)',
    ].join('\n');
  }
  saveDB(db);
  try {
    await client.pushMessage(db.groupId, { type: 'text', text });
  } catch (e) {
    console.error('pet event push failed:', e.originalError?.response?.data || e.message);
  }
}

// เรียกทุกข้อความที่เข้ามาในกลุ่ม เช็คว่าตรงกับเหตุการณ์พิเศษที่ค้างอยู่แบบตรงเป๊ะหรือไม่ (ภายใน 1 นาที)
// คืนค่า null ถ้าไม่ตรง/ไม่มีเหตุการณ์/หมดเวลาแล้ว/คนนี้ได้รางวัลรอบนี้ไปแล้ว, { noPet: true } ถ้าตรงแต่คนพิมพ์
// ยังไม่มีสัตว์เลี้ยง (กรณีนี้ไม่กินสิทธิ์ ให้คนอื่นแข่งต่อได้), หรือ { points, rank, pet } เมื่อได้รางวัลสำเร็จ
function tryClaimPetChallenge(userId, rawText) {
  const ch = db.petChallenge;
  if (!ch) return null;
  if (Date.now() - ch.startedAt > PET_CHALLENGE_ROUND_MS) return null; // หมดเวลารอบพิเศษแล้ว
  if (rawText.trim() !== ch.hookText) return null;
  if (ch.winners.includes(userId)) return null; // คนนี้ได้รางวัลรอบนี้ไปแล้ว พิมพ์ซ้ำไม่ได้เพิ่ม
  if (!db.pets || !db.pets[userId]) return { noPet: true };
  const rank = ch.winners.length; // อันดับก่อนเพิ่มคนนี้ (0-based)
  const points = rank < PET_CHALLENGE_TOP_SLOTS ? PET_CHALLENGE_TOP_POINTS : PET_CHALLENGE_REST_POINTS;
  ch.winners.push(userId);
  const pet = db.pets[userId];
  advancePetDays(pet);
  pet.lifeForce = Math.max(0, pet.lifeForce + points);
  const leveledUp = advancePetStageIfHigher(pet);
  return { points, rank: rank + 1, pet, leveledUp };
}

// เช็คว่ารอบเหตุการณ์พิเศษที่ค้างอยู่ครบ 1 นาทีหรือยัง ถ้าครบแล้วยังไม่ได้แจ้ง ให้ push บอกหมดเวลา
async function checkPetChallengeExpiry() {
  const ch = db.petChallenge;
  if (!ch || ch.timeUpAnnounced) return;
  if (Date.now() - ch.startedAt < PET_CHALLENGE_ROUND_MS) return;
  ch.timeUpAnnounced = true;
  saveDB(db);
  let text;
  if (ch.winners.length > 0) {
    text = `⏰ หมดเวลารอบพิเศษแล้วครับ! รอบนี้มีคนพิมพ์ถูกทันเวลาทั้งหมด ${ch.winners.length} คน 🎉`;
  } else if (ch.kind === 'math') {
    // โจทย์เลขไม่มีใครเห็นคำตอบมาก่อน (ไม่เหมือนโจทย์พิมพ์ตามที่เฉลยอยู่ในตัวข้อความอยู่แล้ว) เลยเฉลยให้ตอนหมดเวลา
    text = `⏰ หมดเวลารอบพิเศษแล้วครับ ไม่มีใครตอบถูกเลย เฉลยคือ ${ch.hookText} รอบหน้ามาลองใหม่นะ!`;
  } else {
    text = '⏰ หมดเวลารอบพิเศษแล้วครับ ไม่มีใครพิมพ์ถูกทันเลย รอบหน้ามาลองใหม่นะ!';
  }
  try {
    await client.pushMessage(db.groupId, { type: 'text', text });
  } catch (e) {
    console.error('pet challenge time-up push failed:', e.originalError?.response?.data || e.message);
  }
}

// เช็คทุก 1 นาทีว่าถึงเวลาที่สุ่มไว้ของวันนี้ (4 รอบ) หรือยัง ถ้าถึงแล้วค่อยส่งเหตุการณ์พิเศษ
// เก็บเวลาที่สุ่มไว้ใน db ด้วย กัน server รีสตาร์ทแล้วสุ่มเวลาใหม่/ส่งซ้ำ พร้อมเช็คปิดรอบที่หมดเวลาไปด้วย
async function petSchedulerTick() {
  const dayIndex = bangkokDayIndex();
  if (!db.petSchedule || db.petSchedule.dayIndex !== dayIndex) {
    db.petSchedule = {
      dayIndex,
      times: pickPetEventTimes(dayIndex),
      kinds: pickDailyRoundKinds(PET_EVENT_ROUNDS_PER_DAY), // ครึ่งนึงโจทย์เลข ครึ่งนึงพิมพ์ตาม สุ่มลำดับ
      sent: new Array(PET_EVENT_ROUNDS_PER_DAY).fill(false),
    };
    saveDB(db);
  }
  const now = Date.now();
  for (let i = 0; i < db.petSchedule.times.length; i++) {
    if (!db.petSchedule.sent[i] && now >= db.petSchedule.times[i]) {
      db.petSchedule.sent[i] = true;
      saveDB(db);
      await triggerPetEvent(i);
    }
  }
  await checkPetChallengeExpiry();
}
setInterval(() => {
  petSchedulerTick().catch((e) => console.error('petSchedulerTick error:', e.message));
}, 60 * 1000);
// ==== จบส่วนเกมเลี้ยงสัตว์เลี้ยง ======================================================

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') return;
  if (event.source.type !== 'group') return; // ใช้งานเฉพาะในกลุ่มไลน์ปกติเท่านั้น

  const groupId = event.source.groupId;
  const userId = event.source.userId;
  const text = event.message.text.trim();

  if (!db.groupId) {
    db.groupId = groupId;
    saveDB(db);
  }

  // เก็บชื่อคนที่พิมพ์ในกลุ่มนี้ไว้ (ใช้สุ่มตอบคำถาม /ใคร) - ทำเงียบๆ ทุกข้อความ ไม่กระทบคำสั่งอื่น
  ensureGroupMemberCached(groupId, userId).catch(() => {});

  // ---- ใครก็พิมพ์ได้ ทุกกลุ่ม ----

  // เกมพิมพ์ตามเหตุการณ์พิเศษของสัตว์เลี้ยง (เช็คก่อนสุด เพราะข้อความที่พิมพ์ตามไม่ใช่คำสั่ง /)
  // ถ้าข้อความไม่ตรงกับเหตุการณ์ที่ค้างอยู่แบบเป๊ะๆ จะคืนค่า null แล้วปล่อยผ่านไปเช็คคำสั่งอื่นตามปกติ
  const petClaim = tryClaimPetChallenge(userId, text);
  if (petClaim) {
    if (petClaim.noPet) {
      await reply(event, '🐾 พิมพ์ถูกแล้วครับ! แต่ต้องมีสัตว์เลี้ยงของตัวเองก่อนถึงจะรับรางวัลได้ พิมพ์ "/รับเลี้ยง <ชื่อ>" ไว้เลย เผื่อรอบหน้า');
      return;
    }
    saveDB(db);
    const winnerName = (db.groupMembers && db.groupMembers[groupId] && db.groupMembers[groupId][userId]) || 'คุณ';
    const claimText =
      `✅ ${winnerName} พิมพ์ถูกเป็นคนที่ ${petClaim.rank}! ${petClaim.pet.species} "${petClaim.pet.name}" ได้พลังชีวิต +${petClaim.points} คะแนน (สะสม ${petClaim.pet.lifeForce})` +
      (petClaim.leveledUp ? `\n🎉 โตขึ้นเป็นระยะ "${PET_STAGE_LABEL[petClaim.pet.stage]}" แล้ว!` : '');
    if (petClaim.leveledUp) {
      const img = petImageMessage(petClaim.pet);
      if (img) {
        await replyMessages(event, [img, { type: 'text', text: claimText }]);
        return;
      }
    }
    await reply(event, claimText);
    return;
  }

  // แท็กชื่อบอทแล้วถามคำถามทั่วไปได้เลย (เช็คก่อนคำสั่งอื่นๆ เพราะข้อความมักมีคำถามต่อท้าย ไม่ใช่คำสั่ง /)
  const mentionQuestion = extractQuestionIfMentioned(event);
  if (mentionQuestion !== null) {
    if (!GEMINI_API_KEY) {
      await reply(event, 'ฟีเจอร์ถามตอบทั่วไปยังไม่ได้ตั้งค่าไว้ครับ (แอดมินต้องตั้งค่า GEMINI_API_KEY ใน Railway ก่อน ขอฟรีได้ที่ aistudio.google.com/apikey)');
      return;
    }
    if (!mentionQuestion) {
      await reply(event, 'แท็กมาแล้วพิมพ์คำถามต่อท้ายด้วยนะครับ เช่น "@ชื่อบอท วันนี้วันอะไร"');
      return;
    }
    try {
      const answer = await askGemini(mentionQuestion);
      await reply(event, answer || 'ขอโทษครับ ตอบคำถามนี้ไม่ได้ ลองถามใหม่อีกครั้งนะครับ');
    } catch (e) {
      console.error('askGemini failed:', e.message);
      await reply(event, 'ขอโทษครับ ตอนนี้ตอบคำถามไม่ได้ (ระบบขัดข้องหรือโควต้าฟรีเต็ม) ลองใหม่อีกครั้งนะครับ');
    }
    return;
  }

  if (/^\/whoami$/i.test(text)) {
    await reply(event, `userId ของคุณ: ${userId}`);
    return;
  }

  if (/^\/groupid$/i.test(text)) {
    await reply(event, `groupId ของกลุ่มนี้คือ:\n${groupId}\n\n(เอาไปตั้งเป็น QUIZ_GROUP_ID หรือ STATUS_GROUP_ID ใน Railway ได้)`);
    return;
  }

  // เช็คสถานะระบบสำรองข้อมูลอัตโนมัติ (กันคะแนน/สัตว์เลี้ยงหายตอน deploy ใหม่หรือรีสตาร์ท)
  if (/^\/(backup|สำรอง)$/i.test(text)) {
    if (!GITHUB_GIST_TOKEN) {
      await reply(event, '⚠️ ยังไม่ได้เปิดระบบสำรองข้อมูลอัตโนมัติ (ไม่ได้ตั้งค่า GITHUB_GIST_TOKEN) ข้อมูลอาจหายได้ถ้ามีการ deploy ใหม่หรือรีสตาร์ทเซิร์ฟเวอร์ ดูวิธีตั้งค่าได้ใน README.md');
      return;
    }
    if (!gistId) {
      await reply(event, '⚠️ ตั้งค่า Token ไว้แล้ว แต่ระบบสำรองข้อมูลเชื่อมต่อ GitHub ไม่สำเร็จ (อาจเพราะ Token ผิดหรือหมดอายุ) เช็ค Logs ของเซิร์ฟเวอร์เพื่อดูรายละเอียด');
      return;
    }
    const savedText = db._backupSavedAt
      ? new Date(db._backupSavedAt + 7 * 3600 * 1000).toISOString().replace('T', ' ').slice(0, 19) + ' น. (เวลาไทย)'
      : 'ยังไม่เคยสำรองข้อมูลรอบแรกเลย (รอ ~30 วินาทีหลังมีการเปลี่ยนแปลงข้อมูลครั้งแรก)';
    await reply(
      event,
      `✅ ระบบสำรองข้อมูลอัตโนมัติเปิดใช้งานอยู่\nสำรองข้อมูลล่าสุดเมื่อ: ${savedText}\nยังไม่ได้สำรองรอบล่าสุด: ${gistDirty ? 'มี (รอรอบถัดไป)' : 'ไม่มี ข้อมูลตรงกันแล้ว'}`
    );
    return;
  }

  // เช็ครายชื่อคนที่บอทรู้จักในกลุ่มนี้ (ใช้เป็นคำตอบของคำถาม /ใคร ได้)
  if (/^\/(สมาชิก|members)$/i.test(text)) {
    const members = (db.groupMembers && db.groupMembers[groupId]) || {};
    const names = Object.values(members);
    if (names.length === 0) {
      await reply(event, 'ยังไม่รู้จักใครในกลุ่มนี้เลยครับ ให้สมาชิกลองพิมพ์อะไรในกลุ่มสักครั้งก่อน บอทถึงจะจำชื่อไว้ได้');
    } else {
      await reply(
        event,
        `👥 คนที่บอทรู้จักในกลุ่มนี้ (${names.length} คน) — ใช้เป็นคำตอบของ /ใคร ได้:\n` + names.join(', ')
      );
    }
    return;
  }

  // คำถาม "/ใคร....." - สุ่มคำตอบเป็นคนในกลุ่ม (ใช้ได้ทุกคน ทุกกลุ่ม)
  const whoMatch = text.match(/^\/ใคร(.*)$/);
  if (whoMatch) {
    await handleWhoQuestion(event, groupId, userId, text);
    return;
  }

  // สุ่มเมนูอาหารไทย - ใช้ได้ทุกคน ทุกกลุ่ม ไม่ต้องเป็นแอดมิน พิมพ์ /food, /menu, /เมนู
  // หรือพิมพ์ประโยคทั่วไปเช่น "กินไรดี", "กินอะไรดี" ก็ได้
  const foodText = text.replace(/^\//, '').trim();
  if (/^(food|menu|เมนู)$/i.test(foodText) || /^กิน(ไร|อะไร)ดี/.test(foodText)) {
    await reply(event, `🍽️ มื้อนี้ลอง: ${randomFoodSuggestion()}`);
    return;
  }

  // ดูดวงรายวัน (เปิดไพ่ทาโร่) - ใช้ได้ทุกคน ทุกกลุ่ม พิมพ์ /ดวง หรือ /ดูดวง
  // จำกัดให้เปิดได้วันละ 1 ครั้งต่อ 1 line (นับตาม userId ไม่ใช่ต่อกลุ่ม เผื่อคนอยู่หลายกลุ่ม)
  if (/^(ดวง|ดูดวง)$/i.test(foodText)) {
    if (!db.fortuneChecks) db.fortuneChecks = {};
    const dayKey = todayKeyBangkok();
    if (db.fortuneChecks[userId] === dayKey) {
      const card = pickTodayCard(userId);
      await reply(event, `🔮 วันนี้เปิดไพ่ไปแล้วนะครับ ได้ใบ "${card.name}" ลองมาเปิดใหม่ได้พรุ่งนี้ครับ`);
      return;
    }
    db.fortuneChecks[userId] = dayKey;
    saveDB(db);
    await replyMessages(event, buildDailyFortuneMessages(userId));
    return;
  }

  // เกมเลี้ยงสัตว์เลี้ยง - ใช้ได้ทุกคน ทุกกลุ่ม
  const adoptMatch = text.match(/^\/(?:รับเลี้ยง|เลี้ยงสัตว์)\s*(.*)$/i);
  if (adoptMatch) {
    if (!db.pets) db.pets = {};
    const existing = db.pets[userId];
    if (existing) {
      await reply(event, `คุณมีสัตว์เลี้ยงอยู่แล้วนะครับ: ${existing.species} "${existing.name}" พิมพ์ /สัตว์เลี้ยง เพื่อดูสถานะได้เลย`);
      return;
    }
    const name = adoptMatch[1].trim();
    if (!name) {
      await reply(event, 'ตั้งชื่อสัตว์เลี้ยงด้วยครับ เช่น "/รับเลี้ยง โมจิ"');
      return;
    }
    db.pets[userId] = adoptPet(name);
    saveDB(db);
    const adoptText =
      `🎉 ยินดีด้วย! คุณได้รับเลี้ยง ${db.pets[userId].species} ชื่อ "${name}" แล้ว\n` +
      'ตอนนี้ยังเป็นไข่ลึกลับอยู่ ดูแลให้ครบทุกวัน (เช้า+บ่าย) เดี๋ยวจะฟักออกมาเอง แล้วค่อยๆ โตขึ้นเรื่อยๆ\n' +
      'พิมพ์ /ให้อาหาร ตอนบอทเตือน หรือเมื่อไหร่ก็ได้ที่นึกขึ้นได้ครับ (พิมพ์ /สัตว์เลี้ยง เพื่อดูสถานะได้ตลอด)';
    const adoptImg = petImageMessage(db.pets[userId]);
    if (adoptImg) {
      await replyMessages(event, [adoptImg, { type: 'text', text: adoptText }]);
      return;
    }
    await reply(event, adoptText);
    return;
  }

  if (/^\/(ให้อาหาร|ป้อนข้าว|ป้อนอาหาร)$/i.test(text)) {
    if (!db.pets || !db.pets[userId]) {
      await reply(event, 'คุณยังไม่มีสัตว์เลี้ยงนะครับ พิมพ์ "/รับเลี้ยง <ชื่อ>" เพื่อเริ่มเลี้ยงได้เลย');
      return;
    }
    const pet = db.pets[userId];
    const result = feedPet(pet);
    saveDB(db);
    if (!result.ok) {
      if (result.reason === 'outside_hours') {
        await reply(event, `ตอนนี้อยู่นอกช่วงเวลาป้อนอาหารครับ ป้อนได้แค่ช่วงเช้า (06:00-12:00 น.) หรือช่วงบ่าย (12:01-18:00 น.) เท่านั้น`);
        return;
      }
      if (result.reason === 'already_fed_this_session') {
        const sessionLabel = result.session === 'morning' ? 'ช่วงเช้า' : 'ช่วงบ่าย';
        await reply(event, `${pet.species} "${pet.name}" ป้อนอาหาร${sessionLabel}ไปแล้วครับ อีกรอบรอช่วงถัดไปนะ`);
        return;
      }
      await reply(event, `${pet.species} "${pet.name}" วันนี้ดูแลครบ 2 รอบแล้วครับ พรุ่งนี้ค่อยมาใหม่นะ`);
      return;
    }
    const sessionLabel = result.session === 'morning' ? 'ช่วงเช้า' : 'ช่วงบ่าย';
    const feedText =
      `🍚 ป้อนอาหาร${sessionLabel}ให้ ${pet.species} "${pet.name}" เรียบร้อย! ได้พลังชีวิต +25 คะแนน (สะสม ${pet.lifeForce})${pet.sick ? ' (แต่ยังป่วยอยู่ อย่าลืม /ป้อนยา ด้วยนะ)' : ' 😊'}` +
      (result.leveledUp ? `\n🎉 ${pet.species} "${pet.name}" โตขึ้นเป็นระยะ "${PET_STAGE_LABEL[pet.stage]}" แล้ว!` : '');
    if (result.leveledUp) {
      const img = petImageMessage(pet);
      if (img) {
        await replyMessages(event, [img, { type: 'text', text: feedText }]);
        return;
      }
    }
    await reply(event, feedText);
    return;
  }

  if (/^\/ป้อนยา$/i.test(text)) {
    if (!db.pets || !db.pets[userId]) {
      await reply(event, 'คุณยังไม่มีสัตว์เลี้ยงนะครับ พิมพ์ "/รับเลี้ยง <ชื่อ>" เพื่อเริ่มเลี้ยงได้เลย');
      return;
    }
    const pet = db.pets[userId];
    const result = givePetMedicine(pet);
    saveDB(db);
    if (!result.ok) {
      await reply(event, `${pet.species} "${pet.name}" ไม่ได้ป่วยนะครับ ไม่ต้องป้อนยา`);
      return;
    }
    await reply(event, `💊 ป้อนยาให้ ${pet.species} "${pet.name}" แล้ว หายป่วยแล้วครับ!`);
    return;
  }

  if (/^\/(สัตว์เลี้ยง|เพ็ท|pet)$/i.test(text)) {
    if (!db.pets || !db.pets[userId]) {
      await reply(event, 'คุณยังไม่มีสัตว์เลี้ยงนะครับ พิมพ์ "/รับเลี้ยง <ชื่อ>" เพื่อเริ่มเลี้ยงได้เลย');
      return;
    }
    const pet = db.pets[userId];
    advancePetDays(pet);
    saveDB(db);
    const statusImg = petImageMessage(pet);
    if (statusImg) {
      await replyMessages(event, [statusImg, { type: 'text', text: formatPetStatus(pet) }]);
      return;
    }
    await reply(event, formatPetStatus(pet));
    return;
  }

  if (/^\/(สัตว์เลี้ยงกลุ่ม|เพ็ทกลุ่ม)$/i.test(text)) {
    if (!db.pets || Object.keys(db.pets).length === 0) {
      await reply(event, 'ยังไม่มีใครเลี้ยงสัตว์เลี้ยงในกลุ่มนี้เลยครับ พิมพ์ "/รับเลี้ยง <ชื่อ>" เป็นคนแรกได้เลย');
      return;
    }
    const rows = Object.entries(db.pets).map(([uid, pet]) => {
      advancePetDays(pet);
      const name = (db.groupMembers && db.groupMembers[groupId] && db.groupMembers[groupId][uid]) || 'สมาชิก';
      return `${pet.sick ? '🤒' : '💚'} ${name}: ${pet.species} "${pet.name}" (${PET_STAGE_LABEL[pet.stage]}, พลังชีวิต ${pet.lifeForce} คะแนน)`;
    });
    saveDB(db);
    await reply(event, '🐾 สัตว์เลี้ยงในกลุ่มนี้:\n' + rows.join('\n'));
    return;
  }

  const inQuizGroup = isQuizGroup(groupId);
  const inStatusGroup = isStatusGroup(groupId);

  // ---- เช็คสถานะงานลูกค้า (เฉพาะกลุ่มที่ตั้งเป็น STATUS_GROUP_ID ถ้ามีการกำหนดไว้) ----
  if (inStatusGroup) {
    const statusMatch = text.match(/^\/?(?:status|สถานะ|เช็คสถานะ)\s*(.*)$/i);
    if (statusMatch) {
      const query = statusMatch[1].trim();
      if (!query) {
        await reply(event, 'พิมพ์ชื่อบริษัทต่อท้ายด้วยครับ เช่น "/สถานะ Alpha"');
      } else {
        await reply(event, buildStatusReply(query));
      }
      return;
    }
  }

  if (/^\/help$/i.test(text)) {
    const sections = [];
    if (inQuizGroup) {
      sections.push(
        [
          '📋 คำสั่งเกมตอบคำถาม (แอดมิน):',
          '/admin <รหัส> - ขอสิทธิ์แอดมิน',
          '/newgame - เริ่มเกมใหม่ (รีเซ็ตคะแนนทั้งหมด เริ่มข้อ 1)',
          '/ans A|B|C|D - ประกาศเฉลยข้อปัจจุบัน แล้วไปข้อถัดไปอัตโนมัติ',
          '/skip - ข้ามข้อปัจจุบันโดยไม่ให้คะแนนใคร',
          '/score - แสดงคะแนนสรุปในแชท',
          '/endgame - จบเกม',
          '',
          'ผู้เล่น: ตอบคำถามโดยพิมพ์ตัวอักษรเดียว เช่น A, B, C, D หรือ ก, ข, ค, ง ก็ได้',
          '(ตอบได้ครั้งเดียวต่อข้อ ตอบซ้ำจะไม่ถูกนับ)',
        ].join('\n')
      );
    }
    if (inStatusGroup) {
      sections.push(
        [
          '📊 คำสั่งเช็คสถานะงานลูกค้า:',
          '/สถานะ <ชื่อบริษัทหรือบางส่วนของชื่อ> - ค้นหาสถานะล่าสุด',
          `ข้อมูลอัปเดตล่าสุด: ${DB.updatedAt ? DB.updatedAt.toLocaleString('th-TH') : 'ยังไม่มีข้อมูล (แอดมินต้องอัปโหลดที่หน้า /upload ก่อน)'}`,
        ].join('\n')
      );
    }
    sections.push(
      [
        'ใช้ได้ทุกคนทุกกลุ่ม:',
        'แท็กชื่อบอทแล้วพิมพ์คำถามต่อท้าย - ถามคำถามทั่วไปได้เลย (ต้องให้แอดมินตั้งค่า GEMINI_API_KEY ไว้ก่อน)',
        'พิมพ์ /food, /menu, /เมนู หรือ "กินไรดี" ให้บอทสุ่มเมนูอาหารไทย',
        '/ดวง หรือ /ดูดวง - เปิดไพ่ทาโร่จริงดูดวงประจำวัน (รูปไพ่ + คำทำนาย + ข้อควรระวัง + สีมงคล + เลขเด่น 2 ชุด + คนที่ควรเลี่ยงคุยด้วยวันนี้) เปิดได้วันละ 1 ครั้งต่อคน เปลี่ยนใหม่ทุกเที่ยงคืน',
        '',
        '🐾 เกมเลี้ยงสัตว์เลี้ยง:',
        '/รับเลี้ยง <ชื่อ> - รับเลี้ยงสัตว์เลี้ยงของตัวเอง (คนละ 1 ตัว)',
        '/ให้อาหาร - ดูแลสัตว์เลี้ยง รอบละ +25 คะแนนทันที (เช้า 06:00-12:00 น. / บ่าย 12:01-18:00 น. เท่านั้น)',
        '/ป้อนยา - รักษาสัตว์เลี้ยงตอนป่วย (ไม่มีตายถาวร หายได้เสมอ)',
        '/สัตว์เลี้ยง - ดูสถานะสัตว์เลี้ยงของตัวเอง (พลังชีวิตสะสม + ระยะการเติบโต)',
        '/สัตว์เลี้ยงกลุ่ม - ดูสัตว์เลี้ยงของทุกคนในกลุ่ม',
        'ดูแลครึ่งวัน (แค่เช้า/แค่บ่าย) ยังได้ +25 เฉยๆ ไม่มีโทษ แต่ถ้าขาดดูแลทั้งวัน (ไม่ป้อนเลย) โดนหักพลังชีวิต -20 ท้ายวัน ขาดติดกัน 2 วันจะป่วยทันที',
        'ระยะการเติบโต: ไข่ลึกลับ → ลูก(50) → เด็ก(400) → วัยรุ่น(1000) → โตเต็มวัย(2000) → พร้อมสืบพันธุ์(4000) → แก่จัด(7500) → ตำนาน(15000) [ตัวเลข = พลังชีวิตสะสม ไม่มีวันถอยกลับ]',
        '🎲 บอทจะสุ่ม "เหตุการณ์พิเศษ" push เข้ากลุ่มวันละ 4 รอบ (เวลาสุ่มช่วง 08:30-17:00 น.)',
        'ครึ่งหนึ่ง (2 รอบ) เป็นโจทย์เลขบวก/ลบ 2-3 หลัก พิมพ์คำตอบให้ถูก, อีกครึ่ง (2 รอบ) เป็นข้อความสั้นๆ แบบสุ่ม ให้พิมพ์ตามให้ตรงเป๊ะ (สุ่มลำดับว่าจะเจอแบบไหนก่อน)',
        'รอบละ 1 นาที 5 คนแรกที่ตอบถูกได้ +30 คะแนน คนที่เหลือได้ +10 คะแนน (ต้องมีสัตว์เลี้ยงก่อนถึงรับได้)',
        '🎨 มีรูปสัตว์เลี้ยงน่ารักๆ แนบให้อัตโนมัติตอนรับเลี้ยง/เช็คสถานะ/โตขึ้นเป็นระยะใหม่',
        '',
        '/ใคร<คำถาม> เช่น "/ใครหล่อที่สุด" - สุ่มคำตอบเป็นคนในกลุ่ม',
        '/สมาชิก - ดูว่าบอทรู้จักใครในกลุ่มนี้บ้าง (ใช้เป็นคำตอบของ /ใคร ได้)',
        '/groupid - ดู groupId ของกลุ่มนี้',
        '/backup (หรือ /สำรอง) - เช็คสถานะระบบสำรองข้อมูลอัตโนมัติ กันคะแนน/สัตว์เลี้ยงหายตอน deploy ใหม่หรือรีสตาร์ท',
      ].join('\n')
    );
    await reply(event, sections.join('\n\n'));
    return;
  }

  if (/^\/admin\s+/i.test(text)) {
    if (!inQuizGroup) return; // คำสั่งแอดมินเกมใช้ได้เฉพาะกลุ่มเกม
    const code = text.replace(/^\/admin\s+/i, '').trim();
    if (code === ADMIN_PASSCODE) {
      if (!db.admins.includes(userId)) db.admins.push(userId);
      saveDB(db);
      await reply(event, '✅ ตั้งเป็นแอดมินเรียบร้อยแล้ว พิมพ์ /help เพื่อดูคำสั่งทั้งหมด');
    } else {
      await reply(event, '❌ รหัสไม่ถูกต้อง');
    }
    return;
  }

  // ---- คำสั่งเฉพาะแอดมิน (เฉพาะกลุ่มเกม) ----
  if (inQuizGroup && isAdmin(userId)) {
    if (/^\/newgame$/i.test(text)) {
      db.active = true;
      db.currentQuestion = 1;
      db.pendingAnswers = {};
      db.players = {};
      db.log = [];
      saveDB(db);
      await reply(event, `🎮 เริ่มเกมใหม่แล้ว! ข้อที่ ${db.currentQuestion} เริ่มตอบได้เลย (พิมพ์ A/B/C/D)`);
      return;
    }

    if (/^\/endgame$/i.test(text)) {
      db.active = false;
      saveDB(db);
      await reply(event, '🏁 จบเกม! สรุปคะแนนรวม:\n' + formatBoard(topScores(10)));
      return;
    }

    if (/^\/skip$/i.test(text)) {
      db.log.push({ q: db.currentQuestion, answer: null, skipped: true });
      db.currentQuestion += 1;
      db.pendingAnswers = {};
      saveDB(db);
      await reply(event, `⏭️ ข้ามข้อที่แล้ว ไปข้อที่ ${db.currentQuestion}`);
      return;
    }

    if (/^\/score$/i.test(text)) {
      await reply(event, '📊 คะแนนล่าสุด:\n' + formatBoard(topScores(10)));
      return;
    }

    const ansMatch = text.match(/^\/ans\s+(.+)$/i);
    if (ansMatch) {
      const correct = normalizeAnswerLetter(ansMatch[1]);
      if (!correct) {
        await reply(event, 'พิมพ์เฉลยเป็น A-F หรือ ก-ฉ เท่านั้น เช่น /ans B หรือ /ans ข');
        return;
      }
      if (!db.active) {
        await reply(event, 'ยังไม่ได้เริ่มเกม พิมพ์ /newgame ก่อนนะครับ');
        return;
      }
      const q = db.currentQuestion;
      let correctCount = 0;
      for (const [uid, ans] of Object.entries(db.pendingAnswers)) {
        ensurePlayer(uid, db.players[uid] && db.players[uid].name);
        db.players[uid].answeredCount += 1;
        if (ans === correct) {
          db.players[uid].score += 1;
          db.players[uid].correctCount += 1;
          correctCount += 1;
        }
      }
      db.log.push({
        q,
        answer: correct,
        respondents: Object.keys(db.pendingAnswers).length,
        correctCount,
      });
      db.currentQuestion += 1;
      db.pendingAnswers = {};
      saveDB(db);
      await reply(
        event,
        `✅ เฉลยข้อ ${q} = ${correct}\nตอบถูก ${correctCount} คน\n➡️ ข้อถัดไปคือข้อ ${db.currentQuestion}`
      );
      return;
    }
  }

  // ---- คำตอบของผู้เล่น (A/B/C/D หรือ ก/ข/ค/ง ...) (เฉพาะกลุ่มเกม) ----
  if (inQuizGroup && db.active) {
    const playerAnswer = normalizeAnswerLetter(text);
    if (playerAnswer) {
      if (db.pendingAnswers[userId] !== undefined) return; // ตอบไปแล้วสำหรับข้อนี้
      const name = await getDisplayName(groupId, userId);
      ensurePlayer(userId, name);
      db.pendingAnswers[userId] = playerAnswer;
      saveDB(db);
    }
  }
}

// --- Dashboard API ------------------------------------------------------
app.get('/api/scores', (req, res) => {
  const board = Object.entries(db.players)
    .map(([uid, p]) => ({ uid, ...p }))
    .sort((a, b) => b.score - a.score);
  res.json({
    active: db.active,
    currentQuestion: db.currentQuestion,
    respondingNow: Object.keys(db.pendingAnswers).length,
    players: board,
    log: db.log,
  });
});

// --- หน้าอัปโหลดไฟล์ Excel สถานะลูกค้า (มีรหัสผ่าน) --------------------
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

function uploadPageHtml({ message, isError } = {}) {
  const statusLine = DB.updatedAt
    ? `อัปเดตล่าสุด: ${DB.updatedAt.toLocaleString('th-TH')} (Initial ${DB.initial.length} / SU ${DB.su.length} / Recer ${DB.recer.length} บริษัท)`
    : 'ยังไม่มีข้อมูลในระบบ';
  return `<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>อัปโหลดไฟล์สถานะลูกค้า</title>
<style>
  body { font-family: sans-serif; max-width: 480px; margin: 40px auto; padding: 0 16px; color: #222; }
  h1 { font-size: 20px; }
  .box { border: 1px solid #ddd; border-radius: 8px; padding: 16px; margin-top: 16px; }
  label { display: block; margin-top: 12px; font-size: 14px; }
  input[type=password], input[type=file] { width: 100%; padding: 8px; margin-top: 4px; box-sizing: border-box; }
  button { margin-top: 16px; padding: 10px 20px; background: #06c755; color: white; border: none; border-radius: 6px; font-size: 15px; cursor: pointer; }
  .msg { padding: 10px; border-radius: 6px; margin-top: 12px; }
  .ok { background: #e6f7e9; color: #157a2e; }
  .err { background: #fdeaea; color: #a12222; }
  .status { color: #666; font-size: 13px; }
</style>
</head>
<body>
  <h1>อัปโหลดไฟล์ Excel สถานะลูกค้า</h1>
  <p class="status">${statusLine}</p>
  ${message ? `<div class="msg ${isError ? 'err' : 'ok'}">${message}</div>` : ''}
  <div class="box">
    <form method="POST" action="/upload" enctype="multipart/form-data">
      <label>ไฟล์ Excel (.xlsx)
        <input type="file" name="file" accept=".xlsx" required />
      </label>
      <label>รหัสผ่าน
        <input type="password" name="password" required />
      </label>
      <button type="submit">อัปโหลด</button>
    </form>
  </div>
</body>
</html>`;
}

app.get('/upload', (req, res) => {
  res.send(uploadPageHtml());
});

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.body.password || req.body.password !== UPLOAD_PASSCODE) {
      return res.status(401).send(uploadPageHtml({ message: '❌ รหัสผ่านไม่ถูกต้อง', isError: true }));
    }
    if (!req.file) {
      return res.status(400).send(uploadPageHtml({ message: '❌ กรุณาเลือกไฟล์ Excel', isError: true }));
    }
    fs.mkdirSync(path.dirname(XLSX_PATH), { recursive: true });
    fs.writeFileSync(XLSX_PATH, req.file.buffer);
    await reloadCustomerDB();
    res.send(
      uploadPageHtml({
        message: `✅ อัปโหลดสำเร็จ! (Initial ${DB.initial.length} / SU ${DB.su.length} / Recer ${DB.recer.length} บริษัท)`,
      })
    );
  } catch (e) {
    console.error('upload error:', e);
    res.status(500).send(uploadPageHtml({ message: '❌ อ่านไฟล์ไม่สำเร็จ: ' + e.message, isError: true }));
  }
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.redirect('/dashboard.html'));

// จับ route ที่ไม่ match อะไรเลย เพื่อ debug ผ่าน Logs
app.use((req, res) => {
  console.log(`[404] ${req.method} ${req.originalUrl}`);
  res.status(404).send('Not Found');
});

const PORT = process.env.PORT || 3000;
dbBackupReady.then(() => {
  app.listen(PORT, () => console.log('LINE quiz scoreboard bot listening on port ' + PORT));
});
