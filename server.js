require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET,
};
const ADMIN_PASSCODE = process.env.ADMIN_PASSCODE || 'changeme';
const UPLOAD_PASSCODE = process.env.UPLOAD_PASSCODE || 'changeme123';

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
}

let db = loadDB();

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
function formatInProgress({ type, rec }, latestRound, certStart, certExpire) {
  if (type === 'initial') {
    if (isDoneStatus(rec.status)) return null;
    const statusText = rec.status ? `${rec.status}(Initial)` : '(ยังไม่ระบุ)';
    const lines = [`🆕 ${rec.companyName}`, `ระบบ: ${rec.system || '-'}`, `สถานะ: ${statusText}`];
    // ระยะเวลาดำเนินการ: นับเฉพาะตอนที่ Status มีการลงข้อมูลแล้ว (ไม่ใช่ช่องว่าง) - นับจากวันที่ยื่นเอกสารขึ้นทะเบียน ถึงวันนี้
    // ถ้าช่องวันที่ตั้งต้นว่าง (ข้อมูลต้นทางไม่มี) ให้นับเป็น 0 วัน แทนที่จะซ่อนบรรทัดไปเลย
    if (rec.status) {
      const d = daysSince(rec.submittedAt) ?? 0;
      lines.push(`ระยะเวลาดำเนินการ: ${d} วัน`);
    }
    return lines.join('\n');
  }
  if (type === 'su') {
    const su1Done = isDoneStatus(rec.su1Status);
    const su2Done = isDoneStatus(rec.su2Status);
    if (su1Done && su2Done) return null; // ผ่านทั้ง SU1 และ SU2 แล้ว ถือว่าจบรอบนี้
    const lines = [`🔄 ${rec.companyName}`, `ระบบ: ${rec.system || '-'}`];
    // ถ้า SU1 ผ่านแล้วแต่ SU2 ยังไม่เสร็จ โชว์เฉพาะ SU2 (ไม่โชว์ SU1 ที่ผ่านแล้วซ้ำ) - ไม่ต้องแปะ (SU1)/(SU2) ซ้ำเพราะบอกอยู่แล้วที่หน้าบรรทัด
    // ระยะเวลาดำเนินการของแต่ละรอบ: นับจากวันที่ Sent Plan SU1/SU2 (เฉพาะรอบที่ Status มีข้อมูลแล้ว) ถึงวันนี้
    if (!su1Done) {
      lines.push(`SU1: ${rec.su1Status || '(ยังไม่ระบุ)'}`);
      if (rec.su1Status) {
        const d = daysSince(rec.su1SentPlan) ?? 0;
        lines.push(`ระยะเวลาดำเนินการ SU1: ${d} วัน`);
      }
    }
    if (!su2Done) {
      lines.push(`SU2: ${rec.su2Status || '(ยังไม่ถึงรอบ/ยังไม่ระบุ)'}`);
      if (rec.su2Status) {
        const d = daysSince(rec.su2SentPlan) ?? 0;
        lines.push(`ระยะเวลาดำเนินการ SU2: ${d} วัน`);
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
    lines.push(`ระยะเวลาดำเนินการ: ${d} วัน`);
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

// บางบริษัทในไฟล์มีชื่อซ้ำแบบมีเลขวงเล็บต่อท้าย เช่น "ABC Co., Ltd." กับ "ABC Co., Ltd. (2)"
// ซึ่งหมายถึงบริษัทเดียวกันแต่เป็นรอบ/สัญญาใหม่ - ตัดวงเล็บออกเพื่อจัดกลุ่มเป็นบริษัทเดียวกัน
// และถือว่าเลขวงเล็บที่สูงกว่า = รอบที่ใหม่กว่าเสมอ (ไม่มีวงเล็บ = รอบแรก/เก่าที่สุด)
function companyNameParts(name) {
  const m = /^(.*?)\s*\((\d+)\)\s*$/.exec((name || '').trim());
  if (m) return { base: m[1].trim(), suffix: Number(m[2]) };
  return { base: (name || '').trim(), suffix: 0 };
}

// เลือก record ที่ "ใหม่ที่สุด" จากลิสต์เดียวกัน - ดูเลขวงเล็บต่อท้ายชื่อบริษัทก่อน (สูงกว่า = ใหม่กว่า)
// ถ้าเลขวงเล็บเท่ากัน ค่อยเทียบจากวันออกใบรับรอง แล้ว fallback ไปวันที่ยื่นขึ้นทะเบียน
function pickLatestRecord(records) {
  if (records.length === 0) return null;
  return records
    .slice()
    .sort((a, b) => {
      const sa = companyNameParts(a.companyName).suffix;
      const sb = companyNameParts(b.companyName).suffix;
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
  if (rec.cerExpireDate) lines.push(`วันที่หมดอายุ: ${rec.cerExpireDate}`);
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

// วันที่ "ใบรับรองฉบับปัจจุบันเริ่มมีผล" - ดูจาก Cer Issue Date ของ Initial หรือ Recer เท่านั้น (เอาที่ล่าสุด)
// ไม่ใช้วันที่ของ SU เลย เพราะ SU เป็นแค่การตรวจติดตามใบรับรองเดิม ไม่ได้ออกใบรับรองใหม่
// (ใบรับรองมีอายุ 3 ปี และวัน issue ของรอบใหม่จะอยู่หลังวันหมดอายุของรอบก่อนหน้า 1 วันเสมอ)
function latestCertStart(records) {
  const candidates = [];
  for (const { type, rec } of records) {
    if (type === 'initial' || type === 'recer') {
      const t = parseThaiDate(rec.cerIssueDate);
      if (t !== null) candidates.push({ t, dateStr: rec.cerIssueDate });
    }
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.t - a.t);
  return candidates[0].dateStr;
}

// วันที่ "ใบรับรองฉบับปัจจุบันหมดอายุ" - คู่กับ latestCertStart (เอาวันหมดอายุที่ล่าสุด จาก Initial (Cer Expire Date)
// หรือ Recer (วันที่ cer จีนหมดอายุ) เท่านั้น) เพื่อให้สอดคล้องกับใบรับรองฉบับเดียวกับที่ certStart อ้างถึง
function latestCertExpire(records) {
  const candidates = [];
  for (const { type, rec } of records) {
    if (type === 'initial' && rec.cerExpireDate) {
      const t = parseThaiDate(rec.cerExpireDate);
      if (t !== null) candidates.push({ t, dateStr: rec.cerExpireDate });
    } else if (type === 'recer' && rec.certExpireDate) {
      const t = parseThaiDate(rec.certExpireDate);
      if (t !== null) candidates.push({ t, dateStr: rec.certExpireDate });
    }
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.t - a.t);
  return candidates[0].dateStr;
}

// รวม record ทั้งหมดของบริษัทเดียวกันเข้าด้วยกัน แล้วตัดสินใจว่าจะโชว์สถานะที่กำลังทำอยู่ หรือสรุปรอบล่าสุด
function buildCompanyBlock(companyName, records) {
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
  // ใช้ชื่อบริษัทแบบตัดเลขวงเล็บต่อท้ายออก (companyNameParts().base) เป็น key เพื่อรวม
  // ชื่อซ้ำที่เป็นรอบใหม่ เช่น "ABC Co., Ltd." กับ "ABC Co., Ltd. (2)" ให้เป็นบริษัทเดียวกัน
  const byCompany = new Map();
  for (const r of results) {
    const key = companyNameParts(r.rec.companyName).base;
    if (!byCompany.has(key)) byCompany.set(key, []);
    byCompany.get(key).push(r);
  }

  if (byCompany.size > MAX_RESULTS) {
    const names = [...byCompany.keys()].slice(0, MAX_RESULTS);
    return (
      `พบ ${byCompany.size} บริษัทที่ตรงกับ "${query}" เยอะเกินไป กรุณาพิมพ์ชื่อให้เจาะจงมากขึ้นครับ\n\n` +
      `ตัวอย่างที่พบ:\n` +
      names.map((n) => '- ' + n).join('\n')
    );
  }

  const blocks = [...byCompany.entries()].map(([name, records]) => buildCompanyBlock(name, records));
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
  if (/^\/whoami$/i.test(text)) {
    await reply(event, `userId ของคุณ: ${userId}`);
    return;
  }

  if (/^\/groupid$/i.test(text)) {
    await reply(event, `groupId ของกลุ่มนี้คือ:\n${groupId}\n\n(เอาไปตั้งเป็น QUIZ_GROUP_ID หรือ STATUS_GROUP_ID ใน Railway ได้)`);
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
        'พิมพ์ /food, /menu, /เมนู หรือ "กินไรดี" ให้บอทสุ่มเมนูอาหารไทย',
        '/ใคร<คำถาม> เช่น "/ใครหล่อที่สุด" - สุ่มคำตอบเป็นคนในกลุ่ม',
        '/สมาชิก - ดูว่าบอทรู้จักใครในกลุ่มนี้บ้าง (ใช้เป็นคำตอบของ /ใคร ได้)',
        '/groupid - ดู groupId ของกลุ่มนี้',
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
app.listen(PORT, () => console.log('LINE quiz scoreboard bot listening on port ' + PORT));
