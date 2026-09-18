require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const fs = require('fs');
const path = require('path');

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET,
};
const ADMIN_PASSCODE = process.env.ADMIN_PASSCODE || 'changeme';

if (!config.channelAccessToken || !config.channelSecret) {
  console.warn('!!! CHANNEL_ACCESS_TOKEN / CHANNEL_SECRET ยังไม่ได้ตั้งค่า (ดูไฟล์ .env.example)');
}

const client = new line.Client(config);
const app = express();

// log ทุก request ที่เข้ามา ไว้ช่วย debug (ดูได้จากแท็บ Logs บน Render)
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

  // ---- ใครก็พิมพ์ได้ ----
  if (/^\/whoami$/i.test(text)) {
    await reply(event, `userId ของคุณ: ${userId}`);
    return;
  }

  if (/^\/help$/i.test(text)) {
    await reply(
      event,
      [
        'คำสั่งแอดมิน:',
        '/admin <รหัส> - ขอสิทธิ์แอดมิน',
        '/newgame - เริ่มเกมใหม่ (รีเซ็ตคะแนนทั้งหมด เริ่มข้อ 1)',
        '/ans A|B|C|D - ประกาศเฉลยข้อปัจจุบัน แล้วไปข้อถัดไปอัตโนมัติ',
        '/skip - ข้ามข้อปัจจุบันโดยไม่ให้คะแนนใคร',
        '/score - แสดงคะแนนสรุปในแชท',
        '/endgame - จบเกม',
        '',
        'ผู้เล่น: ตอบคำถามโดยพิมพ์ตัวอักษรเดียว เช่น A, B, C หรือ D',
        '(ตอบได้ครั้งเดียวต่อข้อ ตอบซ้ำจะไม่ถูกนับ)',
      ].join('\n')
    );
    return;
  }

  if (/^\/admin\s+/i.test(text)) {
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

  // ---- คำสั่งเฉพาะแอดมิน ----
  if (isAdmin(userId)) {
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

    const ansMatch = text.match(/^\/ans\s+([a-fA-F])$/i);
    if (ansMatch) {
      if (!db.active) {
        await reply(event, 'ยังไม่ได้เริ่มเกม พิมพ์ /newgame ก่อนนะครับ');
        return;
      }
      const correct = ansMatch[1].toUpperCase();
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

  // ---- คำตอบของผู้เล่น (A/B/C/D ...) ----
  if (db.active && /^[a-fA-F]$/.test(text)) {
    if (db.pendingAnswers[userId] !== undefined) return; // ตอบไปแล้วสำหรับข้อนี้
    const name = await getDisplayName(groupId, userId);
    ensurePlayer(userId, name);
    db.pendingAnswers[userId] = text.toUpperCase();
    saveDB(db);
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

app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.redirect('/dashboard.html'));

// จับ route ที่ไม่ match อะไรเลย เพื่อ debug ผ่าน Logs
app.use((req, res) => {
  console.log(`[404] ${req.method} ${req.originalUrl}`);
  res.status(404).send('Not Found');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('LINE quiz scoreboard bot listening on port ' + PORT));
