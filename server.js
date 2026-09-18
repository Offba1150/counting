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

  // สุ่มเมนูอาหารไทย - ใช้ได้ทุกคน ไม่ต้องเป็นแอดมิน พิมพ์ /food, /เมนู
  // หรือพิมพ์ประโยคทั่วไปเช่น "กินไรดี", "กินอะไรดี" ก็ได้
  const foodText = text.replace(/^\//, '').trim();
  if (/^(food|เมนู)$/i.test(foodText) || /^กิน(ไร|อะไร)ดี/.test(foodText)) {
    await reply(event, `🍽️ มื้อนี้ลอง: ${randomFoodSuggestion()}`);
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
        '',
        'ใช้ได้ทุกคน: พิมพ์ /food หรือ "กินไรดี" ให้บอทสุ่มเมนูอาหารไทยให้',
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
