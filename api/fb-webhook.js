// Facebook Messenger Webhook（Vercel Serverless Function）
//
// 為什麼要這一層：Meta 的 Webhook 一定要直接回 200，Google Apps Script 會先回 302 轉址，Meta 不接受。
// 所以 Meta → 這裡（通過驗證、檢查簽章、立刻回 200）→ 轉給 Apps Script 訂單系統，由 AI 客服回答。
//
// Vercel → Settings → Environment Variables 要設定（不要寫進程式，這個 repo 是公開的）：
//   FB_VERIFY_TOKEN   自己訂一串字，Meta 後台設定 Webhook 時填一樣的
//   FB_APP_SECRET     Meta 應用程式密鑰（應用程式設定 → 基本資料）
//   FB_RELAY_SECRET   自己訂一串密碼，Apps Script 指令碼屬性 FB_RELAY_SECRET 要一模一樣
//   GAS_URL           訂單系統 API 部署的網址（結尾 /exec）

const crypto = require('crypto');

function readRaw(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function validSignature(raw, header, secret) {
  if (!header || !secret || !header.startsWith('sha256=')) return false;
  const expected = crypto.createHmac('sha256', secret).update(raw).digest('hex');
  const got = header.slice(7);
  if (!/^[0-9a-f]{64}$/i.test(got)) return false;
  return crypto.timingSafeEqual(Buffer.from(got, 'hex'), Buffer.from(expected, 'hex'));
}

module.exports = async (req, res) => {
  // 1. Meta 設定 Webhook 時的驗證
  if (req.method === 'GET') {
    const q = req.query || {};
    if (q['hub.mode'] === 'subscribe' && process.env.FB_VERIFY_TOKEN && q['hub.verify_token'] === process.env.FB_VERIFY_TOKEN) {
      res.setHeader('Content-Type', 'text/plain');
      return res.status(200).send(String(q['hub.challenge'] || ''));
    }
    return res.status(403).send('Forbidden');
  }

  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  // 2. 檢查是 Meta 送來的（簽章）
  const raw = await readRaw(req);
  if (!validSignature(raw, req.headers['x-hub-signature-256'], process.env.FB_APP_SECRET)) {
    return res.status(401).send('Invalid signature');
  }

  let payload;
  try { payload = JSON.parse(raw.toString('utf8')); } catch (e) { return res.status(400).send('Bad JSON'); }
  if (payload.object !== 'page') return res.status(200).send('ignored');

  // 3. 轉給 Apps Script。Apps Script 收到後會自己跑完（AI 回答要幾秒），
  //    這裡最多等 8 秒就先回 200 給 Meta（Meta 約 20 秒內要收到回應），避免 Meta 以為失敗而重送。
  const forward = fetch(process.env.GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ source: 'facebook', relaySecret: process.env.FB_RELAY_SECRET, payload }),
    redirect: 'follow'
  }).catch((err) => console.error('轉送 Apps Script 失敗：', err));

  await Promise.race([forward, new Promise((r) => setTimeout(r, 8000))]);
  return res.status(200).send('EVENT_RECEIVED');
};
