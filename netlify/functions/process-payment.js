// netlify/functions/process-payment.js
// POST /api/process-payment
// Processes a Square payment and creates an Outlook calendar event.

const https  = require('https');
const crypto = require('crypto');

const usedNonces = new Set();

function sanitize(str, maxLen = 200) {
  if (typeof str !== 'string') return '';
  return str.trim().replace(/[<>"'`]/g, '').slice(0, maxLen);
}
function isValidEmail(email) { return /^[^\s@]{1,64}@[^\s@]{1,255}\.[^\s@]{1,63}$/.test(email); }
function isValidPhone(phone) { return /^[\d\s\-\(\)\+]{7,20}$/.test(phone); }

function httpsPost(hostname, path, headers, body) {
  const bodyStr = JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const options = { hostname, path, method: 'POST', headers: { ...headers, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(bodyStr) } };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => { try { resolve({ status: res.statusCode, body: JSON.parse(data) }); } catch { reject(new Error('Invalid JSON response')); } });
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

// FIX: Format ISO date "YYYY-MM-DD" as human-readable string for Square notes
function formatDisplayDate(isoDate) {
  const [year, month, day] = isoDate.split('-').map(Number);
  return new Date(year, month - 1, day).toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
}
function formatDisplayTime(time) {
  const [h, m] = time.split(':').map(Number);
  return `${h % 12 || 12}:${String(m).padStart(2,'0')} ${h >= 12 ? 'PM' : 'AM'}`;
}

async function chargeSquare(sourceId, amountCents, bookingData, idempotencyKey) {
  const { SQUARE_ACCESS_TOKEN, SQUARE_LOCATION_ID, SQUARE_ENVIRONMENT } = process.env;
  const hostname = SQUARE_ENVIRONMENT === 'production' ? 'connect.squareup.com' : 'connect.squareupsandbox.com';
  const res = await httpsPost(hostname, '/v2/payments',
    { 'Square-Version': '2024-01-17', Authorization: `Bearer ${SQUARE_ACCESS_TOKEN}` },
    {
      source_id: sourceId,
      idempotency_key: idempotencyKey,
      amount_money: { amount: amountCents, currency: 'USD' },
      location_id: SQUARE_LOCATION_ID,
      // FIX: use formatted displayDate + displayTime instead of raw ISO string
      note: `WMS Delivery Services LLC — ${bookingData.serviceName} — ${bookingData.displayDate} at ${bookingData.displayTime}`,
      buyer_email_address: bookingData.email,
    }
  );
  if (res.status !== 200 || !res.body.payment) throw new Error(res.body?.errors?.[0]?.detail || 'Payment failed');
  return res.body.payment;
}

async function getMsToken() {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;
  const body = new URLSearchParams({ grant_type: 'client_credentials', client_id: AZURE_CLIENT_ID, client_secret: AZURE_CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default' }).toString();
  return new Promise((resolve, reject) => {
    const buf = Buffer.from(body);
    const req = https.request({ hostname: 'login.microsoftonline.com', path: `/${AZURE_TENANT_ID}/oauth2/v2.0/token`, method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': buf.length } }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => { const p = JSON.parse(data); p.access_token ? resolve(p.access_token) : reject(new Error('Token failed')); });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function createCalendarEvent(token, bookingData) {
  const { OUTLOOK_USER_EMAIL } = process.env;
  const [year, month, day] = bookingData.date.split('-').map(Number);
  const [hour, min] = bookingData.time.split(':').map(Number);
  const startDT = new Date(year, month - 1, day, hour, min);
  const endDT = new Date(startDT.getTime() + 2 * 60 * 60 * 1000);
  const fmt = (d) => d.toISOString().slice(0, 19);
  const res = await httpsPost('graph.microsoft.com', `/v1.0/users/${encodeURIComponent(OUTLOOK_USER_EMAIL)}/events`,
    { Authorization: `Bearer ${token}` },
    {
      subject: `🚙 ${bookingData.serviceName} — ${bookingData.firstName} ${bookingData.lastName}`,
      body: { contentType: 'text', content: [`Service: ${bookingData.serviceName}`, `Date: ${bookingData.displayDate} at ${bookingData.displayTime}`, `Passenger: ${bookingData.firstName} ${bookingData.lastName}`, `Phone: ${bookingData.phone}`, `Email: ${bookingData.email}`, `Passengers: ${bookingData.passengers}`, `Pickup: ${bookingData.pickup}`, `Drop-off: ${bookingData.dropoff}`, bookingData.notes ? `Notes: ${bookingData.notes}` : '', `Confirmation: ${bookingData.confirmationId}`].filter(Boolean).join('\n') },
      start: { dateTime: fmt(startDT), timeZone: 'America/Chicago' },
      end:   { dateTime: fmt(endDT),   timeZone: 'America/Chicago' },
      isReminderOn: true, reminderMinutesBeforeStart: 60, showAs: 'busy',
    }
  );
  if (res.status !== 201) console.error('Calendar event creation failed:', res.body);
  return res.body?.id || null;
}

const PRICES = { airport: 8000, corporate: 37500, wedding: 35000, event: 20000, hourly: 10000 };

exports.handler = async function(event) {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', 'Access-Control-Allow-Origin': process.env.SITE_URL || 'https://your-site.netlify.app', 'Access-Control-Allow-Methods': 'POST' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  let payload;
  try { payload = JSON.parse(event.body); } catch { return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid JSON body' }) }; }
  const { sourceId, bookingData } = payload;
  if (!sourceId || typeof sourceId !== 'string' || sourceId.length > 200) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid payment token' }) };
  for (const f of ['date','time','serviceType','firstName','lastName','email','phone','pickup','dropoff','passengers','nonce']) {
    if (!bookingData?.[f]) return { statusCode: 400, headers, body: JSON.stringify({ error: `Missing field: ${f}` }) };
  }
  const nonce = sanitize(bookingData.nonce, 100);
  if (usedNonces.has(nonce)) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Duplicate submission' }) };
  if (!isValidEmail(bookingData.email)) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid email' }) };
  if (!isValidPhone(bookingData.phone)) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid phone' }) };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(bookingData.date)) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid date format' }) };
  if (!['airport','corporate','wedding','event','hourly'].includes(bookingData.serviceType)) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid service type' }) };
  const clean = {
    date: sanitize(bookingData.date, 20), time: sanitize(bookingData.time, 10),
    serviceType: bookingData.serviceType,
    // FIX: use serviceName from payload (sent by checkout.html) with fallback
    serviceName: sanitize(bookingData.serviceName || bookingData.serviceType, 100),
    firstName: sanitize(bookingData.firstName, 100), lastName: sanitize(bookingData.lastName, 100),
    email: bookingData.email.trim().toLowerCase().slice(0, 254),
    phone: sanitize(bookingData.phone, 30), pickup: sanitize(bookingData.pickup, 300),
    dropoff: sanitize(bookingData.dropoff, 300), passengers: parseInt(bookingData.passengers, 10) || 1,
    notes: sanitize(bookingData.notes || '', 500),
    confirmationId: `NPL-${Date.now()}-${crypto.randomBytes(4).toString('hex').toUpperCase()}`,
    // FIX: generate human-readable date/time for Square note & calendar event
    displayDate: formatDisplayDate(bookingData.date),
    displayTime: formatDisplayTime(bookingData.time),
  };
  usedNonces.add(nonce);
  setTimeout(() => usedNonces.delete(nonce), 60 * 60 * 1000);
  const amountCents = PRICES[clean.serviceType] || 15000;
  const idempotencyKey = crypto.randomUUID();
  try {
    const payment = await chargeSquare(sourceId, amountCents, clean, idempotencyKey);
    try { const msToken = await getMsToken(); await createCalendarEvent(msToken, clean); } catch (calErr) { console.error('Calendar event error (non-fatal):', calErr.message); }
    return { statusCode: 200, headers, body: JSON.stringify({ success: true, confirmationId: clean.confirmationId, amountPaid: (amountCents / 100).toFixed(2), paymentId: payment.id }) };
  } catch (err) {
    console.error('Payment error:', err.message);
    return { statusCode: 402, headers, body: JSON.stringify({ success: false, error: err.message || 'Payment could not be processed.' }) };
  }
};
