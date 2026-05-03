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
      res.on('end', () => {
        // sendMail and other Graph endpoints return 2xx with empty body — don't choke on that
        if (!data) return resolve({ status: res.statusCode, body: null });
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); } catch { reject(new Error('Invalid JSON response')); }
      });
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

function escapeHtml(str) {
  if (typeof str !== 'string') return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
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

function buildCustomerEmailHtml(b, amountPaid) {
  const e = escapeHtml;
  const passengersLabel = `${b.passengers} passenger${b.passengers > 1 ? 's' : ''}`;
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:Helvetica,Arial,sans-serif;color:#222;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:30px 0;"><tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-top:4px solid #c9a84c;max-width:600px;">
<tr><td style="padding:40px 40px 20px;text-align:center;">
<div style="font-family:Georgia,serif;font-size:22px;font-weight:300;letter-spacing:2px;color:#1a1a1a;">WMS <span style="color:#c9a84c;">DELIVERY SERVICES LLC</span></div>
<div style="margin-top:24px;font-size:18px;color:#1a1a1a;font-weight:600;">Your Ride is Confirmed</div>
<div style="margin-top:8px;font-size:13px;color:#999;">Confirmation #${e(b.confirmationId)}</div>
</td></tr>
<tr><td style="padding:0 40px 20px;font-size:14px;line-height:1.6;color:#444;">Hi ${e(b.firstName)},<br><br>Thanks for booking with WMS Delivery Services. Your ride details are below — please review and reach out if anything needs to change.</td></tr>
<tr><td style="padding:0 40px 30px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #eee;border-collapse:collapse;">
<tr><td style="padding:12px 16px;background:#fafafa;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;width:35%;">Service</td><td style="padding:12px 16px;font-size:14px;color:#222;">${e(b.serviceName)}</td></tr>
<tr><td style="padding:12px 16px;background:#fafafa;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;border-top:1px solid #eee;">Date</td><td style="padding:12px 16px;font-size:14px;color:#222;border-top:1px solid #eee;">${e(b.displayDate)}</td></tr>
<tr><td style="padding:12px 16px;background:#fafafa;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;border-top:1px solid #eee;">Time</td><td style="padding:12px 16px;font-size:14px;color:#222;border-top:1px solid #eee;">${e(b.displayTime)}</td></tr>
<tr><td style="padding:12px 16px;background:#fafafa;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;border-top:1px solid #eee;">Pickup</td><td style="padding:12px 16px;font-size:14px;color:#222;border-top:1px solid #eee;">${e(b.pickup)}</td></tr>
<tr><td style="padding:12px 16px;background:#fafafa;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;border-top:1px solid #eee;">Drop-off</td><td style="padding:12px 16px;font-size:14px;color:#222;border-top:1px solid #eee;">${e(b.dropoff)}</td></tr>
<tr><td style="padding:12px 16px;background:#fafafa;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;border-top:1px solid #eee;">Passengers</td><td style="padding:12px 16px;font-size:14px;color:#222;border-top:1px solid #eee;">${e(passengersLabel)}</td></tr>
<tr><td style="padding:12px 16px;background:#fafafa;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;border-top:1px solid #eee;">Total Paid</td><td style="padding:12px 16px;font-size:14px;color:#222;border-top:1px solid #eee;font-weight:600;">$${e(amountPaid)}</td></tr>
</table>
</td></tr>
<tr><td style="padding:0 40px 30px;">
<div style="font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px;">What to expect</div>
<ul style="margin:0;padding-left:20px;font-size:14px;line-height:1.7;color:#444;">
<li>Your driver will arrive 10 minutes before your pickup time.</li>
<li>For airport pickups, we track your flight in real time and adjust automatically for delays.</li>
<li>You'll receive a reminder 24 hours before your ride.</li>
</ul>
</td></tr>
<tr><td style="padding:24px 40px;background:#1a1a1a;color:#f5f0e8;font-size:13px;line-height:1.7;text-align:center;">
<div>Need to change something?</div>
<div style="margin-top:6px;">Text or call <a href="tel:+15047104563" style="color:#c9a84c;text-decoration:none;font-weight:600;">(504) 710-4563</a> · 10 AM – 10 PM</div>
<div style="margin-top:6px;">Or reply directly to this email.</div>
</td></tr>
<tr><td style="padding:16px 40px;text-align:center;font-size:11px;color:#999;">WMS Delivery Services LLC · New Orleans, LA<br><a href="https://wmsdeliveryservicellc.com" style="color:#999;text-decoration:underline;">wmsdeliveryservicellc.com</a></td></tr>
</table>
</td></tr></table>
</body></html>`;
}

function buildCustomerEmailText(b, amountPaid) {
  const passengersLabel = `${b.passengers} passenger${b.passengers > 1 ? 's' : ''}`;
  return `WMS Delivery Services LLC

Your ride is confirmed.

Hi ${b.firstName},

Thanks for booking with WMS Delivery Services. Your ride details are below — please review and reach out if anything needs to change.

Confirmation: ${b.confirmationId}

Service:    ${b.serviceName}
Date:       ${b.displayDate}
Time:       ${b.displayTime}
Pickup:     ${b.pickup}
Drop-off:   ${b.dropoff}
Passengers: ${passengersLabel}
Total Paid: $${amountPaid}

What to expect:
  • Your driver will arrive 10 minutes before your pickup time.
  • For airport pickups, we track your flight in real time.
  • You'll receive a reminder 24 hours before your ride.

Need to change something?
Text or call (504) 710-4563 · 10 AM – 10 PM
Or reply directly to this email.

WMS Delivery Services LLC · New Orleans, LA
https://wmsdeliveryservicellc.com`;
}

async function sendCustomerEmail(token, bookingData, amountPaid) {
  const { OUTLOOK_USER_EMAIL } = process.env;
  const res = await httpsPost('graph.microsoft.com', `/v1.0/users/${encodeURIComponent(OUTLOOK_USER_EMAIL)}/sendMail`,
    { Authorization: `Bearer ${token}` },
    {
      message: {
        subject: `Your WMS Delivery ride is confirmed — ${bookingData.displayDate}`,
        body: { contentType: 'HTML', content: buildCustomerEmailHtml(bookingData, amountPaid) },
        toRecipients: [{ emailAddress: { address: bookingData.email } }],
        replyTo:     [{ emailAddress: { address: OUTLOOK_USER_EMAIL } }],
      },
      saveToSentItems: true,
    }
  );
  // sendMail returns 202 Accepted on success with no body
  if (res.status !== 202 && res.status !== 200) {
    console.error('Customer email send failed:', res.status, res.body);
    return false;
  }
  return true;
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
    confirmationId: `WMS-${Date.now()}-${crypto.randomBytes(4).toString('hex').toUpperCase()}`,
    // FIX: generate human-readable date/time for Square note & calendar event
    displayDate: formatDisplayDate(bookingData.date),
    displayTime: formatDisplayTime(bookingData.time),
  };
  usedNonces.add(nonce);
  setTimeout(() => usedNonces.delete(nonce), 60 * 60 * 1000);
  const amountCents = PRICES[clean.serviceType] || 15000;
  // FIX: use the booking nonce as the Square idempotency key. If the customer
  // double-clicks Submit and two requests fire across separate function instances,
  // Square dedupes them within its 24-hour window — preventing double charges.
  const idempotencyKey = nonce;
  try {
    const payment = await chargeSquare(sourceId, amountCents, clean, idempotencyKey);
    const amountPaid = (amountCents / 100).toFixed(2);
    // Microsoft Graph work — both calendar event and customer email use the same token.
    // Run in parallel; failure of one doesn't block the other or the payment response.
    try {
      const msToken = await getMsToken();
      await Promise.allSettled([
        createCalendarEvent(msToken, clean),
        sendCustomerEmail(msToken, clean, amountPaid),
      ]);
    } catch (msErr) {
      console.error('Microsoft Graph error (non-fatal):', msErr.message);
    }
    return { statusCode: 200, headers, body: JSON.stringify({ success: true, confirmationId: clean.confirmationId, amountPaid, paymentId: payment.id }) };
  } catch (err) {
    console.error('Payment error:', err.message);
    return { statusCode: 402, headers, body: JSON.stringify({ success: false, error: err.message || 'Payment could not be processed.' }) };
  }
};
