// netlify/functions/availability.js
// GET /api/availability?year=2025&month=6
// Fetches busy times from the driver's Outlook calendar via
// Microsoft Graph API using Client Credentials flow.

const https = require('https');

// Simple in-memory rate limiter (per IP)
const rateStore = new Map();
const RATE_LIMIT = 30;
const RATE_WINDOW = 60 * 1000;

function isRateLimited(ip) {
  const now = Date.now();
  const entry = rateStore.get(ip) || { count: 0, resetAt: now + RATE_WINDOW };
  if (now > entry.resetAt) {
    rateStore.set(ip, { count: 1, resetAt: now + RATE_WINDOW });
    return false;
  }
  entry.count++;
  rateStore.set(ip, entry);
  return entry.count > RATE_LIMIT;
}

async function getMsToken() {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;
  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error('Azure credentials not configured');
  }
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
  }).toString();
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.access_token) resolve(parsed.access_token);
          else reject(new Error(parsed.error_description || 'Token request failed'));
        } catch { reject(new Error('Invalid token response')); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function getCalendarEvents(token, startISO, endISO) {
  const { OUTLOOK_USER_EMAIL } = process.env;
  if (!OUTLOOK_USER_EMAIL) throw new Error('Outlook email not configured');
  const params = new URLSearchParams({ startDateTime: startISO, endDateTime: endISO, '$select': 'start,end', '$top': '100' });
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'graph.microsoft.com',
      path: `/v1.0/users/${encodeURIComponent(OUTLOOK_USER_EMAIL)}/calendarView?${params}`,
      method: 'GET',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Prefer: 'outlook.timezone="America/Chicago"' },
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.value) resolve(parsed.value);
          else reject(new Error('Unexpected Graph API response'));
        } catch { reject(new Error('Invalid Graph API response')); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// FIX: Graph API with Prefer: outlook.timezone returns Central Time strings
// WITHOUT a Z suffix e.g. "2026-03-25T09:00:00.0000000".
// Old code appended '-06:00' which is wrong during CDT (Mar-Nov, offset is -05:00).
// Fix: strip fractional seconds and parse as local time directly.
function parseCentralDateTime(dtStr) {
  const clean = dtStr.split('.')[0];
  if (clean.endsWith('Z')) return new Date(clean);
  return new Date(clean.replace('T', ' '));
}

function eventsToBookedSlots(events) {
  const ALL_TIMES = ['06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00','00:00'];
  const slots = {};
  for (const event of events) {
    const start = parseCentralDateTime(event.start.dateTime);
    const end   = parseCentralDateTime(event.end.dateTime);
    // Walk every calendar day the event touches. Old code only used the event's
    // start date, so multi-day Outlook blocks (e.g. all-day Mar 22-25) only
    // marked day 1 — leaving subsequent days fully "available" on the site.
    const dayCursor = new Date(start);
    dayCursor.setHours(0, 0, 0, 0);
    while (dayCursor < end) {
      for (const time of ALL_TIMES) {
        const [h, m] = time.split(':').map(Number);
        const slotStart = new Date(dayCursor);
        slotStart.setHours(h, m, 0, 0);
        const slotEnd = new Date(slotStart);
        slotEnd.setHours(h + 1, m, 0, 0);
        if (slotStart < end && slotEnd > start) {
          const dateKey = [slotStart.getFullYear(), String(slotStart.getMonth()+1).padStart(2,'0'), String(slotStart.getDate()).padStart(2,'0')].join('-');
          if (!slots[dateKey]) slots[dateKey] = [];
          if (!slots[dateKey].includes(time)) slots[dateKey].push(time);
        }
      }
      dayCursor.setDate(dayCursor.getDate() + 1);
    }
    // The booking page's "12 AM" slot for day N is a late-night pickup that
    // resolves to bookedSlots[N+1]["00:00"]. Outlook all-day blocks have end
    // at exact midnight (e.g. blocking 22-26 yields end=27 00:00), and the
    // loop above stops before the end day. Without this, the last covered
    // day's "12 AM" never matches and shows as available.
    if (end > start && end.getHours() === 0 && end.getMinutes() === 0 && end.getSeconds() === 0 && end.getMilliseconds() === 0) {
      const dateKey = [end.getFullYear(), String(end.getMonth()+1).padStart(2,'0'), String(end.getDate()).padStart(2,'0')].join('-');
      if (!slots[dateKey]) slots[dateKey] = [];
      if (!slots[dateKey].includes('00:00')) slots[dateKey].push('00:00');
    }
  }
  return slots;
}

exports.handler = async function(event, context) {
  const headers = {
    'Content-Type': 'application/json',
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
    'Access-Control-Allow-Origin': process.env.SITE_URL || 'https://your-site.netlify.app',
    'Access-Control-Allow-Methods': 'GET',
  };
  if (event.httpMethod !== 'GET') return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  const ip = event.headers['x-nf-client-connection-ip'] || event.headers['x-forwarded-for']?.split(',')[0]?.trim() || 'unknown';
  if (isRateLimited(ip)) return { statusCode: 429, headers, body: JSON.stringify({ error: 'Too many requests' }) };
  const year  = parseInt(event.queryStringParameters?.year,  10);
  const month = parseInt(event.queryStringParameters?.month, 10);
  if (!year || !month || month < 1 || month > 12 || year < 2020 || year > 2100) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid year/month' }) };
  }
  try {
    const token = await getMsToken();
    const startISO = new Date(year, month - 1, 1).toISOString();
    const endISO   = new Date(year, month,     1).toISOString();
    const events   = await getCalendarEvents(token, startISO, endISO);
    const bookedSlots = eventsToBookedSlots(events);
    return { statusCode: 200, headers, body: JSON.stringify({ bookedSlots }) };
  } catch (err) {
    console.error('Availability error:', err.message);
    return { statusCode: 200, headers, body: JSON.stringify({ bookedSlots: {}, _notice: 'Calendar temporarily unavailable' }) };
  }
};
