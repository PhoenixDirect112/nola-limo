// netlify/functions/get-price.js
// ─────────────────────────────────────────────────────────────────
//  POST /api/get-price
//  Returns a price estimate based on service type.
//  In a production build, this can call a distance matrix API
//  (Google Maps, Mapbox) to calculate mileage-based pricing.
// ─────────────────────────────────────────────────────────────────

const PRICES = {
  airport:   18500,  // $185.00
  corporate: 15000,  // $150.00
  wedding:   35000,  // $350.00
  event:     20000,  // $200.00
  hourly:    12500,  // $125.00/hr base
};

exports.handler = async function(event) {
  const headers = {
    'Content-Type':              'application/json',
    'Cache-Control':             'no-store',
    'X-Content-Type-Options':    'nosniff',
    'Access-Control-Allow-Origin': process.env.SITE_URL || 'https://your-site.netlify.app',
  };

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid JSON' }) }; }

  const validServices = Object.keys(PRICES);
  const serviceType = body.serviceType;

  if (!serviceType || !validServices.includes(serviceType)) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid service type' }) };
  }

  const amountCents = PRICES[serviceType];

  return {
    statusCode: 200,
    headers,
    body: JSON.stringify({
      amountCents,
      displayAmount: `$${(amountCents / 100).toFixed(2)}`,
      serviceType,
      note: 'Base rate. Final pricing confirmed before departure.',
    }),
  };
};
