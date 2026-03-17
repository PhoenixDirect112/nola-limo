// netlify/functions/square-config.js
// ─────────────────────────────────────────────────────────────────
//  GET /api/square-config
//  Returns the Square Application ID and Location ID to the client.
//  The ACCESS TOKEN stays server-side — only the public app ID is
//  sent to the browser (this is safe and required by Square SDK).
// ─────────────────────────────────────────────────────────────────

exports.handler = async function(event) {
  const headers = {
    'Content-Type':              'application/json',
    'Cache-Control':             'public, max-age=300',
    'X-Content-Type-Options':    'nosniff',
    'Access-Control-Allow-Origin': process.env.SITE_URL || 'https://your-site.netlify.app',
  };

  if (event.httpMethod !== 'GET') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  const appId      = process.env.SQUARE_APPLICATION_ID;
  const locationId = process.env.SQUARE_LOCATION_ID;

  if (!appId || !locationId) {
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: 'Payment configuration not set up' }),
    };
  }

  return {
    statusCode: 200,
    headers,
    body: JSON.stringify({
      applicationId: appId,
      locationId:    locationId,
    }),
  };
};
