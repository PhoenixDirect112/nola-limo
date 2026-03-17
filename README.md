# NOLA Premier Limo — Website

Dark luxury private chauffeur website for New Orleans.
Built for Netlify with Square payments and Microsoft Outlook calendar integration.

---

## Tech Stack

| Layer       | Technology |
|-------------|-----------|
| Frontend    | Vanilla HTML / CSS / JS (no framework — fast, portable) |
| Hosting     | Netlify (static + serverless functions) |
| Calendar    | Microsoft Graph API → Outlook (Azure App Registration) |
| Payments    | Square Web Payments SDK + Square Payments API |
| Backend     | Netlify Functions (Node.js serverless) |

---

## Project Structure

```
nola-limo/
├── index.html              ← Homepage
├── booking.html            ← Step 1: Date/time + trip details
├── checkout.html           ← Step 2: Square payment
├── confirmation.html       ← Step 3: Booking confirmed
├── css/
│   └── global.css          ← All shared styles
├── netlify/
│   └── functions/
│       ├── availability.js    ← GET  /api/availability
│       ├── process-payment.js ← POST /api/process-payment
│       ├── square-config.js   ← GET  /api/square-config
│       └── get-price.js       ← POST /api/get-price
├── netlify.toml            ← Build config + security headers
├── .env.example            ← Environment variable template
└── .gitignore
```

---

## Setup Guide

### Step 1 — Clone and deploy to Netlify

1. Push this folder to a GitHub/GitLab repo
2. Go to [app.netlify.com](https://app.netlify.com) → **Add new site** → **Import from Git**
3. Set **Build command**: *(leave blank — no build step)*
4. Set **Publish directory**: `.`
5. Click **Deploy**

---

### Step 2 — Set environment variables in Netlify

Go to **Site → Configuration → Environment Variables** and add:

```
SQUARE_ACCESS_TOKEN        ← From Square Developer Dashboard
SQUARE_APPLICATION_ID      ← From Square Developer Dashboard (public app ID)
SQUARE_LOCATION_ID         ← From Square Developer Dashboard
SQUARE_ENVIRONMENT         ← "sandbox" for testing, "production" when live

AZURE_CLIENT_ID            ← From Azure App Registration
AZURE_CLIENT_SECRET        ← From Azure App Registration
AZURE_TENANT_ID            ← From Azure Active Directory
OUTLOOK_USER_EMAIL         ← Driver's Microsoft 365 email (e.g. driver@domain.com)

SITE_URL                   ← Your Netlify URL: https://your-site.netlify.app
```

---

### Step 3 — Configure Square

1. Go to [developer.squareup.com](https://developer.squareup.com) → Create app
2. Copy **Application ID**, **Access Token**, **Location ID** → add to Netlify env vars
3. Under **OAuth → Redirect URLs**, add: `https://your-site.netlify.app`
4. **Sandbox mode**: Use sandbox credentials and `sandbox.web.squarecdn.com` (already set)
5. **Go live**: Change `SQUARE_ENVIRONMENT=production` and update the Square SDK URL in `checkout.html`:
   - Change: `https://sandbox.web.squarecdn.com/v1/square.js`
   - To:     `https://web.squarecdn.com/v1/square.js`

---

### Step 4 — Configure Microsoft Azure (Outlook Calendar)

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name: `NOLA Limo Calendar`, Supported account type: **Single tenant**
3. After registration, go to **Certificates & secrets** → **New client secret** → Copy the value
4. Go to **API permissions** → **Add permission** → **Microsoft Graph** → **Application permissions**:
   - Add: `Calendars.ReadWrite`
   - Add: `User.Read.All`
5. Click **Grant admin consent**
6. Copy **Application (client) ID** → `AZURE_CLIENT_ID`
7. Copy **Directory (tenant) ID** → `AZURE_TENANT_ID`
8. Copy the client secret value → `AZURE_CLIENT_SECRET`
9. Set `OUTLOOK_USER_EMAIL` to the driver's Microsoft 365 email address

> **Important**: The driver's Microsoft 365 account must be in the same Azure AD tenant.
> If the driver uses a personal Outlook.com account (not Microsoft 365), you'll need
> to use the Authorization Code flow instead of Client Credentials.

---

### Step 5 — Customize Content

| File           | What to update |
|---------------|----------------|
| `index.html`  | Business name, phone number, email, testimonials, hero text |
| `booking.html`| Service types, time slots |
| `confirmation.html` | Driver phone number |
| `netlify/functions/get-price.js` | Pricing per service type |
| `netlify.toml` | Replace `yourdomain.com` with actual domain |
| All files     | Replace `(504) 123-4567` with real phone number |

**Add your vehicle photo:**
In `index.html`, find the `.vehicle-visual` div and replace the placeholder with:
```html
<img src="images/suburban.jpg" alt="2023 Chevrolet Suburban" style="width:100%;height:100%;object-fit:cover;border-radius:4px;" />
```

---

## Security Features

| Feature | Implementation |
|---------|---------------|
| HTTPS enforced | `netlify.toml` HSTS header |
| Clickjacking prevention | `X-Frame-Options: DENY` |
| MIME sniffing prevention | `X-Content-Type-Options: nosniff` |
| Content Security Policy | Allows only Square, Google Fonts, and self |
| XSS prevention | All booking data inserted via `textContent`, never `innerHTML` |
| Input validation | Server-side validation of all booking fields |
| Input sanitization | All strings sanitized — HTML chars stripped |
| Rate limiting | 30 req/min per IP on availability endpoint |
| No secrets in browser | Square Access Token stays server-side |
| Replay attack prevention | Per-booking nonces, single-use |
| Session expiry | Booking sessions expire after 30 minutes |
| No sensitive data in logs | Card data tokenized by Square before leaving browser |
| PCI compliance | Square handles all card data; we only receive tokens |

---

## Local Development

```bash
# Install Netlify CLI
npm install -g netlify-cli

# Copy env vars
cp .env.example .env.local
# Fill in your actual values in .env.local

# Start local dev server (serves functions too)
netlify dev
```

Visit: `http://localhost:8888`

---

## Pricing (update in `get-price.js`)

| Service        | Default Rate |
|---------------|-------------|
| Airport Transfer | $185.00 |
| Corporate Travel | $150.00 |
| Wedding         | $350.00 |
| Special Event   | $200.00 |
| Hourly Charter  | $125.00/hr base |

---

## Going Live Checklist

- [ ] All env vars set in Netlify dashboard
- [ ] Square switched from sandbox → production
- [ ] Square SDK URL updated in `checkout.html`
- [ ] Azure app granted admin consent for `Calendars.ReadWrite`
- [ ] Real phone number in all pages
- [ ] Real email in footer
- [ ] Vehicle photo added
- [ ] Business name updated throughout
- [ ] Custom domain connected in Netlify
- [ ] SSL certificate auto-provisioned (Netlify does this)
- [ ] Test full booking flow end-to-end
- [ ] Test calendar event appears in Outlook after booking
