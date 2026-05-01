# Integrations Setup Guide

End-to-end walkthrough for two related tasks:

1. **Custom-domain email** — `rides@<your-domain>.com` (or similar) hosted on the client's existing Microsoft 365 Business tenant, with the domain (registered at GoDaddy) pointed at Microsoft.
2. **Calendar integration** — Netlify functions read busy times and write new bookings to that mailbox's Outlook calendar in real time, via Microsoft Graph + an Azure AD app registration.

> **Order matters.** Do Phase 1 first — the mailbox has to exist before the Azure app can be pointed at it.

---

## Phase 1 — Add the custom domain to Microsoft 365

You need: admin login at <https://admin.microsoft.com>, and GoDaddy login at <https://account.godaddy.com>.

> **GoDaddy shortcut.** When you add a domain in M365 Admin and Microsoft detects the domain is at GoDaddy, it offers a "Connect to GoDaddy" / "Set up automatically" flow that signs into GoDaddy on your behalf and writes most DNS records for you. If that works, you can skip 1.2 and 1.4 below — just sign in when prompted, approve, and Microsoft does the rest. If it fails or you prefer manual control, follow the steps below.

### 1.1 Start the domain add in Microsoft 365 Admin

1. Sign in at <https://admin.microsoft.com> with the client's M365 admin account.
2. Left nav → **Settings** → **Domains** → **Add domain**.
3. Type the custom domain (e.g. `wmsdeliveryservices.com`) → **Use this domain**.
4. When prompted "How do you want to verify the domain?", choose **Add a TXT record to the domain's DNS records**.
5. Microsoft shows a TXT record like:
   - **Type:** `TXT`
   - **Host/Name:** `@`
   - **Value:** `MS=ms########`
   - **TTL:** `1 hour` (or default)

Leave this page open — you'll come back to it to verify.

### 1.2 Add the verification record at GoDaddy

1. Sign in at <https://account.godaddy.com> → **My Products** → find the domain → click the **DNS** button (or the three-dot menu → **Manage DNS**).
2. You'll land on the **DNS Management** page for the domain.
3. Scroll to the **Records** section → click **Add New Record** (or **Add** at the top right).
   - **Type:** `TXT`
   - **Name:** `@`
   - **Value:** `MS=ms########` (the exact value Microsoft showed)
   - **TTL:** `1 Hour` (default is fine)
4. Click **Save**.
5. Wait 5–15 minutes for DNS to propagate.

> **GoDaddy quirk:** GoDaddy creates a "Parked" CNAME and a default `_domainconnect` record automatically. Leave those alone — they don't conflict with M365.

### 1.3 Verify the domain in Microsoft 365

1. Back in M365 Admin → continue the wizard → **Verify**.
2. If it fails, wait another 10 minutes and retry. (You can confirm the TXT propagated using <https://mxtoolbox.com/TXTLookup.aspx>.)

### 1.4 Add the Microsoft DNS records at GoDaddy

After verification, Microsoft shows the full set of records to add. Typically:

| Type  | Name (GoDaddy) | Value                                            | Priority | Purpose            |
|-------|----------------|--------------------------------------------------|----------|---------------------|
| MX    | `@`            | `<your-domain>-com.mail.protection.outlook.com`  | `0`      | Inbound email      |
| TXT   | `@`            | `v=spf1 include:spf.protection.outlook.com -all` | —        | SPF (anti-spoofing) |
| CNAME | `autodiscover` | `autodiscover.outlook.com`                       | —        | Outlook auto-config |
| TXT   | `_dmarc`       | `v=DMARC1; p=none; rua=mailto:dmarc@<your-domain>.com` | —  | DMARC (recommended) |
| CNAME | `selector1._domainkey` | `selector1-<your-domain>-com._domainkey.<tenant>.onmicrosoft.com` | — | DKIM (set up later in Defender) |
| CNAME | `selector2._domainkey` | `selector2-<your-domain>-com._domainkey.<tenant>.onmicrosoft.com` | — | DKIM (set up later in Defender) |

> **GoDaddy MX quirk:** GoDaddy ships every domain with a placeholder MX record (priority 0, value something like `smtp.secureserver.net`). **Delete it** before adding the Microsoft MX, otherwise inbound mail will be intercepted by GoDaddy's email forwarding rather than reach M365.

> **Email Forwarding feature:** if the domain has GoDaddy "Email Forwarding" enabled (a free add-on some packages turn on by default), turn it off under **My Products → Email & Office** for this domain. It conflicts with the new M365 mailbox.

> **Don't enable Skype/Teams DNS records** unless the client uses them — Microsoft sometimes lists them. Add only Email + AutoDiscover + SPF + DMARC + DKIM.

In Microsoft, click **Verify** / **Continue**. If a record fails, double-check exact values — GoDaddy automatically appends the domain to relative names (so `autodiscover` becomes `autodiscover.<your-domain>.com`); use the bare label without trailing dots.

### 1.5 Create the user mailbox

1. M365 Admin → **Users** → **Active users** → **Add a user**.
2. **Username:** `rides` (or `info`, `bookings`, etc.) → select the new domain in the dropdown.
3. Set name, password, and assign a license. The cheapest plan that includes mailbox + Calendar API access is **Microsoft 365 Business Basic** (~$6/user/month).
4. Save.

The new mailbox `rides@<your-domain>.com` is now live. Send a test email to it from a personal address to confirm delivery; reply from Outlook on the web to confirm sending works.

### 1.6 (Optional) Forwarding & shared mailbox notes

- If the client wants emails to also drop into their personal Outlook, set **Mail forwarding** under the user's properties.
- If multiple people will read the inbox, consider converting it to a **Shared mailbox** later (free, no license needed) — but for the booking calendar API to work, **the mailbox must be a licensed user, not a shared mailbox**, so leave it as a licensed user for now.

---

## Phase 2 — Azure AD app registration (calendar API access)

This creates the credentials the Netlify functions use to read/write the mailbox's calendar via Microsoft Graph.

### 2.1 Create the app

1. Sign in at <https://portal.azure.com> with the same M365 admin account.
2. Top search bar → **Azure Active Directory** (or **Microsoft Entra ID** in the new UI) → **App registrations** → **New registration**.
3. **Name:** `WMS Booking Calendar` (cosmetic).
4. **Supported account types:** **Accounts in this organizational directory only** (single tenant).
5. **Redirect URI:** leave blank.
6. **Register**.

On the app's Overview page, copy these two values — you'll need them for Netlify env vars:

- **Application (client) ID** → goes into `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → goes into `AZURE_TENANT_ID`

### 2.2 Create a client secret

1. Left nav inside the app → **Certificates & secrets** → **Client secrets** → **New client secret**.
2. Description: `Netlify functions`. Expiry: **24 months** (calendar a reminder 30 days before).
3. **Add**.
4. **Copy the `Value` immediately** — it's only shown once. This goes into `AZURE_CLIENT_SECRET`.

> If you miss it, just delete the secret and create a new one.

### 2.3 Grant Microsoft Graph permissions

1. Left nav → **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.
2. Add these three:
   - `Calendars.ReadWrite` — read busy times, create new events.
   - `User.Read.All` — needed for the SDK to look up the mailbox owner.
   - `MailboxSettings.Read` — only if you later expand the integration; safe to add now.
3. Back on the **API permissions** screen, click **Grant admin consent for [tenant name]**. The status column should turn green ("Granted for...").

> **Without admin consent, Graph calls will fail with `Insufficient privileges`.** This is the most common setup mistake.

### 2.4 (Optional) Restrict the app to one mailbox

By default, an `Application` permission like `Calendars.ReadWrite` lets the app read every mailbox in the tenant. For a single-mailbox booking system, lock it down with an **Application Access Policy** so it can only touch `rides@<your-domain>.com`.

Run these in **PowerShell** as a tenant admin (one time, ~5 minutes):

```powershell
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName admin@<your-domain>.com

New-ApplicationAccessPolicy `
  -AppId <AZURE_CLIENT_ID> `
  -PolicyScopeGroupId rides@<your-domain>.com `
  -AccessRight RestrictAccess `
  -Description "Booking app may only access rides mailbox"
```

Test:
```powershell
Test-ApplicationAccessPolicy -Identity rides@<your-domain>.com -AppId <AZURE_CLIENT_ID>
# Expect: Granted

Test-ApplicationAccessPolicy -Identity admin@<your-domain>.com -AppId <AZURE_CLIENT_ID>
# Expect: Denied
```

---

## Phase 3 — Wire credentials into Netlify

The functions in `netlify/functions/` already read from environment variables — you just need to set them in the Netlify dashboard.

1. Sign in at <https://app.netlify.com> → select the site → **Site configuration** → **Environment variables**.
2. **Add a variable** for each of:

   | Key                   | Value                              |
   |-----------------------|------------------------------------|
   | `AZURE_TENANT_ID`     | from app Overview (Phase 2.1)      |
   | `AZURE_CLIENT_ID`     | from app Overview (Phase 2.1)      |
   | `AZURE_CLIENT_SECRET` | from secret value (Phase 2.2)      |
   | `OUTLOOK_USER_EMAIL`  | `rides@<your-domain>.com`          |
   | `SITE_URL`            | the live site URL, e.g. `https://nolalimo.netlify.app` (used for CORS) |

3. Trigger a redeploy: **Deploys** → **Trigger deploy** → **Deploy site**. Env var changes only take effect on a fresh build.

> Square keys (`SQUARE_ACCESS_TOKEN`, `SQUARE_LOCATION_ID`, `SQUARE_ENVIRONMENT`) are separate and unrelated — don't touch those during this work.

---

## Phase 4 — Update the site to reflect the new email

There's still a placeholder email in `index.html` that needs to point at the new mailbox. Once you've confirmed the new mailbox is live, share the address and I'll find every occurrence (footer mailto link, any `rides@nolapremierlimo.com`, etc.) and swap it.

I'll also recheck:
- Netlify Forms notification email (under **Forms** → **Notifications** in Netlify) — switch to the new address so the client gets booking submissions.
- Any contact-form `mailto:` fallbacks.

---

## Phase 5 — Test end-to-end

After Phases 1–4 are done:

1. Visit the live site → **Booking** page.
2. Watch the network tab → calendar should fetch from `/api/availability?year=…&month=…` and return `200 OK` with a `bookedSlots` payload (empty `{}` is fine — that just means no events yet).
3. Manually create a 2-hour test event in the `rides@…` Outlook calendar (any date next week, 10:00 AM–12:00 PM).
4. Reload the booking page → that day's 10 AM and 11 AM slots should show as unavailable.
5. Make a real test booking using a real card (or Square sandbox if `SQUARE_ENVIRONMENT=sandbox`):
   - Confirm the Square dashboard shows the charge with the correct WMS note.
   - Confirm the Outlook calendar gets a new event with the booking details in the body.
6. Delete the test event and refund the test charge.

---

## Troubleshooting

| Symptom                                                | Likely cause                                    |
|--------------------------------------------------------|-------------------------------------------------|
| Calendar always shows fully open, even with real events| Env vars not set, or Netlify hasn't redeployed since they were added |
| `401 Unauthorized` in Netlify function logs            | Wrong `AZURE_CLIENT_SECRET`, or secret expired   |
| `403 Forbidden` from Graph API                         | Admin consent not granted (Phase 2.3 last step) |
| `Resource could not be found` from Graph               | `OUTLOOK_USER_EMAIL` typo, or mailbox not licensed |
| Domain verification fails repeatedly at Microsoft      | TXT record at GoDaddy has wrong Name (`@` vs the literal domain), or hasn't propagated yet — wait 30 min |
| Email arrives but can't send out                       | SPF record missing or wrong — recheck Phase 1.4 |
| Calendar reads work but new bookings don't appear      | App is missing `Calendars.ReadWrite` (only has `.Read`) |

---

## What I need from you when ready

To wire the code/config side I'll need:

- The **custom domain** (e.g. `wmsdeliveryservices.com`)
- The chosen **mailbox username** (`rides@…`, `info@…`, `bookings@…`)
- Confirmation when env vars are set in Netlify so I know to do a smoke test of the deployed functions.

Once you have those, I'll update `index.html` and any other code references to the new email address, and double-check the Netlify Forms destination.
