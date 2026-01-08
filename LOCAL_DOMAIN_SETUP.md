# Setting Up local.tansu.co for Secure Add-in Communication

This document explains how to set up `local.tansu.co` to enable secure HTTPS communication between the Word add-in (hosted on GitHub Pages) and the Tansu desktop app.

## Why This Is Needed

- Microsoft AppSource requires add-ins to be served over HTTPS
- HTTPS pages cannot make HTTP requests to localhost (mixed content blocking)
- Solution: Use a real domain `local.tansu.co` that resolves to `127.0.0.1` with a valid SSL certificate

## Step 1: Add DNS Record

Add an A record in your DNS settings for tansu.co:

```
Type: A
Name: local
Value: 127.0.0.1
TTL: 3600 (or lowest available)
```

This makes `local.tansu.co` resolve to `127.0.0.1` (localhost).

## Step 2: Get SSL Certificate

You need a valid SSL certificate for `local.tansu.co`. Since the domain points to localhost, you can't use HTTP-01 challenge. Use DNS-01 challenge instead:

### Option A: Using Cloudflare (if your DNS is on Cloudflare)
Cloudflare can auto-generate certificates for subdomains.

### Option B: Using Let's Encrypt with DNS-01 challenge
```bash
# Install certbot
brew install certbot

# Get certificate using DNS challenge
sudo certbot certonly --manual --preferred-challenges dns -d local.tansu.co

# Follow prompts to add TXT record to DNS
# Certificates will be saved to /etc/letsencrypt/live/local.tansu.co/
```

### Option C: Using acme.sh (easier for automation)
```bash
# Install acme.sh
curl https://get.acme.sh | sh

# Get certificate (replace with your DNS provider API)
acme.sh --issue --dns dns_cf -d local.tansu.co
```

## Step 3: Bundle Certificate with Tansu

Copy the certificate files to the Tansu app bundle:
- `fullchain.pem` (certificate + chain)
- `privkey.pem` (private key)

Place them in `word-addin/certs/` directory.

## Step 4: Update API Server

The API server needs to:
1. Serve on `0.0.0.0:5050` (not just 127.0.0.1)
2. Use the SSL certificate for `local.tansu.co`
3. Handle CORS for the GitHub Pages domain

## Certificate Renewal

Let's Encrypt certificates expire every 90 days. Set up auto-renewal:
```bash
# Add to crontab
0 0 1 * * certbot renew --quiet
```

For the desktop app, you'll need to bundle updated certificates with each release or implement certificate auto-update.

## Testing

1. Add DNS record
2. Verify: `dig local.tansu.co` should return `127.0.0.1`
3. Get certificate
4. Start Tansu with SSL
5. Visit `https://local.tansu.co:5050/ping` - should show no certificate warnings

## Alternative: Wildcard Certificate

Instead of a certificate just for `local.tansu.co`, get a wildcard for `*.tansu.co`:
```bash
certbot certonly --manual --preferred-challenges dns -d "*.tansu.co"
```

This covers all subdomains including `local.tansu.co`.
