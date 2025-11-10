# Network Setup Guide for Printer API

## Problem
- **funPrinting** runs on Render (cloud) at `https://your-app.onrender.com`
- **Printer API** runs locally on `http://localhost:3001`
- They cannot communicate directly because localhost is not accessible from the internet

## Solutions

### Option 1: Use ngrok (Recommended for Development/Testing)

**ngrok** creates a secure tunnel from the internet to your local server.

#### Installation
```bash
# macOS
brew install ngrok

# Or download from https://ngrok.com/download
```

#### Setup
1. **Sign up for free ngrok account** at https://ngrok.com
2. **Get your authtoken** from https://dashboard.ngrok.com/get-started/your-authtoken
3. **Configure ngrok:**
   ```bash
   ngrok config add-authtoken YOUR_AUTH_TOKEN
   ```

#### Start Tunnel
```bash
# Start printer API
cd printer-api
npm start

# In another terminal, start ngrok tunnel
ngrok http 3001
```

#### Get Public URL
ngrok will give you a public URL like:
```
Forwarding: https://abc123.ngrok-free.app -> http://localhost:3001
```

#### Configure funPrinting
Update `.env.local` in funPrinting:
```env
PRINTER_API_URLS=["https://abc123.ngrok-free.app"]
PRINTER_API_KEY=your_api_key_here
```

**Note:** Free ngrok URLs change every time you restart. For production, use ngrok paid plan with static domain.

---

### Option 2: Use localtunnel (Free Alternative)

**localtunnel** is a free alternative to ngrok.

#### Installation
```bash
npm install -g localtunnel
```

#### Start Tunnel
```bash
# Start printer API
cd printer-api
npm start

# In another terminal, start localtunnel
lt --port 3001 --subdomain your-unique-name
```

#### Get Public URL
localtunnel will give you a public URL like:
```
your url is: https://your-unique-name.loca.lt
```

#### Configure funPrinting
Update `.env.local` in funPrinting:
```env
PRINTER_API_URLS=["https://your-unique-name.loca.lt"]
PRINTER_API_KEY=your_api_key_here
```

**Note:** Free localtunnel URLs may change. Use `--subdomain` for more stable URLs.

---

### Option 3: Use Cloudflare Tunnel (Free, More Stable)

**Cloudflare Tunnel** (formerly Argo Tunnel) provides a free, stable tunnel.

#### Installation
```bash
# Download cloudflared from https://developers.cloudflare.com/cloudflare-one/connections/connect-apps/install-and-setup/installation/
```

#### Setup
1. **Login to Cloudflare:**
   ```bash
   cloudflared tunnel login
   ```

2. **Create tunnel:**
   ```bash
   cloudflared tunnel create printer-api
   ```

3. **Configure tunnel:**
   Create `config.yml`:
   ```yaml
   tunnel: printer-api
   credentials-file: /path/to/credentials.json
   
   ingress:
     - hostname: printer-api.yourdomain.com
       service: http://localhost:3001
     - service: http_status:404
   ```

4. **Start tunnel:**
   ```bash
   cloudflared tunnel run printer-api
   ```

#### Configure funPrinting
Update `.env.local` in funPrinting:
```env
PRINTER_API_URLS=["https://printer-api.yourdomain.com"]
PRINTER_API_KEY=your_api_key_here
```

---

### Option 4: Use VPN/Private Network

If both services are on the same private network:

1. **Find your local IP address:**
   ```bash
   # macOS/Linux
   ifconfig | grep "inet "
   
   # Windows
   ipconfig
   ```

2. **Example:** Your IP is `192.168.1.100`

3. **Configure printer API to listen on all interfaces:**
   Update `server.ts` to listen on `0.0.0.0`:
   ```typescript
   app.listen(PORT, '0.0.0.0', () => {
     console.log(`Server running on http://0.0.0.0:${PORT}`);
   });
   ```

4. **Configure funPrinting:**
   Update `.env.local` in funPrinting:
   ```env
   PRINTER_API_URLS=["http://192.168.1.100:3001"]
   PRINTER_API_KEY=your_api_key_here
   ```

**Note:** This only works if Render can access your private network (usually not possible).

---

### Option 5: Deploy Printer API to a Server

Deploy the printer API to a server accessible from Render:

1. **Use a VPS** (DigitalOcean, AWS EC2, etc.)
2. **Install Node.js and dependencies**
3. **Deploy printer API**
4. **Configure firewall** to allow port 3001
5. **Use public IP or domain**

#### Example with VPS
```env
# funPrinting .env.local
PRINTER_API_URLS=["http://YOUR_VPS_IP:3001"]
# or
PRINTER_API_URLS=["https://printer-api.yourdomain.com"]
```

---

## Recommended Setup for Production

### For Production Use:
1. **Use Cloudflare Tunnel** (free, stable, secure)
2. **Or deploy printer API to VPS** (most reliable)
3. **Use static domain** for printer API

### For Development/Testing:
1. **Use ngrok** (easy setup, free tier available)
2. **Or use localtunnel** (completely free)

---

## Testing Connection

### Test from funPrinting
```bash
# Test health endpoint
curl https://your-printer-api-url.com/health

# Test queue status (requires API key)
curl -H "X-API-Key: your_api_key" https://your-printer-api-url.com/api/queue/status
```

### Test from Admin Panel
1. Go to `/admin/printer-monitor`
2. Check if printer API status shows "Healthy"
3. Monitor queue status

---

## Troubleshooting

### Connection Refused
- Check if printer API is running
- Check if tunnel is active
- Verify URL in funPrinting `.env.local`

### Timeout Errors
- Check firewall settings
- Verify tunnel is forwarding correctly
- Check printer API logs

### Authentication Errors
- Verify `API_KEY` matches in both `.env` files
- Check headers are being sent correctly

---

## Security Notes

1. **Always use HTTPS** in production (ngrok, Cloudflare provide this)
2. **Use strong API keys** (generate with: `openssl rand -hex 32`)
3. **Restrict access** if possible (IP whitelist, VPN)
4. **Monitor logs** for unauthorized access attempts

---

## Quick Start (ngrok)

```bash
# 1. Install ngrok
brew install ngrok  # macOS
# or download from https://ngrok.com

# 2. Sign up and get authtoken
# Visit https://ngrok.com and get your authtoken

# 3. Configure
ngrok config add-authtoken YOUR_AUTH_TOKEN

# 4. Start printer API
cd printer-api
npm start

# 5. In another terminal, start tunnel
ngrok http 3001

# 6. Copy the forwarding URL (e.g., https://abc123.ngrok-free.app)

# 7. Update funPrinting .env.local
PRINTER_API_URLS=["https://abc123.ngrok-free.app"]
PRINTER_API_KEY=your_api_key_here

# 8. Restart funPrinting
```

---

**For more help, see the main README.md**

