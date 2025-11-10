# Printer API Server

Local printer API server for USB printer integration with delivery number system.

## Features

- USB printer integration via system print commands
- Print queue management with infinite retry
- Delivery number generation (A-Z cycle: 260 files total)
- Blank page separator after every file
- Full-page letter separator after every 10 files
- API key authentication
- Health check endpoint
- Persistent queue storage

## Prerequisites

- Node.js 18+ installed
- USB printer connected to your computer
- Printer configured in your operating system

## Installation

1. **Clone or download this repository**
   ```bash
   git clone https://github.com/AdityaPandey-DEV/Printer-API.git
   cd Printer-API
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Create `.env` file**
   ```bash
   cp .env.example .env
   # Or create .env manually
   ```

4. **Configure `.env` file**
   ```env
   PORT=3001
   PRINTER_NAME=HP_Deskjet_525
   PRINTER_PATH=/dev/usb/lp0
   PRINT_QUEUE_SIZE=100
   DELIVERY_NUMBER_START=A
   API_KEY=your_secure_api_key_here
   NODE_ENV=production
   ```

   **Important Configuration:**
   - `PRINTER_NAME`: Name of your printer as configured in your OS
     - **Windows**: Check in Control Panel > Devices and Printers
     - **macOS**: Check in System Preferences > Printers & Scanners
     - **Linux**: Check with `lpstat -p` command
   
   - `PRINTER_PATH`: Path to printer (usually not needed, but can be set)
     - **Windows**: Leave empty or use `COM3` format
     - **macOS/Linux**: Use `/dev/usb/lp0` or similar
   
   - `API_KEY`: Generate a secure random string (use this in funPrinting)
     ```bash
     # Generate a secure API key (example)
     node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
     ```

## Running the API

### Development Mode
```bash
npm run dev
```

### Production Mode
```bash
# Build TypeScript
npm run build

# Start server
npm start
```

The server will start on `http://localhost:3001` (or the port specified in `.env`)

## Getting Your API Key

1. **Generate a secure API key:**
   ```bash
   # Option 1: Using Node.js
   node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
   
   # Option 2: Using OpenSSL
   openssl rand -hex 32
   
   # Option 3: Using online generator
   # Visit: https://www.random.org/strings/
   ```

2. **Add to `.env` file:**
   ```env
   API_KEY=your_generated_api_key_here
   ```

3. **Restart the server** after updating `.env`

## API Endpoints

### POST /api/print
Add a print job to the queue.

**Headers:**
```
X-API-Key: your_api_key_here
```
or
```
Authorization: Bearer your_api_key_here
```

**Body:**
```json
{
  "fileUrl": "https://example.com/file.pdf",
  "fileName": "document.pdf",
  "fileType": "application/pdf",
  "printingOptions": {
    "pageSize": "A4",
    "color": "bw",
    "sided": "single",
    "copies": 1
  },
  "printerIndex": 1
}
```

**Response:**
```json
{
  "success": true,
  "message": "Print job added to queue",
  "jobId": "job_1234567890_abc123",
  "deliveryNumber": "A202501151"
}
```

### GET /api/queue/status
Get queue status.

**Headers:**
```
X-API-Key: your_api_key_here
```

**Response:**
```json
{
  "success": true,
  "total": 5,
  "pending": 2,
  "jobs": [...]
}
```

### GET /api/health
Health check endpoint (no authentication required).

**Response:**
```json
{
  "success": true,
  "status": "healthy",
  "printer": {
    "available": true,
    "message": "Printer is available"
  },
  "queue": {
    "total": 0,
    "pending": 0
  },
  "timestamp": "2025-01-15T10:30:00.000Z"
}
```

## Delivery Number System

**Format:** `{LETTER}{YYYYMMDD}{PRINTER_INDEX}`

**Example:** `A202501151` = Letter A, Date 2025-01-15, Printer Index 1

**Logic:**
- Cycles through A-Z (260 files total: 10 files per letter × 26 letters)
- After every file: prints blank page separator
- After every 10 files: prints full-page letter separator (A, B, C...)
- Date format: YYYYMMDD
- Resets letter and count when date changes

## Print Queue

- Jobs are queued and processed sequentially
- **Infinite retry on failure** (payment is done, so printing must succeed)
- Persistent queue stored in `print-queue.json`
- Exponential backoff between retries (max 5 minutes)
- Queue survives server restarts

## How to Configure in funPrinting

### Step 1: Get Your Printer API URL

1. **Find your computer's IP address:**
   ```bash
   # Windows
   ipconfig
   
   # macOS/Linux
   ifconfig
   # or
   ip addr show
   ```

2. **Your Printer API URL will be:**
   ```
   http://YOUR_IP_ADDRESS:3001
   # Example: http://192.168.1.100:3001
   ```

   **Note:** If running on the same machine as funPrinting, use:
   ```
   http://localhost:3001
   ```

### Step 2: Configure funPrinting

1. **Open funPrinting `.env.local` file**

2. **Add printer API configuration:**
   ```env
   # Printer API Configuration
   PRINTER_API_URLS=["http://localhost:3001","http://192.168.1.100:3001"]
   PRINTER_API_TIMEOUT=5000
   PRINTER_API_KEY=your_secure_api_key_here
   INVOICE_ENABLED=true
   RETRY_QUEUE_ENABLED=true
   ```

   **Important:**
   - `PRINTER_API_URLS`: Array of printer API URLs (JSON format)
     - Add multiple URLs if you have multiple printers
     - First URL is used by default
   - `PRINTER_API_KEY`: **Must match** the `API_KEY` in printer-api `.env`
   - `PRINTER_API_TIMEOUT`: Request timeout in milliseconds

3. **Save and restart funPrinting**

### Step 3: Test the Connection

1. **Start Printer API:**
   ```bash
   cd printer-api
   npm start
   ```

2. **Test health endpoint:**
   ```bash
   curl http://localhost:3001/health
   ```

3. **Test from funPrinting:**
   - Place a test order
   - Complete payment
   - Check printer API logs for print job

## Troubleshooting

### Printer Not Found
- **Check printer name:**
  ```bash
  # macOS/Linux
  lpstat -p
  
  # Windows
  # Control Panel > Devices and Printers
  ```
- Update `PRINTER_NAME` in `.env` to match exactly

### API Key Authentication Failed
- Ensure `API_KEY` in printer-api `.env` matches `PRINTER_API_KEY` in funPrinting `.env.local`
- Check headers are being sent correctly

### Print Jobs Not Processing
- Check printer is online and has paper
- Check `print-queue.json` file exists and has jobs
- Check server logs for errors
- Verify printer API is accessible from funPrinting server

### Network Issues
- **If funPrinting is on Render (cloud) and printer-api is local:**
  - Use ngrok or similar tunnel service
  - Or use VPN/private network
  - Or deploy printer-api to a server accessible from Render

### Queue Not Persisting
- Ensure `print-queue.json` file has write permissions
- Check disk space
- Verify `NODE_ENV` is set correctly

## Security Notes

- **Never commit `.env` file** to git
- Use strong, random API keys
- Consider using HTTPS in production
- Restrict network access to printer API if possible
- Use firewall rules to limit access

## Development

### Project Structure
```
printer-api/
├── src/
│   ├── server.ts          # Express server setup
│   ├── routes/
│   │   └── print.ts        # Print job endpoints
│   ├── services/
│   │   ├── printer.ts      # USB printer service
│   │   ├── deliveryNumber.ts # Delivery number logic
│   │   └── queue.ts        # Print queue management
│   ├── models/
│   │   └── PrintJob.ts     # Print job model
│   └── utils/
│       ├── printerUtils.ts # Printer utilities
│       └── auth.ts         # API key authentication
├── dist/                   # Compiled JavaScript
├── package.json
├── tsconfig.json
└── README.md
```

### Building
```bash
npm run build
```

### Running Tests
```bash
npm test
```

## License

ISC

## Support

For issues or questions:
- Open an issue on GitHub
- Contact: adityapandey.dev.in@gmail.com

---

**Made with ❤️ for funPrinting Project**
