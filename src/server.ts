import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import printRoutes from './routes/print';
import { authenticateApiKey } from './utils/auth';

// Load environment variables
dotenv.config();

const app = express();
const PORT = parseInt(process.env.PORT || '3001', 10);

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Request logging
app.use((req, res, next) => {
  console.log(`${new Date().toISOString()} - ${req.method} ${req.path}`);
  next();
});

// Health check (no auth required)
app.get('/health', async (req, res) => {
  res.json({
    success: true,
    status: 'healthy',
    timestamp: new Date().toISOString(),
    service: 'printer-api'
  });
});

// API routes with authentication
app.use('/api', printRoutes);

// Error handling
app.use((err: Error, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error('Error:', err);
  res.status(500).json({
    success: false,
    error: err.message || 'Internal server error'
  });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Route not found'
  });
});

// Start server
const HOST = process.env.HOST || '0.0.0.0'; // Listen on all interfaces for network access
app.listen(PORT, HOST, () => {
  console.log(`🖨️  Printer API server running on ${HOST}:${PORT}`);
  console.log(`📋 Environment: ${process.env.NODE_ENV || 'development'}`);
  console.log(`🔑 API Key authentication: ${process.env.API_KEY ? 'Enabled' : 'Disabled'}`);
  console.log(`🖨️  Printer: ${process.env.PRINTER_NAME || 'Not configured'}`);
  console.log(`🌐 Local URL: http://localhost:${PORT}`);
  console.log(`🌐 Network URL: http://0.0.0.0:${PORT}`);
  console.log(`💡 To expose to internet, use ngrok: ngrok http ${PORT}`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully...');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT received, shutting down gracefully...');
  process.exit(0);
});

