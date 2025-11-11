import { Router, Request, Response } from 'express';
import { addToQueue, getQueueStatus } from '../services/queue';
import { generateDeliveryNumber } from '../services/deliveryNumber';
import { checkPrinterStatus } from '../services/printer';
import { PrintJobData } from '../models/PrintJob';
import { authenticateApiKey } from '../utils/auth';

const router = Router();

// Apply authentication to all routes
router.use(authenticateApiKey);

/**
 * POST /api/print
 * Add a print job to the queue
 */
router.post('/print', async (req: Request, res: Response) => {
  try {
    const jobData: PrintJobData = req.body;
    const printerIndex = parseInt(req.body.printerIndex || '1', 10);

    // Check if multiple files exist
    const hasMultipleFiles = jobData.fileURLs && jobData.fileURLs.length > 0;
    const hasSingleFile = jobData.fileUrl && !hasMultipleFiles;

    // Validate required fields
    if (!hasMultipleFiles && !hasSingleFile) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields: Either fileUrl/fileName or fileURLs/originalFileNames are required'
      });
    }

    if (hasMultipleFiles) {
      // Multiple files: create separate print jobs for each file
      const fileURLs = jobData.fileURLs!;
      const originalFileNames = jobData.originalFileNames || fileURLs.map((_, idx) => `File ${idx + 1}`);
      const fileTypes = jobData.fileTypes || fileURLs.map(() => 'application/octet-stream');

      if (fileURLs.length !== originalFileNames.length) {
        return res.status(400).json({
          success: false,
          error: 'fileURLs and originalFileNames arrays must have the same length'
        });
      }

      const jobIds: string[] = [];
      const deliveryNumbers: string[] = [];

      // Create a print job for each file
      for (let i = 0; i < fileURLs.length; i++) {
        // Generate delivery number for each file
        const deliveryNumber = generateDeliveryNumber(printerIndex);

        // Create print job for this file
        const printJob = {
          fileUrl: fileURLs[i],
          fileName: originalFileNames[i],
          fileType: fileTypes[i] || 'application/octet-stream',
          printingOptions: {
            pageSize: jobData.printingOptions?.pageSize || 'A4',
            color: jobData.printingOptions?.color || 'bw',
            sided: jobData.printingOptions?.sided || 'single',
            copies: jobData.printingOptions?.copies || 1,
            pageCount: jobData.printingOptions?.pageCount,
            pageColors: jobData.printingOptions?.pageColors
          },
          deliveryNumber,
          orderId: jobData.orderId,
          customerInfo: jobData.customerInfo
        };

        // Add to queue
        const jobId = addToQueue(printJob, printerIndex);
        jobIds.push(jobId);
        deliveryNumbers.push(deliveryNumber);

        console.log(`Print job ${i + 1}/${fileURLs.length} queued: ${jobId} (Delivery: ${deliveryNumber}, File: ${originalFileNames[i]})`);
      }

      console.log(`✅ All ${fileURLs.length} print jobs queued successfully`);

      res.json({
        success: true,
        message: `${fileURLs.length} print jobs added to queue`,
        jobIds,
        deliveryNumbers,
        // Return first delivery number for backward compatibility
        jobId: jobIds[0],
        deliveryNumber: deliveryNumbers[0]
      });
    } else {
      // Legacy: single file format (backward compatibility)
      // Generate delivery number
      const deliveryNumber = generateDeliveryNumber(printerIndex);

      // Create print job
      const printJob = {
        fileUrl: jobData.fileUrl!,
        fileName: jobData.fileName || 'document.pdf',
        fileType: jobData.fileType || 'application/pdf',
        printingOptions: {
          pageSize: jobData.printingOptions?.pageSize || 'A4',
          color: jobData.printingOptions?.color || 'bw',
          sided: jobData.printingOptions?.sided || 'single',
          copies: jobData.printingOptions?.copies || 1,
          pageCount: jobData.printingOptions?.pageCount,
          pageColors: jobData.printingOptions?.pageColors
        },
        deliveryNumber,
        orderId: jobData.orderId,
        customerInfo: jobData.customerInfo
      };

      // Add to queue
      const jobId = addToQueue(printJob, printerIndex);

      console.log(`Print job queued: ${jobId} (Delivery: ${deliveryNumber})`);

      res.json({
        success: true,
        message: 'Print job added to queue',
        jobId,
        deliveryNumber
      });
    }
  } catch (error) {
    console.error('Error adding print job:', error);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : 'Failed to add print job'
    });
  }
});

/**
 * GET /api/queue/status
 * Get queue status
 */
router.get('/queue/status', async (req: Request, res: Response) => {
  try {
    const status = getQueueStatus();
    res.json({
      success: true,
      ...status
    });
  } catch (error) {
    console.error('Error getting queue status:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get queue status'
    });
  }
});

/**
 * GET /api/health
 * Health check endpoint
 */
router.get('/health', async (req: Request, res: Response) => {
  try {
    const printerStatus = await checkPrinterStatus();
    const queueStatus = getQueueStatus();

    res.json({
      success: true,
      status: 'healthy',
      printer: printerStatus,
      queue: {
        total: queueStatus.total,
        pending: queueStatus.pending
      },
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    console.error('Error in health check:', error);
    res.status(500).json({
      success: false,
      status: 'unhealthy',
      error: error instanceof Error ? error.message : 'Health check failed'
    });
  }
});

export default router;

