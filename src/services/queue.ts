import * as fs from 'fs';
import * as path from 'path';
import { printJob, PrintJob, PrintResult } from './printer';
import { generateDeliveryNumber } from './deliveryNumber';

interface QueuedJob {
  id: string;
  job: PrintJob;
  printerIndex: number;
  attempts: number;
  createdAt: Date;
  lastAttemptAt?: Date;
}

const QUEUE_FILE = path.join(process.cwd(), 'print-queue.json');
let printQueue: QueuedJob[] = [];
let isProcessing = false;

/**
 * Load queue from file
 */
function loadQueue(): void {
  try {
    if (fs.existsSync(QUEUE_FILE)) {
      const data = fs.readFileSync(QUEUE_FILE, 'utf-8');
      const parsed = JSON.parse(data);
      printQueue = parsed.map((item: any) => ({
        ...item,
        createdAt: new Date(item.createdAt),
        lastAttemptAt: item.lastAttemptAt ? new Date(item.lastAttemptAt) : undefined
      }));
      console.log(`Loaded ${printQueue.length} jobs from queue`);
    }
  } catch (error) {
    console.error('Error loading queue:', error);
    printQueue = [];
  }
}

/**
 * Save queue to file
 */
function saveQueue(): void {
  try {
    fs.writeFileSync(QUEUE_FILE, JSON.stringify(printQueue, null, 2));
  } catch (error) {
    console.error('Error saving queue:', error);
  }
}

/**
 * Add job to queue
 */
export function addToQueue(job: PrintJob, printerIndex: number): string {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  const queuedJob: QueuedJob = {
    id: jobId,
    job,
    printerIndex,
    attempts: 0,
    createdAt: new Date()
  };

  printQueue.push(queuedJob);
  saveQueue();
  console.log(`Added job ${jobId} to queue (Total: ${printQueue.length})`);
  
  // Start processing if not already processing
  if (!isProcessing) {
    processQueue();
  }

  return jobId;
}

/**
 * Process queue with infinite retry
 */
async function processQueue(): Promise<void> {
  if (isProcessing) {
    return;
  }

  isProcessing = true;
  console.log('Starting queue processing...');

  while (printQueue.length > 0) {
    const queuedJob = printQueue[0];
    queuedJob.attempts++;
    queuedJob.lastAttemptAt = new Date();

    console.log(`Processing job ${queuedJob.id} (Attempt ${queuedJob.attempts})...`);

    try {
      // Generate delivery number if not present
      if (!queuedJob.job.deliveryNumber) {
        queuedJob.job.deliveryNumber = generateDeliveryNumber(queuedJob.printerIndex);
      }

      const result: PrintResult = await printJob(queuedJob.job, queuedJob.printerIndex);

      if (result.success) {
        // Remove from queue on success
        printQueue.shift();
        saveQueue();
        console.log(`Job ${queuedJob.id} completed successfully`);
      } else {
        // Keep in queue and retry (infinite retry)
        console.log(`Job ${queuedJob.id} failed, will retry: ${result.error}`);
        saveQueue();
        
        // Wait before retry (exponential backoff, max 5 minutes)
        const waitTime = Math.min(queuedJob.attempts * 10000, 300000);
        console.log(`Waiting ${waitTime}ms before retry...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      }
    } catch (error) {
      console.error(`Error processing job ${queuedJob.id}:`, error);
      // Keep in queue and retry
      saveQueue();
      
      // Wait before retry
      const waitTime = Math.min(queuedJob.attempts * 10000, 300000);
      await new Promise(resolve => setTimeout(resolve, waitTime));
    }
  }

  isProcessing = false;
  console.log('Queue processing completed');
}

/**
 * Get queue status
 */
export function getQueueStatus(): { total: number; pending: number; jobs: QueuedJob[] } {
  return {
    total: printQueue.length,
    pending: printQueue.length,
    jobs: [...printQueue]
  };
}

/**
 * Clear queue (use with caution)
 */
export function clearQueue(): void {
  printQueue = [];
  saveQueue();
  console.log('Queue cleared');
}

// Load queue on startup
loadQueue();

// Start processing if queue has jobs
if (printQueue.length > 0) {
  processQueue();
}

