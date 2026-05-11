import * as fs from 'fs';
import * as path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import { printJob, PrintJob, PrintResult, checkPrinterStatus } from './printer';
import { generateDeliveryNumber } from './deliveryNumber';

const execAsync = promisify(exec);

type JobStatus = 'pending' | 'printing' | 'completed' | 'failed';

interface QueuedJob {
  id: string;
  job: PrintJob;
  printerIndex: number;
  status: JobStatus;
  attempts: number;
  createdAt: Date;
  lastAttemptAt?: Date;
  completedAt?: Date;
  error?: string;
}

const QUEUE_FILE = path.join(process.cwd(), 'print-queue.json');
let printQueue: QueuedJob[] = [];
let isProcessing = false;

// Polling interval to check printer spooler (in ms)
const PRINTER_IDLE_POLL_INTERVAL = 5000; // Check every 5 seconds
const PRINTER_IDLE_TIMEOUT = 600000; // Max wait 10 minutes for printer to become idle

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
        status: item.status || 'pending', // Default to pending for legacy entries
        createdAt: new Date(item.createdAt),
        lastAttemptAt: item.lastAttemptAt ? new Date(item.lastAttemptAt) : undefined,
        completedAt: item.completedAt ? new Date(item.completedAt) : undefined
      }));

      // Reset any jobs that were stuck in 'printing' state (server crashed mid-print)
      let resetCount = 0;
      printQueue.forEach(job => {
        if (job.status === 'printing') {
          job.status = 'pending';
          resetCount++;
        }
      });
      if (resetCount > 0) {
        console.log(`⚠️ Reset ${resetCount} jobs from 'printing' back to 'pending' (server was restarted mid-print)`);
        saveQueue();
      }

      // Remove completed jobs from in-memory queue (keep only pending/failed)
      const activeJobs = printQueue.filter(j => j.status === 'pending' || j.status === 'failed');
      const completedCount = printQueue.length - activeJobs.length;
      printQueue = activeJobs;

      console.log(`📋 Loaded queue: ${printQueue.length} pending jobs (${completedCount} completed jobs cleared)`);
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
 * Check if the printer spooler has any active print jobs
 * Returns true if printer is idle (no jobs in spooler), false if busy
 *
 * IMPORTANT: Queries the SPECIFIC printer from PRINTER_NAME env var,
 * then falls back to checking ALL printers. Never relies on "first printer"
 * which could return a virtual printer (e.g., Microsoft Print to PDF) with 0 jobs.
 *
 * Windows spooler job statuses:
 *   - "Normal"              → queued, waiting to print (ACTIVE)
 *   - "Spooling"            → being spooled (ACTIVE)
 *   - "Printing"            → actively printing (ACTIVE)
 *   - "Printing, Retained"  → actively printing, will be retained after (ACTIVE)
 *   - "Printing, PaperOut, Retained" → printing but out of paper (ACTIVE)
 *   - "Retained"            → finished printing, kept in spooler (COMPLETED - ignore)
 *   - "Printed"             → finished printing (COMPLETED - ignore)
 *   - "Printed, Retained"   → finished, retained (COMPLETED - ignore)
 *   - "Complete"            → done (COMPLETED - ignore)
 *   - "Deleted"             → removed (COMPLETED - ignore)
 */
async function isPrinterIdle(): Promise<{ idle: boolean; jobCount: number; details: string }> {
  try {
    if (process.platform === 'win32') {
      // The active job filter - only count jobs that are genuinely active:
      //   - Include: Normal, Spooling, Printing, "Printing, Retained", "Printing, PaperOut, Retained"
      //   - Exclude: Retained (without Printing/Spooling), Printed, Complete, Deleted
      const activeFilter = `$s = $_.JobStatus; $s -ne 'Complete' -and $s -ne 'Deleted' -and $s -ne 'Printed' -and $s -notlike 'Printed*' -and (-not ($s -like '*Retained*' -and $s -notlike '*Printing*' -and $s -notlike '*Spooling*'))`;

      // Get the specific printer name from environment
      const printerName = process.env.PRINTER_NAME || '';

      if (printerName) {
        // === PRIMARY: Query the SPECIFIC printer being used ===
        // Escape single quotes for PowerShell (parentheses are safe inside single quotes)
        const escapedName = printerName.replace(/'/g, "''");

        const command = `powershell -Command "$jobs = Get-PrintJob -PrinterName '${escapedName}' -ErrorAction SilentlyContinue; if ($jobs) { $active = @($jobs | Where-Object { ${activeFilter} }); Write-Output ('COUNT:' + $active.Count + '|DETAILS:' + (($active | ForEach-Object { $_.DocumentName + '(' + $_.JobStatus + ')' } | Select-Object -First 5) -join ', ')) } else { Write-Output 'COUNT:0|DETAILS:No jobs' }"`;

        try {
          const { stdout } = await execAsync(command, { timeout: 15000 });
          const output = stdout.trim();

          const countMatch = output.match(/COUNT:(\d+)/);
          const detailsMatch = output.match(/DETAILS:(.*)/);
          const jobCount = countMatch ? parseInt(countMatch[1], 10) : 0;
          const details = detailsMatch ? detailsMatch[1].trim() : 'Unknown';

          return { idle: jobCount === 0, jobCount, details };
        } catch (specificError: any) {
          console.warn(`⚠️ Failed to query specific printer '${printerName}': ${specificError.message}`);
          // Fall through to ALL printers check
        }
      }

      // === FALLBACK: Check ALL printers for active jobs ===
      // This ensures we never miss active jobs on any printer
      try {
        const allPrintersCommand = `powershell -Command "$totalActive = 0; $allDetails = @(); Get-Printer -ErrorAction SilentlyContinue | ForEach-Object { $pName = $_.Name; $jobs = Get-PrintJob -PrinterName $pName -ErrorAction SilentlyContinue; if ($jobs) { $active = @($jobs | Where-Object { ${activeFilter} }); if ($active.Count -gt 0) { $totalActive += $active.Count; $active | ForEach-Object { $allDetails += ($_.DocumentName + '(' + $_.JobStatus + ')') } } } }; if ($totalActive -gt 0) { Write-Output ('COUNT:' + $totalActive + '|DETAILS:' + (($allDetails | Select-Object -First 5) -join ', ')) } else { Write-Output 'COUNT:0|DETAILS:No jobs' }"`;
        const { stdout } = await execAsync(allPrintersCommand, { timeout: 20000 });
        const output = stdout.trim();

        const countMatch = output.match(/COUNT:(\d+)/);
        const detailsMatch = output.match(/DETAILS:(.*)/);
        const jobCount = countMatch ? parseInt(countMatch[1], 10) : 0;
        const details = detailsMatch ? detailsMatch[1].trim() : 'Unknown';

        return { idle: jobCount === 0, jobCount, details };
      } catch (allError: any) {
        console.warn(`⚠️ Failed to query all printers: ${allError.message}`);
        // If all queries fail, do NOT assume idle - report as busy to be safe
        console.warn('⚠️ Cannot determine printer spooler status - treating as BUSY to prevent overlap');
        return { idle: false, jobCount: -1, details: 'Could not check spooler (treating as busy for safety)' };
      }
    } else {
      // macOS/Linux: Check lpstat for active jobs
      try {
        const { stdout } = await execAsync('lpstat -o 2>/dev/null || echo "IDLE"', { timeout: 5000 });
        const output = stdout.trim();
        if (output === 'IDLE' || output === '') {
          return { idle: true, jobCount: 0, details: 'No active print jobs' };
        }
        const jobCount = output.split('\n').length;
        return { idle: false, jobCount, details: output };
      } catch {
        return { idle: true, jobCount: 0, details: 'Could not check spooler (assuming idle)' };
      }
    }
  } catch (error) {
    console.warn('⚠️ Error checking printer spooler:', error);
    // Safety: treat as busy when we can't determine status
    return { idle: false, jobCount: -1, details: 'Error checking spooler (treating as busy for safety)' };
  }
}

/**
 * Wait for the printer to become idle (no active print jobs in spooler)
 * This ensures we don't send a new job while the printer is still physically printing
 */
async function waitForPrinterIdle(): Promise<void> {
  console.log('🔍 Checking if printer is idle before sending next job...');

  const startTime = Date.now();
  let checkCount = 0;

  while (true) {
    checkCount++;
    const status = await isPrinterIdle();

    if (status.idle) {
      if (checkCount > 1) {
        console.log(`✅ Printer is now idle after ${checkCount} checks (waited ${Math.round((Date.now() - startTime) / 1000)}s)`);
      } else {
        console.log(`✅ Printer is idle - ready to accept next job`);
      }
      return;
    }

    // Check timeout
    const elapsed = Date.now() - startTime;
    if (elapsed >= PRINTER_IDLE_TIMEOUT) {
      console.warn(`⚠️ Printer idle timeout after ${Math.round(elapsed / 1000)}s - proceeding anyway`);
      console.warn(`   Active jobs in spooler: ${status.jobCount} - ${status.details}`);
      return;
    }

    console.log(`⏳ Printer is busy (${status.jobCount} jobs in spooler: ${status.details}) - waiting ${PRINTER_IDLE_POLL_INTERVAL / 1000}s... [Check ${checkCount}, elapsed: ${Math.round(elapsed / 1000)}s]`);
    await new Promise(resolve => setTimeout(resolve, PRINTER_IDLE_POLL_INTERVAL));
  }
}

/**
 * Add job to queue (transaction - saved to print-queue.json but NOT immediately executed)
 * The job will be picked up by the queue processor when the printer is idle.
 */
export function addToQueue(job: PrintJob, printerIndex: number): string {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  const queuedJob: QueuedJob = {
    id: jobId,
    job,
    printerIndex,
    status: 'pending',
    attempts: 0,
    createdAt: new Date()
  };

  printQueue.push(queuedJob);
  saveQueue();
  console.log(`✅ Added job ${jobId} to queue as PENDING transaction (Total: ${printQueue.length} jobs)`);
  console.log(`📄 Job details: ${job.fileName} (Delivery: ${job.deliveryNumber || 'pending'})`);
  console.log(`💾 Queue saved to print-queue.json - job will wait for printer to be idle`);

  // Start processing if not already processing
  if (!isProcessing) {
    console.log(`🔄 Starting queue processor...`);
    processQueue();
  } else {
    console.log(`⏳ Queue processor is already running - job will be processed when printer finishes current work`);
  }

  return jobId;
}

/**
 * Process queue sequentially - waits for printer to be fully idle between jobs
 * Each job is a "transaction":
 *   1. Status set to 'printing' → saved to file
 *   2. Wait for printer spooler to be idle (no active physical print jobs)
 *   3. Send job to printer
 *   4. Wait for printer to finish (spooler idle again)
 *   5. Status set to 'completed' → removed from queue file
 */
async function processQueue(): Promise<void> {
  if (isProcessing) {
    return;
  }

  isProcessing = true;
  console.log('🖨️ Queue processor started - will process jobs one at a time, waiting for printer idle between jobs');

  while (printQueue.length > 0) {
    // Find the first pending job
    const queuedJob = printQueue.find(j => j.status === 'pending');
    if (!queuedJob) {
      // No pending jobs left - remove completed/failed ones and exit
      printQueue = printQueue.filter(j => j.status === 'pending');
      saveQueue();
      break;
    }

    // === STEP 1: Wait for printer to be idle before starting this job ===
    try {
      await waitForPrinterIdle();
    } catch (idleError) {
      console.warn(`⚠️ Error waiting for printer idle, proceeding anyway:`, idleError);
    }

    // === STEP 2: Mark job as 'printing' (transaction starts) ===
    queuedJob.status = 'printing';
    queuedJob.attempts++;
    queuedJob.lastAttemptAt = new Date();
    saveQueue();

    console.log(`\n${'='.repeat(60)}`);
    console.log(`🖨️ PRINTING JOB: ${queuedJob.id} (Attempt ${queuedJob.attempts})`);
    console.log(`📄 File: ${queuedJob.job.fileName}`);
    console.log(`📋 Delivery: ${queuedJob.job.deliveryNumber || 'pending'}`);
    console.log(`📊 Queue position: 1 of ${printQueue.filter(j => j.status === 'pending').length + 1} jobs`);
    console.log(`${'='.repeat(60)}\n`);

    try {
      // Generate delivery number if not present
      if (!queuedJob.job.deliveryNumber) {
        queuedJob.job.deliveryNumber = generateDeliveryNumber(queuedJob.printerIndex);
      }

      // === STEP 3: Send to printer ===
      const result: PrintResult = await printJob(queuedJob.job, queuedJob.printerIndex);

      if (result.success) {
        // === STEP 4: Wait for printer to finish printing this job physically ===
        console.log(`⏳ Job sent to printer successfully. Waiting for printer to finish physically printing...`);
        // Give the spooler a moment to register the job before we start polling
        await new Promise(resolve => setTimeout(resolve, 3000));
        await waitForPrinterIdle();

        // === STEP 5: Mark as completed and remove from queue ===
        queuedJob.status = 'completed';
        queuedJob.completedAt = new Date();
        // Remove completed job from queue
        printQueue = printQueue.filter(j => j.id !== queuedJob.id);
        saveQueue();

        console.log(`\n✅ Job ${queuedJob.id} COMPLETED successfully`);
        console.log(`📄 File: ${queuedJob.job.fileName}`);
        console.log(`📋 Remaining jobs in queue: ${printQueue.length}`);
        console.log(`${'─'.repeat(60)}\n`);
      } else {
        // Job failed - keep in queue for retry
        const errorMessage = result.error || result.message || 'Unknown error';
        queuedJob.status = 'pending'; // Reset to pending for retry
        queuedJob.error = errorMessage;
        saveQueue();

        console.log(`❌ Job ${queuedJob.id} FAILED (Attempt ${queuedJob.attempts}): ${errorMessage}`);
        console.log(`📋 Job will be retried. Queue: ${printQueue.length} jobs remaining`);

        // Log specific error types
        const errorLower = errorMessage.toLowerCase();
        if (errorLower.includes('printer not connected') ||
            errorLower.includes('printer is offline') ||
            errorLower.includes('printer not found') ||
            errorLower.includes('unable to initialize device') ||
            errorLower.includes('powered off')) {
          console.warn(`⚠️ Printer issue detected: ${errorMessage}`);
          console.warn(`⚠️ Job ${queuedJob.id} will be retried when printer is available`);
        }

        // Wait before retry (exponential backoff, max 5 minutes)
        const baseWaitTime = errorLower.includes('printer not connected') ||
                             errorLower.includes('printer is offline') ||
                             errorLower.includes('printer not found') ||
                             errorLower.includes('unable to initialize device') ||
                             errorLower.includes('powered off')
                             ? 30000 // 30 seconds for printer issues
                             : 10000; // 10 seconds for other errors
        const waitTime = Math.min(queuedJob.attempts * baseWaitTime, 300000);
        console.log(`⏳ Waiting ${waitTime / 1000}s before retry...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      console.error(`❌ Error processing job ${queuedJob.id}:`, errorMessage);

      // Reset to pending for retry
      queuedJob.status = 'pending';
      queuedJob.error = errorMessage;
      saveQueue();

      console.log(`📋 Queue: ${printQueue.length} jobs remaining (Job ${queuedJob.id} will be retried)`);

      // Log specific error types
      const errorLower = errorMessage.toLowerCase();
      if (errorLower.includes('printer not connected') ||
          errorLower.includes('printer is offline') ||
          errorLower.includes('printer not found') ||
          errorLower.includes('unable to initialize device') ||
          errorLower.includes('powered off')) {
        console.warn(`⚠️ Printer issue detected: ${errorMessage}`);
        console.warn(`⚠️ Job ${queuedJob.id} will be retried when printer is available`);
      }

      // Wait before retry (longer wait for printer connection issues)
      const baseWaitTime = errorLower.includes('printer not connected') ||
                           errorLower.includes('printer is offline') ||
                           errorLower.includes('printer not found') ||
                           errorLower.includes('unable to initialize device') ||
                           errorLower.includes('powered off')
                           ? 30000 // 30 seconds for printer issues
                           : 10000; // 10 seconds for other errors
      const waitTime = Math.min(queuedJob.attempts * baseWaitTime, 300000);
      console.log(`⏳ Waiting ${waitTime / 1000}s before retry...`);
      await new Promise(resolve => setTimeout(resolve, waitTime));
    }
  }

  isProcessing = false;
  console.log('🏁 Queue processor finished - all jobs processed');
}

/**
 * Get queue status
 */
export function getQueueStatus(): {
  total: number;
  pending: number;
  printing: number;
  currentJob: QueuedJob | null;
  jobs: QueuedJob[];
} {
  const pendingJobs = printQueue.filter(j => j.status === 'pending');
  const printingJob = printQueue.find(j => j.status === 'printing') || null;

  return {
    total: printQueue.length,
    pending: pendingJobs.length,
    printing: printingJob ? 1 : 0,
    currentJob: printingJob,
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

// Start processing if queue has pending jobs
const pendingOnStartup = printQueue.filter(j => j.status === 'pending');
if (pendingOnStartup.length > 0) {
  console.log(`📋 Found ${pendingOnStartup.length} pending jobs on startup, starting queue processor...`);
  processQueue();
}
