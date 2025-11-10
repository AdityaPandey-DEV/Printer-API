import * as fs from 'fs';
import * as path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import { generateBlankPage, generateLetterSeparator } from '../utils/printerUtils';
import { shouldPrintLetterSeparator, getCurrentLetter } from './deliveryNumber';

const execAsync = promisify(exec);

export interface PrintJob {
  fileUrl: string;
  fileName: string;
  fileType: string;
  printingOptions: {
    pageSize: 'A4' | 'A3';
    color: 'color' | 'bw' | 'mixed';
    sided: 'single' | 'double';
    copies: number;
  };
  deliveryNumber: string;
}

export interface PrintResult {
  success: boolean;
  message: string;
  deliveryNumber?: string;
  error?: string;
}

/**
 * Download file from URL
 */
async function downloadFile(url: string, outputPath: string): Promise<void> {
  const https = require('https');
  const http = require('http');
  const fs = require('fs');

  return new Promise((resolve, reject) => {
    const protocol = url.startsWith('https') ? https : http;
    const file = fs.createWriteStream(outputPath);

    protocol.get(url, (response: any) => {
      if (response.statusCode === 301 || response.statusCode === 302) {
        // Handle redirects
        return downloadFile(response.headers.location!, outputPath)
          .then(resolve)
          .catch(reject);
      }

      if (response.statusCode !== 200) {
        reject(new Error(`Failed to download file: ${response.statusCode}`));
        return;
      }

      response.pipe(file);
      file.on('finish', () => {
        file.close();
        resolve();
      });
    }).on('error', (err: Error) => {
      fs.unlink(outputPath, () => {});
      reject(err);
    });
  });
}

/**
 * Print file using system printer command
 */
async function printFile(filePath: string, options: PrintJob['printingOptions']): Promise<void> {
  const printerName = process.env.PRINTER_NAME || 'HP_Deskjet_525';
  const printerPath = process.env.PRINTER_PATH;

  // Determine print command based on OS
  const isWindows = process.platform === 'win32';
  const isMac = process.platform === 'darwin';
  const isLinux = process.platform === 'linux';

  let printCommand: string;

  if (isWindows) {
    // Windows: Use print command or PowerShell
    printCommand = `print /D:"${printerName}" "${filePath}"`;
  } else if (isMac) {
    // macOS: Use lp command
    const copies = options.copies || 1;
    const colorMode = options.color === 'color' ? 'color' : 'grayscale';
    printCommand = `lp -d "${printerName}" -n ${copies} -o ColorModel=${colorMode} "${filePath}"`;
  } else if (isLinux) {
    // Linux: Use lp command
    const copies = options.copies || 1;
    const colorMode = options.color === 'color' ? 'color' : 'grayscale';
    printCommand = `lp -d "${printerName}" -n ${copies} -o ColorModel=${colorMode} "${filePath}"`;
  } else {
    throw new Error(`Unsupported platform: ${process.platform}`);
  }

  console.log(`Executing print command: ${printCommand}`);
  const { stdout, stderr } = await execAsync(printCommand);
  
  if (stderr && !stderr.includes('request id')) {
    console.warn('Print command stderr:', stderr);
  }
  
  console.log('Print command stdout:', stdout);
}

/**
 * Main print function
 */
export async function printJob(job: PrintJob, printerIndex: number): Promise<PrintResult> {
  try {
    console.log(`Starting print job: ${job.fileName} (Delivery: ${job.deliveryNumber})`);

    // Create temp directory
    const tempDir = path.join(process.cwd(), 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    // Download file
    const fileExtension = path.extname(job.fileName) || '.pdf';
    const tempFilePath = path.join(tempDir, `${job.deliveryNumber}${fileExtension}`);
    
    console.log(`Downloading file from ${job.fileUrl}...`);
    await downloadFile(job.fileUrl, tempFilePath);
    console.log(`File downloaded to ${tempFilePath}`);

    // Print the file
    console.log(`Printing file: ${tempFilePath}`);
    await printFile(tempFilePath, job.printingOptions);
    console.log(`File printed successfully`);

    // Print blank page separator after every file
    console.log('Printing blank page separator...');
    const blankPage = await generateBlankPage();
    const blankPagePath = path.join(tempDir, `blank_${job.deliveryNumber}.pdf`);
    fs.writeFileSync(blankPagePath, blankPage);
    await printFile(blankPagePath, { ...job.printingOptions, copies: 1 });
    fs.unlinkSync(blankPagePath);
    console.log('Blank page separator printed');

    // Print letter separator after every 10 files
    if (shouldPrintLetterSeparator()) {
      const letter = getCurrentLetter();
      console.log(`Printing letter separator: ${letter}`);
      const letterPage = await generateLetterSeparator(letter);
      const letterPagePath = path.join(tempDir, `letter_${letter}_${job.deliveryNumber}.pdf`);
      fs.writeFileSync(letterPagePath, letterPage);
      await printFile(letterPagePath, { ...job.printingOptions, copies: 1 });
      fs.unlinkSync(letterPagePath);
      console.log(`Letter separator ${letter} printed`);
    }

    // Cleanup
    if (fs.existsSync(tempFilePath)) {
      fs.unlinkSync(tempFilePath);
    }

    return {
      success: true,
      message: 'Print job completed successfully',
      deliveryNumber: job.deliveryNumber
    };
  } catch (error) {
    console.error('Error printing job:', error);
    return {
      success: false,
      message: 'Failed to print job',
      error: error instanceof Error ? error.message : 'Unknown error',
      deliveryNumber: job.deliveryNumber
    };
  }
}

/**
 * Check if printer is available
 */
export async function checkPrinterStatus(): Promise<{ available: boolean; message: string }> {
  try {
    const printerName = process.env.PRINTER_NAME || 'HP_Deskjet_525';
    const isWindows = process.platform === 'win32';
    const isMac = process.platform === 'darwin';
    const isLinux = process.platform === 'linux';

    let checkCommand: string;

    if (isWindows) {
      checkCommand = `wmic printer where name="${printerName}" get name,status`;
    } else if (isMac || isLinux) {
      checkCommand = `lpstat -p "${printerName}"`;
    } else {
      return { available: false, message: 'Unsupported platform' };
    }

    await execAsync(checkCommand);
    return { available: true, message: 'Printer is available' };
  } catch (error) {
    return { 
      available: false, 
      message: error instanceof Error ? error.message : 'Printer check failed' 
    };
  }
}

