import * as fs from 'fs';
import * as path from 'path';
import * as http from 'http';
import * as url from 'url';
import { exec } from 'child_process';
import { promisify } from 'util';
import { PDFDocument } from 'pdf-lib';
import { v4 as uuidv4 } from 'uuid';
import { generateOrderSummaryPage } from '../utils/printerUtils';

const execAsync = promisify(exec);

// Cache for detected printer name
let detectedPrinterName: string | null = null;
let lastPrinterDetection: number = 0;
const PRINTER_DETECTION_CACHE_MS = 60000; // Cache for 1 minute

/**
 * Detect available printers automatically
 */
async function detectPrinter(): Promise<string | null> {
  const now = Date.now();
  
  // Return cached printer if still valid
  if (detectedPrinterName && (now - lastPrinterDetection) < PRINTER_DETECTION_CACHE_MS) {
    return detectedPrinterName;
  }

  try {
    const isWindows = process.platform === 'win32';
    const isMac = process.platform === 'darwin';
    const isLinux = process.platform === 'linux';

    let listCommand: string;

    if (isWindows) {
      // PowerShell: Get first available printer (prefer default, then any available)
      try {
        console.log('🔍 Attempting to detect printer on Windows...');
        
        // First try to get default printer
        const defaultCommand = `powershell -Command "$printer = Get-Printer | Where-Object {$_.Default -eq $true} | Select-Object -First 1; if ($printer) { Write-Output $printer.Name }"`;
        try {
          const { stdout: defaultStdout, stderr: defaultStderr } = await execAsync(defaultCommand);
          const defaultPrinter = defaultStdout.trim();
          console.log(`🔍 Default printer detection result: "${defaultPrinter}"`);
          
          if (defaultPrinter && defaultPrinter.length > 0 && !defaultPrinter.includes('Get-Printer') && !defaultPrinter.includes('Where-Object')) {
            detectedPrinterName = defaultPrinter;
            lastPrinterDetection = now;
            console.log(`✅ Auto-detected default printer: ${defaultPrinter}`);
            return defaultPrinter;
          }
        } catch (defaultError: any) {
          console.log(`⚠️ Default printer detection failed: ${defaultError.message}`);
        }
        
        // If no default, get first available printer (any status)
        const availableCommand = `powershell -Command "$printer = Get-Printer | Select-Object -First 1; if ($printer) { Write-Output $printer.Name }"`;
        try {
          const { stdout: availableStdout } = await execAsync(availableCommand);
          const availablePrinter = availableStdout.trim();
          console.log(`🔍 Available printer detection result: "${availablePrinter}"`);
          
          if (availablePrinter && availablePrinter.length > 0 && !availablePrinter.includes('Get-Printer') && !availablePrinter.includes('Select-Object')) {
            detectedPrinterName = availablePrinter;
            lastPrinterDetection = now;
            console.log(`✅ Auto-detected printer: ${availablePrinter}`);
            return availablePrinter;
          }
        } catch (availableError: any) {
          console.log(`⚠️ Available printer detection failed: ${availableError.message}`);
        }
        
        // Try listing all printers with simpler command
        const listCommand = `powershell -Command "Get-Printer | Select-Object -First 1 | ForEach-Object { Write-Output $_.Name }"`;
        try {
          const { stdout: listStdout } = await execAsync(listCommand);
          const listPrinter = listStdout.trim();
          console.log(`🔍 List printer detection result: "${listPrinter}"`);
          
          if (listPrinter && listPrinter.length > 0 && !listPrinter.includes('Get-Printer') && !listPrinter.includes('ForEach-Object')) {
            detectedPrinterName = listPrinter;
            lastPrinterDetection = now;
            console.log(`✅ Auto-detected printer (list): ${listPrinter}`);
            return listPrinter;
          }
        } catch (listError: any) {
          console.log(`⚠️ List printer detection failed: ${listError.message}`);
        }
        
        // Try even simpler command - just get all printer names
        const simpleCommand = `powershell -Command "(Get-Printer).Name | Select-Object -First 1"`;
        try {
          const { stdout: simpleStdout } = await execAsync(simpleCommand);
          const simplePrinter = simpleStdout.trim();
          console.log(`🔍 Simple printer detection result: "${simplePrinter}"`);
          
          if (simplePrinter && simplePrinter.length > 0 && !simplePrinter.includes('Get-Printer') && !simplePrinter.includes('Select-Object')) {
            detectedPrinterName = simplePrinter;
            lastPrinterDetection = now;
            console.log(`✅ Auto-detected printer (simple): ${simplePrinter}`);
            return simplePrinter;
          }
        } catch (simpleError: any) {
          console.log(`⚠️ Simple printer detection failed: ${simpleError.message}`);
        }
      } catch (error: any) {
        // If PowerShell fails on Windows, try wmic
        if (isWindows) {
          try {
            const wmicCommand = `wmic printer where "Default='TRUE'" get Name /value | findstr "Name="`;
            const { stdout } = await execAsync(wmicCommand);
            const match = stdout.match(/Name=(.+)/);
            if (match && match[1]) {
              const printerName = match[1].trim();
              detectedPrinterName = printerName;
              lastPrinterDetection = now;
              console.log(`✅ Auto-detected printer (wmic): ${printerName}`);
              return printerName;
            }
          } catch (wmicError) {
            // Try listing all printers with wmic
            try {
              const wmicListCommand = `wmic printer get Name /value | findstr "Name=" | findstr /v "Name="`;
              const { stdout } = await execAsync(wmicListCommand);
              const lines = stdout.split('\n').filter(line => line.trim().length > 0);
              if (lines.length > 0) {
                const printerName = lines[0].replace('Name=', '').trim();
                detectedPrinterName = printerName;
                lastPrinterDetection = now;
                console.log(`✅ Auto-detected printer (wmic list): ${printerName}`);
                return printerName;
              }
            } catch (listError) {
              console.warn('⚠️ Could not detect printer automatically');
            }
          }
        }
      }
    } else if (isMac || isLinux) {
      // Use lpstat to list all printers
      try {
        // Try lpstat -p to list printers
        const { stdout } = await execAsync(`lpstat -p 2>/dev/null | head -1 | awk '{print $2}' | sed 's/^printer //'`);
        const printerName = stdout.trim();
        if (printerName && printerName.length > 0) {
          detectedPrinterName = printerName;
          lastPrinterDetection = now;
          console.log(`✅ Auto-detected printer: ${printerName}`);
          return printerName;
        }
      } catch (error: any) {
        // Try lpstat -a to list all printers
        try {
          const { stdout } = await execAsync(`lpstat -a 2>/dev/null | head -1 | awk '{print $1}'`);
          const printerName = stdout.trim();
          if (printerName && printerName.length > 0) {
            detectedPrinterName = printerName;
            lastPrinterDetection = now;
            console.log(`✅ Auto-detected printer (lpstat -a): ${printerName}`);
            return printerName;
          }
        } catch (altError) {
          console.warn('⚠️ Could not detect printer automatically');
        }
      }
    }
  } catch (error) {
    console.error('Error detecting printer:', error);
  }

  return null;
}

/**
 * Get printer name (from env, detected, or default)
 */
async function getPrinterName(): Promise<string> {
  // First, try environment variable
  if (process.env.PRINTER_NAME) {
    console.log(`🖨️ Using printer from environment: ${process.env.PRINTER_NAME}`);
    return process.env.PRINTER_NAME;
  }

  // Try to detect automatically
  console.log('🔍 No PRINTER_NAME in environment, attempting auto-detection...');
  const detected = await detectPrinter();
  if (detected) {
    console.log(`✅ Using auto-detected printer: ${detected}`);
    return detected;
  }

  // Fallback to default (will be updated by auto-detection if printer is found)
  console.log(`⚠️ No printer detected, using default: HP_Deskjet_525`);
  return 'HP_Deskjet_525';
}

export interface PrintJob {
  fileUrl: string;
  fileName: string;
  fileType: string;
  printingOptions: {
    pageSize: 'A4' | 'A3';
    color: 'color' | 'bw' | 'mixed';
    sided: 'single' | 'double';
    copies: number;
    pageCount?: number;
    pageColors?: {
      colorPages: number[];
      bwPages: number[];
    };
  };
  deliveryNumber: string;
  orderId?: string;
  customerInfo?: {
    name: string;
    email: string;
    phone: string;
  };
  orderDetails?: {
    orderType: 'file' | 'template';
    pageSize: 'A4' | 'A3';
    color: 'color' | 'bw' | 'mixed';
    sided: 'single' | 'double';
    copies: number;
    pages: number;
    serviceOptions: Array<{
      fileName: string;
      options: string[];
    }>;
    totalAmount: number;
    expectedDelivery: string;
  };
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
 * Fallback print method using rundll32 to set default printer, then print
 */
async function tryFallbackPrintMethod(filePath: string, printerName: string, options: PrintJob['printingOptions']): Promise<void> {
  console.log(`🔄 Using fallback print method: rundll32 + Start-Process`);
  
  const escapedPrinterName = printerName.replace(/"/g, '\\"');
  const escapedFilePath = filePath.replace(/"/g, '\\"');
  
  // Set printer as default using rundll32
  const setDefaultCmd = `rundll32 printui.dll,PrintUIEntry /y /n "${escapedPrinterName}"`;
  // Print using Start-Process with Print verb (now uses default)
  const printCmd = `powershell -Command "Start-Process -FilePath '${escapedFilePath}' -Verb Print -WindowStyle Hidden -ErrorAction Stop"`;
  
  const fallbackCommand = `${setDefaultCmd} && ${printCmd}`;
  
  console.log(`Executing fallback print command: ${fallbackCommand}`);
  
  try {
    const { stdout, stderr } = await execAsync(fallbackCommand);
    
    if (stdout) {
      const stdoutTrimmed = stdout.trim();
      if (stdoutTrimmed) {
        console.log(`✅ Fallback print command output: ${stdoutTrimmed}`);
      }
    }
    
    if (stderr) {
      const stderrTrimmed = stderr.trim();
      const stderrLower = stderrTrimmed.toLowerCase();
      
      if (stderrLower.includes('error') || stderrLower.includes('exception') || stderrLower.includes('failed')) {
        console.error(`❌ Fallback print command error: ${stderrTrimmed}`);
        throw new Error(`Fallback print method failed: ${stderrTrimmed}`);
      }
      
      if (stderrTrimmed && !stderrLower.includes('request id')) {
        console.warn('Fallback print command stderr:', stderrTrimmed);
      }
    }
    
    // Wait a bit for the print job to be queued
    await new Promise(resolve => setTimeout(resolve, 2000));
  } catch (fallbackError: any) {
    const fallbackErrorMessage = fallbackError.message || fallbackError.stdout || fallbackError.stderr || String(fallbackError);
    console.error(`❌ Fallback print method also failed: ${fallbackErrorMessage}`);
    throw new Error(`Both COM object and fallback print methods failed. Last error: ${fallbackErrorMessage}`);
  }
}

/**
 * Find LibreOffice executable path on Windows
 */
function findLibreOfficePath(): string | null {
  if (process.platform === 'win32') {
    // Common installation paths on Windows
    const possiblePaths = [
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
      process.env['PROGRAMFILES'] + '\\LibreOffice\\program\\soffice.exe',
      process.env['PROGRAMFILES(X86)'] + '\\LibreOffice\\program\\soffice.exe'
    ];
    
    for (const possiblePath of possiblePaths) {
      if (possiblePath && fs.existsSync(possiblePath)) {
        console.log(`✅ Found LibreOffice at: ${possiblePath}`);
        return possiblePath;
      }
    }
  }
  return null;
}

/**
 * Convert Word file (DOCX/DOC) to PDF using LibreOffice
 */
async function convertWordToPdf(wordFilePath: string): Promise<string> {
  try {
    console.log(`🔄 Converting Word file to PDF using LibreOffice: ${wordFilePath}`);
    
    const tempDir = path.join(process.cwd(), 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    const fileExt = path.extname(wordFilePath).toLowerCase();
    const baseName = path.basename(wordFilePath, fileExt);
    const pdfPath = path.join(tempDir, `${baseName}_${uuidv4()}.pdf`);
    
    // Find LibreOffice executable
    let sofficeCommand = 'soffice';
    if (process.platform === 'win32') {
      const libreOfficePath = findLibreOfficePath();
      if (libreOfficePath) {
        sofficeCommand = `"${libreOfficePath}"`;
      } else {
        console.warn('⚠️ LibreOffice not found in standard locations, trying PATH...');
      }
    }
    
    // Convert Word to PDF using LibreOffice CLI
    const command = `${sofficeCommand} --headless --convert-to pdf --outdir "${tempDir}" "${wordFilePath}"`;
    console.log(`Running LibreOffice command: ${command}`);
    
    const { stdout, stderr } = await execAsync(command);
    
    if (stderr && !stderr.includes('Warning') && !stderr.includes('Info')) {
      console.warn('LibreOffice stderr:', stderr);
    }
    
    console.log('LibreOffice stdout:', stdout);
    
    // LibreOffice creates PDF with same base name in the output directory
    const expectedPdfPath = path.join(tempDir, `${baseName}.pdf`);
    
    // Check if PDF was created
    if (!fs.existsSync(expectedPdfPath)) {
      throw new Error('LibreOffice conversion failed - no PDF file created');
    }
    
    // If the expected path is different from what we want, rename it
    if (expectedPdfPath !== pdfPath) {
      fs.renameSync(expectedPdfPath, pdfPath);
    }
    
    console.log(`✅ Word to PDF conversion successful: ${pdfPath}`);
    return pdfPath;
    
  } catch (error: any) {
    console.error(`❌ Error converting Word to PDF with LibreOffice: ${error.message}`);
    throw new Error(`Word to PDF conversion failed: ${error.message}`);
  }
}

/**
 * Print PDF with mixed color pages in sequence (maintains page order)
 */
async function printPdfWithMixedColorInSequence(
  pdfPath: string,
  colorPages: number[],
  bwPages: number[],
  printerName: string,
  copies: number,
  pageSize: 'A4' | 'A3',
  sided: 'single' | 'double',
  tempDir: string
): Promise<void> {
  try {
    console.log(`🔍 DEBUG - printPdfWithMixedColorInSequence called:`);
    console.log(`   PDF: ${pdfPath}`);
    console.log(`   Color pages (1-based): [${colorPages.join(', ')}]`);
    console.log(`   B&W pages (1-based): [${bwPages.join(', ')}]`);
    
    const pdfBytes = fs.readFileSync(pdfPath);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const totalPages = pdfDoc.getPageCount();
    console.log(`   Total pages in PDF: ${totalPages}`);
    
    // Convert 1-based page numbers to 0-based indices
    const colorIndices = new Set(colorPages.map(p => p - 1).filter(i => i >= 0 && i < totalPages));
    const bwIndices = new Set(bwPages.map(p => p - 1).filter(i => i >= 0 && i < totalPages));
    
    console.log(`   Color indices (0-based): [${Array.from(colorIndices).join(', ')}]`);
    console.log(`   B&W indices (0-based): [${Array.from(bwIndices).join(', ')}]`);
    
    // Determine page color mode for each page
    // Priority: colorIndices > bwIndices > default (B&W)
    const pageColorMode: Map<number, boolean> = new Map();
    for (let i = 0; i < totalPages; i++) {
      if (colorIndices.has(i)) {
        pageColorMode.set(i, false); // false = color
      } else if (bwIndices.has(i)) {
        pageColorMode.set(i, true); // true = monochrome (B&W)
      } else {
        // Default to B&W if not specified in either array
        pageColorMode.set(i, true); // true = monochrome (B&W)
      }
    }
    
    // Group consecutive pages with the same color mode to minimize print jobs
    const pageGroups: Array<{ startIndex: number; endIndex: number; isMonochrome: boolean }> = [];
    let currentGroup: { startIndex: number; endIndex: number; isMonochrome: boolean } | null = null;
    
    for (let i = 0; i < totalPages; i++) {
      // Fix: Use non-null assertion since all pages are guaranteed to be in the map
      const isMonochrome = pageColorMode.get(i)!;
      
      if (currentGroup === null) {
        // Start new group
        currentGroup = { startIndex: i, endIndex: i, isMonochrome };
      } else if (currentGroup.isMonochrome === isMonochrome) {
        // Extend current group
        currentGroup.endIndex = i;
      } else {
        // Save current group and start new one
        pageGroups.push(currentGroup);
        currentGroup = { startIndex: i, endIndex: i, isMonochrome };
      }
    }
    
    // Add last group
    if (currentGroup !== null) {
      pageGroups.push(currentGroup);
    }
    
    console.log(`📋 Printing ${totalPages} pages in ${pageGroups.length} groups to maintain sequence`);
    console.log(`📋 Page groups created:`);
    pageGroups.forEach((group, idx) => {
      console.log(`   Group ${idx + 1}: Pages ${group.startIndex + 1}-${group.endIndex + 1} (${group.isMonochrome ? 'B&W' : 'Color'})`);
    });
    
    // Reverse the groups array to print in reverse order (stack-based printing)
    // Last group prints first, so pages stack correctly (last printed = top of stack)
    pageGroups.reverse();
    console.log(`🔄 Printing in reverse order (stack-based) to maintain correct page sequence`);
    
    // Print each group in reverse order (last group first)
    for (let groupIndex = 0; groupIndex < pageGroups.length; groupIndex++) {
      const group = pageGroups[groupIndex];
      const pageCount = group.endIndex - group.startIndex + 1;
      
      // Create PDF for this group
      const groupPdf = await PDFDocument.create();
      for (let i = group.startIndex; i <= group.endIndex; i++) {
        const [copiedPage] = await groupPdf.copyPages(pdfDoc, [i]);
        groupPdf.addPage(copiedPage);
      }
      
      const groupPdfBytes = await groupPdf.save();
      const groupPdfPath = path.join(tempDir, `group_${groupIndex}_${group.isMonochrome ? 'bw' : 'color'}_${Date.now()}.pdf`);
      fs.writeFileSync(groupPdfPath, groupPdfBytes);
      
      // Calculate actual group number in original order (before reverse)
      const originalGroupNum = pageGroups.length - groupIndex;
      console.log(`🖨️ Printing Group ${originalGroupNum} (reverse order ${groupIndex + 1}/${pageGroups.length}): Pages ${group.startIndex + 1}-${group.endIndex + 1} (${group.isMonochrome ? 'B&W' : 'Color'})...`);
      
      // For B&W groups, force printer driver to grayscale mode for fast printing
      // For Color groups, force printer driver to color mode
      if (process.platform === 'win32') {
        if (group.isMonochrome) {
          await forceGrayscaleMode(printerName);
          await new Promise(resolve => setTimeout(resolve, 500));
        } else {
          await forceColorMode(printerName);
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      }
      
      // Print with Chrome (primary), SumatraPDF (fallback), or Windows print spooler (last resort)
      try {
        // Pass groupIndex for unique port calculation
        await printPdfWithChrome(groupPdfPath, printerName, group.isMonochrome, copies, groupIndex);
      } catch (error: any) {
        console.warn(`⚠️ Chrome failed for group, trying SumatraPDF: ${error.message}`);
        try {
          await printPdfWithSumatra(groupPdfPath, printerName, group.isMonochrome, copies);
        } catch (sumatraError: any) {
          console.warn(`⚠️ SumatraPDF failed for group, trying Windows spooler: ${sumatraError.message}`);
          await printPdfWithWindowsSpooler(groupPdfPath, printerName, group.isMonochrome, copies);
        }
      }
      console.log(`✅ Group ${originalGroupNum} printed successfully: Pages ${group.startIndex + 1}-${group.endIndex + 1} (${group.isMonochrome ? 'B&W' : 'Color'})`);
      
      // Cleanup: Delete file AFTER printing is complete and HTTP server has closed
      // The printPdfWithChrome function ensures the HTTP server closes before returning
      if (fs.existsSync(groupPdfPath)) {
        fs.unlinkSync(groupPdfPath);
        console.log(`🗑️ Cleaned up group PDF: ${groupPdfPath}`);
      }
      
      // Small delay between groups to ensure proper sequencing
      if (groupIndex < pageGroups.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
    }
    
    console.log(`✅ All ${totalPages} pages printed in correct sequence (stack-based reverse order)`);
    console.log(`📄 Final page order in output: Page 1 (top) → Page ${totalPages} (bottom)`);
  } catch (error: any) {
    console.error(`❌ Error printing PDF with mixed color in sequence: ${error.message}`);
    throw new Error(`Failed to print PDF with mixed color in sequence: ${error.message}`);
  }
}

/**
 * Normalize pageColors structure (handles both array and single object formats)
 * For single file printing, extracts the first element if it's an array
 */
function normalizePageColors(
  pageColors?: { colorPages: number[]; bwPages: number[] } | Array<{ colorPages: number[]; bwPages: number[] }>
): { colorPages: number[]; bwPages: number[] } | undefined {
  if (!pageColors) {
    return undefined;
  }
  
  // Handle array format (per-file) - extract first element for single file
  if (Array.isArray(pageColors)) {
    if (pageColors.length > 0) {
      return pageColors[0];
    }
    return undefined;
  }
  
  // Handle single object format (legacy)
  return pageColors;
}

/**
 * Force printer to grayscale mode using Windows PowerShell
 * This ensures the printer driver processes jobs as grayscale, not color
 */
async function forceGrayscaleMode(printerName: string): Promise<boolean> {
  if (process.platform !== 'win32') {
    return true;
  }

  try {
    console.log(`🎨 Attempting to force grayscale mode for printer: ${printerName}`);
    
    // Method 1: Use Set-PrintConfiguration to set ColorMode to Grayscale
    const setGrayscaleCmd = `powershell -Command "$printer = Get-Printer -Name '${printerName}' -ErrorAction SilentlyContinue; if ($printer) { Set-PrintConfiguration -PrinterName '${printerName}' -ColorMode Grayscale -ErrorAction SilentlyContinue; if ($?) { Write-Output 'SUCCESS' } else { Write-Output 'FAILED' } } else { Write-Output 'PRINTER_NOT_FOUND' }"`;
    
    try {
      const { stdout } = await execAsync(setGrayscaleCmd);
      const result = stdout.trim();
      
      if (result === 'SUCCESS') {
        console.log(`✅ Successfully set printer to grayscale mode via Set-PrintConfiguration`);
        return true;
      } else if (result === 'PRINTER_NOT_FOUND') {
        console.warn(`⚠️ Printer not found for grayscale configuration: ${printerName}`);
      } else {
        console.warn(`⚠️ Set-PrintConfiguration returned: ${result}`);
      }
    } catch (error: any) {
      console.warn(`⚠️ Set-PrintConfiguration failed: ${error.message}`);
    }
    
    // Method 2: Use printer properties via COM object (more reliable for some drivers)
    const comGrayscaleCmd = `powershell -Command "$printer = Get-Printer -Name '${printerName}' -ErrorAction SilentlyContinue; if ($printer) { $printerConfig = Get-PrintConfiguration -PrinterName '${printerName}' -ErrorAction SilentlyContinue; if ($printerConfig) { $printerConfig.ColorMode = 'Grayscale'; Set-PrintConfiguration -InputObject $printerConfig -ErrorAction SilentlyContinue; if ($?) { Write-Output 'SUCCESS' } else { Write-Output 'FAILED' } } else { Write-Output 'NO_CONFIG' } } else { Write-Output 'PRINTER_NOT_FOUND' }"`;
    
    try {
      const { stdout } = await execAsync(comGrayscaleCmd);
      const result = stdout.trim();
      
      if (result === 'SUCCESS') {
        console.log(`✅ Successfully set printer to grayscale mode via COM object`);
        return true;
      }
    } catch (error: any) {
      console.warn(`⚠️ COM object grayscale configuration failed: ${error.message}`);
    }
    
    console.warn(`⚠️ Could not force grayscale mode via automated methods, will rely on print job settings`);
    return false;
  } catch (error: any) {
    console.warn(`⚠️ Error forcing grayscale mode: ${error.message}`);
    return false;
  }
}

/**
 * Force printer to color mode using Windows PowerShell
 * This ensures the printer driver processes jobs as color, not grayscale
 */
async function forceColorMode(printerName: string): Promise<boolean> {
  if (process.platform !== 'win32') {
    return true;
  }

  try {
    console.log(`🎨 Attempting to force color mode for printer: ${printerName}`);
    
    // Method 1: Use Set-PrintConfiguration to set ColorMode to Color
    const setColorCmd = `powershell -Command "$printer = Get-Printer -Name '${printerName}' -ErrorAction SilentlyContinue; if ($printer) { Set-PrintConfiguration -PrinterName '${printerName}' -ColorMode Color -ErrorAction SilentlyContinue; if ($?) { Write-Output 'SUCCESS' } else { Write-Output 'FAILED' } } else { Write-Output 'PRINTER_NOT_FOUND' }"`;
    
    try {
      const { stdout } = await execAsync(setColorCmd);
      const result = stdout.trim();
      
      if (result === 'SUCCESS') {
        console.log(`✅ Successfully set printer to color mode via Set-PrintConfiguration`);
        return true;
      } else if (result === 'PRINTER_NOT_FOUND') {
        console.warn(`⚠️ Printer not found for color configuration: ${printerName}`);
      } else {
        console.warn(`⚠️ Set-PrintConfiguration returned: ${result}`);
      }
    } catch (error: any) {
      console.warn(`⚠️ Set-PrintConfiguration failed: ${error.message}`);
    }
    
    // Method 2: Use printer properties via COM object (more reliable for some drivers)
    const comColorCmd = `powershell -Command "$printer = Get-Printer -Name '${printerName}' -ErrorAction SilentlyContinue; if ($printer) { $printerConfig = Get-PrintConfiguration -PrinterName '${printerName}' -ErrorAction SilentlyContinue; if ($printerConfig) { $printerConfig.ColorMode = 'Color'; Set-PrintConfiguration -InputObject $printerConfig -ErrorAction SilentlyContinue; if ($?) { Write-Output 'SUCCESS' } else { Write-Output 'FAILED' } } else { Write-Output 'NO_CONFIG' } } else { Write-Output 'PRINTER_NOT_FOUND' }"`;
    
    try {
      const { stdout } = await execAsync(comColorCmd);
      const result = stdout.trim();
      
      if (result === 'SUCCESS') {
        console.log(`✅ Successfully set printer to color mode via COM object`);
        return true;
      }
    } catch (error: any) {
      console.warn(`⚠️ COM object color configuration failed: ${error.message}`);
    }
    
    console.warn(`⚠️ Could not force color mode via automated methods, will rely on print job settings`);
    return false;
  } catch (error: any) {
    console.warn(`⚠️ Error forcing color mode: ${error.message}`);
    return false;
  }
}

/**
 * Find Chrome executable path
 */
async function findChromePath(): Promise<string | null> {
  if (process.platform === 'win32') {
    // Common installation paths for Chrome
    const possiblePaths = [
      'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
      'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
      process.env['PROGRAMFILES'] + '\\Google\\Chrome\\Application\\chrome.exe',
      process.env['PROGRAMFILES(X86)'] + '\\Google\\Chrome\\Application\\chrome.exe',
      'C:\\Users\\' + process.env.USERNAME + '\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe',
      'chrome.exe' // Try PATH
    ];
    
    for (const possiblePath of possiblePaths) {
      if (possiblePath && fs.existsSync(possiblePath)) {
        console.log(`✅ Found Chrome at: ${possiblePath}`);
        return possiblePath;
      }
    }
    
    // Try to find in PATH using async method
    try {
      const { stdout } = await execAsync('where chrome.exe');
      if (stdout && stdout.trim()) {
        const path = stdout.trim().split('\n')[0].trim();
        if (path && fs.existsSync(path)) {
          console.log(`✅ Found Chrome in PATH: ${path}`);
          return path;
        }
      }
    } catch {
      // Not in PATH
    }
  }
  return null;
}

/**
 * Find Edge executable path
 */
async function findEdgePath(): Promise<string | null> {
  if (process.platform === 'win32') {
    // Common installation paths for Edge
    const possiblePaths = [
      'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
      'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
      process.env['PROGRAMFILES(X86)'] + '\\Microsoft\\Edge\\Application\\msedge.exe',
      process.env['PROGRAMFILES'] + '\\Microsoft\\Edge\\Application\\msedge.exe',
      'msedge.exe' // Try PATH
    ];
    
    for (const possiblePath of possiblePaths) {
      if (possiblePath && fs.existsSync(possiblePath)) {
        console.log(`✅ Found Edge at: ${possiblePath}`);
        return possiblePath;
      }
    }
    
    // Try to find in PATH using async method
    try {
      const { stdout } = await execAsync('where msedge.exe');
      if (stdout && stdout.trim()) {
        const path = stdout.trim().split('\n')[0].trim();
        if (path && fs.existsSync(path)) {
          console.log(`✅ Found Edge in PATH: ${path}`);
          return path;
        }
      }
    } catch {
      // Not in PATH
    }
  }
  return null;
}

/**
 * Find SumatraPDF executable path
 */
async function findSumatraPdfPath(): Promise<string | null> {
  if (process.platform === 'win32') {
    // Common installation paths for SumatraPDF
    const possiblePaths = [
      'C:\\Program Files\\SumatraPDF\\SumatraPDF.exe',
      'C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe',
      process.env['PROGRAMFILES'] + '\\SumatraPDF\\SumatraPDF.exe',
      process.env['PROGRAMFILES(X86)'] + '\\SumatraPDF\\SumatraPDF.exe',
      'C:\\Users\\' + process.env.USERNAME + '\\AppData\\Local\\SumatraPDF\\SumatraPDF.exe',
      'SumatraPDF.exe' // Try PATH
    ];
    
    for (const possiblePath of possiblePaths) {
      if (possiblePath && fs.existsSync(possiblePath)) {
        console.log(`✅ Found SumatraPDF at: ${possiblePath}`);
        return possiblePath;
      }
    }
    
    // Try to find in PATH using async method
    try {
      const { stdout } = await execAsync('where SumatraPDF.exe');
      if (stdout && stdout.trim()) {
        const path = stdout.trim().split('\n')[0].trim();
        if (path && fs.existsSync(path)) {
          console.log(`✅ Found SumatraPDF in PATH: ${path}`);
          return path;
        }
      }
    } catch {
      // Not in PATH
    }
                            }
  return null;
}

/**
 * Get current default printer name
 */
async function getDefaultPrinter(): Promise<string | null> {
  if (process.platform !== 'win32') {
    return null;
  }
  
  try {
    const { stdout } = await execAsync(`powershell -Command "$printer = Get-Printer | Where-Object {$_.Default -eq $true} | Select-Object -First 1; if ($printer) { Write-Output $printer.Name }"`);
    const printerName = stdout.trim();
    return printerName && printerName.length > 0 ? printerName : null;
  } catch {
    return null;
  }
}

/**
 * Set default printer temporarily
 */
async function setDefaultPrinter(printerName: string): Promise<boolean> {
  if (process.platform !== 'win32') {
    return false;
  }
  
  try {
    const escapedPrinterName = printerName.replace(/'/g, "''").replace(/"/g, '\\"');
    const { stdout } = await execAsync(`powershell -Command "rundll32 printui.dll,PrintUIEntry /y /n '${escapedPrinterName}'"`);
    return true;
  } catch {
    return false;
  }
}

/**
 * Print PDF using Chrome with print dialog (allows color mode control)
 * This opens Chrome normally and uses print dialog interaction to set color mode
 */
async function printPdfWithChrome(
  filePath: string,
  printerName: string,
  isMonochrome: boolean,
  copies: number,
  groupIndex: number = 0
): Promise<void> {
  if (process.platform !== 'win32') {
    throw new Error('Chrome printing only works on Windows');
  }

  // Try Chrome first, then Edge
  let browserPath = await findChromePath();
  let browserName = 'Chrome';
  
  if (!browserPath) {
    browserPath = await findEdgePath();
    browserName = 'Edge';
  }
  
  if (!browserPath) {
    throw new Error('Chrome or Edge not found. Please install Google Chrome or Microsoft Edge.');
  }

  try {
    console.log(`⚡ Using ${browserName} with print dialog for color mode control (monochrome=${isMonochrome})`);
    
    // Force printer driver to appropriate mode BEFORE printing
    if (isMonochrome) {
      console.log(`🎨 Forcing printer driver to grayscale mode before printing...`);
      await forceGrayscaleMode(printerName);
      await new Promise(resolve => setTimeout(resolve, 500));
    } else {
      console.log(`🎨 Forcing printer driver to color mode before printing...`);
      await forceColorMode(printerName);
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    
    // Save current default printer
    const originalDefaultPrinter = await getDefaultPrinter();
    
    // Set target printer as default
    console.log(`🖨️ Setting printer as default: ${printerName}`);
    const setDefaultSuccess = await setDefaultPrinter(printerName);
    if (!setDefaultSuccess) {
      console.warn(`⚠️ Could not set printer as default, Chrome will use current default printer`);
    }
    
    const escapedBrowserPath = browserPath.replace(/"/g, '\\"');
    
    // Print all copies
    for (let i = 0; i < copies; i++) {
      if (i > 0) {
        await new Promise(resolve => setTimeout(resolve, 2000)); // Delay between copies
      }
      
      // Declare server outside try block so it's accessible in catch block
      let server: http.Server | null = null;
      
      try {
        // Use a local HTTP server to serve the PDF (avoids local file access restrictions)
        // Use unique ports for each group and copy: basePort + (groupIndex * 10) + copyIndex
        const serverPort = 8765 + (groupIndex * 10) + i;
        const fileName = path.basename(filePath);
        const httpUrl = `http://localhost:${serverPort}/${encodeURIComponent(fileName)}`;
        
        // Create a simple HTTP server to serve the PDF
        server = http.createServer((req, res) => {
          const parsedUrl = url.parse(req.url || '/');
          const requestedFile = decodeURIComponent(parsedUrl.pathname?.substring(1) || '');
          
          // Check if the requested file matches (handle both encoded and decoded URLs)
          if (requestedFile === fileName || requestedFile === encodeURIComponent(fileName) || decodeURIComponent(requestedFile) === fileName) {
            // Check if file exists before trying to serve it
            if (!fs.existsSync(filePath)) {
              console.error(`❌ File not found: ${filePath}`);
              res.writeHead(404, { 'Content-Type': 'text/plain' });
              res.end('File not found');
              return;
            }
            
            // Serve the PDF file
            try {
              const fileStream = fs.createReadStream(filePath);
              fileStream.on('error', (err: any) => {
                console.error(`❌ Error reading file: ${err.message}`);
                if (!res.headersSent) {
                  res.writeHead(500, { 'Content-Type': 'text/plain' });
                  res.end('Error reading file');
                }
              });
              
              res.writeHead(200, {
                'Content-Type': 'application/pdf',
                'Content-Disposition': `inline; filename="${fileName}"`,
                'Access-Control-Allow-Origin': '*',
                'Cache-Control': 'no-cache'
              });
              fileStream.pipe(res);
            } catch (error: any) {
              console.error(`❌ Error serving file: ${error.message}`);
              if (!res.headersSent) {
                res.writeHead(500, { 'Content-Type': 'text/plain' });
                res.end('Error serving file');
              }
            }
          } else {
            // Redirect to PDF directly - Chrome will open it in PDF viewer
            res.writeHead(302, { 'Location': httpUrl });
            res.end();
          }
        });
        
        // Start the server
        if (!server) {
          throw new Error('Failed to create HTTP server');
        }
        await new Promise<void>((resolve, reject) => {
          server!.listen(serverPort, 'localhost', () => {
            console.log(`📡 Started local HTTP server on port ${serverPort} to serve PDF`);
            resolve();
          });
          server!.on('error', (err: any) => {
            if (err.code === 'EADDRINUSE') {
              // Port in use, try next port
              serverPort + 1;
              resolve();
            } else {
              reject(err);
            }
          });
        });
        
        // Check if file exists before proceeding
        if (!fs.existsSync(filePath)) {
          throw new Error(`File not found: ${filePath}`);
        }
        
        // Calculate wait time based on file size (larger files need more time)
        const fileStats = fs.statSync(filePath);
        const fileSizeMB = fileStats.size / (1024 * 1024);
        const baseWaitTime = 5; // Base wait time in seconds
        const sizeBasedWait = Math.max(0, Math.ceil(fileSizeMB * 2)); // 2 seconds per MB
        const totalWaitTime = baseWaitTime + sizeBasedWait;
        
        console.log(`Executing ${browserName} print command (monochrome=${isMonochrome}, copies=${copies}, copy ${i + 1}/${copies})`);
        console.log(`   Using local HTTP server to serve PDF (avoids local file restrictions)`);
        console.log(`   Opening PDF directly in ${browserName} PDF viewer, then triggering print`);
        console.log(`   File size: ${fileSizeMB.toFixed(2)} MB - will wait ${totalWaitTime} seconds for PDF to load`);
        console.log(`   PDF URL: ${httpUrl}`);
        
        // Verify HTTP server is ready before launching Chrome
        try {
          await new Promise<void>((resolve, reject) => {
            const testReq = http.get(`http://localhost:${serverPort}/`, (res) => {
              res.on('data', () => {});
              res.on('end', () => {
                console.log(`   HTTP server verification: Ready (status ${res.statusCode})`);
                resolve();
              });
            });
            testReq.on('error', (err) => {
              console.log(`   HTTP server test failed (will continue anyway): ${err.message}`);
              resolve(); // Continue anyway
            });
            testReq.setTimeout(2000, () => {
              testReq.destroy();
              console.log(`   HTTP server test timeout (will continue anyway)`);
              resolve(); // Continue anyway
            });
          });
        } catch (testError) {
          console.log(`   HTTP server test skipped, proceeding with Chrome launch`);
        }
        
        // Launch Chrome and keep it running
        // Ensure Chrome recognizes the URL as a PDF - use proper argument formatting
        // Chrome needs the URL to be passed as a single argument, properly quoted
        // Use single quotes around the URL in PowerShell to prevent it from being treated as a search query
        // NOTE: We don't use --kiosk-printing because we need to control color mode via print dialog
        // Added flags to keep Chrome open: --no-first-run, --no-default-browser-check, --disable-extensions
        const escapedPdfUrlForChrome = httpUrl.replace(/'/g, "''");
        const chromeCmd = `powershell -Command "$ErrorActionPreference = 'Stop'; $url = '${escapedPdfUrlForChrome}'; $proc = Start-Process -FilePath '${escapedBrowserPath}' -ArgumentList '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--new-window', $url -PassThru -ErrorAction Stop; $procId = $proc.Id; Write-Output $procId"`;
        
        console.log(`   Launching Chrome with URL: ${httpUrl}`);
        console.log(`   Chrome command: ${chromeCmd.substring(0, 200)}...`); // Log first 200 chars of command
        const { stdout, stderr } = await execAsync(chromeCmd);
        const initialProcId = stdout.trim();
        if (stderr) {
          console.log(`   Chrome launch stderr: ${stderr.trim()}`);
        }
        console.log(`   Chrome launcher process started with PID: ${initialProcId}`);
        
        // Wait longer for Chrome to fully start (launcher exits, actual Chrome process starts)
        console.log(`   Waiting 3 seconds for Chrome to fully initialize...`);
        await new Promise(resolve => setTimeout(resolve, 3000));
        
        // Find the actual Chrome window process (not the launcher)
        // The launcher process exits quickly, we need to find the actual Chrome browser process
        let procId: string | null = null;
        let chromeStillRunning = false;
        
        // First, try to find Chrome process by the initial PID (in case it's still valid)
        try {
          const checkInitialCmd = `powershell -Command "$proc = Get-Process -Id ${initialProcId} -ErrorAction SilentlyContinue; if ($proc -and -not $proc.HasExited) { Write-Output 'RUNNING' } else { Write-Output 'NOT_RUNNING' }"`;
          const { stdout: initialCheck } = await execAsync(checkInitialCmd);
          if (initialCheck.trim() === 'RUNNING') {
            procId = initialProcId;
            chromeStillRunning = true;
            console.log(`   Initial Chrome process ${initialProcId} is still running`);
          }
        } catch (initialError) {
          console.log(`   Initial process ${initialProcId} not found, searching for Chrome window process...`);
        }
        
        // If initial PID doesn't work, find Chrome process by name with a window
        if (!chromeStillRunning) {
          try {
            const browserProcessName = browserName === 'Chrome' ? 'chrome' : 'msedge';
            
            // Extract PDF filename from file path for window title matching
            const pdfFileName = path.basename(filePath);
            const pdfFileNameBase = path.basename(filePath, path.extname(filePath)); // Without extension
            
            console.log(`   Searching for ${browserName} processes with windows...`);
            console.log(`   Looking for window title containing: ${pdfFileName}`);
            
            // Strategy 1: Search by window title containing PDF filename (most reliable)
            let findChromeCmd = `powershell -Command "$procs = Get-Process -Name '${browserProcessName}' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' -and ($_.MainWindowTitle -like '*${pdfFileNameBase}*' -or $_.MainWindowTitle -like '*${pdfFileName}*') } | Sort-Object StartTime -Descending | Select-Object -First 1; if ($procs) { Write-Output $procs.Id } else { Write-Output 'NOT_FOUND' }"`;
            
            let { stdout: chromeProcId, stderr: findStderr } = await execAsync(findChromeCmd);
            let foundProcId = chromeProcId.trim();
            
            if (findStderr) {
              console.log(`   Process search stderr: ${findStderr.trim()}`);
            }
            
            // Strategy 2: If window title search fails, search for most recent Chrome process with any window (within last 60 seconds)
            if (!foundProcId || foundProcId === 'NOT_FOUND' || isNaN(parseInt(foundProcId))) {
              console.log(`   Window title search failed, trying recent processes (last 60 seconds)...`);
              findChromeCmd = `powershell -Command "$procs = Get-Process -Name '${browserProcessName}' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' -and $_.StartTime -gt (Get-Date).AddSeconds(-60) } | Sort-Object StartTime -Descending | Select-Object -First 1; if ($procs) { Write-Output $procs.Id } else { Write-Output 'NOT_FOUND' }"`;
              
              try {
                const result = await execAsync(findChromeCmd);
                foundProcId = result.stdout.trim();
                if (result.stderr) {
                  console.log(`   Recent process search stderr: ${result.stderr.trim()}`);
                }
              } catch (recentError) {
                console.warn(`   Recent process search failed: ${recentError}`);
              }
            }
            
            // Strategy 3: If still not found, find any Chrome process with a window (no time restriction)
            if (!foundProcId || foundProcId === 'NOT_FOUND' || isNaN(parseInt(foundProcId))) {
              console.log(`   Recent process search failed, trying any process with window...`);
              findChromeCmd = `powershell -Command "$procs = Get-Process -Name '${browserProcessName}' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Sort-Object StartTime -Descending | Select-Object -First 1; if ($procs) { Write-Output $procs.Id } else { Write-Output 'NOT_FOUND' }"`;
              
              try {
                const result = await execAsync(findChromeCmd);
                foundProcId = result.stdout.trim();
                if (result.stderr) {
                  console.log(`   Any process search stderr: ${result.stderr.trim()}`);
                }
              } catch (anyError) {
                console.warn(`   Any process search failed: ${anyError}`);
              }
            }
            
            if (foundProcId && foundProcId !== 'NOT_FOUND' && !isNaN(parseInt(foundProcId))) {
              procId = foundProcId;
              chromeStillRunning = true;
              console.log(`   Found Chrome window process with PID: ${procId}`);
              
              // Verify the window title for debugging
              try {
                const verifyCmd = `powershell -Command "$proc = Get-Process -Id ${procId} -ErrorAction SilentlyContinue; if ($proc) { Write-Output $proc.MainWindowTitle } else { Write-Output 'NOT_FOUND' }"`;
                const { stdout: windowTitle } = await execAsync(verifyCmd);
                console.log(`   Chrome window title: ${windowTitle.trim()}`);
              } catch (verifyError) {
                // Ignore verification errors
              }
            } else {
              // List all Chrome processes for debugging
              try {
                const listProcsCmd = `powershell -Command "Get-Process -Name '${browserProcessName}' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object Id, ProcessName, MainWindowTitle, StartTime | Format-List"`;
                const { stdout: allProcs } = await execAsync(listProcsCmd);
                if (allProcs && allProcs.trim()) {
                  console.log(`   All ${browserName} processes with windows found:\n${allProcs}`);
                } else {
                  console.log(`   No ${browserName} processes with windows found`);
                }
              } catch (listError) {
                console.warn(`   Could not list processes: ${listError}`);
              }
              console.warn(`   Could not find Chrome window process`);
            }
          } catch (findError) {
            console.warn(`   Error finding Chrome process: ${findError}`);
          }
        }
        
        // Final verification with retry
        if (procId && chromeStillRunning) {
          for (let retry = 0; retry < 3; retry++) {
            try {
              const checkProcCmd = `powershell -Command "$proc = Get-Process -Id ${procId} -ErrorAction SilentlyContinue; if ($proc -and -not $proc.HasExited) { Write-Output 'RUNNING' } else { Write-Output 'NOT_RUNNING' }"`;
              const { stdout: procCheck } = await execAsync(checkProcCmd);
              chromeStillRunning = procCheck.trim() === 'RUNNING';
              if (chromeStillRunning) {
                console.log(`   Verified Chrome process ${procId} is running`);
                break;
              }
              if (retry < 2) {
                console.log(`   Chrome process check failed, retrying... (${retry + 1}/3)`);
                await new Promise(resolve => setTimeout(resolve, 1000));
              }
            } catch (checkError) {
              console.warn(`   ⚠️ Could not check Chrome process status: ${checkError}`);
              if (retry < 2) {
                await new Promise(resolve => setTimeout(resolve, 1000));
              }
            }
          }
        }
        
        if (!chromeStillRunning || !procId) {
          console.warn(`   ⚠️ Chrome process is not running after launch, cannot send print command`);
          console.warn(`   Initial PID: ${initialProcId}, Found PID: ${procId || 'N/A'}`);
          throw new Error('Chrome process exited immediately after launch or could not be found');
        }
        
        // At this point, procId is guaranteed to be non-null
        const finalProcId = procId;
        
        // Wait for PDF to fully load (longer for larger files)
        console.log(`   Waiting ${totalWaitTime} seconds for PDF to load...`);
        await new Promise(resolve => setTimeout(resolve, totalWaitTime * 1000));
        
        // Verify Chrome is still running before attempting to interact
        try {
          const checkProcCmd = `powershell -Command "$proc = Get-Process -Id ${finalProcId} -ErrorAction SilentlyContinue; if ($proc -and -not $proc.HasExited) { Write-Output 'RUNNING' } else { Write-Output 'NOT_RUNNING' }"`;
          const { stdout: procCheck } = await execAsync(checkProcCmd);
          chromeStillRunning = procCheck.trim() === 'RUNNING';
          if (!chromeStillRunning) {
            console.warn(`   ⚠️ Chrome process ${finalProcId} exited while waiting for PDF to load`);
            throw new Error('Chrome process exited while waiting for PDF to load');
          }
        } catch (checkError) {
          console.warn(`   ⚠️ Could not verify Chrome process status: ${checkError}`);
          throw new Error('Chrome process verification failed');
        }
        
        // Trigger printing with color mode setting
        if (chromeStillRunning) {
          try {
            // Build the print command that interacts with Chrome's print dialog
            // Chrome's print dialog structure after Ctrl+P:
            // - Focus starts on "Pages" dropdown
            // - Color dropdown is at 8 tabs from the start
            // - Color dropdown default is "Black and white"
            // - One Down from "Black and white" = "Color"
            let printCmd = `powershell -Command "Add-Type -AssemblyName Microsoft.VisualBasic; Add-Type -AssemblyName System.Windows.Forms; $proc = Get-Process -Id ${finalProcId} -ErrorAction SilentlyContinue; if ($proc -and -not $proc.HasExited) { [Microsoft.VisualBasic.Interaction]::AppActivate($finalProcId); Start-Sleep -Milliseconds 1000; `;
            
            // Open print dialog (Ctrl+P)
            printCmd += `[System.Windows.Forms.SendKeys]::SendWait('^p'); Start-Sleep -Seconds 3; `;
            
            // Tab 8 times to reach Color dropdown (from Pages dropdown)
            printCmd += `[System.Windows.Forms.SendKeys]::SendWait('{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}'); Start-Sleep -Milliseconds 500; `;
            // Open Color dropdown (Alt+Down or Space)
            printCmd += `[System.Windows.Forms.SendKeys]::SendWait('%{DOWN}'); Start-Sleep -Milliseconds 500; `;
            
            if (isMonochrome) {
              // For B&W printing: Navigate to Color dropdown and explicitly select "Black and white"
              console.log(`   Setting Chrome print dialog to Black and white mode...`);
              // Select "Black and white" option (Up arrow to ensure we're on "Black and white")
              printCmd += `[System.Windows.Forms.SendKeys]::SendWait('{UP}'); Start-Sleep -Milliseconds 400; `;
            } else {
              // For Color printing: Navigate to Color dropdown and select "Color"
              console.log(`   Setting Chrome print dialog to Color mode...`);
              // Select "Color" option (Down arrow once, as "Black and white" is default, "Color" is one down)
              printCmd += `[System.Windows.Forms.SendKeys]::SendWait('{DOWN}'); Start-Sleep -Milliseconds 400; `;
            }
            
            // Confirm selection (Enter)
            printCmd += `[System.Windows.Forms.SendKeys]::SendWait('{ENTER}'); Start-Sleep -Milliseconds 400; `;
            // Tab 2 times to reach Print button after color selection
            printCmd += `[System.Windows.Forms.SendKeys]::SendWait('{TAB}{TAB}'); Start-Sleep -Milliseconds 400; `;
            
            // Press Enter to print
            printCmd += `[System.Windows.Forms.SendKeys]::SendWait('{ENTER}'); Write-Output 'Print command sent with color mode set' } else { Write-Output 'Process not found or exited' }"`;
            
            const printResult = await execAsync(printCmd);
            console.log(`   Print command result: ${printResult.stdout.trim()}`);
          } catch (printError: any) {
            console.warn(`   ⚠️ Failed to send print command: ${printError.message}`);
            // Try fallback: just print without color mode change
            try {
              const fallbackCmd = `powershell -Command "Add-Type -AssemblyName Microsoft.VisualBasic; Add-Type -AssemblyName System.Windows.Forms; $proc = Get-Process -Id ${finalProcId} -ErrorAction SilentlyContinue; if ($proc -and -not $proc.HasExited) { [Microsoft.VisualBasic.Interaction]::AppActivate(${finalProcId}); Start-Sleep -Milliseconds 500; [System.Windows.Forms.SendKeys]::SendWait('^p'); Start-Sleep -Seconds 2; [System.Windows.Forms.SendKeys]::SendWait('{ENTER}'); Write-Output 'Fallback print sent' }"`;
              await execAsync(fallbackCmd);
              console.log(`   Fallback print command sent`);
            } catch (fallbackError) {
              console.warn(`   ⚠️ Fallback print also failed: ${fallbackError}`);
            }
          }
        } else {
          console.warn(`   ⚠️ Cannot print - Chrome process is not running`);
          throw new Error('Chrome process exited before print command could be sent');
        }
        
        // Wait for print job to be queued and processed (longer for larger files)
        const printWaitTime = Math.max(5, Math.ceil(fileSizeMB * 3)); // 3 seconds per MB minimum 5 seconds
        console.log(`   Waiting ${printWaitTime} seconds for print job to be processed...`);
        await new Promise(resolve => setTimeout(resolve, printWaitTime * 1000));
        
        // Keep Chrome open a bit longer to ensure print job completes
        // Wait 10 seconds before closing Chrome
        await new Promise(resolve => setTimeout(resolve, 10000));
        
        // Close Chrome process
        try {
          await execAsync(`powershell -Command "$proc = Get-Process -Id ${finalProcId} -ErrorAction SilentlyContinue; if ($proc -and -not $proc.HasExited) { Stop-Process -Id ${finalProcId} -Force -ErrorAction SilentlyContinue; Write-Output 'Chrome closed' }"`);
        } catch (cleanupError) {
          // Ignore cleanup errors
        }
        
        // Close the HTTP server and wait for it to close before continuing
        // This ensures the file exists for the entire print operation
        if (server) {
          await new Promise<void>((resolve) => {
            server!.close(() => {
              console.log(`   HTTP server closed`);
              resolve();
            });
          });
        }
        
        // File cleanup happens in the calling function after all operations complete
        // This ensures the HTTP server has finished serving the file
        
      } catch (error: any) {
        // Close server even on error
        if (server) {
          try {
            await new Promise<void>((resolve) => {
              server!.close(() => {
                console.log(`   HTTP server closed (error cleanup)`);
                resolve();
              });
            });
          } catch (serverError) {
            // Ignore server close errors
          }
        }
        
        if (i === 0) {
          throw error;
        }
        console.warn(`⚠️ Failed to print copy ${i + 1}: ${error.message}`);
      }
    }
    
    // Restore original default printer if we changed it
    if (originalDefaultPrinter && originalDefaultPrinter !== printerName && setDefaultSuccess) {
      console.log(`🔄 Restoring original default printer: ${originalDefaultPrinter}`);
      await setDefaultPrinter(originalDefaultPrinter);
    }
    
    console.log(`✅ ${browserName} print command completed successfully (${copies} copy/copies)`);
    console.log(`   Mode: ${isMonochrome ? 'Grayscale/B&W (fast mode - direct printing)' : 'Color'}`);
    
    await new Promise(resolve => setTimeout(resolve, 1000));
  } catch (error: any) {
    console.error(`❌ ${browserName} print error: ${error.message}`);
    throw error;
  }
}

/**
 * Print PDF using Windows print spooler API directly (fastest method, similar to Chrome)
 * This bypasses application rendering and sends PDF directly to printer driver
 */
async function printPdfWithWindowsSpooler(
  filePath: string,
  printerName: string,
  isMonochrome: boolean,
  copies: number
): Promise<void> {
  if (process.platform !== 'win32') {
    throw new Error('Windows print spooler only works on Windows');
  }

  try {
    console.log(`⚡ Using Windows print spooler API for fastest printing (monochrome=${isMonochrome})`);
    
    // CRITICAL: Force printer driver to appropriate mode BEFORE printing
    // This ensures the printer processes the job in the correct mode
    if (isMonochrome) {
      console.log(`🎨 Forcing printer driver to grayscale mode before printing...`);
      await forceGrayscaleMode(printerName);
      // Small delay to ensure settings are applied
      await new Promise(resolve => setTimeout(resolve, 500));
    } else {
      console.log(`🎨 Forcing printer driver to color mode before printing...`);
      await forceColorMode(printerName);
      // Small delay to ensure settings are applied
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    
    const escapedPrinterName = printerName.replace(/'/g, "''").replace(/"/g, '\\"');
    const escapedFilePath = filePath.replace(/'/g, "''").replace(/\\/g, '\\\\');
    
    // Use PowerShell to print directly via Windows print spooler
    // This method is similar to how Chrome prints - it uses the default PDF handler
    // but with explicit printer driver settings
    let printCmd = `powershell -Command "$ErrorActionPreference = 'Stop'; `;
    
    // Set printer to appropriate mode (redundant but ensures it's set)
    if (isMonochrome) {
      printCmd += `try { Set-PrintConfiguration -PrinterName '${escapedPrinterName}' -ColorMode Grayscale -ErrorAction SilentlyContinue | Out-Null } catch {}; `;
    } else {
      printCmd += `try { Set-PrintConfiguration -PrinterName '${escapedPrinterName}' -ColorMode Color -ErrorAction SilentlyContinue | Out-Null } catch {}; `;
    }
    
    // Use Start-Process with Print verb - this uses Windows default PDF handler
    // which respects printer driver settings (grayscale mode)
    // The printer driver is already set to grayscale, so it will process as grayscale (fast)
    printCmd += `$proc = Start-Process -FilePath '${escapedFilePath}' -Verb Print -WindowStyle Hidden -PassThru -ErrorAction Stop; `;
    printCmd += `Start-Sleep -Seconds 1; `;
    printCmd += `if (-not $proc.HasExited) { Write-Host 'Print job queued' } else { Write-Host 'Print process completed' }"`;
    
    console.log(`Executing Windows print spooler command (monochrome=${isMonochrome}, copies=${copies})`);
    console.log(`   Printer driver is in ${isMonochrome ? 'grayscale mode for fast B&W printing' : 'color mode'}`);
    
    // Print all copies
    for (let i = 0; i < copies; i++) {
      if (i > 0) {
        await new Promise(resolve => setTimeout(resolve, 1500)); // Delay between copies
      }
      
      try {
        const { stdout, stderr } = await execAsync(printCmd);
        
        if (stdout && stdout.trim()) {
          console.log(`Print spooler output: ${stdout.trim()}`);
        }
        
        if (stderr && !stderr.includes('request id')) {
          const stderrLower = stderr.toLowerCase();
          if (stderrLower.includes('error') || stderrLower.includes('exception')) {
            if (i === 0) {
              throw new Error(`Windows print spooler error: ${stderr.trim()}`);
            }
            console.warn(`⚠️ Failed to print copy ${i + 1}: ${stderr.trim()}`);
          }
        }
      } catch (error: any) {
        // If first copy fails, throw error
        if (i === 0) {
          throw error;
        }
        // For additional copies, log warning but continue
        console.warn(`⚠️ Failed to print copy ${i + 1}: ${error.message}`);
      }
    }
    
    console.log(`✅ Windows print spooler command completed successfully (${copies} copy/copies)`);
    console.log(`   Mode: ${isMonochrome ? 'Grayscale/B&W (fast mode - direct spooler)' : 'Color'}`);
    
    // Wait for print job to be queued
    await new Promise(resolve => setTimeout(resolve, 1000));
  } catch (error: any) {
    console.error(`❌ Windows print spooler error: ${error.message}`);
    throw error;
  }
}

/**
 * Print PDF using SumatraPDF command-line (PRIMARY method for optimal speed)
 * SumatraPDF provides direct control over print settings and respects printer driver grayscale mode
 */
async function printPdfWithSumatra(
  filePath: string,
  printerName: string,
  isMonochrome: boolean,
  copies: number
): Promise<void> {
  if (process.platform !== 'win32') {
    throw new Error('SumatraPDF printing only works on Windows');
  }

  const sumatraPath = await findSumatraPdfPath();
  if (!sumatraPath) {
    throw new Error('SumatraPDF not found. Please install SumatraPDF from https://www.sumatrapdfreader.org/download-free-pdf-viewer');
  }

  try {
    console.log(`⚡ Using SumatraPDF for optimal printing performance (monochrome=${isMonochrome})`);
    
    // Force printer driver to appropriate mode BEFORE printing
    if (isMonochrome) {
      console.log(`🎨 Forcing printer driver to grayscale mode before printing...`);
      await forceGrayscaleMode(printerName);
      await new Promise(resolve => setTimeout(resolve, 500));
    } else {
      console.log(`🎨 Forcing printer driver to color mode before printing...`);
      await forceColorMode(printerName);
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    
    const escapedPrinterName = printerName.replace(/"/g, '\\"');
    const escapedFilePath = filePath.replace(/"/g, '\\"');
    
    // Build SumatraPDF command
    let sumatraCmd = `"${sumatraPath}" -print-to "${escapedPrinterName}" -silent`;
    
    if (isMonochrome) {
      sumatraCmd += ` -print-settings "monochrome"`;
    }
    
    sumatraCmd += ` "${escapedFilePath}"`;
    
    console.log(`Executing SumatraPDF print command (monochrome=${isMonochrome}, copies=${copies})`);
    
    // Print all copies
    for (let i = 0; i < copies; i++) {
      if (i > 0) {
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
      
      try {
        const { stdout, stderr } = await execAsync(sumatraCmd);
        
        if (stdout && stdout.trim()) {
          console.log(`SumatraPDF output: ${stdout.trim()}`);
        }
        
        if (stderr && stderr.trim()) {
          const stderrLower = stderr.toLowerCase();
          if (stderrLower.includes('error') && !stderrLower.includes('warning')) {
            throw new Error(`SumatraPDF error: ${stderr.trim()}`);
          }
        }
      } catch (error: any) {
        if (i === 0) {
          throw error;
        }
        console.warn(`⚠️ Failed to print copy ${i + 1}: ${error.message}`);
      }
    }
    
    console.log(`✅ SumatraPDF print command completed successfully (${copies} copy/copies)`);
    console.log(`   Mode: ${isMonochrome ? 'Grayscale/B&W (fast mode)' : 'Color'}`);
    
    await new Promise(resolve => setTimeout(resolve, 1000));
  } catch (error: any) {
    console.error(`❌ SumatraPDF print error: ${error.message}`);
    throw error;
  }
}

/**
 * Print file using system printer command
 */
async function printFile(filePath: string, options: PrintJob['printingOptions']): Promise<void> {
  console.log(`🔍 DEBUG - printFile called with:`);
  console.log(`   File: ${filePath}`);
  console.log(`   Color mode: ${options.color}`);
  console.log(`   pageColors:`, JSON.stringify(options.pageColors, null, 2));
  
  const printerName = await getPrinterName();
  const printerPath = process.env.PRINTER_PATH;
  
  // Check if printer is available before printing
  const printerStatus = await checkPrinterStatus();
  if (!printerStatus.available) {
    throw new Error(`Printer is not available: ${printerStatus.message}. ${printerStatus.details || ''}`);
  }

  // Determine print command based on OS
  const isWindows = process.platform === 'win32';
  const isMac = process.platform === 'darwin';
  const isLinux = process.platform === 'linux';

  // Determine color mode, copies, page size, and sided based on printing options
  // Default to black and white unless explicitly set to 'color'
  const colorMode = options.color === 'color' ? 'color' : (options.color === 'mixed' ? 'mixed' : 'bw');
  const copies = options.copies || 1;
  const pageSize = options.pageSize || 'A4';
  const sided = options.sided || 'single';
  
  // Ensure monochrome is true by default (black and white) unless explicitly color
  // Note: 'mixed' mode should not default to monochrome - it will be handled separately
  const isMonochrome = colorMode === 'bw';
  
  const fileExt = path.extname(filePath).toLowerCase();

  // For Windows: Use SumatraPDF for PDFs and images, Word COM for DOCX
  if (isWindows) {
    // For Word files with mixed color mode, they will be converted to PDF first
    // Then mixed color mode will be applied to the converted PDF
    if ((fileExt === '.docx' || fileExt === '.doc') && colorMode === 'mixed') {
      console.log(`ℹ️ Word file with mixed color mode detected. Will convert to PDF first, then apply mixed color printing.`);
      // Continue to Word conversion below
    }
    
    if (fileExt === '.pdf' || fileExt === '.jpg' || fileExt === '.jpeg' || fileExt === '.png' || fileExt === '.gif' || fileExt === '.bmp') {
      // Handle mixed color printing for PDFs (maintains page sequence)
      if (fileExt === '.pdf' && colorMode === 'mixed') {
        console.log(`🔍 DEBUG - PDF file with mixed color mode detected`);
        console.log(`🔍 DEBUG - pageColors received:`, JSON.stringify(options.pageColors, null, 2));
        
        // Normalize pageColors (handle both array and single object formats)
        const normalizedPageColors = normalizePageColors(options.pageColors);
        console.log(`🔍 DEBUG - normalized pageColors:`, JSON.stringify(normalizedPageColors, null, 2));
        
        // Validate pageColors structure
        if (!normalizedPageColors) {
          console.warn(`⚠️ Mixed color mode requested but pageColors is missing. Defaulting to B&W mode.`);
          // Default to B&W mode using Chrome (fastest)
          try {
            await printPdfWithChrome(filePath, printerName, true, copies);
          } catch (error: any) {
            console.warn(`⚠️ Chrome failed, using SumatraPDF: ${error.message}`);
            try {
              await printPdfWithSumatra(filePath, printerName, true, copies);
            } catch (sumatraError: any) {
              console.warn(`⚠️ SumatraPDF failed, using Windows spooler: ${sumatraError.message}`);
              await printPdfWithWindowsSpooler(filePath, printerName, true, copies);
            }
          }
          await new Promise(resolve => setTimeout(resolve, 2000));
          return;
        } else if (!Array.isArray(normalizedPageColors.colorPages) || !Array.isArray(normalizedPageColors.bwPages)) {
          console.warn(`⚠️ Mixed color mode requested but pageColors structure is invalid.`);
          console.warn(`   Expected: { colorPages: number[], bwPages: number[] }`);
          console.warn(`   Received:`, JSON.stringify(normalizedPageColors, null, 2));
          console.warn(`   Defaulting to B&W mode.`);
          // Default to B&W mode using Chrome (fastest)
          try {
            await printPdfWithChrome(filePath, printerName, true, copies);
          } catch (error: any) {
            console.warn(`⚠️ Chrome failed, using SumatraPDF: ${error.message}`);
            try {
              await printPdfWithSumatra(filePath, printerName, true, copies);
            } catch (sumatraError: any) {
              console.warn(`⚠️ SumatraPDF failed, using Windows spooler: ${sumatraError.message}`);
              await printPdfWithWindowsSpooler(filePath, printerName, true, copies);
            }
          }
          await new Promise(resolve => setTimeout(resolve, 2000));
          return;
        } else {
          // Valid pageColors structure - use mixed color printing
          try {
            console.log(`🖨️ Printing PDF with mixed color mode (maintaining page sequence)`);
            console.log(`✅ Valid pageColors structure detected for mixed color printing`);
            console.log(`📋 Color pages: ${normalizedPageColors.colorPages.join(', ')}`);
            console.log(`📋 B&W pages: ${normalizedPageColors.bwPages.join(', ')}`);
            
            // Create temp directory
            const tempDir = path.join(process.cwd(), 'temp');
            if (!fs.existsSync(tempDir)) {
              fs.mkdirSync(tempDir, { recursive: true });
            }
            
            // Print PDF with mixed color pages in sequence
            await printPdfWithMixedColorInSequence(
              filePath,
              normalizedPageColors.colorPages,
              normalizedPageColors.bwPages,
              printerName,
              copies,
              pageSize,
              sided,
              tempDir
            );
            
            // Wait a bit for the print jobs to be queued
            await new Promise(resolve => setTimeout(resolve, 2000));
            return; // Success, exit early
          } catch (error: any) {
            console.error(`❌ Mixed color printing error: ${error.message}`);
            throw new Error(`Mixed color print failed: ${error.message}`);
          }
        }
      }
      
      // Use Chrome with --kiosk-printing as PRIMARY method (fastest, similar to Chrome's built-in printing)
      // Falls back to SumatraPDF, then Windows print spooler if Chrome is not available
      try {
        console.log(`🖨️ Printing ${fileExt} file using Chrome with --kiosk-printing (primary method): ${filePath}`);
        console.log(`📋 Options: printer=${printerName}, copies=${copies}, color=${colorMode}, pageSize=${pageSize}, sided=${sided}`);
        
        // Print using Chrome (primary method)
        await printPdfWithChrome(filePath, printerName, isMonochrome, copies);
        console.log(`✅ Print job sent successfully using Chrome`);
        
        // Wait a bit for the print job to be queued
        await new Promise(resolve => setTimeout(resolve, 2000));
        return; // Success, exit early
      } catch (error: any) {
        console.warn(`⚠️ Chrome failed: ${error.message}`);
        console.log(`🔄 Falling back to SumatraPDF...`);
        
        // Fallback to SumatraPDF
        try {
          await printPdfWithSumatra(filePath, printerName, isMonochrome, copies);
          console.log(`✅ Print job sent successfully using SumatraPDF (fallback)`);
          await new Promise(resolve => setTimeout(resolve, 2000));
          return; // Success with fallback
        } catch (sumatraError: any) {
          console.warn(`⚠️ SumatraPDF failed: ${sumatraError.message}`);
          console.log(`🔄 Falling back to Windows print spooler...`);
          
          // Last resort: Windows print spooler
          try {
            await printPdfWithWindowsSpooler(filePath, printerName, isMonochrome, copies);
            console.log(`✅ Print job sent successfully using Windows print spooler (last resort)`);
            await new Promise(resolve => setTimeout(resolve, 2000));
            return; // Success with last resort
          } catch (spoolerError: any) {
            console.error(`❌ All print methods failed`);
            throw new Error(`Print failed: Chrome: ${error.message}, SumatraPDF: ${sumatraError.message}, Windows spooler: ${spoolerError.message}`);
          }
        }
      }
    } else if (fileExt === '.docx' || fileExt === '.doc') {
      // For Word files: Convert to PDF first using LibreOffice, then print as PDF
      try {
        console.log(`🔄 Converting Word file to PDF before printing: ${filePath}`);
        console.log(`🔍 DEBUG - Color mode: ${colorMode}`);
        console.log(`🔍 DEBUG - pageColors received:`, JSON.stringify(options.pageColors, null, 2));
        
        const pdfPath = await convertWordToPdf(filePath);
        
        // Now print the converted PDF
        console.log(`🖨️ Printing converted PDF: ${pdfPath}`);
        
        // Handle mixed color printing for converted PDFs if needed
        if (colorMode === 'mixed') {
          // Normalize pageColors (handle both array and single object formats)
          const normalizedPageColors = normalizePageColors(options.pageColors);
          console.log(`🔍 DEBUG - normalized pageColors:`, JSON.stringify(normalizedPageColors, null, 2));
          
          // Validate pageColors structure
          if (!normalizedPageColors) {
            console.warn(`⚠️ Mixed color mode requested but pageColors is missing. Defaulting to B&W mode.`);
            // Default to B&W mode using Chrome (fastest)
            try {
              await printPdfWithChrome(pdfPath, printerName, true, copies);
            } catch (error: any) {
              console.warn(`⚠️ Chrome failed, using SumatraPDF: ${error.message}`);
              try {
                await printPdfWithSumatra(pdfPath, printerName, true, copies);
              } catch (sumatraError: any) {
                console.warn(`⚠️ SumatraPDF failed, using Windows spooler: ${sumatraError.message}`);
                await printPdfWithWindowsSpooler(pdfPath, printerName, true, copies);
              }
            }
          } else if (!Array.isArray(normalizedPageColors.colorPages) || !Array.isArray(normalizedPageColors.bwPages)) {
            console.warn(`⚠️ Mixed color mode requested but pageColors structure is invalid.`);
            console.warn(`   Expected: { colorPages: number[], bwPages: number[] }`);
            console.warn(`   Received:`, JSON.stringify(normalizedPageColors, null, 2));
            console.warn(`   Defaulting to B&W mode.`);
            // Default to B&W mode using Chrome (fastest)
            try {
              await printPdfWithChrome(pdfPath, printerName, true, copies);
            } catch (error: any) {
              console.warn(`⚠️ Chrome failed, using SumatraPDF: ${error.message}`);
              try {
                await printPdfWithSumatra(pdfPath, printerName, true, copies);
              } catch (sumatraError: any) {
                console.warn(`⚠️ SumatraPDF failed, using Windows spooler: ${sumatraError.message}`);
                await printPdfWithWindowsSpooler(pdfPath, printerName, true, copies);
              }
            }
          } else {
            // Valid pageColors structure - use mixed color printing
            console.log(`✅ Valid pageColors structure detected for mixed color printing`);
            console.log(`📋 Color pages: ${normalizedPageColors.colorPages.join(', ')}`);
            console.log(`📋 B&W pages: ${normalizedPageColors.bwPages.join(', ')}`);
            
            const tempDir = path.join(process.cwd(), 'temp');
            await printPdfWithMixedColorInSequence(
              pdfPath,
              normalizedPageColors.colorPages,
              normalizedPageColors.bwPages,
              printerName,
              copies,
              pageSize,
              sided,
              tempDir
            );
          }
        } else {
          // Regular PDF printing (not mixed mode) using Chrome (fastest, same as regular PDFs)
          try {
            await printPdfWithChrome(pdfPath, printerName, isMonochrome, copies);
            console.log(`✅ Print job sent successfully using Chrome`);
          } catch (error: any) {
            console.warn(`⚠️ Chrome failed: ${error.message}`);
            console.log(`🔄 Falling back to SumatraPDF...`);
            
            // Fallback to SumatraPDF
            try {
              await printPdfWithSumatra(pdfPath, printerName, isMonochrome, copies);
              console.log(`✅ Print job sent successfully using SumatraPDF (fallback)`);
            } catch (sumatraError: any) {
              console.warn(`⚠️ SumatraPDF failed: ${sumatraError.message}`);
              console.log(`🔄 Falling back to Windows print spooler...`);
              
              // Last resort: Windows print spooler
              try {
                await printPdfWithWindowsSpooler(pdfPath, printerName, isMonochrome, copies);
                console.log(`✅ Print job sent successfully using Windows print spooler (last resort)`);
              } catch (spoolerError: any) {
                console.error(`❌ All print methods failed`);
                throw new Error(`Print failed: Chrome: ${error.message}, SumatraPDF: ${sumatraError.message}, Windows spooler: ${spoolerError.message}`);
              }
            }
          }
        }
        
        // Clean up converted PDF file
        try {
          if (fs.existsSync(pdfPath)) {
            fs.unlinkSync(pdfPath);
            console.log(`🗑️ Cleaned up temporary PDF: ${pdfPath}`);
          }
        } catch (cleanupError) {
          console.warn(`⚠️ Could not clean up temporary PDF: ${cleanupError}`);
        }
        
        // Wait a bit for the print job to be queued
        await new Promise(resolve => setTimeout(resolve, 2000));
        console.log(`✅ Word file printed successfully (converted to PDF first)`);
        return; // Success, exit early
        
      } catch (conversionError: any) {
        console.error(`❌ Word to PDF conversion failed: ${conversionError.message}`);
        console.log(`⚠️ Falling back to Word COM object printing...`);
        // Fall through to Word COM object method below
      }
    }
  }

  // For other file types or platforms, use command-based approach
  let printCommand: string;

  if (isWindows) {
    // Windows: Use Word COM object for DOCX (fallback if conversion failed), default application for other types
    const escapedPrinterName = printerName.replace(/'/g, "''").replace(/\\/g, '\\\\');
    const escapedFilePath = filePath.replace(/'/g, "''").replace(/\\/g, '\\\\');
    
    if (fileExt === '.docx' || fileExt === '.doc') {
      // For DOCX: Use Word COM object to print with basic options
      // Simplified version that works with all Word versions
      // Word.PrintOut parameters: Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint
      
      // Determine duplex mode (0=None, 1=Long-edge, 2=Short-edge)
      const duplexMode = sided === 'double' ? 1 : 0; // 1 = Long-edge binding for double-sided
      
      // Simplified PrintOut command - removed problematic Options properties that don't exist in all Word versions
      // Using only essential parameters: Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint
      printCommand = `powershell -Command "$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open('${escapedFilePath}'); $word.ActivePrinter = '${escapedPrinterName}'; $doc.PrintOut([ref]$false, [ref]$false, [ref]0, [ref]'', [ref]0, [ref]0, [ref]0, [ref]${copies}, [ref]'', [ref]0, [ref]$false, [ref]$true, [ref]'', [ref]'', [ref]${duplexMode}); $doc.Close([ref]$false); $word.Quit([ref]$false)"`;
    } else {
      // For other file types: Try default application
      printCommand = `powershell -Command "$file = '${escapedFilePath}'; Start-Process -FilePath $file -Verb Print -WindowStyle Hidden"`;
    }
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
  
  try {
    const { stdout, stderr } = await execAsync(printCommand);
    
    // Shell.Application COM object typically doesn't output much
    // If there's output, check for errors
    if (stdout) {
      const stdoutLower = stdout.toLowerCase();
      const stdoutTrimmed = stdout.trim();
      
      // Check for COM object errors
      if (stdoutLower.includes('new-object') && stdoutLower.includes('cannot create type')) {
        console.error(`❌ COM object error detected in stdout: ${stdoutTrimmed}`);
        throw new Error(`COM object error: Shell.Application not available. ${stdoutTrimmed}`);
      }
      
      if (stdoutLower.includes('invokeverbex') && (stdoutLower.includes('method invocation failed') || stdoutLower.includes('exception'))) {
        console.error(`❌ Print verb error detected in stdout: ${stdoutTrimmed}`);
        throw new Error(`Print verb failed: ${stdoutTrimmed || 'Unable to invoke print verb'}`);
      }
      
      // Check for general printer errors
      if (stdoutLower.includes('unable to initialize device') ||
          stdoutLower.includes('unable to connect') ||
          stdoutLower.includes('printer not found') ||
          stdoutLower.includes('printer does not exist') ||
          stdoutLower.includes('device not found') ||
          stdoutLower.includes('cannot connect to printer') ||
          stdoutLower.includes('cannot find') ||
          (stdoutLower.includes('error') && !stdoutLower.includes('request id')) ||
          (stdoutLower.includes('exception') && !stdoutLower.includes('request id'))) {
        console.error(`❌ Printer error detected in stdout: ${stdoutTrimmed}`);
        throw new Error(`Printer error: ${stdoutTrimmed || 'Unable to print'}`);
      }
      
      if (stdoutTrimmed) {
        console.log(`✅ Print command output: ${stdoutTrimmed}`);
      }
    }
    
    // Check for printer errors in stderr
    if (stderr) {
      const stderrLower = stderr.toLowerCase();
      const stderrTrimmed = stderr.trim();
      
      // COM object specific errors
      if (stderrLower.includes('new-object') && (stderrLower.includes('cannot create type') || stderrLower.includes('comobject'))) {
        console.error(`❌ COM object error detected in stderr: ${stderrTrimmed}`);
        throw new Error(`COM object error: Shell.Application not available. ${stderrTrimmed}`);
      }
      
      if (stderrLower.includes('invokeverbex') && (stderrLower.includes('method invocation failed') || stderrLower.includes('exception'))) {
        console.error(`❌ Print verb error detected in stderr: ${stderrTrimmed}`);
        throw new Error(`Print verb failed: ${stderrTrimmed || 'Unable to invoke print verb'}`);
      }
      
      // Common printer error messages
      if (stderrLower.includes('unable to connect') || 
          stderrLower.includes('printer not found') ||
          stderrLower.includes('no such file or directory') ||
          stderrLower.includes('printer does not exist') ||
          stderrLower.includes('cannot find') ||
          (stderrLower.includes('error') && !stderrLower.includes('request id')) ||
          (stderrLower.includes('exception') && !stderrLower.includes('request id'))) {
        console.error(`❌ Printer error detected in stderr: ${stderrTrimmed}`);
        throw new Error(`Printer error: ${stderrTrimmed || 'Unable to print'}`);
      }
      
      if (stderrLower.includes('printer is not available') ||
          stderrLower.includes('printer is offline') ||
          stderrLower.includes('printer is stopped')) {
        console.error(`❌ Printer error detected in stderr: ${stderrTrimmed}`);
        throw new Error(`Printer is offline or not available: ${printerName}`);
      }
      
      // Some systems output "request id" in stderr which is normal
      if (!stderrLower.includes('request id') && !stderrLower.includes('request-id') && stderrTrimmed) {
        console.warn('Print command stderr:', stderrTrimmed);
      }
    }
    
    // For Shell.Application COM object, if no error, assume success
    // Wait a bit for the print job to be queued
    await new Promise(resolve => setTimeout(resolve, 2000));
  } catch (error: any) {
    // Check if it's a printer-related error
    // Shell.Application COM object may output errors to stdout or stderr
    const errorMessage = error.message || error.stdout || error.stderr || String(error);
    const errorLower = errorMessage.toLowerCase();
    
    // Detect COM object errors and try fallback method
    if (errorLower.includes('new-object') && (errorLower.includes('cannot create type') || errorLower.includes('comobject'))) {
      console.warn('⚠️ COM object method failed, trying fallback method...');
      return await tryFallbackPrintMethod(filePath, printerName, options);
    }
    
    if (errorLower.includes('invokeverbex') && (errorLower.includes('method invocation failed') || errorLower.includes('exception'))) {
      console.warn('⚠️ Print verb failed, trying fallback method...');
      return await tryFallbackPrintMethod(filePath, printerName, options);
    }
    
    // Detect specific printer errors
    if (errorLower.includes('unable to initialize device') ||
        errorLower.includes('unable to connect') ||
        errorLower.includes('printer not found') ||
        errorLower.includes('no such file or directory') ||
        errorLower.includes('printer does not exist') ||
        errorLower.includes('device not found') ||
        errorLower.includes('cannot connect to printer')) {
      throw new Error(`Printer not connected: ${printerName}. Please check USB connection.`);
    }
    
    if (errorLower.includes('printer is not available') ||
        errorLower.includes('printer is offline') ||
        errorLower.includes('printer is idle') ||
        errorLower.includes('printer is stopped') ||
        errorLower.includes('printer is disabled')) {
      throw new Error(`Printer is offline: ${printerName}. Please turn on the printer.`);
    }
    
    if (errorLower.includes('power') && errorLower.includes('off')) {
      throw new Error(`Printer appears to be powered off: ${printerName}. Please turn on the printer.`);
    }
    
    // Re-throw original error if not a known printer error
    throw error;
  }
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

    // Print the file FIRST (will be at bottom of stack)
    console.log(`Printing file: ${tempFilePath}`);
    await printFile(tempFilePath, job.printingOptions);
    console.log(`File printed successfully`);

    // Print order summary page LAST (will appear on top due to stack-based printing - LIFO)
    if (job.orderDetails && job.customerInfo) {
      console.log(`Printing order summary page...`);
      const orderSummaryPage = await generateOrderSummaryPage(job.orderDetails, job.customerInfo);
      const orderSummaryPath = path.join(tempDir, `order_summary_${job.deliveryNumber}.pdf`);
      fs.writeFileSync(orderSummaryPath, orderSummaryPage);
      await printFile(orderSummaryPath, { ...job.printingOptions, copies: 1 });
      fs.unlinkSync(orderSummaryPath);
      console.log(`Order summary page printed successfully`);
    } else {
      console.log(`⏭️ Skipping order summary page (orderDetails or customerInfo not provided)`);
      console.log(`   orderDetails: ${job.orderDetails ? 'provided' : 'missing'}`);
      console.log(`   customerInfo: ${job.customerInfo ? 'provided' : 'missing'}`);
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
    
    // Extract error message
    const errorMessage = error instanceof Error ? error.message : String(error);
    const errorLower = errorMessage.toLowerCase();
    
    // Detect specific error types and provide helpful messages
    let userMessage = 'Failed to print job';
    let errorDetails = errorMessage;
    
    if (errorLower.includes('printer not connected') ||
        errorLower.includes('unable to connect') ||
        errorLower.includes('printer not found')) {
      userMessage = 'Printer not connected';
      errorDetails = 'Please check USB connection and ensure printer is powered on';
    } else if (errorLower.includes('printer is offline') ||
               errorLower.includes('printer is not available') ||
               errorLower.includes('printer is stopped')) {
      userMessage = 'Printer is offline';
      errorDetails = 'Printer may be powered off or disconnected. Please check power and USB connection.';
    } else if (errorLower.includes('power') && errorLower.includes('off')) {
      userMessage = 'Printer appears to be powered off';
      errorDetails = 'Please turn on the printer and try again';
    } else if (errorLower.includes('timeout') || errorLower.includes('timed out')) {
      userMessage = 'Print job timed out';
      errorDetails = 'Printer may be busy or not responding. Will retry automatically.';
    }
    
    return {
      success: false,
      message: userMessage,
      error: errorDetails,
      deliveryNumber: job.deliveryNumber
    };
  }
}

/**
 * Check if printer is available
 */
export async function checkPrinterStatus(): Promise<{ available: boolean; message: string; details?: string }> {
  const printerName = await getPrinterName();
  
  try {
    const isWindows = process.platform === 'win32';
    const isMac = process.platform === 'darwin';
    const isLinux = process.platform === 'linux';

    let checkCommand: string;
    let parseCommand: string | null = null;

    if (isWindows) {
      // Use PowerShell Get-Printer (modern Windows) with fallback to wmic
      // PowerShell command: Get-Printer -Name "HP_Deskjet_525" | Select-Object Name, PrinterStatus
      checkCommand = `powershell -Command "Get-Printer -Name '${printerName}' -ErrorAction SilentlyContinue | Select-Object Name, PrinterStatus | Format-List"`;
    } else if (isMac || isLinux) {
      // Use lpstat to check printer status
      checkCommand = `lpstat -p "${printerName}"`;
      // Also check if printer is enabled and accepting jobs
      parseCommand = `lpstat -p "${printerName}" -l`;
    } else {
      return { available: false, message: 'Unsupported platform' };
    }

    let stdout: string = '';
    let stderr: string = '';
    
    try {
      const result = await execAsync(checkCommand);
      stdout = result.stdout || '';
      stderr = result.stderr || '';
    } catch (error: any) {
      // If PowerShell fails, try wmic as fallback (for older Windows)
      if (isWindows) {
        try {
          const wmicCommand = `wmic printer where name="${printerName}" get name,status`;
          const wmicResult = await execAsync(wmicCommand);
          stdout = wmicResult.stdout || '';
          stderr = wmicResult.stderr || '';
        } catch (wmicError: any) {
          // Both PowerShell and wmic failed
          stderr = wmicError.stderr || wmicError.message || error.message || '';
          stdout = error.stdout || '';
        }
      } else {
        stderr = error.stderr || error.message || '';
        stdout = error.stdout || '';
      }
    }
    
    // Check for errors in stderr
    if (stderr) {
      const stderrLower = stderr.toLowerCase();
      if (stderrLower.includes('printer not found') ||
          stderrLower.includes('unable to connect') ||
          stderrLower.includes('no such file or directory')) {
        return { 
          available: false, 
          message: `Printer not connected: ${printerName}`,
          details: 'Please check USB connection and ensure printer is powered on'
        };
      }
    }
    
    // Parse output for printer status
    const outputLower = stdout.toLowerCase();
    
    // Check if printer was found (PowerShell returns empty if printer not found)
    if (!stdout || stdout.trim().length === 0) {
      return { 
        available: false, 
        message: `Printer not found: ${printerName}`,
        details: 'Printer may not be installed or connected. Please check USB connection and ensure printer is powered on.'
      };
    }
    
    // PowerShell Get-Printer status values: Normal, Warning, Error, Unknown
    // wmic status values: Idle, Printing, Ready, etc.
    // Check for available status
    if (outputLower.includes('normal') || 
        outputLower.includes('idle') || 
        outputLower.includes('printing') || 
        outputLower.includes('ready') ||
        outputLower.includes('online')) {
      // Try to get more details
      if (parseCommand) {
        try {
          const { stdout: details } = await execAsync(parseCommand);
          return { 
            available: true, 
            message: 'Printer is available',
            details: details.trim()
          };
        } catch {
          // Ignore parse command errors
        }
      }
      return { available: true, message: 'Printer is available' };
    }
    
    // Check for offline/stopped/error status
    if (outputLower.includes('offline') || 
        outputLower.includes('stopped') || 
        outputLower.includes('disabled') ||
        outputLower.includes('error') ||
        outputLower.includes('warning')) {
      return { 
        available: false, 
        message: `Printer is offline: ${printerName}`,
        details: 'Printer may be powered off or disconnected. Please check power and USB connection.'
      };
    }
    
    // If we got here, printer exists but status is unclear
    return { available: true, message: 'Printer is available (status unclear)' };
  } catch (error: any) {
    const errorMessage = error.message || error.stderr || String(error);
    const errorLower = errorMessage.toLowerCase();
    
    // Detect specific error types
    if (errorLower.includes('printer not found') ||
        errorLower.includes('unable to connect') ||
        errorLower.includes('no such file or directory') ||
        errorLower.includes('printer does not exist')) {
      return { 
        available: false, 
        message: `Printer not connected: ${printerName}`,
        details: 'Please check USB connection and ensure printer is powered on'
      };
    }
    
    if (errorLower.includes('printer is not available') ||
        errorLower.includes('printer is offline')) {
      return { 
        available: false, 
        message: `Printer is offline: ${printerName}`,
        details: 'Printer may be powered off. Please turn on the printer.'
      };
    }
    
    return { 
      available: false, 
      message: error instanceof Error ? error.message : 'Printer check failed',
      details: 'Unable to determine printer status. Please check printer connection and power.'
    };
  }
}

