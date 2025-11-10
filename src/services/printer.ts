import * as fs from 'fs';
import * as path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import { print as printWithPdfToPrinter, getPrinters } from 'pdf-to-printer';
import { PDFDocument } from 'pdf-lib';
import { generateFileNumberPage, generateLetterSeparator } from '../utils/printerUtils';
import { shouldPrintLetterSeparator, getCurrentLetter, getCurrentFileNumber } from './deliveryNumber';

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
 * Split PDF into color and B&W pages for mixed color printing
 */
async function splitPdfForMixedColor(pdfPath: string, colorPages: number[], bwPages: number[], tempDir: string): Promise<{ colorPdfPath: string | null; bwPdfPath: string | null }> {
  try {
    const pdfBytes = fs.readFileSync(pdfPath);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const totalPages = pdfDoc.getPageCount();
    
    // Convert 1-based page numbers to 0-based indices
    const colorIndices = colorPages.map(p => p - 1).filter(i => i >= 0 && i < totalPages);
    const bwIndices = bwPages.map(p => p - 1).filter(i => i >= 0 && i < totalPages);
    
    // Determine unspecified pages (default to B&W)
    const allSpecifiedIndices = [...new Set([...colorIndices, ...bwIndices])];
    const unspecifiedIndices: number[] = [];
    for (let i = 0; i < totalPages; i++) {
      if (!allSpecifiedIndices.includes(i)) {
        unspecifiedIndices.push(i);
      }
    }
    
    // Add unspecified pages to B&W
    const allBwIndices = [...new Set([...bwIndices, ...unspecifiedIndices])];
    
    let colorPdfPath: string | null = null;
    let bwPdfPath: string | null = null;
    
    // Create color PDF if there are color pages
    if (colorIndices.length > 0) {
      const colorPdf = await PDFDocument.create();
      for (const index of colorIndices) {
        const [copiedPage] = await colorPdf.copyPages(pdfDoc, [index]);
        colorPdf.addPage(copiedPage);
      }
      const colorPdfBytes = await colorPdf.save();
      colorPdfPath = path.join(tempDir, `color_${Date.now()}.pdf`);
      fs.writeFileSync(colorPdfPath, colorPdfBytes);
      console.log(`✅ Created color PDF with ${colorIndices.length} pages: ${colorPdfPath}`);
    }
    
    // Create B&W PDF if there are B&W pages
    if (allBwIndices.length > 0) {
      const bwPdf = await PDFDocument.create();
      for (const index of allBwIndices) {
        const [copiedPage] = await bwPdf.copyPages(pdfDoc, [index]);
        bwPdf.addPage(copiedPage);
      }
      const bwPdfBytes = await bwPdf.save();
      bwPdfPath = path.join(tempDir, `bw_${Date.now()}.pdf`);
      fs.writeFileSync(bwPdfPath, bwPdfBytes);
      console.log(`✅ Created B&W PDF with ${allBwIndices.length} pages: ${bwPdfPath}`);
    }
    
    return { colorPdfPath, bwPdfPath };
  } catch (error: any) {
    console.error(`❌ Error splitting PDF for mixed color: ${error.message}`);
    throw new Error(`Failed to split PDF for mixed color printing: ${error.message}`);
  }
}

/**
 * Print file using system printer command
 */
async function printFile(filePath: string, options: PrintJob['printingOptions']): Promise<void> {
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
  const colorMode = options.color === 'color' ? 'color' : (options.color === 'mixed' ? 'mixed' : 'bw');
  const copies = options.copies || 1;
  const pageSize = options.pageSize || 'A4';
  const sided = options.sided || 'single';
  
  const fileExt = path.extname(filePath).toLowerCase();

  // For Windows: Use pdf-to-printer for PDFs and images, Word COM for DOCX
  if (isWindows) {
    if (fileExt === '.pdf' || fileExt === '.jpg' || fileExt === '.jpeg' || fileExt === '.png' || fileExt === '.gif' || fileExt === '.bmp') {
      // Handle mixed color printing for PDFs
      if (fileExt === '.pdf' && colorMode === 'mixed' && options.pageColors) {
        try {
          console.log(`🖨️ Printing PDF with mixed color mode`);
          console.log(`📋 Color pages: ${options.pageColors.colorPages.join(', ')}`);
          console.log(`📋 B&W pages: ${options.pageColors.bwPages.join(', ')}`);
          
          // Create temp directory
          const tempDir = path.join(process.cwd(), 'temp');
          if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir, { recursive: true });
          }
          
          // Split PDF into color and B&W pages
          const { colorPdfPath, bwPdfPath } = await splitPdfForMixedColor(
            filePath,
            options.pageColors.colorPages,
            options.pageColors.bwPages,
            tempDir
          );
          
          // Print color pages
          if (colorPdfPath) {
            console.log(`🖨️ Printing color pages...`);
            const colorPrintOptions: any = {
              printer: printerName,
              copies: copies,
              monochrome: false, // Color printing
              paperSize: pageSize === 'A3' ? 'A3' : 'A4'
            };
            if (sided === 'double') {
              colorPrintOptions.duplex = 'duplexlong';
            }
            await printWithPdfToPrinter(colorPdfPath, colorPrintOptions);
            console.log(`✅ Color pages printed successfully`);
            
            // Cleanup
            fs.unlinkSync(colorPdfPath);
          }
          
          // Print B&W pages
          if (bwPdfPath) {
            console.log(`🖨️ Printing B&W pages...`);
            const bwPrintOptions: any = {
              printer: printerName,
              copies: copies,
              monochrome: true, // B&W printing
              paperSize: pageSize === 'A3' ? 'A3' : 'A4'
            };
            if (sided === 'double') {
              bwPrintOptions.duplex = 'duplexlong';
            }
            await printWithPdfToPrinter(bwPdfPath, bwPrintOptions);
            console.log(`✅ B&W pages printed successfully`);
            
            // Cleanup
            fs.unlinkSync(bwPdfPath);
          }
          
          // Wait a bit for the print jobs to be queued
          await new Promise(resolve => setTimeout(resolve, 2000));
          return; // Success, exit early
        } catch (error: any) {
          console.error(`❌ Mixed color printing error: ${error.message}`);
          throw new Error(`Mixed color print failed: ${error.message}`);
        }
      }
      
      // Use pdf-to-printer directly for non-mixed color printing (no command string needed)
      try {
        console.log(`🖨️ Printing ${fileExt} file using pdf-to-printer: ${filePath}`);
        console.log(`📋 Options: printer=${printerName}, copies=${copies}, color=${colorMode}, pageSize=${pageSize}, sided=${sided}`);
        
        // Convert options to pdf-to-printer format
        const printOptions: any = {
          printer: printerName,
          copies: copies,
          monochrome: colorMode === 'bw',
          paperSize: pageSize === 'A3' ? 'A3' : 'A4'
        };
        
        // Add duplex if double-sided
        if (sided === 'double') {
          printOptions.duplex = 'duplexlong';
        }
        
        // Print using pdf-to-printer
        await printWithPdfToPrinter(filePath, printOptions);
        console.log(`✅ Print job sent successfully using pdf-to-printer`);
        
        // Wait a bit for the print job to be queued
        await new Promise(resolve => setTimeout(resolve, 2000));
        return; // Success, exit early
      } catch (error: any) {
        console.error(`❌ pdf-to-printer error: ${error.message}`);
        throw new Error(`Print failed: ${error.message}`);
      }
    } else if (fileExt === '.docx' || fileExt === '.doc') {
      // For DOCX: Use Word COM object (keep existing method)
      // This will continue to the command-based approach below
    }
  }

  // For other file types or platforms, use command-based approach
  let printCommand: string;

  if (isWindows) {
    // Windows: Use Word COM object for DOCX, default application for other types
    const escapedPrinterName = printerName.replace(/'/g, "''").replace(/\\/g, '\\\\');
    const escapedFilePath = filePath.replace(/'/g, "''").replace(/\\/g, '\\\\');
    
    if (fileExt === '.docx' || fileExt === '.doc') {
      // For DOCX: Use Word COM object to print with all options
      // Word.PrintOut parameters:
      // Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight
      
      // Determine duplex mode (0=None, 1=Long-edge, 2=Short-edge)
      const duplexMode = sided === 'double' ? 1 : 0; // 1 = Long-edge binding for double-sided
      
      // Determine color mode (True=Color, False=Grayscale)
      // For 'mixed', use color mode (let printer handle it)
      const useColor = colorMode !== 'bw'; // True for 'color' or 'mixed'
      
      // Build PrintOut command with all options
      // Parameters: Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight
      const useColorStr = useColor ? '$true' : '$false';
      printCommand = `powershell -Command "$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open('${escapedFilePath}'); $word.ActivePrinter = '${escapedPrinterName}'; $word.Options.PrintBackgroundColors = ${useColorStr}; $word.Options.PrintBackgroundImages = ${useColorStr}; $word.Options.PrintInColor = ${useColorStr}; $doc.PrintOut([ref]$false, [ref]$false, [ref]0, [ref]'', [ref]0, [ref]0, [ref]0, [ref]${copies}, [ref]'', [ref]0, [ref]$false, [ref]$true, [ref]'', [ref]'', [ref]${duplexMode}, [ref]0, [ref]0, [ref]0, [ref]0); $doc.Close([ref]$false); $word.Quit([ref]$false)"`;
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

    // Extract file number from delivery number (last digit)
    // Delivery number format: {LETTER}{YYYYMMDD}{PRINTER_INDEX}{FILE_NUMBER}
    // Example: A2025111011 = file number 1, A2025111012 = file number 2
    let fileNumber = 1;
    if (job.deliveryNumber && job.deliveryNumber.length > 0) {
      const lastChar = job.deliveryNumber.charAt(job.deliveryNumber.length - 1);
      const parsedFileNumber = parseInt(lastChar, 10);
      if (!isNaN(parsedFileNumber) && parsedFileNumber >= 1 && parsedFileNumber <= 10) {
        fileNumber = parsedFileNumber;
      }
    } else {
      // Fallback to getCurrentFileNumber if delivery number is not available
      fileNumber = getCurrentFileNumber();
    }

    // Check if we need to print letter separator BEFORE the file (after every 10 files, before file 1 of new letter)
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

    // Print file number page separator BEFORE the file
    console.log(`Printing file number separator: File no: ${fileNumber}...`);
    const fileNumberPage = await generateFileNumberPage(fileNumber);
    const fileNumberPagePath = path.join(tempDir, `file_${fileNumber}_${job.deliveryNumber}.pdf`);
    fs.writeFileSync(fileNumberPagePath, fileNumberPage);
    await printFile(fileNumberPagePath, { ...job.printingOptions, copies: 1 });
    fs.unlinkSync(fileNumberPagePath);
    console.log(`File number separator printed: File no: ${fileNumber}`);

    // Add a small delay after printing separator to ensure printer is ready
    await new Promise(resolve => setTimeout(resolve, 2000));

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

