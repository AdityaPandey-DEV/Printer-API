/**
 * Delivery Number Generation Service
 * Format: {LETTER}{YYYYMMDD}{PRINTER_INDEX}{FILE_NUMBER}
 * Example: A2025011511 = Letter A, Date 2025-01-15, Printer Index 1, File Number 1
 * 
 * Logic:
 * - A-Z cycles every 260 files (10 files per letter × 26 letters)
 * - After every 10 files, print a full-page letter separator (A, B, C...)
 * - After every file, print a blank page separator
 */

interface DeliveryNumberState {
  currentLetter: string;
  currentCount: number; // 1-10 for current letter
  currentFileNumber: number; // 1-10 for current file number in letter cycle
  totalFiles: number; // Total files printed since start
  lastDate: string; // YYYYMMDD format
}

let deliveryState: DeliveryNumberState = {
  currentLetter: process.env.DELIVERY_NUMBER_START || 'A',
  currentCount: 0,
  currentFileNumber: 0,
  totalFiles: 0,
  lastDate: getCurrentDateString()
};

function getCurrentDateString(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

function getNextLetter(currentLetter: string): string {
  if (currentLetter === 'Z') {
    return 'A'; // Cycle back to A
  }
  return String.fromCharCode(currentLetter.charCodeAt(0) + 1);
}

/**
 * Generate delivery number for a print job
 * @param printerIndex - Index of the printer (from PRINTER_API_URLS array)
 * @returns Delivery number in format {LETTER}{YYYYMMDD}{PRINTER_INDEX}{FILE_NUMBER}
 */
export function generateDeliveryNumber(printerIndex: number): string {
  const currentDate = getCurrentDateString();
  
  // Reset letter and count if date changed
  if (currentDate !== deliveryState.lastDate) {
    deliveryState.currentLetter = process.env.DELIVERY_NUMBER_START || 'A';
    deliveryState.currentCount = 0;
    deliveryState.currentFileNumber = 0;
    deliveryState.lastDate = currentDate;
  }

  // Increment count for current letter
  deliveryState.currentCount++;
  deliveryState.totalFiles++;
  
  // Increment file number (1-10)
  deliveryState.currentFileNumber++;
  if (deliveryState.currentFileNumber > 10) {
    deliveryState.currentFileNumber = 1;
  }

  // Move to next letter after 10 files
  if (deliveryState.currentCount > 10) {
    deliveryState.currentLetter = getNextLetter(deliveryState.currentLetter);
    deliveryState.currentCount = 1;
    deliveryState.currentFileNumber = 1; // Reset file number for new letter
  }

  // Format: {LETTER}{YYYYMMDD}{PRINTER_INDEX}{FILE_NUMBER}
  const deliveryNumber = `${deliveryState.currentLetter}${currentDate}${printerIndex}${deliveryState.currentFileNumber}`;
  
  console.log(`Generated delivery number: ${deliveryNumber} (Letter: ${deliveryState.currentLetter}, File Number: ${deliveryState.currentFileNumber}, Count: ${deliveryState.currentCount}, Total: ${deliveryState.totalFiles})`);
  
  return deliveryNumber;
}

/**
 * Check if we need to print a letter separator (every 10 files)
 */
export function shouldPrintLetterSeparator(): boolean {
  return deliveryState.currentCount === 1 && deliveryState.totalFiles > 0 && 
         (deliveryState.totalFiles % 10 === 1 || deliveryState.totalFiles === 1);
}

/**
 * Get current letter for separator printing
 */
export function getCurrentLetter(): string {
  return deliveryState.currentLetter;
}

/**
 * Get current file number (1-10)
 */
export function getCurrentFileNumber(): number {
  return deliveryState.currentFileNumber;
}

/**
 * Reset delivery number state (for testing or manual reset)
 */
export function resetDeliveryState() {
  deliveryState = {
    currentLetter: process.env.DELIVERY_NUMBER_START || 'A',
    currentCount: 0,
    currentFileNumber: 0,
    totalFiles: 0,
    lastDate: getCurrentDateString()
  };
}

/**
 * Get current delivery state (for debugging)
 */
export function getDeliveryState(): DeliveryNumberState {
  return { ...deliveryState };
}

