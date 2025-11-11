export interface PrintJobData {
  // Legacy: single file (for backward compatibility)
  fileUrl?: string;
  fileName?: string;
  fileType?: string;
  // Multiple files support
  fileURLs?: string[];
  originalFileNames?: string[];
  fileTypes?: string[];
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
  deliveryNumber?: string;
  orderId?: string;
  customerInfo?: {
    name: string;
    email: string;
    phone: string;
  };
}

export interface PrintJobResponse {
  success: boolean;
  message: string;
  jobId?: string;
  deliveryNumber?: string;
  error?: string;
}

