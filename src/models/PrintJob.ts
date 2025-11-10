export interface PrintJobData {
  fileUrl: string;
  fileName: string;
  fileType: string;
  printingOptions: {
    pageSize: 'A4' | 'A3';
    color: 'color' | 'bw' | 'mixed';
    sided: 'single' | 'double';
    copies: number;
    pageCount?: number;
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

