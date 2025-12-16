export interface ReportHeader {
  buyerName: string;
  supplierName: string;
  fileNo: string;
  invoiceNo: string;
  lcNumber: string;
  invoiceDate: string;
  billingDate: string; // Format: 01-Jan-2025
}

export interface LineItem {
  id: string;
  fabricCode: string;
  itemDescription: string;
  color: string;
  hsCode: string;
  rcvdDate: string;
  challanNo: string;
  piNumber: string;
  unit: string; // e.g., YDS
  invoiceQty: number;
  rcvdQty: number;
  unitPrice: number;
  appstremeNo: string;
}

export interface Totals {
  totalInvoiceQty: number;
  totalRcvdQty: number;
  totalValue: number;
}
