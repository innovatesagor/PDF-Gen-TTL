import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { ReportHeader, LineItem, Totals } from '../types';
import { getFilenameDate } from '../utils';
import { registerPDFFonts, getActiveFontName } from '../fonts';

// Helper to format date for report (DD-Mon-YYYY)
const formatDateForReport = (dateStr: string): string => {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return dateStr;
  
  return new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  }).format(date).replace(/ /g, '-');
};

// Helper to calculate totals
export const calculateTotals = (items: LineItem[]): Totals => {
  return items.reduce(
    (acc, item) => ({
      totalInvoiceQty: acc.totalInvoiceQty + (Number(item.invoiceQty) || 0),
      totalRcvdQty: acc.totalRcvdQty + (Number(item.rcvdQty) || 0),
      totalValue: acc.totalValue + ((Number(item.invoiceQty) || 0) * (Number(item.unitPrice) || 0)),
    }),
    { totalInvoiceQty: 0, totalRcvdQty: 0, totalValue: 0 }
  );
};

export const generateReports = async (header: ReportHeader, items: LineItem[]) => {
  const totals = calculateTotals(items);
  const totalValueStr = Math.round(totals.totalValue); 
  const filenameDate = getFilenameDate(header.billingDate);
  const baseFilename = `Bill of Buyer ${header.buyerName} $${totalValueStr} DATE-${filenameDate}`;

  await generatePDF(header, items, totals, baseFilename);
  await generateExcel(header, items, totals, baseFilename);
};

const generatePDF = async (header: ReportHeader, items: LineItem[], totals: Totals, filename: string) => {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  await registerPDFFonts(doc);
  const fontName = getActiveFontName();

  // --- Layout Constants ---
  const marginX = 10;
  const headerY = 10;
  const titleY = 20;
  const infoBlockY = 26;
  const lineHeight = 6;
  
  // Signature Area Config
  const signatureHeight = 10; 
  const bottomMargin = 7;
  const signatureGap = 20; // Forced gap between table and signatures
  const signatureBlockY = pageHeight - signatureHeight - bottomMargin;
  
  const tableStartY = infoBlockY + (lineHeight * 5) + 4;
  
  // Calculate available vertical space for the table to keep it on one page
  const availableTableHeight = signatureBlockY - tableStartY - signatureGap;

  // --- Dynamic Scaling Logic ---
  const rowCount = items.length + 1; // Items + Footer
  let fontSize = 8;
  let cellPadding = 1.2;
  // Dynamically calculate row height to fill available space or maintain minimum
  let rowHeight = Math.max(6, Math.min(12, availableTableHeight / rowCount));

  if (rowCount > 20) {
    fontSize = 7;
    cellPadding = 0.8;
  }

  // --- Draw Headers ---
  doc.setFont(fontName, 'bold');
  doc.setFontSize(20);
  doc.text("Tusuka Trousers Ltd.", pageWidth / 2, headerY, { align: 'center' });
  
  doc.setFontSize(10);
  doc.setFont(fontName, 'normal');
  doc.text("Neelngar, Konabari, Gazipur", pageWidth / 2, headerY + 6, { align: 'center' });

  doc.setFontSize(15);
  doc.setFont(fontName, 'bold');
  doc.text("Inventory Report", pageWidth / 2, titleY + 2, { align: 'center' });

  // --- Info Block ---
  const drawLabelVal = (label: string, val: string, x: number, y: number) => {
    doc.setFont(fontName, 'bold'); doc.text(label, x, y);
    doc.setFont(fontName, 'normal'); doc.text(val || '', x + 32, y);
  };

  doc.setFontSize(9);
  drawLabelVal("Buyer Name :", header.buyerName, marginX, infoBlockY);
  drawLabelVal("Supplier Name:", header.supplierName, marginX, infoBlockY + lineHeight);
  drawLabelVal("File No :", header.fileNo, marginX, infoBlockY + lineHeight * 2);
  drawLabelVal("Invoice No :", header.invoiceNo, marginX, infoBlockY + lineHeight * 3);
  drawLabelVal("L/C Number :", header.lcNumber, marginX, infoBlockY + lineHeight * 4);

  const rightX = pageWidth - 80;
  drawLabelVal("Invoice Date:", formatDateForReport(header.invoiceDate), rightX, infoBlockY);
  drawLabelVal("Billing Date:", formatDateForReport(header.billingDate), rightX, infoBlockY + lineHeight);

  // --- Table Content ---
  const tableColumn = [
    "Fabric Code", "Item Description", "Rcvd Date", "Challan No", 
    "Pi Number", "Unit", "Invoice Qty", "Rcvd Qty", "Unit Price ($)", "Total Value", "Appstreme No"
  ];

  const tableRows = items.map(item => [
    item.fabricCode,
    item.itemDescription + (item.color ? `\n(Color: ${item.color})` : ""),
    formatDateForReport(item.rcvdDate),
    item.challanNo,
    item.piNumber,
    item.unit,
    Number(item.invoiceQty).toLocaleString(undefined, { minimumFractionDigits: 2 }),
    Number(item.rcvdQty).toLocaleString(undefined, { minimumFractionDigits: 2 }),
    Number(item.unitPrice).toFixed(2),
    (Number(item.invoiceQty) * Number(item.unitPrice)).toLocaleString(undefined, { minimumFractionDigits: 2 }),
    item.appstremeNo
  ]);

  tableRows.push([
    "", "", "", "", "", "TOTAL",
    totals.totalInvoiceQty.toLocaleString(undefined, { minimumFractionDigits: 2 }),
    totals.totalRcvdQty.toLocaleString(undefined, { minimumFractionDigits: 2 }),
    "",
    totals.totalValue.toLocaleString(undefined, { minimumFractionDigits: 2 }),
    ""
  ]);

autoTable(doc, {
    startY: tableStartY,
    head: [tableColumn],
    body: tableRows,
    theme: 'grid',
    // 1. Ensure table respects the margins
    tableWidth: 'auto', 
    margin: { left: marginX, right: marginX }, 
    styles: {
      fontSize: fontSize,
      textColor: [0, 0, 0],
      cellPadding: cellPadding,
      halign: 'center',
      valign: 'middle',
      minCellHeight: rowHeight,
      lineColor: [0, 0, 0],
      lineWidth: 0.1,
      overflow: 'linebreak'
    },
    headStyles: { 
      fillColor: [245, 245, 245], 
      textColor: [0, 0, 0], 
      fontStyle: 'bold' 
    },
    // 2. KEY FIX: Use 'auto' for variable width columns
    columnStyles: {
      0: { cellWidth: 28 },               // Fabric Code (Slightly wider)
      1: { cellWidth: 'auto', halign: 'left' }, // Item Description -> 'auto' allows it to stretch to fill the page
      2: { cellWidth: 24 },               // Rcvd Date
      // 3, 4, 5 (Challan, Pi, Unit) left as default/auto or define if strictly needed
      6: { cellWidth: 22 },               // Invoice Qty
      7: { cellWidth: 22 },               // Rcvd Qty
      8: { cellWidth: 20 },               // Unit Price
      9: { cellWidth: 25 },               // Total Value
      10: { cellWidth: 25 },              // Appstreme No
    },
    didParseCell: (data) => {
      if (data.row.index === tableRows.length - 1) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.textColor = [0, 0, 0];
      }
    }
  });

  // --- Signatures ---
  doc.setLineWidth(0.3);
  doc.setFontSize(9);
  // Left
  doc.line(marginX + 5, signatureBlockY, marginX + 55, signatureBlockY);
  doc.text("Prepared By", marginX + 15, signatureBlockY + 5);
  // Right
  doc.line(pageWidth - marginX - 55, signatureBlockY, pageWidth - marginX - 5, signatureBlockY);
  doc.text("Store In-Charge", pageWidth - marginX - 48, signatureBlockY + 5);

  doc.save(`${filename}.pdf`);
};


const generateExcel = async (header: ReportHeader, items: LineItem[], totals: Totals, filename: string) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Inventory Report');

  // Prepare data
  const data: any[] = [
    ["Tusuka Trousers Ltd."],
    ["Neelngar, Konabari, Gazipur"],
    ["Inventory Report"],
    [],
    ["Buyer Name :", header.buyerName, "", "", "", "", "", "", "", "Invoice Date :", formatDateForReport(header.invoiceDate)],
    ["Supplier Name:", header.supplierName, "", "", "", "", "", "", "", "Billing Date :", formatDateForReport(header.billingDate)],
    ["File No :", header.fileNo],
    ["Invoice No :", header.invoiceNo],
    ["L/C Number :", header.lcNumber],
    [],
    [
      "Fabric Code", "Item Description", "Rcvd Date", "Challan No", "Pi Number", 
      "Unit", "Invoice Qty", "Rcvd Qty", "Unit Price $", "Total Value", "Appstreme No.\n(Receipt no)"
    ]
  ];

  // Add data rows
  items.forEach(item => {
    const invoiceQty = Number(item.invoiceQty) || 0;
    const rcvdQty = Number(item.rcvdQty) || 0;
    const unitPrice = Number(item.unitPrice) || 0;
    const totalVal = invoiceQty * unitPrice;

    let description = item.itemDescription || '';
    const details = [];
    if (item.color && item.color.trim()) details.push(`Color: ${item.color}`);
    if (item.hsCode && item.hsCode.trim()) details.push(`H.S Code: ${item.hsCode}`);
    if (details.length > 0) description += `\n${details.join(', ')}`;

    data.push([
      item.fabricCode,
      description,
      formatDateForReport(item.rcvdDate),
      item.challanNo,
      item.piNumber,
      item.unit,
      invoiceQty,
      rcvdQty,
      unitPrice,
      totalVal,
      item.appstremeNo
    ]);
  });

  // Add Total row
  const totalRowIndex = data.length;
  data.push([
    "", "", "", "", "Total:", "YDS",
    totals.totalInvoiceQty,
    totals.totalRcvdQty,
    "",
    totals.totalValue,
    ""
  ]);

  // Add more empty rows for a larger gap (approx. 3 rows height)
  data.push([], [], [], [], [], [], []);

  // Add signature row
  data.push(["Prepared By", "", "", "", "", "", "", "", "", "Store In-Charge", ""]);

  // Add rows to worksheet
  data.forEach((rowData) => {
    worksheet.addRow(rowData);
  });

  // Set column widths
  worksheet.columns = [
    { width: 14 }, // Fabric Code
    { width: 35 }, // Item Description
    { width: 11 }, // Rcvd Date
    { width: 12 }, // Challan No
    { width: 12 }, // Pi Number
    { width: 8 },  // Unit
    { width: 12 }, // Invoice Qty
    { width: 12 }, // Rcvd Qty
    { width: 12 }, // Unit Price
    { width: 12 }, // Total Value
    { width: 14 }  // Appstreme No
  ];

  // Define border style
  const thinBorder = {
    top: { style: 'thin' as const },
    bottom: { style: 'thin' as const },
    left: { style: 'thin' as const },
    right: { style: 'thin' as const }
  };

  const topBorderOnly = {
    top: { style: 'thin' as const }
  };

  const noBorder = {
    top: undefined,
    bottom: undefined,
    left: undefined,
    right: undefined
  };

  // Style all rows
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      // ROW 1: Company Name - Large Bold Centered (NO BORDER)
      if (rowNumber === 1) {
        cell.font = { bold: true, size: 20 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
      // ROW 2: Address - Centered (NO BORDER)
      else if (rowNumber === 2) {
        cell.font = { size: 11 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
      // ROW 3: Report Title - Bold Centered (NO BORDER)
      else if (rowNumber === 3) {
        cell.font = { bold: true, size: 14 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
      // ROW 4: Empty - skip
      else if (rowNumber === 4) {
        // Skip
      }
      // ROWS 5-9: Bill Information - Bold labels + Values (NO BORDER)
      else if (rowNumber >= 5 && rowNumber <= 9) {
        if (colNumber === 1 || colNumber === 11) {
          // Bold labels
          cell.font = { bold: true, size: 10 };
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          // Values
          cell.font = { size: 10 };
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        }
      }
      // ROW 10: Empty - skip
      else if (rowNumber === 10) {
        // Skip
      }
      // ROW 11: Table Headers - Bold with borders
      else if (rowNumber === 11) {
        cell.font = { bold: true, size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = thinBorder;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      }
      // DATA ROWS: with borders
      else if (rowNumber > 11 && rowNumber <= 11 + items.length) {
        cell.font = { size: 10 };
        cell.border = thinBorder;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        // Left align description and appstreme no
        if (colNumber === 2 || colNumber === 11) {
          cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        }

        // Right align and format numbers
        if (colNumber >= 7 && colNumber <= 10) {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          if (colNumber === 9 || colNumber === 10) {
            cell.numFmt = '0.00';
          }
        }
      }
      // TOTAL ROW: Bold with borders
      else if (rowNumber === totalRowIndex + 1) {
        cell.font = { bold: true, size: 10 };
        cell.border = thinBorder;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };

        if (colNumber === 5) {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        }

        if (colNumber >= 7 && colNumber <= 10) {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          if (colNumber === 9 || colNumber === 10) {
            cell.numFmt = '0.00';
          }
        }
      }
      // SIGNATURE ROW: Top border only on specific cells
      else if (rowNumber === data.length) {
        cell.font = { bold: true, size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'top' };
        
        // Add top border only to cells with content (Prepared By - col 1, and Store In-Charge - col 10)
        if (colNumber === 1 || colNumber === 10) {
          cell.border = topBorderOnly;
        }
        // No border for empty cells
      }
    });
  });

  // Set row heights
  const sheet = worksheet;
  // 1. First, handle specific heights for headers
if (sheet.getRow(1)) sheet.getRow(1).height = 24;
if (sheet.getRow(2)) sheet.getRow(2).height = 16;
if (sheet.getRow(3)) sheet.getRow(3).height = 20;
if (sheet.getRow(4)) sheet.getRow(4).height = 5;

// 2. Handle Info Block
for (let i = 5; i <= 9; i++) {
  if (sheet.getRow(i)) sheet.getRow(i).height = 16;
}
if (sheet.getRow(10)) sheet.getRow(10).height = 5;

// 3. Handle Table Header
if (sheet.getRow(11)) sheet.getRow(11).height = 18;

// 4. Handle EVERYTHING from Table Data down to the Signature
// We use worksheet.rowCount to ensure all rows (including spacers) get height
for (let i = 12; i <= worksheet.rowCount; i++) {
  const row = sheet.getRow(i);
  if (row) {
    // If it's the signature row (the very last one), give it a bit more space
    if (i === worksheet.rowCount) {
      row.height = 20;
    } else {
      row.height = 16; // Standard height for data and empty spacer rows
    }
  }
}

  // Merge cells for headers
  worksheet.mergeCells('A1:K1');
  worksheet.mergeCells('A2:K2');
  worksheet.mergeCells('A3:K3');

  // Generate Excel buffer and trigger download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  });
  
  // Create download link
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${filename}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};
