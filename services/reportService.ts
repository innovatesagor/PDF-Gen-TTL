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

  // Register custom fonts
  await registerPDFFonts(doc);
  const fontName = getActiveFontName();

  // --- Layout Constants ---
  const headerY = 12;
  const addressY = 18;
  const titleY = 25;
  
  const infoBlockY = 32;
  const lineHeight = 5;
  const infoBlockHeight = lineHeight * 4; // 4 lines of details
  
  const tableStartY = infoBlockY + infoBlockHeight + 4;
  
  // Signature Block Constants - FIXED AT BOTTOM WITH GAP
  const signatureGapSize = 35; // Large gap before signatures (in mm)
  const signatureHeight = 20; // Height reserved for signature lines/text
  const bottomMargin = 8;
  const signatureBlockY = pageHeight - signatureHeight - bottomMargin;
  
  // Calculate available height for the table (with generous gap before signatures)
  const maxTableHeight = signatureBlockY - tableStartY - signatureGapSize;
  
  // --- Dynamic Layout Logic for Single Page ---
  // Rows = Header(1) + Items(n) + Footer(1)
  const rowCount = items.length + 2; 
  
  // Default styling - OPTIMIZED FOR SINGLE PAGE
  let finalFontSize = 8;
  let finalCellPadding = 1.5;
  let finalMinCellHeight = 6;

  // DYNAMIC ROW HEIGHT LOGIC:
  // - If fewer rows: Make them TALLER for better readability
  // - If more rows: Shrink proportionally while keeping signature gap
  
  if (rowCount <= 5) {
    // Few rows - EXPAND to fill space nicely
    finalMinCellHeight = maxTableHeight / rowCount; // Distribute all available space
    finalFontSize = 10;
    finalCellPadding = 2;
  } else if (rowCount <= 15) {
    // Medium rows - COMFORTABLE sizing
    finalMinCellHeight = Math.max(8, maxTableHeight / rowCount);
    finalFontSize = 9;
    finalCellPadding = 1.75;
  } else if (rowCount <= 25) {
    // Many rows - COMPACT sizing
    finalMinCellHeight = Math.max(6, maxTableHeight / rowCount);
    finalFontSize = 8;
    finalCellPadding = 1.5;
  } else {
    // Very many rows - MINIMAL sizing (keep signature gap)
    finalMinCellHeight = Math.max(5, maxTableHeight / rowCount);
    finalFontSize = 7;
    finalCellPadding = 1;
  }

  // Ensure minimum readability
  finalFontSize = Math.max(5, finalFontSize);
  finalMinCellHeight = Math.max(5, finalMinCellHeight);

  // --- Draw Static Header Content ---
  doc.setFontSize(18);
  doc.setFont(fontName, 'bold');
  doc.text("Tusuka Trousers Ltd.", pageWidth / 2, headerY, { align: 'center' });
  
  doc.setFontSize(8);
  doc.setFont(fontName, 'normal');
  doc.text("Neelngar, Konabari, Gazipur", pageWidth / 2, addressY, { align: 'center' });

  doc.setFontSize(12);
  doc.setFont(fontName, 'bold');
  doc.text("Inventory Report", pageWidth / 2, titleY, { align: 'center' });

  // --- Info Block with Totals on Right ---
  const leftX = 14;
  const rightX = pageWidth - 75;

  doc.setFontSize(8);
  
  // Helper to draw label: value pair
  const drawLabelVal = (label: string, val: string, x: number, y: number) => {
    doc.setFont(fontName, 'bold'); doc.text(label, x, y);
    doc.setFont(fontName, 'normal'); doc.text(val || '', x + 32, y);
  };

  drawLabelVal("Buyer Name :", header.buyerName, leftX, infoBlockY);
  drawLabelVal("Supplier Name:", header.supplierName, leftX, infoBlockY + lineHeight);
  drawLabelVal("File No :", header.fileNo, leftX, infoBlockY + lineHeight * 2);
  drawLabelVal("Invoice No :", header.invoiceNo, leftX, infoBlockY + lineHeight * 3);

  // Right Column - Dates and Totals
  doc.setFont(fontName, 'bold'); doc.text("Invoice Date:", rightX, infoBlockY);
  doc.setFont(fontName, 'normal'); doc.text(header.invoiceDate || '', rightX + 28, infoBlockY);

  doc.setFont(fontName, 'bold'); doc.text("Billing Date:", rightX, infoBlockY + lineHeight);
  doc.setFont(fontName, 'normal'); doc.text(header.billingDate || '', rightX + 28, infoBlockY + lineHeight);

  // L/C Number on left side
  drawLabelVal("L/C Number :", header.lcNumber, leftX, infoBlockY + lineHeight * 4);

  // --- Table Data Preparation ---
  const tableColumn = [
    "Fabric Code", 
    "Item Description", 
    "Rcvd Date", 
    "Challan No", 
    "Pi Number", 
    "Unit", 
    "Invoice Qty", 
    "Rcvd Qty", 
    "Unit Price $", 
    "Total Value", 
    "Appstreme No.\n(Receipt no)"
  ];

  const tableRows = items.map(item => {
    // FIX: Safely cast to numbers
    const invoiceQty = Number(item.invoiceQty) || 0;
    const rcvdQty = Number(item.rcvdQty) || 0;
    const unitPrice = Number(item.unitPrice) || 0;
    const totalRowVal = invoiceQty * unitPrice;

    // Build Description Conditionally
    let description = item.itemDescription || '';
    const details = [];
    if (item.color && item.color.trim()) details.push(`Color: ${item.color}`);
    if (item.hsCode && item.hsCode.trim()) details.push(`H.S Code: ${item.hsCode}`);
    
    if (details.length > 0) {
      description += `\n${details.join(', ')}`;
    }

    return [
      item.fabricCode,
      description,
      item.rcvdDate,
      item.challanNo,
      item.piNumber,
      item.unit,
      invoiceQty,
      rcvdQty,
      unitPrice.toFixed(2),
      totalRowVal.toFixed(2),
      item.appstremeNo
    ];
  });

  // Footer row with totals - PROFESSIONAL STYLING
  const footerRow = [
    "", "", "", "", "", "TOTAL:", 
    totals.totalInvoiceQty.toFixed(2),
    totals.totalRcvdQty.toFixed(2),
    "",
    totals.totalValue.toFixed(2),
    ""
  ];
  tableRows.push(footerRow);

  // --- AutoTable Generation with Professional Styling ---
  autoTable(doc, {
    startY: tableStartY,
    head: [tableColumn],
    body: tableRows,
    theme: 'grid',
    styles: {
      font: 'helvetica',
      fontSize: finalFontSize,
      cellPadding: finalCellPadding,
      overflow: 'linebreak',
      halign: 'center',
      valign: 'middle',
      minCellHeight: finalMinCellHeight,
      lineColor: [0, 0, 0], // Black border
      lineWidth: 0.3
    },
    headStyles: {
      fillColor: [255, 255, 255], // White background
      textColor: [0, 0, 0], // Black text
      fontStyle: 'bold',
      fontSize: finalFontSize,
      halign: 'center',
      valign: 'middle',
      lineColor: [0, 0, 0], // Black border
      lineWidth: 0.4
    },
    bodyStyles: {
      fillColor: [255, 255, 255], // White background
      textColor: [0, 0, 0], // Black text
      lineColor: [0, 0, 0], // Black border
      lineWidth: 0.3,
    },
    columnStyles: {
      1: { cellWidth: 40, halign: 'left' } // Description column
    },
    didParseCell: (data) => {
      // Style the total row - BOLD without background color
      if (data.row.index === tableRows.length - 1) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.fillColor = [255, 255, 255]; // White background
        data.cell.styles.textColor = [0, 0, 0]; // Black text
        data.cell.styles.lineWidth = 0.4;
      }
      // No alternate row colors - all white
      if (data.row.index % 2 === 0 && data.row.index !== tableRows.length - 1) {
        data.cell.styles.fillColor = [255, 255, 255]; // White background
        data.cell.styles.textColor = [0, 0, 0]; // Black text
      }
    },
  });

  // --- Signatures with Large Gap ---
  // Place signature block at fixed position at bottom with significant gap
  
  let sigY = signatureBlockY;
  const lastTableY = (doc as any).lastAutoTable.finalY;

  // If table overlaps with signature area, add a new page
  if (lastTableY > signatureBlockY - 10) {
      doc.addPage();
      sigY = signatureBlockY; 
  }

  // Draw signature lines with gap area above
  doc.setLineWidth(0.3);
  
  // Left Sig - "Prepared By"
  doc.line(20, sigY, 70, sigY);
  doc.setFontSize(8);
  doc.setFont(fontName, 'bold');
  doc.text("Prepared By", 30, sigY + 4);

  // Right Sig - "Store In-Charge"
  doc.line(pageWidth - 70, sigY, pageWidth - 20, sigY);
  doc.text("Store In-Charge", pageWidth - 65, sigY + 4);

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
    ["Buyer Name :", header.buyerName, "", "", "", "", "", "", "", "", "Invoice Date :", formatDateForReport(header.invoiceDate)],
    ["Supplier Name:", header.supplierName, "", "", "", "", "", "", "", "", "Billing Date :", formatDateForReport(header.billingDate)],
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

  // Add spacing for signatures
  data.push([]);
  data.push([]);
  data.push([]);
  data.push([]);

  // Add signature row
  data.push(["Prepared By", "", "", "", "", "", "", "Store In-Charge", "", "", ""]);

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

  // Style all rows
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      // ROW 1: Company Name - Large Bold Centered
      if (rowNumber === 1) {
        cell.font = { bold: true, size: 20 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
      // ROW 2: Address - Centered
      else if (rowNumber === 2) {
        cell.font = { size: 11 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
      // ROW 3: Report Title - Bold Centered
      else if (rowNumber === 3) {
        cell.font = { bold: true, size: 14 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
      // ROW 4: Empty - skip
      else if (rowNumber === 4) {
        // Skip
      }
      // ROWS 5-9: Bill Information - Bold labels + Values
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
      // SIGNATURE ROW: Top border only
      else if (rowNumber === data.length) {
        cell.font = { bold: true, size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'top' };
        cell.border = topBorderOnly;
      }
    });
  });

  // Set row heights
  const sheet = worksheet;
  if (sheet.getRow(1)) sheet.getRow(1).height = 24;
  if (sheet.getRow(2)) sheet.getRow(2).height = 16;
  if (sheet.getRow(3)) sheet.getRow(3).height = 20;
  if (sheet.getRow(4)) sheet.getRow(4).height = 5;
  for (let i = 5; i <= 9; i++) {
    if (sheet.getRow(i)) sheet.getRow(i).height = 16;
  }
  if (sheet.getRow(10)) sheet.getRow(10).height = 5;
  if (sheet.getRow(11)) sheet.getRow(11).height = 18;

  for (let i = 12; i <= 11 + items.length + 1; i++) {
    if (sheet.getRow(i)) sheet.getRow(i).height = 16;
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
