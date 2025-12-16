import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { ReportHeader, LineItem, Totals } from '../types';
import { getFilenameDate } from '../utils';
import { registerPDFFonts, getActiveFontName } from '../fonts';

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
  generateExcel(header, items, totals, baseFilename);
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

const generateExcel = (header: ReportHeader, items: LineItem[], totals: Totals, filename: string) => {
  const wb = XLSX.utils.book_new();
  
  // Pre-process rows 
  const excelRows = items.map(item => {
    const invoiceQty = Number(item.invoiceQty) || 0;
    const rcvdQty = Number(item.rcvdQty) || 0;
    const unitPrice = Number(item.unitPrice) || 0;
    const totalRowVal = invoiceQty * unitPrice;

    let description = item.itemDescription || '';
    const details = [];
    if (item.color && item.color.trim()) details.push(`Color: ${item.color}`);
    if (item.hsCode && item.hsCode.trim()) details.push(`H.S Code: ${item.hsCode}`);
    if (details.length > 0) description += `\n${details.join(', ')}`;

    return [
      item.fabricCode, description, item.rcvdDate,
      item.challanNo, item.piNumber, item.unit,
      invoiceQty, rcvdQty, unitPrice, totalRowVal, item.appstremeNo
    ];
  });

  const wsData: any[][] = [
    ["Tusuka Trousers Ltd."],
    ["Neelngar, Konabari, Gazipur"],
    ["Inventory Report"],
    [],
    ["Buyer Name :", header.buyerName, null, null, null, null, null, "Invoice Date :", header.invoiceDate],
    ["Supplier Name:", header.supplierName, null, null, null, null, null, "Billing Date :", header.billingDate],
    ["File No :", header.fileNo],
    ["Invoice No :", header.invoiceNo],
    ["L/C Number :", header.lcNumber],
    [],
    // Row 10: Header
    [
      "Fabric Code", "Item Description", "Rcvd Date", "Challan No", "Pi Number", 
      "Unit", "Invoice Qty", "Rcvd Qty", "Unit Price $", "Total Value", "Appstreme No.\n(Receipt no)"
    ]
  ];

  // Add Item Rows
  excelRows.forEach(row => wsData.push(row as any[]));

  // Add Footer Row 
  wsData.push([
    "", "", "", "", "Total:", "YDS", 
    totals.totalInvoiceQty, totals.totalRcvdQty, "", totals.totalValue, ""
  ]);

  // Add empty rows for spacing before signature
  wsData.push([]); 
  wsData.push([]); 
  wsData.push([]); 
  wsData.push([]); 

  // Add Signature Row
  const sigRowIndex = wsData.length;
  const sigRow = new Array(11).fill("");
  sigRow[1] = "Prepared By";
  sigRow[9] = "Store In-Charge";
  wsData.push(sigRow);

  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // --- Styling Constants ---
  const thinBorder = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };

  const centerAlign = { horizontal: "center", vertical: "center", wrapText: true };
  const leftAlign = { horizontal: "left", vertical: "center", wrapText: true };
  const rightAlign = { horizontal: "right", vertical: "center" };

  const labelStyle = { font: { bold: true }, alignment: { horizontal: "left" } };
  const headerStyle = {
    font: { bold: true, sz: 10 },
    alignment: centerAlign,
    border: thinBorder,
    fill: { fgColor: { rgb: "F0F0F0" } } 
  };

  const dataStyleCenter = { font: { sz: 9 }, alignment: centerAlign, border: thinBorder };
  const dataStyleLeft = { font: { sz: 9 }, alignment: leftAlign, border: thinBorder };
  const dataStyleRightCurrency = { font: { sz: 9 }, alignment: rightAlign, border: thinBorder, numFmt: "0.00" };
  const footerStyle = { font: { bold: true, sz: 10 }, alignment: rightAlign, border: thinBorder, numFmt: "0.00" };
  const signatureTextStyle = {
    font: { bold: true, sz: 9 },
    alignment: { horizontal: "center", vertical: "top" },
    border: { top: { style: "thin", color: { rgb: "000000" } } }
  };

  // --- Apply Styles ---
  ws['A1'].s = { font: { bold: true, sz: 18 }, alignment: { horizontal: "center" } };
  ws['A2'].s = { font: { sz: 12 }, alignment: { horizontal: "center" } };
  ws['A3'].s = { font: { bold: true, sz: 14 }, alignment: { horizontal: "center" } };

  const infoRowIndices = [4, 5, 6, 7, 8];
  infoRowIndices.forEach(r => {
    const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
    const cellRef2 = XLSX.utils.encode_cell({ r, c: 7 }); 
    if(ws[cellRef]) ws[cellRef].s = labelStyle;
    if(ws[cellRef2]) ws[cellRef2].s = labelStyle;
  });

  const headerRowIndex = 10;
  for(let c = 0; c <= 10; c++) {
    const cellRef = XLSX.utils.encode_cell({ r: headerRowIndex, c });
    if(ws[cellRef]) ws[cellRef].s = headerStyle;
  }

  const startDataRow = 11;
  const endDataRow = startDataRow + items.length - 1;

  for (let r = startDataRow; r <= endDataRow; r++) {
    for (let c = 0; c <= 10; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      if (!ws[cellRef]) continue;
      let style: any = dataStyleCenter;
      if (c === 1 || c === 10) style = dataStyleLeft; 
      if (c === 6 || c === 7 || c === 8 || c === 9) style = dataStyleRightCurrency; 
      ws[cellRef].s = style;
    }
  }

  const footerRowIndex = endDataRow + 1;
  for (let c = 0; c <= 10; c++) {
    const cellRef = XLSX.utils.encode_cell({ r: footerRowIndex, c });
    if (!ws[cellRef]) XLSX.utils.sheet_add_aoa(ws, [[""]], { origin: cellRef });
    let style = { ...footerStyle };
    if (c === 4) style.alignment = { horizontal: "right", vertical: "center" }; 
    if (c === 5) style.alignment = { horizontal: "center", vertical: "center" }; 
    ws[cellRef].s = style;
  }

  if(!ws['!merges']) ws['!merges'] = [];
  
  // Signatures
  ws['!merges'].push({ s: { r: sigRowIndex, c: 0 }, e: { r: sigRowIndex, c: 2 } });
  const sigLeftRef = XLSX.utils.encode_cell({ r: sigRowIndex, c: 0 });
  if (!ws[sigLeftRef]) XLSX.utils.sheet_add_aoa(ws, [[""]], { origin: sigLeftRef });
  ws[sigLeftRef].v = "Prepared By"; 
  ws[sigLeftRef].s = signatureTextStyle;

  ws['!merges'].push({ s: { r: sigRowIndex, c: 8 }, e: { r: sigRowIndex, c: 10 } });
  const sigRightRef = XLSX.utils.encode_cell({ r: sigRowIndex, c: 8 });
  if (!ws[sigRightRef]) XLSX.utils.sheet_add_aoa(ws, [[""]], { origin: sigRightRef });
  ws[sigRightRef].v = "Store In-Charge";
  ws[sigRightRef].s = signatureTextStyle;

  // Title Merges
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 10 } }); 
  ws['!merges'].push({ s: { r: 1, c: 0 }, e: { r: 1, c: 10 } }); 
  ws['!merges'].push({ s: { r: 2, c: 0 }, e: { r: 2, c: 10 } }); 

  ws['!cols'] = [
    { wch: 15 }, { wch: 40 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, 
    { wch: 6 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 15 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Inventory Report");
  XLSX.writeFile(wb, `${filename}.xlsx`);
};
