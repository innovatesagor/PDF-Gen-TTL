
import React, { useState, useEffect, useRef } from "react";
import { FileText, Plus, Trash2, Download, Eye } from "lucide-react";
import { ReportHeader, LineItem } from "./types";
import { formatCurrency } from "./utils";
import { calculateTotals, generateReports } from "./services/reportService";

const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/**
 * Advanced Date Input
 * - Clamps Day 01-31
 * - Clamps Month 01-12 and shows as text (Jan, Feb...)
 * - Allows natural year input (no auto-expansion)
 */
const SmartDateInput: React.FC<{
  value: string;
  onChange: (val: string) => void;
  label?: string;
  required?: boolean;
  className?: string;
}> = ({ value, onChange, label, required, className }) => {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const input = e.target.value;
    const isDelete = (e.nativeEvent as any).inputType?.includes('delete');
    
    if (isDelete) {
      onChange(input);
      return;
    }

    // Map month names back to digits to allow editing
    let normalized = input;
    months.forEach((m, idx) => {
      const num = (idx + 1).toString().padStart(2, '0');
      normalized = normalized.replace(new RegExp(m, 'g'), num);
    });

    const clean = normalized.replace(/[^\d]/g, "");
    let dayDigits = clean.slice(0, 2);
    let monthDigits = clean.slice(2, 4);
    let yearDigits = clean.slice(4, 8);

    if (dayDigits.length === 2) {
      const d = parseInt(dayDigits);
      if (d > 31) dayDigits = "31";
      if (d === 0) dayDigits = "01";
    }

    let result = dayDigits;
    if (dayDigits.length === 2) result += "/";

    if (monthDigits.length > 0) {
      if (monthDigits.length === 2) {
        let m = parseInt(monthDigits);
        if (m > 12) m = 12;
        if (m === 0) m = 1;
        result += months[m - 1] + "/";
      } else {
        result += monthDigits;
      }
    }

    if (yearDigits.length > 0) {
      result += yearDigits;
    }

    onChange(result);
  };

  return (
    <div className={`input-field ${className || ""}`}>
      {label && (
        <label className="input-label">
          {label} {required && <span style={{ color: "var(--danger)" }}>*</span>}
        </label>
      )}
      <input
        ref={inputRef}
        type="text"
        placeholder="DD/MM/YYYY"
        value={value || ""}
        onChange={handleChange}
        maxLength={11}
        data-is-date-input="true"
        className="date-input-highlight"
      />
    </div>
  );
};

const App: React.FC = () => {
  const [header, setHeader] = useState<ReportHeader>({
    buyerName: "",
    supplierName: "",
    fileNo: "",
    invoiceNo: "",
    lcNumber: "",
    invoiceDate: "",
    billingDate: "",
  });

  const [items, setItems] = useState<LineItem[]>([
    {
      id: crypto.randomUUID(),
      fabricCode: "",
      itemDescription: "",
      color: "",
      hsCode: "",
      rcvdDate: "",
      challanNo: "",
      piNumber: "",
      unit: "YDS",
      invoiceQty: 0,
      rcvdQty: 0,
      unitPrice: 0,
      appstremeNo: "",
    },
  ]);

  const [previewMode, setPreviewMode] = useState(false);

  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Enter") {
        const target = e.target as HTMLInputElement;
        if (target.getAttribute("data-is-date-input") === "true") {
          if (target.value.length > 0 && target.value.length < 10) {
            e.preventDefault();
            return;
          }
        }
        if (target.tagName === "INPUT" || target.tagName === "SELECT") {
          e.preventDefault();
          const formElements = Array.from(
            document.querySelectorAll('input:not([type="hidden"]), select, button:not([disabled])')
          ) as HTMLElement[];
          const index = formElements.indexOf(target);
          if (index > -1 && index < formElements.length - 1) {
            formElements[index + 1].focus();
          }
        }
      }
    };
    window.addEventListener("keydown", handleGlobalKeyDown);
    return () => window.removeEventListener("keydown", handleGlobalKeyDown);
  }, []);

  useEffect(() => {
    const today = new Date();
    const d = String(today.getDate()).padStart(2, '0');
    const m = months[today.getMonth()];
    const y = today.getFullYear();
    setHeader((prev) => ({ ...prev, billingDate: `${d}/${m}/${y}` }));
  }, []);

  const handleHeaderChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setHeader((prev) => ({ ...prev, [name]: value }));
  };

  const handleItemChange = (id: string, field: string, value: any) => {
    setItems((prev) => prev.map((item) => (item.id === id ? { ...item, [field]: value } : item)));
  };

  const addNewRow = () => {
    setItems((prev) => [
      ...prev,
      {
        id: crypto.randomUUID(),
        fabricCode: "",
        itemDescription: "",
        color: "",
        hsCode: "",
        rcvdDate: "",
        challanNo: "",
        piNumber: "",
        unit: "YDS",
        invoiceQty: 0,
        rcvdQty: 0,
        unitPrice: 0,
        appstremeNo: "",
      },
    ]);
  };

  const removeRow = (id: string) => {
    if (items.length > 1) setItems((prev) => prev.filter((item) => item.id !== id));
  };

  const handleGenerate = async () => {
    if (!header.buyerName.trim()) return alert("⚠️ Please enter a Buyer Name.");
    if (!header.billingDate || header.billingDate.length < 8) return alert("⚠️ Please enter a valid billing date.");
    try {
      await generateReports(header, items);
    } catch (error) {
      alert("❌ Error generating reports.");
    }
  };

  const totals = calculateTotals(items);
  const hasQtyMismatch = Math.abs(totals.totalInvoiceQty - totals.totalRcvdQty) > 0.001;

  return (
    <div className="app-container">
      <header className="app-header">
        <div className="header-top">
          <div className="header-title">
            <div className="header-logo">📊</div>
            <span>Bill Of Exchange Report Generator</span>
          </div>
        </div>
      </header>

      <div className="content-wrapper">
        <div className="main-panel">
          <div className="form-section">
            <div className="bill-info-header">
              <FileText size={18} />
              <span>Bill Information</span>
            </div>

            <div className="bill-info-container">
              <div className="bill-info-left">
                <div className="row-30-70">
                  <InputField label="Buyer Name" name="buyerName" value={header.buyerName} onChange={handleHeaderChange} required bold />
                  <InputField label="Supplier Name" name="supplierName" value={header.supplierName} onChange={handleHeaderChange} />
                </div>
                <div className="row-30-70">
                  <InputField label="File No" name="fileNo" value={header.fileNo} onChange={handleHeaderChange} />
                  <InputField label="Invoice No" name="invoiceNo" value={header.invoiceNo} onChange={handleHeaderChange} />
                </div>
              </div>

              <div className="bill-info-right">
                <InputField label="L/C Number" name="lcNumber" value={header.lcNumber} onChange={handleHeaderChange} className="lc-highlight" />
                <div className="row-50-50">
                  <SmartDateInput label="Invoice Date" value={header.invoiceDate} onChange={(val) => setHeader(p => ({ ...p, invoiceDate: val }))} />
                  <SmartDateInput label="Billing Date" value={header.billingDate} onChange={(val) => setHeader(p => ({ ...p, billingDate: val }))} required />
                </div>
              </div>
            </div>
          </div>

          <div className="table-section">
            <div className="table-header">
              <div className="table-title">Line Items ({items.length})</div>
              <div className="table-controls">
                <button className="btn btn-secondary btn-sm" onClick={() => setPreviewMode(!previewMode)}><Eye size={16} />{previewMode ? "Edit" : "Preview"}</button>
                <button className="btn btn-primary btn-sm" onClick={addNewRow}><Plus size={16} />Add Row</button>
              </div>
            </div>

            <div className="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th style={{ minWidth: "220px" }}>Code & Description</th>
                    <th style={{ minWidth: "120px" }}>Color & HS</th>
                    <th style={{ minWidth: "140px" }}>Rcvd Date</th>
                    <th style={{ minWidth: "160px" }}>Challan & PI</th>
                    <th style={{ minWidth: "60px" }}>Unit</th>
                    <th style={{ minWidth: "120px" }}>Quantities</th>
                    <th style={{ minWidth: "100px" }}>Price</th>
                    <th style={{ minWidth: "100px" }}>Total</th>
                    <th style={{ minWidth: "100px" }}>Appstreme</th>
                    <th style={{ minWidth: "40px" }}></th>
                  </tr>
                </thead>
                <tbody>
                  {items.map((item) => (
                    <tr key={item.id}>
                      <td>
                        {previewMode ? <div>{item.fabricCode}<br/>{item.itemDescription}</div> : (
                          <div style={{display:'flex', flexDirection:'column', gap:'4px'}}>
                            <input type="text" value={item.fabricCode} onChange={e => handleItemChange(item.id, 'fabricCode', e.target.value)} placeholder="Code" />
                            <input type="text" value={item.itemDescription} onChange={e => handleItemChange(item.id, 'itemDescription', e.target.value)} placeholder="Desc" />
                          </div>
                        )}
                      </td>
                      <td>
                        {previewMode ? <div>{item.color}<br/>{item.hsCode}</div> : (
                          <div style={{display:'flex', flexDirection:'column', gap:'4px'}}>
                            <input type="text" value={item.color} onChange={e => handleItemChange(item.id, 'color', e.target.value)} placeholder="Color" />
                            <input type="text" value={item.hsCode} onChange={e => handleItemChange(item.id, 'hsCode', e.target.value)} placeholder="HS" />
                          </div>
                        )}
                      </td>
                      <td>
                        {previewMode ? <span>{item.rcvdDate}</span> : <SmartDateInput value={item.rcvdDate} onChange={val => handleItemChange(item.id, 'rcvdDate', val)} />}
                      </td>
                      <td>
                        {previewMode ? <div>{item.challanNo}<br/>{item.piNumber}</div> : (
                          <div style={{display:'flex', flexDirection:'column', gap:'4px'}}>
                            <input type="text" value={item.challanNo} onChange={e => handleItemChange(item.id, 'challanNo', e.target.value)} placeholder="Challan" />
                            <input type="text" value={item.piNumber} onChange={e => handleItemChange(item.id, 'piNumber', e.target.value)} placeholder="PI" />
                          </div>
                        )}
                      </td>
                      <td>
                        {previewMode ? <span>{item.unit}</span> : (
                          <select value={item.unit} onChange={e => handleItemChange(item.id, 'unit', e.target.value)}>
                            <option>YDS</option><option>PCS</option><option>KG</option><option>MTR</option><option>BOX</option>
                          </select>
                        )}
                      </td>
                      <td>
                        {previewMode ? <div>I: {item.invoiceQty}<br/>R: {item.rcvdQty}</div> : (
                          <div style={{display:'flex', flexDirection:'column', gap:'4px'}}>
                            <input 
                              type="number" 
                              value={item.invoiceQty === 0 ? "" : item.invoiceQty} 
                              onChange={e => handleItemChange(item.id, 'invoiceQty', parseFloat(e.target.value)||0)} 
                              placeholder="Inv" 
                            />
                            <input 
                              type="number" 
                              value={item.rcvdQty === 0 ? "" : item.rcvdQty} 
                              onChange={e => handleItemChange(item.id, 'rcvdQty', parseFloat(e.target.value)||0)} 
                              placeholder="Rec" 
                            />
                          </div>
                        )}
                      </td>
                      <td>
                        {previewMode ? <span>${item.unitPrice}</span> : (
                          <input 
                            type="number" 
                            value={item.unitPrice === 0 ? "" : item.unitPrice} 
                            onChange={e => handleItemChange(item.id, 'unitPrice', parseFloat(e.target.value)||0)} 
                            step="0.01" 
                            placeholder="0.01"
                          />
                        )}
                      </td>
                      <td style={{textAlign:'right'}}><strong>${(item.invoiceQty * item.unitPrice).toFixed(2)}</strong></td>
                      <td>{previewMode ? <span>{item.appstremeNo ? item.appstremeNo.toString().padStart(8, '0') : ""}</span> : <input type="text" value={item.appstremeNo} onChange={e => handleItemChange(item.id, 'appstremeNo', e.target.value)} placeholder="No" />}</td>
                      <td>{!previewMode && <button className="btn btn-danger btn-sm" onClick={() => removeRow(item.id)} disabled={items.length===1}><Trash2 size={14}/></button>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

      <footer className="app-footer">
        <div className="footer-summary">
          <div className="footer-summary-item">
            <div className="footer-summary-label">Buyer Name</div>
            <div className="footer-summary-value">{header.buyerName || "Not selected"}</div>
          </div>
          <div className={`footer-summary-item ${hasQtyMismatch ? "qty-mismatch" : ""}`}>
            <div className="footer-summary-label">Invoice Qty</div>
            <div className="footer-summary-value">{totals.totalInvoiceQty.toFixed(2)}</div>
          </div>
          <div className={`footer-summary-item ${hasQtyMismatch ? "qty-mismatch" : ""}`}>
            <div className="footer-summary-label">Rcvd Qty</div>
            <div className="footer-summary-value">{totals.totalRcvdQty.toFixed(2)}</div>
          </div>
          <div className="footer-summary-item">
            <div className="footer-summary-label">Total Value</div>
            <div className="footer-summary-value">{formatCurrency(totals.totalValue)}</div>
          </div>
        </div>
        <div className="footer-actions">
          <button className="btn btn-success" onClick={handleGenerate}>
            <Download size={18} />
            Generate Report (PDF & Excel)
          </button>
        </div>
      </footer>
    </div>
  );
};

const InputField: React.FC<{label:string, name:string, value:any, onChange:any, required?:boolean, bold?:boolean, className?:string}> = ({label, name, value, onChange, required, bold, className}) => (
  <div className={`input-field ${className || ""}`}>
    <label className="input-label">{label} {required && <span style={{color:'var(--danger)'}}>*</span>}</label>
    <input type="text" name={name} value={value} onChange={onChange} className={bold ? "input-bold" : ""} />
  </div>
);

export default App;
