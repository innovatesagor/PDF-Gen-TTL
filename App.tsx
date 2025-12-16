import React, { useState, useEffect } from "react";
import { FileText, Plus, Trash2, Download, Eye } from "lucide-react";
import { ReportHeader, LineItem } from "./types";
import { getFormattedDate, formatCurrency } from "./utils";
import { calculateTotals, generateReports } from "./services/reportService";

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
    const today = new Date();
    setHeader((prev) => ({ ...prev, billingDate: getFormattedDate(today) }));
  }, []);

  const handleHeaderChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setHeader((prev) => ({ ...prev, [name]: value }));
  };

  const handleItemChange = (
    id: string,
    field: string,
    value: string | number
  ) => {
    setItems((prev) =>
      prev.map((item) => {
        if (item.id === id) {
          return { ...item, [field]: value };
        }
        return item;
      })
    );
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
    if (items.length > 1) {
      setItems((prev) => prev.filter((item) => item.id !== id));
    }
  };

  const handleGenerate = async () => {
    if (!header.buyerName.trim()) {
      alert("⚠️ Please enter a Buyer Name.");
      return;
    }
    if (items.length === 0) {
      alert("⚠️ Please add at least one line item.");
      return;
    }
    if (!header.billingDate.trim()) {
      alert("⚠️ Please set a billing date.");
      return;
    }
    try {
      await generateReports(header, items);
    } catch (error) {
      console.error("Error generating reports:", error);
      alert("❌ Error generating reports. Check console for details.");
    }
  };

  const totals = calculateTotals(items);

  return (
    <div className="app-container">
      <header className="app-header">
        <div className="header-top">
          <div className="header-title">
            <div className="header-logo">📊</div>
            <span>Tusuka Inventory Report</span>
          </div>
          <div className="header-subtitle">
            Professional Fabric & Billing Management System
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

  {/* LEFT COLUMN */}
  <div className="bill-info-left">

    {/* Row 1: Buyer 30% | Supplier 70% */}
    <div className="row-30-70">
      <InputField
        label="Buyer Name"
        name="buyerName"
        value={header.buyerName}
        onChange={handleHeaderChange}
        placeholder="Buyer name"
        required
        bold
      />

      <InputField
        label="Supplier Name"
        name="supplierName"
        value={header.supplierName}
        onChange={handleHeaderChange}
        placeholder="Supplier name"
      />
    </div>

    {/* Row 2: File No 30% | Invoice No 70% */}
    <div className="row-30-70">
      <InputField
        label="File No"
        name="fileNo"
        value={header.fileNo}
        onChange={handleHeaderChange}
        placeholder="File No"
      />

      <InputField
        label="Invoice No"
        name="invoiceNo"
        value={header.invoiceNo}
        onChange={handleHeaderChange}
        placeholder="Invoice No"
      />
    </div>

  </div>

  {/* RIGHT COLUMN */}
  <div className="bill-info-right">

    {/* Row 1: L/C Number 100% */}
    <InputField
  label="L/C Number"
  name="lcNumber"
  value={header.lcNumber}
  onChange={handleHeaderChange}
  placeholder="L/C Number"
  className="lc-highlight"
/>


    {/* Row 2: Invoice Date 50% | Billing Date 50% */}
    <div className="row-50-50">
      <InputField
        label="Invoice Date"
        name="invoiceDate"
        type="date"
        value={header.invoiceDate}
        onChange={handleHeaderChange}
      />

      <InputField
        label="Billing Date"
        name="billingDate"
        type="date"
        value={header.billingDate}
        onChange={handleHeaderChange}
        required
      />
    </div>

  </div>
</div>

          </div>

          <div className="table-section">
            <div className="table-header">
              <div className="table-title">
                <FileText size={18} />
                Line Items ({items.length})
              </div>
              <div className="table-controls">
                <button
                  className="btn btn-secondary btn-sm"
                  onClick={() => setPreviewMode(!previewMode)}
                >
                  <Eye size={16} />
                  {previewMode ? "Edit" : "Preview"}
                </button>
                <button className="btn btn-primary btn-sm" onClick={addNewRow}>
                  <Plus size={16} />
                  Add Row
                </button>
              </div>
            </div>

            {items.length === 0 ? (
              <div className="empty-state">
                <p>No line items added yet. Click "Add Row" to start.</p>
              </div>
            ) : (
              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th style={{ minWidth: "80px" }}>Fabric Code</th>
                      <th style={{ minWidth: "140px" }}>Item Description</th>
                      <th style={{ minWidth: "80px" }}>Color</th>
                      <th style={{ minWidth: "80px" }}>HS Code</th>
                      <th style={{ minWidth: "40px" }}>Rcvd Date</th>
                      <th style={{ minWidth: "90px" }}>Challan No</th>
                      <th style={{ minWidth: "90px" }}>PI Number</th>
                      <th style={{ minWidth: "60px" }}>Unit</th>
                      <th style={{ minWidth: "90px" }}>Invoice Qty</th>
                      <th style={{ minWidth: "90px" }}>Rcvd Qty</th>
                      <th style={{ minWidth: "90px" }}>Unit Price</th>
                      <th style={{ minWidth: "90px" }}>Total Value</th>
                      <th style={{ minWidth: "100px" }}>Appstreme No</th>
                      <th style={{ minWidth: "50px" }}>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((item) => {
                      const itemTotal =
                        (Number(item.invoiceQty) || 0) *
                        (Number(item.unitPrice) || 0);
                      return (
                        <tr key={item.id}>
                          <td>
                            {previewMode ? (
                              <span>{item.fabricCode}</span>
                            ) : (
                              <input
                                type="text"
                                value={item.fabricCode}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "fabricCode",
                                    e.target.value
                                  )
                                }
                                placeholder="Code"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.itemDescription}</span>
                            ) : (
                              <input
                                type="text"
                                value={item.itemDescription}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "itemDescription",
                                    e.target.value
                                  )
                                }
                                placeholder="Description"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.color}</span>
                            ) : (
                              <input
                                type="text"
                                value={item.color}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "color",
                                    e.target.value
                                  )
                                }
                                placeholder="Color"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.hsCode}</span>
                            ) : (
                              <input
                                type="text"
                                value={item.hsCode}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "hsCode",
                                    e.target.value
                                  )
                                }
                                placeholder="HS Code"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.rcvdDate}</span>
                            ) : (
                              <input
                                type="date"
                                value={item.rcvdDate}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "rcvdDate",
                                    e.target.value
                                  )
                                }
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.challanNo}</span>
                            ) : (
                              <input
                                type="text"
                                value={item.challanNo}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "challanNo",
                                    e.target.value
                                  )
                                }
                                placeholder="No"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.piNumber}</span>
                            ) : (
                              <input
                                type="text"
                                value={item.piNumber}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "piNumber",
                                    e.target.value
                                  )
                                }
                                placeholder="No"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.unit}</span>
                            ) : (
                              <select
                                value={item.unit}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "unit",
                                    e.target.value
                                  )
                                }
                              >
                                <option>YDS</option>
                                <option>PCS</option>
                                <option>KG</option>
                                <option>MTR</option>
                                <option>BOX</option>
                              </select>
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>
                                {(Number(item.invoiceQty) || 0).toFixed(2)}
                              </span>
                            ) : (
                              <input
                                type="number"
                                value={item.invoiceQty}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "invoiceQty",
                                    parseFloat(e.target.value) || 0
                                  )
                                }
                                placeholder="0"
                                step="0.01"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>
                                {(Number(item.rcvdQty) || 0).toFixed(2)}
                              </span>
                            ) : (
                              <input
                                type="number"
                                value={item.rcvdQty}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "rcvdQty",
                                    parseFloat(e.target.value) || 0
                                  )
                                }
                                placeholder="0"
                                step="0.01"
                              />
                            )}
                          </td>
                          <td>
                            {previewMode ? (
                              <span>
                                ${(Number(item.unitPrice) || 0).toFixed(2)}
                              </span>
                            ) : (
                              <input
                                type="number"
                                value={item.unitPrice}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "unitPrice",
                                    parseFloat(e.target.value) || 0
                                  )
                                }
                                placeholder="0.00"
                                step="0.01"
                              />
                            )}
                          </td>
                          <td className="text-right">
                            <strong>${itemTotal.toFixed(2)}</strong>
                          </td>
                          <td>
                            {previewMode ? (
                              <span>{item.appstremeNo}</span>
                            ) : (
                              <input
                                type="text"
                                value={item.appstremeNo}
                                onChange={(e) =>
                                  handleItemChange(
                                    item.id,
                                    "appstremeNo",
                                    e.target.value
                                  )
                                }
                                placeholder="No"
                              />
                            )}
                          </td>
                          <td>
                            {!previewMode && (
                              <button
                                className="btn btn-danger btn-sm"
                                onClick={() => removeRow(item.id)}
                                title="Delete row"
                                disabled={items.length === 1}
                              >
                                <Trash2 size={14} />
                              </button>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      </div>

      <footer className="app-footer">
        <div className="footer-summary">
          <div className="footer-summary-item">
            <div className="footer-summary-label">Buyer Name</div>
            <div className="footer-summary-value">
              {header.buyerName || "Not selected"}
            </div>
          </div>

          <div
            className={`footer-summary-item ${
              totals.totalInvoiceQty !== totals.totalRcvdQty
                ? "qty-mismatch"
                : ""
            }`}
          >
            <div className="footer-summary-label">Invoice Qty</div>
            <div className="footer-summary-value">
              {totals.totalInvoiceQty.toFixed(2)}
            </div>
          </div>

          <div
            className={`footer-summary-item ${
              totals.totalInvoiceQty !== totals.totalRcvdQty
                ? "qty-mismatch"
                : ""
            }`}
          >
            <div className="footer-summary-label">Rcvd Qty</div>
            <div className="footer-summary-value">
              {totals.totalRcvdQty.toFixed(2)}
            </div>
          </div>

          <div className="footer-summary-item">
            <div className="footer-summary-label">Total Value</div>
            <div className="footer-summary-value">
              {formatCurrency(totals.totalValue)}
            </div>
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

interface InputFieldProps {
  label: string;
  name: string;
  value: string | number;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  placeholder?: string;
  type?: string;
  required?: boolean;
  bold?: boolean;
  error?: boolean;
}

const InputField: React.FC<InputFieldProps> = ({
  label,
  name,
  value,
  onChange,
  placeholder,
  type = "text",
  required,
  bold = false,
  error = false,
}) => (
  <div className="input-field">
    <label className="input-label">
      {label} {required && <span style={{ color: "var(--danger)" }}>*</span>}
    </label>
    <input
      type={type}
      name={name}
      value={value}
      onChange={onChange}
      placeholder={placeholder}
      className={`${bold ? "input-bold" : ""} ${error ? "input-error" : ""}`}
    />
  </div>
);

export default App;
