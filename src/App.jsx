import React, { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import './App.css';

function App() {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState({ type: '', message: '' });
  const [isProcessing, setIsProcessing] = useState(false);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile && (selectedFile.name.endsWith('.xlsx') || selectedFile.name.endsWith('.xls'))) {
      setFile(selectedFile);
      setStatus({ type: '', message: '' });
    } else {
      setStatus({ type: 'error', message: 'Please upload a valid Excel file (.xlsx or .xls)' });
    }
  };

  const processExcel = async () => {
    if (!file) return;
    setIsProcessing(true);
    setStatus({ type: 'info', message: 'Processing Excel data...' });

    try {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.getWorksheet(1);

      const data = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header

        // New Simplified Columns:
        // 1: Code, 2: Name, 3: Address, 4: Pincode, 5: Mobile, 6: Size, 7: Date
        const rowData = {
          code: row.getCell(1).text,
          toName: row.getCell(2).text,
          toAddress: row.getCell(3).text,
          toPincode: row.getCell(4).text,
          toMobile: row.getCell(5).text,
          size: row.getCell(6).text,
          date: row.getCell(7).text,
          // Fixed details
          amount: "950",
          amountWords: "NINE HUNDRED FIFTY",
          fromPincode: "637504",
          fromMobile: "8838996592",
          fromId: "1242064413",
        };
        data.push(rowData);
      });

      if (data.length === 0) {
        throw new Error('No data found in Excel sheet.');
      }

      await generatePDF(data);
      setStatus({ type: 'success', message: `Successfully generated labels for ${data.length} entries!` });
    } catch (error) {
      console.error(error);
      setStatus({ type: 'error', message: 'Error processing file: ' + error.message });
    } finally {
      setIsProcessing(false);
    }
  };

  const generatePDF = async (data) => {
    const doc = new jsPDF('p', 'mm', 'a4');
    const slipWidth = 180;
    const slipHeight = 120;
    const marginX = (210 - slipWidth) / 2;
    const startY = 15;
    const gapY = 35; // 3.5cm gap

    for (let i = 0; i < data.length; i++) {
      const slipIndexOnPage = i % 2;
      const yOffset = startY + slipIndexOnPage * (slipHeight + gapY);

      if (slipIndexOnPage === 0 && i !== 0) {
        doc.addPage();
      }

      const item = data[i];

      // Draw Main Box
      doc.setDrawColor(0);
      doc.setLineWidth(0.5);
      doc.rect(marginX, yOffset, slipWidth, slipHeight);

      // Header Section
      doc.line(marginX, yOffset + 10, marginX + slipWidth, yOffset + 10);
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.text('VP PARCEL          VPP:Rs- 950/-', marginX + slipWidth / 2, yOffset + 7, { align: 'center' });

      // Amount in Words
      doc.line(marginX, yOffset + 20, marginX + slipWidth, yOffset + 20);
      doc.setFontSize(12);
      doc.text(`${item.amountWords} RUPEES ONLY`, marginX + slipWidth / 2, yOffset + 17, { align: 'center' });

      // Body Split Line
      doc.line(marginX + slipWidth / 2, yOffset + 20, marginX + slipWidth / 2, yOffset + 105);

      // FROM Section
      doc.setFontSize(15);
      doc.text('FROM:', marginX + 5, yOffset + 27);
      doc.setFont('helvetica', 'normal');
      doc.text('Murugan textile,', marginX + 5, yOffset + 35);
      doc.text('KALPARAPATTY,', marginX + 5, yOffset + 43);
      doc.text(`Pincode-${item.fromPincode},`, marginX + 5, yOffset + 51);
      doc.text(`Ph.no: ${item.fromMobile}`, marginX + 5, yOffset + 59);
      doc.text(`Customer ID - ${item.fromId}`, marginX + 5, yOffset + 67);

      // TO Section
      doc.setFontSize(15);
      doc.setFont('helvetica', 'bold');
      doc.text('TO:', marginX + slipWidth / 2 + 5, yOffset + 27);
      doc.setFont('helvetica', 'normal');

      const fullAddress = `${item.toName}\n${item.toAddress}\nPincode-${item.toPincode}\nPh.no: ${item.toMobile}`;
      const splitAddress = doc.splitTextToSize(fullAddress, (slipWidth / 2) - 10);
      doc.text(splitAddress, marginX + slipWidth / 2 + 5, yOffset + 35);

      // Footer
      doc.line(marginX, yOffset + 105, marginX + slipWidth, yOffset + 105);
      doc.line(marginX + (slipWidth / 3), yOffset + 105, marginX + (slipWidth / 3), yOffset + slipHeight);
      doc.line(marginX + (2 * slipWidth / 3), yOffset + 105, marginX + (2 * slipWidth / 3), yOffset + slipHeight);

      // Footer Values (Labels removed as per request)
      doc.setFont('helvetica', 'bold');
      doc.text(item.size, marginX + 5, yOffset + 115);
      doc.text(item.code, marginX + (slipWidth / 3) + 5, yOffset + 115);
      doc.text(item.date, marginX + (2 * slipWidth / 3) + 5, yOffset + 115);
    }

    const pdfOutput = doc.output('arraybuffer');
    const success = await window.electronAPI.savePDF(pdfOutput);
    if (!success) {
      setStatus({ type: 'info', message: 'PDF Save Cancelled' });
    }
  };

  return (
    <div className="app-container">
      <div className="card">
        <h1>VP Parcel</h1>
        <p className="subtitle">Shipping Label Generator</p>

        <div
          className="drop-zone"
          onClick={() => document.getElementById('file-input').click()}
        >
          <input
            id="file-input"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileChange}
            style={{ display: 'none' }}
          />
          <div className="drop-zone-content">
            <span className="upload-icon">📄</span>
            {file ? (
              <p>Selected: <strong>{file.name}</strong></p>
            ) : (
              <p>Click or drag Excel file here</p>
            )}
            <span className="file-info">Supports .xlsx and .xls formats</span>
          </div>
        </div>

        <button
          className="btn-generate"
          onClick={processExcel}
          disabled={!file || isProcessing}
        >
          {isProcessing ? 'Processing...' : 'Generate Shipping Labels'}
        </button>

        {status.message && (
          <p className={`status-msg ${status.type}`}>
            {status.message}
          </p>
        )}
      </div>
    </div>
  );
}

export default App;
