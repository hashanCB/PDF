import { ArrowDownTrayIcon, ArrowUpIcon, ArrowDownIcon } from '@heroicons/react/24/solid';
import { AdvisorData } from '../types';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { useState } from 'react';

type SortField = 'annualizedPremium' | 'noOfPolicies';
type SortOrder = 'asc' | 'desc';

interface ExcelTableProps {
  data: AdvisorData[];
}

export default function ExcelTable({ data }: ExcelTableProps) {
  const [sortField, setSortField] = useState<SortField | null>(null);
  const [sortOrder, setSortOrder] = useState<SortOrder>('asc');
  const [headerText, setHeaderText] = useState('ACCEPTED AS @ ANBP DN ZONE 2025');
  const [heightMultiplier, setHeightMultiplier] = useState('1.26');
  const [editableData, setEditableData] = useState(data);

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('asc');
    }
  };

  const handleNameEdit = (advisorCode: string, newName: string) => {
    const newData = editableData.map(row => 
      row.advisorCode === advisorCode 
        ? { ...row, advisorName: newName }
        : row
    );
    setEditableData(newData);
  };

  const sortedData = [...editableData].sort((a, b) => {
    if (!sortField) return 0;

    const multiplier = sortOrder === 'asc' ? 1 : -1;
    return (a[sortField] - b[sortField]) * multiplier;
  });

  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(sortedData.map((row, index) => ({
      '#': index + 1,
      'Advisor Code': row.advisorCode,
      'Advisor Name': row.advisorName,
      'Status': row.advisorStatus,
      'No of Policies': row.noOfPolicies,
      'Annualized New Business Premium (RS)': row.annualizedPremium
    })));
    
    // Set column widths (approximate conversion from mm to Excel units)
    const columnWidths = [
      { wch: 4 },  // #
      { wch: 12 }, // Advisor Code
      { wch: 35 }, // Advisor Name
      { wch: 10 }, // Status
      { wch: 10 }, // No of Policies
      { wch: 20 }, // Annualized Premium
    ];
    worksheet['!cols'] = columnWidths;

    // Add styles for header row
    const headerStyle = {
      fill: { fgColor: { rgb: "662D91" } }, // Purple background
      font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 }, // White, bold text
      alignment: { horizontal: "center", vertical: "center", wrapText: true }
    };

    // Add styles for data rows
    const redRowStyle = {
      fill: { fgColor: { rgb: "FF0000" } },
      font: { color: { rgb: "FFFFFF" }, sz: 9 }
    };
    const blueRowStyle = {
      fill: { fgColor: { rgb: "ADD8E6" } },
      font: { color: { rgb: "000000" }, sz: 9 }
    };
    const greenRowStyle = {
      fill: { fgColor: { rgb: "90EE90" } },
      font: { color: { rgb: "000000" }, sz: 9 }
    };

    // Get the range of cells in the worksheet
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Apply styles to cells
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        if (!worksheet[cellRef]) continue;

        // Initialize style object if it doesn't exist
        if (!worksheet[cellRef].s) worksheet[cellRef].s = {};

        // Apply header styles to first row
        if (R === 0) {
          worksheet[cellRef].s = headerStyle;
        } else {
          // Apply row styles based on index
          const rowIndex = R - 1; // Subtract 1 because R starts from 0 and includes header
          if (rowIndex < 3) {
            worksheet[cellRef].s = redRowStyle;
          } else if (rowIndex < 10) {
            worksheet[cellRef].s = blueRowStyle;
          } else {
            worksheet[cellRef].s = greenRowStyle;
          }

          // Apply specific column alignments
          if (C === 0 || C === 1 || C === 3 || C === 4) { // #, Advisor Code, Status, No of Policies
            worksheet[cellRef].s.alignment = { horizontal: "center", vertical: "center" };
          } else if (C === 5) { // Annualized Premium
            worksheet[cellRef].s.alignment = { horizontal: "right", vertical: "center" };
          }
        }
      }
    }

    // Create workbook and add the worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Active Advisors');
    
    // Add sorting information to the filename
    const sortInfo = sortField ? `_sorted_by_${sortField}_${sortOrder}` : '';
    XLSX.writeFile(workbook, `active_advisors${sortInfo}.xlsx`);
  };

  const exportToPDF = () => {
    // Calculate exact dimensions needed
    const titleHeight = 12; // Height for title section
    const rowHeight = 4.5; // Height per data row
    const headerRowHeight = 8; // Height for table header row
    const totalDataRows = sortedData.length;
    const spaceBetweenHeaderAndTable = 5; // Space between header and table
    
    // Calculate content height
    const tableHeight = (rowHeight * totalDataRows) + headerRowHeight;
    const totalContentHeight = titleHeight + tableHeight + spaceBetweenHeaderAndTable;
    
    // Add margin to total height using custom multiplier
    const pageHeight = Math.ceil(totalContentHeight * parseFloat(heightMultiplier));

    // Calculate total width based on column widths plus margins
    const columnWidths = {
      0: 7, // #
      1: 18, // Advisor Code
      2: 80, // Advisor Name
      3: 20, // Status
      4: 20, // No of Policies
      5: 40, // Annualized Premium
    };
    const totalColumnsWidth = Object.values(columnWidths).reduce((sum, width) => sum + width, 0);
    const margins = 5; // Total margins (left + right)
    const totalWidth = totalColumnsWidth + margins; // Total PDF width

    const doc = new jsPDF({
      format: [totalWidth, pageHeight],
      unit: 'mm',
      orientation: 'portrait'
    });
    
    // Add light purple background for title
    doc.setFillColor(246, 242, 255);
    doc.rect(0, 0, totalWidth, titleHeight, 'F');

    // Add purple border at bottom of header
    doc.setDrawColor(102, 45, 145);
    doc.setLineWidth(0.5);
    doc.line(0, titleHeight, totalWidth, titleHeight);

    // Add logo
    const logoWidth = 25;
    const logoHeight = 8;
    doc.addImage('/softlogic-life.png', 'SVG', 2.5, 2, logoWidth, logoHeight);
    
    // Add title with purple bold text
    doc.setTextColor(102, 45, 145);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(headerText, totalWidth / 2, 8, { align: 'center' });
    
    // Reset text color and font for the rest of the document
    doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'normal');
    doc.setLineWidth(0.1);

    // Function to get first 4 words of a name
    const truncateName = (name: string) => {
      const words = name.split(' ');
      return words.slice(0, 4).join(' ');
    };
    
    // Add table using sortedData with custom styling
    autoTable(doc, {
      theme: 'grid',
      head: [['#', 'Advisor Code', 'Advisor Name', 'Status', 'No of\nPolicies', 'Annualized New\nBusiness Premium']],
      body: sortedData.map((row, index) => [
        index + 1,
        row.advisorCode,
        truncateName(row.advisorName),
        row.advisorStatus,
        row.noOfPolicies,
        row.annualizedPremium.toLocaleString()
      ]),
      startY: titleHeight + spaceBetweenHeaderAndTable,
      styles: {
        fontSize: 8,
        cellPadding: { top: 1, right: 1, bottom: 1, left: 1 },
        fontStyle: 'bold',
        lineWidth: 0.1,
        lineColor: [0, 0, 0],
        minCellHeight: rowHeight,
        valign: 'middle',
        overflow: 'linebreak'
      },
      headStyles: {
        fillColor: [102, 45, 145],
        textColor: [255, 255, 255],
        halign: 'center',
        fontStyle: 'bold',
        fontSize: 10,
        minCellHeight: 15,
        cellPadding: { top: 3, right: 2, bottom: 3, left: 2 },
        lineWidth: 0.5,
        lineColor: [0, 0, 0],
        valign: 'middle'
      },
      willDrawCell: function(data) {
        if (data.row.section === 'head') {
          // Set the fill color for header cells
          doc.setFillColor(102, 45, 145);
          // Set darker border for header cells
          doc.setDrawColor(0, 0, 0);
          doc.setLineWidth(0.5);
        }
      },
      didParseCell: function(data) {
        const rowIndex = data.row.index;
        
        // Make header text bold and ensure header background color
        if (data.row.section === 'head') {
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.fontSize = 10;
          data.cell.styles.fillColor = [102, 45, 145];
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.halign = 'center';
          data.cell.styles.valign = 'middle';
          data.cell.styles.lineWidth = 0.5;
          data.cell.styles.lineColor = [0, 0, 0];
        }
        
        // Apply row colors based on index
        if (rowIndex < 3) {
          data.cell.styles.fillColor = [255, 0, 0];
          data.cell.styles.textColor = [255, 255, 255];
        } else if (rowIndex >= 3 && rowIndex < 10) {
          data.cell.styles.fillColor = [173, 216, 230];
          data.cell.styles.textColor = [0, 0, 0];
        } else {
          data.cell.styles.fillColor = [144, 238, 144];
          data.cell.styles.textColor = [0, 0, 0];
        }
      },
      columnStyles: {
        0: { halign: 'center', cellWidth: columnWidths[0] },
        1: { halign: 'center', cellWidth: columnWidths[1] },
        2: { cellWidth: columnWidths[2] },
        3: { halign: 'center', cellWidth: columnWidths[3] },
        4: { halign: 'center', cellWidth: columnWidths[4] },
        5: { halign: 'right', cellWidth: columnWidths[5] },
      },
      margin: { left: 2.5, right: 2.5 },
      tableWidth: totalColumnsWidth,
    });

    // Add sorting information to the filename
    const sortInfo = sortField ? `_sorted_by_${sortField}_${sortOrder}` : '';
    doc.save(`active_advisors${sortInfo}.pdf`);
  };

  const getSortIcon = (field: SortField) => {
    if (sortField !== field) {
      return <ArrowUpIcon className="w-4 h-4 text-gray-400" />;
    }
    return sortOrder === 'asc' ? 
      <ArrowUpIcon className="w-4 h-4 text-purple-600" /> : 
      <ArrowDownIcon className="w-4 h-4 text-purple-600" />;
  };

  return (
    <div className="bg-white rounded-lg shadow-lg overflow-hidden">
      <div className="p-4 border-b flex justify-between items-center bg-gray-50">
        <div className="flex flex-col gap-2 flex-grow mr-4">
          <h2 className="text-lg font-bold text-gray-900">Active Advisors ({data.length})</h2>
          <div className="flex items-center gap-2">
            <label htmlFor="headerText" className="text-sm font-medium text-gray-700">
              PDF Header Text:
            </label>
            <input
              type="text"
              id="headerText"
              value={headerText}
              onChange={(e) => setHeaderText(e.target.value)}
              className="flex-grow px-3 py-1 border text-black border-gray-300 rounded-md text-sm focus:ring-purple-500 focus:border-purple-500"
              placeholder="Enter header text"
            />
          </div>
          <div className="flex items-center gap-2">
            <label htmlFor="heightMultiplier" className="text-sm font-medium text-gray-700">
              Page Height Margin (default: 1.18):
            </label>
            <input
              type="number"
              id="heightMultiplier"
              value={heightMultiplier}
              onChange={(e) => setHeightMultiplier(e.target.value)}
              step="0.01"
              min="1.05"
              max="2"
              className="w-24 px-3 py-1 border text-black border-gray-300 rounded-md text-sm focus:ring-purple-500 focus:border-purple-500"
            />
          </div>
          {sortField && (
            <p className="text-sm font-semibold text-gray-600">
              Sorted by: {sortField} ({sortOrder === 'asc' ? 'ascending' : 'descending'})
            </p>
          )}
        </div>
        <div className="flex gap-2">
          <button
            onClick={exportToExcel}
            className="flex items-center gap-2 px-3 py-1.5 bg-green-600 text-white rounded hover:bg-green-700 transition-colors text-sm font-semibold"
          >
            <ArrowDownTrayIcon className="w-4 h-4" />
            Export XLSX
          </button>
          <button
            onClick={exportToPDF}
            className="flex items-center gap-2 px-3 py-1.5 bg-red-600 text-white rounded hover:bg-red-700 transition-colors text-sm font-semibold"
          >
            <ArrowDownTrayIcon className="w-4 h-4" />
            Export PDF
          </button>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="min-w-full divide-y divide-gray-200">
          <thead className="bg-[#662D91]">
            <tr>
              <th className="px-6 py-3 text-left text-xs font-medium text-white uppercase tracking-wider">
                Advisor Code
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-white uppercase tracking-wider">
                Advisor Name
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-white uppercase tracking-wider">
                Advisor Status
              </th>
              <th 
                className="px-6 py-3 text-left text-xs font-medium text-white uppercase tracking-wider cursor-pointer hover:bg-[#773da2]"
                onClick={() => handleSort('noOfPolicies')}
              >
                <div className="flex items-center gap-1">
                  No of Policies
                  {getSortIcon('noOfPolicies')}
                </div>
              </th>
              <th 
                className="px-6 py-3 text-left text-xs font-medium text-white uppercase tracking-wider cursor-pointer hover:bg-[#773da2]"
                onClick={() => handleSort('annualizedPremium')}
              >
                <div className="flex items-center gap-1">
                  Annualized Premium (RS)
                  {getSortIcon('annualizedPremium')}
                </div>
              </th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {sortedData.map((row) => (
              <tr key={row.advisorCode} className="hover:bg-gray-50">
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {row.advisorCode}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  <input
                    type="text"
                    value={row.advisorName}
                    onChange={(e) => handleNameEdit(row.advisorCode, e.target.value)}
                    className="w-full px-2 py-1 border border-transparent hover:border-gray-300 focus:border-purple-500 rounded focus:outline-none focus:ring-1 focus:ring-purple-500"
                  />
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-green-100 text-green-800">
                    {row.advisorStatus}
                  </span>
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {row.noOfPolicies}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {row.annualizedPremium.toLocaleString()}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
} 