'use client';

import { useState } from 'react';
import * as XLSX from 'xlsx';
import ExcelTable from './components/ExcelTable';
import FileUpload from './components/FileUpload';
import { AdvisorData } from './types';

export default function Home() {
  const [tableData, setTableData] = useState<AdvisorData[]>([]);

  const processExcelFile = (file: File) => {
    console.log('Processing file:', file.name);
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) {
          console.error('No data read from file');
          return;
        }

        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert to array of arrays to process the data manually
        const rawData = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1 });
        console.log('Raw data rows:', rawData);

        // Find the header row index (looking for 'Advisor Code')
        let headerRowIndex = -1;
        for (let i = 0; i < rawData.length; i++) {
          const row = rawData[i];
          if (row && row.includes('Advisor Code')) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) {
          console.error('Could not find header row');
          return;
        }

        // Get header columns
        const headers = rawData[headerRowIndex];
        console.log('Headers found:', headers);

        // Process data rows
        const processedData: AdvisorData[] = [];
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
          const row = rawData[i];
          if (!row || row.length === 0) continue;

          // Find column indexes (using exact header names from your Excel)
          const advisorCodeIndex = headers.findIndex(h => h === 'Advisor Code');
          const advisorNameIndex = headers.findIndex(h => h === 'Advisor Name');
          const advisorStatusIndex = headers.findIndex(h => h === 'Advisor Status');
          const policiesIndex = headers.findIndex(h => h === 'No of Policies');
          const premiumIndex = headers.findIndex(h => h === 'Annualized New Business Premium (RS)');

          // Skip empty rows or rows without advisor code
          if (!row[advisorCodeIndex]) continue;

          // Get the status, defaulting to empty string if undefined
          const status = (row[advisorStatusIndex] || '').toString().toLowerCase();
          
          // Only process rows with 'active' status
          if (status === 'active') {
            const advisorData: AdvisorData = {
              advisorCode: row[advisorCodeIndex]?.toString() || '',
              advisorName: row[advisorNameIndex]?.toString() || '',
              advisorStatus: row[advisorStatusIndex]?.toString() || '',
              noOfPolicies: Number(row[policiesIndex]) || 0,
              annualizedPremium: Number(row[premiumIndex]) || 0
            };

            // Only add if we have at least an advisor code
            if (advisorData.advisorCode) {
              processedData.push(advisorData);
            }
          }
        }

        console.log('Processed data:', processedData);
        
        if (processedData.length === 0) {
          console.log('No active advisors found in the data');
        }

        setTableData((prev) => [...prev, ...processedData]);
      } catch (error) {
        console.error('Error processing Excel file:', error);
        alert('Error processing Excel file. Please make sure it has the correct format.');
      }
    };

    reader.onerror = (error) => {
      console.error('Error reading file:', error);
      alert('Error reading file. Please try again.');
    };

    reader.readAsArrayBuffer(file);
  };

  const handleFileUpload = (files: FileList) => {
    console.log('Files received:', files.length);
    Array.from(files).forEach(processExcelFile);
  };

  const clearData = () => {
    console.log('Clearing data');
    setTableData([]);
  };

  return (
    <main className="min-h-screen p-8">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-3xl font-bold text-purple-700 mb-8">
          PDF MAKER - Active  Data
        </h1>
        
        <div className="mb-8">
          <FileUpload onFileUpload={handleFileUpload} onClear={clearData} />
        </div>

        {tableData.length > 0 && (
          <ExcelTable data={tableData} />
        )}

        {tableData.length === 0 && (
          <div className="text-center text-gray-500 mt-8">
            No active advisors found. Please upload an Excel file with advisor data.
          </div>
        )}
      </div>
    </main>
  );
}
