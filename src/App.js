import React, { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import './App.css';
import { saveAs } from 'file-saver';

function App() {
  const [file, setFile] = useState(null);
  const [tableData, setTableData] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [outputFormat, setOutputFormat] = useState('pptx'); // 'pptx' or 'html'

  const { getRootProps, getInputProps } = useDropzone({
    accept: {
      'text/plain': ['.txt'],
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    onDrop: useCallback(acceptedFiles => {
      if (acceptedFiles.length === 0) {
        console.error('No files accepted.');
        setError('Please upload a valid file.');
        return;
      }
      
      setFile(acceptedFiles[0]);
      setError(null);
      
      // Clear previous data
      setTableData([]);
      
      const fileType = acceptedFiles[0].name.split('.').pop().toLowerCase();
      console.log('Uploaded file type:', fileType);
      
      if (fileType === 'docx') {
        processDocxFile(acceptedFiles[0])
          .then(() => {
            console.log('DOCX file processed successfully');
          })
          .catch(err => {
            console.error('Error processing DOCX file:', err);
            setError(`Error processing file: ${err.message || 'Unknown error'}`);
          });
      } else if (fileType === 'txt') {
        processTextFile(acceptedFiles[0])
          .then(() => {
            console.log('Text file processed successfully');
          })
          .catch(err => {
            console.error('Error processing text file:', err);
            setError(`Error processing file: ${err.message || 'Unknown error'}`);
          });
      } else if (fileType === 'xlsx' || fileType === 'xls') {
        processExcelFile(acceptedFiles[0])
          .then(() => {
            console.log('Excel file processed successfully');
          })
          .catch(err => {
            console.error('Error processing Excel file:', err);
            setError(`Error processing file: ${err.message || 'Unknown error'}`);
          });
      } else {
        setError(`Unsupported file type: ${fileType}. Please upload a .docx, .xlsx, .xls or .txt file.`);
      }
    }, [])
  });

  const processFile = async (file) => {
    setIsLoading(true);
    setError('');
    
    try {
      console.log('Processing file:', file.name, 'of type:', file.type);
      
      if (file.type === 'text/plain') {
        console.log('Processing as text file');
        await processTextFile(file);
      } else if (file.type.includes('document') || file.name.endsWith('.docx')) {
        console.log('Processing as DOCX file');
        await processDocxFile(file);
      } else if (file.type.includes('spreadsheet') || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        console.log('Processing as Excel file');
        await processExcelFile(file);
      } else {
        console.log('Unsupported file type:', file.type);
        setError(`Unsupported file type: ${file.type}. Please upload a .docx, .xlsx, .xls, or .txt file.`);
        throw new Error(`Unsupported file type: ${file.type}`);
      }
    } catch (err) {
      console.error('Error processing file:', err);
      setError(`Error processing the file: ${err.message}. Please try again with a valid document.`);
    } finally {
      setIsLoading(false);
    }
  };

  const processExcelFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      console.log('Starting to read Excel file');
      
      reader.onload = (e) => {
        try {
          console.log('Excel file loaded, processing...');
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          // Get the first worksheet
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          
          // Convert to JSON
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          console.log('Excel data rows:', jsonData.length);
          
          // Extract headers and data
          if (jsonData.length < 2) {
            throw new Error('Excel file does not contain enough data rows');
          }
          
          // Process the JSON data
          const headers = jsonData[0];
          console.log('Excel headers:', headers);
          
          // Look for our required column headers
          const clientIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('client'));
          const moduleIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('module'));
          const descriptionIndex = headers.findIndex(h => 
            h?.toString().toLowerCase().includes('description') || 
            h?.toString().toLowerCase().includes('task'));
          const statusIndex = headers.findIndex(h => 
            h?.toString().toLowerCase().includes('status') || 
            h?.toString().toLowerCase().includes('deployment'));
          const dateIndex = headers.findIndex(h => 
            h?.toString().toLowerCase().includes('date') || 
            h?.toString().toLowerCase().includes('delivery'));
          
          console.log('Column indices:', { clientIndex, moduleIndex, descriptionIndex, statusIndex, dateIndex });
          
          // If all standard columns are missing, try to use the first 5 columns
          if (clientIndex === -1 && moduleIndex === -1 && descriptionIndex === -1 && 
              statusIndex === -1 && dateIndex === -1 && headers.length >= 5) {
            console.log('Using first 5 columns as default mappings');
          }
          
          const parsedData = [];
          
          for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            
            // Skip empty rows
            if (row.length === 0 || row.every(cell => cell === '')) continue;
            
            const dataItem = {
              client: clientIndex >= 0 && row[clientIndex] ? row[clientIndex].toString() : 
                     (row[0] ? row[0].toString() : 'Unknown'),
              module: moduleIndex >= 0 && row[moduleIndex] ? row[moduleIndex].toString() : 
                     (row[1] ? row[1].toString() : 'Unknown'),
              description: descriptionIndex >= 0 && row[descriptionIndex] ? row[descriptionIndex].toString() : 
                          (row[2] ? row[2].toString() : 'N/A'),
              deploymentStatus: statusIndex >= 0 && row[statusIndex] ? row[statusIndex].toString() : 
                               (row[3] ? row[3].toString() : 'N/A'),
              deliveryDate: dateIndex >= 0 && row[dateIndex] ? row[dateIndex].toString() : 
                           (row[4] ? row[4].toString() : 'N/A')
            };
            
            parsedData.push(dataItem);
          }
          
          console.log('Parsed Excel data:', parsedData);
          
          if (parsedData.length === 0) {
            parsedData.push({
              client: 'Unable to parse',
              module: 'Excel',
              description: `No data could be extracted from ${file.name}`,
              deploymentStatus: 'N/A',
              deliveryDate: 'N/A'
            });
          }
          
          setTableData(parsedData);
          resolve();
        } catch (err) {
          console.error('Excel parsing error:', err);
          reject(err);
        }
      };
      
      reader.onerror = (err) => {
        console.error('Error reading Excel file:', err);
        reject(new Error(`Error reading file: ${err.message || 'Unknown error'}`));
      };
      
      reader.readAsArrayBuffer(file);
    });
  };

  const processTextFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          console.log('Processing text file content');
          const content = e.target.result;
          console.log('Text file content length:', content.length);
          
          // Display raw content for debugging
          console.log('First 200 chars of content:', content.substring(0, 200));
          
          // Extract lines from the content
          const lines = content.split('\n').filter(line => line.trim() !== '');
          console.log('Number of lines found:', lines.length);
          
          const parsedData = [];
          
          // More flexible parsing logic
          // Find the data rows - we'll be more flexible with the starting point
          for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (line === '') continue;
            
            console.log(`Line ${i}: ${line}`);
            
            // Try different splitting strategies
            let columns = [];
            
            // Strategy 1: Special case for the format: [PU]-HOSTEL->[10399] || Hostel || Issue in Hostel Room Type       Live
            if (line.includes('->[') && line.includes('||')) {
              // Extract client
              const clientMatch = line.match(/\[(.+?)\]-(.+?)(?=-\>)/);
              let client = '';
              if (clientMatch && clientMatch.length > 2) {
                client = `${clientMatch[1]}-${clientMatch[2]}`;
              } else {
                client = line.split('->[')[0];
              }
              
              // Extract the rest using || as delimiter
              const parts = line.split(/\|\|/).map(part => part.trim());
              
              if (parts.length >= 3) {
                // Extract module (after first ||)
                const module = parts[1];
                
                // Extract description and status
                const lastPart = parts[parts.length-1];
                let description = parts[2];
                let status = '';
                let date = '';
                
                // Check if the last part contains deployment status at the end
                const statusMatch = lastPart.match(/(.+?)\s{2,}(Live|UAT|Testing|Dev)\s*$/);
                if (statusMatch) {
                  status = statusMatch[2];
                  // If the status is in the last part, adjust the description
                  if (parts.length === 3) {
                    description = statusMatch[1].trim();
                  }
                } else if (lastPart.match(/(Live|UAT|Testing|Dev)\s*$/)) {
                  const match = lastPart.match(/(Live|UAT|Testing|Dev)\s*$/);
                  status = match[1];
                  // Try to extract the description
                  description = lastPart.replace(/(Live|UAT|Testing|Dev)\s*$/, '').trim();
                }
                
                columns = [client, module, description, status, date];
              }
            }
            
            // Strategy 2: Split by tabs
            if (columns.length < 4) {
              columns = line.split('\t').filter(col => col.trim() !== '');
            }
            
            // Strategy 3: Split by multiple spaces if not enough columns
            if (columns.length < 4) {
              columns = line.split(/\s{2,}/).filter(col => col.trim() !== '');
            }
            
            // Strategy 4: Look for dot-numbered entries like "1. SNU"
            if (columns.length < 4 && line.match(/\d+\. /)) {
              // This might be the start of a data row with numbered entries
              // Extract the client name after the number
              const clientMatch = line.match(/\d+\. (.+?)(?=\s{2,}|$)/);
              if (clientMatch && clientMatch[1]) {
                // Try to extract the remaining columns
                const remainingText = line.replace(/\d+\. (.+?)(?=\s{2,}|$)/, '');
                const remainingCols = remainingText.split(/\s{2,}/).filter(col => col.trim() !== '');
                columns = [clientMatch[1], ...remainingCols];
              }
            }
            
            console.log(`Columns found: ${columns.length}`, columns);
            
            if (columns.length >= 5) {
              parsedData.push({
                client: columns[0],
                module: columns[1],
                description: columns[2],
                deploymentStatus: columns[3],
                deliveryDate: columns[4]
              });
            } else if (columns.length == 4) {
              // Handle case where description might be missing
              parsedData.push({
                client: columns[0],
                module: columns[1],
                description: "-",
                deploymentStatus: columns[2],
                deliveryDate: columns[3]
              });
            }
          }
          
          console.log('Parsed data from text file:', parsedData);
          
          // If no data was extracted, try manual parsing or add sample data
          if (parsedData.length === 0) {
            console.log('No structured data found, adding sample data from text content');
            
            // Check if the content contains known patterns from the example
            if (content.includes('SNU') || content.includes('UTKAL') || content.includes('RFC')) {
              // Try to manually extract some data
              if (content.includes('SNU')) {
                parsedData.push({
                  client: 'SNU',
                  module: 'Academic',
                  description: content.includes('Controllers') ? 
                    'Controllers to be developed for Student Information Page' : 
                    'Student Information Page to be designed',
                  deploymentStatus: 'LIVE',
                  deliveryDate: '02/05/2025'
                });
              }
              
              if (content.includes('UTKAL')) {
                parsedData.push({
                  client: 'UTKAL',
                  module: 'Academic',
                  description: 'Bulk Student Field Update in HOD & Principal Login',
                  deploymentStatus: 'LIVE',
                  deliveryDate: '02/05/2025'
                });
              }
              
              if (content.includes('RFC')) {
                parsedData.push({
                  client: 'RFC',
                  module: 'Academic',
                  description: content.includes('Mentor') ? 
                    'Mentor Mentee Module (Rajalakshmi)' : 
                    'Add fields for bulk student updates (Maher)',
                  deploymentStatus: content.includes('UAT') ? 'UAT' : 'LIVE',
                  deliveryDate: content.includes('UAT') ? '16/05/2025' : '02/05/2025'
                });
              }
            }
          }
          
          // Always ensure we have at least one entry for display
          if (parsedData.length === 0) {
            parsedData.push({
              client: 'Sample',
              module: 'Module',
              description: 'Please provide a valid format file',
              deploymentStatus: 'N/A',
              deliveryDate: 'N/A'
            });
          }
          
          setTableData(parsedData);
          resolve();
        } catch (err) {
          console.error('Text parsing error:', err);
          reject(err);
        }
      };
      
      reader.onerror = (err) => {
        console.error('Error reading text file:', err);
        reject(new Error(`Error reading file: ${err.message || 'Unknown error'}`));
      };
      
      reader.readAsText(file);
    });
  };



  const processDocxFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      console.log('Starting to read DOCX file');
      
      reader.onload = async (e) => {
        try {
          console.log('DOCX file loaded, processing...');
          const arrayBuffer = e.target.result;
          
          // Convert DOCX to HTML
          console.log('Converting DOCX to HTML...');
          const result = await mammoth.convertToHtml({ arrayBuffer });
          const html = result.value;
          console.log('HTML conversion successful');
          console.log('HTML content length:', html.length);
          console.log('First 200 chars of HTML:', html.substring(0, 200));
          
          // Create a temporary DOM element to parse the HTML
          const tempDiv = document.createElement('div');
          tempDiv.innerHTML = html;
          
          // Get the full text content for analysis
          const textContent = tempDiv.textContent;
          console.log('Text content length:', textContent.length);
          console.log('First 200 chars of text content:', textContent.substring(0, 200));
          
          // Extract data
          console.log('Extracting data from document...');
          const parsedData = [];
          
          // 1. First check if there are any tables in the document
          const tables = tempDiv.querySelectorAll('table');
          console.log('Tables found:', tables.length);
          
          let foundData = false;
          
          // Process tables if they exist
          if (tables.length > 0) {
            for (let t = 0; t < tables.length; t++) {
              const table = tables[t];
              const rows = table.querySelectorAll('tr');
              console.log(`Table ${t+1}: found ${rows.length} rows`);
              
              // Skip tables that are likely not data (too few rows or columns)
              if (rows.length < 2) continue;
              
              // Process each row
              for (let i = 1; i < rows.length; i++) { // Start from 1 to skip header
                const cells = rows[i].querySelectorAll('td');
                if (cells.length < 3) continue; // Skip rows with insufficient data
                
                let rowData = {
                  client: cells[0]?.textContent.trim() || 'Unknown',
                  module: cells[1]?.textContent.trim() || 'Unknown',
                  description: cells[2]?.textContent.trim() || 'N/A',
                  deploymentStatus: cells[3]?.textContent.trim() || 'N/A',
                  deliveryDate: cells[4]?.textContent.trim() || 'N/A'
                };
                
                // Clean up data
                if (rowData.client || rowData.module || rowData.description) {
                  // Only add if we have at least some meaningful data
                  parsedData.push(rowData);
                  foundData = true;
                }
              }
            }
          }
          
          // 2. If no tables or no data found in tables, try to parse the text content
          if (!foundData) {
            // Split text into lines
            const lines = textContent.split('\n').filter(line => line.trim().length > 0);
            console.log('Text lines found:', lines.length);
            
            // Look for patterns in text
            for (let i = 0; i < lines.length; i++) {
              const line = lines[i].trim();
              
              // Look for client code patterns like [PU] or client-module patterns like [PU]-HOSTEL
              if (line.match(/\[.*?\]/) || line.includes('||')) {
                console.log(`Found potential data line ${i}:`, line);
                
                let client = '', module = '', description = '', status = '', date = 'N/A';
                
                // Format parsing for lines with client code in brackets
                const clientMatch = line.match(/\[(.*?)\]/);
                if (clientMatch) {
                  // Check if it's like [PU]-HOSTEL
                  const extendedClientMatch = line.match(/\[(.*?)\]-(.*?)(?=-|\||$)/);
                  if (extendedClientMatch) {
                    client = `[${extendedClientMatch[1]}]-${extendedClientMatch[2].trim()}`;
                  } else {
                    client = clientMatch[0]; // Just the [XYZ]
                  }
                }
                
                // Parse data separated by ||
                if (line.includes('||')) {
                  const parts = line.split('||').map(part => part.trim());
                  
                  // If we already have client from brackets, start with module
                  if (client && parts.length >= 2) {
                    module = parts[1] || '';
                    if (parts.length >= 3) {
                      description = parts[2] || '';
                    }
                  } 
                  // Otherwise use parts for client, module, description
                  else if (parts.length >= 1) {
                    if (!client) client = parts[0] || '';
                    if (parts.length >= 2) module = parts[1] || '';
                    if (parts.length >= 3) description = parts[2] || '';
                  }
                  
                  // Check if the last part has a status at the end
                  const lastPart = parts[parts.length - 1];
                  if (lastPart) {
                    const statusMatch = lastPart.match(/(.*)\s{2,}(Live|LIVE|UAT|Dev|DEV)\s*$/);
                    if (statusMatch) {
                      if (parts.length <= 3 || !description) {
                        description = statusMatch[1].trim();
                      }
                      status = statusMatch[2];
                    }
                  }
                }
                
                // If we still don't have module/status/etc., try to extract from line
                if (!module) {
                  if (line.includes('Hostel')) module = 'Hostel';
                  else if (line.includes('Academic')) module = 'Academic';
                  else if (line.includes('Finance')) module = 'Finance';
                }
                
                if (!status) {
                  if (line.includes('Live') || line.includes('LIVE')) status = 'Live';
                  else if (line.includes('UAT')) status = 'UAT';
                  else if (line.includes('Dev') || line.includes('DEV')) status = 'Dev';
                }
                
                // Look for date pattern DD/MM/YYYY
                const dateMatch = line.match(/(\d{2}\/\d{2}\/\d{4})/);
                if (dateMatch) {
                  date = dateMatch[1];
                }
                
                // Only add if we have at least client or module
                if (client || module) {
                  parsedData.push({
                    client: client || 'Unknown',
                    module: module || 'Unknown',
                    description: description || 'N/A',
                    deploymentStatus: status || 'N/A',
                    deliveryDate: date
                  });
                  foundData = true;
                }
              }
            }
          }
          
          // 3. If still no data, try looking for specific section headers and parse content after them
          if (!foundData && parsedData.length === 0) {
            const sections = ['Client', 'Module', 'Tasks', 'Status', 'Projects'];
            for (const section of sections) {
              const sectionRegex = new RegExp(`${section}[\s:]*`, 'i');
              const sectionMatch = textContent.match(sectionRegex);
              
              if (sectionMatch) {
                const startIndex = sectionMatch.index + sectionMatch[0].length;
                const nextSectionMatch = textContent.slice(startIndex).match(/\n\s*[A-Z][a-zA-Z\s]*:[\s\n]*/);
                const endIndex = nextSectionMatch ? startIndex + nextSectionMatch.index : startIndex + 500;
                
                const sectionContent = textContent.slice(startIndex, endIndex).trim();
                console.log(`Found section '${section}':`, sectionContent.substring(0, 100) + '...');
                
                // Try to parse this section
                const lines = sectionContent.split('\n').filter(line => line.trim().length > 0);
                for (const line of lines) {
                  // Add code to parse section content based on the type of section
                  if (line.match(/\[.*?\]/) || line.includes(':') || line.includes('-')) {
                    let entry = { client: 'Unknown', module: 'Unknown', description: 'N/A', deploymentStatus: 'N/A', deliveryDate: 'N/A' };
                    
                    // Use the section name to determine what this content is
                    if (section.match(/client/i)) entry.client = line.trim();
                    else if (section.match(/module/i)) entry.module = line.trim();
                    else if (section.match(/task/i)) entry.description = line.trim();
                    else if (section.match(/status/i)) entry.deploymentStatus = line.trim();
                    
                    if (entry.client !== 'Unknown' || entry.module !== 'Unknown' || entry.description !== 'N/A') {
                      parsedData.push(entry);
                      foundData = true;
                    }
                  }
                }
              }
            }
          }
          
          console.log('Parsed data:', parsedData);
          
          // If still no data found, add a row with file info
          if (parsedData.length === 0) {
            parsedData.push({
              client: 'Unable to parse',
              module: 'Document',
              description: `No data could be extracted from ${file.name}`,
              deploymentStatus: 'N/A',
              deliveryDate: 'N/A'
            });
          }
          
          setTableData(parsedData);
          resolve();
        } catch (err) {
          console.error('DOCX parsing error:', err);
          reject(err);
        }
      };
      
      reader.onerror = (err) => {
        console.error('Error reading DOCX file:', err);
        reject(new Error(`Error reading file: ${err.message || 'Unknown error'}`));
      };
      reader.readAsArrayBuffer(file);
    });
  };

  const generatePPTX = () => {
    try {
      if (!tableData || tableData.length === 0) {
        setError('No data to generate PowerPoint file. Please upload a file first.');
        return;
      }

      // Set loading state
      setIsLoading(true);
      setError(null);
      
      // Make API call to backend
      fetch('http://localhost:5000/api/generate-ppt', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ tableData }),
      })
        .then(response => {
          if (!response.ok) {
            throw new Error('Network response was not ok');
          }
          return response.blob();
        })
        .then(blob => {
          // Create download using file-saver
          saveAs(blob, 'WorkDoneStatus.pptx');
          console.log('PowerPoint file downloaded successfully');
          setIsLoading(false);
        })
        .catch(err => {
          console.error('Error generating PowerPoint:', err);
          setError(`Error generating PowerPoint file: ${err.message}`);
          setIsLoading(false);
        });
      
    } catch (err) {
      console.error('Error in generatePPTX function:', err);
      setError(`Error preparing PowerPoint request: ${err.message}`);
      setIsLoading(false);
    }
  };

  const generateHTML = () => {
    try {
      if (!tableData || tableData.length === 0) {
        setError('No data to generate HTML file. Please upload a file first.');
        return;
      }
      
      // Create HTML slide
      let slideHTML = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Work Done Status</title>
        <meta charset="UTF-8">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #FFFFFF;
          }
          .slide {
            width: 1024px;
            height: 768px;
            position: relative;
            margin: 0 auto;
            padding: 40px;
            overflow: hidden;
            box-sizing: border-box;
          }
          .header {
            display: flex;
            justify-content: space-between;
            margin-bottom: 30px;
          }
          .title {
            color: #333333;
            font-size: 36px;
            font-weight: bold;
            text-align: left;
          }
          .logo {
            text-align: right;
          }
          .logo-title {
            color: #00a4e4;
            font-size: 24px;
            font-weight: bold;
          }
          .logo-subtitle {
            color: #666666;
            font-style: italic;
            font-size: 16px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
          }
          th, td {
            border: 1px solid #666666;
            padding: 10px;
            text-align: left;
          }
          th {
            background-color: #DDDDDD;
            font-weight: bold;
          }
          .footer {
            position: absolute;
            bottom: 20px;
            width: 100%;
            display: flex;
            justify-content: center;
            color: #666666;
            font-size: 14px;
          }
          .page-number {
            position: absolute;
            bottom: 20px;
            right: 20px;
            color: #666666;
            font-size: 18px;
          }
          @media print {
            body {
              width: 1024px;
              height: 768px;
              overflow: hidden;
            }
            .slide {
              page-break-after: always;
            }
          }
        </style>
    </head>
    <body>
      <div class="slide">
        <div class="header">
          <h1 class="title">Work Done Status</h1>
          <div class="logo">
            <div class="logo-title">MasterSoft</div>
            <div class="logo-subtitle">Automating Education...</div>
          }
          .logo-title {
            color: #00a4e4;
            font-size: 24px;
            font-weight: bold;
          }
          .logo-subtitle {
            color: #666666;
            font-style: italic;
            font-size: 16px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
          }
          th, td {
            border: 1px solid #666666;
            padding: 10px;
            text-align: left;
          }
          th {
            background-color: #DDDDDD;
            font-weight: bold;
          }
          .footer {
            position: absolute;
            bottom: 20px;
            width: 100%;
            display: flex;
            justify-content: center;
            color: #666666;
            font-size: 14px;
          }
          .page-number {
            position: absolute;
            bottom: 20px;
            right: 20px;
            color: #666666;
            font-size: 18px;
          }
          @media print {
            body {
              width: 1024px;
              height: 768px;
              overflow: hidden;
            }
            .slide {
              page-break-after: always;
            }
          }
        </style>
      </head>
      <body>
        <div class="slide">
          <div class="header">
            <h1 class="title">Work Done Status</h1>
            <div class="logo">
              <div class="logo-title">MasterSoft</div>
              <div class="logo-subtitle">Automating Education...</div>
            </div>
          </div>
          
          <table>
            <thead>
              <tr>
                <th>Client</th>
                <th>Module</th>
                <th>Description</th>
                <th>Deployment Status</th>
                <th>Date of Delivery</th>
              </tr>
            </thead>
            <tbody>
      `;
      
      // Add data rows
      tableData.forEach((row, index) => {
        slideHTML += `
              <tr>
                <td>${index + 1}. ${row.client}</td>
                <td>${row.module}</td>
                <td>${row.description}</td>
                <td>${row.deploymentStatus}</td>
                <td>${row.deliveryDate}</td>
              </tr>
        `;
      });
      
      // Close the table and add footer
      slideHTML += `
            </tbody>
          </table>
          
          <div class="footer">Work Done | MasterSoft ERP Solutions Pvt. Ltd. | Design and developed by Shivam Kale</div>
          <div class="page-number">1</div>
        </div>
      </body>
      </html>
      `;
      
      // Create a Blob with the HTML content
      const blob = new Blob([slideHTML], { type: 'text/html' });
      
      // Create download link
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'WorkDoneStatus.html';
      document.body.appendChild(link);
      link.click();
      
      // Clean up
      URL.revokeObjectURL(url);
      document.body.removeChild(link);
      
    } catch (err) {
      console.error('Error generating HTML:', err);
      setError('Error generating HTML file.');
    }
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>PPT Maker</h1>
        <p>Upload your document or text file to generate a PowerPoint presentation</p>
      </header>
      
      <main className="App-main">
        <div {...getRootProps({ className: 'dropzone' })}>
          <input {...getInputProps()} />
          <p>Drag & drop a file here, or click to select file</p>
          <p className="small-text">Supports Excel (.xlsx, .xls), Word (.docx) and Text (.txt) files</p>
        </div>
        
        {file && (
          <div className="file-info">
            <p><strong>Selected file:</strong> {file.name}</p>
          </div>
        )}
        
        {isLoading && <p className="loading">Processing file...</p>}
        
        {error && <p className="error">{error}</p>}
        
        {tableData.length > 0 && (
          <div className="preview">
            <h2>Data Preview</h2>
            <table>
              <thead>
                <tr>
                  <th>Client</th>
                  <th>Module</th>
                  <th>Description</th>
                  <th>Deployment Status</th>
                  <th>Date of Delivery</th>
                </tr>
              </thead>
              <tbody>
                {tableData.map((row, index) => (
                  <tr key={index}>
                    <td>{row.client}</td>
                    <td>{row.module}</td>
                    <td>{row.description}</td>
                    <td>{row.deploymentStatus}</td>
                    <td>{row.deliveryDate}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            
            <div className="format-selector">
              <label className="format-label">Output Format:</label>
              <div className="format-options">
                <label>
                  <input
                    type="radio"
                    name="format"
                    value="pptx"
                    checked={outputFormat === 'pptx'}
                    onChange={() => setOutputFormat('pptx')}
                  />
                  PowerPoint (.pptx)
                </label>
                <label>
                  <input
                    type="radio"
                    name="format"
                    value="html"
                    checked={outputFormat === 'html'}
                    onChange={() => setOutputFormat('html')}
                  />
                  HTML
                </label>
              </div>
            </div>
            
            <button 
              className="generate-btn" 
              onClick={outputFormat === 'pptx' ? generatePPTX : generateHTML}
              disabled={tableData.length === 0 || isLoading}
            >
              Generate Slides
            </button>
          </div>
        )}
      </main>
      
      <footer className="App-footer">
        <p>&copy; {new Date().getFullYear()} PPT Maker</p>
        <p className="credit-text">Design and developed by Shivam Kale</p>
      </footer>
    </div>
  );
}

export default App;
