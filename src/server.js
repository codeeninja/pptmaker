const express = require('express');
const cors = require('cors');
const multer = require('multer');
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json());

// Set up multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + path.extname(file.originalname));
  }
});

const upload = multer({ storage });

// Ensure uploads directory exists
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// Route for PowerPoint generation
app.post('/api/generate-ppt', (req, res) => {
  try {
    console.log('Received request to generate PowerPoint');
    const data = req.body;
    
    if (!data || !data.tableData || !Array.isArray(data.tableData) || data.tableData.length === 0) {
      console.error('Invalid data received:', data);
      return res.status(400).json({ success: false, message: 'Invalid or missing data' });
    }
    
    console.log(`Received ${data.tableData.length} rows of data for PowerPoint generation`);
    
    // Create a new PowerPoint presentation
    const pptx = new pptxgen();
    
    // Set presentation properties
    pptx.author = 'Shivam Kale';
    pptx.company = 'MasterSoft ERP Solutions Pvt. Ltd.';
    pptx.subject = 'Work Done Status';
    pptx.title = 'Work Done Status';
    
    // Define maximum rows per slide and calculate total slides needed
    const MAX_ROWS_PER_SLIDE = 8; // Reduced for better readability with long text
    const totalRows = data.tableData.length;
    const totalSlides = Math.ceil(totalRows / MAX_ROWS_PER_SLIDE);
    
    console.log(`Creating ${totalSlides} slides for ${totalRows} data rows`);
    
    // Set master slide layout to ensure consistency
    pptx.defineSlideMaster({
      title: 'MASTER_SLIDE',
      background: { color: 'FFFFFF' },
      objects: [
        // Footer with credit line
        { 
          text: { 
            text: 'Work Done | MasterSoft ERP Solutions Pvt. Ltd. | Design and developed by Shivam Kale',
            options: { x: 0.5, y: 6.8, w: 9.0, h: 0.3, fontSize: 10, color: '666666', align: 'center' }
          }
        }
      ]
    });
    
    // Create all slides with paginated content
    for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
      console.log(`Creating slide ${slideIndex + 1} of ${totalSlides}`);
      
      // Create a new slide
      const slide = pptx.addSlide();
      
      // Add title to slide
      slide.addText('Work Done Status', {
        x: 0.5, y: 0.5,
        w: 5.0, h: 0.8,
        fontSize: 36,
        color: '333333',
        bold: true,
        fontFace: 'Arial'
      });
      
      // Add MasterSoft logo and text
      slide.addText('MasterSoft', {
        x: 7.5, y: 0.5,
        w: 2.0, h: 0.5,
        fontSize: 24,
        color: '00a4e4',
        bold: true,
        fontFace: 'Arial',
        align: 'right'
      });
      
      slide.addText('Automating Education...', {
        x: 7.5, y: 1.0,
        w: 2.0, h: 0.5,
        fontSize: 16,
        color: '666666',
        italic: true,
        fontFace: 'Arial',
        align: 'right'
      });
      
      // Calculate which rows to include on this slide
      const startRow = slideIndex * MAX_ROWS_PER_SLIDE;
      const endRow = Math.min(startRow + MAX_ROWS_PER_SLIDE, totalRows);
      const slideData = data.tableData.slice(startRow, endRow);
      
      console.log(`Slide ${slideIndex + 1}: Showing rows ${startRow + 1} to ${endRow} (${slideData.length} rows)`);
      
      // Determine if we need to split any rows with very long descriptions
      const MAX_CHARS_PER_CELL = 250; // Maximum characters before creating a new slide
      let needsSplit = false;
      
      // Check if any description is too long
      slideData.forEach(row => {
        if (row.description && row.description.length > MAX_CHARS_PER_CELL) {
          needsSplit = true;
          console.log(`Found very long description (${row.description.length} chars) that may need splitting`);
        }
      });
      
      // If we don't need to split, proceed normally
      if (!needsSplit) {
        // Create table rows with header
        const tableRows = [
          [
            { text: 'Client', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
            { text: 'Module', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
            { text: 'Description', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
            { text: 'Deployment Status', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
            { text: 'Date of Delivery', options: { bold: true, fill: 'DDDDDD', color: '000000' } }
          ]
        ];
        
        // Add this slide's portion of data rows
        slideData.forEach((row, index) => {
          // Calculate the global index for consistent numbering across slides
          const globalIndex = startRow + index;
          tableRows.push([
            { text: `${globalIndex + 1}. ${row.client}` },
            { text: row.module },
            { text: row.description },
            { text: row.deploymentStatus },
            { text: row.deliveryDate }
          ]);
        });
        
        // Add table to current slide with explicit column widths and text wrapping
        slide.addTable(tableRows, {
          x: 0.5, y: 1.8,
          w: 9.0,
          h: 4.5, // Set a fixed height to prevent overflow
          fontSize: 11, // Slightly reduced font size
          fontFace: 'Arial',
          border: { type: 'solid', color: '666666', pt: 1 },
          colW: [1.3, 1.2, 4.0, 1.5, 1.0], // Adjust column widths (wider for description)
          autoPage: true, // Enable automatic pagination for tall tables
          autoPageRepeatHeader: true, // Repeat headers on new pages
          autoPageHeaderRows: 1, // Number of header rows to repeat
          autoPageLineWeight: 0.5,
          wordWrap: true, // Enable text wrapping in cells
          fill: { color: 'FFFFFF' } // White background for better readability
        });
      } else {
        // For slides with long descriptions, handle each row with special care
        slideData.forEach((row, index) => {
          // Calculate the global index for consistent numbering across slides
          const globalIndex = startRow + index;
          
          // Create a table just for this row
          const singleRowTable = [
            [
              { text: 'Client', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
              { text: 'Module', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
              { text: 'Description', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
              { text: 'Deployment Status', options: { bold: true, fill: 'DDDDDD', color: '000000' } },
              { text: 'Date of Delivery', options: { bold: true, fill: 'DDDDDD', color: '000000' } }
            ],
            [
              { text: `${globalIndex + 1}. ${row.client}` },
              { text: row.module },
              { text: row.description },
              { text: row.deploymentStatus },
              { text: row.deliveryDate }
            ]
          ];
          
          // Calculate vertical position for this row's table
          const yPos = 1.8 + (index * 1.0); // Space each table down the slide
          
          // Add the table with text wrapping enabled
          slide.addTable(singleRowTable, {
            x: 0.5, y: yPos,
            w: 9.0,
            fontSize: 12,
            fontFace: 'Arial',
            border: { type: 'solid', color: '666666', pt: 1 },
            colW: [1.5, 1.5, 3.5, 1.5, 1.0], // Adjust column widths
            autoPage: true, // Enable automatic pagination
            autoPageRepeatHeader: true,
            autoPageHeaderRows: 1,
            autoPageLineWeight: 0.5,
            wordWrap: true // Enable text wrapping
          });
        });
      }
      
      // Add footer with credit line
      slide.addText('Work Done | MasterSoft ERP Solutions Pvt. Ltd. | Design and developed by Shivam Kale', {
        x: 0.5, y: 6.8,
        w: 9.0, h: 0.3,
        fontSize: 10,
        color: '666666',
        fontFace: 'Arial',
        align: 'center'
      });
      
      // Add page number
      slide.addText(`${slideIndex + 1}/${totalSlides}`, {
        x: 9.0, y: 6.8,
        w: 0.5, h: 0.3,
        fontSize: 10,
        color: '666666',
        fontFace: 'Arial',
        align: 'right'
      });
    }
    
    // Generate PowerPoint file
    const fileName = `WorkDoneStatus_${Date.now()}.pptx`;
    const filePath = path.join(__dirname, 'uploads', fileName);
    
    pptx.writeFile({ fileName: filePath })
      .then(() => {
        // Send the file as a download
        res.download(filePath, 'WorkDoneStatus.pptx', (err) => {
          if (err) {
            console.error('Error sending file:', err);
          }
          
          // Clean up the file after sending
          fs.unlink(filePath, (err) => {
            if (err) console.error('Error deleting file:', err);
          });
        });
      })
      .catch(err => {
        console.error('Error writing PowerPoint file:', err);
        res.status(500).json({ success: false, message: 'Error generating PowerPoint file' });
      });
    
  } catch (err) {
    console.error('Server error:', err);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});

// Document parsing endpoint
app.post('/api/parse-document', upload.single('file'), (req, res) => {
  try {
    // Simply return success for now - actual parsing will be implemented as needed
    res.json({ success: true, message: 'Document received', fileName: req.file.filename });
  } catch (err) {
    console.error('Error parsing document:', err);
    res.status(500).json({ success: false, message: 'Error parsing document' });
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`PPT Maker Backend - Design and developed by Shivam Kale`);
});
