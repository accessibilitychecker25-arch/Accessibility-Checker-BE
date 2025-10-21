const express = require('express');
const multer = require('multer');
const cors = require('cors');
const stream = require('stream');
const fs = require('fs');
const path = require('path');

const mammoth = require('mammoth');
const puppeteer = require('puppeteer');

const ruleDescriptions = {
  'Tagged PDF':
    'The document must be properly tagged to support screen readers and assistive technology.',
  Title: 'The document should have a descriptive title displayed in the title bar.',
  'Tagged content':
    'All page content must be tagged so it can be read in the correct order by assistive technologies.',
  'Figures alternate text': 'All images must include meaningful alternative text descriptions.',
  'Nested alternate text':
    'Avoid adding multiple or nested alt texts that will not be read by screen readers.',
  'Associated with content': 'Alternative text must be tied to the correct content element.',
  'Hides annotation': 'Alternative text must not hide or obscure annotations in the document.',
  'Other elements alternate text':
    'Other elements, such as form fields or artifacts, must also have appropriate alt text when needed.',
  Rows: 'Every table row (TR) must be inside a valid table structure (Table, THead, TBody, or TFoot).',
  'TH and TD': 'Table headers (TH) and cells (TD) must always be inside table rows (TR).',
  Headers: 'Tables should use headers to identify row and column meaning.',
  Regularity:
    'All table rows and columns must follow a consistent structure without missing cells.',
  Summary:
    'Tables should include a summary to explain their purpose and structure to users with disabilities.',
  'List items': 'Every list item (LI) must be a child of a list (L).',
  'Lbl and LBody': 'List labels (Lbl) and list bodies (LBody) must be inside list items (LI).',
  'Appropriate nesting': 'Lists and other elements must be nested in a logical, accessible order.',
  'Logical Reading Order':
    'The content must be structured so that it follows a natural and logical reading sequence.',
  'Color contrast':
    'Text and visuals must have sufficient color contrast to be readable by users with visual impairments.',
};

const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  PDFAccessibilityCheckerJob,
  PDFAccessibilityCheckerResult,
} = require('@adobe/pdfservices-node-sdk');
require('dotenv').config();

const app = express();

// CORS configuration: allow GitHub Pages, Vercel frontend, and localhost
const allowedOrigins = [
  'http://localhost:4200',
  'https://kmoreland126.github.io',
  'https://accessibility-checker-rose.vercel.app',
];

app.use(
  cors({
    origin: function (origin, callback) {
      // Allow requests with no origin (like mobile apps, curl, Postman)
      if (!origin) return callback(null, true);
      
      if (allowedOrigins.indexOf(origin) !== -1) {
        callback(null, true);
      } else {
        callback(new Error('Disallowed CORS origin'));
      }
    },
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type'],
    credentials: false,
  })
);

const upload = multer({ storage: multer.memoryStorage() });

app.post('/upload-pdf', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).send('No file uploaded');

  let readStream;
  let mimeType = req.file.mimetype;
  let pdfBuffer;

  try {
    const credentials = new ServicePrincipalCredentials({
      clientId: process.env.ADOBE_CLIENT_ID,
      clientSecret: process.env.ADOBE_CLIENT_SECRET,
    });

    const pdfServices = new PDFServices({ credentials });

    if (
      mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      req.file.originalname.endsWith('.docx')
    ) {
      const mammothResult = await mammoth.convertToHtml({ buffer: req.file.buffer });
      const html = mammothResult.value;

      const browser = await puppeteer.launch();
      const page = await browser.newPage();
      await page.setContent(html);
      pdfBuffer = await page.pdf({ format: 'A4' });
      await browser.close();

      readStream = new stream.PassThrough();
      readStream.end(pdfBuffer);
      mimeType = MimeType.PDF;
    } else {
      readStream = new stream.PassThrough();
      readStream.end(req.file.buffer);
      mimeType = MimeType.PDF;
    }

    const inputAsset = await pdfServices.upload({
      readStream,
      mimeType: MimeType.PDF,
    });

    const job = new PDFAccessibilityCheckerJob({ inputAsset });
    const pollingURL = await pdfServices.submit({ job });

    const pdfServicesResponse = await pdfServices.getJobResult({
      pollingURL,
      resultType: PDFAccessibilityCheckerResult,
    });

    const resultAssetReport = pdfServicesResponse.result.report;
    const streamAssetReport = await pdfServices.getContent({ asset: resultAssetReport });

    // Collect the JSON into memory
    let data = '';
    for await (const chunk of streamAssetReport.readStream) {
      data += chunk.toString();
    }

    // Parse and send JSON response
    const reportJson = JSON.parse(data);
    // Helper: flatten all rule arrays
    const allRules = Object.values(reportJson['Detailed Report'] || {}).flat();

    const enhanceRules = (rules) =>
      rules.map((r) => ({
        ...r,
        Description: ruleDescriptions[r.Rule] || r.Description,
      }));

    const failedRules = enhanceRules(allRules.filter((r) => r.Status === 'Failed'));
    const manualCheckRules = enhanceRules(
      allRules.filter((r) => r.Status === 'Needs manual check')
    );

    const passedCount = allRules.filter((r) => r.Status === 'Passed').length;

    // Build response
    const filteredResult = {
      fileName: req.file.originalname,
      summary: {
        successCount: passedCount,
        failedCount: failedRules.length,
        manualCheckCount: manualCheckRules.length,
      },
      failures: failedRules,
      needsManualCheck: manualCheckRules,
    };

    res.json(filteredResult);
  } catch (err) {
    console.error('Error processing PDF:', err);
    res.status(500).json({ error: 'Error processing PDF' });
  } finally {
    readStream?.destroy();
  }
});

app.listen(3000, () => console.log('Server running on port 3000'));

// const express = require('express');
// const multer = require('multer');
// const cors = require('cors');

// const app = express();

// app.use(
//   cors({
//     origin: 'http://localhost:4200',
//     methods: ['GET', 'POST', 'OPTIONS'],
//     allowedHeaders: ['Content-Type'],
//   })
// );

// const upload = multer({ storage: multer.memoryStorage() });

// app.post('/upload-pdf', upload.single('file'), async (req, res) => {
//   if (!req.file) return res.status(400).send('No file uploaded');

//   console.log('File name:', req.file.originalname);
//   console.log('File size (bytes):', req.file.size);
//   console.log('File mimetype:', req.file.mimetype);

//   // --- Dummy JSON for testing ---
//   const dummyReport = {
//     fileName: req.file.originalname,
//     uploadedAt: new Date().toISOString(),
//     wcagResults: [
//       { rule: '1.1.1 Non-text Content', status: 'pass' },
//       { rule: '1.2.1 Audio-only and Video-only (Prerecorded)', status: 'fail' },
//       { rule: '1.3.1 Info and Relationships', status: 'pass' },
//     ],
//     summary: {
//       totalRules: 3,
//       passed: 2,
//       failed: 1,
//     },
//   };

//   // Send as JSON file download
//   res.setHeader('Content-Disposition', 'attachment; filename="accessibility-report.json"');
//   res.setHeader('Content-Type', 'application/json');
//   res.send(JSON.stringify(dummyReport, null, 2));
// });

// app.listen(3000, () => console.log('Server running on port 3000'));
