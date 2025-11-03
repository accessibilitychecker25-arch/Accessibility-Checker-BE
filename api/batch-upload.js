const Busboy = require('busboy');
const JSZip = require('jszip');
const fs = require('fs').promises;
const path = require('path');
const sessionManager = require('../lib/session-manager');

module.exports = async (req, res) => {
  // CORS: safe allowlist — echo back the requesting Origin when allowed.
  const ALLOWED_ORIGINS = [
    'https://accessibilitychecker25-arch.github.io',
    'https://kmoreland126.github.io',
    'http://localhost:3000',
    'http://localhost:4200'
  ];
  const origin = req.headers.origin;
  if (origin && ALLOWED_ORIGINS.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
  }
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition, Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const busboy = Busboy({ headers: req.headers });
    const uploadedFiles = []; // Store multiple files
    const MAX_FILES = 10; // Allow up to 10 files per batch
    let fileCount = 0;

    busboy.on('file', (fieldname, file, info) => {
      fileCount++;
      
      if (fileCount > MAX_FILES) {
        file.resume(); // Drain the file stream
        return;
      }

      const filename = info.filename;
      const chunks = [];
      
      file.on('data', (chunk) => {
        chunks.push(chunk);
      });
      
      file.on('end', () => {
        const fileData = Buffer.concat(chunks);
        uploadedFiles.push({
          filename: filename,
          data: fileData,
          size: fileData.length
        });
      });
    });

    busboy.on('finish', async () => {
      if (uploadedFiles.length === 0) {
        res.status(400).json({ error: 'No valid files uploaded' });
        return;
      }

      if (fileCount > MAX_FILES) {
        res.status(400).json({ 
          error: `Too many files. Maximum ${MAX_FILES} files allowed per batch.`,
          received: fileCount
        });
        return;
      }

      // Get or create session
      const sessionId = req.headers['x-session-id'] || req.query.sessionId;
      const session = sessionManager.getOrCreateSession(sessionId);

      // Process each file and generate individual reports
      const batchResults = {
        batchId: Date.now(),
        sessionId: session.sessionId,
        timestamp: new Date().toISOString(),
        totalFiles: uploadedFiles.length,
        results: []
      };

      for (let i = 0; i < uploadedFiles.length; i++) {
        const fileInfo = uploadedFiles[i];
        
        try {
          console.log(`Processing file ${i + 1}/${uploadedFiles.length}: ${fileInfo.filename}`);
          
          // Process individual file (reuse existing logic)
          const fileResult = await processSingleFile(fileInfo, session.directory);
          
          // Add file to session
          sessionManager.addFileToSession(session.sessionId, {
            filename: fileInfo.filename,
            reportId: fileResult.reportId,
            originalPath: fileResult.originalFilePath,
            reportPath: fileResult.reportPath,
            processedAt: new Date().toISOString()
          });
          
          batchResults.results.push({
            fileIndex: i + 1,
            filename: fileInfo.filename,
            fileSize: fileInfo.size,
            success: true,
            reportId: fileResult.reportId,
            ...fileResult.report
          });
          
        } catch (error) {
          console.error(`Error processing ${fileInfo.filename}:`, error);
          
          batchResults.results.push({
            fileIndex: i + 1,
            filename: fileInfo.filename,
            fileSize: fileInfo.size,
            success: false,
            error: error.message
          });
        }
      }

      // Save batch summary to session directory
      const batchReportPath = `${session.directory}/batch-${batchResults.batchId}-summary.json`;
      await fs.writeFile(batchReportPath, JSON.stringify(batchResults, null, 2));
      
      // Add batch to session
      sessionManager.addBatchToSession(session.sessionId, {
        batchId: batchResults.batchId,
        timestamp: batchResults.timestamp,
        totalFiles: batchResults.totalFiles,
        successful: batchResults.results.filter(r => r.success).length,
        failed: batchResults.results.filter(r => !r.success).length,
        reportPath: batchReportPath
      });

      // Return batch summary with session info
      res.json({
        message: `Successfully processed batch of ${uploadedFiles.length} files`,
        sessionId: session.sessionId,
        batchId: batchResults.batchId,
        summary: {
          totalFiles: batchResults.totalFiles,
          successful: batchResults.results.filter(r => r.success).length,
          failed: batchResults.results.filter(r => !r.success).length
        },
        results: batchResults.results,
        expiresIn: '1 hour'
      });
    });

    req.pipe(busboy);

  } catch (error) {
    console.error('Batch upload error:', error);
    res.status(500).json({ error: 'Internal server error during batch processing' });
  }
};

// Extract single file processing logic (from existing upload-document.js)
async function processSingleFile(fileInfo, sessionDirectory) {
  const { filename, data } = fileInfo;
  
  // Validate DOCX file
  if (!filename.toLowerCase().endsWith('.docx')) {
    throw new Error(`Invalid file type: ${filename}. Only .docx files are supported.`);
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch (error) {
    throw new Error(`Invalid DOCX file: ${filename}. Unable to read as ZIP archive.`);
  }

  // Generate unique report ID for this file
  const reportId = `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  
  // Initialize report structure
  const report = {
    filename: filename,
    reportId: reportId,
    timestamp: new Date().toISOString(),
    summary: {
      flagged: 0,
      fixed: 0
    },
    details: {
      hasProtection: false,
      removedProtection: false,
      languageDefaultFixed: null,
      titleNeedsFixing: false,
      textShadowsRemoved: false,
      fontsNormalized: false,
      fontSizesNormalized: false
    }
  };

  // Run all analysis functions (copied from existing logic)
  await analyzeDocumentStructure(zip, report);
  await analyzeProtection(zip, report);
  const shadowFontResults = await analyzeShadowsAndFonts(zip);
  
  // Update report with shadow/font analysis
  if (shadowFontResults.hasShadows) {
    report.details.textShadowsRemoved = false; // Will be true after remediation
    report.summary.flagged++;
  }
  
  if (shadowFontResults.hasSerifFonts) {
    report.details.fontsNormalized = false; // Will be true after remediation  
    report.summary.flagged++;
  }
  
  if (shadowFontResults.hasSmallFonts) {
    report.details.fontSizesNormalized = false; // Will be true after remediation
    report.summary.flagged++;
  }

  // Save original file and report to session directory (not permanent storage)
  const originalFilePath = `${sessionDirectory}/original-${reportId}.docx`;
  const reportPath = `${sessionDirectory}/${reportId}-accessibility-report.json`;
  
  await fs.writeFile(originalFilePath, data);
  await fs.writeFile(reportPath, JSON.stringify(report, null, 2));

  return {
    reportId: reportId,
    report: report,
    reportPath: reportPath,
    originalFilePath: originalFilePath
  };
}

// Copy existing analysis functions (you'll need to import these)
async function analyzeDocumentStructure(zip, report) {
  // Implementation from existing upload-document.js
  // ... existing logic ...
}

async function analyzeProtection(zip, report) {
  // Implementation from existing upload-document.js  
  // ... existing logic ...
}

async function analyzeShadowsAndFonts(zip) {
  // Implementation from existing upload-document.js
  // ... existing logic ...
}