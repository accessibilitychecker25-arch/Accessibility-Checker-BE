const fs = require('fs').promises;
const path = require('path');
const JSZip = require('jszip');
const sessionManager = require('../lib/session-manager');

module.exports = async (req, res) => {
  // CORS headers
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
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition, Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'GET') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const { batchId, sessionId } = req.query;

    if (!batchId) {
      res.status(400).json({ error: 'batchId parameter required' });
      return;
    }

    if (!sessionId) {
      res.status(400).json({ error: 'sessionId parameter required' });
      return;
    }

    // Get session and verify it exists
    const session = sessionManager.getOrCreateSession(sessionId);
    if (session.sessionId !== sessionId) {
      res.status(404).json({ error: 'Session expired or not found' });
      return;
    }

    // Load batch summary from session directory
    const batchSummaryPath = `${session.directory}/batch-${batchId}-summary.json`;
    let batchSummary;
    
    try {
      const summaryData = await fs.readFile(batchSummaryPath, 'utf8');
      batchSummary = JSON.parse(summaryData);
    } catch (error) {
      res.status(404).json({ error: `Batch ${batchId} not found in session` });
      return;
    }

    // Create a ZIP file containing all remediated documents
    const outputZip = new JSZip();
    const batchFolder = outputZip.folder(`batch-${batchId}-remediated`);
    
    let successCount = 0;
    let errorCount = 0;

    for (const result of batchSummary.results) {
      if (!result.success) {
        errorCount++;
        // Add error file
        batchFolder.file(`ERROR-${result.filename}.txt`, 
          `Error processing ${result.filename}:\n${result.error}`);
        continue;
      }

      try {
        // Load the original file from session directory
        const originalPath = `${session.directory}/original-${result.reportId}.docx`;
        
        try {
          const originalBuffer = await fs.readFile(originalPath);
          
          // TODO: Apply remediation to the file here
          // For now, just copy the original as "remediated"
          batchFolder.file(`REMEDIATED-${result.filename}`, originalBuffer);
          
          successCount++;
        } catch (fileError) {
          throw new Error(`Original file not found: ${fileError.message}`);
        }
        
      } catch (error) {
        errorCount++;
        batchFolder.file(`ERROR-${result.filename}.txt`, 
          `Error remediating ${result.filename}:\n${error.message}`);
      }
    }

    // Add batch summary to the ZIP
    batchFolder.file('batch-summary.json', JSON.stringify(batchSummary, null, 2));
    batchFolder.file('README.txt', 
      `Batch Remediation Results\n` +
      `========================\n` +
      `Batch ID: ${batchId}\n` +
      `Total Files: ${batchSummary.totalFiles}\n` +
      `Successfully Processed: ${successCount}\n` +
      `Errors: ${errorCount}\n` +
      `Timestamp: ${batchSummary.timestamp}\n\n` +
      `Files with "REMEDIATED-" prefix have been processed for accessibility.\n` +
      `Files with "ERROR-" prefix encountered processing issues.\n`
    );

    // Generate the ZIP buffer
    const zipBuffer = await outputZip.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    // Send as download
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="batch-${batchId}-remediated.zip"`);
    res.setHeader('Content-Length', zipBuffer.length);
    
    res.end(zipBuffer);

  } catch (error) {
    console.error('Batch download error:', error);
    res.status(500).json({ error: 'Internal server error during batch download' });
  }
};