const fs = require('fs').promises;
const path = require('path');

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
  res.setHeader('Access-Control-Allow-Methods', 'GET, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const { action, reportId, batchId, limit = 50 } = req.query;

  try {
    switch (req.method) {
      case 'GET':
        if (action === 'list') {
          await listReports(req, res, { limit: parseInt(limit) });
        } else if (action === 'batches') {
          await listBatches(req, res);
        } else if (reportId) {
          await getReport(req, res, reportId);
        } else if (batchId) {
          await getBatch(req, res, batchId);
        } else {
          res.status(400).json({ error: 'Missing action or ID parameter' });
        }
        break;
        
      case 'DELETE':
        if (reportId) {
          await deleteReport(req, res, reportId);
        } else if (batchId) {
          await deleteBatch(req, res, batchId);
        } else {
          res.status(400).json({ error: 'Missing reportId or batchId parameter' });
        }
        break;
        
      default:
        res.status(405).json({ error: 'Method not allowed' });
    }
  } catch (error) {
    console.error('Reports API error:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
};

async function listReports(req, res, options = {}) {
  const reportsDir = 'reports';
  const files = await fs.readdir(reportsDir);
  
  // Filter for individual reports (not batch summaries)
  const reportFiles = files
    .filter(f => f.endsWith('-accessibility-report.json'))
    .sort((a, b) => {
      // Sort by timestamp (newest first)
      const aTime = parseInt(a.split('-')[0]);
      const bTime = parseInt(b.split('-')[0]);
      return bTime - aTime;
    })
    .slice(0, options.limit);

  const reports = [];
  
  for (const file of reportFiles) {
    try {
      const filePath = path.join(reportsDir, file);
      const content = await fs.readFile(filePath, 'utf8');
      const report = JSON.parse(content);
      
      reports.push({
        reportId: report.reportId,
        filename: report.filename,
        timestamp: report.timestamp,
        summary: report.summary,
        filePath: file
      });
    } catch (error) {
      console.warn(`Failed to read report ${file}:`, error.message);
    }
  }

  res.json({
    totalReports: reports.length,
    reports: reports
  });
}

async function listBatches(req, res) {
  const reportsDir = 'reports';
  const files = await fs.readdir(reportsDir);
  
  // Filter for batch summaries
  const batchFiles = files
    .filter(f => f.startsWith('batch-') && f.endsWith('-summary.json'))
    .sort((a, b) => {
      // Sort by timestamp (newest first)
      const aTime = parseInt(a.split('-')[1]);
      const bTime = parseInt(b.split('-')[1]);
      return bTime - aTime;
    });

  const batches = [];
  
  for (const file of batchFiles) {
    try {
      const filePath = path.join(reportsDir, file);
      const content = await fs.readFile(filePath, 'utf8');
      const batch = JSON.parse(content);
      
      batches.push({
        batchId: batch.batchId,
        timestamp: batch.timestamp,
        totalFiles: batch.totalFiles,
        successful: batch.results.filter(r => r.success).length,
        failed: batch.results.filter(r => !r.success).length,
        filePath: file
      });
    } catch (error) {
      console.warn(`Failed to read batch ${file}:`, error.message);
    }
  }

  res.json({
    totalBatches: batches.length,
    batches: batches
  });
}

async function getReport(req, res, reportId) {
  const reportPath = `reports/${reportId}-accessibility-report.json`;
  
  try {
    const content = await fs.readFile(reportPath, 'utf8');
    const report = JSON.parse(content);
    res.json(report);
  } catch (error) {
    res.status(404).json({ error: `Report ${reportId} not found` });
  }
}

async function getBatch(req, res, batchId) {
  const batchPath = `reports/batch-${batchId}-summary.json`;
  
  try {
    const content = await fs.readFile(batchPath, 'utf8');
    const batch = JSON.parse(content);
    res.json(batch);
  } catch (error) {
    res.status(404).json({ error: `Batch ${batchId} not found` });
  }
}

async function deleteReport(req, res, reportId) {
  const reportPath = `reports/${reportId}-accessibility-report.json`;
  
  try {
    await fs.unlink(reportPath);
    res.json({ message: `Report ${reportId} deleted successfully` });
  } catch (error) {
    res.status(404).json({ error: `Report ${reportId} not found` });
  }
}

async function deleteBatch(req, res, batchId) {
  const batchPath = `reports/batch-${batchId}-summary.json`;
  
  try {
    await fs.unlink(batchPath);
    
    // Also delete individual reports from this batch if they exist
    // This is optional - you might want to keep individual reports
    
    res.json({ message: `Batch ${batchId} deleted successfully` });
  } catch (error) {
    res.status(404).json({ error: `Batch ${batchId} not found` });
  }
}