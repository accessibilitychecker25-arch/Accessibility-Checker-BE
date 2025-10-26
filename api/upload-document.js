const Busboy = require('busboy');
const JSZip = require('jszip');

module.exports = async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', 'https://accessibilitychecker25-arch.github.io');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Expose-Headers', 'Content-Type');

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
    let fileData = null;
    let filename = null;

    busboy.on('file', (fieldname, file, info) => {
      filename = info.filename;
      const chunks = [];
      
      file.on('data', (chunk) => {
        chunks.push(chunk);
      });
      
      file.on('end', () => {
        fileData = Buffer.concat(chunks);
      });
    });

    busboy.on('finish', async () => {
      if (!fileData || !filename) {
        res.status(400).json({ error: 'No file uploaded' });
        return;
      }

      if (!filename.toLowerCase().endsWith('.docx')) {
        res.status(400).json({ error: 'Please upload a .docx file' });
        return;
      }

      try {
        const report = await analyzeDocx(fileData, filename);
        res.status(200).json({
          fileName: filename,
          suggestedFileName: filename,
          report: report
        });
      } catch (error) {
        res.status(500).json({ error: error.message });
      }
    });

    req.pipe(busboy);

  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

async function analyzeDocx(fileData, filename) {
  const report = {
    fileName: filename,
    suggestedFileName: filename,
    summary: { fixed: 0, flagged: 0 },
    details: {
      titleNeedsFixing: false,
      imagesMissingOrBadAlt: 0,
      gifsDetected: [],
      fileNameNeedsFixing: false,
      emptyHeadings: [],
      headingOrderIssues: [],
      tablesHeaderRowSet: [],
      documentProtected: false,
    }
  };

  try {
    const zip = await JSZip.loadAsync(fileData);
    
    // Check title
    const coreXml = await zip.file('docProps/core.xml')?.async('string');
    if (coreXml) {
      if (coreXml.includes('<dc:title></dc:title>') || coreXml.includes('<dc:title/>')) {
        report.details.titleNeedsFixing = true;
        report.summary.flagged += 1;
      }
    }
    
    // Check for images
    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
    if (relsXml) {
      const imageMatches = relsXml.match(/relationships\/image"/g);
      const imageCount = imageMatches ? imageMatches.length : 0;
      if (imageCount > 0) {
        report.details.imagesMissingOrBadAlt = imageCount;
        report.summary.flagged += imageCount;
      }
    }
    
    // Check for GIFs
    const gifFiles = [];
    zip.forEach((relativePath, file) => {
      if (relativePath.startsWith('word/media/') && relativePath.toLowerCase().endsWith('.gif')) {
        gifFiles.push(relativePath);
      }
    });
    if (gifFiles.length > 0) {
      report.details.gifsDetected = gifFiles;
      report.summary.flagged += gifFiles.length;
    }
    
    // Check filename
    if (filename.includes('_') || filename.toLowerCase().startsWith('document') || filename.toLowerCase().startsWith('untitled')) {
      report.details.fileNameNeedsFixing = true;
      report.summary.flagged += 1;
    }
    
    // Check headings for empty content and proper order
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (documentXml) {
      const headingResults = analyzeHeadings(documentXml);
      
      if (headingResults.emptyHeadings.length > 0) {
        report.details.emptyHeadings = headingResults.emptyHeadings;
        report.summary.flagged += headingResults.emptyHeadings.length;
      }
      
      if (headingResults.orderIssues.length > 0) {
        report.details.headingOrderIssues = headingResults.orderIssues;
        report.summary.flagged += headingResults.orderIssues.length;
      }
    }
    
    // Check for document protection
    const settingsXml = await zip.file('word/settings.xml')?.async('string');
    if (settingsXml) {
      if (settingsXml.includes('<w:documentProtection')) {
        report.details.documentProtected = true;
        report.summary.flagged += 1;
      }
    }
    
    return report;
    
  } catch (error) {
    return {
      fileName: filename,
      error: error.message,
      summary: { fixed: 0, flagged: 0 },
      details: {}
    };
  }
}

function analyzeHeadings(documentXml) {
  const emptyHeadings = [];
  const orderIssues = [];
  const headings = [];
  
  // Match all paragraphs with heading styles
  // Pattern: <w:pStyle w:val="Heading1"/> (or Heading2, etc.)
  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];
  
  paragraphs.forEach((para, index) => {
    // Check if this paragraph has a heading style
    const headingMatch = para.match(/<w:pStyle w:val="Heading(\d+)"\/>/);
    
    if (headingMatch) {
      const level = parseInt(headingMatch[1]);
      
      // Extract text content from the paragraph
      const textMatches = para.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
      let text = '';
      
      if (textMatches) {
        text = textMatches
          .map(t => t.replace(/<w:t[^>]*>|<\/w:t>/g, ''))
          .join('')
          .trim();
      }
      
      // Check for empty headings
      if (!text || text.length === 0) {
        emptyHeadings.push({
          level: level,
          position: headings.length + 1,
          message: `Heading ${level} is empty`
        });
      }
      
      // Track heading for order checking
      headings.push({ level, text, position: headings.length + 1 });
    }
  });
  
  // Check heading order
  for (let i = 1; i < headings.length; i++) {
    const prev = headings[i - 1];
    const current = headings[i];
    
    // Headings should not skip levels (e.g., H1 -> H3 is wrong)
    if (current.level > prev.level + 1) {
      orderIssues.push({
        position: current.position,
        message: `Heading ${current.level} follows Heading ${prev.level} - skipped level ${prev.level + 1}`,
        previousLevel: prev.level,
        currentLevel: current.level,
        text: current.text
      });
    }
  }
  
  return { emptyHeadings, orderIssues };
}
