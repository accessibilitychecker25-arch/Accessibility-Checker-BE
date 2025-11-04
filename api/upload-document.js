const Busboy = require('busboy');
const JSZip = require('jszip');

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
  // Expose Content-Disposition for downloads and Content-Type for clients
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
      textShadowsRemoved: false,
      fontsNormalized: false,
      fontSizesNormalized: false,
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

    // Shadow detection deferred to analyzeShadowsAndFonts (single source of truth)
    
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
    
    // Check for document protection (documentProtection, writeProtection, readOnlyRecommended)
    const settingsXml = await zip.file('word/settings.xml')?.async('string');
    if (settingsXml) {
      if (
        settingsXml.includes('<w:documentProtection') ||
        settingsXml.includes('<w:writeProtection') ||
        settingsXml.includes('<w:readOnlyRecommended') ||
        settingsXml.includes('<w:editRestrictions') ||
        settingsXml.includes('<w:formProtection') ||
        settingsXml.includes('<w:locked')
      ) {
        report.details.documentProtected = true;
        report.summary.flagged += 1;
      }
    }
    
    // Check for text shadows, serif fonts, and small font sizes
    const shadowFontResults = await analyzeShadowsAndFonts(zip);
    if (shadowFontResults.hasShadows) {
      report.details.textShadowsRemoved = true;
      report.summary.fixed += 1;
    }
    if (shadowFontResults.hasSerifFonts || shadowFontResults.hasSmallFonts) {
      report.details.fontsNormalized = true;
      report.details.fontSizesNormalized = true;
      report.summary.fixed += 1;
    }
    
    // Scan for any remaining suspicious protection/encryption markers to help debug Protected View
    try {
      const remaining = await scanForRemainingProtection(fileData);
      if (remaining && remaining.length) {
        report.details.remainingProtectionMatches = remaining;
      }
    } catch (e) {
      // ignore scanning errors
    }

    return report;
    
  } catch (error) {
// Helper: scan the package bytes for remaining protection/encryption-related strings
function findRemainingProtectionMatches(zip) {
  const needleRegex = /documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|locked|enforcement|hash|salt|cryptProvider|EncryptedPackage|encryption/gi;
  const matches = new Set();
  zip.forEach((relativePath, file) => {
    if (!relativePath) return;
    const lc = relativePath.toLowerCase();
    // Only scan XML parts and rels
    if (lc.endsWith('.xml') || lc.endsWith('.rels') || lc === '[content_types].xml') {
      const txt = file.async('string').catch(() => '');
      // note: file.async returns promise; we'll collect promises later
    }
  });
  // Synchronous scan isn't possible here because async calls are needed.
  // We'll implement a full async scan where this helper is used.
  return null;
}
    return {
      fileName: filename,
      error: error.message,
      summary: { fixed: 0, flagged: 0 },
      details: {}
    };
  }
// Analyze shadows and fonts in the document
async function analyzeShadowsAndFonts(zip) {
  const results = {
    hasShadows: false,
    hasSerifFonts: false,
    hasSmallFonts: false
  };

  // Function to check for shadow tags only (conservative, avoids false positives)
  const hasShadowEffects = (xmlContent) => {
    if (!xmlContent) return false;
    const tagRegex = /<\s*(?:w|a|w14|w15):(?:shadow|outerShdw|innerShdw|prstShdw|glow|reflection|props3d)\b/i;
    return tagRegex.test(xmlContent);
  };

  // Check document.xml for shadows, fonts, and sizes
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (documentXml) {
    // Check for all shadow types
    if (hasShadowEffects(documentXml)) {
      results.hasShadows = true;
    }
    
    // Check for serif fonts (Times, Times New Roman, Georgia, etc.)
    if (/(Times|Georgia|Garamond|serif)/i.test(documentXml)) {
      results.hasSerifFonts = true;
    }
    
    // Check for small font sizes (less than 22 half-points = 11pt)
    const sizeMatches = documentXml.match(/<w:sz w:val="(\d+)"/g);
    if (sizeMatches) {
      for (const match of sizeMatches) {
        const size = parseInt(match.match(/\d+/)[0]);
        if (size < 22) {
          results.hasSmallFonts = true;
          break;
        }
      }
    }
  }

  // Check styles.xml for shadows, fonts, and sizes
  const stylesXml = await zip.file('word/styles.xml')?.async('string');
  if (stylesXml) {
    if (!results.hasShadows && hasShadowEffects(stylesXml)) {
      results.hasShadows = true;
    }
    
    if (!results.hasSerifFonts && /(Times|Georgia|Garamond|serif)/i.test(stylesXml)) {
      results.hasSerifFonts = true;
    }
    
    if (!results.hasSmallFonts) {
      const sizeMatches = stylesXml.match(/<w:sz w:val="(\d+)"/g);
      if (sizeMatches) {
        for (const match of sizeMatches) {
          const size = parseInt(match.match(/\d+/)[0]);
          if (size < 22) {
            results.hasSmallFonts = true;
            break;
          }
        }
      }
    }
  }

  // Check theme files for advanced shadow effects
  if (!results.hasShadows) {
    const themeFiles = Object.keys(zip.files).filter(name => name.includes('theme') && name.endsWith('.xml'));
    for (const themeFileName of themeFiles) {
      const themeFile = zip.file(themeFileName);
      if (themeFile) {
        const themeXml = await themeFile.async('string');
        if (hasShadowEffects(themeXml)) {
          results.hasShadows = true;
          break;
        }
      }
    }
  }

  return results;
}

// Async scan for remaining protection/encryption related strings in xml parts
async function scanForRemainingProtection(fileData) {
  const zip = await JSZip.loadAsync(fileData);
  const needleRegex = /documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|locked|enforcement|hash|salt|cryptProvider|EncryptedPackage|encryption/gi;
  const findings = [];
  const promises = [];
  zip.forEach((relativePath, file) => {
    const lc = (relativePath || '').toLowerCase();
    if (lc.endsWith('.xml') || lc.endsWith('.rels') || lc === '[content_types].xml') {
      const p = file.async('string').then(txt => {
        const m = txt.match(needleRegex);
        if (m && m.length) {
          findings.push({ part: relativePath, matches: Array.from(new Set(m)) });
        }
      }).catch(() => {});
      promises.push(p);
    }
  });
  await Promise.all(promises);
  return findings;
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
