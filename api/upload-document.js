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
      // Changed from "fixed" to "flagged" - user needs to manually address these
      lineSpacingNeedsFixing: false,
      fontSizeNeedsFixing: false,
      fontTypeNeedsFixing: false,
      linkNamesNeedImprovement: false,
    }
  };

  try {
    const zip = await JSZip.loadAsync(fileData);
    
    // Read core documents once for analysis
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const coreXml = await zip.file('docProps/core.xml')?.async('string');
    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
    
    // Check title - requires user action to fix
    if (coreXml) {
      if (coreXml.includes('<dc:title></dc:title>') || coreXml.includes('<dc:title/>')) {
        report.details.titleNeedsFixing = true;
        report.summary.flagged += 1;
      }
    }
    
    // Check for images with location details
    if (relsXml && documentXml) {
      const imageAnalysis = analyzeImageLocations(documentXml, relsXml);
      if (imageAnalysis.imagesWithoutAlt.length > 0) {
        report.details.imagesMissingOrBadAlt = imageAnalysis.imagesWithoutAlt.length;
        report.details.imageLocations = imageAnalysis.imagesWithoutAlt;
        report.summary.flagged += imageAnalysis.imagesWithoutAlt.length;
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
      
      // Check for non-descriptive link text
      const linkResults = analyzeLinkDescriptiveness(documentXml);
      if (linkResults.nonDescriptiveLinks.length > 0) {
        report.details.linkNamesNeedImprovement = true;
        report.details.linkLocations = linkResults.nonDescriptiveLinks;
        report.summary.flagged += linkResults.nonDescriptiveLinks.length;
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
        report.summary.fixed += 1;
      }
    }
    
    // Check for text shadows, serif fonts, and small font sizes
    const shadowFontResults = await analyzeShadowsAndFonts(zip);
    // attach debug info about which parts contained shadow tags (if any)
    if (shadowFontResults && Array.isArray(shadowFontResults.shadowParts) && shadowFontResults.shadowParts.length > 0) {
      report.details.shadowParts = shadowFontResults.shadowParts;
      report.details.textShadowsRemoved = true;
      report.summary.fixed += 1;
      console.log('[analyzeDocx] shadows detected, fix count now:', report.summary.fixed);
    } else {
      // ensure falsey/empty detection doesn't report a fix
      report.details.textShadowsRemoved = false;
    }
    // These are now flagged for user attention instead of auto-fixed (with location details)
    if (shadowFontResults.hasSerifFonts) {
      report.details.fontTypeNeedsFixing = true;
      report.details.fontTypeLocations = shadowFontResults.fontIssueLocations.filter(issue => issue.type === 'serif');
      report.summary.flagged += 1;
      console.log('[analyzeDocx] serif fonts detected, flagged count now:', report.summary.flagged);
    }
    if (shadowFontResults.hasSmallFonts) {
      report.details.fontSizeNeedsFixing = true;
      report.details.fontSizeLocations = shadowFontResults.fontIssueLocations.filter(issue => issue.type === 'small');
      report.summary.flagged += 1;
      console.log('[analyzeDocx] small fonts detected, flagged count now:', report.summary.flagged);
    }
    if (shadowFontResults.hasInsufficientLineSpacing) {
      report.details.lineSpacingNeedsFixing = true;
      report.details.lineSpacingLocations = shadowFontResults.lineSpacingIssueLocations;
      report.summary.flagged += 1;
      console.log('[analyzeDocx] insufficient line spacing detected, flagged count now:', report.summary.flagged);
    }
    
    console.log('[analyzeDocx] FINAL fix count before return:', report.summary.fixed);
    
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
// Analyze shadows and fonts in the document with location details
async function analyzeShadowsAndFonts(zip) {
  const results = {
    hasShadows: false,
    hasSerifFonts: false,
    hasSmallFonts: false,
    hasInsufficientLineSpacing: false
  };

  // Track detailed locations of issues
  results.shadowParts = [];
  results.fontIssueLocations = [];
  results.lineSpacingIssueLocations = [];

  // Function to check for shadow tags only (conservative, avoids false positives)
  // Returns the matched tag string or null
  // Excludes theme-level shadows that don't represent actual text shadow formatting
  const hasShadowEffects = (xmlContent, partName = '') => {
    if (!xmlContent) return null;
    
    // For theme files, only check for direct text shadow elements, not default theme shadows
    if (partName.toLowerCase().includes('theme')) {
      // Only match w14/w15 text effect shadows in themes, not default a:outerShdw
      const themeTextShadowRegex = /<\s*(?:w14|w15):(?:shadow|glow|reflection|props3d)\b[^>]*>/i;
      const m = xmlContent.match(themeTextShadowRegex);
      return m ? m[0] : null;
    }
    
    // For document/styles, check all shadow types including DrawingML
    const tagRegex = /<\s*(?:w|a|w14|w15):(?:shadow|outerShdw|innerShdw|prstShdw|glow|reflection|props3d)\b[^>]*>/i;
    const m = xmlContent.match(tagRegex);
    return m ? m[0] : null;
  };

  // Analyze document.xml with detailed location tracking
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (documentXml) {
    const locationResults = analyzeDocumentLocations(documentXml);
    
    // Merge results with location details
    if (locationResults.shadows.length > 0) {
      results.hasShadows = true;
      results.shadowParts = locationResults.shadows;
    }
    
    if (locationResults.fontIssues.length > 0) {
      results.hasSerifFonts = locationResults.fontIssues.some(issue => issue.type === 'serif');
      results.hasSmallFonts = locationResults.fontIssues.some(issue => issue.type === 'small');
      results.fontIssueLocations = locationResults.fontIssues;
    }
    
    if (locationResults.lineSpacingIssues.length > 0) {
      results.hasInsufficientLineSpacing = true;
      results.lineSpacingIssueLocations = locationResults.lineSpacingIssues;
    }
  }

  // Check styles.xml for shadows, fonts, and sizes
  const stylesXml = await zip.file('word/styles.xml')?.async('string');
  if (stylesXml) {
    const m = hasShadowEffects(stylesXml, 'word/styles.xml');
    if (!results.hasShadows && m) {
      results.hasShadows = true;
      results.shadowParts.push({ part: 'word/styles.xml', match: m.slice(0, 200) });
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
    
    // Check styles.xml for insufficient line spacing
    if (!results.hasInsufficientLineSpacing) {
      const spacingMatches = stylesXml.match(/<w:spacing[^>]*w:line="(\d+)"[^>]*\/>/g);
      if (spacingMatches) {
        for (const match of spacingMatches) {
          const lineValue = parseInt(match.match(/w:line="(\d+)"/)[1]);
          if (lineValue < 360) {
            results.hasInsufficientLineSpacing = true;
            break;
          }
        }
      }
      
      // Check for exact line spacing in styles
      if (!results.hasInsufficientLineSpacing && stylesXml.includes('w:lineRule="exact"')) {
        results.hasInsufficientLineSpacing = true;
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
        const m = hasShadowEffects(themeXml, themeFileName);
        if (m) {
          results.hasShadows = true;
          results.shadowParts.push({ part: themeFileName, match: m.slice(0, 200) });
          break;
        }
      }
    }
  }

  return results;
}

// Analyze document structure and find specific locations of accessibility issues
function analyzeDocumentLocations(documentXml) {
  const results = {
    shadows: [],
    fontIssues: [],
    lineSpacingIssues: []
  };

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Split into paragraphs for analysis
  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    // Check for page breaks to estimate page numbers
    if (paragraph.includes('<w:br w:type="page"/>') || paragraph.includes('<w:lastRenderedPageBreak/>')) {
      approximatePageNumber++;
    }
    
    // Check if this is a heading to track document structure
    const headingMatch = paragraph.match(/<w:pStyle w:val="(Heading\d+)"\/>/);
    if (headingMatch) {
      const headingText = extractTextFromParagraph(paragraph);
      currentHeading = `${headingMatch[1]}: ${headingText.substring(0, 50)}${headingText.length > 50 ? '...' : ''}`;
    }

    // Check for shadow effects in this paragraph
    const shadowRegex = /<\s*(?:w|a|w14|w15):(?:shadow|outerShdw|innerShdw|prstShdw|glow|reflection|props3d)\b[^>]*>/i;
    if (shadowRegex.test(paragraph)) {
      results.shadows.push({
        location: `Paragraph ${paragraphCount}`,
        approximatePage: approximatePageNumber,
        context: currentHeading || 'Document body',
        preview: extractTextFromParagraph(paragraph).substring(0, 100)
      });
    }

    // Check for font issues
    const fontMatches = paragraph.match(/w:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"/g);
    if (fontMatches) {
      for (const fontMatch of fontMatches) {
        const fontName = fontMatch.match(/"([^"]+)"/)[1];
        if (/(Times|Georgia|Garamond|serif)/i.test(fontName)) {
          results.fontIssues.push({
            type: 'serif',
            font: fontName,
            location: `Paragraph ${paragraphCount}`,
            approximatePage: approximatePageNumber,
            context: currentHeading || 'Document body',
            preview: extractTextFromParagraph(paragraph).substring(0, 100)
          });
          break; // Only record once per paragraph
        }
      }
    }

    // Check for small font sizes
    const sizeMatches = paragraph.match(/<w:sz w:val="(\d+)"/g);
    if (sizeMatches) {
      for (const sizeMatch of sizeMatches) {
        const size = parseInt(sizeMatch.match(/\d+/)[0]);
        if (size < 22) {
          results.fontIssues.push({
            type: 'small',
            size: `${size/2}pt`,
            location: `Paragraph ${paragraphCount}`,
            approximatePage: approximatePageNumber,
            context: currentHeading || 'Document body',
            preview: extractTextFromParagraph(paragraph).substring(0, 100)
          });
          break; // Only record once per paragraph
        }
      }
    }

    // Check for line spacing issues
    const spacingMatch = paragraph.match(/<w:spacing[^>]*w:line="(\d+)"[^>]*\/>/);
    if (spacingMatch) {
      const lineValue = parseInt(spacingMatch[1]);
      if (lineValue < 360) {
        results.lineSpacingIssues.push({
          type: 'insufficient',
          currentSpacing: `${(lineValue/240).toFixed(1)}x`,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 100)
        });
      }
    } else if (paragraph.includes('w:lineRule="exact"')) {
      results.lineSpacingIssues.push({
        type: 'exact',
        currentSpacing: 'Exact spacing',
        location: `Paragraph ${paragraphCount}`,
        approximatePage: approximatePageNumber,
        context: currentHeading || 'Document body',
        preview: extractTextFromParagraph(paragraph).substring(0, 100)
      });
    } else if (!paragraph.includes('<w:spacing') && extractTextFromParagraph(paragraph).trim().length > 0) {
      // Paragraph with text but no spacing defined
      results.lineSpacingIssues.push({
        type: 'missing',
        currentSpacing: 'Default spacing',
        location: `Paragraph ${paragraphCount}`,
        approximatePage: approximatePageNumber,
        context: currentHeading || 'Document body',
        preview: extractTextFromParagraph(paragraph).substring(0, 100)
      });
    }
  });

  return results;
}

// Extract readable text from a paragraph XML element
function extractTextFromParagraph(paragraphXml) {
  const textMatches = paragraphXml.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!textMatches) return '';
  
  return textMatches
    .map(t => t.replace(/<w:t[^>]*>|<\/w:t>/g, ''))
    .join('')
    .trim();
}

// Analyze link descriptiveness in the document
function analyzeLinkDescriptiveness(documentXml) {
  const results = {
    nonDescriptiveLinks: []
  };

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Generic/non-descriptive phrases to detect
  const genericPhrases = [
    'click here', 'here', 'read more', 'more', 'link', 'this link',
    'see more', 'learn more', 'find out more', 'more info', 'more information',
    'view more', 'details', 'continue', 'next', 'go', 'visit', 'see',
    'download', 'open', 'access', 'view', 'show', 'display', 'get',
    'this', 'that', 'these', 'those', 'it', 'page', 'site', 'website',
    'url', 'address', 'location'
  ];

  // Split into paragraphs for analysis
  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    // Check for page breaks
    if (paragraph.includes('<w:br w:type="page"/>') || paragraph.includes('<w:lastRenderedPageBreak/>')) {
      approximatePageNumber++;
    }
    
    // Track headings
    const headingMatch = paragraph.match(/<w:pStyle w:val="(Heading\d+)"\/>/);
    if (headingMatch) {
      const headingText = extractTextFromParagraph(paragraph);
      currentHeading = `${headingMatch[1]}: ${headingText.substring(0, 50)}${headingText.length > 50 ? '...' : ''}`;
    }

    // Check for hyperlinks in this paragraph
    // Word hyperlinks are stored as <w:hyperlink> elements or <w:fldSimple> with HYPERLINK
    const hyperlinkMatches = paragraph.match(/<w:hyperlink[^>]*>[\s\S]*?<\/w:hyperlink>/g) || [];
    const fieldHyperlinkMatches = paragraph.match(/<w:fldSimple[^>]*fldChar[^>]*HYPERLINK[^>]*>[\s\S]*?<\/w:fldSimple>/g) || [];
    
    // Also check for runs that might contain hyperlink formatting
    const runHyperlinkMatches = paragraph.match(/<w:r[^>]*>[\s\S]*?<w:rStyle w:val="Hyperlink"[^>]*\/>[\s\S]*?<\/w:r>/g) || [];
    
    const allLinks = [...hyperlinkMatches, ...fieldHyperlinkMatches, ...runHyperlinkMatches];
    
    allLinks.forEach(link => {
      const linkText = extractTextFromParagraph(link).trim().toLowerCase();
      
      if (linkText && linkText.length > 0) {
        // Check if the link text is non-descriptive
        const isGeneric = genericPhrases.some(phrase => {
          // Exact match or the link text is just the generic phrase
          return linkText === phrase || 
                 linkText === phrase + '.' || 
                 linkText === phrase + '!' ||
                 linkText.startsWith(phrase + ' ') ||
                 linkText.endsWith(' ' + phrase) ||
                 (linkText.length <= 15 && linkText.includes(phrase)); // Short phrases containing generic words
        });
        
        // Also flag very short links (likely non-descriptive)
        const isTooShort = linkText.length <= 3 && !linkText.match(/^[a-z]{2,3}$/); // Allow abbreviations like "FAQ", "PDF"
        
        // Flag URLs or email addresses used as link text
        const isUrl = linkText.includes('www.') || linkText.includes('http') || linkText.includes('.com') || linkText.includes('.org');
        const isEmail = linkText.includes('@') && linkText.includes('.');
        
        if (isGeneric || isTooShort || isUrl || isEmail) {
          let issueType = 'generic';
          if (isTooShort) issueType = 'too-short';
          if (isUrl) issueType = 'url-as-text';
          if (isEmail) issueType = 'email-as-text';
          
          results.nonDescriptiveLinks.push({
            type: issueType,
            linkText: linkText,
            location: `Paragraph ${paragraphCount}`,
            approximatePage: approximatePageNumber,
            context: currentHeading || 'Document body',
            preview: extractTextFromParagraph(paragraph).substring(0, 150),
            recommendation: generateLinkRecommendation(linkText, issueType)
          });
        }
      }
    });
  });

  return results;
}

// Generate recommendations for improving link text
function generateLinkRecommendation(linkText, issueType) {
  switch (issueType) {
    case 'generic':
      if (linkText.includes('click here') || linkText.includes('here')) {
        return 'Replace with descriptive text like "Download the user guide" or "View our services"';
      }
      if (linkText.includes('read more') || linkText.includes('more')) {
        return 'Replace with specific text like "Read the full research report" or "Learn about our methodology"';
      }
      return 'Use descriptive text that explains where the link goes or what action it performs';
    
    case 'too-short':
      return 'Expand to include more context, e.g., "Go" → "Go to our contact page"';
    
    case 'url-as-text':
      return 'Replace URL with descriptive text like "Visit our company website" or "Access the online portal"';
    
    case 'email-as-text':
      return 'Replace email with descriptive text like "Contact our support team" or "Email our sales department"';
    
    default:
      return 'Use clear, descriptive language that tells users where the link will take them';
  }
}

// Analyze image locations and alt text in the document
function analyzeImageLocations(documentXml, relsXml) {
  const results = {
    imagesWithoutAlt: []
  };

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Get image relationships
  const imageRels = {};
  const relMatches = relsXml.match(/<Relationship[^>]*Type="[^"]*\/image"[^>]*>/g) || [];
  relMatches.forEach(rel => {
    const idMatch = rel.match(/Id="([^"]+)"/);
    const targetMatch = rel.match(/Target="([^"]+)"/);
    if (idMatch && targetMatch) {
      imageRels[idMatch[1]] = targetMatch[1];
    }
  });

  // Split into paragraphs for analysis
  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    // Check for page breaks
    if (paragraph.includes('<w:br w:type="page"/>') || paragraph.includes('<w:lastRenderedPageBreak/>')) {
      approximatePageNumber++;
    }
    
    // Track headings
    const headingMatch = paragraph.match(/<w:pStyle w:val="(Heading\d+)"\/>/);
    if (headingMatch) {
      const headingText = extractTextFromParagraph(paragraph);
      currentHeading = `${headingMatch[1]}: ${headingText.substring(0, 50)}${headingText.length > 50 ? '...' : ''}`;
    }

    // Check for images in this paragraph
    const drawingMatches = paragraph.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g) || [];
    drawingMatches.forEach(drawing => {
      // Check for image references
      const blipMatches = drawing.match(/<a:blip r:embed="([^"]+)"/g) || [];
      blipMatches.forEach(blip => {
        const embedId = blip.match(/r:embed="([^"]+)"/)[1];
        const imagePath = imageRels[embedId];
        
        // Check for alt text
        const hasAltText = drawing.includes('<wp:docPr') && 
                          (drawing.includes('descr="') || drawing.includes('title="')) &&
                          !drawing.includes('descr=""') && !drawing.includes('title=""');
        
        if (!hasAltText) {
          results.imagesWithoutAlt.push({
            location: `Paragraph ${paragraphCount}`,
            approximatePage: approximatePageNumber,
            context: currentHeading || 'Document body',
            imagePath: imagePath || 'Unknown image',
            preview: extractTextFromParagraph(paragraph).substring(0, 100) || 'No surrounding text'
          });
        }
      });
    });
  });

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
