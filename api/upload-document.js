const Busboy = require('busboy');
const JSZip = require('jszip');

// Helper function to extract text from paragraph XML - moved to top for availability
function extractTextFromParagraph(paragraphXml) {
  const textMatches = paragraphXml.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!textMatches) return '';
  
  return textMatches
    .map(t => t.replace(/<w:t[^>]*>|<\/w:t>/g, ''))
    .join('')
    .trim();
}

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
      formsDetected: false,
      flashingObjectsDetected: false,
      inlineContentFixed: false,
      inlineContentCount: 0,
      inlineContentDetails: [],
      colorContrastNeedsFixing: false,
      colorContrastIssues: []
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
    
    // Check for GIFs with location information
    const gifFiles = [];
    zip.forEach((relativePath, file) => {
      if (relativePath.startsWith('word/media/') && relativePath.toLowerCase().endsWith('.gif')) {
        gifFiles.push(relativePath);
      }
    });
    
    if (gifFiles.length > 0) {
      // Get location information for GIFs
      const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
      const gifLocations = analyzeGifLocations(documentXml, relsXml, gifFiles);
      
      report.details.gifsDetected = gifFiles;
      report.details.gifLocations = gifLocations;
      report.summary.flagged += gifFiles.length;
      console.log('[analyzeDocx] GIFs detected with locations, flagged count now:', report.summary.flagged);
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
      
      // Check for forms and form fields
      const formResults = analyzeForms(documentXml);
      console.log('[analyzeDocx] Form analysis results:', formResults.length, 'forms found');
      if (formResults.length > 0) {
        report.details.formsDetected = true;
        report.details.formLocations = formResults;
        report.summary.flagged += formResults.length;
        console.log('[analyzeDocx] Forms detected, flagged count now:', report.summary.flagged);
      } else {
        console.log('[analyzeDocx] No forms detected in document');
      }
      
      // Check for flashing/animated objects  
      const flashingResults = analyzeFlashingObjects(documentXml);
      if (flashingResults.length > 0) {
        report.details.flashingObjectsDetected = true;
        report.details.flashingObjectLocations = flashingResults;
        report.summary.flagged += flashingResults.length;
      }
      
      // Check for inline content positioning (images, objects, text boxes)
      const inlineContentResults = await fixInlineContentPositioning(zip, documentXml);
      if (inlineContentResults.fixed > 0) {
        report.details.inlineContentFixed = true;
        report.details.inlineContentCount = inlineContentResults.fixed;
        report.details.inlineContentDetails = inlineContentResults.details;
        report.details.inlineContentExplanation = "In-line content is necessary for users who rely on assistive technology and preferred by all users";
        report.details.inlineContentCategory = "Text Wrapping and Positioning";
        report.summary.fixed += 1; // Count as one consolidated fix
        console.log('[analyzeDocx] Inline content positioning fixed (consolidated from', inlineContentResults.fixed, 'individual fixes), fix count now:', report.summary.fixed);
      }
      
      // Check for color contrast issues with enhanced detection
      const colorContrastResults = await analyzeColorContrast(documentXml, zip);\n      \n      // Add fallback detection if no color issues found\n      const fallbackResults = detectGrayTextFallback(documentXml);\n      const allColorIssues = [...colorContrastResults, ...fallbackResults];
      if (colorContrastResults.length > 0) {
        report.details.colorContrastNeedsFixing = true;
        report.details.colorContrastIssues = colorContrastResults;
        report.summary.flagged += colorContrastResults.length;
        console.log('[analyzeDocx] Color contrast issues detected:', colorContrastResults.length, 'flagged count now:', report.summary.flagged);
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

// Analyze color contrast between text and background with enhanced detection
async function analyzeColorContrast(documentXml, zip) {
  const results = [];
  
  if (!documentXml) {
    return results;
  }

  // First, analyze styles.xml for color definitions
  const styleColorIssues = await analyzeStyleColors(zip);
  results.push(...styleColorIssues);

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

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

    // Extract text content to check if this paragraph has meaningful text
    const textContent = extractTextFromParagraph(paragraph);
    if (!textContent || textContent.trim().length === 0) {
      return; // Skip empty paragraphs
    }

    // Aggressive detection: Check for any potential gray formatting patterns
    const aggressiveColorIssues = detectPotentialGrayFormatting(paragraph, textContent, paragraphCount, approximatePageNumber, currentHeading);
    results.push(...aggressiveColorIssues);

    // Check for color specifications in runs within this paragraph
    const runMatches = paragraph.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) || [];
    
    runMatches.forEach(run => {
      const runText = extractTextFromParagraph(run);
      if (!runText || runText.trim().length === 0) {
        return; // Skip runs without text
      }

      // Look for color properties in run properties
      const colorIssues = analyzeRunColors(run, runText);
      
      if (colorIssues.length > 0) {
        colorIssues.forEach(issue => {
          results.push({
            type: issue.type,
            textColor: issue.textColor,
            backgroundColor: issue.backgroundColor,
            contrastRatio: issue.contrastRatio,
            requiredRatio: issue.requiredRatio,
            fontSize: issue.fontSize,
            location: `Paragraph ${paragraphCount}`,
            approximatePage: approximatePageNumber,
            context: currentHeading || 'Document body',
            preview: runText.substring(0, 100),
            recommendation: generateContrastRecommendation(issue)
          });
        });
      }
    });
  });

  return results;
}

// Analyze colors within a single text run
function analyzeRunColors(runXml, textContent) {
  const issues = [];
  
  // Extract run properties
  const rPrMatch = runXml.match(/<w:rPr[^>]*>([\s\S]*?)<\/w:rPr>/);
  if (!rPrMatch) {
    return issues; // No run properties, assume default colors
  }
  
  const runProperties = rPrMatch[1];
  
  // Extract text color - handle multiple formats
  let textColor = null;
  
  // Check for direct hex color
  const colorMatch = runProperties.match(/<w:color w:val="([^"]+)"\/>/); 
  if (colorMatch) {
    textColor = colorMatch[1];
  }
  
  // Check for theme colors (common in Word)
  const themeColorMatch = runProperties.match(/<w:color w:themeColor="([^"]+)"(?:\s+w:themeTint="([^"]+)")?(?:\s+w:themeShade="([^"]+)")?[^>]*\/>/);  
  if (themeColorMatch && !textColor) {
    const themeColor = themeColorMatch[1];
    const themeTint = themeColorMatch[2];
    const themeShade = themeColorMatch[3];
    
    // Convert common theme colors to hex approximations
    textColor = convertThemeColorToHex(themeColor, themeTint, themeShade);
  }
  
  // Handle auto color (usually black on white)
  if (!textColor || textColor === 'auto') {
    textColor = '000000'; // Assume black text
  }
  
  // Extract highlight/background color - handle multiple formats
  let backgroundColor = null;
  
  // Check for highlight colors
  const highlightMatch = runProperties.match(/<w:highlight w:val="([^"]+)"\/>/); 
  if (highlightMatch) {
    backgroundColor = convertHighlightColorToHex(highlightMatch[1]);
  }
  
  // Check for shading/fill colors
  const shadingMatch = runProperties.match(/<w:shd[^>]*w:fill="([^"]+)"[^>]*\/>/);  
  if (shadingMatch && !backgroundColor) {
    backgroundColor = shadingMatch[1];
  }
  
  // Check for theme background colors
  const themeShadingMatch = runProperties.match(/<w:shd[^>]*w:themeFill="([^"]+)"(?:\s+w:themeFillTint="([^"]+)")?[^>]*\/>/);  
  if (themeShadingMatch && !backgroundColor) {
    backgroundColor = convertThemeColorToHex(themeShadingMatch[1], themeShadingMatch[2]);
  }
  
  // Default to white background if not specified
  if (!backgroundColor || backgroundColor === 'auto') {
    backgroundColor = 'FFFFFF';
  }
  
  // Extract font size for contrast requirements
  let fontSize = 12; // Default assumption
  const sizeMatch = runProperties.match(/<w:sz w:val="(\d+)"\/>/); 
  if (sizeMatch) {
    fontSize = parseInt(sizeMatch[1]) / 2; // Word stores half-points
  }
  
  // Analyze color contrast - we now have more comprehensive color detection  
  if (textColor && backgroundColor) {
    const contrastAnalysis = calculateColorContrast(textColor, backgroundColor, fontSize);
    
    if (!contrastAnalysis.meetsRequirements) {
      issues.push({
        type: contrastAnalysis.type,
        textColor: textColor,
        backgroundColor: backgroundColor,
        contrastRatio: contrastAnalysis.ratio,
        requiredRatio: contrastAnalysis.requiredRatio,
        fontSize: fontSize
      });
    }
  }
  // Always check text color against white background since we default to white
  else if (textColor && textColor !== '000000') { // Skip black text on white (good contrast)
    // Flag potentially problematic colors against assumed white background
    const contrastAnalysis = calculateColorContrast(textColor, 'FFFFFF', fontSize);
    
    if (!contrastAnalysis.meetsRequirements) {
      issues.push({
        type: 'insufficient-contrast-assumed-background',
        textColor: textColor,
        backgroundColor: 'FFFFFF (assumed)',
        contrastRatio: contrastAnalysis.ratio,
        requiredRatio: contrastAnalysis.requiredRatio,
        fontSize: fontSize
      });
    }
  }
  
  return issues;
}

// Calculate color contrast ratio and check requirements
function calculateColorContrast(textColorHex, backgroundColorHex, fontSize) {
  // Remove # if present and ensure 6 digits
  textColorHex = textColorHex.replace('#', '').toUpperCase();
  backgroundColorHex = backgroundColorHex.replace('#', '').toUpperCase();
  
  if (textColorHex.length === 3) {
    textColorHex = textColorHex.split('').map(c => c + c).join('');
  }
  if (backgroundColorHex.length === 3) {
    backgroundColorHex = backgroundColorHex.split('').map(c => c + c).join('');
  }
  
  // Convert hex to RGB
  const textRgb = hexToRgb(textColorHex);
  const backgroundRgb = hexToRgb(backgroundColorHex);
  
  if (!textRgb || !backgroundRgb) {
    return {
      ratio: 0,
      meetsRequirements: false,
      requiredRatio: 4.5,
      type: 'invalid-color'
    };
  }
  
  // Calculate relative luminance
  const textLuminance = getRelativeLuminance(textRgb);
  const backgroundLuminance = getRelativeLuminance(backgroundRgb);
  
  // Calculate contrast ratio
  const lighter = Math.max(textLuminance, backgroundLuminance);
  const darker = Math.min(textLuminance, backgroundLuminance);
  const contrastRatio = (lighter + 0.05) / (darker + 0.05);
  
  // Determine requirements based on font size
  let requiredRatio = 4.5; // Default for normal text
  let type = 'insufficient-contrast-normal';
  
  if (fontSize >= 18 || (fontSize >= 14 && /* assume bold for larger fonts */ true)) {
    requiredRatio = 3.0; // Large text requirements
    type = 'insufficient-contrast-large';
  }
  
  return {
    ratio: Math.round(contrastRatio * 100) / 100,
    meetsRequirements: contrastRatio >= requiredRatio,
    requiredRatio: requiredRatio,
    type: contrastRatio >= requiredRatio ? 'sufficient-contrast' : type
  };
}

// Convert hex color to RGB
function hexToRgb(hex) {
  if (hex.length !== 6) return null;
  
  const r = parseInt(hex.substr(0, 2), 16);
  const g = parseInt(hex.substr(2, 2), 16);
  const b = parseInt(hex.substr(4, 2), 16);
  
  if (isNaN(r) || isNaN(g) || isNaN(b)) return null;
  
  return { r, g, b };
}

// Calculate relative luminance according to WCAG guidelines
function getRelativeLuminance(rgb) {
  const { r, g, b } = rgb;
  
  // Convert to 0-1 range
  const rs = r / 255;
  const gs = g / 255;
  const bs = b / 255;
  
  // Apply gamma correction
  const rLin = rs <= 0.03928 ? rs / 12.92 : Math.pow((rs + 0.055) / 1.055, 2.4);
  const gLin = gs <= 0.03928 ? gs / 12.92 : Math.pow((gs + 0.055) / 1.055, 2.4);
  const bLin = bs <= 0.03928 ? bs / 12.92 : Math.pow((bs + 0.055) / 1.055, 2.4);
  
  // Calculate luminance
  return 0.2126 * rLin + 0.7152 * gLin + 0.0722 * bLin;
}

// Convert Word theme colors to approximate hex values
function convertThemeColorToHex(themeColor, tint, shade) {
  const themeColors = {
    'text1': '000000',     // Black
    'text2': '1F497D',     // Dark blue
    'background1': 'FFFFFF', // White  
    'background2': 'EEECE1', // Light gray
    'accent1': '4F81BD',   // Blue
    'accent2': 'F79646',   // Orange
    'accent3': '9BBB59',   // Green
    'accent4': '8064A2',   // Purple
    'accent5': '4BACC6',   // Light blue
    'accent6': 'F24B4B',   // Red
    'dark1': '000000',     // Black
    'light1': 'FFFFFF',    // White
    'dark2': '1F497D',     // Dark blue
    'light2': 'EEECE1'    // Light gray
  };
  
  let baseColor = themeColors[themeColor] || '000000';
  
  // Apply tint/shade modifications (simplified)
  if (tint) {
    // Tint makes color lighter (mix with white)
    const tintValue = parseInt(tint, 16) / 255;
    baseColor = applyTint(baseColor, tintValue);
  }
  
  if (shade) {
    // Shade makes color darker (mix with black)  
    const shadeValue = parseInt(shade, 16) / 255;
    baseColor = applyShade(baseColor, shadeValue);
  }
  
  return baseColor;
}

// Convert Word highlight colors to hex values
function convertHighlightColorToHex(highlightColor) {
  const highlights = {
    'yellow': 'FFFF00',
    'brightGreen': '00FF00',
    'turquoise': '00FFFF', 
    'pink': 'FF00FF',
    'blue': '0000FF',
    'red': 'FF0000',
    'darkBlue': '000080',
    'darkRed': '800000',
    'darkYellow': '808000',
    'darkGreen': '008000',
    'darkMagenta': '800080',
    'darkCyan': '008080',
    'gray25': 'C0C0C0',
    'gray50': '808080',
    'lightGray': 'D3D3D3',
    'white': 'FFFFFF',
    'black': '000000'
  };
  
  return highlights[highlightColor] || 'FFFF00'; // Default to yellow
}

// Apply tint to a color (mix with white)
function applyTint(hexColor, tintFactor) {
  const rgb = hexToRgb(hexColor);
  if (!rgb) return hexColor;
  
  const r = Math.round(rgb.r + (255 - rgb.r) * tintFactor);
  const g = Math.round(rgb.g + (255 - rgb.g) * tintFactor);
  const b = Math.round(rgb.b + (255 - rgb.b) * tintFactor);
  
  return rgbToHex(r, g, b);
}

// Apply shade to a color (mix with black)
function applyShade(hexColor, shadeFactor) {
  const rgb = hexToRgb(hexColor);
  if (!rgb) return hexColor;
  
  const r = Math.round(rgb.r * (1 - shadeFactor));
  const g = Math.round(rgb.g * (1 - shadeFactor));
  const b = Math.round(rgb.b * (1 - shadeFactor));
  
  return rgbToHex(r, g, b);
}

// Convert RGB to hex
function rgbToHex(r, g, b) {
  return ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

// Generate specific recommendations based on contrast issue
function generateContrastRecommendation(issue) {
  const { type, contrastRatio, requiredRatio, fontSize } = issue;
  
  switch (type) {
    case 'insufficient-contrast-normal':
      return `Increase contrast to at least ${requiredRatio}:1 for ${fontSize}pt text. Current ratio: ${contrastRatio}:1. Consider darker text or lighter background.`;
    
    case 'insufficient-contrast-large':
      return `Increase contrast to at least ${requiredRatio}:1 for large text (${fontSize}pt). Current ratio: ${contrastRatio}:1.`;
    
    case 'insufficient-contrast-assumed-background':
      return `Verify contrast against actual background. Against white background: ${contrastRatio}:1 (needs ${requiredRatio}:1). Consider specifying background color.`;
    
    case 'invalid-color':
      return 'Invalid color specification detected. Use valid hex colors for proper contrast calculation.';
    
    default:
      return `Improve color contrast to meet WCAG guidelines. Required: ${requiredRatio}:1, Current: ${contrastRatio}:1.`;
  }
}

// Fix inline content positioning for images, objects and text boxes
async function fixInlineContentPositioning(zip, documentXml) {
  const results = {
    fixed: 0,
    details: [],
    explanation: "In-line content is necessary for users who rely on assistive technology and preferred by all users",
    category: "Text Wrapping and Positioning"
  };

  if (!documentXml) {
    return results;
  }

  let modifiedDocumentXml = documentXml;
  let changesMade = false;

  // Patterns to fix floating elements and positioning issues
  const floatingPatterns = [
    // DrawingML anchor patterns (modern Word drawings)
    {
      pattern: /<wp:anchor[^>]*>([\s\S]*?)<\/wp:anchor>/g,
      replacement: function(match, content) {
        // Convert anchor (floating) to inline
        return `<wp:inline>${content}</wp:inline>`;
      },
      type: 'text-wrapping-anchor',
      description: 'Converted floating anchor to inline'
    },
    // Text wrapping patterns
    {
      pattern: /<wp:wrapSquare[^>]*\/>/g,
      replacement: '',
      type: 'text-wrapping-square',
      description: 'Removed square text wrapping'
    },
    {
      pattern: /<wp:wrapTight[^>]*>[\s\S]*?<\/wp:wrapTight>/g,
      replacement: '',
      type: 'text-wrapping-tight',
      description: 'Removed tight text wrapping'
    },
    {
      pattern: /<wp:wrapThrough[^>]*>[\s\S]*?<\/wp:wrapThrough>/g,
      replacement: '',
      type: 'text-wrapping-through',
      description: 'Removed through text wrapping'
    },
    {
      pattern: /<wp:wrapTopAndBottom[^>]*\/>/g,
      replacement: '',
      type: 'text-wrapping-topbottom',
      description: 'Removed top and bottom text wrapping'
    },
    {
      pattern: /<wp:wrapNone[^>]*\/>/g,
      replacement: '',
      type: 'text-wrapping-none',
      description: 'Removed no-wrap positioning'
    },
    // Position and alignment patterns
    {
      pattern: /<wp:positionH[^>]*>[\s\S]*?<\/wp:positionH>/g,
      replacement: '',
      type: 'position-horizontal',
      description: 'Removed horizontal positioning'
    },
    {
      pattern: /<wp:positionV[^>]*>[\s\S]*?<\/wp:positionV>/g,
      replacement: '',
      type: 'position-vertical',
      description: 'Removed vertical positioning'
    },
    // VML patterns for legacy compatibility
    {
      pattern: /mso-position-horizontal:[^;]*;?/g,
      replacement: '',
      type: 'vml-horizontal-position',
      description: 'Removed VML horizontal positioning'
    },
    {
      pattern: /mso-position-vertical:[^;]*;?/g,
      replacement: '',
      type: 'vml-vertical-align',
      description: 'Removed VML vertical positioning'
    },
    {
      pattern: /mso-wrap-style:[^;]*;?/g,
      replacement: '',
      type: 'vml-wrap-style',
      description: 'Removed VML wrap style'
    },
    // Remove left/top positioning from style attributes
    {
      pattern: /left:\s*[^;]*;?/g,
      replacement: '',
      type: 'vml-left-position',
      description: 'Removed VML left positioning'
    },
    {
      pattern: /top:\s*[^;]*;?/g,
      replacement: '',
      type: 'vml-top-position',
      description: 'Removed VML top positioning'
    }
  ];

  // Apply fixes for floating elements
  floatingPatterns.forEach(patternObj => {
    const { pattern, replacement, type, description } = patternObj;
    const matches = modifiedDocumentXml.match(pattern);
    
    if (matches) {
      matches.forEach(match => {
        const newContent = typeof replacement === 'function' ? replacement(match) : replacement;
        modifiedDocumentXml = modifiedDocumentXml.replace(match, newContent);
        results.fixed++;
        changesMade = true;
        
        results.details.push({
          type: type,
          description: description,
          location: 'Document content',
          originalContent: match.substring(0, 100) + (match.length > 100 ? '...' : '')
        });
        
        console.log(`[fixInlineContentPositioning] Fixed ${type}: ${description}`);
      });
    }
  });

  // Special handling for drawing elements - ensure they are inline
  const drawingPattern = /<w:drawing[^>]*>[\s\S]*?<\/w:drawing>/g;
  const drawingMatches = modifiedDocumentXml.match(drawingPattern);
  
  if (drawingMatches) {
    drawingMatches.forEach(drawing => {
      // Check if this drawing contains floating elements
      if (drawing.includes('wp:anchor') && !drawing.includes('wp:inline')) {
        // Convert anchor to inline within the drawing
        let fixedDrawing = drawing.replace(/<wp:anchor[^>]*>/g, '<wp:inline>');
        fixedDrawing = fixedDrawing.replace(/<\/wp:anchor>/g, '</wp:inline>');
        
        if (fixedDrawing !== drawing) {
          modifiedDocumentXml = modifiedDocumentXml.replace(drawing, fixedDrawing);
          results.fixed++;
          changesMade = true;
          
          results.details.push({
            type: 'drawing-anchor-conversion',
            description: 'Converted floating drawing to inline',
            location: 'Document content',
            originalContent: drawing.substring(0, 100) + '...'
          });
          
          console.log('[fixInlineContentPositioning] Fixed drawing: Converted floating to inline');
        }
      }
    });
  }

  // Update the document XML in the zip if changes were made
  if (changesMade) {
    try {
      zip.file('word/document.xml', modifiedDocumentXml);
      console.log(`[fixInlineContentPositioning] Updated document.xml with ${results.fixed} inline content fixes`);
    } catch (error) {
      console.error('[fixInlineContentPositioning] Error updating document:', error);
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

// Analyze forms and form fields in the document
function analyzeForms(documentXml) {
  const results = [];

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;
  
  // Track unique form field locations to prevent duplicates
  const seenFormLocations = new Set();

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

    // Check for form fields and form controls
    const formElements = [
      /<w:fldSimple[^>]*FORMTEXT/,                  // Text form fields
      /<w:fldSimple[^>]*FORMCHECKBOX/,              // Checkbox form fields  
      /<w:fldSimple[^>]*FORMDROPDOWN/,              // Dropdown form fields
      /<w:ffData[\s\S]*?<\/w:ffData>/,              // Form field data (complete tags)
      /<w:ffData>/,                                 // Form field data (opening tag)
      /<w:checkBox/,                                // Checkbox controls
      /<w:dropDownList/,                            // Dropdown list controls
      /<w:textInput/,                               // Text input controls
      /<w:sdt>/,                                    // Structured document tags (content controls)
      /<w:sdtContent>/,                             // Content control content
      /<w:fldChar w:fldCharType="begin"\/>/,        // Field character begin
      /FORMTEXT/,                                   // Simple FORMTEXT detection
      /FORMCHECKBOX/,                               // Simple FORMCHECKBOX detection
      /FORMDROPDOWN/                                // Simple FORMDROPDOWN detection
    ];

    // Check for any form field in this paragraph and avoid duplicates
    let formDetectedInParagraph = false;
    let bestFormType = null;
    let detectedPatterns = [];

    formElements.forEach((regex, formIndex) => {
      const matches = paragraph.match(regex);
      if (matches) {
        formDetectedInParagraph = true;
        const formType = getFormType(formIndex);
        detectedPatterns.push(formType);
        
        // Prioritize more specific form types over generic ones
        if (!bestFormType || isPriorityFormType(formType, bestFormType)) {
          bestFormType = formType;
        }
      }
    });

    // Only add one form detection per paragraph to avoid duplicates
    if (formDetectedInParagraph) {
      const locationKey = `${paragraphCount}-${approximatePageNumber}`;
      
      if (!seenFormLocations.has(locationKey)) {
        seenFormLocations.add(locationKey);
        
        results.push({
          type: bestFormType,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150),
          recommendation: 'Consider using alternative formats like accessible web forms or structured tables instead of Word form fields',
          detectedPatterns: detectedPatterns // Debug info showing all patterns that matched
        });
      }
    }
  });

  return results;
}

// Helper function to identify form type
function getFormType(formIndex) {
  const formTypes = [
    'text-field', 'checkbox-field', 'dropdown-field', 'form-data-complete',
    'form-data', 'checkbox-control', 'dropdown-control', 'text-input', 
    'content-control', 'content-control-data', 'field-character',
    'formtext-simple', 'formcheckbox-simple', 'formdropdown-simple'
  ];
  return formTypes[formIndex] || 'form-element';
}

// Helper function to prioritize more specific form types over generic ones
function isPriorityFormType(newType, currentType) {
  // Priority order: specific form types > generic types
  const priorityOrder = {
    'form-data-complete': 10,
    'text-field': 9,
    'checkbox-field': 9,
    'dropdown-field': 9,
    'checkbox-control': 8,
    'dropdown-control': 8,
    'text-input': 8,
    'form-data': 7,
    'content-control': 6,
    'content-control-data': 5,
    'field-character': 4,
    'formtext-simple': 3,
    'formcheckbox-simple': 3,
    'formdropdown-simple': 3,
    'form-element': 1
  };
  
  return (priorityOrder[newType] || 1) > (priorityOrder[currentType] || 1);
}

// Helper function to identify flashing content type
function getFlashingType(flashIndex) {
  const flashTypes = [
    'color-animation', 'rotation-animation', 'scale-animation', 
    'motion-animation', 'generic-animation', 'effect-animation',
    'timing-element', 'looping-video', 'looping-audio'
  ];
  return flashTypes[flashIndex] || 'animated-content';
}

// Analyze flashing objects in the document
function analyzeFlashingObjects(documentXml) {
  const results = [];
  
  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Check for potentially flashing/animated content
  const flashingElements = [
    /<w:drawing>[\s\S]*?<a:animClr/,              // Color animations
    /<w:drawing>[\s\S]*?<a:animRot/,              // Rotation animations  
    /<w:drawing>[\s\S]*?<a:animScale/,            // Scale animations
    /<w:drawing>[\s\S]*?<a:animMotion/,           // Motion animations
    /<w:drawing>[\s\S]*?<a:animate/,              // Generic animations
    /<w:drawing>[\s\S]*?<a:animEffect/,           // Effect animations
    /<p:timing>/,                                 // PowerPoint timing elements (embedded)
    /<a:videoFile[^>]*loop/,                      // Looping videos
    /<a:audioFile[^>]*loop/                       // Looping audio
  ];

  // Split into paragraphs for analysis
  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    // Track page numbers (estimate)
    if (paragraphCount % 15 === 0) {
      approximatePageNumber++;
    }

    // Track headings for context
    if (/<w:pStyle w:val="Heading/.test(paragraph)) {
      currentHeading = extractTextFromParagraph(paragraph);
    }

    flashingElements.forEach((regex, flashIndex) => {
      if (regex.test(paragraph)) {
        const flashType = getFlashingType(flashIndex);
        results.push({
          type: flashType,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150) || 'Animated content detected',
          recommendation: 'Remove animated/flashing content to prevent seizures and improve accessibility for all users'
        });
      }
    });
  });

  return results;
}

// Analyze link descriptiveness in the document
function analyzeLinkDescriptiveness(documentXml) {
  const results = {
    nonDescriptiveLinks: []
  };

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;
  
  // Track unique link texts to prevent duplicates
  const seenLinkTexts = new Set();

  // Generic/non-descriptive phrases to detect
  const genericPhrases = [
    'click here', 'here', 'read more', 'more', 'link', 'this link',
    'see more', 'learn more', 'find out more', 'more info', 'more information',
    'view more', 'details', 'continue', 'next', 'go', 'visit', 'see',
    'download', 'open', 'access', 'view', 'show', 'display', 'get',
    'this', 'that', 'these', 'those', 'it', 'page', 'site', 'website',
    'url', 'address', 'location'
  ];

  // Additional patterns that are commonly non-descriptive
  const genericPatterns = [
    /^click\s+/i,           // "click this", "click the", etc.
    /\bclick\s+\w+\s*:?\s*$/i, // "click this link:", "click button:", etc.
    /^(here|there)\s*:?\s*$/i,  // "here:", "there:"
    /^(this|that)\s+link\s*:?\s*$/i, // "this link:", "that link:"
    /^read\s+(more|on)\s*:?\s*$/i,   // "read more:", "read on:"
    /^see\s+(more|here|this)\s*:?\s*$/i, // "see more:", "see here:", etc.
    /^(more|info|information)\s*:?\s*$/i, // "more:", "info:", etc.
    /^(download|view|open)\s*:?\s*$/i     // "download:", "view:", etc.
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
        // Check if the link text is non-descriptive using exact phrases
        const isGeneric = genericPhrases.some(phrase => {
          // Exact match or the link text is just the generic phrase
          return linkText === phrase || 
                 linkText === phrase + '.' || 
                 linkText === phrase + '!' ||
                 linkText.startsWith(phrase + ' ') ||
                 linkText.endsWith(' ' + phrase) ||
                 (linkText.length <= 15 && linkText.includes(phrase)); // Short phrases containing generic words
        });

        // Check if the link text matches generic patterns
        const isGenericPattern = genericPatterns.some(pattern => pattern.test(linkText));
        
        // Also flag very short links (likely non-descriptive)
        const isTooShort = linkText.length <= 3 && !linkText.match(/^[a-z]{2,3}$/); // Allow abbreviations like "FAQ", "PDF"
        
        // Flag URLs or email addresses used as link text
        const isUrl = linkText.includes('www.') || linkText.includes('http') || linkText.includes('.com') || linkText.includes('.org');
        const isEmail = linkText.includes('@') && linkText.includes('.');
        
        if (isGeneric || isGenericPattern || isTooShort || isUrl || isEmail) {
          // Check if we've already seen this exact link text
          if (!seenLinkTexts.has(linkText)) {
            seenLinkTexts.add(linkText);
            
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

// Analyze GIF locations in the document
function analyzeGifLocations(documentXml, relsXml, gifFiles) {
  const results = [];
  
  if (!relsXml || !gifFiles.length) {
    return results;
  }

  // Create mapping of relationship IDs to GIF files
  const gifRelationships = new Map();
  gifFiles.forEach(gifPath => {
    // Extract filename from path (e.g., "word/media/image1.gif" -> "image1.gif")
    const fileName = gifPath.split('/').pop();
    
    // Find relationship ID for this GIF in rels XML - try multiple patterns
    const patterns = [
      new RegExp(`<Relationship[^>]*Target="media/${fileName.replace('.', '\\.')}"[^>]*Id="([^"]*)"`, 'i'),
      new RegExp(`<Relationship[^>]*Id="([^"]*)"[^>]*Target="media/${fileName.replace('.', '\\.')}"`, 'i'),
      new RegExp(`Id="([^"]*)"[^>]*Target="[^"]*${fileName.replace('.', '\\.')}"`, 'i')
    ];
    
    for (const pattern of patterns) {
      const relMatch = relsXml.match(pattern);
      if (relMatch) {
        gifRelationships.set(relMatch[1], {
          file: gifPath,
          fileName: fileName
        });
        console.log(`[analyzeGifLocations] Found GIF relationship: ${relMatch[1]} -> ${fileName}`);
        break;
      }
    }
  });

  if (gifRelationships.size === 0) {
    return results;
  }

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Split document into paragraphs
  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    // Track page numbers (estimate)
    if (paragraphCount % 15 === 0) {
      approximatePageNumber++;
    }

    // Track headings for context
    if (/<w:pStyle w:val="Heading/.test(paragraph)) {
      currentHeading = extractTextFromParagraph(paragraph);
    }

    // Check if this paragraph contains any GIF references
    gifRelationships.forEach((gifInfo, relationshipId) => {
      // Look for drawing elements that reference this GIF - try multiple patterns
      const patterns = [
        new RegExp(`<w:drawing[\\s\\S]*?r:embed="${relationshipId}"[\\s\\S]*?</w:drawing>`, 'i'),
        new RegExp(`<a:blip[^>]*r:embed="${relationshipId}"`, 'i'),
        new RegExp(`r:embed="${relationshipId}"`, 'i'), // Simple embed reference
        new RegExp(`<w:drawing[\\s\\S]*?${relationshipId}[\\s\\S]*?</w:drawing>`, 'i') // Broader match
      ];
      
      let foundMatch = false;
      for (const pattern of patterns) {
        if (pattern.test(paragraph)) {
          foundMatch = true;
          break;
        }
      }
      
      if (foundMatch) {
        console.log(`[analyzeGifLocations] Found GIF in paragraph ${paragraphCount}: ${gifInfo.fileName}`);
        results.push({
          type: 'animated-gif',
          file: gifInfo.file,
          fileName: gifInfo.fileName,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150) || 'GIF image detected',
          recommendation: 'Replace animated GIFs with static images or accessible alternatives to prevent seizures and improve accessibility for users with vestibular disorders'
        });
      }
    });
  });

  return results;
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
// Add this function to the end of api/upload-document.js

// Fallback detection for gray text - aggressive approach to catch missed cases
function detectGrayTextFallback(documentXml) {
  const results = [];
  
  if (!documentXml) return results;
  
  console.log('[detectGrayTextFallback] Starting aggressive color detection...');
  
  // Look for ANY color-related formatting in the document
  const colorPatterns = [
    /<w:color[^>]*w:val="([^"]+)"[^>]*\/>/g,
    /<w:color[^>]*w:themeColor="([^"]+)"[^>]*\/>/g,
    /<w:highlight[^>]*w:val="([^"]+)"[^>]*\/>/g,
    /<w:shd[^>]*w:fill="([^"]+)"[^>]*\/>/g
  ];
  
  let foundAnyColorFormatting = false;
  
  colorPatterns.forEach((pattern, patternIndex) => {
    const matches = documentXml.match(pattern);
    if (matches) {
      foundAnyColorFormatting = true;
      console.log(`[detectGrayTextFallback] Pattern ${patternIndex} found ${matches.length} matches`);
      
      matches.forEach((match, matchIndex) => {
        console.log(`[detectGrayTextFallback] Match ${matchIndex}:`, match);
        
        // Extract color value
        const colorMatch = match.match(/(?:w:val|w:themeColor|w:fill)="([^"]+)"/);
        if (colorMatch) {
          const colorValue = colorMatch[1];
          
          // Flag any non-default color for manual review
          if (colorValue !== '000000' && colorValue !== 'windowText' && colorValue.toLowerCase() !== 'black') {
            console.log(`[detectGrayTextFallback] Flagging color: ${colorValue}`);
            
            results.push({
              type: 'detected-color-formatting',
              textColor: colorValue,
              backgroundColor: 'FFFFFF (assumed white)',
              contrastRatio: 'Manual verification required',
              requiredRatio: 4.5,
              fontSize: 12,
              location: 'Document with color formatting',
              approximatePage: 1,
              context: 'Color formatting detected',
              preview: `Color value: ${colorValue}`,
              recommendation: `Color '${colorValue}' detected in document. Gray colors like 808080 (3.95:1), 999999 (2.85:1), CCCCCC (1.61:1) fail the 4.5:1 contrast requirement. Please verify this color meets accessibility standards.`
            });
          }
        }
      });
    }
  });
  
  console.log('[detectGrayTextFallback] Completed. Found', results.length, 'issues, foundAnyColorFormatting:', foundAnyColorFormatting);
  return results;
}