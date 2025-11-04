const Busboy = require('busboy');
const JSZip = require('jszip');

module.exports = async (req, res) => {
  // CORS: safe allowlist — echo back the requesting Origin when allowed.
  // This prevents returning a different origin value than the actual request origin
  // which the browser will reject.
  const ALLOWED_ORIGINS = [
    'https://accessibilitychecker25-arch.github.io',
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
        const remediatedFile = await remediateDocx(fileData, filename);
        
        // Always fix filename: replace underscores with hyphens and add -remediated suffix
        let suggestedName = filename
          .replace(/_/g, '-')  // Replace all underscores with hyphens
          .replace(/\.docx$/i, '-remediated.docx');  // Add -remediated before extension

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${suggestedName}"`);
        res.status(200).send(remediatedFile);
        
      } catch (error) {
        res.status(500).json({ error: error.message });
      }
    });

    req.pipe(busboy);

  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

async function remediateDocx(fileData, filename) {
  try {
    const zip = await JSZip.loadAsync(fileData);
    // helper: write to zip only if content actually changed
    const writeIfChanged = (name, originalStr, updatedStr) => {
      if (typeof updatedStr === 'string' && updatedStr !== originalStr) {
        zip.file(name, updatedStr);
      }
    };
    
    // Create a clean title from the filename (remove extension, replace special chars)
    const cleanTitle = filename
      .replace(/\.docx$/i, '')  // Remove extension
      .replace(/_/g, ' ')        // Replace underscores with spaces
      .replace(/-/g, ' ')        // Replace hyphens with spaces
      .replace(/\s+/g, ' ')      // Normalize multiple spaces
      .trim();
    
    // Fix 1: Set title based on filename
    const coreXmlFile = zip.file('docProps/core.xml');
    if (coreXmlFile) {
      let coreXml = await coreXmlFile.async('string');
      const newCore = coreXml.replace(
        /<dc:title>.*?<\/dc:title>|<dc:title\/>/,
        `<dc:title>${cleanTitle}</dc:title>`
      );
      writeIfChanged('docProps/core.xml', coreXml, newCore);
    }
    
    // Fix 2: Set default language to en-US in styles
    const stylesXmlFile = zip.file('word/styles.xml');
    if (stylesXmlFile) {
      let stylesXml = await stylesXmlFile.async('string');
      let newStyles = stylesXml;
      // Check if language is already set
      if (!stylesXml.includes('w:lang w:val="en-US"')) {
        // Find docDefaults section and add language
        if (stylesXml.includes('<w:docDefaults>')) {
          newStyles = newStyles.replace(
            /<w:docDefaults>/,
            '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
          );
        }
      }
      writeIfChanged('word/styles.xml', stylesXml, newStyles);
    }
    
    // Fix 3: Remove ALL document protection / write-protection / read-only recommendation and edit locks
    const settingsXmlFile = zip.file('word/settings.xml');
    if (settingsXmlFile) {
      let settingsXml = await settingsXmlFile.async('string');
      let newSettings = settingsXml;
      
      // Remove all known protection elements (comprehensive list for all Office variants)
      newSettings = newSettings.replace(/<w:(documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection)[^>]*\/>/g, '');
      newSettings = newSettings.replace(/<w:(documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection)[^>]*>[\s\S]*?<\/w:(?:documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection)>/g, '');
      
      // Remove locking elements and attributes
      newSettings = newSettings.replace(/<w:locked[^>]*\/>/g, '');
      newSettings = newSettings.replace(/\s?w:locked="[^"]*"/g, '');
      
      // Remove enforcement and password-related elements
      newSettings = newSettings.replace(/<w:enforcement[^>]*\/>/g, '');
      newSettings = newSettings.replace(/<w:enforcement[^>]*>[\s\S]*?<\/w:enforcement>/g, '');
      
      // Remove any hash/salt protection elements
      newSettings = newSettings.replace(/<w:cryptProviderType[^>]*\/>/g, '');
      newSettings = newSettings.replace(/<w:cryptAlgorithmClass[^>]*\/>/g, '');
      newSettings = newSettings.replace(/<w:cryptAlgorithmType[^>]*\/>/g, '');
      newSettings = newSettings.replace(/<w:cryptAlgorithmSid[^>]*\/>/g, '');
      newSettings = newSettings.replace(/<w:cryptSpinCount[^>]*\/>/g, '');
      
      // Ensure trackRevisions is disabled
      newSettings = newSettings.replace(/<w:trackRevisions[^>]*\/>/g, '');
      newSettings = newSettings.replace(/\s?w:trackRevisions="[^"]*"/g, '');
      
      writeIfChanged('word/settings.xml', settingsXml, newSettings);
    }
    
    // Fix 4: Remove text shadows and normalize fonts/sizes
    const documentXmlFile = zip.file('word/document.xml');
    if (documentXmlFile) {
      const origDocXml = await documentXmlFile.async('string');
      const afterHeadings = fixHeadings(origDocXml);
      const afterShadows = removeShadowsAndNormalizeFonts(afterHeadings);
      if (afterShadows !== null) {
        writeIfChanged('word/document.xml', origDocXml, afterShadows);
      } else if (afterHeadings !== origDocXml) {
        writeIfChanged('word/document.xml', origDocXml, afterHeadings);
      }
    }

    // Fix 5: Apply same shadow/font fixes to styles.xml
    if (stylesXmlFile) {
      const origStylesXml = await stylesXmlFile.async('string');
      const afterStylesShadows = removeShadowsAndNormalizeFonts(origStylesXml);
      if (afterStylesShadows !== null) {
        writeIfChanged('word/styles.xml', origStylesXml, afterStylesShadows);
      }
    }
    
    // Fix 6: Remove advanced shadows from theme files
    const themeFiles = Object.keys(zip.files).filter(name => name.includes('theme') && name.endsWith('.xml'));
    for (const themeFileName of themeFiles) {
      const themeFile = zip.file(themeFileName);
      if (themeFile) {
        const origThemeXml = await themeFile.async('string');
        const afterTheme = removeShadowsAndNormalizeFonts(origThemeXml);
        if (afterTheme !== null) {
          writeIfChanged(themeFileName, origThemeXml, afterTheme);
        }
      }
    }
    
    // Final safety pass: ensure settings.xml contains NO protection tags whatsoever
    try {
      const settingsFinalFile = zip.file('word/settings.xml');
      if (settingsFinalFile) {
        const origSettings = await settingsFinalFile.async('string');
        const hasAnyProt = /<w:(?:documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection|enforcement|locked|trackRevisions|crypt)\b/.test(origSettings);
        if (hasAnyProt) {
          console.log('[remediateDocx] found protection in settings.xml during final pass — removing ALL protection');
          let cleaned = origSettings;
          
          // Remove all protection-related elements (comprehensive final cleanup)
          cleaned = cleaned.replace(/<w:(?:documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection)[^>]*\/>/g, '');
          cleaned = cleaned.replace(/<w:(?:documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection)[^>]*>[\s\S]*?<\/w:(?:documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection)>/g, '');
          cleaned = cleaned.replace(/<w:(?:enforcement|locked|trackRevisions)[^>]*\/>/g, '');
          cleaned = cleaned.replace(/<w:(?:enforcement|locked|trackRevisions)[^>]*>[\s\S]*?<\/w:(?:enforcement|locked|trackRevisions)>/g, '');
          cleaned = cleaned.replace(/<w:crypt[^>]*\/>/g, '');
          cleaned = cleaned.replace(/<w:crypt[^>]*>[\s\S]*?<\/w:crypt[^>]*>/g, '');
          cleaned = cleaned.replace(/\s?w:(?:locked|trackRevisions|enforcement)="[^"]*"/g, '');
          
          writeIfChanged('word/settings.xml', origSettings, cleaned);
        }
      }
    } catch (e) {
      console.warn('[remediateDocx] final protection-clean pass failed', e && e.stack ? e.stack : e);
    }

    // Also ensure document.xml has no section-level protection
    try {
      const docFile = zip.file('word/document.xml');
      if (docFile) {
        const origDoc = await docFile.async('string');
        const hasSectionProt = /<w:sectPrChange[^>]*>[\s\S]*?<w:sectPr[^>]*>[\s\S]*?<w:formProt[^>]*\/?>[\s\S]*?<\/w:sectPr>[\s\S]*?<\/w:sectPrChange>/g.test(origDoc) ||
                              /<w:sectPr[^>]*>[\s\S]*?<w:formProt[^>]*\/?>/g.test(origDoc);
        if (hasSectionProt) {
          console.log('[remediateDocx] removing section-level form protection from document.xml');
          let cleanedDoc = origDoc;
          cleanedDoc = cleanedDoc.replace(/<w:formProt[^>]*\/>/g, '');
          cleanedDoc = cleanedDoc.replace(/<w:formProt[^>]*>[\s\S]*?<\/w:formProt>/g, '');
          writeIfChanged('word/document.xml', origDoc, cleanedDoc);
        }
      }
    } catch (e) {
      console.warn('[remediateDocx] section protection cleanup failed', e && e.stack ? e.stack : e);
    }

    // Generate the remediated DOCX file
    const remediatedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    return remediatedBuffer;
    
  } catch (error) {
    throw new Error(`Failed to remediate document: ${error.message}`);
  }
}

function fixHeadings(documentXml) {
  // Find all paragraphs
  const paragraphRegex = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  
  let fixedXml = documentXml;
  let previousHeadingLevel = 0;
  let headingCount = 0;
  
  fixedXml = fixedXml.replace(paragraphRegex, (match, content) => {
    // Check if this paragraph has a heading style
    const headingMatch = content.match(/<w:pStyle w:val="Heading(\d+)"\/>/);
    
    if (!headingMatch) {
      return match; // Not a heading, leave it unchanged
    }
    
    let level = parseInt(headingMatch[1]);
    headingCount++;
    
    // Fix heading order FIRST before processing text
    let correctedLevel = level;
    let wasCorrect = false;
    if (level > previousHeadingLevel + 1 && previousHeadingLevel > 0) {
      // Heading skips levels - automatically correct to proper level
      correctedLevel = previousHeadingLevel + 1;
      wasCorrect = true;
      
      // Replace the heading style with the corrected level
      content = content.replace(
        /<w:pStyle w:val="Heading\d+"\/>/,
        `<w:pStyle w:val="Heading${correctedLevel}"/>`
      );
      
      // Update level to use corrected level for rest of processing
      level = correctedLevel;
    }
    
    // Extract text content
    const textMatches = content.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
    let hasText = false;
    
    if (textMatches) {
      const text = textMatches
        .map(t => t.replace(/<w:t[^>]*>|<\/w:t>/g, ''))
        .join('')
        .trim();
      hasText = text.length > 0;
    }
    
    // If we corrected the level, update the text to reflect it
    if (wasCorrect && hasText) {
      content = content.replace(
        /(<w:t[^>]*>)(.*?)(<\/w:t>)/,
        `$1[Corrected to H${level}] $2$3`
      );
    }
    
    // Fix empty headings (using corrected level)
    if (!hasText) {
      // Check if there's a run element to add text to
      if (content.includes('<w:r>')) {
        // Add text element to existing run
        const fixedContent = content.replace(
          /(<w:r>(?:(?!<w:t>).)*?)(<\/w:r>)/,
          `$1<w:t>[Empty Heading ${level} - Please add text]</w:t>$2`
        );
        previousHeadingLevel = level;
        return `<w:p>${fixedContent}</w:p>`;
      } else {
        // Create a new run with text
        const newRun = '<w:r><w:t>[Empty Heading ' + level + ' - Please add text]</w:t></w:r>';
        const fixedContent = content.replace(
          /(<w:pPr>[\s\S]*?<\/w:pPr>)/,
          `$1${newRun}`
        );
        previousHeadingLevel = level;
        return `<w:p>${fixedContent}</w:p>`;
      }
    }
    
    previousHeadingLevel = level;
    return `<w:p>${content}</w:p>`;
  });
  
  return fixedXml;
}

function removeShadowsAndNormalizeFonts(xmlContent) {
  const original = xmlContent;
  let fixedXml = xmlContent;
  
  // 1. Remove basic Word text shadows
  fixedXml = fixedXml.replace(/<w:shadow\s*\/>/g, '');
  fixedXml = fixedXml.replace(/<w:shadow[^>]*>.*?<\/w:shadow>/g, '');
  fixedXml = fixedXml.replace(/\s+\w*shadow\w*\s*=\s*"[^"]*"/g, '');
  
  // 2. Remove advanced DrawingML shadow effects
  fixedXml = fixedXml.replace(/<a:outerShdw[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<a:outerShdw[^>]*>.*?<\/a:outerShdw>/g, '');
  fixedXml = fixedXml.replace(/<a:innerShdw[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<a:innerShdw[^>]*>.*?<\/a:innerShdw>/g, '');
  fixedXml = fixedXml.replace(/<a:prstShdw[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<a:prstShdw[^>]*>.*?<\/a:prstShdw>/g, '');
  
  // 3. Remove Office 2010+ shadow effects
  fixedXml = fixedXml.replace(/<w14:shadow[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<w14:shadow[^>]*>.*?<\/w14:shadow>/g, '');
  fixedXml = fixedXml.replace(/<w15:shadow[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<w15:shadow[^>]*>.*?<\/w15:shadow>/g, '');
  
  // 4. Remove shadow-related text effects and 3D properties
  fixedXml = fixedXml.replace(/<w14:glow[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<w14:glow[^>]*>.*?<\/w14:glow>/g, '');
  fixedXml = fixedXml.replace(/<w14:reflection[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<w14:reflection[^>]*>.*?<\/w14:reflection>/g, '');
  fixedXml = fixedXml.replace(/<w14:props3d[^>]*\/>/g, '');
  fixedXml = fixedXml.replace(/<w14:props3d[^>]*>.*?<\/w14:props3d>/g, '');
  
  // 5. Remove shadow properties and attributes
  fixedXml = fixedXml.replace(/outerShdw/g, '');
  fixedXml = fixedXml.replace(/innerShdw/g, '');
  fixedXml = fixedXml.replace(/\s+\w*shdw\w*\s*=\s*"[^"]*"/g, '');
  
  // 6. Normalize fonts to Arial (sans-serif)
  fixedXml = fixedXml.replace(
    /<w:rFonts[^>]*\/?>/g,
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/>'
  );
  
  // 7. Ensure minimum font size of 22 half-points (11pt)
  fixedXml = fixedXml.replace(
    /<w:sz w:val="(\d+)"\s*\/>/g,
    (match, size) => {
      const sizeNum = parseInt(size);
      if (sizeNum < 22) {
        return '<w:sz w:val="22"/>';
      }
      return match;
    }
  );
  
  // 8. Same for complex script font sizes
  fixedXml = fixedXml.replace(
    /<w:szCs w:val="(\d+)"\s*\/>/g,
    (match, size) => {
      const sizeNum = parseInt(size);
      if (sizeNum < 22) {
        return '<w:szCs w:val="22"/>';
      }
      return match;
    }
  );
  
  // If nothing changed, return null so callers can avoid rewriting the part
  if (fixedXml === original) return null;
  return fixedXml;
}
