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
    
    // Helper function to write only if content changed
    const writeIfChanged = (filename, original, modified) => {
      if (original !== modified && modified !== null) {
        zip.file(filename, modified);
        return true;
      }
      return false;
    };
    
    // Process document.xml
    const docFile = zip.file('word/document.xml');
    if (docFile) {
      const origDocXml = await docFile.async('string');
      const afterShadows = removeShadowsAndNormalizeFonts(origDocXml);
      writeIfChanged('word/document.xml', origDocXml, afterShadows);
    }
    
    // Process styles.xml
    const stylesFile = zip.file('word/styles.xml');
    if (stylesFile) {
      const origStylesXml = await stylesFile.async('string');
      const afterStylesShadows = removeShadowsAndNormalizeFonts(origStylesXml);
      writeIfChanged('word/styles.xml', origStylesXml, afterStylesShadows);
    }
    
    // Process theme files
    const themeFile = zip.file('word/theme/theme1.xml');
    if (themeFile) {
      const origThemeXml = await themeFile.async('string');
      const afterTheme = removeShadowsAndNormalizeFonts(origThemeXml);
      writeIfChanged('word/theme/theme1.xml', origThemeXml, afterTheme);
    }
    
    // Protection removal
    try {
      const settingsFile = zip.file('word/settings.xml');
      if (settingsFile) {
        const origSettings = await settingsFile.async('string');
        const hasAnyProt = /<w:(?:documentProtection|writeProtection|readOnlyRecommended|editRestrictions|formProtection|protection|docProtection|enforcement|locked|trackRevisions|crypt)\b/.test(origSettings);
        if (hasAnyProt) {
          let cleaned = origSettings;
          
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
      console.warn('[remediateDocx] Protection removal failed:', e.message);
    }
    
    // Generate with proper compression
    const remediatedBuffer = await zip.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    return remediatedBuffer;
    
  } catch (error) {
    throw new Error(`Failed to remediate document: ${error.message}`);
  }
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
  
  // 5. Remove shadow properties and attributes (safely)
  // Remove only within attribute values, not entire element names
  fixedXml = fixedXml.replace(/\s+\w*shdw\w*\s*=\s*"[^"]*"/g, '');
  
  // 6. Normalize fonts to Arial (sans-serif) - surgically replace font names
  fixedXml = fixedXml.replace(/w:ascii="[^"]*"/g, 'w:ascii="Arial"');
  fixedXml = fixedXml.replace(/w:hAnsi="[^"]*"/g, 'w:hAnsi="Arial"');
  fixedXml = fixedXml.replace(/w:cs="[^"]*"/g, 'w:cs="Arial"');
  fixedXml = fixedXml.replace(/w:eastAsia="[^"]*"/g, 'w:eastAsia="Arial"');
  
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
  
  // 9. Fix line spacing to be at least 1.5 (360 twentieths of a point)
  // Replace any spacing that's less than 1.5 or uses exact spacing
  fixedXml = fixedXml.replace(
    /<w:spacing[^>]*w:line="(\d+)"[^>]*\/>/g,
    (match, lineValue) => {
      const currentSpacing = parseInt(lineValue);
      // If spacing is less than 360 (1.5), set it to 360
      if (currentSpacing < 360) {
        return '<w:spacing w:line="360" w:lineRule="auto"/>';
      }
      // If it has exact spacing, change to auto with minimum 1.5
      if (match.includes('w:lineRule="exact"')) {
        return '<w:spacing w:line="360" w:lineRule="auto"/>';
      }
      return match;
    }
  );
  
  // Also handle spacing elements without line values or with exact rule
  fixedXml = fixedXml.replace(
    /<w:spacing[^>]*w:lineRule="exact"[^>]*\/>/g,
    '<w:spacing w:line="360" w:lineRule="auto"/>'
  );
  
  // Add default spacing to paragraphs that don't have any spacing defined
  fixedXml = fixedXml.replace(
    /<w:pPr>(?![^<]*<w:spacing)/g,
    '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/>'
  );
  
  // Add paragraph properties with proper spacing to paragraphs that have no pPr at all
  fixedXml = fixedXml.replace(
    /<w:p([^>]*)>(?!\s*<w:pPr)/g,
    '<w:p$1><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
  );
  
  // If nothing changed, return null so callers can avoid rewriting the part
  if (fixedXml === original) return null;
  return fixedXml;
}
