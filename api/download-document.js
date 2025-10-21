const Busboy = require('busboy');
const JSZip = require('jszip');

module.exports = async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', 'https://kmoreland126.github.io');
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
      // Always update title to match the clean filename
      coreXml = coreXml.replace(
        /<dc:title>.*?<\/dc:title>|<dc:title\/>/,
        `<dc:title>${cleanTitle}</dc:title>`
      );
      zip.file('docProps/core.xml', coreXml);
    }
    
    // Fix 2: Set default language to en-US in styles
    const stylesXmlFile = zip.file('word/styles.xml');
    if (stylesXmlFile) {
      let stylesXml = await stylesXmlFile.async('string');
      
      // Check if language is already set
      if (!stylesXml.includes('w:lang w:val="en-US"')) {
        // Find docDefaults section and add language
        if (stylesXml.includes('<w:docDefaults>')) {
          stylesXml = stylesXml.replace(
            /<w:docDefaults>/,
            '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
          );
        }
        zip.file('word/styles.xml', stylesXml);
      }
    }
    
    // Fix 3: Fix empty headings and heading order issues
    const documentXmlFile = zip.file('word/document.xml');
    if (documentXmlFile) {
      let documentXml = await documentXmlFile.async('string');
      documentXml = fixHeadings(documentXml);
      zip.file('word/document.xml', documentXml);
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
