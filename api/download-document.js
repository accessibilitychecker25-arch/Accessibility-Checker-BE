const Busboy = require('busboy');
const JSZip = require('jszip');

module.exports = async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', 'https://kmoreland126.github.io');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

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
        
        // Suggest better filename
        let suggestedName = filename;
        if (filename.includes('_') || filename.toLowerCase().startsWith('document') || filename.toLowerCase().startsWith('untitled')) {
          suggestedName = filename.replace(/_/g, '-').replace(/\.docx$/i, '-remediated.docx');
        } else {
          suggestedName = filename.replace(/\.docx$/i, '-remediated.docx');
        }

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
    
    // Fix 1: Set title if missing
    const coreXmlFile = zip.file('docProps/core.xml');
    if (coreXmlFile) {
      let coreXml = await coreXmlFile.async('string');
      if (coreXml.includes('<dc:title></dc:title>') || coreXml.includes('<dc:title/>')) {
        coreXml = coreXml.replace(/<dc:title><\/dc:title>|<dc:title\/>/, '<dc:title>Needs Title</dc:title>');
        zip.file('docProps/core.xml', coreXml);
      }
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
    
    const level = parseInt(headingMatch[1]);
    headingCount++;
    
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
    
    // Fix 1: Add placeholder text to empty headings
    if (!hasText) {
      // Check if there's a run element to add text to
      if (content.includes('<w:r>')) {
        // Add text element to existing run
        const fixedContent = content.replace(
          /(<w:r>(?:(?!<w:t>).)*?)(<\/w:r>)/,
          `$1<w:t>[Empty Heading ${level} - Please add text]</w:t>$2`
        );
        return `<w:p>${fixedContent}</w:p>`;
      } else {
        // Create a new run with text
        const newRun = '<w:r><w:t>[Empty Heading ' + level + ' - Please add text]</w:t></w:r>';
        const fixedContent = content.replace(
          /(<w:pPr>[\s\S]*?<\/w:pPr>)/,
          `$1${newRun}`
        );
        return `<w:p>${fixedContent}</w:p>`;
      }
    }
    
    // Fix 2: Add comment about heading order issues
    if (level > previousHeadingLevel + 1 && previousHeadingLevel > 0) {
      // Heading skips levels - add a warning comment in the text
      const warningText = `[WARNING: Heading ${level} follows Heading ${previousHeadingLevel} - Consider adding Heading ${previousHeadingLevel + 1}] `;
      
      // Insert warning at the beginning of the first text element
      const fixedContent = content.replace(
        /(<w:t[^>]*>)(.*?)(<\/w:t>)/,
        `$1${warningText}$2$3`
      );
      previousHeadingLevel = level;
      return `<w:p>${fixedContent}</w:p>`;
    }
    
    previousHeadingLevel = level;
    return match;
  });
  
  return fixedXml;
}
