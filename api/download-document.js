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
    
    // Generate the remediated DOCX file
    const remediatedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    return remediatedBuffer;
    
  } catch (error) {
    throw new Error(`Failed to remediate document: ${error.message}`);
  }
}
