const JSZip = require('jszip');
const fs = require('fs');

// Import the function from the main file
const downloadDocument = fs.readFileSync('./api/download-document.js', 'utf8');
eval(downloadDocument.match(/function removeShadowsAndNormalizeFonts[\s\S]*?^}/m)[0]);

async function testLineSpacing() {
  try {
    // Test with a sample DOCX file
    const testFile = './tests/fixtures/test_problematic.docx';
    if (!fs.existsSync(testFile)) {
      console.log('Test file not found, creating sample XML for testing...');
      
      // Test with sample XML content that has spacing issues
      const sampleXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:spacing w:line="240" w:lineRule="exact"/>
      </w:pPr>
      <w:r>
        <w:t>This paragraph has tight spacing (240 = 1.0 line spacing)</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:spacing w:line="280"/>
      </w:pPr>
      <w:r>
        <w:t>This paragraph has 280 spacing (less than 1.5)</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
      </w:pPr>
      <w:r>
        <w:t>This paragraph has no spacing defined</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
      
      console.log('Testing line spacing function with sample XML...');
      console.log('Original XML sample:', sampleXml.substring(0, 200) + '...');
      
      const result = removeShadowsAndNormalizeFonts(sampleXml);
      if (result) {
        console.log('\n✅ Changes detected! Function returned modified XML');
        console.log('Modified XML sample:', result.substring(0, 500) + '...');
        
        // Check specific changes
        if (result.includes('w:line="360"')) {
          console.log('✅ Line spacing was updated to 360 (1.5 spacing)');
        }
        if (result.includes('w:lineRule="auto"')) {
          console.log('✅ Line rule was changed to auto');
        }
      } else {
        console.log('❌ No changes detected - function returned null');
      }
      
      return;
    }
    
    // Test with actual DOCX file
    const fileData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(fileData);
    
    const docFile = zip.file('word/document.xml');
    if (docFile) {
      const originalXml = await docFile.async('string');
      console.log('Testing with actual DOCX file...');
      console.log('Original XML length:', originalXml.length);
      console.log('XML sample:', originalXml.substring(0, 500) + '...');
      
      const result = removeShadowsAndNormalizeFonts(originalXml);
      if (result) {
        console.log('\n✅ Changes detected in real file!');
        console.log('Modified XML length:', result.length);
        
        // Look for spacing patterns in original
        const originalSpacing = originalXml.match(/<w:spacing[^>]*>/g);
        const modifiedSpacing = result.match(/<w:spacing[^>]*>/g);
        
        console.log('Original spacing patterns:', originalSpacing?.slice(0, 3) || 'None found');
        console.log('Modified spacing patterns:', modifiedSpacing?.slice(0, 3) || 'None found');
      } else {
        console.log('❌ No changes detected in real file');
        
        // Analyze why no changes were made
        const hasSpacing = originalXml.includes('<w:spacing');
        const hasLowSpacing = /<w:spacing[^>]*w:line="([0-9]+)"/.test(originalXml);
        const hasExactSpacing = originalXml.includes('w:lineRule="exact"');
        
        console.log('Analysis:');
        console.log('- Has spacing elements:', hasSpacing);
        console.log('- Has line values:', hasLowSpacing);
        console.log('- Has exact spacing:', hasExactSpacing);
        
        if (hasSpacing) {
          const spacingMatches = originalXml.match(/<w:spacing[^>]*>/g);
          console.log('Spacing elements found:', spacingMatches?.slice(0, 3));
        }
      }
    }
    
  } catch (error) {
    console.error('Test failed:', error.message);
  }
}

testLineSpacing();