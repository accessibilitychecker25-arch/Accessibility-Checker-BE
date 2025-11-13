const fs = require('fs');
const JSZip = require('jszip');

async function testSimpleDetection() {
  console.log('=== Testing Simplified Detection ===\n');
  
  const testFile = 'reports/Protected_remediated_by_agent.docx';
  
  try {
    const fileData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(fileData);
    
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (documentXml) {
      const totalParagraphs = (documentXml.match(/<w:p[^>]*>/g) || []).length;
      const paragraphsWithSpacing = (documentXml.match(/<w:spacing w:line="360"/g) || []).length;
      
      console.log(`Total paragraphs: ${totalParagraphs}`);
      console.log(`Paragraphs with 1.5 spacing (360): ${paragraphsWithSpacing}`);
      console.log(`Needs line spacing fix: ${totalParagraphs > 0 && paragraphsWithSpacing === 0 ? 'YES' : 'NO'}`);
      
      // Also check for any spacing at all
      const anySpacing = (documentXml.match(/<w:spacing/g) || []).length;
      console.log(`Paragraphs with any spacing: ${anySpacing}`);
      
      // Test the shadow detection too
      const serifFonts = (documentXml.match(/(Times|Georgia|Garamond|serif)/gi) || []).length;
      const smallFonts = (documentXml.match(/<w:sz w:val="(\d+)"/g) || []).map(m => {
        const size = parseInt(m.match(/w:val="(\d+)"/)[1]);
        return size < 22;
      }).filter(Boolean).length;
      
      console.log(`\nOther checks:`);
      console.log(`Serif fonts found: ${serifFonts}`);
      console.log(`Small fonts found: ${smallFonts}`);
      
      // Check what the current logic would return
      const results = {
        hasShadows: false,
        hasSerifFonts: serifFonts > 0,
        hasSmallFonts: smallFonts > 0,
        hasInsufficientLineSpacing: totalParagraphs > 0 && paragraphsWithSpacing === 0
      };
      
      console.log(`\nFinal results:`);
      console.log(`hasShadows: ${results.hasShadows}`);
      console.log(`hasSerifFonts: ${results.hasSerifFonts}`);
      console.log(`hasSmallFonts: ${results.hasSmallFonts}`);
      console.log(`hasInsufficientLineSpacing: ${results.hasInsufficientLineSpacing}`);
    }
    
  } catch (error) {
    console.error('Test failed:', error.message);
  }
}

testSimpleDetection();