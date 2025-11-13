const JSZip = require('jszip');
const fs = require('fs');

// Test the line spacing detection logic
async function testLineSpacingDetection() {
  console.log('=== Testing Line Spacing Detection ===\n');
  
  // Test file path - use an existing document
  const testFile = 'reports/Protected_remediated_by_agent.docx';
  
  if (!fs.existsSync(testFile)) {
    console.log(`Test file ${testFile} not found. Skipping test.`);
    return;
  }
  
  try {
    const fileData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(fileData);
    
    // Simulate the analyzeShadowsAndFonts function logic for line spacing
    const results = {
      hasInsufficientLineSpacing: false
    };
    
    // Check document.xml for line spacing issues
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (documentXml) {
      console.log('Checking document.xml for line spacing issues...');
      
      // Check for insufficient line spacing (less than 1.5 = 360 twentieths of a point)
      const spacingMatches = documentXml.match(/<w:spacing[^>]*w:line="(\d+)"[^>]*\/>/g);
      if (spacingMatches) {
        console.log(`Found ${spacingMatches.length} explicit spacing declarations:`);
        for (const match of spacingMatches) {
          const lineValue = parseInt(match.match(/w:line="(\d+)"/)[1]);
          console.log(`  - Line spacing: ${lineValue} (${lineValue >= 360 ? 'GOOD' : 'NEEDS FIX'})`);
          if (lineValue < 360) {
            results.hasInsufficientLineSpacing = true;
          }
        }
      } else {
        console.log('No explicit spacing declarations found');
      }
      
      // Check for exact line spacing (should be auto for accessibility)
      if (!results.hasInsufficientLineSpacing && documentXml.includes('w:lineRule="exact"')) {
        console.log('Found exact line spacing - NEEDS FIX');
        results.hasInsufficientLineSpacing = true;
      }
      
      // Check for paragraphs without any line spacing
      if (!results.hasInsufficientLineSpacing) {
        const paragraphsWithoutSpacing = documentXml.match(/<w:p[^>]*>(?![^<]*<w:pPr[^>]*>[^<]*<w:spacing)/g);
        const paragraphsWithoutPPr = documentXml.match(/<w:p[^>]*>(?!\s*<w:pPr)/g);
        
        if (paragraphsWithoutSpacing && paragraphsWithoutSpacing.length > 0) {
          console.log(`Found ${paragraphsWithoutSpacing.length} paragraphs without spacing properties - NEEDS FIX`);
          results.hasInsufficientLineSpacing = true;
        }
        
        if (paragraphsWithoutPPr && paragraphsWithoutPPr.length > 0) {
          console.log(`Found ${paragraphsWithoutPPr.length} paragraphs without paragraph properties - NEEDS FIX`);
          results.hasInsufficientLineSpacing = true;
        }
      }
    }
    
    // Check styles.xml
    const stylesXml = await zip.file('word/styles.xml')?.async('string');
    if (stylesXml && !results.hasInsufficientLineSpacing) {
      console.log('\nChecking styles.xml for line spacing issues...');
      
      const spacingMatches = stylesXml.match(/<w:spacing[^>]*w:line="(\d+)"[^>]*\/>/g);
      if (spacingMatches) {
        console.log(`Found ${spacingMatches.length} style spacing declarations:`);
        for (const match of spacingMatches) {
          const lineValue = parseInt(match.match(/w:line="(\d+)"/)[1]);
          console.log(`  - Style line spacing: ${lineValue} (${lineValue >= 360 ? 'GOOD' : 'NEEDS FIX'})`);
          if (lineValue < 360) {
            results.hasInsufficientLineSpacing = true;
          }
        }
      }
      
      if (!results.hasInsufficientLineSpacing && stylesXml.includes('w:lineRule="exact"')) {
        console.log('Found exact line spacing in styles - NEEDS FIX');
        results.hasInsufficientLineSpacing = true;
      }
    }
    
    console.log('\n=== RESULTS ===');
    console.log(`Line spacing needs fixing: ${results.hasInsufficientLineSpacing ? 'YES' : 'NO'}`);
    
    if (results.hasInsufficientLineSpacing) {
      console.log('✅ Detection working - line spacing issues found and will be reported as fixed');
    } else {
      console.log('ℹ️  No line spacing issues detected in this document');
    }
    
  } catch (error) {
    console.error('Test failed:', error.message);
  }
}

testLineSpacingDetection();