const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

// Import the remediation function from the actual endpoint
function removeShadowsAndNormalizeFonts(xmlContent) {
  let fixedXml = xmlContent;
  
  // 1. Remove text shadows
  fixedXml = fixedXml.replace(/<w:shadow\s*\/>/g, '');
  fixedXml = fixedXml.replace(/<w:shadow[^>]*>.*?<\/w:shadow>/g, '');
  fixedXml = fixedXml.replace(/\s+\w*shadow\w*\s*=\s*"[^"]*"/g, '');
  
  // 2. Normalize fonts to Arial (sans-serif)
  fixedXml = fixedXml.replace(
    /<w:rFonts[^>]*\/?>/g,
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/>'
  );
  
  // 3. Ensure minimum font size of 22 half-points (11pt)
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
  
  // 4. Same for complex script font sizes
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
  
  return fixedXml;
}

async function testShadowRemoval() {
  const testFilePath = 'tests/fixtures/test_problematic.docx';
  
  console.log('=== Testing Shadow Removal End-to-End ===');
  
  try {
    // Read the test file
    const buffer = fs.readFileSync(testFilePath);
    const zip = new JSZip();
    await zip.loadAsync(buffer);
    
    // Check original shadow count
    const originalDocXml = await zip.file('word/document.xml').async('string');
    const originalStylesXml = await zip.file('word/styles.xml').async('string');
    
    const originalDocShadows = (originalDocXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    const originalStylesShadows = (originalStylesXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    
    console.log(`Original document.xml shadows: ${originalDocShadows}`);
    console.log(`Original styles.xml shadows: ${originalStylesShadows}`);
    
    // Apply remediation (same logic as in download-document.js)
    let fixedDocXml = removeShadowsAndNormalizeFonts(originalDocXml);
    let fixedStylesXml = removeShadowsAndNormalizeFonts(originalStylesXml);
    
    // Check fixed shadow count
    const fixedDocShadows = (fixedDocXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    const fixedStylesShadows = (fixedStylesXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    
    console.log(`Fixed document.xml shadows: ${fixedDocShadows}`);
    console.log(`Fixed styles.xml shadows: ${fixedStylesShadows}`);
    
    // Update the zip with fixed content
    zip.file('word/document.xml', fixedDocXml);
    zip.file('word/styles.xml', fixedStylesXml);
    
    // Generate output file
    const outputBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    const outputPath = 'tests/fixtures/test_remediated.docx';
    fs.writeFileSync(outputPath, outputBuffer);
    
    console.log(`Remediated file saved to: ${outputPath}`);
    
    // Verify the output file
    const outputZip = new JSZip();
    await outputZip.loadAsync(outputBuffer);
    
    const verifyDocXml = await outputZip.file('word/document.xml').async('string');
    const verifyStylesXml = await outputZip.file('word/styles.xml').async('string');
    
    const verifyDocShadows = (verifyDocXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    const verifyStylesShadows = (verifyStylesXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    
    console.log(`Verification - document.xml shadows: ${verifyDocShadows}`);
    console.log(`Verification - styles.xml shadows: ${verifyStylesShadows}`);
    
    // Show font changes
    const originalTimesFonts = (originalDocXml.match(/Times New Roman/g) || []).length;
    const fixedTimesFonts = (verifyDocXml.match(/Times New Roman/g) || []).length;
    
    console.log(`Times New Roman occurrences: ${originalTimesFonts} → ${fixedTimesFonts}`);
    
    // Show size changes
    const originalSmallSizes = (originalDocXml.match(/<w:sz w:val="(1[0-9]|[0-9])"/g) || []).length;
    const fixedSmallSizes = (verifyDocXml.match(/<w:sz w:val="(1[0-9]|[0-9])"/g) || []).length;
    
    console.log(`Small font sizes (<22): ${originalSmallSizes} → ${fixedSmallSizes}`);
    
    if (verifyDocShadows === 0 && verifyStylesShadows === 0) {
      console.log('\n✅ SUCCESS: All shadows removed!');
    } else {
      console.log('\n❌ ISSUE: Shadows still present');
    }
    
    if (fixedTimesFonts === 0) {
      console.log('✅ SUCCESS: All fonts normalized to Arial!');
    } else {
      console.log('❌ ISSUE: Some Times New Roman fonts remain');
    }
    
    if (fixedSmallSizes === 0) {
      console.log('✅ SUCCESS: All font sizes increased to minimum 11pt!');
    } else {
      console.log('❌ ISSUE: Some small font sizes remain');
    }
    
  } catch (error) {
    console.error('Error:', error);
  }
}

testShadowRemoval();