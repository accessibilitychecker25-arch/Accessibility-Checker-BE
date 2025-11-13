const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

// Enhanced shadow detection to find ALL possible shadow variants
function findAllShadowTypes(xmlContent) {
  const patterns = [
    /<w:shadow[^>]*>/gi,
    /<w14:shadow[^>]*>/gi,
    /<w15:shadow[^>]*>/gi,
    /<a:shadow[^>]*>/gi,
    /<a14:shadow[^>]*>/gi,
    /<a15:shadow[^>]*>/gi,
    /shadow\w*\s*=\s*"[^"]*"/gi,
    /<[^>]*\s+shadow[^>]*>/gi,
  ];
  
  const allMatches = [];
  patterns.forEach((pattern, index) => {
    const matches = xmlContent.match(pattern) || [];
    if (matches.length > 0) {
      allMatches.push({
        pattern: pattern.toString(),
        matches: matches
      });
    }
  });
  
  return allMatches;
}

// Enhanced shadow removal function
function enhancedShadowRemoval(xmlContent) {
  let fixedXml = xmlContent;
  
  console.log('Before removal, shadow analysis:');
  const beforeShadows = findAllShadowTypes(fixedXml);
  beforeShadows.forEach(shadowType => {
    console.log(`  Pattern ${shadowType.pattern}: ${shadowType.matches.length} matches`);
    shadowType.matches.slice(0, 3).forEach(match => console.log(`    "${match}"`));
  });
  
  // Remove all shadow variants
  const removalPatterns = [
    /<w:shadow\s*\/>/gi,
    /<w:shadow[^>]*>.*?<\/w:shadow>/gi,
    /<w14:shadow\s*\/>/gi,
    /<w14:shadow[^>]*>.*?<\/w14:shadow>/gi,
    /<w15:shadow\s*\/>/gi,
    /<w15:shadow[^>]*>.*?<\/w15:shadow>/gi,
    /<a:shadow\s*\/>/gi,
    /<a:shadow[^>]*>.*?<\/a:shadow>/gi,
    /<a14:shadow\s*\/>/gi,
    /<a14:shadow[^>]*>.*?<\/a14:shadow>/gi,
    /<a15:shadow\s*\/>/gi,
    /<a15:shadow[^>]*>.*?<\/a15:shadow>/gi,
    /\s+\w*shadow\w*\s*=\s*"[^"]*"/gi,
  ];
  
  removalPatterns.forEach(pattern => {
    const before = (fixedXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    fixedXml = fixedXml.replace(pattern, '');
    const after = (fixedXml.match(/<[^>]*shadow[^>]*>/gi) || []).length;
    if (before !== after) {
      console.log(`  Removed ${before - after} shadows with pattern: ${pattern}`);
    }
  });
  
  console.log('After removal, remaining shadows:');
  const afterShadows = findAllShadowTypes(fixedXml);
  afterShadows.forEach(shadowType => {
    console.log(`  Pattern ${shadowType.pattern}: ${shadowType.matches.length} matches`);
    shadowType.matches.slice(0, 3).forEach(match => console.log(`    "${match}"`));
  });
  
  return fixedXml;
}

async function comprehensiveShadowTest() {
  console.log('=== Comprehensive Shadow Detection and Removal Test ===\n');
  
  // Test with our known test file
  const testFile = 'tests/fixtures/test_problematic.docx';
  
  if (!fs.existsSync(testFile)) {
    console.log('❌ Test file not found:', testFile);
    return;
  }
  
  try {
    const buffer = fs.readFileSync(testFile);
    const zip = new JSZip();
    await zip.loadAsync(buffer);
    
    // Check all XML files in the document for shadows
    const xmlFiles = ['word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/settings.xml'];
    
    for (const fileName of xmlFiles) {
      const file = zip.file(fileName);
      if (file) {
        console.log(`\n--- Testing ${fileName} ---`);
        const xmlContent = await file.async('string');
        
        // Enhanced shadow detection
        const shadows = findAllShadowTypes(xmlContent);
        if (shadows.length === 0) {
          console.log('✅ No shadows found');
        } else {
          console.log('🔍 Shadows detected:');
          shadows.forEach(shadowType => {
            console.log(`  ${shadowType.pattern}: ${shadowType.matches.length} matches`);
            shadowType.matches.slice(0, 2).forEach(match => console.log(`    "${match}"`));
          });
          
          // Test removal
          console.log('\n🔧 Testing shadow removal:');
          const cleaned = enhancedShadowRemoval(xmlContent);
          
          const remainingShadows = findAllShadowTypes(cleaned);
          if (remainingShadows.length === 0) {
            console.log('✅ All shadows successfully removed');
          } else {
            console.log('❌ Some shadows remain:');
            remainingShadows.forEach(shadowType => {
              console.log(`  ${shadowType.pattern}: ${shadowType.matches.length} matches`);
            });
          }
        }
      }
    }
    
    // Create a fully remediated version for user testing
    console.log('\n=== Creating Fully Remediated File ===');
    const docXml = await zip.file('word/document.xml').async('string');
    const stylesXml = await zip.file('word/styles.xml').async('string');
    
    const cleanedDoc = enhancedShadowRemoval(docXml);
    const cleanedStyles = enhancedShadowRemoval(stylesXml);
    
    zip.file('word/document.xml', cleanedDoc);
    zip.file('word/styles.xml', cleanedStyles);
    
    const outputBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    const outputFile = 'tests/fixtures/test_fully_remediated.docx';
    fs.writeFileSync(outputFile, outputBuffer);
    
    console.log(`\n📁 Fully remediated file created: ${outputFile}`);
    console.log('👀 Please test this file to see if shadows are truly removed.');
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

comprehensiveShadowTest();