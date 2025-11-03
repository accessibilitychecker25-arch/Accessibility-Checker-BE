const fs = require('fs');
const JSZip = require('jszip');

async function checkDocumentForShadows(filePath) {
  console.log(`\n=== Checking ${filePath} for Shadows ===`);
  
  if (!fs.existsSync(filePath)) {
    console.log('❌ File not found:', filePath);
    return false;
  }
  
  try {
    const buffer = fs.readFileSync(filePath);
    const zip = new JSZip();
    await zip.loadAsync(buffer);
    
    let totalShadows = 0;
    const shadowDetails = [];
    
    // Check main XML files
    const xmlFiles = [
      'word/document.xml',
      'word/styles.xml',
      'word/numbering.xml',
      'word/settings.xml'
    ];
    
    for (const fileName of xmlFiles) {
      const file = zip.file(fileName);
      if (file) {
        const xmlContent = await file.async('string');
        
        // Find all shadow-related elements
        const shadowPatterns = [
          /<w:shadow[^>]*>/gi,
          /<w14:shadow[^>]*>/gi,
          /<a:shadow[^>]*>/gi,
          /shadow\w*\s*=\s*"[^"]*"/gi,
        ];
        
        let fileShadows = 0;
        const fileDetails = [];
        
        shadowPatterns.forEach(pattern => {
          const matches = xmlContent.match(pattern) || [];
          if (matches.length > 0) {
            fileShadows += matches.length;
            fileDetails.push({
              pattern: pattern.toString(),
              count: matches.length,
              samples: matches.slice(0, 3)
            });
          }
        });
        
        if (fileShadows > 0) {
          totalShadows += fileShadows;
          shadowDetails.push({
            file: fileName,
            count: fileShadows,
            details: fileDetails
          });
        }
      }
    }
    
    // Report results
    if (totalShadows === 0) {
      console.log('✅ NO SHADOWS FOUND - Document is clean!');
      return true;
    } else {
      console.log(`❌ ${totalShadows} SHADOW ELEMENTS FOUND:`);
      shadowDetails.forEach(fileInfo => {
        console.log(`\n  📄 ${fileInfo.file}: ${fileInfo.count} shadows`);
        fileInfo.details.forEach(detail => {
          console.log(`    Pattern: ${detail.pattern}`);
          console.log(`    Count: ${detail.count}`);
          detail.samples.forEach(sample => {
            console.log(`    Sample: "${sample}"`);
          });
        });
      });
      return false;
    }
    
  } catch (error) {
    console.log('❌ Error reading file:', error.message);
    return false;
  }
}

async function main() {
  console.log('Shadow Detection Utility');
  console.log('========================');
  
  // Check our test files
  const filesToCheck = [
    'tests/fixtures/test_problematic.docx',
    'tests/fixtures/test_remediated.docx',
    'tests/fixtures/test_fully_remediated.docx'
  ];
  
  for (const file of filesToCheck) {
    await checkDocumentForShadows(file);
  }
  
  console.log('\n📋 SUMMARY:');
  console.log('- test_problematic.docx: Original file with intentional shadows');
  console.log('- test_remediated.docx: Processed with Node.js remediation function');
  console.log('- test_fully_remediated.docx: Processed with enhanced removal');
  console.log('\n💡 TO TEST YOUR OWN FILE:');
  console.log('Copy your DOCX file to this directory and modify the filesToCheck array above.');
}

main();