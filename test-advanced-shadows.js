const fs = require('fs');
const JSZip = require('jszip');

// Enhanced shadow detection and removal test
async function testAdvancedShadowRemoval() {
  console.log('=== Advanced Shadow Removal Test ===\n');
  
  const testFile = 'tests/fixtures/test_problematic.docx';
  
  if (!fs.existsSync(testFile)) {
    console.log('❌ Test file not found');
    return;
  }

  // Enhanced shadow removal function (matches the updated Node.js code)
  function removeShadowsAndNormalizeFonts(xmlContent) {
    let fixedXml = xmlContent;
    
    console.log(`Starting with ${(fixedXml.match(/<[^>]*shadow[^>]*>/gi) || []).length} basic shadow elements`);
    console.log(`Starting with ${(fixedXml.match(/outerShdw|innerShdw/gi) || []).length} advanced shadow properties`);
    
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
    
    // 5. Remove shadow properties and attributes
    fixedXml = fixedXml.replace(/outerShdw/g, '');
    fixedXml = fixedXml.replace(/innerShdw/g, '');
    fixedXml = fixedXml.replace(/\s+\w*shdw\w*\s*=\s*"[^"]*"/g, '');
    
    console.log(`Ending with ${(fixedXml.match(/<[^>]*shadow[^>]*>/gi) || []).length} basic shadow elements`);
    console.log(`Ending with ${(fixedXml.match(/outerShdw|innerShdw/gi) || []).length} advanced shadow properties`);
    
    return fixedXml;
  }

  try {
    const buffer = fs.readFileSync(testFile);
    const zip = new JSZip();
    await zip.loadAsync(buffer);
    
    // Test all XML files for shadows
    const xmlFiles = [
      'word/document.xml',
      'word/styles.xml', 
      'word/theme/theme1.xml'
    ];
    
    let totalShadowsFound = 0;
    let totalShadowsRemoved = 0;
    
    for (const fileName of xmlFiles) {
      const file = zip.file(fileName);
      if (file) {
        console.log(`\n--- Processing ${fileName} ---`);
        const xmlContent = await file.async('string');
        
        // Count initial shadows
        const basicShadows = (xmlContent.match(/<[^>]*shadow[^>]*>/gi) || []).length;
        const advancedShadows = (xmlContent.match(/outerShdw|innerShdw/gi) || []).length;
        const totalBefore = basicShadows + advancedShadows;
        
        console.log(`Found ${basicShadows} basic shadows, ${advancedShadows} advanced shadow properties`);
        totalShadowsFound += totalBefore;
        
        if (totalBefore > 0) {
          // Apply shadow removal
          const cleaned = removeShadowsAndNormalizeFonts(xmlContent);
          
          // Count remaining shadows
          const remainingBasic = (cleaned.match(/<[^>]*shadow[^>]*>/gi) || []).length;
          const remainingAdvanced = (cleaned.match(/outerShdw|innerShdw/gi) || []).length;
          const totalAfter = remainingBasic + remainingAdvanced;
          
          const removed = totalBefore - totalAfter;
          totalShadowsRemoved += removed;
          
          console.log(`Removed ${removed} shadow elements/properties`);
          console.log(`Remaining: ${remainingBasic} basic, ${remainingAdvanced} advanced`);
          
          if (totalAfter === 0) {
            console.log('✅ All shadows removed successfully');
          } else {
            console.log('❌ Some shadows remain');
            
            // Show what's left
            if (remainingBasic > 0) {
              const leftBasic = cleaned.match(/<[^>]*shadow[^>]*>/gi);
              console.log('  Remaining basic shadows:', leftBasic);
            }
            if (remainingAdvanced > 0) {
              const leftAdvanced = cleaned.match(/outerShdw|innerShdw/gi);
              console.log('  Remaining advanced shadows:', leftAdvanced);
            }
          }
          
          // Update the zip with cleaned content
          zip.file(fileName, cleaned);
        } else {
          console.log('✅ No shadows found');
        }
      }
    }
    
    console.log(`\n=== SUMMARY ===`);
    console.log(`Total shadows found: ${totalShadowsFound}`);
    console.log(`Total shadows removed: ${totalShadowsRemoved}`);
    
    if (totalShadowsRemoved === totalShadowsFound && totalShadowsFound > 0) {
      console.log('🎉 SUCCESS: All shadows completely removed!');
    } else if (totalShadowsFound === 0) {
      console.log('ℹ️  No shadows found in test file');
    } else {
      console.log('❌ ISSUE: Not all shadows were removed');
    }
    
    // Create enhanced remediated file
    const outputBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    const outputFile = 'tests/fixtures/test_advanced_remediated.docx';
    fs.writeFileSync(outputFile, outputBuffer);
    
    console.log(`\n📁 Enhanced remediated file created: ${outputFile}`);
    console.log('👀 Please test this file - it should have NO shadows of any type!');
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

testAdvancedShadowRemoval();