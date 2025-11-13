const fs = require('fs');
const JSZip = require('jszip');

// Test to see what image location data is actually being returned
async function testImageLocations() {
  console.log('=== Testing Image Location Detection ===\n');
  
  const testFile = 'reports/Protected_remediated_by_agent.docx';
  
  if (!fs.existsSync(testFile)) {
    console.log(`Test file ${testFile} not found. Skipping test.`);
    return;
  }
  
  try {
    const fileData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(fileData);
    
    // Check what's in the document
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
    
    console.log('📄 Document Analysis:');
    if (documentXml) {
      console.log(`   Document XML size: ${documentXml.length} characters`);
      
      // Look for drawing/image elements
      const drawingMatches = documentXml.match(/<w:drawing>/g) || [];
      console.log(`   Found ${drawingMatches.length} <w:drawing> elements`);
      
      const blipMatches = documentXml.match(/<a:blip/g) || [];
      console.log(`   Found ${blipMatches.length} <a:blip> elements (images)`);
      
      const docPrMatches = documentXml.match(/<wp:docPr/g) || [];
      console.log(`   Found ${docPrMatches.length} <wp:docPr> elements (image properties)`);
      
      // Check for alt text patterns
      const descrMatches = documentXml.match(/descr="[^"]*"/g) || [];
      const titleMatches = documentXml.match(/title="[^"]*"/g) || [];
      console.log(`   Found ${descrMatches.length} descr attributes`);
      console.log(`   Found ${titleMatches.length} title attributes`);
      
      if (descrMatches.length > 0) {
        console.log('   Description attributes:', descrMatches);
      }
      if (titleMatches.length > 0) {
        console.log('   Title attributes:', titleMatches);
      }
    }
    
    console.log('\n📋 Relationships Analysis:');
    if (relsXml) {
      console.log(`   Relationships XML size: ${relsXml.length} characters`);
      
      // Look for image relationships
      const imageRelMatches = relsXml.match(/Type="[^"]*\/image"/g) || [];
      console.log(`   Found ${imageRelMatches.length} image relationships`);
      
      const allRelMatches = relsXml.match(/<Relationship[^>]*>/g) || [];
      console.log(`   Total relationships: ${allRelMatches.length}`);
      
      if (imageRelMatches.length > 0) {
        console.log('   Image relationship types:', imageRelMatches);
      }
    }
    
    // Test the actual image analysis function logic
    console.log('\n🔍 Running Image Analysis Logic:');
    
    if (documentXml && relsXml) {
      // Simulate the analyzeImageLocations function
      const imageRels = {};
      const relMatches = relsXml.match(/<Relationship[^>]*Type="[^"]*\/image"[^>]*>/g) || [];
      
      console.log(`   Processing ${relMatches.length} image relationships...`);
      
      relMatches.forEach((rel, index) => {
        const idMatch = rel.match(/Id="([^"]+)"/);
        const targetMatch = rel.match(/Target="([^"]+)"/);
        if (idMatch && targetMatch) {
          imageRels[idMatch[1]] = targetMatch[1];
          console.log(`   Relationship ${index + 1}: ${idMatch[1]} -> ${targetMatch[1]}`);
        }
      });
      
      // Look for paragraphs with images
      const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
      const paragraphs = documentXml.match(paragraphRegex) || [];
      
      console.log(`   Analyzing ${paragraphs.length} paragraphs for images...`);
      
      let imagesFound = 0;
      let imagesWithoutAlt = 0;
      
      paragraphs.forEach((paragraph, index) => {
        const drawingMatches = paragraph.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g) || [];
        
        drawingMatches.forEach(drawing => {
          const blipMatches = drawing.match(/<a:blip r:embed="([^"]+)"/g) || [];
          
          blipMatches.forEach(blip => {
            imagesFound++;
            const embedId = blip.match(/r:embed="([^"]+)"/)[1];
            const imagePath = imageRels[embedId];
            
            // Check for alt text
            const hasAltText = drawing.includes('<wp:docPr') && 
                              (drawing.includes('descr="') || drawing.includes('title="')) &&
                              !drawing.includes('descr=""') && !drawing.includes('title=""');
            
            console.log(`   Image ${imagesFound} in paragraph ${index + 1}:`);
            console.log(`     Embed ID: ${embedId}`);
            console.log(`     Path: ${imagePath}`);
            console.log(`     Has alt text: ${hasAltText}`);
            
            if (!hasAltText) {
              imagesWithoutAlt++;
            }
          });
        });
      });
      
      console.log(`\n📊 Results:`);
      console.log(`   Total images found: ${imagesFound}`);
      console.log(`   Images without alt text: ${imagesWithoutAlt}`);
      
      if (imagesWithoutAlt > 0) {
        console.log(`   ✅ Image location detection should be working`);
      } else {
        console.log(`   ℹ️  No images without alt text found in this document`);
      }
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

testImageLocations();