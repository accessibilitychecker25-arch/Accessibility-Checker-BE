#!/usr/bin/env node

// Mock document XML with GIF references
const mockDocumentWithGif = `
<w:document>
  <w:body>
    <w:p>
      <w:pStyle w:val="Heading1"/>
      <w:r><w:t>Image Gallery</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Here is an animated GIF:</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData>
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rId5">
                      <a:extLst>
                        <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                          <a14:useLocalDpi val="0"/>
                        </a:ext>
                      </a:extLst>
                    </a:blip>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Another section with different content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Mock relationships XML
const mockRelsXml = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.gif"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
`;

// Mock GIF files array
const mockGifFiles = ['word/media/image1.gif'];

// Test functions
function extractTextFromParagraph(paragraph) {
  const textMatch = paragraph.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!textMatch) return '';
  return textMatch.map(match => match.replace(/<[^>]*>/g, '')).join(' ');
}

function analyzeGifLocations(documentXml, relsXml, gifFiles) {
  const results = [];
  
  if (!relsXml || !gifFiles.length) {
    return results;
  }

  // Create mapping of relationship IDs to GIF files
  const gifRelationships = new Map();
  gifFiles.forEach(gifPath => {
    // Extract filename from path (e.g., "word/media/image1.gif" -> "image1.gif")
    const fileName = gifPath.split('/').pop();
    
    // Find relationship ID for this GIF in rels XML - try multiple patterns
    const patterns = [
      new RegExp(`<Relationship[^>]*Target="media/${fileName.replace('.', '\\.')}"[^>]*Id="([^"]*)"`, 'i'),
      new RegExp(`<Relationship[^>]*Id="([^"]*)"[^>]*Target="media/${fileName.replace('.', '\\.')}"`, 'i'),
      new RegExp(`Id="([^"]*)"[^>]*Target="[^"]*${fileName.replace('.', '\\.')}"`, 'i')
    ];
    
    for (const pattern of patterns) {
      const relMatch = relsXml.match(pattern);
      if (relMatch) {
        gifRelationships.set(relMatch[1], {
          file: gifPath,
          fileName: fileName
        });
        console.log(`Found GIF relationship: ${relMatch[1]} -> ${fileName}`);
        break;
      }
    }
  });
  
  console.log(`Total GIF relationships mapped: ${gifRelationships.size}`);

  if (gifRelationships.size === 0) {
    return results;
  }

  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Split document into paragraphs
  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    // Track page numbers (estimate)
    if (paragraphCount % 15 === 0) {
      approximatePageNumber++;
    }

    // Track headings for context
    if (/<w:pStyle w:val="Heading/.test(paragraph)) {
      currentHeading = extractTextFromParagraph(paragraph);
    }

    // Check if this paragraph contains any GIF references
    gifRelationships.forEach((gifInfo, relationshipId) => {
      // Look for drawing elements that reference this GIF - try multiple patterns
      const patterns = [
        new RegExp(`<w:drawing[\\s\\S]*?r:embed="${relationshipId}"[\\s\\S]*?</w:drawing>`, 'i'),
        new RegExp(`<a:blip[^>]*r:embed="${relationshipId}"`, 'i'),
        new RegExp(`r:embed="${relationshipId}"`, 'i'), // Simple embed reference
        new RegExp(`<w:drawing[\\s\\S]*?${relationshipId}[\\s\\S]*?</w:drawing>`, 'i') // Broader match
      ];
      
      let foundMatch = false;
      for (const pattern of patterns) {
        if (pattern.test(paragraph)) {
          console.log(`Found GIF in paragraph ${paragraphCount}: ${gifInfo.fileName} (pattern matched)`);
          foundMatch = true;
          break;
        }
      }
      
      if (foundMatch) {
        results.push({
          type: 'animated-gif',
          file: gifInfo.file,
          fileName: gifInfo.fileName,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150) || 'GIF image detected',
          recommendation: 'Replace animated GIFs with static images or accessible alternatives to prevent seizures and improve accessibility for users with vestibular disorders'
        });
      }
    });
  });

  return results;
}

console.log('🖼️ Testing GIF Location Detection');
console.log('==================================\n');

const results = analyzeGifLocations(mockDocumentWithGif, mockRelsXml, mockGifFiles);

console.log(`GIFs with locations detected: ${results.length}`);
console.log('');

results.forEach((result, index) => {
  console.log(`${index + 1}. GIF Detection:`);
  console.log(`   Type: ${result.type}`);
  console.log(`   File: ${result.file}`);
  console.log(`   File Name: ${result.fileName}`);
  console.log(`   Location: ${result.location}`);
  console.log(`   Page: ${result.approximatePage}`);
  console.log(`   Context: ${result.context}`);
  console.log(`   Preview: ${result.preview}`);
  console.log(`   Recommendation: ${result.recommendation.substring(0, 80)}...`);
  console.log('');
});

if (results.length > 0) {
  console.log('✅ SUCCESS: GIF location detection is working!');
  console.log('   GIFs are now tracked with paragraph-level location information');
  console.log('   Relationship mapping correctly links GIF files to document locations');
} else {
  console.log('❌ ISSUE: GIF location detection not working');
  console.log('   Check relationship mapping and paragraph detection logic');
}

console.log('\n📍 Location Features Added:');
console.log('  ✅ Paragraph-level precision for GIF placement');
console.log('  ✅ Page number estimation');
console.log('  ✅ Heading context tracking');
console.log('  ✅ Content preview for identification');
console.log('  ✅ Relationship ID mapping from rels XML');
console.log('  ✅ File path and filename information');
console.log('  ✅ Accessibility-focused recommendations');

console.log('\n🎯 Expected API Response Structure:');
console.log(`{
  "details": {
    "gifsDetected": ["word/media/image1.gif"],
    "gifLocations": [
      {
        "type": "animated-gif",
        "file": "word/media/image1.gif",
        "fileName": "image1.gif",
        "location": "Paragraph 3",
        "approximatePage": 1,
        "context": "Image Gallery",
        "preview": "GIF content preview...",
        "recommendation": "Replace animated GIFs with static images..."
      }
    ]
  }
}`);