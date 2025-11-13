const fs = require('fs');
const JSZip = require('jszip');

// Test forms and flashing objects detection
async function testFormsAndFlashingDetection() {
  console.log('=== Testing Forms and Flashing Objects Detection ===\n');
  
  try {
    // Create test content with forms and animations
    const testContent = `
      <w:document>
        <w:body>
          <w:p>
            <w:pStyle w:val="Heading1"/>
            <w:r><w:t>Document with Forms</w:t></w:r>
          </w:p>
          <w:p>
            <w:fldSimple w:fldChar="FORMTEXT">
              <w:r><w:t>Enter your name:</w:t></w:r>
            </w:fldSimple>
          </w:p>
          <w:p>
            <w:fldSimple w:fldChar="FORMCHECKBOX">
              <w:r><w:t>I agree to terms</w:t></w:r>
            </w:fldSimple>
          </w:p>
          <w:p>
            <w:checkBox>
              <w:r><w:t>Check this box</w:t></w:r>
            </w:checkBox>
          </w:p>
          <w:p>
            <w:sdt>
              <w:sdtContent>
                <w:r><w:t>Content control field</w:t></w:r>
              </w:sdtContent>
            </w:sdt>
          </w:p>
          <w:p>
            <w:pStyle w:val="Heading2"/>  
            <w:r><w:t>Animated Content Section</w:t></w:r>
          </w:p>
          <w:p>
            <w:drawing>
              <a:animClr type="flash">
                <w:r><w:t>Flashing color animation</w:t></w:r>
              </a:animClr>
            </w:drawing>
          </w:p>
          <w:p>
            <w:drawing>
              <a:animRot angle="360">
                <w:r><w:t>Rotating element</w:t></w:r>
              </a:animRot>
            </w:drawing>
          </w:p>
          <w:p>
            <a:videoFile loop="true" src="video.mp4">
              <w:r><w:t>Looping video content</w:t></w:r>
            </a:videoFile>
          </w:p>
        </w:body>
      </w:document>
    `;
    
    console.log('🧪 Testing with Content Containing:');
    console.log('   - Text form field (FORMTEXT)');
    console.log('   - Checkbox form field (FORMCHECKBOX)');
    console.log('   - Checkbox control');
    console.log('   - Content control (SDT)');
    console.log('   - Color animation (flashing)');
    console.log('   - Rotation animation');  
    console.log('   - Looping video\n');
    
    const results = testFormsAndFlashingAnalysis(testContent);
    
    console.log('📊 Forms Detection Results:');
    console.log(`   Total forms found: ${results.formsDetected.length}`);
    console.log(`   Expected: 4 form elements\n`);
    
    results.formsDetected.forEach((form, index) => {
      console.log(`   ${index + 1}. Form Type: ${form.type}`);
      console.log(`      Location: ${form.location}`);
      console.log(`      Context: ${form.context}`);
      console.log(`      Preview: "${form.preview}"`);
      console.log(`      Recommendation: ${form.recommendation}`);
      console.log('');
    });
    
    console.log('📊 Flashing Objects Detection Results:');
    console.log(`   Total flashing objects found: ${results.flashingObjects.length}`);
    console.log(`   Expected: 3 animated elements\n`);
    
    results.flashingObjects.forEach((flash, index) => {
      console.log(`   ${index + 1}. Animation Type: ${flash.type}`);
      console.log(`      Location: ${flash.location}`);
      console.log(`      Context: ${flash.context}`);
      console.log(`      Preview: "${flash.preview}"`);
      console.log(`      Recommendation: ${flash.recommendation}`);
      console.log('');
    });
    
    if (results.formsDetected.length >= 4 && results.flashingObjects.length >= 3) {
      console.log('✅ Forms and Flashing Objects detection test PASSED!');
      console.log('   All expected form fields and animated content detected.');
    } else {
      console.log('❌ Detection test FAILED!');
      console.log(`   Expected: 4+ forms, 3+ animations`);
      console.log(`   Got: ${results.formsDetected.length} forms, ${results.flashingObjects.length} animations`);
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

// Test function for forms and flashing analysis
function testFormsAndFlashingAnalysis(documentXml) {
  const results = {
    formsDetected: [],
    flashingObjects: []
  };

  let paragraphCount = 0;
  let currentHeading = null;

  // Extract text from XML
  function extractTextFromParagraph(xml) {
    const textMatches = xml.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
    if (!textMatches) return '';
    return textMatches.map(t => t.replace(/<w:t[^>]*>|<\/w:t>/g, '')).join('').trim();
  }

  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    // Track headings
    const headingMatch = paragraph.match(/<w:pStyle w:val="(Heading\d+)"\/>/);
    if (headingMatch) {
      const headingText = extractTextFromParagraph(paragraph);
      currentHeading = `${headingMatch[1]}: ${headingText.substring(0, 50)}`;
    }

    // Check for form fields
    const formElements = [
      { regex: /<w:fldSimple[^>]*fldChar[^>]*FORMTEXT/, type: 'text-field' },
      { regex: /<w:fldSimple[^>]*fldChar[^>]*FORMCHECKBOX/, type: 'checkbox-field' },
      { regex: /<w:checkBox>/, type: 'checkbox-control' },
      { regex: /<w:sdt>/, type: 'content-control' }
    ];

    formElements.forEach(({ regex, type }) => {
      if (regex.test(paragraph)) {
        results.formsDetected.push({
          type: type,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: 1,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150),
          recommendation: 'Consider using alternative formats like accessible web forms instead of Word form fields'
        });
      }
    });

    // Check for flashing/animated content
    const flashingElements = [
      { regex: /<a:animClr/, type: 'color-animation' },
      { regex: /<a:animRot/, type: 'rotation-animation' },
      { regex: /<a:videoFile[^>]*loop/, type: 'looping-video' }
    ];

    flashingElements.forEach(({ regex, type }) => {
      if (regex.test(paragraph)) {
        results.flashingObjects.push({
          type: type,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: 1,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150) || 'Animated content detected',
          recommendation: 'Remove animated/flashing content to prevent seizures and improve accessibility'
        });
      }
    });
  });

  return results;
}

testFormsAndFlashingDetection();