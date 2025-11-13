#!/usr/bin/env node

// Test document content that simulates the form field from the API response
const testDocumentWithForm = `
<w:document>
  <w:body>
    <w:p>
      <w:r><w:t>Form Example</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Please complete the following form:</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
        <w:ffData>
          <w:name w:val="Text1"/>
          <w:enabled/>
          <w:calcOnExit w:val="0"/>
          <w:textInput>
            <w:type w:val="regular"/>
            <w:default w:val=""/>
            <w:maxLength w:val="0"/>
            <w:format w:val=""/>
          </w:textInput>
        </w:ffData>
      </w:r>
      <w:bookmarkStart w:id="1" w:name="Text1"/>
      <w:r>
        <w:instrText xml:space="preserve"> FORMTEXT </w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
  </w:body>
</w:document>
`;

// Test the improved form detection
function extractTextFromParagraph(paragraph) {
  const textMatch = paragraph.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!textMatch) return '';
  return textMatch.map(match => match.replace(/<[^>]*>/g, '')).join(' ');
}

function getFormType(formIndex) {
  const formTypes = [
    'text-field', 'checkbox-field', 'dropdown-field', 'form-data-complete',
    'form-data', 'checkbox-control', 'dropdown-control', 'text-input', 
    'content-control', 'content-control-data', 'field-character',
    'formtext-simple', 'formcheckbox-simple', 'formdropdown-simple'
  ];
  return formTypes[formIndex] || 'form-element';
}

function testImprovedFormDetection(documentXml) {
  const results = [];
  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Updated form detection patterns
  const formElements = [
    /<w:fldSimple[^>]*FORMTEXT/,                  // Text form fields
    /<w:fldSimple[^>]*FORMCHECKBOX/,              // Checkbox form fields  
    /<w:fldSimple[^>]*FORMDROPDOWN/,              // Dropdown form fields
    /<w:ffData[\s\S]*?<\/w:ffData>/,              // Form field data (complete tags)
    /<w:ffData>/,                                 // Form field data (opening tag)
    /<w:checkBox/,                                // Checkbox controls
    /<w:dropDownList/,                            // Dropdown list controls
    /<w:textInput/,                               // Text input controls
    /<w:sdt>/,                                    // Structured document tags (content controls)
    /<w:sdtContent>/,                             // Content control content
    /<w:fldChar w:fldCharType="begin"\/>/,        // Field character begin
    /FORMTEXT/,                                   // Simple FORMTEXT detection
    /FORMCHECKBOX/,                               // Simple FORMCHECKBOX detection
    /FORMDROPDOWN/                                // Simple FORMDROPDOWN detection
  ];

  const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];

  paragraphs.forEach((paragraph, index) => {
    paragraphCount++;
    
    if (paragraphCount % 15 === 0) {
      approximatePageNumber++;
    }

    if (/<w:pStyle w:val="Heading/.test(paragraph)) {
      currentHeading = extractTextFromParagraph(paragraph);
    }

    formElements.forEach((regex, formIndex) => {
      const matches = paragraph.match(regex);
      if (matches) {
        const formType = getFormType(formIndex);
        results.push({
          type: formType,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150),
          recommendation: 'Consider using alternative formats like accessible web forms or structured tables instead of Word form fields',
          detectedPattern: regex.toString(),
          matchedText: matches[0]
        });
      }
    });
  });

  return results;
}

console.log('🔍 Testing Improved Form Detection');
console.log('==================================\n');

const results = testImprovedFormDetection(testDocumentWithForm);

console.log(`Forms detected: ${results.length}`);
console.log('');

results.forEach((result, index) => {
  console.log(`${index + 1}. Form Type: ${result.type}`);
  console.log(`   Location: ${result.location}`);
  console.log(`   Context: ${result.context}`);
  console.log(`   Preview: ${result.preview}`);
  console.log(`   Pattern: ${result.detectedPattern}`);
  console.log(`   Matched: ${result.matchedText.substring(0, 50)}...`);
  console.log('');
});

if (results.length > 0) {
  console.log('✅ SUCCESS: Improved form detection is working!');
  console.log('The patterns now properly detect form fields in Word documents.');
} else {
  console.log('❌ ISSUE: Form detection still not working');
  console.log('Need to further investigate the patterns.');
}

console.log('\n📝 Key Improvements Made:');
console.log('  • Added complete <w:ffData>...</w:ffData> pattern');
  console.log('  • Added <w:fldChar w:fldCharType="begin"/> detection');
console.log('  • Added simple FORMTEXT/FORMCHECKBOX/FORMDROPDOWN patterns');
console.log('  • Removed global flag (g) from regex patterns');
console.log('  • Added more comprehensive form element detection');