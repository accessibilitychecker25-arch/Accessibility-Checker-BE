#!/usr/bin/env node

// Test document with overlapping form patterns that would cause duplicates
const testDocumentWithDuplicateForms = `
<w:document>
  <w:body>
    <w:p>
      <w:r><w:t>Form with multiple detectable patterns:</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
        <w:ffData>
          <w:name w:val="Text1"/>
          <w:textInput>
            <w:type w:val="regular"/>
          </w:textInput>
        </w:ffData>
      </w:r>
      <w:r>
        <w:instrText xml:space="preserve"> FORMTEXT </w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Another form field:</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
        <w:ffData>
          <w:checkBox/>
        </w:ffData>
      </w:r>
      <w:r>
        <w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Duplicate prevention test functions
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

function isPriorityFormType(newType, currentType) {
  const priorityOrder = {
    'form-data-complete': 10,
    'text-field': 9,
    'checkbox-field': 9,
    'dropdown-field': 9,
    'checkbox-control': 8,
    'dropdown-control': 8,
    'text-input': 8,
    'form-data': 7,
    'content-control': 6,
    'content-control-data': 5,
    'field-character': 4,
    'formtext-simple': 3,
    'formcheckbox-simple': 3,
    'formdropdown-simple': 3,
    'form-element': 1
  };
  
  return (priorityOrder[newType] || 1) > (priorityOrder[currentType] || 1);
}

function testDuplicatePrevention(documentXml) {
  const results = [];
  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;
  
  // Track unique form field locations to prevent duplicates
  const seenFormLocations = new Set();

  const formElements = [
    /<w:fldSimple[^>]*FORMTEXT/,
    /<w:fldSimple[^>]*FORMCHECKBOX/,
    /<w:fldSimple[^>]*FORMDROPDOWN/,
    /<w:ffData[\s\S]*?<\/w:ffData>/,
    /<w:ffData>/,
    /<w:checkBox/,
    /<w:dropDownList/,
    /<w:textInput/,
    /<w:sdt>/,
    /<w:sdtContent>/,
    /<w:fldChar w:fldCharType="begin"\/>/,
    /FORMTEXT/,
    /FORMCHECKBOX/,
    /FORMDROPDOWN/
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

    // Check for any form field in this paragraph and avoid duplicates
    let formDetectedInParagraph = false;
    let bestFormType = null;
    let detectedPatterns = [];

    formElements.forEach((regex, formIndex) => {
      const matches = paragraph.match(regex);
      if (matches) {
        formDetectedInParagraph = true;
        const formType = getFormType(formIndex);
        detectedPatterns.push(formType);
        
        // Prioritize more specific form types over generic ones
        if (!bestFormType || isPriorityFormType(formType, bestFormType)) {
          bestFormType = formType;
        }
      }
    });

    // Only add one form detection per paragraph to avoid duplicates
    if (formDetectedInParagraph) {
      const locationKey = `${paragraphCount}-${approximatePageNumber}`;
      
      if (!seenFormLocations.has(locationKey)) {
        seenFormLocations.add(locationKey);
        
        results.push({
          type: bestFormType,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150),
          recommendation: 'Consider using alternative formats...',
          detectedPatterns: detectedPatterns
        });
      }
    }
  });

  return results;
}

console.log('🔍 Testing Form Duplicate Prevention');
console.log('====================================\n');

const results = testDuplicatePrevention(testDocumentWithDuplicateForms);

console.log(`Forms detected: ${results.length}`);
console.log('');

results.forEach((result, index) => {
  console.log(`${index + 1}. Form Type: ${result.type}`);
  console.log(`   Location: ${result.location}`);
  console.log(`   Preview: ${result.preview}`);
  console.log(`   Detected Patterns: ${result.detectedPatterns.join(', ')}`);
  console.log(`   Priority Selection: Chose "${result.type}" from [${result.detectedPatterns.join(', ')}]`);
  console.log('');
});

console.log('📊 Test Analysis:');
console.log('=================');

// Count total possible matches without deduplication
let totalPossibleMatches = 0;
const paragraphRegex = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
const paragraphs = testDocumentWithDuplicateForms.match(paragraphRegex) || [];

const formElements = [
  /<w:fldSimple[^>]*FORMTEXT/, /<w:fldSimple[^>]*FORMCHECKBOX/, /<w:fldSimple[^>]*FORMDROPDOWN/,
  /<w:ffData[\s\S]*?<\/w:ffData>/, /<w:ffData>/, /<w:checkBox/, /<w:dropDownList/, /<w:textInput/,
  /<w:sdt>/, /<w:sdtContent>/, /<w:fldChar w:fldCharType="begin"\/>/, /FORMTEXT/, /FORMCHECKBOX/, /FORMDROPDOWN/
];

paragraphs.forEach(paragraph => {
  formElements.forEach(regex => {
    if (paragraph.match(regex)) {
      totalPossibleMatches++;
    }
  });
});

console.log(`Total possible matches without deduplication: ${totalPossibleMatches}`);
console.log(`Actual results after deduplication: ${results.length}`);
console.log(`Duplicates prevented: ${totalPossibleMatches - results.length}`);

if (results.length < totalPossibleMatches) {
  console.log('\n✅ SUCCESS: Duplicate prevention is working!');
  console.log('   Each paragraph with form fields is reported only once');
  console.log('   Higher priority form types are selected when multiple patterns match');
} else {
  console.log('\n❌ ISSUE: Duplicate prevention may not be working properly');
}

console.log('\n🎯 Key Features:');
console.log('  • One form detection per paragraph maximum');
console.log('  • Priority-based form type selection');
console.log('  • Location-based deduplication using Set()');
console.log('  • Debug info showing all detected patterns');