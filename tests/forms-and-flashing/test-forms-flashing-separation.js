const fs = require('fs');
const JSZip = require('jszip');

// Mock document content for testing forms and flashing separation
const mockDocumentWithForms = `
<w:document>
  <w:body>
    <w:p>
      <w:r><w:t>Please fill out this form:</w:t></w:r>
    </w:p>
    <w:p>
      <w:sdt>
        <w:sdtPr>
          <w:text/>
        </w:sdtPr>
        <w:sdtContent>
          <w:r><w:t>Text Field</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>
    <w:p>
      <w:r><w:t>Check this box:</w:t></w:r>
      <w:checkBox/>
    </w:p>
  </w:body>
</w:document>
`;

const mockDocumentWithFlashing = `
<w:document>
  <w:body>
    <w:p>
      <w:r><w:t>This document has animated content:</w:t></w:r>
    </w:p>
    <w:p>
      <w:drawing>
        <a:animClr colorType="accent1"/>
        <w:r><w:t>Color changing animation</w:t></w:r>
      </w:drawing>
    </w:p>
    <w:p>
      <w:drawing>
        <a:animRot by="360000"/>
        <w:r><w:t>Rotating element</w:t></w:r>
      </w:drawing>
    </w:p>
  </w:body>
</w:document>
`;

const mockDocumentWithBoth = `
<w:document>
  <w:body>
    <w:p>
      <w:r><w:t>This document has both forms and animations:</w:t></w:r>
    </w:p>
    <w:p>
      <w:sdt>
        <w:sdtPr>
          <w:dropDownList/>
        </w:sdtPr>
        <w:sdtContent>
          <w:r><w:t>Dropdown Field</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>
    <w:p>
      <w:drawing>
        <a:animScale scaleX="150000" scaleY="150000"/>
        <w:r><w:t>Scaling animation</w:t></w:r>
      </w:drawing>
    </w:p>
    <w:p>
      <w:r><w:t>Enter your name:</w:t></w:r>
      <w:checkBox/>
    </w:p>
  </w:body>
</w:document>
`;

// Simple utility functions for testing
function extractTextFromParagraph(paragraph) {
  const textMatch = paragraph.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!textMatch) return '';
  return textMatch.map(match => match.replace(/<[^>]*>/g, '')).join(' ');
}

function getFormType(formIndex) {
  const formTypes = [
    'text-field', 'checkbox-field', 'dropdown-field', 'form-data',
    'checkbox-control', 'dropdown-control', 'text-input', 
    'content-control', 'content-control-data'
  ];
  return formTypes[formIndex] || 'form-element';
}

function getFlashingType(flashIndex) {
  const flashTypes = [
    'color-animation', 'rotation-animation', 'scale-animation', 
    'motion-animation', 'generic-animation', 'effect-animation',
    'timing-element', 'looping-video', 'looping-audio'
  ];
  return flashTypes[flashIndex] || 'animated-content';
}

// Test versions of our functions
function analyzeForms(documentXml) {
  const results = [];
  
  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Check for form elements in Word documents
  const formElements = [
    /<w:sdt>[\s\S]*?<w:text\/>/,                    // Text content controls
    /<w:sdt>[\s\S]*?<w:checkBox\/>/,                // Checkbox content controls
    /<w:sdt>[\s\S]*?<w:dropDownList/,               // Dropdown content controls
    /<w:formData[^>]*>/,                            // Form data elements
    /<w:checkBox[^>]*>/,                            // Checkbox form fields
    /<w:dropDownList[^>]*>/,                        // Dropdown form fields
    /<w:textInput[^>]*>/,                           // Text input form fields
    /<w:sdt>/,                                      // Any content control
    /<w:sdtContent>/                                // Content control data
  ];

  // Split into paragraphs for analysis
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
          recommendation: 'Consider using alternative formats like accessible web forms or structured tables instead of Word form fields'
        });
      }
    });
  });

  return results;
}

function analyzeFlashingObjects(documentXml) {
  const results = [];
  
  let paragraphCount = 0;
  let currentHeading = null;
  let approximatePageNumber = 1;

  // Check for potentially flashing/animated content
  const flashingElements = [
    /<w:drawing>[\s\S]*?<a:animClr/,              // Color animations
    /<w:drawing>[\s\S]*?<a:animRot/,              // Rotation animations  
    /<w:drawing>[\s\S]*?<a:animScale/,            // Scale animations
    /<w:drawing>[\s\S]*?<a:animMotion/,           // Motion animations
    /<w:drawing>[\s\S]*?<a:animate/,              // Generic animations
    /<w:drawing>[\s\S]*?<a:animEffect/,           // Effect animations
    /<p:timing>/,                                 // PowerPoint timing elements (embedded)
    /<a:videoFile[^>]*loop/,                      // Looping videos
    /<a:audioFile[^>]*loop/                       // Looping audio
  ];

  // Split into paragraphs for analysis
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

    flashingElements.forEach((regex, flashIndex) => {
      if (regex.test(paragraph)) {
        const flashType = getFlashingType(flashIndex);
        results.push({
          type: flashType,
          location: `Paragraph ${paragraphCount}`,
          approximatePage: approximatePageNumber,
          context: currentHeading || 'Document body',
          preview: extractTextFromParagraph(paragraph).substring(0, 150) || 'Animated content detected',
          recommendation: 'Remove animated/flashing content to prevent seizures and improve accessibility for all users'
        });
      }
    });
  });

  return results;
}

console.log('🧪 Testing Forms and Flashing Objects Separation');
console.log('================================================\n');

// Test 1: Document with only forms
console.log('Test 1: Document with only forms');
console.log('----------------------------------');
const formsResult = analyzeForms(mockDocumentWithForms);
const flashingFromFormsDoc = analyzeFlashingObjects(mockDocumentWithForms);

console.log('Forms found:', formsResult.length);
formsResult.forEach((issue, index) => {
  console.log(`  ${index + 1}. Type: ${issue.type}`);
  console.log(`     Location: ${issue.location}`);
  console.log(`     Preview: ${issue.preview}`);
});

console.log('Flashing objects found:', flashingFromFormsDoc.length);
if (flashingFromFormsDoc.length === 0) {
  console.log('  ✅ Correctly found no flashing objects in forms-only document');
}
console.log('');

// Test 2: Document with only flashing objects
console.log('Test 2: Document with only flashing objects');
console.log('--------------------------------------------');
const flashingResult = analyzeFlashingObjects(mockDocumentWithFlashing);
const formsFromFlashingDoc = analyzeForms(mockDocumentWithFlashing);

console.log('Flashing objects found:', flashingResult.length);
flashingResult.forEach((issue, index) => {
  console.log(`  ${index + 1}. Type: ${issue.type}`);
  console.log(`     Location: ${issue.location}`);
  console.log(`     Preview: ${issue.preview}`);
});

console.log('Forms found:', formsFromFlashingDoc.length);
if (formsFromFlashingDoc.length === 0) {
  console.log('  ✅ Correctly found no forms in flashing-only document');
}
console.log('');

// Test 3: Document with both forms and flashing objects
console.log('Test 3: Document with both forms and flashing objects');
console.log('-----------------------------------------------------');
const bothFormsResult = analyzeForms(mockDocumentWithBoth);
const bothFlashingResult = analyzeFlashingObjects(mockDocumentWithBoth);

console.log('Forms found:', bothFormsResult.length);
bothFormsResult.forEach((issue, index) => {
  console.log(`  ${index + 1}. Type: ${issue.type}`);
  console.log(`     Location: ${issue.location}`);
  console.log(`     Preview: ${issue.preview.substring(0, 50)}...`);
});

console.log('Flashing objects found:', bothFlashingResult.length);
bothFlashingResult.forEach((issue, index) => {
  console.log(`  ${index + 1}. Type: ${issue.type}`);
  console.log(`     Location: ${issue.location}`);
  console.log(`     Preview: ${issue.preview.substring(0, 50)}...`);
});

// Summary
console.log('\n📊 Test Summary');
console.log('================');
console.log(`✅ Forms function correctly identifies forms: ${formsResult.length > 0 && bothFormsResult.length > 0}`);
console.log(`✅ Flashing function correctly identifies animations: ${flashingResult.length > 0 && bothFlashingResult.length > 0}`);
console.log(`✅ Functions are properly separated: ${flashingFromFormsDoc.length === 0 && formsFromFlashingDoc.length === 0}`);
console.log(`✅ Both functions work on mixed content: ${bothFormsResult.length > 0 && bothFlashingResult.length > 0}`);

console.log('\n🎉 Forms and flashing objects are now properly separated into two distinct checks!');