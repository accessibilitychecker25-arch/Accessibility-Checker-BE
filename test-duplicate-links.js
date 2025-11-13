const fs = require('fs');
const JSZip = require('jszip');

// Test duplicate link detection
async function testDuplicateLinkHandling() {
  console.log('=== Testing Duplicate Link Handling ===\n');
  
  try {
    // Create test content with duplicate non-descriptive links
    const testContent = `
      <w:document>
        <w:body>
          <w:p>
            <w:hyperlink r:id="rId1">
              <w:r><w:t>click here</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId2">
              <w:r><w:t>click here</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId3">
              <w:r><w:t>Click this link:</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId4">
              <w:r><w:t>read more</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId5">
              <w:r><w:t>read more</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId6">
              <w:r><w:t>read more</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId7">
              <w:r><w:t>www.example.com</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId8">
              <w:r><w:t>www.example.com</w:t></w:r>
            </w:hyperlink>
          </w:p>
        </w:body>
      </w:document>
    `;
    
    console.log('🧪 Testing with Content Containing Duplicates:');
    console.log('   - "click here" appears 2 times');  
    console.log('   - "Click this link:" appears 1 time');
    console.log('   - "read more" appears 3 times');
    console.log('   - "www.example.com" appears 2 times');
    console.log('   Total link elements: 8');
    console.log('   Expected unique issues: 4\n');
    
    const results = testLinkAnalysis(testContent);
    
    console.log('📊 Results:');
    console.log(`   Total non-descriptive links found: ${results.nonDescriptiveLinks.length}`);
    console.log(`   Should be 4 (no duplicates)\n`);
    
    const linkTextCounts = {};
    results.nonDescriptiveLinks.forEach((link, index) => {
      console.log(`   ${index + 1}. "${link.linkText}" (${link.type})`);
      console.log(`      Location: ${link.location}`);
      console.log(`      Recommendation: ${link.recommendation}`);
      console.log('');
      
      // Count occurrences to verify no duplicates
      linkTextCounts[link.linkText] = (linkTextCounts[link.linkText] || 0) + 1;
    });
    
    console.log('🔍 Verification:');
    let hasDuplicates = false;
    Object.entries(linkTextCounts).forEach(([linkText, count]) => {
      if (count > 1) {
        console.log(`   ❌ DUPLICATE: "${linkText}" appears ${count} times`);
        hasDuplicates = true;
      } else {
        console.log(`   ✅ UNIQUE: "${linkText}" appears ${count} time`);
      }
    });
    
    if (!hasDuplicates && results.nonDescriptiveLinks.length === 4) {
      console.log('\n✅ Duplicate handling test PASSED!');
      console.log('   All duplicate link texts were properly deduplicated.');
    } else {
      console.log('\n❌ Duplicate handling test FAILED!');
      console.log('   Expected 4 unique issues, got:', results.nonDescriptiveLinks.length);
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

// Updated test function with deduplication logic
function testLinkAnalysis(documentXml) {
  const results = { nonDescriptiveLinks: [] };
  const seenLinkTexts = new Set(); // Track unique link texts
  
  const genericPhrases = [
    'click here', 'here', 'read more', 'more', 'link', 'this link',
    'see more', 'learn more', 'find out more', 'more info', 'more information'
  ];

  const genericPatterns = [
    /^click\s+/i,           // "click this", "click the", etc.
    /\bclick\s+\w+\s*:?\s*$/i, // "click this link:", "click button:", etc.
    /^(here|there)\s*:?\s*$/i,  // "here:", "there:"
    /^(this|that)\s+link\s*:?\s*$/i, // "this link:", "that link:"
    /^read\s+(more|on)\s*:?\s*$/i,   // "read more:", "read on:"
    /^see\s+(more|here|this)\s*:?\s*$/i, // "see more:", "see here:", etc.
    /^(more|info|information)\s*:?\s*$/i, // "more:", "info:", etc.
    /^(download|view|open)\s*:?\s*$/i     // "download:", "view:", etc.
  ];
  
  const hyperlinkMatches = documentXml.match(/<w:hyperlink[^>]*>[\s\S]*?<\/w:hyperlink>/g) || [];
  
  hyperlinkMatches.forEach((link, index) => {
    const textMatch = link.match(/<w:t[^>]*>(.*?)<\/w:t>/);
    if (textMatch) {
      const linkText = textMatch[1].toLowerCase().trim();
      
      // Only process if we haven't seen this link text before  
      if (!seenLinkTexts.has(linkText)) {
        const isGeneric = genericPhrases.some(phrase => linkText === phrase);
        const isGenericPattern = genericPatterns.some(pattern => pattern.test(linkText));
        const isUrl = linkText.includes('www.') || linkText.includes('http');
        
        let issueType = null;
        if (isGeneric || isGenericPattern) issueType = 'generic';
        if (isUrl) issueType = 'url-as-text';
        
        if (issueType) {
          seenLinkTexts.add(linkText); // Mark as seen
          results.nonDescriptiveLinks.push({
            type: issueType,
            linkText: linkText,
            location: `Paragraph ${index + 1}`,
            approximatePage: 1,
            context: 'Document body',
            recommendation: generateRecommendation(linkText, issueType)
          });
        }
      }
    }
  });
  
  return results;
}

function generateRecommendation(linkText, issueType) {
  if (issueType === 'generic') {
    return 'Replace with descriptive text that explains where the link goes';
  }
  if (issueType === 'url-as-text') {
    return 'Replace URL with descriptive text like "Visit our website"';
  }
  return 'Use clear, descriptive language';
}

testDuplicateLinkHandling();