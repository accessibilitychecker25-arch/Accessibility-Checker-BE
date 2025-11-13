const fs = require('fs');
const JSZip = require('jszip');

// Test the link descriptiveness detection
async function testLinkDescriptiveness() {
  console.log('=== Testing Link Descriptiveness Detection ===\n');
  
  const testFile = 'reports/Protected_remediated_by_agent.docx';
  
  if (!fs.existsSync(testFile)) {
    console.log(`Test file ${testFile} not found. Skipping test.`);
    return;
  }
  
  try {
    const fileData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(fileData);
    const documentXml = await zip.file('word/document.xml')?.async('string');
    
    if (documentXml) {
      console.log('📄 Document Analysis:');
      console.log(`   Document XML size: ${documentXml.length} characters`);
      
      // Look for hyperlink elements
      const hyperlinkMatches = documentXml.match(/<w:hyperlink[^>]*>/g) || [];
      console.log(`   Found ${hyperlinkMatches.length} <w:hyperlink> elements`);
      
      const fieldHyperlinkMatches = documentXml.match(/<w:fldSimple[^>]*HYPERLINK/g) || [];
      console.log(`   Found ${fieldHyperlinkMatches.length} field hyperlinks`);
      
      const hyperlinkStyleMatches = documentXml.match(/<w:rStyle w:val="Hyperlink"/g) || [];
      console.log(`   Found ${hyperlinkStyleMatches.length} hyperlink style runs`);
      
      // Look for any text that might be a link
      const urlPatterns = documentXml.match(/https?:\/\/[^\s<]+/g) || [];
      const emailPatterns = documentXml.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
      console.log(`   Found ${urlPatterns.length} URL patterns`);
      console.log(`   Found ${emailPatterns.length} email patterns`);
      
      if (hyperlinkMatches.length === 0 && fieldHyperlinkMatches.length === 0 && hyperlinkStyleMatches.length === 0) {
        console.log('   ℹ️  No hyperlinks found in this document');
        
        // Create a test with simulated link content
        console.log('\n🧪 Testing with Simulated Link Content:');
        
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
                  <w:r><w:t>read more</w:t></w:r>
                </w:hyperlink>
              </w:p>
              <w:p>
                <w:hyperlink r:id="rId3">
                  <w:r><w:t>Download the accessibility guide</w:t></w:r>
                </w:hyperlink>
              </w:p>
              <w:p>
                <w:hyperlink r:id="rId4">
                  <w:r><w:t>www.example.com</w:t></w:r>
                </w:hyperlink>
              </w:p>
            </w:body>
          </w:document>
        `;
        
        const results = testLinkAnalysis(testContent);
        console.log(`   Found ${results.nonDescriptiveLinks.length} non-descriptive links:`);
        
        results.nonDescriptiveLinks.forEach((link, index) => {
          console.log(`   ${index + 1}. "${link.linkText}" (${link.type})`);
          console.log(`      Location: ${link.location}`);
          console.log(`      Recommendation: ${link.recommendation}`);
          console.log('');
        });
        
      } else {
        console.log('\n🔍 Analyzing Actual Links:');
        // Here we'd run the actual analysis if there were links
      }
      
    } else {
      console.log('❌ Could not load document XML');
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

// Simplified version of the analysis for testing
function testLinkAnalysis(documentXml) {
  const results = { nonDescriptiveLinks: [] };
  
  const genericPhrases = [
    'click here', 'here', 'read more', 'more', 'link', 'this link',
    'see more', 'learn more', 'find out more', 'more info', 'more information'
  ];
  
  const hyperlinkMatches = documentXml.match(/<w:hyperlink[^>]*>[\s\S]*?<\/w:hyperlink>/g) || [];
  
  hyperlinkMatches.forEach((link, index) => {
    const textMatch = link.match(/<w:t[^>]*>(.*?)<\/w:t>/);
    if (textMatch) {
      const linkText = textMatch[1].toLowerCase().trim();
      
      const isGeneric = genericPhrases.some(phrase => linkText === phrase);
      const isUrl = linkText.includes('www.') || linkText.includes('http');
      
      let issueType = null;
      if (isGeneric) issueType = 'generic';
      if (isUrl) issueType = 'url-as-text';
      
      if (issueType) {
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

testLinkDescriptiveness();