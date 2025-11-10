const fs = require('fs');
const JSZip = require('jszip');

async function debugDetection() {
  console.log('=== Debugging Detection Issues ===\n');
  
  // Test with an actual document
  const testFile = 'reports/Protected_remediated_by_agent.docx';
  
  if (!fs.existsSync(testFile)) {
    console.log('Test file not found, trying other files...');
    const reports = fs.readdirSync('reports');
    const docxFiles = reports.filter(f => f.endsWith('.docx'));
    if (docxFiles.length === 0) {
      console.log('No .docx files found in reports folder');
      return;
    }
    console.log(`Using ${docxFiles[0]} instead`);
  }
  
  try {
    const fileData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(fileData);
    
    console.log('1. CHECKING DOCUMENT.XML');
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (documentXml) {
      console.log(`Document XML length: ${documentXml.length}`);
      
      // Check for shadows
      const shadowTests = [
        /<w:shadow\s*\/>/,
        /<w:shadow[^>]*>/,
        /<a:outerShdw[^>]*>/,
        /<w14:shadow[^>]*>/
      ];
      
      console.log('\nShadow detection:');
      shadowTests.forEach((regex, i) => {
        const matches = documentXml.match(regex);
        console.log(`  Test ${i+1}: ${matches ? matches.length + ' matches' : 'no matches'}`);
        if (matches) console.log(`    First match: ${matches[0].slice(0, 100)}`);
      });
      
      // Check for serif fonts
      console.log('\nFont detection:');
      const serifMatches = documentXml.match(/(Times|Georgia|Garamond|serif)/gi);
      console.log(`  Serif fonts: ${serifMatches ? serifMatches.length + ' matches' : 'none found'}`);
      if (serifMatches) console.log(`    Found: ${[...new Set(serifMatches)].join(', ')}`);
      
      // Check font declarations
      const fontMatches = documentXml.match(/w:ascii="[^"]*"/g);
      if (fontMatches) {
        console.log(`  Font declarations: ${fontMatches.length}`);
        const uniqueFonts = [...new Set(fontMatches.map(m => m.match(/w:ascii="([^"]*)"/)[1]))];
        console.log(`    Fonts found: ${uniqueFonts.join(', ')}`);
      }
      
      // Check for small font sizes
      console.log('\nFont size detection:');
      const sizeMatches = documentXml.match(/<w:sz w:val="(\d+)"/g);
      if (sizeMatches) {
        console.log(`  Size declarations: ${sizeMatches.length}`);
        const sizes = sizeMatches.map(m => parseInt(m.match(/w:val="(\d+)"/)[1]));
        const smallSizes = sizes.filter(s => s < 22);
        console.log(`    Sizes found: ${[...new Set(sizes)].sort((a,b) => a-b).join(', ')}`);
        console.log(`    Small sizes (< 22): ${smallSizes.length > 0 ? smallSizes.join(', ') : 'none'}`);
      } else {
        console.log('  No size declarations found');
      }
      
      // Check line spacing
      console.log('\nLine spacing detection:');
      const spacingMatches = documentXml.match(/<w:spacing[^>]*w:line="(\d+)"[^>]*\/>/g);
      if (spacingMatches) {
        console.log(`  Spacing declarations: ${spacingMatches.length}`);
        spacingMatches.forEach(match => {
          const lineValue = parseInt(match.match(/w:line="(\d+)"/)[1]);
          console.log(`    ${match} -> ${lineValue} ${lineValue < 360 ? '(NEEDS FIX)' : '(OK)'}`);
        });
      } else {
        console.log('  No explicit spacing declarations found');
      }
      
      // Check for exact spacing
      if (documentXml.includes('w:lineRule="exact"')) {
        console.log('  Found exact line spacing rule (NEEDS FIX)');
      }
      
      // Check for paragraphs without spacing
      const totalParas = (documentXml.match(/<w:p[^>]*>/g) || []).length;
      const parasWithSpacing = (documentXml.match(/<w:p[^>]*>.*?<w:pPr[^>]*>.*?<w:spacing/gs) || []).length;
      console.log(`  Total paragraphs: ${totalParas}`);
      console.log(`  Paragraphs with spacing: ${parasWithSpacing}`);
      console.log(`  Paragraphs without spacing: ${totalParas - parasWithSpacing} ${totalParas - parasWithSpacing > 0 ? '(NEEDS FIX)' : '(OK)'}`);
    }
    
    console.log('\n2. CHECKING STYLES.XML');
    const stylesXml = await zip.file('word/styles.xml')?.async('string');
    if (stylesXml) {
      console.log(`Styles XML length: ${stylesXml.length}`);
      
      // Quick checks for styles
      const styleSerifMatches = stylesXml.match(/(Times|Georgia|Garamond|serif)/gi);
      console.log(`Serif fonts in styles: ${styleSerifMatches ? styleSerifMatches.length : 0}`);
      
      const styleSizeMatches = stylesXml.match(/<w:sz w:val="(\d+)"/g);
      if (styleSizeMatches) {
        const sizes = styleSizeMatches.map(m => parseInt(m.match(/w:val="(\d+)"/)[1]));
        const smallSizes = sizes.filter(s => s < 22);
        console.log(`Small font sizes in styles: ${smallSizes.length > 0 ? smallSizes.join(', ') : 'none'}`);
      }
    }
    
  } catch (error) {
    console.error('Debug failed:', error.message);
  }
}

debugDetection();