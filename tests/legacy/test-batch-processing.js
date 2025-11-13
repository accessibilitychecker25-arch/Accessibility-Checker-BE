const fs = require('fs');
const FormData = require('form-data');
const fetch = require('node-fetch');

async function testBatchProcessing() {
  console.log('🔬 Testing Batch Processing System...\n');

  const API_BASE = 'http://localhost:3000'; // Adjust as needed
  
  // 1. Test batch upload with multiple files
  console.log('1. Testing batch upload...');
  
  const testFiles = [
    'tests/fixtures/test_problematic.docx',
    'tests/fixtures/test_remediated.docx',
    'tests/fixtures/test_advanced_remediated.docx'
  ];

  const form = new FormData();
  
  testFiles.forEach((filePath, index) => {
    if (fs.existsSync(filePath)) {
      const fileStream = fs.createReadStream(filePath);
      form.append(`file${index}`, fileStream, {
        filename: `test-file-${index + 1}.docx`
      });
      console.log(`  ✓ Added ${filePath}`);
    } else {
      console.log(`  ⚠️  File not found: ${filePath}`);
    }
  });

  try {
    console.log('\n📤 Uploading batch...');
    const uploadResponse = await fetch(`${API_BASE}/api/batch-upload`, {
      method: 'POST',
      body: form
    });

    if (!uploadResponse.ok) {
      console.log('❌ Upload failed:', uploadResponse.status, uploadResponse.statusText);
      return;
    }

    const uploadResult = await uploadResponse.json();
    console.log('✅ Batch upload successful!');
    console.log(`   Batch ID: ${uploadResult.batchId}`);
    console.log(`   Total files: ${uploadResult.summary.totalFiles}`);
    console.log(`   Successful: ${uploadResult.summary.successful}`);
    console.log(`   Failed: ${uploadResult.summary.failed}`);

    const batchId = uploadResult.batchId;

    // 2. Test batch listing
    console.log('\n2. Testing batch listing...');
    const listResponse = await fetch(`${API_BASE}/api/reports?action=batches`);
    
    if (listResponse.ok) {
      const listResult = await listResponse.json();
      console.log(`✅ Found ${listResult.totalBatches} batches`);
      
      if (listResult.batches.length > 0) {
        const latestBatch = listResult.batches[0];
        console.log(`   Latest batch: ${latestBatch.batchId} (${latestBatch.totalFiles} files)`);
      }
    } else {
      console.log('❌ Failed to list batches');
    }

    // 3. Test batch download
    console.log('\n3. Testing batch download...');
    const downloadResponse = await fetch(`${API_BASE}/api/batch-download?batchId=${batchId}`);
    
    if (downloadResponse.ok) {
      const zipBuffer = await downloadResponse.buffer();
      const outputPath = `batch-${batchId}-test-download.zip`;
      fs.writeFileSync(outputPath, zipBuffer);
      console.log(`✅ Batch downloaded: ${outputPath} (${zipBuffer.length} bytes)`);
    } else {
      console.log('❌ Failed to download batch');
    }

    // 4. Test individual report listing
    console.log('\n4. Testing report listing...');
    const reportsResponse = await fetch(`${API_BASE}/api/reports?action=list&limit=5`);
    
    if (reportsResponse.ok) {
      const reportsResult = await reportsResponse.json();
      console.log(`✅ Found ${reportsResult.totalReports} recent reports`);
      
      reportsResult.reports.forEach((report, index) => {
        console.log(`   ${index + 1}. ${report.filename} (${report.reportId})`);
      });
    } else {
      console.log('❌ Failed to list reports');
    }

    console.log('\n🎉 Batch processing test completed!');
    console.log('\nNext steps:');
    console.log('- Open docs/batch-processing.html in your browser');
    console.log('- Test with your own DOCX files');
    console.log('- Check the reports/ directory for stored results');

  } catch (error) {
    console.error('❌ Test error:', error);
  }
}

// Add CLI support
if (require.main === module) {
  testBatchProcessing();
}

module.exports = testBatchProcessing;