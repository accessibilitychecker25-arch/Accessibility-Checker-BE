const fs = require('fs');
const uploadHandler = require('./api/upload-document');
const downloadHandler = require('./api/download-document');

// Mock request/response objects for testing
function createMockReq(filePath) {
  const fileData = fs.readFileSync(filePath);
  const filename = filePath.split('/').pop();
  
  return {
    method: 'POST',
    headers: {
      'content-type': 'multipart/form-data; boundary=test',
      'origin': 'http://localhost:3000'
    },
    pipe: (busboy) => {
      // Simulate file upload
      setTimeout(() => {
        busboy.emit('file', 'file', {
          on: (event, callback) => {
            if (event === 'data') {
              callback(fileData);
            } else if (event === 'end') {
              callback();
            }
          }
        }, { filename });
        
        setTimeout(() => {
          busboy.emit('finish');
        }, 10);
      }, 10);
    }
  };
}

function createMockRes() {
  let statusCode = 200;
  let headers = {};
  let responseData = null;
  
  return {
    setHeader: (key, value) => headers[key] = value,
    status: (code) => {
      statusCode = code;
      return {
        json: (data) => responseData = { statusCode, data },
        send: (data) => responseData = { statusCode, data },
        end: () => responseData = { statusCode }
      };
    },
    json: (data) => responseData = { statusCode, data },
    send: (data) => responseData = { statusCode, data },
    end: () => responseData = { statusCode },
    getResponse: () => responseData
  };
}

async function testNewFlaggingSystem() {
  console.log('=== Testing New Flagging System ===\n');
  
  const testFile = 'reports/Protected_remediated_by_agent.docx';
  
  if (!fs.existsSync(testFile)) {
    console.log(`Test file ${testFile} not found. Skipping test.`);
    return;
  }
  
  try {
    console.log('1. Testing upload analysis (should flag issues, not fix them)...');
    
    const req = createMockReq(testFile);
    const res = createMockRes();
    
    await uploadHandler(req, res);
    
    const uploadResult = res.getResponse();
    
    if (uploadResult && uploadResult.data && uploadResult.data.report) {
      const report = uploadResult.data.report;
      
      console.log('\n📊 Upload Analysis Results:');
      console.log(`   Fixed: ${report.summary.fixed}`);
      console.log(`   Flagged: ${report.summary.flagged}`);
      
      console.log('\n🔍 Issues detected:');
      if (report.details.titleNeedsFixing) console.log('   📝 Title needs fixing (flagged)');
      if (report.details.lineSpacingNeedsFixing) console.log('   📏 Line spacing needs fixing (flagged)');
      if (report.details.fontSizeNeedsFixing) console.log('   🔤 Font size needs fixing (flagged)');  
      if (report.details.fontTypeNeedsFixing) console.log('   🎨 Font type needs fixing (flagged)');
      if (report.details.textShadowsRemoved) console.log('   👤 Text shadows removed (fixed)');
      if (report.details.documentProtected) console.log('   🔒 Document protection removed (fixed)');
      if (report.details.imagesMissingOrBadAlt > 0) console.log(`   🖼️  ${report.details.imagesMissingOrBadAlt} images missing alt text (flagged)`);
      
      console.log('\n✅ Expected behavior:');
      console.log('   - Line spacing, font size, font type should be FLAGGED (not fixed)');
      console.log('   - Text shadows and document protection should be FIXED');
      console.log('   - Alt text and title issues should be FLAGGED');
      
    } else {
      console.log('❌ Upload analysis failed:', uploadResult);
    }
    
    console.log('\n2. Testing download processing (should only fix shadows and protection)...');
    
    const downloadReq = createMockReq(testFile);
    const downloadRes = createMockRes();
    
    await downloadHandler(downloadReq, downloadRes);
    
    const downloadResult = downloadRes.getResponse();
    
    if (downloadResult && downloadResult.statusCode === 200) {
      console.log('✅ Download processing completed successfully');
      console.log('   - Only shadows and document protection should be modified');
      console.log('   - Line spacing, font sizes, and font types should remain unchanged');
    } else {
      console.log('❌ Download processing failed:', downloadResult);
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

testNewFlaggingSystem();