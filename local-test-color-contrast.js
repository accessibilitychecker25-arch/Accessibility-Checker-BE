// local-test-color-contrast.js
// Locally invoke the backend's `analyzeDocx` function to test logic such as color contrast and line spacing.
//local testing feature for the backend.Command: node local-test-color-contrast.js
const fs = require('fs');
const path = require('path');

// Reference the modified upload-document handler function
const uploadHandler = require('./api/upload-document');
const analyzeDocx = uploadHandler.analyzeDocx;

async function run() {
  try {
    //  Test docx files are located in the test-docs directory.
    const testPath = path.join(
      __dirname,
      'test-docs',
      'Set one row to a very light gray.docx'
    );

    const fileData = fs.readFileSync(testPath);
    const report = await analyzeDocx(fileData, path.basename(testPath));

    console.log('=== Local analyzeDocx report ===');
    console.log(JSON.stringify(report, null, 2));
  } catch (err) {
    console.error('Local test failed:', err);
  }
}

run();
