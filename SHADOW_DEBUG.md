**SHADOW DEBUGGING GUIDE**

The shadow removal is working correctly in our tests. Here's how to debug why you might still see shadows:

## Step 1: Verify File Processing
1. Copy your problematic DOCX file to this directory
2. Rename it to 'user_test.docx'
3. Edit check-shadows.js and add 'user_test.docx' to the filesToCheck array
4. Run: node check-shadows.js

## Step 2: Test the Full Workflow
1. Upload your file through the frontend
2. Download the remediated version
3. Check if the downloaded file has shadows using the tool above

## Step 3: Visual vs XML Shadows
The shadows we remove are XML-level text shadows (<w:shadow/>). If you're still seeing visual shadows, they might be:
- CSS shadows from the document viewer
- Theme-based formatting
- Different shadow types (drawing objects, shapes, etc.)

## Step 4: Common Issues
- **Browser caching**: Clear cache and re-download
- **Wrong file**: Make sure you're opening the remediated file, not the original
- **File corruption**: Check if the file opens correctly in Word
- **Different shadow types**: Some shadows might be in drawing objects, not text runs

## Test Files Available:
- test_problematic.docx: Has shadows (for testing detection)
- test_remediated.docx: Shadows removed (for testing removal)

## Contact Info:
If shadows persist after these checks, please:
1. Share the specific file you're testing
2. Describe where you see the shadows (which text, which page)
3. Confirm you're opening the downloaded/remediated file