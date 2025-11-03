# Frontend Integration Guide - Session-Based Batch Processing

## 🚀 NEW ENDPOINTS AVAILABLE

### 1. **Session Management** - `/api/session`
**Purpose**: Initialize and maintain user sessions for temporary file storage

```javascript
// Initialize session when user opens the app
POST /api/session
Response: { sessionId: "1762145344331-h6evl2etm", success: true }

// Keep session alive (call every 5 minutes while user is active)
POST /api/session
Headers: { "X-Session-ID": "session-id-here" }
Response: { success: true, message: "Session refreshed" }

// Get session info and existing batches
GET /api/session?sessionId=session-id-here
Response: { 
  sessionId: "...", 
  files: [...], 
  batches: [...],
  expiresIn: "1 hour from last activity"
}
```

### 2. **Batch Upload** - `/api/batch-upload`
**Purpose**: Upload and process multiple DOCX files at once (up to 10 files)

```javascript
// Upload multiple files
POST /api/batch-upload
Headers: { "X-Session-ID": "session-id-here" }
Body: FormData with multiple files

Response: {
  sessionId: "session-id-here",
  batchId: 1762145344343,
  summary: {
    totalFiles: 5,
    successful: 4,
    failed: 1
  },
  results: [
    {
      fileIndex: 1,
      filename: "document1.docx",
      success: true,
      reportId: "report-123",
      summary: { flagged: 2, fixed: 1 },
      details: { ... }
    },
    // ... more files
  ],
  expiresIn: "1 hour"
}
```

### 3. **Batch Download** - `/api/batch-download`
**Purpose**: Download all remediated files as a ZIP

```javascript
// Download remediated files
GET /api/batch-download?batchId=1762145344343&sessionId=session-id-here
Response: ZIP file containing all remediated documents
```

---

## 📋 FRONTEND IMPLEMENTATION CHECKLIST

### Step 1: **Session Initialization** (Required)
```javascript
class AccessibilityChecker {
  constructor() {
    this.sessionId = null;
    this.heartbeatInterval = null;
  }

  async initializeSession() {
    try {
      const response = await fetch('/api/session', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' }
      });
      const data = await response.json();
      this.sessionId = data.sessionId;
      
      // Start heartbeat to keep session alive
      this.startHeartbeat();
      
      return this.sessionId;
    } catch (error) {
      console.error('Session initialization failed:', error);
    }
  }

  startHeartbeat() {
    // Send heartbeat every 5 minutes while user is active
    this.heartbeatInterval = setInterval(async () => {
      if (this.sessionId) {
        try {
          await fetch('/api/session', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'X-Session-ID': this.sessionId
            }
          });
        } catch (error) {
          console.warn('Heartbeat failed:', error);
        }
      }
    }, 5 * 60 * 1000); // 5 minutes
  }

  cleanup() {
    if (this.heartbeatInterval) {
      clearInterval(this.heartbeatInterval);
    }
    // Note: Server will auto-cleanup files after 1 hour
  }
}

// Initialize when app loads
const checker = new AccessibilityChecker();
checker.initializeSession();

// Cleanup when user leaves
window.addEventListener('beforeunload', () => checker.cleanup());
```

### Step 2: **Multi-File Upload UI** (Recommended)
```javascript
async function uploadMultipleFiles(files) {
  if (!checker.sessionId) {
    throw new Error('Session not initialized');
  }

  const formData = new FormData();
  files.forEach((file, index) => {
    formData.append(`file${index}`, file);
  });

  const response = await fetch('/api/batch-upload', {
    method: 'POST',
    headers: {
      'X-Session-ID': checker.sessionId
    },
    body: formData
  });

  if (!response.ok) {
    throw new Error(`Upload failed: ${response.statusText}`);
  }

  return await response.json();
}

// Usage example:
document.getElementById('fileInput').addEventListener('change', async (e) => {
  const files = Array.from(e.target.files);
  try {
    const result = await uploadMultipleFiles(files);
    console.log(`Processed ${result.summary.totalFiles} files`);
    console.log(`Batch ID: ${result.batchId}`);
    
    // Show results to user
    displayBatchResults(result);
  } catch (error) {
    console.error('Upload error:', error);
  }
});
```

### Step 3: **Download Remediated Files** (Required)
```javascript
function downloadBatch(batchId) {
  if (!checker.sessionId) {
    alert('Session expired. Please refresh the page.');
    return;
  }

  const downloadUrl = `/api/batch-download?batchId=${batchId}&sessionId=${checker.sessionId}`;
  
  // Create temporary download link
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = `batch-${batchId}-remediated.zip`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
```

---

## 🔄 MIGRATION FROM EXISTING ENDPOINTS

### If you're currently using single-file endpoints:

**Old way:**
```javascript
// Single file upload
POST /api/upload-document
POST /api/download-document
```

**New way (backward compatible):**
```javascript
// Keep using single file endpoints for 1 file
// OR use batch endpoints for 1+ files

// For multiple files:
POST /api/batch-upload (new)
GET /api/batch-download (new)
```

### **Integration Options:**

1. **Quick Integration** (minimal changes):
   - Add session initialization on app start
   - Keep existing single-file flow
   - Add optional multi-file upload as new feature

2. **Full Integration** (recommended):
   - Replace single-file with batch endpoints
   - Add drag-and-drop for multiple files
   - Show batch progress and results

---

## 🎯 UI/UX RECOMMENDATIONS

### **File Upload Area:**
```html
<!-- Support both single and multiple files -->
<input type="file" multiple accept=".docx" id="fileInput">

<!-- Or drag-and-drop area -->
<div id="dropArea">
  <p>Drop up to 10 DOCX files here, or click to select</p>
  <button>Select Files</button>
</div>
```

### **Progress Display:**
```javascript
// Show batch processing progress
function displayBatchResults(result) {
  const container = document.getElementById('results');
  
  container.innerHTML = `
    <h3>Batch Processing Complete</h3>
    <p>Processed: ${result.summary.totalFiles} files</p>
    <p>Successful: ${result.summary.successful}</p>
    <p>Failed: ${result.summary.failed}</p>
    
    <button onclick="downloadBatch('${result.batchId}')">
      Download All Remediated Files
    </button>
    
    <div class="file-list">
      ${result.results.map(file => `
        <div class="file-result ${file.success ? 'success' : 'error'}">
          <strong>${file.filename}</strong>
          ${file.success ? 
            `<span>✓ ${file.summary.fixed} issues fixed</span>` :
            `<span>✗ ${file.error}</span>`
          }
        </div>
      `).join('')}
    </div>
  `;
}
```

---

## 🚨 IMPORTANT NOTES

1. **Session Required**: All new endpoints require a valid session ID
2. **Auto-Cleanup**: Files expire after 1 hour of inactivity 
3. **No Permanent Storage**: Files are NOT saved permanently on the server
4. **Batch Limit**: Maximum 10 files per batch upload
5. **File Size**: Standard DOCX file size limits apply per file

---

## 📞 IMPLEMENTATION SUPPORT

**Ready-to-use example**: See `docs/batch-processing.html` for complete working implementation

**Test endpoints**: Use the existing test files in `tests/fixtures/` for testing

**Questions?** The backend is ready - just implement the session management and you're good to go! 🚀