const sessionManager = require('./lib/session-manager');
const fs = require('fs');
const path = require('path');

async function testSessionSystem() {
  console.log('🧪 Testing Session-Based Storage System\n');

  // 1. Create a session
  console.log('1. Creating new session...');
  const session = sessionManager.createSession();
  console.log(`✅ Session created: ${session.sessionId}`);
  console.log(`   Directory: ${session.directory}`);
  console.log(`   Created at: ${new Date(session.createdAt).toLocaleString()}`);

  // 2. Add some mock files to the session
  console.log('\n2. Adding files to session...');
  
  const mockFiles = [
    { filename: 'test1.docx', reportId: 'report-1', size: 1024 },
    { filename: 'test2.docx', reportId: 'report-2', size: 2048 },
    { filename: 'test3.docx', reportId: 'report-3', size: 1536 }
  ];

  mockFiles.forEach(file => {
    sessionManager.addFileToSession(session.sessionId, {
      filename: file.filename,
      reportId: file.reportId,
      originalPath: `${session.directory}/original-${file.reportId}.docx`,
      reportPath: `${session.directory}/${file.reportId}-accessibility-report.json`,
      processedAt: new Date().toISOString()
    });
  });

  console.log(`✅ Added ${mockFiles.length} files to session`);

  // 3. Add a batch to the session
  console.log('\n3. Adding batch to session...');
  const batchId = Date.now();
  sessionManager.addBatchToSession(session.sessionId, {
    batchId: batchId,
    timestamp: new Date().toISOString(),
    totalFiles: mockFiles.length,
    successful: mockFiles.length,
    failed: 0,
    reportPath: `${session.directory}/batch-${batchId}-summary.json`
  });

  console.log(`✅ Added batch ${batchId} to session`);

  // 4. Test heartbeat
  console.log('\n4. Testing heartbeat...');
  const heartbeatResult = sessionManager.heartbeat(session.sessionId);
  console.log(`✅ Heartbeat result: ${heartbeatResult}`);

  // 5. Get session data
  console.log('\n5. Getting session data...');
  const sessionFiles = sessionManager.getSessionFiles(session.sessionId);
  const sessionBatches = sessionManager.getSessionBatches(session.sessionId);
  
  console.log(`✅ Session has ${sessionFiles.length} files and ${sessionBatches.length} batches`);
  console.log('   Files:', sessionFiles.map(f => f.filename).join(', '));
  console.log('   Batches:', sessionBatches.map(b => b.batchId).join(', '));

  // 6. Test session stats
  console.log('\n6. Getting session statistics...');
  const stats = sessionManager.getSessionStats();
  console.log(`✅ Total active sessions: ${stats.activeSessions}`);
  
  // 7. Test automatic cleanup (simulate expired session)
  console.log('\n7. Testing session cleanup...');
  console.log('   (Creating expired session for cleanup test)');
  
  const expiredSession = sessionManager.createSession();
  // Manually set old timestamp to simulate expiration
  const session_obj = sessionManager.sessions.get(expiredSession.sessionId);
  session_obj.lastActivity = Date.now() - (2 * 60 * 60 * 1000); // 2 hours ago
  
  console.log(`   Created expired session: ${expiredSession.sessionId}`);
  
  // Force cleanup
  await sessionManager.cleanupExpiredSessions();
  
  const statsAfterCleanup = sessionManager.getSessionStats();
  console.log(`✅ Sessions after cleanup: ${statsAfterCleanup.activeSessions}`);

  // 8. Test session destruction
  console.log('\n8. Testing manual session destruction...');
  await sessionManager.destroySession(session.sessionId);
  
  const finalStats = sessionManager.getSessionStats();
  console.log(`✅ Final session count: ${finalStats.activeSessions}`);

  console.log('\n🎉 Session system test completed!');
  console.log('\n📝 Session Features:');
  console.log('   ✓ Automatic session creation');
  console.log('   ✓ File and batch tracking per session');  
  console.log('   ✓ Heartbeat to keep sessions alive');
  console.log('   ✓ Automatic cleanup after 1 hour of inactivity');
  console.log('   ✓ Manual session destruction');
  console.log('   ✓ Temporary file storage (no permanent accumulation)');
  
  console.log('\n💡 Usage:');
  console.log('   - Files are kept only during the user session');
  console.log('   - Sessions expire 1 hour after last activity');
  console.log('   - All files are automatically cleaned up');
  console.log('   - No need to manually manage file deletion');
}

// Run the test
testSessionSystem().catch(console.error);