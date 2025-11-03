// Session-based file storage with automatic cleanup
const fs = require('fs').promises;
const path = require('path');

class SessionManager {
  constructor() {
    this.sessions = new Map();
    this.cleanupInterval = 30 * 60 * 1000; // 30 minutes
    this.sessionTimeout = 60 * 60 * 1000; // 1 hour
    
    // Start cleanup timer
    setInterval(() => this.cleanupExpiredSessions(), this.cleanupInterval);
  }

  // Create a new session
  createSession() {
    const sessionId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);
    const sessionDir = `temp-sessions/${sessionId}`;
    
    const session = {
      sessionId,
      createdAt: Date.now(),
      lastActivity: Date.now(),
      directory: sessionDir,
      files: [],
      batches: [],
      reports: []
    };
    
    this.sessions.set(sessionId, session);
    
    // Create session directory
    this.ensureSessionDirectory(sessionDir);
    
    return session;
  }

  // Get existing session or create new one
  getOrCreateSession(sessionId) {
    if (sessionId && this.sessions.has(sessionId)) {
      const session = this.sessions.get(sessionId);
      session.lastActivity = Date.now();
      return session;
    }
    return this.createSession();
  }

  // Update session activity (keeps it alive)
  heartbeat(sessionId) {
    if (this.sessions.has(sessionId)) {
      const session = this.sessions.get(sessionId);
      session.lastActivity = Date.now();
      return true;
    }
    return false;
  }

  // Add file to session
  addFileToSession(sessionId, fileInfo) {
    const session = this.sessions.get(sessionId);
    if (session) {
      session.files.push(fileInfo);
      session.lastActivity = Date.now();
    }
  }

  // Add batch to session
  addBatchToSession(sessionId, batchInfo) {
    const session = this.sessions.get(sessionId);
    if (session) {
      session.batches.push(batchInfo);
      session.lastActivity = Date.now();
    }
  }

  // Get session files
  getSessionFiles(sessionId) {
    const session = this.sessions.get(sessionId);
    return session ? session.files : [];
  }

  // Get session batches
  getSessionBatches(sessionId) {
    const session = this.sessions.get(sessionId);
    return session ? session.batches : [];
  }

  // Clean up expired sessions
  async cleanupExpiredSessions() {
    const now = Date.now();
    const expiredSessions = [];

    for (const [sessionId, session] of this.sessions) {
      if (now - session.lastActivity > this.sessionTimeout) {
        expiredSessions.push(sessionId);
      }
    }

    for (const sessionId of expiredSessions) {
      await this.destroySession(sessionId);
    }

    if (expiredSessions.length > 0) {
      console.log(`🧹 Cleaned up ${expiredSessions.length} expired sessions`);
    }
  }

  // Manually destroy a session
  async destroySession(sessionId) {
    const session = this.sessions.get(sessionId);
    if (!session) return;

    try {
      // Delete all session files
      await this.deleteDirectory(session.directory);
      console.log(`🗑️ Deleted session directory: ${session.directory}`);
    } catch (error) {
      console.warn(`Failed to delete session directory ${session.directory}:`, error.message);
    }

    // Remove from memory
    this.sessions.delete(sessionId);
  }

  // Ensure session directory exists
  async ensureSessionDirectory(sessionDir) {
    try {
      await fs.mkdir(sessionDir, { recursive: true });
    } catch (error) {
      if (error.code !== 'EEXIST') {
        throw error;
      }
    }
  }

  // Recursively delete directory
  async deleteDirectory(dirPath) {
    try {
      const stats = await fs.stat(dirPath);
      if (stats.isDirectory()) {
        const files = await fs.readdir(dirPath);
        await Promise.all(
          files.map(file => this.deleteDirectory(path.join(dirPath, file)))
        );
        await fs.rmdir(dirPath);
      } else {
        await fs.unlink(dirPath);
      }
    } catch (error) {
      if (error.code !== 'ENOENT') {
        throw error;
      }
    }
  }

  // Get session stats
  getSessionStats() {
    return {
      activeSessions: this.sessions.size,
      sessions: Array.from(this.sessions.values()).map(s => ({
        sessionId: s.sessionId,
        createdAt: s.createdAt,
        lastActivity: s.lastActivity,
        filesCount: s.files.length,
        batchesCount: s.batches.length
      }))
    };
  }
}

// Global session manager instance
const sessionManager = new SessionManager();

module.exports = sessionManager;