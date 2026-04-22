const test = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');

const HOST = '127.0.0.1';
const PORT = 3101;
const BASE_URL = `http://${HOST}:${PORT}`;

function waitForReady(proc, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (!settled) {
        settled = true;
        reject(new Error('Timed out waiting for server startup.'));
      }
    }, timeoutMs);

    proc.stdout.on('data', (chunk) => {
      const text = chunk.toString();
      if (text.includes('Assignment grader is running on')) {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          resolve();
        }
      }
    });

    proc.on('exit', (code) => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        reject(new Error(`Server exited before startup with code ${code}.`));
      }
    });
  });
}

test('smoke: homepage and profiles endpoint respond with 200', async () => {
  const proc = spawn(process.execPath, ['server.js'], {
    env: { ...process.env, PORT: String(PORT) },
    stdio: ['ignore', 'pipe', 'pipe']
  });

  try {
    await waitForReady(proc);

    const homeRes = await fetch(`${BASE_URL}/`);
    assert.equal(homeRes.status, 200, 'GET / should return 200');

    const profilesRes = await fetch(`${BASE_URL}/api/profiles`);
    assert.equal(profilesRes.status, 200, 'GET /api/profiles should return 200');

    const json = await profilesRes.json();
    assert.ok(Array.isArray(json.profiles), 'profiles payload should be an array');
  } finally {
    proc.kill('SIGTERM');
  }
});
