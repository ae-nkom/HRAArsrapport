import { defineConfig, devices } from '@playwright/test';

const basePath = process.env.BASE_PATH || '';
const localURL = `http://127.0.0.1:4173${basePath}/`;

export default defineConfig({
  testDir: './test/browser',
  fullyParallel: false,
  workers: 1,
  reporter: 'list',
  timeout: 30_000,
  use: {
    ...devices['Desktop Chrome'],
    baseURL: process.env.PLAYWRIGHT_BASE_URL || localURL,
    launchOptions: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH } : {},
    // Tests use fictional data. Do not enable recording for local HR sources.
    trace: 'off', screenshot: 'off', video: 'off'
  },
  webServer: process.env.PLAYWRIGHT_BASE_URL ? undefined : {
    command: 'node test/serve-build.mjs',
    url: localURL,
    reuseExistingServer: false
  }
});
