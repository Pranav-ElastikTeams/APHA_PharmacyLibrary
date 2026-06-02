import * as dotenv from 'dotenv';
import * as path from 'path';
import { chromium } from 'playwright';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

export const USER_DATA_DIR = path.resolve(__dirname, '..', '.browser-profile');

// Stealth args: tell Edge not to expose automation flags
const STEALTH_ARGS = [
  '--disable-blink-features=AutomationControlled',
  '--no-first-run',
  '--no-default-browser-check',
  '--disable-infobars',
  '--disable-notifications',
];

// Injected before every page load: removes the navigator.webdriver property
// that Cloudflare and other bot-detection services check first.
const STEALTH_INIT_SCRIPT = `
  Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
  Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
  window.chrome = { runtime: {} };
`;

export async function createStealthContext(headless: boolean) {
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless,
    channel: 'msedge',
    ignoreHTTPSErrors: true,
    viewport: { width: 1280, height: 900 },
    args: STEALTH_ARGS,
    // Suppress the "Chrome is being controlled by automated software" banner
    ignoreDefaultArgs: ['--enable-automation'],
  });
  await context.addInitScript(STEALTH_INIT_SCRIPT);
  return context;
}

async function main(): Promise<void> {
  console.log('='.repeat(60));
  console.log('Manual Login Browser');
  console.log('='.repeat(60));
  console.log('');
  console.log('Steps:');
  console.log('  1. Log in to pharmacylibrary.com in this browser');
  console.log('  2. Optionally navigate to a staging page to confirm access');
  console.log('  3. CLOSE this browser window when done');
  console.log('');
  console.log('Then run:  npm run check');
  console.log('');
  console.log(`Session will be saved to: ${USER_DATA_DIR}`);
  console.log('');

  const context = await createStealthContext(false);

  const page = await context.newPage();
  await page.goto('https://pharmacylibrary.com/').catch(() => {});

  await context.waitForEvent('close').catch(() => {});
  console.log('Browser closed — session saved.');
  console.log('Run: npm run check');
}

main().catch(err => {
  console.error('[Fatal]', err);
  process.exit(1);
});
