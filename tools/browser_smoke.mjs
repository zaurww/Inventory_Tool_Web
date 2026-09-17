// npm install --no-save playwright, or set PLAYWRIGHT_MODULE to an installed module URL.
// Uses only synthetic workbooks made by browser_fixtures.py; never a working database.
import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const [url, directoryArg] = process.argv.slice(2);
if (!url || !directoryArg) throw new Error('Usage: node browser_smoke.mjs URL QA_DIRECTORY');
const directory = path.resolve(directoryArg);
const server = await chromium.launchServer({channel: process.env.BROWSER_CHANNEL || 'msedge', headless: true, host:'127.0.0.1'});
const browser = await chromium.connect(server.wsEndpoint());
const context = await browser.newContext({acceptDownloads: true, viewport: {width: 1440, height: 1000}});
const page = await context.newPage();
const requests = [], failures = [], errors = [];
context.on('request', r => {
  // Edge exposes its LOCAL downloads-hub resources too; they are not network traffic.
  if (/^https?:/.test(r.url())) requests.push({method:r.method(), url:r.url(), hasBody:!!r.postDataBuffer()});
});
context.on('requestfailed', r => failures.push({url:r.url(), error:r.failure()?.errorText}));
page.on('pageerror', e => errors.push(e.message));
const started = Date.now();
try {
  // Starter files must work even when JavaScript/Pyodide is unavailable.
  const starterContext = await browser.newContext({acceptDownloads: true, javaScriptEnabled: false});
  const starterPage = await starterContext.newPage();
  const starterResponse = await starterPage.goto(url, {waitUntil:'domcontentloaded'});
  assert.equal(starterResponse.status(), 200);
  for (const [id, filename] of [
    ['template-download', 'Inventory_Template_AZ.xlsx'],
    ['demo-download', 'Inventory_Demo_AZ.xlsx'],
  ]) {
    assert(await starterPage.locator(`#${id}`).isVisible());
    const pending = starterPage.waitForEvent('download');
    await starterPage.locator(`#${id}`).click();
    const result = await pending;
    assert.equal(result.suggestedFilename(), filename);
    await result.saveAs(path.join(directory, filename));
    assert.equal((await fs.readFile(path.join(directory, filename))).subarray(0, 2).toString(), 'PK');
  }
  await starterPage.close();

  const response = await page.goto(url, {waitUntil:'domcontentloaded'});
  assert.equal(response.status(), 200);
  const displayedVersion = await page.locator('#app-version').innerText();
  assert.match(displayedVersion, /^Inventory Tool \d+\.\d+\.\d+$/);
  await page.screenshot({path:path.join(directory,'starter-desktop.png'), fullPage:true});
  await page.setViewportSize({width:390, height:844});
  await page.screenshot({path:path.join(directory,'starter-mobile.png'), fullPage:true});
  assert(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth));
  await page.setViewportSize({width:1440, height:1000});
  await page.locator('#file-input').setInputFiles(path.join(directory, 'synthetic.xlsx'));
  await page.locator('#status-panel.is-ready, #status-panel.is-error').waitFor({timeout:180000});
  assert(await page.locator('#status-panel').getAttribute('class').then(c=>c.includes('is-ready')),
    await page.locator('#status-detail').innerText());
  const engineSeconds = (Date.now() - started) / 1000;
  const requestsAfterLoad = requests.length;

  async function calculate() {
    await page.locator('#calculate-button').click();
    await page.locator('#status-panel.is-success, #status-panel.is-error').waitFor({timeout:180000});
  }
  async function download(index, destination) {
    const pending = page.waitForEvent('download');
    await page.locator('#download-list a').nth(index).click();
    const result = await pending;
    await result.saveAs(path.join(directory, destination));
  }

  await calculate();
  assert(await page.locator('#results').isVisible(), await page.locator('#status-detail').innerText());
  assert.match(await page.locator('#result-note').innerText(), /Yeni kodların sayı: 2/);
  await download(0, 'report.xlsx');
  await download(1, 'updated.xlsx');
  const firstResult = await page.locator('#result-time').innerText();
  await page.screenshot({path:path.join(directory,'desktop.png'), fullPage:true});

  await page.locator('#file-input').setInputFiles(path.join(directory, 'updated.xlsx'));
  await calculate();
  assert.match(await page.locator('#result-note').innerText(), /Yeni kod tələb olunmadı/);
  await download(0, 'repeat-report.xlsx');
  await download(1, 'repeat-updated.xlsx');

  await page.locator('#file-input').setInputFiles(path.join(directory, 'Inventory_Demo_AZ.xlsx'));
  await calculate();
  assert(await page.locator('#results').isVisible(), await page.locator('#status-detail').innerText());
  assert.match(await page.locator('#result-note').innerText(), /Yeni kod tələb olunmadı/);
  await download(0, 'demo-report.xlsx');
  await download(1, 'demo-updated.xlsx');

  await page.setViewportSize({width:390, height:844});
  await page.screenshot({path:path.join(directory,'mobile.png'), fullPage:true});
  assert(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth));

  await page.locator('#file-input').setInputFiles(path.join(directory, 'invalid.xlsx'));
  await calculate();
  assert(await page.locator('#error-panel').isVisible());
  assert.equal(await page.locator('#results').isVisible(), false);
  assert.equal(await page.locator('#error-downloads a').count(), 2);
  const errorSummary = await page.locator('#error-summary').innerText();

  await page.locator('#file-input').setInputFiles({name:'broken.xlsx', mimeType:'application/octet-stream', buffer:Buffer.from('not a zip file')});
  await calculate();
  assert.equal(await page.locator('#results').isVisible(), false);
  assert.equal(await page.locator('#download-list a').count(), 0);
  assert.match(await page.locator('#status-panel').getAttribute('class'), /is-error/);
  assert.match(await page.locator('#status-detail').innerText(), /düzgün XLSX kitabı deyil/);

  assert(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth));
  assert.equal(errors.length, 0, JSON.stringify(errors));
  assert.equal(failures.length, 0, JSON.stringify(failures));
  assert(requests.every(r=>['GET','HEAD'].includes(r.method) && !r.hasBody));
  assert.equal(requests.length, requestsAfterLoad, 'Unexpected network activity after choosing/calculating workbooks');
  const result = {url, browser:browser.version(), engineSeconds, firstCalculation:firstResult,
    displayedVersion, starterDownloadsWithoutJavascript:true, demoCalculated:true,
    stableCodes:true, dataErrorHandled:errorSummary, corruptFileHandled:true, mobileFits:true,
    requests:requests.length, requestsAfterEngineReady:requests.length-requestsAfterLoad,
    pageErrors:errors, failedRequests:failures};
  await fs.writeFile(path.join(directory,'browser-results.json'), JSON.stringify(result,null,2));
  console.log(JSON.stringify(result,null,2));
} catch (error) {
  await page.screenshot({path:path.join(directory,'failure.png'), fullPage:true}).catch(()=>{});
  console.error(JSON.stringify({error:error.stack, status:await page.locator('#status-panel').innerText().catch(()=>''), errors, failures, requests},null,2));
  throw error;
} finally {
  // Edge's internal downloads hub can stall graceful close. This is only the
  // isolated headless test process; downloaded artifacts have already been saved.
  await server.kill();
}
