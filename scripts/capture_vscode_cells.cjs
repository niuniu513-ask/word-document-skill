/* Capture actual VS Code Electron/Jupyter UI, never an HTML replica.
   node capture_vscode_cells.cjs capture-plan.json screenshots.json [CDP URL]
   Set PLAYWRIGHT_MODULE to an absolute module path when require('playwright') is unavailable. */
const fs=require('fs'),path=require('path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
(async()=>{
 const [planFile,outFile,endpoint='http://127.0.0.1:9223']=process.argv.slice(2);
 if(!planFile||!outFile)throw Error('Usage: node capture_vscode_cells.cjs PLAN OUTPUT [CDP]');
 const plan=JSON.parse(fs.readFileSync(planFile,'utf8').replace(/^\uFEFF/,''));
 const root=path.dirname(path.resolve(outFile));fs.mkdirSync(path.join(root,'raw'),{recursive:true});
 const browser=await chromium.connectOverCDP(endpoint);
 try {
 const pages=browser.contexts().flatMap(c=>c.pages());
 let candidates=pages.filter(p=>/vscode|workbench/i.test(p.url()));
 if(process.env.VSCODE_WINDOW_TITLE){
  const matches=[];
  for(const p of candidates)if((await p.title()).includes(process.env.VSCODE_WINDOW_TITLE))matches.push(p);
  candidates=matches;
 }
 if(candidates.length!==1)throw Error('Expected one VS Code window; set VSCODE_WINDOW_TITLE to a unique existing title substring');
 const page=candidates[0];await page.bringToFront();await page.keyboard.press('Escape');
 const results=[];
 for(const item of plan.cells){
  await page.keyboard.press('Control+p');
  const input=page.locator('.quick-input-widget input:visible');await input.waitFor();
  await input.fill(path.resolve(item.evidence_notebook));await page.waitForTimeout(300);await page.keyboard.press('Enter');
  await page.waitForTimeout(900);
  if(!await page.locator('.notebookOverlay:visible').count()){
   await page.keyboard.press('Control+Shift+p');await input.waitFor();
   await input.fill('>View: Reopen Editor With...');await page.waitForTimeout(250);
   await page.locator('.quick-input-widget .monaco-list-row').filter({hasText:'View: Reopen Editor With...'}).first().click();
   await page.getByText('Jupyter Notebook',{exact:true}).click();await page.waitForTimeout(700);
  }
  const first=page.locator('.notebookOverlay:visible .monaco-list-row[data-index="0"]');
  const sentinel=page.locator('.notebookOverlay:visible .monaco-list-row[data-index="1"]');
  await first.waitFor();await sentinel.waitFor({state:'attached'});await page.mouse.move(5,5);
  const actual=(await first.innerText()).replaceAll('\u00a0',' ');
  const source=Array.isArray(item.source)?item.source.join(''):item.source;
  if(!actual.includes(source.split('\n')[0]))throw Error(`Wrong cell visible: ${item.cell}; expected=${source.split('\n')[0]}; actual=${actual.slice(0,200)}`);
  const env=await page.evaluate(()=>({w:innerWidth,h:innerHeight,d:devicePixelRatio}));
  const bb=await first.boundingBox();
  const firstTop=await first.evaluate(e=>parseFloat(e.style.top));
  const nextTop=await sentinel.evaluate(e=>parseFloat(e.style.top));
  if(!bb||!Number.isFinite(firstTop)||!Number.isFinite(nextTop))throw Error('Notebook DOM changed; calibrate selectors, do not guess');
  const outputBox={x:bb.x+44,y:bb.y+bb.height+2,width:bb.width-65,height:nextTop-firstTop-bb.height-25};
  if(bb.y<0||bb.y+bb.height>env.h-25||
      (item.hasOutput&&(outputBox.height<=0||outputBox.y+outputBox.height>env.h-25)))
    throw Error(`Cell ${item.cell} does not fit. Wrap source / expand outputs / calibrate UI and retry in new directory.`);
  const raw=path.join(root,'raw',`cell-${String(item.cell).padStart(3,'0')}.png`);
  await page.screenshot({path:raw});
  // Electron may pad the screenshot canvas beyond the rendered viewport.
  // Preserve the calibrated CSS-to-rendered-pixel scale rather than infer it from padded dimensions.
  results.push({...item,env,screenshotScale:{x:env.d,y:env.d},codeBox:bb,outputBox,codeFile:raw,outputFile:raw});
  fs.writeFileSync(outFile,JSON.stringify({...plan,cells:results},null,2));
  console.log(`Captured actual VS Code cell ${item.cell}`);
 }
 } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1});
