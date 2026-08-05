import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const appCode=fs.readFileSync(new URL('../app/web/app.js',import.meta.url),'utf8');
const workerCode=fs.readFileSync(new URL('../app/web/diff-worker.js',import.meta.url),'utf8');

function functionSource(code,name){
  const start=code.indexOf(`function ${name}(`);if(start<0)throw new Error(`missing function: ${name}`);
  const brace=code.indexOf('{',start);let depth=0,quote='',escaped=false;
  for(let index=brace;index<code.length;index++){
    const ch=code[index];
    if(quote){if(escaped)escaped=false;else if(ch==='\\')escaped=true;else if(ch===quote)quote='';continue;}
    if(ch==='\''||ch==='"'||ch==='`'){quote=ch;continue;}
    if(ch==='{')depth++;else if(ch==='}'&&--depth===0)return code.slice(start,index+1);
  }
  throw new Error(`unterminated function: ${name}`);
}
const names=['normalizeDiffPdfText','diffPdfNumericFragments','diffPdfTextTemplate','diffPdfTextTemplatesMatch',
  'diffPdfNonNumericFingerprint','diffPdfNumericItems','diffPdfNumericCenterDistance','unmatchedDiffPdfNumericItems',
  'pairChangedDiffPdfNumbers','diffPdfTextPixelBox','dedupeDiffPdfRowItems','groupDiffPdfTextRows','matchDiffPdfTextRows',
  'buildTextRowStructureDiffResult','buildNumericTextDiffResult','diffRegionsOverlap','diffRegionsShareTextRow',
  'mergeDiffRegionsWithText','selectDiffSemanticResult','diffPdfTextLayoutFingerprint',
  'diffRegionOverlapsPdfText','shouldSuppressDiffRasterNoise'];
const app={Uint8Array,Uint16Array,Math,Number,Array,Map,Set,Object,String,Error};vm.createContext(app);
vm.runInContext(names.map(name=>functionSource(appCode,name)).join('\n'),app);
const item=(text,x,y,width=60,height=12)=>({text,x,y,width,height});

const beforeNumbers=[item('Report',40,20,80),item('Header',40,60,180),item('1 Item 580,400',40,90,290),
  item('2 Review 42,000',40,120,290),item('Total 622,400',190,150,140)];
const afterNumbers=beforeNumbers.map(value=>({...value}));
afterNumbers[2].text='1 Item 580,410';afterNumbers[4].text='Total 622,410';
afterNumbers.push(item('580,410',185,104,145,8));
const numeric=app.selectDiffSemanticResult(beforeNumbers,afterNumbers,600,400,
  {rowStructureAdjusted:false,alignmentAdjusted:false,fallbackUsed:false},[]);
assert.equal(numeric.mode,'numeric');
assert.equal(numeric.regions.length,2);
assert.ok(numeric.regions.every(region=>region.source==='pdf-text'));

const beforeRows=[item('Report',40,20,80),item('Header',40,60,180),item('1 Finance Close 100',40,90,280),
  item('2 HR Recruit 200',40,120,280),item('Total 300',200,150,120)];
const afterRows=[item('Report',40,20,80),item('Header',40,60,180),item('1 Finance Close 100',40,90,280),
  item('2 Legal Contract 50',40,120,280),item('3 HR Recruit 200',40,150,280),item('Total 350',200,180,120)];
const row=app.selectDiffSemanticResult(beforeRows,afterRows,600,400,
  {rowStructureAdjusted:true,alignmentAdjusted:true,fallbackUsed:false},[]);
assert.equal(row.mode,'row');assert.equal(row.regions.length,1);assert.equal(row.regions[0].kind,'added');
const duplicateAfterRows=[...afterRows,{...afterRows[2],y:afterRows[2].y+10}];
const duplicateRow=app.selectDiffSemanticResult(beforeRows,duplicateAfterRows,600,400,
  {rowStructureAdjusted:true,alignmentAdjusted:true,fallbackUsed:false},[]);
assert.equal(duplicateRow.mode,'row');assert.equal(duplicateRow.regions.length,1,'duplicate PDF glyph rows must not multiply an inserted row');

const stableText=[item('提出状況確認',100,200,120,18)];
const glyphNoise=[{x:.165,y:.22,width:.045,height:.025,pixelCount:20}];
assert.equal(app.shouldSuppressDiffRasterNoise(stableText,stableText,{changedRatio:.0002,alignmentAdjusted:false,fallbackUsed:false},glyphNoise,1200,900),true,
  'sparse noise overlapping unchanged text must be suppressed');
const fontColorChange=[{...glyphNoise[0],pixelCount:190}];
assert.equal(app.shouldSuppressDiffRasterNoise(stableText,stableText,{changedRatio:.0002,alignmentAdjusted:false,fallbackUsed:false},fontColorChange,1200,900),false,
  'dense text color changes must remain visible');

function analyze(before,after,width,height){
  let result=null;const self={postMessage:value=>{result=value;}};
  vm.runInNewContext(workerCode,{self,Uint8Array,Uint8ClampedArray,Uint16Array,Uint32Array,Int32Array,Float32Array,Math,Number,Array,Map,Set,Object,String,Error});
  self.onmessage({data:{id:1,width,height,before:before.buffer,after:after.buffer}});
  if(result?.error)throw new Error(result.error);return result;
}
const WIDTH=800,HEIGHT=600;
function white(){const data=new Uint8ClampedArray(WIDTH*HEIGHT*4);data.fill(255);return data;}
function gray(data,x,y,value=90){const index=(y*WIDTH+x)*4;data[index]=data[index+1]=data[index+2]=value;data[index+3]=255;}
function horizontal(data,x1,x2,y){for(let x=x1;x<=x2;x++)gray(data,x,y);}
function vertical(data,x,y1,y2){for(let y=y1;y<=y2;y++)gray(data,x,y);}
function glyphs(data,x,y,length=8){for(let glyph=0;glyph<length;glyph++)for(let dy=0;dy<7;dy++)for(let dx=0;dx<2;dx++)gray(data,x+glyph*4+dx,y+dy,105);}

function rowTable(rows){
  const data=white(),columns=[50,100,220,320,420,520];
  for(const row of rows)horizontal(data,50,520,row.y);horizontal(data,50,520,rows.at(-1).y+20);
  for(const x of columns)vertical(data,x,rows[0].y,rows.at(-1).y+20);
  for(const row of rows)for(let column=0;column<columns.length-1;column++)glyphs(data,columns[column]+5,row.y+7,3+(row.id+column)%7);
  return data;
}
const beforeRowLayout=Array.from({length:10},(_,id)=>({id,y:100+id*20}));
const afterRowLayout=[...beforeRowLayout.slice(0,5),{id:99,y:200},...beforeRowLayout.slice(5).map(row=>({...row,y:row.y+20}))];
const rowRaster=analyze(rowTable(beforeRowLayout),rowTable(afterRowLayout),WIDTH,HEIGHT);
assert.equal(rowRaster.rowStructureAdjusted,true);assert.equal(rowRaster.rowStructureRegions.length,1);
assert.equal(rowRaster.regions.length,1);assert.ok(rowRaster.regions[0].height<.08);

function weakColumnTable(columns){
  const data=white();for(let row=0;row<12;row++){const y=100+row*20;horizontal(data,50,520,y);for(const x of columns)glyphs(data,x,y+7);}
  horizontal(data,50,520,340);return data;
}
const columnRaster=analyze(weakColumnTable([60,110,250,330,410]),weakColumnTable([60,110,300,380,460]),WIDTH,HEIGHT);
assert.ok(columnRaster.alignmentAdjusted);assert.ok(columnRaster.regions.length<=3);
assert.ok(columnRaster.regions.every(region=>region.width<.15));

function ruledTable(horizontalRules,verticalRules){
  const data=white(),top=horizontalRules[0],bottom=horizontalRules.at(-1),left=verticalRules[0],right=verticalRules.at(-1);
  for(const y of horizontalRules)horizontal(data,left,right,y);
  for(const x of verticalRules)vertical(data,x,top,bottom);
  for(let row=0;row<horizontalRules.length-1;row++)for(let column=0;column<verticalRules.length-1;column++){
    glyphs(data,verticalRules[column]+5,horizontalRules[row]+6,Math.max(2,Math.min(8,(verticalRules[column+1]-verticalRules[column]-10)>>2)));
  }
  return data;
}

const resizedColumn=analyze(ruledTable([100,120,140,160,180,200,220,240,260,280,300,320],[50,100,220,320,420,520]),
  ruledTable([100,120,140,160,180,200,220,240,260,280,300,320],[50,100,280,380,480,520]),WIDTH,HEIGHT);
assert.equal(resizedColumn.regions.length,1,'one width change remains one full-column region');
assert.ok(resizedColumn.regions[0].before.width>.12&&resizedColumn.regions[0].after.width>.18,
  'column-width change highlights the whole column on both sides');

const normalRows=[100,120,140,160,180,200,220,240,260,280,300,320];
const tallRows=[100,120,140,160,200,220,240,260,280,300,320,340];
const rowHeight=analyze(ruledTable(normalRows,[50,100,220,320,420,520]),ruledTable(tallRows,[50,100,220,320,420,520]),WIDTH,HEIGHT);
assert.equal(rowHeight.regions.length,1,'one resized row remains one region');
assert.ok(rowHeight.regions[0].before&&rowHeight.regions[0].after,'row-height region needs side-specific geometry');
assert.ok(rowHeight.regions[0].after.height>rowHeight.regions[0].before.height*1.5,'after side shows the taller row');
const noisyInsertedRows=[100,120,140,160,180,200,240,260,280,300,320,340];
const noisyRowInsert=analyze(ruledTable(normalRows,[50,100,220,320,420,520]),ruledTable(noisyInsertedRows,[50,100,220,320,420,520]),WIDTH,HEIGHT);
assert.equal(noisyRowInsert.regions.length,1,'row insertion must discard page-wide residual regions');

const baseColumns=[50,100,220,320,420,520,620];
const insertedColumns=[50,100,220,320,420,470,520,620];
const columnAdded=analyze(ruledTable(normalRows,baseColumns),ruledTable(normalRows,insertedColumns),WIDTH,HEIGHT);
assert.equal(columnAdded.regions.length,1);assert.equal(columnAdded.regions[0].kind,'added');
assert.ok(columnAdded.regions[0].after.width<.09&&columnAdded.regions[0].after.width>.04,'added column is localized');
const columnRemoved=analyze(ruledTable(normalRows,insertedColumns),ruledTable(normalRows,baseColumns),WIDTH,HEIGHT);
assert.equal(columnRemoved.regions.length,1);assert.equal(columnRemoved.regions[0].kind,'removed');
assert.ok(columnRemoved.regions[0].before.width<.09&&columnRemoved.regions[0].before.width>.04,'removed column is localized');

console.log('diff regression tests OK');
