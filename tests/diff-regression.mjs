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
  'diffPdfTextBoxUnion','mergeAdjacentDiffTextRegions','buildTextRowStructureDiffResult','findDiffPdfTextColumnSplit',
  'buildColumnTextRowStructureDiffResult','buildDocumentTextRowStructureDiffResult','buildTextLayoutShiftDiffResult',
  'buildNumericTextDiffResult','diffPdfTextItemDistance','buildTextFragmentDiffResult','diffRegionsOverlap','diffRegionsShareTextRow',
  'diffPdfTextItemsInBand','buildLocalizedTextRowStructureDiffResult','mergeDiffRegionsWithText','diffHasLocalizedRowSignal','selectDiffSemanticResult','diffPdfTextLayoutFingerprint',
  'diffRegionOverlapsPdfText','shouldSuppressDiffRasterNoise','diffKindMeta','diffCsvCell','buildDiffSummaryCsv'];
const app={Uint8Array,Uint16Array,Math,Number,Array,Map,Set,Object,String,Error};vm.createContext(app);
vm.runInContext(names.map(name=>functionSource(appCode,name)).join('\n'),app);
const item=(text,x,y,width=60,height=12)=>({text,x,y,width,height});

app.asArray=value=>Array.isArray(value)?value:[];
app.formatDateTime=value=>String(value||'');
const diffCsv=app.buildDiffSummaryCsv({comparison:{scope:'history',baselineAt:'2026-08-01',currentAt:'2026-08-07'},sheets:[
  {kind:'modified',beforeSheetName:'Page 1',afterSheetName:'Page 2',matchConfidence:.8,matchMethod:'sequence-between-anchors',beforePages:1,afterPages:1,message:'見出しを「A」から「B」に変更'},
  {kind:'added',beforeSheetName:'',afterSheetName:'Page 3',matchConfidence:1,matchMethod:'added',beforePages:0,afterPages:1,message:'追加ページ'}
]},'部門"報告.csv');
assert.match(diffCsv,/"任意2版"/);assert.match(diffCsv,/"80%"/);assert.match(diffCsv,/"sequence-between-anchors"/);
assert.match(diffCsv,/"部門""報告.csv"/,'CSV quotes must be escaped');
assert.match(diffCsv,/"追加"/);assert.doesNotMatch(diffCsv,/"added","0%"/,'added pages have no match confidence');

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

const cumulativeTextBefore=[item('Resume',40,30,100),item('Name',40,80,90),item('Address',40,120,120),item('History',40,160,110)];
const cumulativeTextAfter=[...cumulativeTextBefore.map(value=>({...value})),item('38',720,195,22),item('a',520,370,12),item('d',90,420,12),item('11',120,445,20)];
const structuralColumnRegion={x:.38,y:.1,width:.14,height:.39,before:{x:.39,y:.1,width:.11,height:.39},after:{x:.38,y:.1,width:.13,height:.39},confidence:1,pixelCount:70000};
const rasterNoiseRegion={x:.06,y:.03,width:.09,height:.08,confidence:.55,pixelCount:500};
const cumulativeText=app.selectDiffSemanticResult(cumulativeTextBefore,cumulativeTextAfter,800,600,
  {rowStructureAdjusted:false,alignmentAdjusted:true,fallbackUsed:false},[structuralColumnRegion,rasterNoiseRegion]);
assert.equal(cumulativeText.mode,'text-fragment','mixed layout and text changes use exact PDF text fragments');
assert.equal(cumulativeText.regions.length,5,'one structural marker and four cumulative text edits remain');
assert.equal(cumulativeText.regions.filter(region=>region.source==='pdf-fragment').length,4);
assert.ok(cumulativeText.regions.every(region=>region.source==='pdf-fragment'||region.confidence>=.9),
  'low-confidence alignment fragments are omitted when exact text boxes exist');

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

// Independent side-by-side tables can occupy the same Y positions. A row added
// only to the left table must be compared inside the raster-confirmed table band;
// otherwise the unchanged right table keeps the page-level row count constant.
const sideBySideBefore=[item('L1',40,100,180),item('L2',40,120,180),item('L3',40,140,180),
  item('R1',360,100,120),item('R2',360,120,120),item('R3',360,140,120),item('R4',360,160,120)];
const sideBySideAfter=[item('L1',40,100,180),item('L-new',40,120,180),item('L2',40,140,180),item('L3',40,160,180),
  item('R1',360,100,120),item('R2',360,120,120),item('R3',360,140,120),item('R4',360,160,120)];
assert.equal(app.buildTextRowStructureDiffResult(sideBySideBefore,sideBySideAfter,600,400).confident,false,
  'page-level row count is intentionally masked by the side table');
const sideBySideAnalysis={tableRowStructureDetected:true,tableRowStructureBand:{x:.04,y:.2,width:.48,height:.3},fallbackUsed:false};
const localizedSideBySide=app.buildLocalizedTextRowStructureDiffResult(sideBySideBefore,sideBySideAfter,600,400,sideBySideAnalysis);
assert.equal(localizedSideBySide.confident,true,'raster-confirmed table band restores the inserted text row');
assert.equal(localizedSideBySide.regions.length,1);assert.equal(localizedSideBySide.regions[0].kind,'added');
const sideBySideSemantic=app.selectDiffSemanticResult(sideBySideBefore,sideBySideAfter,600,400,sideBySideAnalysis,[]);
assert.equal(sideBySideSemantic.mode,'row');assert.equal(sideBySideSemantic.regions.length,1,'side table must not split the inserted row');

// A wrapped prose paragraph is several PDF rows but one human edit. Adjacent
// inserted lines must collapse into one marker instead of one marker per line.
const proseBefore=[item('A',40,40,300),item('B',40,60,300),item('C',40,80,300),item('D',40,100,300)];
const proseAfter=[item('A',40,40,300),item('B',40,60,300),item('NEW first line',40,80,300),
  item('NEW second line',40,100,260),item('C',40,120,300),item('D',40,140,300)];
const proseInsert=app.buildTextRowStructureDiffResult(proseBefore,proseAfter,600,400);
assert.equal(proseInsert.confident,true);assert.equal(proseInsert.regions.length,1,'wrapped paragraph lines merge into one edit');

// Page-wide rows can also be masked by a two-column resume. Detect the split
// from text geometry and compare the affected column independently.
const columnBefore=[item('L1',40,80,180),item('L2',40,100,180),item('L3',40,120,180),item('L4',40,140,180),
  item('R1',360,80,180),item('R2',360,100,180),item('R3',360,120,180),item('R4',360,140,180),item('R5',360,160,180)];
const columnAfter=[item('L1',40,80,180),item('L-new',40,100,180),item('L2',40,120,180),item('L3',40,140,180),item('L4',40,160,180),
  item('R1',360,80,180),item('R2',360,100,180),item('R3',360,120,180),item('R4',360,140,180),item('R5',360,160,180)];
assert.equal(app.buildTextRowStructureDiffResult(columnBefore,columnAfter,600,400).confident,false);
const columnInsert=app.buildColumnTextRowStructureDiffResult(columnBefore,columnAfter,600,400);
assert.equal(columnInsert.confident,true);assert.equal(columnInsert.regions.length,1);assert.equal(columnInsert.regions[0].kind,'added');
const columnDelete=app.buildColumnTextRowStructureDiffResult(columnAfter,columnBefore,600,400);
assert.equal(columnDelete.confident,true);assert.equal(columnDelete.regions.length,1);assert.equal(columnDelete.regions[0].kind,'removed');

// Whole-document LCS distinguishes a real paragraph inserted on page 1 from
// unchanged rows merely repaginated onto page 2.
const docBefore=[[item('A',40,40),item('B',40,60),item('C',40,80),item('D',40,100),item('E',40,120),item('F',40,140)],
  [item('G',40,40),item('H',40,60),item('I',40,80)]];
const docAfter=[[item('A',40,40),item('B',40,60),item('NEW 1',40,80),item('NEW 2',40,100),item('C',40,120),item('D',40,140)],
  [item('E',40,40),item('F',40,60),item('G',40,80),item('H',40,100),item('I',40,120)]];
const docPage1=app.buildDocumentTextRowStructureDiffResult(docBefore,docAfter,600,200,1);
assert.equal(docPage1.confident,true);assert.equal(docPage1.regions.length,1);assert.equal(docPage1.regions[0].kind,'added');
const docPage2=app.buildDocumentTextRowStructureDiffResult(docBefore,docAfter,600,200,2);
assert.equal(docPage2.confident,true);assert.equal(docPage2.regions.length,0,'repaginated unchanged text is not an edit on page 2');
const docDeletePage1=app.buildDocumentTextRowStructureDiffResult(docAfter,docBefore,600,200,1);
assert.equal(docDeletePage1.confident,true);assert.equal(docDeletePage1.regions.length,1);assert.equal(docDeletePage1.regions[0].kind,'removed');
const insertedPhysicalPage=[item('X1',40,40),item('X2',40,60),item('X3',40,80),item('X4',40,100),item('X5',40,120),item('X6',40,140)];
const shiftedDocument=[insertedPhysicalPage,...docBefore];
const mappedAfterInsertion=app.buildDocumentTextRowStructureDiffResult(docBefore,shiftedDocument,600,200,1,2);
assert.equal(mappedAfterInsertion.confident,true);assert.equal(mappedAfterInsertion.regions.length,0,
  'page text comparison uses separate before/after physical page numbers after an insertion');
const mappedAfterRemoval=app.buildDocumentTextRowStructureDiffResult(shiftedDocument,docBefore,600,200,2,1);
assert.equal(mappedAfterRemoval.confident,true);assert.equal(mappedAfterRemoval.regions.length,0,
  'reverse page text comparison preserves the shifted physical page mapping');

// A local line-spacing/border change shifts all following rows. Highlight the
// transition band once instead of every downstream text line.
const spacingBefore=[item('A',40,40,300),item('B',40,60,300),item('C',40,80,300),item('D',40,100,300),item('E',40,120,300)];
const spacingAfter=[item('A',40,40,300),item('B',40,60,300),item('C',40,88,300),item('D',40,108,300),item('E',40,128,300)];
const spacing=app.buildTextLayoutShiftDiffResult(spacingBefore,spacingAfter,600,400);
assert.equal(spacing.confident,true);assert.equal(spacing.regions.length,1);assert.equal(spacing.regions[0].source,'pdf-layout');

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
function glyphs(data,x,y,length=8,value=105){for(let glyph=0;glyph<length;glyph++)for(let dy=0;dy<7;dy++)for(let dx=0;dx<2;dx++)gray(data,x+glyph*4+dx,y+dy,value);}

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
assert.equal(resizedColumn.columnStructureKind,'','a width change is not an inserted column');
assert.match(resizedColumn.alignmentMode,/column-boundary-width/);
assert.ok(resizedColumn.regions[0].before.width>.12&&resizedColumn.regions[0].after.width>.18,
  'column-width change highlights the whole column on both sides');
assert.ok(resizedColumn.regions[0].before.y>.1,'column-width highlight does not climb into a distant title band');

const normalRows=[100,120,140,160,180,200,220,240,260,280,300,320];
const tallRows=[100,120,140,160,200,220,240,260,280,300,320,340];
const rowHeight=analyze(ruledTable(normalRows,[50,100,220,320,420,520]),ruledTable(tallRows,[50,100,220,320,420,520]),WIDTH,HEIGHT);
assert.equal(rowHeight.regions.length,1,'one resized row remains one region');
assert.equal(rowHeight.rowHeightAdjusted,true);assert.equal(rowHeight.columnStructureKind,'','a taller row is not an inserted column');
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

const manyColumns=Array.from({length:25},(_,index)=>50+index*20);
const manyColumnsWithExtra=[...manyColumns.slice(0,18),manyColumns[17]+10,...manyColumns.slice(18)];
const mergedGridNoise=analyze(ruledTable(normalRows,manyColumns),ruledTable(normalRows,manyColumnsWithExtra),WIDTH,HEIGHT);
assert.equal(mergedGridNoise.columnStructureKind,'','more than twenty vertical rules are not one credible table column model');

// A dense report page: title, explanatory paragraphs, two independent tables,
// a KPI box and notes.  The fixtures below deliberately keep unrelated content
// on the page so layout detection cannot succeed by treating the only table as
// the whole document.
function fill(data,x1,y1,x2,y2,value=220){for(let y=y1;y<=y2;y++)for(let x=x1;x<=x2;x++)gray(data,x,y,value);}
function textLine(data,x,y,length=28,value=105){glyphs(data,x,y,length,value);}
function drawRuledTable(data,horizontalRules,verticalRules,tone=105,rowIds=null){
  const top=horizontalRules[0],bottom=horizontalRules.at(-1),left=verticalRules[0],right=verticalRules.at(-1);
  fill(data,left,top,right,top+8,225);
  for(const y of horizontalRules)horizontal(data,left,right,y);
  for(const x of verticalRules)vertical(data,x,top,bottom);
  for(let rowIndex=0;rowIndex<horizontalRules.length-1;rowIndex++)for(let columnIndex=0;columnIndex<verticalRules.length-1;columnIndex++){
    const available=verticalRules[columnIndex+1]-verticalRules[columnIndex]-10;
    const identity=rowIds?.[rowIndex]??rowIndex,base=Math.max(2,Math.min(12,available>>2));
    glyphs(data,verticalRules[columnIndex]+5,horizontalRules[rowIndex]+7,Math.max(2,base-(identity+columnIndex)%3),tone);
  }
}
function richReport({mainRows=[110,130,150,170,190,210,230,250,270,290,310],mainColumns=[40,100,240,330,420,510],mainRowIds=null,colorCell=false,noise=[]}={}){
  const data=white();
  fill(data,40,25,750,45,75);textLine(data,48,31,42,250);
  textLine(data,40,58,70);textLine(data,40,72,54);textLine(data,430,72,34);
  drawRuledTable(data,mainRows,mainColumns,105,mainRowIds);
  drawRuledTable(data,[110,135,160,185,210,235],[550,610,680,750]);
  fill(data,550,255,750,310,238);textLine(data,565,267,34);textLine(data,565,286,26);
  textLine(data,40,342,78);textLine(data,40,356,74);textLine(data,40,370,65);textLine(data,40,384,82);
  drawRuledTable(data,[430,452,474,496,518,540,562],[40,110,300,410,520,640,750]);
  textLine(data,40,580,62);textLine(data,430,580,44);
  if(colorCell)for(let y=177;y<=183;y++)for(let x=247;x<=267;x++)if(data[(y*WIDTH+x)*4]<200)gray(data,x,y,190);
  for(const point of noise)gray(data,point.x,point.y,point.value??150);
  return data;
}
function richTextRows(includeAdded=false){
  const items=[item('月次業績報告 2026年8月',40,25,280,18),item('当月の業績と主要指標を報告します。',40,58,310,11),
    item('単位 百万円',620,72,90,11)];
  const labels=['売上高 12,500','営業利益 980','材料費 4,200','物流費 750','人件費 3,600','開発費 1,250','その他 420','小計 23,700','調整額 80','合計 23,780'];
  labels.forEach((text,index)=>items.push(item(text,45,117+index*20,320,11)));
  if(includeAdded){
    items.splice(3+5,0,item('追加監査費 150',45,117+5*20,320,11));
    for(let index=3+6;index<items.length;index++)items[index]={...items[index],y:items[index].y+20};
  }
  items.push(item('主要KPI 前年比 103.2%',555,117,170,11),item('為替影響 240',555,142,140,11));
  items.push(item('概況：国内販売は堅調に推移し、海外販売は一部地域で減少しました。',40,342,650,11));
  items.push(item('今後の見通し：原材料価格と為替変動を継続して注視します。',40,370,610,11));
  for(let index=0;index<6;index++)items.push(item(`${index+1} 部門別明細 ${1000+index*125}`,45,437+index*22,500,11));
  items.push(item('注：数値は速報値であり、確定値と異なる場合があります。',40,580,520,11));
  return items;
}

const richBeforeText=richTextRows(false),richAfterText=richTextRows(true);
const richRowSemantic=app.selectDiffSemanticResult(richBeforeText,richAfterText,WIDTH,HEIGHT,
  {rowStructureAdjusted:true,alignmentAdjusted:true,fallbackUsed:false},[]);
assert.equal(richRowSemantic.mode,'row','dense mixed-content report still recognizes one inserted row');
assert.equal(richRowSemantic.regions.length,1,'unrelated paragraphs and tables must not multiply the inserted row');

const richInsertedRows=[110,130,150,170,190,210,230,250,270,290,310,330];
const richRowBefore=richReport({mainRowIds:[0,1,2,3,4,5,6,7,8,9]}),
  richRowAfter=richReport({mainRows:richInsertedRows,mainRowIds:[0,1,2,3,4,99,5,6,7,8,9]});
const richRowRaster=analyze(richRowBefore,richRowAfter,WIDTH,HEIGHT);
assert.equal(richRowRaster.tableRowStructureDetected,true,'rich report row-count evidence survives unrelated content');
assert.ok(richRowRaster.regions.every(region=>region.width*region.height<.08),'raw rich-report candidates never cover the page');
const richRowIntegrated=app.selectDiffSemanticResult(richBeforeText,richAfterText,WIDTH,HEIGHT,richRowRaster,richRowRaster.regions);
assert.equal(richRowIntegrated.mode,'row','localized raster evidence promotes the exact PDF text row on a rich report');
assert.equal(richRowIntegrated.regions.length,1,'integrated rich report row result keeps one M marker');
const richRowRemovedRaster=analyze(richRowAfter,richRowBefore,WIDTH,HEIGHT);
assert.equal(richRowRemovedRaster.tableRowStructureDetected,true,'reverse rich report comparison keeps row-count evidence');
const richRowRemovedIntegrated=app.selectDiffSemanticResult(richAfterText,richBeforeText,WIDTH,HEIGHT,richRowRemovedRaster,richRowRemovedRaster.regions);
assert.equal(richRowRemovedIntegrated.mode,'row');assert.equal(richRowRemovedIntegrated.regions.length,1);

const richWidthRaster=analyze(richReport(),richReport({mainColumns:[40,100,280,370,460,510]}),WIDTH,HEIGHT);
assert.equal(richWidthRaster.regions.length,1,'rich report column resize remains one region');
assert.ok(richWidthRaster.regions[0].before&&richWidthRaster.regions[0].after,'rich report column resize preserves side-specific boxes');
assert.ok(richWidthRaster.regions[0].before.height>.3&&richWidthRaster.regions[0].after.height>.3,
  'rich report column resize highlights the complete table column');

// Six history generations can accumulate independent cell edits around a column
// structure change. Every arbitrary pair, in either direction, must retain the edits
// introduced between those generations instead of collapsing to the nearest change.
const cumulativeA=richReport(),cumulativeB=richReport({mainColumns:[40,100,280,370,460,510]});
fill(cumulativeB,84,177,90,183,35);fill(cumulativeB,716,296,722,302,35);
const cumulativeC=new Uint8ClampedArray(cumulativeB);fill(cumulativeC,696,366,702,372,35);
const cumulativeD=new Uint8ClampedArray(cumulativeC);fill(cumulativeD,316,397,322,403,35);
const cumulativeE=new Uint8ClampedArray(cumulativeD);fill(cumulativeE,606,407,612,413,35);
const cumulativeF=new Uint8ClampedArray(cumulativeE);fill(cumulativeF,146,573,152,579,35);
const cumulativeAB=analyze(cumulativeA,cumulativeB,WIDTH,HEIGHT),cumulativeBC=analyze(cumulativeB,cumulativeC,WIDTH,HEIGHT),
  cumulativeAC=analyze(cumulativeA,cumulativeC,WIDTH,HEIGHT),cumulativeAF=analyze(cumulativeA,cumulativeF,WIDTH,HEIGHT);
function diffRegionsCoverPoint(regions,x,y){
  const nx=x/WIDTH,ny=y/HEIGHT;
  return regions.some(region=>[region,region.before,region.after].filter(Boolean).some(box=>
    nx>=box.x&&nx<=box.x+box.width&&ny>=box.y&&ny<=box.y+box.height));
}
assert.ok(cumulativeAB.regions.length>=3,'column resize keeps both earlier independent cell edits');
assert.ok(diffRegionsCoverPoint(cumulativeAB.regions,87,180)&&diffRegionsCoverPoint(cumulativeAB.regions,719,299),
  'A-to-B comparison covers every earlier edit outside the resized column');
assert.ok(diffRegionsCoverPoint(cumulativeBC.regions,699,369),'B-to-C comparison covers the latest edit');
assert.ok(cumulativeAC.regions.length>=4,'A-to-C comparison keeps the structural marker and all cumulative edits');
assert.ok([[87,180],[719,299],[699,369]].every(([x,y])=>diffRegionsCoverPoint(cumulativeAC.regions,x,y)),
  'A-to-C comparison covers edits introduced in every generation');
assert.ok([[87,180],[719,299],[699,369],[319,400],[609,410],[149,576]].every(([x,y])=>diffRegionsCoverPoint(cumulativeAF.regions,x,y)),
  'oldest-to-latest comparison covers edits introduced across six generations');

const cumulativeGenerations=[cumulativeA,cumulativeB,cumulativeC,cumulativeD,cumulativeE,cumulativeF];
const cumulativeEdits=[
  {generation:1,points:[[87,180],[719,299]]},
  {generation:2,points:[[699,369]]},
  {generation:3,points:[[319,400]]},
  {generation:4,points:[[609,410]]},
  {generation:5,points:[[149,576]]}
];
const cumulativeAnalysisCache=new Map([
  ['0:1',cumulativeAB],['1:2',cumulativeBC],['0:2',cumulativeAC],['0:5',cumulativeAF]
]);
function analyzeCumulativePair(from,to){
  const key=`${from}:${to}`;
  if(!cumulativeAnalysisCache.has(key))cumulativeAnalysisCache.set(key,analyze(cumulativeGenerations[from],cumulativeGenerations[to],WIDTH,HEIGHT));
  return cumulativeAnalysisCache.get(key);
}
for(let from=0;from<cumulativeGenerations.length-1;from++)for(let to=from+1;to<cumulativeGenerations.length;to++){
  const expected=cumulativeEdits.filter(edit=>edit.generation>from&&edit.generation<=to).flatMap(edit=>edit.points);
  const forward=analyzeCumulativePair(from,to);
  assert.ok(expected.every(([x,y])=>diffRegionsCoverPoint(forward.regions,x,y)),`generation ${from}-to-${to} forward comparison keeps every intervening edit`);
}
for(const [from,to] of [[5,0],[4,1],[3,2]]){
  const expected=cumulativeEdits.filter(edit=>edit.generation>to&&edit.generation<=from).flatMap(edit=>edit.points);
  const reverse=analyzeCumulativePair(from,to);
  assert.ok(expected.every(([x,y])=>diffRegionsCoverPoint(reverse.regions,x,y)),`generation ${from}-to-${to} reverse comparison keeps every intervening edit`);
}

const richStableText=richTextRows(false),manySparseTextNoise=Array.from({length:12},(_,index)=>({
  x:(48+(index%6)*38)/WIDTH,y:(58+Math.floor(index/6)*14)/HEIGHT,width:7/WIDTH,height:7/HEIGHT,pixelCount:1
}));
assert.equal(app.shouldSuppressDiffRasterNoise(richStableText,richStableText,
  {changedRatio:.00018,alignmentAdjusted:false,fallbackUsed:false},manySparseTextNoise,WIDTH,HEIGHT),true,
  'many scattered antialiasing specks on a dense report must be suppressed together');
const meaningfulColorRegion=[{x:240/WIDTH,y:174/HEIGHT,width:32/WIDTH,height:16/HEIGHT,pixelCount:220}];
assert.equal(app.shouldSuppressDiffRasterNoise(richStableText,richStableText,
  {changedRatio:.0004,alignmentAdjusted:false,fallbackUsed:false},meaningfulColorRegion,WIDTH,HEIGHT),false,
  'a real font-color change inside a dense report must remain visible');

const richNumericAfter=richTextRows(false).map(value=>({...value}));
richNumericAfter.find(value=>value.text==='売上高 12,500').text='売上高 12,510';
richNumericAfter.find(value=>value.text==='合計 23,780').text='合計 23,790';
const richNumeric=app.selectDiffSemanticResult(richStableText,richNumericAfter,WIDTH,HEIGHT,
  {rowStructureAdjusted:false,alignmentAdjusted:false,fallbackUsed:false},manySparseTextNoise);
assert.equal(richNumeric.mode,'numeric','two changed values stay detectable among paragraphs and multiple tables');
assert.equal(richNumeric.regions.length,2,'only the two changed numeric fragments are highlighted');

const richColorRaster=analyze(richReport(),richReport({colorCell:true}),WIDTH,HEIGHT);
assert.ok(richColorRaster.regions.length>=1&&richColorRaster.regions.length<=2,'font-color edit remains localized on a rich report');
assert.ok(richColorRaster.regions.every(region=>region.width<.12&&region.height<.08),'font-color edit does not spread across its row or page');
assert.equal(app.shouldSuppressDiffRasterNoise(richStableText,richStableText,richColorRaster,richColorRaster.regions,WIDTH,HEIGHT),false,
  'actual rich-report font-color pixels are not classified as raster noise');

console.log('diff regression tests OK');
