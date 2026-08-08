import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';

const manifestPath=path.resolve(process.argv[2]||'');
const renderRoot=path.resolve(process.argv[3]||'');
if(!fs.existsSync(manifestPath)||!fs.existsSync(renderRoot))throw new Error('usage: node pdf-diff-corpus.mjs <manifest> <render-dir>');
const manifest=JSON.parse(fs.readFileSync(manifestPath,'utf8'));
const workerCode=fs.readFileSync(new URL('../app/web/diff-worker.js',import.meta.url),'utf8');

function readBmp(file){
  const data=fs.readFileSync(file);
  assert.equal(data.toString('ascii',0,2),'BM',`${file} is not a BMP`);
  const offset=data.readUInt32LE(10),width=data.readInt32LE(18),signedHeight=data.readInt32LE(22),bits=data.readUInt16LE(28);
  assert.ok(width>0&&signedHeight!==0&&[24,32].includes(bits),`unsupported BMP layout: ${width}x${signedHeight} ${bits}bpp`);
  const height=Math.abs(signedHeight),bottomUp=signedHeight>0,bytes=bits/8,stride=Math.ceil(width*bytes/4)*4;
  const rgba=new Uint8ClampedArray(width*height*4);
  for(let y=0;y<height;y++){
    const sourceY=bottomUp?height-1-y:y;
    for(let x=0;x<width;x++){
      const source=offset+sourceY*stride+x*bytes,target=(y*width+x)*4;
      rgba[target]=data[source+2];rgba[target+1]=data[source+1];rgba[target+2]=data[source];rgba[target+3]=bits===32?data[source+3]:255;
    }
  }
  return {width,height,rgba};
}

function analyze(before,after,width,height){
  let result=null;const self={postMessage:value=>{result=value;}};
  vm.runInNewContext(workerCode,{self,Uint8Array,Uint8ClampedArray,Uint16Array,Uint32Array,Int32Array,Float32Array,Math,Number,Array,Map,Set,Object,String,Error});
  self.onmessage({data:{id:1,width,height,before:before.buffer,after:after.buffer}});
  if(result?.error)throw new Error(result.error);return result;
}

function overlapArea(a,b){
  const left=Math.max(a.x,b.x),top=Math.max(a.y,b.y),right=Math.min(a.x+a.width,b.x+b.width),bottom=Math.min(a.y+a.height,b.y+b.height);
  return Math.max(0,right-left)*Math.max(0,bottom-top);
}
function regionBoxes(region){return [region,region.before,region.after].filter(Boolean);}
function hitsExpected(region,expected){
  return regionBoxes(region).some(box=>{
    const overlap=overlapArea(box,expected),expectedArea=expected.width*expected.height,boxArea=box.width*box.height;
    const centerX=expected.x+expected.width/2,centerY=expected.y+expected.height/2;
    return overlap/Math.max(Math.min(expectedArea,boxArea),1e-9)>=.08||(centerX>=box.x&&centerX<=box.x+box.width&&centerY>=box.y&&centerY<=box.y+box.height);
  });
}

const metrics=[];
for(const testCase of manifest.cases){
  const before=readBmp(path.join(renderRoot,`${path.parse(testCase.before).name}-${testCase.page}.bmp`));
  const after=readBmp(path.join(renderRoot,`${path.parse(testCase.after).name}-${testCase.page}.bmp`));
  assert.equal(before.width,after.width);assert.equal(before.height,after.height);
  const result=analyze(before.rgba,after.rgba,before.width,before.height),regions=result.regions||[],expected=testCase.expected||[];
  const recalled=expected.filter(box=>regions.some(region=>hitsExpected(region,box))).length;
  const relevant=regions.filter(region=>expected.some(box=>hitsExpected(region,box))).length;
  const recall=expected.length?recalled/expected.length:regions.length?0:1;
  const precision=regions.length?relevant/regions.length:expected.length?0:1;
  metrics.push({name:testCase.name,width:before.width,height:before.height,expected:expected.length,predicted:regions.length,recall,precision});
  console.log(`${testCase.name}: expected=${expected.length} predicted=${regions.length} recall=${recall.toFixed(3)} precision=${precision.toFixed(3)}`);
  if(recall<manifest.minimumRecall||precision<manifest.minimumPrecision)console.log(JSON.stringify({expected,regions},null,2));
  assert.ok(recall>=manifest.minimumRecall,`${testCase.name} recall ${recall.toFixed(3)} < ${manifest.minimumRecall}`);
  assert.ok(precision>=manifest.minimumPrecision,`${testCase.name} precision ${precision.toFixed(3)} < ${manifest.minimumPrecision}`);
  const stable=analyze(before.rgba,new Uint8ClampedArray(before.rgba),before.width,before.height);
  assert.equal(stable.regions.length,0,`${testCase.name} unchanged PDF produced false-positive regions`);
}
console.log(JSON.stringify({ok:true,minimumRecall:manifest.minimumRecall,minimumPrecision:manifest.minimumPrecision,cases:metrics},null,2));
