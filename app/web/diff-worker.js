'use strict';
const PIXEL_THRESHOLD=30;
const BLOCK=4;
const MIN_PIXELS=28;
const PADDING=5;
const MAX_REGIONS=100;

function pixelDifference(before,after,beforeIndex,afterIndex){
  const dr=Math.abs(before[beforeIndex]-after[afterIndex]);
  const dg=Math.abs(before[beforeIndex+1]-after[afterIndex+1]);
  const db=Math.abs(before[beforeIndex+2]-after[afterIndex+2]);
  return Math.max(dr,dg,db);
}
function luminance(data,index){
  return (77*data[index]+150*data[index+1]+29*data[index+2])>>8;
}
const ROW_STEP=4,ROW_BINS=48,ROW_ALIGNMENT_BAND=64;
function buildRowDescriptors(data,width,height){
  const rowCount=Math.ceil(height/ROW_STEP),values=new Float32Array(rowCount*ROW_BINS),ink=new Float32Array(rowCount);
  for(let row=0;row<rowCount;row++){
    const y0=row*ROW_STEP,y1=Math.min(height,y0+ROW_STEP);
    let total=0;
    for(let bin=0;bin<ROW_BINS;bin++){
      const x0=Math.floor(bin*width/ROW_BINS),x1=Math.max(x0+1,Math.floor((bin+1)*width/ROW_BINS));
      let sum=0,count=0;
      for(let y=y0;y<y1;y++)for(let x=x0;x<x1;x+=2){sum+=255-luminance(data,(y*width+x)*4);count++;}
      const value=count?sum/count:0;values[row*ROW_BINS+bin]=value;total+=value;
    }
    ink[row]=total/ROW_BINS;
  }
  return {rowCount,values,ink};
}
function rowDistance(a,b,ai,bi){
  let total=0;
  const a0=ai*ROW_BINS,b0=bi*ROW_BINS;
  for(let k=0;k<ROW_BINS;k++)total+=Math.abs(a.values[a0+k]-b.values[b0+k]);
  return total/ROW_BINS+Math.abs(a.ink[ai]-b.ink[bi])*.35;
}
function rowGapCost(descriptors,index){return 5+Math.min(24,descriptors.ink[index]*.12);}
function smoothMappingScore(a,b,scale,offset,split=-1,jump=0){
  let total=0,count=0;
  for(let i=0;i<a.rowCount;i+=3){
    const j=Math.round(i*scale+offset+(split>=0&&i>=split?jump:0));
    if(j<0||j>=b.rowCount)continue;
    if(a.ink[i]<2&&b.ink[j]<2)continue;
    total+=Math.min(45,rowDistance(a,b,i,j));count++;
  }
  return count?total/count:Number.MAX_VALUE;
}
function chooseSmoothRowMapping(a,b){
  let best={scale:1,offset:0,split:-1,jump:0,score:smoothMappingScore(a,b,1,0)};
  best.fitness=best.score;
  for(let scaleStep=-6;scaleStep<=6;scaleStep++){
    const scale=1+scaleStep*.01;
    for(let offset=-6;offset<=6;offset+=2){
      const score=smoothMappingScore(a,b,scale,offset);
      const fitness=score+Math.abs(scale-1)*20;
      if(fitness<best.fitness)best={scale,offset,split:-1,jump:0,score,fitness};
    }
  }
  const coarse=best;
  for(let scaleStep=-4;scaleStep<=4;scaleStep++){
    const scale=coarse.scale+scaleStep*.0025;
    for(let offset=coarse.offset-2;offset<=coarse.offset+2;offset++){
      const score=smoothMappingScore(a,b,scale,offset);
      const fitness=score+Math.abs(scale-1)*20;
      if(fitness<best.fitness)best={scale,offset,split:-1,jump:0,score,fitness};
    }
  }
  const base=best,minSplit=Math.floor(a.rowCount*.1),maxSplit=Math.ceil(a.rowCount*.9);
  const scaleCandidates=[base.scale-.02,base.scale-.01,base.scale,base.scale+.01,base.scale+.02,1]
    .map(value=>Math.max(.94,Math.min(1.06,Math.round(value*1000)/1000)))
    .filter((value,index,array)=>array.indexOf(value)===index);
  for(const candidateScale of scaleCandidates){
    const offsetStats=new Map();
    for(let offset=-12;offset<=12;offset++){
      const sums=new Float32Array(a.rowCount+1),counts=new Uint16Array(a.rowCount+1);
      for(let i=0;i<a.rowCount;i++){
        sums[i+1]=sums[i];counts[i+1]=counts[i];
        const j=Math.round(i*candidateScale+offset);
        if(j<0||j>=b.rowCount||(a.ink[i]<2&&b.ink[j]<2))continue;
        sums[i+1]+=Math.min(45,rowDistance(a,b,i,j));counts[i+1]++;
      }
      offsetStats.set(offset,{sums,counts});
    }
    for(let split=minSplit;split<=maxSplit;split+=4)for(let offset1=-12;offset1<=12;offset1++)for(let offset2=-12;offset2<=12;offset2++){
      const first=offsetStats.get(offset1),second=offsetStats.get(offset2);
      const sum=first.sums[split]+second.sums[a.rowCount]-second.sums[split];
      const count=first.counts[split]+second.counts[a.rowCount]-second.counts[split];
      if(!count)continue;
      const jump=offset2-offset1,score=sum/count,fitness=score+Math.abs(candidateScale-1)*20+Math.abs(jump)*1.5;
      if(fitness<best.fitness)best={scale:candidateScale,offset:offset1,split,jump,score,fitness};
    }
  }
  const mapping=new Float32Array(a.rowCount);
  for(let i=0;i<a.rowCount;i++)mapping[i]=i*best.scale+best.offset+(best.split>=0&&i>=best.split?best.jump:0);
  return {...best,mapping};
}
function choosePixelRowMapping(before,after,width,height,rowCount){
  let best={scale:1,offset:0,split:-1,jump:0,score:Number.MAX_VALUE,fitness:Number.MAX_VALUE};
  const offsetValues=[];for(let offset=-32;offset<=32;offset+=4)offsetValues.push(offset);
  for(let scaleStep=-6;scaleStep<=6;scaleStep++){
    const scale=1+scaleStep*.01,stats=new Map();
    for(const offset of offsetValues){
      const sums=new Float32Array(rowCount+1),counts=new Uint16Array(rowCount+1);
      for(let row=0;row<rowCount;row++){
        sums[row+1]=sums[row];counts[row+1]=counts[row];
        const y=Math.min(height-1,row*ROW_STEP),ay=Math.round(y*scale+offset);
        if(ay<0||ay>=height)continue;
        let total=0,count=0;
        for(let x=8;x<width-8;x+=12){
          const bi=(y*width+x)*4,ai=(ay*width+x)*4;
          if(luminance(before,bi)>247&&luminance(after,ai)>247)continue;
          total+=Math.min(45,pixelDifference(before,after,bi,ai));count++;
        }
        if(count){sums[row+1]+=total/count;counts[row+1]++;}
      }
      stats.set(offset,{sums,counts});
      const count=counts[rowCount],score=count?sums[rowCount]/count:Number.MAX_VALUE,fitness=score+Math.abs(scale-1)*20+Math.abs(offset)*.04;
      if(fitness<best.fitness)best={scale,offset,split:-1,jump:0,score,fitness};
    }
    const minSplit=Math.floor(rowCount*.1),maxSplit=Math.ceil(rowCount*.9);
    for(let split=minSplit;split<=maxSplit;split+=8)for(const offset1 of offsetValues)for(const offset2 of offsetValues){
      const first=stats.get(offset1),second=stats.get(offset2);
      const sum=first.sums[split]+second.sums[rowCount]-second.sums[split];
      const count=first.counts[split]+second.counts[rowCount]-second.counts[split];
      if(!count)continue;
      const jump=offset2-offset1,score=sum/count,fitness=score+Math.abs(scale-1)*20+Math.abs(jump)*.08+(Math.abs(offset1)+Math.abs(offset2))*.02;
      if(fitness<best.fitness)best={scale,offset:offset1,split,jump,score,fitness};
    }
  }
  const mapping=new Float32Array(rowCount);
  for(let row=0;row<rowCount;row++){
    const y=row*ROW_STEP;
    mapping[row]=(y*best.scale+best.offset+(best.split>=0&&row>=best.split?best.jump:0))/ROW_STEP;
  }
  return {...best,mapping};
}
function alignRows(before,after,width,height){
  const a=buildRowDescriptors(before,width,height),b=buildRowDescriptors(after,width,height);
  const n=a.rowCount,m=b.rowCount,stride=m+1,size=(n+1)*stride,inf=1e20;
  const cost=new Float32Array(size),direction=new Uint8Array(size);cost.fill(inf);cost[0]=0;
  for(let i=0;i<=n;i++){
    const minJ=Math.max(0,i-ROW_ALIGNMENT_BAND),maxJ=Math.min(m,i+ROW_ALIGNMENT_BAND);
    for(let j=minJ;j<=maxJ;j++){
      if(!i&&!j)continue;
      const index=i*stride+j;
      let best=inf,dir=0;
      if(i&&j){
        const candidate=cost[(i-1)*stride+j-1]+rowDistance(a,b,i-1,j-1);
        if(candidate<best){best=candidate;dir=1;}
      }
      if(i&&Math.abs((i-1)-j)<=ROW_ALIGNMENT_BAND){
        const candidate=cost[(i-1)*stride+j]+rowGapCost(a,i-1);
        if(candidate<best-.001){best=candidate;dir=2;}
      }
      if(j&&Math.abs(i-(j-1))<=ROW_ALIGNMENT_BAND){
        const candidate=cost[i*stride+j-1]+rowGapCost(b,j-1);
        if(candidate<best-.001){best=candidate;dir=3;}
      }
      cost[index]=best;direction[index]=dir;
    }
  }
  const mapping=new Float32Array(n);mapping.fill(-1);
  const inserted=[],deleted=[],pairs=[];
  let i=n,j=m;
  while(i>0||j>0){
    const dir=direction[i*stride+j];
    if(dir===1){mapping[i-1]=j-1;pairs.push([i-1,j-1]);i--;j--;}
    else if(dir===2){deleted.push(i-1);i--;}
    else if(dir===3){inserted.push(j-1);j--;}
    else{if(i&&j){mapping[i-1]=j-1;pairs.push([i-1,j-1]);i--;j--;}else if(i)i--;else j--;}
  }
  pairs.reverse();
  let previous=-1,nextPair=0;
  for(let row=0;row<n;row++){
    if(mapping[row]>=0){previous=row;while(nextPair<pairs.length&&pairs[nextPair][0]<=row)nextPair++;continue;}
    const next=nextPair<pairs.length?pairs[nextPair][0]:-1;
    if(previous>=0&&next>=0){
      const t=(row-previous)/(next-previous);mapping[row]=mapping[previous]+(mapping[next]-mapping[previous])*t;
    }else if(previous>=0)mapping[row]=mapping[previous]+(row-previous);
    else if(next>=0)mapping[row]=mapping[next]-(next-row);
    else mapping[row]=row;
  }
  let maxLocalShiftJump=0;
  for(let row=1;row<n;row++)maxLocalShiftJump=Math.max(maxLocalShiftJump,Math.abs((mapping[row]-mapping[row-1])-1)*ROW_STEP);
  const first=pairs[0]||[0,0],last=pairs[pairs.length-1]||[Math.max(1,n-1),Math.max(1,m-1)];
  let scaleY=(last[1]-first[1])/Math.max(1,last[0]-first[0]),alignmentMode='sequence';
  const sequenceScore=(()=>{
    let total=0,count=0;
    for(let row=0;row<n;row+=3){
      const target=Math.max(0,Math.min(m-1,Math.round(mapping[row])));
      if(a.ink[row]<2&&b.ink[target]<2)continue;
      total+=Math.min(45,rowDistance(a,b,row,target));count++;
    }
    return count?total/count:Number.MAX_VALUE;
  })();
  const smooth=chooseSmoothRowMapping(a,b);
  if(smooth.score<2||smooth.fitness<=sequenceScore*.96){
    mapping.set(smooth.mapping);scaleY=smooth.scale;alignmentMode=smooth.split>=0?'scale-and-row-shift':'scale';
    inserted.length=0;deleted.length=0;
    maxLocalShiftJump=Math.abs(smooth.jump)*ROW_STEP;
  }
  const pixel=choosePixelRowMapping(before,after,width,height,n);
  if(pixel.score<12||pixel.fitness<=sequenceScore*.9){
    mapping.set(pixel.mapping);scaleY=pixel.scale;alignmentMode=pixel.split>=0?'pixel-scale-and-row-shift':'pixel-scale';
    inserted.length=0;deleted.length=0;
    maxLocalShiftJump=Math.abs(pixel.jump);
  }
  return {a,b,mapping,inserted,deleted,scaleY,maxLocalShiftJump,alignmentMode};
}
function mappedCoarseRowY(alignment,y,height){
  const position=y/ROW_STEP,index=Math.max(0,Math.min(alignment.mapping.length-1,Math.floor(position))),next=Math.min(alignment.mapping.length-1,index+1),fraction=position-index;
  return Math.max(0,Math.min(height-1,Math.round((alignment.mapping[index]*(1-fraction)+alignment.mapping[next]*fraction)*ROW_STEP)));
}
function mappedRowY(alignment,y,height){
  if(alignment.pixelMapping&&y>=0&&y<alignment.pixelMapping.length)return alignment.pixelMapping[y];
  return mappedCoarseRowY(alignment,y,height);
}
function chooseHorizontalAlignment(before,after,width,height,rowAlignment){
  let bestX=0,bestScore=Number.MAX_VALUE;
  for(let dx=-6;dx<=6;dx++){
    let total=0,count=0;
    for(let y=8;y<height-8;y+=12){
      const ay=mappedRowY(rowAlignment,y,height);
      for(let x=8;x<width-8;x+=12){
        const ax=x+dx;if(ax<1||ax>=width-1)continue;
        const bi=(y*width+x)*4,ai=(ay*width+ax)*4;
        if(luminance(before,bi)>247&&luminance(after,ai)>247)continue;
        total+=pixelDifference(before,after,bi,ai);count++;
      }
    }
    const score=count?total/count:Number.MAX_VALUE;
    if(score<bestScore-.001||(Math.abs(score-bestScore)<=.001&&Math.abs(dx)<Math.abs(bestX))){bestScore=score;bestX=dx;}
  }
  return bestX;
}
function refinePixelRowMapping(before,after,width,height,rowAlignment,offsetX){
  const mapping=new Int32Array(height);
  for(let y=0;y<height;y+=2){
    const expected=mappedCoarseRowY(rowAlignment,y,height);
    let bestY=expected,bestScore=Number.MAX_VALUE,foundInk=false;
    for(let candidate=Math.max(0,expected-3);candidate<=Math.min(height-1,expected+3);candidate++){
      let total=0,count=0;
      for(let x=8;x<width-8;x+=8){
        const ax=x+offsetX;if(ax<0||ax>=width)continue;
        const bi=(y*width+x)*4,ai=(candidate*width+ax)*4;
        if(luminance(before,bi)>247&&luminance(after,ai)>247)continue;
        total+=pixelDifference(before,after,bi,ai);count++;
      }
      if(!count)continue;
      foundInk=true;
      const score=total/count+Math.abs(candidate-expected)*.8;
      if(score<bestScore){bestScore=score;bestY=candidate;}
    }
    mapping[y]=foundInk?bestY:expected;
  }
  for(let y=1;y<height;y+=2)mapping[y]=Math.round((mapping[y-1]+mapping[Math.min(height-1,y+1)])/2);
  for(let y=1;y<height;y++)mapping[y]=Math.max(mapping[y-1],mapping[y]);
  rowAlignment.pixelMapping=mapping;
}
function tolerantPixelDifference(before,after,width,height,x,y,ax,ay){
  const bi=(y*width+x)*4;
  const exactIndex=(ay*width+ax)*4;
  let best=pixelDifference(before,after,bi,exactIndex);
  if(best<=PIXEL_THRESHOLD)return best;
  for(let oy=-1;oy<=1;oy++)for(let ox=-1;ox<=1;ox++){
    if(!ox&&!oy)continue;
    const sx=ax+ox,sy=ay+oy;
    if(sx<0||sx>=width||sy<0||sy>=height)continue;
    best=Math.min(best,pixelDifference(before,after,bi,(sy*width+sx)*4));
  }
  return best;
}
function mergeNearbyComponents(components,width,height){
  const sorted=components.slice().sort((a,b)=>a.minY-b.minY||a.minX-b.minX),merged=[];
  for(const component of sorted){
    let target=null;
    for(let i=merged.length-1;i>=Math.max(0,merged.length-12);i--){
      const candidate=merged[i],verticalGap=component.minY-candidate.maxY-1;
      if(verticalGap>12)break;
      const overlap=Math.max(0,Math.min(candidate.maxX,component.maxX)-Math.max(candidate.minX,component.minX)+1);
      const minWidth=Math.max(1,Math.min(candidate.maxX-candidate.minX+1,component.maxX-component.minX+1));
      const combinedHeight=Math.max(candidate.maxY,component.maxY)-Math.min(candidate.minY,component.minY)+1;
      if(verticalGap<=8&&overlap/minWidth>=.25&&combinedHeight<=height*.25){target=candidate;break;}
      if(verticalGap<=12&&(candidate.maxX-candidate.minX+1)>width*.55&&(component.maxX-component.minX+1)>width*.55&&combinedHeight<=height*.2){target=candidate;break;}
    }
    if(target){
      target.minX=Math.min(target.minX,component.minX);target.minY=Math.min(target.minY,component.minY);
      target.maxX=Math.max(target.maxX,component.maxX);target.maxY=Math.max(target.maxY,component.maxY);target.pixels+=component.pixels;
    }else merged.push({...component});
  }
  return merged;
}
function analyzeBrowserDiff(before,after,width,height){
  const rowAlignment=alignRows(before,after,width,height);
  let offsetX=chooseHorizontalAlignment(before,after,width,height,rowAlignment);
  refinePixelRowMapping(before,after,width,height,rowAlignment,offsetX);
  offsetX=chooseHorizontalAlignment(before,after,width,height,rowAlignment);
  const gridWidth=Math.ceil(width/BLOCK),gridHeight=Math.ceil(height/BLOCK),cellCount=gridWidth*gridHeight;
  const mask=new Uint8Array(cellCount),counts=new Uint32Array(cellCount);
  let totalChanged=0;
  for(let gy=0;gy<gridHeight;gy++)for(let gx=0;gx<gridWidth;gx++){
    let changed=0;
    const x0=gx*BLOCK,y0=gy*BLOCK,x1=Math.min(width,x0+BLOCK),y1=Math.min(height,y0+BLOCK);
    for(let y=y0;y<y1;y++)for(let x=x0;x<x1;x++){
      const ax=x+offsetX,ay=mappedRowY(rowAlignment,y,height),bi=(y*width+x)*4;
      let different=false;
      if(ax<0||ax>=width||ay<0||ay>=height){
        different=before[bi]<248||before[bi+1]<248||before[bi+2]<248;
      }else different=tolerantPixelDifference(before,after,width,height,x,y,ax,ay)>PIXEL_THRESHOLD;
      if(different)changed++;
    }
    const index=gy*gridWidth+gx;
    counts[index]=changed;totalChanged+=changed;
    if(changed>=Math.max(2,Math.floor((x1-x0)*(y1-y0)*.2)))mask[index]=1;
  }
  // Rows that have no counterpart are the human-visible insertion/deletion bands.
  const gapRows=[];
  for(const row of rowAlignment.deleted){if(rowAlignment.a.ink[row]>4)gapRows.push({beforeRow:row,source:'before'});}
  for(const row of rowAlignment.inserted){
    if(rowAlignment.b.ink[row]<=4)continue;
    let beforeRow=0,bestDistance=Number.MAX_VALUE;
    for(let i=0;i<rowAlignment.mapping.length;i++){
      const distance=Math.abs(rowAlignment.mapping[i]-row);
      if(distance<bestDistance){bestDistance=distance;beforeRow=i;}
    }
    gapRows.push({beforeRow,afterRow:row,source:'after'});
  }
  for(const gap of gapRows){
    const centerY=Math.min(height-1,gap.beforeRow*ROW_STEP),half=6;
    let minGX=gridWidth,maxGX=-1;
    for(let y=Math.max(0,centerY-half);y<Math.min(height,centerY+half);y+=4)for(let x=8;x<width-8;x+=8){
      const bi=(y*width+x)*4,ay=gap.source==='after'?Math.min(height-1,(gap.afterRow||0)*ROW_STEP):mappedRowY(rowAlignment,y,height),ai=(ay*width+Math.max(0,Math.min(width-1,x+offsetX)))*4;
      if(luminance(before,bi)<245||luminance(after,ai)<245){minGX=Math.min(minGX,Math.floor(x/BLOCK));maxGX=Math.max(maxGX,Math.floor(x/BLOCK));}
    }
    if(maxGX<minGX){minGX=2;maxGX=gridWidth-3;}
    for(let gy=Math.max(0,Math.floor((centerY-half)/BLOCK));gy<=Math.min(gridHeight-1,Math.floor((centerY+half)/BLOCK));gy++)for(let gx=minGX;gx<=maxGX;gx++){
      const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
    }
  }
  // Grow horizontally to combine one logical row, but not vertically across many spreadsheet rows.
  const grown=new Uint8Array(cellCount),gapX=2;
  for(let gy=0;gy<gridHeight;gy++)for(let gx=0;gx<gridWidth;gx++){
    const index=gy*gridWidth+gx;if(!mask[index])continue;
    for(let ox=-gapX;ox<=gapX;ox++){
      const nx=gx+ox;if(nx>=0&&nx<gridWidth)grown[gy*gridWidth+nx]=1;
    }
  }
  const visited=new Uint8Array(cellCount),queue=new Int32Array(cellCount),raw=[];
  for(let start=0;start<cellCount;start++){
    if(!grown[start]||visited[start])continue;
    let head=0,tail=0;queue[tail++]=start;visited[start]=1;
    let minX=width,minY=height,maxX=-1,maxY=-1,pixels=0;
    while(head<tail){
      const current=queue[head++],cx=current%gridWidth,cy=Math.floor(current/gridWidth);
      if(mask[current]){
        const x0=cx*BLOCK,y0=cy*BLOCK;
        minX=Math.min(minX,x0);minY=Math.min(minY,y0);maxX=Math.max(maxX,Math.min(width,x0+BLOCK)-1);maxY=Math.max(maxY,Math.min(height,y0+BLOCK)-1);pixels+=counts[current];
      }
      for(let oy=-1;oy<=1;oy++)for(let ox=-1;ox<=1;ox++){
        if(!ox&&!oy)continue;
        const nx=cx+ox,ny=cy+oy;if(nx<0||nx>=gridWidth||ny<0||ny>=gridHeight)continue;
        const next=ny*gridWidth+nx;if(grown[next]&&!visited[next]){visited[next]=1;queue[tail++]=next;}
      }
    }
    if(pixels<MIN_PIXELS||maxX<minX||maxY<minY)continue;
    raw.push({minX:Math.max(0,minX-PADDING),minY:Math.max(0,minY-PADDING),maxX:Math.min(width-1,maxX+PADDING),maxY:Math.min(height-1,maxY+PADDING),pixels});
  }
  let components=mergeNearbyComponents(raw,width,height);
  if(components.length>MAX_REGIONS)components.sort((a,b)=>b.pixels-a.pixels).splice(MAX_REGIONS);
  components.sort((a,b)=>a.minY-b.minY||a.minX-b.minX);
  const regions=components.map((component,index)=>({
    regionId:`browser-r${String(index+1).padStart(4,'0')}`,kind:'modified',x:component.minX/width,y:component.minY/height,
    width:Math.max(1,component.maxX-component.minX+1)/width,height:Math.max(1,component.maxY-component.minY+1)/height,
    confidence:Math.max(.55,Math.min(1,component.pixels/Math.max(MIN_PIXELS,(component.maxX-component.minX+1)*(component.maxY-component.minY+1)))),pixelCount:component.pixels
  }));
  const alignmentAdjusted=Math.abs(rowAlignment.scaleY-1)>.003||Math.abs(offsetX)>1||gapRows.length>0||rowAlignment.maxLocalShiftJump>=4;
  return {regions,changedRatio:totalChanged/Math.max(1,width*height),offsetX,offsetY:0,scaleY:rowAlignment.scaleY,maxLocalShiftJump:rowAlignment.maxLocalShiftJump,alignmentMode:rowAlignment.alignmentMode,alignmentAdjusted};
}
self.onmessage=function handleDiffWorkerMessage(event){
  const payload=event.data||{},id=payload.id;
  try{
    const width=Number(payload.width||0),height=Number(payload.height||0);
    if(width<=0||height<=0)throw new Error('画像サイズが不正です。');
    const before=new Uint8ClampedArray(payload.before),after=new Uint8ClampedArray(payload.after);
    if(before.length!==width*height*4||after.length!==width*height*4)throw new Error('画像データが不正です。');
    const result=analyzeBrowserDiff(before,after,width,height);self.postMessage(Object.assign({id},result));
  }catch(error){self.postMessage({id,error:String(error?.message||error)});}
};
