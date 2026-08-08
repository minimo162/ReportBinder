'use strict';
const PIXEL_THRESHOLD=30;
const BLOCK=4;
const MIN_PIXELS=28;
const PADDING=5;
const MAX_REGIONS=100;
const MAX_LAYOUT_REGIONS=12;

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
const COLUMN_STEP=4,COLUMN_BINS=48,MIN_LAYOUT_SCALE=.82,MAX_LAYOUT_SCALE=1.18;
function clearlyImproves(candidateScore,identityScore,ratio=.76){
  if(!Number.isFinite(candidateScore)||!Number.isFinite(identityScore))return false;
  return candidateScore<2.5&&identityScore>3.5||candidateScore<identityScore*ratio;
}
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
function buildColumnDescriptors(data,width,height,verticalBand=null){
  const columnCount=Math.ceil(width/COLUMN_STEP),values=new Float32Array(columnCount*COLUMN_BINS),ink=new Float32Array(columnCount);
  const minY=verticalBand?Math.max(0,Math.floor(verticalBand.start)):0,maxY=verticalBand?Math.min(height-1,Math.ceil(verticalBand.last)):height-1,bandHeight=Math.max(1,maxY-minY+1);
  const minX=verticalBand&&Number.isFinite(verticalBand.minX)?Math.max(0,verticalBand.minX-PADDING):0,maxX=verticalBand&&Number.isFinite(verticalBand.maxX)?Math.min(width-1,verticalBand.maxX+PADDING):width-1;
  for(let column=0;column<columnCount;column++){
    const x0=column*COLUMN_STEP,x1=Math.min(width,x0+COLUMN_STEP);
    if(x1-1<minX||x0>maxX)continue;
    let total=0;
    for(let bin=0;bin<COLUMN_BINS;bin++){
      const y0=minY+Math.floor(bin*bandHeight/COLUMN_BINS),y1=Math.min(maxY+1,Math.max(y0+1,minY+Math.floor((bin+1)*bandHeight/COLUMN_BINS)));
      let sum=0,count=0;
      for(let x=x0;x<x1;x++)for(let y=y0;y<y1;y+=2){sum+=255-luminance(data,(y*width+x)*4);count++;}
      const value=count?sum/count:0;values[column*COLUMN_BINS+bin]=value;total+=value;
    }
    ink[column]=total/COLUMN_BINS;
  }
  return {columnCount,values,ink};
}
function columnDistance(a,b,ai,bi){
  let total=0;const a0=ai*COLUMN_BINS,b0=bi*COLUMN_BINS;
  for(let k=0;k<COLUMN_BINS;k++)total+=Math.abs(a.values[a0+k]-b.values[b0+k]);
  return total/COLUMN_BINS+Math.abs(a.ink[ai]-b.ink[bi])*.35;
}
function smoothColumnScore(a,b,scale,offset,split=-1,jump=0){
  let total=0,count=0;
  for(let i=0;i<a.columnCount;i+=2){
    const j=Math.round(i*scale+offset+(split>=0&&i>=split?jump:0));
    if(j<0||j>=b.columnCount||(a.ink[i]<2&&b.ink[j]<2))continue;
    total+=Math.min(45,columnDistance(a,b,i,j));count++;
  }
  return count?total/count:Number.MAX_VALUE;
}
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
  for(let scaleStep=-8;scaleStep<=8;scaleStep++){
    const scale=1+scaleStep*.02;
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
    .map(value=>Math.max(MIN_LAYOUT_SCALE,Math.min(MAX_LAYOUT_SCALE,Math.round(value*1000)/1000)))
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
      const jump=offset2-offset1,score=sum/count,fitness=score+Math.abs(candidateScale-1)*20+Math.abs(jump)*.15;
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
  for(let scaleStep=-8;scaleStep<=8;scaleStep++){
    const scale=1+scaleStep*.02,stats=new Map();
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
          let difference=pixelDifference(before,after,bi,ai);
          if(ay>0)difference=Math.min(difference,pixelDifference(before,after,bi,ai-width*4));
          if(ay+1<height)difference=Math.min(difference,pixelDifference(before,after,bi,ai+width*4));
          total+=Math.min(45,difference);count++;
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
  const identityScore=(()=>{
    let total=0,count=0;
    for(let row=0;row<rowCount;row++){
      const y=Math.min(height-1,row*ROW_STEP);
      let rowTotal=0,rowCounted=0;
      for(let x=8;x<width-8;x+=12){
        const index=(y*width+x)*4;
        if(luminance(before,index)>247&&luminance(after,index)>247)continue;
        let difference=pixelDifference(before,after,index,index);
        if(y>0)difference=Math.min(difference,pixelDifference(before,after,index,index-width*4));
        if(y+1<height)difference=Math.min(difference,pixelDifference(before,after,index,index+width*4));
        rowTotal+=Math.min(45,difference);rowCounted++;
      }
      if(rowCounted){total+=rowTotal/rowCounted;count++;}
    }
    return count?total/count:Number.MAX_VALUE;
  })();
  return {...best,mapping,identityScore};
}
function splitHasLayoutSupport(descriptors,split){
  if(split<0)return true;
  let before=0,after=0;
  for(let i=0;i<descriptors.rowCount;i++)if(descriptors.ink[i]>=2){if(i<split)before++;else after++;}
  const total=before+after,minimum=Math.max(6,Math.ceil(total*.12));
  return before>=minimum&&after>=minimum;
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
  let sequenceMaxLocalShiftJump=0;
  for(let row=1;row<n;row++)sequenceMaxLocalShiftJump=Math.max(sequenceMaxLocalShiftJump,Math.abs((mapping[row]-mapping[row-1])-1)*ROW_STEP);
  const first=pairs[0]||[0,0],last=pairs[pairs.length-1]||[Math.max(1,n-1),Math.max(1,m-1)];
  const sequenceScale=(last[1]-first[1])/Math.max(1,last[0]-first[0]);
  const sequenceScore=(()=>{
    let total=0,count=0;
    for(let row=0;row<n;row+=3){
      const target=Math.max(0,Math.min(m-1,Math.round(mapping[row])));
      if(a.ink[row]<2&&b.ink[target]<2)continue;
      total+=Math.min(45,rowDistance(a,b,row,target));count++;
    }
    return count?total/count:Number.MAX_VALUE;
  })();
  const identityScore=smoothMappingScore(a,b,1,0);
  const sequenceMapping=new Float32Array(mapping),sequenceInserted=inserted.slice(),sequenceDeleted=deleted.slice();
  for(let row=0;row<n;row++)mapping[row]=row;
  inserted.length=0;deleted.length=0;
  let scaleY=1,maxLocalShiftJump=0,alignmentMode='identity',adjusted=false,splitRow=-1,jumpPixels=0;
  let chosenRelative=1;
  const sequenceGapCount=sequenceInserted.length+sequenceDeleted.length;
  if(sequenceGapCount<=Math.max(8,Math.ceil(n*.08))&&sequenceMaxLocalShiftJump<=48&&clearlyImproves(sequenceScore,identityScore,.72)){
    mapping.set(sequenceMapping);inserted.push(...sequenceInserted);deleted.push(...sequenceDeleted);
    scaleY=sequenceScale;maxLocalShiftJump=sequenceMaxLocalShiftJump;alignmentMode='sequence';adjusted=true;
    chosenRelative=sequenceScore/Math.max(.1,identityScore);
  }
  const smooth=chooseSmoothRowMapping(a,b);
  const smoothRelative=smooth.score/Math.max(.1,identityScore);
  // A one-sided row descriptor can make a newly added object look like a layout shift by
  // mapping around it. Piecewise shifts therefore require pixel/sequence corroboration.
  if(smooth.split<0&&clearlyImproves(smooth.score,identityScore,.65)&&smoothRelative<chosenRelative*.98){
    mapping.set(smooth.mapping);scaleY=smooth.scale;alignmentMode=smooth.split>=0?'scale-and-row-shift':'scale';
    inserted.length=0;deleted.length=0;maxLocalShiftJump=Math.abs(smooth.jump)*ROW_STEP;adjusted=true;
    splitRow=smooth.split;jumpPixels=smooth.jump*ROW_STEP;chosenRelative=smoothRelative;
  }
  const pixel=choosePixelRowMapping(before,after,width,height,n);
  const pixelRelative=pixel.score/Math.max(.1,pixel.identityScore);
  if(splitHasLayoutSupport(a,pixel.split)&&clearlyImproves(pixel.score,pixel.identityScore,.7)&&(pixelRelative<chosenRelative*.98||alignmentMode==='identity')){
    mapping.set(pixel.mapping);scaleY=pixel.scale;alignmentMode=pixel.split>=0?'pixel-scale-and-row-shift':'pixel-scale';
    inserted.length=0;deleted.length=0;maxLocalShiftJump=Math.abs(pixel.jump);adjusted=true;
    splitRow=pixel.split;jumpPixels=pixel.jump;
  }
  return {a,b,mapping,inserted,deleted,scaleY,maxLocalShiftJump,alignmentMode,adjusted,splitRow,jumpPixels};
}
function mappedCoarseRowY(alignment,y,height){
  const position=y/ROW_STEP,index=Math.max(0,Math.min(alignment.mapping.length-1,Math.floor(position))),next=Math.min(alignment.mapping.length-1,index+1),fraction=position-index;
  return Math.max(0,Math.min(height-1,Math.round((alignment.mapping[index]*(1-fraction)+alignment.mapping[next]*fraction)*ROW_STEP)));
}
function mappedRowY(alignment,y,height,x=-1){
  if(Number.isFinite(alignment.activeMinY)&&Number.isFinite(alignment.activeMaxY)&&(y<alignment.activeMinY||y>alignment.activeMaxY))return Math.max(0,Math.min(height-1,y));
  if(x>=0&&Number.isFinite(alignment.activeMinX)&&Number.isFinite(alignment.activeMaxX)&&(x<alignment.activeMinX||x>alignment.activeMaxX))return Math.max(0,Math.min(height-1,y));
  if(alignment.pixelMapping&&y>=0&&y<alignment.pixelMapping.length)return alignment.pixelMapping[y];
  return mappedCoarseRowY(alignment,y,height);
}
function choosePixelColumnMapping(before,after,width,height,rowAlignment,verticalBand=null){
  const a=buildColumnDescriptors(before,width,height,verticalBand),b=buildColumnDescriptors(after,width,height,verticalBand),columnCount=a.columnCount;
  const identityScore=smoothColumnScore(a,b,1,0);
  if(!verticalBand){
    const mapping=new Float32Array(columnCount);for(let column=0;column<columnCount;column++)mapping[column]=column;
    return {scale:1,offset:0,split:-1,jump:0,structuralSplit:-1,structuralJump:0,score:identityScore,fitness:identityScore,mapping,identityScore,adjusted:false,a,b};
  }
  let best={scale:1,offset:0,split:-1,jump:0,score:identityScore,fitness:identityScore};
  for(let scaleStep=-8;scaleStep<=8;scaleStep++){
    const scale=1+scaleStep*.02;
    for(let offset=-24;offset<=24;offset+=2){
      const score=smoothColumnScore(a,b,scale,offset),fitness=score+Math.abs(scale-1)*4+Math.abs(offset)*.02;
      if(fitness<best.fitness)best={scale,offset,split:-1,jump:0,score,fitness};
    }
  }
  const coarse=best;
  for(let scaleStep=-4;scaleStep<=4;scaleStep++)for(let offset=coarse.offset-2;offset<=coarse.offset+2;offset++){
    const scale=coarse.scale+scaleStep*.0025,score=smoothColumnScore(a,b,scale,offset),fitness=score+Math.abs(scale-1)*4+Math.abs(offset)*.02;
    if(fitness<best.fitness)best={scale,offset,split:-1,jump:0,score,fitness};
  }
  const base=best,minSplit=Math.floor(columnCount*.08),maxSplit=Math.ceil(columnCount*.92),offsetValues=[];
  for(let offset=-24;offset<=24;offset++)offsetValues.push(offset);
  const scaleCandidates=[base.scale-.01,base.scale,base.scale+.01,1]
    .map(value=>Math.max(MIN_LAYOUT_SCALE,Math.min(MAX_LAYOUT_SCALE,Math.round(value*1000)/1000))).filter((value,index,array)=>array.indexOf(value)===index);
  let bestPiecewise=null;
  for(const scale of scaleCandidates){
    const stats=new Map();
    for(const offset of offsetValues){
      const sums=new Float32Array(columnCount+1),counts=new Uint16Array(columnCount+1);
      for(let i=0;i<columnCount;i++){
        sums[i+1]=sums[i];counts[i+1]=counts[i];const j=Math.round(i*scale+offset);
        if(j<0||j>=b.columnCount||(a.ink[i]<2&&b.ink[j]<2))continue;
        sums[i+1]+=Math.min(45,columnDistance(a,b,i,j));counts[i+1]++;
      }
      stats.set(offset,{sums,counts});
    }
    for(let split=minSplit;split<=maxSplit;split+=6)for(const offset1 of offsetValues)for(const offset2 of offsetValues){
      const first=stats.get(offset1),second=stats.get(offset2),sum=first.sums[split]+second.sums[columnCount]-second.sums[split];
      const count=first.counts[split]+second.counts[columnCount]-second.counts[split];if(!count)continue;
      const jump=offset2-offset1,score=sum/count,fitness=score+Math.abs(scale-1)*4+Math.abs(jump)*.04+(Math.abs(offset1)+Math.abs(offset2))*.01;
      const candidate={scale,offset:offset1,split,jump,score,fitness};
      if(!bestPiecewise||fitness<bestPiecewise.fitness)bestPiecewise=candidate;
      if(fitness<best.fitness)best=candidate;
    }
  }
  const scaleChange=Math.abs(best.scale-1);
  const requiredScoreRatio=verticalBand&&verticalBand.detection&&verticalBand.detection.startsWith('horizontal')?(scaleChange>.04?.985:.82):(scaleChange>.04?.94:.65);
  const adjusted=(scaleChange>.003||Math.abs(best.offset*COLUMN_STEP)>1||Math.abs(best.jump*COLUMN_STEP)>1)&&clearlyImproves(best.score,identityScore,requiredScoreRatio);
  const mapping=new Float32Array(columnCount);
  for(let column=0;column<columnCount;column++){
    mapping[column]=adjusted?column*best.scale+best.offset+(best.split>=0&&column>=best.split?best.jump:0):column;
  }
  if(!adjusted)best={scale:1,offset:0,split:-1,jump:0,score:identityScore,fitness:identityScore};
  const piecewiseClose=bestPiecewise&&Math.abs(bestPiecewise.jump*COLUMN_STEP)>=3&&
    clearlyImproves(bestPiecewise.score,identityScore,.96)&&bestPiecewise.fitness<=best.fitness*1.18+.8;
  const structuralSplit=best.split>=0?best.split:(piecewiseClose?bestPiecewise.split:-1);
  const structuralJump=best.split>=0?best.jump*COLUMN_STEP:(piecewiseClose?bestPiecewise.jump*COLUMN_STEP:0);
  return {...best,offset:best.offset*COLUMN_STEP,jump:best.jump*COLUMN_STEP,structuralSplit,structuralJump,mapping,identityScore,adjusted,a,b};
}
function mappedCoarseColumnX(alignment,x,width){
  const position=x/COLUMN_STEP,index=Math.max(0,Math.min(alignment.mapping.length-1,Math.floor(position))),next=Math.min(alignment.mapping.length-1,index+1),fraction=position-index;
  return Math.max(0,Math.min(width-1,Math.round((alignment.mapping[index]*(1-fraction)+alignment.mapping[next]*fraction)*COLUMN_STEP)));
}
function mappedColumnX(alignment,x,width,y=-1){
  if(y>=0&&Number.isFinite(alignment.activeMinY)&&Number.isFinite(alignment.activeMaxY)&&(y<alignment.activeMinY||y>alignment.activeMaxY))return Math.max(0,Math.min(width-1,x));
  if(Number.isFinite(alignment.activeMinX)&&Number.isFinite(alignment.activeMaxX)&&(x<alignment.activeMinX||x>alignment.activeMaxX))return Math.max(0,Math.min(width-1,x));
  if(alignment.pixelMapping&&x>=0&&x<alignment.pixelMapping.length)return alignment.pixelMapping[x];
  return mappedCoarseColumnX(alignment,x,width);
}
function refinePixelRowMapping(before,after,width,height,rowAlignment,columnAlignment){
  if(!rowAlignment.adjusted)return;
  const mapping=new Int32Array(height);
  for(let y=0;y<height;y+=2){
    const expected=mappedCoarseRowY(rowAlignment,y,height);
    let bestY=expected,bestScore=Number.MAX_VALUE,foundInk=false;
    for(let candidate=Math.max(0,expected-1);candidate<=Math.min(height-1,expected+1);candidate++){
      let total=0,count=0;
      for(let x=8;x<width-8;x+=8){
        const ax=mappedColumnX(columnAlignment,x,width,y);if(ax<0||ax>=width)continue;
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
function refinePixelColumnMapping(before,after,width,height,rowAlignment,columnAlignment){
  if(!columnAlignment.adjusted)return;
  const mapping=new Int32Array(width);
  for(let x=0;x<width;x+=2){
    if(Number.isFinite(columnAlignment.activeMinX)&&Number.isFinite(columnAlignment.activeMaxX)&&(x<columnAlignment.activeMinX||x>columnAlignment.activeMaxX)){mapping[x]=x;continue;}
    const expected=mappedCoarseColumnX(columnAlignment,x,width);
    let bestX=expected,bestScore=Number.MAX_VALUE,foundInk=false;
    for(let candidate=Math.max(0,expected-1);candidate<=Math.min(width-1,expected+1);candidate++){
      let total=0,count=0;
      const minY=Number.isFinite(columnAlignment.activeMinY)?Math.max(8,Math.floor(columnAlignment.activeMinY)):8;
      const maxY=Number.isFinite(columnAlignment.activeMaxY)?Math.min(height-8,Math.ceil(columnAlignment.activeMaxY)):height-8;
      for(let y=minY;y<maxY;y+=8){
        const ay=mappedRowY(rowAlignment,y,height,x),bi=(y*width+x)*4,ai=(ay*width+candidate)*4;
        if(luminance(before,bi)>247&&luminance(after,ai)>247)continue;
        total+=pixelDifference(before,after,bi,ai);count++;
      }
      if(!count)continue;
      foundInk=true;
      const score=total/count+Math.abs(candidate-expected)*.8;
      if(score<bestScore){bestScore=score;bestX=candidate;}
    }
    mapping[x]=foundInk?bestX:expected;
  }
  for(let x=1;x<width;x+=2)mapping[x]=Math.round((mapping[x-1]+mapping[Math.min(width-1,x+1)])/2);
  for(let x=1;x<width;x++)mapping[x]=Math.max(mapping[x-1],mapping[x]);
  columnAlignment.pixelMapping=mapping;
}
function tolerantPixelDifference(before,after,width,height,x,y,ax,ay){
  const bi=(y*width+x)*4;
  const exactIndex=(ay*width+ax)*4;
  const exact=pixelDifference(before,after,bi,exactIndex);
  if(exact<=PIXEL_THRESHOLD)return exact;
  let forward=exact,reverse=exact;
  for(let oy=-1;oy<=1;oy++)for(let ox=-1;ox<=1;ox++){
    if(!ox&&!oy)continue;
    const sx=ax+ox,sy=ay+oy;
    if(sx<0||sx>=width||sy<0||sy>=height)continue;
    forward=Math.min(forward,pixelDifference(before,after,bi,(sy*width+sx)*4));
    const bx=x+ox,by=y+oy;
    if(bx>=0&&bx<width&&by>=0&&by<height)reverse=Math.min(reverse,pixelDifference(before,after,(by*width+bx)*4,exactIndex));
  }
  // The one-way minimum can erase a genuine thin addition: a white source
  // pixel simply finds neighboring white beside the new glyph or rule.  Both
  // directions must have a nearby match before the difference is treated as
  // harmless sub-pixel raster movement.
  return Math.max(forward,reverse);
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
function collectMaskComponents(mask,counts,gridWidth,gridHeight,width,height,gapX=2){
  const cellCount=gridWidth*gridHeight,grown=new Uint8Array(cellCount);
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
  return raw;
}
function clearLongMaskRuns(mask,gridWidth,gridHeight){
  const horizontalMinimum=Math.max(12,Math.ceil(gridWidth*.18)),verticalMinimum=Math.max(12,Math.ceil(gridHeight*.18));
  for(let gy=0;gy<gridHeight;gy++){
    let start=-1;
    for(let gx=0;gx<=gridWidth;gx++){
      const active=gx<gridWidth&&mask[gy*gridWidth+gx];
      if(active&&start<0)start=gx;
      if((!active||gx===gridWidth)&&start>=0){if(gx-start>=horizontalMinimum)for(let x=start;x<gx;x++)mask[gy*gridWidth+x]=0;start=-1;}
    }
  }
  for(let gx=0;gx<gridWidth;gx++){
    let start=-1;
    for(let gy=0;gy<=gridHeight;gy++){
      const active=gy<gridHeight&&mask[gy*gridWidth+gx];
      if(active&&start<0)start=gy;
      if((!active||gy===gridHeight)&&start>=0){if(gy-start>=verticalMinimum)for(let y=start;y<gy;y++)mask[y*gridWidth+gx]=0;start=-1;}
    }
  }
}
function buildFallbackRegion(counts,gridWidth,gridHeight,width,height){
  let seed=-1,peak=0;
  for(let index=0;index<counts.length;index++)if(counts[index]>peak){peak=counts[index];seed=index;}
  if(seed<0||peak<=0)return null;
  const threshold=Math.max(1,Math.ceil(peak*.25)),visited=new Uint8Array(counts.length),queue=new Int32Array(counts.length);
  let head=0,tail=0;queue[tail++]=seed;visited[seed]=1;
  let minGX=seed%gridWidth,maxGX=minGX,minGY=Math.floor(seed/gridWidth),maxGY=minGY,pixels=0,cells=0;
  while(head<tail){
    const current=queue[head++],cx=current%gridWidth,cy=Math.floor(current/gridWidth),value=counts[current];
    if(value<threshold)continue;
    minGX=Math.min(minGX,cx);maxGX=Math.max(maxGX,cx);minGY=Math.min(minGY,cy);maxGY=Math.max(maxGY,cy);pixels+=value;cells++;
    for(let oy=-1;oy<=1;oy++)for(let ox=-1;ox<=1;ox++){
      if(!ox&&!oy)continue;
      const nx=cx+ox,ny=cy+oy;if(nx<0||nx>=gridWidth||ny<0||ny>=gridHeight)continue;
      const next=ny*gridWidth+nx;if(!visited[next]&&counts[next]>=threshold){visited[next]=1;queue[tail++]=next;}
    }
  }
  if(!cells)return null;
  const boxWidth=(maxGX-minGX+1)*BLOCK,boxHeight=(maxGY-minGY+1)*BLOCK;
  // Sparse render noise can connect distant cells. In that case emphasize the strongest
  // local cell instead of returning another misleading whole-page rectangle.
  if(boxWidth>width*.75&&boxHeight>height*.75&&pixels/Math.max(1,boxWidth*boxHeight)<.01){
    const sx=seed%gridWidth,sy=Math.floor(seed/gridWidth);minGX=Math.max(0,sx-2);maxGX=Math.min(gridWidth-1,sx+2);minGY=Math.max(0,sy-2);maxGY=Math.min(gridHeight-1,sy+2);pixels=peak;
  }
  const padding=PADDING+4,minX=Math.max(0,minGX*BLOCK-padding),minY=Math.max(0,minGY*BLOCK-padding);
  const maxX=Math.min(width-1,(maxGX+1)*BLOCK-1+padding),maxY=Math.min(height-1,(maxGY+1)*BLOCK-1+padding);
  return {regionId:'browser-r0001',kind:'modified',x:minX/width,y:minY/height,width:Math.max(1,maxX-minX+1)/width,height:Math.max(1,maxY-minY+1)/height,confidence:.5,pixelCount:pixels};
}
function strongestLocalBox(component,counts,gridWidth,gridHeight,width,height){
  const minGX=Math.max(0,Math.floor(component.minX/BLOCK)),maxGX=Math.min(gridWidth-1,Math.floor(component.maxX/BLOCK));
  const minGY=Math.max(0,Math.floor(component.minY/BLOCK)),maxGY=Math.min(gridHeight-1,Math.floor(component.maxY/BLOCK));
  let best=null;
  for(let gy=minGY;gy<=maxGY;gy++)for(let gx=minGX;gx<=maxGX;gx++){
    let score=0;
    for(let oy=-1;oy<=1;oy++)for(let ox=-2;ox<=2;ox++){
      const nx=gx+ox,ny=gy+oy;if(nx<minGX||nx>maxGX||ny<minGY||ny>maxGY)continue;
      score+=counts[ny*gridWidth+nx];
    }
    if(!best||score>best.score)best={gx,gy,score};
  }
  if(!best||best.score<MIN_PIXELS)return null;
  return {minX:Math.max(0,(best.gx-2)*BLOCK-PADDING),maxX:Math.min(width-1,(best.gx+3)*BLOCK-1+PADDING),
    minY:Math.max(0,(best.gy-1)*BLOCK-PADDING),maxY:Math.min(height-1,(best.gy+2)*BLOCK-1+PADDING),pixels:best.score};
}
function findTableGridBand(before,after,width,height){
  const candidates=[];
  for(let x=0;x<width;x++){
    let start=-1,last=-1,inkRows=0,bestAtX=null;
    const finish=()=>{
      if(start<0)return;
      const candidate={start,last,inkRows,span:last-start+1};
      if(candidate.span>height*.08&&candidate.inkRows/candidate.span>.35&&(!bestAtX||candidate.span>bestAtX.span||candidate.span===bestAtX.span&&candidate.inkRows>bestAtX.inkRows))bestAtX={...candidate,x};
      start=-1;last=-1;inkRows=0;
    };
    for(let y=0;y<height;y++){
      const index=(y*width+x)*4,dark=luminance(before,index)<252||luminance(after,index)<252;
      if(dark){if(start<0)start=y;last=y;inkRows++;}
      else if(start>=0&&y-last>3)finish();
    }
    finish();
    if(bestAtX)candidates.push(bestAtX);
  }
  // A rectangle or enlarged shape has only two long edges. Treat a region as a table
  // only when at least four separated vertical rules share substantially the same span.
  const clusters=[];
  for(const candidate of candidates){
    const previous=clusters[clusters.length-1];
    const overlap=previous?Math.max(0,Math.min(previous.last,candidate.last)-Math.max(previous.start,candidate.start)+1):0;
    if(previous&&candidate.x-previous.lastX<=1&&overlap/Math.max(1,Math.min(previous.span,candidate.span))>.72){
      previous.lastX=candidate.x;
      if(candidate.span>previous.span||candidate.inkRows>previous.inkRows)Object.assign(previous,{start:candidate.start,last:candidate.last,span:candidate.span,inkRows:candidate.inkRows});
    }else clusters.push({...candidate,firstX:candidate.x,lastX:candidate.x});
  }
  const lines=clusters.filter(cluster=>cluster.lastX-cluster.firstX+1<=8).map(cluster=>({...cluster,x:Math.round((cluster.firstX+cluster.lastX)/2)}));
  let verticalBest=null;
  for(const seed of lines){
    const group=lines.filter(line=>{
      const overlap=Math.max(0,Math.min(seed.last,line.last)-Math.max(seed.start,line.start)+1);
      const spanRatio=Math.min(seed.span,line.span)/Math.max(1,Math.max(seed.span,line.span));
      return overlap/Math.max(1,Math.min(seed.span,line.span))>.78&&spanRatio>.72;
    });
    if(group.length<4)continue;
    const starts=group.map(line=>line.start).sort((a,b)=>a-b),lasts=group.map(line=>line.last).sort((a,b)=>a-b);
    const start=starts[Math.floor(starts.length/2)],last=lasts[Math.floor(lasts.length/2)],span=last-start+1;
    const score=group.length*span;
    if(!verticalBest||score>verticalBest.score)verticalBest={start,last,span,score,lineCount:group.length,minX:Math.min(...group.map(line=>line.x)),maxX:Math.max(...group.map(line=>line.x)),detection:'vertical'};
  }
  const horizontalCandidates=[];
  for(let y=0;y<height;y++){
    let start=-1,last=-1,inkColumns=0,bestAtY=null;
    const finish=()=>{
      if(start<0)return;
      const candidate={start,last,inkColumns,span:last-start+1};
      if(candidate.span>width*.16&&candidate.inkColumns/candidate.span>.35&&(!bestAtY||candidate.span>bestAtY.span||candidate.span===bestAtY.span&&candidate.inkColumns>bestAtY.inkColumns))bestAtY={...candidate,y};
      start=-1;last=-1;inkColumns=0;
    };
    for(let x=0;x<width;x++){
      const index=(y*width+x)*4,dark=luminance(before,index)<252||luminance(after,index)<252;
      if(dark){if(start<0)start=x;last=x;inkColumns++;}
      else if(start>=0&&x-last>3)finish();
    }
    finish();if(bestAtY)horizontalCandidates.push(bestAtY);
  }
  const horizontalClusters=[];
  for(const candidate of horizontalCandidates){
    const previous=horizontalClusters[horizontalClusters.length-1];
    const overlap=previous?Math.max(0,Math.min(previous.last,candidate.last)-Math.max(previous.start,candidate.start)+1):0;
    if(previous&&candidate.y-previous.lastY<=1&&overlap/Math.max(1,Math.min(previous.span,candidate.span))>.72){
      previous.lastY=candidate.y;
      if(candidate.span>previous.span||candidate.inkColumns>previous.inkColumns)Object.assign(previous,{start:candidate.start,last:candidate.last,span:candidate.span,inkColumns:candidate.inkColumns});
    }else horizontalClusters.push({...candidate,firstY:candidate.y,lastY:candidate.y});
  }
  const horizontalLines=horizontalClusters.filter(cluster=>cluster.lastY-cluster.firstY+1<=8)
    .map(cluster=>({...cluster,y:Math.round((cluster.firstY+cluster.lastY)/2)}));
  let horizontalBest=null;
  for(const seed of horizontalLines){
    const group=horizontalLines.filter(line=>{
      const overlap=Math.max(0,Math.min(seed.last,line.last)-Math.max(seed.start,line.start)+1);
      const spanRatio=Math.min(seed.span,line.span)/Math.max(1,Math.max(seed.span,line.span));
      return overlap/Math.max(1,Math.min(seed.span,line.span))>.78&&spanRatio>.72;
    });
    if(group.length<4)continue;
    const xs=group.map(line=>line.start).sort((a,b)=>a-b),lastXs=group.map(line=>line.last).sort((a,b)=>a-b),ys=group.map(line=>line.y).sort((a,b)=>a-b);
    const gaps=[];for(let index=1;index<ys.length;index++){const gap=ys[index]-ys[index-1];if(gap>=6&&gap<=height*.12)gaps.push(gap);}
    gaps.sort((a,b)=>a-b);
    const start=ys[0],last=ys[ys.length-1],span=last-start+1,score=group.length*span;
    const candidate={start,last,span,score,lineCount:group.length,minX:xs[Math.floor(xs.length/2)],maxX:lastXs[Math.floor(lastXs.length/2)],
      rowHeight:gaps.length?gaps[Math.floor(gaps.length/2)]:null,horizontalLines:ys,detection:'horizontal'};
    if(!horizontalBest||score>horizontalBest.score)horizontalBest=candidate;
  }
  if(horizontalBest&&verticalBest){
    const overlap=Math.max(0,Math.min(horizontalBest.last,verticalBest.last)-Math.max(horizontalBest.start,verticalBest.start)+1);
    if(overlap/Math.max(1,Math.min(horizontalBest.span,verticalBest.span))>.65){
      horizontalBest.start=Math.min(horizontalBest.start,verticalBest.start);horizontalBest.last=Math.max(horizontalBest.last,verticalBest.last);
      horizontalBest.span=horizontalBest.last-horizontalBest.start+1;horizontalBest.detection='horizontal+vertical';
    }
  }
  // Repeated horizontal rules are more reliable for Excel tables: a resized rectangle
  // has two edges, while a table has a rule for every row even when vertical gridlines
  // are too faint or slightly broken by PDF rasterization.
  return horizontalBest||verticalBest;
}
function findTableGridBands(before,after,width,height){
  const segments=[];
  for(let y=0;y<height;y++){
    let start=-1,last=-1,inkColumns=0;
    const finish=()=>{
      if(start<0)return;
      const span=last-start+1,density=inkColumns/Math.max(1,span);
      if(span>width*.14&&density>.82)segments.push({y,start,last,span,inkColumns});
      start=-1;last=-1;inkColumns=0;
    };
    for(let x=0;x<width;x++){
      const index=(y*width+x)*4,dark=luminance(before,index)<254||luminance(after,index)<254;
      if(dark){if(start<0)start=x;last=x;inkColumns++;}
      else if(start>=0&&x-last>3)finish();
    }
    finish();
  }
  const clusters=[];
  for(const segment of segments){
    let target=null;
    for(let index=clusters.length-1;index>=Math.max(0,clusters.length-24);index--){
      const candidate=clusters[index];
      if(segment.y-candidate.lastY>1)break;
      const overlap=Math.max(0,Math.min(segment.last,candidate.last)-Math.max(segment.start,candidate.start)+1);
      const spanRatio=Math.min(segment.span,candidate.span)/Math.max(1,Math.max(segment.span,candidate.span));
      if(overlap/Math.max(1,Math.min(segment.span,candidate.span))>.78&&spanRatio>.72){target=candidate;break;}
    }
    if(target){
      target.firstY=Math.min(target.firstY,segment.y);target.lastY=Math.max(target.lastY,segment.y);
      target.start=Math.min(target.start,segment.start);target.last=Math.max(target.last,segment.last);
      target.span=Math.max(target.span,segment.span);target.inkColumns=Math.max(target.inkColumns,segment.inkColumns);
    }else clusters.push({...segment,firstY:segment.y,lastY:segment.y});
  }
  const lines=clusters.filter(cluster=>cluster.lastY-cluster.firstY+1<=8)
    .map(cluster=>({...cluster,y:Math.round((cluster.firstY+cluster.lastY)/2)}));
  const candidates=[];
  for(const seed of lines){
    const group=lines.filter(line=>{
      const overlap=Math.max(0,Math.min(seed.last,line.last)-Math.max(seed.start,line.start)+1);
      const spanRatio=Math.min(seed.span,line.span)/Math.max(1,Math.max(seed.span,line.span));
      return overlap/Math.max(1,Math.min(seed.span,line.span))>.76&&spanRatio>.7;
    });
    if(group.length<4)continue;
    const starts=group.map(line=>line.start).sort((a,b)=>a-b),lasts=group.map(line=>line.last).sort((a,b)=>a-b);
    const ys=group.map(line=>line.y).sort((a,b)=>a-b),gaps=[];
    for(let index=1;index<ys.length;index++){
      const gap=ys[index]-ys[index-1];
      if(gap>=6&&gap<=height*.12)gaps.push(gap);
    }
    gaps.sort((a,b)=>a-b);
    const rowHeight=gaps.length?gaps[Math.floor(gaps.length/2)]:null;
    let firstIndex=0;
    if(rowHeight)for(let index=0;index+1<ys.length;index++)if(ys[index+1]-ys[index]>rowHeight*3)firstIndex=index+1;
    const start=Math.max(0,Math.round(ys[firstIndex]-(rowHeight||0)*.55));
    const last=Math.min(height-1,Math.round(ys[ys.length-1]+(rowHeight||0)*1.8)),span=last-start+1;
    const minX=starts[Math.floor(starts.length/2)],maxX=lasts[Math.floor(lasts.length/2)];
    candidates.push({start,last,span,score:group.length*span*(maxX-minX+1)/width,lineCount:group.length,
      minX,maxX,
      rowHeight,horizontalLines:ys,detection:'horizontal-multi'});
  }
  candidates.sort((a,b)=>b.score-a.score||b.maxX-b.minX-(a.maxX-a.minX));
  const deduped=[];
  for(const candidate of candidates){
    const duplicate=deduped.some(existing=>{
      const overlapX=Math.max(0,Math.min(existing.maxX,candidate.maxX)-Math.max(existing.minX,candidate.minX)+1);
      const overlapY=Math.max(0,Math.min(existing.last,candidate.last)-Math.max(existing.start,candidate.start)+1);
      const spanRatioX=Math.min(existing.maxX-existing.minX+1,candidate.maxX-candidate.minX+1)/
        Math.max(1,Math.max(existing.maxX-existing.minX+1,candidate.maxX-candidate.minX+1));
      return overlapX/Math.max(1,Math.min(existing.maxX-existing.minX+1,candidate.maxX-candidate.minX+1))>.78&&spanRatioX>.68&&
        overlapY/Math.max(1,Math.min(existing.span,candidate.span))>.72;
    });
    if(!duplicate)deduped.push(candidate);
  }
  const primary=findTableGridBand(before,after,width,height);
  if(primary&&!deduped.length)deduped.push(primary);
  else if(primary&&!deduped.some(candidate=>{
    const overlapX=Math.max(0,Math.min(candidate.maxX,primary.maxX)-Math.max(candidate.minX,primary.minX)+1);
    const overlapY=Math.max(0,Math.min(candidate.last,primary.last)-Math.max(candidate.start,primary.start)+1);
    const spanRatioX=Math.min(candidate.maxX-candidate.minX+1,primary.maxX-primary.minX+1)/
      Math.max(1,Math.max(candidate.maxX-candidate.minX+1,primary.maxX-primary.minX+1));
    // A narrow text/rule strip can sit inside the real table and overlap it on
    // both axes. It is not a duplicate unless their horizontal spans also agree.
    return overlapX/Math.max(1,Math.min(candidate.maxX-candidate.minX+1,primary.maxX-primary.minX+1))>.65&&spanRatioX>.68&&
      overlapY/Math.max(1,Math.min(candidate.span,primary.span))>.65;
  }))deduped.push(primary);
  deduped.sort((a,b)=>b.score-a.score||b.maxX-b.minX-(a.maxX-a.minX));
  return deduped.slice(0,6);
}
function extractHeaderCellRules(data,width,height,tableBand){
  if(!tableBand)return null;
  const margin=Math.max(PADDING,24),minX=Math.max(0,Math.floor(tableBand.minX-margin)),maxX=Math.min(width-1,Math.ceil(tableBand.maxX+margin));
  const rowHeight=Math.max(12,tableBand.rowHeight||18),searchMinY=Math.max(0,Math.floor(tableBand.start-rowHeight*5.5));
  const searchMaxY=Math.min(height-1,Math.ceil(tableBand.start+rowHeight*2.5));
  let best=null,start=-1,last=-1;
  const finish=()=>{
    if(start<0)return;
    const span=last-start+1,center=(start+last)/2,distance=Math.abs(center-tableBand.start);
    if(span>=4&&(!best||distance<best.distance||distance===best.distance&&span>best.span))best={start,last,span,distance};
    start=-1;last=-1;
  };
  for(let y=searchMinY;y<=searchMaxY;y++){
    let dark=0;
    for(let x=minX;x<=maxX;x+=2)if(luminance(data,(y*width+x)*4)<205)dark++;
    const dense=dark/Math.max(1,Math.ceil((maxX-minX+1)/2))>.42;
    if(dense){if(start<0)start=y;last=y;}else if(start>=0)finish();
  }
  finish();if(!best)return null;
  const candidates=[];let headerMinX=maxX,headerMaxX=minX;
  for(let x=minX;x<=maxX;x++){
    let darkRows=0;
    for(let y=best.start;y<=best.last;y++)if(luminance(data,(y*width+x)*4)<220)darkRows++;
    const density=darkRows/Math.max(1,best.span);
    if(density>.35){headerMinX=Math.min(headerMinX,x);headerMaxX=Math.max(headerMaxX,x);}
    if(density<.08)candidates.push(x);
  }
  if(headerMaxX<headerMinX)return null;
  const rules=[headerMinX];let clusterStart=-1,clusterLast=-1;
  const flush=()=>{
    if(clusterStart<0)return;
    const clusterWidth=clusterLast-clusterStart+1;
    if(clusterWidth<=8&&clusterStart>headerMinX&&clusterLast<headerMaxX)rules.push(Math.round((clusterStart+clusterLast)/2));
    clusterStart=-1;clusterLast=-1;
  };
  for(const x of candidates){
    if(clusterStart<0){clusterStart=x;clusterLast=x;}
    else if(x-clusterLast<=1)clusterLast=x;
    else{flush();clusterStart=x;clusterLast=x;}
  }
  flush();rules.push(headerMaxX);
  const normalized=rules.sort((a,b)=>a-b).filter((value,index,array)=>!index||value-array[index-1]>=4);
  return {rules:normalized,start:best.start,last:best.last,valid:normalized.length>=4&&normalized.length<=20};
}
function extractTableVerticalRules(data,width,height,tableBand){
  if(!tableBand)return [];
  // Fit-to-page width changes can move the outer table edge farther than normal
  // highlight padding. Keep enough horizontal context to retain both edge rules.
  const ruleMargin=Math.max(PADDING,24);
  const minX=Math.max(0,Math.floor(tableBand.minX-ruleMargin)),maxX=Math.min(width-1,Math.ceil(tableBand.maxX+ruleMargin));
  const minY=Math.max(0,Math.floor(tableBand.start)),maxY=Math.min(height-1,Math.ceil(tableBand.last)),bandHeight=Math.max(1,maxY-minY+1),candidates=[];
  const minRuleSpan=Math.max(bandHeight*.14,Math.max(10,tableBand.rowHeight||18)*2.2);
  for(let x=minX;x<=maxX;x++){
    let start=-1,last=-1,inkRows=0,best=null;
    const finish=()=>{
      if(start<0)return;
      const span=last-start+1,density=inkRows/Math.max(1,span);
      if(span>minRuleSpan&&density>.28&&(!best||span>best.span||span===best.span&&inkRows>best.inkRows))best={x,start,last,span,inkRows};
      start=-1;last=-1;inkRows=0;
    };
    for(let y=minY;y<=maxY;y++){
      const dark=luminance(data,(y*width+x)*4)<252;
      if(dark){if(start<0)start=y;last=y;inkRows++;}
      else if(start>=0&&y-last>3)finish();
    }
    finish();if(best)candidates.push(best);
  }
  const clusters=[];
  for(const candidate of candidates){
    const previous=clusters[clusters.length-1];
    if(previous&&candidate.x-previous.lastX<=1){
      previous.lastX=candidate.x;previous.weight+=candidate.span;previous.weightedX+=candidate.x*candidate.span;
      previous.maxSpan=Math.max(previous.maxSpan,candidate.span);
    }else clusters.push({firstX:candidate.x,lastX:candidate.x,weight:candidate.span,weightedX:candidate.x*candidate.span,maxSpan:candidate.span});
  }
  const continuous=clusters.filter(cluster=>cluster.lastX-cluster.firstX+1<=8&&cluster.maxSpan>minRuleSpan)
    .map(cluster=>Math.round(cluster.weightedX/Math.max(1,cluster.weight))).filter((value,index,array)=>!index||value-array[index-1]>=4);
  const header=extractHeaderCellRules(data,width,height,tableBand);
  const headerNearTable=header&&header.start>=tableBand.start-Math.max(36,(tableBand.rowHeight||18)*3);
  if(headerNearTable){
    tableBand.headerStart=Number.isFinite(tableBand.headerStart)?Math.min(tableBand.headerStart,header.start):header.start;
  }
  if(headerNearTable&&header.valid){
    if(continuous.length<4||header.rules.length===continuous.length)return header.rules;
  }
  return continuous;
}
function extractTableHorizontalRules(data,width,height,tableBand){
  if(!tableBand)return [];
  // Excel fit-to-page can move the first/last rule by more than the ordinary
  // component padding when one row becomes taller. Keep two nominal rows of
  // vertical context so the unchanged table edge is not dropped.
  const margin=Math.max(PADDING,12,Math.min(36,Math.round(tableBand.rowHeight||18)*2)),minX=Math.max(0,Math.floor(tableBand.minX)),maxX=Math.min(width-1,Math.ceil(tableBand.maxX));
  const minY=Math.max(0,Math.floor(tableBand.start-margin)),maxY=Math.min(height-1,Math.ceil(tableBand.last+margin));
  const span=Math.max(1,maxX-minX+1),candidates=[];
  for(let y=minY;y<=maxY;y++){
    let dark=0;
    for(let x=minX;x<=maxX;x+=2)if(luminance(data,(y*width+x)*4)<248)dark++;
    if(dark/Math.max(1,Math.ceil(span/2))>.28)candidates.push({y,weight:dark});
  }
  const clusters=[];
  for(const candidate of candidates){
    const previous=clusters[clusters.length-1];
    if(previous&&candidate.y-previous.lastY<=1){previous.lastY=candidate.y;previous.weight+=candidate.weight;previous.weightedY+=candidate.y*candidate.weight;}
    else clusters.push({firstY:candidate.y,lastY:candidate.y,weight:candidate.weight,weightedY:candidate.y*candidate.weight});
  }
  return clusters.filter(cluster=>cluster.lastY-cluster.firstY+1<=7)
    .map(cluster=>Math.round(cluster.weightedY/Math.max(1,cluster.weight)))
    .filter((value,index,array)=>!index||value-array[index-1]>=4);
}
function tableRowRasterDistance(before,after,width,beforeTop,beforeBottom,afterTop,afterBottom,minX,maxX){
  let total=0,count=0;
  for(let band=0;band<8;band++){
    const beforeY=Math.max(0,Math.round(beforeTop+(beforeBottom-beforeTop)*(band+.5)/8));
    const afterY=Math.max(0,Math.round(afterTop+(afterBottom-afterTop)*(band+.5)/8));
    for(let x=minX+3;x<=maxX-3;x+=4){
      const beforeIndex=(beforeY*width+x)*4,afterIndex=(afterY*width+x)*4;
      if(luminance(before,beforeIndex)>247&&luminance(after,afterIndex)>247)continue;
      let difference=pixelDifference(before,after,beforeIndex,afterIndex);
      if(x>minX)difference=Math.min(difference,pixelDifference(before,after,beforeIndex,afterIndex-4));
      if(x<maxX)difference=Math.min(difference,pixelDifference(before,after,beforeIndex,afterIndex+4));
      total+=Math.min(80,difference);count++;
    }
  }
  return count?total/count:0;
}
function detectTableRowInsertionDeletion(before,after,width,height,tableBand){
  const beforeRules=extractTableHorizontalRules(before,width,height,tableBand),afterRules=extractTableHorizontalRules(after,width,height,tableBand);
  const beforeCount=beforeRules.length-1,afterCount=afterRules.length-1;
  if(Math.abs(beforeCount-afterCount)!==1||Math.min(beforeCount,afterCount)<4||Math.max(beforeCount,afterCount)>80)return null;
  const added=afterCount>beforeCount,longRules=added?afterRules:beforeRules,shortRules=added?beforeRules:afterRules;
  const longData=added?after:before,shortData=added?before:after;
  const minX=Math.max(0,Math.floor(tableBand.minX)),maxX=Math.min(width-1,Math.ceil(tableBand.maxX));
  const rowDistance=(shortIndex,longIndex)=>tableRowRasterDistance(shortData,longData,width,
    shortRules[shortIndex],shortRules[shortIndex+1],longRules[longIndex],longRules[longIndex+1],minX,maxX);
  let best=null,identity=0;
  for(let index=0;index<shortRules.length-1;index++)identity+=rowDistance(index,index);
  identity/=Math.max(1,shortRules.length-1);
  for(let extra=0;extra<longRules.length-1;extra++){
    let score=0;
    for(let index=0;index<shortRules.length-1;index++)score+=rowDistance(index,index<extra?index:index+1);
    score/=Math.max(1,shortRules.length-1);
    if(!best||score<best.score)best={extra,score};
  }
  // PDF text-row LCS is still required before this raster hint becomes a final
  // insertion/deletion result. Allow modest Excel glyph antialiasing here while
  // retaining a clear improvement over identity row matching.
  if(!best||best.score>24||identity>1&&best.score>identity*.85)return null;
  const shortHeights=[];for(let index=0;index<shortRules.length-1;index++)shortHeights.push(shortRules[index+1]-shortRules[index]);
  shortHeights.sort((a,b)=>a-b);const nominalHeight=shortHeights[Math.floor(shortHeights.length/2)]||1;
  const extraHeight=longRules[best.extra+1]-longRules[best.extra];
  // A newly drawn border can split one existing row into two short intervals.
  // Do not reinterpret that border-format edit as a newly inserted data row.
  if(extraHeight<nominalHeight*.55||extraHeight>nominalHeight*2.5)return null;
  const longTop=longRules[best.extra],longBottom=longRules[best.extra+1];
  const anchorIndex=Math.min(best.extra,shortRules.length-2),shortTop=shortRules[anchorIndex],shortBottom=shortRules[anchorIndex+1],padding=3;
  const longBox={minX:Math.max(0,minX-PADDING),maxX:Math.min(width-1,maxX+PADDING),minY:Math.max(0,longTop-padding),maxY:Math.min(height-1,longBottom+padding)};
  const shortBox={minX:longBox.minX,maxX:longBox.maxX,minY:Math.max(0,shortTop-padding),maxY:Math.min(height-1,shortBottom+padding)};
  return {kind:added?'added':'removed',score:best.score,identityScore:identity,tableBand,
    minX:longBox.minX,maxX:longBox.maxX,minY:longBox.minY,maxY:longBox.maxY,pixels:(longBox.maxX-longBox.minX+1)*(longBox.maxY-longBox.minY+1),
    beforeBox:added?shortBox:longBox,afterBox:added?longBox:shortBox};
}
function detectRowHeightChanges(before,after,width,height,tableBand){
  const beforeRules=extractTableHorizontalRules(before,width,height,tableBand),afterRules=extractTableHorizontalRules(after,width,height,tableBand);
  if(beforeRules.length<5||beforeRules.length!==afterRules.length||beforeRules.length>80)return [];
  const beforeSpan=beforeRules.at(-1)-beforeRules[0],afterSpan=afterRules.at(-1)-afterRules[0];
  if(beforeSpan<height*.08||afterSpan<height*.08)return [];
  const candidates=[];
  for(let index=0;index<beforeRules.length-1;index++){
    const beforeHeight=beforeRules[index+1]-beforeRules[index],afterHeight=afterRules[index+1]-afterRules[index];
    const normalizedDelta=Math.abs(afterHeight/afterSpan-beforeHeight/beforeSpan),pixelDelta=Math.abs(afterHeight-beforeHeight);
    if(normalizedDelta>=.004&&pixelDelta>=3)candidates.push({index,beforeHeight,afterHeight,normalizedDelta,pixelDelta});
  }
  candidates.sort((a,b)=>b.normalizedDelta-a.normalizedDelta||b.pixelDelta-a.pixelDelta);
  if(!candidates.length)return [];
  const best=candidates[0],threshold=Math.max(.004,best.normalizedDelta*.45),changes=[];
  for(const candidate of candidates.filter(value=>value.normalizedDelta>=threshold&&value.pixelDelta>=3)){
    const beforeTop=beforeRules[candidate.index],beforeBottom=beforeRules[candidate.index+1];
    const afterTop=afterRules[candidate.index],afterBottom=afterRules[candidate.index+1];
    const padding=3,minX=Math.max(0,tableBand.minX-PADDING),maxX=Math.min(width-1,tableBand.maxX+PADDING);
    changes.push({kind:'row-height',index:candidate.index,minX,maxX,
      minY:Math.max(0,Math.min(beforeTop,afterTop)-padding),maxY:Math.min(height-1,Math.max(beforeBottom,afterBottom)+padding),
      beforeBox:{minX,maxX,minY:Math.max(0,beforeTop-padding),maxY:Math.min(height-1,beforeBottom+padding)},
      afterBox:{minX,maxX,minY:Math.max(0,afterTop-padding),maxY:Math.min(height-1,afterBottom+padding)},
      normalizedDelta:candidate.normalizedDelta,pixelDelta:candidate.pixelDelta,tableBand});
    if(changes.length>=3)break;
  }
  return changes;
}
function detectAlignedRowHeightChange(rowAlignment,width,height,tableBand){
  if(!tableBand||!rowAlignment.adjusted||rowAlignment.splitRow<0||rowAlignment.deleted.length||rowAlignment.inserted.length)return null;
  const beforeBottom=Math.min(height-1,rowAlignment.splitRow*ROW_STEP);
  const mappedRow=rowAlignment.mapping[rowAlignment.splitRow];
  if(mappedRow<0)return null;
  const afterBottom=Math.min(height-1,mappedRow*ROW_STEP),pixelDelta=Math.abs(afterBottom-beforeBottom);
  const nominalHeight=Math.max(10,Math.min(48,Math.round(tableBand.rowHeight||18)));
  if(pixelDelta<3||pixelDelta>nominalHeight*1.75||beforeBottom<tableBand.start||beforeBottom>tableBand.last+nominalHeight)return null;
  const top=Math.max(0,Math.min(beforeBottom,afterBottom)-nominalHeight),padding=3;
  const minX=Math.max(0,tableBand.minX-PADDING),maxX=Math.min(width-1,tableBand.maxX+PADDING);
  const beforeBox={minX,maxX,minY:Math.max(0,top-padding),maxY:Math.min(height-1,beforeBottom+padding)};
  const afterBox={minX,maxX,minY:Math.max(0,top-padding),maxY:Math.min(height-1,afterBottom+padding)};
  return {kind:'row-height',minX,maxX,minY:Math.min(beforeBox.minY,afterBox.minY),maxY:Math.max(beforeBox.maxY,afterBox.maxY),
    beforeBox,afterBox,pixelDelta,normalizedDelta:pixelDelta/Math.max(1,tableBand.span),tableBand};
}
function detectColumnInsertionDeletion(before,after,width,height,tableBand){
  const beforeRules=extractTableVerticalRules(before,width,height,tableBand),afterRules=extractTableVerticalRules(after,width,height,tableBand);
  // A merged band spanning several neighboring tables can expose dozens of
  // unrelated vertical edges. It is not a credible single-table column model.
  if(Math.abs(beforeRules.length-afterRules.length)!==1||Math.min(beforeRules.length,afterRules.length)<4||Math.max(beforeRules.length,afterRules.length)>20)return null;
  // One taller Excel row can rescale the printed page and bring an adjacent
  // table edge into this band's horizontal margin. Prefer direct row-boundary
  // evidence over interpreting that edge as a newly inserted column.
  if(detectRowHeightChanges(before,after,width,height,tableBand).length)return null;
  const added=afterRules.length>beforeRules.length,longRules=added?afterRules:beforeRules,shortRules=added?beforeRules:afterRules;
  const normalize=rules=>rules.map(value=>(value-rules[0])/Math.max(1,rules.at(-1)-rules[0]));
  const shortNormalized=normalize(shortRules);let best=null;
  for(let extra=1;extra<longRules.length-1;extra++){
    const reduced=longRules.filter((_,index)=>index!==extra),reducedNormalized=normalize(reduced);
    const score=reducedNormalized.reduce((sum,value,index)=>sum+Math.abs(value-shortNormalized[index]),0)/shortNormalized.length;
    if(!best||score<best.score)best={extra,score};
  }
  if(!best||best.score>.045)return null;
  const left={min:longRules[best.extra-1],max:longRules[best.extra]},right={min:longRules[best.extra],max:longRules[best.extra+1]};
  const interval=(left.max-left.min)<=(right.max-right.min)?left:right;
  const minY=Math.max(0,(tableBand.headerStart??tableBand.start)-PADDING),maxY=Math.min(height-1,tableBand.last+PADDING),edgePadding=3;
  const fullBox={minX:Math.max(0,interval.min-edgePadding),maxX:Math.min(width-1,interval.max+edgePadding),minY,maxY};
  const anchor=longRules[best.extra],anchorBox={minX:Math.max(0,anchor-edgePadding),maxX:Math.min(width-1,anchor+edgePadding),minY,maxY};
  return {kind:added?'added':'removed',score:best.score,tableBand,
    minX:fullBox.minX,maxX:fullBox.maxX,minY,maxY,pixels:(fullBox.maxX-fullBox.minX+1)*(maxY-minY+1),
    beforeBox:added?anchorBox:fullBox,afterBox:added?fullBox:anchorBox,
    beforeRules,afterRules};
}
function detectColumnWidthBoundaryChanges(before,after,width,height,tableBand){
  const beforeRules=extractTableVerticalRules(before,width,height,tableBand),afterRules=extractTableVerticalRules(after,width,height,tableBand);
  if(beforeRules.length<4||beforeRules.length!==afterRules.length||beforeRules.length>20)return [];
  const beforeSpan=beforeRules[beforeRules.length-1]-beforeRules[0],afterSpan=afterRules[afterRules.length-1]-afterRules[0];
  if(beforeSpan<width*.12||afterSpan<width*.12)return [];
  const candidates=[];
  for(let index=0;index<beforeRules.length-1;index++){
    const beforeWidth=beforeRules[index+1]-beforeRules[index],afterWidth=afterRules[index+1]-afterRules[index];
    const normalizedBefore=beforeWidth/beforeSpan,normalizedAfter=afterWidth/afterSpan,normalizedDelta=Math.abs(normalizedAfter-normalizedBefore);
    const pixelDelta=normalizedDelta*(beforeSpan+afterSpan)/2;
    candidates.push({index,beforeWidth,afterWidth,normalizedDelta,pixelDelta});
  }
  candidates.sort((a,b)=>b.normalizedDelta-a.normalizedDelta);
  const best=candidates[0];
  if(!best||best.normalizedDelta<.006||best.pixelDelta<2)return [];
  const threshold=Math.max(.006,best.normalizedDelta*.35);
  const changes=[];
  for(const candidate of candidates.filter(candidate=>candidate.normalizedDelta>=threshold&&candidate.pixelDelta>=2)){
    const left=beforeRules[candidate.index],right=beforeRules[candidate.index+1],afterLeft=afterRules[candidate.index],afterRight=afterRules[candidate.index+1];
    const mappedAfterLeft=beforeRules[0]+(afterLeft-afterRules[0])*beforeSpan/afterSpan,mappedAfterRight=beforeRules[0]+(afterRight-afterRules[0])*beforeSpan/afterSpan;
    const useLeft=Math.abs(mappedAfterLeft-left)>Math.abs(mappedAfterRight-right),beforeBoundary=useLeft?left:right,afterBoundary=useLeft?afterLeft:afterRight;
    const mappedAfterBoundary=useLeft?mappedAfterLeft:mappedAfterRight;
    if(changes.some(change=>Math.abs(change.beforeBoundary-beforeBoundary)<=3&&Math.abs(change.mappedAfterBoundary-mappedAfterBoundary)<=3))continue;
    const edgePadding=3;
    const beforeMinX=Math.max(tableBand.minX,left-edgePadding),beforeMaxX=Math.min(tableBand.maxX,right+edgePadding);
    const afterMinX=Math.max(tableBand.minX,afterLeft-edgePadding),afterMaxX=Math.min(tableBand.maxX,afterRight+edgePadding);
    const minX=Math.min(beforeMinX,afterMinX),maxX=Math.max(beforeMaxX,afterMaxX);
    changes.push({kind:'column-width',index:candidate.index,minX,maxX,centerX:(minX+maxX)/2,half:Math.max(5,(maxX-minX+1)/2),
      pixelDelta:candidate.pixelDelta,normalizedDelta:candidate.normalizedDelta,beforeBoundary,afterBoundary,mappedAfterBoundary,
      beforeMinX,beforeMaxX,afterMinX,afterMaxX,
      beforeRules,afterRules,tableBand});
    if(changes.length>=3)break;
  }
  return changes;
}
function detectColumnWidthBoundaryChange(before,after,width,height,tableBand){
  return detectColumnWidthBoundaryChanges(before,after,width,height,tableBand)[0]||null;
}
function strongestStructuralRowBox(centerY,tableBand,counts,gridWidth,gridHeight,width,height){
  const rowHeight=Math.max(10,Math.min(40,Math.round(tableBand&&tableBand.rowHeight||18)));
  const minX=tableBand?Math.max(0,tableBand.minX-PADDING):BLOCK,maxX=tableBand?Math.min(width-1,tableBand.maxX+PADDING):width-BLOCK-1;
  const minGX=Math.max(0,Math.floor(minX/BLOCK)),maxGX=Math.min(gridWidth-1,Math.floor(maxX/BLOCK));
  let best=null;
  for(let start=Math.round(centerY-rowHeight*1.5);start<=Math.round(centerY+rowHeight*.5);start+=BLOCK){
    const minY=Math.max(0,start),maxY=Math.min(height-1,minY+rowHeight-1);let score=0;
    for(let gy=Math.max(0,Math.floor(minY/BLOCK));gy<=Math.min(gridHeight-1,Math.floor(maxY/BLOCK));gy++)for(let gx=minGX;gx<=maxGX;gx++)score+=counts[gy*gridWidth+gx];
    const distance=Math.abs((minY+maxY)/2-centerY),fitness=score-distance*.25;
    if(!best||fitness>best.fitness)best={minX,maxX,minY:Math.max(0,minY-2),maxY:Math.min(height-1,maxY+2),pixels:Math.max(MIN_PIXELS,score),fitness};
  }
  return best||{minX,maxX,minY:Math.max(0,centerY-Math.floor(rowHeight/2)),maxY:Math.min(height-1,centerY+Math.ceil(rowHeight/2)),pixels:MIN_PIXELS};
}
function structuralColumnInkBounds(before,after,width,height,columnAlignment,centerX,half){
  const verticalBand=findTableGridBand(before,after,width,height);
  if(verticalBand)return {minY:Math.max(0,verticalBand.start-PADDING),maxY:Math.min(height-1,verticalBand.last+PADDING)};
  const rows=new Uint8Array(height),probeHalf=Math.max(24,half+12),left=Math.max(0,centerX-probeHalf),right=Math.min(width-1,centerX+probeHalf);
  for(let y=0;y<height;y++){
    let ink=0;
    for(let x=left;x<=right;x+=2){
      const mapped=mappedColumnX(columnAlignment,x,width,y),bi=(y*width+x)*4,ai=(y*width+mapped)*4;
      if(luminance(before,bi)<245||luminance(after,ai)<245){ink++;if(ink>=2){rows[y]=1;break;}}
    }
  }
  let best=null,start=-1,last=-1,inkRows=0;
  const finish=()=>{
    if(start<0)return;
    const candidate={start,last,inkRows,span:last-start+1};
    if(!best||candidate.span>best.span||(candidate.span===best.span&&candidate.inkRows>best.inkRows))best=candidate;
    start=-1;last=-1;inkRows=0;
  };
  for(let y=0;y<height;y++){
    if(rows[y]){if(start<0)start=y;last=y;inkRows++;}
    else if(start>=0&&y-last>8)finish();
  }
  finish();
  if(!best)return {minY:BLOCK,maxY:height-BLOCK-1};
  return {minY:Math.max(0,best.start-PADDING),maxY:Math.min(height-1,best.last+PADDING)};
}
function analyzeBrowserDiff(before,after,width,height){
  const rowAlignment=alignRows(before,after,width,height);
  const tableBands=findTableGridBands(before,after,width,height);
  let tableBand=tableBands[0]||findTableGridBand(before,after,width,height);
  const strongestBandScore=tableBands[0]?.score||0;
  const candidateTableBands=tableBands.filter(band=>band.span>=height*.08&&(!strongestBandScore||band.score>=strongestBandScore*.2));
  const tableRowStructureChanges=candidateTableBands.map(band=>detectTableRowInsertionDeletion(before,after,width,height,band))
    .filter(Boolean).sort((a,b)=>a.score-b.score);
  const tableRowStructureChange=tableRowStructureChanges[0]||null;
  let rowHeightChanges=candidateTableBands.flatMap(band=>detectRowHeightChanges(before,after,width,height,band))
    .sort((a,b)=>b.normalizedDelta-a.normalizedDelta||b.pixelDelta-a.pixelDelta).slice(0,1);
  let rowHeightChange=rowHeightChanges[0]||null;
  const columnStructureChanges=candidateTableBands.map(band=>detectColumnInsertionDeletion(before,after,width,height,band))
    .filter(Boolean).sort((a,b)=>a.score-b.score);
  const columnStructureChange=columnStructureChanges[0]||null;
  const columnBoundaryChanges=columnStructureChange?[]:candidateTableBands.flatMap(band=>detectColumnWidthBoundaryChanges(before,after,width,height,band))
    .sort((a,b)=>b.normalizedDelta-a.normalizedDelta||b.pixelDelta-a.pixelDelta).slice(0,1);
  const columnBoundaryChange=columnBoundaryChanges[0]||null;
  if(columnStructureChange)tableBand=columnStructureChange.tableBand;
  else if(rowHeightChange)tableBand=rowHeightChange.tableBand;
  else if(columnBoundaryChange)tableBand=columnBoundaryChange.tableBand;
  const reliableColumnRules=tableBand&&extractTableVerticalRules(before,width,height,tableBand).length>=4&&
    extractTableVerticalRules(after,width,height,tableBand).length>=4;
  if(rowAlignment.adjusted&&tableBand){
    rowAlignment.activeMinY=Math.max(0,tableBand.start-PADDING);rowAlignment.activeMaxY=Math.min(height-1,tableBand.last+PADDING);
    rowAlignment.activeMinX=Math.max(0,tableBand.minX-PADDING);rowAlignment.activeMaxX=Math.min(width-1,tableBand.maxX+PADDING);
  }
  const columnAlignment=choosePixelColumnMapping(before,after,width,height,rowAlignment,tableBand);
  if(!rowHeightChange&&!columnAlignment.adjusted&&!columnStructureChange&&!columnBoundaryChange){
    const aligned=candidateTableBands.map(band=>detectAlignedRowHeightChange(rowAlignment,width,height,band)).filter(Boolean)
      .sort((a,b)=>b.normalizedDelta-a.normalizedDelta||b.pixelDelta-a.pixelDelta)[0];
    if(aligned){rowHeightChanges=[aligned];rowHeightChange=aligned;tableBand=aligned.tableBand;}
  }
  if(columnAlignment.adjusted){
    let activeBounds=tableBand?{minY:Math.max(0,tableBand.start-PADDING),maxY:Math.min(height-1,tableBand.last+PADDING)}:null;
    if(columnAlignment.structuralSplit>=0&&Math.abs(columnAlignment.structuralJump)>=3){
      const centerX=Math.min(width-1,columnAlignment.structuralSplit*COLUMN_STEP),half=Math.max(6,Math.ceil(Math.abs(columnAlignment.structuralJump)/2)+3);
      activeBounds=structuralColumnInkBounds(before,after,width,height,columnAlignment,centerX,half);
    }
    if(activeBounds){
      columnAlignment.activeMinY=activeBounds.minY;columnAlignment.activeMaxY=activeBounds.maxY;
      columnAlignment.activeMinX=Math.max(0,tableBand.minX-PADDING);columnAlignment.activeMaxX=Math.min(width-1,tableBand.maxX+PADDING);
    }
  }
  refinePixelRowMapping(before,after,width,height,rowAlignment,columnAlignment);
  refinePixelColumnMapping(before,after,width,height,rowAlignment,columnAlignment);
  const gridWidth=Math.ceil(width/BLOCK),gridHeight=Math.ceil(height/BLOCK),cellCount=gridWidth*gridHeight;
  const mask=new Uint8Array(cellCount),counts=new Uint32Array(cellCount),looseCounts=new Uint32Array(cellCount);
  let totalChanged=0;
  for(let gy=0;gy<gridHeight;gy++)for(let gx=0;gx<gridWidth;gx++){
    let changed=0,looseChanged=0;
    const x0=gx*BLOCK,y0=gy*BLOCK,x1=Math.min(width,x0+BLOCK),y1=Math.min(height,y0+BLOCK);
    for(let y=y0;y<y1;y++)for(let x=x0;x<x1;x++){
      const ax=mappedColumnX(columnAlignment,x,width,y),ay=mappedRowY(rowAlignment,y,height,x),bi=(y*width+x)*4;
      let difference=0;
      if(ax<0||ax>=width||ay<0||ay>=height){
        difference=255-Math.min(before[bi],before[bi+1],before[bi+2]);
      }else difference=tolerantPixelDifference(before,after,width,height,x,y,ax,ay);
      if(difference>8)looseChanged++;
      if(difference>PIXEL_THRESHOLD)changed++;
    }
    const index=gy*gridWidth+gx;
    counts[index]=changed;looseCounts[index]=looseChanged;totalChanged+=changed;
    if(changed>=Math.max(2,Math.floor((x1-x0)*(y1-y0)*.2)))mask[index]=1;
  }
  // Rows that have no counterpart are the human-visible insertion/deletion bands.
  const gapRows=[],structuralRowBoxes=[],structuralColumnBoxes=[];
  for(const change of rowHeightChanges){
    change.pixels=Math.max(MIN_PIXELS,(change.maxX-change.minX+1)*(change.maxY-change.minY+1));
    structuralRowBoxes.push(change);
    for(let gy=Math.max(0,Math.floor(change.minY/BLOCK));gy<=Math.min(gridHeight-1,Math.floor(change.maxY/BLOCK));gy++)for(let gx=Math.max(0,Math.floor(change.minX/BLOCK));gx<=Math.min(gridWidth-1,Math.floor(change.maxX/BLOCK));gx++){
      const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
    }
  }
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
  for(const gap of rowHeightChange?[]:gapRows){
    const centerY=Math.min(height-1,gap.beforeRow*ROW_STEP),half=6;
    let minGX=gridWidth,maxGX=-1;
    for(let y=Math.max(0,centerY-half);y<Math.min(height,centerY+half);y+=4)for(let x=8;x<width-8;x+=8){
      const bi=(y*width+x)*4,ay=gap.source==='after'?Math.min(height-1,(gap.afterRow||0)*ROW_STEP):mappedRowY(rowAlignment,y,height,x),ai=(ay*width+mappedColumnX(columnAlignment,x,width,y))*4;
      if(luminance(before,bi)<245||luminance(after,ai)<245){minGX=Math.min(minGX,Math.floor(x/BLOCK));maxGX=Math.max(maxGX,Math.floor(x/BLOCK));}
    }
    if(maxGX<minGX){minGX=2;maxGX=gridWidth-3;}
    const rowBox=strongestStructuralRowBox(centerY,tableBand,counts,gridWidth,gridHeight,width,height);structuralRowBoxes.push(rowBox);
    for(let gy=Math.max(0,Math.floor(rowBox.minY/BLOCK));gy<=Math.min(gridHeight-1,Math.floor(rowBox.maxY/BLOCK));gy++)for(let gx=Math.max(minGX,Math.floor(rowBox.minX/BLOCK));gx<=Math.min(maxGX,Math.floor(rowBox.maxX/BLOCK));gx++){
      const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
    }
  }
  if(!rowHeightChange&&rowAlignment.splitRow>=0&&Math.abs(rowAlignment.jumpPixels)>=3){
    const centerY=Math.min(height-1,rowAlignment.splitRow*ROW_STEP),rowBox=strongestStructuralRowBox(centerY,tableBand,counts,gridWidth,gridHeight,width,height);
    structuralRowBoxes.push(rowBox);
    for(let gy=Math.max(0,Math.floor(rowBox.minY/BLOCK));gy<=Math.min(gridHeight-1,Math.floor(rowBox.maxY/BLOCK));gy++)for(let gx=Math.max(0,Math.floor(rowBox.minX/BLOCK));gx<=Math.min(gridWidth-1,Math.floor(rowBox.maxX/BLOCK));gx++){
      const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
    }
  }
  if(columnStructureChange){
    structuralColumnBoxes.push(columnStructureChange);
    for(let gx=Math.max(0,Math.floor(columnStructureChange.minX/BLOCK));gx<=Math.min(gridWidth-1,Math.floor(columnStructureChange.maxX/BLOCK));gx++)for(let gy=Math.max(0,Math.floor(columnStructureChange.minY/BLOCK));gy<=Math.min(gridHeight-1,Math.floor(columnStructureChange.maxY/BLOCK));gy++){
      const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
    }
  }else if((columnAlignment.adjusted||columnBoundaryChange)&&tableBand){
    const changes=columnBoundaryChanges.length?columnBoundaryChanges:[null];
    for(const boundaryChange of changes){
      let centerX,half;
      if(boundaryChange){
        centerX=boundaryChange.centerX;half=boundaryChange.half;
      }else if(columnAlignment.structuralSplit>=0&&Math.abs(columnAlignment.structuralJump)>=3){
        centerX=Math.min(width-1,columnAlignment.structuralSplit*COLUMN_STEP);half=Math.max(6,Math.ceil(Math.abs(columnAlignment.structuralJump)/2)+3);
      }else if(!reliableColumnRules){
        // Without a detected boundary or a local discontinuity, highlighting the
        // whole table is misleading. Leave localization to the residual regions.
        continue;
      }else{
        let peak=-1,peakScore=-1;
        for(let column=0;column<columnAlignment.a.columnCount;column++){
          const target=Math.round(columnAlignment.mapping[column]);if(target<0||target>=columnAlignment.b.columnCount)continue;
          const score=columnDistance(columnAlignment.a,columnAlignment.b,column,target);
          if(score>peakScore){peakScore=score;peak=column;}
        }
        centerX=Math.max(0,Math.min(width-1,(peak>=0?peak:Math.floor(columnAlignment.a.columnCount/2))*COLUMN_STEP));
        half=Math.max(8,Math.min(Math.ceil(width*.1),Math.ceil(Math.abs(columnAlignment.offset))+COLUMN_STEP*2));
      }
      const activeBand=boundaryChange?.tableBand||tableBand;
      const inkBounds=activeBand?{minY:Math.max(0,(activeBand.headerStart??activeBand.start)-PADDING),maxY:Math.min(height-1,activeBand.last+PADDING)}:
        structuralColumnInkBounds(before,after,width,height,columnAlignment,centerX,half);
      const minGY=Math.max(0,Math.floor(inkBounds.minY/BLOCK)),maxGY=Math.min(gridHeight-1,Math.floor(inkBounds.maxY/BLOCK));
      const bandMinX=Math.max(0,centerX-half-PADDING),bandMaxX=Math.min(width-1,centerX+half+PADDING);
      const structuralColumnBox={minX:bandMinX,maxX:bandMaxX,minY:inkBounds.minY,maxY:inkBounds.maxY,pixels:(bandMaxX-bandMinX+1)*(inkBounds.maxY-inkBounds.minY+1)};
      if(boundaryChange){
        const edgePadding=2;
        structuralColumnBox.beforeBox={minX:Math.max(0,boundaryChange.beforeMinX-edgePadding),maxX:Math.min(width-1,boundaryChange.beforeMaxX+edgePadding),minY:inkBounds.minY,maxY:inkBounds.maxY};
        structuralColumnBox.afterBox={minX:Math.max(0,boundaryChange.afterMinX-edgePadding),maxX:Math.min(width-1,boundaryChange.afterMaxX+edgePadding),minY:inkBounds.minY,maxY:inkBounds.maxY};
      }
      structuralColumnBoxes.push(structuralColumnBox);
      for(let gx=Math.max(0,Math.floor((centerX-half)/BLOCK));gx<=Math.min(gridWidth-1,Math.floor((centerX+half)/BLOCK));gx++)for(let gy=minGY;gy<=maxGY;gy++){
        const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
      }
    }
  }
  // Grow horizontally to combine one logical row, but not vertically across many spreadsheet rows.
  const raw=collectMaskComponents(mask,counts,gridWidth,gridHeight,width,height,2);
  let components=mergeNearbyComponents(raw,width,height),rowStructureBoxes=[];
  if(structuralRowBoxes.length){
    const rowBoxes=[];
    for(const box of structuralRowBoxes)if(!rowBoxes.some(item=>Math.abs(item.minY-box.minY)<BLOCK*2))rowBoxes.push(box);
    rowStructureBoxes=rowBoxes;
    const localized=[];
    components=components.filter(component=>{
      if(rowBoxes.some(box=>component.maxY>=box.minY&&component.minY<=box.maxY))return false;
      const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
      if(componentWidth>width*.5&&componentHeight<height*.14){
        const local=strongestLocalBox(component,counts,gridWidth,gridHeight,width,height);
        if(local&&!rowBoxes.some(box=>local.maxY>=box.minY&&local.minY<=box.maxY))localized.push(local);
        return false;
      }
      return true;
    });
    components.push(...localized.slice(0,3),...rowBoxes);
    if(rowAlignment.adjusted||rowHeightChange){
      // Page-height growth and shifted totals can leave strong residuals far from
      // the inserted row. Keep the structural band and localized edits elsewhere;
      // only broad residuals are alignment noise. Dropping every off-band component
      // erases cell edits accumulated across older history generations.
      if(tableRowStructureChange)components=components.filter(component=>rowBoxes.some(box=>component.maxY>=box.minY&&component.minY<=box.maxY));
      else components=components.filter(component=>{
          if(rowBoxes.some(box=>component.maxY>=box.minY&&component.minY<=box.maxY))return true;
          const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
          return componentWidth<=width*.42&&componentHeight<=height*.14&&componentWidth*componentHeight<=width*height*.04;
        });
    }
  }
  if(structuralColumnBoxes.length){
    // Residual antialiasing after an Excel fit-to-page scale can connect every table
    // gridline into one huge rectangle. Keep the discontinuity as the primary signal,
    // but do not discard independent cell edits elsewhere on the same page. Historical
    // comparisons often combine a column resize in one generation with text edits from
    // earlier generations, and replacing the component list hid those cumulative edits.
    if(columnStructureChange||columnBoundaryChange||!reliableColumnRules){
      const bandTop=Math.min(...structuralColumnBoxes.map(box=>box.minY)),bandBottom=Math.max(...structuralColumnBoxes.map(box=>box.maxY));
      const localizedResiduals=components.filter(component=>{
        const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
        const componentArea=componentWidth*componentHeight;
        const overlapsStructuralBand=component.maxY>=bandTop&&component.minY<=bandBottom;
        const overlapsStructuralColumn=structuralColumnBoxes.some(box=>{
          const overlapX=Math.min(component.maxX,box.maxX)-Math.max(component.minX,box.minX)+1;
          const overlapY=Math.min(component.maxY,box.maxY)-Math.max(component.minY,box.minY)+1;
          return overlapX>0&&overlapY>0;
        });
        if(overlapsStructuralColumn)return false;
        // Outside the affected table band the alignment is identity, so even a wider
        // component is a real page edit. Inside the band, retain only localized residuals
        // and suppress the broad gridline/antialiasing artifacts caused by the resize.
        if(!overlapsStructuralBand)return true;
        return componentWidth<=width*.42&&componentHeight<=height*.35&&componentArea<=width*height*.08;
      });
      components=[...localizedResiduals,...structuralColumnBoxes];
    }else{
      const structuralPixels=Math.max(...structuralColumnBoxes.map(box=>box.pixels||MIN_PIXELS));
      components=components.filter(component=>{
        const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
        const componentArea=componentWidth*componentHeight;
        return component.pixels>=structuralPixels*.01&&
          !(componentWidth>width*.42||componentHeight>height*.35||componentArea>width*height*.08);
      });
      components.push(...structuralColumnBoxes);
    }
    if(components.length>MAX_LAYOUT_REGIONS)components.sort((a,b)=>b.pixels-a.pixels).splice(MAX_LAYOUT_REGIONS);
  }
  if(structuralColumnBoxes.length||structuralRowBoxes.length&&!tableRowStructureChange){
    // Structural rules can connect otherwise independent text edits into one page-wide
    // component. Remove the known structural bands and long rule runs, then recover the
    // remaining compact components as supplementary cumulative changes.
    const recoveryMask=new Uint8Array(mask),exclusionBoxes=[...structuralRowBoxes,...structuralColumnBoxes];
    for(const box of exclusionBoxes){
      const minGX=Math.max(0,Math.floor((box.minX-PADDING)/BLOCK)),maxGX=Math.min(gridWidth-1,Math.floor((box.maxX+PADDING)/BLOCK));
      const minGY=Math.max(0,Math.floor((box.minY-PADDING)/BLOCK)),maxGY=Math.min(gridHeight-1,Math.floor((box.maxY+PADDING)/BLOCK));
      for(let gy=minGY;gy<=maxGY;gy++)for(let gx=minGX;gx<=maxGX;gx++)recoveryMask[gy*gridWidth+gx]=0;
    }
    clearLongMaskRuns(recoveryMask,gridWidth,gridHeight);
    const recovered=mergeNearbyComponents(collectMaskComponents(recoveryMask,counts,gridWidth,gridHeight,width,height,1),width,height)
      .filter(component=>{
        const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
        const aspect=Math.max(componentWidth/componentHeight,componentHeight/componentWidth);
        return componentWidth<=width*.25&&componentHeight<=height*.12&&componentWidth*componentHeight<=width*height*.02&&aspect<=6;
      })
      .sort((a,b)=>{
        const score=component=>component.pixels*Math.sqrt(component.pixels/Math.max(1,(component.maxX-component.minX+1)*(component.maxY-component.minY+1)));
        return score(b)-score(a);
      }).slice(0,12);
    for(const component of recovered){
      const duplicate=components.some(existing=>{
        const overlapX=Math.min(component.maxX,existing.maxX)-Math.max(component.minX,existing.minX)+1;
        const overlapY=Math.min(component.maxY,existing.maxY)-Math.max(component.minY,existing.minY)+1;
        if(overlapX<=0||overlapY<=0)return false;
        const componentArea=(component.maxX-component.minX+1)*(component.maxY-component.minY+1);
        return overlapX*overlapY/componentArea>=.35;
      });
      if(!duplicate)components.push(component);
    }
  }
  const dominantHorizontal=components
    .filter(component=>(component.maxX-component.minX+1)>width*.2&&(component.maxY-component.minY+1)<height*.05)
    .sort((a,b)=>b.pixels-a.pixels)[0]||null;
  if(dominantHorizontal&&(rowAlignment.adjusted||columnAlignment.adjusted||columnBoundaryChange||gapRows.length>0)){
    components=components.filter(component=>{
      if(component===dominantHorizontal)return true;
      const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
      const topTextJitter=component.maxY<height*.25&&componentWidth<width*.25&&componentHeight<height*.12&&component.pixels<dominantHorizontal.pixels*.8;
      return !topTextJitter;
    });
  }
  if(components.length>MAX_REGIONS)components.sort((a,b)=>b.pixels-a.pixels).splice(MAX_REGIONS);
  components.sort((a,b)=>a.minY-b.minY||a.minX-b.minX);
  const normalizeBox=box=>({x:box.minX/width,y:box.minY/height,width:Math.max(1,box.maxX-box.minX+1)/width,height:Math.max(1,box.maxY-box.minY+1)/height});
  let regions=components.map((component,index)=>{
    const region={regionId:`browser-r${String(index+1).padStart(4,'0')}`,kind:component.kind==='added'||component.kind==='removed'?component.kind:'modified',...normalizeBox(component),
      confidence:Math.max(.55,Math.min(1,component.pixels/Math.max(MIN_PIXELS,(component.maxX-component.minX+1)*(component.maxY-component.minY+1)))),pixelCount:component.pixels};
    if(component.beforeBox&&component.afterBox){region.before=normalizeBox(component.beforeBox);region.after=normalizeBox(component.afterBox);}
    return region;
  });
  let fallbackUsed=false;
  if(!regions.length){
    const fallback=buildFallbackRegion(totalChanged?counts:looseCounts,gridWidth,gridHeight,width,height);
    if(fallback){regions=[fallback];fallbackUsed=true;}
  }
  const rowHeightAdjusted=!!rowHeightChange;
  const rowStructureAdjusted=rowStructureBoxes.length>0&&(rowAlignment.adjusted||gapRows.length>0||rowHeightAdjusted);
  const rowStructureRegions=rowStructureBoxes.map(box=>{
    if(!box.beforeBox||!box.afterBox)return normalizeBox(box);
    const beforeHeight=box.beforeBox.maxY-box.beforeBox.minY,afterHeight=box.afterBox.maxY-box.afterBox.minY;
    return normalizeBox(beforeHeight<=afterHeight?box.beforeBox:box.afterBox);
  });
  const tableRowStructureBand=tableRowStructureChange?normalizeBox({
    minX:Math.max(0,tableRowStructureChange.tableBand.minX-PADDING),maxX:Math.min(width-1,tableRowStructureChange.tableBand.maxX+PADDING),
    minY:Math.max(0,tableRowStructureChange.tableBand.start-PADDING),maxY:Math.min(height-1,tableRowStructureChange.tableBand.last+PADDING)
  }):null;
  const alignmentAdjusted=rowAlignment.adjusted||columnAlignment.adjusted||!!columnBoundaryChange||!!columnStructureChange||rowHeightAdjusted||gapRows.length>0;
  const rowMode=rowHeightAdjusted?'row-boundary-height':rowAlignment.alignmentMode;
  const columnMode=columnStructureChange?`column-${columnStructureChange.kind}`:columnBoundaryChange?'column-boundary-width':columnAlignment.adjusted?(columnAlignment.split>=0?'column-scale-and-shift':'column-scale'):'column-identity';
  return {regions,changedRatio:totalChanged/Math.max(1,width*height),offsetX:columnAlignment.offset,offsetY:0,scaleX:columnAlignment.scale,scaleY:rowAlignment.scaleY,maxLocalShiftJump:Math.max(rowAlignment.maxLocalShiftJump,Math.abs(columnAlignment.jump),columnBoundaryChange?.pixelDelta||0),alignmentMode:`${rowMode}/${columnMode}`,alignmentAdjusted,rowStructureAdjusted,rowStructureRegions,rowHeightAdjusted,tableRowStructureDetected:!!tableRowStructureChange,tableRowStructureBand,columnStructureKind:columnStructureChange?.kind||'',fallbackUsed};
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
