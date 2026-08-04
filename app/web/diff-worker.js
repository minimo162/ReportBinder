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
function extractTableVerticalRules(data,width,height,tableBand){
  if(!tableBand)return [];
  const minX=Math.max(0,Math.floor(tableBand.minX-PADDING)),maxX=Math.min(width-1,Math.ceil(tableBand.maxX+PADDING));
  const minY=Math.max(0,Math.floor(tableBand.start)),maxY=Math.min(height-1,Math.ceil(tableBand.last)),bandHeight=Math.max(1,maxY-minY+1),candidates=[];
  for(let x=minX;x<=maxX;x++){
    let start=-1,last=-1,inkRows=0,best=null;
    const finish=()=>{
      if(start<0)return;
      const span=last-start+1,density=inkRows/Math.max(1,span);
      if(span>bandHeight*.32&&density>.28&&(!best||span>best.span||span===best.span&&inkRows>best.inkRows))best={x,start,last,span,inkRows};
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
  return clusters.filter(cluster=>cluster.lastX-cluster.firstX+1<=8&&cluster.maxSpan>bandHeight*.32)
    .map(cluster=>Math.round(cluster.weightedX/Math.max(1,cluster.weight))).filter((value,index,array)=>!index||value-array[index-1]>=4);
}
function detectColumnWidthBoundaryChange(before,after,width,height,tableBand){
  const beforeRules=extractTableVerticalRules(before,width,height,tableBand),afterRules=extractTableVerticalRules(after,width,height,tableBand);
  if(beforeRules.length<4||beforeRules.length!==afterRules.length||beforeRules.length>20)return null;
  const beforeSpan=beforeRules[beforeRules.length-1]-beforeRules[0],afterSpan=afterRules[afterRules.length-1]-afterRules[0];
  if(beforeSpan<width*.12||afterSpan<width*.12)return null;
  const candidates=[];
  for(let index=0;index<beforeRules.length-1;index++){
    const beforeWidth=beforeRules[index+1]-beforeRules[index],afterWidth=afterRules[index+1]-afterRules[index];
    const normalizedBefore=beforeWidth/beforeSpan,normalizedAfter=afterWidth/afterSpan,normalizedDelta=Math.abs(normalizedAfter-normalizedBefore);
    const pixelDelta=normalizedDelta*(beforeSpan+afterSpan)/2;
    candidates.push({index,beforeWidth,afterWidth,normalizedDelta,pixelDelta});
  }
  candidates.sort((a,b)=>b.normalizedDelta-a.normalizedDelta);
  const best=candidates[0],second=candidates[1]||{normalizedDelta:0};
  if(!best||best.normalizedDelta<.012||best.pixelDelta<3||best.normalizedDelta<second.normalizedDelta*1.35)return null;
  const left=beforeRules[best.index],right=beforeRules[best.index+1],afterLeft=afterRules[best.index],afterRight=afterRules[best.index+1];
  const mappedAfterLeft=beforeRules[0]+(afterLeft-afterRules[0])*beforeSpan/afterSpan,mappedAfterRight=beforeRules[0]+(afterRight-afterRules[0])*beforeSpan/afterSpan;
  const minX=Math.max(tableBand.minX,Math.floor(Math.min(left,right,mappedAfterLeft,mappedAfterRight))),maxX=Math.min(tableBand.maxX,Math.ceil(Math.max(left,right,mappedAfterLeft,mappedAfterRight)));
  return {kind:'column-width',index:best.index,minX,maxX,centerX:(minX+maxX)/2,half:Math.max(7,(maxX-minX+1)/2),pixelDelta:best.pixelDelta,beforeRules,afterRules};
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
  const tableBand=findTableGridBand(before,after,width,height);
  const columnBoundaryChange=detectColumnWidthBoundaryChange(before,after,width,height,tableBand);
  if(rowAlignment.adjusted&&tableBand){
    rowAlignment.activeMinY=Math.max(0,tableBand.start-PADDING);rowAlignment.activeMaxY=Math.min(height-1,tableBand.last+PADDING);
    rowAlignment.activeMinX=Math.max(0,tableBand.minX-PADDING);rowAlignment.activeMaxX=Math.min(width-1,tableBand.maxX+PADDING);
  }
  const columnAlignment=choosePixelColumnMapping(before,after,width,height,rowAlignment,tableBand);
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
      if(difference>12)looseChanged++;
      if(difference>PIXEL_THRESHOLD)changed++;
    }
    const index=gy*gridWidth+gx;
    counts[index]=changed;looseCounts[index]=looseChanged;totalChanged+=changed;
    if(changed>=Math.max(2,Math.floor((x1-x0)*(y1-y0)*.2)))mask[index]=1;
  }
  // Rows that have no counterpart are the human-visible insertion/deletion bands.
  const gapRows=[],structuralRowBoxes=[];let structuralColumnBox=null;
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
      const bi=(y*width+x)*4,ay=gap.source==='after'?Math.min(height-1,(gap.afterRow||0)*ROW_STEP):mappedRowY(rowAlignment,y,height,x),ai=(ay*width+mappedColumnX(columnAlignment,x,width,y))*4;
      if(luminance(before,bi)<245||luminance(after,ai)<245){minGX=Math.min(minGX,Math.floor(x/BLOCK));maxGX=Math.max(maxGX,Math.floor(x/BLOCK));}
    }
    if(maxGX<minGX){minGX=2;maxGX=gridWidth-3;}
    const rowBox=strongestStructuralRowBox(centerY,tableBand,counts,gridWidth,gridHeight,width,height);structuralRowBoxes.push(rowBox);
    for(let gy=Math.max(0,Math.floor(rowBox.minY/BLOCK));gy<=Math.min(gridHeight-1,Math.floor(rowBox.maxY/BLOCK));gy++)for(let gx=Math.max(minGX,Math.floor(rowBox.minX/BLOCK));gx<=Math.min(maxGX,Math.floor(rowBox.maxX/BLOCK));gx++){
      const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
    }
  }
  if(rowAlignment.splitRow>=0&&Math.abs(rowAlignment.jumpPixels)>=3){
    const centerY=Math.min(height-1,rowAlignment.splitRow*ROW_STEP),rowBox=strongestStructuralRowBox(centerY,tableBand,counts,gridWidth,gridHeight,width,height);
    structuralRowBoxes.push(rowBox);
    for(let gy=Math.max(0,Math.floor(rowBox.minY/BLOCK));gy<=Math.min(gridHeight-1,Math.floor(rowBox.maxY/BLOCK));gy++)for(let gx=Math.max(0,Math.floor(rowBox.minX/BLOCK));gx<=Math.min(gridWidth-1,Math.floor(rowBox.maxX/BLOCK));gx++){
      const index=gy*gridWidth+gx;mask[index]=1;counts[index]=Math.max(counts[index],2);
    }
  }
  if((columnAlignment.adjusted||columnBoundaryChange)&&tableBand){
    let centerX,half;
    if(columnBoundaryChange){
      centerX=columnBoundaryChange.centerX;half=columnBoundaryChange.half;
    }else if(columnAlignment.structuralSplit>=0&&Math.abs(columnAlignment.structuralJump)>=3){
      centerX=Math.min(width-1,columnAlignment.structuralSplit*COLUMN_STEP);half=Math.max(6,Math.ceil(Math.abs(columnAlignment.structuralJump)/2)+3);
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
    const inkBounds=structuralColumnInkBounds(before,after,width,height,columnAlignment,centerX,half);
    const minGY=Math.max(0,Math.floor(inkBounds.minY/BLOCK)),maxGY=Math.min(gridHeight-1,Math.floor(inkBounds.maxY/BLOCK));
    const bandMinX=Math.max(0,centerX-half-PADDING),bandMaxX=Math.min(width-1,centerX+half+PADDING);
    structuralColumnBox={minX:bandMinX,maxX:bandMaxX,minY:inkBounds.minY,maxY:inkBounds.maxY,pixels:(bandMaxX-bandMinX+1)*(inkBounds.maxY-inkBounds.minY+1)};
    for(let gx=Math.max(0,Math.floor((centerX-half)/BLOCK));gx<=Math.min(gridWidth-1,Math.floor((centerX+half)/BLOCK));gx++)for(let gy=minGY;gy<=maxGY;gy++){
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
  if(structuralRowBoxes.length){
    const rowBoxes=[];
    for(const box of structuralRowBoxes)if(!rowBoxes.some(item=>Math.abs(item.minY-box.minY)<BLOCK*2))rowBoxes.push(box);
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
  }
  if(structuralColumnBox){
    // Residual antialiasing after an Excel fit-to-page scale can connect every table
    // gridline into one huge rectangle. The discontinuity itself is the useful human
    // signal, so replace such broad residuals with the actual inserted/resized band.
    if(columnBoundaryChange)components=[structuralColumnBox];
    else{
      components=components.filter(component=>{
        const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
        const componentArea=componentWidth*componentHeight;
        return !(componentWidth>width*.42||componentHeight>height*.35||componentArea>width*height*.08);
      });
      components.push(structuralColumnBox);
    }
    if(components.length>MAX_LAYOUT_REGIONS)components.sort((a,b)=>b.pixels-a.pixels).splice(MAX_LAYOUT_REGIONS);
  }
  const dominantHorizontal=components
    .filter(component=>(component.maxX-component.minX+1)>width*.2&&(component.maxY-component.minY+1)<height*.05)
    .sort((a,b)=>b.pixels-a.pixels)[0]||null;
  if(dominantHorizontal){
    components=components.filter(component=>{
      if(component===dominantHorizontal)return true;
      const componentWidth=component.maxX-component.minX+1,componentHeight=component.maxY-component.minY+1;
      const topTextJitter=component.maxY<height*.25&&componentWidth<width*.25&&componentHeight<height*.12&&component.pixels<dominantHorizontal.pixels*.8;
      return !topTextJitter;
    });
  }
  if(components.length>MAX_REGIONS)components.sort((a,b)=>b.pixels-a.pixels).splice(MAX_REGIONS);
  components.sort((a,b)=>a.minY-b.minY||a.minX-b.minX);
  let regions=components.map((component,index)=>({
    regionId:`browser-r${String(index+1).padStart(4,'0')}`,kind:'modified',x:component.minX/width,y:component.minY/height,
    width:Math.max(1,component.maxX-component.minX+1)/width,height:Math.max(1,component.maxY-component.minY+1)/height,
    confidence:Math.max(.55,Math.min(1,component.pixels/Math.max(MIN_PIXELS,(component.maxX-component.minX+1)*(component.maxY-component.minY+1)))),pixelCount:component.pixels
  }));
  let fallbackUsed=false;
  if(!regions.length){
    const fallback=buildFallbackRegion(totalChanged?counts:looseCounts,gridWidth,gridHeight,width,height);
    if(fallback){regions=[fallback];fallbackUsed=true;}
  }
  const alignmentAdjusted=rowAlignment.adjusted||columnAlignment.adjusted||!!columnBoundaryChange||gapRows.length>0;
  const columnMode=columnBoundaryChange?'column-boundary-width':columnAlignment.adjusted?(columnAlignment.split>=0?'column-scale-and-shift':'column-scale'):'column-identity';
  return {regions,changedRatio:totalChanged/Math.max(1,width*height),offsetX:columnAlignment.offset,offsetY:0,scaleX:columnAlignment.scale,scaleY:rowAlignment.scaleY,maxLocalShiftJump:Math.max(rowAlignment.maxLocalShiftJump,Math.abs(columnAlignment.jump),columnBoundaryChange?.pixelDelta||0),alignmentMode:`${rowAlignment.alignmentMode}/${columnMode}`,alignmentAdjusted,fallbackUsed};
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
