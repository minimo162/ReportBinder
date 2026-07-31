'use strict';
function pixelDifference(before,after,beforeIndex,afterIndex){
  const dr=Math.abs(before[beforeIndex]-after[afterIndex]);
  const dg=Math.abs(before[beforeIndex+1]-after[afterIndex+1]);
  const db=Math.abs(before[beforeIndex+2]-after[afterIndex+2]);
  return Math.max(dr,dg,db);
}
function alignmentScore(before,after,width,height,dx,dy){
  let total=0,count=0;
  const step=12,margin=8;
  for(let y=margin;y<height-margin;y+=step){
    const ay=y+dy;if(ay<0||ay>=height)continue;
    for(let x=margin;x<width-margin;x+=step){
      const ax=x+dx;if(ax<0||ax>=width)continue;
      const bi=(y*width+x)*4,ai=(ay*width+ax)*4;
      total+=pixelDifference(before,after,bi,ai);count++;
    }
  }
  return count?total/count:Number.MAX_VALUE;
}
function chooseAlignment(before,after,width,height){
  let bestX=0,bestY=0,bestScore=alignmentScore(before,after,width,height,0,0);
  for(let dy=-4;dy<=4;dy++)for(let dx=-4;dx<=4;dx++){
    if(!dx&&!dy)continue;
    const score=alignmentScore(before,after,width,height,dx,dy);
    if(score<bestScore){bestScore=score;bestX=dx;bestY=dy;}
  }
  return {x:bestX,y:bestY,score:bestScore};
}
function analyzeBrowserDiff(before,after,width,height){
  const threshold=24,block=4,minPixels=24,gap=2,padding=5,maxRegions=120;
  const alignment=chooseAlignment(before,after,width,height);
  const gridWidth=Math.ceil(width/block),gridHeight=Math.ceil(height/block),cellCount=gridWidth*gridHeight;
  const mask=new Uint8Array(cellCount),counts=new Uint32Array(cellCount);
  let totalChanged=0;
  for(let gy=0;gy<gridHeight;gy++)for(let gx=0;gx<gridWidth;gx++){
    let changed=0;
    const x0=gx*block,y0=gy*block,x1=Math.min(width,x0+block),y1=Math.min(height,y0+block);
    for(let y=y0;y<y1;y++)for(let x=x0;x<x1;x++){
      const ax=x+alignment.x,ay=y+alignment.y,bi=(y*width+x)*4;
      let different=false;
      if(ax<0||ax>=width||ay<0||ay>=height){
        different=before[bi]<250||before[bi+1]<250||before[bi+2]<250;
      }else{
        const ai=(ay*width+ax)*4;
        different=pixelDifference(before,after,bi,ai)>threshold;
      }
      if(different)changed++;
    }
    const index=gy*gridWidth+gx;
    counts[index]=changed;totalChanged+=changed;
    if(changed>=Math.max(2,Math.floor((x1-x0)*(y1-y0)*0.18)))mask[index]=1;
  }
  const grown=new Uint8Array(cellCount);
  for(let gy=0;gy<gridHeight;gy++)for(let gx=0;gx<gridWidth;gx++){
    const index=gy*gridWidth+gx;if(!mask[index])continue;
    for(let oy=-gap;oy<=gap;oy++)for(let ox=-gap;ox<=gap;ox++){
      const nx=gx+ox,ny=gy+oy;
      if(nx>=0&&nx<gridWidth&&ny>=0&&ny<gridHeight)grown[ny*gridWidth+nx]=1;
    }
  }
  const visited=new Uint8Array(cellCount),queue=new Int32Array(cellCount),components=[];
  for(let start=0;start<cellCount;start++){
    if(!grown[start]||visited[start])continue;
    let head=0,tail=0;queue[tail++]=start;visited[start]=1;
    let minX=width,minY=height,maxX=-1,maxY=-1,pixels=0;
    while(head<tail){
      const current=queue[head++],cx=current%gridWidth,cy=Math.floor(current/gridWidth);
      if(mask[current]){
        const x0=cx*block,y0=cy*block;
        minX=Math.min(minX,x0);minY=Math.min(minY,y0);
        maxX=Math.max(maxX,Math.min(width,x0+block)-1);maxY=Math.max(maxY,Math.min(height,y0+block)-1);
        pixels+=counts[current];
      }
      for(let oy=-1;oy<=1;oy++)for(let ox=-1;ox<=1;ox++){
        if(!ox&&!oy)continue;
        const nx=cx+ox,ny=cy+oy;
        if(nx<0||nx>=gridWidth||ny<0||ny>=gridHeight)continue;
        const next=ny*gridWidth+nx;
        if(grown[next]&&!visited[next]){visited[next]=1;queue[tail++]=next;}
      }
    }
    if(pixels<minPixels||maxX<minX||maxY<minY)continue;
    minX=Math.max(0,minX-padding);minY=Math.max(0,minY-padding);
    maxX=Math.min(width-1,maxX+padding);maxY=Math.min(height-1,maxY+padding);
    components.push({minX,minY,maxX,maxY,pixels});
  }
  if(components.length>maxRegions)components.sort((a,b)=>b.pixels-a.pixels).splice(maxRegions);
  components.sort((a,b)=>a.minY-b.minY||a.minX-b.minX);
  const regions=components.map((component,index)=>({
    regionId:`browser-r${String(index+1).padStart(4,'0')}`,
    kind:'modified',
    x:component.minX/width,
    y:component.minY/height,
    width:Math.max(1,component.maxX-component.minX+1)/width,
    height:Math.max(1,component.maxY-component.minY+1)/height,
    confidence:Math.max(0.55,Math.min(1,component.pixels/Math.max(24,(component.maxX-component.minX+1)*(component.maxY-component.minY+1)))),
    pixelCount:component.pixels
  }));
  return {regions,changedRatio:totalChanged/Math.max(1,width*height),offsetX:alignment.x,offsetY:alignment.y};
}
self.onmessage=function handleDiffWorkerMessage(event){
  const payload=event.data||{},id=payload.id;
  try{
    const width=Number(payload.width||0),height=Number(payload.height||0);
    if(width<=0||height<=0)throw new Error('画像サイズが不正です。');
    const before=new Uint8ClampedArray(payload.before),after=new Uint8ClampedArray(payload.after);
    if(before.length!==width*height*4||after.length!==width*height*4)throw new Error('画像データが不正です。');
    const result=analyzeBrowserDiff(before,after,width,height);
    self.postMessage(Object.assign({id},result));
  }catch(error){self.postMessage({id,error:String(error?.message||error)});}
};
