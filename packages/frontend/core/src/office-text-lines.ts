/** Logical LF-delimited lines. The last LF always starts a final empty line. */
export function textLineStarts(value:string){
 let count=1;for(let i=0;i<value.length;i++)if(value.charCodeAt(i)===10)count++;
 const starts=new Uint32Array(count);let line=1;
 for(let i=0;i<value.length;i++)if(value.charCodeAt(i)===10)starts[line++]=i+1;
 return starts;
}
function boundary(value:string,offset:number){
 let position=Number.isFinite(offset)?Math.max(0,Math.min(value.length,Math.trunc(offset))):0;
 if(position>0&&position<value.length&&value.charCodeAt(position)>=0xdc00&&value.charCodeAt(position)<=0xdfff&&value.charCodeAt(position-1)>=0xd800&&value.charCodeAt(position-1)<=0xdbff)position--;
 return position;
}
export function textCharacterCount(value:string,start=0,end=value.length){
 const stop=boundary(value,end);let count=0;
 for(let i=boundary(value,start);i<stop;i++,count++){const unit=value.charCodeAt(i),next=value.charCodeAt(i+1);if(unit>=0xd800&&unit<=0xdbff&&next>=0xdc00&&next<=0xdfff&&i+1<stop)i++;}
 return count;
}
export function textLinePosition(value:string,starts:Uint32Array,offset:number){
 const position=boundary(value,offset);let low=0,high=starts.length-1;
 while(low<high){const middle=Math.ceil((low+high)/2);if(starts[middle]<=position)low=middle;else high=middle-1;}
 return {line:low+1,column:textCharacterCount(value,starts[low],position)+1};
}
export function visibleTextLines(count:number,scrollTop:number,height:number,lineHeight:number,paddingTop:number){
 const size=Number.isFinite(lineHeight)&&lineHeight>0?lineHeight:25.5;
 const top=Math.max(0,Number.isFinite(scrollTop)?scrollTop:0),viewport=Math.max(0,Number.isFinite(height)?height:0),padding=Math.max(0,Number.isFinite(paddingTop)?paddingTop:0);
 const first=Math.max(0,Math.floor((top-padding)/size)-3),last=Math.min(count,Math.ceil((top+viewport-padding)/size)+3);
 return {first:Math.min(first,count),last:Math.max(0,last)};
}
