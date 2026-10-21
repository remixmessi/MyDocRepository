import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {MAX_OFFICE_TEXT_LENGTH} from '../../../common/nbstore/src/office';
export const MAX_TEXT_MATCHES=10_000;
export type TextRange={start:number;end:number};
export type TextMatch={start:Y.RelativePosition;end:Y.RelativePosition;expected:string};
function validUnicode(value:string){
 for(let i=0;i<value.length;i++){const unit=value.charCodeAt(i);if(unit>=0xd800&&unit<=0xdbff){const next=value.charCodeAt(++i);if(!(next>=0xdc00&&next<=0xdfff))return false;}else if(unit>=0xdc00&&unit<=0xdfff)return false;}return true;
}
export function findTextRanges(source:string,query:string){
 const ranges:TextRange[]=[];
 if(!query||!validUnicode(query))return {ranges,truncated:false};
 let offset=0;
 while(offset<=source.length-query.length){const start=source.indexOf(query,offset);if(start<0)break;if(ranges.length===MAX_TEXT_MATCHES)return {ranges,truncated:true};ranges.push({start,end:start+query.length});offset=start+query.length;}
 return {ranges,truncated:false};
}
export function captureTextMatch(text:Y.Text,range:TextRange):TextMatch{
 return {start:Y.createRelativePositionFromTypeIndex(text,range.start,0),end:Y.createRelativePositionFromTypeIndex(text,range.end,-1),expected:text.toString().slice(range.start,range.end)};
}
export function resolveTextMatch(text:Y.Text,match:TextMatch):TextRange|undefined{
 if(!text.doc)return;
 const start=Y.createAbsolutePositionFromRelativePosition(match.start,text.doc),end=Y.createAbsolutePositionFromRelativePosition(match.end,text.doc);
 if(!start||!end||start.type!==text||end.type!==text||!match.expected||text.toString().slice(start.index,end.index)!==match.expected)return;
 return {start:start.index,end:end.index};
}
export function replaceTextMatches(text:Y.Text,query:string,replacement:string,match?:TextMatch){
 if(!text.doc||!query||!validUnicode(query)||!validUnicode(replacement))throw new Error('查找或替换文本包含无效字符');
 const source=text.toString();let ranges:TextRange[];
 if(match){const range=resolveTextMatch(text,match);if(!range||match.expected!==query)throw new Error('匹配内容已改变，请重新查找');ranges=[range];}
 else{const result=findTextRanges(source,query);if(result.truncated)throw new Error('匹配超过 1 万处，请缩小查找范围');ranges=result.ranges;}
 const length=source.length+ranges.length*(replacement.length-query.length);
 if(length>MAX_OFFICE_TEXT_LENGTH)throw new Error('替换后文本超过 100 万字符');
 if(!ranges.length||query===replacement)return 0;
 text.doc.transact(()=>{for(let i=ranges.length-1;i>=0;i--){const range=ranges[i];text.delete(range.start,range.end-range.start);if(replacement)text.insert(range.start,replacement);}},LOCAL_ORIGIN);
 return ranges.length;
}
