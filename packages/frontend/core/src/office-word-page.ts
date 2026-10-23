import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {setWordParagraphAttribute} from './office-word-format';
import {DEFAULT_WORD_PAGE,validWordPage,wordPageKey,type WordPageLayout} from '../../../common/nbstore/src/office-word-page';
import {validWordHeaderFooter,parseWordHeaderImageKey,WORD_HEADER_IMAGE_MAX,WORD_HEADER_LINES,type WordHeaderFooter} from '../../../common/nbstore/src/office-word-header';
import {createWordImage,isWordImage,wordHeaderImages} from '../../../common/nbstore/src/office-image';
export function getWordPage(doc:Y.Doc):WordPageLayout{
 const value=doc.getMap('pageSettings').get('layout');if(value===undefined)return {...DEFAULT_WORD_PAGE};if(!validWordPage(value))throw new Error('页面设置无效');return {...value};
}
export function setWordPage(doc:Y.Doc,page:WordPageLayout,base?:WordPageLayout){
 if(!validWordPage(page))throw new Error('页面尺寸或页边距无效，请保留至少 18 磅正文区域');
 if(base&&wordPageKey(getWordPage(doc))!==wordPageKey(base))throw new Error('页面设置已被其他成员修改，请重新载入后再保存');
 doc.transact(()=>doc.getMap('pageSettings').set('layout',{...page}),LOCAL_ORIGIN);
}
export function getWordHeaderFooter(doc:Y.Doc):WordHeaderFooter|undefined{
 const value=doc.share.has('pageSettings')?doc.getMap('pageSettings').get('headerFooter'):undefined;return validWordHeaderFooter(value)?value:undefined;
}
export type WordHeaderPicture={assetId:string;width:number;height:number;alt:string};
/** Header/footer pictures keyed by `wordHeaderImageKey`; pictures of a variant that is switched off are left out. */
export function getWordHeaderImages(doc:Y.Doc,value=getWordHeaderFooter(doc)):Record<string,WordHeaderPicture>{
 const result:Record<string,WordHeaderPicture>={};if(!value)return result;
 for(const [key,image]of wordHeaderImages(doc)?.entries()??[]){const place=parseWordHeaderImageKey(key);if(!place||!isWordImage(image)||place.scope!=='main'&&!value[place.scope])continue;result[key]={assetId:String(image.get('assetId')),width:Number(image.get('width')),height:Number(image.get('height')),alt:String(image.get('alt'))};}
 return result;
}
/**
 * `undefined` removes headers, footers and their pictures; `base`/`baseImages` guard against overwriting another
 * member's change. Pictures are written in the same transaction, and unused ones are removed.
 */
export function setWordHeaderFooter(doc:Y.Doc,value:WordHeaderFooter|undefined,base:WordHeaderFooter|undefined,images:Record<string,WordHeaderPicture>={},baseImages=getWordHeaderImages(doc,base)){
 if(value!==undefined&&!validWordHeaderFooter(value))throw new Error(`页眉页脚内容无效，每处最多 255 个字符、${WORD_HEADER_LINES} 行`);
 if(JSON.stringify(getWordHeaderFooter(doc)??null)!==JSON.stringify(base??null)||JSON.stringify(getWordHeaderImages(doc))!==JSON.stringify(baseImages))throw new Error('页眉页脚已被其他成员修改，请重新载入后再保存');
 const kept=Object.entries(value?images:{}).filter(([key])=>{const place=parseWordHeaderImageKey(key);return !!place&&(place.scope==='main'||!!value?.[place.scope]);});
 for(const [,image]of kept)if(!(image.width>0&&image.height>0&&image.width<=WORD_HEADER_IMAGE_MAX.width&&image.height<=WORD_HEADER_IMAGE_MAX.height))throw new Error('页眉页脚图片尺寸无效');
 doc.transact(()=>{
  const map=doc.getMap('pageSettings');if(value)map.set('headerFooter',value);else map.delete('headerFooter');
  if(!kept.length){map.delete('headerImages');return;}
  let pictures=wordHeaderImages(doc);if(!pictures){pictures=new Y.Map<Y.Map<unknown>>();map.set('headerImages',pictures);}
  const next=new Map(kept);for(const key of [...pictures.keys()])if(!next.has(key))pictures.delete(key);
  // Unchanged pictures keep their map, so a concurrent resize by someone else is not overwritten.
  for(const [key,image]of next){const current=pictures.get(key);if(isWordImage(current)&&current.get('assetId')===image.assetId&&current.get('width')===image.width&&current.get('height')===image.height&&current.get('alt')===image.alt)continue;pictures.set(key,createWordImage(image.assetId,image.width,image.height,image.alt));}
 },LOCAL_ORIGIN);
}
export function setWordPageBreak(text:Y.Text,value:boolean){
 if(typeof value!=='boolean')throw new Error('段前分页设置无效');if(!text.doc||text.doc.getArray<Y.Text>('paragraphs').toArray().indexOf(text)<0)throw new Error('段落已删除');
 text.doc.transact(()=>setWordParagraphAttribute(text,'pageBreakBefore',value),LOCAL_ORIGIN);
}
