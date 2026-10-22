import * as Y from 'yjs';
export const LEGACY_SLIDE='first';
export const MAX_PRESENTATION_SLIDES=200;
export interface PresentationSlide {id:string;objects:Y.Array<Y.Map<unknown>>;notes:string;}
export function presentationSlides(doc:Y.Doc):PresentationSlide[]{
 const order=doc.share.has('slideOrder')?doc.getArray<string>('slideOrder').toArray():[LEGACY_SLIDE];
 return order.map(id=>{const slide=id===LEGACY_SLIDE?undefined:doc.getMap<Y.Map<unknown>>('slides').get(id);return {id,objects:id===LEGACY_SLIDE?doc.getArray<Y.Map<unknown>>('objects'):slide!.get('objects') as Y.Array<Y.Map<unknown>>,notes:String(id===LEGACY_SLIDE?(doc.share.has('presentation')?doc.getMap('presentation').get('notes')??'':''):slide!.get('notes')??'')};});
}
export function presentationSize(doc:Y.Doc){const meta=doc.share.has('presentation')?doc.getMap('presentation'):undefined;return {width:Number(meta?.get('width')??12192000),height:Number(meta?.get('height')??6858000)};}
export function createPresentationObject(type='text'){
 const object=new Y.Map<unknown>();for(const [key,value]of Object.entries({type,x:0.1,y:0.1,width:0.5,height:0.2,rotation:0,fill:'FFFFFF',stroke:'333333',text:''}))object.set(key,value);return object;
}
export function changePresentationSlide(doc:Y.Doc,action:'add'|'duplicate'|'delete'|'up'|'down',id:string){
 const current=presentationSlides(doc),index=current.findIndex(slide=>slide.id===id);if(index<0)throw new Error('幻灯片已不存在');
 if((action==='add'||action==='duplicate')&&current.length>=MAX_PRESENTATION_SLIDES)throw new Error('幻灯片最多 200 张');if(action==='delete'&&current.length===1)throw new Error('至少保留一张幻灯片');
 let selected=id;doc.transact(()=>{
  const order=doc.getArray<string>('slideOrder');if(!order.length)order.push([LEGACY_SLIDE]);
  if(action==='add'||action==='duplicate'){
   selected=crypto.randomUUID();const slide=new Y.Map<unknown>(),objects=new Y.Array<Y.Map<unknown>>();slide.set('objects',objects);slide.set('notes',action==='duplicate'?current[index].notes:'');
   if(action==='duplicate')for(const object of current[index].objects)objects.push([object.clone()]);
   doc.getMap<Y.Map<unknown>>('slides').set(selected,slide);order.insert(index+1,[selected]);
  }else if(action==='delete'){order.delete(index,1);if(id===LEGACY_SLIDE){const objects=doc.getArray('objects');objects.delete(0,objects.length);if(doc.share.has('presentation'))doc.getMap('presentation').delete('notes');}else doc.getMap('slides').delete(id);selected=order.get(Math.min(index,order.length-1));}
  else{const target=action==='up'?index-1:index+1;if(target<0||target>=order.length)return;order.delete(index,1);order.insert(target,[id]);}
 });return selected;
}
export function setPresentationNotes(doc:Y.Doc,id:string,notes:string){if(id===LEGACY_SLIDE)doc.getMap('presentation').set('notes',notes);else doc.getMap<Y.Map<unknown>>('slides').get(id)?.set('notes',notes);}
