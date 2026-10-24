import type * as Y from 'yjs';
export const PRESENTATION_FIELDS=['text','x','y','width','height','rotation','fill','stroke','fontSize','bold','italic','underline','color','align','animation','animationTrigger','animationOrder'] as const;
/** Entrance animations ('' = none). `animationTrigger`: click starts a new step, with/after join the current one. */
export const PRESENTATION_ANIMATIONS=['appear','fade','flyBottom','flyLeft','wipe','zoom'] as const;
export type PresentationAnimation=typeof PRESENTATION_ANIMATIONS[number];
export const PRESENTATION_ANIMATION_LABELS:Record<PresentationAnimation,string>={appear:'出现',fade:'淡入',flyBottom:'飞入（自底部）',flyLeft:'飞入（自左侧）',wipe:'擦除',zoom:'缩放'};
export const PRESENTATION_TRIGGERS=['click','with','after'] as const;
export type PresentationTrigger=typeof PRESENTATION_TRIGGERS[number];
export const PRESENTATION_TRIGGER_LABELS:Record<PresentationTrigger,string>={click:'单击时',with:'与上一动画同时',after:'上一动画之后'};
export const PRESENTATION_ANIMATION_MS=500;
/** 出现 is instant; the other entrance effects run for PRESENTATION_ANIMATION_MS. */
export const presentationEffectDuration=(effect:PresentationAnimation)=>effect==='appear'?0:PRESENTATION_ANIMATION_MS;
export type PresentationField=typeof PRESENTATION_FIELDS[number];
/** Character/paragraph formatting applies to the whole text box (one run style per object). */
export const PRESENTATION_TEXT_FIELDS=['fontSize','bold','italic','underline','color','align'] as const;
export const PRESENTATION_ALIGNS=['left','center','right'] as const;
export const DEFAULT_PRESENTATION_FONT_SIZE=18;
export interface PresentationFieldOperation {action:'setField';operationId:string;officeGeneration?:string;slide:string;object:string;field:PresentationField;value:string|number|boolean;base:string;}
export function presentationObjectIdentity(object:Y.Map<unknown>){const id=object._item?.id;if(!id)throw new Error('对象尚未保存');return id.client+'-'+id.clock;}
export function presentationFieldRevision(object:Y.Map<unknown>,field:PresentationField){const id=presentationObjectIdentity(object),doc=object.doc!;return JSON.stringify([id,field,object.get(field)??null,doc.share.has('presentationFields')?doc.getMap('presentationFields').get(id+':'+field)??'':'']);}
export const PRESENTATION_OBJECT_TYPES=['text','rect','ellipse','line','image','roundRect','triangle','diamond','rightArrow','star5'] as const;
export type PresentationObjectType=typeof PRESENTATION_OBJECT_TYPES[number];
export const PRESENTATION_SHAPE_LABELS:Record<string,string>={text:'文本框',rect:'矩形',ellipse:'圆形',line:'直线',image:'图片',roundRect:'圆角矩形',triangle:'三角形',diamond:'菱形',rightArrow:'右箭头',star5:'五角星'};
export type PresentationObjectGeometry=Partial<Record<'x'|'y'|'width'|'height',number>>;
export type PresentationArrange='front'|'back'|'forward'|'backward';
export type PresentationObjectOperation={action:'addObject';operationId:string;officeGeneration?:string;slide:string;type:PresentationObjectType;assetId?:string;geometry?:PresentationObjectGeometry}|{action:'deleteObject';operationId:string;officeGeneration?:string;slide:string;object:string;base:string}|{action:'arrangeObject';operationId:string;officeGeneration?:string;slide:string;object:string;to:PresentationArrange;base:string};
/** Whole-object revision: a delete only applies to the exact object state the author saw. */
export function presentationObjectRevision(object:Y.Map<unknown>){return JSON.stringify([presentationObjectIdentity(object),object.get('type')??null,object.get('assetId')??null,...PRESENTATION_FIELDS.map(field=>presentationFieldRevision(object,field))]);}
/** Per-slide object identities, order and immutable attributes; only the object endpoint may change it. */
export function presentationObjectLayout(slides:{id:string;objects:Y.Array<Y.Map<unknown>>}[],orphan?:Y.Array<Y.Map<unknown>>){return JSON.stringify([...slides.map(slide=>[slide.id,slide.objects.toArray().map(object=>[presentationObjectIdentity(object),object.get('type')??null,object.get('assetId')??null,object.get('z')??0])]),orphan?.toArray().map(object=>presentationObjectIdentity(object))??[]]);}
/** Paint order: ascending `z`, ties keep array order. Indices refer to the slide's object array. */
export function presentationPaintOrder(objects:Y.Map<unknown>[]){return objects.map((object,index)=>({object,index,z:Number(object.get('z')??0)})).sort((a,b)=>a.z-b.z||a.index-b.index);}
/** Layer revision of one slide: an arrange request only applies to the stacking order its author saw. */
export function presentationArrangeRevision(objects:Y.Map<unknown>[]){return JSON.stringify(presentationPaintOrder(objects).map(entry=>presentationObjectIdentity(entry.object)));}
/** New `z` values (by array index) after moving one object; every object on the slide gets a dense 0..n-1 rank. */
export function arrangePresentationObjects(objects:Y.Map<unknown>[],index:number,to:PresentationArrange){
 const order=presentationPaintOrder(objects).map(entry=>entry.index),at=order.indexOf(index);if(at<0)throw new Error('对象已不存在');
 order.splice(at,1);const target=to==='front'?order.length:to==='back'?0:to==='forward'?Math.min(order.length,at+1):Math.max(0,at-1);order.splice(target,0,index);
 return new Map(order.map((value,rank)=>[value,rank]));
}
export function validPresentationField(field:unknown,value:unknown){
 if(!PRESENTATION_FIELDS.includes(field as PresentationField))return false;
 if(field==='text')return typeof value==='string'&&value.length<=100000;
 if(field==='fill'||field==='stroke'||field==='color')return typeof value==='string'&&/^[0-9A-Fa-f]{6}$/.test(value);
 if(field==='bold'||field==='italic'||field==='underline')return typeof value==='boolean';
 if(field==='align')return PRESENTATION_ALIGNS.includes(value as never);
 if(field==='fontSize')return typeof value==='number'&&Number.isFinite(value)&&value>=6&&value<=400;
 if(field==='animation')return value===''||PRESENTATION_ANIMATIONS.includes(value as never);
 if(field==='animationTrigger')return PRESENTATION_TRIGGERS.includes(value as never);
 if(field==='animationOrder')return Number.isSafeInteger(value)&&Number(value)>=0&&Number(value)<=100000;
 return typeof value==='number'&&Number.isFinite(value)&&Math.abs(value)<=(field==='rotation'?3600:100)&&(!['width','height'].includes(String(field))||value>=0);
}
/**
 * A slide's animation sequence: animated objects by `animationOrder` (then array order), grouped into steps. A click
 * item starts a step; with/after items join the current step (leading ones form step 0, which plays on entering the
 * slide). Within a step, `after` waits until everything started so far in the step has finished; `with` shares the
 * start of the previous item.
 */
export type PresentationBuild={object:Y.Map<unknown>;index:number;effect:PresentationAnimation;trigger:PresentationTrigger;step:number;delay:number};
export function presentationBuilds(objects:Y.Map<unknown>[]){
 const items=objects.map((object,index)=>({object,index,effect:object.get('animation') as PresentationAnimation,trigger:(PRESENTATION_TRIGGERS.includes(object.get('animationTrigger') as never)?object.get('animationTrigger'):'click') as PresentationTrigger,order:Number(object.get('animationOrder')??0)})).filter(item=>PRESENTATION_ANIMATIONS.includes(item.effect)).sort((a,b)=>a.order-b.order||a.index-b.index);
 const builds:PresentationBuild[]=[];let step=0,delay=0,end=0;
 for(const item of items){
  const previous=builds[builds.length-1],same=!!previous&&previous.step===step;
  if(item.trigger==='click'){step++;delay=0;end=0;}else if(item.trigger==='after')delay=same?end:0;else delay=same?previous.delay:0;
  end=Math.max(end,delay+presentationEffectDuration(item.effect));
  builds.push({object:item.object,index:item.index,effect:item.effect,trigger:item.trigger,step,delay});
 }
 return builds;
}
