import type * as Y from 'yjs';
export const PRESENTATION_FIELDS=['text','x','y','width','height','rotation','fill','stroke'] as const;
export type PresentationField=typeof PRESENTATION_FIELDS[number];
export interface PresentationFieldOperation {action:'setField';operationId:string;officeGeneration?:string;slide:string;object:string;field:PresentationField;value:string|number;base:string;}
export function presentationObjectIdentity(object:Y.Map<unknown>){const id=object._item?.id;if(!id)throw new Error('对象尚未保存');return id.client+'-'+id.clock;}
export function presentationFieldRevision(object:Y.Map<unknown>,field:PresentationField){const id=presentationObjectIdentity(object),doc=object.doc!;return JSON.stringify([id,field,object.get(field)??null,doc.share.has('presentationFields')?doc.getMap('presentationFields').get(id+':'+field)??'':'']);}
export const PRESENTATION_OBJECT_TYPES=['text','rect','ellipse','line','image'] as const;
export type PresentationObjectType=typeof PRESENTATION_OBJECT_TYPES[number];
export type PresentationObjectGeometry=Partial<Record<'x'|'y'|'width'|'height',number>>;
export type PresentationObjectOperation={action:'addObject';operationId:string;officeGeneration?:string;slide:string;type:PresentationObjectType;assetId?:string;geometry?:PresentationObjectGeometry}|{action:'deleteObject';operationId:string;officeGeneration?:string;slide:string;object:string;base:string};
/** Whole-object revision: a delete only applies to the exact object state the author saw. */
export function presentationObjectRevision(object:Y.Map<unknown>){return JSON.stringify([presentationObjectIdentity(object),object.get('type')??null,object.get('assetId')??null,...PRESENTATION_FIELDS.map(field=>presentationFieldRevision(object,field))]);}
/** Per-slide object identities, order and immutable attributes; only the object endpoint may change it. */
export function presentationObjectLayout(slides:{id:string;objects:Y.Array<Y.Map<unknown>>}[],orphan?:Y.Array<Y.Map<unknown>>){return JSON.stringify([...slides.map(slide=>[slide.id,slide.objects.toArray().map(object=>[presentationObjectIdentity(object),object.get('type')??null,object.get('assetId')??null])]),orphan?.toArray().map(object=>presentationObjectIdentity(object))??[]]);}
export function validPresentationField(field:unknown,value:unknown){
 if(!PRESENTATION_FIELDS.includes(field as PresentationField))return false;
 if(field==='text')return typeof value==='string'&&value.length<=100000;
 if(field==='fill'||field==='stroke')return typeof value==='string'&&/^[0-9A-Fa-f]{6}$/.test(value);
 return typeof value==='number'&&Number.isFinite(value)&&Math.abs(value)<=(field==='rotation'?3600:100)&&(!['width','height'].includes(String(field))||value>=0);
}
