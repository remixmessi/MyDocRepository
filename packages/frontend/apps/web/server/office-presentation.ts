import {presentationObjectIdentity,presentationObjectLayout,PRESENTATION_FIELDS} from '../../../../common/nbstore/src/presentation-field';
import {validateOfficeComments} from '../../../../common/nbstore/src/office-comments';
import {packageImageReader,type OfficeMedia,validateOfficeMedia} from './office-media';
import {validOfficeAssetId} from '../../../../common/nbstore/src/office-image';
import * as Y from 'yjs';
import {posix} from 'node:path';
import {readOfficePackage} from './office-package';
import {xlsxXml,attribute,child,children,type XlsxXml} from './office-xlsx-xml';
import {presentationSlides,MAX_PRESENTATION_SLIDES,LEGACY_SLIDE,presentationSize,createPresentationObject} from '../../../../common/nbstore/src/office-presentation';
export const MAX_SLIDE_OBJECTS=10_000;
export function emptyPresentationUpdate(){const doc=new Y.Doc();doc.getArray<Y.Map<unknown>>('objects');return {doc,update:Y.encodeStateAsUpdate(doc)};}
const descendants=(node:XlsxXml,name:string):XlsxXml[]=>node.children.flatMap(value=>[...(value.name===name?[value]:[]),...descendants(value,name)]);
function target(base:string,value:string|undefined){if(!value||/[\\:#?]/.test(value))throw new Error('PPTX 关系路径无效');const path=value.startsWith('/')?value.slice(1):posix.normalize(posix.join(base,value));if(!path.startsWith('ppt/')||path.split('/').includes('..'))throw new Error('PPTX 关系路径无效');return path;}
/** Placeholder geometry and solid fill/line inherit slide → layout → master (ECMA-376 19.3.1.36); other styling is not modelled. */
function relationTargets(files:Record<string,Uint8Array>,path:string,type:string){const rels=posix.join(posix.dirname(path),'_rels',posix.basename(path)+'.rels');if(!files[rels])return [];return children(xlsxXml(files[rels]),'Relationship').filter(rel=>attribute(rel,'Type')?.endsWith('/'+type)).map(rel=>{if(attribute(rel,'TargetMode')==='External')throw new Error('PPTX 不支持外部版式或母版');return target(posix.dirname(path),attribute(rel,'Target'));});}
type Placeholder={type:string;idx?:string;shape:XlsxXml};
function placeholderOf(shape:XlsxXml){const ph=descendants(shape,'ph')[0];return ph?{type:attribute(ph,'type')??'obj',idx:attribute(ph,'idx')}:undefined;}
function placeholders(files:Record<string,Uint8Array>,path:string|undefined):Placeholder[]{if(!path||!files[path])return [];const root=xlsxXml(files[path]),tree=descendants(root,'spTree')[0]??root;return tree.children.filter(node=>node.name==='sp').flatMap(shape=>{const key=placeholderOf(shape);return key?[{...key,shape}]:[];});}
const masterType=(type:string)=>type==='ctrTitle'?'title':['subTitle','obj','body'].includes(type)?'body':type;
function inheritedShapes(key:{type:string;idx?:string}|undefined,layout:Placeholder[],master:Placeholder[]){
 if(!key)return [];const fromLayout=(key.idx!==undefined?layout.find(p=>p.idx===key.idx):undefined)??layout.find(p=>p.type===key.type)??layout.find(p=>masterType(p.type)===masterType(key.type)&&key.idx===undefined);
 const fromMaster=master.find(p=>p.type===masterType(fromLayout?.type??key.type))??master.find(p=>masterType(p.type)===masterType(fromLayout?.type??key.type));return [fromLayout?.shape,fromMaster?.shape].filter((shape):shape is XlsxXml=>!!shape);
}
export function encodePptxPresentation(bytes:Uint8Array){
 const files=readOfficePackage(bytes),doc=new Y.Doc();try{
  const media:OfficeMedia[]=[],imageRegistry=new Map<string,OfficeMedia>();let mediaBytes=0;let paths:string[],width=12192000,height=6858000;
  if(files['ppt/presentation.xml']){
   const root=xlsxXml(files['ppt/presentation.xml']),rels=files['ppt/_rels/presentation.xml.rels'];if(root.name!=='presentation'||!rels)throw new Error('PPTX 演示结构或关系缺失');
   const relRoot=xlsxXml(rels);if(relRoot.name!=='Relationships')throw new Error('PPTX 演示关系无效');const links=children(relRoot,'Relationship'),ids=child(root,'sldIdLst');const seen=new Set<string>();
   paths=(ids?children(ids,'sldId'):[]).map(slide=>{const relationIds=Object.entries(slide.attributes).filter(([key])=>key.includes(':')&&key.endsWith(':id'));if(relationIds.length!==1)throw new Error('PPTX 幻灯片关系编号无效');const id=relationIds[0][1];const matches=links.filter(link=>attribute(link,'Id')===id);if(!id||seen.has(id)||matches.length!==1||attribute(matches[0],'TargetMode')==='External'||!attribute(matches[0],'Type')?.endsWith('/slide'))throw new Error('PPTX 幻灯片关系无效');seen.add(id);return target('ppt',attribute(matches[0],'Target'));});
   const size=child(root,'sldSz');if(size){width=Number(attribute(size,'cx'));height=Number(attribute(size,'cy'));}
  }else paths=Object.keys(files).filter(key=>/^ppt\/slides\/slide\d+\.xml$/.test(key)).sort((a,b)=>Number(/slide(\d+)/.exec(a)![1])-Number(/slide(\d+)/.exec(b)![1]));
  if(!paths.length||paths.length>MAX_PRESENTATION_SLIDES||new Set(paths).size!==paths.length)throw new Error('PPTX 幻灯片数量无效');
  if(!Number.isSafeInteger(width)||!Number.isSafeInteger(height)||width<1||height<1||width>100000000||height>100000000)throw new Error('PPTX 页面尺寸无效');
  doc.getMap('presentation').set('width',width);doc.getMap('presentation').set('height',height);
  for(const [index,path]of paths.entries()){
   if(!files[path])throw new Error('PPTX 工作页缺失');const root=xlsxXml(files[path]);if(root.name!=='sld')throw new Error('PPTX 幻灯片结构无效');
   const objects=index===0?doc.getArray<Y.Map<unknown>>('objects'):new Y.Array<Y.Map<unknown>>(),id=index===0?LEGACY_SLIDE:'slide-'+index;
   if(index){const slide=new Y.Map<unknown>();slide.set('objects',objects);slide.set('notes','');doc.getMap('slides').set(id,slide);}doc.getArray('slideOrder').push([id]);
   const images=packageImageReader(files,path,imageRegistry);
   const tree=descendants(root,'spTree')[0]??root;if(descendants(tree,'grpSp').some(group=>descendants(group,'pic').length))throw new Error('组合内图片尚未支持，请先取消组合');
   const layoutPath=relationTargets(files,path,'slideLayout')[0],masterPath=layoutPath?relationTargets(files,layoutPath,'slideMaster')[0]:undefined,layout=placeholders(files,layoutPath),master=placeholders(files,masterPath);
   for(const shape of tree.children.filter(node=>['sp','pic'].includes(node.name))){
    const properties=child(shape,'spPr'),geometry=properties?child(properties,'prstGeom'):undefined,preset=geometry?attribute(geometry,'prst'):undefined;
    const paragraphs=descendants(shape,'p'),text=paragraphs.length?paragraphs.map(p=>descendants(p,'t').map(t=>t.text).join('')).join('\n'):descendants(shape,'t').map(t=>t.text).join('');
    if(!text&&preset&&!['rect','ellipse','line'].includes(preset))continue;
    const image=shape.name==='pic';
    const object=createPresentationObject(image?'image':text||child(shape,'txBody')?'text':preset==='ellipse'?'ellipse':preset==='line'?'line':'rect');object.set('text',text);if(image){const blip=descendants(shape,'blip')[0],embed=blip&&Object.entries(blip.attributes).find(([key])=>key.endsWith(':embed'))?.[1];if(!embed)throw new Error('PPTX 图片缺少内嵌关系');object.set('assetId',images.read(embed));}
    const sources=[shape,...inheritedShapes(placeholderOf(shape),layout,master)],spPr=(node:XlsxXml)=>child(node,'spPr');
    const transform=sources.map(node=>{const own=spPr(node);return own?child(own,'xfrm'):undefined;}).find(node=>node&&child(node,'off')&&child(node,'ext')),off=transform?child(transform,'off'):undefined,ext=transform?child(transform,'ext'):undefined;
    if(off){object.set('x',Number(attribute(off,'x'))/width);object.set('y',Number(attribute(off,'y'))/height);}if(ext){object.set('width',Number(attribute(ext,'cx'))/width);object.set('height',Number(attribute(ext,'cy'))/height);}if(transform&&attribute(transform,'rot'))object.set('rotation',Number(attribute(transform,'rot'))/60000);
    // The nearest level that states a fill (solid or none) decides it; only explicit RGB is carried over.
    for(const field of ['fill','stroke']){for(const source of sources){const own=spPr(source),node=field==='fill'?own:own?child(own,'ln'):undefined;if(!node||!child(node,'solidFill')&&!child(node,'noFill'))continue;const rgb=child(node,'solidFill')?child(child(node,'solidFill')!,'srgbClr'):undefined;if(rgb)object.set(field,attribute(rgb,'val'));break;}}
    objects.push([object]);if(objects.length>MAX_SLIDE_OBJECTS)throw new Error('幻灯片对象超过 1 万');
   }
   mediaBytes+=images.media.reduce((total,value)=>total+Buffer.from(value.data,'base64').length,0);if(mediaBytes>8*1024*1024||media.length+images.media.length>1000)throw new Error('办公图片总量或数量超限');media.push(...images.media);
   const relPath=posix.join(posix.dirname(path),'_rels',posix.basename(path)+'.rels');if(files[relPath]){const notes=children(xlsxXml(files[relPath]),'Relationship').filter(rel=>attribute(rel,'Type')?.endsWith('/notesSlide'));if(notes.length>1)throw new Error('PPTX 备注关系重复');if(notes[0]){if(attribute(notes[0],'TargetMode')==='External')throw new Error('PPTX 不支持外部备注');const notesPath=target(posix.dirname(path),attribute(notes[0],'Target'));if(!files[notesPath])throw new Error('PPTX 备注缺失');const notesRoot=xlsxXml(files[notesPath]);if(notesRoot.name!=='notes')throw new Error('PPTX 备注结构无效');const note=descendants(notesRoot,'sp').filter(sp=>!descendants(sp,'ph').some(ph=>['sldNum','sldImg'].includes(attribute(ph,'type')??''))).flatMap(sp=>descendants(sp,'p').map(p=>descendants(p,'t').map(t=>t.text).join(''))).join('\n');if(index===0)doc.getMap('presentation').set('notes',note);else doc.getMap<Y.Map<unknown>>('slides').get(id)!.set('notes',note);}}
  }
  validateOfficeMedia(media);const update=Y.encodeStateAsUpdate(doc);validatePresentationUpdate(null,update);return {doc,update,media};
 }catch(error){doc.destroy();throw error;}
}
export function validatePresentationUpdate(current:Uint8Array|null,update:Uint8Array){
 const doc=new Y.Doc();try{
  if(current)Y.applyUpdate(doc,current);Y.applyUpdate(doc,update);validateOfficeComments(doc);if([...doc.share.keys()].some(key=>!['officeComments','officeCommentOperations','objects','slideOrder','slides','presentation','presentationOperations','presentationFields'].includes(key)))throw new Error('演示文稿更新包含非法根结构');
  if(doc.share.has('presentation')){const meta=doc.getMap('presentation');if(['width','height'].some(key=>meta.has(key)&&typeof meta.get(key)!=='number')||[...meta.keys()].some(key=>!['width','height','notes'].includes(key))||meta.has('notes')&&(typeof meta.get('notes')!=='string'||String(meta.get('notes')).length>100000))throw new Error('演示设置无效');}
  const size=presentationSize(doc);if(!Object.values(size).every(n=>Number.isSafeInteger(n)&&n>=1&&n<=100000000))throw new Error('演示尺寸无效');
  if(doc.share.has('slideOrder')){const order=doc.getArray('slideOrder').toArray();if(!order.length||order.length>MAX_PRESENTATION_SLIDES||order.some(id=>typeof id!=='string'||!/^[a-zA-Z0-9-]{1,80}$/.test(id))||new Set(order).size!==order.length)throw new Error('幻灯片顺序无效');const slides=doc.getMap('slides');if([...slides.keys()].some(id=>id===LEGACY_SLIDE||!order.includes(id)))throw new Error('幻灯片引用无效');for(const id of order){if(id===LEGACY_SLIDE)continue;const slide=slides.get(id as string);if(!(slide instanceof Y.Map)||[...slide.keys()].some(key=>!['objects','notes'].includes(key))||!(slide.get('objects') instanceof Y.Array)||typeof slide.get('notes')!=='string'||slide.get('notes').length>100000)throw new Error('幻灯片结构无效');}}
  else if(doc.share.has('slides')&&doc.getMap('slides').size)throw new Error('幻灯片缺少顺序');
  if(doc.share.has('presentationOperations')){const receipts=doc.getMap('presentationOperations');if(receipts.size>10000||[...receipts.values()].some(value=>typeof value!=='string'||! /^[a-f0-9]{64}:[a-zA-Z0-9-]{1,80}$/.test(value)))throw new Error('演示操作回执无效');}
  if(doc.share.has('presentationFields')){const fields=doc.getMap('presentationFields');if(fields.size>80000||[...fields].some(([key,value])=>!/^\d+-\d+:(text|x|y|width|height|rotation|fill|stroke)$/.test(key)||typeof value!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(value)))throw new Error('演示字段修订无效');}
  if(current){const previous=new Y.Doc();try{Y.applyUpdate(previous,current);
   if(JSON.stringify(previous.share.has('presentationFields')?previous.getMap('presentationFields').toJSON():{})!==JSON.stringify(doc.share.has('presentationFields')?doc.getMap('presentationFields').toJSON():{}))throw new Error('请通过专用接口修改演示字段修订');
   if(previous.share.has('presentationFields')){const objects=new Map(presentationSlides(doc).flatMap(slide=>slide.objects.toArray().map(object=>[presentationObjectIdentity(object),object] as const)));for(const slide of presentationSlides(previous))for(const old of slide.objects){const value=objects.get(presentationObjectIdentity(old));if(value&&PRESENTATION_FIELDS.some(field=>old.get(field)!==value.get(field)))throw new Error('请通过专用接口修改演示对象字段');}}
   const orphan=(value:Y.Doc)=>value.share.has('objects')&&value.share.has('slideOrder')&&!value.getArray('slideOrder').toArray().includes(LEGACY_SLIDE)?value.getArray<Y.Map<unknown>>('objects'):undefined;
   if(presentationObjectLayout(presentationSlides(previous),orphan(previous))!==presentationObjectLayout(presentationSlides(doc),orphan(doc)))throw new Error('请通过专用接口新增或删除演示对象');
   if(JSON.stringify(presentationSlides(previous).map(slide=>slide.id))!==JSON.stringify(presentationSlides(doc).map(slide=>slide.id))||JSON.stringify(previous.share.has('presentationOperations')?previous.getMap('presentationOperations').toJSON():{})!==JSON.stringify(doc.share.has('presentationOperations')?doc.getMap('presentationOperations').toJSON():{}))throw new Error('请通过专用接口修改幻灯片结构');}finally{previous.destroy();}}
  let total=0;const arrays=presentationSlides(doc).map(slide=>slide.objects);if(doc.share.has('slideOrder')&&!doc.getArray('slideOrder').toArray().includes(LEGACY_SLIDE))arrays.push(doc.getArray('objects'));
  for(const objects of arrays){if(!(objects instanceof Y.Array))throw new Error('幻灯片对象结构无效');total+=objects.length;if(total>MAX_SLIDE_OBJECTS)throw new Error('演示对象超过 1 万');for(const value of objects){if(!(value instanceof Y.Map))throw new Error('演示对象结构无效');if(value._start||[...value.keys()].some(key=>!['type','text','x','y','width','height','rotation','fill','stroke','assetId'].includes(key)))throw new Error('演示对象字段无效');const type=value.get('type'),text=value.get('text');if(type==='image'&&!validOfficeAssetId(value.get('assetId')))throw new Error('演示图片引用无效');if(!['text','rect','ellipse','line','image'].includes(type as string)||['x','y','width','height'].some(key=>typeof value.get(key)!=='number'||!Number.isFinite(value.get(key))||Math.abs(Number(value.get(key)))>100)||Number(value.get('width'))<0||Number(value.get('height'))<0||type==='text'&&typeof text!=='string'||typeof text==='string'&&text.length>100000)throw new Error('演示对象字段无效');for(const key of ['fill','stroke'])if(value.has(key)&&(typeof value.get(key)!=='string'||!/^[0-9a-f]{6}$/i.test(String(value.get(key)))))throw new Error('演示颜色无效');if(value.has('rotation')&&(typeof value.get('rotation')!=='number'||!Number.isFinite(value.get('rotation'))||Math.abs(Number(value.get('rotation')))>3600))throw new Error('演示旋转无效');}}
 }finally{doc.destroy();}
}
