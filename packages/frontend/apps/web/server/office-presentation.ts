import {presentationObjectIdentity,presentationObjectLayout,validPresentationField,PRESENTATION_FIELDS,PRESENTATION_OBJECT_TYPES,PRESENTATION_TEXT_FIELDS} from '../../../../common/nbstore/src/presentation-field';
import {validateOfficeComments} from '../../../../common/nbstore/src/office-comments';
import {packageImageReader,type OfficeMedia,validateOfficeMedia} from './office-media';
import {validOfficeAssetId} from '../../../../common/nbstore/src/office-image';
import * as Y from 'yjs';
import {posix} from 'node:path';
import {readOfficePackage} from './office-package';
import {xlsxXml,attribute,child,children,type XlsxXml} from './office-xlsx-xml';
import {presentationSlides,MAX_PRESENTATION_SLIDES,LEGACY_SLIDE,presentationSize,createPresentationObject,validPresentationBackground,validPresentationTransition,PRESENTATION_TRANSITIONS,type PresentationTransition} from '../../../../common/nbstore/src/office-presentation';
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
const SHAPE_PRESETS=['rect','roundRect','ellipse','line','triangle','diamond','rightArrow','star5'] as const;
/** Whole-box text style from the first run and paragraph (the model keeps one style per text box). */
function textStyle(shape:XlsxXml){
 const style:Record<string,unknown>={},paragraph=descendants(shape,'p')[0],run=paragraph?descendants(paragraph,'r')[0]:undefined,properties=run?child(run,'rPr'):undefined,align=paragraph&&child(paragraph,'pPr')?attribute(child(paragraph,'pPr')!,'algn'):undefined;
 if(properties){const size=Number(attribute(properties,'sz'))/100;if(Number.isFinite(size)&&size>=6&&size<=400)style.fontSize=Math.round(size*10)/10;if(attribute(properties,'b')==='1')style.bold=true;if(attribute(properties,'i')==='1')style.italic=true;if(attribute(properties,'u')&&attribute(properties,'u')!=='none')style.underline=true;const fill=child(properties,'solidFill'),rgb=fill?child(fill,'srgbClr'):undefined,value=rgb?attribute(rgb,'val'):undefined;if(value&&/^[0-9A-Fa-f]{6}$/.test(value))style.color=value.toUpperCase();}
 if(align==='ctr')style.align='center';else if(align==='r')style.align='right';return style;
}
/** Solid RGB background of a slide, layout or master; `null` means the level states a background the model cannot show. */
function slideBackground(root:XlsxXml):string|null|undefined{const bg=descendants(root,'bg')[0];if(!bg)return undefined;const fill=descendants(bg,'solidFill')[0],rgb=fill?child(fill,'srgbClr'):undefined,value=rgb?attribute(rgb,'val'):undefined;return value&&/^[0-9A-Fa-f]{6}$/.test(value)?value.toUpperCase():null;}
/** Maps p:transition (including the mc:Fallback of newer effects) to the supported effects; others become fade. */
function slideTransition(root:XlsxXml):PresentationTransition|undefined{
 const nodes=descendants(root,'transition');if(!nodes.length)return undefined;const names=nodes.flatMap(node=>node.children.map(value=>value.name));
 const known=names.find(name=>PRESENTATION_TRANSITIONS.includes(name as never)) as PresentationTransition|undefined;if(known)return known;return names.some(name=>name&&!['sndAc','extLst','cut'].includes(name))?'fade':undefined;
}
/** Entrance effects of the main sequence (other classes, paths and text builds are left out); unknown presets become fade. */
function readTiming(root:XlsxXml,shapes:Map<string,Y.Map<unknown>>){
 const timing=descendants(root,'timing')[0];if(!timing)return;let order=0;
 for(const node of descendants(timing,'cTn')){
  if(attribute(node,'presetClass')!=='entr')continue;const target=descendants(node,'spTgt')[0],object=target&&shapes.get(attribute(target,'spid')??'');if(!object||object.has('animation'))continue;
  const preset=Number(attribute(node,'presetID')),subtype=Number(attribute(node,'presetSubtype')??0),type=attribute(node,'nodeType');
  // 飞入 from the left/top-left/bottom-left becomes 自左侧, every other direction 自底部.
  object.set('animation',preset===1?'appear':preset===2?([8,9,12].includes(subtype)?'flyLeft':'flyBottom'):preset===22?'wipe':[53,23,31].includes(preset)?'zoom':'fade');object.set('animationTrigger',type==='withEffect'?'with':type==='afterEffect'?'after':'click');object.set('animationOrder',++order);
 }
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
   {const settings=index===0?doc.getMap('presentation'):doc.getMap<Y.Map<unknown>>('slides').get(id)!,transition=slideTransition(root);if(transition)settings.set('transition',transition);const layoutFile=relationTargets(files,path,'slideLayout')[0],masterFile=layoutFile?relationTargets(files,layoutFile,'slideMaster')[0]:undefined,background=[root,...[layoutFile,masterFile].flatMap(file=>file&&files[file]?[xlsxXml(files[file])]:[])].map(slideBackground).find(value=>value!==undefined);if(background)settings.set('background',background);}
   const tree=descendants(root,'spTree')[0]??root;if(descendants(tree,'grpSp').some(group=>descendants(group,'pic').length))throw new Error('组合内图片尚未支持，请先取消组合');
   const layoutPath=relationTargets(files,path,'slideLayout')[0],masterPath=layoutPath?relationTargets(files,layoutPath,'slideMaster')[0]:undefined,layout=placeholders(files,layoutPath),master=placeholders(files,masterPath);
   const shapeIds=new Map<string,Y.Map<unknown>>();
   for(const shape of tree.children.filter(node=>['sp','pic'].includes(node.name))){
    const properties=child(shape,'spPr'),geometry=properties?child(properties,'prstGeom'):undefined,preset=geometry?attribute(geometry,'prst'):undefined;
    const paragraphs=descendants(shape,'p'),text=paragraphs.length?paragraphs.map(p=>descendants(p,'t').map(t=>t.text).join('')).join('\n'):descendants(shape,'t').map(t=>t.text).join('');
    if(!text&&preset&&!SHAPE_PRESETS.includes(preset as never))continue;
    const image=shape.name==='pic',box=attribute(child(child(shape,'nvSpPr')??shape,'cNvSpPr')??shape,'txBox')==='1'||!!placeholderOf(shape);
    // An empty drawn shape keeps its geometry even though PowerPoint gives it an empty text body.
    const type=image?'image':!text&&!box&&preset&&SHAPE_PRESETS.includes(preset as never)?preset:text||child(shape,'txBody')?'text':preset==='ellipse'?'ellipse':preset==='line'?'line':'rect';
    const object=createPresentationObject(type);object.set('text',text);if(type==='text')for(const [key,value]of Object.entries(textStyle(shape)))object.set(key,value);if(image){const blip=descendants(shape,'blip')[0],embed=blip&&Object.entries(blip.attributes).find(([key])=>key.endsWith(':embed'))?.[1];if(!embed)throw new Error('PPTX 图片缺少内嵌关系');object.set('assetId',images.read(embed));}
    const sources=[shape,...inheritedShapes(placeholderOf(shape),layout,master)],spPr=(node:XlsxXml)=>child(node,'spPr');
    const transform=sources.map(node=>{const own=spPr(node);return own?child(own,'xfrm'):undefined;}).find(node=>node&&child(node,'off')&&child(node,'ext')),off=transform?child(transform,'off'):undefined,ext=transform?child(transform,'ext'):undefined;
    if(off){object.set('x',Number(attribute(off,'x'))/width);object.set('y',Number(attribute(off,'y'))/height);}if(ext){object.set('width',Number(attribute(ext,'cx'))/width);object.set('height',Number(attribute(ext,'cy'))/height);}if(transform&&attribute(transform,'rot'))object.set('rotation',Number(attribute(transform,'rot'))/60000);
    // The nearest level that states a fill (solid or none) decides it; only explicit RGB is carried over.
    for(const field of ['fill','stroke']){for(const source of sources){const own=spPr(source),node=field==='fill'?own:own?child(own,'ln'):undefined;if(!node||!child(node,'solidFill')&&!child(node,'noFill'))continue;const rgb=child(node,'solidFill')?child(child(node,'solidFill')!,'srgbClr'):undefined;if(rgb)object.set(field,attribute(rgb,'val'));break;}}
    {const cNvPr=descendants(child(shape,'nvSpPr')??child(shape,'nvPicPr')??shape,'cNvPr')[0],spid=cNvPr&&attribute(cNvPr,'id');if(spid)shapeIds.set(spid,object);}
    objects.push([object]);if(objects.length>MAX_SLIDE_OBJECTS)throw new Error('幻灯片对象超过 1 万');
   }
   readTiming(root,shapeIds);
   mediaBytes+=images.media.reduce((total,value)=>total+Buffer.from(value.data,'base64').length,0);if(mediaBytes>8*1024*1024||media.length+images.media.length>1000)throw new Error('办公图片总量或数量超限');media.push(...images.media);
   const relPath=posix.join(posix.dirname(path),'_rels',posix.basename(path)+'.rels');if(files[relPath]){const notes=children(xlsxXml(files[relPath]),'Relationship').filter(rel=>attribute(rel,'Type')?.endsWith('/notesSlide'));if(notes.length>1)throw new Error('PPTX 备注关系重复');if(notes[0]){if(attribute(notes[0],'TargetMode')==='External')throw new Error('PPTX 不支持外部备注');const notesPath=target(posix.dirname(path),attribute(notes[0],'Target'));if(!files[notesPath])throw new Error('PPTX 备注缺失');const notesRoot=xlsxXml(files[notesPath]);if(notesRoot.name!=='notes')throw new Error('PPTX 备注结构无效');const note=descendants(notesRoot,'sp').filter(sp=>!descendants(sp,'ph').some(ph=>['sldNum','sldImg'].includes(attribute(ph,'type')??''))).flatMap(sp=>descendants(sp,'p').map(p=>descendants(p,'t').map(t=>t.text).join(''))).join('\n');if(index===0)doc.getMap('presentation').set('notes',note);else doc.getMap<Y.Map<unknown>>('slides').get(id)!.set('notes',note);}}
  }
  validateOfficeMedia(media);const update=Y.encodeStateAsUpdate(doc);validatePresentationUpdate(null,update);return {doc,update,media};
 }catch(error){doc.destroy();throw error;}
}
const OBJECT_KEYS=['type','text','x','y','width','height','rotation','fill','stroke','assetId','fontSize','bold','italic','underline','color','align','z','animation','animationTrigger','animationOrder'];
function validatePresentationObject(value:unknown){
 if(!(value instanceof Y.Map))throw new Error('演示对象结构无效');if(value._start||[...value.keys()].some(key=>!OBJECT_KEYS.includes(key)))throw new Error('演示对象字段无效');
 const type=value.get('type'),text=value.get('text');if(type==='image'&&!validOfficeAssetId(value.get('assetId')))throw new Error('演示图片引用无效');
 if(!PRESENTATION_OBJECT_TYPES.includes(type as never)||['x','y','width','height'].some(key=>typeof value.get(key)!=='number'||!Number.isFinite(value.get(key))||Math.abs(Number(value.get(key)))>100)||Number(value.get('width'))<0||Number(value.get('height'))<0||type==='text'&&typeof text!=='string'||typeof text==='string'&&text.length>100000)throw new Error('演示对象字段无效');
 for(const key of ['fill','stroke','color'])if(value.has(key)&&!validPresentationField(key,value.get(key)))throw new Error('演示颜色无效');
 if(value.has('rotation')&&!validPresentationField('rotation',value.get('rotation')))throw new Error('演示旋转无效');
 // Text formatting belongs to text boxes only; other shapes carry no text.
 for(const key of PRESENTATION_TEXT_FIELDS)if(value.has(key)&&(type!=='text'||!validPresentationField(key,value.get(key))))throw new Error('演示文字格式无效');
 if(value.has('animationOrder')&&!validPresentationField('animationOrder',value.get('animationOrder'))||value.has('animationTrigger')&&!validPresentationField('animationTrigger',value.get('animationTrigger'))||value.has('animation')&&!validPresentationField('animation',value.get('animation')))throw new Error('演示动画设置无效');
 if(value.has('z')&&(!Number.isSafeInteger(value.get('z'))||Number(value.get('z'))<0||Number(value.get('z'))>=MAX_SLIDE_OBJECTS))throw new Error('演示对象层次无效');
}
export function validatePresentationUpdate(current:Uint8Array|null,update:Uint8Array){
 const doc=new Y.Doc();try{
  if(current)Y.applyUpdate(doc,current);Y.applyUpdate(doc,update);validateOfficeComments(doc);if([...doc.share.keys()].some(key=>!['officeComments','officeCommentOperations','objects','slideOrder','slides','presentation','presentationOperations','presentationFields'].includes(key)))throw new Error('演示文稿更新包含非法根结构');
  if(doc.share.has('presentation')){const meta=doc.getMap('presentation');if(['width','height'].some(key=>meta.has(key)&&typeof meta.get(key)!=='number')||[...meta.keys()].some(key=>!['width','height','notes','background','transition'].includes(key))||meta.has('background')&&!validPresentationBackground(meta.get('background'))||meta.has('transition')&&!validPresentationTransition(meta.get('transition'))||meta.has('notes')&&(typeof meta.get('notes')!=='string'||String(meta.get('notes')).length>100000))throw new Error('演示设置无效');}
  const size=presentationSize(doc);if(!Object.values(size).every(n=>Number.isSafeInteger(n)&&n>=1&&n<=100000000))throw new Error('演示尺寸无效');
  if(doc.share.has('slideOrder')){const order=doc.getArray('slideOrder').toArray();if(!order.length||order.length>MAX_PRESENTATION_SLIDES||order.some(id=>typeof id!=='string'||!/^[a-zA-Z0-9-]{1,80}$/.test(id))||new Set(order).size!==order.length)throw new Error('幻灯片顺序无效');const slides=doc.getMap('slides');if([...slides.keys()].some(id=>id===LEGACY_SLIDE||!order.includes(id)))throw new Error('幻灯片引用无效');for(const id of order){if(id===LEGACY_SLIDE)continue;const slide=slides.get(id as string);if(!(slide instanceof Y.Map)||[...slide.keys()].some(key=>!['objects','notes','background','transition'].includes(key))||slide.has('background')&&!validPresentationBackground(slide.get('background'))||slide.has('transition')&&!validPresentationTransition(slide.get('transition'))||!(slide.get('objects') instanceof Y.Array)||typeof slide.get('notes')!=='string'||slide.get('notes').length>100000)throw new Error('幻灯片结构无效');}}
  else if(doc.share.has('slides')&&doc.getMap('slides').size)throw new Error('幻灯片缺少顺序');
  if(doc.share.has('presentationOperations')){const receipts=doc.getMap('presentationOperations');if(receipts.size>10000||[...receipts.values()].some(value=>typeof value!=='string'||! /^[a-f0-9]{64}:[a-zA-Z0-9-]{1,80}$/.test(value)))throw new Error('演示操作回执无效');}
  if(doc.share.has('presentationFields')){const fields=doc.getMap('presentationFields');if(fields.size>80000||[...fields].some(([key,value])=>!/^\d+-\d+:[a-zA-Z]{1,16}$/.test(key)||!PRESENTATION_FIELDS.includes(key.split(':')[1] as never)||typeof value!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(value)))throw new Error('演示字段修订无效');}
  if(current){const previous=new Y.Doc();try{Y.applyUpdate(previous,current);
   if(JSON.stringify(previous.share.has('presentationFields')?previous.getMap('presentationFields').toJSON():{})!==JSON.stringify(doc.share.has('presentationFields')?doc.getMap('presentationFields').toJSON():{}))throw new Error('请通过专用接口修改演示字段修订');
   if(previous.share.has('presentationFields')){const objects=new Map(presentationSlides(doc).flatMap(slide=>slide.objects.toArray().map(object=>[presentationObjectIdentity(object),object] as const)));for(const slide of presentationSlides(previous))for(const old of slide.objects){const value=objects.get(presentationObjectIdentity(old));if(value&&PRESENTATION_FIELDS.some(field=>old.get(field)!==value.get(field)))throw new Error('请通过专用接口修改演示对象字段');}}
   const orphan=(value:Y.Doc)=>value.share.has('objects')&&value.share.has('slideOrder')&&!value.getArray('slideOrder').toArray().includes(LEGACY_SLIDE)?value.getArray<Y.Map<unknown>>('objects'):undefined;
   if(presentationObjectLayout(presentationSlides(previous),orphan(previous))!==presentationObjectLayout(presentationSlides(doc),orphan(doc)))throw new Error('请通过专用接口新增或删除演示对象');
   if(JSON.stringify(presentationSlides(previous).map(slide=>slide.id))!==JSON.stringify(presentationSlides(doc).map(slide=>slide.id))||JSON.stringify(previous.share.has('presentationOperations')?previous.getMap('presentationOperations').toJSON():{})!==JSON.stringify(doc.share.has('presentationOperations')?doc.getMap('presentationOperations').toJSON():{}))throw new Error('请通过专用接口修改幻灯片结构');}finally{previous.destroy();}}
  let total=0;const arrays=presentationSlides(doc).map(slide=>slide.objects);if(doc.share.has('slideOrder')&&!doc.getArray('slideOrder').toArray().includes(LEGACY_SLIDE))arrays.push(doc.getArray('objects'));
  for(const objects of arrays){if(!(objects instanceof Y.Array))throw new Error('幻灯片对象结构无效');total+=objects.length;if(total>MAX_SLIDE_OBJECTS)throw new Error('演示对象超过 1 万');for(const value of objects)validatePresentationObject(value);}
 }finally{doc.destroy();}
}
