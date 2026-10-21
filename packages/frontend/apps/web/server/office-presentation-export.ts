import {unpackOfficeMedia,checkOfficeMedia} from './office-media';
import * as Y from 'yjs';
import {zipSync,strToU8} from 'fflate';
import {presentationSlides,presentationSize} from '../../../../common/nbstore/src/office-presentation';
import {validatePresentationUpdate} from './office-presentation';
import {readOfficePackage,readOfficeXml} from './office-package';
import {validOfficeComment,type OfficeCommentThread} from '../../../../common/nbstore/src/office-comments';
import {presentationObjectIdentity} from '../../../../common/nbstore/src/presentation-field';
/** p:pos is in PowerPoint's comment grid (EMU / COMMENT_UNIT); measured against LibreOffice's import. */
export const COMMENT_UNIT=1587.5;
const ns='http://schemas.openxmlformats.org';
function xml(value:string){for(const ch of value){const n=ch.codePointAt(0)!;if(!(n===9||n===10||n===13||n>=32&&n<=0xd7ff||n>=0xe000&&n<=0xfffd||n>=0x10000&&n<=0x10ffff))throw new Error('演示文本包含非法 XML 字符');}return value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll('\r','&#13;');}
const rels=(items:Array<[string,string,string]>)=>`<Relationships xmlns="${ns}/package/2006/relationships">${items.map(([id,type,target])=>`<Relationship Id="${id}" Type="${ns}/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const group='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
export function exportPptxPresentation(snapshot:Uint8Array){
 const payload=unpackOfficeMedia(snapshot);snapshot=payload.state;const media=checkOfficeMedia(snapshot,'presentation',payload.media);
 validatePresentationUpdate(null,snapshot);const doc=new Y.Doc();try{
  Y.applyUpdate(doc,snapshot);if(doc.store.pendingStructs||doc.store.pendingDs)throw new Error('演示状态缺少依赖');const slides=presentationSlides(doc),size=presentationSize(doc),files:Record<string,Uint8Array>={};let total=0;
  const put=(name:string,value:string)=>{const bytes=strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+value);if(bytes.length>8*1024*1024)throw new Error('PPTX XML 超过 8 MB');total+=bytes.length;if(total>32*1024*1024)throw new Error('PPTX 内容超过 32 MB');files[name]=bytes;};
  put('_rels/.rels',rels([['rId1','officeDocument','ppt/presentation.xml']]));
  put('ppt/presentation.xml',`<p:presentation xmlns:a="${ns}/drawingml/2006/main" xmlns:r="${ns}/officeDocument/2006/relationships" xmlns:p="${ns}/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="master"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_,i)=>`<p:sldId id="${256+i}" r:id="r${i+1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${size.width}" cy="${size.height}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  const hasComments=doc.share.has('officeComments')&&[...doc.getMap<OfficeCommentThread>('officeComments').values()].some(validOfficeComment);
  put('ppt/_rels/presentation.xml.rels',rels([['master','slideMaster','slideMasters/slideMaster1.xml'],...slides.map((_,i)=>['r'+(i+1),'slide',`slides/slide${i+1}.xml`] as [string,string,string]),...(hasComments?[['authors','commentAuthors','commentAuthors.xml'] as [string,string,string]]:[])]));
  put('ppt/slideMasters/slideMaster1.xml',`<p:sldMaster xmlns:a="${ns}/drawingml/2006/main" xmlns:r="${ns}/officeDocument/2006/relationships" xmlns:p="${ns}/presentationml/2006/main"><p:cSld><p:spTree>${group}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst></p:sldMaster>`);
  put('ppt/slideMasters/_rels/slideMaster1.xml.rels',rels([['layout','slideLayout','../slideLayouts/slideLayout1.xml']]));
  put('ppt/slideLayouts/slideLayout1.xml',`<p:sldLayout xmlns:a="${ns}/drawingml/2006/main" xmlns:r="${ns}/officeDocument/2006/relationships" xmlns:p="${ns}/presentationml/2006/main" type="blank"><p:cSld><p:spTree>${group}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
  put('ppt/slideLayouts/_rels/slideLayout1.xml.rels',rels([['master','slideMaster','../slideMasters/slideMaster1.xml']]));
  const overrides:Array<[string,string]>=[['/ppt/presentation.xml','presentation.main'],['/ppt/slideMasters/slideMaster1.xml','slideMaster'],['/ppt/slideLayouts/slideLayout1.xml','slideLayout']];
  // Discussions become slide comments next to the commented object (each message one comment, replies after the first).
  // Threads whose object was deleted stay on their slide; threads without a slide go to the first slide.
  const authors:string[]=[],slideComments=new Map<number,{x:number;y:number;author:number;date:string;text:string}[]>();
  const threads=doc.share.has('officeComments')?[...doc.getMap<OfficeCommentThread>('officeComments').values()].filter(validOfficeComment).sort((a,b)=>a.messages[0].created-b.messages[0].created||a.id.localeCompare(b.id)):[];
  for(const thread of threads){
   const anchor=thread.anchor?.kind==='presentation'?thread.anchor:undefined,index=Math.max(0,anchor?slides.findIndex(slide=>slide.id===anchor.slide):0),object=anchor?slides[index]?.objects.toArray().find(value=>presentationObjectIdentity(value)===anchor.object):undefined;
   const x=object?Math.round((Number(object.get('x'))+Number(object.get('width')))*size.width/COMMENT_UNIT):0,y=object?Math.round(Number(object.get('y'))*size.height/COMMENT_UNIT):0,list=slideComments.get(index)??[];
   thread.messages.forEach((message,i)=>{let author=authors.indexOf(message.authorName);if(author<0)author=authors.push(message.authorName)-1;list.push({x,y,author,date:new Date(message.created).toISOString().replace(/\.\d{3}Z$/,''),text:(i===0&&thread.resolved?'[已解决] ':'')+(i?'回复：':'')+message.body});});
   slideComments.set(index,list);
  }
  if(authors.length){
   const counts=authors.map((_,author)=>[...slideComments.values()].flat().filter(comment=>comment.author===author).length);
   put('ppt/commentAuthors.xml',`<p:cmAuthorLst xmlns:p="${ns}/presentationml/2006/main">${authors.map((name,i)=>`<p:cmAuthor id="${i}" name="${xml(name)}" initials="${xml(([...name.trim()][0]??'?'))}" lastIdx="${counts[i]}" clrIdx="${i%8}"/>`).join('')}</p:cmAuthorLst>`);
   overrides.push(['/ppt/commentAuthors.xml','commentAuthors']);
  }
  const paragraphs=(value:string)=>value.split('\n').map(line=>`<a:p><a:r><a:rPr lang="zh-CN" sz="1800"/><a:t xml:space="preserve">${xml(line)}</a:t></a:r></a:p>`).join('');
  for(const [i,slide]of slides.entries()){
   const imageRels:Array<[string,string,string]>=[];
   const shapes=slide.objects.toArray().map((object,index)=>{
    const type=object.get('type');if(!['text','rect','ellipse','line','image'].includes(String(type)))throw new Error('PPTX 副本暂不支持图片或未知对象，请保留原件或办公文件备份');
    const number=(key:string,scale:number)=>Math.round(Number(object.get(key))*scale),rotation=Math.round(Number(object.get('rotation')??0)*60000),fill=xml(String(object.get('fill')??'FFFFFF')),stroke=xml(String(object.get('stroke')??'333333'));
    if(type==='image'){const asset=media.get(String(object.get('assetId')))!;const id='image'+index,name=`${asset.entry.id}.${asset.entry.mime==='image/png'?'png':'jpg'}`;files['ppt/media/'+name]=asset.bytes;imageRels.push([id,'image','../media/'+name]);return `<p:pic><p:nvPicPr><p:cNvPr id="${index+2}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm rot="${rotation}"><a:off x="${number('x',size.width)}" y="${number('y',size.height)}"/><a:ext cx="${number('width',size.width)}" cy="${number('height',size.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;}
    const geometry=type==='ellipse'?'ellipse':type==='line'?'line':'rect';
    return `<p:sp><p:nvSpPr><p:cNvPr id="${index+2}" name="对象 ${index+1}"/><p:cNvSpPr${type==='text'?' txBox="1"':''}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="${rotation}"><a:off x="${number('x',size.width)}" y="${number('y',size.height)}"/><a:ext cx="${number('width',size.width)}" cy="${number('height',size.height)}"/></a:xfrm><a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>${type==='text'?'<a:noFill/>':`<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`}<a:ln><a:solidFill><a:srgbClr val="${stroke}"/></a:solidFill></a:ln></p:spPr>${type==='text'?`<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs(String(object.get('text')??''))}</p:txBody>`:''}</p:sp>`;
   }).join('');
   put(`ppt/slides/slide${i+1}.xml`,`<p:sld xmlns:a="${ns}/drawingml/2006/main" xmlns:r="${ns}/officeDocument/2006/relationships" xmlns:p="${ns}/presentationml/2006/main"><p:cSld><p:spTree>${group}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
   const comments=slideComments.get(i)??[],seen=new Map<number,number>();
   if(comments.length){put(`ppt/comments/comment${i+1}.xml`,`<p:cmLst xmlns:p="${ns}/presentationml/2006/main">${comments.map(comment=>{const idx=(seen.get(comment.author)??0)+1;seen.set(comment.author,idx);return `<p:cm authorId="${comment.author}" dt="${comment.date}" idx="${idx}"><p:pos x="${comment.x}" y="${comment.y}"/><p:text>${xml(comment.text)}</p:text></p:cm>`;}).join('')}</p:cmLst>`);overrides.push([`/ppt/comments/comment${i+1}.xml`,'comments']);}
   put(`ppt/slides/_rels/slide${i+1}.xml.rels`,rels([...imageRels,['layout','slideLayout','../slideLayouts/slideLayout1.xml'],['notes','notesSlide',`../notesSlides/notesSlide${i+1}.xml`],...(comments.length?[['comments','comments',`../comments/comment${i+1}.xml`] as [string,string,string]]:[])]));
   put(`ppt/notesSlides/notesSlide${i+1}.xml`,`<p:notes xmlns:a="${ns}/drawingml/2006/main" xmlns:p="${ns}/presentationml/2006/main"><p:cSld><p:spTree>${group}<p:sp><p:nvSpPr><p:cNvPr id="2" name="备注"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs(slide.notes)}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`);
   put(`ppt/notesSlides/_rels/notesSlide${i+1}.xml.rels`,rels([['slide','slide',`../slides/slide${i+1}.xml`]]));
   overrides.push([`/ppt/slides/slide${i+1}.xml`,'slide'],[`/ppt/notesSlides/notesSlide${i+1}.xml`,'notesSlide']);
  }
  put('[Content_Types].xml',`<Types xmlns="${ns}/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/>${overrides.map(([path,type])=>`<Override PartName="${path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.${type}+xml"/>`).join('')}</Types>`);
  const bytes=zipSync(files);for(const [name,value]of Object.entries(readOfficePackage(bytes)))if(/\.(xml|rels)$/.test(name))readOfficeXml(value);return bytes;
 }finally{doc.destroy();}
}
