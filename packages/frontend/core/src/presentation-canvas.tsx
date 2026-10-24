import {useLayoutEffect,useRef,type CSSProperties} from 'react';
import type * as Y from 'yjs';
import {presentationObjectIdentity,presentationPaintOrder,DEFAULT_PRESENTATION_FONT_SIZE,PRESENTATION_ANIMATION_MS,type PresentationAnimation} from '../../../common/nbstore/src/presentation-field';
/** Web Animations keyframes of each entrance effect (shared by the slideshow and the editor's 预览). */
export function presentationEffectKeyframes(effect:PresentationAnimation,element:HTMLElement):Keyframe[]{
 const slide=element.parentElement?.getBoundingClientRect(),box=element.getBoundingClientRect(),below=slide?slide.bottom-box.top:400,left=slide?box.right-slide.left:400;
 return effect==='appear'?[{opacity:0},{opacity:0,offset:.999},{opacity:1}]:effect==='fade'?[{opacity:0},{opacity:1}]:effect==='flyBottom'?[{translate:`0 ${below}px`},{translate:'0 0'}]:effect==='flyLeft'?[{translate:`${-left}px 0`},{translate:'0 0'}]:effect==='wipe'?[{clipPath:'inset(0 100% 0 0)'},{clipPath:'inset(0 0 0 0)'}]:[{scale:'0',opacity:0},{scale:'1',opacity:1}];
}
/** With “减少动态效果” the effect is instant, but delayed items still wait their turn so the build order stays visible. */
export const playPresentationEffect=(element:HTMLElement,effect:PresentationAnimation,delay=0)=>{const reduced=window.matchMedia?.('(prefers-reduced-motion: reduce)').matches;return element.animate(presentationEffectKeyframes(reduced?'appear':effect,element),{duration:effect==='appear'||reduced?1:PRESENTATION_ANIMATION_MS,delay,easing:'ease-out',fill:'backwards'});};
import type {PresentationSlide} from '../../../common/nbstore/src/office-presentation';
import {OfficeImage} from './office-image-ui';
import type {LocalWorkspace} from './workspace';
/** Polygon outlines (0..100 box) for preset shapes that CSS borders cannot draw. */
const POLYGONS:Record<string,string>={triangle:'50,0 100,100 0,100',diamond:'50,0 100,50 50,100 0,50',rightArrow:'0,25 60,25 60,0 100,50 60,100 60,75 0,75',star5:'50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35'};
export const isPolygonShape=(type:string)=>Object.hasOwn(POLYGONS,type);
/** Box geometry shared by the editor, thumbnails and the slideshow. Polygon shapes paint their own fill/outline. */
export function presentationObjectStyle(object:Y.Map<unknown>,preview?:{x:number;y:number;width:number;height:number}):CSSProperties{
 const type=String(object.get('type')),box=preview??{x:Number(object.get('x')),y:Number(object.get('y')),width:Number(object.get('width')),height:Number(object.get('height'))},plain=type==='text'||type==='line'||type==='image'||isPolygonShape(type);
 return {left:`${box.x*100}%`,top:`${box.y*100}%`,width:`${box.width*100}%`,height:`${box.height*100}%`,transform:`rotate(${Number(object.get('rotation')??0)}deg)`,background:plain?'transparent':'#'+String(object.get('fill')??'FFFFFF'),borderColor:isPolygonShape(type)?'transparent':'#'+String(object.get('stroke')??'333333'),borderRadius:type==='ellipse'?'50%':type==='roundRect'?'12%':undefined};
}
/**
 * Text box style. Font sizes are points on the slide; `cqw` scales them with the canvas width
 * (slide width in points = EMU / 12700), so thumbnails, editor and slideshow keep the same layout.
 */
export function presentationTextStyle(object:Y.Map<unknown>,slideWidth:number):CSSProperties{
 const size=Number(object.get('fontSize')??DEFAULT_PRESENTATION_FONT_SIZE),color=object.get('color');
 return {fontSize:`${size*100/(slideWidth/12700)}cqw`,lineHeight:1.2,fontWeight:object.get('bold')?700:undefined,fontStyle:object.get('italic')?'italic':undefined,textDecoration:object.get('underline')?'underline':undefined,color:color?'#'+String(color):undefined,textAlign:(object.get('align') as CSSProperties['textAlign'])??undefined};
}
export function PresentationShapeGlyph({object}:{object:Y.Map<unknown>}){
 const type=String(object.get('type')),stroke='#'+String(object.get('stroke')??'333333');
 if(type==='line')return <svg width="100%" height="100%" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><line x1="0" y1="0" x2="100" y2="100" stroke={stroke} vectorEffect="non-scaling-stroke"/></svg>;
 if(isPolygonShape(type))return <svg width="100%" height="100%" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><polygon points={POLYGONS[type]} fill={'#'+String(object.get('fill')??'FFFFFF')} stroke={stroke} vectorEffect="non-scaling-stroke"/></svg>;
 return null;
}
/** Read-only slide rendering (thumbnails, slideshow, presenter view). */
/** `build` (slideshow): objects not yet shown are hidden; `play` objects run their entrance effect when `token` changes. */
export type PresentationBuildState={hidden:Set<string>;play:{object:string;effect:PresentationAnimation;delay:number}[];token:number};
export function PresentationSlideView({slide,size,workspace,id,className,label,build}:{slide:PresentationSlide;size:{width:number;height:number};workspace:LocalWorkspace;id:string;className:string;label?:string;build?:PresentationBuildState}){
 const root=useRef<HTMLDivElement>(null);
 useLayoutEffect(()=>{const element=root.current;if(!element||!build?.play.length)return;const running=build.play.flatMap(item=>{const target=element.querySelector<HTMLElement>(`[data-object="${item.object}"]`);return target?[playPresentationEffect(target,item.effect,item.delay)]:[];});return()=>{for(const animation of running)animation.finish();};},[build?.token]);
 return <div ref={root} className={className} style={{aspectRatio:`${size.width} / ${size.height}`,background:'#'+(slide.background??'FFFFFF')}} aria-label={label} role={label?'img':undefined}>{presentationPaintOrder(slide.objects.toArray()).map(({object})=>{
  const type=String(object.get('type'));
  const identity=presentationObjectIdentity(object);
  return <article key={identity} data-object={identity} className={'presentation-thumb-object presentation-'+type} style={{...presentationObjectStyle(object),...(build?.hidden.has(identity)?{visibility:'hidden' as const}:{})}}>{type==='text'&&<div className="presentation-text" style={presentationTextStyle(object,size.width)}>{String(object.get('text')??'')}</div>}<PresentationShapeGlyph object={object}/>{type==='image'&&<OfficeImage workspace={workspace} id={id} assetId={String(object.get('assetId'))}/>}</article>;
 })}</div>;
}
