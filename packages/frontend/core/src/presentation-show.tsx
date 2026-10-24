import {useEffect,useRef,useState,type CSSProperties} from 'react';
import {ChevronLeft,ChevronRight,MonitorPlay,Square,X,RotateCcw} from 'lucide-react';
import type {PresentationSlide} from '../../../common/nbstore/src/office-presentation';
import {PresentationSlideView,type PresentationBuildState} from './presentation-canvas';
import {presentationBuilds,presentationObjectIdentity} from '../../../common/nbstore/src/presentation-field';
import type {LocalWorkspace} from './workspace';
const clock=(ms:number)=>{const s=Math.floor(ms/1000);return String(Math.floor(s/60)).padStart(2,'0')+':'+String(s%60).padStart(2,'0');};
/**
 * Full-screen slideshow with slide transitions, keyboard control (→/空格/PageDown 下一张, ←/PageUp 上一张,
 * Home/End, B 黑屏, W 白屏, Esc 结束) and an optional presenter view (当前页、下一页、备注、计时).
 * Going back shows the slide without its transition, as in PowerPoint/WPS.
 */
export function PresentationShow({slides,size,start,workspace,id,presenter:initialPresenter=false,onClose}:{slides:PresentationSlide[];size:{width:number;height:number};start:number;workspace:LocalWorkspace;id:string;presenter?:boolean;onClose:(slide:string)=>void}){
 const [index,setIndex]=useState(Math.max(0,Math.min(start,slides.length-1))),[leaving,setLeaving]=useState<{index:number;transition:string}|null>(null),[presenter,setPresenter]=useState(initialPresenter),[blank,setBlank]=useState<''|'black'|'white'>(''),[ended,setEnded]=useState(false),[now,setNow]=useState(Date.now()),[started,setStarted]=useState(Date.now());
 const root=useRef<HTMLDivElement>(null),at=Math.min(index,slides.length-1),current=slides[at],next=slides[at+1];
 // Animation steps of the current slide: `step` is how far it has been built; `played` names the step to animate now.
 const builds=presentationBuilds(current.objects.toArray()),steps=builds.reduce((max,build)=>Math.max(max,build.step),0),[step,setStep]=useState(0),[played,setPlayed]=useState({token:1,step:0});
 const build:PresentationBuildState={hidden:new Set(builds.filter(item=>item.step>step).map(item=>presentationObjectIdentity(item.object))),play:played.step===step?builds.filter(item=>item.step===step).map(item=>({object:presentationObjectIdentity(item.object),effect:item.effect,delay:item.delay})):[],token:played.token};
 const show=(stepValue:number,play:boolean)=>{setStep(stepValue);setPlayed(value=>({token:value.token+1,step:play?stepValue:-1}));};
 const close=useRef(onClose);close.current=onClose;const currentId=useRef(current.id);currentId.current=current.id;
 useEffect(()=>{const timer=setInterval(()=>setNow(Date.now()),1000);return()=>clearInterval(timer);},[]);
 // The outgoing layer is dropped after the transition time, so re-mounting the stage (演讲者视图) never replays it.
 useEffect(()=>{if(!leaving)return;const timer=setTimeout(()=>setLeaving(null),650);return()=>clearTimeout(timer);},[leaving]);
 // Fullscreen is best effort (it can be refused without a user gesture); leaving fullscreen with the browser's own Esc ends the show.
 useEffect(()=>{
  const element=root.current,restore=document.activeElement as HTMLElement|null;let entered=false;element?.focus();
  const change=()=>{if(entered&&!document.fullscreenElement)close.current(currentId.current);};document.addEventListener('fullscreenchange',change);
  element?.requestFullscreen?.().then(()=>{entered=true;},()=>{});
  return()=>{document.removeEventListener('fullscreenchange',change);if(document.fullscreenElement===element)void document.exitFullscreen().catch(()=>{});restore?.focus?.();};
 },[]);
 // Forward entry plays step 0 (with/after animations); going back shows the slide fully built, without animation.
 const go=(target:number,built=false)=>{
  setBlank('');if(target<0)return;if(target>=slides.length){setEnded(true);return;}
  setEnded(false);if(target===at)return;setLeaving(target>at&&slides[target].transition?{index:at,transition:slides[target].transition!}:null);setIndex(target);
  show(built?presentationBuilds(slides[target].objects.toArray()).reduce((max,item)=>Math.max(max,item.step),0):0,!built);
 };
 const forward=()=>{if(ended){onClose(current.id);return;}setBlank('');if(step<steps)show(step+1,true);else go(at+1);},back=()=>{if(ended){setEnded(false);return;}setBlank('');if(step>0)show(step-1,false);else go(at-1,true);};
 const key=(event:KeyboardEvent)=>{
  const k=event.key;let handled=true;
  // Enter/Space on a toolbar button activate that button.
  if((k===' '||k==='Enter')&&event.target instanceof HTMLElement&&event.target.closest('button'))return;
  if(['ArrowRight','ArrowDown','PageDown',' ','Enter','n','N'].includes(k))forward();else if(['ArrowLeft','ArrowUp','PageUp','Backspace','p','P'].includes(k))back();
  else if(k==='Home'){if(at)go(0);else show(0,true);}else if(k==='End')go(slides.length-1,true);else if(k==='Escape')onClose(current.id);
  else if(k==='b'||k==='B'||k==='.')setBlank(value=>value==='black'?'':'black');else if(k==='w'||k==='W'||k===',')setBlank(value=>value==='white'?'':'white');else handled=false;
  if(handled){event.preventDefault();event.stopPropagation();}
 };
 // Listen on the document: switching to the presenter view remounts the toolbar and focus may fall back to the page.
 const keyRef=useRef(key);keyRef.current=key;
 useEffect(()=>{const listener=(event:KeyboardEvent)=>{if(!event.defaultPrevented&&!event.isComposing)keyRef.current(event);};document.addEventListener('keydown',listener,true);return()=>document.removeEventListener('keydown',listener,true);},[]);
 const frame=(slide:PresentationSlide,className:string,extra?:{transition?:string;leaving?:boolean;onAnimationEnd?:()=>void},label?:string,state?:PresentationBuildState)=><div className={'presentation-show-layer '+className} data-transition={extra?.transition} data-leaving={extra?.leaving||undefined} onAnimationEnd={event=>{if(event.target===event.currentTarget)extra?.onAnimationEnd?.();}}><PresentationSlideView slide={slide} size={size} workspace={workspace} id={id} className="presentation-thumb-canvas presentation-show-slide" label={label} build={state}/></div>;
 const stage=<div className="presentation-show-frame" style={{'--slide-ratio':size.width/size.height} as CSSProperties} onClick={forward}>
  {ended?<div className="presentation-show-end" role="status">放映结束，单击或按任意翻页键退出。</div>:<>
   {leaving&&slides[leaving.index]&&frame(slides[leaving.index],'',{transition:leaving.transition,leaving:true})}
   <div key={current.id+':'+at} className="presentation-show-current">{frame(current,'',{transition:leaving?.transition,onAnimationEnd:()=>setLeaving(null)},`第 ${at+1} 张幻灯片`,build)}</div>
  </>}
  {blank&&<div className={'presentation-show-blank presentation-show-'+blank} role="status" aria-label={blank==='black'?'黑屏':'白屏'}/>}
 </div>;
 const toolbar=<div className="presentation-show-toolbar" onClick={event=>event.stopPropagation()}>
  <button type="button" onClick={back} disabled={!ended&&at===0&&step===0} aria-label="上一张" title="上一张（←）"><ChevronLeft size={16} aria-hidden="true"/></button>
  <span aria-live="polite">{ended?'已结束':`${at+1} / ${slides.length}`}{!ended&&steps>0&&<small className="presentation-show-steps">{` · 动画 ${step}/${steps}`}</small>}</span>
  <button type="button" onClick={forward} aria-label="下一张" title="下一张（→ 或空格）"><ChevronRight size={16} aria-hidden="true"/></button>
  <button type="button" aria-pressed={presenter} onClick={()=>setPresenter(value=>!value)} title="演讲者视图"><MonitorPlay size={15} aria-hidden="true"/>演讲者视图</button>
  <button type="button" aria-pressed={blank==='black'} onClick={()=>setBlank(value=>value==='black'?'':'black')} title="黑屏（B）"><Square size={15} aria-hidden="true"/>黑屏</button>
  <button type="button" onClick={()=>onClose(current.id)} title="结束放映（Esc）"><X size={15} aria-hidden="true"/>结束放映</button>
 </div>;
 return <div ref={root} className={'presentation-show'+(presenter?' presentation-show-presenter':'')} role="dialog" aria-modal="true" aria-label="幻灯片放映" tabIndex={-1}>
  {presenter?<div className="presentation-presenter">
   <section className="presentation-presenter-main" aria-label="当前幻灯片">{stage}</section>
   <aside className="presentation-presenter-side">
    <section aria-label="下一张幻灯片"><h3>下一张</h3>{next&&!ended?<div className="presentation-presenter-next" style={{'--slide-ratio':size.width/size.height} as CSSProperties}><PresentationSlideView slide={next} size={size} workspace={workspace} id={id} className="presentation-thumb-canvas presentation-show-slide"/></div>:<p>放映结束</p>}</section>
    <section className="presentation-presenter-timer" aria-label="计时"><span role="timer" aria-label="已用时间">{clock(now-started)}</span><button type="button" onClick={()=>{setStarted(Date.now());setNow(Date.now());}} title="重新计时"><RotateCcw size={14} aria-hidden="true"/>重新计时</button><span>{new Date(now).toLocaleTimeString('zh-CN',{hour:'2-digit',minute:'2-digit'})}</span></section>
    <section className="presentation-presenter-notes" aria-label="演讲备注"><h3>备注</h3><p>{current.notes||'（本页无备注）'}</p></section>
   </aside>
   {toolbar}
  </div>:<>{stage}{toolbar}</>}
 </div>;
}
