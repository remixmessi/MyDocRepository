import {splitWordParagraph,mergeWordParagraph,adjacentWordParagraph,wordTextIdentity} from './office-word-flow';
import {Bold,Italic,Underline,Strikethrough,Superscript,Subscript,Highlighter,RemoveFormatting,Paintbrush,Link,Unlink,AlignLeft,AlignCenter,AlignRight,AlignJustify} from 'lucide-react';
import {WORD_FONT_FAMILIES} from '../../../common/nbstore/src/office-word-font';
const WORD_HIGHLIGHT_LABELS:Record<string,string>={yellow:'黄色',green:'鲜绿',cyan:'青绿',magenta:'粉红',blue:'蓝色',red:'红色',darkYellow:'深黄',lightGray:'浅灰',darkGray:'深灰'};
import {useContext} from 'react';
import {createPortal} from 'react-dom';
import {WordEditingContext} from './word-editing-context';
import {WordRichInput,wordDomSelection} from './word-rich-input';
import {WORD_CELL_PARAGRAPH} from '../../../common/nbstore/src/office-word-table';
import type {WordAlignment} from '../../../common/nbstore/src/office-word-layout';
import {useEffect,useRef,useState} from 'react';
import {Y} from '@mydoc/crdt';
import {setWordPageBreak} from './office-word-page';
import {assertWordTextLive} from './office-word-location';
import {setWordFont,wordSelectionFont,setWordRightIndent,setWordFirstLineIndent,setWordAlignment,setWordLeftIndent,setWordLink,wordSelectionLink,setWordList,setWordHeading,captureWordSelection,resolveWordSelection,toggleWordMark,wordMarkActive,clearWordFormatting,setWordSpacing,wordFormatAt,type WordSelection,type WordMark,type WordFontKey} from './office-word-format';
import {applyOfficeTextEdit} from './office-text';
import {Modal} from './dialogs';

export function WordParagraphEditor({paragraph,index=0,label=`第 ${index+1} 段`,inputLabel=label,readOnly,marker,assertWritable,cell=false}:{paragraph:Y.Text;index?:number;label?:string;inputLabel?:string;marker:string;readOnly:boolean;assertWritable:()=>void;cell?:boolean}){
  const editing=useContext(WordEditingContext);
  const active=editing.active===paragraph;
  const writable=()=>{assertWritable();assertWordTextLive(paragraph);};

  const [error,setError]=useState('');
  const [isComposing,setComposing]=useState(false);
  const [selection,setSelection]=useState<WordSelection>();
  const [sizeSelection,setSizeSelection]=useState<WordSelection>(),[sizeDraft,setSizeDraft]=useState('12'),[sizeError,setSizeError]=useState('');
  const [linkSelection,setLinkSelection]=useState<WordSelection>(),[linkUrl,setLinkUrl]=useState(''),[linkError,setLinkError]=useState('');
  const [,refresh]=useState(0);
  const composition=useRef<{doc:Y.Doc;text:Y.Text;before:string}>();
  useEffect(()=>{
    const update=()=>{refresh(value=>value+1);};
    paragraph.observe(update);update();
    return ()=>{paragraph.unobserve(update);composition.current?.doc.destroy();composition.current=undefined;};
  },[paragraph]);
  const commit=(next:string)=>{
    if(readOnly)return;
    try{writable();applyOfficeTextEdit(paragraph,paragraph.toString(),next);setError('');}
    catch{setError('段落更新失败，请保留输入内容后重试。');}
  };
  const leftIndent=Number(paragraph.getAttribute('leftIndent')??0),alignment=(paragraph.getAttribute('align')??'left') as WordAlignment;
  const rightIndent=Number(paragraph.getAttribute('rightIndent')??0),firstLineIndent=Number(paragraph.getAttribute('firstLineIndent')??0);
  const indents=[0,360,720,1080,1440,1800,2160,2520,2880],spaces=[0,120,240,360,480,720];
  const range=selection&&resolveWordSelection(paragraph,selection);
  const fontSize=range?wordSelectionFont(paragraph,range.start,range.end,'fontSize'):undefined,color=range?wordSelectionFont(paragraph,range.start,range.end,'color'):undefined;
  const fontSizes=[8,9,10,10.5,11,12,14,16,18,20,24,28,32,36,48,72,96,144,200];
  const colors=[['#000000','黑色'],['#FFFFFF','白色'],['#FF0000','红色'],['#0000FF','蓝色'],['#008000','绿色'],['#FF8000','橙色'],['#808080','灰色']];
  // `selectionchange` is asynchronous, so a toolbar action right after a fast Shift+Arrow could see the previous
  // selection. Read the live DOM selection when it is still inside this paragraph; otherwise use the saved one.
  const current=()=>{if(composition.current)return selection;const input=document.querySelector<HTMLElement>(`.word-rich-input[data-word-text="${wordTextIdentity(paragraph)}"]`),live=input?wordDomSelection(input):undefined;return live?captureWordSelection(paragraph,live.start,live.end):selection;};
  const font=(key:WordFontKey,value:number|string|undefined)=>{
    if(readOnly||composition.current)return;if(editing.formatFlow?.(key,value))return;const target=current();if(!target)return;
    try{writable();setWordFont(paragraph,target,key,value);setError('');}catch(error){setError(error instanceof Error?error.message:'文字格式更新失败');}
  };
  const clear=()=>{if(readOnly||composition.current)return;if(editing.formatFlow?.('clear'))return;const target=current();if(!target)return;try{writable();clearWordFormatting(paragraph,target);setError('');}catch(error){setError(error instanceof Error?error.message:'清除格式失败');}};
  const spacing=(key:'lineSpacing'|'spaceBefore'|'spaceAfter',value:number|undefined)=>{try{writable();setWordSpacing(paragraph,key,value);setError('');}catch(error){setError(error instanceof Error?error.message:'间距更新失败');}};
  const fontFamily=range?wordSelectionFont(paragraph,range.start,range.end,'fontFamily'):undefined,highlight=range?wordSelectionFont(paragraph,range.start,range.end,'highlight'):undefined,vertAlign=range?wordSelectionFont(paragraph,range.start,range.end,'vertAlign'):undefined;
  const lineSpacing=paragraph.getAttribute('lineSpacing') as number|undefined,spaceBefore=paragraph.getAttribute('spaceBefore') as number|undefined,spaceAfter=paragraph.getAttribute('spaceAfter') as number|undefined;
  const noText=readOnly||!range&&!editing.flowActive||isComposing;
  const format=(mark:WordMark)=>{
    if(readOnly||composition.current)return;if(editing.formatFlow?.(mark))return;const target=current();if(!target)return;
    try{writable();toggleWordMark(paragraph,target,mark);setError('');}
    catch(error){setError(error instanceof Error?error.message:'格式更新失败');}
  };
  return <div className="office-word-paragraph">
    {active&&editing.toolbar&&createPortal(<div className="office-word-format" role="group" aria-label={`${label}文字样式`}>
      <div className="word-font-tools ribbon-group">
        <div className="ribbon-control-row"><select aria-label={`${label}字体`} title="字体" className="word-font-family" value={typeof fontFamily==='string'&&fontFamily!=='mixed'?fontFamily:fontFamily==='mixed'?'mixed':'inherit'} disabled={noText} onChange={event=>font('fontFamily',event.target.value==='inherit'?undefined:event.target.value)}>
          <option value="inherit">默认字体</option>{fontFamily==='mixed'&&<option value="mixed" disabled>混合</option>}{typeof fontFamily==='string'&&fontFamily!=='mixed'&&!WORD_FONT_FAMILIES.includes(fontFamily)&&<option value={fontFamily}>{fontFamily}</option>}{WORD_FONT_FAMILIES.map(value=><option key={value} value={value} style={{fontFamily:`"${value}"`}}>{value}</option>)}
        </select><select aria-label={`${label}字号`} title="字号" value={fontSize??'inherit'} disabled={readOnly||!range&&!editing.flowActive||isComposing} onChange={event=>{if(event.target.value==='custom'){const target=current();if(target){setSizeSelection(target);setSizeDraft(typeof fontSize==='number'?String(fontSize):'12');setSizeError('');}}else font('fontSize',event.target.value==='inherit'?undefined:Number(event.target.value));}}>
          <option value="inherit">默认字号</option><option value="custom">自定义…</option>{fontSize==='mixed'&&<option value="mixed" disabled>混合</option>}{typeof fontSize==='number'&&!fontSizes.includes(fontSize)&&<option value={fontSize}>{fontSize}</option>}{fontSizes.map(value=><option key={value} value={value}>{value}</option>)}
        </select><select aria-label={`${label}文字颜色`} title="文字颜色" value={color??'inherit'} disabled={readOnly||!range&&!editing.flowActive||isComposing} onChange={event=>font('color',event.target.value==='inherit'?undefined:event.target.value)}>
          <option value="inherit">默认颜色</option><option value="auto">自动</option>{color==='mixed'&&<option value="mixed" disabled>混合</option>}{typeof color==='string'&&color!=='auto'&&color!=='mixed'&&!colors.some(([value])=>value===color)&&<option value={color}>{color}</option>}{colors.map(([value,label])=><option key={value} value={value}>{label}</option>)}
        </select><input type="color" aria-label={`${label}自定义文字颜色`} title="自定义文字颜色" value={typeof color==='string'&&color.startsWith('#')?color:'#000000'} disabled={readOnly||!range&&!editing.flowActive||isComposing} onChange={event=>font('color',event.target.value)}/></div>
        <div className="ribbon-control-row">{([['bold','加粗',Bold],['italic','斜体',Italic],['underline','下划线',Underline],['strike','删除线',Strikethrough]] as const).map(([mark,label,Icon])=><button key={mark} type="button" className="ribbon-icon-button" aria-label={label} title={label} disabled={readOnly||!range&&!editing.flowActive||isComposing} aria-pressed={!!range&&wordMarkActive(paragraph,range.start,range.end,mark)} onMouseDown={event=>event.preventDefault()} onClick={()=>format(mark)}><Icon size={16}/></button>)}
          {([['superscript','上标',Superscript],['subscript','下标',Subscript]] as const).map(([value,title,Icon])=><button key={value} type="button" className="ribbon-icon-button" aria-label={title} title={title} disabled={noText} aria-pressed={vertAlign===value} onMouseDown={event=>event.preventDefault()} onClick={()=>font('vertAlign',vertAlign===value?undefined:value)}><Icon size={16}/></button>)}
          <label className="word-highlight" title="突出显示"><Highlighter size={15} aria-hidden="true"/><select aria-label={`${label}突出显示`} value={typeof highlight==='string'&&highlight!=='mixed'?highlight:'none'} disabled={noText} onChange={event=>font('highlight',event.target.value==='none'?undefined:event.target.value)}><option value="none">无突出显示</option>{Object.entries(WORD_HIGHLIGHT_LABELS).map(([value,name])=><option key={value} value={value}>{name}</option>)}</select></label>
          <button type="button" className="ribbon-icon-button" aria-label="清除格式" title="清除格式" disabled={noText} onMouseDown={event=>event.preventDefault()} onClick={clear}><RemoveFormatting size={16}/></button>
          <button type="button" className="ribbon-icon-button" aria-label="格式刷" title="格式刷：复制光标处文字格式，再选中目标文字" aria-pressed={!!editing.painter} disabled={readOnly||isComposing} onMouseDown={event=>event.preventDefault()} onClick={()=>{if(editing.painter){editing.setPainter?.(undefined);return;}const dom=window.getSelection(),element=dom?.anchorNode?(dom.anchorNode instanceof HTMLElement?dom.anchorNode:dom.anchorNode.parentElement)?.closest<HTMLElement>('.word-rich-input'):null,offset=element?wordDomSelection(element)?.start??0:range?.start??0;editing.setPainter?.(wordFormatAt(paragraph,offset));}}><Paintbrush size={16}/></button>
          <button type="button" className="ribbon-icon-button" aria-label="设置链接" title="设置链接" disabled={readOnly||!range&&!editing.flowActive||isComposing} onMouseDown={event=>event.preventDefault()} onClick={()=>{const target=current();if(!target||!range)return;setLinkSelection(target);setLinkUrl(wordSelectionLink(paragraph,range.start,range.end));setLinkError('');}}><Link size={16}/></button>
          <button type="button" className="ribbon-icon-button" aria-label="移除链接" title="移除链接" disabled={readOnly||!range&&!editing.flowActive||isComposing} onMouseDown={event=>event.preventDefault()} onClick={()=>{if(readOnly||!selection)return;try{writable();setWordLink(paragraph,selection);setError('');}catch(error){setError(error instanceof Error?error.message:'移除链接失败');}}}><Unlink size={16}/></button>
        </div><small>字体</small>
      </div>
      <div className="word-paragraph-tools ribbon-group"><div className="ribbon-control-row">
        <select aria-label={`${label}类型`} title="段落样式" value={Number(paragraph.getAttribute('heading')??0)} disabled={readOnly||isComposing} onChange={event=>{try{writable();setWordHeading(paragraph,Number(event.target.value));setError('');}catch(error){setError(error instanceof Error?error.message:'段落类型更新失败');}}}><option value="0">正文</option><option value="1">标题 1</option><option value="2">标题 2</option><option value="3">标题 3</option></select>
        <select aria-label={`${label}列表`} title="列表" value={(paragraph.getAttribute('list') as {kind:string}|undefined)?.kind??'none'} disabled={readOnly||isComposing} onChange={event=>{try{writable();setWordList(paragraph,event.target.value as 'none'|'bullet'|'decimal');setError('');}catch(error){setError(error instanceof Error?error.message:'列表更新失败');}}}><option value="none">无列表</option><option value="bullet">项目符号</option><option value="decimal">编号列表</option></select>
        <select aria-label={`${label}行距`} title="行距" value={lineSpacing??''} disabled={readOnly||isComposing} onChange={event=>spacing('lineSpacing',event.target.value?Number(event.target.value):undefined)}><option value="">默认行距</option>{lineSpacing!==undefined&&![1,1.15,1.5,2,2.5,3].includes(lineSpacing)&&<option value={lineSpacing}>{lineSpacing} 倍</option>}{[1,1.15,1.5,2,2.5,3].map(value=><option key={value} value={value}>{value} 倍行距</option>)}</select>
        </div><div className="ribbon-control-row">{([['left','左对齐',AlignLeft],['center','居中',AlignCenter],['right','右对齐',AlignRight],['justify','两端对齐',AlignJustify]]as const).map(([value,title,Icon])=><button type="button" className="ribbon-icon-button" aria-label={title} title={title} key={value} aria-pressed={alignment===value} disabled={readOnly||isComposing} onMouseDown={event=>event.preventDefault()} onClick={()=>{try{writable();setWordAlignment(paragraph,value);setError('');}catch(error){setError(error instanceof Error?error.message:'对齐更新失败');}}}><Icon size={16}/></button>)}</div><small>段落</small>
      </div>
      <div className="word-layout-tools ribbon-group"><div className="ribbon-control-row">
        {([['左缩进',leftIndent,setWordLeftIndent],['右缩进',rightIndent,setWordRightIndent]]as const).map(([title,current,apply])=><label key={title}>{title}<select aria-label={`${label}${title}`} value={current} disabled={readOnly||isComposing} onChange={event=>{try{writable();apply(paragraph,Number(event.target.value));setError('');}catch(error){setError(error instanceof Error?error.message:'缩进更新失败');}}}>{!indents.includes(current)&&<option value={current}>{current/20} 磅</option>}{indents.map(value=><option key={value} value={value}>{value/20} 磅</option>)}</select></label>)}
        <label>特殊缩进<select aria-label={`${label}特殊缩进`} value={firstLineIndent} disabled={readOnly||isComposing} onChange={event=>{try{writable();setWordFirstLineIndent(paragraph,Number(event.target.value));setError('');}catch(error){setError(error instanceof Error?error.message:'缩进更新失败');}}}>{!indents.includes(Math.abs(firstLineIndent))&&<option value={firstLineIndent}>{firstLineIndent/20} 磅</option>}<option value="0">无</option>{indents.slice(1).flatMap(value=>[<option key={value} value={value}>首行 {value/20} 磅</option>,<option key={-value} value={-value}>悬挂 {value/20} 磅</option>])}</select></label>
        {([['段前',spaceBefore,'spaceBefore'],['段后',spaceAfter,'spaceAfter']] as const).map(([title,current,key])=><label key={key}>{title}<select aria-label={`${label}${title}间距`} value={current??''} disabled={readOnly||isComposing} onChange={event=>spacing(key,event.target.value?Number(event.target.value):undefined)}><option value="">默认</option>{current!==undefined&&!spaces.includes(current)&&<option value={current}>{current/20} 磅</option>}{spaces.map(value=><option key={value} value={value}>{value/20} 磅</option>)}</select></label>)}
        {!cell&&<label><input type="checkbox" aria-label={`${label}段前分页`} checked={paragraph.getAttribute('pageBreakBefore')===true} disabled={readOnly||isComposing} onChange={event=>{try{writable();setWordPageBreak(paragraph,event.target.checked);setError('');}catch(error){setError(error instanceof Error?error.message:'分页更新失败');}}}/>本段另起一页</label>}
      </div><small>缩进与分页</small></div>

    </div>,editing.toolbar)}
    <div className="word-paragraph-content" data-heading={Number(paragraph.getAttribute('heading')??0)} data-page-break={paragraph.getAttribute('pageBreakBefore')===true||undefined} style={{textAlign:alignment,paddingLeft:leftIndent/15,paddingRight:rightIndent/15,textIndent:firstLineIndent/15,lineHeight:lineSpacing!==undefined?lineSpacing*1.2:undefined,marginTop:spaceBefore!==undefined?spaceBefore/15:undefined,marginBottom:spaceAfter!==undefined?spaceAfter/15:undefined}}>
    {marker&&<span className="office-word-list-marker" contentEditable={false}>{marker} </span>}
    <WordRichInput text={paragraph} label={inputLabel} readOnly={readOnly} paragraphText={cell?WORD_CELL_PARAGRAPH:undefined} onEnter={cell?undefined:(start,end)=>{try{writable();if(!paragraph.length&&paragraph.getAttribute('list')){setWordList(paragraph,'none');return;}editing.focus(splitWordParagraph(paragraph,start,end));}catch(error){setError(error instanceof Error?error.message:'无法换段');}}}
    onBoundary={cell?undefined:direction=>{try{writable();const merged=mergeWordParagraph(paragraph,direction);if(!merged)return false;editing.focus(merged.text,merged.offset);return true;}catch(error){setError(error instanceof Error?error.message:'无法合并正文');return true;}}}
    onNavigate={cell?undefined:direction=>{const text=adjacentWordParagraph(paragraph,direction);if(!text)return false;editing.focus(text,direction==='backward'?text.length:0);return true;}} onFocus={()=>editing.activate(paragraph)} onSelection={(start,end)=>{if(!composition.current)setSelection(captureWordSelection(paragraph,start,end));}} onFormat={format} onCommit={commit}
    onCompositionStart={()=>{
      if(readOnly)return;setComposing(true);setSelection(undefined);
      const doc=new Y.Doc();Y.applyUpdate(doc,Y.encodeStateAsUpdate(paragraph.doc!));
      const position=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromTypeIndex(paragraph,0),doc);
      if(position?.type instanceof Y.Text){composition.current?.doc.destroy();composition.current={doc,text:position.type,before:position.type.toString()};}else doc.destroy();
    }} onCompositionEnd={value=>{
      setComposing(false);const pending=composition.current;composition.current=undefined;if(!pending)return false;
      try{writable();if(!readOnly)applyOfficeTextEdit(paragraph,pending.before,value,pending.text);setError('');return true;}
      catch{setError('本次输入未提交，请复制保留草稿，再检查权限或重新打开文档。');return false;}finally{pending.doc.destroy();}
    }}/></div>
    {sizeSelection&&<Modal title="设置字号" onClose={()=>setSizeSelection(undefined)}><form onSubmit={event=>{
      event.preventDefault();if(readOnly)return;try{writable();setWordFont(paragraph,sizeSelection,'fontSize',Number(sizeDraft));setSizeSelection(undefined);setError('');}catch(error){setSizeError(error instanceof Error?error.message:'字号更新失败');}
    }}><label className="form-field">字号（磅）<input autoFocus type="number" required min="1" max="200" step="0.5" value={sizeDraft} disabled={readOnly} onChange={event=>setSizeDraft(event.target.value)}/></label>
      {sizeError&&<p role="alert" className="form-error">{sizeError}</p>}{readOnly&&<p role="status">当前文档只读。</p>}
      <div className="dialog-actions"><button type="button" onClick={()=>setSizeSelection(undefined)}>取消</button><button className="primary" disabled={readOnly}>保存字号</button></div>
    </form></Modal>}
    {linkSelection&&<Modal title="设置网页链接" onClose={()=>setLinkSelection(undefined)}><form onSubmit={event=>{
      event.preventDefault();if(readOnly)return;
      try{writable();setWordLink(paragraph,linkSelection,linkUrl.trim());setLinkSelection(undefined);setError('');}catch(error){setLinkError(error instanceof Error?error.message:'链接设置失败');}
    }}><label className="form-field">网页地址<input autoFocus type="url" required maxLength={2048} value={linkUrl} disabled={readOnly} onChange={event=>setLinkUrl(event.target.value)}/></label><p>支持 HTTP/HTTPS 网页链接，链接只会在点击时打开。</p>
      {readOnly&&<p role="alert">当前文档已变为只读，无法保存链接。</p>}{linkError&&<p role="alert" className="form-error">{linkError}</p>}
      <div className="dialog-actions"><button type="button" onClick={()=>setLinkSelection(undefined)}>取消</button><button className="primary" disabled={readOnly||!linkUrl.trim()}>保存链接</button></div>
    </form></Modal>}
    {error&&<p role="alert">{error}</p>}
</div>;
}

