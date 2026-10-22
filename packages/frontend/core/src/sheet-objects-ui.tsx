import {useEffect,useLayoutEffect,useRef,useState,type RefObject} from 'react';
import {BarChart3,BarChartHorizontal,LineChart,PieChart,AreaChart,Palette,Trash2,ArrowUp,ArrowDown,Pencil} from 'lucide-react';
import type {Y} from '@mydoc/crdt';
import {Modal} from './dialogs';
import {OfficeMenu,RibbonButton,RibbonGroup} from './office-ribbon';
import {SheetChartView,CHART_TYPE_LABELS} from './sheet-chart';
import {cellName,type SheetRange} from './spreadsheet-commands';
import {chartData,resolveSheetRange,sheetObjectRevision,sheetObjects,sheetRangeKeys,SHEET_CHART_TYPES,type ConditionalRule,type ConditionalStyle,type SheetChart,type SheetChartType,type SheetConditionalFormat,type SheetObject} from '../../../common/nbstore/src/spreadsheet-objects';
import type {SheetOperation} from '../../../common/nbstore/src/spreadsheet-operation';
import {sheetRowIdentity} from '../../../common/nbstore/src/sheet-columns';

type Run=(operation:SheetOperation,done?:string)=>Promise<boolean>;
type Sheet=Y.Array<Y.Array<string>>;
/** Excel's preset highlight styles. */
export const CONDITIONAL_STYLES:{label:string;style:ConditionalStyle}[]=[
 {label:'浅红填充色深红色文本',style:{fill:'FFC7CE',color:'9C0006'}},{label:'黄填充色深黄色文本',style:{fill:'FFEB9C',color:'9C5700'}},
 {label:'绿填充色深绿色文本',style:{fill:'C6EFCE',color:'006100'}},{label:'浅红色填充',style:{fill:'FFC7CE'}},{label:'红色文本',style:{color:'9C0006'}},{label:'红色加粗文本',style:{color:'C00000',bold:true}},
];
const SCALES:{label:string;colors:[string,string]|[string,string,string]}[]=[{label:'红-黄-绿',colors:['F8696B','FFEB84','63BE7B']},{label:'绿-黄-红',colors:['63BE7B','FFEB84','F8696B']},{label:'白-红',colors:['FFFFFF','F8696B']},{label:'白-绿',colors:['FFFFFF','63BE7B']},{label:'蓝-白-红',colors:['5A8AC6','FCFCFF','F8696B']}];
const BARS:{label:string;color:string}[]=[{label:'蓝色数据条',color:'638EC6'},{label:'绿色数据条',color:'63C384'},{label:'红色数据条',color:'FF555A'},{label:'橙色数据条',color:'FFB628'}];
const RULE_LABELS:Record<string,string>={greaterThan:'大于',lessThan:'小于',greaterOrEqual:'大于或等于',lessOrEqual:'小于或等于',equal:'等于',notEqual:'不等于',between:'介于',notBetween:'未介于',textContains:'文本包含',textNotContains:'文本不包含',beginsWith:'开头是',endsWith:'结尾是',duplicate:'重复值',unique:'唯一值',blank:'空值',notBlank:'非空值',top:'前',bottom:'后',aboveAverage:'高于平均值',belowAverage:'低于平均值',colorScale:'色阶',dataBar:'数据条'};
export function describeRule(rule:ConditionalRule){
 if(rule.type==='top'||rule.type==='bottom')return `${RULE_LABELS[rule.type]} ${rule.count}${rule.percent?'%':' 项'}`;
 if(rule.type==='between'||rule.type==='notBetween')return `${RULE_LABELS[rule.type]} ${rule.value} 和 ${rule.value2}`;
 return 'value' in rule?`${RULE_LABELS[rule.type]} ${rule.value}`:RULE_LABELS[rule.type];
}
const rangeLabel=(range:SheetRange|undefined)=>range?cellName(range.top,range.left)+(range.bottom>range.top||range.right>range.left?':'+cellName(range.bottom,range.right):''):'（区域已删除）';
const objectId=(prefix:string)=>prefix+'-'+crypto.randomUUID();
function Swatch({rule}:{rule:ConditionalRule}){
 if(rule.type==='colorScale')return <span className="cf-swatch" style={{background:`linear-gradient(90deg,${rule.colors.map(c=>'#'+c).join(',')})`}} aria-hidden="true"/>;
 if(rule.type==='dataBar')return <span className="cf-swatch" style={{background:`linear-gradient(90deg,#${rule.color} 60%,transparent 60%)`}} aria-hidden="true"/>;
 return <span className="cf-swatch" style={{background:rule.style.fill?'#'+rule.style.fill:undefined,color:rule.style.color?'#'+rule.style.color:undefined,fontWeight:rule.style.bold?700:undefined}} aria-hidden="true">Aa</span>;
}

// ---- conditional formatting ----------------------------------------------------------------------------------------
/** Rule types created through the dialog (the ones with a highlight style). */
type DialogRule=Extract<ConditionalRule,{style:ConditionalStyle}>['type'];
export type ConditionalDraft={type:DialogRule;id?:string;rule?:ConditionalRule;range?:SheetRange};
/** 开始 → 样式 → 条件格式: highlight rules, top/bottom, data bars, colour scales, clear and manage. */
export function ConditionalFormatMenu({disabled,onRule,onQuick,onClear,onManage}:{disabled:boolean;onRule:(type:DialogRule)=>void;onQuick:(rule:ConditionalRule)=>void;onClear:(scope:'selection'|'sheet')=>void;onManage:()=>void}){
 const item=(label:string,action:()=>void,close:()=>void)=><button key={label} role="menuitem" disabled={disabled} onClick={()=>{close();action();}}>{label}</button>;
 return <OfficeMenu label="条件格式" icon={Palette}>{close=><>
  <p className="office-menu-heading" role="presentation">突出显示单元格规则</p>
  {(['greaterThan','lessThan','between','equal','textContains','duplicate'] as DialogRule[]).map(type=>item(RULE_LABELS[type]+'…',()=>onRule(type),close))}
  <p className="office-menu-heading" role="presentation">最前/最后规则</p>
  {item('前 10 项…',()=>onRule('top'),close)}{item('后 10 项…',()=>onRule('bottom'),close)}{item('高于平均值…',()=>onRule('aboveAverage'),close)}{item('低于平均值…',()=>onRule('belowAverage'),close)}
  <p className="office-menu-heading" role="presentation">数据条</p>{BARS.map(bar=>item(bar.label,()=>onQuick({type:'dataBar',color:bar.color}),close))}
  <p className="office-menu-heading" role="presentation">色阶</p>{SCALES.map(scale=>item('色阶 '+scale.label,()=>onQuick({type:'colorScale',colors:scale.colors}),close))}
  <hr/>{item('清除所选单元格的规则',()=>onClear('selection'),close)}{item('清除整个工作表的规则',()=>onClear('sheet'),close)}<button role="menuitem" onClick={()=>{close();onManage();}}>管理规则…</button>
 </>}</OfficeMenu>;
}
/** New-rule dialog (Excel's 突出显示单元格规则 / 最前最后规则 boxes). */
export function ConditionalRuleDialog({draft,busy,onCancel,onSave}:{draft:ConditionalDraft;busy:boolean;onCancel:()=>void;onSave:(rule:ConditionalRule)=>void}){
 const initial=draft.rule as Record<string,unknown>|undefined;
 const [type,setType]=useState<DialogRule>(draft.type),[value,setValue]=useState(String(initial?.value??(draft.type==='top'||draft.type==='bottom'?'':''))),[value2,setValue2]=useState(String(initial?.value2??''));
 const [count,setCount]=useState(String(initial?.count??10)),[percent,setPercent]=useState(!!initial?.percent),[style,setStyle]=useState(()=>{const at=CONDITIONAL_STYLES.findIndex(item=>JSON.stringify(item.style)===JSON.stringify(initial?.style));return at<0?0:at;}),[error,setError]=useState('');
 const kind=['greaterThan','lessThan','greaterOrEqual','lessOrEqual','equal','notEqual'].includes(type)?'compare':['between','notBetween'].includes(type)?'between':['textContains','textNotContains','beginsWith','endsWith'].includes(type)?'text':['duplicate','unique'].includes(type)?'duplicate':['top','bottom'].includes(type)?'rank':'plain';
 const options:Record<string,DialogRule[]>={compare:['greaterThan','lessThan','greaterOrEqual','lessOrEqual','equal','notEqual'],between:['between','notBetween'],text:['textContains','textNotContains','beginsWith','endsWith'],duplicate:['duplicate','unique'],rank:['top','bottom'],plain:['aboveAverage','belowAverage','blank','notBlank']};
 const save=()=>{
  const chosen=CONDITIONAL_STYLES[style].style;let rule:ConditionalRule;
  if(kind==='compare'||kind==='text'){if(!value.trim()){setError('请输入比较的值');return;}rule={type,value:value.trim(),style:chosen} as ConditionalRule;}
  else if(kind==='between'){if(!value.trim()||!value2.trim()||!Number.isFinite(Number(value))||!Number.isFinite(Number(value2))){setError('请输入两个数值');return;}rule={type,value:value.trim(),value2:value2.trim(),style:chosen} as ConditionalRule;}
  else if(kind==='rank'){const n=Number(count);if(!Number.isInteger(n)||n<1||n>(percent?100:1000)){setError(percent?'百分比为 1–100':'项数为 1–1000');return;}rule={type,count:n,...(percent?{percent:true}:{}),style:chosen} as ConditionalRule;}
  else rule={type,style:chosen} as ConditionalRule;
  onSave(rule);
 };
 return <Modal title={draft.id?'编辑条件格式规则':'新建条件格式规则'} onClose={()=>{if(!busy)onCancel();}}>
  <form className="cf-dialog" onSubmit={event=>{event.preventDefault();save();}}>
   <p>应用于：{rangeLabel(draft.range)}</p>
   <label className="form-field">规则<select aria-label="规则类型" value={type} onChange={event=>{setType(event.target.value as DialogRule);setError('');}}>{options[kind].map(option=><option key={option} value={option}>{RULE_LABELS[option]}</option>)}</select></label>
   {(kind==='compare'||kind==='text'||kind==='between')&&<label className="form-field">{kind==='between'?'最小值':'值'}<input aria-label={kind==='between'?'最小值':'比较值'} autoFocus value={value} maxLength={255} onChange={event=>setValue(event.target.value)}/></label>}
   {kind==='between'&&<label className="form-field">最大值<input aria-label="最大值" value={value2} maxLength={255} onChange={event=>setValue2(event.target.value)}/></label>}
   {kind==='rank'&&<><label className="form-field">{percent?'百分比':'项数'}<input aria-label="项数" type="number" min={1} max={percent?100:1000} value={count} onChange={event=>setCount(event.target.value)}/></label><label><input type="checkbox" checked={percent} onChange={event=>setPercent(event.target.checked)}/>按百分比</label></>}
   <label className="form-field">设置为<select aria-label="格式样式" value={style} onChange={event=>setStyle(Number(event.target.value))}>{CONDITIONAL_STYLES.map((item,i)=><option key={i} value={i}>{item.label}</option>)}</select></label>
   <p className="cf-preview" style={{background:CONDITIONAL_STYLES[style].style.fill?'#'+CONDITIONAL_STYLES[style].style.fill:undefined,color:CONDITIONAL_STYLES[style].style.color?'#'+CONDITIONAL_STYLES[style].style.color:undefined,fontWeight:CONDITIONAL_STYLES[style].style.bold?700:undefined}}>示例 AaBbCc 123</p>
   {error&&<p role="alert" className="form-error">{error}</p>}
   <div className="dialog-actions"><button type="button" disabled={busy} onClick={onCancel}>取消</button><button className="primary" disabled={busy}>确定</button></div>
  </form>
 </Modal>;
}
/** 管理规则: rules of this sheet in priority order (first wins), with move, edit and delete. */
export function ConditionalRulesManager({doc,sheet,sheetId,busy,run,onEdit,onClose}:{doc:Y.Doc;sheet:Sheet;sheetId:string;busy:boolean;run:Run;onEdit:(draft:ConditionalDraft)=>void;onClose:()=>void}){
 const rules=sheetObjects(doc,sheetId).filter((entry):entry is [string,SheetConditionalFormat]=>entry[1].kind==='conditional').sort((a,b)=>b[1].priority-a[1].priority);
 const put=(id:string,value:SheetConditionalFormat,previous:SheetConditionalFormat)=>run({operationId:crypto.randomUUID(),type:'putSheetObject',objectId:id,object:value,base:sheetObjectRevision(previous)});
 // Swapping priorities with the neighbour moves a rule; two puts, each with its own revision check.
 const move=async(index:number,delta:-1|1)=>{const [a,b]=[rules[index],rules[index+delta]];if(!a||!b)return;if(await put(a[0],{...a[1],priority:b[1].priority},a[1]))await put(b[0],{...b[1],priority:a[1].priority},b[1]);};
 return <Modal title="条件格式规则管理器" wide onClose={()=>{if(!busy)onClose();}}>
  <p>规则按列表顺序优先：同一单元格的同一项格式取最先满足的规则。</p>
  {!rules.length?<p>此工作表没有条件格式规则。</p>:<ol className="cf-rules">{rules.map(([id,rule],i)=><li key={id}>
   <Swatch rule={rule.rule}/><span className="cf-rule-text">{describeRule(rule.rule)}</span><span className="cf-rule-range">{rangeLabel(resolveSheetRange(sheet,rule.range))}</span>
   <span className="cf-rule-actions"><button type="button" aria-label={`上移规则 ${i+1}`} disabled={busy||i===0} onClick={()=>void move(i,-1)}><ArrowUp size={14} aria-hidden="true"/></button><button type="button" aria-label={`下移规则 ${i+1}`} disabled={busy||i===rules.length-1} onClick={()=>void move(i,1)}><ArrowDown size={14} aria-hidden="true"/></button>
    {'style' in rule.rule&&<button type="button" aria-label={`编辑规则 ${i+1}`} disabled={busy} onClick={()=>onEdit({type:rule.rule.type as DialogRule,id,rule:rule.rule,range:resolveSheetRange(sheet,rule.range)})}><Pencil size={14} aria-hidden="true"/></button>}
    <button type="button" aria-label={`删除规则 ${i+1}`} disabled={busy} onClick={()=>void run({operationId:crypto.randomUUID(),type:'deleteSheetObject',objectId:id,base:sheetObjectRevision(rule)})}><Trash2 size={14} aria-hidden="true"/></button></span>
  </li>)}</ol>}
  <div className="dialog-actions"><button disabled={busy} onClick={onClose}>关闭</button></div>
 </Modal>;
}
/** Creates or replaces a rule on `range`; new rules go on top (highest priority), as in Excel. */
export function saveConditionalRule(doc:Y.Doc,sheet:Sheet,sheetId:string,range:SheetRange,rule:ConditionalRule,run:Run,existing?:string){
 const current=existing?sheetObjects(doc,sheetId).find(([id])=>id===existing)?.[1] as SheetConditionalFormat|undefined:undefined;
 const top=Math.max(0,...sheetObjects(doc,sheetId).filter(([,object])=>object.kind==='conditional').map(([,object])=>(object as SheetConditionalFormat).priority));
 const value:SheetConditionalFormat={kind:'conditional',sheet:sheetId,range:current?current.range:sheetRangeKeys(sheet,range),rule,priority:current?current.priority:top+1};
 return run({operationId:crypto.randomUUID(),type:'putSheetObject',objectId:existing??objectId('cf'),object:value,base:sheetObjectRevision(current)},'已应用条件格式');
}
/** Removes rules on this sheet (all, or those overlapping the selection). Ranges are not split, unlike Excel. */
export async function clearConditionalRules(doc:Y.Doc,sheet:Sheet,sheetId:string,run:Run,range?:SheetRange){
 const targets=sheetObjects(doc,sheetId).filter(([,object])=>{if(object.kind!=='conditional')return false;if(!range)return true;const at=resolveSheetRange(sheet,object.range);return !at||!(at.right<range.left||at.left>range.right||at.bottom<range.top||at.top>range.bottom);});
 for(const [id,object]of targets)if(!await run({operationId:crypto.randomUUID(),type:'deleteSheetObject',objectId:id,base:sheetObjectRevision(object)}))return false;
 return targets.length;
}

// ---- charts ---------------------------------------------------------------------------------------------------------
const CHART_ICONS={column:BarChart3,bar:BarChartHorizontal,line:LineChart,pie:PieChart,area:AreaChart};
export function ChartInsertGroup({disabled,onInsert}:{disabled:boolean;onInsert:(type:SheetChartType)=>void}){
 return <RibbonGroup label="图表">{SHEET_CHART_TYPES.map(type=><RibbonButton key={type} icon={CHART_ICONS[type]} label={CHART_TYPE_LABELS[type]} disabled={disabled} onClick={()=>onInsert(type)}/>)}</RibbonGroup>;
}
/** Contiguous non-empty block around a single cell, like Excel's quick chart range. */
export function currentRegion(value:(r:number,c:number)=>string,rows:number,columns:number,start:[number,number]):SheetRange{
 const seen=new Set<string>(),queue=[start];let range={top:start[0],bottom:start[0],left:start[1],right:start[1]};
 while(queue.length&&seen.size<10000){const [r,c]=queue.pop()!,key=r+':'+c;if(seen.has(key)||r<0||c<0||r>=rows||c>=columns)continue;seen.add(key);if(!value(r,c).trim()&&!(r===start[0]&&c===start[1]))continue;
  range={top:Math.min(range.top,r),bottom:Math.max(range.bottom,r),left:Math.min(range.left,c),right:Math.max(range.right,c)};for(const [dr,dc]of [[1,0],[-1,0],[0,1],[0,-1]])queue.push([r+dr,c+dc]);}
 return range;
}
export function newChart(sheet:Sheet,sheetId:string,range:SheetRange,type:SheetChartType,columns:number,value:(r:number,c:number)=>string,label:(r:number,c:number)=>string):SheetChart{
 const data=chartData(range,value,false,label),rows=sheet.toArray().length,anchorColumn=range.right+2<columns?range.right+2:range.left,anchorRow=range.right+2<columns?range.top:Math.min(rows-1,range.bottom+1);
 const keys=sheetRangeKeys(sheet,range),anchor=sheetRangeKeys(sheet,{top:anchorRow,bottom:anchorRow,left:anchorColumn,right:anchorColumn});
 return {kind:'chart',sheet:sheetId,range:keys,chartType:type,title:data.series.length===1?data.series[0].name:'',legend:type==='pie'||data.series.length>1,anchor:{row:anchor.top,column:anchor.left},width:480,height:288};
}
type Placement={left:number;top:number};
/** Charts float over the grid, anchored at a cell by stable identity; drag to move, corner to resize. */
export function SheetChartLayer({doc,sheet,sheetId,scroller,zoom,editable,selection,value,label,run}:{doc:Y.Doc;sheet:Sheet;sheetId:string;scroller:RefObject<HTMLDivElement>;zoom:number;editable:boolean;selection:SheetRange;value:(r:number,c:number)=>string;label:(r:number,c:number)=>string;run:Run}){
 const charts=sheetObjects(doc,sheetId).filter((entry):entry is [string,SheetChart]=>entry[1].kind==='chart');
 const [places,setPlaces]=useState<Record<string,Placement>>({}),[active,setActive]=useState<string>(),[gesture,setGesture]=useState<{id:string;dx:number;dy:number;width?:number;height?:number}>(),[title,setTitle]=useState<string>();
 const layer=useRef<HTMLDivElement>(null),scale=zoom/100;
 const anchorCell=(chart:SheetChart)=>{const root=scroller.current;if(!root)return;const r=sheet.toArray().findIndex(row=>sheetRowIdentity(row)===chart.anchor.row);
  const cells=[...root.querySelectorAll<HTMLElement>('td[data-cell]')],exact=cells.find(td=>td.dataset.cell===chart.anchor.row+'|'+chart.anchor.column);if(exact)return exact;
  // A hidden (merged or filtered) anchor falls back to the nearest rendered cell in the same column.
  const column=cells.filter(td=>td.dataset.cell!.endsWith('|'+chart.anchor.column));return column.find(td=>Number(td.dataset.row)>=r)??column[column.length-1];};
 useLayoutEffect(()=>{
  const root=scroller.current;if(!root)return;
  const measure=()=>{const box=root.getBoundingClientRect(),next:Record<string,Placement>={};
   for(const [id,chart]of charts){const cell=anchorCell(chart);if(!cell)continue;const rect=cell.getBoundingClientRect();next[id]={left:rect.left-box.left+root.scrollLeft-root.clientLeft,top:rect.top-box.top+root.scrollTop-root.clientTop};}
   setPlaces(previous=>JSON.stringify(previous)===JSON.stringify(next)?previous:next);};
  measure();const observer=new ResizeObserver(measure);const table=root.querySelector('table');if(table)observer.observe(table);return()=>observer.disconnect();
 });
 useEffect(()=>{if(!active)return;const outside=(event:MouseEvent)=>{if(!layer.current?.contains(event.target as Node))setActive(undefined);};document.addEventListener('mousedown',outside);return()=>document.removeEventListener('mousedown',outside);},[active]);
 const put=(id:string,chart:SheetChart,next:SheetChart)=>run({operationId:crypto.randomUUID(),type:'putSheetObject',objectId:id,object:next,base:sheetObjectRevision(chart)});
 const remove=(id:string,chart:SheetChart)=>run({operationId:crypto.randomUUID(),type:'deleteSheetObject',objectId:id,base:sheetObjectRevision(chart)},'已删除图表').then(ok=>{if(ok)setActive(undefined);return ok;});
 const drag=(event:React.PointerEvent<HTMLElement>,id:string,chart:SheetChart,mode:'move'|'resize')=>{
  if(event.button!==0)return;setActive(id);if(!editable)return;event.preventDefault();const target=event.currentTarget,startX=event.clientX,startY=event.clientY;target.setPointerCapture(event.pointerId);let moved=false,last={dx:0,dy:0};
  const move=(next:PointerEvent)=>{last={dx:next.clientX-startX,dy:next.clientY-startY};if(!moved&&Math.hypot(last.dx,last.dy)<4)return;moved=true;
   setGesture(mode==='move'?{id,...last}:{id,dx:0,dy:0,width:Math.round(Math.min(2000,Math.max(160,chart.width+last.dx/scale))),height:Math.round(Math.min(2000,Math.max(160,chart.height+last.dy/scale)))});};
  const up=async()=>{target.removeEventListener('pointermove',move);target.removeEventListener('pointerup',up);target.removeEventListener('pointercancel',up);if(!moved){setGesture(undefined);return;}
   let next=chart;
   if(mode==='resize')next={...chart,width:Math.round(Math.min(2000,Math.max(160,chart.width+last.dx/scale))),height:Math.round(Math.min(2000,Math.max(160,chart.height+last.dy/scale)))};
   else{const frame=target.closest<HTMLElement>('.sheet-chart')!,rect=frame.getBoundingClientRect();layer.current!.style.pointerEvents='none';frame.style.visibility='hidden';
    const hit=document.elementFromPoint(rect.left+4,rect.top+4)?.closest<HTMLElement>('td[data-cell]');frame.style.visibility='';layer.current!.style.pointerEvents='';
    if(hit){const [row,column]=hit.dataset.cell!.split('|');next={...chart,anchor:{row,column}};}}
   if(JSON.stringify(next)!==JSON.stringify(chart))await put(id,chart,next);setGesture(undefined);};
  target.addEventListener('pointermove',move);target.addEventListener('pointerup',up);target.addEventListener('pointercancel',up);
 };
 return <div ref={layer} className="sheet-chart-layer">{charts.map(([id,chart])=>{
  const place=places[id];if(!place)return null;const range=resolveSheetRange(sheet,chart.range),data=range?chartData(range,value,chart.seriesInRows,label):{categories:[],series:[]};
  const g=gesture?.id===id?gesture:undefined,width=g?.width??chart.width,height=g?.height??chart.height,selected=active===id;
  return <div key={id} className="sheet-chart" data-selected={selected||undefined} style={{left:place.left+(g?.dx??0),top:place.top+(g?.dy??0)}} tabIndex={0} aria-label={`图表：${chart.title||CHART_TYPE_LABELS[chart.chartType]}`}
   onPointerDown={event=>{if((event.target as HTMLElement).closest('.sheet-chart-toolbar,.sheet-chart-resize'))return;drag(event,id,chart,'move');}} onFocus={()=>setActive(id)}
   onKeyDown={event=>{if(event.target!==event.currentTarget)return;if(event.key==='Escape')setActive(undefined);if(editable&&(event.key==='Delete'||event.key==='Backspace')){event.preventDefault();void remove(id,chart);}}}>
   <div className="sheet-chart-body" style={{zoom:scale}}><SheetChartView type={chart.chartType} title={chart.title} categories={data.categories} series={data.series} legend={chart.legend} width={width} height={height}/>{!range&&<p className="sheet-chart-warning" role="status">数据区域的行列已删除，请重新选择数据</p>}</div>
   {selected&&editable&&<>
    <span className="sheet-chart-resize" aria-hidden="true" onPointerDown={event=>{event.stopPropagation();drag(event,id,chart,'resize');}}/>
    <div className="sheet-chart-toolbar" role="toolbar" aria-label="图表工具" onPointerDown={event=>event.stopPropagation()}>
     <select aria-label="图表类型" value={chart.chartType} onChange={event=>void put(id,chart,{...chart,chartType:event.target.value as SheetChartType})}>{SHEET_CHART_TYPES.map(type=><option key={type} value={type}>{CHART_TYPE_LABELS[type]}</option>)}</select>
     <input aria-label="图表标题" placeholder="图表标题" maxLength={200} value={title??chart.title} onFocus={()=>setTitle(chart.title)} onChange={event=>setTitle(event.target.value)} onBlur={()=>{if(title!==undefined&&title!==chart.title)void put(id,chart,{...chart,title});setTitle(undefined);}} onKeyDown={event=>{if(event.key==='Enter')event.currentTarget.blur();}}/>
     <button type="button" onClick={()=>{const next={...chart};if(chart.seriesInRows)delete next.seriesInRows;else next.seriesInRows=true;void put(id,chart,next);}}>切换行/列</button>
     <label><input type="checkbox" checked={chart.legend} onChange={event=>void put(id,chart,{...chart,legend:event.target.checked})}/>图例</label>
     <button type="button" title={'改用当前选区 '+rangeLabel(selection)} onClick={()=>void put(id,chart,{...chart,range:sheetRangeKeys(sheet,selection)})}>使用选区 {rangeLabel(selection)}</button>
     <button type="button" className="sheet-chart-delete" onClick={()=>void remove(id,chart)}><Trash2 size={13} aria-hidden="true"/>删除图表</button>
    </div>
   </>}
  </div>;})}</div>;
}
export type {SheetObject};
