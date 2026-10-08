import {useEffect,useRef,useState} from 'react';
import {ChevronLeft,ChevronRight} from 'lucide-react';
import {Modal} from './dialogs';
import type {LocalWorkspace,DocumentMeta} from './workspace';
import {localJournalDate,validJournalDate,journalMonthDays,shiftJournalDay,shiftJournalMonth,journalWeekday,type JournalTemplate} from './journal';

export function JournalDialog({workspace,initialDate,onClose,onOpen}:{workspace:LocalWorkspace;initialDate?:string;onClose:()=>void;onOpen:(id:string)=>void}) {
  const [date,setDate]=useState(initialDate??localJournalDate()),[today,setToday]=useState(localJournalDate),[template,setTemplate]=useState<JournalTemplate>('daily');
  const [busy,setBusy]=useState(false),[error,setError]=useState('');const [,refresh]=useState(0);
  useEffect(()=>workspace.subscribe(()=>refresh(n=>n+1)),[workspace]);
  useEffect(()=>{const timer=setInterval(()=>setToday(localJournalDate()),30000);return()=>clearInterval(timer);},[]);
  const records=new Map(workspace.list().filter(meta=>meta.journalDate).map(meta=>[meta.journalDate!,meta]));
  const recycled=new Map(workspace.trashList().filter(entry=>entry.meta.journalDate).map(entry=>[entry.meta.journalDate!,entry.meta]));
  const select=(next:string)=>{setDate(next);setError('');};
  const record=records.get(date),trash=recycled.get(date);
  const perform=async()=>{if(busy)return;setBusy(true);setError('');try{const id=trash?await workspace.restoreDocument(trash.id):await workspace.getOrCreateJournal(date,template);onOpen(id);}catch(e){setError(e instanceof Error?e.message:'日记操作失败，请重试');}finally{setBusy(false);}};
  return <Modal title="日历笔记" onClose={()=>{if(!busy)onClose();}}>
    <div className="journal-toolbar"><button disabled={busy} onClick={()=>select(localJournalDate())}>回到今天</button><label>跳转日期<input aria-label="日记日期" type="date" min="0001-01-01" max="9999-12-31" disabled={busy} value={date} onChange={e=>{if(validJournalDate(e.target.value))select(e.target.value);else setError('请选择有效日期');}}/></label></div>
    <MonthCalendar date={date} today={today} records={records} recycled={recycled} busy={busy} onSelect={select}/>
    <JournalDayDetail date={date} record={record} trash={trash} template={template} onTemplate={setTemplate} busy={busy} onOpen={()=>void perform()}/>
    {busy&&<p role="status" className="organization-help">正在打开日记…</p>}{error&&<p role="alert" className="form-error">{error} 可再次点击重试。</p>}
    <p className="organization-help">日期归属与标题独立，日记也可使用标签、收藏、搜索和工作区备份。</p>
  </Modal>;
}
function MonthCalendar({date,today,records,recycled,busy,onSelect}:{date:string;today:string;records:Map<string,DocumentMeta>;recycled:Map<string,DocumentMeta>;busy:boolean;onSelect:(date:string)=>void}) {
  const host=useRef<HTMLDivElement>(null),month=date.slice(0,7),days=journalMonthDays(date);
  const count=[...records.keys()].filter(day=>day.startsWith(month)).length;
  const move=(next:string)=>{onSelect(next);requestAnimationFrame(()=>host.current?.querySelector<HTMLButtonElement>(`button[data-date="${next}"]`)?.focus());};
  return <section aria-label="日记月历" className="journal-calendar"><header><button aria-label="上个月" disabled={busy||shiftJournalMonth(date,-1)===date} onClick={()=>onSelect(shiftJournalMonth(date,-1))}><ChevronLeft size={16}/></button><h3 aria-live="polite">{Number(date.slice(0,4))} 年 {Number(date.slice(5,7))} 月</h3><button aria-label="下个月" disabled={busy||shiftJournalMonth(date,1)===date} onClick={()=>onSelect(shiftJournalMonth(date,1))}><ChevronRight size={16}/></button></header>
    <div className="journal-weekdays" aria-hidden="true">{['一','二','三','四','五','六','日'].map(day=><span key={day}>{day}</span>)}</div>
    <div className="journal-days" ref={host}>{days.map((day,i)=>day?<button key={day} type="button" data-date={day} disabled={busy} tabIndex={day===date?0:-1} aria-pressed={day===date} aria-current={day===today?'date':undefined} aria-label={`${day}${day===today?' 今天':''}${records.has(day)?' 有日记':recycled.has(day)?' 日记在回收站':''}`} className={`${day.startsWith(month)?'':'outside-month'} ${day===date?'selected':''}`} onClick={()=>onSelect(day)} onKeyDown={e=>{const keys:Record<string,()=>string>={ArrowLeft:()=>shiftJournalDay(day,-1),ArrowRight:()=>shiftJournalDay(day,1),ArrowUp:()=>shiftJournalDay(day,-7),ArrowDown:()=>shiftJournalDay(day,7),Home:()=>shiftJournalDay(day,-journalWeekday(day)),End:()=>shiftJournalDay(day,6-journalWeekday(day)),PageUp:()=>shiftJournalMonth(day,-1),PageDown:()=>shiftJournalMonth(day,1)};if(keys[e.key]){e.preventDefault();move(keys[e.key]());}}}><span>{Number(day.slice(8))}</span><small>{records.has(day)?'记录':recycled.has(day)?'回收':''}</small></button>:<span key={'empty-'+i}/>)}</div>
    <p className="journal-month-count">本月 {count} 篇日记 · 方向键选择日期</p>
  </section>;
}
function JournalDayDetail({date,record,trash,template,onTemplate,busy,onOpen}:{date:string;record?:DocumentMeta;trash?:DocumentMeta;template:JournalTemplate;onTemplate:(value:JournalTemplate)=>void;busy:boolean;onOpen:()=>void}) {
  return <section className="journal-day-detail" aria-label="所选日期日记"><h3>{date}</h3>{record?<><p className="journal-record-title">{record.title||'无标题文档'}</p><p className="organization-help">这一天已有记录，打开后继续编辑。</p></>:trash?<><p className="journal-record-title">{trash.title||'无标题文档'}</p><p className="organization-help">这一天的日记在回收站中。恢复后继续编辑原来的内容。</p></>:<><p className="organization-help">这一天还没有日记。选择模板开始记录。</p><label className="journal-template">日记模板<select aria-label="日记模板" value={template} disabled={busy} onChange={e=>onTemplate(e.target.value as JournalTemplate)}><option value="daily">日常记录（记录／待办／回顾）</option><option value="blank">空白正文</option></select></label></>}
    <div className="dialog-actions"><button className="primary" disabled={busy} onClick={onOpen}>{record?'打开日记':trash?'恢复并打开日记':'创建日记'}</button></div>
  </section>;
}
