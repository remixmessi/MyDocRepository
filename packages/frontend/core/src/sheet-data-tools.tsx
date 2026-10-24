import {useState} from 'react';
import {Modal} from './dialogs';
import {columnName,cellName,dedupePlan,splitPlan,type SheetRange} from './spreadsheet-commands';
import {validSheetObject,SHEET_VALIDATION_OPERATORS,type SheetValidation,type SheetValidationOperator,type SheetValidationRule} from '../../../common/nbstore/src/spreadsheet-objects';
const OPERATOR_LABELS:Record<SheetValidationOperator,string>={between:'介于',notBetween:'不介于',equal:'等于',notEqual:'不等于',greaterThan:'大于',lessThan:'小于',greaterOrEqual:'大于或等于',lessOrEqual:'小于或等于'};
const rangeLabel=(range:SheetRange)=>cellName(range.top,range.left)+(range.bottom>range.top||range.right>range.left?':'+cellName(range.bottom,range.right):'');
/** 数据有效性 dialog (WPS 数据 → 有效性): allowed values, error style and messages for the selected range. */
export function SheetValidationDialog({range,existing,busy,onSave,onClear,onCancel}:{range:SheetRange;existing?:SheetValidation;busy:boolean;onSave:(value:Omit<SheetValidation,'kind'|'sheet'|'range'>)=>void;onClear?:()=>void;onCancel:()=>void}){
 const rule=existing?.rule,[type,setType]=useState<SheetValidationRule['type']>(rule?.type??'list'),[source,setSource]=useState(rule?.type==='list'?rule.items.join(','):''),[operator,setOperator]=useState<SheetValidationOperator>(rule&&rule.type!=='list'?rule.operator:'between');
 const [value,setValue]=useState(rule&&rule.type!=='list'?rule.value:''),[value2,setValue2]=useState(rule&&rule.type!=='list'?rule.value2??'':''),[style,setStyle]=useState<'stop'|'warning'>(existing?.errorStyle??'stop'),[message,setMessage]=useState(existing?.message??''),[prompt,setPrompt]=useState(existing?.prompt??''),[error,setError]=useState('');
 const between=operator==='between'||operator==='notBetween';
 const save=()=>{
  const next:SheetValidationRule=type==='list'?{type,items:[...new Set(source.split(/[,，]/).map(item=>item.trim()).filter(Boolean))]}:between?{type,operator,value:value.trim(),value2:value2.trim()}:{type,operator,value:value.trim()};
  const value_={rule:next,errorStyle:style,...(message.trim()?{message:message.trim()}:{}),...(prompt.trim()?{prompt:prompt.trim()}:{})};
  if(!validSheetObject({kind:'validation',sheet:'first',range:{top:'0-0',left:'c',bottom:'0-0',right:'c'},...value_})){setError(type==='list'?'请填写 1–100 个不重复的选项，用逗号分隔，合计不超过 255 个字符':type==='decimal'?'请填写有效的数字':type==='whole'?'请填写整数':'请填写不小于 0 的整数');return;}
  onSave(value_);
 };
 return <Modal title="数据有效性" onClose={()=>{if(!busy)onCancel();}}>
  <p>应用范围：{rangeLabel(range)}</p>
  <label className="form-field">允许<select aria-label="允许的值" value={type} onChange={event=>setType(event.target.value as SheetValidationRule['type'])}><option value="list">序列（下拉列表）</option><option value="whole">整数</option><option value="decimal">小数</option><option value="textLength">文本长度</option></select></label>
  {type==='list'?<label className="form-field">来源（用逗号分隔）<input aria-label="序列来源" value={source} placeholder="例如：北京,上海,广州" onChange={event=>setSource(event.target.value)}/></label>:<>
   <label className="form-field">数据<select aria-label="比较方式" value={operator} onChange={event=>setOperator(event.target.value as SheetValidationOperator)}>{SHEET_VALIDATION_OPERATORS.map(value=><option key={value} value={value}>{OPERATOR_LABELS[value]}</option>)}</select></label>
   <label className="form-field">{between?'最小值':'数值'}<input aria-label={between?'最小值':'数值'} inputMode="decimal" value={value} onChange={event=>setValue(event.target.value)}/></label>
   {between&&<label className="form-field">最大值<input aria-label="最大值" inputMode="decimal" value={value2} onChange={event=>setValue2(event.target.value)}/></label>}
  </>}
  <label className="form-field">出错时<select aria-label="出错样式" value={style} onChange={event=>setStyle(event.target.value as 'stop'|'warning')}><option value="stop">停止：拒绝输入</option><option value="warning">警告：允许输入并标出</option></select></label>
  <label className="form-field">错误信息（可选）<input aria-label="错误信息" maxLength={255} value={message} onChange={event=>setMessage(event.target.value)}/></label>
  <label className="form-field">输入提示（可选）<input aria-label="输入提示" maxLength={255} value={prompt} onChange={event=>setPrompt(event.target.value)}/></label>
  <p className="dialog-note">与 WPS、Excel 一样，空单元格不检查；粘贴、排序与替换不受限制，不符合的值以红色虚线标出。</p>
  {error&&<p role="alert">{error}</p>}
  <div className="dialog-actions">{onClear&&<button disabled={busy} onClick={onClear}>全部清除</button>}<button disabled={busy} onClick={onCancel}>取消</button><button className="primary" disabled={busy} onClick={save}>确定</button></div>
 </Modal>;
}
/** 删除重复项: choose the compared columns and whether the first row is a header. */
export function SheetDedupeDialog({range,grid,busy,onApply,onCancel}:{range:SheetRange;grid:string[][];busy:boolean;onApply:(keys:number[],header:boolean)=>void;onCancel:()=>void}){
 const columns=Array.from({length:range.right-range.left+1},(_,i)=>range.left+i),[keys,setKeys]=useState(columns),[header,setHeader]=useState(range.bottom>range.top&&grid[range.top]?.slice(range.left,range.right+1).some(value=>value.trim()&&!/^[+-]?\d/.test(value.trim())));
 const plan=keys.length?dedupePlan(grid,range,keys,header):undefined;
 return <Modal title="删除重复项" onClose={()=>{if(!busy)onCancel();}}>
  <p>区域：{rangeLabel(range)}。保留每组重复值中的第一行，其余行删除后下方数据上移，区域末尾清空。</p>
  <label><input type="checkbox" checked={header} onChange={event=>setHeader(event.target.checked)}/>数据包含标题</label>
  <fieldset><legend>比较的列</legend>{columns.map(column=><label key={column}><input type="checkbox" checked={keys.includes(column)} onChange={event=>setKeys(event.target.checked?[...keys,column].sort((a,b)=>a-b):keys.filter(value=>value!==column))}/>{header&&grid[range.top]?.[column]?.trim()?`${grid[range.top][column]}（${columnName(column)} 列）`:columnName(column)+' 列'}</label>)}</fieldset>
  <p role="status">{plan?(plan.removed?`将删除 ${plan.removed} 行重复值，保留 ${plan.kept} 行`:'未发现重复值'):'请至少选择一列'}</p>
  <div className="dialog-actions"><button disabled={busy} onClick={onCancel}>取消</button><button className="primary" disabled={busy||!plan?.removed} onClick={()=>onApply(keys,header)}>删除重复项</button></div>
 </Modal>;
}
const DELIMITERS:[string,string][]=[[',','逗号'],['\t','Tab 键'],[' ','空格'],[';','分号'],['，','中文逗号'],['、','顿号']];
/** 分列: split the first selected column by a delimiter into the columns to its right. */
export function SheetSplitDialog({range,grid,width,busy,onApply,onCancel}:{range:SheetRange;grid:string[][];width:number;busy:boolean;onApply:(delimiter:string)=>void;onCancel:()=>void}){
 const [choice,setChoice]=useState(','),[custom,setCustom]=useState(''),delimiter=choice==='custom'?custom:choice,plan=splitPlan(grid,range,range.left,delimiter,width);
 return <Modal title="分列" onClose={()=>{if(!busy)onCancel();}}>
  <p>按分隔符拆分 {columnName(range.left)} 列第 {range.top+1}–{range.bottom+1} 行，结果写入右侧相邻列。公式单元格不拆分。</p>
  <fieldset><legend>分隔符号</legend>{DELIMITERS.map(([value,label])=><label key={label}><input type="radio" name="sheet-split" checked={choice===value} onChange={()=>setChoice(value)}/>{label}</label>)}<label><input type="radio" name="sheet-split" checked={choice==='custom'} onChange={()=>setChoice('custom')}/>其他<input aria-label="其他分隔符" maxLength={5} value={custom} onChange={event=>{setCustom(event.target.value);setChoice('custom');}}/></label></fieldset>
  <p role="status">{!delimiter?'请输入分隔符':plan.overflow?'部分行拆分后超出工作表右边界，请先在右侧插入列':!plan.edits.length?'没有可拆分的内容':`将拆分为最多 ${plan.columns} 列，修改 ${plan.edits.length} 个单元格`+(plan.overwrite?`，其中 ${plan.overwrite} 个已有内容会被覆盖`:'')}</p>
  <div className="dialog-actions"><button disabled={busy} onClick={onCancel}>取消</button><button className="primary" disabled={busy||!plan.edits.length||plan.overflow} onClick={()=>onApply(delimiter)}>{plan.overwrite?'覆盖并分列':'分列'}</button></div>
 </Modal>;
}
/** 自定义排序: up to three keys, each a column and a direction; the first row can stay as the header. */
export function SheetSortDialog({range,grid,busy,onApply,onCancel}:{range:SheetRange;grid:string[][];busy:boolean;onApply:(keys:{column:number;descending:boolean}[],header:boolean)=>void;onCancel:()=>void}){
 const columns=Array.from({length:range.right-range.left+1},(_,i)=>range.left+i),[header,setHeader]=useState(range.bottom>range.top&&!!grid[range.top]?.slice(range.left,range.right+1).some(value=>value.trim()&&!/^[+-]?\d/.test(value.trim())));
 const [keys,setKeys]=useState([{column:range.left,descending:false}]),name=(column:number)=>header&&grid[range.top]?.[column]?.trim()?`${grid[range.top][column]}（${columnName(column)} 列）`:columnName(column)+' 列';
 const update=(index:number,patch:Partial<{column:number;descending:boolean}>)=>setKeys(keys.map((key,i)=>i===index?{...key,...patch}:key));
 return <Modal title="排序" onClose={()=>{if(!busy)onCancel();}}>
  <p>区域：{rangeLabel(range)}，整行一起移动；数字在前、文本按拼音，空值始终在最后。</p>
  <label><input type="checkbox" checked={header} onChange={event=>setHeader(event.target.checked)}/>数据包含标题</label>
  <ol className="sheet-sort-keys">{keys.map((key,i)=><li key={i}><span>{i?'次要关键字':'主要关键字'}</span><select aria-label={`第 ${i+1} 关键字列`} value={key.column} onChange={event=>update(i,{column:Number(event.target.value)})}>{columns.map(column=><option key={column} value={column}>{name(column)}</option>)}</select><select aria-label={`第 ${i+1} 关键字次序`} value={key.descending?'desc':'asc'} onChange={event=>update(i,{descending:event.target.value==='desc'})}><option value="asc">升序</option><option value="desc">降序</option></select>{i>0&&<button type="button" onClick={()=>setKeys(keys.filter((_,j)=>j!==i))}>删除条件</button>}</li>)}</ol>
  <button type="button" disabled={keys.length>=3||keys.length>=columns.length} onClick={()=>setKeys([...keys,{column:columns.find(column=>!keys.some(key=>key.column===column))??range.left,descending:false}])}>添加条件</button>
  <div className="dialog-actions"><button disabled={busy} onClick={onCancel}>取消</button><button className="primary" disabled={busy} onClick={()=>onApply(keys,header)}>确定</button></div>
 </Modal>;
}
/** 行高 (points) / 列宽 (characters of the default font), the units WPS and Excel show. */
export function SheetSizeDialog({axis,initial,busy,onApply,onCancel}:{axis:'row'|'column';initial:number;busy:boolean;onApply:(pixels:number)=>void;onCancel:()=>void}){
 const toUnit=(px:number)=>axis==='row'?Math.round(px*0.75*10)/10:Math.round((px-5)/7*10)/10,fromUnit=(value:number)=>Math.round(axis==='row'?value*4/3:value*7+5);
 const [value,setValue]=useState(String(toUnit(initial))),pixels=fromUnit(Number(value)),limits=axis==='row'?{min:16,max:600}:{min:20,max:1000},ok=Number.isFinite(Number(value))&&value.trim()!==''&&pixels>=limits.min&&pixels<=limits.max;
 return <Modal title={axis==='row'?'行高':'列宽'} onClose={()=>{if(!busy)onCancel();}}><form onSubmit={event=>{event.preventDefault();if(ok)onApply(pixels);}}>
  <label className="form-field">{axis==='row'?'行高（磅）':'列宽（字符）'}<input autoFocus aria-label={axis==='row'?'行高':'列宽'} inputMode="decimal" value={value} onChange={event=>setValue(event.target.value)}/></label>
  <p className="dialog-note">{axis==='row'?`范围 ${toUnit(limits.min)}–${toUnit(limits.max)} 磅，默认 ${toUnit(25)} 磅。`:`范围 ${toUnit(limits.min)}–${toUnit(limits.max)} 个字符，默认 ${toUnit(96)}。`}</p>
  <div className="dialog-actions"><button type="button" disabled={busy} onClick={onCancel}>取消</button><button className="primary" disabled={busy||!ok}>确定</button></div>
 </form></Modal>;
}
