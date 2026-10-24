import {useState,type ReactNode} from 'react';
import {SHEET_FUNCTION_GROUPS} from '../../../common/nbstore/src/spreadsheet-function-catalog';
import {Bold,Italic,Underline,Strikethrough,AlignLeft,AlignCenter,AlignRight,AlignVerticalJustifyStart,AlignVerticalJustifyCenter,AlignVerticalJustifyEnd,WrapText,ClipboardPaste,Copy,Paintbrush,Merge,Split,Percent,Sigma,ArrowDownAZ,ArrowUpZA,Filter,Search,Eraser,RemoveFormatting,Rows3,Columns3,Trash2,Snowflake,Grid3x3,Grid2x2,SquareFunction,MessageSquarePlus,Baseline,PaintBucket,Hash,ListChecks,CopyMinus,SplitSquareHorizontal,ListOrdered} from 'lucide-react';
import {OfficeRibbon,RibbonGroup,RibbonButton,OfficeMenu} from './office-ribbon';
import {sheetCellFormat,SHEET_NUMBER_FORMATS,SHEET_FONT_SIZES,SHEET_DEFAULT_FONT_SIZE,SHEET_BORDER_STYLES,type SheetCellFormat,type SheetMergeLayout,type SheetBorderStyle} from '../../../common/nbstore/src/spreadsheet-format';
import type {SheetFormatPatch} from '../../../common/nbstore/src/spreadsheet-operation';
import type {SpreadsheetCellType} from '../../../common/nbstore/src/spreadsheet-cell-type';
import {adjustDecimals,type SheetRange,type BorderPreset} from './spreadsheet-commands';
import type {SheetCellHandle} from './spreadsheet-cell';
import type {OfficeSpreadsheetDocument} from './workspace';
export const SHEET_TABS=['开始','插入','公式','数据','视图'];
const FORMAT_LABELS:Record<string,string>={'':'常规','0':'整数','0.00':'两位小数','#,##0':'千分位','#,##0.00':'千分位两位小数','0%':'百分比','0.00%':'百分比两位小数','¥#,##0.00':'人民币','$#,##0.00':'美元','yyyy-mm-dd':'日期 2026-10-20','yyyy/m/d':'日期 2026/10/20','yyyy年m月d日':'日期 2026年10月20日','hh:mm':'时间 14:30','hh:mm:ss':'时间 14:30:00','yyyy-mm-dd hh:mm':'日期时间','@':'文本'};
/** Quick aggregate buttons (自动求和); the full library is SHEET_FUNCTION_GROUPS. */
export const SHEET_FUNCTIONS=[['SUM','求和'],['AVERAGE','平均值'],['COUNT','计数'],['MAX','最大值'],['MIN','最小值']] as const;
/** Grouped function picker used by 公式 → 函数库 and the formula bar's fx box. */
export function SheetFunctionSelect({label,placeholder,disabled,onPick}:{label:string;placeholder:string;disabled:boolean;onPick:(name:string)=>void}){
 return <select aria-label={label} value="" disabled={disabled} onChange={event=>{if(event.target.value)onPick(event.target.value);}}><option value="">{placeholder}</option>{SHEET_FUNCTION_GROUPS.map(group=><optgroup key={group.group} label={group.group}>{group.functions.map(entry=><option key={entry.name} value={entry.name} title={entry.syntax}>{entry.name} · {entry.label}</option>)}</optgroup>)}</select>;
}
/** Local view settings. `freezeRows`/`freezeColumns` count the rows above and columns left of the split (冻结窗格). */
export type SheetView={freezeRows:number;freezeColumns:number;gridlines:boolean;formulas:boolean;filter:boolean};
const BORDER_PRESETS:[BorderPreset,string][]=[['all','所有框线'],['outside','外侧框线'],['thickOutside','粗外侧框线'],['top','上框线'],['bottom','下框线'],['left','左框线'],['right','右框线'],['none','无框线']];
const BORDER_LABELS:Record<SheetBorderStyle,string>={thin:'细实线',medium:'中等实线',thick:'粗实线',dashed:'虚线',dotted:'点线',double:'双线'};
export interface SpreadsheetActions {
 format:(patch:SheetFormatPatch)=>void;merge:()=>void;unmerge:()=>void;structure:(type:'insertRow'|'deleteRow'|'insertColumn'|'deleteColumn'|'appendRow'|'deleteLastRow',after?:boolean)=>void;
 autoSum:(fn:string)=>void;sort:(descending:boolean)=>void;clear:()=>void;copy:()=>void;paste:()=>void;find:()=>void;insertFunction:(fn:string)=>void;comment:()=>void;
 painter:()=>void;error:(message:string)=>void;
 borders:(preset:BorderPreset,line:string)=>void;clearFormat:()=>void;validation:()=>void;dedupe:()=>void;split:()=>void;freezeAtSelection:()=>void;
 size:(axis:'row'|'column')=>void;autofit:(axis:'row'|'column')=>void;hide:(axis:'row'|'column',hidden:boolean)=>void;customSort:()=>void;
}
export function SpreadsheetRibbon({item,range,merges,handle,disabled,readOnly,actions,painter,view,onView,rowCount,columnCount,styles,insert}:{item:OfficeSpreadsheetDocument;range:SheetRange;merges:SheetMergeLayout;handle?:SheetCellHandle;disabled:boolean;readOnly:boolean;actions:SpreadsheetActions;painter:boolean;view:SheetView;onView:(view:SheetView)=>void;rowCount:number;columnCount:number;styles?:ReactNode;insert?:ReactNode}){
 const [tab,setTab]=useState('开始');
 const formats:(SheetCellFormat|undefined)[]=[];for(let r=range.top;r<=range.bottom&&r<rowCount;r++)for(let c=range.left;c<=range.right&&c<columnCount;c++)formats.push(sheetCellFormat(item.doc,item.sheetId!,item.sheet,r,c));
 const [line,setLine]=useState<SheetBorderStyle>('thin'),[lineColor,setLineColor]=useState('000000');
 const first=formats[0],all=(key:'bold'|'italic'|'underline'|'strike'|'wrap')=>formats.length>0&&formats.every(format=>format?.[key]),count=(range.bottom-range.top+1)*(range.right-range.left+1);
 const owner=merges.owner.get(range.top+':'+range.left),merged=owner?merges.anchors.get(owner):undefined,off=disabled||!formats.length;
 const decimals=(delta:1|-1)=>{const next=adjustDecimals(first?.numberFormat,delta);if(next)actions.format({numberFormat:next});else actions.error('当前数字格式不支持调整小数位');};
 const toggle=(key:'gridlines'|'formulas'|'filter')=>onView({...view,[key]:!view[key]}),frozen=view.freezeRows>0||view.freezeColumns>0;
 const freeze=<RibbonGroup label="冻结窗格"><RibbonButton icon={Snowflake} label="冻结至当前单元格" onClick={actions.freezeAtSelection}/><RibbonButton icon={Snowflake} label="冻结首行" pressed={view.freezeRows===1&&view.freezeColumns===0} onClick={()=>onView({...view,freezeRows:view.freezeRows===1&&!view.freezeColumns?0:1,freezeColumns:0})}/><RibbonButton icon={Snowflake} label="冻结首列" pressed={view.freezeColumns===1&&view.freezeRows===0} onClick={()=>onView({...view,freezeColumns:view.freezeColumns===1&&!view.freezeRows?0:1,freezeRows:0})}/><RibbonButton label="取消冻结" disabled={!frozen} onClick={()=>onView({...view,freezeRows:0,freezeColumns:0})}/></RibbonGroup>;
 const borderMenu=<OfficeMenu label="框线" icon={Grid2x2}>{close=><>
  {BORDER_PRESETS.map(([preset,label])=><button key={preset} role="menuitem" disabled={off} onClick={()=>{close();actions.borders(preset,line+(lineColor!=='000000'?':'+lineColor:''));}}>{label}</button>)}
  <p className="office-menu-heading" role="presentation">线条样式</p>{SHEET_BORDER_STYLES.filter(style=>style!=='thick').map(style=><button key={style} role="menuitemradio" aria-checked={line===style} onClick={()=>setLine(style)}>{BORDER_LABELS[style]}</button>)}
  <label className="office-menu-color">线条颜色<input type="color" aria-label="框线颜色" value={'#'+lineColor} onChange={event=>setLineColor(event.target.value.slice(1).toUpperCase())}/></label>
 </>}</OfficeMenu>;
 const rows=<RibbonGroup label="行和列"><RibbonButton icon={Rows3} label="在当前行上方插入" disabled={disabled} onClick={()=>actions.structure('insertRow')}/><RibbonButton icon={Trash2} label="删除当前行" disabled={disabled||!rowCount} onClick={()=>actions.structure('deleteRow')}/><RibbonButton icon={Columns3} label="在当前列左侧插入" disabled={disabled} onClick={()=>actions.structure('insertColumn')}/><RibbonButton icon={Trash2} label="删除当前列" disabled={disabled||columnCount<=1} onClick={()=>actions.structure('deleteColumn')}/><RibbonButton label="新增行" disabled={disabled} onClick={()=>actions.structure('appendRow')}/><RibbonButton label="删除末行" disabled={disabled||!rowCount} onClick={()=>actions.structure('deleteLastRow')}/></RibbonGroup>;
 return <OfficeRibbon tabs={SHEET_TABS} tab={tab} onTab={setTab} label="表格功能区">
  {tab==='开始'&&<>
   <RibbonGroup label="剪贴板"><RibbonButton icon={ClipboardPaste} label="批量粘贴" large disabled={disabled} onClick={actions.paste}/><RibbonButton icon={Copy} label="复制选区" compact onClick={actions.copy}/><RibbonButton icon={Paintbrush} label="格式刷" compact pressed={painter} disabled={off} onClick={actions.painter}/></RibbonGroup>
   <RibbonGroup label="字体"><select aria-label="字号" title="字号" disabled={off} value={first?.fontSize??SHEET_DEFAULT_FONT_SIZE} onChange={event=>{const size=Number(event.target.value);actions.format({fontSize:size===SHEET_DEFAULT_FONT_SIZE?null:size});}}>{[...new Set([...SHEET_FONT_SIZES,first?.fontSize??SHEET_DEFAULT_FONT_SIZE])].sort((a,b)=>a-b).map(size=><option key={size} value={size}>{size}</option>)}</select>
    {(['bold','italic','underline','strike'] as const).map(key=><RibbonButton key={key} icon={{bold:Bold,italic:Italic,underline:Underline,strike:Strikethrough}[key]} label={{bold:'加粗',italic:'斜体',underline:'下划线',strike:'删除线'}[key]} compact pressed={all(key)} disabled={off} onClick={()=>actions.format({[key]:all(key)?null:true})}/>)}{borderMenu}
    <label className="ribbon-color" title="文字颜色"><Baseline size={15} aria-hidden="true"/><input type="color" aria-label="单元格文字颜色" disabled={off} value={'#'+(first?.color??'000000')} onChange={event=>actions.format({color:event.target.value.slice(1).toUpperCase()})}/></label>
    <label className="ribbon-color" title="填充颜色"><PaintBucket size={15} aria-hidden="true"/><input type="color" aria-label="单元格填充色" disabled={off} value={'#'+(first?.fill??'FFFFFF')} onChange={event=>actions.format({fill:event.target.value.slice(1).toUpperCase()})}/></label>
    <RibbonButton icon={Eraser} label="清除颜色" compact disabled={off} onClick={()=>actions.format({color:null,fill:null})}/></RibbonGroup>
   <RibbonGroup label="对齐方式">{(['left','center','right'] as const).map(align=><RibbonButton key={align} icon={{left:AlignLeft,center:AlignCenter,right:AlignRight}[align]} label={{left:'左对齐',center:'居中',right:'右对齐'}[align]} compact pressed={first?.align===align} disabled={off} onClick={()=>actions.format({align:first?.align===align?null:align})}/>)}
    {(['top','middle','bottom'] as const).map(value=><RibbonButton key={value} icon={{top:AlignVerticalJustifyStart,middle:AlignVerticalJustifyCenter,bottom:AlignVerticalJustifyEnd}[value]} label={{top:'顶端对齐',middle:'垂直居中',bottom:'底端对齐'}[value]} compact pressed={(first?.verticalAlign??'middle')===value} disabled={off} onClick={()=>actions.format({verticalAlign:value==='middle'?null:value})}/>)}
    <RibbonButton icon={WrapText} label="自动换行" compact pressed={all('wrap')} disabled={off} onClick={()=>actions.format({wrap:all('wrap')?null:true})}/>
    <RibbonButton icon={Merge} label="合并单元格" disabled={off||count<2||!!merged} onClick={actions.merge}/><RibbonButton icon={Split} label="取消合并" disabled={disabled||!merged} onClick={actions.unmerge}/></RibbonGroup>
   <RibbonGroup label="数字"><select aria-label="数字格式" disabled={off} value={first?.numberFormat??''} onChange={event=>actions.format({numberFormat:event.target.value||null})}>{first?.numberFormat&&!SHEET_NUMBER_FORMATS.includes(first.numberFormat)&&<option value={first.numberFormat}>导入格式 {first.numberFormat}</option>}{['',...SHEET_NUMBER_FORMATS].map(code=><option key={code} value={code}>{FORMAT_LABELS[code]??code}</option>)}</select>
    <RibbonButton icon={Percent} label="百分比样式" compact disabled={off} onClick={()=>actions.format({numberFormat:'0%'})}/><RibbonButton icon={Hash} label="千位分隔" compact disabled={off} onClick={()=>actions.format({numberFormat:'#,##0'})}/><RibbonButton label=".0+" aria-label="增加小数位" title="增加小数位" disabled={off} onClick={()=>decimals(1)}/><RibbonButton label=".0-" aria-label="减少小数位" title="减少小数位" disabled={off} onClick={()=>decimals(-1)}/>
    <select aria-label="单元格内容类型" title="内容类型" value={handle?.kind??'auto'} disabled={readOnly||!handle||handle.busy} onChange={event=>handle?.stage(handle.value,event.target.value as SpreadsheetCellType)} onBlur={()=>{if(handle?.dirty)void handle.commit();}}><option value="text">文本</option><option value="formula">公式（以 = 开头）</option><option value="auto">自动（兼容旧表格）</option></select></RibbonGroup>
   {styles&&<RibbonGroup label="样式">{styles}</RibbonGroup>}
   {rows}
   <RibbonGroup label="单元格"><OfficeMenu label="行和列" icon={Rows3}>{close=><>{([['行高…',()=>actions.size('row')],['最适合的行高',()=>actions.autofit('row')],['列宽…',()=>actions.size('column')],['最适合的列宽',()=>actions.autofit('column')]] as const).map(([label,run])=><button key={label} role="menuitem" disabled={disabled} onClick={()=>{close();run();}}>{label}</button>)}<hr/>{([['隐藏行',()=>actions.hide('row',true)],['取消隐藏行',()=>actions.hide('row',false)],['隐藏列',()=>actions.hide('column',true)],['取消隐藏列',()=>actions.hide('column',false)]] as const).map(([label,run])=><button key={label} role="menuitem" disabled={disabled} onClick={()=>{close();run();}}>{label}</button>)}</>}</OfficeMenu></RibbonGroup>
   <RibbonGroup label="编辑"><RibbonButton icon={Sigma} label="求和" disabled={disabled} onClick={()=>actions.autoSum('SUM')}/><RibbonButton icon={ArrowDownAZ} label="升序" disabled={disabled} onClick={()=>actions.sort(false)}/><RibbonButton icon={Filter} label="筛选" pressed={view.filter} onClick={()=>toggle('filter')}/><RibbonButton icon={Search} label="查找替换" onClick={actions.find}/><RibbonButton icon={Eraser} label="清除内容" disabled={disabled} onClick={actions.clear}/><RibbonButton icon={RemoveFormatting} label="清除格式" disabled={off} onClick={actions.clearFormat}/></RibbonGroup>
   {freeze}
  </>}
  {tab==='插入'&&<>
   {insert}
   {rows}
   <RibbonGroup label="行和列（之后）"><RibbonButton icon={Rows3} label="在当前行下方插入" disabled={disabled} onClick={()=>actions.structure('insertRow',true)}/><RibbonButton icon={Columns3} label="在当前列右侧插入" disabled={disabled} onClick={()=>actions.structure('insertColumn',true)}/></RibbonGroup>
   <RibbonGroup label="函数与批注"><RibbonButton icon={SquareFunction} label="插入函数" disabled={disabled||!handle} onClick={()=>actions.insertFunction('SUM')}/><RibbonButton icon={MessageSquarePlus} label="新建批注" onClick={actions.comment}/></RibbonGroup>
  </>}
  {tab==='公式'&&<>
   <RibbonGroup label="自动求和">{SHEET_FUNCTIONS.slice(0,5).map(([fn,label])=><RibbonButton key={fn} icon={fn==='SUM'?Sigma:undefined} label={label} disabled={disabled} onClick={()=>actions.autoSum(fn)}/>)}</RibbonGroup>
   <RibbonGroup label="函数库"><SheetFunctionSelect label="插入函数" placeholder="全部函数…" disabled={disabled||!handle} onPick={actions.insertFunction}/>{SHEET_FUNCTION_GROUPS.slice(1,5).map(group=><SheetFunctionSelect key={group.group} label={group.group+'函数'} placeholder={group.group} disabled={disabled||!handle} onPick={actions.insertFunction}/>)}</RibbonGroup>
   <RibbonGroup label="公式审核"><RibbonButton icon={SquareFunction} label="显示公式" pressed={view.formulas} onClick={()=>toggle('formulas')}/></RibbonGroup>
  </>}
  {tab==='数据'&&<>
   <RibbonGroup label="排序和筛选"><RibbonButton icon={ArrowDownAZ} label="升序" disabled={disabled} onClick={()=>actions.sort(false)}/><RibbonButton icon={ArrowUpZA} label="降序" disabled={disabled} onClick={()=>actions.sort(true)}/><RibbonButton icon={ListOrdered} label="自定义排序" disabled={disabled} onClick={actions.customSort}/><RibbonButton icon={Filter} label="筛选" pressed={view.filter} onClick={()=>toggle('filter')}/></RibbonGroup>
   <RibbonGroup label="数据工具"><RibbonButton icon={ListChecks} label="有效性" large disabled={disabled} onClick={actions.validation}/><RibbonButton icon={CopyMinus} label="删除重复项" disabled={disabled} onClick={actions.dedupe}/><RibbonButton icon={SplitSquareHorizontal} label="分列" disabled={disabled} onClick={actions.split}/><RibbonButton icon={ClipboardPaste} label="批量粘贴" disabled={disabled} onClick={actions.paste}/><RibbonButton icon={Search} label="查找替换" onClick={actions.find}/><RibbonButton icon={Eraser} label="清除内容" disabled={disabled} onClick={actions.clear}/></RibbonGroup>
  </>}
  {tab==='视图'&&<>
   {freeze}
   <RibbonGroup label="显示"><RibbonButton icon={Grid3x3} label="网格线" pressed={view.gridlines} onClick={()=>toggle('gridlines')}/><RibbonButton icon={SquareFunction} label="显示公式" pressed={view.formulas} onClick={()=>toggle('formulas')}/></RibbonGroup>
  </>}
 </OfficeRibbon>;
}
