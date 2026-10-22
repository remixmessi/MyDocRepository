import {useId} from 'react';
import type {ChartSeries,SheetChartType} from '../../../common/nbstore/src/spreadsheet-objects';

/** Office's default series colours (accent 1–6, then darker variants). */
export const CHART_COLORS=['4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47','264478','9E480E','636363','997300'];
export const CHART_TYPE_LABELS:Record<SheetChartType,string>={column:'柱形图',bar:'条形图',line:'折线图',pie:'饼图',area:'面积图'};
/** 1, 2 or 5 × 10ⁿ ticks covering the data and zero. */
function niceScale(values:number[],ticks=5){
 let min=Math.min(0,...values),max=Math.max(0,...values);if(min===max){max=min+1;}
 const raw=(max-min)/ticks,power=10**Math.floor(Math.log10(raw)),step=[1,2,5,10].map(n=>n*power).find(n=>n>=raw)!;
 min=Math.floor(min/step)*step;max=Math.ceil(max/step)*step;const list:number[]=[];for(let v=min;v<=max+step/2;v+=step)list.push(Number(v.toPrecision(12)));return {min,max,ticks:list};
}
const short=(text:string,limit:number)=>[...text].length>limit?[...text].slice(0,Math.max(1,limit-1)).join('')+'…':text;
const tick=(value:number)=>Math.abs(value)>=1e6?Number((value/1e6).toPrecision(3))+'M':Math.abs(value)>=1e4?Number((value/1e3).toPrecision(3))+'k':String(Number(value.toPrecision(6)));

/** Chart drawn as SVG; a visually hidden table carries the same data for screen readers. */
export function SheetChartView({type,title,categories,series,legend,width,height}:{type:SheetChartType;title:string;categories:string[];series:ChartSeries[];legend:boolean;width:number;height:number}){
 const id=useId(),titleHeight=title?28:8,legendHeight=legend?24:0,values=series.flatMap(item=>item.values.filter((value):value is number=>value!==null));
 const color=(i:number)=>'#'+CHART_COLORS[i%CHART_COLORS.length],label=title||CHART_TYPE_LABELS[type];
 const table=<table className="sr-only"><caption>{label}</caption><thead><tr><th scope="col">分类</th>{series.map((item,i)=><th key={i} scope="col">{item.name}</th>)}</tr></thead><tbody>{categories.map((category,c)=><tr key={c}><th scope="row">{category}</th>{series.map((item,i)=><td key={i}>{item.values[c]??''}</td>)}</tr>)}</tbody></table>;
 const empty=!values.length||!categories.length;
 let body:JSX.Element|null=null,legendItems:{name:string;color:string}[]=series.map((item,i)=>({name:item.name,color:color(i)}));
 const top=titleHeight,bottom=height-legendHeight;
 if(empty)body=<text x={width/2} y={(top+bottom)/2} textAnchor="middle" className="sheet-chart-empty">所选区域没有可绘制的数值</text>;
 else if(type==='pie'){
  const data=series[0].values.map((value,i)=>({value:value&&value>0?value:0,name:categories[i],color:color(i)})).filter(item=>item.value>0),total=data.reduce((a,b)=>a+b.value,0);
  legendItems=data.map(item=>({name:item.name,color:item.color}));const cx=width/2,cy=(top+bottom)/2,r=Math.max(10,Math.min(width/2-16,(bottom-top)/2-8));let angle=-Math.PI/2;
  body=<g>{data.map((item,i)=>{const sweep=item.value/total*Math.PI*2,start=angle,end=(angle+=sweep),mid=(start+end)/2;
   if(data.length===1)return <circle key={i} cx={cx} cy={cy} r={r} fill={item.color}/>;
   const path=`M${cx},${cy}L${cx+r*Math.cos(start)},${cy+r*Math.sin(start)}A${r},${r} 0 ${sweep>Math.PI?1:0} 1 ${cx+r*Math.cos(end)},${cy+r*Math.sin(end)}Z`;
   return <g key={i}><path d={path} fill={item.color} stroke="#fff" strokeWidth={1}><title>{item.name}：{item.value}（{Math.round(item.value/total*1000)/10}%）</title></path>{sweep>0.35&&<text x={cx+r*0.62*Math.cos(mid)} y={cy+r*0.62*Math.sin(mid)} textAnchor="middle" dominantBaseline="middle" className="sheet-chart-slice">{Math.round(item.value/total*100)}%</text>}</g>;})}</g>;
 }else{
  const scale=niceScale(values),horizontal=type==='bar',axisLeft=horizontal?Math.min(110,Math.max(48,Math.max(...categories.map(c=>[...c].length))*7+12)):Math.max(36,...scale.ticks.map(t=>tick(t).length*7+10)),plot={left:axisLeft,right:width-14,top:top+6,bottom:bottom-(horizontal?22:26)};
  const span=scale.max-scale.min,valueAt=(v:number)=>horizontal?plot.left+(v-scale.min)/span*(plot.right-plot.left):plot.bottom-(v-scale.min)/span*(plot.bottom-plot.top),zero=valueAt(0);
  const n=categories.length,band=(horizontal?plot.bottom-plot.top:plot.right-plot.left)/n,catAt=(i:number)=>horizontal?plot.bottom-band*(i+0.5):plot.left+band*(i+0.5);
  const grid=scale.ticks.map(t=><g key={t}>{horizontal?<line x1={valueAt(t)} x2={valueAt(t)} y1={plot.top} y2={plot.bottom} className="sheet-chart-grid"/>:<line x1={plot.left} x2={plot.right} y1={valueAt(t)} y2={valueAt(t)} className="sheet-chart-grid"/>}<text x={horizontal?valueAt(t):plot.left-6} y={horizontal?plot.bottom+14:valueAt(t)} textAnchor={horizontal?'middle':'end'} dominantBaseline={horizontal?'auto':'middle'} className="sheet-chart-axis">{tick(t)}</text></g>);
  const step=Math.max(1,Math.ceil(n/Math.max(1,Math.floor((horizontal?plot.bottom-plot.top:plot.right-plot.left)/(horizontal?14:42))))),limit=Math.max(2,Math.floor((horizontal?axisLeft-10:band*step)/8));
  const cats=categories.map((name,i)=>i%step?null:<text key={i} x={horizontal?plot.left-6:catAt(i)} y={horizontal?catAt(i):plot.bottom+16} textAnchor={horizontal?'end':'middle'} dominantBaseline={horizontal?'middle':'auto'} className="sheet-chart-axis">{short(name,limit)}</text>);
  let marks:JSX.Element[];
  if(type==='column'||type==='bar'){const inner=band*0.72,bar=inner/series.length;
   marks=series.flatMap((item,s)=>item.values.map((value,i)=>{if(value===null)return null as unknown as JSX.Element;const offset=-inner/2+bar*s,start=Math.min(zero,valueAt(value)),size=Math.abs(valueAt(value)-zero);
    return horizontal?<rect key={s+':'+i} x={start} y={catAt(i)+offset} width={Math.max(size,0.5)} height={Math.max(bar-1,1)} fill={color(s)}><title>{item.name} · {categories[i]}：{value}</title></rect>:<rect key={s+':'+i} x={catAt(i)+offset} y={start} width={Math.max(bar-1,1)} height={Math.max(size,0.5)} fill={color(s)}><title>{item.name} · {categories[i]}：{value}</title></rect>;}).filter(Boolean));
  }else marks=series.map((item,s)=>{
   const points=item.values.map((value,i)=>value===null?null:[catAt(i),valueAt(value)] as const),segments:string[]=[];let current='';
   for(const point of points){if(!point){if(current)segments.push(current);current='';continue;}current+=(current?'L':'M')+point[0]+','+point[1];}if(current)segments.push(current);
   return <g key={s}>{type==='area'&&segments.map((d,k)=>{const xs=[...d.matchAll(/[ML]([\d.-]+),/g)].map(m=>Number(m[1]));return <path key={'a'+k} d={d+`L${xs[xs.length-1]},${zero}L${xs[0]},${zero}Z`} fill={color(s)} opacity={0.35}/>;})}
    {segments.map((d,k)=><path key={k} d={d} fill="none" stroke={color(s)} strokeWidth={2}/>)}{points.map((point,i)=>point&&<circle key={i} cx={point[0]} cy={point[1]} r={3} fill={color(s)}><title>{item.name} · {categories[i]}：{item.values[i]}</title></circle>)}</g>;});
  body=<g>{grid}<line x1={horizontal?zero:plot.left} x2={horizontal?zero:plot.right} y1={horizontal?plot.top:zero} y2={horizontal?plot.bottom:zero} className="sheet-chart-zero"/>{marks}{cats}</g>;
 }
 return <figure className="sheet-chart-figure" aria-labelledby={id}>
  <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} role="img" aria-labelledby={id}><title id={id}>{label}</title>
   {title&&<text x={width/2} y={19} textAnchor="middle" className="sheet-chart-title">{short(title,Math.floor(width/9))}</text>}{body}
   {legend&&!empty&&<g transform={`translate(0,${height-16})`}>{(()=>{const each=Math.max(60,Math.min(140,(width-16)/Math.max(1,legendItems.length)));const shown=legendItems.slice(0,Math.max(1,Math.floor((width-16)/each)));const start=(width-each*shown.length)/2;
    return shown.map((item,i)=><g key={i} transform={`translate(${start+i*each},0)`}><rect width={10} height={10} y={-9} fill={item.color}/><text x={14} y={0} className="sheet-chart-legend">{short(item.name,Math.floor((each-18)/7))}</text></g>);})()}</g>}
  </svg>{table}
 </figure>;
}
