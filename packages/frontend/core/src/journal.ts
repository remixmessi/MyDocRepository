/** Journal dates are local calendar dates, independent of document titles and UTC. */
export type JournalTemplate = 'daily' | 'blank';
export function validJournalDate(value: unknown): value is string {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const [year,month,day] = value.split('-').map(Number);
  return year >= 1 && month >= 1 && month <= 12 && day >= 1 && day <= daysInMonth(year,month);
}
export function daysInMonth(year:number,month:number) { return month===2 ? (year%4===0&&(year%100!==0||year%400===0)?29:28) : [4,6,9,11].includes(month)?30:31; }
export function localJournalDate(date=new Date()) { return `${String(date.getFullYear()).padStart(4,'0')}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`; }
function dateValue(value:string) { if(!validJournalDate(value))throw new Error('请选择有效日期');const [y,m,d]=value.split('-').map(Number),date=new Date(0);date.setFullYear(y,m-1,d);date.setHours(12,0,0,0);return date; }
export function shiftJournalDay(value:string,offset:number) { const date=dateValue(value);date.setDate(date.getDate()+offset);const next=localJournalDate(date);return validJournalDate(next)?next:value; }
export function shiftJournalMonth(value:string,offset:number) { const [y,m,d]=value.split('-').map(Number);const total=y*12+m-1+offset,year=Math.floor(total/12),month=total%12+1;if(year<1||year>9999)return value;return `${String(year).padStart(4,'0')}-${String(month).padStart(2,'0')}-${String(Math.min(d,daysInMonth(year,month))).padStart(2,'0')}`; }
export function journalWeekday(value:string) { return (dateValue(value).getDay()+6)%7; }
export function journalMonthDays(value:string) { const first=value.slice(0,7)+'-01',offset=journalWeekday(first);const [year,month]=first.split('-').map(Number),length=Math.ceil((offset+daysInMonth(year,month))/7)*7;return Array.from({length},(_,i)=>{const date=dateValue(first);date.setDate(date.getDate()-offset+i);const next=localJournalDate(date);return validJournalDate(next)?next:null;}); }
export function journalTitle(date:string) { if(!validJournalDate(date))throw new Error('请选择有效日期');return `${date} 日记`; }
export function journalMarkdown(date:string,template:JournalTemplate) { journalTitle(date);if(template==='blank')return '';return `日期：${date}\n\n## 今日记录\n\n\n## 待办\n\n- [ ] \n\n## 回顾\n\n`; }
