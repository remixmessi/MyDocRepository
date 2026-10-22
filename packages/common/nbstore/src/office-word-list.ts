export type WordList={kind:'bullet'|'decimal';id:string;start:number};
export function validWordList(value:unknown):value is WordList{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;
 const item=value as Record<string,unknown>;
 return Object.keys(item).length===3&&(item.kind==='bullet'||item.kind==='decimal')&&typeof item.id==='string'&&/^[a-zA-Z0-9_-]{1,100}$/.test(item.id)&&typeof item.start==='number'&&Number.isInteger(item.start)&&item.start>=1&&item.start<=1000000;
}
