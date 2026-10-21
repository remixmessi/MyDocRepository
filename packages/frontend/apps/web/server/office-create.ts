import * as Y from 'yjs';
import type {OfficeKind} from '../../../../common/nbstore/src/office';

export function createOfficeState(kind:OfficeKind){
 const doc=new Y.Doc();
 try{
  if(kind==='text')doc.getText('text');
  if(kind==='word')doc.getArray<Y.Text>('paragraphs').push([new Y.Text()]);
  if(kind==='spreadsheet'){
   const sheet=doc.getArray<Y.Array<string>>('sheet');
   for(let r=0;r<20;r++){const row=new Y.Array<string>();row.insert(0,Array(8).fill(''));sheet.push([row]);}
  }
  if(kind==='pdf'){doc.getArray('annotations');doc.getMap('pdfOperations').set('$schema','0'.repeat(64));}
  if(kind==='presentation'){
   const object=new Y.Map<unknown>();
   for(const [key,value] of Object.entries({type:'text',x:0.1,y:0.1,width:0.8,height:0.3,text:''}))object.set(key,value);
   doc.getArray('objects').push([object]);
  }
  return Y.encodeStateAsUpdate(doc);
 }finally{doc.destroy();}
}

/** A single blank A4 page. Offsets are byte offsets because the body is ASCII. */
export function createBlankPdf(){
 let body='%PDF-1.4\n';const offsets:number[]=[0];
 const objects=['<< /Type /Catalog /Pages 2 0 R >>','<< /Type /Pages /Kids [3 0 R] /Count 1 >>','<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << >> >>'];
 for(let i=0;i<objects.length;i++){offsets.push(Buffer.byteLength(body));body+=`${i+1} 0 obj\n${objects[i]}\nendobj\n`;}
 const start=Buffer.byteLength(body);body+='xref\n0 4\n0000000000 65535 f \n';
 for(const offset of offsets.slice(1))body+=String(offset).padStart(10,'0')+' 00000 n \n';
 return Buffer.from(body+`trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n${start}\n%%EOF\n`);
}
