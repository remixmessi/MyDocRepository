export function isPdfAttachment(name:string,mime:string,bytes?:Uint8Array){return /\.pdf$/i.test(name)||mime.toLowerCase()==='application/pdf'||!!bytes&&new TextDecoder().decode(bytes.subarray(0,5))==='%PDF-';}
export function pdfPageNumber(value:string,count:number){const number=Number(value);return Number.isInteger(number)&&number>=1&&number<=count?number:null;}
export function pdfCanvasSize(width:number,height:number,scale:number,dpr:number){
 const cssWidth=width*scale,cssHeight=height*scale;
 const ratio=Math.min(Math.max(1,dpr),2,4096/Math.max(cssWidth,cssHeight),Math.sqrt(8_000_000/(cssWidth*cssHeight)));
 return {width:Math.max(1,Math.floor(cssWidth*ratio)),height:Math.max(1,Math.floor(cssHeight*ratio)),cssWidth,cssHeight,ratio};
}
