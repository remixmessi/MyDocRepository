import {existsSync,readFileSync,statSync} from 'node:fs';
import {resolve} from 'node:path';
import {MAX_PDF_FONT_BYTES} from './office-pdf-font';

/** Bundled Noto Sans SC (SIL OFL 1.1); see fonts/README.md. Read by the parent and handed to the parser process. */
export const PDF_FONT_FILE='NotoSansSC-Regular.ttf';
let cached:Uint8Array|null|undefined;
/** Server/desktop bundles copy `fonts/` next to the bundle; development and tests read it from the source tree. */
export function pdfFontCandidates(){
 const here=typeof __dirname==='string'?__dirname:undefined;
 return [process.env.MYDOC_PDF_FONT,here&&resolve(here,'fonts',PDF_FONT_FILE),resolve(process.cwd(),'server/fonts',PDF_FONT_FILE),resolve(process.cwd(),'packages/frontend/apps/web/server/fonts',PDF_FONT_FILE)].filter((path):path is string=>!!path);
}
export function loadPdfFont(){
 if(cached!==undefined)return cached??undefined;
 for(const path of pdfFontCandidates()){try{if(existsSync(path)&&statSync(path).size<=MAX_PDF_FONT_BYTES){cached=new Uint8Array(readFileSync(path));return cached;}}catch{/* try the next location */}}
 cached=null;return undefined;
}
