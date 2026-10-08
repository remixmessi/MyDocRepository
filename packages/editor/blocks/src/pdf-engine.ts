import { GlobalWorkerOptions, getDocument } from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
GlobalWorkerOptions.workerSrc=workerUrl;
export function loadPdf(bytes:Uint8Array){
 const base=import.meta.env.BASE_URL+'pdf-assets/';
 return getDocument({data:bytes.slice(),cMapUrl:base+'cmaps/',standardFontDataUrl:base+'standard_fonts/',wasmUrl:base+'wasm/',iccUrl:base+'iccs/',enableXfa:false,stopAtErrors:true});
}
