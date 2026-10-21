import {createContext} from 'react';
import type {Y} from '@mydoc/crdt';
import type {WordFontKey,WordMark} from './office-word-format';
/** `formatFlow` applies to a cross-paragraph selection; `painter` is the format painter's copied character format. */
export const WordEditingContext=createContext<{flowActive?:boolean;formatFlow?:(key:WordMark|WordFontKey|'clear',value?:string|number)=>boolean;active:Y.Text|null;activate:(text:Y.Text)=>void;toolbar:HTMLElement|null;focus:(text:Y.Text,offset?:number)=>void;painter?:Record<string,unknown>;setPainter?:(format:Record<string,unknown>|undefined)=>void}>({active:null,activate:()=>{},toolbar:null,focus:()=>{}});
