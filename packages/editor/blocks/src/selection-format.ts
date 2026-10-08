import {TextSelection,AllSelection,type EditorState,type Command} from 'prosemirror-state';
import {toggleMark} from 'prosemirror-commands';
import type {MarkType} from 'prosemirror-model';
export const formatMarks=['strong','em','strike','code'] as const;
export function formatSelectionAllowed(state:EditorState){
 if(!(state.selection instanceof TextSelection)&&!(state.selection instanceof AllSelection))return false;if(state.selection.empty&&!state.selection.$from.parent.isTextblock)return false;
 let code=!!state.selection.$from.parent.type.spec.code,text=state.selection.empty;state.doc.nodesBetween(state.selection.from,state.selection.to,node=>{if(node.type.spec.code)code=true;if(node.isText)text=true;});return !code&&text;
}
export function selectionHasMark(state:EditorState,type:MarkType){
 const {from,to,empty,$from}=state.selection;if(empty)return !!type.isInSet(state.storedMarks??$from.marks());let any=false,all=true;
 state.doc.nodesBetween(from,to,node=>{if(node.isText){any=true;if(!type.isInSet(node.marks))all=false;}});return any&&all;
}
export function safeEditorLink(value:string){const href=value.trim();if(!href||href.length>2048||/[\u0000-\u0020\u007f]/.test(href))return null;if(/^https?:\/\//i.test(href)){try{const url=new URL(href);return url.hostname?href:null;}catch{return null;}}return /^mailto:[^@\s]+@[^@\s]+$/i.test(href)?href:null;}
export function formatCommand(name:string):Command{return (state,dispatch,view)=>{if(!formatSelectionAllowed(state))return false;
 if(name==='clear'){let tr=state.tr;for(const type of Object.values(state.schema.marks))tr=state.selection.empty?tr.removeStoredMark(type):tr.removeMark(state.selection.from,state.selection.to,type);dispatch?.(tr);return true;}
 const type=state.schema.marks[name];return !!type&&toggleMark(type,undefined,{removeWhenPresent:false})(state,dispatch,view);
};}
