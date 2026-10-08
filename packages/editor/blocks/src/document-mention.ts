import type {EditorState} from 'prosemirror-state';
export function typedReference(state:EditorState,composing=false){
 const {$from,empty}=state.selection;
 if(!empty||composing||!$from.parent.isTextblock||$from.parent.type.spec.code||$from.marks().some(mark=>mark.type.spec.code))return null;
 const text=$from.parent.textBetween(0,$from.parentOffset,'\n',node=>node.type.name==='hard_break'?'\n':'\ufffc');
 const brackets=/\[\[([^\[\]\n]*)$/.exec(text);
 const mention=/(?:^|[\s([{，。！？：；、])@([^@\n\ufffc]*)$/.exec(text);
 const match=brackets??mention;if(!match)return null;
 const from=$from.pos-match[1].length-(brackets?2:1);
 let code=false;state.doc.nodesBetween(from,$from.pos,node=>{if(node.marks.some(mark=>mark.type.spec.code))code=true;});
 return code?null:{from,to:$from.pos,query:match[1]};
}
