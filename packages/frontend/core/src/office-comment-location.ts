import {Y} from '@mydoc/crdt';
import {captureOfficeTextAnchor,officeAnchorTarget,type OfficeCommentAnchor} from '../../../common/nbstore/src/office-comment-anchor';
import {wordBlockTexts,type WordBlock} from '../../../common/nbstore/src/office-word-table';
import {wordTextIdentity} from './office-word-flow';
import {wordDomPoint} from './word-flow-dom';
export function captureWordComment(root:HTMLElement|null,doc:Y.Doc){
 const selection=window.getSelection();if(!root||!selection?.rangeCount||selection.isCollapsed)return;const range=selection.getRangeAt(0),texts=wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray());
 const point=(node:Node,index:number)=>{const element=(node instanceof HTMLElement?node:node.parentElement)?.closest<HTMLElement>('.word-rich-input');if(!element||!root.contains(element))return;const text=texts.find(text=>wordTextIdentity(text)===element.dataset.wordText);if(!text)return;const before=document.createRange();before.selectNodeContents(element);before.setEnd(node,index);return {text,index:before.toString().length};};
 const a=point(range.startContainer,range.startOffset),b=point(range.endContainer,range.endOffset);if(a&&b)return captureOfficeTextAnchor(a.text,a.index,b.text,b.index,range.toString());
}
export function locateWordComment(root:HTMLElement|null,doc:Y.Doc,anchor:OfficeCommentAnchor){
 const target=officeAnchorTarget(doc,anchor);if(!root||target?.kind!=='text')return;const first=root.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(target.first)}"]`),last=root.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(target.last)}"]`);if(!first||!last)return;first.focus();const range=document.createRange();range.setStart(...wordDomPoint(first,target.start));range.setEnd(...wordDomPoint(last,target.end));const selection=window.getSelection();selection?.removeAllRanges();selection?.addRange(range);first.scrollIntoView({block:'center'});
}
