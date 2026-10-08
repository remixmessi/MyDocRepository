import type {BlockStore} from '@mydoc/editor-framework';
import {LOCAL_ORIGIN} from '@mydoc/crdt';
import {prosemirrorToYXmlFragment} from 'y-prosemirror';
import {documentSchema,migrateDocument,readDocument} from '../document-format';
import type {BoardImage} from './image-model';
export function documentImages(store:BlockStore) {
  const images=new Map<string,string>();readDocument(store).descendants(node=>{if(node.type.name==='image'&&String(node.attrs.src).startsWith('asset:'))images.set(node.attrs.src.slice(6),node.attrs.alt||node.attrs.title||'正文图片');});return [...images].map(([id,name])=>({id,name}));
}
export function appendImageToDocument(store:BlockStore,image:BoardImage) {
  const previous=readDocument(store),node=documentSchema.nodes.paragraph.create(null,documentSchema.nodes.image.create({src:`asset:${image.assetId}`,alt:image.text,title:image.text}));
  const next=previous.copy(previous.content.append(documentSchema.nodes.doc.create(null,[node]).content)),fragment=migrateDocument(store);
  store.history.stopCapturing();store.doc.transact(()=>prosemirrorToYXmlFragment(next,fragment),LOCAL_ORIGIN);store.history.stopCapturing();
}
