import type {PersistentDocument} from '@mydoc/nbstore';
import type {CloudStorage} from '../../../common/nbstore/src/cloud';
import {useEffect,useReducer} from 'react';
export function SaveStatus({persistence,sync}:{persistence:PersistentDocument;sync?:CloudStorage}){
  const [,refresh]=useReducer(n=>n+1,0);
  useEffect(()=>{const off=persistence.subscribe(refresh),offSync=sync?.subscribe(refresh);return()=>{off();offSync?.();};},[persistence,sync]);
  const message=sync?.status==='blocked'?sync.reason:sync?.memoryOnly?'未落盘，请保留草稿':persistence.state==='error'?'保存失败':sync?.status==='offline'?'已断线，等待同步':sync?.status==='connecting'?'正在连接':persistence.state==='saving'?'正在保存…':sync?.pendingCount?'等待提交':sync?.isReadOnly(persistence.id)?'只读':sync?'已同步':'已保存到本机';
  return <span className="save-status" role="status">{message}{(persistence.state==='error'||sync?.status==='offline')&&<button onClick={()=>void (async()=>{await sync?.reconnect();if(persistence.state==='error'&&!sync?.isReadOnly(persistence.id))await persistence.retry();})().catch(()=>{})}>重试保存</button>}</span>;
}
