import {useEffect,useReducer} from 'react';
import type {Y} from '@mydoc/crdt';
import type {LocalWorkspace} from './workspace';
import {observeOfficeDocument} from './office-observer';

export function useOfficeDocument(doc:Y.Doc,workspace:LocalWorkspace){
  const [,refresh]=useReducer(value=>value+1,0);
  useEffect(()=>{
    const stop=observeOfficeDocument(doc,refresh),stopWorkspace=workspace.subscribe(refresh);
    refresh();
    return ()=>{stop();stopWorkspace();};
  },[doc,workspace]);
}
