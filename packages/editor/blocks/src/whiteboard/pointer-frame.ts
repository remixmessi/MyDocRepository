import {useEffect,useRef} from 'react';
export function createFrameQueue<T>(run:(value:T)=>void,request:(callback:()=>void)=>number,cancelRequest:(id:number)=>void){
  let frame:number|undefined,pending:T|undefined;
  const cancel=()=>{if(frame!==undefined)cancelRequest(frame);frame=undefined;pending=undefined;};
  const flush=()=>{const value=pending;cancel();if(value!==undefined)run(value);};
  return {push(value:T){pending=value;if(frame===undefined)frame=request(flush);},flush,cancel};
}
export function usePointerFrame<T>(run:(value:T)=>void){
  const handler=useRef(run);handler.current=run;const queue=useRef<ReturnType<typeof createFrameQueue<T>>>();
  if(!queue.current)queue.current=createFrameQueue<T>(value=>handler.current(value),callback=>requestAnimationFrame(callback),id=>cancelAnimationFrame(id));
  useEffect(()=>()=>queue.current!.cancel(),[]);return queue.current;
}
