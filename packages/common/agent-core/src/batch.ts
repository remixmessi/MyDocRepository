import type {ToolCall,ToolSpec} from '@mydoc/agent-contracts';
/** Contiguous reads share bounded slots. Writes form barriers in declaration order. */
export async function executeBatch(
 calls:readonly ToolCall[],tools:readonly ToolSpec[],concurrency:number,
 execute:(call:ToolCall)=>Promise<void>,signal:AbortSignal,
):Promise<void>{
 const effects=new Map(tools.map(tool=>[tool.name,tool.effect]));
 for(let at=0;at<calls.length;){
  if(signal.aborted)throw signal.reason;
  if(effects.get(calls[at].name)!=='read'){await execute(calls[at++]);continue;}
  let end=at;while(end<calls.length&&effects.get(calls[end].name)==='read')end++;
  let next=at,failure:unknown,failed=false;
  const worker=async()=>{
   while(next<end&&!failed&&!signal.aborted){
    const index=next++;
    try{await execute(calls[index]);}catch(error){failed=true;failure=error;}
   }
  };
  // Settle all workers before reporting failure, so journal cleanup cannot race them.
  await Promise.all(Array.from({length:Math.min(concurrency,end-at)},worker));
  if(failed)throw failure;if(signal.aborted)throw signal.reason;at=end;
 }
}
