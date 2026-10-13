import {useCallback,useEffect,useState} from 'react';
import {cloudJson,type CloudAccount} from '../../../common/nbstore/src/cloud';
import {CloudLogin} from './cloud-ui';
import {ReleaseManagement} from './release-management';
import {ClientReleaseCheck} from './client-release-check';

export function WebReleaseSettings({busy,run:perform}:{busy:boolean;run:(work:()=>Promise<void>)=>void}){
 const [user,setUser]=useState<CloudAccount|null>(),[error,setError]=useState('');
 // Publishing always targets the Web service serving this application.
 const json:typeof cloudJson=useCallback((path,method,value)=>cloudJson(path,method,value,''),[]);
 const load=useCallback(async()=>{const result=await json<{user:CloudAccount|null}>('/session');setUser(result.user);},[json]);
 useEffect(()=>{let alive=true;void json<{user:CloudAccount|null}>('/session').then(value=>{if(alive)setUser(value.user);}).catch(e=>{if(alive)setError(e instanceof Error?e.message:'读取管理员权限失败');});return()=>{alive=false;};},[json]);
 const run=(work:()=>Promise<void>)=>perform(async()=>{setError('');try{await work();}catch(e){setError(e instanceof Error?e.message:'操作失败');}});
 return <section>{error?<p role="alert" className="form-error">{error}<button disabled={busy} onClick={()=>run(load)}>重试</button></p>:user===undefined?<p role="status">正在读取本服务器管理权限…</p>:!user?<><h2>版本维护</h2><p>请登录当前 Web 服务器的管理员账号。</p><CloudLogin requestJson={json} onLogin={()=>run(load)}/></>:!user.serverAdmin?<><h2>版本维护</h2><p role="status">已登录 {user.email}，当前账号没有本服务器的版本维护权限。</p></>:<ReleaseManagement server="" accountId={user.id} busy={busy} run={run}/>}</section>;
}

export function DesktopVersionSettings({busy}:{busy:boolean}){
 const [version,setVersion]=useState<string>(),[error,setError]=useState(''),[retry,setRetry]=useState(0);
 useEffect(()=>{let alive=true;setError('');void window.mydocDesktop!.updateStatus().then(state=>{if(alive)setVersion(state.version);}).catch(()=>{if(alive)setError('无法读取当前客户端版本');});return()=>{alive=false;};},[retry]);
 return <section><h2>版本与更新</h2>{version&&<p>当前版本：{version}</p>}{error&&<p role="alert">{error}<button onClick={()=>setRetry(value=>value+1)}>重试读取</button></p>}<ClientReleaseCheck current={version} locked={busy}/></section>;
}
