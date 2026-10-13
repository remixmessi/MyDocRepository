import {useEffect,useState} from 'react';
import {sqliteFetch} from '../../../common/nbstore/src/desktop';
import {projectPreferences} from '@mydoc/nbstore';
import {cloudJson} from '../../../common/nbstore/src/cloud';
type LocalAccount={id:string;email:string;name:string};
let sessionToken='';
export let localAccount:LocalAccount|null=null;
export async function restoreLocalAccount(){sessionToken=projectPreferences.getItem('local-account-session')??'';const result=await request({action:'session',token:sessionToken});localAccount=result.user;return localAccount;}
async function request(value:unknown){const response=await sqliteFetch('/local-account',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(value)});const result=await response.json();if(!response.ok)throw new Error(result.error??'本地账号操作失败');return result as {user:LocalAccount|null;token?:string};}
export async function registerLocalOnServer(password:string){
 if(!localAccount)throw new Error('请先在设置中登录本地账号');
 const checked=await request({action:'login',email:localAccount.email,password});
 if(checked.token)await request({action:'logout',token:checked.token});
 await cloudJson('/register','POST',{email:localAccount.email,name:localAccount.name,password});
}
export function LocalAccountPanel(){
 const [user,setUser]=useState(localAccount),[register,setRegister]=useState(false),[busy,setBusy]=useState(false),[error,setError]=useState(''),[loading,setLoading]=useState(true);
 useEffect(()=>{let alive=true;void restoreLocalAccount().then(user=>{if(alive)setUser(user);}).catch(e=>{if(alive)setError(e.message);}).finally(()=>{if(alive)setLoading(false);});return()=>{alive=false;};},[]);
 const run=async(work:()=>Promise<void>)=>{setBusy(true);setError('');try{await work();}catch(e){setError(e instanceof Error?e.message:'操作失败');}finally{setBusy(false);}};
 return <section className="local-account-panel"><h2>本地账号</h2><p className="organization-help">账号保存在本地 SQLite 中，可用于注册团队服务器账号。</p>{error&&<p role="alert" className="form-error">{error}</p>}
   {loading?<p role="status">正在读取本地账号…</p>:user?<><h3>{user.name}</h3><p>{user.email}</p><button disabled={busy} onClick={()=>void run(async()=>{await request({action:'logout',token:sessionToken});await projectPreferences.setItem('local-account-session','');sessionToken='';localAccount=null;setUser(null);})}>退出本地账号</button></>:<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);void run(async()=>{const result=await request({action:register?'register':'login',email:data.get('email'),name:data.get('name'),password:data.get('password')});sessionToken=result.token!;await projectPreferences.setItem('local-account-session',sessionToken);localAccount=result.user;setUser(result.user);});}}>
    {register&&<label className="form-field">显示名称<input name="name" required maxLength={160} disabled={busy}/></label>}
    <label className="form-field">邮箱<input type="email" name="email" required autoComplete="email" disabled={busy}/></label><label className="form-field">密码<input type="password" name="password" required minLength={10} maxLength={256} autoComplete={register?'new-password':'current-password'} disabled={busy}/></label>
    <div className="dialog-actions"><button type="button" disabled={busy} onClick={()=>setRegister(!register)}>{register?'已有本地账号，登录':'创建本地账号'}</button><button className="primary" disabled={busy}>{register?'创建并登录':'登录本地账号'}</button></div>
   </form>}
  </section>;
}
