import {useState} from 'react';
import {projectPreferences} from '@mydoc/nbstore';
import {cloudServerUrl,configureCloudServer} from '../../../common/nbstore/src/cloud';
export interface ServerConnection {url:string;name:string}
export function normalizeServer(value:string){const url=new URL(value.trim());if(!['http:','https:'].includes(url.protocol)||url.username||url.password||url.search||url.hash)throw new Error('请输入有效的 http 或 https 服务器地址');return url.origin+url.pathname.replace(/\/+$/,'');}
export function savedServers():ServerConnection[]{
 let list:ServerConnection[]=[];try{const data=JSON.parse(projectPreferences.getItem('cloud-servers')??'[]');if(Array.isArray(data))for(const item of data){if(typeof item?.url==='string'&&typeof item?.name==='string'){const url=normalizeServer(item.url);if(!list.some(server=>server.url===url))list.push({url,name:item.name});}}}catch{}
 const legacy=projectPreferences.getItem('cloud-server-url');if(legacy){try{const url=normalizeServer(legacy);if(!list.some(server=>server.url===url))list.push({url,name:new URL(url).host});}catch{}}
 return list;
}
export function ServerConnections({prepare,onServerChange,busy=false}:{prepare:()=>Promise<void>;onServerChange?:()=>Promise<void>|void;busy?:boolean}){
 const [servers,setServers]=useState(savedServers),[url,setUrl]=useState(''),[name,setName]=useState(''),[saving,setSaving]=useState(false),[error,setError]=useState('');
 const locked=busy||saving,current=cloudServerUrl();
 const act=async(work:()=>Promise<void>)=>{setSaving(true);setError('');try{await work();}catch(e){setError(e instanceof Error?e.message:'服务器配置保存失败');}finally{setSaving(false);}};
 const connect=async(next:string)=>{await prepare();configureCloudServer(next);if(next&&!servers.some(server=>server.url===next)){const entry={url:next,name:new URL(next).host};const updated=[...servers,entry];await projectPreferences.setItem('cloud-servers',JSON.stringify(updated));setServers(updated);}await projectPreferences.setItem('cloud-server-url',next);await projectPreferences.flush();await onServerChange?.();};
 return <section className="server-connections" aria-label="服务器连接"><h3>服务器连接</h3><p className="organization-help">仅显示你添加的远程服务器，连接后查看对应账号和团队空间。</p>
 {error&&<p role="alert" className="form-error">{error}</p>}
 {!servers.length&&<p className="organization-help">尚未添加远程服务器。添加后即可连接团队工作空间。</p>}
 <ul className="cloud-members">{servers.map(server=><li key={server.url}><div><strong>{server.name}</strong><small>{server.url}</small></div><button disabled={locked} onClick={()=>void act(()=>connect(current===server.url?'':server.url))}>{current===server.url?'断开':'连接'}</button><button disabled={locked||current===server.url} onClick={()=>void act(async()=>{const next=servers.filter(value=>value.url!==server.url);await projectPreferences.setItem('cloud-servers',JSON.stringify(next));setServers(next);})}>移除<span className="sr-only"> {server.name}</span></button></li>)}</ul>
 <form className="settings-server-form" onSubmit={event=>{event.preventDefault();void act(async()=>{const address=normalizeServer(url),next=[...servers.filter(server=>server.url!==address),{url:address,name:name.trim()||new URL(address).host}];await projectPreferences.setItem('cloud-servers',JSON.stringify(next));setServers(next);setUrl('');setName('');});}}><label className="form-field">服务器名称<input value={name} onChange={event=>setName(event.target.value)} maxLength={100} disabled={locked} placeholder="例如：公司服务器"/></label><label className="form-field">服务器地址<input type="url" required value={url} onChange={event=>setUrl(event.target.value)} placeholder="https://docs.example.com" disabled={locked}/></label><button disabled={locked}>添加服务器</button></form>
 </section>;
}
