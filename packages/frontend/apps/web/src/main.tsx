import React from 'react';
import { createRoot } from 'react-dom/client';
import { App,initializeTheme,CloudLogin } from '@mydoc/frontend-core';
import {initializeProjectStorage} from '@mydoc/nbstore';
async function start(){let hosted=false;try{const response=await fetch('/api/hosting');if(response.ok)hosted=(await response.json()).hosted===true;}catch{}if(hosted){const response=await fetch('/api/cloud/session');const session=await response.json();if(!session.user){createRoot(document.getElementById('root')!).render(<main className="hosting-login"><h1>MyDoc 工作空间</h1><CloudLogin onLogin={()=>location.reload()}/></main>);return;}}await initializeProjectStorage();initializeTheme();
createRoot(document.getElementById('root')!).render(<React.StrictMode><App /></React.StrictMode>);}
void start().catch(error=>{const root=document.getElementById('root')!;root.setAttribute('role','alert');root.textContent='项目数据库启动失败：'+(error instanceof Error?error.message:String(error))+'。请确认存储服务已启动后重试。';const retry=document.createElement('button');retry.textContent='重新打开';retry.onclick=()=>location.reload();root.appendChild(retry);});
