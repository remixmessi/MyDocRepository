import React from 'react';
import { createRoot } from 'react-dom/client';
import { App,initializeTheme } from '@mydoc/frontend-core';
import {initializeProjectStorage} from '@mydoc/nbstore';
async function start(){await initializeProjectStorage();initializeTheme();
createRoot(document.getElementById('root')!).render(<React.StrictMode><App /></React.StrictMode>);}
void start().catch(error=>{const root=document.getElementById('root')!;root.setAttribute('role','alert');root.textContent='项目数据库启动失败：'+(error instanceof Error?error.message:String(error))+'。请确认存储服务已启动后重试。';const retry=document.createElement('button');retry.textContent='重新打开';retry.onclick=()=>location.reload();root.appendChild(retry);});
