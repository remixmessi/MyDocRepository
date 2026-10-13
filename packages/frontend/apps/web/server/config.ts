import {existsSync,readFileSync} from 'node:fs';
import {dirname,resolve} from 'node:path';
import {defaultAdminConfig,type DefaultAdminConfig} from './default-admin';

export interface ServerConfig {
 host?:string;
 port?:number;
 cloudDbPath?:string;
 admin?:{email?:string;password?:string;name?:string};
}

function projectRoot(start:string){
 let current=resolve(start);
 for(;;){if(existsSync(resolve(current,'data/server-config.json'))||existsSync(resolve(current,'data/server-config.example.json')))return current;const parent=dirname(current);if(parent===current)return resolve(start);current=parent;}
}
function readConfig(path:string):ServerConfig{
 if(!existsSync(path))return {};
 let parsed:unknown;
 try{parsed=JSON.parse(readFileSync(path,'utf8'));}catch(error){throw new Error(`服务器配置文件无法读取：${path}（${error instanceof Error?error.message:'JSON 无效'}）`);}
 if(!parsed||typeof parsed!=='object'||Array.isArray(parsed))throw new Error(`服务器配置文件必须是 JSON 对象：${path}`);
 return parsed as ServerConfig;
}

export function serverConfig(env:NodeJS.ProcessEnv=process.env){
 const root=projectRoot(process.cwd()),path=resolve(env.MYDOC_CONFIG_PATH??resolve(root,'data/server-config.json'));
 const file=readConfig(path);
 const admin=file.admin??{};
 // Container environment variables are the deployment override. The config file
 // remains useful as a local fallback, especially when no env values are set.
 const mergedEnv={...env,
  MYDOC_ADMIN_EMAIL:env.MYDOC_ADMIN_EMAIL??admin.email,
  MYDOC_ADMIN_PASSWORD:env.MYDOC_ADMIN_PASSWORD??admin.password,
  MYDOC_ADMIN_NAME:env.MYDOC_ADMIN_NAME??admin.name,
 };
 const configuredAdmin=defaultAdminConfig(mergedEnv);
 const portValue=file.port??(env.MYDOC_PORT===undefined?8080:Number(env.MYDOC_PORT));
 if(!Number.isInteger(portValue)||portValue<1||portValue>65535)throw new Error('服务器配置中的 port 必须是 1 至 65535 的整数');
 return {path,host:file.host??env.MYDOC_HOST??'127.0.0.1',port:portValue,
  cloudDbPath:resolve(root,file.cloudDbPath??env.MYDOC_CLOUD_DB_PATH??'data/cloud.sqlite'),admin:configuredAdmin as DefaultAdminConfig|undefined};
}
