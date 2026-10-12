export interface DefaultAdminConfig {email:string;password:string;name:string}

export function defaultAdminConfig(env:NodeJS.ProcessEnv=process.env):DefaultAdminConfig|undefined{
 const email=env.MYDOC_ADMIN_EMAIL?.trim().toLowerCase(),password=env.MYDOC_ADMIN_PASSWORD,name=env.MYDOC_ADMIN_NAME?.trim()||'管理员';
 if(!email&&password===undefined&&env.MYDOC_ADMIN_NAME===undefined)return undefined;
 if(!email||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||email.length>254)throw new Error('MYDOC_ADMIN_EMAIL 必须配置为有效邮箱');
 if(!password||password.length<10||password.length>256||password!==password.trim())throw new Error('MYDOC_ADMIN_PASSWORD 必须为 10 至 256 位，且首尾不能有空白');
 if(name.length>160)throw new Error('MYDOC_ADMIN_NAME 最多 160 个字符');
 return {email,password,name};
}
