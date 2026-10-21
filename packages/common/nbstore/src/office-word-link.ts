/** Shared by import, server validation and the preview; never fetches the URL. */
export function validWordLink(value:unknown):value is string{
 if(typeof value!=='string'||value.length>2048||!/^https?:\/\//i.test(value)||/[\s\u0000-\u001f\u007f\\]/.test(value))return false;
 for(const character of value){const code=character.codePointAt(0)!;if(code>=0xd800&&code<=0xdfff||code===0xfffe||code===0xffff)return false;}
 try{const url=new URL(value);return ['http:','https:'].includes(url.protocol)&&!!url.hostname&&!url.username&&!url.password;}catch{return false;}
}
