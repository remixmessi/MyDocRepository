import {net} from 'electron';
import {Readable,Writable} from 'node:stream';
import type {CloudRequestFactory,CloudResponse} from './cloud-proxy';

export const electronCloudRequest:CloudRequestFactory=(url,method,headers,respond)=>{
  const request=net.request({url:url.href,method,headers,redirect:'error',useSessionCookies:false,partition:'mydoc-cloud-transport'});
  const body=new Writable({autoDestroy:false,
    write(chunk,encoding,done){request.write(chunk,encoding,done);},
    final(done){request.end();done();},
    destroy(error,done){request.abort();done(error);},
  });
  request.on('login',(_info,callback)=>callback());
  request.on('error',error=>body.destroy(Object.assign(error,{code:/net::(ERR_[A-Z_]+)/.exec(error.message)?.[1]??'UNKNOWN'})));
  request.on('response',incoming=>{
    const response=new Readable({read(){}}) as CloudResponse;
    response.statusCode=incoming.statusCode;response.headers=incoming.headers;
    incoming.on('end',()=>response.push(null));
    incoming.on('error',error=>response.destroy(error));
    incoming.on('aborted',()=>response.destroy(new Error('upstream response aborted')));
    incoming.on('data',chunk=>response.push(chunk));
    respond(response);
  });
  return {body,cancel:()=>{body.destroy();request.abort();}};
};
