export function downloadOfficeBytes(name:string,bytes:Uint8Array,mime='application/octet-stream'){
  const url=URL.createObjectURL(new Blob([bytes.slice().buffer],{type:mime}));
  const link=document.createElement('a');link.href=url;link.download=name.replace(/[\\/:*?"<>|]/g,'_');link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
