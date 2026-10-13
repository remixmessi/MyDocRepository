const {join}=require('node:path');
const {execFileSync}=require('node:child_process');
module.exports=async context=>{
 if(context.electronPlatformName!=='darwin')return;
 const {verifyMacBundle}=await import('./mac-architecture.mjs');
 // app-builder-lib Arch enum: x64=1, arm64=3, universal=4.
 const expected=context.arch===1?['x64']:context.arch===3?['arm64']:context.arch===4?['x64','arm64']:undefined;
 if(!expected)throw new Error(`不支持的 macOS 打包架构：${context.arch}`);
 const app=join(context.appOutDir,context.packager.appInfo.productFilename+'.app');
 const count=await verifyMacBundle(app,expected);
 console.log(`已验证 macOS ${expected.join('+')}：${count} 个 Mach-O／原生模块架构正确。`);
 const {prepareMacIdentity}=await import('./mac-identity.mjs');
 const executables=await prepareMacIdentity(app,context.packager.appInfo.id,context.packager.appInfo.productFilename);
 // PR builds may skip electron-builder's signing stage; changed UUIDs invalidate the stock signature.
 execFileSync('/usr/bin/codesign',['--force','--deep','--sign','-',app],{stdio:'inherit'});
 console.log(`已为 ${executables} 个 MyDoc 主程序／辅助程序设置独立的 macOS UUID。`);
};
