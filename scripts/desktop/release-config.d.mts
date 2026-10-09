export function updateFeed(value?:string):string|undefined;
export function releaseConfig(base:Record<string,any>,env:NodeJS.ProcessEnv):Record<string,any>;
export function validateRelease(target:string,env:NodeJS.ProcessEnv,platform:string):void;
