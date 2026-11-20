import {MODEL_API_TYPES,type ModelApiType} from '@mydoc/agent-contracts';
import {CompatibleProvider,type CompatibleModelConfig} from './compatible-provider';
import {NativeProvider} from './native-provider';
export function createModelProvider(config:CompatibleModelConfig&{apiType?:ModelApiType},transport?:typeof fetch){
 const type=config.apiType??'openai-chat-completions';
 if(!MODEL_API_TYPES.includes(type))throw new Error('Invalid model API type');
 return type==='openai-chat-completions'?new CompatibleProvider(config,transport):new NativeProvider(config,type,transport);
}
