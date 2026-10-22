import type {OfficeKind} from './office';
export const MAX_WORKSPACE_ARCHIVE_BYTES=50*1024*1024;
export interface ArchiveBinary {data:string;sha256:string;}
/** Content snapshot only: accounts, membership, permissions, sessions and local queues are excluded. */
export interface CloudWorkspaceArchive {
 type:'mydoc-team-workspace';version:1|2;exportedAt:number;workspace:{id:string;name:string};
 documents:Array<{id:string;state:ArchiveBinary}>;
 office:Array<{id:string;kind:OfficeKind;name:string;mime:string;sourceFormat:string;originalAssetId:string|null;modelVersion:1;created:number;updated:number;deleted:0|1}>;
 assets:Array<{id:string;name:string;mime:string;doc:string|null;state:ArchiveBinary}>;
 versions?:Array<{id:string;resource:string;name:string;created:number;state:ArchiveBinary}>;
 deleted:string[];
}
export function isCloudWorkspaceArchive(value:unknown):value is CloudWorkspaceArchive{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const v=value as Record<string,unknown>,w=v.workspace as Record<string,unknown>|undefined;
 return v.type==='mydoc-team-workspace'&&(v.version===1||v.version===2&&Array.isArray(v.versions))&&!!w&&typeof w==='object'&&typeof w.id==='string'&&typeof w.name==='string'&&Array.isArray(v.documents)&&Array.isArray(v.office)&&Array.isArray(v.assets)&&Array.isArray(v.deleted);
}
