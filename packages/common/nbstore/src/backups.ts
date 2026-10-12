export interface BackupSettings {enabled:boolean;intervalHours:number;keep:number;directory:string;lastAutomaticAt?:number;lastError?:string;}
export interface DatabaseBackup {id:string;createdAt:number;reason:'automatic'|'manual'|'before-restore'|'before-update';bytes:number;sha256:string;schemaVersion:number;documents:number;assets:number;workspaces:number;}
export interface BackupStatus {settings:BackupSettings;entries:DatabaseBackup[];error?:string;running:boolean;}
