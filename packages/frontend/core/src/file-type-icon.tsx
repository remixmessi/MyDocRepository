import {FileText,FileSpreadsheet,Presentation,File} from 'lucide-react';
import type {CloudOfficeResource} from '../../../common/nbstore/src/cloud';
export function FileTypeIcon({resource,size=16}:{resource?:CloudOfficeResource;size?:number}){
 const Icon=resource?.kind==='spreadsheet'?FileSpreadsheet:resource?.kind==='presentation'?Presentation:resource?.kind==='pdf'?File:FileText;
 return <Icon size={size} className={resource?'file-type-office':undefined}/>;
}
