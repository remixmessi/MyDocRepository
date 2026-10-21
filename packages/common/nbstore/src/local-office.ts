/** Routes exposed through the local SQLite boundary and desktop IPC. */
export const LOCAL_OFFICE_METHODS:Record<string,string[]>={
 '/office-workspace-full-restore':['POST'],
 '/office-pdf-forms':['GET'],'/office-pdf-form-operation':['POST'],'/office-pdf-native-annotations':['GET','POST'],
 '/office-version-revert':['POST'],
 '/office-comment-operation':['POST'],
 '/office-workspace-backup':['GET'],'/office-workspace-restore':['POST'],
 '/office-purge':['DELETE'],
 '/office-resources':['GET','POST'],'/office-resource':['PATCH','DELETE'],
 '/office-import':['POST'],'/office-text-import':['POST'],'/office-original':['GET'],
 '/office-export':['GET'],'/office-pdf-export':['GET'],'/office-exports':['GET','POST','DELETE'],
 '/office-sheet-operation':['POST'],'/office-word-table-operation':['POST'],'/office-presentation-operation':['POST'],'/office-pdf-operation':['POST'],
 '/office-backup':['GET'],'/office-restore':['POST'],
 '/office-versions':['GET','POST'],'/office-version':['GET','DELETE'],'/office-version-restore':['POST'],
 '/office-deleted-resources':['GET'],'/access':['GET'],'/document':['GET','POST'],'/asset':['GET','PUT','DELETE'],
};
