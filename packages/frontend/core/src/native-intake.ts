let owner:symbol|undefined;
export function acquireIntake(token:symbol){if(owner&&owner!==token)return false;owner=token;return true;}
export function releaseIntake(token:symbol){if(owner!==token)return;owner=undefined;window.dispatchEvent(new Event('mydoc:intake-ready'));}
