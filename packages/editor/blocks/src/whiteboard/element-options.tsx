import {memo} from 'react';import {elementLabel,type SurfaceElement} from './model';
// Native selects remain keyboard-accessible; stable snapshots skip rebuilding thousands of options during gestures.
export const ElementOptions=memo(function ElementOptions({elements}:{elements:SurfaceElement[]}){return <>{elements.map((s,i)=><option key={s.id} value={s.id}>{i+1}. {elementLabel(s)}</option>)}</>;});
