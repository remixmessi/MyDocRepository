import {extraShapes} from './shape-geometry';
import {lineHeight} from './text-style';
import type {Shape} from './model';
export type EditableShape = Shape & {kind:Exclude<Shape['kind'],'frame'>};
export function isEditableShape(shape:unknown):shape is EditableShape {
  return !!shape && ['text','rectangle','ellipse',...extraShapes].includes((shape as Shape).kind);
}
export function shapeTextLayout(shape:Shape,measure:(text:string)=>number) {
  const centered=(shape.textStyle?.align??(shape.kind==='text'?'left':'center'))==='center',leading=lineHeight(shape);
  const width=Math.max(1,shape.width*(shape.kind==='ellipse'?.7:1)-24);
  const height=Math.max(1,shape.height*(shape.kind==='ellipse'?.65:1)-16);
  const lines:string[]=[];let line='';
  for(const char of shape.text){if(char==='\n'||measure(line+char)>width){lines.push(line);line='';if(char==='\n')continue;}line+=char;}
  lines.push(line);
  const limit=Math.max(0,Math.floor(height/leading));
  return {width,height,leading,lines,limit,centered,x:centered?shape.x+(shape.width-width)/2:shape.x+12,y:centered?shape.y+(shape.height-height)/2:shape.y+12,paddingTop:centered?Math.max(0,(height-Math.min(lines.length,limit)*leading)/2):0};
}
