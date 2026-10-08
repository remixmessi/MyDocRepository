export function GroupNavigation({depth,canEnter,enter,back,root}:{depth:number;canEnter:boolean;enter:()=>void;back:()=>void;root:()=>void}){
  return <div className="whiteboard-selection-tools" role="toolbar" aria-label="分组层级">
    <span role="status">{depth?`组内编辑 · 第 ${depth} 层`:'画布 · 最外层'}</span>
    <button disabled={!canEnter} onClick={enter}>进入分组</button>
    <button disabled={!depth} onClick={back}>返回上级</button>
    {depth>1&&<button onClick={root}>返回画布</button>}
    {depth>0&&<span>仅选择和操作当前组内成员；Escape 返回上级。</span>}
  </div>;
}
