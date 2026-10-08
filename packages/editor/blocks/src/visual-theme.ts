import {useEffect,useState} from 'react';
export function isDarkTheme(){return typeof document!=='undefined'&&document.documentElement.dataset.theme==='dark';}
export function useVisualTheme(){const [dark,setDark]=useState(isDarkTheme);useEffect(()=>{const refresh=()=>setDark(isDarkTheme());window.addEventListener('mydoc-theme-change',refresh);refresh();return()=>window.removeEventListener('mydoc-theme-change',refresh);},[]);return dark;}
