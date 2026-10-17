/* CRP GitHub Pages -> Apps Script iframe connector. Set deployed /exec URL below. */
window.CRP_GAS_URL = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_EXEC_URL_HERE';
(function(){
'use strict';
const pending=new Map();
const channel=Array.from(crypto.getRandomValues(new Uint8Array(24)),v=>v.toString(16).padStart(2,'0')).join('');
let frame=null,ready=false,seq=0;
let bridgeOrigin='';
let readyResolve,readyReject;
const readyPromise=new Promise((res,rej)=>{readyResolve=res;readyReject=rej;});
function start(){
 const url=window.CRP_GAS_URL;
 if(!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(url)){
   readyReject(new Error('กรุณาตั้งค่า CRP_GAS_URL ใน connector.js เป็นลิงก์ Google Apps Script /exec'));return;
 }
 frame=document.createElement('iframe');
 frame.style.cssText='position:absolute;width:0;height:0;border:0;opacity:0;pointer-events:none';
 frame.setAttribute('title','CRP Connector');
 const src=new URL(url);src.searchParams.set('page','bridge');src.searchParams.set('origin',location.origin);src.searchParams.set('channel',channel);
 frame.src=src.href;
 document.body.appendChild(frame);
 const timer=setInterval(()=>{if(ready){clearInterval(timer);return;}if(frame.contentWindow)frame.contentWindow.postMessage({crpBridge:true,channel,type:'hello'},'*');},750);
 setTimeout(()=>{clearInterval(timer);if(!ready)readyReject(new Error('หมดเวลารอ Google Apps Script: ตรวจสอบ /exec, Deploy, CRP_GITHUB_ORIGIN และ Bridge.html'));},20000);
}
window.addEventListener('message',e=>{
 if(!frame||e.source!==frame.contentWindow||!e.data||e.data.crpBridge!==true||e.data.channel!==channel)return;
 // HtmlService is hosted on Google's sandboxed googleusercontent.com iframe.
 if(!/^https:\/\/([a-z0-9-]+\.)?googleusercontent\.com$/.test(e.origin) && e.origin!=='null')return;
 const m=e.data;
 if(m.type==='ready'){ready=true;bridgeOrigin=e.origin;readyResolve();return;}
 if(m.type!=='response')return;
 const entry=pending.get(m.id);if(!entry)return;
 pending.delete(m.id);clearTimeout(entry.timer);
 m.ok?entry.resolve(m.result):entry.reject(new Error(m.error||'คำสั่งล้มเหลว'));
});
window.callServer=async function(action,payload={}){
 await readyPromise;
 const id=channel+'_'+(++seq);
 return new Promise((resolve,reject)=>{
  const timer=setTimeout(()=>{pending.delete(id);reject(new Error('หมดเวลารอคำตอบจาก Apps Script'));},60000);
  pending.set(id,{resolve,reject,timer});
  frame.contentWindow.postMessage({crpBridge:true,channel,type:'request',id,action,payload},bridgeOrigin==='null'?'*':bridgeOrigin);
 });
};
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();
