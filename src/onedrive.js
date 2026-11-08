/* LNStatik – úložisko v OneDrive (Microsoft Graph) namiesto claude.ai.
   Poskytuje rovnaké rozhranie ako claude.use('db'|'user'|'downloads'|'assets'). */
(function(){
const CFG=window.LNS_CONFIG||{};
const SCOPES=['User.Read','Files.ReadWrite'];
const GRAPH='https://graph.microsoft.com/v1.0';
const BASES=['Dokumenty/1-- LNStatik/4_LNStatik/2_Financie','Documents/1-- LNStatik/4_LNStatik/2_Financie','1-- LNStatik/4_LNStatik/2_Financie'];
const DATA_DIR='LNStatik_data',DATA_FILE='lnstatik_data.json',ATT_DIR='Prilohy';
const enc=p=>p.split('/').map(encodeURIComponent).join('/');

let msal=null,account=null;
const AUTH={ok:false,ready:false,err:'',user:''};
window.AUTH=AUTH;

AUTH.loginHTML=function(){
  if(!CFG.clientId)return `<div class="gate"><div class="lg">${(typeof sig!=='undefined'?sig(15,'#8B929B','#1F2328'):'')}</div><p class="dim">Appka ešte nie je prepojená s Microsoft 365 (chýba Client ID v config.js).</p></div>`;
  return `<div class="gate"><div class="lg">${(typeof sig!=='undefined'?sig(15,'#8B929B','#1F2328'):'')}</div>
   <div class="day" style="font-size:26px">Prihlásenie</div><div class="dt" style="text-transform:none">Dáta sú uložené v tvojom OneDrive.</div>
   ${AUTH.err?`<p style="color:var(--bad);font-size:14px">${AUTH.err}</p>`:''}
   <button class="gb" onclick="AUTH.login()"><span>Prihlásiť sa cez Microsoft<small>Konto LN-Statika</small></span><span class="ar">›</span></button></div>`};
AUTH.login=()=>msal.loginRedirect({scopes:SCOPES,prompt:'select_account'});
AUTH.logout=()=>msal.logoutRedirect({account});

async function token(sc){sc=sc||SCOPES;
  try{return (await msal.acquireTokenSilent({scopes:sc,account})).accessToken}
  catch(e){if(sc!==SCOPES){return (await msal.acquireTokenPopup({scopes:sc,account})).accessToken}await msal.acquireTokenRedirect({scopes:sc,account});throw e}}
async function g(path,opt={}){
  const t=await token();
  return fetch(path.startsWith('http')?path:GRAPH+path,{...opt,headers:{Authorization:'Bearer '+t,...(opt.headers||{})}})}

// ---------- dáta ----------
let base=null,state={items:{},u:{},del:{}},etag=null,dirty={},listeners=[],saving=null,timer=null,loaded=false;
const FILE=()=>`/me/drive/root:/${enc(base+'/'+DATA_DIR+'/'+DATA_FILE)}`;
async function findBase(){
  if(CFG.basePath){base=CFG.basePath;return}
  for(const b of BASES){const r=await g(`/me/drive/root:/${enc(b)}?$select=id`);if(r.ok){base=b;return}}
  base='LNStatik';}
function emit(fromCache){const docs=Object.entries(state.items).map(([id,d])=>({id,exists:true,data:()=>JSON.parse(JSON.stringify(d))}));
  listeners.forEach(f=>{try{f({docs,metadata:{fromCache:!!fromCache}})}catch(e){console.error(e)}})}
async function pull(){
  const m=await g(FILE()+'?$select=eTag,id');
  if(m.status===404){etag=null;return {items:{},u:{},del:{}}}
  if(!m.ok)throw new Error('read '+m.status);
  const meta=await m.json();const c=await g(FILE()+':/content');if(!c.ok)throw new Error('read '+c.status);
  const j=await c.json();etag=meta.eTag;return {items:j.items||{},u:j.u||{},del:j.del||{}}}
function merge(remote){
  const out={items:{...remote.items},u:{...remote.u},del:{...remote.del}};
  for(const [id,ts] of Object.entries(dirty)){
    const rts=Math.max(remote.u[id]||0,remote.del[id]||0);if(rts>ts)continue;
    if(state.items[id]!==undefined){out.items[id]=state.items[id];out.u[id]=ts;delete out.del[id]}
    else{delete out.items[id];delete out.u[id];out.del[id]=ts}}
  state=out}
async function push(){
  const body=JSON.stringify({app:'LNStatik',verzia:1,ulozene:new Date().toISOString(),items:state.items,u:state.u,del:state.del});
  const sent={...dirty};
  const r=await g(FILE()+':/content',{method:'PUT',headers:{'Content-Type':'application/json',...(etag?{'If-Match':etag}:{})},body});
  if(r.status===412||r.status===409){merge(await pull());emit(false);return push()}
  if(!r.ok)throw new Error('write '+r.status);
  etag=(await r.json()).eTag;
  for(const [id,ts] of Object.entries(sent))if(dirty[id]===ts)delete dirty[id]}
function schedule(){clearTimeout(timer);setStatus('Ukladám…');timer=setTimeout(flush,800)}
async function flush(){if(saving)return saving;
  saving=(async()=>{try{while(Object.keys(dirty).length)await push();setStatus('Uložené v OneDrive')}
    catch(e){console.error(e);setStatus('Neuložené – skúsim znova');clearTimeout(timer);timer=setTimeout(flush,8000)}
    finally{saving=null}})();return saving}
function setStatus(t){const s=document.getElementById('status');if(s)s.textContent=t}
async function poll(){if(!loaded||document.hidden||Object.keys(dirty).length||saving)return;
  try{const m=await g(FILE()+'?$select=eTag');if(!m.ok)return;const e=(await m.json()).eTag;if(e!==etag){state=await pull();emit(false)}}catch(e){}}
setInterval(poll,20000);document.addEventListener('visibilitychange',()=>{if(!document.hidden)poll()});
window.addEventListener('beforeunload',e=>{if(Object.keys(dirty).length){flush();e.preventDefault();e.returnValue=''}});

const col={doc:id=>({
    set:async d=>{state.items[id]=JSON.parse(JSON.stringify(d));state.u[id]=Date.now();delete state.del[id];dirty[id]=state.u[id];schedule()},
    delete:async()=>{delete state.items[id];delete state.u[id];state.del[id]=Date.now();dirty[id]=state.del[id];schedule()}}),
  onSnapshot:(f,err)=>{listeners.push(f);ready.then(()=>emit(false)).catch(e=>err&&err(e));return()=>{listeners=listeners.filter(x=>x!==f)}}};

// ---------- prílohy ----------
async function attFolder(){return base+'/'+DATA_DIR+'/'+ATT_DIR}
const assets={
  upload:async(file,{type}={})=>{const name=Date.now().toString(36)+'_'+(file.name||'priloha').replace(/[\\/:*?"<>|#%]/g,'_');
    const path=`/me/drive/root:/${enc((await attFolder())+'/'+name)}`;let item;
    if(file.size<=4*1024*1024){const r=await g(path+':/content',{method:'PUT',headers:{'Content-Type':type||file.type||'application/octet-stream'},body:file});if(!r.ok)throw new Error('upload '+r.status);item=await r.json()}
    else{const s=await g(path+':/createUploadSession',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});const {uploadUrl}=await s.json();
      const CH=5*327680;for(let o=0;o<file.size;o+=CH){const part=file.slice(o,o+CH);const r=await fetch(uploadUrl,{method:'PUT',headers:{'Content-Range':`bytes ${o}-${o+part.size-1}/${file.size}`},body:part});if(!r.ok)throw new Error('upload '+r.status);if(r.status===200||r.status===201)item=await r.json()}}
    return {id:item.id,contentType:(item.file&&item.file.mimeType)||type||file.type,sizeBytes:item.size}},
  delete:async id=>{await g(`/me/drive/items/${encodeURIComponent(id)}`,{method:'DELETE'})}};
async function openAtt(id){const w=window.open('about:blank','_blank');
  try{const r=await g(`/me/drive/items/${encodeURIComponent(id)}?$select=@microsoft.graph.downloadUrl,webUrl,file`);const j=await r.json();
    const u=j['@microsoft.graph.downloadUrl']||j.webUrl;if(w)w.location=u;else location.href=u}catch(e){if(w)w.close();alert('Prílohu sa nepodarilo otvoriť.')}}
document.addEventListener('click',e=>{const a=e.target.closest&&e.target.closest('a[href^="/_blob/"]');if(!a)return;e.preventDefault();e.stopPropagation();openAtt(decodeURIComponent(a.getAttribute('href').slice(7)))},true);

// ---------- Outlook koncept ----------
AUTH.mailDraft=async({subject,html,to,pdf,logo})=>{const t=await token(['Mail.ReadWrite']);
  const att=[{'@odata.type':'#microsoft.graph.fileAttachment',name:pdf.name,contentType:'application/pdf',contentBytes:pdf.b64}];
  if(logo)att.push({'@odata.type':'#microsoft.graph.fileAttachment',name:'lnstatik.png',contentType:'image/png',contentBytes:logo,isInline:true,contentId:'lnlogo@lnstatik'});
  const r=await fetch(GRAPH+'/me/messages',{method:'POST',headers:{Authorization:'Bearer '+t,'Content-Type':'application/json'},
    body:JSON.stringify({subject,body:{contentType:'HTML',content:html},toRecipients:String(to||'').split(/[,;]\s*/).filter(Boolean).map(a=>({emailAddress:{address:a}})),attachments:att})});
  if(!r.ok){let m='';try{m=(await r.json()).error.message}catch(e){}throw new Error(m||('Outlook '+r.status))}
  const j=await r.json();return j.webLink||'https://outlook.office.com/mail/drafts'};

// ---------- sťahovanie ----------
const downloads={save:async({filename,data})=>{const b=data instanceof Blob?data:new Blob([data],{type:'application/octet-stream'});
  const u=URL.createObjectURL(b);const a=document.createElement('a');a.href=u;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),30000)}};

// ---------- štart ----------
let readyRes,readyRej;const ready=new Promise((a,b)=>{readyRes=a;readyRej=b});
async function start(){
  if(!CFG.clientId||!window.msal){AUTH.ready=true;return}
  msal=new window.msal.PublicClientApplication({auth:{clientId:CFG.clientId,authority:'https://login.microsoftonline.com/'+(CFG.tenantId||'organizations'),redirectUri:location.origin+location.pathname},cache:{cacheLocation:'localStorage'}});
  await msal.initialize();
  try{const r=await msal.handleRedirectPromise();if(r&&r.account)account=r.account}catch(e){AUTH.err='Prihlásenie zlyhalo: '+(e.errorMessage||e.message||e)}
  account=account||msal.getAllAccounts()[0]||null;
  AUTH.ready=true;
  if(!account){if(window.render)render();return}
  AUTH.ok=true;AUTH.user=account.username||'';if(window.render)render();
  try{await findBase();state=await pull();loaded=true;readyRes()}catch(e){console.error(e);AUTH.err='Nepodarilo sa načítať dáta z OneDrive.';readyRej(e)}
}
const started=start();
window.claude={use:async k=>{await started;if(!AUTH.ok)return null;
  if(k==='db')return {collection:()=>col};
  if(k==='user')return {id:async()=>'me'};
  if(k==='downloads')return downloads;
  if(k==='assets')return assets;return null}};
})();
