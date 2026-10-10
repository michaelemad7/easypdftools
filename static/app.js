(()=>{'use strict';
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)],MB=30,MAX=MB*1048576;
const SLUGS=['pdf-to-word','pdf-to-excel','pdf-to-jpg','pdf-to-text','word-to-pdf','excel-to-pdf','jpg-to-pdf','merge-pdf','split-pdf','compress-pdf','sign-pdf','ocr-pdf'];
const CFG={'pdf-to-word':{api:'pdf-to-word',acc:'.pdf'},'pdf-to-excel':{api:'pdf-to-excel',acc:'.pdf'},'pdf-to-jpg':{api:'pdf-to-jpg',acc:'.pdf'},'pdf-to-text':{api:'pdf-to-text',acc:'.pdf'},
'word-to-pdf':{api:'word-to-pdf',acc:'.docx'},'excel-to-pdf':{api:'excel-to-pdf',acc:'.xlsx'},'jpg-to-pdf':{api:'jpg-to-pdf',acc:'.jpg,.jpeg,.png',multi:1},'merge-pdf':{api:'merge',acc:'.pdf',multi:1},
'split-pdf':{api:'split',acc:'.pdf',opt:{n:'ranges',k:'text'}},'compress-pdf':{api:'reduce',acc:'.pdf',opt:{n:'level',k:'sel',v:['low','medium','high']}},
'ocr-pdf':{api:'edit',acc:'.pdf',opt:{n:'languages',k:'sel',v:['','ar,en','ar','en','de','fr','es','it']}}};
const OPT={ranges:'Page ranges (optional, e.g. 1-3,5)',level:'Compression level',languages:'Document language'};
const OPL={low:'Light',medium:'Balanced',high:'Maximum','':'Auto-detect','ar,en':'Arabic + English',ar:'Arabic',en:'English',de:'German',fr:'French',es:'Spanish',it:'Italian'};
const L={
en:{dir:'ltr',hero:'Every PDF tool you need in one place',sub:'Fast document conversions with nothing to install and no signup.',search:'Search tools…',pick:'Click or drop a file',note:'Processed in memory, never stored. Max 30 MB.',go:'Convert & download',t:['PDF → Word','PDF → Excel','PDF → JPG','PDF → Text','Word → PDF','Excel → PDF','JPG → PDF','Merge PDF','Split PDF','Compress PDF','Sign PDF','Scan to Word (OCR)']},
ar:{dir:'rtl',hero:'جميع أدوات PDF التي تحتاجها في مكان واحد',sub:'تحويل سريع للمستندات دون تثبيت برامج ودون تسجيل.',search:'ابحث عن أداة…',pick:'اضغط أو أفلت ملفاً هنا',note:'تتم المعالجة في الذاكرة دون حفظ الملفات. الحد الأقصى 30 ميجابايت.',go:'تحويل وتنزيل',t:['PDF إلى Word','PDF إلى Excel','PDF إلى JPG','PDF إلى نص','Word إلى PDF','Excel إلى PDF','JPG إلى PDF','دمج PDF','تقسيم PDF','ضغط PDF','توقيع PDF','مسح إلى Word (OCR)']},
de:{dir:'ltr',hero:'Alle PDF-Tools an einem Ort',sub:'Schnelle Dokumentkonvertierung ohne Installation und Anmeldung.',search:'Tools suchen…',pick:'Klicken oder Datei hierher ziehen',note:'Im Speicher verarbeitet, nie gespeichert. Max. 30 MB.',go:'Umwandeln & herunterladen',t:['PDF in Word','PDF in Excel','PDF in JPG','PDF in Text','Word in PDF','Excel in PDF','JPG in PDF','PDFs zusammenfügen','PDF teilen','PDF komprimieren','PDF signieren','Scan zu Word (OCR)']},
fr:{dir:'ltr',hero:'Tous les outils PDF au même endroit',sub:'Conversions rapides, sans installation ni inscription.',search:'Rechercher un outil…',pick:'Cliquez ou déposez un fichier',note:'Traité en mémoire, jamais conservé. Max 30 Mo.',go:'Convertir et télécharger',t:['PDF en Word','PDF en Excel','PDF en JPG','PDF en texte','Word en PDF','Excel en PDF','JPG en PDF','Fusionner PDF','Diviser PDF','Compresser PDF','Signer PDF','Scan vers Word (OCR)']},
es:{dir:'ltr',hero:'Todas las herramientas PDF en un solo lugar',sub:'Conversiones rápidas, sin instalar nada y sin registro.',search:'Buscar herramientas…',pick:'Haz clic o suelta un archivo',note:'Procesado en memoria, nunca guardado. Máx. 30 MB.',go:'Convertir y descargar',t:['PDF a Word','PDF a Excel','PDF a JPG','PDF a texto','Word a PDF','Excel a PDF','JPG a PDF','Unir PDF','Dividir PDF','Comprimir PDF','Firmar PDF','Escaneo a Word (OCR)']},
it:{dir:'ltr',hero:'Tutti gli strumenti PDF in un unico posto',sub:'Conversioni rapide, senza installare nulla e senza registrazione.',search:'Cerca strumenti…',pick:'Clicca o trascina un file',note:'Elaborato in memoria, mai salvato. Max 30 MB.',go:'Converti e scarica',t:['PDF in Word','PDF in Excel','PDF in JPG','PDF in testo','Word in PDF','Excel in PDF','JPG in PDF','Unisci PDF','Dividi PDF','Comprimi PDF','Firma PDF','Scansione in Word (OCR)']}};
let lang='en',cur=null,files=[];

/* helpers */
const say=(k,t,id='#msg')=>{const m=$(id);if(!m)return;m.className=k?'st-'+k:'';m.textContent=t||''};
const smsg=(k,t)=>say(k,t,'#smsg'),show=d=>{if(d&&!d.open)d.showModal()},mb=n=>(n/1048576).toFixed(1)+' MB';
function btn(t,fn,label){const b=document.createElement('button');b.type='button';b.className='btn';b.textContent=t;if(label)b.setAttribute('aria-label',label);b.onclick=fn;return b}
function save(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),15000)}
function fname(r,fb){const d=r.headers.get('Content-Disposition')||'';let m=d.match(/filename\*=UTF-8''([^;]+)/i);if(m)try{return decodeURIComponent(m[1])}catch{}m=d.match(/filename="([^"]+)"/);return m?m[1]:fb}
async function errText(r){try{const j=await r.json();if(typeof j.detail==='string')return j.detail}catch{}return r.status===413?`File too large (max ${MB} MB).`:r.status===429?'Too many requests. Please wait a minute and try again.':'The server had a problem. Please try again.'}
function dz(sel,fn){const d=$(sel);if(!d)return;['dragenter','dragover'].forEach(t=>d.addEventListener(t,e=>{e.preventDefault();d.classList.add('on')}));['dragleave','drop'].forEach(t=>d.addEventListener(t,e=>{e.preventDefault();d.classList.remove('on');if(t==='drop')fn(e.dataTransfer.files)}))}

/* language, theme, search */
function setLang(l){lang=L[l]?l:'en';const d=L[lang];document.documentElement.lang=lang;document.documentElement.dir=d.dir;$('#lang').value=lang;
$$('[data-i]').forEach(e=>e.textContent=d[e.dataset.i]);$$('[data-t]').forEach(e=>e.textContent=d.t[e.dataset.t]);$('#q').placeholder=d.search;try{localStorage.lang=lang}catch{}}
$('#lang').onchange=e=>setLang(e.target.value);
$('#theme').onclick=()=>{const t=document.documentElement.dataset.theme==='dark'?'light':'dark';document.documentElement.dataset.theme=t;try{localStorage.theme=t}catch{}};
$('#q').oninput=e=>{const q=e.target.value.toLowerCase().trim();let n=0;
$$('#grid a').forEach(a=>{const ok=!q||a.textContent.toLowerCase().includes(q)||a.getAttribute('href').includes(q.replace(/\s+/g,'-'));a.style.display=ok?'':'none';if(ok)n++});
$('#empty').style.display=n?'none':'block'};

/* generic tool dialog */
function opts(c){const el=$('#opts');el.textContent='';const o=c.opt;if(!o)return;const lab=document.createElement('label');lab.className='opt';lab.append(OPT[o.n]);let f;
if(o.k==='text'){f=document.createElement('input');f.placeholder='1-3,5,8-10'}else{f=document.createElement('select');o.v.forEach(v=>f.add(new Option(OPL[v],v)));if(o.n==='level')f.value='medium'}
f.id='opt';lab.append(f);el.append(lab)}
function render(){const ul=$('#files');ul.textContent='';files.forEach((f,i)=>{const li=document.createElement('li'),s=document.createElement('span');s.textContent=f.name;li.append(s);
if(CFG[cur].multi)[['↑',-1,'Move up'],['↓',1,'Move down']].forEach(([t,d,a])=>li.append(btn(t,()=>{const j=i+d;if(j<0||j>=files.length)return;[files[i],files[j]]=[files[j],files[i]];render()},a)));
li.append(btn('✕',()=>{files.splice(i,1);render()},'Remove'));ul.append(li)})}
function add(list){const c=CFG[cur],inc=[...list],arr=c.multi?files.concat(inc):inc.slice(0,1);
if(!arr.every(f=>c.acc.split(',').some(e=>f.name.toLowerCase().endsWith(e))))return say('err','Unsupported file type. Accepted: '+c.acc);
if(arr.reduce((s,f)=>s+f.size,0)>MAX)return say('err',`Too large. Max ${MB} MB${c.multi?' in total':''}.`);files=arr;say();render()}
function openTool(slug){
  cur=slug;files=[];const c=CFG[slug],f=$('#file');
  $('#tool-title').textContent=L[lang].t[SLUGS.indexOf(slug)];
  f.accept=c.acc;f.multiple=!!c.multi;
  $('[data-i="note"]').textContent = slug === 'ocr-pdf' ? L[lang].note + ' (First 2 pages only)' : L[lang].note;
  opts(c);say();render();show($('#tool'))
}
$('#file').onchange=e=>{add(e.target.files);e.target.value=''};dz('#drop',add);
$('#go').onclick=async()=>{const c=CFG[cur];if(!files.length)return say('err','Please choose a file first.');
const fd=new FormData();files.forEach(f=>fd.append(c.multi?'files':'file',f));if(c.opt)fd.append(c.opt.n,$('#opt').value);
const b=$('#go');b.disabled=true;say('wait','Working…');
try{const r=await fetch('/api/'+c.api,{method:'POST',body:fd});if(!r.ok)throw new Error(await errText(r));
save(await r.blob(),fname(r,'converted'));let m='Done! Your download has started.';const a=+r.headers.get('X-Original-Size'),z=+r.headers.get('X-Result-Size');
if(a)m=z<a?`Done! Reduced by ${Math.round(100-z*100/a)}% (${mb(a)} → ${mb(z)}).`:'This file is already well optimised, so it was returned unchanged.';
if(r.headers.get('X-Pages-Total'))m+=` Scanned ${r.headers.get('X-Pages-Processed')} of ${r.headers.get('X-Pages-Total')} pages.`;say('ok',m)}
catch(e){say('err',e instanceof TypeError?'Network error. Check your connection and try again.':e.message)}finally{b.disabled=false}};

/* signer (runs fully in the browser; libraries load only when needed) */
const CDN='https://cdnjs.cloudflare.com/ajax/libs/',S={bytes:null,doc:null,n:1,p:1,marks:[],name:'',task:null,col:'#000000',g:0};
const pad=$('#pad'),px=pad.getContext('2d',{willReadFrequently:true});
const load=src=>new Promise((ok,no)=>{const s=document.createElement('script');s.src=src;s.onload=ok;s.onerror=()=>no(new Error('Could not load the signing library. Check your connection.'));document.head.append(s)});
let libs;
const needLibs=()=>libs||(libs=Promise.all([load(CDN+'pdf.js/3.11.174/pdf.min.js'),load(CDN+'pdf-lib/1.17.1/pdf-lib.min.js')])
  .then(()=>{pdfjsLib.GlobalWorkerOptions.workerSrc=CDN+'pdf.js/3.11.174/pdf.worker.min.js'})
  .catch(e=>{libs=null;throw e}));
let drawing=false,ready=false;
const pt=e=>{const r=pad.getBoundingClientRect();return{x:e.clientX-r.left,y:e.clientY-r.top}};
function initPad(){
  const r=pad.getBoundingClientRect();
  if(ready||!r.width)return; /* dialog must be visible so the canvas has a size */
  const d=devicePixelRatio||1;
  pad.width=r.width*d;pad.height=r.height*d;
  px.setTransform(d,0,0,d,0,0);
  px.lineCap=px.lineJoin='round';px.lineWidth=3;px.strokeStyle=S.col;ready=true}
const clearPad=()=>{px.save();px.setTransform(1,0,0,1,0,0);px.clearRect(0,0,pad.width,pad.height);px.restore()};
pad.onpointerdown=e=>{initPad();pad.setPointerCapture(e.pointerId);drawing=true;const p=pt(e);px.beginPath();px.moveTo(p.x,p.y);px.lineTo(p.x+.1,p.y);px.stroke()};
pad.onpointermove=e=>{if(!drawing)return;const p=pt(e);px.lineTo(p.x,p.y);px.stroke()};
pad.onpointerup=pad.onpointercancel=()=>{drawing=false};
$$('.dot').forEach(b=>b.onclick=()=>{S.col=b.dataset.c;px.strokeStyle=S.col;$$('.dot').forEach(x=>x.setAttribute('aria-pressed',x===b))});
$('#sclr').onclick=clearPad;
$('#supl').onclick=()=>$('#simg').click();
$('#simg').onchange=e=>{const f=e.target.files[0];e.target.value='';if(!f)return;initPad();const im=new Image();
im.onload=()=>{clearPad();const w=pad.clientWidth,h=pad.clientHeight,s=Math.min((w-16)/im.width,(h-16)/im.height);
px.drawImage(im,(w-im.width*s)/2,(h-im.height*s)/2,im.width*s,im.height*s);
const id=px.getImageData(0,0,pad.width,pad.height),d=id.data;
for(let i=0;i<d.length;i+=4)if(d[i]>235&&d[i+1]>235&&d[i+2]>235)d[i+3]=0; /* make white background transparent */
px.putImageData(id,0,0);URL.revokeObjectURL(im.src)};
im.onerror=()=>{URL.revokeObjectURL(im.src);smsg('err','Could not read this image.')};
im.src=URL.createObjectURL(f)};
function crop(){const w=pad.width,h=pad.height;if(!w||!h)return null;const d=px.getImageData(0,0,w,h).data;let x0=w,y0=h,x1=-1,y1=-1;
for(let y=0;y<h;y++)for(let x=0;x<w;x++)if(d[(y*w+x)*4+3]>8){if(x<x0)x0=x;if(x>x1)x1=x;if(y<y0)y0=y;if(y>y1)y1=y}
if(x1<0)return null;const cw=x1-x0+1,ch=y1-y0+1,c=document.createElement('canvas');c.width=cw;c.height=ch;c.getContext('2d').drawImage(pad,x0,y0,cw,ch,0,0,cw,ch);return{url:c.toDataURL('image/png'),ar:ch/cw}}
function stamps(){const o=$('#ov');o.textContent='';S.marks.filter(m=>m.p===S.p).forEach(m=>{const i=new Image();i.src=m.url;i.alt='';i.style.left=m.x*100+'%';i.style.top=m.y*100+'%';o.append(i)})}
async function draw(){try{S.task&&S.task.cancel()}catch{}
const pg=await S.doc.getPage(S.p),v=pg.getViewport({scale:1.4}),c=$('#pdfc');c.width=v.width;c.height=v.height;$('#pn').textContent=`${S.p} / ${S.n}`;
S.task=pg.render({canvasContext:c.getContext('2d'),viewport:v});
try{await S.task.promise}catch(e){if(e.name!=='RenderingCancelledException')throw e}stamps()}
async function loadPdf(f){if(!f)return;if(f.size>MAX)return smsg('err',`File too large (max ${MB} MB).`);smsg('wait','Loading…');
try{await needLibs();S.bytes=await f.arrayBuffer();S.name=f.name;S.doc=await pdfjsLib.getDocument({data:new Uint8Array(S.bytes.slice(0))}).promise;
S.n=S.doc.numPages;S.p=1;S.marks=[];$('#viewer').hidden=false;smsg();await draw()}
catch(e){smsg('err',e.name==='PasswordException'?'This PDF is password-protected.':(e.message||'Could not open this PDF.'))}}
$('#sfile').onchange=e=>{const f=e.target.files[0];e.target.value='';loadPdf(f)};dz('#sdrop',fs=>loadPdf(fs[0]));
$('#prev').onclick=()=>{if(S.p>1){S.p--;draw()}};
$('#next').onclick=()=>{if(S.p<S.n){S.p++;draw()}};
$('#undo').onclick=()=>{if(!S.marks.length)return;const g=Math.max(...S.marks.map(m=>m.g));S.marks=S.marks.filter(m=>m.g!==g);stamps()};
$('#pdfc').onclick=e=>{const s=crop();if(!s)return smsg('err','Draw or upload your signature first.');smsg();
const r=e.target.getBoundingClientRect(),x=(e.clientX-r.left)/r.width,y=(e.clientY-r.top)/r.height,all=$('#all').checked,g=++S.g;
for(let p=all?1:S.p;p<=(all?S.n:S.p);p++)S.marks.push({p,x,y,g,...s});stamps()};
$('#sdl').onclick=async()=>{if(!S.marks.length)return smsg('err','Place at least one signature first.');const b=$('#sdl');b.disabled=true;smsg('wait','Generating…');
try{const doc=await PDFLib.PDFDocument.load(S.bytes.slice(0)),cache=new Map();
for(const m of S.marks){let im=cache.get(m.url);if(!im){im=await doc.embedPng(m.url);cache.set(m.url,im)}
const pg=doc.getPage(m.p-1),{width:W,height:H}=pg.getSize(),w=W*.22,h=w*m.ar;
pg.drawImage(im,{x:m.x*W-w/2,y:H-m.y*H-h/2,width:w,height:h})}
save(new Blob([await doc.save()],{type:'application/pdf'}),'signed_'+S.name);smsg('ok','Done! Your download has started.')}
catch(e){smsg('err',e.message||'Signing failed.')}finally{b.disabled=false}};

/* routing: every tool has its own URL, e.g. /merge-pdf */
function resolveCoreTool(slug){
  if(CFG[slug]||slug==='sign-pdf')return slug;
  const meta=$('meta[name="core-tool"]');
  if(meta&&meta.content&&(CFG[meta.content]||meta.content==='sign-pdf'))return meta.content;
  for(const k of SLUGS)if(slug.startsWith(k)||slug.includes(k))return k;
  return null}
function openSlug(slug,push){
  const core=resolveCoreTool(slug);if(!core)return;
  if(core==='sign-pdf'){show($('#sign'));initPad()}else openTool(core);
  if(push)history.pushState({},'','/'+slug)}
function route(){
  $$('dialog[open]').forEach(d=>{d.dataset.skip=1;d.close()});
  const p=location.pathname.slice(1);if(resolveCoreTool(p))openSlug(p,false)}
$$('#grid a').forEach(a=>a.onclick=e=>{if(e.metaKey||e.ctrlKey||e.shiftKey||e.button)return;e.preventDefault();openSlug(a.getAttribute('href').slice(1),true)});
$$('[data-info]').forEach(b=>b.onclick=()=>show($('#'+b.dataset.info)));
$$('[data-close]').forEach(b=>b.onclick=()=>b.closest('dialog').close());
$$('dialog').forEach(d=>{
  d.addEventListener('click',e=>{const r=d.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)d.close()});
  d.addEventListener('close',()=>{if(d.dataset.skip){delete d.dataset.skip;return}if((d.id==='tool'||d.id==='sign')&&location.pathname!=='/')history.pushState({},'','/')})});
addEventListener('popstate',route);
let want;try{want=localStorage.lang}catch{}
setLang(want||(navigator.language||'en').slice(0,2));route();

/* ad-blocker notice: a soft, dismissible banner; the tools stay usable */
function checkAdBlocker(){
  const bait=document.createElement('div');
  bait.innerHTML='&nbsp;';
  bait.className='adsbox ad-placement doubleclick';
  bait.style.cssText='position:absolute;top:-999px;left:-999px';
  bait.setAttribute('aria-hidden','true');
  document.body.appendChild(bait);
  setTimeout(()=>{
    const blocked=bait.offsetHeight===0||getComputedStyle(bait).display==='none';
    bait.remove();
    if(blocked)showAdNotice()},100)}
function showAdNotice(){
  const main=document.querySelector('main');
  if(!main||$('#adnote'))return;
  const b=document.createElement('div');
  b.id='adnote';
  b.style.cssText='display:block;background:#e11d4822;color:var(--err);border:2px solid var(--err);border-radius:16px;padding:20px;margin:0 auto 24px;max-width:640px;text-align:center';
  b.innerHTML=`<h2 style="margin:0 0 10px">Help us keep this free</h2>
    <p style="color:var(--fg);margin:0 0 14px">EasyPDFTools is free, but processing PDF and Excel files needs servers, and ads help cover that. If you can, please whitelist our site.</p>`;
  const x=document.createElement('button');x.type='button';x.className='btn pri';x.style.maxWidth='250px';x.textContent='Got it';x.onclick=()=>b.remove();
  b.append(x);main.prepend(b)}
addEventListener('load',checkAdBlocker);

})();