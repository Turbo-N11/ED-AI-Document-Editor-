const $ = s => document.querySelector(s);
const messages = $('#messages');
const input = $('#messageInput');
const toast = $('#toast');
const dropzone = $('#dropzone');
const fileInput = $('#fileInput');

let documents = [];
let activeId = null;
let uploadBusy = false;

function applyTheme(theme){
  const root=document.documentElement;
  root.dataset.theme=theme;
  const dark=theme==='dark';
  const icon=$('#themeIcon');
  const label=$('#themeLabel');
  if(icon) icon.textContent=dark?'☀':'☾';
  if(label) label.textContent=dark?'Light mode':'Dark mode';
  const meta=$('#themeColorMeta');
  if(meta) meta.content=dark?'#090808':'#f6ead6';
  localStorage.setItem('ed-theme',theme);
}

(function initTheme(){
  const saved=localStorage.getItem('ed-theme');
  const theme=saved==='dark' || saved==='light' ? saved : 'light';
  applyTheme(theme);
})();

$('#themeToggle').onclick=()=>{
  applyTheme(document.documentElement.dataset.theme==='dark'?'light':'dark');
};

function showToast(text){
  toast.textContent = text;
  toast.classList.add('show');
  clearTimeout(window._toast);
  window._toast = setTimeout(() => toast.classList.remove('show'), 3200);
}
function setStatus(text){ $('#statusText').textContent = text; }
function addMessage(role,text){
  const wrap=document.createElement('div');
  wrap.className=`message ${role}`;
  const bubble=document.createElement('div');
  bubble.className='bubble';
  bubble.textContent=text;
  wrap.appendChild(bubble);
  messages.appendChild(wrap);
  messages.scrollTop=messages.scrollHeight;
  return wrap;
}
function addTyping(){
  const wrap=document.createElement('div');
  wrap.className='message assistant typing';
  wrap.innerHTML='<div class="bubble"><span class="typing-dots"><span></span><span></span><span></span></span></div>';
  messages.appendChild(wrap);
  messages.scrollTop=messages.scrollHeight;
  return wrap;
}

function renderDocuments(list){
  documents=list || [];
  const box=$('#documentList');
  $('#fileCount').textContent=documents.length;
  box.innerHTML='';
  if(!documents.length){
    box.innerHTML='<div class="list-empty">No documents loaded yet.</div>';
    return;
  }
  documents.forEach(doc=>{
    const button=document.createElement('button');
    button.type='button';
    button.className='document-item'+(doc.active?' active':'');
    const ext=(doc.name.split('.').pop()||'DOCX').toUpperCase();
    button.innerHTML=`<span class="file-badge">${ext==='DOCX'?'W':ext.slice(0,3)}</span><span class="file-copy"><strong></strong><small></small></span>${doc.active?'<span class="active-check">●</span>':''}`;
    button.querySelector('strong').textContent=doc.name;
    const st=doc.stats||{};
    button.querySelector('small').textContent=`${st.paragraphs||0} paragraphs · ${st.tables||0} tables`;
    button.onclick=()=>selectDocument(doc.id);
    box.appendChild(button);
  });
}

function updateView(data){
  if(data.documents) renderDocuments(data.documents);
  if(data.active) activeId=data.active;
  const stats=data.stats;
  if(stats){
    $('#workspaceTitle').textContent=stats.name||'Your documents';
    $('#activeMeta').textContent=`${stats.paragraphs||0} paragraphs · ${stats.tables||0} tables · ${stats.sections||0} sections`;
    $('#previewStats').textContent=`${stats.paragraphs||0} paragraphs · ${stats.tables||0} tables · ${stats.images||0} images`;
  }else{
    $('#workspaceTitle').textContent='Your documents';
    $('#activeMeta').textContent='No active document';
    $('#previewStats').textContent='—';
  }
  if(data.preview_url){
    $('#documentEmpty').hidden=true;
    $('#pdfViewer').hidden=false;
    const frame=$('#documentFrame');
    const url=data.preview_url;
    frame.title=`LibreOffice preview — ${stats?.name||'Document preview'}`;
    if(frame.src !== new URL(url,window.location.href).href) frame.src=url;
  }else if(data.loaded===false || !stats){
    $('#documentEmpty').hidden=false;
    $('#pdfViewer').hidden=true;
    $('#documentFrame').src='about:blank';
  }
}

function resetWelcome(){
  messages.innerHTML=`<div class="welcome"><div class="welcome-mark">ED</div><h2>Work across your documents</h2><p>Upload multiple files, select a file, or add <code>--all</code> to an editing command to apply it everywhere.</p><div class="suggestions"><button>What is this document about?</button><button>Make all headings bold and dark blue</button><button>Improve the overall formatting and spacing</button><button>change "Load" to "Road" --all</button></div></div>`;
  messages.querySelectorAll('.suggestions button').forEach(b=>b.onclick=()=>{input.value=b.textContent;input.focus()});
}

async function uploadOne(file){
  const fd=new FormData();
  fd.append('files',file,file.name);
  const r=await fetch('/api/upload',{method:'POST',body:fd,credentials:'same-origin'});
  const type=r.headers.get('content-type')||'';
  const d=type.includes('application/json') ? await r.json() : {error:await r.text()};
  if(!r.ok) throw new Error(d.error||`Upload failed (${r.status})`);
  return d;
}

async function uploadFiles(fileList){
  if(uploadBusy)return;
  const files=[...fileList];
  if(!files.length)return;
  const allowed=['.docx','.pdf','.png','.jpg','.jpeg','.webp','.bmp'];
  const valid=files.filter(f=>allowed.includes(f.name.slice(f.name.lastIndexOf('.')).toLowerCase()));
  const skipped=files.filter(f=>!valid.includes(f));
  if(!valid.length){showToast('Supported: DOCX, PDF and image files.');return}

  uploadBusy=true;
  dropzone.classList.add('loading');
  let loaded=0, auxiliary=0, errors=[];
  try{
    for(let i=0;i<valid.length;i++){
      const file=valid[i];
      setStatus(`Uploading ${i+1} of ${valid.length}: ${file.name}`);
      try{
        const d=await uploadOne(file);
        loaded += (d.added||[]).length;
        auxiliary += (d.auxiliary||[]).length;
        (d.errors||[]).forEach(e=>errors.push(e));
        // Every response contains the complete workspace state. Refresh after
        // each file so one bad/large file cannot prevent the others appearing.
        updateView(d);
      }catch(err){
        errors.push(`${file.name}: ${err.message}`);
      }
    }
    skipped.forEach(f=>errors.push(`${f.name}: unsupported file type`));

    const state=await fetch('/api/state',{credentials:'same-origin'}).then(async r=>{
      const d=await r.json();
      if(!r.ok) throw new Error(d.error||'Could not refresh workspace');
      return d;
    });
    activeId=state.active;
    updateView(state);

    if(loaded || auxiliary){
      resetWelcome();
      let msg=`${loaded} document${loaded===1?'':'s'} loaded`;
      if(auxiliary) msg += ` · ${auxiliary} auxiliary file${auxiliary===1?'':'s'}`;
      if(errors.length) msg += ` · ${errors.length} skipped/failed`;
      addMessage('assistant',msg+'. Select a document on the left or use --all for batch edits.');
      showToast(errors.length ? `${loaded} loaded · ${errors.length} issue${errors.length===1?'':'s'}` : `${loaded} document${loaded===1?'':'s'} loaded`);
    }else{
      showToast(errors[0]||'No supported documents were uploaded.');
    }
    if(errors.length) console.warn('Upload issues:',errors);
    setStatus('Workspace ready');
  }finally{
    uploadBusy=false;
    dropzone.classList.remove('loading');
    input.focus();
  }
}

async function selectDocument(id){
  if(id===activeId)return;
  setStatus('Opening…');
  try{
    const r=await fetch('/api/select',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({id})});
    const d=await r.json();
    if(!r.ok)throw new Error(d.error||'Could not open document');
    activeId=id; updateView(d); setStatus('Ready');
  }catch(e){showToast(e.message);setStatus('Ready');}
}

$('#chooseBtn').onclick=(e)=>{e.stopPropagation();fileInput.click()};
fileInput.onchange=e=>{const files=[...e.target.files];e.target.value='';uploadFiles(files)};
dropzone.ondragover=e=>{e.preventDefault();e.dataTransfer.dropEffect='copy';dropzone.classList.add('drag')};
dropzone.ondragleave=e=>{if(!dropzone.contains(e.relatedTarget))dropzone.classList.remove('drag')};
dropzone.ondrop=e=>{e.preventDefault();dropzone.classList.remove('drag');uploadFiles(e.dataTransfer.files)};

document.addEventListener('dragover',e=>e.preventDefault());

async function sendMessage(text){
  addMessage('user',text); input.value=''; setStatus('Working…'); $('#chatForm').classList.add('loading');
  const typing=addTyping();
  try{
    const r=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({message:text})});
    const d=await r.json();
    if(!r.ok)throw new Error(d.error||'Request failed');
    typing.remove(); addMessage('assistant',d.message||'Completed.'); updateView(d); addFileLinks(d.files);
    setStatus(d.changed?'Changes applied':'Ready'); if(d.changed)showToast('Document updated');
  }catch(err){typing.remove();addMessage('assistant',err.message);setStatus('Ready');}
  finally{$('#chatForm').classList.remove('loading');input.focus();}
}
$('#chatForm').onsubmit=e=>{e.preventDefault();const text=input.value.trim();if(text)sendMessage(text)};

document.querySelectorAll('.suggestions button').forEach(b=>b.onclick=()=>{input.value=b.textContent;input.focus()});
function addFileLinks(files){
  if(!files||!files.length)return;
  const box=document.createElement('div');box.className='command-files';
  files.slice(0,10).forEach(f=>{const a=document.createElement('a');a.href=f.url;a.textContent=`Download ${f.name}`;a.className='download-link';a.target='_blank';box.appendChild(a)});
  messages.appendChild(box);messages.scrollTop=messages.scrollHeight;
}
async function action(name){
  setStatus('Working…');const typing=addTyping();
  try{
    const r=await fetch('/api/action',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({action:name})});
    const d=await r.json();if(!r.ok)throw new Error(d.error||'Action failed');
    typing.remove();addMessage('assistant',d.message||'Action completed.');updateView(d);setStatus(d.changed?'Changes applied':'Ready');
  }catch(e){typing.remove();addMessage('assistant',e.message);setStatus('Ready');}
}
document.querySelectorAll('.side-action').forEach(b=>b.onclick=()=>action(b.dataset.action));
$('#resetBtn').onclick=()=>action('reset');
$('#doneBtn').onclick=async()=>{
  setStatus('Preparing export…');
  try{
    const r=await fetch('/api/export',{credentials:'same-origin'});
    if(!r.ok){const d=await r.json().catch(()=>({}));throw new Error(d.error||'Nothing to export');}
    const blob=await r.blob(),url=URL.createObjectURL(blob),a=document.createElement('a');
    a.href=url;a.download=documents.length>1?'edited-documents.zip':(documents[0]?.name||'edited-document.docx');document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1000);setStatus('Export ready');showToast('Export downloaded');
  }catch(e){showToast(e.message);setStatus('Ready');}
};
$('#allHint').onclick=()=>{input.value='change "Load" to "Road" --all';input.focus();input.setSelectionRange(input.value.length,input.value.length)};
input.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();$('#chatForm').requestSubmit()}});

fetch('/api/state',{credentials:'same-origin'}).then(async r=>{const d=await r.json();if(!r.ok)throw new Error(d.error||'Could not load workspace');activeId=d.active;updateView(d)}).catch(err=>console.warn(err));
