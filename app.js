/* Shared storage client. No admin password or database secret is shipped here. */
'use strict';
const $ = s => document.querySelector(s);
const uid = () => crypto.randomUUID();
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const cfg = window.DAYBOOK_CONFIG || {};
let workspaceTimeZone = cfg.timeZone || 'Pacific/Auckland';
try{new Intl.DateTimeFormat('en',{timeZone:workspaceTimeZone}).format();}catch{workspaceTimeZone='Pacific/Auckland';}
const iso = d => {
  if(d)return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  const parts=new Intl.DateTimeFormat('en',{timeZone:workspaceTimeZone,year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());
  return ['year','month','day'].map(k=>parts.find(p=>p.type===k).value).join('-');
};
let data = {name:'daybook', labels:['Must be done by today','If you have time, do this','User notes'], nav:['Today','This month','Settings'], users:[], templates:[], recurring:{labels:['Opening up','Closing up','Front of house responsibilities','Have a go coach responsibilities'],items:[]}, days:{}};
const DAILY_LABELS=['Opening up','Closing up','Front of house responsibilities','Have a go coach responsibilities'],OLD_DAILY_LABELS=['Daily section 1','Daily section 2','Daily section 3','Daily section 4'];
function migrateDailyLabels(l){return DAILY_LABELS.map((n,i)=>!l?.[i]||l[i]===OLD_DAILY_LABELS[i]?n:l[i]);}
function ensureRecurring(s){s.recurring||={labels:DAILY_LABELS.slice(),items:[]};s.recurring.items||=[];s.recurring.labels=migrateDailyLabels(s.recurring.labels);return s.recurring;}
let view = 'today', selected = iso(), month = new Date(selected+'T12:00:00');month.setDate(1);
let admin = false, pending = null, actualToday = iso(), revision = -1, token = '', busy = false, loaded = false;
let connection = 'Connecting…', lastFailedDraft = null, lastChecked = null;
const configured = typeof google !== 'undefined' && !!google.script?.run;

function toast(message){ $('.toast')?.remove(); const el=document.createElement('div'); el.className='toast'; el.setAttribute('role','status'); el.textContent=message; document.body.append(el); setTimeout(()=>el.remove(),6500); }
async function rpc(name,args={}){
  if(!configured) throw new Error('Open this dashboard through the deployed Google Apps Script web app. See README.md.');
  return new Promise((resolve,reject)=>{
    let settled=false;
    const timer=setTimeout(()=>{settled=true;reject(new Error('Connection timed out. The save may have reached Drive. Refresh before retrying.'));},45000);
    google.script.run.withSuccessHandler(response=>{
      if(settled)return;settled=true;clearTimeout(timer);
      if(!response||!response.ok){const error=new Error(response?.error?.message||'Google Drive request failed');error.code=response?.error?.code;reject(error);return;}
      resolve(response.result);
    }).withFailureHandler(failure=>{
      if(settled)return;settled=true;clearTimeout(timer);reject(new Error(failure?.message||'Cannot reach Google Drive. This save was not confirmed.'));
    }).daybookRequest(name,args);
  });
}

function accept(result){
  if(!result || !result.state || !Number.isSafeInteger(result.revision)) throw new Error('Drive setup is incomplete. Run setupDaybook_ in the Apps Script editor.');
  if(result.revision<revision)return;
  data=result.state; revision=result.revision; loaded=true; connection='Google Drive · Connected'; lastChecked=new Date();
}
function captureInputs(){return [...document.querySelectorAll('#app input,#app textarea,#app select')].map(el=>({name:el.name,id:el.id,type:el.type,value:el.value})).filter(el=>el.type!=='password'&&el.type!=='file');}
function restoreInputs(items){for(const item of items){const el=[...document.querySelectorAll('#app input,#app textarea,#app select')].find(el=>item.name?el.name===item.name:item.id?el.id===item.id:el.type===item.type);if(el)el.value=item.value;}const picker=$('#taskform [name="template"]'),button=$('#savetemplatebutton');if(picker&&button)button.textContent=picker.value?'Update selected template':'Save as new template';}
function setBusy(value){busy=value;document.body.classList.toggle('busy',value);document.querySelectorAll('#app button,#app input,#app select,#app textarea,dialog button,dialog input,dialog select').forEach(el=>el.disabled=value);}
function safeToRefresh(){return !busy&&!document.querySelector('dialog[open]')&&!['INPUT','SELECT','TEXTAREA'].includes(document.activeElement?.tagName)&&![...document.querySelectorAll('#app form input,#app form textarea')].some(el=>el.type!=='file'&&el.value!==el.defaultValue);}
async function refresh(force=false){
  if(busy || (!force&&!safeToRefresh())) return;
  if(force&&lastFailedDraft&&!confirm('Refresh from the Google Drive? Your failed admin draft remains available to download.'))return;
  const inputs=loaded?captureInputs():[];
  try {const result=await rpc('daybook_read');const changed=result.revision!==revision;accept(result);if(changed||force||!$('#app .content')&&view==='today'){render();restoreInputs(inputs);}else updateConnection();}
  catch(error){connection='Google Drive · Disconnected';updateConnection();if(force||!loaded)showError(error.message);}
}
function updateConnection(){const el=$('.saved');if(el)el.textContent=connection;}
function showError(message){if(!loaded){$('#app').innerHTML=`<div class="panel"><div class="panelhead"><h2>Shared storage setup</h2></div><div class="form"><p class="sub">${esc(message)}</p><p class="sub">This version saves to a Google Drive. Complete the included Google Apps Script setup, then reload.</p><button class="primary" onclick="location.reload()">Retry connection</button><span class="sub">The setup guide is included in your GitHub package.</span></div></div>`;}else toast(message);}
function render(){
  document.title=data.name+' · Task Manager';$('#brand').textContent=data.name;
  $('#nav').innerHTML=['today','month','settings'].map((v,i)=>`<button class="${view===v?'active':''}" onclick="navigate('${v}')" aria-label="${esc(data.nav[i])}" ${view===v?'aria-current="page"':''}>${icon(v)}${esc(data.nav[i])}</button>`).join('');
  $('#app').innerHTML=`<div class="topbar"><span>WORKSPACE &nbsp; / &nbsp; ${view==='settings'?'ADMINISTRATION':view==='month'?'CALENDAR':'DAILY OVERVIEW'}</span><span class="saved" role="status">${esc(connection)}</span></div>`+
    (view==='settings'?renderAdmin():view==='month'?renderMonth():renderDay())+
    `<footer>Shared records are saved to Google Drive after confirmation. Other visitors are checked for updates every 30 seconds. <button class="smallbtn" onclick="refresh(true)">Refresh records</button>${lastChecked?' · Checked '+esc(timestamp(lastChecked)):''}</footer>`;
  if(view==='settings'&&lastFailedDraft){const block=document.createElement('div');block.className='tip';block.innerHTML='<strong>An admin edit was not saved.</strong>Your draft is available to download. Review the latest shared records before reapplying it. <button class="secondary" onclick="downloadFailedDraft()">Download unsaved draft</button>';$('#app').append(block);}
}
function lockAdmin(){const oldToken=token;token='';admin=false;if(oldToken)rpc('daybook_logout',{p_token:oldToken}).catch(()=>{});}
function navigate(v){if(busy||!loaded)return;if(v==='settings'){if(!admin){$('#password').value='';$('#loginerror').textContent='';$('#login').showModal();return}view='settings';}else{lockAdmin();view=v;if(v==='today')selected=iso();}render();}
async function mutateAdmin(change,message='Saved to Google Drive'){
  if(!admin||busy||!loaded)return false;
  const next=structuredClone(data), expectedRevision=revision, inputs=captureInputs();
  if(change(next)===false)return false;
  setBusy(true);
  try {const result=await rpc('daybook_write',{p_state:next,p_revision:expectedRevision,p_token:token});accept(result);lastFailedDraft=null;render();toast(message);return true;}
  catch(error){lastFailedDraft=next;
    if(error.code==='42501'){lockAdmin();view='today';}
    if(error.code==='40001'){try{accept(await rpc('daybook_read'));}catch{}}
    render();restoreInputs(inputs);toast(error.message);return false;
  }finally{setBusy(false);}
}
async function saveCompletion(task,done,userId=null){
  if(busy)return false;setBusy(true);
  try{const rec=!!pending?.recurring,args={p_date:pending?.date||selected,p_user_id:userId,p_done:done,p_expected:task.completed||null};args[rec?'p_item_id':'p_task_id']=task.id;const result=await rpc(rec?'daybook_recurring_complete':'daybook_complete',args);accept(result);$('#complete').close();pending=null;render();toast(done?'Completion recorded in Google Drive':'Task reopened in Google Drive');return true;}
  catch(error){if(error.code==='40001'){try{accept(await rpc('daybook_read'));render();}catch{}}toast(error.message);return false;}
  finally{setBusy(false);}
}
function toggleTask(id){
  if(busy||!loaded)return;const t=day().tasks.find(t=>t.id===id);if(!t)return;
  if(t.completed){if(confirm('Mark this task as unfinished? The previous completion remains in its activity history.')){pending={id,date:selected};saveCompletion(t,false);}return;}
  if(!data.users.length){toast('Add a user in Settings before completing a task.');return;}
  pending={id,date:selected};$('#completetitle').textContent=t.title;$('#completionuser').innerHTML=userOptions();$('#complete').showModal();
}
$('#completeform').onsubmit=async e=>{e.preventDefault();if(busy||!pending)return;const u=data.users.find(u=>u.id===$('#completionuser').value),t=pending.recurring?recurringTask(pending.date,pending.id):day(pending.date).tasks.find(t=>t.id===pending.id);if(!u||!t){toast('Refresh and choose a current task and user.');return;}await saveCompletion(t,true,u.id);};
$('#loginform').onsubmit=async e=>{
  e.preventDefault();if(busy)return;setBusy(true);$('#loginerror').textContent='';
  try{const result=await rpc('daybook_login',{p_password:$('#password').value});$('#password').value='';if(!result.ok){$('#loginerror').textContent=result.message||'Access denied';return;}token=result.token;admin=true;view='settings';accept(await rpc('daybook_read'));$('#login').close();render();}
  catch(error){$('#loginerror').textContent=error.message;}
  finally{setBusy(false);}
};
document.addEventListener('submit',async e=>{
  const id=e.target.id;if(!['noteform','taskform','userform','namesform','passwordform','dailyform'].includes(id))return;e.preventDefault();if(busy||!loaded)return;const f=new FormData(e.target);
  if(id==='noteform'){
    const text=String(f.get('note')).trim();if(!text)return;const date=selected;
    // Retain an id across retries so an uncertain network response cannot duplicate a note.
    let noteId=e.target.dataset.noteId;if(!noteId){noteId=uid();e.target.dataset.noteId=noteId;}
    setBusy(true);
    try{accept(await rpc('daybook_note',{p_date:date,p_text:text,p_user_id:f.get('author')||null,p_note_id:noteId}));render();toast('Note saved to Google Drive');}
    catch(error){toast(error.message);}
    finally{setBusy(false);}return;
  }
  if(!admin)return;
  if(id==='passwordform'){
    if(f.get('next')!==f.get('confirm'))return toast('New passwords do not match');
    if(new TextEncoder().encode(String(f.get('next'))).length>72)return toast('New password must be at most 72 bytes');
    setBusy(true);
    try{await rpc('daybook_change_password',{p_token:token,p_current:f.get('current'),p_next:f.get('next')});lockAdmin();view='today';render();toast('Password changed. Sign in with the new password.');}
    catch(error){toast(error.message);}finally{setBusy(false);}return;
  }
  await mutateAdmin(next=>{
    if(id==='taskform'){const title=String(f.get('title')).trim();if(!title)return false;const date=f.get('date'),u=next.users.find(u=>u.id===f.get('assignee'));next.days[date]||={tasks:[],notes:[]};next.days[date].tasks.push({id:uid(),title,section:Number(f.get('section')),assignee:u?{id:u.id,name:u.name}:null,completed:null,history:[]});selected=date;}
    if(id==='userform'){const name=String(f.get('name')).trim();if(!name)return false;if(next.users.some(u=>u.name.toLowerCase()===name.toLowerCase())){toast('That user already exists');return false;}next.users.push({id:uid(),name});}
    if(id==='dailyform'){const title=String(f.get('title')).trim();if(!title)return false;const r=ensureRecurring(next);if(r.items.length>=200){toast('Daily item limit reached (200). Remove an unused item first.');return false;}r.items.push({id:uid(),title,section:Number(f.get('section')),from:iso()});}
    if(id==='namesform'){if([...f.values()].some(v=>!String(v).trim())){toast('Names cannot be blank');return false;}next.name=f.get('brand').trim();next.labels=next.labels.map((_,i)=>f.get('label'+i).trim());next.nav=next.nav.map((_,i)=>f.get('nav'+i).trim());ensureRecurring(next).labels=[0,1,2,3].map(i=>f.get('daily'+i).trim());}
  });
});
document.addEventListener('input',e=>{if(e.target.closest('#noteform'))delete e.target.closest('#noteform').dataset.noteId;});
function renameUser(id){if(!admin||busy)return;const u=data.users.find(u=>u.id===id),name=prompt('User name',u.name)?.trim();if(!name)return;if(name.length>80)return toast('Use at most 80 characters');if(data.users.some(v=>v.id!==id&&v.name.toLowerCase()===name.toLowerCase()))return toast('That user already exists');mutateAdmin(next=>{next.users.find(u=>u.id===id).name=name;});}
function removeUser(id){if(!admin||busy)return;if(confirm('Remove this user? Existing completion records will keep their name.'))mutateAdmin(next=>{next.users=next.users.filter(u=>u.id!==id);});}
function removeTask(id){if(!admin||busy)return;const date=selected;if(confirm('Delete this task from this day? Export a backup first if you need to keep it.'))mutateAdmin(next=>{next.days[date].tasks=next.days[date].tasks.filter(t=>t.id!==id);});}
function editTask(id){if(!admin||busy)return;const date=selected,t=day().tasks.find(t=>t.id===id),name=prompt('Task name',t.title)?.trim();if(name){if(name.length>200)return toast('Use at most 200 characters');mutateAdmin(next=>{next.days[date].tasks.find(t=>t.id===id).title=name;});}}
function applyTemplate(id,focus=false){
  if(!admin||busy)return;
  const form=$('#taskform');if(!form)return;
  const field=name=>form.querySelector('[name="'+name+'"]');
  const template=(data.templates||[]).find(t=>t.id===id);
  if(id&&!template){toast('This template was removed. Refresh the records.');return;}
  field('template').value=id||'';field('title').value=template?.title||'';field('section').value=String(template?.section??0);
  const assignee=template?.assignee&&data.users.find(u=>u.id===template.assignee.id);
  field('assignee').value=assignee?.id||'';
  $('#savetemplatebutton').textContent=template?'Update selected template':'Save as new template';
  if(focus){form.scrollIntoView({behavior:'smooth',block:'start'});field('date').focus();}
  if(template?.assignee&&!assignee)toast('Template loaded. Its previous assignee is no longer active; choose a person if needed.');
}
async function storeTemplate(task,existingId=''){
  if(!admin||busy)return false;
  const title=String(task.title||'').trim();
  if(!title||title.length>200||![0,1].includes(task.section)){toast('Enter a task name (up to 200 characters) and select a section.');return false;}
  const assigned=data.users.find(u=>u.id===task.assignee?.id);
  const template={id:existingId||uid(),title,section:task.section,assignee:assigned?{id:assigned.id,name:assigned.name}:null};
  const duplicate=(data.templates||[]).find(t=>t.id!==existingId&&t.title.toLowerCase()===title.toLowerCase()&&t.section===template.section&&(t.assignee?.id||'')===(template.assignee?.id||''));
  if(duplicate){toast('This task is already saved as a template. Select it from Task template.');return false;}
  const saved=await mutateAdmin(next=>{
    next.templates||=[];
    if(existingId){const index=next.templates.findIndex(t=>t.id===existingId);if(index<0){toast('This template was removed. Refresh the records.');return false;}next.templates[index]=template;}
    else{if(next.templates.length>=500){toast('Template limit reached (500). Remove an unused template first.');return false;}next.templates.push(template);}
  },existingId?'Task template updated in Google Drive':'Task template saved to Google Drive');
  return saved?template.id:false;
}
async function saveTaskTemplate(){
  if(!admin||busy)return;
  const form=$('#taskform');if(!form)return;const f=new FormData(form),date=f.get('date');
  const id=await storeTemplate({title:f.get('title'),section:Number(f.get('section')),assignee:{id:f.get('assignee')}},f.get('template')||'');
  if(id){applyTemplate(id);$('#taskform [name="date"]').value=date||selected;}
}
async function saveScheduledTemplate(id){
  if(!admin||busy)return;const task=day().tasks.find(t=>t.id===id);if(task)await storeTemplate(task);
}
async function removeTemplate(id){
  if(!admin||busy)return;if(confirm('Delete this reusable template? Tasks already scheduled from it will remain unchanged.'))await mutateAdmin(next=>{next.templates=(next.templates||[]).filter(t=>t.id!==id);},'Task template deleted');
}
function toggleRecurring(id){
  if(busy||!loaded)return;const t=recurringTask(selected,id);if(!t)return;
  if(t.completed){if(confirm('Mark this item as unfinished? The previous completion remains in its activity history.')){pending={id,date:selected,recurring:true};saveCompletion(t,false);}return;}
  if(!data.users.length){toast('Add a user in Settings before completing a task.');return;}
  pending={id,date:selected,recurring:true};$('#completetitle').textContent=t.title;$('#completionuser').innerHTML=userOptions();$('#complete').showModal();
}
function editDaily(id){if(!admin||busy)return;const it=recurringState().items.find(i=>i.id===id);if(!it)return;const name=prompt('Daily item name',it.title)?.trim();if(!name)return;if(name.length>200)return toast('Use at most 200 characters');mutateAdmin(next=>{const x=ensureRecurring(next).items.find(i=>i.id===id);if(!x)return false;x.title=name;});}
function removeDaily(id){if(!admin||busy)return;if(confirm('Delete this daily item? It will disappear from every day, including past records. Export a backup first if you need to keep its history.'))mutateAdmin(next=>{const r=ensureRecurring(next);r.items=r.items.filter(i=>i.id!==id);},'Daily item deleted');}
function downloadJSON(value,filename){const url=URL.createObjectURL(new Blob([JSON.stringify(value,null,2)],{type:'application/json'})),a=document.createElement('a');a.href=url;a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
async function exportData(){if(!admin||busy)return;setBusy(true);try{accept(await rpc('daybook_read'));downloadJSON({format:'daybook-shared-v1',exportedAt:new Date().toISOString(),state:data},'daybook-backup-'+iso()+'.json');toast('Latest shared records exported');}catch(error){toast(error.message);}finally{setBusy(false);}}
function downloadFailedDraft(){if(lastFailedDraft&&admin)downloadJSON({format:'daybook-shared-v1',state:lastFailedDraft},'daybook-unsaved-draft-'+iso()+'.json');}
function normalizeBackup(value){
  const d=structuredClone(value.state||value);delete d.password;
  const validId=id=>typeof id==='string'&&/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id);
  const text=(v,max)=>typeof v==='string'&&v.trim().length>0&&v.length<=max;
  if(!text(d.name,40)||!Array.isArray(d.labels)||d.labels.length!==3||!d.labels.every(v=>text(v,80))||!Array.isArray(d.nav)||d.nav.length!==3||!d.nav.every(v=>text(v,40))||!Array.isArray(d.users)||d.users.length>500||!d.users.every(u=>validId(u.id)&&text(u.name,80))||!d.days||typeof d.days!=='object'||Array.isArray(d.days))throw new Error('Invalid Daybook backup');
  const ids=new Set();for(const u of d.users){if(ids.has(u.id))throw new Error('Duplicate user id');ids.add(u.id);}const recordIds=new Set();
  if(d.templates===undefined)d.templates=[];
  if(!Array.isArray(d.templates)||d.templates.length>500)throw new Error('Invalid task templates');
  const templateIds=new Set();for(const t of d.templates){if(!t||!validId(t.id)||templateIds.has(t.id.toLowerCase())||!text(t.title,200)||![0,1].includes(t.section)||Object.hasOwn(t,'completed')||Object.hasOwn(t,'history')||(t.assignee&&(!validId(t.assignee.id)||!text(t.assignee.name,80))))throw new Error('Invalid task template');templateIds.add(t.id.toLowerCase());}
  for(const [date,day] of Object.entries(d.days)){
    if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||Number.isNaN(Date.parse(date))||new Date(date).toISOString().slice(0,10)!==date||!Array.isArray(day.tasks)||!Array.isArray(day.notes)||day.tasks.length>1000||day.notes.length>5000)throw new Error('Invalid daily record');
    for(const t of day.tasks){if(!validId(t.id)||recordIds.has(t.id)||!text(t.title,200)||![0,1].includes(t.section))throw new Error('Invalid task');recordIds.add(t.id);t.history||=[];if(!Array.isArray(t.history))throw new Error('Invalid activity history');if(t.completed&&(!text(t.completed.name,80)||!validId(t.completed.userId)||!Number.isFinite(Date.parse(t.completed.time))))throw new Error('Invalid completion');if(t.assignee&&(!validId(t.assignee.id)||!text(t.assignee.name,80)))throw new Error('Invalid assignee');}
    for(const n of day.notes){n.id ||= uid();n.author ||= '';if(!validId(n.id)||recordIds.has(n.id)||!text(n.text,10000)||typeof n.author!=='string'||!Number.isFinite(Date.parse(n.time)))throw new Error('Invalid note');recordIds.add(n.id);}
  }
  if(JSON.stringify(d).length>5000000)throw new Error('Backup exceeds workspace size limit');return d;
}
document.addEventListener('change',async e=>{
  if(e.target.id!=='importfile'||!admin||busy)return;
  try{const file=e.target.files[0];if(!file)return;if(file.size>6000000)throw new Error('Backup file is too large');const d=normalizeBackup(JSON.parse(await file.text()));if(confirm('Replace ALL shared records for EVERY visitor with this backup? Passwords are not imported. Export a backup first.'))await mutateAdmin(next=>{for(const k of Object.keys(next))delete next[k];Object.assign(next,d);},'Backup restored to Google Drive');}
  catch(error){toast(error.message||'Invalid backup');}
});
setInterval(()=>{if(iso()!==actualToday){const previous=actualToday;actualToday=iso();if(view==='today'&&selected===previous){if(!safeToRefresh()){toast('A new day has started. Select Today when you finish your draft.');return;}selected=actualToday;if(loaded)render();}}},30000);
setInterval(()=>{if(configured&&document.visibilityState==='visible')refresh();},30000);
document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'&&configured)refresh();});
window.addEventListener('online',()=>refresh());
if(configured)refresh(true);else showError('This dashboard must run inside the deployed Google Apps Script web app.');
