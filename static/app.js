const $=s=>document.querySelector(s), types={phone:'电话',email:'邮箱',social:'社交账号',address:'地址'};
let contacts=[], favorites=false, editing=null, token=sessionStorage.getItem('token')||'';
function message(text,error=false){$('#message').textContent=text;$('#message').style.color=error?'#b42318':'#366b54';}
async function api(path,options={}){const res=await fetch('/api/'+path,{...options,headers:{Authorization:'Bearer '+token,...options.headers}});if(!res.ok){let err=await res.json();throw Error(err.error||'请求失败');}return res;}
async function refresh(){try{contacts=await(await api('contacts')).json();render();}catch(e){message(e.message,true);$('#count').textContent='读取失败，请检查访问密码或网络';}}
function el(tag,text,className){const n=document.createElement(tag);if(text!==undefined)n.textContent=text;if(className)n.className=className;return n;}
function render(){const query=$('#search').value.trim().toLowerCase();const list=contacts.filter(c=>(!favorites||c.favorite)&&JSON.stringify(c).toLowerCase().includes(query)).sort((a,b)=>Number(b.favorite)-Number(a.favorite)||a.name.localeCompare(b.name,'zh'));
$('#count').textContent=`共 ${contacts.length} 位联系人 · 当前显示 ${list.length} 位`;$('#cards').replaceChildren();$('#empty').hidden=!!list.length;
for(const c of list){let card=el('article',undefined,'card'),top=el('div',undefined,'card-top'),avatar=el('div',c.name.slice(0,1),'avatar'),title=el('h2',c.name),star=el('button',c.favorite?'★':'☆','star');star.setAttribute('aria-label',c.favorite?'取消收藏':'收藏联系人');star.onclick=()=>mutate('contacts/'+c.id,'PUT',{...c,favorite:!c.favorite});top.append(avatar,title,star);card.append(top);
for(const m of c.methods){let line=el('p',undefined,'method');line.append(el('small',m.platform||types[m.type]),el('span',m.value));card.append(line);}if(c.notes)card.append(el('p',c.notes,'note'));let actions=el('div',undefined,'actions');let edit=el('button','编辑','secondary'),del=el('button','删除','danger');edit.onclick=()=>openEditor(c);del.onclick=()=>{if(confirm(`确定删除「${c.name}」？`))mutate('contacts/'+c.id,'DELETE');};actions.append(edit,del);card.append(actions);$('#cards').append(card);}}
async function mutate(path,method,data){try{await api(path,{method,body:data?JSON.stringify(data):undefined,headers:{'Content-Type':'application/json'}});await refresh();message('已保存');return true;}catch(e){message(e.message,true);return false;}}
function methodRow(m={type:'phone',value:''}) {
 const row=el('div',undefined,'method-row');
 const picker=el('div',undefined,'type-picker'), trigger=el('button',undefined,'secondary type-trigger'), menu=el('div',undefined,'type-menu');
 const input=el('input'), custom=el('input'), remove=el('button','×','secondary');
 let kind=m.type, platform=m.platform||'';
 trigger.type='button'; trigger.setAttribute('aria-label','选择联系方式类型');
 custom.className='platform-name';custom.placeholder='输入社交软件名称';custom.maxLength=50;custom.setAttribute('aria-label','社交软件名称');
 input.className='method-value';input.value=m.value;input.required=true;input.maxLength=1000;input.setAttribute('aria-label','联系方式内容');
 function update(){trigger.textContent=kind==='social'?(platform||'社交账号')+' ▾':types[kind]+' ▾';custom.hidden=kind!=='social'||platform!=='其他';custom.required=!custom.hidden;input.placeholder=kind==='social'?'输入账号':'输入'+types[kind];}
 function choose(t,p=''){kind=t;platform=p;menu.classList.remove('open');update();if(!custom.hidden)custom.focus();}
 for(const [t,label] of Object.entries(types)){
   if(t==='social'){
     const group=el('div',undefined,'social-group'), b=el('button','社交账号 ›','secondary'), sub=el('div',undefined,'social-submenu');b.type='button';b.onclick=()=>group.classList.toggle('expanded');
     for(const name of ['QQ','微信','其他']){const option=el('button',name,'secondary');option.type='button';option.onclick=()=>choose('social',name);sub.append(option);}
     group.append(b,sub);menu.append(group);
   }else{const b=el('button',label,'secondary');b.type='button';b.onclick=()=>choose(t);menu.append(b);}
 }
 if(platform&&!['QQ','微信'].includes(platform)){custom.value=platform;platform='其他';}
 trigger.onclick=()=>menu.classList.toggle('open');picker.append(trigger,menu);
 remove.type='button';remove.setAttribute('aria-label','移除联系方式');remove.onclick=()=>row.remove();
 row.readMethod=()=>{const result={type:kind,value:input.value};if(kind==='social'){const name=platform==='其他'?custom.value.trim():platform;if(name)result.platform=name;}return result;};
 row.append(picker,custom,input,remove);update();$('#methods').append(row);
}
function openEditor(c){editing=c?.id??null;$('#form-title').textContent=c?'编辑联系人':'新建联系人';$('#name').value=c?.name||'';$('#favorite').checked=c?.favorite||false;$('#notes').value=c?.notes||'';$('#methods').replaceChildren();for(const m of c?.methods||[{type:'phone',value:''}])methodRow(m);$('#editor').showModal();}
$('#create').onclick=()=>openEditor();$('#cancel').onclick=()=>$('#editor').close();$('#add-method').onclick=()=>methodRow();
$('#form').onsubmit=async e=>{e.preventDefault();let data={name:$('#name').value,favorite:$('#favorite').checked,notes:$('#notes').value,methods:[...$('#methods').children].map(r=>r.readMethod())};const btn=e.submitter;btn.disabled=true;try{if(await mutate(editing?'contacts/'+editing:'contacts',editing?'PUT':'POST',data))$('#editor').close();}finally{btn.disabled=false;}};
$('#search').oninput=render;for(const [id,value]of [['all',false],['favorites',true]])$('#'+id).onclick=()=>{favorites=value;$('#heading').textContent=value?'我的收藏':'所有联系人';$('#all').classList.toggle('active',!value);$('#favorites').classList.toggle('active',value);render();};
$('#login').onclick=()=>{let value=prompt('请输入服务器设置的访问密码（本地未设置可留空）',token);if(value!==null){token=value;sessionStorage.setItem('token',token);refresh();}};
$('#export').onclick=async()=>{try{const blob=await(await api('export')).blob(),url=URL.createObjectURL(blob),a=el('a');a.href=url;a.download='contacts.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);message('Excel 已导出');}catch(e){message(e.message,true);}};
$('#import').onclick=()=>$('#file').click();$('#file').onchange=async e=>{const file=e.target.files[0];if(!file)return;if(!confirm('导入将新增联系人（不会覆盖或去重），是否继续？')){e.target.value='';return;}try{let result=await(await api('import',{method:'POST',body:file})).json();await refresh();message(`成功导入 ${result.imported} 位联系人`);}catch(err){message(err.message,true);}finally{e.target.value='';}};refresh();
