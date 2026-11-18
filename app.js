const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const state={zip:null,file:null,paragraphs:[],blob:null,view:'diff'};
const esc=s=>s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const toast=msg=>{const el=$('#toast');el.textContent=msg;el.classList.add('show');setTimeout(()=>el.classList.remove('show'),2400)};

function looksLikeSpeaker(text){
  const s=text.trim().replace(/[：:]$/,'');
  return s.length>0&&s.length<28&&!/[。！？!?，,]/.test(s)&&(/^[A-Z][A-Z .'-]{1,26}$/.test(s)||/^[\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9·._-]{0,15}$/.test(s));
}
function isExcluded(text){return /^(INT\.|EXT\.|INT\/EXT\.|场景|镜头|画面|转场|第.{0,5}场|△|▲|【|\[)/i.test(text.trim())}
function findTranslations(text,prev){
  if(isExcluded(text))return [];
  // Allow short parenthetical performance directions between the speaker name
  // and the colon, e.g. 马库斯(皱眉)：... Keep those directions intact.
  const performanceDirection=String.raw`(?:[（(][^()（）\r\n]{1,40}[）)]\s*)*`;
  // Screenplay delivery markers may appear after the name/direction and
  // before the colon: OS, O.S., VO, V.O. or OFF.
  const deliveryMarker=String.raw`(?:(?:O\.?\s*S\.?|V\.?\s*O\.?|OFF)\s*)?`;
  const speakerColon=new RegExp(String.raw`^\s*(?:[A-Z][A-Z .'-]{1,26}|[\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9·._-]{0,15})\s*${performanceDirection}${deliveryMarker}[：:]\s*`,'i').test(text);
  const context=speakerColon||looksLikeSpeaker(prev||'');
  if(!context)return [];
  const out=[]; const re=/[（(]([^()（）]{1,120})[）)]/g; let m;
  while((m=re.exec(text))){
    const before=text.slice(0,m.index);
    const inner=m[1];
    const after=text.slice(m.index+m[0].length);
    // A parenthesis immediately before a colon belongs to the speaker cue,
    // including a second speaker later in the same Word paragraph.
    if(/^\s*[：:]/.test(after)) continue;
    const latin=(before.match(/[A-Za-z]/g)||[]).length;
    const han=(inner.match(/[\u3400-\u9fff]/g)||[]).length;
    // Punctuation (especially "......") should not dilute a short Chinese
    // translation. Compare Han characters only with semantic characters.
    const semantic=(inner.match(/[A-Za-z0-9\u3400-\u9fff]/g)||[]).length;
    // Explicit dialogue context makes short replies such as Good/Yes/No safe
    // to recognize; keep a two-letter minimum to avoid isolated initials.
    if(latin>=2&&han>=1&&han/Math.max(semantic,1)>=.5) out.push({start:m.index,end:m.index+m[0].length,text:m[0]});
  }
  return out;
}
function textNodes(p){return [...p.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','t')]}
function applyRanges(nodes,ranges){
  let offset=0;
  for(const node of nodes){const src=node.textContent||'';let out='';for(let i=0;i<src.length;i++){const g=offset+i;if(!ranges.some(r=>g>=r.start&&g<r.end))out+=src[i]}node.textContent=out;offset+=src.length}
}
async function processFile(file){
  if(!file.name.toLowerCase().endsWith('.docx')){toast('请选择 .docx 格式的 Word 文件');return}
  try{
    const zip=await JSZip.loadAsync(file); const entry=zip.file('word/document.xml');
    if(!entry)throw new Error('缺少 document.xml');
    const xml=await entry.async('string'); const doc=new DOMParser().parseFromString(xml,'application/xml');
    if(doc.querySelector('parsererror'))throw new Error('文档 XML 无法解析');
    const ps=[...doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','p')];
    const paragraphs=[];let prev='';
    ps.forEach(p=>{const nodes=textNodes(p),before=nodes.map(n=>n.textContent||'').join('');const ranges=findTranslations(before,prev);applyRanges(nodes,ranges);const after=nodes.map(n=>n.textContent||'').join('');paragraphs.push({before,after,ranges});if(before.trim())prev=before});
    zip.file('word/document.xml',new XMLSerializer().serializeToString(doc));
    const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
    Object.assign(state,{zip,file,paragraphs,blob});
    const count=paragraphs.reduce((n,p)=>n+p.ranges.length,0), chars=paragraphs.reduce((n,p)=>n+p.ranges.reduce((a,r)=>a+r.text.length,0),0);
    $('#fileName').textContent=file.name;$('#fileMeta').textContent=`${(file.size/1024).toFixed(1)} KB · 已在本地完成分析`;
    $('#matchCount').textContent=count;$('#charCount').textContent=chars;$('#uploadCard').classList.add('hidden');$('#resultCard').classList.remove('hidden');render();
    toast(count?`已识别 ${count} 处括号翻译`:'未发现符合保守规则的括号翻译');
  }catch(e){console.error(e);toast('无法读取该文件，请确认它是有效的 .docx 文档')}
}
function diffHtml(p){
  if(!p.ranges.length)return esc(p.before);
  let html='',at=0;p.ranges.forEach(r=>{html+=esc(p.before.slice(at,r.start))+`<span class="deleted" title="将删除">${esc(r.text)}</span>`;at=r.end});return html+esc(p.before.slice(at));
}
function render(){
  $('#preview').innerHTML=state.paragraphs.map(p=>{const text=state.view==='before'?esc(p.before):state.view==='after'?esc(p.after):diffHtml(p);return `<div class="para ${p.ranges.length?'changed':''}">${text||'&nbsp;'}</div>`}).join('');
}
function reset(){Object.assign(state,{zip:null,file:null,paragraphs:[],blob:null,view:'diff'});$('#fileInput').value='';$('#resultCard').classList.add('hidden');$('#uploadCard').classList.remove('hidden')}
$('#chooseBtn').onclick=$('#replaceBtn').onclick=()=>$('#fileInput').click();$('#fileInput').onchange=e=>e.target.files[0]&&processFile(e.target.files[0]);$('#resetBtn').onclick=reset;
$('#downloadBtn').onclick=()=>{if(!state.blob)return;const a=document.createElement('a');a.href=URL.createObjectURL(state.blob);a.download=state.file.name.replace(/\.docx$/i,'')+'_已清理.docx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('处理后的 Word 已开始下载')};
$$('.tabs button').forEach(b=>b.onclick=()=>{$$('.tabs button').forEach(x=>x.classList.remove('active'));b.classList.add('active');state.view=b.dataset.view;render()});
const card=$('#uploadCard');['dragenter','dragover'].forEach(n=>card.addEventListener(n,e=>{e.preventDefault();card.classList.add('drag')}));['dragleave','drop'].forEach(n=>card.addEventListener(n,e=>{e.preventDefault();card.classList.remove('drag')}));card.addEventListener('drop',e=>e.dataTransfer.files[0]&&processFile(e.dataTransfer.files[0]));
