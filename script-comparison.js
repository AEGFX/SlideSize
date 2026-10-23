/* script-comparison.js — Script Comparison, script version compare for SlideSize
 *
 * Two versions of a script in, the new version out with every change marked.
 * Paragraphs are aligned with a Myers diff, changed paragraphs are paired by
 * word overlap (so a paragraph that moved and was edited is still recognised),
 * then each pair gets a word level diff. Everything runs in the browser.
 *
 * Readers: TXT and other plain text, DOCX, PPTX (slides and notes), PDF via
 * pdf.js, RTF, ODT, Final Draft FDX, HTML.
 */
(function(){
"use strict";
const $=id=>document.getElementById(id);
const state={A:null,B:null,example:true,changes:[],cur:-1,view:"all",out:null};

/* ---------- Example scripts ---------- */
const EX_A=`WALK-IN MUSIC FADES. HOUSE LIGHTS DOWN.

VOICE OF GOD: Ladies and gentlemen, please welcome to the stage your host for this evening, Sarah Mitchell.

SARAH: Good evening and welcome to the 2026 Partner Excellence Awards. It's wonderful to see so many of you here in London tonight.

SARAH: Before we begin, a quick reminder that dinner will be served after the first three categories, so please stay in your seats until then.

SARAH: Over the last twelve months our partners have delivered record growth across every region. Tonight is about celebrating the people behind those numbers.

SARAH: Our first award is for Innovation of the Year. To present it, please welcome our Chief Technology Officer, James Okafor.

JAMES: Thank you Sarah. This year we received over 200 nominations for this category, and choosing a winner was not easy.

SARAH: And finally, a huge thank you to our production team, who have made tonight possible.`;
const EX_B=`WALK-IN MUSIC FADES. HOUSE LIGHTS DOWN.

VOICE OF GOD: Ladies and gentlemen, please welcome to the stage your host for this evening, Sarah Mitchell.

SARAH: Good evening and welcome to the 2026 Partner Excellence Awards. It's fantastic to see so many of you here in London tonight.

SARAH: Over the last twelve months our partners have delivered record growth in every single region. Tonight is about celebrating the people behind those numbers.

SARAH: Before we begin, a quick reminder that dinner will be served after the first four categories, so please stay in your seats until then.

SARAH: Our first award is for Innovation of the Year. To present it, please welcome our Chief Product Officer, Priya Nair.

PRIYA: Thank you Sarah. This year we received over 300 nominations for this category, and choosing a winner was not easy.

PRIYA: The winner of Innovation of the Year is...

SARAH: And finally, a huge thank you to our production team, who have made tonight possible.`;

/* ---------- Reading files: shared with Script Format, see script-read.js ---------- */
const SR=window.ScriptRead;
function keyOf(w,loose){
  let k=SR.canon(w);
  if(loose){k=k.toLowerCase().replace(/[^\p{L}\p{N}]+/gu,"")}
  return k;
}

/* ---------- Myers diff ---------- */
// Returns ops: [type, aIndex, bIndex] with type "=", "-", "+"
function diff(a,b,limit){
  const ops=[];let s=0,ea=a.length,eb=b.length;
  while(s<ea&&s<eb&&a[s]===b[s]){ops.push(["=",s,s]);s++}
  const tail=[];
  while(ea>s&&eb>s&&a[ea-1]===b[eb-1]){ea--;eb--;tail.push(["=",ea,eb])}
  const n=ea-s,m=eb-s,max=n+m;
  if(max===0){return ops.concat(tail.reverse())}
  const off=max+1,v=new Int32Array(2*max+3),trace=[];
  let D=-1;
  outer:for(let d=0;d<=max;d++){
    if(limit&&d>limit){D=-1;break}
    trace.push(v.slice(off-d-1,off+d+2));
    for(let k=-d;k<=d;k+=2){
      let x=(k===-d||(k!==d&&v[off+k-1]<v[off+k+1]))?v[off+k+1]:v[off+k-1]+1;
      let y=x-k;
      while(x<n&&y<m&&a[s+x]===b[s+y]){x++;y++}
      v[off+k]=x;
      if(x>=n&&y>=m){D=d;break outer}
    }
  }
  const mid=[];
  if(D<0){
    for(let i=0;i<n;i++)mid.push(["-",s+i,-1]);
    for(let j=0;j<m;j++)mid.push(["+",-1,s+j]);
  }else{
    let x=n,y=m;
    const V=(t,d,k)=>t[k+d+1];
    for(let d=D;d>0;d--){
      const t=trace[d],k=x-y;
      const pk=(k===-d||(k!==d&&V(t,d,k-1)<V(t,d,k+1)))?k+1:k-1;
      const px=V(t,d,pk),py=px-pk;
      while(x>px&&y>py){x--;y--;mid.push(["=",s+x,s+y])}
      if(pk===k+1){mid.push(["+",-1,s+py])}else{mid.push(["-",s+px,-1])}
      x=px;y=py;
    }
    while(x>0&&y>0){x--;y--;mid.push(["=",s+x,s+y])}
    mid.reverse();
  }
  return ops.concat(mid,tail.reverse());
}

/* ---------- Compare ---------- */
function sim(x,y){
  if(!x.length||!y.length)return 0;
  const m=new Map();for(const w of x)m.set(w,(m.get(w)||0)+1);
  let c=0;for(const w of y){const v=m.get(w);if(v){c++;m.set(w,v-1)}}
  return 2*c/(x.length+y.length);
}
const esc=s=>s.replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
function compare(A,B,opt){
  const PA=SR.paragraphs(A,opt.lb),PB=SR.paragraphs(B,opt.lb);
  const pa=PA.map(p=>p.text),pb=PB.map(p=>p.text);
  const wa=PA.map(p=>p.words.map(x=>keyOf(x.w,opt.loose)));
  const wb=PB.map(p=>p.words.map(x=>keyOf(x.w,opt.loose)));
  const ka=wa.map(w=>w.join(" ")),kb=wb.map(w=>w.join(" "));
  // Formatting is compared only when both files carry it (Word and PowerPoint)
  const fmtBoth=!!(A.fmt&&B.fmt),showB=!!B.fmt;
  const ops=diff(ka,kb,0);

  const groups=[];let h=null;
  for(const o of ops){
    if(o[0]==="="){if(h){groups.push(h);h=null}groups.push({eq:true,a:o[1],b:o[2]})}
    else{if(!h)h={eq:false,dels:[],ins:[]};(o[0]==="-"?h.dels:h.ins).push(o[0]==="-"?o[1]:o[2])}
  }
  if(h)groups.push(h);
  const hunkA=new Map(),hunkB=new Map();
  groups.forEach((g,gi)=>{if(!g.eq){g.dels.forEach(a=>hunkA.set(a,gi));g.ins.forEach(b=>hunkB.set(b,gi))}});

  // Pair changed paragraphs: B index -> A index
  const pairB=new Map(),pairA=new Map();
  const link=(b,a)=>{pairB.set(b,a);pairA.set(a,b)};
  // 1 exact matches anywhere (moved paragraphs)
  const exact=new Map();
  for(const a of hunkA.keys()){if(ka[a].length>=12){const l=exact.get(ka[a])||[];l.push(a);exact.set(ka[a],l)}}
  for(const b of hunkB.keys()){const l=exact.get(kb[b]);if(l&&l.length)link(b,l.shift())}
  // 2 similar paragraphs in the same hunk
  for(const g of groups){
    if(g.eq)continue;
    const cand=[];
    for(const b of g.ins){if(pairB.has(b))continue;for(const a of g.dels){if(pairA.has(a))continue;const s=sim(wa[a],wb[b]);if(s>=0.4)cand.push([s,a,b])}}
    cand.sort((x,y)=>y[0]-x[0]);
    for(const [s,a,b] of cand)if(!pairA.has(a)&&!pairB.has(b))link(b,a);
  }
  // 3 strongly similar paragraphs in other hunks (moved and edited)
  const ua=[...hunkA.keys()].filter(a=>!pairA.has(a)&&wa[a].length>=5);
  const ub=[...hunkB.keys()].filter(b=>!pairB.has(b)&&wb[b].length>=5);
  if(ua.length*ub.length<250000){
    const cand=[];
    for(const b of ub)for(const a of ua){const s=sim(wa[a],wb[b]);if(s>=0.6)cand.push([s,a,b])}
    cand.sort((x,y)=>y[0]-x[0]);
    for(const [s,a,b] of cand)if(!pairA.has(a)&&!pairB.has(b))link(b,a);
  }

  const out=[];let lastB=-1,chgId=0;
  const stats={add:0,rem:0,moved:0,fmt:0,paras:0};
  // one word of B, shown with its highlight and weight
  const wordB=(x)=>{const st=showB?SR.fmtStyle(x.f):"";return st?'<span class="w" style="'+st+'">'+esc(x.w)+'</span>':esc(x.w)};

  function wordPara(a,b,note,moved){
    const ta=PA[a].words,tb=PB[b].words;
    let w=diff(wa[a],wb[b],2000);w=cleanup(w,tb.map(x=>x.w));
    const p={n:b+1,changed:false,moved,ids:[],fmts:[],was:pa[a],now:pb[b],html:note?'<span class="note">'+note+'</span>':""};
    let mode="",curFmt="",fmtWords=[],first=true;
    const sp=()=>{const r=first?"":" ";first=false;return r};
    const close=()=>{
      if(mode==="chg")p.html+="</span>";
      else if(mode==="fmt"){p.html+='</span><span class="fmt-tag">'+esc(curFmt)+'</span>';p.fmts.push({text:fmtWords.join(" "),desc:curFmt});stats.fmt++}
      mode="";fmtWords=[];
    };
    for(const [t,ai,bi] of w){
      if(t==="="){
        const d=fmtBoth?SR.describeFmt(ta[ai].f,tb[bi].f):"";
        if(d){
          if(mode!=="fmt"||curFmt!==d){close();const id=chgId++;p.ids.push(id);p.html+=sp()+'<span class="chg fmtchg" id="c'+id+'" title="'+esc(d)+'">';first=true;mode="fmt";curFmt=d}
          p.html+=sp()+wordB(tb[bi]);fmtWords.push(tb[bi].w);
        }else{close();p.html+=sp()+wordB(tb[bi])}
        continue;
      }
      if(mode!=="chg"){close();const id=chgId++;p.ids.push(id);p.html+=sp()+'<span class="chg" id="c'+id+'">';first=true;mode="chg"}
      if(t==="-"){p.html+=sp()+"<del>"+esc(ta[ai].w)+"</del>";stats.rem++}
      else{p.html+=sp()+"<ins>"+esc(tb[bi].w)+"</ins>";stats.add++}
    }
    close();
    p.html=p.html.replace(/<\/del> <del>/g," ").replace(/<\/ins> <ins>/g," ");
    if(moved&&!p.ids.length){const id=chgId++;p.ids.push(id);p.html=p.html.replace(/(<\/span>)/,'$1<span class="chg" id="c'+id+'">')+"</span>"}
    p.changed=p.ids.length>0;
    p.textChanged=ka[a]!==kb[b];
    if(p.changed)stats.paras++;
    if(moved)stats.moved++;
    return p;
  }
  function removedPara(a){
    const id=chgId++;stats.rem+=wa[a].length;stats.paras++;
    return {n:null,gone:true,changed:true,ids:[id],afterN:lastB+1,was:pa[a],now:"",html:'<span class="chg" id="c'+id+'"><del>'+esc(pa[a])+'</del></span>'};
  }
  function addedPara(b){
    const id=chgId++;stats.add+=wb[b].length;stats.paras++;
    return {n:b+1,changed:true,ids:[id],was:"",now:pb[b],html:'<span class="chg" id="c'+id+'"><ins>'+esc(pb[b])+'</ins></span>'};
  }
  const fmtDiffers=(a,b)=>fmtBoth&&PA[a].words.some((x,i)=>x.f!==PB[b].words[i].f);

  groups.forEach((g,gi)=>{
    if(g.eq){
      if(fmtDiffers(g.a,g.b))out.push(wordPara(g.a,g.b,"",false));
      else out.push({n:g.b+1,html:PB[g.b].words.map(wordB).join(" "),changed:false});
      lastB=g.b;return;
    }
    // A paragraphs from this hunk that belong elsewhere now
    const leftovers=[];
    for(const a of g.dels){
      if(pairA.has(a)&&hunkB.get(pairA.get(a))!==gi){
        out.push({n:null,placeholder:true,gone:true,changed:false,html:'<span class="note">Moved to ¶ '+(pairA.get(a)+1)+'</span><del>'+esc(trunc(pa[a],90))+'</del>'});
      }else if(!pairA.has(a))leftovers.push(a);
    }
    let lastPairedA=-1;
    for(const b of g.ins){
      const a=pairB.get(b);
      const local=a!==undefined&&hunkA.get(a)===gi;
      if(local){
        while(leftovers.length&&leftovers[0]<a)out.push(removedPara(leftovers.shift()));
        const crossed=a<lastPairedA;
        out.push(wordPara(a,b,crossed?"Moved from ¶ "+(a+1)+" in the previous version":"",crossed));
        lastPairedA=Math.max(lastPairedA,a);
      }else if(a!==undefined){
        const same=ka[a]===kb[b];
        out.push(wordPara(a,b,same?"Moved from ¶ "+(a+1)+" in the previous version":"Moved from ¶ "+(a+1)+" and edited",true));
      }else{
        out.push(addedPara(b));
      }
      lastB=b;
    }
    while(leftovers.length)out.push(removedPara(leftovers.shift()));
  });
  return {out,count:chgId,stats,fmtBoth,paraCountA:pa.length,paraCountB:pb.length};
}
function trunc(s,n){return s.length>n?s.slice(0,n).replace(/\s+\S*$/,"")+"…":s}

// Absorb single unchanged words stranded between changes, then show removed words before added ones
function cleanup(ops,tb){
  const r=ops.slice();
  for(let i=1;i<r.length-1;i++){
    if(r[i][0]==="="&&r[i-1][0]!=="="&&r[i+1][0]!=="="&&tb[r[i][2]].length<=4){
      r.splice(i,1,["-",r[i][1],-1],["+",-1,r[i][2]]);i++;
    }
  }
  const res=[];let i=0;
  while(i<r.length){
    if(r[i][0]==="="){res.push(r[i]);i++;continue}
    const d=[],a=[];
    while(i<r.length&&r[i][0]!=="="){(r[i][0]==="-"?d:a).push(r[i]);i++}
    res.push(...d,...a);
  }
  return res;
}

/* ---------- Render ---------- */
function render(){
  const res=$("result"),stats=$("stats");
  if(!state.A||!state.B){
    res.innerHTML='<div class="empty-state"><strong>Waiting for two versions</strong>Add the '+(state.A?"new":"previous")+' version to see what changed.</div>';
    stats.innerHTML="";state.changes=[];updateNav();return;
  }
  const opt={loose:$("optCase").checked,lb:$("optLB").checked};
  const t0=performance.now();
  const r=compare(state.A,state.B,opt);
  state.out=r;
  const s=r.stats;
  if(!r.count){
    stats.innerHTML='<span class="chip same">No changes found</span><span class="chip">'+r.paraCountB+' paragraphs</span>';
  }else{
    stats.innerHTML='<span class="chip big">'+r.count+' change'+(r.count===1?"":"s")+'</span>'
      +'<span class="chip">'+s.paras+' ¶ affected</span>'
      +(s.add?'<span class="chip add">+'+s.add+' words</span>':"")
      +(s.rem?'<span class="chip rem">−'+s.rem+' words</span>':"")
      +(s.moved?'<span class="chip mov">'+s.moved+' moved</span>':"")
      +(s.fmt?'<span class="chip fmt">'+s.fmt+' formatting</span>':"");
  }
  let html="";
  if(state.view==="chg"){
    let skipped=0;
    r.out.forEach((p,i)=>{
      if(p.changed||p.placeholder){if(skipped)html+='<div class="gap">'+skipped+' unchanged</div>';skipped=0;html+=paraHTML(p)}
      else skipped++;
    });
    if(skipped&&html)html+='<div class="gap">'+skipped+' unchanged</div>';
    if(!html)html='<div class="empty-state"><strong>The two versions match</strong>Nothing to show in Changes only.</div>';
  }else{
    html=r.out.map(paraHTML).join("");
  }
  res.innerHTML=html;
  state.changes=Array.from(res.querySelectorAll(".chg"));
  state.cur=-1;updateNav();
}
function paraHTML(p){
  const cls=["para"];if(p.changed)cls.push("changed");if(p.moved)cls.push("moved");if(p.gone||p.placeholder)cls.push("gone");
  const n=p.n!=null?"¶ "+p.n:(p.placeholder?"":"cut");
  return '<div class="'+cls.join(" ")+'"><span class="n">'+n+'</span><div class="t">'+(p.gone&&!p.placeholder?'<span class="note">Removed</span>':"")+(p.html||"&nbsp;")+'</div></div>';
}
function updateNav(){
  const n=state.changes.length;
  $("count").textContent=(state.cur<0?"–":state.cur+1)+" / "+n;
  $("prev").disabled=$("next").disabled=!n;$("copy").disabled=$("dl").disabled=!state.out||!state.out.count;
}
function go(dir){
  const n=state.changes.length;if(!n)return;
  if(state.cur>=0)state.changes[state.cur].classList.remove("current");
  state.cur=state.cur<0?(dir>0?0:n-1):(state.cur+dir+n)%n;
  const el=state.changes[state.cur];el.classList.add("current");
  el.scrollIntoView({block:"center",behavior:matchMedia("(prefers-reduced-motion: reduce)").matches?"auto":"smooth"});
  updateNav();
}

/* ---------- Copy list ---------- */
function changeList(){
  const r=state.out;if(!r)return"";
  const lines=["Script changes","A  "+state.A.name,"B  "+state.B.name,r.count+" change"+(r.count===1?"":"s")+" in "+r.stats.paras+" paragraph"+(r.stats.paras===1?"":"s"),""];
  for(const p of r.out){
    if(!p.changed)continue;
    if(p.moved){lines.push("¶ "+p.n+"  MOVED here","Now  "+p.now,"");continue}
    if(p.gone){lines.push("Cut (after ¶ "+(p.afterN||0)+")","Was  "+p.was,"");continue}
    lines.push("¶ "+p.n);
    if(p.textChanged!==false){
      if(p.was)lines.push("Was  "+p.was);
      lines.push("Now  "+(p.now||"(removed)"));
    }
    (p.fmts||[]).forEach(f=>lines.push("Formatting  \""+trunc(f.text,60)+"\"  "+f.desc));
    lines.push("");
  }
  return lines.join("\n");
}
async function copyList(){
  const t=changeList();
  try{await navigator.clipboard.writeText(t);toast("Change list copied")}
  catch(e){
    const ta=document.createElement("textarea");ta.value=t;ta.style.position="fixed";ta.style.opacity="0";document.body.appendChild(ta);ta.select();
    let ok=false;try{ok=document.execCommand("copy")}catch(_){}
    ta.remove();toast(ok?"Change list copied":"Copy was blocked. Try again from a normal browser tab.");
  }
}
function toast(m){SS.showToast(m)}

/* ---------- Inputs ---------- */
function setNotice(html,err){$("notice").innerHTML=html?'<div class="notice'+(err?" err":"")+'">'+html+'</div>':""}
function showSlot(side){
  const s=state[side],el=$("slot"+side);
  el.classList.toggle("filled",!!s);
  const nm=$("name"+side);
  if(s){nm.textContent=s.name;nm.classList.remove("empty");
    const words=s.text.split(/\s+/).filter(Boolean).length;
    $("meta"+side).textContent=s.kind+" · "+words.toLocaleString("en-GB")+" words";
  }else{nm.textContent="Drop a file here";nm.classList.add("empty");$("meta"+side).textContent=""}
}
function clearExample(){
  if(state.example){state.example=false;
    if(state.A&&state.A.example)state.A=null;if(state.B&&state.B.example)state.B=null;showSlot("A");showSlot("B");setNotice("")}
}
async function loadFiles(files,side){
  files=Array.from(files).slice(0,2);if(!files.length)return;
  clearExample();
  document.body.classList.add("busy");setNotice("Reading "+files.map(f=>f.name).join(" and ")+"…");
  try{
    if(files.length===2){
      const [x,y]=await Promise.all(files.map(SR.extract));
      const older=x.modified<=y.modified?[x,y]:[y,x];
      state.A=older[0];state.B=older[1];closePaste("A");closePaste("B");
    }else{
      state[side]=await SR.extract(files[0]);closePaste(side);
    }
    setNotice("");
  }catch(e){setNotice(escapeT(e.message||String(e)),true)}
  finally{document.body.classList.remove("busy")}
  showSlot("A");showSlot("B");render();
}
function escapeT(s){return s.replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]))}
function nextEmpty(){return !state.A?"A":!state.B?"B":"B"}

["A","B"].forEach(side=>{
  const slot=$("slot"+side),input=$("file"+side);
  slot.addEventListener("dragover",e=>{e.preventDefault();e.stopPropagation();slot.classList.add("over")});
  slot.addEventListener("dragleave",()=>slot.classList.remove("over"));
  slot.addEventListener("drop",e=>{e.preventDefault();e.stopPropagation();slot.classList.remove("over");loadFiles(e.dataTransfer.files,side)});
  input.addEventListener("change",()=>{loadFiles(input.files,side);input.value=""});
  document.querySelector('[data-pick="'+side+'"]').addEventListener("click",()=>input.click());
  document.querySelector('[data-paste="'+side+'"]').addEventListener("click",()=>openPaste(side));
  const ta=$("paste"+side);let deb;
  ta.addEventListener("input",()=>{clearTimeout(deb);deb=setTimeout(()=>{
    if(state.example)clearExample();
    const v=ta.value;state[side]=v.trim()?SR.fromText(v):null;showSlot(side);render()},250)});
});
function openPaste(side){
  clearExample();
  const ta=$("paste"+side);ta.hidden=false;
  if(state[side]&&!ta.value)ta.value=state[side].text;
  ta.focus();
}
function closePaste(side){const ta=$("paste"+side);ta.hidden=true;ta.value=""}
document.addEventListener("dragover",e=>e.preventDefault());
document.addEventListener("drop",e=>{e.preventDefault();if(e.dataTransfer.files.length)loadFiles(e.dataTransfer.files,e.dataTransfer.files.length>1?"A":nextEmpty())});

$("swap").addEventListener("click",()=>{[state.A,state.B]=[state.B,state.A];
  const a=$("pasteA"),b=$("pasteB");[a.value,b.value]=[b.value,a.value];[a.hidden,b.hidden]=[b.hidden,a.hidden];
  showSlot("A");showSlot("B");render()});
$("vAll").addEventListener("click",()=>{state.view="all";$("vAll").setAttribute("aria-pressed","true");$("vChg").setAttribute("aria-pressed","false");render()});
$("vChg").addEventListener("click",()=>{state.view="chg";$("vChg").setAttribute("aria-pressed","true");$("vAll").setAttribute("aria-pressed","false");render()});
$("prev").addEventListener("click",()=>go(-1));
$("next").addEventListener("click",()=>go(1));
$("copy").addEventListener("click",copyList);
$("dl").addEventListener("click",()=>{
  const base=(state.B&&state.B.name||"script").replace(/\.[^.]+$/,"");
  SS.download(changeList()+"\n\nMade with Script Comparison, slidesize.com\n","changes-"+base+".txt","text/plain;charset=utf-8");
});
$("optDel").addEventListener("change",e=>document.body.classList.toggle("hide-del",!e.target.checked));
$("optCase").addEventListener("change",render);
$("optLB").addEventListener("change",render);
$("optHl").addEventListener("change",e=>document.body.classList.toggle("no-hl",!e.target.checked));
document.addEventListener("keydown",e=>{
  if(e.target.closest("textarea,input"))return;
  if(e.key==="n"||e.key==="j"){go(1);e.preventDefault()}
  else if(e.key==="p"||e.key==="k"){go(-1);e.preventDefault()}
});

/* ---------- Start with the example ---------- */
// Example files carry highlights like a colour coded Word script, to show formatting changes
function exampleSrc(name,text,marks){
  const keys=[""],fmt=new Int32Array(text.length);
  marks.forEach(([str,key])=>{let id=keys.indexOf(key);if(id<0){id=keys.length;keys.push(key)}
    let i=text.indexOf(str);while(i>=0){fmt.fill(id,i,i+str.length);i=text.indexOf(str,i+str.length)}});
  return {name,kind:"Example",text,fmt,fmtKeys:keys,flow:false,example:true};
}
state.A=exampleSrc("awards-script-v3.docx",EX_A,[["SARAH:","h=FFFF00"],["JAMES:","h=00FF00"],["VOICE OF GOD:","h=00FFFF"]]);
state.B=exampleSrc("awards-script-v4.docx",EX_B,[["SARAH:","h=FFFF00"],["PRIYA:","h=00FF00"],["VOICE OF GOD:","h=FF00FF"],["fantastic","b"]]);
showSlot("A");showSlot("B");
setNotice('<span class="tag">Example</span> Two versions of a sample awards script are loaded. Drop your own files to replace them.');
render();

/* Files dropped on the Script Comparison tile on the home page */
SS.consumeHandoff(files=>{clearExample();loadFiles(files,"A")});
})();
