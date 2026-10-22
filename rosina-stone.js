/* rosina-stone.js — Rosina Stone, script version compare for SlideSize
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

/* ---------- Script loaders (lazy) ---------- */
const loaded={};
function loadScript(src){
  if(loaded[src])return loaded[src];
  return loaded[src]=new Promise((res,rej)=>{const s=document.createElement("script");s.src=src;s.onload=res;s.onerror=()=>{delete loaded[src];rej(new Error("Couldn't load a helper library. Check your internet connection and try again."))};document.head.appendChild(s)});
}
function getZip(){return SS.loadJSZip()}
async function getPdf(){
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js");
  const lib=window.pdfjsLib;lib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  return lib;
}

/* ---------- Text extraction ---------- */
function decodeText(buf){
  let t=new TextDecoder("utf-8").decode(buf);
  if((t.match(/\uFFFD/g)||[]).length>3)t=new TextDecoder("windows-1252").decode(buf);
  return t.replace(/^﻿/,"");
}
const xml=s=>new DOMParser().parseFromString(s,"application/xml");
const byTag=(n,t)=>Array.from(n.getElementsByTagName(t));

function rtfToText(s){
  const skipDest=new Set(["fonttbl","colortbl","stylesheet","info","pict","header","footer","headerl","headerr","footerl","footerr","themedata","datastore","latentstyles","xmlnstbl","listtable","listoverridetable","rsidtbl","generator","mmathPr","object","fldinst","filetbl","revtbl","listtext","pntext","pntxtb","pntxta"]);
  let out="",i=0,depth=0,skipAt=-1,uc=1,pendingSkip=0;const n=s.length;
  while(i<n){
    const c=s[i];
    if(c==="{"){depth++;i++;continue}
    if(c==="}"){if(skipAt===depth)skipAt=-1;depth--;i++;continue}
    if(c==="\\"){
      const d=s[i+1];
      if(d==="\\"||d==="{"||d==="}"){if(skipAt<0)out+=d;i+=2;continue}
      if(d==="*"){if(skipAt<0)skipAt=depth;i+=2;continue}
      if(d==="'"){const h=parseInt(s.substr(i+2,2),16);if(pendingSkip>0)pendingSkip--;else if(skipAt<0)out+=String.fromCharCode(h);i+=4;continue}
      if(d==="~"){if(skipAt<0)out+=" ";i+=2;continue}
      if(d==="-"||d==="_"){if(skipAt<0&&d==="_")out+="-";i+=2;continue}
      if(d==="\n"||d==="\r"){if(skipAt<0)out+="\n\n";i+=2;continue}
      const m=/^([a-zA-Z]+)(-?\d+)? ?/.exec(s.slice(i+1,i+40));
      if(!m){i+=2;continue}
      i+=1+m[0].length;const w=m[1],p=m[2];
      if(skipDest.has(w)){if(skipAt<0)skipAt=depth;continue}
      if(skipAt>=0)continue;
      if(w==="par"||w==="sect"||w==="page")out+="\n\n";
      else if(w==="line")out+="\n";
      else if(w==="tab"||w==="cell")out+="\t";
      else if(w==="row")out+="\n";
      else if(w==="uc")uc=+p||1;
      else if(w==="u"){let v=+p;if(v<0)v+=65536;out+=String.fromCharCode(v);pendingSkip=uc}
      else if(w==="emdash"||w==="endash")out+="-";
      else if(w==="lquote"||w==="rquote")out+="'";
      else if(w==="ldblquote"||w==="rdblquote")out+='"';
      continue;
    }
    if(c==="\n"||c==="\r"){i++;continue}
    if(skipAt<0){if(pendingSkip>0)pendingSkip--;else out+=c}
    i++;
  }
  return out;
}

function htmlToText(s){
  const doc=new DOMParser().parseFromString(s,"text/html");
  doc.querySelectorAll("script,style,noscript,head").forEach(e=>e.remove());
  const block=/^(P|DIV|H[1-6]|LI|TR|BLOCKQUOTE|SECTION|ARTICLE|HEADER|FOOTER|PRE|TABLE|UL|OL|DT|DD)$/;
  let out="";
  (function walk(n){
    for(const c of n.childNodes){
      if(c.nodeType===3)out+=c.nodeValue.replace(/\s+/g," ");
      else if(c.nodeType===1){
        if(c.tagName==="BR"){out+="\n";continue}
        const b=block.test(c.tagName);if(b)out+="\n\n";walk(c);if(b)out+="\n\n";
      }
    }
  })(doc.body||doc);
  return out;
}

function docxParas(x){
  const body=x.getElementsByTagName("w:body")[0]||x;
  const parts=[];
  for(const p of byTag(body,"w:p")){
    let t="";
    (function walk(n){
      for(const c of n.childNodes){
        if(c.nodeType!==1)continue;
        const tn=c.tagName;
        if(tn==="w:t")t+=c.textContent;
        else if(tn==="w:tab")t+="\t";
        else if(tn==="w:br"||tn==="w:cr")t+="\n";
        else if(tn==="w:noBreakHyphen")t+="-";
        else if(tn==="w:del"||tn==="w:delText"||tn==="w:instrText"||tn==="w:p")continue;
        else walk(c);
      }
    })(p);
    parts.push(t);
  }
  return parts.join("\n\n");
}

async function pptxText(buf){
  const JSZip=await getZip();const zip=await JSZip.loadAsync(buf);
  const slides=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a,b)=>+a.match(/\d+/)[0]-+b.match(/\d+/)[0]);
  // Use presentation order when available
  try{
    const pres=xml(await zip.file("ppt/presentation.xml").async("string"));
    const rels=xml(await zip.file("ppt/_rels/presentation.xml.rels").async("string"));
    const map={};byTag(rels,"Relationship").forEach(r=>map[r.getAttribute("Id")]=r.getAttribute("Target"));
    const ordered=byTag(pres,"p:sldId").map(s=>"ppt/"+map[s.getAttribute("r:id")]).filter(f=>zip.file(f));
    if(ordered.length)slides.splice(0,slides.length,...ordered);
  }catch(e){}
  const paraText=p=>{let t="";for(const c of byTag(p,"*")){if(c.tagName==="a:t")t+=c.textContent;else if(c.tagName==="a:br")t+="\n"}return t};
  let out=[];
  for(let i=0;i<slides.length;i++){
    const f=slides[i];
    const x=xml(await zip.file(f).async("string"));
    out.push("SLIDE "+(i+1));
    byTag(x,"a:p").forEach(p=>{const t=paraText(p);if(t.trim())out.push(t)});
    const relF=f.replace("slides/","slides/_rels/")+".rels";
    if(zip.file(relF)){
      const r=xml(await zip.file(relF).async("string"));
      const nr=byTag(r,"Relationship").find(e=>/notesSlide$/.test(e.getAttribute("Type")));
      if(nr){
        const nf="ppt/"+nr.getAttribute("Target").replace(/^\.\.\//,"");
        if(zip.file(nf)){
          const nx=xml(await zip.file(nf).async("string"));
          const notes=[];
          byTag(nx,"p:sp").forEach(sp=>{const ph=sp.getElementsByTagName("p:ph")[0];if(ph&&ph.getAttribute("type")==="body")byTag(sp,"a:p").forEach(p=>{const t=paraText(p);if(t.trim())notes.push(t)})});
          if(notes.length){out.push("NOTES (slide "+(i+1)+")");out.push(...notes)}
        }
      }
    }
  }
  return out.join("\n\n");
}

async function pdfText(buf){
  const lib=await getPdf();
  const pdf=await lib.getDocument({data:new Uint8Array(buf)}).promise;
  const paras=[];let cur="";
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);const tc=await page.getTextContent();
    const lines=[];let line="",lastY=null,lastH=12;
    for(const it of tc.items){
      if(!("str" in it))continue;
      const y=it.transform[5];const h=Math.abs(it.transform[3])||lastH;
      if(lastY!==null&&Math.abs(y-lastY)>h*0.5){lines.push({t:line,y:lastY,h:lastH});line=""}
      line+=it.str;lastY=y;lastH=h;
      if(it.hasEOL){lines.push({t:line,y,h});line="";lastY=null}
    }
    if(line)lines.push({t:line,y:lastY,h:lastH});
    const L=lines.filter(l=>l.t.trim());
    const gaps=[];for(let i=1;i<L.length;i++)gaps.push(Math.abs(L[i-1].y-L[i].y));
    const med=gaps.length?gaps.slice().sort((a,b)=>a-b)[Math.floor(gaps.length/2)]:12;
    for(let i=0;i<L.length;i++){
      if(i>0){const g=Math.abs(L[i-1].y-L[i].y);if(g>med*1.45||g<0.1){if(cur.trim())paras.push(cur);cur=""}}
      cur+=(cur?" ":"")+L[i].t.trim();
    }
    if(cur.trim()&&/[.!?:"')\]]$/.test(cur.trim())){paras.push(cur);cur=""}
  }
  if(cur.trim())paras.push(cur);
  return paras.join("\n\n");
}

async function extract(file){
  const name=file.name;const ext=(name.split(".").pop()||"").toLowerCase();
  const buf=await file.arrayBuffer();
  let text="",kind=ext.toUpperCase(),flow=false;
  if(ext==="docx"||ext==="docm"||ext==="dotx"){
    const JSZip=await getZip();const z=await JSZip.loadAsync(buf);
    text=docxParas(xml(await z.file("word/document.xml").async("string")));kind="Word";
  }else if(ext==="pptx"||ext==="pptm"||ext==="ppsx"){text=await pptxText(buf);kind="PowerPoint"}
  else if(ext==="pdf"){text=await pdfText(buf);kind="PDF";flow=true}
  else if(ext==="odt"){
    const JSZip=await getZip();const z=await JSZip.loadAsync(buf);
    const x=xml(await z.file("content.xml").async("string"));
    const ps=byTag(x,"*").filter(e=>e.tagName==="text:p"||e.tagName==="text:h");
    text=ps.map(p=>{let t="";(function w(n){for(const c of n.childNodes){if(c.nodeType===3)t+=c.nodeValue;else if(c.nodeType===1){if(c.tagName==="text:s")t+=" ".repeat(+c.getAttribute("text:c")||1);else if(c.tagName==="text:tab")t+="\t";else if(c.tagName==="text:line-break")t+="\n";else if(c.tagName!=="text:note")w(c)}}})(p);return t}).join("\n\n");kind="OpenDocument";
  }else if(ext==="rtf"){text=rtfToText(decodeText(buf));kind="RTF"}
  else if(ext==="fdx"){
    const x=xml(decodeText(buf));
    text=byTag(x,"Paragraph").map(p=>{const ty=p.getAttribute("Type")||"";const t=byTag(p,"Text").map(e=>e.textContent).join("");return ty==="Character"||ty==="Scene Heading"?t.toUpperCase():t}).join("\n\n");kind="Final Draft";
  }else if(ext==="html"||ext==="htm"){text=htmlToText(decodeText(buf));kind="HTML"}
  else if(ext==="doc"||ext==="ppt"||ext==="pages"||ext==="key"){
    throw new Error(name+" is an older or Apple format this page can't read. Save it as ."+(ext==="ppt"||ext==="key"?"pptx":"docx")+" or PDF and drop it again.");
  }else{
    text=decodeText(buf);kind=ext?ext.toUpperCase():"Text";
    if(/\0/.test(text.slice(0,2000)))throw new Error(name+" doesn't look like a text file. Try saving it as .txt, .docx or PDF.");
  }
  if(!text.trim())throw new Error("Couldn't find any text in "+name+". If it's a scanned PDF it has no real text to compare.");
  return {name,kind,text,flow,modified:file.lastModified};
}

/* ---------- Normalising and splitting ---------- */
function canon(s){
  return s.replace(/[‘’‚′]/g,"'").replace(/[“”„″]/g,'"')
    .replace(/[–—−]/g,"-").replace(/…/g,"...").replace(/[   ]/g," ");
}
function splitParas(src,ignoreLB){
  const t=src.text.replace(/\r\n?/g,"\n");
  let raw;
  if(ignoreLB||src.flow)raw=t.split(/\n\s*\n/).map(p=>p.replace(/\s*\n\s*/g," "));
  else raw=t.split(/\n/);
  return raw.map(p=>p.replace(/[ \t]+/g," ").trim()).filter(Boolean);
}
function keyOf(w,loose){
  let k=canon(w);
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
  const pa=splitParas(A,opt.lb),pb=splitParas(B,opt.lb);
  const tok=p=>p.split(" ");
  const wa=pa.map(p=>tok(p).map(w=>keyOf(w,opt.loose)));
  const wb=pb.map(p=>tok(p).map(w=>keyOf(w,opt.loose)));
  const ka=wa.map(w=>w.join(" ")),kb=wb.map(w=>w.join(" "));
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
  const stats={add:0,rem:0,moved:0,paras:0};

  function wordPara(a,b,note,moved){
    const ta=tok(pa[a]),tb=tok(pb[b]);
    let w=diff(wa[a],wb[b],2000);w=cleanup(w,tb);
    const p={n:b+1,changed:false,moved,ids:[],was:pa[a],now:pb[b],html:note?'<span class="note">'+note+'</span>':""};
    let inChg=false,first=true;
    const sp=()=>{const r=first?"":" ";first=false;return r};
    for(const [t,ai,bi] of w){
      if(t==="="){if(inChg){p.html+="</span>";inChg=false}p.html+=sp()+esc(tb[bi]);continue}
      if(!inChg){const id=chgId++;p.ids.push(id);p.html+=sp()+'<span class="chg" id="c'+id+'">';inChg=true;first=true}
      if(t==="-"){p.html+=sp()+"<del>"+esc(ta[ai])+"</del>";stats.rem++}
      else{p.html+=sp()+"<ins>"+esc(tb[bi])+"</ins>";stats.add++}
    }
    if(inChg)p.html+="</span>";
    p.html=p.html.replace(/<\/del> <del>/g," ").replace(/<\/ins> <ins>/g," ");
    if(moved&&!p.ids.length){const id=chgId++;p.ids.push(id);p.html=p.html.replace(/(<\/span>)/,'$1<span class="chg" id="c'+id+'">')+"</span>"}
    p.changed=p.ids.length>0;
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

  groups.forEach((g,gi)=>{
    if(g.eq){out.push({n:g.b+1,html:esc(pb[g.b]),changed:false});lastB=g.b;return}
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
  return {out,count:chgId,stats,paraCountA:pa.length,paraCountB:pb.length};
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
      +(s.moved?'<span class="chip mov">'+s.moved+' moved</span>':"");
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
    if(p.was)lines.push("Was  "+p.was);
    lines.push("Now  "+(p.now||"(removed)"),"");
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
      const [x,y]=await Promise.all(files.map(extract));
      const older=x.modified<=y.modified?[x,y]:[y,x];
      state.A=older[0];state.B=older[1];closePaste("A");closePaste("B");
    }else{
      state[side]=await extract(files[0]);closePaste(side);
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
    const v=ta.value;state[side]=v.trim()?{name:"Pasted text",kind:"Pasted",text:v,flow:false,modified:Date.now()}:null;showSlot(side);render()},250)});
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
  SS.download(changeList()+"\n\nMade with Rosina Stone, slidesize.com\n","changes-"+base+".txt","text/plain;charset=utf-8");
});
$("optDel").addEventListener("change",e=>document.body.classList.toggle("hide-del",!e.target.checked));
$("optCase").addEventListener("change",render);
$("optLB").addEventListener("change",render);
document.addEventListener("keydown",e=>{
  if(e.target.closest("textarea,input"))return;
  if(e.key==="n"||e.key==="j"){go(1);e.preventDefault()}
  else if(e.key==="p"||e.key==="k"){go(-1);e.preventDefault()}
});

/* ---------- Start with the example ---------- */
state.A={name:"awards-script-v3.txt",kind:"Example",text:EX_A,flow:false,example:true};
state.B={name:"awards-script-v4.txt",kind:"Example",text:EX_B,flow:false,example:true};
showSlot("A");showSlot("B");
setNotice('<span class="tag">Example</span> Two versions of a sample awards script are loaded. Drop your own files to replace them.');
render();

/* Files dropped on the Rosina Stone tile on the home page */
SS.consumeHandoff(files=>{clearExample();loadFiles(files,"A")});
})();
