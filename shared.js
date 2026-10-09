/* shared.js — site-wide helpers
 *
 * Loaded by every page. Exposes window.SS (SlideSize) with the
 * common helpers so pages do not each re-declare them:
 *
 *   SS.fmtBytes(n)
 *   SS.download(blobOrBytes, filename, mime)
 *   SS.loadJSZip()
 *   SS.showToast(msg)
 *   SS.esc(str)
 *   SS.handoff.put(files)   -> saves dropped files for the next page to pick up
 *   SS.handoff.take()       -> promise, resolves to the files or null
 *
 * The handoff uses IndexedDB because sessionStorage cannot hold File blobs.
 */

(function(){
'use strict';

var SS = window.SS = window.SS || {};

/* ---------- byte formatting ---------- */
SS.fmtBytes = function(n) {
if (n < 1024) return n + ' B';
if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
if (n < 1073741824) return (n / 1048576).toFixed(2) + ' MB';
return (n / 1073741824).toFixed(2) + ' GB';
};

/* ---------- download trigger ---------- */
SS.download = function(data, filename, mime) {
var blob = (data instanceof Blob) ? data : new Blob([data], { type: mime || 'application/octet-stream' });
var url = URL.createObjectURL(blob);
var a = document.createElement('a');
a.href = url; a.download = filename;
document.body.appendChild(a); a.click(); a.remove();
setTimeout(function(){ URL.revokeObjectURL(url); }, 4000);
};

/* ---------- HTML escaping ---------- */
SS.esc = function(s) {
return String(s == null ? '' : s)
.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
};

/* ---------- JSZip lazy loader (cached across pages via browser cache) ---------- */
var JSZIP_SRC = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
var jszipReady = null;
SS.loadJSZip = function() {
if (jszipReady) return jszipReady;
jszipReady = new Promise(function(res, rej){
if (window.JSZip) { res(window.JSZip); return; }
var sc = document.createElement('script');
sc.src = JSZIP_SRC;
sc.onload = function(){ window.JSZip ? res(window.JSZip) : rej(new Error('JSZip not available')); };
sc.onerror = function(){ rej(new Error('JSZip failed to load')); };
document.head.appendChild(sc);
});
return jszipReady;
};

/* ---------- toast ---------- */
var toastTimer = null;
SS.showToast = function(msg) {
var t = document.getElementById('ss-toast');
if (!t) {
t = document.createElement('div');
t.id = 'ss-toast';
t.className = 'toast';
document.body.appendChild(t);
}
t.textContent = msg || 'Copied';
t.classList.add('visible');
if (toastTimer) clearTimeout(toastTimer);
toastTimer = setTimeout(function(){ t.classList.remove('visible'); }, 2600);
};

/* ---------- IndexedDB handoff for drag-and-drop from a tile ----------
 * When the user drops a file on a tile on the launchpad, the tile stashes
 * the File into IndexedDB and navigates to the tool page with ?pickup=1.
 * The tool page's init calls SS.handoff.take() which resolves to the
 * File list (and clears the stash) or null if nothing was waiting. */
var DB_NAME = 'slidesize-handoff';
var STORE = 'files';

function openDb() {
return new Promise(function(res, rej){
if (!window.indexedDB) { rej(new Error('IndexedDB not available')); return; }
var req = indexedDB.open(DB_NAME, 1);
req.onupgradeneeded = function(){ req.result.createObjectStore(STORE); };
req.onsuccess = function(){ res(req.result); };
req.onerror = function(){ rej(req.error); };
});
}
function txFiles(mode) {
return openDb().then(function(db){
return db.transaction(STORE, mode).objectStore(STORE);
});
}

SS.handoff = {
put: function(fileList) {
var files = Array.prototype.slice.call(fileList || []);
if (!files.length) return Promise.resolve(false);
/* Read each file to an ArrayBuffer, store ArrayBuffers in IDB. Real
   browsers also handle File/Blob through structured clone, but
   ArrayBuffer is a universal lowest-common-denominator and keeps the
   reconstruction logic the same everywhere. */
return Promise.all(files.map(function(f){
return f.arrayBuffer().then(function(ab){
return { bytes: ab, name: f.name, type: f.type || 'application/octet-stream', size: f.size, lastModified: f.lastModified || Date.now() };
});
})).then(function(items){
return txFiles('readwrite').then(function(store){
return new Promise(function(res, rej){
var payload = { items: items, at: Date.now() };
var req = store.put(payload, 'current');
req.onsuccess = function(){ res(true); };
req.onerror = function(){ rej(req.error); };
});
});
}).catch(function(){ return false; });
},
take: function() {
return txFiles('readwrite').then(function(store){
return new Promise(function(res){
var getReq = store.get('current');
getReq.onsuccess = function(){
var val = getReq.result;
/* Delete after reading so a refresh does not re-trigger */
store.delete('current');
if (val && val.items && val.items.length && Date.now() - val.at < 60000) {
/* Reconstruct File objects from the stored byte wrappers */
var out = val.items.map(function(it){
try { return new File([it.bytes], it.name, { type: it.type, lastModified: it.lastModified }); }
catch (e) {
/* Older browsers that cannot construct File this way get a Blob
   with .name tacked on, which every call site here can still use */
var b = new Blob([it.bytes], { type: it.type }); b.name = it.name; return b;
}
});
res(out);
} else {
res(null);
}
};
getReq.onerror = function(){ res(null); };
});
}).catch(function(){ return null; });
},
clear: function() {
return txFiles('readwrite').then(function(store){ store.delete('current'); }).catch(function(){});
}
};

/* ---------- Attach drag-and-drop pickup to a tile ----------
 * Call SS.wireTileDrop(linkElement, { accept: ['pptx', 'pdf'] }) to make a
 * tile accept drops of those extensions. On drop, it stashes the files and
 * navigates to the tile's href with ?pickup=1 appended. */
SS.wireTileDrop = function(el, opts) {
var accept = (opts && opts.accept) || null;
function isAccepted(name) {
if (!accept) return true;
var m = name.match(/\.([a-z0-9]+)$/i);
if (!m) return false;
return accept.indexOf(m[1].toLowerCase()) >= 0;
}
['dragenter','dragover'].forEach(function(ev){
el.addEventListener(ev, function(e){
if (!e.dataTransfer || !e.dataTransfer.types || e.dataTransfer.types.indexOf('Files') < 0) return;
e.preventDefault();
el.classList.add('drag');
});
});
['dragleave','drop'].forEach(function(ev){
el.addEventListener(ev, function(e){
e.preventDefault();
if (ev === 'drop' || !el.contains(e.relatedTarget)) el.classList.remove('drag');
});
});
el.addEventListener('drop', function(e){
e.preventDefault();
el.classList.remove('drag');
var fs = (e.dataTransfer && e.dataTransfer.files) || [];
if (!fs.length) return;
var usable = Array.prototype.slice.call(fs).filter(function(f){ return isAccepted(f.name); });
if (!usable.length) {
SS.showToast('This tool does not accept that file type');
return;
}
SS.handoff.put(usable).then(function(ok){
var href = el.getAttribute('href');
if (!href) return;
var sep = href.indexOf('?') < 0 ? '?' : '&';
window.location.href = href + (ok ? sep + 'pickup=1' : '');
});
});
};

/* ---------- Consume handoff on a tool page ----------
 * Call SS.consumeHandoff(handler) at the end of a tool page's init.
 * handler(files) is called if the URL has ?pickup=1 and the IndexedDB
 * stash has files. */
SS.consumeHandoff = function(handler) {
var qs = new URLSearchParams(window.location.search);
if (qs.get('pickup') !== '1') return;
SS.handoff.take().then(function(files){
if (files && files.length) handler(files);
/* Clean the ?pickup=1 off the URL so a refresh doesn't try again */
if (window.history && window.history.replaceState) {
var clean = window.location.pathname + window.location.hash;
window.history.replaceState({}, '', clean);
}
});
};

})();
