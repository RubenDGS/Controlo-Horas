const C='controlo-horas-v3-10-1-melhorias-recebimentos';
const A=['./','./index.html','./styles.css','./app.js?v=3.10.1','./manifest.webmanifest','./logotipo.png','./icon-192.png','./icon-512.png','./locations.json'];
const SHARED_CACHE='controlo-horas-shared-backup-v1';
const SHARED_KEY=new URL('./__shared_backup__.json',self.registration.scope).href;

self.addEventListener('message',e=>{if(e.data?.type==='SKIP_WAITING')self.skipWaiting()});

self.addEventListener('install',e=>{
 e.waitUntil(caches.open(C).then(c=>c.addAll(A)));
 self.skipWaiting();
});

self.addEventListener('activate',e=>{
 e.waitUntil(
  caches.keys().then(keys=>Promise.all(
   keys.filter(k=>k!==C && k!==SHARED_CACHE).map(k=>caches.delete(k))
  ))
 );
 self.clients.claim();
});

self.addEventListener('fetch',e=>{
 const u=new URL(e.request.url);

 // Receber um backup JSON enviado pelo menu Partilhar do Android.
 if(e.request.method==='POST' && u.pathname.endsWith('/share-target')){
  e.respondWith((async()=>{
   try{
    const form=await e.request.formData();
    const file=form.get('backup');
    if(!file || typeof file.text!=='function')throw new Error('ficheiro em falta');
    const text=await file.text();

    // Validar JSON antes de o guardar.
    JSON.parse(text);

    const c=await caches.open(SHARED_CACHE);
    await c.put(SHARED_KEY,new Response(text,{
     headers:{'Content-Type':'application/json;charset=utf-8'}
    }));

    return Response.redirect(new URL('./index.html?sharedBackup=1',self.registration.scope).href,303);
   }catch(err){
    return Response.redirect(new URL('./index.html?sharedBackupError=1',self.registration.scope).href,303);
   }
  })());
  return;
 }

 if(e.request.method!=='GET')return;

 if(u.origin===location.origin&&(u.pathname.endsWith('/app.js')||u.pathname.endsWith('/index.html')||u.pathname.endsWith('/'))){
  e.respondWith(
   fetch(e.request,{cache:'no-store'})
    .then(r=>{const x=r.clone();caches.open(C).then(c=>c.put(e.request,x));return r})
    .catch(()=>caches.match(e.request))
  );
  return;
 }

 e.respondWith(
  fetch(e.request)
   .then(r=>{const x=r.clone();caches.open(C).then(c=>c.put(e.request,x));return r})
   .catch(()=>caches.match(e.request).then(r=>r||caches.match('./index.html')))
 );
});
