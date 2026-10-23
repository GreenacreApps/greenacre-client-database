import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {stripTypeScriptTypes} from 'node:module';
import {webcrypto} from 'node:crypto';
const root = new URL('../', import.meta.url);
const read = path => fs.readFileSync(new URL(path, root), 'utf8');
const api = read('netlify/functions/api.mts');
const app = read('app.js');
assert(!/test-restore|test-attachment-restore|isolated-test|greenacre-assets-test-20261022/.test(api));
assert(!read('index.html').includes('TEST ENVIRONMENT'));
assert(api.includes('const STORE_NAME = "greenacre-client-db"'));
assert(api.includes('return isProduction ? getStore(STORE_NAME, { consistency: "strong" }) : getDeployStore({ name: STORE_NAME })'));
assert(api.includes('return postReminderEmail(req)'));
const prod = new Map(), preview = new Map();
const calls = [];
const backend = map => ({get: async key => structuredClone(map.get(key) ?? null), setJSON: async (key,value) => map.set(key, structuredClone(value))});
const context = {Request,Response,URL,TextEncoder,Uint8Array,btoa,atob,crypto:webcrypto,console,
  Netlify:{context:{deploy:{context:'production'}},env:{get:()=>''}},
  getStore:(name)=>{calls.push(['production',name]);return backend(prod);},
  getDeployStore:({name})=>{calls.push(['preview',name]);return backend(preview);},
  XLSX:{},fetch:()=>{throw Error('Outbound access forbidden in release checks');}};
vm.createContext(context);
let executable = stripTypeScriptTypes(api.replace(/^import .*$/mg,''));
executable = executable.replace('export default async','const handler = async').replace('export const config','const config');
vm.runInContext(executable+'\nglobalThis.checkApi={handler,applyImportedAssetFields};',context);
const {handler,applyImportedAssetFields} = context.checkApi;
assert.equal((await handler(new Request('https://release.invalid/api/health'),{})).status,200);
assert.equal((await handler(new Request('https://release.invalid/api/data'),{})).status,401);
const token='release-test-fixture';
const hash=Buffer.from(await webcrypto.subtle.digest('SHA-256',new TextEncoder().encode(token))).toString('base64');
for(const map of [prod,preview]) {
  map.set('users',[{id:'fixture-user',username:'fixture',role:'user'}]);
  map.set('sessions/'+hash,{userId:'fixture-user',expiresAt:'2099-01-01T00:00:00Z'});
  map.set('data',{version:7,state:{clients:[{id:'client',sites:[{id:'site',assets:[{id:'asset',coverPhotoId:'photo-b',customData:{custom_job_numbers:'W1',custom_recent_project_numbers:'W2'}}]}]}]},assetCustomFields:[]});
}
const request=(method='GET',body)=>new Request('https://release.invalid/api/data',{method,headers:{cookie:'greenacre_session='+token,...(body?{'Content-Type':'application/json'}:{})},...(body?{body:JSON.stringify(body)}:{})});
const original=structuredClone(prod.get('data'));
assert.deepEqual(await (await handler(request(),{})).json(),original);
assert.deepEqual(prod.get('data'),original);
assert.equal((await handler(request('PUT',{...original,version:6}),{})).status,409);
assert.deepEqual(prod.get('data'),original);
const saved=await (await handler(request('PUT',original),{})).json();
assert.equal(saved.version,8);assert.deepEqual(saved.state,original.state);
assert.equal(preview.get('data').version,7);
context.Netlify.context.deploy.context='deploy-preview';calls.length=0;
assert.equal((await (await handler(request(),{})).json()).version,7);
assert(calls.every(([scope])=>scope==='preview'));
const asset={customData:{custom_site_contacts:'existing-history'}};const fields=[];
applyImportedAssetFields({'Site Contacts':'new-contact','Email':'test@example.invalid','Extra: Site Contacts':'new-contact','Fan Model':'Fan A'},asset,fields);
assert.equal(asset.customData.custom_site_contacts,'existing-history');assert.equal(asset.fanModel,'Fan A');assert(!fields.some(f=>/contact|email/i.test(f.label)));
const ui={assetCustomFields:[{id:'job',label:'Job Numbers / Project Numbers'},{id:'recent',label:'Project Numbers'},{id:'contacts',label:'Site Contacts'}]};vm.createContext(ui);
for(const name of ['cleanHeader','isContactField','isProjectNumberField','visibleAssetFields','projectNumbersForAsset','isPhotoAttachment','preferredAssetPhoto']) {
  const a=app.indexOf('function '+name+'('),b=app.indexOf('\n}',a)+2;assert(a>=0);vm.runInContext(app.slice(a,b),ui);
}
assert.equal(ui.visibleAssetFields(ui.assetCustomFields).length,1);
assert.equal(ui.projectNumbersForAsset({customData:{job:'W1',recent:'W2'}}),'W1, W2');
const photo=id=>({id,name:id+'.jpg',type:'image/jpeg',url:'/api/asset-attachments/asset/'+id});
const a=photo('a'),b=photo('b'),pdf={id:'doc.pdf',name:'doc.pdf',type:'application/pdf',url:'/api/asset-attachments/asset/doc.pdf'};
assert.equal(ui.preferredAssetPhoto([pdf,a,b],{coverPhotoId:'b'}).id,'b');
assert.equal(ui.preferredAssetPhoto([pdf,a],{coverPhotoId:'doc.pdf'}).id,'a');
assert.equal(ui.preferredAssetPhoto([pdf],{coverPhotoId:'doc.pdf'}),undefined);
assert.equal(ui.preferredAssetPhoto([a],{coverPhotoId:'deleted'}).id,'a');
assert(!ui.isPhotoAttachment({...a,name:'mislabeled.PDF'}));
console.log('PASS: production/preview storage isolation; authentication; conflict protection; exact roundtrip incl. cover/history; contact import exclusion; project values; PDF exclusions. No network or real storage used.');
