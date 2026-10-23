import assert from 'node:assert/strict';
const module=await import('./companions.mjs').catch(()=>({}));
assert.equal(typeof module.validateCompanions,'function','The optional companion catalogue must validate exact release metadata');
const {plannedCompanions,validateCompanions,companionReady}=module;
assert.equal(validateCompanions(plannedCompanions),plannedCompanions);
assert.equal(companionReady(plannedCompanions,'standalone'),false);
const complete=structuredClone(plannedCompanions);
complete.publishedAt='2026-10-23T12:00:00Z';
for(const kind of ['standalone','combined'])Object.assign(complete[kind],{sha256:'a'.repeat(64),sizeBytes:123});
assert.equal(companionReady(validateCompanions(complete),'combined'),true);
for(const change of [value=>value.releaseTag='../latest',value=>value.standalone.asset='wordplai-setup.exe',value=>value.combined.asset=value.standalone.asset,value=>value.standalone.sha256='unknown',value=>value.combined.sizeBytes=-1,value=>value.publishedAt='unknown']){
  const value=structuredClone(complete);change(value);assert.throws(()=>validateCompanions(value));
}
console.log('PASS: companion catalogue rejects unsafe tags, wrong/duplicate assets and invalid digest/size/date metadata.');
