export const plannedCompanions=Object.freeze({
  releaseTag:'companions-v0.1.0-preview.1',publishedAt:null,
  standalone:{asset:'secret-tunnel-lite-setup.exe',version:'5.0.0',sha256:null,sizeBytes:null},
  combined:{asset:'wordplai-and-secret-tunnel-setup.exe',version:'Wordplai preview.8 + Secret Tunnel Lite 5.0.0',sha256:null,sizeBytes:null},
});
export function validateCompanions(value){
  if(!value||Object.keys(value).sort().join(',')!=='combined,publishedAt,releaseTag,standalone'||!/^companions-v\d+\.\d+\.\d+-preview\.\d+$/.test(value.releaseTag))throw new Error('Invalid companion release identity.');
  if(value.publishedAt!==null&&(typeof value.publishedAt!=='string'||!Number.isFinite(Date.parse(value.publishedAt))))throw new Error('Invalid companion publication time.');
  for(const [kind,asset] of [['standalone','secret-tunnel-lite-setup.exe'],['combined','wordplai-and-secret-tunnel-setup.exe']]){
    const file=value[kind];
    if(!file||Object.keys(file).sort().join(',')!=='asset,sha256,sizeBytes,version'||file.asset!==asset||typeof file.version!=='string'||!file.version||file.version.length>100)throw new Error('Invalid companion asset.');
    if(file.sha256!==null&&!/^[a-f0-9]{64}$/.test(file.sha256))throw new Error('Invalid companion digest.');
    if(file.sizeBytes!==null&&(!Number.isSafeInteger(file.sizeBytes)||file.sizeBytes<=0))throw new Error('Invalid companion size.');
  }
  return value;
}
export const companionReady=(value,kind)=>value.publishedAt!==null&&value[kind].sha256!==null&&value[kind].sizeBytes!==null;
export const companionUrl=(value,kind)=>`https://github.com/GeorgeFejer91/wordplai/releases/download/${value.releaseTag}/${value[kind].asset}`;
