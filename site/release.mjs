export const plannedRelease = Object.freeze({
  version: 'v0.1.0-preview.1', asset: 'wordplai-setup.exe',
  downloadUrl: 'https://github.com/GeorgeFejer91/wordplai/releases/download/v0.1.0-preview.1/wordplai-setup.exe',
  sha256: null, sizeBytes: null, publishedAt: null,
});

export function validateRelease(value) {
  if (!value || Object.keys(value).sort().join(',') !== Object.keys(plannedRelease).sort().join(',')) throw new Error('Invalid release metadata fields.');
  if (!/^v\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/.test(value.version) || value.asset !== 'wordplai-setup.exe' ||
      value.downloadUrl !== `https://github.com/GeorgeFejer91/wordplai/releases/download/${value.version}/${value.asset}`) throw new Error('Release URL must identify this repository and its explicit release asset.');
  if (value.sha256 !== null && !/^[a-fA-F0-9]{64}$/.test(value.sha256)) throw new Error('Invalid installer SHA-256.');
  if (value.sizeBytes !== null && (!Number.isSafeInteger(value.sizeBytes) || value.sizeBytes <= 0)) throw new Error('Invalid installer byte size.');
  if (value.publishedAt !== null && (typeof value.publishedAt !== 'string' || !Number.isFinite(Date.parse(value.publishedAt)))) throw new Error('Invalid publication time.');
  return value;
}

export const releaseReady = value => value.sha256 !== null && value.sizeBytes !== null && value.publishedAt !== null;
