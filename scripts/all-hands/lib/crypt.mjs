import crypto from 'crypto';

// デッキの数字は AES-256-GCM で暗号化して HTML に埋める。
// public repo でソースを見られても平文の売上が出ないようにするための仕組み。
// blob = base64( salt(16) | iv(12) | authTag(16) | ciphertext )
// WebCrypto 側は ct||tag の順で受け取るので、ブラウザ側で並べ替えている。
const ITERATIONS = 250000;

function deriveKey(password, salt) {
  return crypto.pbkdf2Sync(password, salt, ITERATIONS, 32, 'sha256');
}

export function encrypt(obj, password) {
  const salt = crypto.randomBytes(16);
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', deriveKey(password, salt), iv);
  const ct = Buffer.concat([c.update(Buffer.from(JSON.stringify(obj), 'utf8')), c.final()]);
  return Buffer.concat([salt, iv, c.getAuthTag(), ct]).toString('base64');
}

export function decrypt(b64, password) {
  const buf = Buffer.from(b64, 'base64');
  const d = crypto.createDecipheriv('aes-256-gcm', deriveKey(password, buf.subarray(0, 16)), buf.subarray(16, 28));
  d.setAuthTag(buf.subarray(28, 44));
  return JSON.parse(Buffer.concat([d.update(buf.subarray(44)), d.final()]).toString('utf8'));
}
