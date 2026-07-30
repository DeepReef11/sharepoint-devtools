// RFC 6238 TOTP, standard 30s/6-digit/SHA1 — what Entra expects for a
// "different authenticator app" enrolment. No dependencies.
const crypto = require('crypto');

function base32Decode(input) {
  const A = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  let bits = '';
  for (const ch of input.replace(/=+$/, '').replace(/\s+/g, '').toUpperCase()) {
    const v = A.indexOf(ch);
    if (v === -1) continue;
    bits += v.toString(2).padStart(5, '0');
  }
  const bytes = [];
  for (let i = 0; i + 8 <= bits.length; i += 8) bytes.push(parseInt(bits.slice(i, i + 8), 2));
  return Buffer.from(bytes);
}

function totp(secret, atMs = Date.now(), step = 30, digits = 6) {
  const key = base32Decode(secret);
  const counter = Math.floor(atMs / 1000 / step);
  const buf = Buffer.alloc(8);
  buf.writeBigUInt64BE(BigInt(counter));
  const mac = crypto.createHmac('sha1', key).update(buf).digest();
  const off = mac[mac.length - 1] & 0x0f;
  const bin = mac.readUInt32BE(off) & 0x7fffffff;
  return (bin % 10 ** digits).toString().padStart(digits, '0');
}

/** Seconds until the current code rolls over. */
function secondsRemaining(atMs = Date.now(), step = 30) {
  return step - Math.floor((atMs / 1000) % step);
}

module.exports = { totp, base32Decode, secondsRemaining };

if (require.main === module) {
  const secret = process.argv[2] || process.env.SP_TOTP_SECRET;
  if (!secret) {
    console.error('usage: node totp.js <base32-secret>');
    process.exit(1);
  }
  console.log(totp(secret));
}
