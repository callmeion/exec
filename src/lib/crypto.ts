// รูปแบบไฟล์ data.enc = "NSL1" + salt(16) + iv(12) + AES-GCM ciphertext ; กุญแจมาจากรหัสผ่านด้วย PBKDF2-SHA256 (200,000 รอบ)
const te = new TextEncoder()
async function deriveKey(pw:string, salt:Uint8Array) {
  const base = await crypto.subtle.importKey('raw', te.encode(pw), 'PBKDF2', false, ['deriveKey'])
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt as unknown as BufferSource, iterations: 200000, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['decrypt'])
}
/** คืน ArrayBuffer ของไฟล์ xlsx ถ้ารหัสถูก; โยน error ถ้ารหัสผิด */
export async function decryptData(buf:ArrayBuffer, pw:string): Promise<ArrayBuffer> {
  const u = new Uint8Array(buf)
  if (new TextDecoder().decode(u.slice(0, 4)) !== 'NSL1') throw new Error('bad-format')
  return crypto.subtle.decrypt({ name: 'AES-GCM', iv: u.slice(20, 32) }, await deriveKey(pw, u.slice(4, 20)), u.slice(32))
}
