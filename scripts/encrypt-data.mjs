// ใช้: node scripts/encrypt-data.mjs   (อ่าน .xlsx ตัวแรกใน data-source/ → เขียน public/data/data.enc)
import { webcrypto as crypto } from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import readline from 'node:readline'

const dir = 'data-source'
const files = fs.existsSync(dir) ? fs.readdirSync(dir).filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$')) : []
if (!files.length) { console.error(`ไม่พบไฟล์ .xlsx ในโฟลเดอร์ ${dir}/`); process.exit(1) }
const src = path.join(dir, files[0])

const ask = q => new Promise(res => {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true })
  rl._writeToOutput = s => { if (s.includes(q)) rl.output.write(s); else rl.output.write('*') }   // ซ่อนรหัสที่พิมพ์
  rl.question(q, a => { rl.close(); console.log(); res(a) })
})
const pw = process.env.DASH_PASSWORD || await ask('ตั้งรหัสผ่านสำหรับเข้าดูแดชบอร์ด: ')
if (!pw) { console.error('ต้องใส่รหัสผ่าน'); process.exit(1) }

const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12))
const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey'])
const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 200000, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt'])
const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, fs.readFileSync(src)))
const out = Buffer.concat([Buffer.from('NSL1'), salt, iv, ct])
fs.mkdirSync('public/data', { recursive: true }); fs.writeFileSync('public/data/data.enc', out)
console.log(`✓ เข้ารหัส ${src} → public/data/data.enc (${(out.length / 1024).toFixed(0)} KB)`)
