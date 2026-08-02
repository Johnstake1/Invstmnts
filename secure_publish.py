#!/usr/bin/env python3
"""
secure_publish.py  —  encrypt the portfolio & build a passphrase-protected dashboard
====================================================================================
Lets you host the dashboard on a PUBLIC url (e.g. GitHub Pages) while keeping your
data private: everything is AES-256-GCM encrypted with a passphrase, and only
ciphertext is ever published. The page asks for the passphrase and decrypts in the
browser (Web Crypto) — nothing sensitive leaves your device.

Passphrase comes from the env var  PORTFOLIO_PASSPHRASE  (a GitHub Actions secret in
the cloud; just set it in your shell locally). It is NEVER written to any file.

Commands
--------
  python secure_publish.py encrypt-store   # dataset.json      -> portfolio.enc   (do this once, commit portfolio.enc)
  python secure_publish.py decrypt-store    # portfolio.enc     -> dataset.json    (used by the CI job)
  python secure_publish.py build            # dataset.json + template -> index.html (encrypted, safe to publish)

Crypto: PBKDF2-HMAC-SHA256 (200k iters) -> AES-256-GCM. Matches the browser's
Web Crypto exactly, so index.html can decrypt what this script encrypts.
"""
import os, sys, json, base64, hashlib

HERE = os.path.dirname(os.path.abspath(__file__))
DATASET  = os.path.join(HERE, "dataset.json")
STORE    = os.path.join(HERE, "portfolio.enc")
TEMPLATE = os.path.join(HERE, "dashboard_template.html")
INDEX    = os.path.join(HERE, "index.html")
ITERS    = 200000

def _passphrase():
    pw = os.environ.get("PORTFOLIO_PASSPHRASE")
    if not pw:
        sys.exit("PORTFOLIO_PASSPHRASE is not set. Set it in your shell (or as a GitHub secret).")
    return pw.encode()

def _encrypt(plaintext_bytes):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt = os.urandom(16); iv = os.urandom(12)
    key = hashlib.pbkdf2_hmac("sha256", _passphrase(), salt, ITERS, dklen=32)
    ct = AESGCM(key).encrypt(iv, plaintext_bytes, None)          # ct = ciphertext || 16-byte tag
    return {"v": 1, "iters": ITERS,
            "salt": base64.b64encode(salt).decode(),
            "iv":   base64.b64encode(iv).decode(),
            "ct":   base64.b64encode(ct).decode()}

def _decrypt(payload):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt = base64.b64decode(payload["salt"]); iv = base64.b64decode(payload["iv"])
    ct   = base64.b64decode(payload["ct"])
    key  = hashlib.pbkdf2_hmac("sha256", _passphrase(), salt, payload.get("iters", ITERS), dklen=32)
    return AESGCM(key).decrypt(iv, ct, None)

# ---------------------------------------------------------------- overlay + boot
def _secure_html(payload):
    tpl = open(TEMPLATE, encoding="utf-8").read()
    # 1) data now comes from decryption at runtime, not an injected literal
    tpl = tpl.replace("let DATA = __DATA__;", "let DATA = window.__DATA;")
    # 2) turn the app <script> into an inert data block that runs only after unlock
    tpl = tpl.replace("<script>", '<script type="text/plain" id="__appcode">', 1)
    overlay = '''
<div id="__lock" style="position:fixed;inset:0;z-index:9999;background:#0d0d0d;color:#eee;
  display:flex;align-items:center;justify-content:center;font-family:system-ui,-apple-system,'Segoe UI',sans-serif">
  <form onsubmit="return __unlock(event)" style="text-align:center;max-width:340px;padding:24px">
    <div style="font-size:20px;font-weight:700;margin-bottom:6px">Investment Dashboard</div>
    <div style="font-size:13px;color:#9a9a9a;margin-bottom:18px">Enter your passphrase to decrypt</div>
    <input id="__pw" type="password" autocomplete="current-password" autofocus
      style="width:100%;padding:11px 12px;border-radius:10px;border:1px solid #333;background:#1a1a1a;color:#fff;font-size:15px">
    <button style="margin-top:12px;width:100%;padding:11px;border:0;border-radius:10px;background:#2a78d6;color:#fff;font-weight:600;font-size:15px;cursor:pointer">Unlock</button>
    <div id="__err" style="color:#e66767;font-size:12.5px;margin-top:10px;min-height:16px"></div>
    <div style="color:#6a6a6a;font-size:11px;margin-top:14px">Your data is AES-256 encrypted. It is decrypted only in this browser — nothing is sent anywhere.</div>
  </form>
</div>'''
    boot = '''
<script>
const __PAYLOAD = %s;
const __b = b => Uint8Array.from(atob(b), c => c.charCodeAt(0));
async function __unlock(e){
  e.preventDefault();
  const err = document.getElementById('__err'); err.textContent = 'Decrypting…';
  try{
    const pw = document.getElementById('__pw').value;
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
    const key = await crypto.subtle.deriveKey(
      {name:'PBKDF2', salt:__b(__PAYLOAD.salt), iterations:__PAYLOAD.iters, hash:'SHA-256'},
      base, {name:'AES-GCM', length:256}, false, ['decrypt']);
    const pt = await crypto.subtle.decrypt({name:'AES-GCM', iv:__b(__PAYLOAD.iv)}, key, __b(__PAYLOAD.ct));
    window.__DATA = JSON.parse(new TextDecoder().decode(pt));
    window.__PASSPHRASE = pw;   // kept in memory only, so the dashboard can re-encrypt on "Save encrypted file"
    const s = document.createElement('script'); s.textContent = document.getElementById('__appcode').textContent;
    document.body.appendChild(s);
    const lock = document.getElementById('__lock'); if(lock) lock.remove();
  }catch(_){ err.textContent = 'Wrong passphrase — try again.'; }
  return false;
}
</script>'''
    tpl = tpl.replace("<body>", "<body>\n" + overlay, 1)
    tpl = tpl.replace("</body>", boot % json.dumps(payload) + "\n</body>", 1)
    return tpl

def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "build"
    if cmd == "encrypt-store":
        json.dump(_encrypt(open(DATASET, "rb").read()), open(STORE, "w"), indent=1)
        print("wrote portfolio.enc (encrypted). Commit THIS, never dataset.json.")
    elif cmd == "decrypt-store":
        try:
            plaintext = _decrypt(json.load(open(STORE)))   # decrypt FIRST; a bad passphrase must not touch dataset.json
        except Exception:
            sys.exit("Could not decrypt portfolio.enc — wrong PORTFOLIO_PASSPHRASE? (dataset.json left untouched)")
        open(DATASET, "wb").write(plaintext)
        print("wrote dataset.json (decrypted) — transient; don't commit it.")
    elif cmd == "build":
        payload = _encrypt(open(DATASET, "rb").read())
        open(INDEX, "w", encoding="utf-8").write(_secure_html(payload))
        print("wrote index.html (passphrase-protected, safe to publish).")
    else:
        sys.exit(__doc__)

if __name__ == "__main__":
    main()
