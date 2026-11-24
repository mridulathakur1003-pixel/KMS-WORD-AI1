"""
KMS WORD AI online activation API.
Deploy behind HTTPS. Set KMS_ADMIN_KEY and KMS_SIGNING_SECRET as environment variables.
"""
import os, sqlite3, hashlib, hmac, secrets, string, time, json
from datetime import datetime, timezone
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel, Field

DB_PATH = os.environ.get("KMS_LICENSE_DB", "licenses.sqlite3")
ADMIN_KEY = os.environ.get("KMS_ADMIN_KEY", "")
SIGNING_SECRET = os.environ.get("KMS_SIGNING_SECRET", "")
if not ADMIN_KEY or len(ADMIN_KEY) < 24:
    raise RuntimeError("Set KMS_ADMIN_KEY to a random secret of at least 24 characters.")
if not SIGNING_SECRET or len(SIGNING_SECRET) < 32:
    raise RuntimeError("Set KMS_SIGNING_SECRET to a random secret of at least 32 characters.")

app = FastAPI(title="KMS WORD AI Licensing API", version="1.0.0")

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS licenses (
            key_hash TEXT PRIMARY KEY, status TEXT NOT NULL DEFAULT 'active',
            max_devices INTEGER NOT NULL DEFAULT 1, expires_at INTEGER,
            created_at INTEGER NOT NULL, notes TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS activations (
            key_hash TEXT NOT NULL, device_id TEXT NOT NULL, activated_at INTEGER NOT NULL,
            PRIMARY KEY(key_hash, device_id)
        )""")
init_db()

class ActivateRequest(BaseModel):
    license_key: str = Field(min_length=8, max_length=80)
    device_id: str = Field(min_length=32, max_length=128)
    app_version: str = "20.0"

class ValidateRequest(BaseModel):
    token: str = Field(min_length=20, max_length=2000)
    device_id: str = Field(min_length=32, max_length=128)

class GenerateRequest(BaseModel):
    count: int = Field(default=1, ge=1, le=500)
    max_devices: int = Field(default=1, ge=1, le=20)
    days_valid: int | None = Field(default=None, ge=1, le=3650)
    notes: str = Field(default="", max_length=200)

def key_digest(key):
    return hashlib.sha256(key.strip().upper().encode()).hexdigest()

def sign(payload):
    raw = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode()
    sig = hmac.new(SIGNING_SECRET.encode(), raw, hashlib.sha256).hexdigest()
    return secrets.token_urlsafe(8) + "." + __import__("base64").urlsafe_b64encode(raw).decode().rstrip("=") + "." + sig

def verify(token):
    try:
        _, body, sig = token.split(".")
        raw = __import__("base64").urlsafe_b64decode(body + "=" * (-len(body) % 4))
        expected = hmac.new(SIGNING_SECRET.encode(), raw, hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return None
        return json.loads(raw)
    except Exception:
        return None

def require_admin(value):
    if not value or not hmac.compare_digest(value, ADMIN_KEY):
        raise HTTPException(401, "Admin authentication failed.")

@app.get("/health")
def health():
    return {"status": "ok", "service": "KMS WORD AI licensing"}

@app.post("/activate")
def activate(data: ActivateRequest):
    digest = key_digest(data.license_key)
    now = int(time.time())
    with db() as c:
        lic = c.execute("SELECT * FROM licenses WHERE key_hash=?", (digest,)).fetchone()
        if not lic or lic["status"] != "active":
            raise HTTPException(403, "Invalid or disabled product key.")
        if lic["expires_at"] and lic["expires_at"] < now:
            raise HTTPException(403, "This product key has expired.")
        existing = c.execute("SELECT 1 FROM activations WHERE key_hash=? AND device_id=?",
                             (digest, data.device_id)).fetchone()
        if not existing:
            count = c.execute("SELECT COUNT(*) FROM activations WHERE key_hash=?", (digest,)).fetchone()[0]
            if count >= lic["max_devices"]:
                raise HTTPException(403, "This key has reached its device activation limit.")
            c.execute("INSERT INTO activations VALUES(?,?,?)", (digest, data.device_id, now))
    payload = {"key_hash": digest, "device_id": data.device_id,
               "issued_at": now, "expires_at": lic["expires_at"]}
    return {"token": sign(payload),
            "activated_at": datetime.fromtimestamp(now, timezone.utc).isoformat()}

@app.post("/validate")
def validate(data: ValidateRequest):
    payload = verify(data.token)
    if not payload or payload.get("device_id") != data.device_id:
        return {"valid": False}
    now = int(time.time())
    if payload.get("expires_at") and payload["expires_at"] < now:
        return {"valid": False}
    with db() as c:
        lic = c.execute("SELECT status, expires_at FROM licenses WHERE key_hash=?",
                        (payload.get("key_hash"),)).fetchone()
        activation = c.execute("SELECT 1 FROM activations WHERE key_hash=? AND device_id=?",
                               (payload.get("key_hash"), data.device_id)).fetchone()
    valid = bool(lic and lic["status"] == "active" and activation and
                 (not lic["expires_at"] or lic["expires_at"] >= now))
    return {"valid": valid}

@app.post("/admin/licenses")
def generate_licenses(data: GenerateRequest, x_admin_key: str = Header(default="")):
    require_admin(x_admin_key)
    alphabet = string.ascii_uppercase + string.digits
    generated = []
    now = int(time.time())
    expires = now + data.days_valid * 86400 if data.days_valid else None
    with db() as c:
        for _ in range(data.count):
            key = "-".join("".join(secrets.choice(alphabet) for _ in range(4)) for _ in range(4))
            c.execute("INSERT INTO licenses(key_hash,status,max_devices,expires_at,created_at,notes) VALUES(?,?,?,?,?,?)",
                      (key_digest(key), "active", data.max_devices, expires, now, data.notes))
            generated.append(key)
    return {"created": len(generated), "keys": generated,
            "max_devices": data.max_devices, "expires_at": expires}

@app.post("/admin/revoke")
def revoke_license(license_key: str, x_admin_key: str = Header(default="")):
    require_admin(x_admin_key)
    with db() as c:
        cur = c.execute("UPDATE licenses SET status='revoked' WHERE key_hash=?", (key_digest(license_key),))
        if cur.rowcount == 0:
            raise HTTPException(404, "Key not found.")
    return {"revoked": True}
