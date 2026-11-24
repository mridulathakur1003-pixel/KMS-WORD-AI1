"""Create licence keys using the hosted KMS WORD AI licensing API."""
import argparse, json, urllib.request
p = argparse.ArgumentParser()
p.add_argument("--url", required=True, help="Base URL, e.g. https://licenses.example.com")
p.add_argument("--admin-key", required=True)
p.add_argument("--count", type=int, default=10)
p.add_argument("--devices", type=int, default=1)
p.add_argument("--days", type=int, default=None)
p.add_argument("--notes", default="")
a = p.parse_args()
payload = {"count": a.count, "max_devices": a.devices, "notes": a.notes}
if a.days: payload["days_valid"] = a.days
req = urllib.request.Request(a.url.rstrip("/") + "/admin/licenses",
    data=json.dumps(payload).encode(),
    headers={"Content-Type":"application/json", "X-Admin-Key":a.admin_key},
    method="POST")
with urllib.request.urlopen(req, timeout=20) as r:
    print(json.dumps(json.loads(r.read()), indent=2))
