"""Creates release/version.json (read by the app's Update button)."""
import argparse
import hashlib
import json
import os

ap = argparse.ArgumentParser()
ap.add_argument("--version", required=True)
ap.add_argument("--installer", required=True)
ap.add_argument("--repo", required=True)  # owner/name
ap.add_argument("--notes", default="")
a = ap.parse_args()

name = os.path.basename(a.installer)
digest = hashlib.sha256(open(a.installer, "rb").read()).hexdigest()
manifest = {
    "version": a.version,
    "url": "https://github.com/%s/releases/download/v%s/%s" % (a.repo, a.version, name),
    "sha256": digest,
    "notes": a.notes.strip()[:300],
}
out = os.path.join(os.path.dirname(a.installer), "version.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(manifest, f, indent=2)
print(json.dumps(manifest, indent=2))
