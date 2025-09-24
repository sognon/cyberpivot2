
#!/usr/bin/env python3
"""Reset admin password in users_demo.yaml (bcrypt)."""
import argparse, yaml, bcrypt
from pathlib import Path
USERS_PATH = Path(__file__).parent / "users_demo.yaml"

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--email", required=True)
    p.add_argument("--password", required=True)
    a = p.parse_args()
    data = {}
    if USERS_PATH.exists():
        data = yaml.safe_load(USERS_PATH.read_text(encoding="utf-8")) or {}
    data.setdefault("credentials", {}).setdefault("usernames", {})
    key = a.email.strip().lower()
    entry = data["credentials"]["usernames"].get(key, {})
    entry["email"] = key
    entry.setdefault("name", key)
    entry["role"] = "admin"
    entry["password"] = bcrypt.hashpw(a.password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    data["credentials"]["usernames"][key] = entry
    USERS_PATH.write_text(yaml.safe_dump(data, sort_keys=False, allow_unicode=True), encoding="utf-8")
    print("OK - password set for", key)

if __name__ == "__main__":
    main()
