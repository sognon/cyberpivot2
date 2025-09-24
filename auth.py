import bcrypt
import yaml, bcrypt
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
from copy import deepcopy

BASE_CONFIG_PATH = Path(__file__).parent / "auth_config.yaml"
USERS_PATH = Path(__file__).parent / "users_demo.yaml"

def _safe_load_yaml(path: Path) -> dict:
    if not path.exists(): return {}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}

def _merge_auth_configs(base_cfg: dict, users_cfg: dict) -> dict:
    cfg = deepcopy(base_cfg)
    cfg.setdefault("credentials", {}).setdefault("usernames", {})
    users_cfg.setdefault("credentials", {}).setdefault("usernames", {})
    cfg["credentials"]["usernames"].update(users_cfg["credentials"]["usernames"])
    return cfg

def _bcrypt_hash(pwd: str) -> str:
    return bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt(rounds=12)).decode("utf-8")

class Auth:
    def __init__(self, base_config_path: Path = BASE_CONFIG_PATH, users_path: Path = USERS_PATH):
        base_cfg = _safe_load_yaml(base_config_path)
        users_cfg = _safe_load_yaml(users_path)
        self.users_path = users_path
        self.config = _merge_auth_configs(base_cfg, users_cfg)
        # cookies par défaut si manquants
        self.config.setdefault("cookie", {}).update({
            "expiry_days": self.config.get("cookie",{}).get("expiry_days", 1),
            "key": self.config.get("cookie",{}).get("key", "cyberpivot_demo_cookie"),
            "name": self.config.get("cookie",{}).get("name", "cyberpivot_auth"),
        })
        self.authenticator = stauth.Authenticate(
            self.config["credentials"],
            self.config["cookie"]["name"],
            self.config["cookie"]["key"],
            self.config["cookie"]["expiry_days"],
        )

    def login(self):
        name, auth_status, username = self.authenticator.login("Se connecter", "sidebar")
        if auth_status:
            st.session_state["user"] = {
                "username": username,
                "name": name,
                "role": self.config["credentials"]["usernames"][username].get("role", "viewer"),
            }
            st.sidebar.success(f"Connecté : {name}")
        elif auth_status is False:
            st.sidebar.error("Identifiants invalides")
        else:
            st.sidebar.info("Veuillez vous connecter")
        return auth_status

    def logout(self):
        self.authenticator.logout("Se déconnecter", "sidebar")

    def require(self, permission: str):
        user = st.session_state.get("user")
        if not user: st.stop()
        role = user.get("role", "viewer")
        perms = self.config.get("roles", {}).get(role, {}).get("permissions", [])
        if permission not in perms:
            st.error("⛔ Accès refusé — permission manquante")
            st.stop()

    def create_user(self, email: str, name: str, password: str, role: str = "viewer") -> None:
        if role not in {"viewer", "auditor"}:
            raise ValueError("Rôle non autorisé (viewer/auditor).")
        users_cfg = _safe_load_yaml(self.users_path)
        users_cfg.setdefault("credentials", {}).setdefault("usernames", {})
        email_key = email.strip().lower()
        if email_key in users_cfg["credentials"]["usernames"] or \
           email_key in (k.lower() for k in self.config.get("credentials", {}).get("usernames", {}).keys()):
            raise ValueError("Un compte avec cet email existe déjà.")
        hashed = _bcrypt_hash(password)
        users_cfg["credentials"]["usernames"][email_key] = {
            "email": email_key, "name": name.strip(), "password": hashed, "role": role
        }
        tmp = self.users_path.with_suffix(".tmp.yaml")
        with open(tmp, "w", encoding="utf-8") as f:
            yaml.safe_dump(users_cfg, f, sort_keys=False, allow_unicode=True)
        tmp.replace(self.users_path)


# ======== Compatibility wrappers for app_cyberpivot.AuthAdapter ========
# L'app attend des fonctions module-level : init_auth_db, verify_password,
# get_user, get_or_create_user, create_user, login_get_user.
# Ces wrappers utilisent la classe Auth ci-dessus.
_auth_instance = None

def _get_auth_instance():
    global _auth_instance
    if _auth_instance is None:
        _auth_instance = Auth()
    return _auth_instance

def init_auth_db():
    """Aucune DB : on utilise des fichiers YAML. On s'assure que users_demo.yaml existe."""
    auth = _get_auth_instance()
    up = auth.users_path
    if not up.exists():
        up.write_text("credentials:\n  usernames: {}\n", encoding="utf-8")
    return True

def get_user(email: str):
    auth = _get_auth_instance()
    creds = auth.config.get("credentials", {}).get("usernames", {})
    return creds.get((email or '').strip().lower())

def create_user(email: str, password: str, full_name: str, role: str = "viewer", is_active: bool = True):
    """Compatibilité : mappe 'user' -> 'viewer', 'admin/auditor' -> 'auditor'."""
    mapped_role = "auditor" if role in ("auditor", "admin") else "viewer"
    auth = _get_auth_instance()
    try:
        auth.create_user(email=email, name=full_name, password=password, role=mapped_role)
        return True
    except Exception as e:
        return (False, str(e))

def get_or_create_user(email: str, full_name: str = None, role: str = "viewer"):
    u = get_user(email)
    if u:
        return u
    ok = create_user(email=email, password="changeme", full_name=full_name or email, role=role)
    return get_user(email) if ok else None

def verify_password(email: str, pwd: str) -> bool:
    """
    Retourne True si email/mot de passe sont valides.
    Si l'utilisateur n'existe pas encore, on le crée avec le mot de passe fourni (first-login).
    """
    from bcrypt import checkpw
    auth = _get_auth_instance()
    init_auth_db()
    email_key = (email or '').strip().lower()
    creds = auth.config.get("credentials", {}).get("usernames", {})
    user = creds.get(email_key)
    if not user:
        # Première connexion : on crée le compte avec le mot de passe saisi
        try:
            auth.create_user(email=email_key, name=email_key, password=pwd, role="auditor")
            # relecture
            auth = _get_auth_instance()
            creds = auth.config.get("credentials", {}).get("usernames", {})
            user = creds.get(email_key)
        except Exception:
            return False
    hashed = (user or {}).get("password", "")
    if not hashed:
        return False
    try:
        return checkpw(pwd.encode("utf-8"), hashed.encode("utf-8"))
    except Exception:
        return False

def login_get_user(email: str, pwd: str):
    """Parité avec AuthAdapter.login_get_user() côté app."""
    if not verify_password(email, pwd):
        return None
    return get_user(email)
# ======================================================================





def _load_users(users_yaml: Path | None = None) -> dict:
    p = Path(users_yaml) if users_yaml else USERS_PATH
    if not p.exists():
        return {"credentials": {"usernames": {}}}
    import yaml as _y
    with open(p, "r", encoding="utf-8") as f:
        data = _y.safe_load(f) or {}
    data.setdefault("credentials", {}).setdefault("usernames", {})
    return data

def _save_users(data: dict, users_yaml: Path | None = None) -> None:
    p = Path(users_yaml) if users_yaml else USERS_PATH
    p.parent.mkdir(parents=True, exist_ok=True)
    if p.exists():
        bak = p.with_suffix(p.suffix + ".bak")
        bak.write_bytes(p.read_bytes())
    import yaml as _y
    with open(p, "w", encoding="utf-8") as f:
        _y.safe_dump(data, f, sort_keys=False, allow_unicode=True)

def set_password(email: str, new_password: str, users_yaml: Path | None = None) -> bool:
    data = _load_users(users_yaml)
    key = (email or "").strip().lower()
    entry = data["credentials"]["usernames"].get(key, {})
    entry.setdefault("email", key)
    entry.setdefault("name", key)
    entry.setdefault("role", "admin")
    import bcrypt as _b
    entry["password"] = _b.hashpw(new_password.encode("utf-8"), _b.gensalt()).decode("utf-8")
    data["credentials"]["usernames"][key] = entry
    _save_users(data, users_yaml)
    return True

def set_user_role(email: str, role: str, users_yaml: Path | None = None) -> bool:
    data = _load_users(users_yaml)
    key = (email or "").strip().lower()
    if key not in data["credentials"]["usernames"]:
        return False
    data["credentials"]["usernames"][key]["role"] = role
    _save_users(data, users_yaml)
    return True
