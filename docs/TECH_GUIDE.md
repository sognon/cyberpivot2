# Guide Technique
- Auth: `streamlit_authenticator` + YAML (`users_demo.yaml`), bcrypt
- Rôles: admin/auditor/viewer (voir `auth_config.yaml`)
- Données: `data/` (norms, projects, uploads, exports, templates)
- Mapping risques↔contrôles: multiselect + reverse panel
- Exports: DOCX (section Risques), PPTX (Top risques)