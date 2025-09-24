# CyberPivot — Pack complet

## Installation
```bash
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
```

## Comptes (YAML)
Crée/maj le compte admin (hash bcrypt dans `users_demo.yaml`) :
```bash
python3 reset_admin_pw.py --email admin@local --password "Demo#2025"
```

## Lancer
```bash
streamlit run app_cyberpivot_risk.py
```

### Normes
Importe un Excel/CSV/JSON via **Admin → Charger** (alias → `data/norms/<alias>.csv`).

### Risques ↔ Contrôles
- **Multiselect** (aucune saisie d'ID) pour lier les contrôles à chaque risque.
- Panneau **Contrôle → Risques** pour éditer le mapping inverse.

### Exports
- **DOCX** : inclut la section **Risques** automatiquement.
- **PPTX** : diapo titre propre (sans double en-tête) + **Top risques**.