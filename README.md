# Capitole Énergie — Import Odoo

Application de transformation de fichiers Salesforce → Odoo.

## Structure du projet

```
capitole-odoo/
├── api/
│   └── transform.py        ← Serverless function Vercel (endpoint POST /api/transform)
├── public/
│   └── index.html          ← Interface utilisateur (charte graphique 2026)
├── data/
│   └── Comptes_analytiques.xlsx   ← ⚠️ À placer ici (référentiel — non versionné)
├── transfo_odoo.py          ← Logique de transformation (inchangée)
├── requirements.txt         ← Dépendances Python
├── vercel.json              ← Configuration Vercel
└── .gitignore
```

## Déploiement sur Vercel

### 1. Prérequis
- Compte [Vercel](https://vercel.com)
- [Vercel CLI](https://vercel.com/docs/cli) installé : `npm i -g vercel`
- Dépôt GitHub créé

### 2. Placer le référentiel
Copier le fichier `Comptes_analytiques.xlsx` dans le dossier `data/` :
```
data/Comptes_analytiques.xlsx
```
> Ce fichier est listé dans `.gitignore` pour éviter de versionner des données sensibles.
> Sur Vercel, il sera inclus dans le build via la configuration ci-dessous.

### 3. Pousser sur GitHub
```bash
git init
git add .
git commit -m "feat: application import odoo capitole"
git remote add origin https://github.com/VOTRE_ORG/capitole-odoo.git
git push -u origin main
```

### 4. Connecter à Vercel
```bash
vercel
```
Ou via l'interface Vercel :
1. **New Project** → importer le dépôt GitHub
2. Framework : **Other**
3. Build & Output Settings : laisser par défaut
4. Deploy

### 5. Utilisation
- Accéder à l'URL Vercel générée
- Déposer un fichier `Salesforce.xlsx` avec l'onglet **"Import Odoo"**
- Cliquer sur **Lancer la transformation**
- Télécharger le fichier `Salesforce_transforme.xlsx`

## Notes techniques

- L'endpoint `/api/transform` accepte une requête `multipart/form-data` avec le champ `file`
- Le référentiel `Comptes_analytiques.xlsx` est intégré dans le bundle de déploiement
- Aucune donnée n'est persistée : tout est traité en mémoire
- Le sheet_name attendu est **"Import Odoo"** (modifié dans `transfo_odoo.py`)

## Modifier le référentiel sans redéployer

Si les comptes analytiques changent fréquemment, envisager de passer le référentiel
en variable d'environnement Vercel (stockage objet S3/R2) ou d'ajouter un second
champ de téléversement dans l'interface.
