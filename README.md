# PAP/DUE Backend - Vercel

Backend pour le générateur de documents PAP/DUE.

## 📁 Structure

```
.
├── api/
│   └── generate.js    ← API endpoint
└── package.json       ← Dependencies
```

## 🚀 Déploiement

1. Push ce dossier sur GitHub
2. Connecter sur vercel.com
3. Deploy

## 📝 Endpoint

Une fois déployé:
```
POST https://VOTRE-PROJET.vercel.app/api/generate
```

## 🔗 Configuration

Après le deployment, modifier `wizard-pap-due.html`:
```javascript
const response = await fetch('https://VOTRE-PROJET.vercel.app/api/generate', {
```

Voir `VERCEL_SUPER_FACIL.md` pour le guide complet.
