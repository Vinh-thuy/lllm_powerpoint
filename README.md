# Générateur de Planning Interactif

## Prérequis
- Python 3.8+
- Ollama installé et configuré

## Installation
1. Cloner le dépôt
2. Créer un environnement virtuel
```bash
python3 -m venv venv
source venv/bin/activate
```

3. Installer les dépendances
```bash
pip install -r requirements.txt
```

## Utilisation
1. Lancer l'application
```bash
python app.py
```

2. Ouvrir un navigateur à l'adresse `http://localhost:5000`

## Exemples de Commandes
- Créer un projet : `Créer projet 'P1' du 15 mai au 29 décembre (couleur : vert)`
- Modifier un projet : `Modifier projet 'P1' pour prolonger jusqu'au 15 juin`
- Supprimer un projet : `Supprimer le projet 'P1'`

## Fonctionnalités
- Saisie de projets en langage naturel
- Prévisualisation en temps réel
- Téléchargement du PowerPoint généré
