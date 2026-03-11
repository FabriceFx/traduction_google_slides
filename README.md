# 🌐 Traducteur Google Slides

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

## 📝 Description
**Traducteur Google Slides** est un script Google Apps Script (GAS) complet qui ajoute une barre latérale personnalisée à vos présentations Google Slides. Il vous permet de traduire automatiquement et fidèlement tout le contenu textuel de votre présentation vers plus de 60 langues en utilisant le moteur natif de Google Translate.

## ✨ Fonctionnalités clés
- **Traduction intégrale** : Parcourt et traduit intelligemment les zones de texte classiques, les cellules de tableaux, ainsi que les notes du présentateur.
- **Copie de sécurité** : Utilise l'API Drive Avancée (v3) pour générer une nouvelle copie de la présentation avec le titre traduit, garantissant que votre document original reste intact.
- **Plage de slides personnalisée** : Permet de traduire l'intégralité du document ou seulement une sélection spécifique de diapositives (ex: `1, 3, 5-7`).
- **Détection automatique de la langue** : Identifie dynamiquement la langue source du document pour une traduction simplifiée.
- **Interface utilisateur moderne** : Barre latérale élégante développée avec les principes du Material Design 3 (M3).


## 🚀 Instructions d'installation manuelle

Pour installer ce script dans votre propre environnement Google Workspace, suivez ces étapes méthodiques :

1. **Ouvrir le projet Slides** : Créez une nouvelle présentation Google Slides ou ouvrez-en une existante.
2. **Accéder à l'éditeur de script** : Dans le menu supérieur, cliquez sur **Extensions** > **Apps Script**.
3. **Copier le code principal** : 
   - Renommez le fichier par défaut `Code.gs` (si nécessaire).
   - Collez l'intégralité du contenu de votre fichier local `Code.gs` dans l'éditeur.
4. **Créer l'interface utilisateur** :
   - Cliquez sur l'icône **+** à côté de "Fichiers" et sélectionnez **HTML**.
   - Nommez le fichier **strictement** `Sidebar` (l'extension `.html` s'ajoute automatiquement).
   - Collez le code de votre fichier `Sidebar.html` dans ce nouveau fichier.
5. **Activer le service avancé Google Drive** :
   - Sur la gauche de l'éditeur, cliquez sur le bouton **+** à côté de **Services**.
   - Cherchez **Google Drive API** et ajoutez-le (version v3, identifiant `Drive`).
6. **Configurer le Manifeste (appsscript.json)** :
   - Allez dans les paramètres du projet (l'icône engrenage ⚙️) et cochez *"Afficher le fichier manifeste 'appsscript.json' dans l'éditeur"*.
   - Retournez dans l'éditeur de code, ouvrez `appsscript.json` et remplacez son contenu par celui fourni dans votre projet (pour inclure les `oauthScopes` et les `dependencies`).
7. **Sauvegarder et Exécuter** :
   - Cliquez sur l'icône de disquette 💾 pour tout sauvegarder.
   - Fermez l'onglet Apps Script et rafraîchissez votre présentation Google Slides (F5).
   - Un nouveau menu **"🌐 Traducteur"** apparaîtra en haut. Lors de la première exécution, Google vous demandera d'autoriser les permissions d'accès.
