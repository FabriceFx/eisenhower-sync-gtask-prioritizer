# Matrice Eisenhower - Google Tasks Sync 🔄

[English version below](#english-version)

## 🇫🇷 Description du projet
Ce script Google Apps Script automatise la synchronisation de vos Google Tasks vers un Google Sheets préconfiguré, en les classant selon la fameuse matrice d'Eisenhower.
L'outil analyse vos tâches (titres, dates d'échéance et notes) pour déterminer leur urgence et leur importance, puis met à jour le tableau et vous envoie un rapport visuel quotidien par email pour vous aider à vous concentrer sur la stratégie (Q2) et non sur le bruit.

**Mantra :** *"Ne confondez pas le bruit et la musique."*

## 🚀 Fonctionnalités
- Synchronisation automatique avec la liste par défaut de Google Tasks.
- Classement par quadrant (Q1, Q2, Q3, Q4) basé sur :
  - **Urgence** : Date d'échéance ≤ 3 jours OU présence du tag `#urgent` dans les notes.
  - **Importance** : Présence des tags `#important` ou `#strategie` dans les notes.
- Nettoyage et mise à jour automatique des données dans l'onglet `Tâches` avec **sauvegarde préalable** dans un onglet `Archives`.
- Envoi automatique d'un rapport email esthétique (HTML/CSS) résumant la journée.
- Alerte visuelle si le quadrant des "Crises" (Q1) dépasse celui de la "Stratégie" (Q2).
- Menu personnalisé Google Sheets avec interface "À propos" en Material Design 3.

## 🛠️ Prérequis
- Un compte Google Workspace ou Gmail standard.
- Un document Google Sheets contenant un onglet nommé `Tâches`. Les entêtes recommandées (Ligne 1) sont : `Titre de la tâche`, `Date d'échéance`, `Urgent`, `Important`, `Quadrant`, `Statut`, `Notes`.

## ⚙️ Installation & Configuration
1. **Copier le code** :
   - Ouvrez votre Google Sheets.
   - Allez dans `Extensions > Apps Script`.
   - Copiez l'ensemble des fichiers (`Code.gs`, `Email.gs`, `Menu.gs`, `Triggers.gs`, `TemplateEmail.html`).
   - Assurez-vous d'avoir copié le fichier `appsscript.json` (pour l'afficher dans l'éditeur : Paramètres > Cocher "Afficher le fichier manifeste appsscript.json").
2. **Activer l'API Tasks** :
   - Dans l'éditeur Apps Script, cliquez sur le "+" à côté de "Services".
   - Sélectionnez "Tasks API" et cliquez sur "Ajouter".
3. **Configurer le déclencheur (Trigger)** :
   - Vous pouvez soit exécuter la fonction `installerDeclencheurQuotidien` depuis le fichier `Triggers.gs`.
   - Soit aller dans l'horloge (Déclencheurs) sur la gauche de l'éditeur > "Ajouter un déclencheur" > Fonction : `synchroniserMatriceEisenhower`, Source : "Temporel", Type : "Quotidien", Heure : celle de votre choix.
4. **Première exécution** :
   - Lancez `onOpen` ou rafraîchissez votre Sheets pour faire apparaître le menu personnalisé `⚙️ Eisenhower`.
   - Cliquez sur `⚙️ Eisenhower > 🔄 Synchroniser les tâches`. Google vous demandera d'autoriser l'accès à vos tâches, emails et feuilles de calcul.

## 📄 Licence
Ce projet est sous licence MIT.

Développé par [Fabrice Faucheux](https://faucheux.bzh)

---

<a name="english-version"></a>
## 🇬🇧 Project Description
This Google Apps Script automates the synchronization of your Google Tasks to a pre-configured Google Sheet, categorizing them according to the famous Eisenhower matrix.
The tool analyzes your tasks (titles, due dates, and notes) to determine their urgency and importance. It updates the sheet and sends you a daily visual email report to help you focus on strategy (Q2) rather than noise.

**Mantra:** *"Do not confuse noise with music."*

## 🚀 Features
- Automatic sync with the default Google Tasks list.
- Quadrant categorization (Q1, Q2, Q3, Q4) based on:
  - **Urgency**: Due date ≤ 3 days OR `#urgent` tag in notes.
  - **Importance**: `#important` or `#strategie` tag in notes.
- Automatic clearing and updating of data in the `Tâches` tab with **prior backup** in an `Archives` tab.
- Automatic dispatch of an aesthetic HTML/CSS email report summarizing the day.
- Visual alert if the "Crises" quadrant (Q1) exceeds the "Strategy" (Q2) quadrant.
- Custom Google Sheets menu with a Material Design 3 "About" interface.

## 🛠️ Prerequisites
- A Google Workspace or standard Gmail account.
- A Google Sheets document containing a tab named `Tâches`. Recommended headers (Row 1) are: `Task Title`, `Due Date`, `Urgent`, `Important`, `Quadrant`, `Status`, `Notes`.

## ⚙️ Installation & Configuration
1. **Copy the code**:
   - Open your Google Sheets.
   - Go to `Extensions > Apps Script`.
   - Copy all the files (`Code.gs`, `Email.gs`, `Menu.gs`, `Triggers.gs`, `TemplateEmail.html`).
   - Make sure you copied the `appsscript.json` file (to display it in the editor: Settings > Check "Show appsscript.json manifest file").
2. **Enable Tasks API**:
   - In the Apps Script editor, click the "+" next to "Services".
   - Select "Tasks API" and click "Add".
3. **Configure the Trigger**:
   - You can either run the `installerDeclencheurQuotidien` function from the `Triggers.gs` file.
   - Or go to the clock icon (Triggers) on the left of the editor > "Add Trigger" > Function: `synchroniserMatriceEisenhower`, Source: "Time-driven", Type: "Day timer", Time: of your choice.
4. **First Execution**:
   - Run `onOpen` or refresh your Sheets to display the custom `⚙️ Eisenhower` menu.
   - Click `⚙️ Eisenhower > 🔄 Synchroniser les tâches`. Google will ask you to authorize access to your tasks, emails, and spreadsheets.

## 📄 License
This project is licensed under the MIT License.

Developed by [Fabrice Faucheux](https://faucheux.bzh)
