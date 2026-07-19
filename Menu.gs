/**
 * @fileoverview Gestion du menu personnalisé et de la boîte de dialogue "À propos".
 */

/**
 * Fonction exécutée à l'ouverture du classeur.
 * Ajoute le menu personnalisé.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Eisenhower')
    .addItem('🔄 Synchroniser les tâches', 'synchroniserMatriceEisenhower')
    .addSeparator()
    .addItem('ℹ️ À propos', 'afficherAPropos')
    .addToUi();
}

/**
 * Affiche la boîte de dialogue "À propos" en utilisant Material Design 3.
 */
function afficherAPropos() {
  const htmlStr = `
    <!DOCTYPE html>
    <html lang="fr">
    <head>
      <base target="_top">
      <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Outfit', sans-serif;
          background-color: #f0f4f8;
          color: #1a1a1a;
          padding: 20px;
          margin: 0;
          display: flex;
          flex-direction: column;
          align-items: center;
        }
        .card {
          background: rgba(255, 255, 255, 0.7);
          backdrop-filter: blur(10px);
          border-radius: 16px;
          box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
          border: 1px solid rgba(255, 255, 255, 0.3);
          padding: 24px;
          max-width: 100%;
          text-align: center;
        }
        h2 {
          color: #0b5394;
          margin-top: 0;
        }
        p {
          font-size: 14px;
          line-height: 1.6;
        }
        .dev-info {
          margin-top: 20px;
          padding-top: 15px;
          border-top: 1px solid #e0e0e0;
          font-size: 13px;
        }
        a {
          color: #3c78d8;
          text-decoration: none;
          font-weight: 600;
        }
        a:hover {
          text-decoration: underline;
        }
        .btn {
          margin-top: 20px;
          background-color: #0b5394;
          color: white;
          border: none;
          padding: 10px 24px;
          border-radius: 20px;
          font-family: 'Outfit', sans-serif;
          font-weight: 600;
          cursor: pointer;
          transition: background-color 0.3s, transform 0.2s;
        }
        .btn:hover {
          background-color: #083c6b;
          transform: translateY(-2px);
        }
      </style>
    </head>
    <body>
      <div class="card">
        <h2>Matrice Eisenhower</h2>
        <p>Cet outil synchronise vos tâches Google Tasks pour les catégoriser selon la matrice d'Eisenhower.</p>
        <p>Gérez vos priorités, concentrez-vous sur la stratégie (Q2) et recevez un rapport quotidien directement dans votre boîte de réception.</p>
        <div class="dev-info">
          Développé par <strong>Fabrice Faucheux</strong><br>
          <a href="https://faucheux.bzh" target="_blank">https://faucheux.bzh</a>
        </div>
        <button class="btn" onclick="google.script.host.close()">Fermer</button>
      </div>
    </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlStr)
    .setWidth(400)
    .setHeight(320)
    .setTitle('À propos - Matrice Eisenhower');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'À propos');
}

/**
 * Affiche une boîte de dialogue d'erreur en Material Design 3.
 * @param {string} message Le message d'erreur.
 */
function afficherErreur(message) {
  const htmlStr = `
    <!DOCTYPE html>
    <html lang="fr">
    <head>
      <base target="_top">
      <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Outfit', sans-serif;
          background-color: #fce8e6;
          color: #c5221f;
          padding: 20px;
          margin: 0;
          display: flex;
          flex-direction: column;
          align-items: center;
        }
        .card {
          background: white;
          border-radius: 16px;
          box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
          border: 1px solid rgba(197, 34, 31, 0.3);
          padding: 24px;
          max-width: 100%;
          text-align: center;
        }
        h2 {
          color: #c5221f;
          margin-top: 0;
        }
        p {
          font-size: 14px;
          line-height: 1.6;
          color: #333;
        }
        .btn {
          margin-top: 20px;
          background-color: #c5221f;
          color: white;
          border: none;
          padding: 10px 24px;
          border-radius: 20px;
          font-family: 'Outfit', sans-serif;
          font-weight: 600;
          cursor: pointer;
          transition: background-color 0.3s, transform 0.2s;
        }
        .btn:hover {
          background-color: #a51d1a;
          transform: translateY(-2px);
        }
      </style>
    </head>
    <body>
      <div class="card">
        <h2>Erreur</h2>
        <p><?= message ?></p>
        <button class="btn" onclick="google.script.host.close()">Fermer</button>
      </div>
    </body>
    </html>
  `;
  
  const template = HtmlService.createTemplate(htmlStr);
  template.message = message;
  
  const htmlOutput = template.evaluate()
    .setWidth(400)
    .setHeight(250)
    .setTitle('Erreur');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Erreur');
}
