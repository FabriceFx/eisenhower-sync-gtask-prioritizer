/**
 * @fileoverview Logique principale pour la synchronisation Google Tasks vers la Matrice d'Eisenhower.
 */

/**
 * Fonction principale déclenchée manuellement ou par trigger.
 * Synchronise les tâches Google, les classe et met à jour le Google Sheets.
 */
function synchroniserMatriceEisenhower() {
  try {
    const data = recupererEtClasserTaches();
    mettreAJourFeuille(data.lignes);
    envoyerRapportEmail(data.stats);
    
    // Notification de succès pour un lancement manuel
    if (SpreadsheetApp.getActiveSpreadsheet()) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Synchronisation terminée avec succès', 'Succès', 5);
    }
  } catch (e) {
    console.error('Erreur lors de la synchronisation :', e);
    // En cas d'exécution depuis l'interface Sheets, afficher une alerte propre (MD3)
    if (SpreadsheetApp.getActiveSpreadsheet()) {
      afficherErreur('Une erreur est survenue lors de la synchronisation : ' + e.message);
    }
  }
}

/**
 * Récupère les tâches non terminées depuis Google Tasks et les classe par quadrant.
 * @returns {Object} Un objet contenant les lignes pour Sheets et les statistiques pour l'email.
 */
function recupererEtClasserTaches() {
  const listesTaches = Tasks.Tasklists.list();
  if (!listesTaches.items || listesTaches.items.length === 0) {
    throw new Error('Aucune liste de tâches trouvée.');
  }
  
  // Utiliser la liste par défaut (souvent la première) ou '@default'
  const idListe = '@default'; 
  
  let taches;
  let retryCount = 0;
  const maxRetries = 3;
  let success = false;
  
  while (!success && retryCount < maxRetries) {
    try {
      try {
        taches = Tasks.Tasks.list(idListe, { showHidden: false });
      } catch (e) {
        // Si '@default' échoue, on prend la première liste disponible
        taches = Tasks.Tasks.list(listesTaches.items[0].id, { showHidden: false });
      }
      success = true;
    } catch (e) {
      retryCount++;
      if (retryCount >= maxRetries) throw new Error('Échec de la récupération des tâches après plusieurs tentatives : ' + e.message);
      Utilities.sleep(Math.pow(2, retryCount) * 1000); // Backoff exponentiel : 2s, 4s...
    }
  }
  
  const lignesSheets = [];
  const stats = { Q1: 0, Q2: 0, Q3: 0, Q4: 0, total: 0 };
  const maintenant = new Date();
  
  if (taches.items) {
    taches.items.forEach(tache => {
      // Ignorer les tâches terminées (au cas où, bien que l'API filtre souvent avec d'autres params)
      if (tache.status === 'completed') return;
      
      const titre = tache.title || 'Sans titre';
      const notes = tache.notes || '';
      let dateEcheanceStr = '';
      let estUrgent = false;
      let estImportant = false;
      
      // Déterminer l'importance
      if (notes.toLowerCase().includes('#important') || notes.toLowerCase().includes('#strategie')) {
        estImportant = true;
      }
      
      // Déterminer l'urgence
      if (notes.toLowerCase().includes('#urgent')) {
        estUrgent = true;
      } else if (tache.due) {
        const dateEcheance = new Date(tache.due);
        dateEcheanceStr = Utilities.formatDate(dateEcheance, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        
        // Différence en jours
        const diffTemps = dateEcheance.getTime() - maintenant.getTime();
        const diffJours = Math.ceil(diffTemps / (1000 * 3600 * 24));
        
        if (diffJours <= 3) {
          estUrgent = true;
        }
      }
      
      // Déduire le quadrant
      let quadrant = '';
      if (estUrgent && estImportant) {
        quadrant = 'Q1 (Crises)';
        stats.Q1++;
      } else if (!estUrgent && estImportant) {
        quadrant = 'Q2 (Stratégie)';
        stats.Q2++;
      } else if (estUrgent && !estImportant) {
        quadrant = 'Q3 (Bruit)';
        stats.Q3++;
      } else {
        quadrant = 'Q4 (Distractions)';
        stats.Q4++;
      }
      
      stats.total++;
      
      // Construire la ligne pour Sheets
      lignesSheets.push([
        titre,
        dateEcheanceStr,
        estUrgent ? 'Oui' : 'Non',
        estImportant ? 'Oui' : 'Non',
        quadrant,
        'En cours',
        notes
      ]);
    });
  }
  
  return { lignes: lignesSheets, stats: stats };
}

/**
 * Met à jour la feuille Google Sheets avec les nouvelles données.
 * @param {Array<Array>} lignes Les données à insérer.
 */
function mettreAJourFeuille(lignes) {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  if (!classeur) {
    console.log('Script exécuté hors classeur, mise à jour de la feuille ignorée si non liée.');
    return;
  }
  
  // Chercher une feuille nommée "Tâches" ou prendre la première
  let feuille = classeur.getSheetByName('Tâches');
  if (!feuille) {
    feuille = classeur.getSheets()[0];
  }
  
  // Effacer les anciennes données à partir de la ligne 2
  const lastRow = feuille.getLastRow();
  const derniereColonne = Math.max(feuille.getLastColumn(), 7);
  
  if (lastRow >= 2) {
    const numRows = lastRow - 1;
    
    // Règle globale : Sauvegarde avant opération destructrice
    let archiveFeuille = classeur.getSheetByName('Archives');
    if (!archiveFeuille) {
      archiveFeuille = classeur.insertSheet('Archives');
      archiveFeuille.appendRow(['Date Archivage', 'Titre', 'Échéance', 'Urgent', 'Important', 'Quadrant', 'Statut', 'Notes']);
    }
    const valeursASauvegarder = feuille.getRange(2, 1, numRows, derniereColonne).getValues();
    const dateArchive = new Date();
    // Filtrer les lignes vides et ajouter la date d'archivage
    const lignesArchive = valeursASauvegarder
      .filter(row => row.join('').trim() !== '')
      .map(row => [dateArchive].concat(row));
      
    if (lignesArchive.length > 0) {
      archiveFeuille.getRange(archiveFeuille.getLastRow() + 1, 1, lignesArchive.length, lignesArchive[0].length).setValues(lignesArchive);
    }

    feuille.getRange(2, 1, numRows, derniereColonne).clearContent();
  }
  
  // Si on a des données, les insérer en bloc (performance)
  if (lignes.length > 0) {
    feuille.getRange(2, 1, lignes.length, lignes[0].length).setValues(lignes);
  }
}
