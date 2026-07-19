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
/**
 * Exécute un appel API avec backoff exponentiel.
 */
function fetchAvecBackoff(apiCall) {
  let retryCount = 0;
  const maxRetries = 3;
  while (retryCount < maxRetries) {
    try {
      return apiCall();
    } catch (e) {
      retryCount++;
      if (retryCount >= maxRetries) throw new Error('Échec API après plusieurs tentatives : ' + e.message);
      Utilities.sleep(Math.pow(2, retryCount) * 1000);
    }
  }
}

function recupererEtClasserTaches() {
  const listesTaches = fetchAvecBackoff(() => Tasks.Tasklists.list());
  if (!listesTaches || !listesTaches.items || listesTaches.items.length === 0) {
    throw new Error('Aucune liste de tâches trouvée.');
  }
  
  let toutesLesTaches = [];
  
  // Parcourir toutes les listes de tâches de l'utilisateur
  listesTaches.items.forEach(liste => {
    const titreListe = liste.title.toLowerCase();
    const estListeStrategique = titreListe.includes('stratégie') || titreListe.includes('strategie') || titreListe.includes('important');
    
    let pageToken = null;
    do {
      const option = { showHidden: false, maxResults: 100 };
      if (pageToken) option.pageToken = pageToken;
      
      const response = fetchAvecBackoff(() => Tasks.Tasks.list(liste.id, option));
      
      if (response && response.items) {
        response.items.forEach(t => t._estListeStrategique = estListeStrategique);
        toutesLesTaches = toutesLesTaches.concat(response.items);
      }
      pageToken = response ? response.nextPageToken : null;
    } while (pageToken);
  });
  
  const lignesSheets = [];
  const stats = { Q1: 0, Q2: 0, Q3: 0, Q4: 0, total: 0, tasksQ1: [], tasksQ2: [], tasksQ3: [], tasksQ4: [] };
  const maintenant = new Date();
  
  toutesLesTaches.forEach(tache => {
    // Ignorer les tâches terminées (au cas où, bien que l'API filtre souvent avec d'autres params)
    if (tache.status === 'completed') return;
      
      const titre = tache.title || 'Sans titre';
      const notes = tache.notes || '';
      let dateEcheanceStr = '';
      let estUrgent = false;
      let estImportant = false;
      
      const contenuRecherche = (titre + ' ' + notes).toLowerCase();
      
      // Déterminer l'importance (Tags, Emojis, ou Liste)
      if (contenuRecherche.includes('#important') || contenuRecherche.includes('#strategie') || contenuRecherche.includes('⭐') || tache._estListeStrategique) {
        estImportant = true;
      }
      
      // Déterminer l'urgence (Tags, Emojis, ou Date)
      if (contenuRecherche.includes('#urgent') || contenuRecherche.includes('🔴')) {
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
      
      // Déterminer l'Effort (ICE Scoring)
      const estQuickWin = contenuRecherche.includes('#quickwin') || contenuRecherche.includes('⚡');
      const estLourd = contenuRecherche.includes('#lourd') || contenuRecherche.includes('🏋️');
      const effortTexte = estQuickWin ? 'Quick Win ⚡' : (estLourd ? 'Lourd 🏋️' : 'Normal');
      
      // Objet tâche pour l'email et le tri
      const objTache = { titre: titre, quickwin: estQuickWin, lourd: estLourd, due: tache.due };
      
      // Déduire le quadrant
      let quadrant = '';
      if (estUrgent && estImportant) {
        quadrant = 'Q1 (Crises)';
        stats.Q1++;
        stats.tasksQ1.push(objTache);
      } else if (!estUrgent && estImportant) {
        quadrant = 'Q2 (Stratégie)';
        stats.Q2++;
        stats.tasksQ2.push(objTache);
      } else if (estUrgent && !estImportant) {
        quadrant = 'Q3 (Bruit)';
        stats.Q3++;
        stats.tasksQ3.push(objTache);
      } else {
        quadrant = 'Q4 (Distractions)';
        stats.Q4++;
        stats.tasksQ4.push(objTache);
      }
      
      stats.total++;
      
      // Construire la ligne pour Sheets (avec la nouvelle colonne Effort)
      lignesSheets.push([
        titre,
        dateEcheanceStr,
        estUrgent ? 'Oui' : 'Non',
        estImportant ? 'Oui' : 'Non',
        effortTexte,
        quadrant,
        'En cours',
        notes
      ]);
  });
  
  // Trier les tâches pour l'email (Quick Wins en premier, Lourds à la fin)
  const triEffort = (a, b) => {
    if (a.quickwin && !b.quickwin) return -1;
    if (!a.quickwin && b.quickwin) return 1;
    if (a.lourd && !b.lourd) return 1;
    if (!a.lourd && b.lourd) return -1;
    return 0;
  };
  stats.tasksQ1.sort(triEffort);
  stats.tasksQ2.sort(triEffort);
  stats.tasksQ3.sort(triEffort);
  stats.tasksQ4.sort(triEffort);
  
  // Timeboxing Calendrier pour les tâches Q1 et Q2
  try {
    timeboxerTachesCritiques(stats.tasksQ1.concat(stats.tasksQ2));
  } catch (e) {
    console.warn("Erreur lors du timeboxing calendrier : ", e.message);
  }
  
  return { lignes: lignesSheets, stats: stats };
}

/**
 * Crée des événements "Focus" dans l'agenda pour les tâches urgentes/importantes dues aujourd'hui ou demain.
 */
function timeboxerTachesCritiques(tachesCritiques) {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const endOfTomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 2);
  
  const cal = CalendarApp.getDefaultCalendar();
  const events = cal.getEvents(today, endOfTomorrow);
  const eventTitles = events.map(e => e.getTitle());
  
  tachesCritiques.forEach(t => {
    if (t.due) {
      const dueDate = new Date(t.due);
      // Vérifier si la tâche est prévue pour aujourd'hui ou demain
      if (dueDate >= today && dueDate < endOfTomorrow) {
        const focusTitle = "🎯 Focus : " + t.titre;
        if (!eventTitles.includes(focusTitle)) {
          try {
            // Tentative de création d'un événement "Temps de concentration" (focusTime) 
            // via l'API Calendar avancée. Dispo surtout sur Workspace.
            const dateStr = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            const nextDate = new Date(dueDate.getTime() + 24 * 60 * 60 * 1000);
            const nextDateStr = Utilities.formatDate(nextDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            
            const eventPayload = {
              summary: focusTitle,
              eventType: 'focusTime',
              start: { date: dateStr },
              end: { date: nextDateStr }
            };
            Calendar.Events.insert(eventPayload, 'primary');
          } catch (e) {
            console.warn("Création focusTime a échoué (souvent car compte gratuit). Fallback sur événement classique :", e.message);
            cal.createAllDayEvent(focusTitle, dueDate);
          }
        }
      }
    }
  });
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
  
  // Chercher une feuille nommée "Tâches" ou la créer si elle n'existe pas
  let feuille = classeur.getSheetByName('Tâches');
  if (!feuille) {
    feuille = classeur.insertSheet('Tâches', 0);
  }
  
  // Valider et forcer les en-têtes (Ajout de la colonne Effort)
  const entetesAttendus = ['Titre de la tâche', 'Date d\'échéance', 'Urgent', 'Important', 'Effort', 'Quadrant', 'Statut', 'Notes'];
  
  // Sécurité : Étendre la feuille si elle n'a pas assez de colonnes (évite l'erreur "hors plage")
  if (feuille.getMaxColumns() < entetesAttendus.length) {
    feuille.insertColumnsAfter(feuille.getMaxColumns(), entetesAttendus.length - feuille.getMaxColumns());
  }
  
  feuille.getRange(1, 1, 1, entetesAttendus.length).setValues([entetesAttendus]);
  
  // Effacer les anciennes données à partir de la ligne 2
  const lastRow = feuille.getLastRow();
  const derniereColonne = Math.max(feuille.getLastColumn(), 8);
  
  if (lastRow >= 2) {
    const numRows = lastRow - 1;
    
    // Règle globale : Sauvegarde avant opération destructrice
    let archiveFeuille = classeur.getSheetByName('Archives');
    if (!archiveFeuille) {
      archiveFeuille = classeur.insertSheet('Archives');
    }
    const entetesArchives = ['Date Archivage', 'Titre', 'Échéance', 'Urgent', 'Important', 'Effort', 'Quadrant', 'Statut', 'Notes'];
    if (archiveFeuille.getMaxColumns() < entetesArchives.length) {
      archiveFeuille.insertColumnsAfter(archiveFeuille.getMaxColumns(), entetesArchives.length - archiveFeuille.getMaxColumns());
    }
    if (archiveFeuille.getLastRow() === 0) {
      archiveFeuille.appendRow(entetesArchives);
    }
    const valeursASauvegarder = feuille.getRange(2, 1, numRows, derniereColonne).getValues();
    const dateArchive = new Date();
    const titresNouveaux = lignes.map(row => row[0]); // Le titre est dans la première colonne
    // Archiver uniquement les lignes (tâches) qui ont disparu de la nouvelle liste
    const lignesArchive = valeursASauvegarder
      .filter(row => row.join('').trim() !== '')
      .filter(row => !titresNouveaux.includes(row[0]))
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
  
  // Nettoyage automatique des doublons dans l'onglet Archives
  const sheetArchives = classeur.getSheetByName('Archives');
  if (sheetArchives && sheetArchives.getLastRow() > 1) {
    // Supprime les doublons en ignorant la colonne Date Archivage (col 1).
    // On se base sur le Titre (2), Échéance (3), Effort (6) et Notes (9)
    sheetArchives.getDataRange().removeDuplicates([2, 3, 6, 9]);
  }
}
