/**
 * @fileoverview Gestion des déclencheurs automatiques (Triggers).
 */

/**
 * Configure le déclencheur temporel pour lancer la synchronisation quotidiennement.
 * Supprime les anciens déclencheurs pour éviter les doublons.
 */
function installerDeclencheurQuotidien() {
  const nomFonction = 'synchroniserMatriceEisenhower';
  
  // 1. Supprimer les déclencheurs existants pour cette fonction
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === nomFonction) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 2. Créer le nouveau déclencheur (ex: tous les jours entre 7h et 8h)
  ScriptApp.newTrigger(nomFonction)
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
    
  console.log('Déclencheur quotidien installé avec succès pour 7h-8h.');
  
  // Notification optionnelle
  if (SpreadsheetApp.getActiveSpreadsheet()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Déclencheur configuré. Le script tournera tous les jours entre 7h et 8h.', 'Succès', 5);
  }
}
