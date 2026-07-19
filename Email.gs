/**
 * @fileoverview Gestion de la création et de l'envoi de l'email avec la matrice.
 */

/**
 * Génère et envoie le rapport par email.
 * @param {Object} stats Les statistiques des quadrants (Q1, Q2, Q3, Q4, total).
 */
function envoyerRapportEmail(stats) {
  if (stats.total === 0) {
    console.log('Aucune tâche à analyser, pas d\'email envoyé.');
    return;
  }
  
  // Calcul de la part du Q2
  const partQ2 = Math.round((stats.Q2 / stats.total) * 100);
  const alerteQ1 = stats.Q1 > stats.Q2;
  
  // Préparation du template
  const template = HtmlService.createTemplateFromFile('TemplateEmail');
  template.stats = stats;
  template.partQ2 = partQ2;
  template.alerteQ1 = alerteQ1;
  
  const htmlBody = template.evaluate().getContent();
  const emailUtilisateur = Session.getActiveUser().getEmail();
  
  // Envoi de l'email
  MailApp.sendEmail({
    to: emailUtilisateur,
    subject: 'Votre matrice Eisenhower - Rapport quotidien',
    htmlBody: htmlBody
  });
  
  console.log('Email envoyé à ' + emailUtilisateur);
}
