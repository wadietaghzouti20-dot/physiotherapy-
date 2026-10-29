# RENDEVO — Cabinet de kinésithérapie

Frontend spécialisé adapté aux mobiles, en français, avec formulaire multi-étapes et motifs de consultation propres à la spécialité.

## Configuration avant mise en service
1. Ouvrir `config.js` et ajouter l'URL du déploiement Apps Script correspondant à ce cabinet.
2. Ajouter le nom public et le téléphone du cabinet si souhaité.
3. Vérifier que le backend traite l'action `availability` et les champs `service`, `preferredDate`, `preferredTime`, `fullName`, `phone`, `email`, `notes`, `privacyConsent`, `source`, `timezone`, `submittedAt`.
4. Tester la disponibilité, l'envoi, les emails, les doubles réservations et les rappels avant d'utiliser avec de vrais patients.

L'URL Apps Script est volontairement vide pour empêcher les demandes de cette spécialité d'être envoyées au tableur dentaire par erreur. Les emails nécessitent une configuration et des tests du backend.
