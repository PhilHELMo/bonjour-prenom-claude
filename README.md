# Complément Outlook « Formule de politesse »

Dans un **nouveau message** (Outlook web / Microsoft 365), le volet lit le destinataire du champ À,
en déduit le prénom et insère :

    Bonjour Marie,


    Bien à toi        (ou « Bien à vous » en un clic)

## 1. Héberger les fichiers (HTTPS obligatoire)

Le complément est un site statique, sans donnée sensible. Le plus simple : GitHub Pages.

1. Crée un dépôt, envoie-y `taskpane.html`, `taskpane.js` et le dossier `assets/`.
2. Active GitHub Pages (Settings → Pages). Tu obtiens une adresse du type
   `https://TON-COMPTE.github.io/NOM-DU-DEPOT`.
   (Azure Static Web Apps, Netlify ou Cloudflare Pages conviennent aussi.)

## 2. Renseigner ton adresse dans le manifeste

    sed -i 's#https://VOTRE-HOTE#https://TON-COMPTE.github.io/NOM-DU-DEPOT#g' manifest.xml

(ou remplace « https://VOTRE-HOTE » à la main dans `manifest.xml`).

## 3. Installer

**Test personnel** : Outlook web → Applications / Obtenir des compléments → Mes compléments →
Compléments personnalisés → Ajouter à partir d'un fichier → `manifest.xml`.
(L'intitulé exact des menus peut varier ; l'option peut être bloquée par la politique du tenant.)

**Déploiement** : centre d'administration Microsoft 365 → Paramètres → Applications intégrées →
Charger des applications personnalisées → Complément Office → fichier manifeste, puis choisis
les utilisateurs. La propagation peut prendre de quelques minutes à plusieurs heures.

## 4. Utiliser

Nouveau message → ajoute le destinataire → bouton « Bonjour + formule » du ruban
(parfois sous « … » ou « Applications ») → **épingle le volet** pour qu'il reste ouvert
et se mette à jour quand tu changes de destinataire → « Insérer dans le message ».
