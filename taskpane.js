/* Formule de politesse – complément Outlook (mode rédaction d'un nouveau message) */

const MOTS_GENERIQUES = new Set([
  'info', 'infos', 'contact', 'admin', 'administration', 'secretariat', 'secrétariat',
  'direction', 'service', 'support', 'noreply', 'no-reply', 'accueil', 'scolarite',
  'scolarité', 'rh', 'compta', 'comptabilite', 'bureau', 'mail', 'team', 'equipe', 'équipe'
]);

/* ---------- Extraction du prénom ---------- */

function capitaliser(mot) {
  return mot
    .toLowerCase()
    .split(/([-'’])/)
    .map(p => (/^[-'’]$/.test(p) ? p : p.charAt(0).toUpperCase() + p.slice(1)))
    .join('');
}

function estGenerique(mot) {
  return MOTS_GENERIQUES.has(mot.toLowerCase());
}

// "Marie Dupont", "Dupont, Marie", "DUPONT Marie", "Marie DUPONT (Service X)"
function prenomDepuisNom(nom) {
  let s = (nom || '').replace(/\(.*?\)|\[.*?\]/g, ' ').replace(/["“”]/g, '').trim();
  if (!s || s.includes('@')) return '';
  if (s.includes(',')) s = s.split(',').slice(1).join(' ').trim(); // "Dupont, Marie"
  const mots = s.split(/\s+/).filter(m => /\p{L}/u.test(m));
  if (!mots.length) return '';
  const estMajuscules = m => m.length > 1 && m === m.toUpperCase() && m !== m.toLowerCase();
  const candidats = mots.filter(m => !estMajuscules(m));
  const choisi = (candidats.length ? candidats : mots)[0];
  return estGenerique(choisi) ? '' : capitaliser(choisi);
}

// "marie.dupont@..." -> "Marie" ; "marie-claire_dupont@..." -> "Marie-Claire"
function prenomDepuisEmail(email) {
  const local = (email || '').split('@')[0].split('+')[0];
  const premier = local.split(/[._]/)[0].replace(/[\d]+/g, '');
  if (premier.length < 2 || estGenerique(premier)) return '';
  return capitaliser(premier);
}

// dest = { displayName, emailAddress } tel que fourni par Office.js
function prenomDestinataire(dest) {
  const nom = (dest.displayName || '').trim();
  const adresse = (dest.emailAddress || '').trim();
  if (nom && nom.toLowerCase() !== adresse.toLowerCase()) {
    const p = prenomDepuisNom(nom);
    if (p) return { prenom: p, deduitDeLAdresse: false };
  }
  const p = prenomDepuisEmail(adresse);
  return { prenom: p, deduitDeLAdresse: !!p };
}

function joindre(prenoms) {
  if (prenoms.length <= 1) return prenoms.join('');
  return prenoms.slice(0, -1).join(', ') + ' et ' + prenoms[prenoms.length - 1];
}

/* ---------- Construction du texte ---------- */

function echapperHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function salutation(prenoms) {
  const liste = joindre(prenoms.filter(Boolean));
  return liste ? `Bonjour ${liste},` : 'Bonjour,';
}

function construireHtml(prenoms, formule) {
  return (
    `<div>${echapperHtml(salutation(prenoms))}</div>` +
    '<div><br></div><div><br></div>' +
    `<div>${echapperHtml(formule)}</div>`
  );
}


/* ---------- Insertion sans lignes vides parasites avant la signature ---------- */

const SELECTEUR_SIGNATURE =
  '#Signature, #x_Signature, [id^="ms-outlook-mobile-signature"], [id^="x_ms-outlook-mobile-signature"]';

function estBlocVide(noeud) {
  if (noeud.nodeType === 3) return noeud.textContent.trim() === '';
  if (noeud.nodeType !== 1) return true; // commentaires, etc.
  return noeud.textContent.trim() === '' && !noeud.querySelector('img, table, hr, a');
}

/*
 * htmlCorps : corps actuel du message ; htmlInsere : notre bloc.
 * Renvoie le nouveau corps si le message est « vierge avec signature »
 * (uniquement des blocs vides avant la signature), sinon null :
 * l'appelant insère alors au curseur, sans toucher au reste.
 */
function fusionnerAvecSignature(htmlCorps, htmlInsere) {
  const doc = new DOMParser().parseFromString(htmlCorps, 'text/html');
  const signature = doc.body.querySelector(SELECTEUR_SIGNATURE);
  if (!signature) return null;

  const aSupprimer = [];
  let noeud = signature;
  while (noeud && noeud !== doc.body) {
    for (let p = noeud.previousSibling; p; p = p.previousSibling) {
      if (!estBlocVide(p)) return null; // du texte a déjà été saisi
      aSupprimer.push(p);
    }
    if (noeud.parentNode === doc.body) break;
    noeud = noeud.parentNode;
  }
  const racine = noeud; // enfant direct du body qui contient la signature
  aSupprimer.forEach(n => n.parentNode && n.parentNode.removeChild(n));

  const conteneur = doc.createElement('div');
  conteneur.innerHTML = htmlInsere;
  while (conteneur.firstChild) doc.body.insertBefore(conteneur.firstChild, racine);
  return doc.body.innerHTML;
}

/* ---------- Interface ---------- */

if (typeof Office !== 'undefined') {
  let destinataires = [];
  let modifieManuellement = false;

  const el = id => document.getElementById(id);
  const formuleChoisie = () => document.querySelector('input[name="formule"]:checked').value;
  const prenomsSaisis = () =>
    el('prenoms').value.split(/\s*(?:,|\bet\b|&)\s*/i).map(s => s.trim()).filter(Boolean);

  function dire(message, erreur) {
    const s = el('statut');
    s.textContent = message || '';
    s.className = erreur ? 'statut erreur' : 'statut';
  }

  function afficherApercu() {
    const prenoms = prenomsSaisis();
    el('apercu').textContent = `${salutation(prenoms)}\n\n\n${formuleChoisie()}`;
  }

  function appliquerDestinataires() {
    const resultats = destinataires.map(prenomDestinataire).filter(r => r.prenom);
    if (!modifieManuellement) {
      el('prenoms').value = joindre(resultats.map(r => r.prenom));
    }
    const note = el('note');
    if (!destinataires.length) {
      note.textContent = 'Ajoute un destinataire dans le champ À : son prénom apparaîtra ici.';
    } else if (resultats.some(r => r.deduitDeLAdresse)) {
      note.textContent = 'Prénom déduit de l’adresse e-mail : vérifie-le avant d’insérer.';
    } else {
      note.textContent = '';
    }
    el('reinit').hidden = !modifieManuellement;
    afficherApercu();
  }

  function lireDestinataires() {
    Office.context.mailbox.item.to.getAsync(res => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        destinataires = res.value || [];
        appliquerDestinataires();
      } else {
        dire('Impossible de lire les destinataires.', true);
      }
    });
  }

  function insererAuCurseur(html, texte) {
    const corps = Office.context.mailbox.item.body;
    corps.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, res => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        dire('Texte inséré dans le message.');
        return;
      }
      // Message en texte brut : on réessaie sans mise en forme
      corps.setSelectedDataAsync(texte, { coercionType: Office.CoercionType.Text }, res2 => {
        if (res2.status === Office.AsyncResultStatus.Succeeded) {
          dire('Texte inséré dans le message.');
        } else {
          dire('Insertion impossible : ' + res2.error.message, true);
        }
      });
    });
  }

  function inserer() {
    const prenoms = prenomsSaisis();
    const formule = formuleChoisie();
    const html = construireHtml(prenoms, formule);
    const texte = `${salutation(prenoms)}\n\n\n${formule}`;
    const corps = Office.context.mailbox.item.body;

    // Nouveau message avec signature : on place notre bloc tout en haut et on retire
    // les lignes vides que Outlook laisse entre le curseur et la signature.
    corps.getAsync(Office.CoercionType.Html, lecture => {
      if (lecture.status !== Office.AsyncResultStatus.Succeeded) {
        insererAuCurseur(html, texte);
        return;
      }
      let nouveauCorps = null;
      try { nouveauCorps = fusionnerAvecSignature(lecture.value, html); } catch (e) { /* repli */ }
      if (nouveauCorps === null) {
        insererAuCurseur(html, texte);
        return;
      }
      corps.setAsync(nouveauCorps, { coercionType: Office.CoercionType.Html }, res => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          dire('Texte inséré dans le message.');
        } else {
          insererAuCurseur(html, texte);
        }
      });
    });
  }

  Office.onReady(info => {
    if (info.host !== Office.HostType.Outlook) return;
    const item = Office.context.mailbox.item;
    if (!item || !item.to || typeof item.to.getAsync !== 'function') {
      dire('Ce volet fonctionne uniquement lors de la rédaction d’un message.', true);
      el('inserer').disabled = true;
      return;
    }

    el('inserer').addEventListener('click', inserer);
    document.querySelectorAll('input[name="formule"]').forEach(r =>
      r.addEventListener('change', () => { dire(''); afficherApercu(); }));
    el('prenoms').addEventListener('input', () => {
      modifieManuellement = true;
      dire('');
      appliquerDestinataires();
    });
    el('reinit').addEventListener('click', () => {
      modifieManuellement = false;
      appliquerDestinataires();
    });

    // Se met à jour quand les destinataires changent (volet épinglé)
    try {
      item.addHandlerAsync(Office.EventType.RecipientsChanged, lireDestinataires);
    } catch (e) { /* non bloquant */ }

    lireDestinataires();
  });
}

if (typeof module !== 'undefined') {
  module.exports = { prenomDepuisNom, prenomDepuisEmail, prenomDestinataire, joindre, salutation, construireHtml, fusionnerAvecSignature };
}
