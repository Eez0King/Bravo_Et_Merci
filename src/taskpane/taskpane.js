// // Attendre que Office soit prêt
// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook) {
//     init();
//   }
// });

// // Initialisation
// function init() {
//   prefillNominateur();
//   document.getElementById("sendBtn").onclick = sendMessage;
// }

// // Pré-remplir les champs du formulaire
// function prefillNominateur() {
//   // showDebug("Pre-remplissage des champs...");
//   const item = Office.context.mailbox.item;

//   if (item && item.itemType === Office.MailboxEnums.ItemType.Message) {

//     const user = Office.context.mailbox.userProfile;
//     const displayName = user?.displayName || "Nom inconnu";
//     const emailAddress = user?.emailAddress || "email inconnu";

//     // showDebug("User connecté - Nom: " + displayName + ", Email: " + emailAddress);
//     fillForm(displayName, emailAddress);
//   } else {
//     const user = Office.context.mailbox.userProfile;
//     const displayName = user?.displayName || "Nom inconnu";
//     const emailAddress = user?.emailAddress || "email inconnu";

//     // showDebug("userProfile - Nom: " + displayName + ", Email: " + emailAddress);
//     fillForm(displayName, emailAddress);
//   }
// }

// // Remplit les champs du formulaire
// function fillForm(name, email) {
//   document.getElementById("nominateur-nom").innerText = name;
//   document.getElementById("nominateur-email").innerText = email;
//   document.getElementById("date-envoi").innerText = new Date().toLocaleDateString();
// }

// const EMAIL_REGEX = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;

// function isValidEmail(email) {
//   return EMAIL_REGEX.test(email);
// }

// function validateEmailArray(emails, label) {
//   const invalidEmails = emails.filter((e) => !isValidEmail(e));
//   if (invalidEmails.length) {
//     throw new Error(`Adresse(s) email invalide(s) dans ${label} : ${invalidEmails.join(", ")}`);
//   }
// }

// // Rassembler les données du formulaire
// function buildMessageBody() {
//   const messageTextElement = document.getElementById("message");
//   const messageText = messageTextElement ? messageTextElement.value : "";

//   const valeurs = Array.from(document.querySelectorAll(".valeur:checked")).map((cb) => cb.value);
//   const criteres = Array.from(document.querySelectorAll(".critere:checked")).map((cb) => cb.value);

//   let body = "";

//   if (valeurs.length) {
//     body += `<p><strong>Valeurs reconnues :</strong> ${valeurs.join(", ")}</p>`;
//   }

//   if (criteres.length) {
//     body += `<p><strong>Critères :</strong> ${criteres.join(", ")}</p>`;
//   }

//   if (messageText) {
//     body += `<p>${messageText}</p>`;
//   }

//   return body;
// }

// // Convertir une liste d'emails en objets EmailAddressDetails
// function toEmailAddressDetails(emails) {
//   return emails.map((email) => {
//     return {
//       emailAddress: email,
//       displayName: email,
//     };
//   });
// }

// // Envoi du message via le formulaire Outlook
// async function sendMessage() {
//   // Récupération et nettoyage des destinataires
//   const to = document
//     .getElementById("to")
//     .value.split(/[,;]+/)
//     .map((e) => e.trim())
//     .filter((e) => e);
//   const cc = document
//     .getElementById("cc")
//     .value.split(/[,;]+/)
//     .map((e) => e.trim())
//     .filter((e) => e);
//   const bcc = document
//     .getElementById("bcc")
//     .value.split(/[,;]+/)
//     .map((e) => e.trim())
//     .filter((e) => e);
//   const subject = document.getElementById("subject").value.trim();
//   const messageBody = buildMessageBody();

//   // Validation
//   const EMAIL_REGEX = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;

// function getInvalidEmails(emails) {
//   return emails.filter(email => !EMAIL_REGEX.test(email));
// }

//   const invalidEmails = [
//     ...getInvalidEmails(to),
//     ...getInvalidEmails(cc),
//     ...getInvalidEmails(bcc),
//   ];

//   if (invalidEmails.length) {
//     document.getElementById("error").innerText =
//       "Adresse(s) email invalide(s) : " + invalidEmails.join(", ");
//     return;
//   }

//   if (!to.length || !subject) {
//     document.getElementById("error").innerText =
//       "Veuillez remplir tous les champs obligatoires (destinataire et sujet).";
//     return;
//   }

//   // Effacer les erreurs précédentes
//   document.getElementById("error").innerText = "";

//   try {
//     // Envoyer à la base de données avant d'ouvrir le formulaire
//     await saveEmail();
//     // Préparer les paramètres du message avec le format EmailAddressDetails
//     const messageParams = {
//       toRecipients: toEmailAddressDetails(to),
//       subject: subject,
//       htmlBody: messageBody,
//     };

//     console.log("Tentative d'ouverture du formulaire avec:", messageParams);

//     // Méthode 1 : Essayer avec displayNewMessageForm
//     if (typeof Office.context.mailbox.displayNewMessageForm === "function") {
//       Office.context.mailbox.displayNewMessageForm(messageParams);
//       console.log("Formulaire ouvert avec displayNewMessageForm");
//     }
//     // Méthode 2 : Utiliser l'API REST pour composer le message
//     else if (Office.context.mailbox.item) {
//       composeNewMessage(to, cc, bcc, subject, messageBody);
//     }
//     // Méthode 3 : Fallback avec mailto
//     else {
//       const mailtoLink = createMailtoLink(to, cc, bcc, subject, messageBody);
//       window.open(mailtoLink, "_blank");
//       console.log("Formulaire ouvert avec mailto");
//     }
//   } catch (error) {
//     document.getElementById("error").innerText = "Erreur lors de l'envoi : " + error.message;
//     console.error("Erreur sendMessage:", error);
//   }
// }

// // Composer un nouveau message en utilisant l'API compose
// function composeNewMessage(to, cc, bcc, subject, body) {
//   Office.context.mailbox.item.subject.setAsync(subject);

//   Office.context.mailbox.item.to.setAsync(toEmailAddressDetails(to));

//   if (cc.length) {
//     Office.context.mailbox.item.cc.setAsync(toEmailAddressDetails(cc));
//   }

//   if (bcc.length) {
//     Office.context.mailbox.item.bcc.setAsync(toEmailAddressDetails(bcc));
//   }

//   Office.context.mailbox.item.body.setAsync(body, { coercionType: Office.CoercionType.Html });

//   console.log("Message composé avec succès");
// }

// // Fallback : Créer un lien mailto
// function createMailtoLink(to, cc, bcc, subject, body) {
//   let mailto = "mailto:" + to.join(",");
//   const params = [];

//   if (cc.length) params.push("cc=" + encodeURIComponent(cc.join(",")));
//   if (bcc.length) params.push("bcc=" + encodeURIComponent(bcc.join(",")));
//   if (subject) params.push("subject=" + encodeURIComponent(subject));
//   if (body) params.push("body=" + encodeURIComponent(body.replace(/<[^>]*>/g, "")));

//   if (params.length) {
//     mailto += "?" + params.join("&");
//   }

//   return mailto;
// }

// // Preparer les données à envoyer à la base de données
// function sendToDatabase() {
//   // Récupération des valeurs du formulaire
//   const to = document.getElementById("to").value.split(/[,;]+/).map(e => e.trim()).filter(e => e);
//   const cc = document.getElementById("cc").value.split(/[,;]+/).map(e => e.trim()).filter(e => e);
//   const bcc = document.getElementById("bcc").value.split(/[,;]+/).map(e => e.trim()).filter(e => e);
//   const subject = document.getElementById("subject").value.trim();
//   const message_body = document.getElementById("message")?.value || "";

//   // Récupération des checkboxes sélectionnées
//   const valeurs = Array.from(document.querySelectorAll(".valeur:checked")).map(cb => cb.value);
//   const criteres = Array.from(document.querySelectorAll(".critere:checked")).map(cb => cb.value);

//   // Récupération du nominateur
//   const nominateur_nom = document.getElementById("nominateur-nom")?.innerText || "";
//   const nominateur_email = document.getElementById("nominateur-email")?.innerText || "";
//   const date_envoi = document.getElementById("date-envoi")?.innerText || new Date().toLocaleDateString();

//   // Construction de l'objet transaction
//   const data = {
//     to,
//     cc,
//     bcc,
//     subject,
//     message_body,
//     valeurs,
//     criteres,
//     nominateur_nom,
//     nominateur_email,
//     // date_envoi
//   };

//   return data;
// }

// //Sauvegarde dans la base de données
// async function saveEmail() {
//   const data = sendToDatabase();

//   try {
//     const response = await fetch("http://127.0.0.1:8000/api/emails", {
//       method: "POST",
//       headers: {
//         "Content-Type": "application/json",
//         "Accept": "application/json"
//       },
//       body: JSON.stringify(data)
//     });

//     if (!response.ok) {
//       const err = await response.json();
//       console.error("Erreur API:", err);
//       return;
//     }

//     const result = await response.json();
//     console.log("Mail enregistré :", result);
//     alert("Mail enregistré avec succès !");
//   } catch (error) {
//     console.error("Erreur réseau:", error);
//   }
// }

/**
 * Bravo & Merci - Add-in Outlook
 * Script principal du formulaire de reconnaissance
 */

const CONFIG = {
  API_URL: "http://127.0.0.1:8000/api/emails",
  EMAIL_REGEX: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
  RH_EMAIL: "bravo-merci@entreprise.com",
};

// ========================================
// Initialisation
// ========================================
let isInitialized = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook && !isInitialized) {
    isInitialized = true;
    init();
  }
});

function init() {
  prefillNominateur();
  attachEventHandlers();
}

function attachEventHandlers() {
  const sendBtn = document.getElementById("sendBtn");

  // Retirer les anciens listeners avant d'en ajouter
  if (sendBtn) {
    const newBtn = sendBtn.cloneNode(true);
    sendBtn.parentNode.replaceChild(newBtn, sendBtn);
    newBtn.addEventListener("click", handleSendMessage);
  }

  // Effacer les erreurs lors de la saisie
  const inputs = document.querySelectorAll("input, textarea");
  inputs.forEach((input) => {
    input.addEventListener("input", clearMessages);
  });
}

// ========================================
// Pré-remplissage du formulaire
// ========================================
function prefillNominateur() {
  try {
    const user = Office.context.mailbox.userProfile;
    const displayName = user?.displayName || "Nom inconnu";
    const emailAddress = user?.emailAddress || "email@inconnu.com";

    fillNominateurFields(displayName, emailAddress);
  } catch (error) {
    console.error("Erreur lors du pré-remplissage:", error);
    showError("Impossible de récupérer les informations utilisateur.");
  }
}

function fillNominateurFields(name, email) {
  document.getElementById("nominateur-nom").innerText = name;
  document.getElementById("nominateur-email").innerText = email;
  document.getElementById("date-envoi").innerText = new Date().toLocaleDateString("fr-FR");
}

// ========================================
// Validation des données
// ========================================
function isValidEmail(email) {
  return CONFIG.EMAIL_REGEX.test(email);
}

function getInvalidEmails(emails) {
  return emails.filter((email) => !isValidEmail(email));
}

function parseEmailField(fieldId) {
  const value = document.getElementById(fieldId)?.value || "";
  return value
    .split(/[,;]+/)
    .map((e) => e.trim())
    .filter((e) => e);
}

function validateForm() {
  const errors = [];

  // Récupération des champs
  const to = parseEmailField("to");
  const cc = parseEmailField("cc");
  const bcc = parseEmailField("bcc");
  const subject = document.getElementById("subject")?.value.trim() || "";
  const message = document.getElementById("message")?.value.trim() || "";
  const valeurs = getCheckedValues(".valeur");
  const criteres = getCheckedValues(".critere");

  // Validation des champs obligatoires
  if (!to.length) {
    errors.push("Le champ 'Destinataire(s)' est obligatoire.");
  }

  if (!subject) {
    errors.push("Le champ 'Objet' est obligatoire.");
  }

  if (!valeurs.length) {
    errors.push("Veuillez sélectionner au moins une valeur reconnue.");
  }

  if (!criteres.length) {
    errors.push("Veuillez sélectionner au moins un critère de reconnaissance.");
  }

  if (!message) {
    errors.push("Le champ 'Message' est obligatoire.");
  }

  // Validation des emails
  const allEmails = [...to, ...cc, ...bcc];
  const invalidEmails = getInvalidEmails(allEmails);

  if (invalidEmails.length) {
    errors.push(`Adresse(s) email invalide(s) : ${invalidEmails.join(", ")}`);
  }

  return {
    isValid: errors.length === 0,
    errors,
    data: { to, cc, bcc, subject, message, valeurs, criteres },
  };
}

function getCheckedValues(selector) {
  return Array.from(document.querySelectorAll(`${selector}:checked`)).map((cb) => cb.value);
}

// ========================================
// Construction du message
// ========================================
function buildMessageBody(subject = "", message = "", valeurs = [], criteres = []) {
  const logoUrl = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRg-Me-63CZ3XXkow-vhNX4xV_lsFk-gYIxzA&s";

  const dateEnvoi = new Date().toLocaleDateString("fr-FR");
  let nomNominateur = "";
  let emailNominateur = "";

  if (typeof Office !== "undefined" && Office.context && Office.context.mailbox) {
    const profile = Office.context.mailbox.userProfile;
    if (profile) {
      nomNominateur = profile.displayName || nomNominateur;
      emailNominateur = profile.emailAddress || emailNominateur;
    }
  }
  return `
<!doctype html>
<html lang="fr">
  <head>
    <meta charset="UTF-8" />
    <title>Bravo & Merci</title>
  </head>

  <body style="margin:0;padding:0;background-color:#fff;">
    <div style="width:100%;padding:24px 0;background-color:#f3f2f1;">

      <!-- Container -->
      <div style="
        max-width:600px;
        margin:0 auto;
        background-color:#ffffff;
        border-radius:4px;
        padding:28px;
        font-family:Segoe UI, Arial, sans-serif;
        color:#1f1f1f;
      ">

        <!-- Header -->
        <div style="margin-bottom:24px;">
          <img src="${logoUrl}" width="160" alt="SGCI" />
        </div>

        <!-- Meta informations -->
        <div style="
          font-size:13px;
          line-height:18px;
          color:#4a4a4a;
          margin-bottom:24px;
          padding-left:12px;
          border-left:3px solid #e60028;
        ">
          <div><strong>Date d’envoi :</strong> ${dateEnvoi}</div>
          <div><strong>Nom du nominateur :</strong> ${nomNominateur}</div>
          <div><strong>Email du nominateur :</strong> ${emailNominateur}</div>
        </div>

        <!-- Objet -->
        <div style="margin-bottom:20px;">
          <div style="
            font-size:12px;
            text-transform:uppercase;
            letter-spacing:0.4px;
            color:#605e5c;
            margin-bottom:6px;
          ">
            Objet
          </div>
          <div style="
            font-size:17px;
            font-weight:600;
            line-height:22px;
            color:#000000;
          ">
           ${ subject }
          </div>
        </div>

        <!-- Divider -->
        <div style="height:1px;background-color:#edebe9;margin:24px 0;"></div>

        <!-- Valeurs -->
        ${
          valeurs.length
            ? `
        <div style="margin-bottom:20px;">
          <div style="
            font-size:14px;
            font-weight:600;
            margin-bottom:8px;
            color:#000000;
          ">
            Valeur Société Générale reconnue
          </div>
          ${valeurs
            .map(
              (v) => `
          <span style="
            display:inline-block;
            background-color:#fde7ea;
            color:#a4001d;
            padding:5px 12px;
            border-radius:14px;
            font-size:12px;
            margin:4px 6px 0 0;
          ">
            ${v}
          </span>`
            )
            .join("")}
        </div>`
            : ""
        }

        <!-- Critères -->
        ${
          criteres.length
            ? `
        <div style="margin-bottom:24px;">
          <div style="
            font-size:14px;
            font-weight:600;
            margin-bottom:8px;
            color:#000000;
          ">
            Critère(s) de reconnaissance
          </div>
          ${criteres
            .map(
              (c) => `
          <span style="
            display:inline-block;
            background-color:#f2f2f2;
            color:#1f1f1f;
            padding:5px 12px;
            border-radius:14px;
            font-size:12px;
            margin:4px 6px 0 0;
          ">
            ${c}
          </span>`
            )
            .join("")}
        </div>`
            : ""
        }

        <!-- Message -->
        <div style="margin-bottom:24px;">
          <div style="
            font-size:14px;
            font-weight:600;
            margin-bottom:8px;
            color:#000000;
          ">
            Message de reconnaissance
          </div>

          <div style="
            font-size:15px;
            line-height:22px;
            color:#1f1f1f;
          ">
            ${message.replace(/\n/g, "<br>")}
          </div>
        </div>

        <!-- Footer -->
        <div style="
          border-top:1px solid #edebe9;
          padding-top:12px;
          font-size:11px;
          color:#6b6b6b;
          text-align:left;
        ">
          Message envoyé via le programme de reconnaissance interne SGCI.
        </div>

      </div>
    </div>
  </body>
</html>
  `;
}

// ========================================
// Gestion de l'envoi
// ========================================
let isSending = false;

async function handleSendMessage(event) {
  event.preventDefault();

  // Bloquer si déjà en cours d'envoi
  if (isSending) {
    return;
  }

  isSending = true;

  // Nettoyer les messages précédents
  clearMessages();

  // Validation
  const validation = validateForm();

  if (!validation.isValid) {
    showError(validation.errors.join("<br>"));
    isSending = false;
    return;
  }

  // Afficher le loader
  setLoading(true);

  try {
    const { to, cc, bcc, subject, message, valeurs, criteres } = validation.data;

    // Ajouter l'email RH en CCI automatiquement
    const bccWithRH = [...bcc];
    if (CONFIG.RH_EMAIL && !bccWithRH.includes(CONFIG.RH_EMAIL)) {
      bccWithRH.push(CONFIG.RH_EMAIL);
    }

    // Construire le corps du message
    htmlBody = buildMessageBody(subject, message, valeurs, criteres);

    // Préparer les données pour la BDD
    const dbData = prepareDataForDatabase(to, cc, bccWithRH, subject, message, valeurs, criteres);

    // Ouvrir le formulaire de composition Outlook
    const messageParams = {
      toRecipients: toEmailAddressDetails(to),
      ccRecipients: cc.length ? toEmailAddressDetails(cc) : undefined,
      bccRecipients: bccWithRH.length ? toEmailAddressDetails(bccWithRH) : undefined,
      subject: subject,
      htmlBody: htmlBody,
    };

    // Envoyer à Outlook
    await openOutlookCompose(messageParams);

    // Sauvegarder en BDD (après ouverture du formulaire)
    await saveToDatabase(dbData);

    // Succès
    showSuccess("Reconnaissance envoyée avec succès ! Le message a été sauvegardé.");

    // Réinitialiser le formulaire après 2 secondes
    setTimeout(() => {
      resetForm();
      isSending = false;
    }, 2000);
  } catch (error) {
    showError(`Erreur lors de l'envoi : ${error.message}`);
    isSending = false;
  } finally {
    setLoading(false);
  }
}

// ========================================
// Interaction avec Outlook
// ========================================
function toEmailAddressDetails(emails) {
  return emails.map((email) => ({
    emailAddress: email,
    displayName: email,
  }));
}

async function openOutlookCompose(messageParams) {
  return new Promise((resolve, reject) => {
    try {
      // Méthode 1 : displayNewMessageForm (préférée)
      if (typeof Office.context.mailbox.displayNewMessageForm === "function") {
        Office.context.mailbox.displayNewMessageForm(messageParams);
        resolve();
      }
      // Méthode 2 : Composer dans le contexte actuel
      else if (Office.context.mailbox.item) {
        composeInCurrentContext(messageParams);
        resolve();
      }
      // Méthode 3 : Fallback mailto
      else {
        const mailtoUrl = createMailtoUrl(messageParams);
        window.open(mailtoUrl, "_blank");
        resolve();
      }
    } catch (error) {
      reject(new Error(`Impossible d'ouvrir le formulaire Outlook : ${error.message}`));
    }
  });
}

function composeInCurrentContext(params) {
  const item = Office.context.mailbox.item;

  if (params.subject) {
    item.subject.setAsync(params.subject);
  }

  if (params.toRecipients) {
    item.to.setAsync(params.toRecipients);
  }

  if (params.ccRecipients) {
    item.cc.setAsync(params.ccRecipients);
  }

  if (params.bccRecipients) {
    item.bcc.setAsync(params.bccRecipients);
  }

  if (params.htmlBody) {
    item.body.setAsync(params.htmlBody, { coercionType: Office.CoercionType.Html });
  }
}

function createMailtoUrl(params) {
  const to = params.toRecipients?.map((r) => r.emailAddress).join(",") || "";
  let mailto = `mailto:${to}`;
  const queryParams = [];

  if (params.ccRecipients?.length) {
    const cc = params.ccRecipients.map((r) => r.emailAddress).join(",");
    queryParams.push(`cc=${encodeURIComponent(cc)}`);
  }

  if (params.bccRecipients?.length) {
    const bcc = params.bccRecipients.map((r) => r.emailAddress).join(",");
    queryParams.push(`bcc=${encodeURIComponent(bcc)}`);
  }

  if (params.subject) {
    queryParams.push(`subject=${encodeURIComponent(params.subject)}`);
  }

  if (params.htmlBody) {
    const plainText = params.htmlBody.replace(/<[^>]*>/g, "");
    queryParams.push(`body=${encodeURIComponent(plainText)}`);
  }

  if (queryParams.length) {
    mailto += `?${queryParams.join("&")}`;
  }

  return mailto;
}

// ========================================
// Sauvegarde en base de données
// ========================================
function prepareDataForDatabase(to, cc, bcc, subject, message, valeurs, criteres) {
  const nominateurNom = document.getElementById("nominateur-nom")?.innerText || "";
  const nominateurEmail = document.getElementById("nominateur-email")?.innerText || "";
  // const dateEnvoi = new Date().toISOString();

  return {
    nominateur_nom: nominateurNom,
    nominateur_email: nominateurEmail,
    // date_envoi: dateEnvoi,
    to: to,
    cc: cc,
    bcc: bcc,
    subject: subject,
    message_body: message,
    valeurs: valeurs,
    criteres: criteres,
  };
}

async function saveToDatabase(data) {
  try {
    const response = await fetch(CONFIG.API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: JSON.stringify(data),
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.message || `Erreur API : ${response.status}`);
    }

    const result = await response.json();
    console.log("Données sauvegardées dans la BD:", result);
    return result;
  } catch (error) {
    console.error("Erreur lors de la sauvegarde en BD:", error);
    // Note : On ne bloque pas l'envoi si la BDD échoue
    throw new Error(`Sauvegarde en base de données échouée : ${error.message}`);
  }
}

// ========================================
// UI - Messages et états
// ========================================
function showError(message) {
  const errorEl = document.getElementById("error");
  if (errorEl) {
    errorEl.innerHTML = message;
    errorEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

function showSuccess(message) {
  const successEl = document.getElementById("success");
  if (successEl) {
    successEl.innerText = message;
    successEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

function clearMessages() {
  const errorEl = document.getElementById("error");
  const successEl = document.getElementById("success");

  if (errorEl) errorEl.innerHTML = "";
  if (successEl) successEl.innerText = "";
}

function setLoading(isLoading) {
  const btn = document.getElementById("sendBtn");

  if (!btn) return;

  if (isLoading) {
    btn.disabled = true;
    btn.classList.add("loading");
    btn.innerHTML = '<span class="ms-Button-label">Envoi en cours...</span>';
  } else {
    btn.disabled = false;
    btn.classList.remove("loading");
    btn.innerHTML = '<span class="ms-Button-label">Envoyer la reconnaissance</span>';
  }
}

function resetForm() {
  // Réinitialiser tous les champs sauf le nominateur
  document.getElementById("to").value = "";
  document.getElementById("cc").value = "";
  document.getElementById("bcc").value = "";
  document.getElementById("subject").value = "";
  document.getElementById("message").value = "";

  // Décocher toutes les checkboxes
  document.querySelectorAll(".valeur, .critere").forEach((cb) => {
    cb.checked = false;
  });

  clearMessages();
}

// ========================================
// Utilitaires de debug (optionnel)
// ========================================
function showDebug(message) {
  const debugEl = document.getElementById("debugOutput");
  if (debugEl) {
    const timestamp = new Date().toLocaleTimeString();
    debugEl.textContent += `[${timestamp}] ${message}\n`;
  }
}

// // ========================================
// // Configuration Microsoft Graph
// // ========================================
// const GRAPH_CONFIG = {
//   clientId: "1eafde30-d8fe-4eac-9a7b-8c0026f960fc",
//   tenantId: "901cb4ca-b862-4029-9306-e5cd0f6d9f86",

//   // Permissions nécessaires
//   scopes: ["User.Read", "Mail.Send"],

//   // Endpoints
//   graphEndpoint: "https://graph.microsoft.com/v1.0/me/sendMail"
// };

// // Variables globales pour l'authentification
// let msalInstance = null;
// let currentAccount = null;

// // ========================================
// // Initialisation MSAL
// // ========================================
// function initializeMSAL() {
//   const msalConfig = {
//     auth: {
//       clientId: GRAPH_CONFIG.clientId,
//       authority: `https://login.microsoftonline.com/${GRAPH_CONFIG.tenantId}`,
//       redirectUri: window.location.origin
//     },
//     cache: {
//       cacheLocation: "localStorage",
//       storeAuthStateInCookie: false
//     }
//   };

//   msalInstance = new msal.PublicClientApplication(msalConfig);

//   // Gérer les redirections après authentification
//   msalInstance.handleRedirectPromise()
//     .then(response => {
//       if (response) {
//         currentAccount = response.account;
//         console.log("Authentification réussie:", currentAccount.username);
//       } else {
//         const accounts = msalInstance.getAllAccounts();
//         if (accounts.length > 0) {
//           currentAccount = accounts[0];
//           console.log("Compte existant trouvé:", currentAccount.username);
//         }
//       }
//     })
//     .catch(error => {
//       console.error("Erreur lors de la gestion de la redirection:", error);
//     });
// }

// // ========================================
// // Authentification Microsoft
// // ========================================
// async function getAccessToken() {
//   if (!msalInstance) {
//     throw new Error("MSAL n'est pas initialisé");
//   }

//   const accounts = msalInstance.getAllAccounts();

//   if (accounts.length === 0) {
//     // Aucun compte connecté, demander connexion
//     return await loginWithPopup();
//   }

//   // Tentative d'acquisition silencieuse du token
//   const account = accounts[0];
//   const silentRequest = {
//     scopes: GRAPH_CONFIG.scopes,
//     account: account
//   };

//   try {
//     const response = await msalInstance.acquireTokenSilent(silentRequest);
//     return response.accessToken;
//   } catch (error) {
//     console.log("Token silencieux échoué, popup requise:", error);
//     // Si le token silencieux échoue, demander connexion popup
//     return await loginWithPopup();
//   }
// }

// async function loginWithPopup() {
//   const loginRequest = {
//     scopes: GRAPH_CONFIG.scopes,
//     prompt: "select_account"
//   };

//   try {
//     const response = await msalInstance.loginPopup(loginRequest);
//     currentAccount = response.account;
//     return response.accessToken;
//   } catch (error) {
//     console.error("Erreur lors de la connexion popup:", error);
//     throw new Error("Impossible de se connecter à Microsoft. Veuillez réessayer.");
//   }
// }

// // ========================================
// // Envoi d'email via Microsoft Graph
// // ========================================
// async function sendEmailViaGraph(emailParams) {
//   try {
//     // Obtenir un token d'accès valide
//     const accessToken = await getAccessToken();

//     // Construire le payload au format Microsoft Graph
//     const emailPayload = {
//       message: {
//         subject: emailParams.subject,
//         body: {
//           contentType: "HTML",
//           content: emailParams.htmlBody
//         },
//         toRecipients: emailParams.toRecipients.map(r => ({
//           emailAddress: {
//             address: r.emailAddress,
//             name: r.displayName || r.emailAddress
//           }
//         })),
//         ccRecipients: (emailParams.ccRecipients || []).map(r => ({
//           emailAddress: {
//             address: r.emailAddress,
//             name: r.displayName || r.emailAddress
//           }
//         })),
//         bccRecipients: (emailParams.bccRecipients || []).map(r => ({
//           emailAddress: {
//             address: r.emailAddress,
//             name: r.displayName || r.emailAddress
//           }
//         }))
//       },
//       saveToSentItems: true
//     };

//     // Appel à l'API Microsoft Graph
//     const response = await fetch(GRAPH_CONFIG.graphEndpoint, {
//       method: "POST",
//       headers: {
//         "Authorization": `Bearer ${accessToken}`,
//         "Content-Type": "application/json"
//       },
//       body: JSON.stringify(emailPayload)
//     });

//     if (!response.ok) {
//       const errorData = await response.json().catch(() => ({}));
//       const errorMessage = errorData.error?.message || `Erreur HTTP ${response.status}`;
//       throw new Error(errorMessage);
//     }

//     // L'API Graph retourne 202 Accepted sans body
//     return {
//       success: true,
//       message: "Email envoyé avec succès"
//     };

//   } catch (error) {
//     console.error("Erreur lors de l'envoi via Microsoft Graph:", error);
//     throw error;
//   }
// }

// // ========================================
// // Initialisation Office
// // ========================================
// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook) {
//     // Initialiser MSAL en premier
//     initializeMSAL();

//     // Initialiser le reste de l'application
//     init();
//   }
// });

// function init() {
//   prefillNominateur();
//   attachEventHandlers();
// }

// function attachEventHandlers() {
//   document.getElementById("sendBtn").addEventListener("click", handleSendMessage);

//   // Effacer les erreurs lors de la saisie
//   const inputs = document.querySelectorAll("input, textarea");
//   inputs.forEach(input => {
//     input.addEventListener("input", clearMessages);
//   });
// }

// // ========================================
// // Pré-remplissage du formulaire
// // ========================================
// function prefillNominateur() {
//   try {
//     const user = Office.context.mailbox.userProfile;
//     const displayName = user?.displayName || "Nom inconnu";
//     const emailAddress = user?.emailAddress || "email@inconnu.com";

//     fillNominateurFields(displayName, emailAddress);
//   } catch (error) {
//     console.error("Erreur lors du pré-remplissage:", error);
//     showError("Impossible de récupérer les informations utilisateur.");
//   }
// }

// function fillNominateurFields(name, email) {
//   document.getElementById("nominateur-nom").innerText = name;
//   document.getElementById("nominateur-email").innerText = email;
//   document.getElementById("date-envoi").innerText = new Date().toLocaleDateString("fr-FR");
// }

// // ========================================
// // Validation des données
// // ========================================
// function isValidEmail(email) {
//   return CONFIG.EMAIL_REGEX.test(email);
// }

// function getInvalidEmails(emails) {
//   return emails.filter(email => !isValidEmail(email));
// }

// function parseEmailField(fieldId) {
//   const value = document.getElementById(fieldId)?.value || "";
//   return value
//     .split(/[,;]+/)
//     .map(e => e.trim())
//     .filter(e => e);
// }

// function validateForm() {
//   const errors = [];

//   // Récupération des champs
//   const to = parseEmailField("to");
//   const cc = parseEmailField("cc");
//   const bcc = parseEmailField("bcc");
//   const subject = document.getElementById("subject")?.value.trim() || "";
//   const message = document.getElementById("message")?.value.trim() || "";
//   const valeurs = getCheckedValues(".valeur");
//   const criteres = getCheckedValues(".critere");

//   // Validation des champs obligatoires
//   if (!to.length) {
//     errors.push("Le champ 'Destinataire(s)' est obligatoire.");
//   }

//   if (!subject) {
//     errors.push("Le champ 'Objet' est obligatoire.");
//   }

//   if (!valeurs.length) {
//     errors.push("Veuillez sélectionner au moins une valeur reconnue.");
//   }

//   if (!criteres.length) {
//     errors.push("Veuillez sélectionner au moins un critère de reconnaissance.");
//   }

//   if (!message) {
//     errors.push("Le champ 'Message' est obligatoire.");
//   }

//   // Validation des emails
//   const allEmails = [...to, ...cc, ...bcc];
//   const invalidEmails = getInvalidEmails(allEmails);

//   if (invalidEmails.length) {
//     errors.push(`Adresse(s) email invalide(s) : ${invalidEmails.join(", ")}`);
//   }

//   return {
//     isValid: errors.length === 0,
//     errors,
//     data: { to, cc, bcc, subject, message, valeurs, criteres }
//   };
// }

// function getCheckedValues(selector) {
//   return Array.from(document.querySelectorAll(`${selector}:checked`))
//     .map(cb => cb.value);
// }

// // ========================================
// // Construction du message
// // ========================================
// function buildMessageBody(message, valeurs, criteres) {
//   let body = "<div style='font-family: Segoe UI, sans-serif;'>";

//   if (valeurs.length) {
//     body += `<p><strong>Valeurs reconnues :</strong><br>${valeurs.join(", ")}</p>`;
//   }

//   if (criteres.length) {
//     body += `<p><strong>Critères :</strong><br>${criteres.join(", ")}</p>`;
//   }

//   body += `<hr style='border: none; border-top: 1px solid #edebe9; margin: 16px 0;'>`;
//   body += `<p>${message.replace(/\n/g, "<br>")}</p>`;
//   body += "</div>";

//   return body;
// }

// // ========================================
// // Gestion de l'envoi (NOUVELLE LOGIQUE)
// // ========================================
// async function handleSendMessage(event) {
//   event.preventDefault();

//   // Nettoyer les messages précédents
//   clearMessages();

//   // Validation
//   const validation = validateForm();

//   if (!validation.isValid) {
//     showError(validation.errors.join("<br>"));
//     return;
//   }

//   // Afficher le loader
//   setLoading(true);

//   try {
//     const { to, cc, bcc, subject, message, valeurs, criteres } = validation.data;

//     // Ajouter l'email RH en CCI automatiquement
//     const bccWithRH = [...bcc];
//     if (CONFIG.RH_EMAIL && !bccWithRH.includes(CONFIG.RH_EMAIL)) {
//       bccWithRH.push(CONFIG.RH_EMAIL);
//     }

//     // Construire le corps du message HTML
//     const htmlBody = buildMessageBody(message, valeurs, criteres);

//     // Préparer les paramètres pour Microsoft Graph
//     const emailParams = {
//       toRecipients: toEmailAddressDetails(to),
//       ccRecipients: cc.length ? toEmailAddressDetails(cc) : [],
//       bccRecipients: bccWithRH.length ? toEmailAddressDetails(bccWithRH) : [],
//       subject: subject,
//       htmlBody: htmlBody
//     };

//     // ENVOI AUTOMATIQUE via Microsoft Graph API
//     await sendEmailViaGraph(emailParams);

//     // Préparer les données pour la BDD
//     const dbData = prepareDataForDatabase(to, cc, bccWithRH, subject, message, valeurs, criteres);

//     // Sauvegarder en BDD après envoi réussi
//     await saveToDatabase(dbData);

//     // Succès
//     showSuccess("Email envoyé automatiquement et sauvegardé avec succès !");

//     // Réinitialiser le formulaire après 2 secondes
//     setTimeout(resetForm, 2000);

//   } catch (error) {
//     console.error("Erreur lors de l'envoi:", error);

//     // Messages d'erreur plus détaillés
//     let errorMessage = "Erreur lors de l'envoi : ";

//     if (error.message.includes("MSAL")) {
//       errorMessage += "Problème d'authentification Microsoft. Veuillez vous reconnecter.";
//     } else if (error.message.includes("MailboxNotEnabledForRESTAPI")) {
//       errorMessage += "Votre boîte mail n'est pas configurée pour l'API. Contactez votre administrateur.";
//     } else if (error.message.includes("InvalidAuthenticationToken")) {
//       errorMessage += "Session expirée. Veuillez rafraîchir la page.";
//     } else {
//       errorMessage += error.message;
//     }

//     showError(errorMessage);
//   } finally {
//     setLoading(false);
//   }
// }

// // ========================================
// // Utilitaire de conversion d'emails
// // ========================================
// function toEmailAddressDetails(emails) {
//   return emails.map(email => ({
//     emailAddress: email,
//     displayName: email
//   }));
// }

// // ========================================
// // Sauvegarde en base de données
// // ========================================
// function prepareDataForDatabase(to, cc, bcc, subject, message, valeurs, criteres) {
//   const nominateurNom = document.getElementById("nominateur-nom")?.innerText || "";
//   const nominateurEmail = document.getElementById("nominateur-email")?.innerText || "";

//   return {
//     nominateur_nom: nominateurNom,
//     nominateur_email: nominateurEmail,
//     to: to,
//     cc: cc,
//     bcc: bcc,
//     subject: subject,
//     message_body: message,
//     valeurs: valeurs,
//     criteres: criteres
//   };
// }

// async function saveToDatabase(data) {
//   try {
//     const response = await fetch(CONFIG.API_URL, {
//       method: "POST",
//       headers: {
//         "Content-Type": "application/json",
//         "Accept": "application/json"
//       },
//       body: JSON.stringify(data)
//     });

//     if (!response.ok) {
//       const errorData = await response.json().catch(() => ({}));
//       throw new Error(errorData.message || `Erreur API : ${response.status}`);
//     }

//     const result = await response.json();
//     console.log("Données sauvegardées dans la BD:", result);
//     return result;

//   } catch (error) {
//     console.error("Erreur lors de la sauvegarde en BD:", error);
//     // Note : On affiche un warning mais on ne bloque pas l'envoi
//     console.warn("L'email a été envoyé mais n'a pas pu être sauvegardé en base de données");
//     throw error;
//   }
// }

// // ========================================
// // UI - Messages et états
// // ========================================
// function showError(message) {
//   const errorEl = document.getElementById("error");
//   if (errorEl) {
//     errorEl.innerHTML = message;
//     errorEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
//   }
// }

// function showSuccess(message) {
//   const successEl = document.getElementById("success");
//   if (successEl) {
//     successEl.innerText = message;
//     successEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
//   }
// }

// function clearMessages() {
//   const errorEl = document.getElementById("error");
//   const successEl = document.getElementById("success");

//   if (errorEl) errorEl.innerHTML = "";
//   if (successEl) successEl.innerText = "";
// }

// function setLoading(isLoading) {
//   const btn = document.getElementById("sendBtn");

//   if (!btn) return;

//   if (isLoading) {
//     btn.disabled = true;
//     btn.classList.add("loading");
//     btn.innerHTML = '<span class="ms-Button-label">Envoi en cours...</span>';
//   } else {
//     btn.disabled = false;
//     btn.classList.remove("loading");
//     btn.innerHTML = '<span class="ms-Button-label">Envoyer la reconnaissance</span>';
//   }
// }

// function resetForm() {
//   // Réinitialiser tous les champs sauf le nominateur
//   document.getElementById("to").value = "";
//   document.getElementById("cc").value = "";
//   document.getElementById("bcc").value = "";
//   document.getElementById("subject").value = "";
//   document.getElementById("message").value = "";

//   // Décocher toutes les checkboxes
//   document.querySelectorAll(".valeur, .critere").forEach(cb => {
//     cb.checked = false;
//   });

//   clearMessages();
// }

// // ========================================
// // Utilitaires de debug (optionnel)
// // ========================================
// function showDebug(message) {
//   const debugEl = document.getElementById("debugOutput");
//   if (debugEl) {
//     const timestamp = new Date().toLocaleTimeString();
//     debugEl.textContent += `[${timestamp}] ${message}\n`;
//   }
// }
