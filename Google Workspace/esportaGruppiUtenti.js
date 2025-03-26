/*

Zucchetti SpA - Estrazione utenti e gruppi da Google Workspace
Versione 1.0 by Fabrizio Alberti - 26.03.2025

ISTRUZIONI

Entrare come Workspace Global Admin

Vai su https://script.new

Incolla lo script

Aggiungi il servizio: Admin Directory API

Esegui lo script confermando le autorizzazioni richieste.

Ad estrazione terminata, verrà presentato un link di un foglio di workspace da salvare in formato excel

*/

function esportaGruppiUtenti() {

  // === Recupera dominio principale
  let dominioPrincipale = "dominio_sconosciuto";
  try {
    const domains = AdminDirectory.Domains.list("my_customer").domains || [];
    dominioPrincipale = domains.find(d => d.isPrimary)?.domainName || dominioPrincipale;
    Logger.log(`Inizio Estrazione utenti e gruppi da Google Workspace per il dominio (principale) ${dominioPrincipale}`);
  } catch (e) {
    Logger.log("KO!!! Impossibile recuperare il dominio principale: " + e.message);
  }

// === Data corrente nel formato GG/MM/AAAA
  const oggi = new Date();
  const giorno = String(oggi.getDate()).padStart(2, '0');
  const mese = String(oggi.getMonth() + 1).padStart(2, '0'); // I mesi partono da 0
  const anno = oggi.getFullYear();
  const dataFormattata = `${giorno}/${mese}/${anno}`;

  const nomeFile = `Esportazione Directory Workspace – ${dominioPrincipale} – ${dataFormattata}`;
  const ss = SpreadsheetApp.create(nomeFile);

  // === FOGLIO UTENTI ===
  const utentiSheet = ss.insertSheet("Utenti");
  utentiSheet.appendRow(["Nome", "Cognome", "Email"]);

  let userToken;
  let totaleUtenti = 0;
  let utentiElenco = [];

  do {
    const response = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      pageToken: userToken,
      orderBy: 'email'
    });
    const users = response.users || [];
    utentiElenco = utentiElenco.concat(users);
    userToken = response.nextPageToken;
  } while (userToken);

  totaleUtenti = utentiElenco.length;
  Logger.log("Totale utenti trovati: " + totaleUtenti);

  utentiElenco.forEach((u, index) => {
    try {
      const nome = u.name?.givenName || "";
      const cognome = u.name?.familyName || "";
      const email = u.primaryEmail || "";

      utentiSheet.appendRow([nome, cognome, email]);

      const percent = Math.round(((index + 1) / totaleUtenti) * 100);
      Logger.log(`[${index + 1}/${totaleUtenti}] ${email} – ${percent}% completato`);
    } catch (e) {
      const email = u.primaryEmail || "???";
      Logger.log(`KO!!! Errore con l'utente ${email}: ${e.message}`);
    }
  });

  // === FOGLIO GRUPPI E MEMBRI ===
  const datiSheet = ss.insertSheet("Gruppi e Membri");
  datiSheet.appendRow(["Gruppo", "Membri"]);

  const gruppi = [];
  let pageToken;
  do {
    const response = AdminDirectory.Groups.list({
      customer: 'my_customer',
      maxResults: 200,
      pageToken: pageToken
    });
    const groups = response.groups || [];
    groups.forEach(g => gruppi.push(g));
    pageToken = response.nextPageToken;
  } while (pageToken);

  const totaleGruppi = gruppi.length;
  Logger.log("Totale gruppi trovati: " + totaleGruppi);

  gruppi.forEach((g, index) => {
    const gruppoEmail = g.email;
    let membriEmails = [];

    try {
      const members = AdminDirectory.Members.list(gruppoEmail).members || [];
      membriEmails = members.map(m => m.email);
    } catch (e) {
      Logger.log(`KO!! Errore con il gruppo "${gruppoEmail}" – errore: ${e.message}`);
    }

    datiSheet.appendRow([gruppoEmail, membriEmails.join("|")]);

    const percent = Math.round(((index + 1) / totaleGruppi) * 100);
    Logger.log(`[${index + 1}/${totaleGruppi}] ${gruppoEmail} – ${percent}% completato`);
  });

  Logger.log(`Esportazione completata di ${totaleGruppi} gruppi e ${totaleUtenti} utenti.`);
  Logger.log("File generato: " + ss.getUrl());
  
  // === Rimuove "Foglio1" se presente ===
  const defaultSheet = ss.getSheets()[0];
  if (defaultSheet.getName() === "Foglio1") {
    ss.deleteSheet(defaultSheet);
  }  
}
