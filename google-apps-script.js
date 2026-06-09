const SHEET_NAME_BOOKINGS = 'Réservations';
const SHEET_NAME_SUBS     = 'Abonnements';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'booking') {
      let sheet = ss.getSheetByName(SHEET_NAME_BOOKINGS);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_BOOKINGS);
        sheet.appendRow(['Référence','Prénom','Nom','Téléphone','Adresse','Date','Type','Montant (F CFA)','Statut','Enregistré le']);
        sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#0D2B4E').setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([
        data.reference, data.prenom, data.nom, data.tel, data.addr,
        data.date, data.type, data.montant, data.statut,
        new Date().toLocaleString('fr-FR')
      ]);
      // Couleur selon statut
      colorLastRow(sheet, data.statut);
    }

    else if (data.action === 'subscription') {
      let sheet = ss.getSheetByName(SHEET_NAME_SUBS);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_SUBS);
        sheet.appendRow(['Référence','Prénom','Nom','Téléphone','Adresse','Type','Validité','Montant (F CFA)','Statut','Enregistré le']);
        sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#0D2B4E').setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([
        data.reference, data.prenom, data.nom, data.tel, data.addr,
        data.type, data.validite, data.montant, data.statut,
        new Date().toLocaleString('fr-FR')
      ]);
      colorLastRow(sheet, data.statut);
    }

    else if (data.action === 'update_status') {
      // Met à jour le statut quand le passager clique "J'ai payé"
      updateStatus(ss, data.reference, data.statut);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'FOCUS Script actif ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function colorLastRow(sheet, statut) {
  const last = sheet.getLastRow();
  const color = statut === 'Payé Wave' ? '#E1F5EE'
              : statut.includes('attente') ? '#FAEEDA'
              : '#F9F5EE';
  sheet.getRange(last, 1, 1, sheet.getLastColumn()).setBackground(color);
}

function updateStatus(ss, reference, newStatut) {
  [SHEET_NAME_BOOKINGS, SHEET_NAME_SUBS].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reference) {
        const statutCol = name === SHEET_NAME_BOOKINGS ? 9 : 9; // colonne Statut
        sheet.getRange(i + 1, statutCol).setValue(newStatut);
        colorLastRow(sheet, newStatut);
        break;
      }
    }
  });
}
