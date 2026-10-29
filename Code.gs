/**
 * RENDEVO Physiotherapy — complete Apps Script backend
 * Time zone: Africa/Algiers
 *
 * Setup:
 * 1) Paste this entire file into Code.gs and save.
 * 2) Run setupPhysiotherapySystem() once and approve permissions.
 * 3) Deploy > Manage deployments > Edit > New version > Deploy.
 *
 * Notifications use the executing Google account. Set OWNER_EMAIL below if
 * you want clinic notifications sent to a different address.
 */

const BOOKING_SETTINGS = Object.freeze({
  TIME_ZONE: "Africa/Algiers",
  SPREADSHEET_PROPERTY: "PHYSIOTHERAPY_SPREADSHEET_ID",
  SHEET_NAME: "Demandes kinésithérapie",
  OWNER_EMAIL: "", // Optional: clinic admin notification email.
  CALENDAR_ID: "primary",
  WEBSITE_URL: "https://wadietaghzouti20-dot.github.io/physiotherapy-/",
  CANCEL_PAGE_URL: "https://wadietaghzouti20-dot.github.io/physiotherapy-/cancel.html",
  ALLOWED_SERVICES: [
    "Bilan kinésithérapique",
    "Rééducation fonctionnelle",
    "Rééducation post-opératoire",
    "Douleurs du dos et du cou",
    "Rééducation sportive",
    "Séance de suivi",
    "Rééducation respiratoire",
    "Autre motif"
  ],
  SERVICE_DURATIONS_MINUTES: {
    "Bilan kinésithérapique": 60,
    "Rééducation fonctionnelle": 45,
    "Rééducation post-opératoire": 45,
    "Douleurs du dos et du cou": 45,
    "Rééducation sportive": 45,
    "Séance de suivi": 30,
    "Rééducation respiratoire": 45,
    "Autre motif": 30
  },
  MAX_NAME_LENGTH: 100,
  MAX_PHONE_LENGTH: 25,
  MAX_EMAIL_LENGTH: 150,
  MAX_NOTES_LENGTH: 1000
});

const BOOKING_HEADERS = [
  "Référence", "Créée le", "Statut", "Motif", "Date souhaitée",
  "Moment préféré", "Date confirmée", "Heure confirmée", "Nom et prénom",
  "Téléphone", "E-mail", "Message", "Notes internes", "Jeton annulation",
  "Rappel envoyé le", "ID déclencheur rappel", "ID événement calendrier"
];

function setupPhysiotherapySystem() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(BOOKING_SETTINGS.SPREADSHEET_PROPERTY);
  let spreadsheet = null;

  if (spreadsheetId) {
    try { spreadsheet = SpreadsheetApp.openById(spreadsheetId); } catch (err) {}
  }
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create("RENDEVO Physiotherapy — Demandes de rendez-vous");
    props.setProperty(BOOKING_SETTINGS.SPREADSHEET_PROPERTY, spreadsheet.getId());
  }

  spreadsheet.setSpreadsheetTimeZone(BOOKING_SETTINGS.TIME_ZONE);
  let sheet = spreadsheet.getSheetByName(BOOKING_SETTINGS.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(BOOKING_SETTINGS.SHEET_NAME);
  ensureBookingColumns_(sheet);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  const headers = getHeaderMap_(sheet);
  if (headers["Créée le"]) sheet.getRange(2, headers["Créée le"], Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  if (headers["Rappel envoyé le"]) sheet.getRange(2, headers["Rappel envoyé le"], Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  Logger.log("Spreadsheet ready: " + spreadsheet.getUrl());
  Logger.log("Admin notification email: " + getOwnerEmail_());
  Logger.log("Calendar: " + BOOKING_SETTINGS.CALENDAR_ID);
  return spreadsheet.getUrl();
}

function doGet() {
  return jsonResponse_({
    success: true,
    service: "RENDEVO Physiotherapy",
    message: "Le service de réservation est disponible."
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || typeof e.postData.contents !== "string") {
      return jsonResponse_({ success: false, message: "Requête invalide." });
    }
    const raw = JSON.parse(e.postData.contents || "{}");

    if (raw.action === "cancel") {
      return jsonResponse_({ success: true, message: cancelAppointment_(raw.reference, raw.token) });
    }
    if (raw.action === "availability") {
      const slots = getAvailableSlots_(raw.preferredDate, raw.service);
      return jsonResponse_({
        success: true,
        availableSlots: slots,
        message: slots.length
          ? "Choisissez un créneau disponible."
          : "Aucun créneau disponible à cette date. Choisissez une autre date."
      });
    }

    const data = validateRequest_(raw);
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    let reference, token;
    try {
      const sheet = getBookingSheet_();
      assertSlotCanBeBooked_(sheet, data.preferredDate, data.preferredTime, data.service, "");
      reference = "RDV-" + Utilities.getUuid().replace(/-/g, "").slice(0, 8).toUpperCase();
      token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "").toLowerCase();
      const row = {};
      row["Référence"] = reference;
      row["Créée le"] = new Date();
      row["Statut"] = "Confirmé";
      row["Motif"] = data.service;
      row["Date souhaitée"] = data.preferredDate;
      row["Moment préféré"] = data.preferredTime;
      row["Date confirmée"] = data.preferredDate;
      row["Heure confirmée"] = data.preferredTime;
      row["Nom et prénom"] = safeCell_(data.fullName);
      row["Téléphone"] = safeCell_(data.phone);
      row["E-mail"] = data.email;
      row["Message"] = safeCell_(data.notes);
      row["Notes internes"] = "";
      row["Jeton annulation"] = token;
      row["Rappel envoyé le"] = "";
      row["ID déclencheur rappel"] = "";
      row["ID événement calendrier"] = "";
      appendBookingRow_(sheet, row);
    } finally {
      lock.releaseLock();
    }

    // These integrations must not undo a successfully saved booking if a
    // Google service temporarily fails. Each failure is written to execution logs.
    const calendarId = createCalendarEvent_(reference, token, data);
    if (calendarId) saveBookingField_(reference, "ID événement calendrier", calendarId);

    try {
      scheduleAppointmentReminder_(reference, token, data);
    } catch (err) {
      console.error("Reminder scheduling failed for " + reference + ": " + err);
    }

    sendClientConfirmation_(reference, token, data);
    sendClinicNotification_(reference, data);

    return jsonResponse_({
      success: true,
      reference: reference,
      status: "Confirmé",
      confirmedDate: data.preferredDate,
      confirmedTime: data.preferredTime,
      message: "Votre rendez-vous est confirmé le " + data.preferredDate + " à " + data.preferredTime + "."
    });
  } catch (error) {
    console.error("Booking request failed: " + (error && error.stack ? error.stack : error));
    return jsonResponse_({
      success: false,
      message: error && error.message
        ? error.message
        : "Une erreur est survenue. Veuillez vérifier les informations et réessayer."
    });
  }
}

function validateRequest_(input) {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    throw new Error("Les informations envoyées sont invalides.");
  }
  const service = cleanText_(input.service, 80);
  const preferredDate = cleanText_(input.preferredDate, 10);
  const preferredTime = cleanText_(input.preferredTime, 5);
  const fullName = cleanText_(input.fullName, BOOKING_SETTINGS.MAX_NAME_LENGTH);
  const phone = cleanText_(input.phone, BOOKING_SETTINGS.MAX_PHONE_LENGTH);
  const email = cleanText_(input.email, BOOKING_SETTINGS.MAX_EMAIL_LENGTH);
  const notes = cleanText_(input.notes, BOOKING_SETTINGS.MAX_NOTES_LENGTH);

  if (BOOKING_SETTINGS.ALLOWED_SERVICES.indexOf(service) === -1) throw new Error("Veuillez choisir un motif de rendez-vous valide.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(preferredDate) || !isRealDate_(preferredDate)) throw new Error("Veuillez choisir une date valide.");
  const today = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  if (preferredDate < today) throw new Error("La date choisie est déjà passée. Veuillez en sélectionner une autre.");
  if (!/^([01]\d|2[0-3]):([0-5]\d)$/.test(preferredTime)) throw new Error("Veuillez choisir une heure valide parmi les créneaux proposés.");
  if (fullName.length < 2) throw new Error("Veuillez saisir votre nom et prénom.");
  const digits = phone.replace(/\D/g, "");
  if (digits.length < 8 || digits.length > 15) throw new Error("Veuillez saisir un numéro de téléphone valide.");
  if (!email) throw new Error("Une adresse e-mail est nécessaire pour envoyer la confirmation et le rappel.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Veuillez vérifier votre adresse e-mail.");
  if (notes.length > BOOKING_SETTINGS.MAX_NOTES_LENGTH) throw new Error("Le message ne doit pas dépasser 1 000 caractères.");
  if (input.privacyConsent !== true) throw new Error("Veuillez accepter l’utilisation de vos informations pour traiter la demande.");

  return { service: service, preferredDate: preferredDate, preferredTime: preferredTime,
    fullName: fullName, phone: phone, email: email, notes: notes };
}

function getAvailableSlots_(dateValue, serviceValue) {
  const date = cleanText_(dateValue, 10);
  const service = cleanText_(serviceValue, 80);
  if (BOOKING_SETTINGS.ALLOWED_SERVICES.indexOf(service) === -1) throw new Error("Veuillez choisir un motif de rendez-vous valide.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !isRealDate_(date)) throw new Error("Veuillez choisir une date valide.");
  const today = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  if (date < today) throw new Error("La date choisie est déjà passée. Veuillez en sélectionner une autre.");

  const schedule = getScheduleForDate_(date);
  if (!schedule.open) return [];
  const duration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[service] || 30;
  let firstStart = schedule.openingMinutes;
  if (date === today) {
    const nowParts = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "HH:mm").split(":").map(Number);
    firstStart = Math.max(firstStart, Math.ceil((nowParts[0] * 60 + nowParts[1] + 15) / 15) * 15);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const values = getBookingSheet_().getDataRange().getValues();
    const slots = [];
    for (let start = firstStart; start + duration <= schedule.closingMinutes; start += 15) {
      if (!hasBookingConflict_(values, date, start, duration, "")) slots.push(minutesToTime_(start));
    }
    return slots;
  } finally {
    lock.releaseLock();
  }
}

function getScheduleForDate_(date) {
  const day = new Date(date + "T12:00:00Z");
  const weekday = Number(Utilities.formatDate(day, BOOKING_SETTINGS.TIME_ZONE, "u"));
  if (weekday === 5) return { open: false, openingMinutes: 0, closingMinutes: 0 };
  if (weekday === 6) return { open: true, openingMinutes: 600, closingMinutes: 960 };
  return { open: true, openingMinutes: 480, closingMinutes: 1020 };
}

function assertSlotCanBeBooked_(sheet, date, time, service, ignoredReference) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !isRealDate_(date)) throw new Error("Date de rendez-vous invalide.");
  const today = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  if (date < today) throw new Error("La date choisie est déjà passée.");
  if (!/^([01]\d|2[0-3]):([0-5]\d)$/.test(time)) throw new Error("Veuillez choisir une heure de rendez-vous valide.");
  const parts = time.split(":").map(Number);
  const start = parts[0] * 60 + parts[1];
  if (parts[1] % 15 !== 0) throw new Error("Choisissez un début de rendez-vous sur un quart d'heure.");
  if (date === today) {
    const now = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "HH:mm").split(":").map(Number);
    if (start < Math.ceil((now[0] * 60 + now[1] + 15) / 15) * 15) throw new Error("Ce créneau n’est plus disponible. Choisissez une heure future.");
  }
  const schedule = getScheduleForDate_(date);
  if (!schedule.open) throw new Error("Le cabinet est fermé le vendredi.");
  const duration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[service] || 30;
  if (start < schedule.openingMinutes || start + duration > schedule.closingMinutes) {
    throw new Error("Ce créneau ne respecte pas les horaires du cabinet (" + minutesToTime_(schedule.openingMinutes) + "–" + minutesToTime_(schedule.closingMinutes) + ").");
  }
  if (hasBookingConflict_(sheet.getDataRange().getValues(), date, start, duration, ignoredReference)) {
    throw new Error("Ce créneau vient d'être réservé. Veuillez choisir un autre horaire.");
  }
}

function hasBookingConflict_(values, date, startMinutes, duration, ignoredReference) {
  if (!values || values.length < 2) return false;
  const col = {};
  values[0].forEach(function (name, index) { col[name] = index; });
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[col["Référence"]] || "") === ignoredReference) continue;
    if (String(row[col["Statut"]] || "") !== "Confirmé") continue;
    if (normalizeDate_(row[col["Date confirmée"]]) !== date) continue;
    const match = normalizeTime_(row[col["Heure confirmée"]]).match(/^(\d{2}):(\d{2})$/);
    if (!match) continue;
    const existingStart = Number(match[1]) * 60 + Number(match[2]);
    const existingService = String(row[col["Motif"]] || "");
    const existingDuration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[existingService] || 30;
    if (startMinutes < existingStart + existingDuration && existingStart < startMinutes + duration) return true;
  }
  return false;
}

function confirmAppointment(reference, confirmedDate, confirmedTime) {
  reference = cleanText_(reference, 30);
  confirmedDate = cleanText_(confirmedDate, 10);
  confirmedTime = cleanText_(confirmedTime, 5);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getBookingSheet_();
    const rowNumber = findBookingRow_(sheet, reference);
    if (rowNumber < 2) throw new Error("Référence introuvable.");
    const headers = getHeaderMap_(sheet);
    const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const service = String(row[headers["Motif"] - 1] || "");
    assertSlotCanBeBooked_(sheet, confirmedDate, confirmedTime, service, reference);
    sheet.getRange(rowNumber, headers["Date confirmée"]).setNumberFormat("@").setValue(confirmedDate);
    sheet.getRange(rowNumber, headers["Heure confirmée"]).setNumberFormat("@").setValue(confirmedTime);
    sheet.getRange(rowNumber, headers["Statut"]).setValue("Confirmé");
    return "Rendez-vous confirmé : " + reference + " le " + confirmedDate + " à " + confirmedTime + ".";
  } finally {
    lock.releaseLock();
  }
}

function createCalendarEvent_(reference, token, data) {
  try {
    const calendar = CalendarApp.getCalendarById(BOOKING_SETTINGS.CALENDAR_ID);
    if (!calendar) throw new Error("Calendar not accessible: " + BOOKING_SETTINGS.CALENDAR_ID);
    const start = appointmentDateTime_(data.preferredDate, data.preferredTime);
    const duration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[data.service] || 30;
    const end = new Date(start.getTime() + duration * 60000);
    const details = [
      "Statut : Confirmé", "Référence : " + reference, "Motif : " + data.service,
      "Patient : " + data.fullName, "Téléphone : " + data.phone,
      "E-mail : " + (data.email || "Non renseigné"), "Message : " + (data.notes || "Aucun"),
      "Annulation patient : " + getCancellationUrl_(reference, token)
    ].join("\n");
    const options = { description: details, sendInvites: Boolean(data.email) };
    if (data.email) options.guests = data.email;
    const event = calendar.createEvent("Rendez-vous — " + data.fullName + " (" + data.service + ")", start, end, options);
    return event.getId();
  } catch (err) {
    console.error("Calendar event creation failed for " + reference + ": " + err);
    return "";
  }
}

function sendClientConfirmation_(reference, token, data) {
  if (!data.email) {
    console.log("No client email provided; skipped confirmation for " + reference);
    return;
  }
  const cancelUrl = getCancellationUrl_(reference, token);
  const dateLabel = formatDateFr_(data.preferredDate);
  const timeLabel = data.preferredTime;
  const subject = "Rendez-vous confirmé — " + reference;
  const body = [
    "Bonjour " + data.fullName + ",", "",
    "Votre rendez-vous est confirmé.", "",
    "Motif : " + data.service, "Date : " + dateLabel, "Heure : " + timeLabel,
    "Référence : " + reference, "",
    "Pour annuler ce rendez-vous, ouvrez ce lien : " + cancelUrl, "",
    "Un rappel vous sera envoyé environ 24 heures avant le rendez-vous si celui-ci est prévu dans plus de 24 heures.",
    "Vous pouvez également recevoir une invitation Google Agenda si votre adresse e-mail a été renseignée.",
    "", "À bientôt."
  ].join("\n");
  const html = "<div style='font-family:Arial,sans-serif;line-height:1.6;color:#202b35;max-width:600px'>" +
    "<h2 style='color:#176b68'>Votre rendez-vous est confirmé</h2>" +
    "<p>Bonjour " + escapeHtml_(data.fullName) + ",</p><p>Votre rendez-vous a bien été confirmé.</p>" +
    "<p><b>Motif :</b> " + escapeHtml_(data.service) + "<br><b>Date :</b> " + escapeHtml_(dateLabel) +
    "<br><b>Heure :</b> " + escapeHtml_(timeLabel) + "<br><b>Référence :</b> " + escapeHtml_(reference) + "</p>" +
    "<p>Un rappel vous sera envoyé environ 24 heures avant le rendez-vous, si celui-ci est prévu dans plus de 24 heures.</p>" +
    "<p><a href='" + escapeHtml_(cancelUrl) + "' style='display:inline-block;padding:12px 18px;background:#b42318;color:#fff;text-decoration:none;border-radius:8px;font-weight:bold'>Annuler mon rendez-vous</a></p>" +
    "<p>Si le bouton ne fonctionne pas, copiez ce lien :<br><a href='" + escapeHtml_(cancelUrl) + "'>" + escapeHtml_(cancelUrl) + "</a></p>" +
    "<p>À bientôt.</p></div>";
  try {
    MailApp.sendEmail({ to: data.email, subject: subject, body: body, htmlBody: html, name: "Cabinet de kinésithérapie" });
  } catch (err) {
    console.error("Client confirmation email failed for " + reference + ": " + err);
  }
}

function sendClinicNotification_(reference, data) {
  const recipient = getOwnerEmail_();
  if (!recipient) {
    console.error("Admin email skipped: OWNER_EMAIL is blank and Apps Script could not identify the executing account.");
    return;
  }
  const body = [
    "NOUVEAU RENDEZ-VOUS CONFIRMÉ", "",
    "Statut : Confirmé", "Référence : " + reference,
    "Nom et prénom : " + data.fullName, "Téléphone : " + data.phone,
    "E-mail : " + (data.email || "Non renseigné"), "Motif : " + data.service,
    "Date : " + formatDateFr_(data.preferredDate), "Heure : " + data.preferredTime,
    "Message du patient : " + (data.notes || "Aucun"),
    "", "Le rendez-vous a été ajouté au Google Agenda principal si Calendar était accessible."
  ].join("\n");
  const html = "<div style='font-family:Arial,sans-serif;line-height:1.6;color:#202b35'>" +
    "<h2>Nouveau rendez-vous confirmé</h2><p><b>Statut :</b> Confirmé<br><b>Référence :</b> " + escapeHtml_(reference) +
    "<br><b>Patient :</b> " + escapeHtml_(data.fullName) + "<br><b>Téléphone :</b> " + escapeHtml_(data.phone) +
    "<br><b>E-mail :</b> " + escapeHtml_(data.email || "Non renseigné") + "<br><b>Motif :</b> " + escapeHtml_(data.service) +
    "<br><b>Date :</b> " + escapeHtml_(formatDateFr_(data.preferredDate)) + "<br><b>Heure :</b> " + escapeHtml_(data.preferredTime) +
    "</p><h3>Message du patient</h3><p>" + escapeHtml_(data.notes || "Aucun") + "</p></div>";
  try {
    MailApp.sendEmail({ to: recipient, subject: "Nouveau rendez-vous confirmé — " + reference, body: body, htmlBody: html, name: "RENDEVO Physiotherapy" });
  } catch (err) {
    console.error("Admin notification failed for " + reference + ": " + err);
  }
}

function scheduleAppointmentReminder_(reference, token, data) {
  if (!data.email) return;
  const appointmentAt = appointmentDateTime_(data.preferredDate, data.preferredTime);
  const reminderAt = new Date(appointmentAt.getTime() - 24 * 60 * 60 * 1000);
  if (reminderAt.getTime() <= Date.now()) {
    console.log("Skipped 24-hour reminder because booking is less than 24 hours away: " + reference);
    return;
  }
  const trigger = ScriptApp.newTrigger("sendScheduledReminder").timeBased().at(reminderAt).create();
  const sheet = getBookingSheet_();
  const row = findBookingRow_(sheet, reference);
  if (row < 2) {
    ScriptApp.deleteTrigger(trigger);
    return;
  }
  const headers = getHeaderMap_(sheet);
  sheet.getRange(row, headers["ID déclencheur rappel"]).setValue(trigger.getUniqueId());
}

function sendScheduledReminder(e) {
  const triggerUid = e && e.triggerUid ? String(e.triggerUid) : "";
  if (!triggerUid) return;
  const sheet = getBookingSheet_();
  const headers = getHeaderMap_(sheet);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[headers["ID déclencheur rappel"] - 1] || "") !== triggerUid) continue;
    const rowNumber = i + 1;
    const reference = String(row[headers["Référence"] - 1] || "");
    const status = String(row[headers["Statut"] - 1] || "");
    const email = String(row[headers["E-mail"] - 1] || "").trim();
    const token = String(row[headers["Jeton annulation"] - 1] || "");
    const date = normalizeDate_(row[headers["Date confirmée"] - 1]);
    const time = normalizeTime_(row[headers["Heure confirmée"] - 1]);
    const name = String(row[headers["Nom et prénom"] - 1] || "");
    const service = String(row[headers["Motif"] - 1] || "");
    try {
      if (status === "Confirmé" && email && token && date && time && !row[headers["Rappel envoyé le"] - 1]) {
        const cancelUrl = getCancellationUrl_(reference, token);
        const dateLabel = formatDateFr_(date);
        const body = "Bonjour " + name + ",\n\nRappel : votre rendez-vous est prévu dans environ 24 heures.\n\nMotif : " + service +
          "\nDate : " + dateLabel + "\nHeure : " + time + "\nRéférence : " + reference +
          "\n\nPour annuler ce rendez-vous, ouvrez ce lien : " + cancelUrl + "\n\nSi vous souhaitez conserver le rendez-vous, aucune action n'est nécessaire.";
        const html = "<div style='font-family:Arial,sans-serif;line-height:1.6;color:#202b35'><h2>Rappel de rendez-vous</h2>" +
          "<p>Bonjour " + escapeHtml_(name) + ", votre rendez-vous est prévu dans environ 24 heures.</p><p><b>Motif :</b> " +
          escapeHtml_(service) + "<br><b>Date :</b> " + escapeHtml_(dateLabel) + "<br><b>Heure :</b> " + escapeHtml_(time) +
          "<br><b>Référence :</b> " + escapeHtml_(reference) + "</p><p><a href='" + escapeHtml_(cancelUrl) +
          "' style='display:inline-block;padding:12px 18px;background:#b42318;color:#fff;text-decoration:none;border-radius:8px;font-weight:bold'>Annuler mon rendez-vous</a></p>" +
          "<p>Si vous souhaitez conserver le rendez-vous, aucune action n'est nécessaire.</p></div>";
        MailApp.sendEmail({ to: email, subject: "Rappel : votre rendez-vous est demain — " + reference, body: body, htmlBody: html, name: "Cabinet de kinésithérapie" });
        sheet.getRange(rowNumber, headers["Rappel envoyé le"]).setValue(new Date());
      }
    } catch (err) {
      console.error("Reminder email failed for " + reference + ": " + err);
    } finally {
      sheet.getRange(rowNumber, headers["ID déclencheur rappel"]).clearContent();
      const trigger = ScriptApp.getProjectTriggers().filter(function (t) { return t.getUniqueId() === triggerUid; })[0];
      if (trigger) ScriptApp.deleteTrigger(trigger);
    }
    return;
  }
}

function cancelAppointment_(referenceValue, tokenValue) {
  const reference = cleanText_(referenceValue, 30);
  const token = cleanText_(tokenValue, 100);
  if (!/^RDV-[A-F0-9]{8}$/.test(reference) || !/^[a-f0-9]{64}$/.test(token)) throw new Error("Lien d’annulation invalide.");

  let details;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getBookingSheet_();
    const headers = getHeaderMap_(sheet);
    const rowNumber = findBookingRow_(sheet, reference);
    if (rowNumber < 2) throw new Error("Rendez-vous introuvable.");
    const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const savedToken = String(row[headers["Jeton annulation"] - 1] || "");
    if (!savedToken || savedToken !== token) throw new Error("Lien d’annulation invalide ou expiré.");
    const status = String(row[headers["Statut"] - 1] || "");
    if (status === "Annulée") return "Ce rendez-vous est déjà annulé.";
    if (status !== "Confirmé") throw new Error("Ce rendez-vous ne peut plus être annulé en ligne.");

    const triggerId = String(row[headers["ID déclencheur rappel"] - 1] || "");
    const calendarEventId = String(row[headers["ID événement calendrier"] - 1] || "");
    sheet.getRange(rowNumber, headers["Statut"]).setValue("Annulée");
    sheet.getRange(rowNumber, headers["Jeton annulation"]).clearContent();
    sheet.getRange(rowNumber, headers["ID déclencheur rappel"]).clearContent();

    if (triggerId) {
      ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getUniqueId() === triggerId) ScriptApp.deleteTrigger(trigger);
      });
    }

    details = {
      reference: reference,
      name: String(row[headers["Nom et prénom"] - 1] || ""),
      phone: String(row[headers["Téléphone"] - 1] || ""),
      email: String(row[headers["E-mail"] - 1] || "").trim(),
      service: String(row[headers["Motif"] - 1] || ""),
      date: normalizeDate_(row[headers["Date confirmée"] - 1]),
      time: normalizeTime_(row[headers["Heure confirmée"] - 1]),
      notes: String(row[headers["Message"] - 1] || ""),
      calendarEventId: calendarEventId
    };
  } finally {
    lock.releaseLock();
  }

  // Status=Annulée immediately frees the slot because availability only counts
  // rows with status "Confirmé".
  if (details.calendarEventId) {
    try {
      const calendar = CalendarApp.getCalendarById(BOOKING_SETTINGS.CALENDAR_ID);
      const event = calendar && calendar.getEventById(details.calendarEventId);
      if (event) event.deleteEvent();
    } catch (err) {
      console.error("Could not remove calendar event for " + reference + ": " + err);
    }
  }
  sendCancellationNotices_(details);
  return "Votre rendez-vous a été annulé. Le créneau est de nouveau disponible et le cabinet a été informé.";
}

function sendCancellationNotices_(data) {
  const dateLabel = formatDateFr_(data.date);
  const body = "Rendez-vous annulé\n\nRéférence : " + data.reference + "\nPatient : " + data.name +
    "\nMotif : " + data.service + "\nDate : " + dateLabel + "\nHeure : " + data.time;
  const html = "<div style='font-family:Arial,sans-serif;line-height:1.6'><h2>Rendez-vous annulé</h2><p><b>Référence :</b> " +
    escapeHtml_(data.reference) + "<br><b>Patient :</b> " + escapeHtml_(data.name) + "<br><b>Motif :</b> " +
    escapeHtml_(data.service) + "<br><b>Date :</b> " + escapeHtml_(dateLabel) + "<br><b>Heure :</b> " + escapeHtml_(data.time) + "</p></div>";
  const owner = getOwnerEmail_();
  if (owner) {
    try { MailApp.sendEmail({ to: owner, subject: "ANNULATION de rendez-vous — " + data.reference, body: body, htmlBody: html }); }
    catch (err) { console.error("Admin cancellation email failed: " + err); }
  }
  if (data.email) {
    try {
      MailApp.sendEmail({
        to: data.email,
        subject: "Confirmation d’annulation — " + data.reference,
        body: "Bonjour " + data.name + ",\n\nVotre rendez-vous a bien été annulé.\n\n" + body + "\n\nLe créneau est libéré. Si vous souhaitez un autre rendez-vous, vous pouvez utiliser : " + BOOKING_SETTINGS.WEBSITE_URL,
        htmlBody: "<div style='font-family:Arial,sans-serif;line-height:1.6'><h2>Votre rendez-vous a été annulé</h2><p>Bonjour " +
          escapeHtml_(data.name) + ", votre annulation a bien été enregistrée.</p>" + html +
          "<p>Le créneau est libéré. Pour réserver un autre rendez-vous, <a href='" + BOOKING_SETTINGS.WEBSITE_URL + "'>ouvrez le site de réservation</a>.</p></div>"
      });
    } catch (err) { console.error("Client cancellation email failed: " + err); }
  }
}

function getOwnerEmail_() {
  const configured = String(BOOKING_SETTINGS.OWNER_EMAIL || "").trim();
  if (configured) return configured;
  try { return String(Session.getEffectiveUser().getEmail() || "").trim(); } catch (err) { return ""; }
}

function getBookingSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(BOOKING_SETTINGS.SPREADSHEET_PROPERTY);
  if (!id) throw new Error("Le système n’est pas initialisé. Exécutez setupPhysiotherapySystem() dans Apps Script.");
  const spreadsheet = SpreadsheetApp.openById(id);
  const sheet = spreadsheet.getSheetByName(BOOKING_SETTINGS.SHEET_NAME);
  if (!sheet) throw new Error("L’onglet des demandes est introuvable. Exécutez setupPhysiotherapySystem().");
  ensureBookingColumns_(sheet);
  return sheet;
}

function appendBookingRow_(sheet, valuesByHeader) {
  const headers = getHeaderMap_(sheet);
  const row = new Array(sheet.getLastColumn()).fill("");
  Object.keys(valuesByHeader).forEach(function (name) {
    if (headers[name]) row[headers[name] - 1] = valuesByHeader[name];
  });
  sheet.appendRow(row);
}

function saveBookingField_(reference, fieldName, value) {
  const sheet = getBookingSheet_();
  const headers = getHeaderMap_(sheet);
  const row = findBookingRow_(sheet, reference);
  if (row >= 2 && headers[fieldName]) sheet.getRange(row, headers[fieldName]).setValue(value);
}

function ensureBookingColumns_(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS]);
    return;
  }
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  BOOKING_HEADERS.forEach(function (header) {
    if (existing.indexOf(header) === -1) {
      const newColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColumn).setValue(header).setFontWeight("bold");
      existing.push(header);
    }
  });
}

function getHeaderMap_(sheet) {
  const names = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  names.forEach(function (name, index) { if (name) map[name] = index + 1; });
  return map;
}

function findBookingRow_(sheet, reference) {
  const headers = getHeaderMap_(sheet);
  const refColumn = headers["Référence"];
  if (!refColumn || sheet.getLastRow() < 2) return -1;
  const found = sheet.getRange(2, refColumn, sheet.getLastRow() - 1, 1).createTextFinder(reference).matchEntireCell(true).findNext();
  return found ? found.getRow() : -1;
}

function appointmentDateTime_(date, time) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^([01]\d|2[0-3]):([0-5]\d)$/.test(time) || !isRealDate_(date)) {
    throw new Error("Date ou heure de rendez-vous invalide.");
  }
  // Algeria is UTC+1. Explicit offset avoids dependence on server timezone.
  return new Date(date + "T" + time + ":00+01:00");
}

function getCancellationUrl_(reference, token) {
  return BOOKING_SETTINGS.CANCEL_PAGE_URL + "?ref=" + encodeURIComponent(reference) + "&token=" + encodeURIComponent(token);
}

function isRealDate_(value) {
  const parts = value.split("-").map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  return date.getUTCFullYear() === parts[0] && date.getUTCMonth() === parts[1] - 1 && date.getUTCDate() === parts[2];
}

function normalizeDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  const text = String(value == null ? "" : value).trim();
  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : text;
}

function normalizeTime_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, BOOKING_SETTINGS.TIME_ZONE, "HH:mm");
  const text = String(value == null ? "" : value).trim();
  const match = text.match(/^([0-2]?\d):([0-5]\d)/);
  return match ? String(match[1]).padStart(2, "0") + ":" + match[2] : text;
}

function minutesToTime_(minutes) {
  return String(Math.floor(minutes / 60)).padStart(2, "0") + ":" + String(minutes % 60).padStart(2, "0");
}

function formatDateFr_(date) {
  const parts = String(date).split("-").map(Number);
  return Utilities.formatDate(new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 12, 0, 0)), BOOKING_SETTINGS.TIME_ZONE, "d MMMM yyyy");
}

function cleanText_(value, maxLength) {
  if (typeof value !== "string" && typeof value !== "number") return "";
  return String(value).replace(/[\u0000-\u001F\u007F]/g, " ").trim().slice(0, maxLength);
}

function safeCell_(value) {
  const text = String(value == null ? "" : value);
  return /^[=+@\-]/.test(text) ? "'" + text : text;
}

function escapeHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
