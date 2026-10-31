(() => {
  "use strict";

  const form = document.getElementById("appointment-form");
  if (!form) return;

  const steps = Array.from(document.querySelectorAll(".form-step"));
  const backButton = document.getElementById("back-button");
  const nextButton = document.getElementById("next-button");
  const submitButton = document.getElementById("submit-button");
  const stepCounter = document.getElementById("step-counter");
  const progressTrack = document.getElementById("progress-track");
  const progressFill = document.getElementById("progress-fill");
  const alertBox = document.getElementById("form-alert");
  const reviewList = document.getElementById("review-list");
  const successPanel = document.getElementById("success-panel");
  const formActions = document.getElementById("form-actions");
  const totalSteps = steps.length;
  let currentStep = 1;
  let submitting = false;
  let availabilityLoading = false;
  let availabilityRequest = 0;
  let lastSlotDate = "";
  let lastSlotService = "";

  const config = window.BOOKING_CONFIG || {};
  const labels = {
    service: "Motif du rendez-vous",
    preferredDate: "Date souhaitée",
    preferredTime: "Heure du rendez-vous",
    fullName: "Nom et prénom",
    phone: "Téléphone",
    email: "E-mail",
    notes: "Message"
  };

  function localDateString(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
  }

  const dateInput = document.getElementById("preferred-date");
  const timeSelect = document.getElementById("preferred-time");
  const slotStatus = document.getElementById("slot-status");
  if (dateInput) {
    dateInput.min = localDateString(new Date());
    dateInput.addEventListener("change", () => loadAvailableSlots());
  }

  function showAlert(message) {
    alertBox.textContent = message;
    alertBox.hidden = false;
    alertBox.scrollIntoView({ block: "nearest" });
  }

  function clearAlert() {
    alertBox.textContent = "";
    alertBox.hidden = true;
  }

  function showStep(number) {
    currentStep = Math.min(Math.max(number, 1), totalSteps);
    steps.forEach((step) => {
      const active = Number(step.dataset.step) === currentStep;
      step.hidden = !active;
      step.classList.toggle("is-active", active);
    });
    stepCounter.textContent = "Étape " + currentStep + " sur " + totalSteps;
    progressFill.style.width = ((currentStep / totalSteps) * 100) + "%";
    progressTrack.setAttribute("aria-valuenow", String(currentStep));
    backButton.disabled = currentStep === 1 || submitting;
    nextButton.hidden = currentStep === totalSteps;
    submitButton.hidden = currentStep !== totalSteps;
    nextButton.disabled = submitting;
    submitButton.disabled = submitting;
    clearAlert();
    const heading = document.getElementById("step-title-" + currentStep);
    if (heading) heading.setAttribute("tabindex", "-1");
    if (heading) heading.focus({ preventScroll: true });
    if (currentStep === 2) loadAvailableSlots();
  }

  function validateCurrentStep() {
    clearAlert();
    const activeStep = steps.find((step) => Number(step.dataset.step) === currentStep);
    if (!activeStep) return true;
    const controls = Array.from(activeStep.querySelectorAll("input, textarea, select"));
    for (const control of controls) {
      if (control.type === "radio") {
        const group = controls.filter((item) => item.type === "radio" && item.name === control.name);
        if (control.required && !group.some((item) => item.checked)) {
          showAlert("Veuillez sélectionner une option pour continuer.");
          control.focus();
          return false;
        }
        continue;
      }
      if (control.type === "checkbox") {
        if (control.required && !control.checked) {
          showAlert("Veuillez confirmer votre accord avant d’envoyer la demande.");
          control.focus();
          return false;
        }
        continue;
      }
      if (control.required && !control.value.trim()) {
        showAlert("Veuillez remplir le champ « " + (control.labels && control.labels[0] ? control.labels[0].innerText.replace("*", "").trim() : control.name) + " ».");
        control.focus();
        return false;
      }
      if (control.value.trim() && !control.checkValidity()) {
        showAlert(control.type === "email" ? "Veuillez vérifier votre adresse e-mail." : "Veuillez vérifier les informations saisies.");
        control.focus();
        return false;
      }
    }
    if (currentStep === 2) {
      const chosenDate = dateInput.value;
      if (chosenDate && chosenDate < localDateString(new Date())) {
        showAlert("Veuillez choisir une date à partir d’aujourd’hui.");
        dateInput.focus();
        return false;
      }
      if (availabilityLoading) {
        showAlert("Veuillez attendre le chargement des créneaux disponibles.");
        return false;
      }
      if (!timeSelect || !timeSelect.value) {
        showAlert("Veuillez sélectionner une heure disponible avant de continuer.");
        if (timeSelect) timeSelect.focus();
        return false;
      }
    }
    if (currentStep === 3) {
      const phone = document.getElementById("phone");
      const digits = phone.value.replace(/\D/g, "");
      if (digits.length < 8 || digits.length > 15) {
        showAlert("Veuillez saisir un numéro de téléphone valide.");
        phone.focus();
        return false;
      }
    }
    return true;
  }

  function selectedValue(name) {
    if (name === "preferredTime") return timeSelect ? timeSelect.value : "";
    const selected = form.querySelector('input[name="' + name + '"]:checked');
    return selected ? selected.value : "";
  }

  async function loadAvailableSlots() {
    if (!timeSelect || !slotStatus || !dateInput) return;
    const requestId = ++availabilityRequest;
    const date = dateInput.value;
    const service = selectedValue("service");
    const previousSelection = date === lastSlotDate && service === lastSlotService ? timeSelect.value : "";
    timeSelect.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "";

    if (!date || !service) {
      placeholder.textContent = "Choisissez d’abord une date";
      timeSelect.append(placeholder);
      timeSelect.disabled = true;
      slotStatus.textContent = "Choisissez une date pour afficher les heures disponibles.";
      availabilityLoading = false;
      return;
    }

    placeholder.textContent = "Chargement des créneaux…";
    timeSelect.append(placeholder);
    timeSelect.disabled = true;
    slotStatus.textContent = "Recherche des heures disponibles…";
    availabilityLoading = true;

    try {
      const endpoint = typeof config.appsScriptUrl === "string" ? config.appsScriptUrl.trim() : "";
      if (!endpoint || !/^https:\/\/script\.google\.com\/macros\/s\//.test(endpoint)) {
        throw new Error("La connexion au service de réservation n’est pas configurée.");
      }
      const response = await fetch(endpoint, {
        method: "POST",
        redirect: "follow",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "availability", preferredDate: date, service: service })
      });
      if (!response.ok) throw new Error("Le service de réservation est temporairement indisponible.");
      const result = await response.json();
      if (!result || result.success !== true) {
        throw new Error(result && result.message ? result.message : "Impossible de charger les créneaux.");
      }
      if (requestId !== availabilityRequest) return;

      const slots = Array.isArray(result.availableSlots) ? result.availableSlots : [];
      timeSelect.replaceChildren();
      const firstOption = document.createElement("option");
      firstOption.value = "";
      firstOption.textContent = slots.length ? "Sélectionnez une heure" : "Aucun créneau disponible";
      timeSelect.append(firstOption);
      slots.forEach((slot) => {
        const option = document.createElement("option");
        option.value = slot;
        option.textContent = slot;
        if (slot === previousSelection) option.selected = true;
        timeSelect.append(option);
      });
      timeSelect.disabled = slots.length === 0;
      lastSlotDate = date;
      lastSlotService = service;
      slotStatus.textContent = result.message || (slots.length
        ? "Ces créneaux sont disponibles au moment de la vérification."
        : "Aucun créneau disponible. Essayez une autre date.");
    } catch (error) {
      if (requestId !== availabilityRequest) return;
      timeSelect.replaceChildren();
      const errorOption = document.createElement("option");
      errorOption.value = "";
      errorOption.textContent = "Créneaux indisponibles";
      timeSelect.append(errorOption);
      timeSelect.disabled = true;
      slotStatus.textContent = error instanceof Error ? error.message : "Impossible de charger les créneaux.";
    } finally {
      if (requestId === availabilityRequest) availabilityLoading = false;
    }
  }

  function getFormData() {
    return {
      service: selectedValue("service"),
      preferredDate: document.getElementById("preferred-date").value,
      preferredTime: selectedValue("preferredTime"),
      fullName: document.getElementById("full-name").value.trim(),
      phone: document.getElementById("phone").value.trim(),
      email: document.getElementById("email").value.trim(),
      notes: document.getElementById("notes").value.trim(),
      privacyConsent: document.getElementById("privacy-consent").checked,
      source: "rendevo-physiotherapy",
      timezone: "Africa/Algiers",
      submittedAt: new Date().toISOString()
    };
  }

  function readableDate(value) {
    if (!value) return "Non précisée";
    const date = new Date(value + "T12:00:00");
    return new Intl.DateTimeFormat("fr-FR", { day: "numeric", month: "long", year: "numeric" }).format(date);
  }

  function renderReview() {
    const data = getFormData();
    const rows = [
      [labels.service, data.service],
      [labels.preferredDate, readableDate(data.preferredDate)],
      [labels.preferredTime, data.preferredTime],
      [labels.fullName, data.fullName],
      [labels.phone, data.phone],
      [labels.email, data.email || "Non renseignée"],
      [labels.notes, data.notes || "Aucun message"]
    ];
    reviewList.replaceChildren();
    rows.forEach(([label, value]) => {
      const row = document.createElement("div");
      row.className = "review-row";
      const term = document.createElement("dt");
      term.textContent = label;
      const detail = document.createElement("dd");
      detail.textContent = value;
      row.append(term, detail);
      reviewList.append(row);
    });
  }

  async function sendRequest(data) {
    const endpoint = typeof config.appsScriptUrl === "string" ? config.appsScriptUrl.trim() : "";
    if (!endpoint || !/^https:\/\/script\.google\.com\/macros\/s\//.test(endpoint)) {
      throw new Error("La connexion au service de réservation n’est pas encore configurée. Votre demande n’a pas été envoyée.");
    }
    const response = await fetch(endpoint, {
      method: "POST",
      redirect: "follow",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(data)
    });
    if (!response.ok) throw new Error("Le service de réservation est temporairement indisponible. Veuillez réessayer.");
    let result;
    try {
      result = await response.json();
    } catch (_) {
      throw new Error("La réponse du service est illisible. Votre demande n’a pas pu être confirmée.");
    }
    if (!result || result.success !== true) {
      throw new Error(result && result.message ? result.message : "La demande n’a pas été acceptée. Veuillez réessayer.");
    }
    return result;
  }

  function setSubmitting(value) {
    submitting = value;
    backButton.disabled = value || currentStep === 1;
    nextButton.disabled = value;
    submitButton.disabled = value;
    submitButton.textContent = value ? "Confirmation en cours…" : "Confirmer mon rendez-vous →";
  }

  backButton.addEventListener("click", () => {
    if (!submitting && currentStep > 1) showStep(currentStep - 1);
  });

  nextButton.addEventListener("click", () => {
    if (!validateCurrentStep()) return;
    if (currentStep < totalSteps) {
      if (currentStep + 1 === totalSteps) renderReview();
      showStep(currentStep + 1);
    }
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (currentStep !== totalSteps || submitting || !validateCurrentStep()) return;
    const data = getFormData();
    setSubmitting(true);
    clearAlert();
    try {
      const result = await sendRequest(data);
      form.hidden = true;
      formActions.hidden = true;
      document.querySelector(".progress-track").hidden = true;
      document.querySelector(".card-topline").hidden = true;
      successPanel.hidden = false;
      document.getElementById("success-message").textContent =
        result.message || "Votre rendez-vous est confirmé.";
      if (result.reference) {
        const ref = document.getElementById("request-reference");
        ref.textContent = "Référence de réservation : " + String(result.reference);
        ref.hidden = false;
      }
      const cancelNote = document.getElementById("cancel-note");
      const cancelLink = document.getElementById("cancel-link");
      if (result.cancelUrl && cancelNote && cancelLink) {
        cancelLink.href = result.cancelUrl;
        cancelNote.hidden = false;
      }
      successPanel.scrollIntoView({ block: "nearest" });
    } catch (error) {
      showAlert(error instanceof Error ? error.message : "Une erreur est survenue. Veuillez réessayer.");
    } finally {
      setSubmitting(false);
    }
  });

  document.getElementById("new-request-button").addEventListener("click", () => {
    form.reset();
    form.hidden = false;
    formActions.hidden = false;
    document.querySelector(".progress-track").hidden = false;
    document.querySelector(".card-topline").hidden = false;
    successPanel.hidden = true;
    document.getElementById("request-reference").hidden = true;
    document.getElementById("notes-count").textContent = "0 / 1000";
    lastSlotDate = "";
    lastSlotService = "";
    if (timeSelect) { timeSelect.replaceChildren(); timeSelect.disabled = true; }
    if (slotStatus) slotStatus.textContent = "Choisissez une date pour afficher les heures disponibles.";
    showStep(1);
  });

  document.getElementById("notes").addEventListener("input", (event) => {
    document.getElementById("notes-count").textContent = event.target.value.length + " / 1000";
  });

  const clinicName = typeof config.clinicName === "string" && config.clinicName.trim() ? config.clinicName.trim() : "Cabinet médical";
  document.getElementById("clinic-name").textContent = clinicName;
  document.getElementById("footer-clinic-name").textContent = clinicName;
  document.title = "Prendre rendez-vous | " + clinicName;
  document.getElementById("current-year").textContent = String(new Date().getFullYear());

  const phone = typeof config.clinicPhone === "string" ? config.clinicPhone.trim() : "";
  const phoneLink = document.getElementById("header-phone");
  if (phone) {
    phoneLink.href = "tel:" + phone.replace(/[^+\d]/g, "");
    phoneLink.querySelector("strong").textContent = phone;
  } else {
    phoneLink.hidden = true;
  }

  showStep(1);
})();
