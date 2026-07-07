const sidebarToggle = document.getElementById("sidebarToggle");

if (localStorage.getItem("sidebarCollapsed") === "1") {
  document.body.classList.add("sidebar-collapsed");
}

sidebarToggle?.addEventListener("click", () => {
  document.body.classList.toggle("sidebar-collapsed");
  localStorage.setItem(
    "sidebarCollapsed",
    document.body.classList.contains("sidebar-collapsed") ? "1" : "0",
  );
});

function renderDownloadsInto(container, files) {
  if (!container) return;
  container.innerHTML = "";
  for (const file of files) {
    const link = document.createElement("a");
    link.className = "download";
    link.href = file.url;
    link.innerHTML = `<strong>${file.label}</strong><span>Descargar</span>`;
    container.appendChild(link);
  }
}

const monthNames = {
  "01": "Ene",
  "02": "Feb",
  "03": "Mar",
  "04": "Abr",
  "05": "May",
  "06": "Jun",
  "07": "Jul",
  "08": "Ago",
  "09": "Sep",
  "10": "Oct",
  "11": "Nov",
  "12": "Dic",
};

const xmForm = document.getElementById("runForm");

if (xmForm) {
  const statusBox = document.getElementById("status");
  const downloads = document.getElementById("downloads");
  const submitBtn = document.getElementById("submitBtn");
  const summaryInput = document.getElementById("summary");
  const instancesInput = document.getElementById("instances");
  const resourceInput = document.getElementById("resourceImage");
  const resourcePreview = document.getElementById("resourcePreview");
  const monthFallback = document.getElementById("monthFallback");
  const monthDetected = document.querySelector("#monthDetected strong");
  const duplicateModal = document.getElementById("duplicateModal");
  const duplicateMessage = document.getElementById("duplicateMessage");
  const updateExisting = document.getElementById("updateExisting");
  const addDuplicate = document.getElementById("addDuplicate");
  const cancelDuplicate = document.getElementById("cancelDuplicate");
  const processModal = document.getElementById("processModal");
  const processMessage = document.getElementById("processMessage");
  const processSteps = [...document.querySelectorAll(".process-step")];

  let pendingFormData = null;
  let progressTimer = null;
  let selectedResourceImages = [];

  function setStatus(text, kind = "") {
    statusBox.className = `status${kind ? ` ${kind}` : ""}`;
    statusBox.textContent = text;
  }

  function inferMonth(fileName) {
    const match = fileName.match(/(20\d{2})[-_](0[1-9]|1[0-2])/);
    if (!match) return null;
    return `${monthNames[match[2]]}-${match[1].slice(-2)}`;
  }

  function updateMonthState() {
    const months = [summaryInput.files[0], instancesInput.files[0]]
      .filter(Boolean)
      .map((file) => inferMonth(file.name))
      .filter(Boolean);
    const uniqueMonths = [...new Set(months)];

    if (uniqueMonths.length === 1) {
      monthDetected.textContent = uniqueMonths[0];
      monthFallback.classList.add("hidden");
      return;
    }

    monthDetected.textContent = uniqueMonths.length > 1 ? "Meses distintos" : "No reconocido";
    monthFallback.classList.remove("hidden");
  }

  function renderImagePreview() {
    resourcePreview.innerHTML = "";
    selectedResourceImages.forEach((file, index) => {
      const card = document.createElement("div");
      card.className = "preview-card";

      const image = document.createElement("img");
      image.src = URL.createObjectURL(file);
      image.alt = file.name;
      image.onload = () => URL.revokeObjectURL(image.src);

      const label = document.createElement("span");
      label.textContent = `${index + 1}. ${file.name}`;

      const remove = document.createElement("button");
      remove.type = "button";
      remove.textContent = "Borrar";
      remove.addEventListener("click", () => {
        selectedResourceImages.splice(index, 1);
        renderImagePreview();
      });

      card.append(image, label, remove);
      resourcePreview.appendChild(card);
    });
  }

  function isAcceptedResourceImage(file) {
    return file.type.startsWith("image/") || /\.(png|jpe?g)$/i.test(file.name);
  }

  async function normalizeResourceImage(file) {
    try {
      const bitmap = await createImageBitmap(file);
      const canvas = document.createElement("canvas");
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      const context = canvas.getContext("2d");
      context.drawImage(bitmap, 0, 0);
      bitmap.close();
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.9));
      if (!blob) return file;
      const cleanName = file.name.replace(/\.[^.]+$/, "") || "imagen_consumo_recurso";
      return new File([blob], `${cleanName}.jpg`, { type: "image/jpeg" });
    } catch {
      return file;
    }
  }

  async function appendResourceImages(files) {
    for (const file of files) {
      if (!isAcceptedResourceImage(file)) continue;
      selectedResourceImages.push(await normalizeResourceImage(file));
    }
    resourceInput.value = "";
    renderImagePreview();
  }

  function buildFormData() {
    const formData = new FormData(xmForm);
    formData.delete("resource_images");
    selectedResourceImages.forEach((file) => {
      formData.append("resource_images", file, file.name);
    });
    return formData;
  }

  function stepByName(name) {
    return processSteps.find((step) => step.dataset.step === name);
  }

  function setStep(name, state) {
    const step = stepByName(name);
    if (!step) return;
    step.classList.remove("active", "done", "error");
    if (state) step.classList.add(state);
  }

  function startProgress(hasComparativo) {
    clearInterval(progressTimer);
    processSteps.forEach((step) => step.classList.remove("active", "done", "error"));
    processMessage.textContent = "Preparando archivos para la automatizacion...";
    processModal.classList.remove("hidden");

    setStep("summary", "active");
    setStep("instances", "active");
    if (hasComparativo) {
      setStep("comparativo", "active");
      setStep("word", "active");
    }

    let elapsed = 0;
    progressTimer = setInterval(() => {
      elapsed += 15;
      processMessage.textContent = hasComparativo
        ? `Proceso activo (${elapsed}s). Agregando datos al comparativo y preparando el informe.`
        : `Proceso activo (${elapsed}s). Generando archivos XM.`;
    }, 15000);
  }

  function finishProgress(files) {
    clearInterval(progressTimer);
    const labels = files.map((file) => file.label.toLowerCase());
    setStep("summary", labels.some((label) => label.includes("summary")) ? "done" : "");
    setStep("instances", labels.some((label) => label.includes("instances")) ? "done" : "");
    setStep("comparativo", labels.some((label) => label.includes("comparativo")) ? "done" : "");
    setStep("word", labels.some((label) => label.endsWith(".docx")) ? "done" : "");
    processMessage.textContent = "Proceso completado. Ya puedes descargar los archivos.";
    setTimeout(() => processModal.classList.add("hidden"), 1200);
  }

  function failProgress(message) {
    clearInterval(progressTimer);
    const active = processSteps.find((step) => step.classList.contains("active"));
    if (active) active.classList.replace("active", "error");
    processMessage.textContent = message;
  }

  async function submitFormData(formData) {
    downloads.innerHTML = "";
    submitBtn.disabled = true;
    setStatus("Procesando archivos...");
    startProgress(Boolean(formData.get("comparativo")?.name));

    try {
      const response = await fetch("/run", {
        method: "POST",
        body: formData,
      });
      const raw = await response.text();
      let data;
      try {
        data = JSON.parse(raw);
      } catch {
        throw new Error(raw || "El servidor termino sin devolver una respuesta valida.");
      }
      if (!response.ok || !data.ok) {
        throw new Error(data.error || "No se pudo generar la salida.");
      }
      if (data.needs_confirmation) {
        pendingFormData = formData;
        duplicateMessage.textContent = `${data.message} Puedes actualizar la informacion de ese mes o agregar otro bloque.`;
        duplicateModal.classList.remove("hidden");
        processModal.classList.add("hidden");
        clearInterval(progressTimer);
        setStatus("Se requiere confirmacion para continuar.");
        return;
      }
      setStatus(`Archivos generados correctamente para ${data.month}.`, "ok");
      renderDownloadsInto(downloads, data.files);
      finishProgress(data.files);
    } catch (error) {
      const message = error.message;
      setStatus(message, "err");
      failProgress(message);
    } finally {
      submitBtn.disabled = false;
    }
  }

  function closeDuplicateModal() {
    duplicateModal.classList.add("hidden");
  }

  xmForm.addEventListener("reset", () => {
    downloads.innerHTML = "";
    selectedResourceImages = [];
    renderImagePreview();
    setStatus("Selecciona los archivos XM para comenzar.");
    setTimeout(updateMonthState, 0);
  });

  xmForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    await submitFormData(buildFormData());
  });

  summaryInput.addEventListener("change", updateMonthState);
  instancesInput.addEventListener("change", updateMonthState);
  resourceInput.addEventListener("change", async () => {
    setStatus("Preparando imagenes adjuntas...");
    await appendResourceImages(resourceInput.files);
    setStatus("Imagenes listas para enviar.");
  });

  updateExisting.addEventListener("click", async () => {
    if (!pendingFormData) return;
    closeDuplicateModal();
    pendingFormData.set("duplicate_action", "actualizar");
    await submitFormData(pendingFormData);
  });

  addDuplicate.addEventListener("click", async () => {
    if (!pendingFormData) return;
    closeDuplicateModal();
    pendingFormData.set("duplicate_action", "agregar");
    await submitFormData(pendingFormData);
  });

  cancelDuplicate.addEventListener("click", () => {
    pendingFormData = null;
    closeDuplicateModal();
    setStatus("Operacion cancelada.");
  });

  updateMonthState();
}

const helpdeskForm = document.getElementById("helpdeskForm");

if (helpdeskForm) {
  const helpdeskStatus = document.getElementById("helpdeskStatus");
  const helpdeskDownloads = document.getElementById("helpdeskDownloads");
  const helpdeskSubmitBtn = document.getElementById("helpdeskSubmitBtn");

  function setHelpdeskStatus(text, kind = "") {
    helpdeskStatus.className = `status${kind ? ` ${kind}` : ""}`;
    helpdeskStatus.textContent = text;
  }

  helpdeskForm.addEventListener("reset", () => {
    helpdeskDownloads.innerHTML = "";
    setHelpdeskStatus("Configura el rango de fechas para consultar Freshdesk.");
  });

  helpdeskForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    helpdeskDownloads.innerHTML = "";
    helpdeskSubmitBtn.disabled = true;
    setHelpdeskStatus("Consultando Freshdesk y generando informe...");

    try {
      const response = await fetch("/helpdesk", {
        method: "POST",
        body: new FormData(helpdeskForm),
      });
      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || "No se pudo generar el informe Helpdesk.");
      }
      setHelpdeskStatus(
        `Informe generado correctamente. Responsable: Manuel Alcalá. Tickets encontrados: ${data.count}. Nuevos agregados: ${data.added_count}.`,
        "ok",
      );
      renderDownloadsInto(helpdeskDownloads, data.files);
    } catch (error) {
      setHelpdeskStatus(error.message, "err");
    } finally {
      helpdeskSubmitBtn.disabled = false;
    }
  });
}
