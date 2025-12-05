// Estado simple en memoria para múltiples auditorías
const state = {
  audits: {},
  currentAuditCode: "",
  roles: [ // roles dinámicos que pueden editarse
    { id: "auditor_lider", label: "Auditor líder" },
    { id: "auditor", label: "Auditor" },
    { id: "responsable_proceso", label: "Responsable de proceso" },
    { id: "observador", label: "Observador" },
  ],
};

// Añadimos imports de jspdf y html2canvas (usando importmap)
import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";

// Helpers
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function uuid() {
  return crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2);
}

// Debounce helper para agrupar renderizados frecuentes
function debounce(fn, wait = 120) {
  let t = null;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), wait);
  };
}

// Small helper to format date strings as DD/MM/YYYY for display in reports
function formatDateDisplay(dateStr) {
  if (!dateStr) return "?";
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// Modal helpers
let modalResolve = null;

function openTextPrompt({ title = "", message = "", placeholder = "", initialValue = "" } = {}) {
  const backdrop = $("#app-modal-backdrop");
  const titleEl = $("#modal-title");
  const msgEl = $("#modal-message");
  const inputWrapper = $("#modal-input-wrapper");
  const inputEl = $("#modal-input");
  const okBtn = $("#modal-ok");
  const cancelBtn = $("#modal-cancel");

  if (!backdrop || !titleEl || !msgEl || !inputWrapper || !inputEl || !okBtn || !cancelBtn) {
    // Fallback muy básico si por alguna razón no existe el modal
    const v = window.prompt(message || title, initialValue || "");
    return Promise.resolve(v || null);
  }

  titleEl.textContent = title;
  msgEl.textContent = message;
  inputWrapper.classList.remove("hidden");
  inputEl.value = initialValue || "";
  inputEl.placeholder = placeholder || "";
  backdrop.classList.remove("hidden");

  return new Promise((resolve) => {
    modalResolve = resolve;

    const close = (value) => {
      backdrop.classList.add("hidden");
      inputWrapper.classList.add("hidden");
      modalResolve = null;
      okBtn.removeEventListener("click", onOk);
      cancelBtn.removeEventListener("click", onCancel);
      inputEl.removeEventListener("keydown", onKey);
      resolve(value);
    };

    const onOk = () => {
      // Accept empty string as a valid value (user may want default name)
      const val = inputEl.value;
      // If user explicitly pressed OK, return the string (can be empty)
      close(val);
    };

    const onCancel = () => {
      close(null);
    };

    const onKey = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        onOk();
      } else if (e.key === "Escape") {
        e.preventDefault();
        onCancel();
      }
    };

    okBtn.addEventListener("click", onOk);
    cancelBtn.addEventListener("click", onCancel);
    inputEl.addEventListener("keydown", onKey);

    setTimeout(() => {
      inputEl.focus();
      inputEl.select();
    }, 0);
  });
}

/**
 * Abre un modal específico para editar un participante (nombre, correo, rol) y devuelve
 * la información editada o null si se cancela.
 */
function openConfirmDialog({ title = "", message = "" } = {}) {
  const backdrop = $("#app-modal-backdrop");
  const titleEl = $("#modal-title");
  const msgEl = $("#modal-message");
  const inputWrapper = $("#modal-input-wrapper");
  const okBtn = $("#modal-ok");
  const cancelBtn = $("#modal-cancel");

  if (!backdrop || !titleEl || !msgEl || !okBtn || !cancelBtn) {
    return Promise.resolve(false);
  }

  titleEl.textContent = title;
  msgEl.textContent = message;
  inputWrapper.classList.add("hidden");
  backdrop.classList.remove("hidden");

  // Cambiar textos de botones
  okBtn.textContent = "Confirmar";
  cancelBtn.textContent = "Cancelar";

  return new Promise((resolve) => {
    const close = (result) => {
      backdrop.classList.add("hidden");
      okBtn.removeEventListener("click", onOk);
      cancelBtn.removeEventListener("click", onCancel);
      document.removeEventListener("keydown", onKey);
      resolve(result);
    };

    const onOk = () => {
      close(true);
    };

    const onCancel = () => {
      close(false);
    };

    const onKey = (e) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onCancel();
      }
    };

    okBtn.addEventListener("click", onOk);
    cancelBtn.addEventListener("click", onCancel);
    document.addEventListener("keydown", onKey);
  });
}

function openParticipantEditor(participant = {}, roles = []) {
  const backdrop = $("#app-modal-backdrop");
  const titleEl = $("#modal-title");
  const msgEl = $("#modal-message");
  const inputWrapper = $("#modal-input-wrapper");
  const inputEl = $("#modal-input");
  const okBtn = $("#modal-ok");
  const cancelBtn = $("#modal-cancel");

  titleEl.textContent = "Editar participante";
  msgEl.textContent = "";
  // Limpiar wrapper y crear inputs personalizados
  inputWrapper.classList.remove("hidden");
  inputWrapper.innerHTML = ""; // prepare to inject multiple inputs

  // Nombre
  const nameInput = document.createElement("input");
  nameInput.className = "field-input";
  nameInput.type = "text";
  nameInput.placeholder = "Nombre completo";
  nameInput.value = participant.name || "";

  // Correo
  const emailInput = document.createElement("input");
  emailInput.className = "field-input";
  emailInput.type = "email";
  emailInput.placeholder = "Correo electrónico (opcional)";
  emailInput.value = participant.email || "";

  // Rol (select)
  const roleSelect = document.createElement("select");
  roleSelect.className = "field-input";
  roles.forEach(r => {
    const opt = document.createElement("option");
    opt.value = r.id;
    opt.textContent = r.label;
    roleSelect.appendChild(opt);
  });
  roleSelect.value = participant.role || (roles[0] ? roles[0].id : "");

  // Layout small spacing
  const wrapperInner = document.createElement("div");
  wrapperInner.style.display = "flex";
  wrapperInner.style.flexDirection = "column";
  wrapperInner.style.gap = "8px";
  wrapperInner.appendChild(nameInput);
  wrapperInner.appendChild(emailInput);
  wrapperInner.appendChild(roleSelect);
  inputWrapper.appendChild(wrapperInner);

  backdrop.classList.remove("hidden");

  return new Promise((resolve) => {
    const cleanUp = () => {
      backdrop.classList.add("hidden");
      inputWrapper.classList.add("hidden");
      inputWrapper.innerHTML = "";
      okBtn.removeEventListener("click", onOk);
      cancelBtn.removeEventListener("click", onCancel);
      nameInput.removeEventListener("keydown", onKey);
      emailInput.removeEventListener("keydown", onKey);
      roleSelect.removeEventListener("keydown", onKey);
    };

    const onOk = () => {
      const result = {
        name: nameInput.value.trim(),
        email: emailInput.value.trim(),
        role: roleSelect.value,
      };
      cleanUp();
      resolve(result);
    };
    const onCancel = () => {
      cleanUp();
      resolve(null);
    };
    const onKey = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        onOk();
      } else if (e.key === "Escape") {
        e.preventDefault();
        onCancel();
      }
    };

    okBtn.addEventListener("click", onOk);
    cancelBtn.addEventListener("click", onCancel);
    nameInput.addEventListener("keydown", onKey);
    emailInput.addEventListener("keydown", onKey);
    roleSelect.addEventListener("keydown", onKey);

    setTimeout(() => {
      nameInput.focus();
      nameInput.select();
    }, 0);
  });
}

function createEmptyAudit(code = "") {
  return {
    code,
    name: "",
    leadAuditorId: null,
    participants: [],
    stages: [],
  };
}

function getCurrentAudit() {
  if (!state.currentAuditCode) {
    return null;
  }
  if (!state.audits[state.currentAuditCode]) {
    state.audits[state.currentAuditCode] = createEmptyAudit(state.currentAuditCode);
  }
  return state.audits[state.currentAuditCode];
}

function setCurrentAuditCode(code) {
  state.currentAuditCode = code;
  if (!state.audits[code]) {
    state.audits[code] = createEmptyAudit(code);
  }
}

// Inicialización
window.addEventListener("DOMContentLoaded", () => {
  // Código por defecto
  const defaultCode = "AUD-001";
  setCurrentAuditCode(defaultCode);

  setupTabs();
  setupAuditBinding();
  setupParticipants();
  setupStages();
  setupCalendar();
  renderAll();
});

// Programador de render global (debounced)
const scheduleRender = debounce(() => {
  renderParticipants();
  renderStages();
  updateLeadAuditorOptions();
  updateSummary();
  renderCalendar();
}, 120);

function setupTabs() {
  const buttons = $$( ".tab-button" );
  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const tab = btn.dataset.tab;
      buttons.forEach((b) => b.classList.toggle("active", b === btn));
      $$( ".tab-content" ).forEach((c) => {
        c.classList.toggle("active", c.id === `tab-${tab}`);
      });
    });
  });
}

// Auditoría básica
function setupAuditBinding() {
  const codeInput = $("#audit-code");
  const nameInput = $("#audit-name");
  const leadSel = $("#lead-auditor");
  const auditSelector = $("#audit-selector");
  const saveBtn = $("#save-audit");
  const downloadBtn = $("#download-json");
  const sampleBtn = $("#load-sample");
  const loadJsonBtn = $("#load-json");

  const current = getCurrentAudit();
  if (current) {
    codeInput.value = current.code;
    nameInput.value = current.name;
  }

  codeInput.addEventListener("input", () => {
    // Solo cambia el código en memoria, se consolida al guardar
  });

  nameInput.addEventListener("input", () => {
    const audit = getCurrentAudit();
    if (!audit) return;
    audit.name = nameInput.value.trim();
    updateAuditSelector();
    scheduleRender();
  });

  leadSel.addEventListener("change", () => {
    const audit = getCurrentAudit();
    if (!audit) return;
    audit.leadAuditorId = leadSel.value || null;
  });

  saveBtn.addEventListener("click", () => {
    const newCode = codeInput.value.trim();
    if (!newCode) return;

    const oldCode = state.currentAuditCode;
    let audit = getCurrentAudit() || createEmptyAudit();

    if (!oldCode) {
      // Nunca se había definido
      state.currentAuditCode = newCode;
      audit.code = newCode;
      state.audits[newCode] = audit;
    } else if (oldCode !== newCode) {
      // Renombrar código
      delete state.audits[oldCode];
      audit.code = newCode;
      state.audits[newCode] = audit;
      state.currentAuditCode = newCode;
    } else {
      audit.code = newCode;
    }

    updateAuditSelector();
  });

  auditSelector.addEventListener("change", () => {
    const code = auditSelector.value;
    if (!code) {
      // Nueva auditoría
      const newCode = "AUD-" + (Object.keys(state.audits).length + 1).toString().padStart(3, "0");
      setCurrentAuditCode(newCode);
    } else {
      setCurrentAuditCode(code);
    }
    loadAuditIntoUI();
  });

  downloadBtn.addEventListener("click", () => {
    const payload = {
      audits: state.audits,
      currentAuditCode: state.currentAuditCode,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "auditor-monitor-data.json";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  sampleBtn.addEventListener("click", () => {
    loadSampleData();
  });

  // Cargar JSON desde archivo: abre selector de archivos, parsea y carga estado
  if (loadJsonBtn) {
    loadJsonBtn.addEventListener("click", () => {
      const fileInput = document.createElement("input");
      fileInput.type = "file";
      fileInput.accept = "application/json,application/*+json";
      fileInput.addEventListener("change", (e) => {
        const f = fileInput.files && fileInput.files[0];
        if (!f) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
          try {
            const parsed = JSON.parse(String(ev.target.result));
            // If structure includes audits and currentAuditCode, load directly
            if (parsed && typeof parsed === "object") {
              if (parsed.audits && typeof parsed.audits === "object") {
                state.audits = parsed.audits;
                state.currentAuditCode = parsed.currentAuditCode || Object.keys(state.audits)[0] || "";
              } else {
                // Backwards compat: assume file contains a single audit object
                if (parsed.code) {
                  state.audits = { [parsed.code]: parsed };
                  state.currentAuditCode = parsed.code;
                }
              }
            }
            updateAuditSelector();
            loadAuditIntoUI();
            scheduleRender();
          } catch (err) {
            console.error("Error parsing JSON file:", err);
            // mostrar fallback simple
            openTextPrompt({ title: "Error", message: "No se pudo leer el archivo JSON (archivo inválido)." }).then(()=>{});
          }
        };
        reader.readAsText(f);
      });
      fileInput.click();
    });
  }

  // Reset application button
  const resetBtn = $("#reset-app");
  if (resetBtn) {
    resetBtn.addEventListener("click", async () => {
      const confirmed = await openConfirmDialog({
        title: "Actualizar página",
        message: "¿Está seguro de que desea actualizar la página? Se recargará el navegador."
      });
      if (!confirmed) return;

      // Reload page
      location.reload();
    });
  }

  updateAuditSelector();
  loadAuditIntoUI();
}

function loadAuditIntoUI() {
  const audit = getCurrentAudit();
  const codeInput = $("#audit-code");
  const nameInput = $("#audit-name");

  if (!audit) {
    codeInput.value = "";
    nameInput.value = "";
    $("#lead-auditor").value = "";
    $("#participants-list").innerHTML = "";
    $("#stages-list").innerHTML = "";
    $("#calendar-list").innerHTML = "";
    updateSummary();
    return;
  }

  codeInput.value = audit.code || "";
  nameInput.value = audit.name || "";

  renderParticipants();
  updateLeadAuditorOptions();
  renderStages();
  renderCalendar();
  updateSummary();
}

function updateAuditSelector() {
  const selector = $("#audit-selector");
  if (!selector) return;

  const prev = selector.value;
  selector.innerHTML = "";

  const optNew = document.createElement("option");
  optNew.value = "";
  optNew.textContent = "Nueva auditoría";
  selector.appendChild(optNew);

  Object.values(state.audits)
    .sort((a, b) => (a.code || "").localeCompare(b.code || ""))
    .forEach((audit) => {
      if (!audit.code) return;
      const opt = document.createElement("option");
      opt.value = audit.code;
      opt.textContent = audit.name ? `${audit.code} - ${audit.name}` : audit.code;
      selector.appendChild(opt);
    });

  if (state.currentAuditCode && state.audits[state.currentAuditCode]) {
    selector.value = state.currentAuditCode;
  } else {
    selector.value = "";
  }
}

function updateLeadAuditorOptions() {
  const leadSel = $("#lead-auditor");
  const audit = getCurrentAudit();
  if (!leadSel || !audit) return;

  const prev = audit.leadAuditorId;
  leadSel.innerHTML = "";

  const optNone = document.createElement("option");
  optNone.value = "";
  optNone.textContent = "Selecciona auditor líder";
  leadSel.appendChild(optNone);

  audit.participants.forEach((p) => {
    if (["auditor_lider", "auditor"].includes(p.role)) {
      const opt = document.createElement("option");
      opt.value = p.id;
      opt.textContent = `${p.name} (${roleLabel(p.role)})`;
      leadSel.appendChild(opt);
    }
  });

  if (prev && audit.participants.some((p) => p.id === prev)) {
    leadSel.value = prev;
  } else {
    audit.leadAuditorId = null;
    leadSel.value = "";
  }
}

function renderRoleOptions(selectEl, selected) {
  if (!selectEl) return;
  selectEl.innerHTML = "";
  state.roles.forEach(r => {
    const opt = document.createElement("option");
    opt.value = r.id;
    opt.textContent = r.label;
    selectEl.appendChild(opt);
  });
  if (selected) selectEl.value = selected;
}

function setupParticipants() {
  // render initial options
  renderRoleOptions($("#participant-role"));

  $("#manage-roles").addEventListener("click", async () => {
    // show current roles as comma-separated "id:label"
    const current = state.roles.map(r => `${r.id}:${r.label}`).join(", ");
    const edited = await openTextPrompt({
      title: "Editar roles",
      message: "Edita la lista de roles como una lista separada por comas en formato id:Etiqueta. Ejemplo: auditor:Auditor,observador:Observador",
      placeholder: "id:Etiqueta,otro:Etiqueta2",
      initialValue: current
    });
    if (edited === null) return;
    const items = edited.split(",").map(s => s.trim()).filter(Boolean);
    const nextRoles = [];
    items.forEach(it => {
      const parts = it.split(":").map(p => p.trim());
      if (parts.length === 2 && parts[0]) {
        nextRoles.push({ id: parts[0], label: parts[1] || parts[0] });
      }
    });
    if (nextRoles.length === 0) return;
    // update roles
    state.roles = nextRoles;
    // ensure participants with removed roles fallback to 'observador' if exists else first role
    const fallback = state.roles.find(r => r.id === "observador") ? "observador" : (state.roles[0] ? state.roles[0].id : null);
    Object.values(state.audits).forEach(audit => {
      (audit.participants || []).forEach(p => {
        if (!state.roles.some(r => r.id === p.role)) {
          p.role = fallback;
        }
      });
    });
    // re-render role selects and participants
    renderRoleOptions($("#participant-role"));
    updateLeadAuditorOptions();
    scheduleRender();
  });

  $("#add-participant").addEventListener("click", () => {
    const audit = getCurrentAudit();
    if (!audit) return;

    const name = $("#participant-name").value.trim();
    const email = $("#participant-email").value.trim();
    const role = $("#participant-role").value;

    if (!name) return;

    audit.participants.push({
      id: uuid(),
      name,
      email,
      role,
    });

    $("#participant-name").value = "";
    $("#participant-email").value = "";

    // Usar scheduler para agrupar renders
    scheduleRender();
    updateLeadAuditorOptions();
    renderCalendar();
  });
}

function roleLabel(role) {
  switch (role) {
    case "auditor_lider": return "Auditor líder";
    case "auditor": return "Auditor";
    case "responsable_proceso": return "Responsable de proceso";
    case "observador": return "Observador";
    default: return role;
  }
}

function renderParticipants() {
  const audit = getCurrentAudit();
  const list = $("#participants-list");
  if (!list) return;

  list.innerHTML = "";

  if (!audit || audit.participants.length === 0) return;

  audit.participants.forEach((p) => {
    const li = document.createElement("li");
    li.className = "list-item";

    const main = document.createElement("div");
    main.className = "list-main";

    const title = document.createElement("span");
    title.className = "list-title";
    title.textContent = p.name;

    const sub = document.createElement("span");
    sub.className = "list-sub";
    sub.textContent = p.email || "Sin correo";

    const meta = document.createElement("div");
    meta.className = "list-meta-row";

    const badgeRole = document.createElement("span");
    badgeRole.className = "badge badge-role";
    badgeRole.textContent = roleLabel(p.role);

    meta.appendChild(badgeRole);
    main.appendChild(title);
    main.appendChild(sub);
    main.appendChild(meta);

    const actions = document.createElement("div");
    const del = document.createElement("button");
    del.className = "icon-btn";
    del.textContent = "🗑";
    del.title = "Eliminar participante";
    del.addEventListener("click", () => {
      const idx = audit.participants.findIndex((x) => x.id === p.id);
      if (idx !== -1) audit.participants.splice(idx, 1);
      scheduleRender();
      updateLeadAuditorOptions();
      renderCalendar();
    });

    const edit = document.createElement("button");
    edit.className = "icon-btn";
    edit.textContent = "✎";
    edit.title = "Editar participante";
    edit.addEventListener("click", async () => {
      // Abrir editor de participante con formulario (nombre, email, rol)
      const result = await openParticipantEditor(
        { name: p.name, email: p.email, role: p.role },
        state.roles
      );
      if (!result) return; // cancelado

      // Validaciones ligeras: nombre requerido
      if (result.name && result.name.trim()) {
        p.name = result.name.trim();
      }
      p.email = result.email ? result.email.trim() : "";
      if (result.role && state.roles.some(r => r.id === result.role)) {
        p.role = result.role;
      }

      // Asegurar que el select principal de agregar participante contenga los roles actualizados
      renderRoleOptions($("#participant-role"));

      scheduleRender();
      updateLeadAuditorOptions();
      renderCalendar();
    });

    actions.appendChild(edit);
    actions.appendChild(del);
    li.appendChild(main);
    li.appendChild(actions);
    list.appendChild(li);
  });
}

// Etapas
function setupStages() {
  $("#add-stage").addEventListener("click", () => {
    const audit = getCurrentAudit();
    if (!audit) return;

    const name = $("#stage-name").value.trim();
    const risk = $("#stage-risk").value;
    const filesInput = $("#stage-files");
    const files = Array.from(filesInput.files || []);

    if (!name) return;

    audit.stages.push({
      id: uuid(),
      name,
      risk,
      activities: [],
      attachments: files.map((f) => ({
        name: f.name,
        type: f.type,
      })),
      startDate: "",
      endDate: "",
    });

    $("#stage-name").value = "";
    filesInput.value = "";

    scheduleRender();
    renderCalendar();
  });
}

/**
 * Abre editor de etapa (nombre, riesgo, fechas) y devuelve objeto o null si cancela.
 */
function openStageEditor(stage = {}) {
  const backdrop = $("#app-modal-backdrop");
  const titleEl = $("#modal-title");
  const msgEl = $("#modal-message");
  const inputWrapper = $("#modal-input-wrapper");
  const okBtn = $("#modal-ok");
  const cancelBtn = $("#modal-cancel");

  titleEl.textContent = "Editar etapa";
  msgEl.textContent = "";
  inputWrapper.classList.remove("hidden");
  inputWrapper.innerHTML = "";

  const nameInput = document.createElement("input");
  nameInput.className = "field-input";
  nameInput.type = "text";
  nameInput.placeholder = "Nombre de la etapa";
  nameInput.value = stage.name || "";

  const riskSelect = document.createElement("select");
  riskSelect.className = "field-input";
  ["bajo", "medio", "alto"].forEach(r => {
    const o = document.createElement("option");
    o.value = r;
    o.textContent = `Riesgo ${r}`;
    riskSelect.appendChild(o);
  });
  riskSelect.value = stage.risk || "bajo";

  const startInput = document.createElement("input");
  startInput.type = "date";
  startInput.className = "field-input";
  startInput.value = stage.startDate || "";

  const endInput = document.createElement("input");
  endInput.type = "date";
  endInput.className = "field-input";
  endInput.value = stage.endDate || "";

  const wrapperInner = document.createElement("div");
  wrapperInner.style.display = "flex";
  wrapperInner.style.flexDirection = "column";
  wrapperInner.style.gap = "8px";
  wrapperInner.appendChild(nameInput);
  wrapperInner.appendChild(riskSelect);

  const datesRow = document.createElement("div");
  datesRow.style.display = "flex";
  datesRow.style.gap = "8px";
  datesRow.appendChild(startInput);
  datesRow.appendChild(endInput);
  wrapperInner.appendChild(datesRow);

  inputWrapper.appendChild(wrapperInner);
  backdrop.classList.remove("hidden");

  return new Promise((resolve) => {
    const clean = () => {
      backdrop.classList.add("hidden");
      inputWrapper.classList.add("hidden");
      inputWrapper.innerHTML = "";
      okBtn.removeEventListener("click", ok);
      cancelBtn.removeEventListener("click", cancel);
    };
    const ok = () => {
      const res = {
        name: nameInput.value.trim(),
        risk: riskSelect.value,
        startDate: startInput.value || "",
        endDate: endInput.value || "",
      };
      clean();
      resolve(res);
    };
    const cancel = () => {
      clean();
      resolve(null);
    };
    okBtn.addEventListener("click", ok);
    cancelBtn.addEventListener("click", cancel);
    setTimeout(()=>{ nameInput.focus(); nameInput.select(); }, 0);
  });
}

/**
 * Abre editor de actividad (nombre, fechas, descripción)
 */
function openActivityEditor(activity = {}) {
  const backdrop = $("#app-modal-backdrop");
  const titleEl = $("#modal-title");
  const msgEl = $("#modal-message");
  const inputWrapper = $("#modal-input-wrapper");
  const okBtn = $("#modal-ok");
  const cancelBtn = $("#modal-cancel");

  titleEl.textContent = "Editar actividad";
  msgEl.textContent = "";
  inputWrapper.classList.remove("hidden");
  inputWrapper.innerHTML = "";

  const nameInput = document.createElement("input");
  nameInput.className = "field-input";
  nameInput.type = "text";
  nameInput.placeholder = "Nombre de la actividad";
  nameInput.value = activity.name || "";

  const startInput = document.createElement("input");
  startInput.type = "date";
  startInput.className = "field-input";
  startInput.value = activity.startDate || "";

  const endInput = document.createElement("input");
  endInput.type = "date";
  endInput.className = "field-input";
  endInput.value = activity.endDate || "";

  const descInput = document.createElement("input");
  descInput.className = "field-input";
  descInput.type = "text";
  descInput.placeholder = "Descripción (opcional)";
  descInput.value = activity.description || "";

  const wrapperInner = document.createElement("div");
  wrapperInner.style.display = "flex";
  wrapperInner.style.flexDirection = "column";
  wrapperInner.style.gap = "8px";
  wrapperInner.appendChild(nameInput);

  const datesRow = document.createElement("div");
  datesRow.style.display = "flex";
  datesRow.style.gap = "8px";
  datesRow.appendChild(startInput);
  datesRow.appendChild(endInput);
  wrapperInner.appendChild(datesRow);
  wrapperInner.appendChild(descInput);

  inputWrapper.appendChild(wrapperInner);
  backdrop.classList.remove("hidden");

  return new Promise((resolve) => {
    const clean = () => {
      backdrop.classList.add("hidden");
      inputWrapper.classList.add("hidden");
      inputWrapper.innerHTML = "";
      okBtn.removeEventListener("click", ok);
      cancelBtn.removeEventListener("click", cancel);
    };
    const ok = () => {
      const res = {
        name: nameInput.value.trim(),
        startDate: startInput.value || "",
        endDate: endInput.value || "",
        description: descInput.value.trim() || ""
      };
      clean();
      resolve(res);
    };
    const cancel = () => {
      clean();
      resolve(null);
    };
    okBtn.addEventListener("click", ok);
    cancelBtn.addEventListener("click", cancel);
    setTimeout(()=>{ nameInput.focus(); nameInput.select(); }, 0);
  });
}

function renderStages() {
  const audit = getCurrentAudit();
  const list = $("#stages-list");
  if (!list) return;

  list.innerHTML = "";

  if (!audit || audit.stages.length === 0) return;

  audit.stages.forEach((s) => {
    const li = document.createElement("li");
    li.className = "list-item";

    const main = document.createElement("div");
    main.className = "list-main";

    const titleRow = document.createElement("div");
    titleRow.style.display = "flex";
    titleRow.style.justifyContent = "space-between";
    titleRow.style.alignItems = "center";

    const title = document.createElement("span");
    title.className = "list-title";
    title.textContent = s.name || "(sin nombre)";

    const rightActions = document.createElement("div");

    const editStageBtn = document.createElement("button");
    editStageBtn.className = "btn-chip";
    editStageBtn.textContent = "Editar etapa";
    editStageBtn.addEventListener("click", async () => {
      const res = await openStageEditor(s);
      if (!res) return;
      s.name = res.name || s.name;
      s.risk = res.risk || s.risk;
      s.startDate = res.startDate || "";
      s.endDate = res.endDate || "";
      scheduleRender();
      renderCalendar();
    });

    rightActions.appendChild(editStageBtn);
    // place delete button under the edit button with a small separation
    rightActions.style.display = "flex";
    rightActions.style.flexDirection = "column";
    rightActions.style.gap = "8px";

    const delStageBtn = document.createElement("button");
    delStageBtn.className = "btn-danger";
    delStageBtn.textContent = "Eliminar etapa";
    delStageBtn.title = "Eliminar etapa";
    delStageBtn.addEventListener("click", () => {
      const idx = audit.stages.findIndex((x) => x.id === s.id);
      if (idx !== -1) audit.stages.splice(idx, 1);
      scheduleRender();
      renderCalendar();
    });

    rightActions.appendChild(delStageBtn);
    titleRow.appendChild(title);
    titleRow.appendChild(rightActions);

    const meta = document.createElement("div");
    meta.className = "list-meta-row";

    const badgeRisk = document.createElement("span");
    badgeRisk.className = `badge badge-risk-${s.risk}`;
    badgeRisk.textContent = `Riesgo ${s.risk}`;

    const tasksInfo = document.createElement("span");
    tasksInfo.className = "list-sub";
    const activitiesCount = (s.activities && s.activities.length) || 0;
    tasksInfo.textContent = `${activitiesCount} actividades`;

    meta.appendChild(badgeRisk);
    meta.appendChild(tasksInfo);

    const attachmentsInfo = document.createElement("span");
    attachmentsInfo.className = "list-sub";
    const count = (s.attachments && s.attachments.length) || 0;
    if (count > 0) {
      const names = s.attachments.map(a => a.name).join(", ");
      attachmentsInfo.textContent = `Adjuntos (${count}): ${names}`;
    } else {
      attachmentsInfo.textContent = "Sin adjuntos";
    }

    const scheduleInfo = document.createElement("div");
    scheduleInfo.className = "list-sub";
    scheduleInfo.style.display = "flex";
    scheduleInfo.style.gap = "8px";
    scheduleInfo.style.alignItems = "center";

    const startInput = document.createElement("input");
    startInput.type = "date";
    startInput.className = "field-input";
    startInput.style.maxWidth = "140px";
    startInput.value = s.startDate || "";
    startInput.addEventListener("change", () => {
      s.startDate = startInput.value;
      scheduleRender();
      renderGantt();
      updateSummary();
      // keep calendar view in sync
      renderCalendar();
    });

    const endInput = document.createElement("input");
    endInput.type = "date";
    endInput.className = "field-input";
    endInput.style.maxWidth = "140px";
    endInput.value = s.endDate || "";
    endInput.addEventListener("change", () => {
      s.endDate = endInput.value;
      scheduleRender();
      renderGantt();
      updateSummary();
      // keep calendar view in sync
      renderCalendar();
    });

    scheduleInfo.appendChild(startInput);
    scheduleInfo.appendChild(endInput);

    // Listado de actividades con editar/eliminar por actividad
    const activitiesWrapper = document.createElement("div");
    activitiesWrapper.style.marginTop = "8px";
    if (activitiesCount > 0) {
      const actList = document.createElement("ul");
      actList.className = "list sub-list";
      s.activities.forEach((a) => {
        const actItem = document.createElement("li");
        actItem.className = "list-item list-item-compact";

        const actMain = document.createElement("div");
        actMain.className = "list-main";

        const actTitleRow = document.createElement("div");
        actTitleRow.style.display = "flex";
        actTitleRow.style.justifyContent = "space-between";
        actTitleRow.style.alignItems = "center";

        const actTitle = document.createElement("span");
        actTitle.className = "list-sub";
        actTitle.textContent = a.name || "(sin nombre)";

        const actBtns = document.createElement("div");

        const editActBtn = document.createElement("button");
        editActBtn.className = "icon-btn";
        editActBtn.textContent = "✎"; // icono lápiz
        editActBtn.title = "Editar actividad";
        editActBtn.addEventListener("click", async () => {
          const res = await openActivityEditor(a);
          if (!res) return;
          a.name = res.name || a.name;
          a.startDate = res.startDate || "";
          a.endDate = res.endDate || "";
          a.description = res.description || "";
          scheduleRender();
          renderGantt();
          updateSummary();
          // keep calendar view in sync
          renderCalendar();
        });

        const delActBtn = document.createElement("button");
        delActBtn.className = "icon-btn";
        delActBtn.textContent = "🗑"; // convertir a icono de basurero
        delActBtn.title = "Eliminar actividad";
        delActBtn.addEventListener("click", () => {
          const idx = s.activities.findIndex(x => x.id === a.id);
          if (idx !== -1) s.activities.splice(idx, 1);
          scheduleRender();
          renderGantt();
          updateSummary();
          // keep calendar view in sync
          renderCalendar();
        });

        actBtns.appendChild(editActBtn);
        actBtns.appendChild(delActBtn);

        actTitleRow.appendChild(actTitle);
        actTitleRow.appendChild(actBtns);

        const actMeta = document.createElement("div");
        actMeta.className = "list-meta-row";

        const actDates = document.createElement("span");
        actDates.className = "list-sub";
        // Mostrar fechas de actividad en formato DD/MM/YYYY (usar helper formatDateDisplay)
        const startDisplay = a.startDate ? formatDateDisplay(a.startDate) : "?";
        const endDisplay = a.endDate ? formatDateDisplay(a.endDate) : "?";
        actDates.textContent = a.startDate || a.endDate ? `${startDisplay} → ${endDisplay}` : "?";

        actMeta.appendChild(actDates);

        actMain.appendChild(actTitleRow);
        actMain.appendChild(actMeta);

        actItem.appendChild(actMain);
        actList.appendChild(actItem);
      });
      activitiesWrapper.appendChild(actList);
    } else {
      const noAct = document.createElement("div");
      noAct.className = "list-sub";
      noAct.textContent = "Sin actividades";
      activitiesWrapper.appendChild(noAct);
    }

    // Botón para agregar actividad
    const addActivityRow = document.createElement("div");
    addActivityRow.className = "list-meta-row";
    const addActivityBtn = document.createElement("button");
    addActivityBtn.className = "btn-chip";
    addActivityBtn.textContent = "Agregar actividad";
    addActivityBtn.addEventListener("click", async () => {
      const name = await openTextPrompt({
        title: "Nueva actividad",
        message: `Agrega una actividad para la etapa "${s.name}"`,
        placeholder: "Nombre de la actividad",
      });
      if (!name) return;
      if (!s.activities) s.activities = [];
      s.activities.push({
        id: uuid(),
        name: name.trim(),
        startDate: "",
        endDate: "",
      });
      scheduleRender();
      renderGantt();
      updateSummary();
      // keep calendar view in sync
      renderCalendar();
    });
    addActivityRow.appendChild(addActivityBtn);

    main.appendChild(titleRow);
    main.appendChild(meta);
    main.appendChild(attachmentsInfo);
    main.appendChild(scheduleInfo);
    main.appendChild(activitiesWrapper);
    main.appendChild(addActivityRow);

    li.appendChild(main);
    // note: delete button was moved into rightActions above
    list.appendChild(li);
  });
}

// Calendario
function setupCalendar() {
  // Solo renderizado reactivo por ahora
  renderCalendar();
}

function renderCalendar() {
  const audit = getCurrentAudit();
  const list = $("#calendar-list");
  if (!list) return;

  list.innerHTML = "";

  if (!audit || audit.stages.length === 0) {
    renderGantt();
    return;
  }

  audit.stages.forEach((s) => {
    const li = document.createElement("li");
    li.className = "calendar-column";

    const title = document.createElement("div");
    title.className = "calendar-column-title";
    title.textContent = s.name;

    const datesWrapper = document.createElement("div");
    datesWrapper.className = "calendar-column-dates";

    const startLabel = document.createElement("div");
    startLabel.className = "calendar-date-label";
    startLabel.textContent = "Inicio";

    const startInput = document.createElement("input");
    startInput.type = "date";
    startInput.className = "field-input";
    startInput.value = s.startDate || "";
    startInput.addEventListener("change", () => {
      s.startDate = startInput.value;
      scheduleRender();
      renderGantt();
      updateSummary();
    });

    const endLabel = document.createElement("div");
    endLabel.className = "calendar-date-label";
    endLabel.textContent = "Fin";

    const endInput = document.createElement("input");
    endInput.type = "date";
    endInput.className = "field-input";
    endInput.value = s.endDate || "";
    endInput.addEventListener("change", () => {
      s.endDate = endInput.value;
      scheduleRender();
      renderGantt();
      updateSummary();
    });

    datesWrapper.appendChild(startLabel);
    datesWrapper.appendChild(startInput);
    datesWrapper.appendChild(endLabel);
    datesWrapper.appendChild(endInput);

    li.appendChild(title);
    li.appendChild(datesWrapper);

    // Actividades por etapa dentro de la columna del calendario
    if (s.activities && s.activities.length > 0) {
      const actList = document.createElement("ul");
      actList.className = "list sub-list";

      s.activities.forEach((a) => {
        const actItem = document.createElement("li");
        actItem.className = "list-item list-item-compact";

        const actMain = document.createElement("div");
        actMain.className = "list-main";

        const actTitle = document.createElement("span");
        actTitle.className = "list-sub";
        actTitle.textContent = a.name;

        const actMetaRow = document.createElement("div");
        actMetaRow.className = "list-meta-row";

        const actStart = document.createElement("input");
        actStart.type = "date";
        actStart.className = "field-input";
        actStart.style.maxWidth = "120px";
        actStart.value = a.startDate || "";
        actStart.addEventListener("change", () => {
          a.startDate = actStart.value;
          scheduleRender();
          renderGantt();
          updateSummary();
        });

        const actEnd = document.createElement("input");
        actEnd.type = "date";
        actEnd.style.maxWidth = "120px";
        actEnd.className = "field-input";
        actEnd.value = a.endDate || "";
        actEnd.addEventListener("change", () => {
          a.endDate = actEnd.value;
          scheduleRender();
          renderGantt();
          updateSummary();
        });

        actMetaRow.appendChild(actStart);
        actMetaRow.appendChild(actEnd);

        actMain.appendChild(actTitle);
        actMain.appendChild(actMetaRow);

        actItem.appendChild(actMain);
        actList.appendChild(actItem);
      });

      li.appendChild(actList);
    }

    list.appendChild(li);
  });

  renderGantt();
}

function renderGantt() {
  const container = $("#gantt-container");
  if (!container) return;

  container.innerHTML = "";

  const audit = getCurrentAudit();
  if (!audit || !audit.stages || audit.stages.length === 0) {
    const empty = document.createElement("div");
    empty.className = "gantt-empty";
    empty.textContent = "Sin programación aún";
    container.appendChild(empty);
    return;
  }

  // Recolectar todas las fechas
  const ranges = [];
  audit.stages.forEach((s) => {
    if (s.startDate && s.endDate) {
      ranges.push({
        type: "stage",
        id: s.id,
        name: s.name,
        start: new Date(s.startDate),
        end: new Date(s.endDate),
      });
    }
    (s.activities || []).forEach((a) => {
      if (a.startDate && a.endDate) {
        ranges.push({
          type: "activity",
          id: a.id,
          stageId: s.id,
          stageName: s.name,
          name: a.name,
          start: new Date(a.startDate),
          end: new Date(a.endDate),
          description: a.description || ""
        });
      }
    });
  });

  if (ranges.length === 0) {
    const empty = document.createElement("div");
    empty.className = "gantt-empty";
    empty.textContent = "Agrega fechas a las etapas o actividades para ver el diagrama de Gantt";
    container.appendChild(empty);
    return;
  }

  let minStart = ranges[0].start;
  let maxEnd = ranges[0].end;
  ranges.forEach((r) => {
    if (r.start < minStart) minStart = r.start;
    if (r.end > maxEnd) maxEnd = r.end;
  });

  const msPerDay = 1000 * 60 * 60 * 24;
  const totalDays = Math.max(1, Math.round((maxEnd - minStart) / msPerDay) + 1);

  const formatDateShort = (d) => {
    const day = String(d.getDate()).padStart(2, "0");
    const month = String(d.getMonth() + 1).padStart(2, "0");
    return `${day}/${month}`;
  };

  // Encabezado y leyenda del Gantt
  const header = document.createElement("div");
  header.className = "gantt-header";

  const rangeLabel = document.createElement("div");
  rangeLabel.className = "gantt-header-range";
  rangeLabel.textContent = `Del ${formatDateShort(minStart)} al ${formatDateShort(maxEnd)}`;

  const spanLabel = document.createElement("div");
  spanLabel.className = "gantt-header-span";
  spanLabel.textContent = `${totalDays} día${totalDays === 1 ? "" : "s"}`;

  header.appendChild(rangeLabel);
  header.appendChild(spanLabel);
  container.appendChild(header);

  const legend = document.createElement("div");
  legend.className = "gantt-legend";

  const legendStage = document.createElement("div");
  legendStage.className = "gantt-legend-item";
  const legendStageSwatch = document.createElement("span");
  legendStageSwatch.className = "gantt-legend-swatch gantt-legend-swatch-stage";
  const legendStageText = document.createElement("span");
  legendStageText.textContent = "Etapa";
  legendStage.appendChild(legendStageSwatch);
  legendStage.appendChild(legendStageText);

  const legendActivity = document.createElement("div");
  legendActivity.className = "gantt-legend-item";
  const legendActivitySwatch = document.createElement("span");
  legendActivitySwatch.className = "gantt-legend-swatch gantt-legend-swatch-activity";
  const legendActivityText = document.createElement("span");
  legendActivityText.textContent = "Actividad";
  legendActivity.appendChild(legendActivitySwatch);
  legendActivity.appendChild(legendActivityText);

  legend.appendChild(legendStage);
  legend.appendChild(legendActivity);
  container.appendChild(legend);

  // Agrupar por etapa para mostrar filas
  const stagesMap = new Map();
  audit.stages.forEach((s) => {
    stagesMap.set(s.id, {
      stage: s,
      activities: (s.activities || []),
    });
  });

  stagesMap.forEach(({ stage, activities }) => {
    const row = document.createElement("div");
    row.className = "gantt-row";

    const label = document.createElement("div");
    label.className = "gantt-label";
    label.textContent = stage.name;

    const barArea = document.createElement("div");
    barArea.className = "gantt-bar-area";

    // Barra de etapa
    if (stage.startDate && stage.endDate) {
      const sStart = new Date(stage.startDate);
      const sEnd = new Date(stage.endDate);
      const startOffset = Math.max(0, Math.round((sStart - minStart) / msPerDay));
      const duration = Math.max(1, Math.round((sEnd - sStart) / msPerDay) + 1);

      const bar = document.createElement("div");
      bar.className = "gantt-bar gantt-bar-stage";
      bar.style.left = `${(startOffset / totalDays) * 100}%`;
      bar.style.width = `${(duration / totalDays) * 100}%`;
      bar.title = `${stage.name} — ${stage.startDate || ""} → ${stage.endDate || ""}`;
      barArea.appendChild(bar);
    }

    // Barras de actividades
    activities.forEach((a, index) => {
      if (!a.startDate || !a.endDate) return;
      const aStart = new Date(a.startDate);
      const aEnd = new Date(a.endDate);
      const startOffset = Math.max(0, Math.round((aStart - minStart) / msPerDay));
      const duration = Math.max(1, Math.round((aEnd - aStart) / msPerDay) + 1);

      const bar = document.createElement("div");
      bar.className = "gantt-bar gantt-bar-activity";
      // Separar ligeramente verticalmente por actividad
      bar.style.top = `${4 + index * 6}px`;
      bar.style.left = `${(startOffset / totalDays) * 100}%`;
      bar.style.width = `${(duration / totalDays) * 100}%`;

      // Tooltip con descripción (si existe) y datos básicos
      const desc = a.description ? `Descripción: ${a.description}\n` : "";
      bar.title = `${a.name} — ${a.startDate} → ${a.endDate}\n${desc}`.trim();

      barArea.appendChild(bar);
    });

    row.appendChild(label);
    row.appendChild(barArea);
    container.appendChild(row);
  });
}

// Nuevo: función para generar reporte descargable en PDF (usa jsPDF + html2canvas)
async function generateGanttReport(filename) {
  const audit = getCurrentAudit();
  const doc = new jsPDF({
    unit: "pt",
    format: "a4",
    compress: true
  });

  // Construir un contenedor temporal con el contenido a renderizar
  const tmp = document.createElement("div");
  tmp.style.width = "840px"; // ancho aproximado para mejor resolución en A4
  tmp.style.padding = "16px";
  tmp.style.fontFamily = 'system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif';
  tmp.style.color = "#222";

  const titleForDoc = filename && filename.trim() ? filename.trim() : (audit ? (audit.name || audit.code) : "Informe");

  // Secciones: Datos generales
  const genHtml = `
    <h1 style="color:#f57c00;margin-bottom:8px;">${titleForDoc}</h1>
    <div style="font-size:12px;margin-bottom:8px;">
      <strong>Auditoría:</strong> ${audit ? (audit.code + (audit.name ? ` — ${audit.name}` : "")) : "N/A"}<br/>
      <strong>Auditor líder:</strong> ${audit && audit.leadAuditorId ? (audit.participants.find(p=>p.id===audit.leadAuditorId)?.name || "") : "No asignado"}
    </div>
  `;
  tmp.innerHTML = genHtml;

  // Participantes en 3 columnas (HTML sencillo)
  const participantsHTML = (audit && audit.participants && audit.participants.length)
    ? `<div style="margin-bottom:12px;"><h3 style="margin:6px 0 8px 0;">Listado de participantes</h3><div style="column-count:3;column-gap:16px;">${audit.participants.map(p => `<div style="break-inside:avoid;margin-bottom:6px;"><strong>${p.name}</strong><div style="font-size:11px;color:#555;">${p.email || "sin correo"} — ${roleLabel(p.role)}</div></div>`).join("")}</div></div>`
    : `<div style="margin-bottom:12px;"><h3 style="margin:6px 0 8px 0;">Listado de participantes</h3><div>No hay participantes</div></div>`;

  tmp.insertAdjacentHTML("beforeend", participantsHTML);

  // Etapas y actividades con calendario
  const stagesHTML = (audit && audit.stages && audit.stages.length)
    ? `<div style="margin-bottom:12px;"><h3 style="margin:6px 0 8px 0;">Definición de las etapas</h3>${audit.stages.map(s => {
        const acts = (s.activities || []).map(a => {
          const start = formatDateDisplay(a.startDate);
          const end = formatDateDisplay(a.endDate);
          const dates = a.startDate ? ` — ${start} → ${end}` : "";
          return `<li style="font-size:12px;margin-bottom:4px;">${a.name}${dates}${a.description ? ` — ${a.description}` : ""}</li>`;
        }).join("");
        const atts = (s.attachments || []).map(at => at.name).join(", ");
        const sStart = formatDateDisplay(s.startDate);
        const sEnd = formatDateDisplay(s.endDate);
        return `<div style="margin-bottom:8px;"><strong style="font-size:13px;">${s.name}</strong> <div style="font-size:12px;color:#555;">Riesgo: ${s.risk} · Fechas: ${sStart} → ${sEnd} · Adjuntos: ${atts || "Ninguno"}</div><ul style="margin-top:6px;padding-left:16px;">${acts || "<li style='font-size:12px;'>Sin actividades</li>"}</ul></div>`;
      }).join("")}</div>`
    : `<div style="margin-bottom:12px;"><h3 style="margin:6px 0 8px 0;">Definición de las etapas</h3><div>Sin etapas</div></div>`;

  tmp.insertAdjacentHTML("beforeend", stagesHTML);

  // Diagrama: reutilizar el gantt actual pero simplificar estilo para captura
  const gantt = $("#gantt-container");
  const ganttClone = document.createElement("div");
  ganttClone.style.border = "1px solid #ddd";
  ganttClone.style.padding = "8px";
  ganttClone.style.background = "#fafafa";
  ganttClone.innerHTML = gantt ? gantt.innerHTML : "<div>Sin diagrama</div>";

  tmp.appendChild(ganttClone);

  // Renderizar a canvas con html2canvas y añadir a PDF
  // html2canvas necesita que el elemento esté en el DOM
  tmp.style.position = "fixed";
  tmp.style.left = "-20000px";
  tmp.style.top = "-20000px";
  tmp.style.visibility = "hidden";
  
  // Crear un contenedor limpio sin referencias de performance
  const tmpClone = document.createElement("div");
  tmpClone.style.width = "840px";
  tmpClone.style.padding = "16px";
  tmpClone.style.fontFamily = 'system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif';
  tmpClone.style.color = "#222";
  tmpClone.style.position = "fixed";
  tmpClone.style.left = "-20000px";
  tmpClone.style.top = "-20000px";
  tmpClone.style.visibility = "hidden";
  tmpClone.innerHTML = tmp.innerHTML;
  
  document.body.appendChild(tmpClone);

  try {
    // Try with minimal, conservative options to avoid PerformanceServerTiming serialization
    const canvas = await html2canvas(tmpClone, {
      scale: 1,
      useCORS: false,
      allowTaint: true,
      backgroundColor: "#ffffff",
      logging: false,
      imageTimeout: 0,
      removeContainer: false
    });

    const imgData = canvas.toDataURL("image/jpeg", 0.85);
    const pdfWidth = doc.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height / canvas.width) * (pdfWidth - 40);

    doc.addImage(imgData, 'JPEG', 20, 20, pdfWidth - 40, pdfHeight);
  } catch (err) {
    // fallback: generar PDF con contenido textual si html2canvas falla
    console.warn("html2canvas error during PDF generation:", err);
    doc.setFontSize(12);
    doc.text("Informe de Auditoría", 20, 30);
    doc.setFontSize(10);
    if (audit) {
      doc.text(`Auditoría: ${audit.code || "N/A"}`, 20, 50);
      if (audit.name) doc.text(`Nombre: ${audit.name}`, 20, 65);
      doc.text(`Participantes: ${audit.participants ? audit.participants.length : 0}`, 20, 80);
      doc.text(`Etapas: ${audit.stages ? audit.stages.length : 0}`, 20, 95);
    }
  } finally {
    // Limpieza del nodo temporal añadido
    if (tmpClone && tmpClone.parentNode) {
      tmpClone.remove();
    }
    if (tmp && tmp.parentNode) {
      tmp.remove();
    }
  }

  // Guardar archivo
  const outName = (filename && filename.trim()) ? `${filename.trim()}.pdf` : `${(audit ? (audit.code || "informe") : "informe")}.pdf`;
  try {
    doc.save(outName);
  } catch (e) {
    console.error("Error saving PDF:", e);
  }
}

// Bind PDF button
window.addEventListener("DOMContentLoaded", () => {
  const pdfBtn = $("#gantt-pdf-btn");
  if (pdfBtn) {
    pdfBtn.addEventListener("click", () => {
      openTextPrompt({
        title: "Descargar Informe en PDF",
        message: "El informe se generará y descargará en formato PDF. Ingresa un nombre para el archivo (opcional).",
        placeholder: "Nombre del archivo (sin extensión)",
        initialValue: ""
      }).then((val) => {
        if (val === null) return;
        generateGanttReport(val || undefined);
      });
    });
  }
});

// Resumen
function updateSummary() {
  const audit = getCurrentAudit();
  const participantsCount = audit ? audit.participants.length : 0;
  const stagesCount = audit ? audit.stages.length : 0;
  const totalTasks = audit
    ? audit.stages.reduce((acc, s) => acc + ((s.activities && s.activities.length) || 0), 0)
    : 0;

  $("#summary-participants").textContent = participantsCount;
  $("#summary-stages").textContent = stagesCount;
  $("#summary-tasks").textContent = totalTasks;

  // Calculate project days: difference between earliest start and latest end among stages and activities
  let projectDays = 0;
  if (audit && audit.stages && audit.stages.length) {
    const msPerDay = 1000 * 60 * 60 * 24;
    let minStart = null;
    let maxEnd = null;
    audit.stages.forEach((s) => {
      if (s.startDate) {
        const d = new Date(s.startDate);
        if (!minStart || d < minStart) minStart = d;
      }
      if (s.endDate) {
        const d = new Date(s.endDate);
        if (!maxEnd || d > maxEnd) maxEnd = d;
      }
      (s.activities || []).forEach((a) => {
        if (a.startDate) {
          const d = new Date(a.startDate);
          if (!minStart || d < minStart) minStart = d;
        }
        if (a.endDate) {
          const d = new Date(a.endDate);
          if (!maxEnd || d > maxEnd) maxEnd = d;
        }
      });
    });
    if (minStart && maxEnd && maxEnd >= minStart) {
      projectDays = Math.round((maxEnd - minStart) / msPerDay) + 1;
    }
  }
  $("#summary-days").textContent = projectDays;
}

// Cargar datos de ejemplo (añadimos descripciones a actividades para demostrar tooltip)
function loadSampleData() {
  const sampleCode = "AUD-2025-01";
  const sampleAudit = createEmptyAudit(sampleCode);
  sampleAudit.name = "Auditoría integral 2025";
  sampleAudit.participants = [
    { id: uuid(), name: "Ana Torres", email: "ana.torres@empresa.com", role: "auditor_lider" },
    { id: uuid(), name: "Luis Pérez", email: "luis.perez@empresa.com", role: "auditor" },
    { id: uuid(), name: "María Gómez", email: "maria.gomez@empresa.com", role: "responsable_proceso" },
    { id: uuid(), name: "Carlos Rivas", email: "carlos.rivas@empresa.com", role: "responsable_proceso" },
    { id: uuid(), name: "Elena Núñez", email: "elena.nunez@empresa.com", role: "observador" },
  ];
  sampleAudit.leadAuditorId = sampleAudit.participants[0].id;
  sampleAudit.stages = [
    {
      id: uuid(),
      name: "Planeación de la auditoría",
      risk: "bajo",
      activities: [
        {
          id: uuid(),
          name: "Definir alcance",
          startDate: "2025-01-10",
          endDate: "2025-01-11",
          description: "Reunión con dirección para delimitar objetivos y alcance."
        },
        {
          id: uuid(),
          name: "Identificar riesgos clave",
          startDate: "2025-01-12",
          endDate: "2025-01-15",
          description: "Análisis preliminar de procesos y controles."
        },
      ],
      attachments: [
        { name: "alcance_auditoria.pdf", type: "application/pdf" },
      ],
      startDate: "2025-01-10",
      endDate: "2025-01-15",
    },
    {
      id: uuid(),
      name: "Revisión de documentación financiera",
      risk: "alto",
      activities: [
        {
          id: uuid(),
          name: "Revisión estados financieros",
          startDate: "2025-01-16",
          endDate: "2025-01-20",
          description: "Comparar balances con libros auxiliares y registros."
        },
        {
          id: uuid(),
          name: "Validación soportes contables",
          startDate: "2025-01-21",
          endDate: "2025-01-25",
          description: "Verificación de facturas y comprobantes asociados."
        },
      ],
      attachments: [
        { name: "lista_documentos.xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      ],
      startDate: "2025-01-16",
      endDate: "2025-01-25",
    },
    {
      id: uuid(),
      name: "Pruebas de control interno",
      risk: "medio",
      activities: [
        {
          id: uuid(),
          name: "Levantamiento de controles",
          startDate: "2025-01-26",
          endDate: "2025-01-28",
          description: "Documentación de controles existentes y responsables."
        },
        {
          id: uuid(),
          name: "Ejecución de pruebas",
          startDate: "2025-01-29",
          endDate: "2025-02-02",
          description: "Aplicación de pruebas de cumplimiento y eficacia."
        },
      ],
      attachments: [
        { name: "matriz_controles.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      ],
      startDate: "2025-01-26",
      endDate: "2025-02-02",
    },
    {
      id: uuid(),
      name: "Informe y cierre de auditoría",
      risk: "medio",
      activities: [
        {
          id: uuid(),
          name: "Redacción del informe",
          startDate: "2025-02-03",
          endDate: "2025-02-07",
          description: "Consolidar hallazgos y recomendaciones."
        },
        {
          id: uuid(),
          name: "Reunión de cierre",
          startDate: "2025-02-10",
          endDate: "2025-02-10",
          description: "Presentación de resultados a la dirección."
        },
      ],
      attachments: [
        { name: "borrador_informe.pdf", type: "application/pdf" },
      ],
      startDate: "2025-02-03",
      endDate: "2025-02-10",
    },
  ];

  state.audits = {
    [sampleCode]: sampleAudit,
  };
  state.currentAuditCode = sampleCode;

  updateAuditSelector();
  loadAuditIntoUI();
  scheduleRender();
}

// Render global
function renderAll() {
  renderParticipants();
  renderStages();
  updateLeadAuditorOptions();
  updateSummary();
  renderCalendar();
}

// Setup help button in tab bar
window.addEventListener("DOMContentLoaded", () => {
  const helpBtn = $("#help-button");
  const helpBackdrop = $("#help-modal-backdrop");
  const helpCloseBtn = $("#help-close-btn");

  if (helpBtn && helpBackdrop && helpCloseBtn) {
    helpBtn.addEventListener("click", (e) => {
      e.preventDefault();
      helpBackdrop.classList.remove("hidden");
    });

    helpCloseBtn.addEventListener("click", () => {
      helpBackdrop.classList.add("hidden");
    });

    helpBackdrop.addEventListener("click", (e) => {
      if (e.target === helpBackdrop) {
        helpBackdrop.classList.add("hidden");
      }
    });

    // Close on Escape key
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !helpBackdrop.classList.contains("hidden")) {
        helpBackdrop.classList.add("hidden");
      }
    });
  }
});

// Setup help TOC toggle
window.addEventListener("DOMContentLoaded", () => {
  const tocToggle = document.querySelector(".help-toc-toggle");
  const tocList = document.querySelector(".help-toc-list");
  const tocLinks = document.querySelectorAll(".help-toc-link");

  if (tocToggle && tocList) {
    tocToggle.addEventListener("click", () => {
      tocList.classList.toggle("hidden");
    });

    tocLinks.forEach(link => {
      link.addEventListener("click", (e) => {
        e.preventDefault();
        const target = link.getAttribute("href");
        const el = document.querySelector(target);
        if (el) {
          el.scrollIntoView({ behavior: "smooth" });
          // Close TOC after selection
          tocList.classList.add("hidden");
        }
      });
    });
  }
});

// Setup help sections to show only selected topic
window.addEventListener("DOMContentLoaded", () => {
  const tocLinks = document.querySelectorAll(".help-toc-link");
  const helpSections = document.querySelectorAll(".help-section");

  // Set first topic as active by default
  if (helpSections.length > 0) {
    helpSections[0].classList.add("active");
    if (tocLinks.length > 0) {
      tocLinks[0].classList.add("active");
    }
  }

  tocLinks.forEach(link => {
    link.addEventListener("click", (e) => {
      e.preventDefault();
      const topicId = link.getAttribute("data-topic");
      if (!topicId) return;

      // Remove active class from all sections and links
      helpSections.forEach(section => section.classList.remove("active"));
      tocLinks.forEach(l => l.classList.remove("active"));

      // Add active class to selected section and link
      const targetSection = document.getElementById(topicId);
      if (targetSection) {
        targetSection.classList.add("active");
        link.classList.add("active");
        
        // Reset scroll to top for new section
        const helpContent = document.querySelector(".help-content");
        if (helpContent) {
          helpContent.scrollTop = 0;
        }
      }
    });
  });
});