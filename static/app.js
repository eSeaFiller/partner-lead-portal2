const state = {
  schema: null,
  leads: [],
  currentView: "partner",
  adminKey: sessionStorage.getItem("leadPortalAdminKey") || "",
};

const fields = [
  "firstName",
  "lastName",
  "country",
  "companyName",
  "mobileNumber",
  "workEmail",
  "jobTitle",
  "companySize",
  "industry",
  "subIndustry",
  "trackingCode",
];

const form = document.querySelector("#leadForm");
const uploadForm = document.querySelector("#uploadForm");
const duplicateBox = document.querySelector("#duplicateBox");
const submitStatus = document.querySelector("#submitStatus");
const uploadStatus = document.querySelector("#uploadStatus");
const uploadResult = document.querySelector("#uploadResult");

function option(label, value = "") {
  const item = document.createElement("option");
  item.value = value;
  item.textContent = label;
  return item;
}

function fillOptions(select, items, placeholder) {
  select.replaceChildren(option(placeholder));
  for (const item of items) {
    if (typeof item === "string") {
      select.append(option(item, item));
    } else {
      select.append(option(`${item.value} - ${item.label}`, item.value));
    }
  }
}

function activeIndustry() {
  const value = form.elements.industry.value;
  return state.schema.industries.find((item) => item.value === value);
}

function refreshSubIndustries() {
  const industry = activeIndustry();
  fillOptions(
    form.elements.subIndustry,
    industry ? industry.subIndustries : [],
    industry ? "Select sub industry" : "Select industry first",
  );
}

function setView(view) {
  state.currentView = view;
  document.querySelector("#partnerView").classList.toggle("active", view === "partner");
  document.querySelector("#adminView").classList.toggle("active", view === "admin");
  document.querySelector("#partnerTab").classList.toggle("active", view === "partner");
  document.querySelector("#adminTab").classList.toggle("active", view === "admin");
  document.querySelector("#pageTitle").textContent = view === "admin" ? "Admin review" : "Lead submission";
  if (view === "admin") unlockAdmin(false);
}

function valuesFromForm() {
  const data = {};
  for (const field of fields) data[field] = form.elements[field]?.value.trim() || "";
  return data;
}

function clearErrors() {
  for (const item of document.querySelectorAll("[data-error-for]")) item.textContent = "";
  duplicateBox.classList.add("hidden");
  duplicateBox.textContent = "";
}

function showErrors(errors = {}, warnings = {}) {
  clearErrors();
  for (const [field, message] of Object.entries({ ...warnings, ...errors })) {
    const target = document.querySelector(`[data-error-for="${field}"]`);
    if (target) target.textContent = message;
  }
}

function duplicateText(duplicates) {
  const lines = [];
  if (Array.isArray(duplicates.email) ? duplicates.email.length : duplicates.email) {
    lines.push("A lead with this work email already exists");
  }
  if (Array.isArray(duplicates.company) ? duplicates.company.length : duplicates.company) {
    lines.push("A lead from this company and country already exists");
  }
  return lines.join(" · ");
}

async function checkDuplicates() {
  const values = valuesFromForm();
  const params = new URLSearchParams({
    workEmail: values.workEmail,
    companyName: values.companyName,
    country: values.country,
  });
  if (!values.workEmail && !values.companyName) return;
  const response = await fetch(`/api/check?${params}`);
  const payload = await response.json();
  const text = duplicateText(payload.duplicates);
  duplicateBox.textContent = text;
  duplicateBox.classList.toggle("hidden", !text);
}

async function submitLead(event) {
  event.preventDefault();
  clearErrors();
  submitStatus.textContent = "Submitting...";

  const response = await fetch("/api/leads", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      partner: document.querySelector("#partner").value.trim(),
      fields: valuesFromForm(),
    }),
  });
  const payload = await response.json();

  if (!response.ok) {
    showErrors(payload.errors, payload.warnings);
    const text = duplicateText(payload.duplicates || {});
    if (text) {
      duplicateBox.textContent = text;
      duplicateBox.classList.remove("hidden");
    }
    submitStatus.textContent = "Please fix the highlighted fields.";
    return;
  }

  submitStatus.textContent = "Submitted. Status: pending review.";
  form.reset();
  setPartnerFromUrl();
  refreshSubIndustries();
}

async function uploadLeads(event) {
  event.preventDefault();
  uploadResult.classList.add("hidden");
  uploadResult.textContent = "";
  const file = document.querySelector("#leadFile").files[0];
  if (!file) {
    uploadStatus.textContent = "Choose an .xlsx file first.";
    return;
  }
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    uploadStatus.textContent = "Only .xlsx files are supported.";
    return;
  }

  uploadStatus.textContent = "Uploading and parsing...";
  const body = new FormData();
  body.append("partner", document.querySelector("#partner").value.trim());
  body.append("file", file);

  const response = await fetch("/api/upload", {
    method: "POST",
    body,
  });
  const payload = await response.json();

  if (!response.ok) {
    uploadStatus.textContent = "Upload failed.";
    uploadResult.innerHTML = escapeHtml(payload.error || "Unknown error");
    uploadResult.classList.remove("hidden");
    return;
  }

  uploadStatus.textContent = "Upload processed.";
  uploadResult.innerHTML = uploadSummary(payload);
  uploadResult.classList.remove("hidden");
  uploadForm.reset();
}

function uploadSummary(payload) {
  const lines = [
    `<strong>${payload.imported}</strong> leads submitted for review.`,
    `${payload.parsed} rows parsed · ${payload.failed} rows failed`,
  ];
  if (payload.errors?.length) {
    const items = payload.errors
      .map((item) => {
        const messages = Object.entries(item.errors)
          .map(([field, message]) => `${field}: ${message}`)
          .join("; ");
        return `<li>Row ${escapeHtml(item.row)} · ${escapeHtml(item.email || item.companyName || "No identifier")} · ${escapeHtml(messages)}</li>`;
      })
      .join("");
    lines.push(`<ul>${items}</ul>`);
  }
  return lines.join("<br>");
}

function warningSummary(lead) {
  const warnings = lead.warnings || {};
  const values = Object.values(warnings).filter(Boolean);
  return values.length ? values.join("; ") : "None";
}

function formatDate(value) {
  if (!value) return "";
  return new Date(value).toLocaleString();
}

async function updateLead(id, patch) {
  const response = await fetch(`/api/leads/${id}`, {
    method: "PATCH",
    headers: adminHeaders(),
    body: JSON.stringify(patch),
  });
  if (!response.ok) throw new Error("Failed to update lead");
  await loadLeads();
}

function renderLeads() {
  const tbody = document.querySelector("#leadRows");
  const filter = document.querySelector("#statusFilter").value;
  const leads = state.leads.filter((lead) => filter === "all" || lead.status === filter);
  tbody.replaceChildren();

  for (const lead of leads) {
    const tr = document.createElement("tr");
    const f = lead.fields;
    tr.innerHTML = `
      <td><span class="status ${lead.status}">${lead.status}</span></td>
      <td>${escapeHtml(lead.partner)}</td>
      <td><strong>${escapeHtml(f.firstName)} ${escapeHtml(f.lastName)}</strong><br>${escapeHtml(f.workEmail)}</td>
      <td>${escapeHtml(f.companyName)}</td>
      <td>${escapeHtml(f.country)}</td>
      <td>${escapeHtml(f.jobTitle)}</td>
      <td>${escapeHtml(warningSummary(lead))}</td>
      <td>${escapeHtml(formatDate(lead.createdAt))}</td>
      <td>
        <div class="review-buttons">
          <button type="button" data-action="approved" data-id="${lead.id}">Approve</button>
          <button type="button" data-action="rejected" data-id="${lead.id}">Reject</button>
          <button type="button" data-action="pending" data-id="${lead.id}">Pending</button>
        </div>
      </td>
    `;
    tbody.append(tr);
  }

  const counts = state.leads.reduce((acc, lead) => {
    acc[lead.status] = (acc[lead.status] || 0) + 1;
    return acc;
  }, {});
  document.querySelector("#leadSummary").textContent =
    `${state.leads.length} total · ${counts.pending || 0} pending · ${counts.approved || 0} approved · ${counts.rejected || 0} rejected`;
}

function escapeHtml(value) {
  return String(value || "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[char]);
}

async function loadLeads() {
  const response = await fetch("/api/leads", { headers: adminHeaders() });
  const payload = await response.json();
  if (!response.ok) {
    lockAdmin("Admin key is invalid.");
    return;
  }
  state.leads = payload.leads;
  renderLeads();
}

async function exportApproved() {
  const box = document.querySelector("#exportResult");
  box.textContent = "Generating export...";
  box.classList.remove("hidden");
  const response = await fetch("/api/export?status=approved", { headers: adminHeaders() });
  const payload = await response.json();
  if (!response.ok) {
    box.textContent = "Export failed.";
    return;
  }
  const fileName = payload.path.split("/").pop();
  box.innerHTML = `Exported ${payload.count} approved leads: <a href="/exports/${fileName}?admin_key=${encodeURIComponent(state.adminKey)}">${escapeHtml(fileName)}</a>`;
}

function setPartnerFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const partner = params.get("partner");
  if (partner) document.querySelector("#partner").value = partner;
}

function adminHeaders() {
  return {
    "Content-Type": "application/json",
    "X-Admin-Key": state.adminKey,
  };
}

function lockAdmin(message = "") {
  document.querySelector("#adminLogin").classList.remove("hidden");
  document.querySelector("#adminPanel").classList.add("hidden");
  document.querySelector("#adminLoginStatus").textContent = message;
}

async function unlockAdmin(showError = true) {
  if (!state.adminKey) {
    lockAdmin(showError ? "Admin key required." : "");
    return;
  }
  document.querySelector("#adminLogin").classList.add("hidden");
  document.querySelector("#adminPanel").classList.remove("hidden");
  await loadLeads();
}

function configureEntryMode() {
  const path = window.location.pathname;
  const isAdmin = path === "/admin";
  document.querySelector(".tabs").classList.add("hidden");
  document.querySelector("#adminTab").classList.toggle("hidden", !isAdmin);
  document.querySelector("#partnerTab").classList.toggle("hidden", isAdmin);
  setView(isAdmin ? "admin" : "partner");
}

async function init() {
  const schemaResponse = await fetch("/api/schema");
  state.schema = await schemaResponse.json();
  document.querySelector("#templatePath").textContent = "Required fields are checked before submission.";
  fillOptions(form.elements.country, state.schema.countries, "Select country");
  fillOptions(form.elements.jobTitle, state.schema.jobTitles, "Select job title");
  fillOptions(form.elements.companySize, state.schema.companySizes, "Select company size");
  fillOptions(form.elements.industry, state.schema.industries, "Select industry");
  refreshSubIndustries();
  setPartnerFromUrl();

  document.querySelector("#partnerTab").addEventListener("click", () => {
    window.history.pushState({}, "", "/partner");
    setView("partner");
  });
  document.querySelector("#adminTab").addEventListener("click", () => {
    window.history.pushState({}, "", "/admin");
    setView("admin");
  });
  document.querySelector("#adminLogin").addEventListener("submit", (event) => {
    event.preventDefault();
    state.adminKey = document.querySelector("#adminKey").value.trim();
    sessionStorage.setItem("leadPortalAdminKey", state.adminKey);
    unlockAdmin();
  });
  document.querySelector("#refreshLeads").addEventListener("click", loadLeads);
  document.querySelector("#statusFilter").addEventListener("change", renderLeads);
  document.querySelector("#exportApproved").addEventListener("click", exportApproved);
  document.querySelector("#leadRows").addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (button) updateLead(button.dataset.id, { status: button.dataset.action });
  });
  form.addEventListener("submit", submitLead);
  uploadForm.addEventListener("submit", uploadLeads);
  form.elements.industry.addEventListener("change", refreshSubIndustries);
  form.elements.workEmail.addEventListener("blur", checkDuplicates);
  form.elements.companyName.addEventListener("blur", checkDuplicates);
  form.elements.country.addEventListener("change", checkDuplicates);
  configureEntryMode();
}

init().catch((error) => {
  console.error(error);
  submitStatus.textContent = "Failed to load portal.";
});
