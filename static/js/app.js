const projectSwitcher = document.getElementById("projectSwitcher");
if (projectSwitcher) {
  projectSwitcher.addEventListener("change", () => {
    projectSwitcher.form?.requestSubmit();
  });
}

document.querySelectorAll(".js-confirm-delete").forEach((form) => {
  form.addEventListener("submit", (event) => {
    if (form.dataset.confirmed === "1") return;
    event.preventDefault();
    const modal = document.getElementById("deleteConfirmModal");
    if (!modal || !window.bootstrap?.Modal) {
      if (window.confirm("Esta accion no se puede deshacer.")) {
        form.dataset.confirmed = "1";
        form.submit();
      }
      return;
    }

    const actionLabel = modal.querySelector("[data-delete-action]");
    const confirmButton = modal.querySelector("[data-delete-confirm]");
    if (actionLabel) {
      actionLabel.textContent = form.dataset.deleteLabel || "este registro";
    }
    if (confirmButton) {
      confirmButton.onclick = () => {
        form.dataset.confirmed = "1";
        form.submit();
      };
    }

    window.bootstrap.Modal.getOrCreateInstance(modal).show();
  });
});

document.querySelectorAll("[data-open-modal]").forEach((button) => {
  button.addEventListener("click", () => {
    const modal = document.querySelector(button.dataset.openModal);
    if (!modal || !window.bootstrap?.Modal) return;
    const toDataKey = (key) => key.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
    const fields = modal.querySelectorAll("[data-field]");
    fields.forEach((field) => {
      const key = field.getAttribute("data-field");
      const value = button.dataset[toDataKey(key)] ?? "";
      if (field.tagName === "SELECT" || field.tagName === "INPUT" || field.tagName === "TEXTAREA") {
        field.value = value;
      } else {
        field.textContent = value;
      }
    });
    const form = modal.querySelector("form");
    if (form && button.dataset.action) {
      form.action = button.dataset.action;
    }
    const title = modal.querySelector("[data-modal-title]");
    if (title && button.dataset.title) {
      title.textContent = button.dataset.title;
    }
    const submit = modal.querySelector("[data-submit-label]");
    if (submit && button.dataset.submitLabel) {
      submit.textContent = button.dataset.submitLabel;
    }
    window.bootstrap.Modal.getOrCreateInstance(modal).show();
  });
});

document.querySelectorAll("[data-reset-modal]").forEach((button) => {
  button.addEventListener("click", () => {
    const modal = document.querySelector(button.dataset.resetModal);
    if (!modal) return;
    const form = modal.querySelector("form");
    if (form) form.reset();
    modal.querySelectorAll("[data-field]").forEach((field) => {
      if (field.tagName === "INPUT" || field.tagName === "TEXTAREA") field.value = field.dataset.defaultValue || "";
      if (field.tagName === "SELECT") {
        if (field.dataset.defaultValue) {
          field.value = field.dataset.defaultValue;
        } else {
          field.selectedIndex = 0;
        }
      }
    });
  });
});
