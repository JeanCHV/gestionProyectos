document.querySelectorAll(".side-link").forEach((link) => {
  link.addEventListener("click", () => {
    document.querySelectorAll(".side-link").forEach((item) => item.classList.remove("active"));
    link.classList.add("active");
  });
});

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

document.addEventListener("DOMContentLoaded", () => {
  if (!window.jQuery || !jQuery.fn?.DataTable) return;

  document.querySelectorAll("table.datatable").forEach((table) => {
    if (jQuery.fn.DataTable.isDataTable(table)) return;

    const headerCount = table.querySelectorAll("thead th").length;
    const bodyRows = Array.from(table.tBodies[0]?.rows || []);
    const hasRealRows = bodyRows.some((row) => {
      if (row.cells.length !== 1) return true;
      return row.cells[0].colSpan < headerCount;
    });

    if (!hasRealRows) {
      return;
    }

    jQuery(table).DataTable({
      paging: false,
      autoWidth: false,
      scrollX: true,
      scrollCollapse: true,
      order: [],
      dom: "<'d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3'Bf>rt<'d-flex flex-wrap justify-content-between align-items-center gap-2 mt-3'i>",
      buttons: [
        { extend: "copyHtml5", text: "Copiar", className: "btn btn-glass btn-sm" },
        { extend: "csvHtml5", text: "CSV", className: "btn btn-glass btn-sm" },
        { extend: "excelHtml5", text: "Excel", className: "btn btn-glass btn-sm" },
        { extend: "pdfHtml5", text: "PDF", className: "btn btn-glass btn-sm" },
        { extend: "print", text: "Imprimir", className: "btn btn-glass btn-sm" },
        { extend: "colvis", text: "Columnas", className: "btn btn-glass btn-sm" },
      ],
      language: {
        search: "Buscar:",
        lengthMenu: "Mostrar _MENU_ registros",
        info: "Mostrando _START_ a _END_ de _TOTAL_ registros",
        infoEmpty: "Sin registros disponibles",
        infoFiltered: "(filtrado de _MAX_ registros totales)",
        zeroRecords: "No se encontraron coincidencias",
        paginate: {
          first: "Primero",
          last: "Ultimo",
          next: "Siguiente",
          previous: "Anterior",
        },
      },
      initComplete() {
        const wrapper = this.api().table().container();
        wrapper.querySelectorAll(".dt-buttons .btn").forEach((btn) => btn.classList.add("me-1", "mb-1"));
      },
    });
  });
});
