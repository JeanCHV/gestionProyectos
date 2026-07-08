const chartData = [
  { x: 0.18, y: 22, label: "Acceso" },
  { x: 0.36, y: 54, label: "Proveedor" },
  { x: 0.58, y: 68, label: "Plazos" },
  { x: 0.74, y: 82, label: "Integracion" },
  { x: 0.88, y: 91, label: "Critico" }
];

const pointColors = ["#60a5fa", "#34d399", "#f59e0b", "#fb7185", "#a78bfa"];

const canvas = document.getElementById("riskChart");

if (canvas && window.Chart) {
  const ctx = canvas.getContext("2d");

  new Chart(ctx, {
    type: "scatter",
    data: {
      datasets: [
        {
          label: "Riesgos",
          data: chartData,
          pointRadius: 8,
          pointHoverRadius: 11,
          pointBorderWidth: 2,
          pointBorderColor: "rgba(255,255,255,0.9)",
          pointBackgroundColor: pointColors
        }
      ]
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "rgba(2, 6, 23, 0.92)",
          borderColor: "rgba(255,255,255,0.12)",
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label(context) {
              const point = context.raw;
              return `${point.label} | P: ${point.x.toFixed(2)} | I: ${point.y}`;
            }
          }
        }
      },
      scales: {
        x: {
          min: 0,
          max: 1,
          grid: { color: "rgba(255,255,255,0.08)" },
          ticks: {
            color: "rgba(229,238,252,0.7)",
            callback: (value) => Number(value).toFixed(1)
          },
          title: {
            display: true,
            text: "Probabilidad",
            color: "rgba(229,238,252,0.72)"
          }
        },
        y: {
          min: 0,
          max: 100,
          grid: { color: "rgba(255,255,255,0.08)" },
          ticks: { color: "rgba(229,238,252,0.7)" },
          title: {
            display: true,
            text: "Impacto",
            color: "rgba(229,238,252,0.72)"
          }
        }
      }
    }
  });
}

document.querySelectorAll(".side-link").forEach((link) => {
  link.addEventListener("click", () => {
    document.querySelectorAll(".side-link").forEach((item) => item.classList.remove("active"));
    link.classList.add("active");
  });
});

const sidebar = document.getElementById("sidebar");
const sidebarClose = document.getElementById("sidebarClose");

if (sidebar && sidebarClose && window.bootstrap?.Offcanvas) {
  sidebarClose.addEventListener("click", () => {
    const instance = bootstrap.Offcanvas.getOrCreateInstance(sidebar);
    instance.hide();
  });
}
