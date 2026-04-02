// compare-plugin/static/js/trends.js

/**
 * trends.js
 * ---------
 * This module manages the rendering, layout, and updates of trend-related charts in the application.
 * It uses Chart.js to create and manipulate charts, ensuring proper responsiveness and layout.
 *
 * Purpose:
 * - Initializes and updates charts for visualizing trends.
 * - Ensures charts are properly wrapped and styled to avoid layout issues.
 * - Manages chart instances and prevents overlapping renders.
 * - Handles dynamic resizing and responsiveness of charts.
 *
 * Key Features:
 * - `ensureTrendBox(canvas, heightPx)`:
 *   - Ensures that the `<canvas>` element used by Chart.js is wrapped in a fixed-height container.
 *   - Prevents layout issues caused by Chart.js's responsive behavior.
 * - Chart Instances:
 *   - Manages `overallLineChart` and `countsStackedChart` instances to avoid redundant renders.
 * - Request Sequence Management:
 *   - Tracks asynchronous requests with `trendsReqSeq` to ensure only the latest data is used for rendering.
 * - Dynamic Chart Resizing:
 *   - Adjusts chart dimensions dynamically to fit the UI layout.
 *
 * Usage:
 * - This file is included in the application's frontend to manage trend-related charts.
 * - Functions are called to initialize, update, and manage charts dynamically based on user interactions or data updates.
 */

// Hold chart instances and a request sequence to avoid overlapping renders.
let overallLineChart = null;
let countsStackedChart = null;
let trendsReqSeq = 0;

/**
 * Ensure the canvas has a fixed-height wrapper that Chart.js will respect.
 * This avoids growth caused by responsive: true + maintainAspectRatio: false.
 */
function ensureTrendBox(canvas, heightPx) {
  if (!canvas) return null;

  // If already wrapped from a previous call, just update the height.
  if (canvas.parentElement && canvas.parentElement.classList.contains('trendbox')) {
    canvas.parentElement.style.height = heightPx + 'px';
    return canvas.parentElement;
  }

  // Create a wrapper right before the canvas and move the canvas inside it.
  const wrapper = document.createElement('div');
  wrapper.className = 'trendbox';
  wrapper.style.height = heightPx + 'px';
  wrapper.style.width = '100%';
  wrapper.style.position = 'relative';

  const parent = canvas.parentElement;
  parent.insertBefore(wrapper, canvas);
  wrapper.appendChild(canvas);

  // Let Chart.js use 100% to fill the wrapper height.
  canvas.style.height = '100%';
  canvas.style.width  = '100%';

  return wrapper;
}

/**
 * Desktop-only context menu to delete the currently selected snapshot.
 * Creates the menu dynamically and wires events to the Snapshot selector.
 */
function initSnapshotDeleteMenu() {
  const domainSel = document.getElementById('domain');
  const snapshotSel = document.getElementById('snapshot');
  if (!domainSel || !snapshotSel) return;

  // Inject lightweight styles for the context menu (once).
  if (!document.getElementById('snapshotMenuStyles')) {
    const style = document.createElement('style');
    style.id = 'snapshotMenuStyles';
    style.textContent = `
      .contextmenu{position:fixed;background:#0f1620;border:1px solid rgba(255,255,255,0.12);
        border-radius:8px;padding:6px;display:none;min-width:220px;box-shadow:0 8px 24px rgba(0,0,0,0.4);z-index:1000}
      .contextmenu .item{padding:8px 10px;cursor:pointer;font-size:13px;color:#e8eef7}
      .contextmenu .item:hover{background:rgba(255,255,255,0.06)}
      .contextmenu .item.danger:hover{background:rgba(192,0,0,0.15);color:#ffd7d7}
    `;
    document.head.appendChild(style);
  }

  // Build the menu once and attach to body.
  const menu = document.createElement('div');
  menu.id = 'snapshotMenu';
  menu.className = 'contextmenu';
  menu.setAttribute('role', 'menu');
  menu.setAttribute('aria-hidden', 'true');

  const btnDelete = document.createElement('div');
  btnDelete.className = 'item danger';
  btnDelete.id = 'menuDelete';
  btnDelete.textContent = 'Delete snapshot…';

  const btnCancel = document.createElement('div');
  btnCancel.className = 'item';
  btnCancel.id = 'menuCancel';
  btnCancel.textContent = 'Cancel';

  menu.appendChild(btnDelete);
  menu.appendChild(btnCancel);
  document.body.appendChild(menu);

  let menuOpen = false;

  function hideMenu() {
    if (!menuOpen) return;
    menu.style.display = 'none';
    menu.setAttribute('aria-hidden', 'true');
    menuOpen = false;
  }

  function showMenuAt(x, y) {
    menu.style.display = 'block';
    menu.setAttribute('aria-hidden', 'false');
    // Clamp within viewport
    const margin = 8;
    const w = menu.offsetWidth;
    const h = menu.offsetHeight;
    let left = x;
    let top = y;
    if (left + w > window.innerWidth - margin) left = window.innerWidth - w - margin;
    if (top + h > window.innerHeight - margin) top = window.innerHeight - h - margin;
    menu.style.left = left + 'px';
    menu.style.top = top + 'px';
    menuOpen = true;
  }

  // Right-click on Snapshot selector opens the menu (except for LATEST).
  snapshotSel.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const file = snapshotSel.value;
    if (!file || file === 'LATEST') return;
    showMenuAt(e.clientX, e.clientY);
  });

  // Close on outside click or Escape.
  document.addEventListener('click', (e) => {
    if (menuOpen && !menu.contains(e.target)) hideMenu();
  });
  document.addEventListener('keydown', (e) => {
    if (menuOpen && e.key === 'Escape') hideMenu();
    // Optional: open via Shift+F10 when snapshot select is focused.
    if (document.activeElement === snapshotSel && e.shiftKey && e.key === 'F10') {
      e.preventDefault();
      const rect = snapshotSel.getBoundingClientRect();
      showMenuAt(rect.left + 10, rect.top + 10);
    }
  });

  btnCancel.addEventListener('click', hideMenu);

  // Delete handler posts to backend and refreshes UI if globals exist.
  btnDelete.addEventListener('click', async () => {
    const domain = (domainSel.value || '').toLowerCase();
    const file = snapshotSel.value;
    if (!file || file === 'LATEST') { hideMenu(); return; }

    const ok = window.confirm(`Delete snapshot:\n${file}\n\nThis will remove it from history and trends.`);
    if (!ok) return;

    try {
      const res = await fetch('/api/history/delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ domain, file })
      });
      if (!res.ok) {
        alert(`Delete failed (status ${res.status}).`);
        return;
      }
      hideMenu();

      // Refresh UI safely if functions are available.
      if (typeof loadHistory === 'function') await loadHistory();
      if (typeof loadApps === 'function') await loadApps();
      if (typeof loadInsights === 'function') await loadInsights();
      if (typeof triggerTrendsRefresh === 'function') triggerTrendsRefresh();
    } catch (err) {
      alert('Delete failed: ' + err);
    }
  });
}

// Initialize the desktop menu after DOM is ready.
document.addEventListener('DOMContentLoaded', initSnapshotDeleteMenu);

/**
 * Load and render controller trends for a domain.
 * Safe against concurrent calls and reuses existing Chart instances.
 */
async function loadTrends(domain, controller) {
  // Sequence guard: ignore stale responses.
  const seq = ++trendsReqSeq;

  try {
    const q = new URLSearchParams({ domain, controller, limit: 20, baseline: 'earliestprev' });
    const res = await fetch(`/api/trends/runs?${q.toString()}`);
    const data = await res.json();

    // If a newer call superseded this one, abort.
    if (seq !== trendsReqSeq) return;

    const items    = Array.isArray(data.items) ? data.items : [];
    const labels   = items.map(i => `${i.previousDate} → ${i.currentDate}`);
    // Use counts for both charts.
    const improved = items.map(i => Number(i.improved ?? 0));
    const degraded = items.map(i => Number(i.degraded ?? 0));

    const overallCanvas = document.getElementById('overallLine');
    const countsCanvas  = document.getElementById('countsStacked');
    if (!overallCanvas || !countsCanvas) return;

    // Fix heights via wrapper so responsive resizing fills these exact values.
    ensureTrendBox(overallCanvas, 140);
    ensureTrendBox(countsCanvas, 160);

    // Update-or-create: Two-line chart (Improved vs Declined) based on counts.
    const existingOverall = Chart.getChart ? Chart.getChart(overallCanvas) : null;
    if (existingOverall && existingOverall.config?.type === 'line' && existingOverall.data?.datasets?.length === 2) {
      existingOverall.data.labels = labels;
      existingOverall.data.datasets[0].label = 'Improved';
      existingOverall.data.datasets[0].data = improved;
      existingOverall.data.datasets[0].borderColor = '#35c56d';
      existingOverall.data.datasets[0].backgroundColor = 'rgba(53,197,109,0.15)';
      existingOverall.data.datasets[1].label = 'Declined';
      existingOverall.data.datasets[1].data = degraded;
      existingOverall.data.datasets[1].borderColor = '#c00000';
      existingOverall.data.datasets[1].backgroundColor = 'rgba(192,0,0,0.15)';
      existingOverall.update();
      overallLineChart = existingOverall;
    } else {
      // Recreate to ensure correct structure if the existing chart was percentage-based.
      try { existingOverall?.destroy(); } catch {}
      overallLineChart = new Chart(overallCanvas, {
        type: 'line',
        data: {
          labels,
          datasets: [
            {
              label: 'Improved',
              data: improved,
              borderColor: '#35c56d',
              backgroundColor: 'rgba(53,197,109,0.15)',
              pointBackgroundColor: '#35c56d',
              pointRadius: 2,
              tension: 0.25,
              fill: false
            },
            {
              label: 'Declined',
              data: degraded,
              borderColor: '#c00000',
              backgroundColor: 'rgba(192,0,0,0.15)',
              pointBackgroundColor: '#c00000',
              pointRadius: 2,
              tension: 0.25,
              fill: false
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false, // fill wrapper height
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'bottom' },
            tooltip: {
              callbacks: { label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y}` }
            }
          },
          scales: {
            x: { title: { display: true, text: 'Run (previous → current)' } },
            y: { beginAtZero: true, title: { display: true, text: 'Metric count' } }
          }
        }
      });
    }

    // Update-or-create: Improved vs Degraded stacked bars.
    const existingCounts = Chart.getChart ? Chart.getChart(countsCanvas) : null;
    if (existingCounts) {
      existingCounts.data.labels = labels;
      existingCounts.data.datasets[0].data = improved;
      existingCounts.data.datasets[1].data = degraded;
      existingCounts.update();
      countsStackedChart = existingCounts;
    } else {
      countsStackedChart = new Chart(countsCanvas, {
        type: 'bar',
        data: {
          labels,
          datasets: [
            { label: 'Improved', data: improved, backgroundColor: '#2ecc71' },
            { label: 'Degraded', data: degraded, backgroundColor: '#e74c3c' }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false, // fill wrapper height
          scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } },
          plugins: { legend: { position: 'bottom' } }
        }
      });
    }
  } catch (err) {
    console.error('loadTrends failed:', err);
  }
}

// Expose for insights.html to call.
window.loadTrends = loadTrends;

// Optional cleanup on page unload.
window.addEventListener('beforeunload', () => {
  try { overallLineChart?.destroy(); } catch {}
  try { countsStackedChart?.destroy(); } catch {}
});