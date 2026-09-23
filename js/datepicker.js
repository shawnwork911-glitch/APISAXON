/* =====================================================================
   Date picker
   ---------------------------------------------------------------------
   Native <input type="date"> always shows the browser/OS locale's date
   format (mm/dd/yyyy, dd/mm/yyyy, etc.) — a page can't force it either
   way. This is a small self-contained calendar popup instead, so the
   displayed format (dd/mm/yyyy) is guaranteed regardless of the
   viewer's locale.

   Usage: DatePicker.attach(inputEl) — the input should be a plain text
   input (readonly, so typing can't produce an unparseable value). The
   visible text is dd/mm/yyyy; the underlying ISO value (yyyy-mm-dd, what
   the rest of the app actually needs) is kept in inputEl.dataset.iso —
   read it with DatePicker.getISO(inputEl) rather than inputEl.value.
   ===================================================================== */

const DatePicker = (() => {
  const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const DOW = ["Su","Mo","Tu","We","Th","Fr","Sa"];
  let openPopup = null;
  let openInput = null;

  function pad(n) { return String(n).padStart(2, "0"); }
  function toISO(y, m, d) { return `${y}-${pad(m + 1)}-${pad(d)}`; }
  function toDisplay(y, m, d) { return `${pad(d)}/${pad(m + 1)}/${y}`; }

  function attach(inputEl) {
    const wrap = inputEl.closest(".date-field-wrap");
    const btn = wrap ? wrap.querySelector(".date-picker-btn") : null;
    const open = () => togglePopup(inputEl, wrap);
    inputEl.addEventListener("click", open);
    if (btn) btn.addEventListener("click", open);
  }

  function getISO(inputEl) { return inputEl?.dataset.iso || ""; }

  function setValue(inputEl, y, m, d) {
    inputEl.value = toDisplay(y, m, d);
    inputEl.dataset.iso = toISO(y, m, d);
    inputEl.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function togglePopup(inputEl, wrap) {
    if (openPopup && openInput === inputEl) { closePopup(); return; }
    closePopup();
    const iso = getISO(inputEl);
    const base = iso ? new Date(`${iso}T00:00:00`) : new Date();
    const popup = document.createElement("div");
    popup.className = "date-picker-popup";
    (wrap || inputEl.parentElement).appendChild(popup);
    openPopup = popup;
    openInput = inputEl;
    renderCalendar(popup, inputEl, base.getFullYear(), base.getMonth());

    setTimeout(() => document.addEventListener("click", onOutsideClick), 0);
  }

  function onOutsideClick(e) {
    if (openPopup && !openPopup.contains(e.target) && e.target !== openInput && !e.target.closest(".date-picker-btn")) {
      closePopup();
    }
  }

  function closePopup() {
    if (openPopup) openPopup.remove();
    openPopup = null;
    openInput = null;
    document.removeEventListener("click", onOutsideClick);
  }

  function renderCalendar(popup, inputEl, year, month) {
    const today = new Date();
    const selectedIso = getISO(inputEl);
    const firstDow = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();

    let cells = "";
    for (let i = 0; i < firstDow; i++) {
      const d = daysInPrevMonth - firstDow + 1 + i;
      cells += `<div class="day muted" data-y="${month === 0 ? year - 1 : year}" data-m="${month === 0 ? 11 : month - 1}" data-d="${d}">${d}</div>`;
    }
    for (let d = 1; d <= daysInMonth; d++) {
      const iso = toISO(year, month, d);
      const isToday = iso === toISO(today.getFullYear(), today.getMonth(), today.getDate());
      const isSelected = iso === selectedIso;
      cells += `<div class="day${isToday ? " today" : ""}${isSelected ? " selected" : ""}" data-y="${year}" data-m="${month}" data-d="${d}">${d}</div>`;
    }
    const totalCells = firstDow + daysInMonth;
    const trailing = (7 - (totalCells % 7)) % 7;
    for (let d = 1; d <= trailing; d++) {
      cells += `<div class="day muted" data-y="${month === 11 ? year + 1 : year}" data-m="${month === 11 ? 0 : month + 1}" data-d="${d}">${d}</div>`;
    }

    popup.innerHTML = `
      <div class="date-picker-header">
        <button type="button" class="date-picker-nav" data-nav="-1">‹</button>
        <span>${MONTH_NAMES[month]} ${year}</span>
        <button type="button" class="date-picker-nav" data-nav="1">›</button>
      </div>
      <div class="date-picker-grid">
        ${DOW.map(d => `<div class="dow">${d}</div>`).join("")}
        ${cells}
      </div>`;

    popup.querySelector('[data-nav="-1"]').addEventListener("click", (e) => {
      e.stopPropagation();
      const m = month === 0 ? 11 : month - 1;
      renderCalendar(popup, inputEl, month === 0 ? year - 1 : year, m);
    });
    popup.querySelector('[data-nav="1"]').addEventListener("click", (e) => {
      e.stopPropagation();
      const m = month === 11 ? 0 : month + 1;
      renderCalendar(popup, inputEl, month === 11 ? year + 1 : year, m);
    });
    popup.querySelectorAll(".day").forEach(el => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        setValue(inputEl, Number(el.dataset.y), Number(el.dataset.m), Number(el.dataset.d));
        closePopup();
      });
    });
  }

  return { attach, getISO };
})();
