(() => {
  "use strict";

  const navRoot = document.getElementById("courseNav");
  const frame = document.getElementById("contentFrame");
  const pageTitle = document.getElementById("pageTitle");
  const search = document.getElementById("courseSearch");
  const completeButton = document.getElementById("completeButton");
  const completeText = completeButton.querySelector("span");
  const progressText = document.getElementById("progressText");
  const progressBar = document.getElementById("progressBar");
  const openPageButton = document.getElementById("openPageButton");
  const previousButton = document.getElementById("previousButton");
  const nextButton = document.getElementById("nextButton");
  const previousLabel = document.getElementById("previousLabel");
  const nextLabel = document.getElementById("nextLabel");
  const sidebar = document.getElementById("sidebar");
  const menuButton = document.getElementById("menuButton");
  const mobileScrim = document.getElementById("mobileScrim");

  const STORAGE_KEY = "larson-instructor-completed-pages";
  const allItems = window.COURSE_NAV.flatMap(section =>
    section.items.map(item => ({ ...item, section: section.section }))
  );
  let activeIndex = 0;
  let completed = new Set(readCompleted());

  function readCompleted() {
    try {
      const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      return Array.isArray(value) ? value : [];
    } catch (error) {
      return [];
    }
  }

  function saveCompleted() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify([...completed]));
    } catch (error) {
      // Local storage can be disabled; navigation still works without persistence.
    }
  }

  function iconMarkup() {
    return `<svg aria-hidden="true" viewBox="0 0 24 24"><path d="m5 12 4 4L19 6"/></svg>`;
  }

  function renderNavigation() {
    navRoot.innerHTML = window.COURSE_NAV.map((section, sectionIndex) => {
      if (section.direct && section.items.length === 1) {
        const item = section.items[0];
        return `
          <a class="nav-link nav-direct" href="${item.path}" target="contentFrame" data-id="${item.id}" data-title="${escapeHtml(item.title)}" data-section="${escapeHtml(section.section)}">
            <span>${escapeHtml(item.title)}</span>
            <span class="nav-status" aria-label="Not complete">${iconMarkup()}</span>
          </a>`;
      }

      const links = section.items.map(item => {
        const simulationNumber = item.id.match(/^simulation-(\d+)$/)?.[1] || "";
        return `
          <a class="nav-link" href="${item.path}" target="contentFrame" data-id="${item.id}" data-title="${escapeHtml(item.title)}" data-section="${escapeHtml(section.section)}">
            <span class="nav-number${simulationNumber ? "" : " is-empty"}" aria-hidden="true">${simulationNumber}</span>
            <span>${escapeHtml(item.title)}</span>
            <span class="nav-status" aria-label="Not complete">${iconMarkup()}</span>
          </a>`;
      }).join("");

      return `
        <details class="nav-section" ${section.open ? "open" : ""} data-section-index="${sectionIndex}">
          <summary><svg class="chevron" aria-hidden="true" viewBox="0 0 24 24"><path d="m9 18 6-6-6-6"/></svg>${escapeHtml(section.section)}</summary>
          <div class="nav-items">${links}</div>
        </details>`;
    }).join("");

    navRoot.querySelectorAll(".nav-link").forEach(link => {
      link.addEventListener("click", event => {
        if (event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
        event.preventDefault();
        const index = allItems.findIndex(item => item.id === link.dataset.id);
        openItem(index);
      });
    });
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function openItem(index) {
    if (index < 0 || index >= allItems.length) return;
    activeIndex = index;
    const item = allItems[activeIndex];

    frame.src = item.path;
    frame.title = item.title;
    pageTitle.textContent = item.title;
    document.title = `${item.title} | Clinical Simulations for the Athletic Trainer`;

    navRoot.querySelectorAll(".nav-link").forEach(link => {
      const active = link.dataset.id === item.id;
      if (active) link.setAttribute("aria-current", "page");
      else link.removeAttribute("aria-current");
    });

    const activeLink = navRoot.querySelector(`[data-id="${CSS.escape(item.id)}"]`);
    activeLink?.closest("details")?.setAttribute("open", "");
    activeLink?.scrollIntoView({ block: "nearest" });

    updateCompleteButton();
    updatePageControls();
    closeMobileMenu();
  }

  function updatePageControls() {
    const previous = allItems[activeIndex - 1];
    const next = allItems[activeIndex + 1];

    previousButton.disabled = !previous;
    nextButton.disabled = !next;
    previousLabel.textContent = previous?.title || "Beginning of course";
    nextLabel.textContent = next?.title || "End of course";
  }

  function updateCompleteButton() {
    const item = allItems[activeIndex];
    const isComplete = completed.has(item.id);
    completeButton.setAttribute("aria-pressed", String(isComplete));
    completeText.textContent = isComplete ? "Completed" : "Mark complete";
  }

  function updateProgress() {
    navRoot.querySelectorAll(".nav-link").forEach(link => {
      const isComplete = completed.has(link.dataset.id);
      link.classList.toggle("is-complete", isComplete);
      const status = link.querySelector(".nav-status");
      status.setAttribute("aria-label", isComplete ? "Complete" : "Not complete");
    });

    const count = allItems.filter(item => completed.has(item.id)).length;
    progressText.textContent = `${count} of ${allItems.length} complete`;
    progressBar.style.width = `${allItems.length ? (count / allItems.length) * 100 : 0}%`;
  }

  function toggleComplete() {
    const id = allItems[activeIndex].id;
    if (completed.has(id)) completed.delete(id);
    else completed.add(id);
    saveCompleted();
    updateCompleteButton();
    updateProgress();
  }

  function filterNavigation(query) {
    const normalized = query.trim().toLowerCase();
    navRoot.querySelectorAll(".nav-section").forEach(section => {
      let visibleCount = 0;
      section.querySelectorAll(".nav-link").forEach(link => {
        const visible = !normalized || link.dataset.title.toLowerCase().includes(normalized);
        link.hidden = !visible;
        if (visible) visibleCount += 1;
      });
      section.hidden = visibleCount === 0;
      if (normalized && visibleCount) section.open = true;
    });
  }

  function openMobileMenu() {
    sidebar.classList.add("is-open");
    mobileScrim.hidden = false;
    menuButton.setAttribute("aria-expanded", "true");
  }

  function closeMobileMenu() {
    sidebar.classList.remove("is-open");
    mobileScrim.hidden = true;
    menuButton.setAttribute("aria-expanded", "false");
  }

  renderNavigation();
  updateProgress();
  openItem(0);

  search.addEventListener("input", event => filterNavigation(event.target.value));
  completeButton.addEventListener("click", toggleComplete);
  openPageButton.addEventListener("click", () => window.open(allItems[activeIndex].path, "_blank", "noopener"));
  previousButton.addEventListener("click", () => openItem(activeIndex - 1));
  nextButton.addEventListener("click", () => openItem(activeIndex + 1));
  menuButton.addEventListener("click", () => sidebar.classList.contains("is-open") ? closeMobileMenu() : openMobileMenu());
  mobileScrim.addEventListener("click", closeMobileMenu);
  window.addEventListener("keydown", event => { if (event.key === "Escape") closeMobileMenu(); });
})();
