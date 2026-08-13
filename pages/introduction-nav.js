(() => {
  "use strict";

  const main = document.querySelector("main.page");
  if (!main || document.getElementById("introduction-navigation")) return;

  const headings = [...main.querySelectorAll("section.content-card > h2[data-manuscript-block]")];
  if (!headings.length) return;

  const navigation = document.createElement("details");
  navigation.className = "content-card introduction-navigation";
  navigation.id = "introduction-navigation";
  navigation.open = true;

  const summary = document.createElement("summary");
  summary.textContent = "Section navigation";

  const grid = document.createElement("div");
  grid.className = "section-jump-grid";

  headings.forEach((heading, index) => {
    const anchor = `introduction-section-${index + 1}`;
    heading.id = anchor;

    const jump = document.createElement("a");
    jump.className = "section-jump";
    jump.href = `#${anchor}`;
    jump.textContent = heading.textContent;
    grid.append(jump);

    const back = document.createElement("a");
    back.className = "section-return";
    back.href = "#introduction-navigation";
    back.textContent = "↑ Return to section navigation";
    heading.closest("section").append(back);
  });

  const copyright = document.createElement("section");
  copyright.className = "content-card";
  copyright.id = "introduction-copyright";

  const copyrightHeading = document.createElement("h2");
  copyrightHeading.textContent = "Copyright";

  const downloads = document.createElement("div");
  downloads.className = "download-grid activity-grid";

  const download = document.createElement("a");
  download.className = "download-card featured";
  download.href = "../assets/downloads/copyright-page-placeholder.docx";
  download.setAttribute("download", "");
  download.innerHTML = '<span class="file-icon" aria-hidden="true">DOCX</span><span><strong>Download Copyright Page</strong><small>Placeholder Word document</small></span><span class="download-arrow" aria-hidden="true">↓</span>';

  const copyrightBack = document.createElement("a");
  copyrightBack.className = "section-return";
  copyrightBack.href = "#introduction-navigation";
  copyrightBack.textContent = "↑ Return to section navigation";

  downloads.append(download);
  copyright.append(copyrightHeading, downloads, copyrightBack);

  const copyrightJump = document.createElement("a");
  copyrightJump.className = "section-jump";
  copyrightJump.href = "#introduction-copyright";
  copyrightJump.textContent = "Copyright";
  grid.append(copyrightJump);

  navigation.append(summary, grid);
  main.prepend(navigation);
  main.append(copyright);
})();
