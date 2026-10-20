"use strict";

// Capture a separate snapshot so switching tabs or changing files cannot mix reports.
function buildSprintReportSnapshot() {
  const host = document.createElement("div");
  host.className = "pdf-report-host";
  host.setAttribute("aria-hidden", "true");
  const dashboard = document.createElement("div");
  dashboard.className = "pdf-report-content";
  document.querySelectorAll("#dashboard > .panel").forEach(panel => dashboard.append(panel.cloneNode(true)));
  // Textareas scroll on screen; export the complete, currently edited comment.
  const inputs = document.querySelectorAll("#dashboard textarea");
  dashboard.querySelectorAll("textarea").forEach((input, index) => {
    const text = document.createElement("span");
    text.className = "pdf-comment";
    text.textContent = inputs[index].value;
    input.replaceWith(text);
  });
  const commitments = document.createElement("div");
  commitments.className = "pdf-report-content pdf-commitments";
  const cards = document.getElementById("burndown-cards").cloneNode(true);
  if (cards.children.length !== 6) throw new Error("The report requires all six sprint commitment cards.");
  commitments.append(cards);
  host.append(dashboard, commitments);
  const stories = document.getElementById("user-story-cards").cloneNode(true);
  const storyMessage = document.getElementById("user-story-status-message").textContent;
  const burndown = document.getElementById("burndown-site-charts").cloneNode(true);
  const burndownMessage = document.getElementById("burndown-shared-status").textContent;
  const burndownTitle = document.querySelector("#burndown-chart h1").textContent.trim();
  // Keep export IDs separate from the live dashboard while uploads/tabs change.
  host.querySelectorAll("[id]").forEach(element => { element.id = `pdf-${element.id}`; });
  const chart = host.querySelector("#pdf-progress-chart");
  if (chart) chart.style.minWidth = "0";
  document.body.append(host);
  return { host, pages: [dashboard, commitments], stories, storyMessage, burndown, burndownMessage, burndownTitle };
}

// Keep each chart, its dates, metrics and legend together. Two columns allow
// readable axes without reducing the entire section to a single-page image.
function paginateReportBurndown(snapshot) {
  const pages = [];
  let grid;
  const newPage = () => {
    const page = document.createElement("div");
    page.className = "pdf-report-content pdf-burndown-page";
    page.dataset.reportTitle = snapshot.burndownTitle || "Sprint burndown chart report";
    grid = document.createElement("div");
    grid.className = "pdf-burndown-grid";
    page.append(grid);
    snapshot.host.append(page);
    pages.push(page);
  };
  newPage();
  if (snapshot.burndownMessage) {
    const message = document.createElement("p");
    message.className = "pdf-burndown-source";
    message.textContent = snapshot.burndownMessage;
    grid.append(message);
  }
  for (const source of snapshot.burndown.children) {
    const card = source.cloneNode(true);
    card.querySelectorAll("[id]").forEach(node => node.removeAttribute("id"));
    card.removeAttribute("aria-labelledby");
    grid.append(card);
    if (grid.getBoundingClientRect().height > 638) {
      card.remove();
      if (grid.children.length) newPage();
      grid.append(card);
      if (grid.getBoundingClientRect().height > 638) {
        throw new Error("A burndown chart card is too tall to fit legibly on a PDF page.");
      }
    }
  }
  return pages;
}

// Three columns at a fixed reading size. Pack cards first, split only at story
// or task boundaries, and repeat the site/story identity on continuation cards.
function paginateReportStories(snapshot) {
  const pages = [];
  let column, columns, columnIndex = 2;
  const nextColumn = () => {
    if (++columnIndex >= 3) {
      const page = document.createElement("div");
      page.className = "pdf-report-content pdf-story-page";
      page.dataset.reportTitle = "User Story Status";
      columns = Array.from({ length: 3 }, () => {
        const node = document.createElement("div");
        node.className = "pdf-story-column";
        page.append(node);
        return node;
      });
      snapshot.host.append(page);
      pages.push(page);
      columnIndex = 0;
    }
    column = columns[columnIndex];
  };
  const fits = () => !column.lastElementChild ||
    column.lastElementChild.getBoundingClientRect().bottom - column.getBoundingClientRect().top <= 638;
  nextColumn();
  const sourceCards = [...snapshot.stories.children];
  if (!sourceCards.length) {
    const message = document.createElement("p");
    message.textContent = snapshot.storyMessage || "No user story data available.";
    column.append(message);
  }
  for (const source of sourceCards) {
    const whole = source.cloneNode(true);
    column.append(whole);
    if (fits()) continue;
    whole.remove();
    // Move an intact card to the next column if it can fit there.
    const probe = source.cloneNode(true);
    probe.style.position = "absolute";
    probe.style.width = `${column.clientWidth}px`;
    column.append(probe);
    const wholeHeight = probe.getBoundingClientRect().height;
    probe.remove();
    if (wholeHeight <= 638) {
      nextColumn();
      column.append(whole);
      continue;
    }
    let fragment, trees, continued = false;
    const startFragment = () => {
      fragment = source.cloneNode(true);
      trees = fragment.querySelector(".user-story-trees");
      trees.replaceChildren();
      if (continued) {
        const note = document.createElement("span");
        note.className = "pdf-story-continuation";
        note.textContent = "Continued";
        fragment.querySelector(".user-story-card-header").append(note);
      }
      column.append(fragment);
    };
    const advance = () => {
      if (!trees.children.length) fragment.remove();
      else continued = true;
      nextColumn();
      startFragment();
    };
    startFragment();
    for (const sourceTree of source.querySelector(".user-story-trees").children) {
      let tree = sourceTree.cloneNode(true);
      trees.append(tree);
      if (fits()) continue;
      tree.remove();
      // Try the whole story in an empty column before splitting its tasks.
      if (column.children.length > 1 || trees.children.length) advance();
      trees.append(tree);
      if (fits()) continue;
      tree.remove();
      const sourceTasks = [...sourceTree.querySelectorAll(".user-story-task")];
      if (!sourceTasks.length) throw new Error("A user story is too large to fit on a PDF page.");
      let taskList;
      const startTree = (isContinuation) => {
        tree = sourceTree.cloneNode(true);
        taskList = tree.querySelector(".user-story-tasks");
        taskList.replaceChildren();
        if (isContinuation) tree.querySelector(".story-progress-label").append(" (continued)");
        trees.append(tree);
      };
      startTree(false);
      for (const task of sourceTasks) {
        const row = task.cloneNode(true);
        taskList.append(row);
        if (fits()) continue;
        row.remove();
        if (!taskList.children.length) throw new Error("A task is too large to fit on a PDF page.");
        advance();
        startTree(true);
        taskList.append(row);
        if (!fits()) throw new Error("A task is too large to fit on a PDF page.");
      }
    }
  }
  // Balance the columns within each already-packed page (especially a short
  // final page), without increasing page count or changing reading order.
  pages.forEach(page => {
    const pageColumns = [...page.children];
    const cards = pageColumns.flatMap(node => [...node.children]);
    if (cards.length && cards.every(card => card.classList.contains("user-story-card"))) {
      const heights = cards.map(card => card.getBoundingClientRect().height);
      const columnsNeeded = limit => {
        let count = 1, height = 0;
        heights.forEach(h => {
          if (height && height + 12 + h > limit) { count++; height = 0; }
          height += (height ? 12 : 0) + h;
        });
        return count;
      };
      let low = Math.max(...heights), high = 638;
      for (let i = 0; i < 12; i++) {
        const middle = (low + high) / 2;
        if (columnsNeeded(middle) <= 3) high = middle;
        else low = middle;
      }
      let index = 0, height = 0;
      cards.forEach((card, i) => {
        if (height && height + 12 + heights[i] > Math.ceil(high)) { index++; height = 0; }
        pageColumns[index].append(card);
        height += (height ? 12 : 0) + heights[i];
      });
    }
    page.querySelectorAll("[id]").forEach(node => node.removeAttribute("id"));
  });
  return pages;
}

function addSprintReportBranding(pdf, title, pageNumber) {
  const width = pdf.internal.pageSize.getWidth();
  const height = pdf.internal.pageSize.getHeight();
  pdf.setTextColor(7, 19, 62);
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(15);
  pdf.text(title, 12, 17, { maxWidth: width - 24 });
  pdf.setDrawColor(0, 137, 155);
  pdf.setLineWidth(0.6);
  pdf.line(12, 23, width - 12, 23);
  // The supplied image has large transparent margins. Clip to its banner band in
  // PDF coordinates so the original asset stays intact and is never stretched.
  const scale = width / 2170;
  const bandTop = 275, bandBottom = 420;
  const footerTop = height - (bandBottom - bandTop) * scale;
  pdf.saveGraphicsState();
  // null leaves the path open for clip(); the default stroke consumes it,
  // producing an empty clipping path that some PDF viewers reject.
  pdf.rect(0, footerTop, width, height - footerTop, null);
  pdf.clip();
  pdf.discardPath();
  pdf.addImage(sprintReportFooter, "PNG", 0, footerTop - bandTop * scale, width, 725 * scale, "report-footer", "FAST");
  pdf.restoreGraphicsState();
  // Replace the sample's printed "2" with the actual PDF page number.
  pdf.setFillColor(255, 255, 255);
  pdf.rect(77 * scale, footerTop + (327 - bandTop) * scale, 34 * scale, 39 * scale, "F");
  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(26 * scale * 72 / 25.4);
  pdf.setTextColor(20, 73, 120);
  pdf.text(String(pageNumber), 90 * scale, footerTop + (361 - bandTop) * scale, { align: "center" });
}

function formatSprintReportCoverDate(date = new Date()) {
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit", month: "long", year: "numeric", timeZone: "Asia/Kolkata"
  }).format(date);
}

function addSprintReportCover(pdf, date) {
  const width = pdf.internal.pageSize.getWidth();
  const height = pdf.internal.pageSize.getHeight();
  pdf.addImage(sprintReportCover, "PNG", 0, 0, width, height, "report-cover", "FAST");
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(32 * width / 1672 * 72 / 25.4);
  pdf.setTextColor(10, 10, 10);
  pdf.text(formatSprintReportCoverDate(date), width * 98 / 1672, height * 760 / 941);
}

async function downloadSprintReport() {
  const button = document.getElementById("create-report-button");
  if (button.disabled) return;
  const label = button.querySelector("span");
  const status = document.getElementById("report-status");
  button.disabled = true;
  button.setAttribute("aria-busy", "true");
  label.textContent = "Creating PDF…";
  status.textContent = "Preparing your report…";
  let snapshot;
  const reportDate = new Date();
  try {
    if (!window.jspdf?.jsPDF || typeof html2canvas !== "function") {
      throw new Error("PDF libraries could not load. Refresh the page with the vendor folder available.");
    }
    const period = document.getElementById("reporting-period").textContent.trim().replace(/^[–—-]\s*/, "");
    const title = `Sprint health Status - ${period || "Reporting period unavailable"}`;
    const filename = `Sprint-health-status-${selectedDashboardPeriod ? `${selectedDashboardPeriod.from}_to_${selectedDashboardPeriod.to}` : "report"}.pdf`;
    snapshot = buildSprintReportSnapshot();
    await document.fonts.ready;
    snapshot.pages.push(...paginateReportStories(snapshot));
    snapshot.pages.push(...paginateReportBurndown(snapshot));
    // Preserve the supplied cover's aspect ratio; report pages remain A4 landscape.
    const pdf = new window.jspdf.jsPDF({ orientation: "landscape", unit: "mm", format: [297, 297 * 941 / 1672], compress: true });
    pdf.setProperties({ title, subject: "Sprint health, commitment, user story status and burndown charts", creator: "Sprint Monitor" });
    addSprintReportCover(pdf, reportDate);
    const totalPages = snapshot.pages.length + 1;
    for (let index = 0; index < snapshot.pages.length; index++) {
      status.textContent = `Creating page ${index + 2} of ${totalPages}…`;
      const page = snapshot.pages[index];
      const canvas = await html2canvas(page, {
        backgroundColor: "#ffffff", scale: 2, logging: false,
        windowWidth: 1600, windowHeight: Math.max(1200, page.scrollHeight),
        scrollX: 0, scrollY: 0
      });
      pdf.addPage("a4", "landscape");
      addSprintReportBranding(pdf, page.dataset.reportTitle || (index ? "Sprint Commitment" : title), index + 2);
      const availableWidth = pdf.internal.pageSize.getWidth() - 24;
      const availableHeight = pdf.internal.pageSize.getHeight() - 51;
      const scale = Math.min(availableWidth / canvas.width, availableHeight / canvas.height);
      const width = canvas.width * scale, height = canvas.height * scale;
      pdf.addImage(canvas, "PNG", (pdf.internal.pageSize.getWidth() - width) / 2, 28, width, height, `page-${index}`, "FAST");
      canvas.width = canvas.height = 0;
    }
    await pdf.save(filename, { returnPromise: true });
    status.textContent = "PDF created. Download started.";
  } catch (error) {
    status.textContent = `Could not create PDF. ${error.message || "Please try again."}`;
  } finally {
    snapshot?.host.remove();
    button.disabled = false;
    button.removeAttribute("aria-busy");
    label.textContent = "Create Report";
  }
}
