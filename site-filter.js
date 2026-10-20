"use strict";

const dashboardSiteNames = ["Alingsas", "Vasteras", "Rakovski", "Heidelberg", "Senatobia", "Chassieu"];

function initializeSiteFilterDismissal(id) {
  const filter = document.getElementById(`${id}-site-filter`);
  document.addEventListener("click", event => {
    if (!filter.contains(event.target)) filter.open = false;
  });
  filter.addEventListener("keydown", event => {
    if (event.key === "Escape") {
      filter.open = false;
      filter.querySelector("summary").focus();
    }
  });
}

// Keep hidden selections across uploads, including additional workbook sites.
function renderSiteFilter(id, names, hiddenSites, onChange) {
  const count = names.filter(name => !hiddenSites.has(name)).length;
  document.getElementById(`${id}-site-selection`).textContent = count === names.length ? "All" : count === 0 ? "None" : `${count} selected`;
  const options = document.getElementById(`${id}-site-options`);
  options.replaceChildren();
  let nextOptionIndex = 0;
  const addOption = (name, checked, change) => {
    const label = document.createElement("label");
    const text = document.createElement("span");
    text.textContent = name;
    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = checked;
    const optionIndex = nextOptionIndex++;
    input.addEventListener("change", () => {
      change(input.checked);
      onChange();
      renderSiteFilter(id, names, hiddenSites, onChange);
      // Restore keyboard focus after rebuilding the options.
      options.children[optionIndex]?.querySelector("input")?.focus();
    });
    label.append(text, input);
    options.append(label);
    return input;
  };
  const all = addOption("All sites", count === names.length, checked => {
    hiddenSites.clear();
    if (!checked) names.forEach(name => hiddenSites.add(name));
  });
  all.indeterminate = count > 0 && count < names.length;
  names.forEach(name => addOption(name, !hiddenSites.has(name), checked => {
    if (checked) hiddenSites.delete(name);
    else hiddenSites.add(name);
  }));
}
