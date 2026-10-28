"use strict";

// Sorting has to intercept OWA's own requests, so the code is loaded into the
// page context before the application starts.
//
// Preferences (folders sorted by category, folded categories) are kept in the
// extension's storage rather than the page's: Firefox keeps it when site data is
// cleared, and OWA cannot wipe it on sign-out. The page script asks for them with
// a "hello" event and sends them back with a "save" event.
const PREFS_KEY = `prefs:${location.origin}`;
const prefs = browser.storage.local
  .get(PREFS_KEY)
  .then((stored) => stored[PREFS_KEY] || null)
  .catch(() => null);

document.addEventListener("cat-filter-for-owa:hello", () => {
  prefs.then((value) => {
    document.dispatchEvent(new CustomEvent("cat-filter-for-owa:prefs", { detail: JSON.stringify(value) }));
  });
});

document.addEventListener("cat-filter-for-owa:save", (event) => {
  try {
    browser.storage.local.set({ [PREFS_KEY]: JSON.parse(event.detail) });
  } catch {
    // Malformed message: ignored.
  }
});

for (const file of ["src/sort-core.js", "src/page.js"]) {
  const script = document.createElement("script");
  script.src = browser.runtime.getURL(file);
  script.async = false;
  script.addEventListener("load", () => script.remove());
  (document.head || document.documentElement).append(script);
}
