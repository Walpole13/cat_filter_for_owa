"use strict";

// Runs in the OWA page context. Exchange refuses to sort by category, so for
// folders where this sort is chosen, OWA's list requests are intercepted: the
// whole folder is fetched and sorted here, then OWA receives the page it asked
// for and renders it as usual.
(function () {
  const core = window.__catFilterForOwaCore;
  if (!core || window.__catFilterForOwaLoaded) return;
  window.__catFilterForOwaLoaded = true;

  const { ADAPTERS, sortByCategory, pageStart, groupLabels, categoryColors, swatches, locateRun } = core;

  const FRENCH = /^fr/i.test(document.documentElement.lang || navigator.language || "fr");
  const LABEL = FRENCH ? "Catégorie" : "Category";
  const UNCATEGORIZED = FRENCH ? "Sans catégorie" : "No category";
  const FOLDERS_KEY = "cat-filter-for-owa:folders";
  const COLLAPSED_KEY = "cat-filter-for-owa:collapsed";
  const PAGE_SIZE = 200;
  const PARALLEL_REQUESTS = 4;
  const REUSE_MS = 10000;
  const PENDING_MS = 5000;
  const HEADING_HEIGHT = 29;
  const BLANK_IMAGE = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
  const HELLO_EVENT = "cat-filter-for-owa:hello";
  const PREFS_EVENT = "cat-filter-for-owa:prefs";
  const SAVE_EVENT = "cat-filter-for-owa:save";
  const PREFS_TIMEOUT_MS = 1500;
  const CALL = Symbol("cat-filter-for-owa");
  const READ_ONLY_ACTION = /^(Find|Get|Subscribe|Unsubscribe|Ping|Log|Sync)/;

  const enabledFolders = new Set(readStorage(FOLDERS_KEY, []));
  // Collapsed groups, per folder.
  const collapsed = readStorage(COLLAPSED_KEY, {});
  // Whole folders already sorted, reused for the following pages.
  const lists = new Map();
  // Last sorted list handed to OWA, per folder: row ids, group labels and categories by position.
  const served = new Map();
  // Category colors from OWA's master category list, by lowercase name.
  let colors = new Map();
  let colorsRequested = false;
  // Category-sorted folder currently displayed, inferred from the rows on screen.
  let displayed = null;
  // Activation requested from the menu: applies to the folder of the next list loaded.
  let pendingEnableUntil = 0;
  let ownClick = false;

  function readStorage(key, fallback) {
    try {
      return JSON.parse(localStorage.getItem(key)) || fallback;
    } catch {
      return fallback;
    }
  }

  function writeStorage(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // Preference not persisted: harmless for the current session.
    }
  }

  // Preferences live in the extension's storage (see content.js); the page's own
  // storage is only a fallback for when the page script runs on its own.
  function savePrefs() {
    const folders = [...enabledFolders];
    writeStorage(FOLDERS_KEY, folders);
    writeStorage(COLLAPSED_KEY, collapsed);
    document.dispatchEvent(new CustomEvent(SAVE_EVENT, { detail: JSON.stringify({ folders, collapsed }) }));
  }

  function applyPrefs(detail) {
    let stored = null;
    try {
      stored = JSON.parse(detail);
    } catch {
      // Unreadable answer: keep what the page had.
    }
    if (stored && Array.isArray(stored.folders)) {
      enabledFolders.clear();
      for (const folder of stored.folders) enabledFolders.add(folder);
      for (const folder of Object.keys(collapsed)) delete collapsed[folder];
      Object.assign(collapsed, stored.collapsed || {});
      writeStorage(FOLDERS_KEY, stored.folders);
      writeStorage(COLLAPSED_KEY, collapsed);
    } else if (enabledFolders.size || Object.keys(collapsed).length) {
      // Nothing saved by the extension yet: carry over what the page had kept.
      savePrefs();
    }
  }

  // OWA's first list request waits for the preferences, so a folder sorted by
  // category opens sorted right away.
  let prefsLoaded = false;
  const prefsReady = new Promise((resolve) => {
    const done = () => {
      prefsLoaded = true;
      resolve();
    };
    document.addEventListener(
      PREFS_EVENT,
      (event) => {
        applyPrefs(event.detail);
        done();
      },
      { once: true },
    );
    setTimeout(done, PREFS_TIMEOUT_MS);
  });
  document.dispatchEvent(new CustomEvent(HELLO_EVENT));

  function setEnabled(folder, enabled) {
    if (enabled) enabledFolders.add(folder);
    else {
      enabledFolders.delete(folder);
      served.delete(folder);
    }
    savePrefs();
  }

  function toggleCollapsed(folder, label) {
    const labels = new Set(collapsed[folder] || []);
    if (labels.has(label)) labels.delete(label);
    else labels.add(label);
    if (labels.size) collapsed[folder] = [...labels];
    else delete collapsed[folder];
    savePrefs();
  }

  // --- List request interception ---

  const proto = XMLHttpRequest.prototype;
  const native = { open: proto.open, send: proto.send, setRequestHeader: proto.setRequestHeader };
  const nativeResponseText = Object.getOwnPropertyDescriptor(proto, "responseText");
  const nativeResponse = Object.getOwnPropertyDescriptor(proto, "response");

  proto.open = function (method, url) {
    delete this.responseText;
    delete this.response;
    const action = (/[?&]action=(\w+)/.exec(String(url)) || [])[1];
    this[CALL] = ADAPTERS[action] ? { action, headers: [] } : null;
    if (action && !READ_ONLY_ACTION.test(action)) lists.clear();
    if (action === "GetOwaUserConfiguration") {
      this.addEventListener("load", () => {
        try {
          readColors(JSON.parse(nativeResponseText.get.call(this)));
        } catch {
          // Not the expected JSON: colors are requested separately when needed.
        }
      });
    }
    return native.open.apply(this, arguments);
  };

  proto.setRequestHeader = function (name, value) {
    const call = this[CALL];
    if (call) call.headers.push([name, value]);
    else native.setRequestHeader.call(this, name, value);
  };

  proto.send = function (body) {
    const call = this[CALL];
    if (!call) return native.send.apply(this, arguments);
    this[CALL] = null;
    intercept(this, call, body);
  };

  function readRequest(headers, body) {
    const posted = headers.find(([name]) => /^x-owa-urlpostdata$/i.test(name));
    try {
      return { request: JSON.parse(posted ? decodeURIComponent(posted[1]) : body), viaHeader: !!posted };
    } catch {
      return null;
    }
  }

  function intercept(xhr, call, body) {
    if (!prefsLoaded) {
      prefsReady.then(() => intercept(xhr, call, body));
      return;
    }
    let sent = false;
    const send = (headers, payload) => {
      if (sent) return;
      sent = true;
      for (const [name, value] of headers) native.setRequestHeader.call(xhr, name, value);
      native.send.call(xhr, payload);
    };
    const passThrough = () => send(call.headers, body);

    const adapter = ADAPTERS[call.action];
    const parsed = readRequest(call.headers, body);
    const requestBody = parsed && parsed.request && parsed.request.Body;
    const folder = requestBody && adapter.isList(requestBody) && adapter.folderOf(requestBody);
    if (!folder) return passThrough();

    if (pendingEnableUntil > Date.now()) {
      pendingEnableUntil = 0;
      setEnabled(folder, true);
    }
    const paging = requestBody.Paging;
    if (!enabledFolders.has(folder) || !paging || paging.BasePoint !== "Beginning") return passThrough();

    serve(call.action, adapter, parsed.request)
      .then(({ slice, offset, total, rows }) => {
        // Send a minimal request to get a genuine server response,
        // then replace its content with the sorted page.
        const envelope = JSON.parse(JSON.stringify(parsed.request));
        envelope.Body.Paging = { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: 1 };
        const encoded = JSON.stringify(envelope);
        const headers = parsed.viaHeader
          ? call.headers.map(([name, value]) => (/^x-owa-urlpostdata$/i.test(name) ? [name, encodeURIComponent(encoded)] : [name, value]))
          : call.headers;
        overrideResponse(xhr, (text) => {
          const json = JSON.parse(text);
          if (!adapter.read(json.Body)) return text;
          adapter.write(json.Body, slice, offset, total);
          return JSON.stringify(json);
        });
        served.set(folder, describe(rows, adapter));
        requestColors();
        scheduleSync();
        send(headers, parsed.viaHeader ? body : encoded);
      })
      .catch((error) => {
        console.warn("[cat-filter-for-owa] category sort failed, showing the original list.", error);
        passThrough();
      });
  }

  // Everything the page needs to decorate the list, by position in the sorted rows.
  function describe(rows, adapter) {
    const { labels, starts } = groupLabels(rows, UNCATEGORIZED);
    const realPositions = [];
    rows.forEach((row, pos) => row.copy || realPositions.push(pos));
    return {
      adapter,
      rows,
      order: rows.map((row) => adapter.rowIdOf(row.entry)),
      realOrder: realPositions.map((pos) => adapter.rowIdOf(rows[pos].entry)),
      realPositions,
      labels,
      starts,
    };
  }

  async function serve(action, adapter, request) {
    const body = request.Body;
    const cacheKey = JSON.stringify([action, adapter.folderOf(body), body.ViewFilter, body.FocusedViewFilter, body.CategoryFilter || null]);
    const fresh = !/^SeekToCondition/.test(body.Paging.__type || "") && !body.Paging.Offset;

    const load = async () => {
      const rows = sortByCategory(await fetchAll(action, adapter, request), adapter.categoriesOf);
      const list = { time: Date.now(), rows, entries: rows.filter((row) => !row.copy).map((row) => row.entry) };
      lists.set(cacheKey, list);
      return list;
    };

    let list = lists.get(cacheKey);
    let reloaded = false;
    if (!list || (fresh && Date.now() - list.time > REUSE_MS)) {
      list = await load();
      reloaded = true;
    }
    let start = pageStart(body.Paging, list.entries);
    if (start < 0 && !reloaded) {
      list = await load();
      start = pageStart(body.Paging, list.entries);
    }
    if (start < 0) start = list.entries.length;

    const slice = list.entries.slice(start, start + (body.Paging.MaxEntriesReturned || 25));
    return { slice, offset: start + slice.length, total: list.entries.length, rows: list.rows };
  }

  async function fetchAll(action, adapter, request) {
    const pageAt = async (offset) => {
      const query = JSON.parse(JSON.stringify(request));
      query.Body.Paging = { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: offset, MaxEntriesReturned: PAGE_SIZE };
      query.Body.SortOrder = adapter.dateSort;
      adapter.prepare(query.Body);
      const page = adapter.read(await post(action, query));
      if (!page) throw new Error(`${action}: unexpected server response`);
      return page;
    };

    // The server caps the page size: the first page gives the step and the total.
    const first = await pageAt(0);
    const step = first.entries.length;
    const offsets = [];
    if (step > 0) for (let offset = step; offset < first.total; offset += step) offsets.push(offset);

    const pages = new Array(offsets.length);
    let next = 0;
    const worker = async () => {
      while (next < offsets.length) {
        const index = next++;
        pages[index] = (await pageAt(offsets[index])).entries;
      }
    };
    await Promise.all(Array.from({ length: Math.min(PARALLEL_REQUESTS, offsets.length) }, worker));

    // A message arriving during the load shifts the pages: drop duplicates.
    const seen = new Set();
    return first.entries.concat(...pages).filter((entry) => {
      const key = adapter.keyOf(entry);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function readColors(configuration) {
    const master = configuration && configuration.MasterCategoryList;
    if (!master) return;
    colors = categoryColors(master.MasterList);
    colorsRequested = true;
    scheduleSync();
  }

  // The colors chosen for each category live in OWA's user configuration.
  function requestColors() {
    if (colorsRequested) return;
    colorsRequested = true;
    post("GetOwaUserConfiguration", {}, true)
      .then(readColors)
      .catch((error) => {
        colorsRequested = false;
        console.warn("[cat-filter-for-owa] could not read the category colors.", error);
      });
  }

  async function post(action, request, raw = false) {
    const canary = (/(?:^|;\s*)X-OWA-CANARY=([^;]+)/i.exec(document.cookie) || [])[1] || "";
    const response = await fetch(`/owa/service.svc?action=${action}`, {
      method: "POST",
      credentials: "same-origin",
      headers: {
        Action: action,
        "Content-Type": "application/json; charset=utf-8",
        "X-OWA-CANARY": canary,
        "X-Requested-With": "XMLHttpRequest",
      },
      body: JSON.stringify(request),
    });
    if (!response.ok) throw new Error(`${action}: HTTP ${response.status}`);
    const json = JSON.parse(await response.text());
    return raw ? json : json.Body;
  }

  function overrideResponse(xhr, transform) {
    let transformed = null;
    const text = () => {
      const raw = nativeResponseText.get.call(xhr);
      if (xhr.readyState !== 4) return raw;
      if (transformed === null) {
        try {
          transformed = transform(raw);
        } catch {
          transformed = raw;
        }
      }
      return transformed;
    };
    Object.defineProperty(xhr, "responseText", { configurable: true, get: text });
    Object.defineProperty(xhr, "response", {
      configurable: true,
      get() {
        if (xhr.responseType === "json") {
          try {
            return JSON.parse(text());
          } catch {
            return null;
          }
        }
        return !xhr.responseType || xhr.responseType === "text" ? text() : nativeResponse.get.call(xhr);
      },
    });
  }

  // --- Group headings in the list ---

  const listRows = () => [...document.querySelectorAll('[role="listbox"] [data-convid]')];
  const groupAttributes = ["data-cfo-group", "data-cfo-head-collapsed", "data-cfo-folded"];
  let headingColorSet = false;
  let swatchStyle = null;
  // Rows drawn by the page for the extra categories of a message, by sorted position.
  const copies = new Map();
  let copiesView = null;
  // Last rendering of each conversation row seen in OWA's list, used to draw its copies.
  const snapshots = new Map();

  function setAttribute(element, name, value) {
    if (value === null) {
      if (element.hasAttribute(name)) element.removeAttribute(name);
    } else if (element.getAttribute(name) !== value) {
      element.setAttribute(name, value);
    }
  }

  // Position of each displayed row in the sorted list, or null when the displayed
  // list is not a list sorted by the extension.
  function mapRows(rows) {
    const shown = rows.slice(0, 8).map((row) => row.getAttribute("data-convid"));
    for (const [folder, view] of served) {
      if (!enabledFolders.has(folder)) continue;
      const run = locateRun(shown, view.realOrder);
      if (!run) continue;
      const positions = rows.map(() => -1);
      for (let i = run.shownOffset, real = run.sortedStart; i < rows.length; i++, real++) {
        if (view.realOrder[real] !== rows[i].getAttribute("data-convid")) break;
        positions[i] = view.realPositions[real];
      }
      return { folder, view, positions };
    }
    return null;
  }

  // OWA draws category bars as <img> elements with no image, only a colored border and
  // background. Firefox does not render an <img> without a source, so the bars vanish:
  // a transparent image makes them show up, and changes nothing elsewhere.
  function showCategoryBars() {
    for (const bar of document.querySelectorAll('[role="listbox"] img.listViewCategoryIcon:not([src])')) {
      bar.setAttribute("src", BLANK_IMAGE);
    }
  }

  function syncList() {
    showCategoryBars();
    const rows = listRows();
    const mapped = mapRows(rows);
    displayed = mapped && mapped.folder;

    const root = document.documentElement;
    root.classList.toggle("cfo-by-category", !!mapped);
    if (mapped && !headingColorSet) {
      const themed = document.querySelector('[role="listbox"] [role="heading"] span') || document.querySelector(".ms-fcl-tp");
      if (themed) {
        root.style.setProperty("--cfo-heading-color", getComputedStyle(themed).color);
        headingColorSet = true;
      }
    }

    if (!mapped) {
      for (const row of rows) for (const name of groupAttributes) setAttribute(row, name, null);
      removeCopies();
      return;
    }

    const { folder, view, positions } = mapped;
    const folded = new Set(collapsed[folder] || []);
    const decorate = (element, pos) => {
      const label = pos < 0 ? null : view.labels[pos];
      const head = pos >= 0 && view.starts.has(pos);
      const isFolded = label !== null && folded.has(label);
      setAttribute(element, "data-cfo-group", head ? label : null);
      setAttribute(element, "data-cfo-head-collapsed", head && isFolded ? "" : null);
      setAttribute(element, "data-cfo-folded", !head && isFolded ? "" : null);
    };

    rows.forEach((row, i) => {
      decorate(row, positions[i]);
      if (positions[i] < 0) return;
      const id = row.getAttribute("data-convid");
      const html = row.innerHTML;
      const known = snapshots.get(id);
      if (!known || known.html !== html) snapshots.set(id, { html, node: row.cloneNode(true) });
    });
    placeCopies(rows, positions, view, decorate);
    updateSwatches(view);
  }

  // Copies of a row follow the row that precedes them in the sorted list.
  function placeCopies(rows, positions, view, decorate) {
    if (copiesView !== view) removeCopies();
    copiesView = view;
    const wanted = new Set();
    rows.forEach((row, i) => {
      let anchor = row;
      for (let pos = positions[i] + 1; positions[i] >= 0 && view.rows[pos] && view.rows[pos].copy; pos++) {
        const copy = copyAt(pos, view);
        if (anchor.nextSibling !== copy) anchor.after(copy);
        decorate(copy, pos);
        wanted.add(pos);
        anchor = copy;
      }
    });
    for (const [pos, copy] of copies) {
      if (wanted.has(pos)) continue;
      copy.remove();
      copies.delete(pos);
    }
  }

  function copyAt(pos, view) {
    const snapshot = snapshots.get(view.order[pos]);
    let copy = copies.get(pos);
    if (!copy) {
      copy = document.createElement("div");
      copy.setAttribute("data-cfo-copy", String(pos));
      copies.set(pos, copy);
    }
    if (copy.cfoSource === snapshot && copy.firstChild) return copy;
    copy.cfoSource = snapshot;
    if (snapshot) {
      const content = snapshot.node.cloneNode(true);
      for (const element of content.querySelectorAll("*")) {
        element.removeAttribute("id");
        element.removeAttribute("tabindex");
        element.removeAttribute("aria-selected");
        if (element.getAttribute("role") === "option") element.setAttribute("role", "presentation");
      }
      copy.replaceChildren(...content.childNodes);
    } else {
      // OWA has not rendered the original row yet: show who and what, until it does.
      const { from, subject } = view.adapter.summaryOf(view.rows[pos].entry);
      const summary = document.createElement("div");
      summary.className = "cfo-copy-summary";
      const sender = document.createElement("div");
      sender.textContent = from;
      const topic = document.createElement("div");
      topic.textContent = subject;
      summary.append(sender, topic);
      copy.replaceChildren(summary);
    }
    return copy;
  }

  function removeCopies() {
    for (const copy of copies.values()) copy.remove();
    copies.clear();
    copiesView = null;
  }

  // A copy opens the original row in the reading pane, or in its own window when
  // OWA has not rendered the original row.
  function openCopy(copy) {
    const view = copiesView;
    if (!view) return;
    const pos = Number(copy.getAttribute("data-cfo-copy"));
    const id = view.order[pos];
    const original = listRows().find((row) => row.getAttribute("data-convid") === id);
    const option = original && original.querySelector('[role="option"]');
    if (option) {
      // OWA scrolls to the row it selects; keep the list where the user clicked.
      const scroller = copy.closest(".scrollContainer");
      const top = scroller ? scroller.scrollTop : 0;
      option.click();
      if (scroller) {
        const restore = () => (scroller.scrollTop = top);
        requestAnimationFrame(restore);
        setTimeout(restore, 150);
      }
      return;
    }
    const itemId = view.adapter.itemIdOf(view.rows[pos].entry);
    if (itemId) {
      window.open(`/owa/?ItemID=${encodeURIComponent(itemId)}&exvsurl=1&viewmodel=ReadMessageItem`, "_blank", "popup,width=1000,height=800");
    }
  }

  function updateSwatches(view) {
    const rules = [];
    for (const pos of view.starts) {
      const label = view.labels[pos];
      const style = colors.get(label.toLowerCase());
      if (!style) continue;
      const { image, width } = swatches([style.color]);
      rules.push(
        `html.cfo-by-category [role="listbox"] [data-cfo-group=${JSON.stringify(label)}]::before ` +
          `{ --cfo-swatches: ${image}; --cfo-swatch-width: ${width}px; --cfo-tint: ${style.tint}; --cfo-accent: ${style.color}; --cfo-text: #212121; }`,
      );
    }
    const text = rules.join("\n");
    if (!swatchStyle) {
      swatchStyle = document.createElement("style");
      (document.head || document.documentElement).append(swatchStyle);
    }
    if (swatchStyle.textContent !== text) swatchStyle.textContent = text;
  }

  // A heading is drawn on top of its group's first row: clicking it folds the group
  // instead of selecting that row.
  function headingRow(event) {
    const row = event.target instanceof Element && event.target.closest("[data-cfo-group]");
    if (!row || !displayed || event.target !== row) return null;
    return event.clientY - row.getBoundingClientRect().top < HEADING_HEIGHT ? row : null;
  }

  // Mouse events on headings and copies are handled here; OWA never sees them.
  for (const type of ["mousedown", "mouseup", "click", "dblclick", "contextmenu"]) {
    document.addEventListener(
      type,
      (event) => {
        const heading = headingRow(event);
        const copy = !heading && event.target instanceof Element && event.target.closest("[data-cfo-copy]");
        if (!heading && !copy) return;
        event.preventDefault();
        event.stopImmediatePropagation();
        if (type !== "click") return;
        if (heading) {
          toggleCollapsed(displayed, heading.getAttribute("data-cfo-group"));
          scheduleSync();
        } else {
          openCopy(copy);
        }
      },
      true,
    );
  }

  // --- "Category" option in the "Sort by" menu ---

  const isShown = (el) => !!el && el.style.display !== "none" && el.offsetParent !== null;
  const popups = () => [...document.querySelectorAll('[role="menu"][ispopup="1"]')].filter(isShown);
  const isSortMenu = (menu) => !!menu.querySelector('button[role="menuitemcheckbox"] [class*="Icon--arrow"]');
  const sortMenu = () => popups().find(isSortMenu) || null;
  const nativeButtons = (menu) =>
    [...menu.querySelectorAll('button[role="menuitemcheckbox"]')].filter((b) => !b.closest("[data-cfo-item]"));
  const labelOf = (button) => (button ? button.textContent.trim() : "");
  const nativelyChecked = (button) => (button.getAttribute("data-cfo-native") || button.getAttribute("aria-checked")) === "true";

  function iconOf(button) {
    return button.querySelector('span[class*="Icon--"]');
  }

  function showChecked(button, checked) {
    const value = String(checked);
    if (button.getAttribute("aria-checked") !== value) button.setAttribute("aria-checked", value);
    button.classList.toggle("ms-fwt-sb", checked);
    const icon = iconOf(button);
    if (icon && (icon.style.display === "none") === checked) icon.style.display = checked ? "" : "none";
  }

  function syncMenu() {
    for (const menu of document.querySelectorAll('[role="menu"][ispopup="1"]')) {
      if (!isSortMenu(menu)) continue;
      const natives = nativeButtons(menu);
      if (natives.length === 0) continue;

      let item = menu.querySelector("[data-cfo-item]");
      if (!item) item = createItem(natives[natives.length - 1].closest('[role="menuitem"]'));
      const on = displayed !== null;
      showChecked(item.querySelector("button"), on);

      for (const button of natives) {
        if (on) {
          if (!button.hasAttribute("data-cfo-native")) button.setAttribute("data-cfo-native", button.getAttribute("aria-checked") || "false");
          showChecked(button, false);
        } else if (button.hasAttribute("data-cfo-native")) {
          const checked = button.getAttribute("data-cfo-native") === "true";
          button.removeAttribute("data-cfo-native");
          showChecked(button, checked);
        }
      }
    }
  }

  function createItem(template) {
    const item = template.cloneNode(true);
    item.setAttribute("data-cfo-item", "");
    item.setAttribute("tabindex", "-1");
    item.removeAttribute("aria-selected");
    const button = item.querySelector("button");
    button.removeAttribute("title");
    button.classList.remove("ms-bgc-tl");
    const label = [...item.querySelectorAll("span")].find((span) => span.children.length === 0 && span.textContent.trim());
    if (label) label.textContent = LABEL;
    const icon = iconOf(button);
    if (icon) icon.className = icon.className.replace(/Icon--arrow\w+/, "Icon--arrowUp");
    item.addEventListener("click", onCategoryClick, true);
    item.addEventListener("mouseenter", () => button.classList.add("ms-bgc-tl"));
    item.addEventListener("mouseleave", () => button.classList.remove("ms-bgc-tl"));
    template.after(item);
    return item;
  }

  function onCategoryClick(event) {
    event.preventDefault();
    event.stopImmediatePropagation();
    const menu = event.currentTarget.closest('[role="menu"][ispopup="1"]');
    const current = menu && nativeButtons(menu).find(nativelyChecked);
    const opener = menuOpener();
    if (displayed !== null || !current) {
      if (opener) opener.click();
      return;
    }
    // Clicking the current sort twice makes OWA reload the list without changing its setting;
    // the reloaded list is the displayed folder's, which then switches to category sort.
    pendingEnableUntil = Date.now() + PENDING_MS;
    clickAsExtension(current);
    toggleAgain(labelOf(current), opener);
  }

  // Leaving category sort: if the user picks the sort OWA was already applying,
  // OWA would flip its direction, so it is restored right away.
  document.addEventListener(
    "click",
    (event) => {
      const button = event.target instanceof Element && event.target.closest('button[role="menuitemcheckbox"]');
      const menu = button && !button.closest("[data-cfo-item]") && button.closest('[role="menu"][ispopup="1"]');
      if (ownClick || !menu || !isSortMenu(menu) || displayed === null) return;
      const sameSort = nativelyChecked(button);
      const opener = sameSort ? menuOpener() : null;
      setEnabled(displayed, false);
      scheduleSync();
      if (sameSort) toggleAgain(labelOf(button), opener);
    },
    true,
  );

  // "Filter" button that opened the menu, so it can be reopened.
  function menuOpener() {
    return (
      [...document.querySelectorAll('button[aria-haspopup="true"][aria-expanded="true"]')].find((b) => !b.closest('[ispopup="1"]')) || null
    );
  }

  // The "Sort by" submenu cannot be identified by its label (it depends on the language):
  // try each submenu until the one holding the sort options shows up.
  async function openSortMenu(opener) {
    opener.click();
    const root = await waitFor(() => popups().find((m) => !isSortMenu(m) && m.querySelector('button[aria-haspopup="true"]')));
    for (const parent of root.querySelectorAll('button[aria-haspopup="true"]')) {
      parent.click();
      const menu = await waitFor(sortMenu, 800).catch(() => null);
      if (menu) return menu;
    }
    throw new Error("\"Sort by\" menu not found");
  }

  async function toggleAgain(label, opener) {
    if (!opener) return;
    const root = document.documentElement;
    root.classList.add("cfo-busy");
    try {
      await waitFor(() => !sortMenu());
      const menu = await openSortMenu(opener);
      const button = nativeButtons(menu).find((b) => labelOf(b) === label);
      if (button) clickAsExtension(button);
      else opener.click();
    } catch (error) {
      console.warn("[cat-filter-for-owa] could not restore the sort direction.", error);
    } finally {
      root.classList.remove("cfo-busy");
    }
  }

  function clickAsExtension(element) {
    ownClick = true;
    try {
      element.click();
    } finally {
      ownClick = false;
    }
  }

  function waitFor(predicate, timeout = 4000) {
    return new Promise((resolve, reject) => {
      const started = Date.now();
      const tick = () => {
        const value = predicate();
        if (value) return resolve(value);
        if (Date.now() - started > timeout) return reject(new Error("timed out"));
        setTimeout(tick, 50);
      };
      tick();
    });
  }

  // --- Sync with OWA's DOM ---

  let syncScheduled = false;

  function scheduleSync() {
    if (syncScheduled) return;
    syncScheduled = true;
    requestAnimationFrame(() => {
      syncScheduled = false;
      syncList();
      syncMenu();
    });
  }

  new MutationObserver(scheduleSync).observe(document.documentElement, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ["style", "aria-checked"],
  });
})();
