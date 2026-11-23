"use strict";

// Runs in the OWA page context. Exchange refuses to sort by category, so for
// folders where this sort is chosen, OWA's list requests are intercepted: the
// whole folder is fetched and sorted here, then OWA receives the page it asked
// for and renders it as usual.
(function () {
  const core = window.__catFilterForOwaCore;
  if (!core || window.__catFilterForOwaLoaded) return;
  window.__catFilterForOwaLoaded = true;

  const {
    ADAPTERS,
    sortByCategory,
    pageStart,
    groupLabels,
    groupCounts,
    countLabel,
    categoryColors,
    categoriesChanged,
    parseColor,
    isDark,
    swatches,
    mapDisplayedRows,
  } = core;

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
  const REFRESH_DELAY_MS = 1500;
  const REFRESH_GAP_MS = 8000;
  const REFRESH_GAP_MAX_MS = 120000;
  const THEME_TTL_MS = 1000;
  const SETTLE_MS = 3000;
  const CALL = Symbol("cat-filter-for-owa");
  const GENERATION = Symbol("cat-filter-for-owa-generation");
  const SERVED_FOLDERS_KEPT = 3;
  const MISMATCH_STRIKES = 2;
  const READ_ONLY_ACTION = /^(Find|Get|Subscribe|Unsubscribe|Ping|Log|Sync)/;
  const SEARCH_START = /^(StartSearchSession|ExecuteSearch|PerformInstantSearch)$/;
  const SEARCH_END = /^(EndSearchSession|EndInstantSearchSession)$/;

  const enabledFolders = new Set(readStorage(FOLDERS_KEY, []));
  // Collapsed groups, per folder.
  const collapsed = readStorage(COLLAPSED_KEY, {});
  // Whole folders already sorted, reused for the following pages; loads in progress.
  const lists = new Map();
  const loading = new Map();
  // Last sorted list handed to OWA, per folder (the few most recent): row ids, groups and
  // categories by position.
  const served = new Map();
  // Category colors from OWA's master category list, by lowercase name.
  let colors = new Map();
  let colorsRequested = false;
  // Category-sorted folder currently displayed, inferred from the rows on screen.
  let displayed = null;
  // Activation requested from the menu: applies to the folder of the next list loaded.
  let pendingEnableUntil = 0;
  let ownClick = false;
  // Search results replace the folder's rows: no decoration until the search ends.
  let searching = false;

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
    // Sent again whenever another OWA tab changes them.
    document.addEventListener(PREFS_EVENT, (event) => {
      applyPrefs(event.detail);
      done();
      scheduleSync();
    });
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

  function toggleCollapsed(folder, key) {
    const keys = new Set(collapsed[folder] || []);
    if (keys.has(key)) keys.delete(key);
    else keys.add(key);
    if (keys.size) collapsed[folder] = [...keys];
    else delete collapsed[folder];
    savePrefs();
  }

  // --- List request interception ---

  const proto = XMLHttpRequest.prototype;
  const native = { open: proto.open, send: proto.send, setRequestHeader: proto.setRequestHeader, abort: proto.abort };
  const nativeResponseText = Object.getOwnPropertyDescriptor(proto, "responseText");
  const nativeResponse = Object.getOwnPropertyDescriptor(proto, "response");

  proto.open = function (method, url) {
    delete this.responseText;
    delete this.response;
    const action = (/[?&]action=(\w+)/.exec(String(url)) || [])[1];
    // Each open() and abort() starts a new generation: a list request held while the
    // folder loads is only sent if OWA has not cancelled or reused it meanwhile.
    this[GENERATION] = (this[GENERATION] || 0) + 1;
    this[CALL] = ADAPTERS[action] ? { action, headers: [], generation: this[GENERATION] } : null;
    if (action && (SEARCH_START.test(action) || SEARCH_END.test(action))) {
      searching = SEARCH_START.test(action);
      scheduleSync();
    }
    if (action === "UpdateMasterCategoryList") {
      // A category was added, renamed or recolored: read the colors again.
      this.addEventListener("load", () => {
        colorsRequested = false;
        mismatchStrikes.clear();
        requestColors();
      });
    }
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

  proto.abort = function () {
    this[GENERATION] = (this[GENERATION] || 0) + 1;
    return native.abort.apply(this, arguments);
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
      // OWA may have aborted or reused the request while the folder was loading.
      if (sent || xhr[GENERATION] !== call.generation || xhr.readyState !== XMLHttpRequest.OPENED) return;
      sent = true;
      for (const [name, value] of headers) native.setRequestHeader.call(xhr, name, value);
      native.send.call(xhr, payload);
    };
    const passThrough = () => send(call.headers, body);

    const adapter = ADAPTERS[call.action];
    const parsed = readRequest(call.headers, body);
    const requestBody = parsed && parsed.request && parsed.request.Body;
    const folder = requestBody && adapter.isList(requestBody) && adapter.folderOf(requestBody);
    if (!folder) {
      if (requestBody && (requestBody.QueryString || requestBody.SearchFolderId)) searching = true;
      return passThrough();
    }

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
        served.delete(folder);
        served.set(folder, describe(rows, adapter));
        for (const old of served.keys()) if (served.size > SERVED_FOLDERS_KEPT) served.delete(old);
        settle();
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
    const { labels, keys, starts } = groupLabels(rows, UNCATEGORIZED);
    const realPositions = [];
    rows.forEach((row, pos) => row.copy || realPositions.push(pos));
    const unread = new Map();
    for (const row of rows) unread.set(adapter.keyOf(row.entry), adapter.isUnread(row.entry));
    const copied = new Set(rows.filter((row) => row.copy).map((row) => adapter.keyOf(row.entry)));
    return {
      adapter,
      rows,
      realOrder: realPositions.map((pos) => adapter.rowIdOf(rows[pos].entry)),
      realPositions,
      labels,
      keys,
      starts,
      // Bumped whenever read states or dropped messages change, to recount only then.
      version: 0,
      counts: null,
      rest: null,
      // Read state by message, kept up to date from the rows OWA displays.
      unread,
      // Messages that also appear as copies under another category.
      copied,
      // Messages OWA has dropped from its list since (deleted, moved to another folder).
      removed: new Set(),
    };
  }

  async function serve(action, adapter, request) {
    const body = request.Body;
    const cacheKey = JSON.stringify([action, adapter.folderOf(body), body.ViewFilter, body.FocusedViewFilter, body.CategoryFilter || null]);
    const fresh = !/^SeekToCondition/.test(body.Paging.__type || "") && !body.Paging.Offset;

    const load = () => {
      if (!loading.has(cacheKey)) {
        const task = fetchAll(action, adapter, request)
          .then((entries) => {
            const rows = sortByCategory(entries, adapter.categoriesOf);
            const list = { time: Date.now(), rows, entries: rows.filter((row) => !row.copy).map((row) => row.entry) };
            lists.set(cacheKey, list);
            return list;
          })
          .finally(() => loading.delete(cacheKey));
        loading.set(cacheKey, task);
      }
      return loading.get(cacheKey);
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

  async function fetchAll(action, adapter, request, retry = true) {
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

    // A message arriving or leaving during the load shifts the pages: drop duplicates,
    // and load again once if some were skipped.
    const seen = new Set();
    const entries = first.entries.concat(...pages).filter((entry) => {
      const key = adapter.keyOf(entry);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    return entries.length < first.total && retry ? fetchAll(action, adapter, request, false) : entries;
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
  const groupAttributes = ["data-cfo-group", "data-cfo-key", "data-cfo-count", "data-cfo-head-collapsed", "data-cfo-folded"];
  const cssEscape = (value) => (window.CSS && CSS.escape ? CSS.escape(value) : JSON.stringify(String(value)).slice(1, -1));
  let themeStyle = null;
  // Rows drawn by the page for the extra categories of a message, by sorted position.
  const copies = new Map();
  let copiesView = null;
  // Last rendering of each message row seen in OWA's list, used to draw its copies, and
  // the rows on screen by message.
  const snapshots = new Map();
  let shownRows = new Map();
  // Messages whose bars kept disagreeing with their categories after reloads.
  const mismatchStrikes = new Map();
  let mismatched = new Set();
  // Element whose height OWA sets for the whole list, and the height OWA gave it.
  let fitted = null;

  function setAttribute(element, name, value) {
    if (value === null) {
      if (element.hasAttribute(name)) element.removeAttribute(name);
    } else if (element.getAttribute(name) !== value) {
      element.setAttribute(name, value);
    }
  }

  function mapRows(rows) {
    const shown = rows.map((row) => row.getAttribute("data-convid"));
    for (const [folder, view] of served) {
      if (!enabledFolders.has(folder)) continue;
      const mapped = mapDisplayedRows(shown, view);
      if (mapped) return { folder, view, ...mapped };
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

  const isRowUnread = (row) => !!row.querySelector('[role="option"] .ms-fwt-sb');

  function syncList() {
    showCategoryBars();
    const rows = listRows();
    const mapped = searching ? null : mapRows(rows);
    displayed = mapped && mapped.folder;
    document.documentElement.classList.toggle("cfo-by-category", !!mapped);

    if (!mapped) {
      for (const row of rows) for (const name of groupAttributes) setAttribute(row, name, null);
      removeCopies();
      unfitListHeight();
      return;
    }

    const { folder, view, positions } = mapped;
    const keyAt = (pos) => view.adapter.keyOf(view.rows[pos].entry);
    for (const pos of mapped.missing) {
      if (view.removed.has(keyAt(pos))) continue;
      view.removed.add(keyAt(pos));
      view.version += 1;
    }
    mismatched = new Set();
    shownRows = new Map();
    rows.forEach((row, i) => {
      if (positions[i] < 0) return;
      const key = keyAt(positions[i]);
      shownRows.set(key, row);
      const unread = isRowUnread(row);
      if (view.unread.get(key) !== unread) {
        view.unread.set(key, unread);
        view.version += 1;
      }
      if ((mismatchStrikes.get(key) || 0) < MISMATCH_STRIKES && rowCategoriesChanged(row, view.rows[positions[i]].entry, view.adapter)) {
        mismatched.add(key);
      }
      // Copies are drawn from the original row: keep its latest rendering.
      if (!view.copied.has(key)) return;
      const html = row.innerHTML;
      const known = snapshots.get(key);
      if (!known || known.html !== html) snapshots.set(key, { html, node: row.cloneNode(true) });
    });

    if (!view.counts || view.counts.version !== view.version) {
      const counts = groupCounts(
        view.rows,
        view.starts,
        (row) => !view.removed.has(view.adapter.keyOf(row.entry)),
        (row) => view.unread.get(view.adapter.keyOf(row.entry)),
      );
      view.counts = { version: view.version, counts };
    }
    const counts = view.counts.counts;
    const folded = new Set(collapsed[folder] || []);
    const decorate = (element, pos) => {
      const key = pos < 0 ? null : view.keys[pos];
      const head = pos >= 0 && view.starts.has(pos);
      const isFolded = key !== null && folded.has(key);
      setAttribute(element, "data-cfo-group", head ? view.labels[pos] : null);
      setAttribute(element, "data-cfo-key", head ? key : null);
      setAttribute(element, "data-cfo-count", head ? countLabel(counts.get(pos), FRENCH) : null);
      setAttribute(element, "data-cfo-head-collapsed", head && isFolded ? "" : null);
      setAttribute(element, "data-cfo-folded", !head && isFolded ? "" : null);
    };
    rows.forEach((row, i) => decorate(row, positions[i]));
    placeCopies(rows, positions, view, decorate);
    updateStyles(view);
    fitListHeight(rows, positions, view, folded);

    // A new message, or a conversation moved to the top by a reply, sits where the date
    // sort would put it; a message whose categories changed sits in its old group. In
    // both cases the list is reloaded to sort them in.
    if (mapped.unknown || mismatched.size) scheduleRefresh();
    else listInOrder = true;
  }

  // The colored bars OWA draws in a row show its categories as they are now. OWA keeps a
  // hidden bar in rows without categories, and folded rows show none: only visible bars
  // of visible rows count.
  function rowCategoriesChanged(row, entry, adapter) {
    const option = row.querySelector('[role="option"]');
    if (!colors.size || !option || !option.getClientRects().length) return false;
    const expected = [];
    for (const name of adapter.categoriesOf(entry) || []) {
      const style = colors.get(String(name).trim().toLowerCase());
      // A category without a color (or unknown here) draws a bar we cannot predict.
      if (!style) return false;
      expected.push(style.color);
    }
    const shown = [...row.querySelectorAll("img.listViewCategoryIcon")]
      .filter((bar) => bar.getClientRects().length > 0)
      .map((bar) => bar.style.backgroundColor);
    return categoriesChanged(expected, shown);
  }

  // Copies of a row follow the row that precedes them in the sorted list.
  function placeCopies(rows, positions, view, decorate) {
    if (copiesView !== view) {
      removeCopies();
      for (const key of snapshots.keys()) if (!view.copied.has(key)) snapshots.delete(key);
    }
    copiesView = view;
    const wanted = new Set();
    rows.forEach((row, i) => {
      let anchor = row;
      for (let pos = positions[i] + 1; positions[i] >= 0 && view.rows[pos] && view.rows[pos].copy; pos++) {
        if (view.removed.has(view.adapter.keyOf(view.rows[pos].entry))) continue;
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
    const snapshot = snapshots.get(view.adapter.keyOf(view.rows[pos].entry));
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
    const original = shownRows.get(view.adapter.keyOf(view.rows[pos].entry));
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

  // --- List height ---

  // OWA sizes the list for all of its rows, folded ones included, which leaves blank
  // space at the bottom. The height is worked out again from what is shown: the rows
  // on screen, plus an estimate for those OWA has not loaded yet.
  function fitListHeight(rows, positions, view, folded) {
    const content = rows[0] && rows[0].parentElement;
    const shifted = content && content.parentElement;
    const sized = shifted && shifted.parentElement;
    if (!sized || !/px$/.test(sized.style.height)) return;
    if (fitted && fitted.element !== sized) unfitListHeight();
    if (!fitted) fitted = { element: sized, owa: sized.style.height, ours: null };
    else if (sized.style.height !== fitted.ours) fitted.owa = sized.style.height;

    let last = -1;
    let visibleHeight = 0;
    let visibleRows = 0;
    rows.forEach((row, i) => {
      last = Math.max(last, positions[i]);
      if (positions[i] < 0 || row.hasAttribute("data-cfo-folded") || row.hasAttribute("data-cfo-head-collapsed")) return;
      visibleHeight += row.offsetHeight - (row.hasAttribute("data-cfo-group") ? HEADING_HEIGHT : 0);
      visibleRows += 1;
    });
    if (last < 0) return;
    const rowHeight = visibleRows ? visibleHeight / visibleRows : 80;

    const restKey = `${last}|${view.version}|${[...folded].join("\u0000")}`;
    if (!view.rest || view.rest.key !== restKey) {
      let headings = 0;
      let pending = 0;
      let pos = last + 1;
      while (pos < view.rows.length && view.rows[pos].copy) pos++;
      for (; pos < view.rows.length; pos++) {
        if (view.starts.has(pos)) headings += 1;
        const gone = view.removed.has(view.adapter.keyOf(view.rows[pos].entry));
        if (!gone && !folded.has(view.keys[pos])) pending += 1;
      }
      view.rest = { key: restKey, headings, pending };
    }
    const rest = view.rest.headings * HEADING_HEIGHT + view.rest.pending * rowHeight;
    const height = `${Math.ceil((parseFloat(shifted.style.top) || 0) + content.offsetHeight + rest)}px`;
    if (sized.style.height !== height) sized.style.setProperty("height", height, "important");
    fitted.ours = sized.style.height;
  }

  function unfitListHeight() {
    if (fitted && fitted.element.isConnected && fitted.element.style.height === fitted.ours) {
      fitted.element.style.height = fitted.owa;
    }
    fitted = null;
  }

  // --- Colors ---

  // The headings take their colors from the list as it is drawn, so they follow the
  // page's theme, including dark themes applied by other extensions such as Dark Reader.
  let themeCache = null;

  function readTheme() {
    if (themeCache && Date.now() - themeCache.time < THEME_TTL_MS) return themeCache.theme;
    const option = document.querySelector('[role="listbox"] [data-convid] [role="option"]');
    if (!option) return null;
    let surface = "rgb(255, 255, 255)";
    for (let element = option; element && element !== document.documentElement; element = element.parentElement) {
      const background = parseColor(getComputedStyle(element).backgroundColor);
      if (background && background[3] > 0) {
        surface = getComputedStyle(element).backgroundColor;
        break;
      }
    }
    const heading = document.querySelector('[role="listbox"] [role="heading"] span') || document.querySelector(".ms-fcl-tp");
    const theme = {
      surface,
      ink: getComputedStyle(option).color,
      accent: heading ? getComputedStyle(heading).color : "#0078d7",
      dark: isDark(surface),
    };
    themeCache = { time: Date.now(), theme };
    return theme;
  }

  function chevron(path, fill) {
    const svg = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 8 8'><path d='${path}' fill='${fill}'/></svg>`;
    return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
  }

  function updateStyles(view) {
    const theme = readTheme();
    if (!theme) return;
    const muted = theme.dark ? theme.ink : "#666666";
    const rules = [
      `html.cfo-by-category { --cfo-surface: ${theme.surface}; --cfo-heading-color: ${theme.accent}; ` +
        `--cfo-line: ${theme.dark ? `color-mix(in srgb, ${theme.ink} 20%, ${theme.surface})` : "#eaeaea"}; ` +
        `--cfo-chevron-open: ${chevron("M0 2h8L4 6.5z", muted)}; --cfo-chevron-closed: ${chevron("M2 0v8l4.5-4z", muted)}; }`,
    ];
    for (const pos of view.starts) {
      const style = view.rows[pos].category && colors.get(view.rows[pos].category.trim().toLowerCase());
      if (!style) continue;
      const { image, width } = swatches([style.color]);
      const tint = theme.dark ? `color-mix(in srgb, ${style.color} 30%, ${theme.surface})` : style.tint;
      rules.push(
        `html.cfo-by-category [role="listbox"] [data-cfo-key="${cssEscape(view.keys[pos])}"]::before ` +
          `{ --cfo-swatches: ${image}; --cfo-swatch-width: ${width}px; --cfo-tint: ${tint}; ` +
          `--cfo-accent: ${style.color}; --cfo-text: ${theme.dark ? theme.ink : "#212121"}; }`,
      );
    }
    const text = rules.join("\n");
    if (!themeStyle) {
      themeStyle = document.createElement("style");
      (document.head || document.documentElement).append(themeStyle);
    }
    if (themeStyle.textContent !== text) themeStyle.textContent = text;
  }

  // --- Reloading the list ---

  let refreshTimer = 0;
  let refreshing = false;
  let lastRefresh = 0;
  let settledAt = 0;
  // Reloads in a row that left the list still out of order: each one doubles the wait.
  let refreshFailures = 0;
  let listInOrder = true;

  // A list just handed to OWA takes a moment to replace the old rows on screen.
  function settle() {
    settledAt = Date.now() + SETTLE_MS;
    setTimeout(scheduleSync, SETTLE_MS + 50);
  }

  function scheduleRefresh() {
    if (refreshTimer || refreshing || Date.now() < settledAt) return;
    const gap = Math.min(REFRESH_GAP_MS * 2 ** refreshFailures, REFRESH_GAP_MAX_MS);
    const delay = Math.max(REFRESH_DELAY_MS, lastRefresh + gap - Date.now());
    refreshTimer = setTimeout(refresh, delay);
  }

  // Never while the user types (a reply, a search) or has a menu open.
  function canRefresh() {
    const active = document.activeElement;
    // OWA's message editor lives in an iframe: focus there means the user is writing.
    const typing = !!active && (active.isContentEditable || /^(INPUT|TEXTAREA|SELECT|IFRAME)$/.test(active.tagName));
    return !typing && document.visibilityState === "visible" && popups().length === 0;
  }

  async function refresh() {
    refreshTimer = 0;
    if (!displayed) return;
    if (!canRefresh()) {
      scheduleRefresh();
      return;
    }
    refreshing = true;
    for (const key of mismatched) mismatchStrikes.set(key, (mismatchStrikes.get(key) || 0) + 1);
    refreshFailures = listInOrder ? 0 : refreshFailures + 1;
    listInOrder = false;
    lastRefresh = Date.now();
    lists.clear();
    const scroller = () => document.querySelector('[role="listbox"] .scrollContainer');
    const top = scroller() ? scroller().scrollTop : 0;
    const focused = document.activeElement;
    try {
      await reloadList();
      // Give the focus back (OWA's menus took it), let OWA render the reloaded list, then
      // scroll back to where the user was.
      if (focused && focused.isConnected && focused !== document.body) focused.focus({ preventScroll: true });
      await new Promise((resolve) => setTimeout(resolve, SETTLE_MS / 2));
      await waitFor(() => displayed, 6000).catch(() => null);
      if (scroller()) scroller().scrollTop = top;
    } finally {
      refreshing = false;
    }
  }

  // Clicking the current sort twice makes OWA reload the list without changing its setting.
  async function reloadList() {
    const opener = listOpener();
    if (!opener) return;
    const root = document.documentElement;
    let label = null;
    root.classList.add("cfo-busy");
    try {
      const menu = await openSortMenu(opener);
      const current = nativeButtons(menu).find(nativelyChecked);
      if (current) {
        label = labelOf(current);
        clickAsExtension(current);
      } else {
        opener.click();
      }
    } catch (error) {
      console.warn("[cat-filter-for-owa] could not reload the list.", error);
      if (popups().length) opener.click();
    } finally {
      root.classList.remove("cfo-busy");
    }
    if (label) await toggleAgain(label, opener);
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
          toggleCollapsed(displayed, heading.getAttribute("data-cfo-key"));
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
      [...document.querySelectorAll('button[aria-haspopup="true"][aria-expanded="true"]')].find((b) => !b.closest('[ispopup="1"]')) ||
      listOpener()
    );
  }

  // The list's "Filter" button: the only menu button in the list's header.
  function listOpener() {
    const list = document.querySelector('[role="listbox"]');
    for (let element = list, depth = 0; element && depth < 8; element = element.parentElement, depth++) {
      const button = [...element.querySelectorAll('button[aria-haspopup="true"]')].find(
        (b) => !list.contains(b) && !b.closest('[ispopup="1"]'),
      );
      if (button) return button;
    }
    return null;
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
      if (popups().length) opener.click();
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

  const RELEVANT = '[role="listbox"], [ispopup="1"]';
  const isRelevant = (node) =>
    node instanceof Element && (!!node.closest(RELEVANT) || !!node.querySelector(RELEVANT) || node === document.documentElement);

  new MutationObserver((records) => {
    for (const record of records) {
      const target = record.target instanceof Element ? record.target : record.target.parentElement;
      if ((target && target.closest(RELEVANT)) || [...record.addedNodes].some(isRelevant) || [...record.removedNodes].some(isRelevant)) {
        scheduleSync();
        return;
      }
    }
  }).observe(document.documentElement, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ["style", "aria-checked"],
  });
})();
