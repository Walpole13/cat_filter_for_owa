"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = (file) => fs.readFileSync(path.join(__dirname, "..", "src", file), "utf8");
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const conversation = (id, categories) => ({
  ConversationId: { Id: `conv-${id}` },
  InstanceKey: `key-${id}`,
  ...(categories ? { Categories: categories } : {}),
});
const FOLDER = [conversation(1, ["Urgent"]), conversation(2), conversation(3, ["Admin"])];

const listRequest = {
  __type: "FindConversationJsonRequest:#Exchange",
  Body: {
    __type: "FindConversationRequest:#Exchange",
    ParentFolderId: { BaseFolderId: { Id: "inbox" } },
    ShapeName: "ConversationListView",
    Paging: { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: 25 },
    SortOrder: [],
  },
};

// Minimal page: enough DOM, XHR and storage for page.js to run against a fake OWA server.
function createPage({ stored = {}, answer }) {
  const document = new EventTarget();
  Object.assign(document, {
    cookie: "X-OWA-CANARY=canary",
    documentElement: { lang: "fr-FR", classList: { toggle() {} }, style: { setProperty() {} } },
    querySelectorAll: () => [],
    querySelector: () => null,
  });

  class FakeXHR {
    constructor() {
      this.readyState = 0;
      this.responseType = "";
    }
    open(method, url) {
      this.url = url;
    }
    setRequestHeader(name, value) {
      (this.sentHeaders ||= []).push([name, value]);
    }
    send(body) {
      this.sentBody = body;
      sent.push(this);
    }
    get responseText() {
      return this.raw || "";
    }
    get response() {
      return this.responseText;
    }
  }

  const sent = [];
  const fetches = [];
  const saves = [];
  const storage = new Map(Object.entries(stored));

  const window = {
    document,
    navigator: { language: "fr-FR" },
    XMLHttpRequest: FakeXHR,
    CustomEvent,
    EventTarget,
    Element: class {},
    MutationObserver: class {
      observe() {}
    },
    requestAnimationFrame: (fn) => setTimeout(fn, 0),
    setTimeout,
    clearTimeout,
    console,
    localStorage: {
      getItem: (key) => (storage.has(key) ? storage.get(key) : null),
      setItem: (key, value) => storage.set(key, String(value)),
    },
    async fetch(url, init) {
      fetches.push(JSON.parse(init.body));
      return {
        ok: true,
        text: async () =>
          JSON.stringify({ Body: { ResponseClass: "Success", Conversations: FOLDER, TotalConversationsInView: FOLDER.length } }),
      };
    },
  };
  window.window = window;
  window.globalThis = window;

  // Stand-in for content.js: answers the page script's request for preferences.
  document.addEventListener("cat-filter-for-owa:hello", () => {
    if (answer !== undefined) {
      setTimeout(() => document.dispatchEvent(new CustomEvent("cat-filter-for-owa:prefs", { detail: JSON.stringify(answer) })), 20);
    }
  });
  document.addEventListener("cat-filter-for-owa:save", (event) => saves.push(JSON.parse(event.detail)));

  const context = vm.createContext(window);
  vm.runInContext(source("sort-core.js"), context);
  vm.runInContext(source("page.js"), context);

  function requestList() {
    const xhr = new FakeXHR();
    xhr.open("POST", "/owa/service.svc?action=FindConversation&ID=-1");
    xhr.setRequestHeader("Action", "FindConversation");
    xhr.setRequestHeader("X-OWA-UrlPostData", encodeURIComponent(JSON.stringify(listRequest)));
    xhr.send("");
    return xhr;
  }

  // What OWA reads once the (minimal) request sent to the server comes back.
  function respond(xhr) {
    xhr.raw = JSON.stringify({ Body: { ResponseClass: "Success", Conversations: [FOLDER[0]], TotalConversationsInView: 3 } });
    xhr.readyState = 4;
    return JSON.parse(xhr.responseText).Body.Conversations.map((c) => c.ConversationId.Id);
  }

  return { sent, fetches, saves, storage, requestList, respond };
}

test("the first list waits for the saved preferences, then opens sorted by category", async () => {
  const page = createPage({ answer: { folders: ["inbox"], collapsed: {} } });
  const xhr = page.requestList();
  assert.equal(page.sent.length, 0, "OWA's request is held until the preferences arrive");

  await wait(100);
  assert.equal(page.sent.length, 1);
  assert.ok(page.fetches.length >= 1, "the whole folder was fetched");
  assert.deepEqual(page.respond(xhr), ["conv-2", "conv-3", "conv-1"], "uncategorized first, then A to Z");
  assert.equal(page.storage.get("cat-filter-for-owa:folders"), JSON.stringify(["inbox"]), "page storage kept in sync");
});

test("without saved preferences, the list is left to OWA", async () => {
  const page = createPage({ answer: null });
  const xhr = page.requestList();
  await wait(100);
  assert.equal(page.sent.length, 1);
  assert.equal(page.fetches.length, 0);
  assert.equal(xhr.sentHeaders.find(([name]) => name === "X-OWA-UrlPostData")[1], encodeURIComponent(JSON.stringify(listRequest)));
});

test("preferences kept by an earlier version in the page are carried over", async () => {
  const page = createPage({ stored: { "cat-filter-for-owa:folders": JSON.stringify(["inbox"]) }, answer: null });
  await wait(100);
  assert.deepEqual(page.saves, [{ folders: ["inbox"], collapsed: {} }]);
});

test("if the extension never answers, the page's own preferences are used", async () => {
  const page = createPage({ stored: { "cat-filter-for-owa:folders": JSON.stringify(["inbox"]) } });
  page.requestList();
  await wait(200);
  assert.equal(page.sent.length, 0, "still waiting");
  await wait(1500);
  assert.equal(page.sent.length, 1);
  assert.ok(page.fetches.length >= 1, "sorted from the page's own preferences");
});
