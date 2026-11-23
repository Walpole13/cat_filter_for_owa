"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const {
  ADAPTERS,
  categoryNames,
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
} = require("../src/sort-core.js");

const conversation = (id, categories) => ({
  ConversationId: { Id: `conv-${id}` },
  InstanceKey: `key-${id}`,
  ...(categories ? { Categories: categories } : {}),
});

// Server order: newest first.
const inbox = [
  conversation(1),
  conversation(2, ["Urgent"]),
  conversation(3, ["développement"]),
  conversation(4, ["Urgent"]),
  conversation(5, ["Urgent", "AJ en attente"]),
  conversation(6, ["Développement"]),
];

const ids = (entries) => entries.map((e) => e.ConversationId.Id);

test("categoryNames sorts, trims and dedupes category names", () => {
  assert.deepEqual(categoryNames([" Urgent ", "AJ en attente", "Urgent", ""]), ["AJ en attente", "Urgent"]);
  assert.deepEqual(categoryNames(undefined), []);
});

test("sortByCategory lists a message once per category, uncategorized first, then A to Z", () => {
  const rows = sortByCategory(inbox, ADAPTERS.FindConversation.categoriesOf);
  assert.deepEqual(
    rows.map((row) => `${row.category || "-"}:${row.entry.ConversationId.Id}${row.copy ? "*" : ""}`),
    ["-:conv-1", "AJ en attente:conv-5", "développement:conv-3", "Développement:conv-6", "Urgent:conv-2", "Urgent:conv-4", "Urgent:conv-5*"],
  );
  assert.equal(rows.filter((row) => !row.copy).length, inbox.length, "OWA gets each message once");
});

test("pageStart handles offsets and seek-to-item paging (inclusive)", () => {
  assert.equal(pageStart({ __type: "IndexedPageView:#Exchange", Offset: 25 }, inbox), 25);
  const seek = (value) => ({
    __type: "SeekToConditionPageView:#Exchange",
    Condition: { Item: { FieldURIOrConstant: { Item: { Value: value } } } },
  });
  assert.equal(pageStart(seek("key-3"), inbox), 2);
  assert.equal(pageStart(seek("missing"), inbox), -1);
});

test("groupLabels finds group starts, case-insensitively", () => {
  const rows = sortByCategory(inbox, ADAPTERS.FindConversation.categoriesOf);
  const { labels, keys, starts } = groupLabels(rows, "No category");
  assert.deepEqual([...starts], [0, 1, 2, 4]);
  assert.deepEqual(labels, ["No category", "AJ en attente", "développement", "développement", "Urgent", "Urgent", "Urgent"]);
  assert.deepEqual(keys, ["", "aj en attente", "développement", "développement", "urgent", "urgent", "urgent"]);
});

test("categories differing only by an accent stay apart; a category named like the uncategorized group too", () => {
  const entries = [conversation(1, ["Relance"]), conversation(2, ["Relancé"]), conversation(3, ["No category"]), conversation(4)];
  const rows = sortByCategory(entries, ADAPTERS.FindConversation.categoriesOf);
  const { labels, keys, starts } = groupLabels(rows, "No category");
  assert.equal(starts.size, 4);
  assert.deepEqual(labels, ["No category", "No category", "Relance", "Relancé"]);
  assert.notEqual(keys[0], keys[1], "the real category and the uncategorized group have different keys");
});

test("categoryColors maps OWA's master list to Outlook preset colors", () => {
  const colors = categoryColors([
    { Name: "URGENT", Color: 15 },
    { Name: "AJ EN ATTENTE", Color: 5 },
    { Name: "Sans couleur", Color: -1 },
  ]);
  assert.deepEqual(colors.get("urgent"), { color: "#910A19", tint: "#FADFE3" });
  assert.deepEqual(colors.get("aj en attente"), { color: "#33BAB1", tint: "#D4F5F2" });
  assert.equal(colors.has("sans couleur"), false);
  assert.equal(categoryColors(undefined).size, 0);
});

test("swatches draws one bar per color", () => {
  assert.deepEqual(swatches(["red"]), { image: "linear-gradient(to right, red 0px 5px, transparent 5px 8px)", width: 5 });
  assert.equal(swatches(["red", "blue"]).width, 13);
});

test("FindConversation adapter reads and rewrites the response body", () => {
  const body = { ResponseClass: "Success", Conversations: inbox.slice(0, 2), IndexedOffset: 2, TotalConversationsInView: 6 };
  assert.deepEqual(ADAPTERS.FindConversation.read(body), { entries: inbox.slice(0, 2), total: 6 });
  ADAPTERS.FindConversation.write(body, inbox.slice(2, 4), 4, 6);
  assert.deepEqual(body, { ResponseClass: "Success", Conversations: inbox.slice(2, 4), IndexedOffset: 4, TotalConversationsInView: 6 });
  assert.equal(ADAPTERS.FindConversation.read({ ResponseClass: "Error" }), null);
});

test("FindItem adapter requests categories and rewrites the root folder", () => {
  const request = { ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly" } };
  ADAPTERS.FindItem.prepare(request);
  ADAPTERS.FindItem.prepare(request);
  assert.deepEqual(request.ItemShape.AdditionalProperties.map((p) => p.FieldURI), ["Categories"]);

  const body = { ResponseMessages: { Items: [{ ResponseClass: "Success", RootFolder: { Items: [], TotalItemsInView: 3, IndexedPagingOffset: 0 } }] } };
  ADAPTERS.FindItem.write(body, [{ ItemId: { Id: "x" } }], 3, 3);
  const root = body.ResponseMessages.Items[0].RootFolder;
  assert.equal(root.IndexedPagingOffset, 3);
  assert.equal(root.IncludesLastItemInRange, true);
});

test("list requests are recognized, searches are not", () => {
  assert.equal(ADAPTERS.FindConversation.isList({ ShapeName: "ConversationListView" }), true);
  assert.equal(ADAPTERS.FindConversation.isList({ ShapeName: "ConversationListView", QueryString: "x" }), false);
  assert.equal(ADAPTERS.FindItem.isList({ ShapeName: "MailListItem", Traversal: "Shallow" }), true);
  assert.equal(ADAPTERS.FindConversation.folderOf({ ParentFolderId: { BaseFolderId: { Id: "inbox" } } }), "inbox");
});

test("groupCounts counts messages and unread ones per group, skipping removed ones", () => {
  const rows = sortByCategory(inbox, ADAPTERS.FindConversation.categoriesOf);
  const { starts } = groupLabels(rows, "No category");
  const unread = new Set(["conv-2", "conv-5"]);
  const counts = groupCounts(rows, starts, (row) => row.entry.ConversationId.Id !== "conv-4", (row) => unread.has(row.entry.ConversationId.Id));
  assert.deepEqual([...counts], [
    [0, { count: 1, unread: 0 }],
    [1, { count: 1, unread: 1 }],
    [2, { count: 2, unread: 0 }],
    [4, { count: 2, unread: 2 }],
  ]);
});

test("countLabel shows the message count and, when any, the unread count", () => {
  assert.equal(countLabel({ count: 12, unread: 0 }, true), " (12)");
  assert.equal(countLabel({ count: 12, unread: 1 }, true), " (12 · 1 non lu)");
  assert.equal(countLabel({ count: 12, unread: 3 }, true), " (12 · 3 non lus)");
  assert.equal(countLabel({ count: 12, unread: 3 }, false), " (12 · 3 unread)");
});

test("isDark tells dark list backgrounds from light ones", () => {
  assert.deepEqual(parseColor("rgba(0, 0, 0, 0)"), [0, 0, 0, 0]);
  assert.equal(parseColor("transparent"), null);
  assert.equal(isDark("rgb(255, 255, 255)"), false);
  assert.equal(isDark("rgb(24, 26, 27)"), true);
});

test("mapDisplayedRows follows OWA's rows through new, removed and scrolled-away rows", () => {
  const view = { realOrder: ["a", "b", "c", "d", "e", "f", "g", "h"], realPositions: [0, 1, 2, 4, 5, 6, 7, 8] };
  assert.deepEqual(mapDisplayedRows(["a", "b", "c", "d", "e", "f"], view), { positions: [0, 1, 2, 4, 5, 6], missing: [], unknown: false });
  assert.deepEqual(
    mapDisplayedRows(["new", "a", "b", "c", "d", "e", "f"], view),
    { positions: [-1, 0, 1, 2, 4, 5, 6], missing: [], unknown: true },
    "new message at the top",
  );
  assert.deepEqual(
    mapDisplayedRows(["a", "b", "c", "e", "f", "g", "h"], view),
    { positions: [0, 1, 2, 5, 6, 7, 8], missing: [4], unknown: false },
    "a message deleted from the list",
  );
  assert.deepEqual(
    mapDisplayedRows(["e", "a", "b", "c", "d", "f", "g"], view),
    { positions: [-1, 0, 1, 2, 4, 6, 7], missing: [5], unknown: true },
    "a conversation moved to the top by a reply",
  );
  assert.equal(mapDisplayedRows(["x", "y", "z"], view), null, "another folder");
  assert.equal(mapDisplayedRows(["b", "c"], view), null, "a small folder sharing two conversations");
  const tiny = { realOrder: ["b", "c", "d"], realPositions: [0, 1, 2] };
  assert.equal(mapDisplayedRows(["x", "b", "c", "d", "y", "z", "w", "v"], tiny), null, "a long list containing a small folder's rows");
  const small = { realOrder: ["a", "b"], realPositions: [0, 1] };
  assert.deepEqual(mapDisplayedRows(["a", "b"], small), { positions: [0, 1], missing: [], unknown: false }, "a small sorted folder");
});

test("categoriesChanged spots a category added, changed or removed in a row", () => {
  const urgent = "#910A19";
  assert.equal(categoriesChanged([urgent], ["rgb(145, 10, 25)"]), false);
  assert.equal(categoriesChanged([], ["rgb(145, 10, 25)"]), true, "category added");
  assert.equal(categoriesChanged([urgent], ["rgb(51, 186, 177)"]), true, "category changed");
  assert.equal(categoriesChanged([urgent], []), true, "categories removed");
  assert.equal(categoriesChanged([urgent, "#33BAB1", "#F07D88", "#FF8C00"], ["rgb(145, 10, 25)"]), false, "not every bar drawn");
  assert.equal(categoriesChanged([], []), false);
});
