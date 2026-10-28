"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { ADAPTERS, categoryNames, sortByCategory, pageStart, groupLabels, categoryColors, swatches, locateRun } = require("../src/sort-core.js");

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
  const { labels, starts } = groupLabels(rows, "No category");
  assert.deepEqual([...starts], [0, 1, 2, 4]);
  assert.deepEqual(labels, ["No category", "AJ en attente", "développement", "développement", "Urgent", "Urgent", "Urgent"]);
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

test("locateRun finds the displayed rows in the sorted list, wherever the list is scrolled", () => {
  const order = ["a", "b", "c", "d", "e", "f", "g", "h"];
  assert.deepEqual(locateRun(["a", "b", "c", "d", "e", "f"], order), { shownOffset: 0, sortedStart: 0 });
  assert.deepEqual(locateRun(["c", "d", "e", "f", "g", "h"], order), { shownOffset: 0, sortedStart: 2 }, "top rows not rendered");
  assert.deepEqual(locateRun(["new", "a", "b", "c", "d", "e", "f"], order), { shownOffset: 1, sortedStart: 0 }, "message added at the top");
  assert.deepEqual(locateRun(["a", "b"], order), { shownOffset: 0, sortedStart: 0 }, "short list");
  assert.equal(locateRun(["c", "a", "b", "d", "e", "f"], order), null, "date order of another folder");
  assert.equal(locateRun([], order), null);
  const repeated = ["x", "a", "y", "a", "b", "c"];
  assert.deepEqual(locateRun(["a", "b", "c"], repeated), { shownOffset: 0, sortedStart: 3 }, "message listed in two groups");
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
