"use strict";

// Pure category-sort logic, shared by the OWA page script and the tests.
(function (root) {
  const collator = new Intl.Collator("fr", { sensitivity: "base", numeric: true });

  const sortResult = (order, field) => ({
    __type: "SortResults:#Exchange",
    Order: order,
    Path: { __type: "PropertyUri:#Exchange", FieldURI: field },
  });

  function categoryNames(categories) {
    if (!Array.isArray(categories)) return [];
    return [...new Set(categories.map((c) => String(c).trim()).filter(Boolean))].sort(collator.compare);
  }

  // One row per category, as Outlook does: a message with two categories is listed in
  // both groups. OWA merges rows of the same conversation, so only the first row of a
  // message is handed to OWA; the others are marked as copies and drawn by the page.
  // Uncategorized messages first, then categories A to Z; the order received from the
  // server (newest first) is kept within each group.
  function sortByCategory(entries, categoriesOf) {
    const rows = [];
    entries.forEach((entry, index) => {
      const names = categoryNames(categoriesOf(entry));
      if (names.length === 0) rows.push({ entry, category: "", index, copy: false });
      names.forEach((category, n) => rows.push({ entry, category, index, copy: n > 0 }));
    });
    return rows
      .sort((a, b) => {
        if (!a.category !== !b.category) return a.category ? 1 : -1;
        return collator.compare(a.category, b.category) || a.index - b.index;
      })
      .map(({ entry, category, copy }) => ({ entry, category, copy }));
  }

  // Start position requested by OWA: a plain offset, or resuming from a given
  // item (inclusive), as the server does. -1 when the item is not found.
  function pageStart(paging, entries) {
    if (/^SeekToConditionPageView/.test(paging.__type || "")) {
      const condition = paging.Condition && paging.Condition.Item;
      const value = condition && condition.FieldURIOrConstant && condition.FieldURIOrConstant.Item && condition.FieldURIOrConstant.Item.Value;
      return entries.findIndex((entry) => entry.InstanceKey === value);
    }
    return paging.Offset || 0;
  }

  // Group label of each sorted row, and the positions where a new group starts.
  function groupLabels(rows, uncategorizedLabel) {
    const labels = [];
    const starts = new Set();
    rows.forEach((row, index) => {
      const previous = rows[index - 1];
      if (index === 0 || collator.compare(row.category, previous.category) !== 0) {
        starts.add(index);
        labels.push(row.category || uncategorizedLabel);
      } else {
        labels.push(labels[index - 1]);
      }
    });
    return { labels, starts };
  }

  // Outlook's 25 category colors (preset0 to preset24), as OWA draws them:
  // the solid color of the category bars and the light tint of its backgrounds.
  const CATEGORY_COLORS = [
    ["#F07D88", "#FADFE3"], ["#FF8C00", "#FDE5D3"], ["#FECB6F", "#EEE2D8"], ["#FFF100", "#FBF4CF"],
    ["#5FBE7D", "#DCF1DE"], ["#33BAB1", "#D4F5F2"], ["#A3B367", "#EAF0D3"], ["#55ABE5", "#DDF3FE"],
    ["#A895E2", "#E6E0FF"], ["#E48BB5", "#FBE4F1"], ["#B9C0CB", "#E2E6EC"], ["#4C596E", "#E2E6EC"],
    ["#ABABAB", "#EEEEEE"], ["#666666", "#EEEEEE"], ["#474747", "#EEEEEE"], ["#910A19", "#FADFE3"],
    ["#CE4B28", "#FDE5D3"], ["#996E36", "#EEE2D8"], ["#B0A923", "#FBF4CF"], ["#026802", "#DCF1DE"],
    ["#1C6367", "#D4F5F2"], ["#5C6A22", "#EAF0D3"], ["#254069", "#DDF3FE"], ["#562685", "#E6E0FF"],
    ["#80275D", "#FBE4F1"],
  ];

  // Color of each category from OWA's master category list, by lowercase name.
  function categoryColors(masterList) {
    const colors = new Map();
    for (const category of Array.isArray(masterList) ? masterList : []) {
      const preset = CATEGORY_COLORS[category && category.Color];
      if (preset && category.Name) colors.set(String(category.Name).trim().toLowerCase(), { color: preset[0], tint: preset[1] });
    }
    return colors;
  }

  // Category markers drawn like OWA's own: 5px wide bars, 3px apart.
  function swatches(colors) {
    const stops = colors.map((color, i) => `${color} ${i * 8}px ${i * 8 + 5}px, transparent ${i * 8 + 5}px ${i * 8 + 8}px`);
    return { image: `linear-gradient(to right, ${stops.join(", ")})`, width: colors.length * 8 - 3 };
  }

  // Where do the rows on screen sit in a list sorted by the extension? OWA only keeps
  // the rows around the scroll position, so they must form a run of consecutive rows
  // of the sorted list. A few leading rows may be skipped (a message that just arrived
  // at the top). Returns null when the displayed list is not that sorted list.
  function locateRun(shown, sorted, length = 6, slack = 2) {
    const k = Math.min(length, shown.length);
    if (k === 0) return null;
    for (let a = 0; a <= slack && a + k <= shown.length; a++) {
      for (let start = sorted.indexOf(shown[a]); start !== -1; start = sorted.indexOf(shown[a], start + 1)) {
        let same = true;
        for (let i = 1; i < k && same; i++) same = sorted[start + i] === shown[a + i];
        if (same) return { shownOffset: a, sortedStart: start };
      }
    }
    return null;
  }

  const ADAPTERS = {
    // "Conversations" view.
    FindConversation: {
      isList: (body) => !!body && body.ShapeName === "ConversationListView" && !body.QueryString && !body.SearchFolderId,
      folderOf: (body) => body.ParentFolderId && body.ParentFolderId.BaseFolderId && body.ParentFolderId.BaseFolderId.Id,
      dateSort: [sortResult("Descending", "ConversationLastDeliveryOrRenewTime"), sortResult("Descending", "ConversationLastDeliveryTime")],
      prepare: () => {},
      read(body) {
        if (!body || body.ResponseClass !== "Success") return null;
        return { entries: body.Conversations || [], total: body.TotalConversationsInView || 0 };
      },
      write(body, entries, offset, total) {
        body.Conversations = entries;
        body.IndexedOffset = offset;
        body.TotalConversationsInView = total;
      },
      categoriesOf: (entry) => entry.Categories,
      rowIdOf: (entry) => entry.ConversationId && entry.ConversationId.Id,
      keyOf: (entry) => entry.InstanceKey || (entry.ConversationId && entry.ConversationId.Id),
      itemIdOf: (entry) => entry.ItemIds && entry.ItemIds[0] && entry.ItemIds[0].Id,
      summaryOf: (entry) => ({ from: (entry.UniqueSenders || []).join("; "), subject: entry.ConversationTopic || "" }),
    },

    // "Messages" view: the shape OWA uses does not return categories, so they are requested too.
    FindItem: {
      isList: (body) => !!body && body.ShapeName === "MailListItem" && !body.QueryString && body.Traversal !== "Associated",
      folderOf: (body) => body.ParentFolderIds && body.ParentFolderIds[0] && body.ParentFolderIds[0].Id,
      dateSort: [sortResult("Descending", "ReceivedOrRenewTime"), sortResult("Descending", "DateTimeReceived")],
      prepare(body) {
        const shape = body.ItemShape || (body.ItemShape = { __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly" });
        const extra = shape.AdditionalProperties || (shape.AdditionalProperties = []);
        if (!extra.some((p) => p.FieldURI === "Categories")) extra.push({ __type: "PropertyUri:#Exchange", FieldURI: "Categories" });
      },
      read(body) {
        const message = body && body.ResponseMessages && body.ResponseMessages.Items && body.ResponseMessages.Items[0];
        if (!message || message.ResponseClass !== "Success" || !message.RootFolder) return null;
        return { entries: message.RootFolder.Items || [], total: message.RootFolder.TotalItemsInView || 0 };
      },
      write(body, entries, offset, total) {
        const rootFolder = body.ResponseMessages.Items[0].RootFolder;
        rootFolder.Items = entries;
        rootFolder.IndexedPagingOffset = offset;
        rootFolder.TotalItemsInView = total;
        rootFolder.IncludesLastItemInRange = offset >= total;
      },
      categoriesOf: (entry) => entry.Categories,
      rowIdOf: (entry) => entry.ConversationId && entry.ConversationId.Id,
      keyOf: (entry) => entry.InstanceKey || (entry.ItemId && entry.ItemId.Id),
      itemIdOf: (entry) => entry.ItemId && entry.ItemId.Id,
      summaryOf: (entry) => {
        const mailbox = (entry.From && entry.From.Mailbox) || (entry.Sender && entry.Sender.Mailbox) || {};
        return { from: mailbox.Name || mailbox.EmailAddress || "", subject: entry.Subject || "" };
      },
    },
  };

  const api = { ADAPTERS, categoryNames, sortByCategory, pageStart, groupLabels, categoryColors, swatches, locateRun };
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.__catFilterForOwaCore = api;
})(globalThis);
