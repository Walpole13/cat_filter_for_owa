# Cat Filter for OWA

A Firefox extension that adds **sorting by category** to Outlook Web App (on-premises Exchange 2016/2019), which OWA does not offer.

It blends into OWA: a **Category** entry appears in the **Filter › Sort by** menu, after Size. Once chosen for a folder:

- the message list is grouped by category: uncategorized messages first, then categories A to Z;
- inside each category, messages stay newest first;
- OWA's date headings ("Yesterday", "Last week"…) are replaced by category headings in the category's color (the color chosen in Outlook);
- clicking a heading folds or unfolds its category;
- a message with several categories is listed under each of them;
- everything else is OWA's own behaviour: reading pane, selection, scrolling, filters (Unread, Flagged…), Conversations or Messages view.

The choice is remembered per folder, as are folded categories, from one session to the next: they are kept in the extension's own storage, which survives signing out of OWA and clearing the browser's site data. Picking another sort option (Date, From…) goes back to OWA's normal behaviour.

## How it works

Exchange refuses to sort by category on the server side. For folders where this sort is enabled, the extension intercepts OWA's list requests, fetches the whole folder (read-only, using the session already open in the tab), sorts it, and hands OWA the page it asked for. OWA then renders it as usual.

- **Read-only**: no message is modified, moved or marked as read.
- No data leaves the browser: the only requests go to the Exchange server itself, and nothing is written to disk except the folders sorted by category and the folded categories (in the extension's storage, and the page's as a fallback).
- Large folders take a little longer to open in this mode, since the whole folder has to be loaded first.
- OWA shows each conversation only once, so when a message has several categories, OWA's own row sits in its first category and the other groups show a copy of it. Clicking a copy opens the message in the reading pane.
- New messages arriving while the folder is open appear at the top until the list is reloaded (by switching folders, for instance).
- In Firefox, OWA's colored category bars in the message rows do not show up on their own; the extension makes them visible, in every folder.

The new Outlook on the web (outlook.office.com) is not supported.

## Installation

### Quick try (until Firefox restarts)

A temporary add-on is removed when Firefox closes, together with its saved preferences: use the permanent installation below for everyday use.

1. Download the repository (**Code → Download ZIP**) and unzip it.
2. Open `about:debugging#/runtime/this-firefox` in Firefox.
3. Click **Load Temporary Add-on…** and select the `manifest.json` file.
4. Reload the OWA tab.

### Permanent installation

Release builds of Firefox only keep extensions signed by Mozilla. Signing is free and automatic for an unlisted extension:

1. Create an account on [addons.mozilla.org](https://addons.mozilla.org/developers/).
2. Generate API credentials on the [API keys](https://addons.mozilla.org/developers/addon/api/key/) page.
3. From the project folder:

   ```bash
   npx web-ext sign --channel=unlisted --api-key=<JWT issuer> --api-secret=<JWT secret>
   ```

4. Open the generated `.xpi` file from `web-ext-artifacts/` in Firefox to install it.

> On a managed workstation, a group policy may forbid installing add-ons. Check with your IT department in that case.

## Troubleshooting

If the category sort cannot be applied, OWA simply shows its normal list, and the browser console (`Ctrl+Shift+K`) shows a message prefixed with `[cat-filter-for-owa]`.

- **The Category entry does not show up**: make sure the page address contains `/owa/` and that the extension is enabled in `about:addons`, then reload the tab.

## Development

```bash
npm test        # unit tests (Node 18+)
npm run lint    # extension validation
npm run build   # package into web-ext-artifacts/
```

- `src/content.js`: loads the scripts into the page context;
- `src/page.js`: request interception, "Sort by" menu entry and group headings;
- `src/sort-core.js`: pure sorting logic, covered by `tests/`;
- `src/page.css`: category headings.

## License

[MIT](LICENSE)
