# Project conventions — AA Scooters app

## Modal / lightbox "click outside to close" pattern

Every popup that closes when the user clicks its backdrop/overlay (outside
the card itself) must guard against closing mid-drag. If the check is only
`if (e.target === backdropEl) close();` on a `click` listener, then
click-and-drag text selection inside an input near the edge of the modal
can end (mouseup) over the backdrop and incorrectly close the whole popup.

Fix: track whether the *mousedown* also started on the backdrop, and only
close if both mousedown and the click's target are the backdrop:

```js
let mouseDownOnBackdrop = false;
backdropEl.addEventListener('mousedown', (e) => {
  mouseDownOnBackdrop = (e.target === backdropEl);
});
backdropEl.addEventListener('click', (e) => {
  if (e.target === backdropEl && mouseDownOnBackdrop) close();
  mouseDownOnBackdrop = false;
});
```

Applied so far in: `accounts.html` (Add/Edit modal), `bikephotos.html`
(lightbox). Apply this same pattern to any new modal/lightbox/overlay
added to this project going forward.

## Code.gs edits always need a manual redeploy — flag it loudly

`Code.gs` is Google Apps Script. Saving/editing the file on disk (or in this
repo) does **not** update the live web app that every page calls through
`scriptUrl` — that only happens when the code is pasted into the Apps Script
editor and a new version is deployed (Deploy → Manage deployments → Edit →
New version). Until that's done, the pages keep hitting the OLD deployed
code, so a new/changed backend action can silently fail or 404 with no
obvious cause.

Whenever `Code.gs` is created or edited, this must be called out
unmissably at the end of the response — not folded into a paragraph.
Always use this exact callout, on its own line:

🔴 **Code.gs changed — redeploy required.** Paste the updated Code.gs into
the Apps Script editor, then Deploy → Manage deployments → Edit → New
version.

Do this every single time Code.gs changes, even if it was already flagged
earlier in the same conversation.
