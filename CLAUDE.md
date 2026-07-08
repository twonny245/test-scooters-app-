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
