// Shared site header/nav for the AA Scooters app.
//
// Every page includes this once (`<script src="nav.js" defer></script>` in
// <head>) and has an empty `<div id="topbar-mount"></div>` where the old
// inline topbar used to be. This script injects the topbar markup AND its
// CSS at load time, and marks whichever link matches the current page as
// active -- so adding, renaming, or reordering a nav link only ever needs
// editing here, not in every page.
//
// Relies on the CSS custom properties --petrol, --cone, and --line already
// being defined on :root by each page's own stylesheet (they all define
// the same brand palette).
(function () {
  var NAV_ITEMS = [
    { href: 'index.html', label: 'Home' },
    { href: 'customers.html', label: 'Customer Record' },
    { href: 'contract.html', label: 'Contract' },
    { href: 'pricing.html', label: 'Price Calculator' },
    { href: 'parts.html', label: 'Parts &amp; Oil' },
    { href: 'oilchange.html', label: 'Oil Change' },
    { href: 'bikes.html', label: 'Bikes Status' },
    { href: 'bikephotos.html', label: 'Bike Photos' },
    { href: 'available-bikes.html', label: 'Available Bikes' },
    { href: 'accounts.html', label: 'Accounts' }
  ];

  var TOPBAR_CSS = '\n' +
    '  .topbar{\n' +
    '    background:var(--petrol);\n' +
    '    padding:14px 16px;\n' +
    '    display:flex;\n' +
    '    align-items:center;\n' +
    '    justify-content:space-between;\n' +
    '    flex-wrap:wrap;\n' +
    '    gap:10px;\n' +
    '  }\n' +
    '  .topbar .brand{\n' +
    '    display:flex;\n' +
    '    align-items:center;\n' +
    '    gap:8px;\n' +
    '    text-decoration:none;\n' +
    '  }\n' +
    '  .topbar .brand img{\n' +
    '    width:28px; height:28px;\n' +
    '    object-fit:contain;\n' +
    '  }\n' +
    '  .topbar .brand span{\n' +
    "    font-family:'Barlow Condensed',sans-serif;\n" +
    '    font-weight:700;\n' +
    '    font-size:15px;\n' +
    '    color:#fff;\n' +
    '    letter-spacing:.02em;\n' +
    '  }\n' +
    '  .topbar nav{\n' +
    '    display:flex;\n' +
    '    gap:16px;\n' +
    '    flex-wrap:wrap;\n' +
    '  }\n' +
    '  .topbar nav a{\n' +
    '    color:#CFE3E0;\n' +
    '    text-decoration:none;\n' +
    '    font-size:12.5px;\n' +
    '    font-weight:500;\n' +
    '    padding:3px 0;\n' +
    '    border-bottom:2px solid transparent;\n' +
    '  }\n' +
    '  .topbar nav a:hover{ color:#fff; }\n' +
    '  .topbar nav a.active{ color:#fff; border-bottom-color:var(--cone); }\n';

  function currentPage() {
    var path = window.location.pathname.split('/').pop();
    return path || 'index.html';
  }

  function injectCss() {
    if (document.getElementById('shared-topbar-css')) return;
    var style = document.createElement('style');
    style.id = 'shared-topbar-css';
    style.textContent = TOPBAR_CSS;
    document.head.appendChild(style);
  }

  function buildLinksHtml() {
    var current = currentPage();
    return NAV_ITEMS.map(function (item) {
      var active = item.href === current ? ' class="active"' : '';
      return '<a href="' + item.href + '"' + active + '>' + item.label + '</a>';
    }).join('\n    ');
  }

  function renderTopbar() {
    var mount = document.getElementById('topbar-mount');
    if (!mount) return; // page opted out of the shared header

    injectCss();

    mount.outerHTML =
      '<div class="topbar">\n' +
      '  <a class="brand" href="index.html">\n' +
      '    <img src="https://scooterrentalchiangmai.com/wp-content/uploads/2025/02/cropped-logo-3333-101x105.png" alt="AA Scooters logo">\n' +
      '    <span>AA Scooter Rental</span>\n' +
      '  </a>\n' +
      '  <nav>\n' +
      '    ' + buildLinksHtml() + '\n' +
      '  </nav>\n' +
      '</div>';
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', renderTopbar);
  } else {
    renderTopbar();
  }
})();
