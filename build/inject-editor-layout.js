'use strict';

const fs = require('node:fs');
const path = require('node:path');

const MARKER = '<!-- editor-layout-bootstrap -->';
const EDITORS = ['documenteditor', 'spreadsheeteditor', 'presentationeditor'];

function initializeEditorLayout() {
  'use strict';

  let loadingScreen = null;
  let isDocumentReady = false;
  const editorTypes = new Set(['word', 'cell', 'slide']);

  function updateLoadingScreen(role) {
    const shouldPreserveLoadingScreen = role && editorTypes.has(role.type);

    if (!shouldPreserveLoadingScreen) {
      removeLoadingScreen();

      return;
    }

    if (isDocumentReady || loadingScreen) {
      return;
    }

    const mask = document.getElementById('loading-mask');

    if (!mask) {
      return;
    }

    loadingScreen = mask.cloneNode(true);
    loadingScreen.id = 'editor-loading-screen';
    loadingScreen.setAttribute('aria-hidden', 'true');

    // Sheets keep their ribbon in both roles; Word and Slides use a compact loader
    if (role.type !== 'cell') {
      loadingScreen.querySelectorAll('.sktoolbar, .brendpanel > :nth-child(2)').forEach(element => element.remove());
    }

    loadingScreen.querySelectorAll('[id]').forEach(element => element.removeAttribute('id'));
    document.body.appendChild(loadingScreen);
  }

  function removeLoadingScreen() {
    if (loadingScreen) {
      loadingScreen.remove();
      loadingScreen = null;
    }
  }

  window.finishEditorLoading = function () {
    if (isDocumentReady) {
      return;
    }

    isDocumentReady = true;
    requestAnimationFrame(() => requestAnimationFrame(removeLoadingScreen));
  };

  function setEditorLayoutRole(role) {
    const classes = document.documentElement.classList;
    const isKnown = role && typeof role.canEdit === 'boolean';
    const isViewer = isKnown && !role.canEdit;

    classes.toggle('editor-fixed-header', editorTypes.has(role.type));
    classes.toggle('editor-sheet-skeleton', role.type === 'cell');
    classes.toggle('editor-role-pending', !isKnown);
    classes.toggle('sk-role-viewer', !!isViewer);
    classes.toggle('sk-viewer-compact', !!isViewer && role.type !== 'cell');
    classes.toggle('sk-viewer-sheet', !!isViewer && role.type === 'cell');
    updateLoadingScreen(role);
  }

  function applyHostEditorLayoutRole() {
    let role;

    try {
      role = window.parent.editorLayoutRole;
    } catch (error) {}

    setEditorLayoutRole(role || { canEdit: true });
  }

  window.setEditorLayoutRole = setEditorLayoutRole;
  document.addEventListener('DOMContentLoaded', applyHostEditorLayoutRole, { once: true });
  applyHostEditorLayoutRole();
}

function injectEditorLayout(siteRoot) {
  for (const editor of EDITORS) {
    const file = path.join(siteRoot, 'web-apps', 'apps', editor, 'main', 'index.html');

    if (!fs.existsSync(file)) {
      continue;
    }

    const html = fs.readFileSync(file, 'utf8');

    if (!html.includes('</head>')) {
      throw new Error('Missing </head> in ' + file);
    }

    const newline = html.includes('\r\n') ? '\r\n' : '\n';
    const bootstrap = [
      MARKER,
      '    <link id="sk-editor-ui-styles" rel="stylesheet" href="/editor-ui.css">',
      '    <script>(' + initializeEditorLayout.toString().replace(/\r?\n/g, newline) + ')();</script>'
    ].join(newline);
    const updated = html.includes(MARKER)
      ? html.replace(/<!-- editor-layout-bootstrap -->[\s\S]*?<\/script>/, () => bootstrap)
      : html.replace('</head>', () => '    ' + bootstrap + newline + '</head>');

    if (updated !== html) {
      fs.writeFileSync(file, updated);
    }
  }
}

module.exports = injectEditorLayout;

if (require.main === module) {
  if (!process.argv[2]) {
    throw new Error('Usage: node inject-editor-layout.js <site-root>');
  }

  injectEditorLayout(process.argv[2]);
}
