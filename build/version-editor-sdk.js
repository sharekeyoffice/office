'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const EDITORS = {
  documenteditor: 'word',
  spreadsheeteditor: 'cell',
  presentationeditor: 'slide',
};
const CONFIG = /\bvar\s+require\s*=\s*\{/g;
const PREVIOUS_VERSION = /\s*\/\* editor-sdk-version:start \*\/[\s\S]*?\/\* editor-sdk-version:end \*\//g;

function versionEditorSdk(outputDir) {
  for (const [editor, type] of Object.entries(EDITORS)) {
    const sdkPath = path.join(outputDir, 'sdkjs', type, 'sdk-all-min.js');
    const htmlPath = path.join(outputDir, 'web-apps/apps', editor, 'main/index.html');
    const hash = crypto.createHash('sha256').update(fs.readFileSync(sdkPath)).digest('hex').slice(0, 16);
    const original = fs.readFileSync(htmlPath, 'utf8');
    const html = original.replace(PREVIOUS_VERSION, '');

    if ((html.match(CONFIG) || []).length !== 1) {
      throw new Error(`Expected one RequireJS bootstrap in ${htmlPath}`);
    }

    if (/\burlArgs\s*:/.test(html)) {
      throw new Error(`RequireJS urlArgs already configured in ${htmlPath}`);
    }

    // сhange the SDK URL only when its contents change
    const config = `
            /* editor-sdk-version:start */
            urlArgs: function (id, url) {
                return id === 'sdk' ? (url.indexOf('?') === -1 ? '?' : '&') + 'sdkv=${hash}' : '';
            },
            /* editor-sdk-version:end */`;
    const result = html.replace(CONFIG, match => match + config);

    if (result !== original) {
      fs.writeFileSync(htmlPath, result);
    }
  }
}

module.exports = versionEditorSdk;

if (require.main === module) {
  const outputDir = process.argv[2];

  if (!outputDir) {
    throw new Error('Usage: node version-editor-sdk.js <output directory>');
  }

  versionEditorSdk(path.resolve(outputDir));
}
