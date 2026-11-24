'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const ASSET_DIRECTORIES = ['web-apps', 'sdkjs', 'fonts', 'x2t'];
const EDITORS = ['documenteditor', 'spreadsheeteditor', 'presentationeditor'];
const PREVIOUS_SDK_VERSION = /\s*\/\* editor-sdk-version:start \*\/[\s\S]*?\/\* editor-sdk-version:end \*\//g;
const ASSET_PATH = /(["']|url\(\s*)(?:\/editor-assets\/[a-f0-9]{16})?\/(web-apps|sdkjs|fonts|x2t)\//g;

function writeIfChanged(file, contents) {
  if (fs.readFileSync(file, 'utf8') === contents) {
    return;
  }

  fs.writeFileSync(file, contents);

  for (const extension of ['.gz', '.br']) {
    fs.rmSync(file + extension, { force: true });
  }
}

function versionEditorAssets(outputDir) {
  // remove the old SDK-only version before hashing the vendor files
  for (const editor of EDITORS) {
    const file = path.join(outputDir, 'web-apps/apps', editor, 'main/index.html');
    writeIfChanged(file, fs.readFileSync(file, 'utf8').replace(PREVIOUS_SDK_VERSION, ''));
  }

  const hash = crypto.createHash('sha256');

  function hashDirectory(directory) {
    for (const name of fs.readdirSync(directory).sort()) {
      const file = path.join(directory, name);

      if (fs.statSync(file).isDirectory()) {
        hashDirectory(file);
      } else if (!/\.(gz|br)$/.test(name)) {
        const relative = path.relative(outputDir, file).split(path.sep).join('/');
        hash.update(relative + '\0').update(fs.readFileSync(file)).update('\0');
      }
    }
  }

  for (const directory of ASSET_DIRECTORIES) {
    hashDirectory(path.join(outputDir, directory));
  }

  const assetPath = '/editor-assets/' + hash.digest('hex').slice(0, 16);

  // relative URLs inside DocsAPI, RequireJS, CSS and WASM inherit this prefix
  for (const entry of fs.readdirSync(outputDir, { withFileTypes: true })) {
    if (!entry.isFile() || !/\.(html|js|css)$/.test(entry.name)) {
      continue;
    }

    const file = path.join(outputDir, entry.name);
    const original = fs.readFileSync(file, 'utf8');
    const result = original.replace(ASSET_PATH, (_, start, directory) => `${start}${assetPath}/${directory}/`);
    writeIfChanged(file, result);
  }

  return assetPath;
}

module.exports = versionEditorAssets;

if (require.main === module) {
  const outputDir = process.argv[2];

  if (!outputDir) {
    throw new Error('Usage: node version-editor-assets.js <output directory>');
  }

  versionEditorAssets(path.resolve(outputDir));
}
