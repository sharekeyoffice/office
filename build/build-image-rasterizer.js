'use strict';

const path = require('node:path');
const fs = require('node:fs');
const esbuild = require('esbuild');

function buildImageRasterizer(outputDirectory) {
  const licenses = ['utif', 'pako'].map(function (name) {
    return '/* ' + name + '\n' +
      fs.readFileSync(require.resolve(name + '/LICENSE'), 'utf8') + '\n*/';
  }).join('\n');

  esbuild.buildSync({
    entryPoints: [path.join(__dirname, '../wrapper/v1/image-rasterizer.js')],
    outfile: path.join(outputDirectory, 'image-rasterizer.js'),
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'OfficeImageRasterizer',
    minify: true,
    legalComments: 'inline',
    banner: { js: licenses }
  });
}

module.exports = buildImageRasterizer;

if (require.main === module) {
  buildImageRasterizer(process.argv[2]);
}
