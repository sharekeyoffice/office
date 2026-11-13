const tiffDecoder = require('./tiff-decoder');

function bitmapToPng(file) {
  const url = URL.createObjectURL(file);
  const image = new Image();

  return new Promise(function (resolve, reject) {
    const timeout = setTimeout(function () {
      reject(new Error('BMP image loading timed out'));
    }, 15e3);

    image.onload = function () {
      clearTimeout(timeout);

      try {
        const canvas = document.createElement('canvas');
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        canvas.getContext('2d').drawImage(image, 0, 0);
        resolve(canvas.toDataURL('image/png'));
      } catch (error) {
        reject(error);
      }
    };
    image.onerror = function () {
      clearTimeout(timeout);
      reject(new Error('BMP image could not be decoded'));
    };
    image.src = url;
  }).finally(function () {
    image.onload = null;
    image.onerror = null;
    URL.revokeObjectURL(url);
  });
}

function toPngDataUrl(file) {
  const isTiffFile = /\.(tif|tiff)$/i.test(file.name || '') ||
    /^image\/tiff?$/i.test(file.type || '');

  if (isTiffFile) {
    return file.arrayBuffer().then(tiffDecoder.toPngDataUrl);
  }

  return bitmapToPng(file);
}

module.exports = { toPngDataUrl };
