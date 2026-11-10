const UTIF = require('utif');

function decode(bytes) {
  const directories = UTIF.decode(bytes);
  const image = directories.find(function (directory) {
    return directory.t256 && directory.t257;
  });

  if (!image) {
    throw new Error('TIFF contains no image');
  }

  const width = image.t256[0];
  const height = image.t257[0];
  const hasValidDimensions = Number.isSafeInteger(width) && width > 0 &&
    Number.isSafeInteger(height) && height > 0 &&
    Number.isSafeInteger(width * height * 4);

  if (!hasValidDimensions) {
    throw new Error('TIFF image dimensions are invalid');
  }

  UTIF.decodeImage(bytes, image);

  if (!image.data || !image.data.length) {
    throw new Error('TIFF image could not be decoded');
  }

  const photometric = image.t262 ? image.t262[0] : 2;

  if (![0, 1, 2, 3, 5].includes(photometric)) {
    throw new Error('TIFF color format is not supported');
  }

  const data = UTIF.toRGBA8(image);

  if (data.length !== width * height * 4) {
    throw new Error('TIFF pixel data is incomplete');
  }

  const hasAssociatedAlpha = photometric === 2 && image.t338 && image.t338[0] === 1;

  if (hasAssociatedAlpha) {
    for (let offset = 0; offset < data.length; offset += 4) {
      const alpha = data[offset + 3];

      if (alpha > 0 && alpha < 255) {
        for (let channel = 0; channel < 3; channel++) {
          data[offset + channel] = Math.min(255,
            Math.round(data[offset + channel] * 255 / alpha));
        }
      }
    }
  }

  return {
    width,
    height,
    data,
    orientation: image.t274 ? image.t274[0] : 1
  };
}

function toPngDataUrl(bytes) {
  const image = decode(bytes);
  const canvas = document.createElement('canvas');
  canvas.width = image.width;
  canvas.height = image.height;

  const context = canvas.getContext('2d');
  const pixels = context.createImageData(image.width, image.height);
  pixels.data.set(image.data);
  context.putImageData(pixels, 0, 0);

  if (image.orientation < 2 || image.orientation > 8) {
    return canvas.toDataURL('image/png');
  }

  const orientedCanvas = document.createElement('canvas');
  const swapsDimensions = image.orientation >= 5;
  orientedCanvas.width = swapsDimensions ? image.height : image.width;
  orientedCanvas.height = swapsDimensions ? image.width : image.height;

  const transforms = {
    2: [-1, 0, 0, 1, image.width, 0],
    3: [-1, 0, 0, -1, image.width, image.height],
    4: [1, 0, 0, -1, 0, image.height],
    5: [0, 1, 1, 0, 0, 0],
    6: [0, 1, -1, 0, image.height, 0],
    7: [0, -1, -1, 0, image.height, image.width],
    8: [0, -1, 1, 0, 0, image.width]
  };
  const orientedContext = orientedCanvas.getContext('2d');
  orientedContext.setTransform(...transforms[image.orientation]);
  orientedContext.drawImage(canvas, 0, 0);

  return orientedCanvas.toDataURL('image/png');
}

module.exports = { decode, toPngDataUrl };
