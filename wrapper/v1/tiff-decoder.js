const UTIF = require('utif');

const TIFF_ORIENTATION = {
  NORMAL: 1,
  MIRROR_HORIZONTAL: 2,
  ROTATE_180: 3,
  MIRROR_VERTICAL: 4,
  MIRROR_MAIN_DIAGONAL: 5,
  ROTATE_90_CLOCKWISE: 6,
  MIRROR_ANTI_DIAGONAL: 7,
  ROTATE_90_COUNTERCLOCKWISE: 8
};

// UTIF returns one byte per channel, in red, green, blue, alpha order
const RGBA_CHANNEL_COUNT = 4;
const RGB_CHANNEL_COUNT = 3;
const ALPHA_CHANNEL_OFFSET = 3;
const MAX_CHANNEL_VALUE = 255;

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
    Number.isSafeInteger(width * height * RGBA_CHANNEL_COUNT);

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

  if (data.length !== width * height * RGBA_CHANNEL_COUNT) {
    throw new Error('TIFF pixel data is incomplete');
  }

  const hasAssociatedAlpha = photometric === 2 && image.t338 && image.t338[0] === 1;

  if (hasAssociatedAlpha) {
    for (let offset = 0; offset < data.length; offset += RGBA_CHANNEL_COUNT) {
      const alpha = data[offset + ALPHA_CHANNEL_OFFSET];

      if (alpha > 0 && alpha < MAX_CHANNEL_VALUE) {
        for (let channel = 0; channel < RGB_CHANNEL_COUNT; channel++) {
          data[offset + channel] = Math.min(MAX_CHANNEL_VALUE, Math.round(data[offset + channel] * MAX_CHANNEL_VALUE / alpha));
        }
      }
    }
  }

  return {
    width,
    height,
    data,
    orientation: image.t274 ? image.t274[0] : TIFF_ORIENTATION.NORMAL
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

  const canSkipOrientationTransform = image.orientation < TIFF_ORIENTATION.MIRROR_HORIZONTAL ||
      image.orientation > TIFF_ORIENTATION.ROTATE_90_COUNTERCLOCKWISE;

  if (canSkipOrientationTransform) {
    return canvas.toDataURL('image/png');
  }

  const orientedCanvas = document.createElement('canvas');
  const swapsDimensions = [
    TIFF_ORIENTATION.MIRROR_MAIN_DIAGONAL,
    TIFF_ORIENTATION.ROTATE_90_CLOCKWISE,
    TIFF_ORIENTATION.MIRROR_ANTI_DIAGONAL,
    TIFF_ORIENTATION.ROTATE_90_COUNTERCLOCKWISE
  ].includes(image.orientation);
  orientedCanvas.width = swapsDimensions ? image.height : image.width;
  orientedCanvas.height = swapsDimensions ? image.width : image.height;

  const transforms = {
    [TIFF_ORIENTATION.MIRROR_HORIZONTAL]: [-1, 0, 0, 1, image.width, 0],
    [TIFF_ORIENTATION.ROTATE_180]: [-1, 0, 0, -1, image.width, image.height],
    [TIFF_ORIENTATION.MIRROR_VERTICAL]: [1, 0, 0, -1, 0, image.height],
    [TIFF_ORIENTATION.MIRROR_MAIN_DIAGONAL]: [0, 1, 1, 0, 0, 0],
    [TIFF_ORIENTATION.ROTATE_90_CLOCKWISE]: [0, 1, -1, 0, image.height, 0],
    [TIFF_ORIENTATION.MIRROR_ANTI_DIAGONAL]: [0, -1, -1, 0, image.height, image.width],
    [TIFF_ORIENTATION.ROTATE_90_COUNTERCLOCKWISE]: [0, -1, 1, 0, 0, image.width]
  };
  const orientedContext = orientedCanvas.getContext('2d');
  orientedContext.setTransform(...transforms[image.orientation]);
  orientedContext.drawImage(canvas, 0, 0);

  return orientedCanvas.toDataURL('image/png');
}

module.exports = { decode, toPngDataUrl };
