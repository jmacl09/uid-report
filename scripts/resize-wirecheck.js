const { Jimp } = require('jimp');
const path = require('path');

async function resize() {
  const inPath = path.join(__dirname, '..', 'src', 'assets', 'wirecheck.png');
  const outPath = inPath; // overwrite the original file
  console.log('Reading', inPath);
  const image = await Jimp.read(inPath);
  const maxWidth = 1600;
  if (image.bitmap.width > maxWidth) {
    console.log(`Resizing from ${image.bitmap.width} -> ${maxWidth}`);
    image.resize(maxWidth, Jimp.AUTO);
  } else {
    console.log('Image smaller than max width, skipping resize.');
  }
  // If PNG, Jimp will write a PNG; use deflateLevel to help size
  if (path.extname(inPath).toLowerCase() === '.png') {
    image.deflateLevel(9);
  }
  // write back (will overwrite)
  await image.writeAsync(outPath);
  const { size } = require('fs').statSync(outPath);
  console.log('Wrote', outPath, 'size bytes:', size);
}

resize().catch(err => {
  console.error('Resize failed:', err);
  process.exit(1);
});
