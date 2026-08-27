const fs = require("fs");
const path = require("path");

function pngToIco(pngBuffer) {
  if (pngBuffer.length < 24 || pngBuffer.toString("ascii", 1, 4) !== "PNG") {
    throw new Error("Not a PNG file");
  }
  const width = pngBuffer.readUInt32BE(16);
  const height = pngBuffer.readUInt32BE(20);
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(1, 4);
  const entry = Buffer.alloc(16);
  entry.writeUInt8(width >= 256 ? 0 : width, 0);
  entry.writeUInt8(height >= 256 ? 0 : height, 1);
  entry.writeUInt8(0, 2);
  entry.writeUInt8(0, 3);
  entry.writeUInt16LE(1, 4);
  entry.writeUInt16LE(32, 6);
  entry.writeUInt32LE(pngBuffer.length, 8);
  entry.writeUInt32LE(22, 12);
  return Buffer.concat([header, entry, pngBuffer]);
}

function writeIconFiles(dir, pngBuffer) {
  fs.mkdirSync(dir, { recursive: true });
  const pngPath = path.join(dir, "icon.png");
  const icoPath = path.join(dir, "icon.ico");
  fs.writeFileSync(pngPath, pngBuffer);
  fs.writeFileSync(icoPath, pngToIco(pngBuffer));
  console.log("Wrote", pngPath);
  console.log("Wrote", icoPath);
}

module.exports = { pngToIco, writeIconFiles };
