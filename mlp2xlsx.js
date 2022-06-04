// Playlist converter for mAirList playlist files to Excel
// (c) frankmathy@gmail.com, 2022
const xml2js = require("xml2js");
const fs = require("fs");
var xl = require("excel4node");

if (process.argv.length < 4) {
  console.log(
    `Usage: ${process.argv[0]} ${process.argv[1]} <ExcelFileName.xlsx> <Input1.mlp,Input2.mlp,...>`
  );
  process.exit();
}
const excelFileName = process.argv[2];
const mlpFiles = process.argv.slice(3);

var wb = new xl.Workbook();
var ws = wb.addWorksheet("Playlist");
ws.cell(1, 1).string("Playlist: ");
let row = 3;

mlpFiles.forEach((fileName) => {
  var songCount = 0;

  // read XML from a file
  const xml = fs.readFileSync(fileName);

  // convert XML to JSON
  xml2js.parseString(xml, { mergeAttrs: true }, (err, data) => {
    if (err) {
      throw err;
    }
    data.Playlist.PlaylistItem.forEach((item) => {
      const [itemType] = item.Type;
      if (itemType === "Music") {
        const [artist] = item.Artist;
        const [title] = item.Title;
        ws.cell(row, 1).string(artist);
        ws.cell(row, 2).string(title);
        ws.cell(row++, 3).string(artist + " - " + title);
        songCount++;
      }
    });
  });
  console.log(`${songCount} songs in ${fileName}`);
});

wb.write(excelFileName);
console.log(`Playlist written to ${excelFileName}`);
