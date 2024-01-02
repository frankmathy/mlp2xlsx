// Playlist converter for mAirList playlist files to Excel
// (c) frankmathy@gmail.com, 2022
const xml2js = require("xml2js");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
var xl = require("excel4node");

const isValidTrackEntry = (entryString) => {
  return (
    entryString !== undefined &&
    entryString !== "" &&
    !entryString.startsWith("Playlist") &&
    !entryString.startsWith("Künstler") &&
    !entryString.startsWith("Artist")
  );
};

const getDateFromFileName = (fileName) => {
  if (fileName !== undefined && fileName.length > 0) {
    const dirName = fileName.split("/")[0];
    const dates = dirName.split("-");
    return dates.length === 3 ? `${dates[2]}.${dates[1]}.${dates[0]}` : dirName;
  } else {
    return undefined;
  }
};

if (process.argv.length < 4) {
  console.log(`Usage: ${process.argv[0]} ${process.argv[1]} <DirectoryName> <Output.xlsx>`);
  process.exit();
}
const directoryName = process.argv[2];
const excelFileName = process.argv[3];

const files = fs.readdirSync(directoryName, {
  recursive: true,
});
const xlsxFiles = files.filter((file) => file.toLowerCase().endsWith(".xlsx"));
xlsxFiles.forEach((xlsFileName) => {
  if (xlsFileName.indexOf("~") < 0) {
    const xlsPathName = path.join(directoryName, xlsFileName);
    console.log(xlsPathName);
    const file = xlsx.readFile(xlsPathName);
    const sheets = file.SheetNames;
    if (sheets.length > 0) {
      const trackData = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[0]], { header: 1 });
      const filteredTrackData = trackData
        .filter((line) => isValidTrackEntry(line[0]))
        .map((filteredLine) => [filteredLine[0], filteredLine[1], getDateFromFileName(xlsFileName)]);
      console.log(JSON.stringify(filteredTrackData));
    }
  }
});

/*
var wb = new xl.Workbook();
var ws = wb.addWorksheet("Songs");
ws.cell(1, 1).string("Playlist: ");
let row = 3;
ws.cell(row, 1).string("Artist");
ws.cell(row, 2).string("Title");
ws.cell(row++, 3).string("Artist - Title");

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
      if (item.Type !== undefined) {
        const [itemType] = item.Type;
        if (itemType === "Music") {
          const [artist] = item.Artist;
          const [title] = item.Title;
          ws.cell(row, 1).string(artist);
          ws.cell(row, 2).string(title);
          ws.cell(row++, 3).string(artist + " - " + title);
          songCount++;
        }
      }
    });
  });
  console.log(`${songCount} songs in ${fileName}`);
});

wb.write(excelFileName);
console.log(`Playlist written to ${excelFileName}`);
*/
