// Read Excel Playlists and create one single Playlist with dates
// (c) frankmathy@gmail.com, 2024
const xlsx = require("xlsx");
const nodeXlsx = require("node-xlsx");
const fs = require("fs");
const path = require("path");
const moment = require("moment");
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
    try {
      const fileDate = moment(dirName, "YYYY-MM-DD");
      return fileDate.isValid() ? fileDate.toDate() : dirName;
    } catch (error) {
      return dirname;
    }
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

var allSongs = [["Artist", "Title", "Date"]];

const files = fs.readdirSync(directoryName, {
  recursive: true,
});
const xlsxFiles = files.filter((file) => file.toLowerCase().endsWith(".xlsx"));
xlsxFiles.forEach((xlsFileName) => {
  if (xlsFileName.indexOf("~") < 0) {
    const xlsPathName = path.join(directoryName, xlsFileName);
    const file = xlsx.readFile(xlsPathName);
    const sheets = file.SheetNames;
    if (sheets.length > 0) {
      const trackData = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[0]], { header: 1 });
      const filteredTrackData = trackData
        .filter((line) => isValidTrackEntry(line[0]))
        .map((filteredLine) => [
          filteredLine[0].trim(),
          filteredLine[1].trim(),
          getDateFromFileName(xlsFileName),
          xlsPathName,
        ]);
      console.log(`${xlsPathName} => ${filteredTrackData.length} songs`);
      allSongs = allSongs.concat(filteredTrackData);
    }
  }
});

const xlsXBuffer = nodeXlsx.build([{ name: " Playlists Songs", data: allSongs }]);
fs.writeFileSync(excelFileName, xlsXBuffer);
