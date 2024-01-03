// Read Excel Playlists and create one single Playlist with dates
// (c) frankmathy@gmail.com, 2024
const xlsx = require("xlsx");
const nodeXlsx = require("node-xlsx");
const fs = require("fs");
const path = require("path");
const moment = require("moment");
var xl = require("excel4node");

const isValidString = (value) => typeof value === "string";

const getDateFromFileName = (fileName) => {
  if (fileName !== undefined && fileName.length > 0) {
    const pathNameParts = fileName.split("/");
    const dirName = pathNameParts[pathNameParts.length - 2];
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

var allSongs = [["Artist", "Title", "Date", "File Name"]];

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
      trackData.forEach((line, lineIndex) => {
        const artistColumn = line.findIndex((element) => element === "Artist" || element === "Künstler");
        const titleColumn = line.findIndex((element) => element === "Titel" || element === "Title");
        if (artistColumn >= 0 && titleColumn >= 0) {
          const filteredTrackData = trackData.slice(lineIndex + 1);
          outputTrackData = filteredTrackData
            .filter((line) => isValidString(line[artistColumn]) && isValidString(line[titleColumn]))
            .map((filteredLine) => [
              filteredLine[artistColumn].trim(),
              filteredLine[titleColumn].trim(),
              getDateFromFileName(xlsFileName),
              xlsPathName,
            ]);
          console.log(`${xlsPathName} => ${filteredTrackData.length} songs`);
          allSongs = allSongs.concat(outputTrackData);
          return false;
        }
      });
    }
  }
});

const xlsXBuffer = nodeXlsx.build([{ name: " Playlists Songs", data: allSongs }]);
fs.writeFileSync(excelFileName, xlsXBuffer);
