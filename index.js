const XLSX = require('xlsx');
const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');

const app = express();
const port = 3000;

app.use(bodyParser.json());

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

app.post('/convert', (req, res) => {
  const jsonData = req.body.jsonData;

  jsonToExcel(jsonData);
  const filePath = __dirname + '/output.xlsx';

  res.download(filePath, 'output.xlsx', (err) => {
    fs.unlinkSync(filePath);
  });
});

app.listen(port, () => {
  console.log(`App listening at http://localhost:${port}`);
});

function main() {
  const jsondata = require('./json-excel.json');
  jsonToExcel(jsondata);
}

// main();

function jsonToExcel(jsondata) {
  var refinedData = [];
  for (const i in jsondata) {
    const currentProduct = jsondata[i];
    const pageNo = currentProduct['row_capture']['PageNo'];
    const SubType = currentProduct['row_capture']['SubType'];
    const Text = currentProduct['row_capture']['Text'];
    const Characterization = currentProduct['row_capture']['Characterization'];
    const Category = currentProduct['row_capture']['Category'];
    const Category_Score = currentProduct['row_capture']['Category_Score'];
    const Char_Score = currentProduct['row_capture']['Char_Score'];

    if (
      currentProduct.allocations == [] ||
      currentProduct.allocations == undefined ||
      null
    ) {
      const PBS = 'pbs';
      const ABS = 'abs';
      const OBS = 'OBS';
      const Receivers = 'Receivers';

      const currentRefinedProdcut = {
        pageNo,
        SubType,
        Text,
        Char_Score,
        Characterization,
        Category,
        Category_Score,
        PBS,
        ABS,
        OBS,
        Receivers,
      };
      refinedData.push(currentRefinedProdcut);
      continue;
    }

    for (const j in currentProduct.allocations) {
      const allocations = currentProduct.allocations[j];
      const PBS =
        allocations.pbs == (null || undefined) ? 'ABCD' : allocations.pbs;
      const OBS =
        allocations.obs == (null || undefined) ? 'OBS' : allocations.obs;
      const ABS =
        allocations.abs == (null || undefined) ? 'ABS' : allocations.abs;
      const Receivers =
        allocations.receiver == null ? 'Receivers' : allocations.receiver;

      const currentRefinedProdcut = {
        pageNo,
        SubType,
        Text,
        Char_Score,
        Characterization,
        Category,
        Category_Score,
        PBS,
        OBS,
        ABS,
        Receivers,
      };
      refinedData.push(currentRefinedProdcut);
    }
  }

  const worksheet = XLSX.utils.json_to_sheet(refinedData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Allocation');

  customMergeCells(worksheet);

  XLSX.writeFile(workbook, 'output.xlsx');
}

function customMergeCells(worksheet) {
  const range = XLSX.utils.decode_range(worksheet['!ref']);

  for (let C = 0; C <= range.e.c; C++) {
    const mergedCells = {};
    let previouscellvalue = undefined;
    for (let R = 1; R <= range.e.r; R++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cellValue = worksheet[cellAddress]?.v;
      previouscellvalue = cellValue;

      if (!mergedCells[cellValue]) {
        mergedCells[cellValue] = { start: cellAddress, end: cellAddress };
      } else {
        const prevStart = mergedCells[cellValue].start;
        const prevEnd = mergedCells[cellValue].end;

        const currentCellAddress = cellAddress;
        const currentStart = currentCellAddress;
        const currentEnd = currentCellAddress;

        const prevCellAddr = XLSX.utils.decode_cell(prevStart);
        const currCellAddr = XLSX.utils.decode_cell(currentStart);

        const prevCol = prevCellAddr.c;
        const prevRow = prevCellAddr.r;

        const currentCol = currCellAddr.c;
        const currentRow = currCellAddr.r;

        if (
          previouscellvalue !== cellValue &&
          mergedCells[cellValue] &&
          R != 1
        ) {
          addToMerges(
            worksheet,
            mergedCells[cellValue].start,
            mergedCells[cellValue].end
          );
          mergedCells[cellValue].start = cellAddress;
          mergedCells[cellValue].end = cellAddress;
        } else if (prevCol === currentCol && prevRow === currentRow - 1) {
          mergedCells[cellValue].end = currentEnd;
        } else {
          mergedCells[cellValue] = { start: prevStart, end: currentEnd };
        }
        previouscellvalue = cellValue;
      }
    }

    for (const key in mergedCells) {
      addToMerges(worksheet, mergedCells[key].start, mergedCells[key].end);
    }
  }
}
function addToMerges(worksheet, start, end) {
  const mergedRange = `${start}:${end}`;
  worksheet['!merges'] = worksheet['!merges'] || [];
  const range = XLSX.utils.decode_range(mergedRange);
  worksheet['!merges'].push({
    s: range.s,
    e: range.e,
  });
}

function hashCode(str) {
  let hash = 0;
  if (str.length === 0) {
    return hash;
  }
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = (hash << 5) - hash + char;
    hash &= hash;
  }
  return hash;
}
