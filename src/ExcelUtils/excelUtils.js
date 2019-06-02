/* eslint-disable no-restricted-syntax */
/* eslint-disable prefer-template */
/* eslint-disable radix */
/* eslint-disable no-nested-ternary */
/* eslint-disable no-unused-expressions */
/* eslint-disable no-plusplus */
/* eslint-disable guard-for-in */
import XLSX from 'xlsx';

export default {
  loadExcelFile(file) {
    return new Promise(rel => {
      const rABS = false;
      const reader = new FileReader();
      reader.onload = e => {
        let data = e.target.result;
        if (!rABS) data = new Uint8Array(data);
        const workbook = XLSX.read(data, {
          type: rABS ? 'binary' : 'array',
        });

        rel(workbook);
      };
      if (rABS) reader.readAsBinaryString(file);
      else reader.readAsArrayBuffer(file);
    });
  },
  formatDate(numb, format) {
    const time = new Date((numb - 1) * 24 * 3600000 + 1);
    time.setYear(time.getFullYear() - 70);
    const year = time.getFullYear() + '';
    const month = time.getMonth() + 1 + '';
    const date = time.getDate() + '';
    if (format && format.length === 1) {
      return year + format + month + format + date;
    }
    return year + (month < 10 ? '0' + month : month) + (date < 10 ? '0' + date : date);
  },
  excelSheetToJson(workbook, sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    const sheetObj = XLSX.utils.sheet_to_json(worksheet);
    for (let i = 0; i < sheetObj.length; i++) {
      if (sheetObj[i].更新日期) {
        sheetObj[i].更新日期 = this.formatDate(sheetObj[i].更新日期, '-');
      }
    }
    return sheetObj;
  },
  initialFullPartMiddleSheet(worksheet, columns) {
    const middleSheet = [];
    for (let sheetIndex = 0; sheetIndex < worksheet.length; sheetIndex++) {
      const row = {
        APN: worksheet[sheetIndex].APN,
        SUPPLIER: worksheet[sheetIndex].appleKey,
        Commodity: worksheet[sheetIndex].commodity,
        'Supplier Desc': worksheet[sheetIndex].supplier,
        Buyer: worksheet[sheetIndex].buyer,
        BRShipCommitDate: '',
        ManualShipCommitDate: '',
        'DIS-COUNT': 0,
      };

      for (let colIndex = 0; colIndex < columns.length; colIndex++) {
        row[columns[colIndex]] = '';
      }
      middleSheet.push(row);
    }
    return middleSheet;
  },
  initialMiddleSheet(worksheet, headers, columns) {
    const middleSheet = [];
    for (let sheetIndex = 0; sheetIndex < worksheet.length; sheetIndex++) {
      const row = {};
      for (let headerIndex = 0; headerIndex < headers.length; headerIndex++) {
        if (headers[headerIndex] === 'DIS-COUNT') {
          row[headers[headerIndex]] = 0;
        } else {
          row[headers[headerIndex]] = worksheet[sheetIndex][headers[headerIndex]];
        }
      }
      for (let colIndex = 0; colIndex < columns.length; colIndex++) {
        row[columns[colIndex]] = '';
      }
      row['DIS-COUNT'] = 0;
      middleSheet.push(row);
    }
    return middleSheet;
  },
  initialDisMiddleSheet(worksheet, headers, columns, supplier) {
    const middleSheet = [];
    for (let sheetIndex = 0; sheetIndex < worksheet.length; sheetIndex++) {
      if (worksheet[sheetIndex].SUPPLIER === supplier) {
        let row = {};
        for (let headerIndex = 0; headerIndex < headers.length; headerIndex++) {
          if (headers[headerIndex] === 'DIS-COUNT') {
            row[headers[headerIndex]] = '';
          } else {
            row[headers[headerIndex]] = worksheet[sheetIndex][headers[headerIndex]];
          }
        }
        for (let colIndex = 0; colIndex < columns.length; colIndex++) {
          row[columns[colIndex]] = 'No Plan';
        }
        middleSheet.push(row);
        row = JSON.parse(JSON.stringify(row));
        middleSheet.push(row);
        break;
      }
    }
    return middleSheet;
  },
  getOriginalCompareData(
    middleSheet,
    BRSheet,
    manualSheet,
    BRColumns,
    manualColumns,
    middleSheetColumns
  ) {
    const BRRow = middleSheet[0];
    BRRow['DIS-COUNT'] = 'BuildRabbit';
    const ManualRow = middleSheet[1];
    ManualRow['DIS-COUNT'] = 'ManualPlan';
    for (let brIndex = 0; brIndex < BRSheet.length; brIndex++) {
      if (BRSheet[brIndex].SUPPLIER === BRRow.SUPPLIER) {
        const oriBRRow = BRSheet[brIndex];
        if (ManualRow['更新日期']) {
          ManualRow['更新日期'] = oriBRRow.ShipCommitDate;
        }
        for (let middleColIndex = 0; middleColIndex < middleSheetColumns.length; middleColIndex++) {
          BRRow[middleSheetColumns[middleColIndex]] =
            middleColIndex < BRColumns.length
              ? oriBRRow[BRColumns[middleColIndex]]
                ? oriBRRow[BRColumns[middleColIndex]]
                : 0
              : 0;
        }
        break;
      }
    }
    for (let manualIndex = 0; manualIndex < manualSheet.length; manualIndex++) {
      if (manualSheet[manualIndex].SUPPLIER === ManualRow.SUPPLIER) {
        const oriManualRow = manualSheet[manualIndex];
        if (ManualRow.ShipCommitDate) {
          ManualRow.ShipCommitDate = oriManualRow['更新日期'];
        }
        for (let middleColIndex = 0; middleColIndex < middleSheetColumns.length; middleColIndex++) {
          ManualRow[middleSheetColumns[middleColIndex]] =
            middleColIndex < manualColumns.length
              ? oriManualRow[manualColumns[middleColIndex]]
                ? oriManualRow[manualColumns[middleColIndex]]
                : 0
              : 0;
        }
        break;
      }
    }
    return middleSheet;
  },
  compareFullWorksheet(
    middleSheet,
    BRSheet,
    manualSheet,
    BRColumns,
    manualColumns,
    middleSheetColumns
  ) {
    for (let sheetIndex = 0; sheetIndex < middleSheet.length; sheetIndex++) {
      const middleRow = middleSheet[sheetIndex];
      let BRRow = {};
      let manualRow = {};
      for (let brIndex = 0; brIndex < BRSheet.length; brIndex++) {
        if (BRSheet[brIndex].SUPPLIER === middleRow.SUPPLIER) {
          BRRow = BRSheet[brIndex];
          middleRow.BRShipCommitDate = BRRow.ShipCommitDate;
          break;
        }
      }
      for (let manualIndex = 0; manualIndex < manualSheet.length; manualIndex++) {
        if (manualSheet[manualIndex].SUPPLIER === middleRow.SUPPLIER) {
          manualRow = manualSheet[manualIndex];
          middleRow.ManualShipCommitDate = manualRow['更新日期'];
          break;
        }
      }
      for (let middleColIndex = 0; middleColIndex < middleSheetColumns.length; middleColIndex++) {
        const BRValue =
          middleColIndex < BRColumns.length
            ? BRRow[BRColumns[middleColIndex]]
              ? BRRow[BRColumns[middleColIndex]]
              : 0
            : 0;
        const manualValue =
          middleColIndex < manualColumns.length
            ? manualRow[manualColumns[middleColIndex]]
              ? manualRow[manualColumns[middleColIndex]]
              : 0
            : 0;
        middleRow[middleSheetColumns[middleColIndex]] =
          Number.parseInt(BRValue) - Number.parseInt(manualValue);

        middleRow['DIS-COUNT'] += Number.parseInt(BRValue) - Number.parseInt(manualValue);
      }
    }
    return middleSheet;
  },
  compareWorksheet(
    middleSheet,
    BRSheet,
    manualSheet,
    BRColumns,
    manualColumns,
    middleSheetColumns
  ) {
    for (let sheetIndex = 0; sheetIndex < middleSheet.length; sheetIndex++) {
      const middleRow = middleSheet[sheetIndex];
      let BRRow = {};
      let manualRow = {};
      for (let brIndex = 0; brIndex < BRSheet.length; brIndex++) {
        if (BRSheet[brIndex].SUPPLIER === middleRow.SUPPLIER) {
          BRRow = BRSheet[brIndex];
          break;
        }
      }
      for (let manualIndex = 0; manualIndex < manualSheet.length; manualIndex++) {
        if (manualSheet[manualIndex].SUPPLIER === middleRow.SUPPLIER) {
          manualRow = manualSheet[manualIndex];
          break;
        }
      }
      for (let middleColIndex = 0; middleColIndex < middleSheetColumns.length; middleColIndex++) {
        const BRValue =
          middleColIndex < BRColumns.length
            ? BRRow[BRColumns[middleColIndex]]
              ? BRRow[BRColumns[middleColIndex]]
              : 0
            : 0;
        const manualValue =
          middleColIndex < manualColumns.length
            ? manualRow[manualColumns[middleColIndex]]
              ? manualRow[manualColumns[middleColIndex]]
              : 0
            : 0;
        middleRow[middleSheetColumns[middleColIndex]] =
          Number.parseInt(BRValue) - Number.parseInt(manualValue);

        middleRow['DIS-COUNT'] += Number.parseInt(BRValue) - Number.parseInt(manualValue);
      }
    }
    return middleSheet;
  },
  excelToJsonData(data, rABS) {
    const workbook = XLSX.read(data, { type: rABS ? 'binary' : 'array' });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const sheetObj = XLSX.utils.sheet_to_json(worksheet);
    return sheetObj;
  },
  jsonDataToExcel(jsonData, fileName = 'BuildRabbit.xlsx', sheetName = 'Sheet0') {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(jsonData);

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);
  },
  /**
   * 根据JsonData中的key生成多个sheet
   */
  jsonDataToExcelMoreSheet(JsonData, fileName = 'cer_report.xlsx') {
    const wb = XLSX.utils.book_new();
    for (const key in JsonData) {
      const ws = XLSX.utils.json_to_sheet(JsonData[key]);
      XLSX.utils.book_append_sheet(wb, ws, key);
    }
    XLSX.writeFile(wb, fileName);
  },
};
