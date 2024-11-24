// IMPORT PACKAGES
import * as ExcelJS from "exceljs";

// IMPORT CONSTANSTS
import { borderProperties, MASTER_LIST_DELIMITER } from "../constants";

// IMPORT TYPES
import {
  CellStyling,
  MasterListDataType,
  WorkbookListTemplate,
} from "../types";

/**
 * Returns the name of the column for given column index
 */
const getColumnLetter = (columnIndex: number) => {
  let columnName = "";
  while (columnIndex >= 0) {
    const remainder = columnIndex % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    columnIndex = Math.floor(columnIndex / 26) - 1;
  }
  return columnName;
};

/**
 * Returns a excel consisting of color-coded template with a master list of values
 */
export const generateExcelTemplate = (
  dataSets: WorkbookListTemplate,
  excelData?: Array<any>,
  displayMasterSheet: boolean = true,
  withErrorValidation: boolean = true
) => {
  dataSets.fetchExcelGenerate.forEach((dataSet) => {
    dataSet.modulesWrkBook.forEach((workbookDetails) => {
      //Generate work book
      const workbook = new ExcelJS.Workbook();

      let masterDataList: MasterListDataType[] = [];

      workbookDetails.workBooKSheets.forEach((workSheet) => {
        workSheet.wrkBookSheetsDtl.forEach((column, index) => {
          // store the array of master values in a master list. This will be used to generate the master sheet in excel
          if (column?.exlWBSheetColDSValue?.length > 1) {
            const masterData = column.exlWBSheetColDSValue
              .split(MASTER_LIST_DELIMITER)
              .map((item, index) => {
                return { code: index, value: item };
              });

            const record: MasterListDataType = {
              id: column.exlWBSheetColDispName,
              listName: column.exlWBSheetColDispName,
              usage: `used in ${workSheet.exlWBSheetName}`,
              header: {
                code: `Sl No`,
                value: column.exlWBSheetColDispName,
              },
              values: masterData.filter((item) => item.value.length > 0),
            };
            masterDataList.push(record);
          }
        });
      });

      // Creating master list in Master sheet to get the cell positions that will be used for data validation
      const master = workbook.addWorksheet("Master-1");

      const masterData = masterDataList.map((masterList, index) => {
        const firstColumn = getColumnLetter(3 * index + 1);
        const secondColumn = getColumnLetter(3 * index + 2);

        // print the headers till the 3rd row
        let row = 4;

        // render the master values from 4th row
        masterList.values.forEach((data) => {
          const firstCell = `${firstColumn}${row}`;
          const secondCell = `${secondColumn}${row}`;
          master.getCell(firstCell).value = data.code;
          master.getCell(secondCell).value = data.value;
          row += 1;
        });

        // storing the absolute cell indices of master lists
        return {
          id: masterList.id,
          listName: masterList.listName,
          startIndex: `$${secondColumn}$4`,
          endIndex: `$${secondColumn}$${masterList.values.length + 3}`,
        };
      });

      workbookDetails.workBooKSheets.forEach((workSheet) => {
        // store the count of rows to which data validation is to be provided
        let rowCount = workSheet.customIntegerOutput1;

        // Generate worksheets
        const sheet = workbook.addWorksheet(workSheet.exlWBSheetName);
        workSheet.wrkBookSheetsDtl.forEach((column, index) => {
          let rowIndex = 1; // This will be incremented to insert all the column values later
          const columnName = getColumnLetter(index);

          const cell = sheet.getCell(`${columnName}${rowIndex}`);

          // Setting the width of the column
          cell.value = column.exlWBSheetColDispName;
          sheet.getColumn(columnName).width =
            column.exlWBSheetColDispName.length + 2;

          try {
            const jsonObject = stringToJSON(column.exlWBSheetColFormat);

            const cellStyling: CellStyling = JSON.parse(jsonObject);

            // removing the hash in front of the hex-color code (if at all there)
            const fgColor =
              cellStyling.font.fill.fgColor.argb.indexOf("#") === 0
                ? cellStyling.font.fill.fgColor.argb.split("#")[1]
                : cellStyling.font.fill.fgColor.argb;
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: fgColor },
            };

            cell.font = {
              bold: cellStyling.font.bold,
              italic: cellStyling.font.italic,
              name: cellStyling.font.name,
              size: cellStyling.font.size,
              color: cellStyling.font.color,
            };

            cell.alignment = cellStyling.font.alignment;

            if (cellStyling.font.border) {
              cell.border = borderProperties;
            }

            // adding comment to the cell
            if (cellStyling.font.comment && cellStyling.font.comment != null) {
              cell.note = {
                texts: [{ text: cellStyling.font.comment }],
              };
            }
          } catch (e) {
            console.log("Error in styling", e);
          }

          if (column.exlWBSheetColDSValue?.length > 1) {
            const masterListAddress = masterData.filter(
              (data) => data.id === column.exlWBSheetColDispName
            );

            if (masterListAddress.length !== 0) {
              const formula = `${"Master"}!${masterListAddress[0].startIndex}:${
                masterListAddress[0].endIndex
              }`;

              for (let rowNumber = 2; rowNumber <= rowCount + 1; rowNumber++) {
                sheet.getCell(`${columnName}${rowNumber}`).dataValidation = {
                  type: "list",
                  formulae: [formula],
                  allowBlank: true,
                  error: "Please choose an input from the master list",
                  errorTitle: "Invalid input!",
                  showErrorMessage: withErrorValidation,
                };
              }
            }
          }
        });
        if (excelData) {
          // render the data below headers
          Object.keys(excelData).forEach((sheetName: any) => {
            const sheetData = excelData[sheetName];
            let rowIndex = 1;

            if (Array.isArray(sheetData)) {
              const sheet = workbook.getWorksheet(sheetName);

              // For multiple rows of data
              sheetData.forEach((rowData: any) => {
                rowIndex++; // This will be incremented to insert all the column values later
                const keys = Object.keys(rowData);

                keys.forEach((item: string, index: number) => {
                  const columnName = getColumnLetter(index);
                  const cell = sheet?.getCell(`${columnName}${rowIndex}`);
                  if (cell) cell.value = rowData[item];
                  if (sheet)
                    sheet.getColumn(columnName).width =
                      rowData[item].length + 2;
                });
              });
            } else {
              excelData?.forEach((rowData) => {
                rowIndex++; // This will be incremented to insert all the column values later

                const keys = Object.keys(rowData);

                keys.forEach((item: string, index: number) => {
                  const columnName = getColumnLetter(index);

                  const cell = sheet.getCell(`${columnName}${rowIndex}`);

                  cell.value = rowData?.[item];
                });
              });
            }
          });
        }
      });

      // remove the sheet that was generated to get the master list data positions
      const sheetToRemove = workbook.getWorksheet("Master-1");
      if (sheetToRemove) {
        workbook.removeWorksheet(sheetToRemove.id);
      }
      if (displayMasterSheet) {
        // Creating master list in "Master" sheet
        const master = workbook.addWorksheet("Master");

        // font color for first row
        master.getRow(1).font = { color: { argb: "548235" }, bold: true };
        master.getRow(1).alignment = { horizontal: "center" };

        // font color for second row
        master.getRow(2).font = { color: { argb: "2F75B5" }, bold: true };
        master.getRow(2).alignment = { horizontal: "center" };

        master.getRow(3).font = { bold: true };
        master.getRow(3).alignment = { horizontal: "center" };

        masterDataList.forEach((masterList, index) => {
          // NOTE: index starts from 0

          // variable to store column width
          let firstColumnWidth = 0;
          let secondColumnWidth = 0;

          const firstColumn = getColumnLetter(3 * index + 1);
          const secondColumn = getColumnLetter(3 * index + 2);

          // render the header and description alone by merging 2 cells of 1st and 2nd rows

          master.mergeCells(`${firstColumn}1: ${secondColumn}1`);
          master.getCell(`${firstColumn}1`).value = masterList.listName;
          master.getCell(`${firstColumn}1`).border = borderProperties;

          master.mergeCells(`${firstColumn}2: ${secondColumn}2`);
          master.getCell(`${firstColumn}2`).value = masterList.usage;
          master.getCell(`${firstColumn}2`).border = borderProperties;

          // print the headers in the 3rd row; can add comments here itself if any
          let row = 3;
          master.getCell(`${firstColumn}${row}`).value = masterList.header.code;
          master.getCell(`${firstColumn}${row}`).border = borderProperties;
          master.getCell(`${secondColumn}${row}`).value =
            masterList.header.value;
          master.getCell(`${secondColumn}${row}`).border = borderProperties;

          if (firstColumnWidth < masterList.header.code.toString().length)
            firstColumnWidth = masterList.header.code.toString().length;

          if (secondColumnWidth < masterList.header.value.length)
            secondColumnWidth = masterList.header.value.length;

          // To add comment
          if (masterList.header.comment)
            master.getCell(`${secondColumn}${row}`).note = {
              texts: [{ text: masterList.header.comment }],
            };

          row += 1;

          // render the master values from 4th row
          masterList.values.forEach((data: any) => {
            const firstCell = `${firstColumn}${row}`;
            const secondCell = `${secondColumn}${row}`;
            master.getCell(firstCell).value = data.code;
            master.getCell(secondCell).value = data.value;
            row += 1;

            // styling
            master.getCell(firstCell).alignment = {
              horizontal: "center",
            };

            master.getCell(firstCell).border = borderProperties;
            master.getCell(secondCell).border = borderProperties;

            if (firstColumnWidth < data.code.toString().length)
              firstColumnWidth = data.code.toString().length;

            if (secondColumnWidth < data.value.length)
              secondColumnWidth = data.value.length;
          });

          if (row < masterList.values.length + 4)
            row = masterList.values.length + 4;

          master.getColumn(firstColumn).width = firstColumnWidth + 2;
          master.getColumn(secondColumn).width = secondColumnWidth + 2;
        });
        // protecting the master sheet from being altered
        /* 
          "Secure random number generation is not supported by this browser. 
          Use Chrome, Firefox or Internet Explorer 11"
        */
        master.protect("", {});
      }

      // Generate the Excel file
      workbook.xlsx.writeBuffer().then((buffer: any) => {
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = workbookDetails.exlWBName;
        a.click();
        window.URL.revokeObjectURL(url);
      });
    });
  });
};

/**
 * Converts JSON string to JSON
 */
export const stringToJSON = (serializedData: string) => {
  const cleanedString = serializedData
    ?.replace(/\\u0022/g, '"')
    ?.replace(/\\\\/g, "\\");

  return JSON.parse(cleanedString);
};

/**
 * Converts JSON obtained from API response to excel
 */
export const generateExcelFromJson = (
  data: any,
  columns: any,
  excelName: string
) => {
  // Create a new instance of a Workbook class
  const workbook = new ExcelJS.Workbook();

  // Add a worksheet
  const worksheet = workbook.addWorksheet("Sheet 1");

  worksheet.columns = columns;

  // Add rows from JSON data
  data.forEach((row: any) => {
    worksheet.addRow(row);
  });
  // Generate the Excel file
  workbook.xlsx.writeBuffer().then((buffer: any) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = excelName;
    a.click();
    window.URL.revokeObjectURL(url);
  });
};
