// IMPORT PACKAGES
import * as ExcelJS from "exceljs";

export const MASTER_LIST_DELIMITER = "~";
export const FORMULA_DELIMITER = ",";

export const borderProperties: Partial<ExcelJS.Borders> = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" },
};
