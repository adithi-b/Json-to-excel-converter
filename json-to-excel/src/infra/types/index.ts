export type CellStyling = {
  font: FontProperties;
};

export type FontProperties = {
  name: string;
  size: number;
  color: ColorProperties;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  alignment: AlignmentProperties;
  fill: FillProperties;
  border: boolean;
  comment: string;
};

export type ColorProperties = {
  argb: string;
};

export type AlignmentProperties = {
  horizontal:
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous"
    | "distributed";
};

export type FillProperties = {
  type: "pattern";
  pattern: "solid";
  fgColor: ColorProperties;
};

export type MasterListDataType = {
  id: string;
  listName: string;
  usage: string;
  header: HeaderData;
  values: Array<MasterDataItem>;
};

export type HeaderData = {
  code: string;
  value: string;
  comment?: string;
};

export type MasterDataItem = {
  code: number;
  value: string;
};

export type WorkbookListTemplate = {
  fetchExcelGenerate: Array<WorkbookTemplate>;
};

export type WorkbookTemplate = {
  moduleCode: string;
  modulesWrkBook: Array<WorkbookDetails>;
  primModuleCode: string;
  templateCode: string;
};

export type WorkbookDetails = {
  exlWBName: string;
  exlWBProperties: string;
  exlWBVersion: string;
  moduleCode: string;
  primModuleCode: string;
  templateCode: string;
  templateDesc: string;
  workBooKSheets: Array<SheetTemplate>;
};

export type SheetTemplate = {
  customIntegerOutput1: number;
  exlWBName: string;
  exlWBSheetCode: string;
  exlWBSheetName: string;
  exlWBSheetProperties: string;
  wrkBookSheetsDtl: Array<SheetDetails>;
};

export type SheetDetails = {
  exlWBName: string;
  exlWBSheetCode: string;
  exlWBSheetColDispName: string;
  exlWBSheetColDSType: string;
  exlWBSheetColDSValue: string;
  exlWBSheetColFormat: string;
};
