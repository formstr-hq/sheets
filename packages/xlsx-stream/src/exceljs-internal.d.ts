// Types for the exceljs internals deep-imported by excelInternals.ts (the only importer).
// They are CommonJS modules whose module.exports is the class / function itself.
declare module "exceljs/lib/stream/xlsx/worksheet-reader" {
  const WorksheetReader: new (options: {
    workbook: unknown;
    id: string | number;
    iterator: AsyncIterable<string | Uint8Array>;
    options?: { worksheets?: string; hyperlinks?: string };
  }) => {
    workbook: any;
    iterator: AsyncIterable<string | Uint8Array>;
    [Symbol.asyncIterator](): AsyncIterator<any>;
  };
  export default WorksheetReader;
}
declare module "exceljs/lib/doc/row" {
  const Row: new (worksheet: unknown, number: number) => any;
  export default Row;
}
declare module "exceljs/lib/utils/utils" {
  const utils: {
    excelToDate(value: number, date1904?: boolean): Date;
    isDateFmt(format?: string): boolean;
    xmlDecode(text: string): string;
  };
  export default utils;
}
declare module "exceljs/lib/xlsx/xform/style/styles-xform" {
  const StylesXform: new (initialise?: boolean) => {
    init(): void;
    parseStream(stream: AsyncIterable<string | Uint8Array>): Promise<unknown>;
    getStyleModel(id: number): unknown;
  };
  export default StylesXform;
}
declare module "exceljs/lib/xlsx/xform/book/workbook-xform" {
  const WorkbookXform: new () => {
    parseStream(stream: AsyncIterable<string | Uint8Array>): Promise<unknown>;
  };
  export default WorkbookXform;
}
declare module "exceljs/lib/xlsx/xform/core/relationships-xform" {
  const RelationshipsXform: new () => {
    parseStream(stream: AsyncIterable<string | Uint8Array>): Promise<unknown>;
  };
  export default RelationshipsXform;
}
declare module "exceljs/lib/utils/parse-sax" {
  /** One SAX event; `value` is a saxes tag ({ name, attributes, ... }) or a text string. */
  export interface SaxEvent {
    eventType: "opentag" | "text" | "closetag";
    value: any;
  }
  const parseSax: (iterable: AsyncIterable<string | Uint8Array>) => AsyncGenerator<SaxEvent[]>;
  export default parseSax;
}
