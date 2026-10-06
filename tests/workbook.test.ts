import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  createTemplateWorkbook,
  detectHeaderRow,
  getSourceColumns,
  normalizeHeader,
  readSpreadsheetFile,
  suggestColumnMappings,
} from "../src/workbook";
import type { SheetIngestTemplate } from "../src/types";

const template: SheetIngestTemplate = {
  fileName: "people.xlsx",
  sheetName: "People",
  groups: [
    {
      label: "Contact details",
      fields: [
        { key: "firstName", label: "First name", aliases: ["Given name"], required: true },
        { key: "email", label: "Email address", aliases: ["Email"] },
      ],
    },
  ],
};

describe("spreadsheet template helpers", () => {
  it("normalizes spreadsheet labels without punctuation or accents", () => {
    expect(normalizeHeader("  Prénom / Email  ")).toBe("prenomemail");
  });

  it("detects a two-row grouped template's field header", () => {
    const rows = [
      ["Contact details", ""],
      ["firstName", "email"],
      ["Ada", "ada@example.com"],
    ];
    expect(detectHeaderRow(rows, template.groups.flatMap((group) => group.fields))).toBe(1);
  });

  it("suggests unique mappings using field labels and aliases", () => {
    const mappings = suggestColumnMappings(
      ["Email", "Given name", "Ignore me"],
      template.groups.flatMap((group) => group.fields),
    );
    expect(mappings).toEqual({ firstName: 1, email: 0 });
  });

  it("carries merged group headings across source columns", () => {
    const columns = getSourceColumns(
      [
        ["Contact details", ""],
        ["firstName", "email"],
        ["Ada", "ada@example.com"],
      ],
      1,
    );
    expect(columns.map(({ header, group }) => [header, group])).toEqual([
      ["firstName", "Contact details"],
      ["email", "Contact details"],
    ]);
  });

  it("writes grouped headers, merge ranges, and frozen header rows", async () => {
    const workbook = createTemplateWorkbook(template);
    const output = await workbook.xlsx.writeBuffer();
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(output);
    const worksheet = loaded.getWorksheet("People");

    expect(worksheet?.model.merges).toContain("A1:B1");
    expect(worksheet?.getCell(1, 1).value).toBe("Contact details");
    expect(worksheet?.getCell(2, 1).value).toBe("firstName");
    expect(worksheet?.getCell(2, 2).value).toBe("email");
    expect(worksheet?.views[0]).toMatchObject({ state: "frozen", ySplit: 2 });
  });

  it("adds host-generated records beneath the same grouped template headers", async () => {
    const rows = [
      { firstName: "Ada", email: "ada@example.com" },
      { firstName: "Grace", email: "grace@example.com" },
    ];
    const workbook = createTemplateWorkbook(template, rows);
    const output = await workbook.xlsx.writeBuffer();
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(output);
    const loadedSheet = loaded.getWorksheet("People");

    expect(loadedSheet?.getCell(3, 1).value).toBe("Ada");
    expect(loadedSheet?.getCell(3, 2).value).toBe("ada@example.com");
    expect(loadedSheet?.getCell(4, 1).value).toBe("Grace");
    expect(loadedSheet?.getCell(4, 2).value).toBe("grace@example.com");
  });

  it("reads grouped XLSX headers and values from a browser file", async () => {
    const workbook = createTemplateWorkbook(template);
    workbook.getWorksheet("People")?.addRow(["Ada", "ada@example.com"]);
    const output = await workbook.xlsx.writeBuffer();
    const file = new File([new Uint8Array(output)], "people.xlsx");
    Object.defineProperty(file, "arrayBuffer", {
      value: async () => output,
    });

    const sheet = await readSpreadsheetFile(file);
    expect(sheet.rows).toEqual([
      ["Contact details", ""],
      ["firstName", "email"],
      ["Ada", "ada@example.com"],
    ]);
  });

  it("parses comma-separated values from a browser file", async () => {
    const file = new File(["firstName,email\nAda,ada@example.com"], "people.csv");
    Object.defineProperty(file, "text", {
      value: async () => "firstName,email\nAda,ada@example.com",
    });

    const sheet = await readSpreadsheetFile(file);
    expect(sheet.rows).toEqual([
      ["firstName", "email"],
      ["Ada", "ada@example.com"],
    ]);
  });
});
