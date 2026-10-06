import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { SheetIngest } from "../src/SheetIngest";
import type { SheetIngestTemplate } from "../src/types";

vi.mock("@simplishelf/super-data-grid", () => ({ default: () => null }));

const template: SheetIngestTemplate = {
  fileName: "contacts.xlsx",
  sheetName: "Contacts",
  groups: [
    {
      label: "Contact details",
      fields: [
        { key: "firstName", label: "First name", required: true },
        { key: "email", label: "Email address", aliases: ["Email"] },
      ],
    },
  ],
};

describe("SheetIngest import flow", () => {
  it("detects grouped headers during upload and goes straight to column mapping", async () => {
    const csv = [
      "Contact details,,",
      "firstName,email,phone",
      "Ada,ada@example.com,+1 415 555 0100",
    ].join("\n");
    const file = new File([csv], "contacts.csv", { type: "text/csv" });
    Object.defineProperty(file, "text", { value: async () => csv });

    render(
      <SheetIngest
        isOpen
        onClose={vi.fn()}
        onSubmit={vi.fn()}
        template={template}
      />,
    );

    fireEvent.change(screen.getByLabelText("Upload spreadsheet"), {
      target: { files: [file] },
    });

    expect(await screen.findByText("Match each template field to the column that contains its data.")).toBeInTheDocument();
    expect(screen.getByText("contacts.csv")).toBeInTheDocument();
    expect(screen.getByText("contacts · 1 row · 3 columns")).toBeInTheDocument();
    expect(screen.getByText("Contact details")).toBeInTheDocument();
    expect(screen.queryByText("Select header rows")).not.toBeInTheDocument();
  });
});
