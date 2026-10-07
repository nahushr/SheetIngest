import { useState } from "react";
import FileUploadOutlinedIcon from "@mui/icons-material/FileUploadOutlined";
import PeopleAltOutlinedIcon from "@mui/icons-material/PeopleAltOutlined";
import {
  Alert,
  Box,
  Button,
  Chip,
  Container,
  Paper,
  Stack,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Typography,
} from "@mui/material";
import {
  SheetIngest,
  downloadTemplate,
} from "@simplishelf/sheet-ingest";
import type {
  SheetIngestIssue,
  SheetIngestTemplate,
} from "@simplishelf/sheet-ingest";

type ContactRow = {
  firstName: string;
  lastName: string;
  email: string;
  company: string;
  role: string;
  phone: string;
};

type Contact = ContactRow & { id: string };

const contactTemplate: SheetIngestTemplate = {
  fileName: "contacts_import_template.xlsx",
  sheetName: "Contacts",
  groups: [
    {
      label: "Contact details",
      fields: [
        {
          key: "firstName",
          label: "First name",
          aliases: ["Given name"],
          required: true,
          example: "Amara",
        },
        {
          key: "lastName",
          label: "Last name",
          aliases: ["Surname", "Family name"],
          required: true,
          example: "Cole",
        },
        {
          key: "email",
          label: "Email address",
          aliases: ["Email"],
          required: true,
          example: "amara@northstar.example",
        },
      ],
    },
    {
      label: "Organization",
      fields: [
        {
          key: "company",
          label: "Company",
          aliases: ["Organization", "Account"],
          example: "Northstar Labs",
        },
        { key: "role", label: "Role", aliases: ["Job title"] },
        { key: "phone", label: "Phone", aliases: ["Telephone"] },
      ],
    },
  ],
};

const initialContacts: Contact[] = [
  {
    id: "contact-1001",
    firstName: "Amara",
    lastName: "Cole",
    email: "amara.cole@northstar.example",
    company: "Northstar Labs",
    role: "Operations lead",
    phone: "+1 415 555 0132",
  },
  {
    id: "contact-1002",
    firstName: "Theo",
    lastName: "Bennett",
    email: "theo.bennett@fieldwork.example",
    company: "Fieldwork Studio",
    role: "Product manager",
    phone: "+1 212 555 0168",
  },
  {
    id: "contact-1003",
    firstName: "Priya",
    lastName: "Shah",
    email: "priya.shah@brightline.example",
    company: "Brightline Health",
    role: "Program director",
    phone: "+1 312 555 0181",
  },
  {
    id: "contact-1004",
    firstName: "Mateo",
    lastName: "Rivera",
    email: "mateo.rivera@redwood.example",
    company: "Redwood Supply",
    role: "Account executive",
    phone: "+1 617 555 0144",
  },
];

const samplePeople = [
  ["Amara", "Cole"],
  ["Theo", "Bennett"],
  ["Priya", "Shah"],
  ["Mateo", "Rivera"],
  ["Nina", "Patel"],
  ["Eli", "Morgan"],
  ["Sofia", "Reed"],
  ["Lucas", "Brooks"],
];

const sampleOrganizations = [
  "Northstar Labs",
  "Fieldwork Studio",
  "Brightline Health",
  "Redwood Supply",
  "Juniper Works",
];

const sampleRoles = [
  "Operations lead",
  "Product manager",
  "Program director",
  "Account executive",
  "Customer success",
];

function makeTestContacts(count: number): ContactRow[] {
  return Array.from({ length: count }, (_, index) => {
    const [firstName, lastName] = samplePeople[index % samplePeople.length];
    const number = index + 1;

    return {
      firstName,
      lastName,
      email: `${firstName.toLowerCase()}.${lastName.toLowerCase()}${number}@example.com`,
      company: sampleOrganizations[index % sampleOrganizations.length],
      role: sampleRoles[index % sampleRoles.length],
      phone: `+1 415 555 ${String(1000 + number).slice(-4)}`,
    };
  });
}

export default function App() {
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [contacts, setContacts] = useState(initialContacts);
  const [notice, setNotice] = useState("");

  const mapContactRow = (values: Readonly<Record<string, string>>): ContactRow => ({
    firstName: values.firstName?.trim() ?? "",
    lastName: values.lastName?.trim() ?? "",
    email: values.email?.trim().toLowerCase() ?? "",
    company: values.company?.trim() ?? "",
    role: values.role?.trim() ?? "",
    phone: values.phone?.trim() ?? "",
  });

  const validateContact = (contact: ContactRow): SheetIngestIssue | undefined => {
    const isEmailValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contact.email);

    if (!isEmailValid) {
      return {
        field: "email",
        message: "Enter a valid email address.",
      };
    }

    return undefined;
  };

  const generateTestData = async (count: number) => {
    await downloadTemplate(contactTemplate, makeTestContacts(count));
  };

  const importContacts = async (rows: readonly ContactRow[], file: File) => {
    const imported = rows.map((contact, index) => ({
      ...contact,
      id: `import-${Date.now()}-${index}`,
    }));

    setContacts((current) => [...imported, ...current]);
    setNotice(`Imported ${rows.length} contacts from ${file.name}.`);
  };

  return (
    <main className="demo-shell">
      <Container maxWidth="lg" disableGutters>
        <header className="demo-topbar">
          <img className="demo-logo" src="/sheet-ingest-logo.svg" alt="SheetIngest" />
          <span className="demo-topbar__tag">React example</span>
        </header>

        <section className="demo-intro">
          <p className="demo-kicker">SPREADSHEET IMPORT FOR REACT</p>
          <Typography component="h1" variant="h3">
            Bring a spreadsheet into your app.
          </Typography>
          <Typography className="demo-intro__copy">
            Try the full import flow with grouped contact fields, column matching,
            row validation, and a host-owned submit callback.
          </Typography>
        </section>

        <section className="directory" aria-labelledby="directory-title">
          <div className="directory-heading">
            <Stack direction="row" spacing={1.5} alignItems="center">
              <div className="directory-icon" aria-hidden="true">
                <PeopleAltOutlinedIcon fontSize="small" />
              </div>
              <Box>
                <Typography id="directory-title" component="h2" variant="h5">
                  Contacts
                </Typography>
                <Typography className="directory-count">
                  {contacts.length} records in this example
                </Typography>
              </Box>
            </Stack>
            <Button
              variant="contained"
              startIcon={<FileUploadOutlinedIcon />}
              onClick={() => {
                setNotice("");
                setIsImportOpen(true);
              }}
            >
              Import contacts
            </Button>
          </div>

          {notice && (
            <Alert severity="success" onClose={() => setNotice("")}>
              {notice}
            </Alert>
          )}

          <TableContainer component={Paper} variant="outlined" className="contacts-table">
            <Table aria-label="Example contact directory">
              <TableHead>
                <TableRow>
                  <TableCell>Contact</TableCell>
                  <TableCell>Company</TableCell>
                  <TableCell>Role</TableCell>
                  <TableCell>Phone</TableCell>
                  <TableCell>Status</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {contacts.slice(0, 8).map((contact) => (
                  <TableRow key={contact.id}>
                    <TableCell>
                      <div className="contact-name">
                        <strong>{contact.firstName} {contact.lastName}</strong>
                        <span>{contact.email}</span>
                      </div>
                    </TableCell>
                    <TableCell>{contact.company}</TableCell>
                    <TableCell>{contact.role}</TableCell>
                    <TableCell>{contact.phone}</TableCell>
                    <TableCell><Chip size="small" label="Active" color="success" /></TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
          {contacts.length > 8 && (
            <Typography className="table-footnote">
              Showing the first 8 rows. {contacts.length - 8} more imported records are in the list.
            </Typography>
          )}
        </section>

        <footer className="demo-footer">
          Templates, parsing, column matching, validation display, and preview come from SheetIngest.
          The example app owns generated data and contact submission.
        </footer>
      </Container>

      <SheetIngest<ContactRow>
        isOpen={isImportOpen}
        onClose={() => setIsImportOpen(false)}
        template={contactTemplate}
        mapRow={mapContactRow}
        validateRow={validateContact}
        onGenerateTestData={generateTestData}
        onSubmit={importContacts}
        defaultTestDataCount={25}
        maxTestDataCount={500}
      />
    </main>
  );
}
