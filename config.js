/*
 * KC Timesheet: settings.
 * This file is published with the add-in, so it must NOT contain client names.
 * Client names are read from SharePoint using the member of staff's own
 * Microsoft 365 sign-in.
 */
window.KC_CONFIG = {
  // Microsoft Entra ID
  tenantId: "d15beb8b-78b2-4fe4-9ead-0d1204377b43",   // Directory (tenant) ID: confirmed for kdpc.uk
  clientId: "5c0081ab-6727-49c1-8b6e-ede656d392fd",   // Application (client) ID: confirmed

  // Where the client list lives in SharePoint
  sharePointHost: "kaleidoscopeconsultants.sharepoint.com",
  sitePath: "",                        // "" = main KC site (KC Operations home)
  source: "file",                      // "file" = CSV in a document library; "list" = SharePoint list
  // Path inside the site's Documents library (do not include "Shared Documents")
  filePath: "/Timesheet guidance/KC Timesheets.csv",
  listName: "",                        // only used when source is "list"

  // Column headings in the file (the first row of the CSV)
  nameColumn: "Title",                 // client name exactly as written in the subject
  aliasColumn: "Aliases",              // other recognised names, separated by semicolons

  // Subject format
  separator: " – ",

  // Internal codes (not confidential, so kept here)
  internalCodes: [
    { label: "Administration", subject: "Administration", hint: "Day to day running of the organisation, e.g. inbox, expenses, invoicing, filing, diary." },
    { label: "Meetings", subject: "Meetings", hint: "Internal gatherings only, not client meetings." },
    { label: "Projects", subject: "Projects", hint: "Internal build work that is not for a client." },
    { label: "Protected time", subject: "Protected time", hint: "For blocking your diary ahead. If you did client work in it, replace it with a separate External entry." },
    { label: "CPD and networking", subject: "CPD and networking", hint: "Learning and professional development, including learning a new tool." },
    { label: "Sales and business development", subject: "Sales and business development", hint: "Stays Internal even when a client or prospect is named, up to the first billable delivery. Name them in the activity." },
    { label: "Travel time", subject: "Travel time", hint: "Recorded so it can be seen, not charged." }
  ],
  quickActivities: {
    "Administration": ["diary and email management"]
  }
};
