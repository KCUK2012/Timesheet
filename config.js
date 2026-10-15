/*
 * Outlook Timesheet: settings.
 * This file is published with the add-in, so it must NOT contain client names.
 * Client names and internal codes are read from Outlook Timesheet Source Lookups.xlsx in
 * SharePoint using the member of staff's own Microsoft 365 sign-in.
 */
window.KC_CONFIG = {
  // Microsoft Entra ID
  tenantId: "d15beb8b-78b2-4fe4-9ead-0d1204377b43",   // Directory (tenant) ID: confirmed for kdpc.uk
  clientId: "5c0081ab-6727-49c1-8b6e-ede656d392fd",   // Application (client) ID: confirmed

  // Where the lists live in SharePoint
  sharePointHost: "kaleidoscopeconsultants.sharepoint.com",
  sitePath: "",                        // "" = main KC site (KC Operations home)
  source: "file",                      // "file" = .xlsx or .csv in a document library; "list" = SharePoint list
  // Path inside the site's Documents library (do not include "Shared Documents")
  filePath: "/Policies, Guidelines, Handbooks and ISO's/Timesheet guidance/Outlook Timesheet Source Lookups.xlsx",
  // Tab names in the workbook
  sheets: { clients: "Clients", internalCodes: "Internal codes", projects: "Projects" },
  listName: "",                        // only used when source is "list"

  // Column headings on the Clients tab (row 1)
  nameColumn: "Title",                 // client name exactly as written in the subject
  aliasColumn: "Aliases",              // other recognised names, separated by semicolons

  // Subject format
  separator: " – ",

  // Work type for External entries: one must be chosen. Written into the subject after the client.
  workTypes: ["Subscription Service", "Project Work"],
  workTypeAliases: { "Full Service": "Subscription Service" },  // older names still recognised when an entry is reopened

  // Always shown last in the Internal code list. The Activity box (mandatory) describes the work:
  // "Internal – Option not available – <activity>".
  otherCode: { label: "Option not available" },

  // Load attempts before showing an error (the Reload lists button tries again)
  loadAttempts: 2,

  // Internal codes: used only if the workbook has no "Internal codes" tab
  internalCodes: [
    { label: "Administration and queries", subject: "Administration and queries", hint: "" },
    { label: "Meeting", subject: "Meeting", hint: "" },
    { label: "Project", subject: "Project", hint: "" },
    { label: "Placeholder (keep free)", subject: "Placeholder (keep free)", hint: "", nft: true },
    { label: "Sales and business development", subject: "Sales and business development", hint: "" },
    { label: "Learning and events", subject: "Learning and events", hint: "" },
    { label: "Business Services", subject: "Business Services", hint: "" },
    { label: "People Management", subject: "People Management", hint: "" },
    { label: "Research and AI", subject: "Research and AI", hint: "" },
    { label: "Not available", subject: "Not available", hint: "" }
  ]
};
