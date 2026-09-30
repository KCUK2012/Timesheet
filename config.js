/*
 * KC Timesheet: settings.
 * This file is published with the add-in, so it must NOT contain client names.
 * Client names are read from the SharePoint list below, using the member of
 * staff's own Microsoft 365 sign-in.
 *
 * IT: fill in the four values marked FILL IN.
 */
window.KC_CONFIG = {
  // Microsoft Entra ID app registration (Entra admin centre > App registrations)
  tenantId: "FILL IN: Directory (tenant) ID",
  clientId: "FILL IN: Application (client) ID",

  // SharePoint list holding the client names
  sharePointHost: "kaleidoscopeconsultants.sharepoint.com",
  sitePath: "FILL IN: e.g. /sites/KCcorpserv/KCOps", // leave empty ("") for the root site
  listName: "FILL IN: e.g. Timesheet clients",
  // Column internal names in that list: Title holds the client name as written
  // in the subject; Aliases holds other recognised forms separated by semicolons.
  nameColumn: "Title",
  aliasColumn: "Aliases",

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
