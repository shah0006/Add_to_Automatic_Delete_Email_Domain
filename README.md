---
title: Add_to_Automatic_Delete_Email_Domain
description: Outlook VBA macro that adds the domains of selected messages to a rule and moves matching mail to Deleted Items.
modified: 2025-10-01
---

# Add_to_Automatic_Delete_Email_Domain

Visual Basic for Applications (VBA) macro for Microsoft Outlook that:
- Collects sender **domains** from the currently selected emails.
- Adds those domains to (or creates) a **client-side rule** matching “sender’s address includes”.
- Ensures the rule **moves** matching messages to **Deleted Items**.
- Immediately **sweeps the current folder** to move existing matches.

## Files
- `AutoDeleteDomains.bas` — Module containing the macro and helper functions.
  - Entry point macro: `AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted`

## Requirements
- Outlook for Windows (desktop), with VBA enabled.
- Ability to create and edit **client-side rules** in the default mailbox store.
- Macro security configured to allow running your code.
- Outlook must be running for the client-side rule to act on incoming mail.

## Install
- Open Outlook.
- Press **ALT+F11** to open the **VBA Editor**.
- In the Project pane, right-click your VBA project → **Insert → Module**.
- Paste the contents of `AutoDeleteDomains.bas` into the new module.
- Save, then return to Outlook.

## Quick start
- Select one or more messages from senders you want to filter.
- Run the macro **AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted**:
  - Press **ALT+F8**, choose the macro, then **Run**.
- When prompted, accept the default rule name **Blocked Domains – Delete** or enter a different name.
- Review the detected domains and confirm.
- The rule is saved, enabled, and messages from those domains are moved to **Deleted Items**.

## Optional: Ribbon / QAT button
- Right-click the Ribbon → **Customize the Ribbon…** → create/select a custom group.
- From **Choose commands from:** pick **Macros**, select:
  - `Project1.AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted`
- Click **Add >>**, rename/icon as desired, then **OK**.
- For Quick Access Toolbar (QAT): **Customize Quick Access Toolbar…** and add the same macro.
  - Trigger via **ALT** + the QAT index number.

## How it works
- Extracts the true SMTP sender using:
  - Exchange Primary SMTP (`AddressEntry.GetExchangeUser.PrimarySmtpAddress`)
  - MAPI property `PR_SMTP_ADDRESS`
  - Fallback to `MailItem.SenderEmailAddress`
- Parses and normalizes the **domain** (string after `@`, trimmed of `< >`, lower-cased).
- Loads any **existing domains** already present in the rule condition and merges them (no duplicates).
- Updates the rule condition **Sender’s address includes** with the merged domain list.
- Enables and binds the rule’s **Move to folder** action to **Deleted Items**.
- Moves any matching messages in:
  - The current selection.
  - The **current folder** (sweeping from bottom up for stability).

## Managing / editing the rule
- Open **File → Manage Rules & Alerts**.
- Locate the chosen rule (default **Blocked Domains – Delete**).
- Edit the **Sender’s address includes** list to add/remove domains.
- Re-run the macro anytime to append more domains from selected messages.
- Re-order the rule if needed so it runs **before** other rules that might stop processing.

## Notes and limitations
- This is a **client-side** rule; Outlook must be open for new mail to be moved.
- Very large domain lists may slow rule evaluation; curate occasionally.
- Many bulk senders rotate subdomains. Matching on the parent domain (e.g., `example.com`) typically catches subdomains because the rule checks for **“words in the sender’s address”** as a substring.
- Some internal or system-generated items may not expose an SMTP domain; these are skipped.
- Organization policy, add-ins, or shared mailbox setups may constrain rules/macros.

## Troubleshooting
- **Macro doesn’t run**: check **File → Options → Trust Center → Trust Center Settings → Macro Settings** and enable a safe option (e.g., “Notifications for all macros”).
- **Deleted Items not found**: confirm the default store is your primary mailbox and the Deleted Items folder exists and is accessible.
- **No domains extracted**: ensure the selection contains **MailItem** messages (not reports/meetings); check that the sender has a resolvable SMTP address.
- **Rule not triggering**: verify it’s enabled in **Rules & Alerts**, appears early enough in the rule order, and that Outlook stays open.

## Updating / removing
- Press **ALT+F11** to edit the module.
- Export a backup: right-click the module → **Export File…** to save `AutoDeleteDomains.bas`.
- Remove by deleting the module in the VBA project and saving.
- Remove or disable the rule in **Rules & Alerts**.

## Security
- Keep macro security enabled; only run code you trust.
- Review the domain list periodically to avoid over-blocking legitimate senders.

## License
- Suggested: MIT License for the macro/module. Add a `LICENSE` file if publishing publicly.

## Acknowledgments
- Built with Outlook’s Rules and MAPI **PropertyAccessor** patterns for robust SMTP resolution.

## Macro reference (for discoverability)
```vb
' Entry point
Sub AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted()
