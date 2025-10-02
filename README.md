# Add_to_Automatic_Delete_Email_Domain

Add the selected email domain to the list of domains that are automatically deleted.

## What this macro does
- Collects sender domains from the currently selected emails in Outlook.
- Adds those domains to (or creates) a client-side rule that detects sender address domains.
- Ensures the rule moves matching messages to **Deleted Items**.
- Immediately moves any already-present matching messages in the current folder to **Deleted Items**.

## Requirements
- Outlook for Windows (desktop), with VBA enabled.
- Permission to create and edit client-side rules in your default mailbox.
- Macro security set to allow signed macros or temporarily enable macros while installing.

## Install
- Open Outlook.
- Press **ALT+F11** to open the **VBA Editor**.
- In the Project pane, right-click **Project1 (VbaProject.OTM)** → **Insert** → **Module**.
- Paste the contents of `AutoDeleteDomains.bas` into the new module.
- Close the VBA Editor and return to Outlook.

## First run & rule setup
- In Outlook, select one or more emails whose sender domains you want to block.
- Run the macro **AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted**:
  - Press **ALT+F8**, choose the macro, then **Run**.
- When prompted for a rule name, accept the default `Blocked Domains - Delete` or enter your own.
- Review the list of domains and confirm.
- The rule will be saved and messages matching the listed domains will be moved to **Deleted Items**.

## Adding a toolbar/ribbon button
- Right-click the Ribbon → **Customize the Ribbon…**.
- Choose a tab/group (or create a custom group) and click **New Group**.
- From **Choose commands from:** select **Macros**, pick **Project1.AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted**.
- Click **Add >>**, optionally rename and assign an icon, then **OK**.

## Keyboard shortcut (Quick Access Toolbar)
- Right-click the Ribbon → **Customize Quick Access Toolbar…**.
- From **Choose commands from:** select **Macros**, add the macro to the QAT.
- Use **ALT** plus the QAT index number to trigger it quickly.

## Usage
- Select one or more unwanted messages.
- Trigger the macro (button, ALT+F8, or QAT shortcut).
- Confirm the discovered domains.
- The rule is updated; matching messages in the current folder are moved to **Deleted Items**.

## How it determines the domain
- Attempts to resolve the sender’s **primary SMTP address** for Exchange senders.
- Falls back to the item’s **SenderEmailAddress** if necessary.
- Extracts the string after `@`, trims `< >`, and normalizes to lowercase.

## Unblocking or editing domains
- File → **Manage Rules & Alerts**.
- Open the chosen rule (default `Blocked Domains - Delete`).
- Edit the **Sender’s address includes** condition list to remove or change domains.
- Save the rule.

## Notes & limitations
- Client-side rule; Outlook must be running for it to act on incoming mail.
- Very large domain lists can slow rule evaluation; consider curating occasionally.
- Some bulk senders use subdomains or changing MAIL FROM domains; you may need to block the parent domain thoughtfully.
- Organization policies or add-ins may restrict macros or user-created rules.
- The macro processes the **current folder** for immediate cleanup; run again in other folders if needed.

## Troubleshooting
- Macro disabled: check **File → Options → Trust Center → Trust Center Settings → Macro Settings**.
- “Cannot access Deleted Items”: confirm the default store and folder are available.
- No domains extracted: verify the items are MailItems and not meeting invites or reports.
- Rule not triggering on new mail: confirm it’s enabled and appears in **Rules & Alerts**; ensure Outlook remains open.

## Updating or removing the macro
- Press **ALT+F11** to open the **VBA Editor** and update the module.
- Export a backup by right-clicking the module → **Export File…** to save `AutoDeleteDomains.bas`.
- Remove by deleting the module from the VBA project and saving.
