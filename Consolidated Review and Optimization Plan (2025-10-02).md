---
title: Obsidian Note — Add_to_Automatic_Delete_Email_Domain (Consolidated Review & Optimization Plan)
tags: [Outlook, VBA, rules, repo-review, documentation, optimization, security]
modified: 2025-10-02
---

# Repository Snapshot
- Contents: `AutoDeleteDomains.bas` (VBA module) and `README.md` (documentation).
- README front matter is rendered visibly as a table-like block at the top (and/or bottom), creating visual noise; full usage and workflow documentation follows.
- Macro entry point: `AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted`, which gathers domains from selected Outlook messages, updates/creates a client-side rule moving matches to **Deleted Items**, then sweeps both the **selection** and the **active folder**.

# Repository Status (from Reviewer #1)
- Public repo confirmed. Exactly two files: `AutoDeleteDomains.bas`, `README.md` (see References).
- README covers purpose, install, quick start, how it works, troubleshooting, and an “Overview & Workflow” section.

## Strengths
- Clear single-purpose macro with a documented entry point and helper routines (also reflected in module comments).
- Installation and run steps are actionable (ALT+F11/ALT+F8; rule-naming flow; security reminders).
- SMTP extraction approach is sound and documented: **Exchange Primary SMTP → `PR_SMTP_ADDRESS` → `SenderEmailAddress`** with normalization helpers.

# Gaps, Weaknesses & Proposed Solutions
1. Visible front matter metadata in README  
	 - **Issue:** YAML-like metadata is rendered as plain text at top/bottom.  
	 - **Solution:** Convert to proper YAML front matter (`---` … `---`) recognized by GitHub or remove entirely for a cleaner presentation. Optionally move the Workflow/Overview section into the main body.

2. No explicit license file  
	 - **Issue:** Repo ships only macro module + README; README suggests adding a license but none is present.  
	 - **Solution:** Add `LICENSE` (e.g., MIT) and surface license details in README (“License” section).

3. Compatibility guidance incomplete  
	 - **Issue:** README states “Outlook for Windows (desktop)” but does not warn that the **New Outlook** for Windows **lacks VBA/COM support**.  
	 - **Solution:** Add an early **Compatibility** callout: “Works in **classic Outlook for Windows**; **not supported** in **New Outlook** (no VBA/COM).” Reference Microsoft guidance.

4. Security best practices limited  
	 - **Issue:** Macro security is mentioned but emphasis on **signing (SelfCert)** and avoiding **‘Enable all macros’** could be stronger.  
	 - **Solution:** Expand the Security section with SelfCert steps and safe Trust Center settings.

5. No visual onboarding aids  
	 - **Issue:** README lacks a screenshot/GIF showing macro execution or the resulting rule.  
	 - **Solution:** Capture and embed a simple image/GIF of running the macro and the “Rules & Alerts” configuration.

6. Unconditional full-folder sweep  
	 - **Issue:** Macro always processes the entire active folder after updating the rule, which can be slow on large folders and unexpected if the user intended to act only on the selection.  
	 - **Solution:** Prompt before sweeping, allow bypass, and consider progress feedback for large folders.

7. Existing-domain parsing assumes semicolon delimiters  
	 - **Issue:** When loading existing rule entries (string branch), it splits only on `;`, potentially missing comma- or line-delimited entries.  
	 - **Solution:** Normalize delimiters (replace CR/LF with `;`, also split on commas) before trimming/storing domains.

8. SMTP resolution skips Exchange distribution lists (DLs)  
	 - **Issue:** `ResolveSmtpAddress` checks `GetExchangeUser` but not `GetExchangeDistributionList`, so DL senders may fall back to non-SMTP values.  
	 - **Solution:** Extend resolution to DLs (retrieve `PrimarySmtpAddress` when available), guard for COM errors.

9. Performance/scalability considerations  
	 - **Issue:** Large domain lists slow rule evaluation and may hit rule size quotas.  
	 - **Solutions:**  
		 - Optional **caching** of sender→domain during sweeps.  
		 - **Guard rails** (soft max domain count with warning).  
		 - Optional **advanced sweep** (Inbox + subfolders) as an opt-in feature.

10. Audit/export enhancements  
		- **Issue:** No audit trail of the merged blocked list.  
		- **Solution:** Optional export of the final merged domain list to a text file for backup/audit.

11. Additional repo hygiene (from Reviewer #1)  
		- **Suggestion:** Add `/LICENSE` (MIT), `CHANGELOG.md`, and a simple `.gitignore` to avoid committing `VbaProject.OTM` or generated artifacts.

12. Clarify rule semantics (from Reviewer #1)  
		- **Suggestion:** Explicitly state that **“Sender’s address includes”** performs substring matching, so blocking `example.com` will also match `mail.example.com`.

13. International domains (from Reviewer #1)  
		- **Suggestion:** Consider trimming trailing dots; optionally convert IDN → Punycode if internationalized domains become a target.


# Comprehensive Solutions Summary
- **Documentation:** Clean metadata; add Compatibility and Security callouts; include one visual; describe parent-domain matching behavior to reassure users.  
- **Licensing:** Commit MIT license and reference it in README.  
- **Macro Options:** Prompt/config for folder sweep, optional cached lookups to reduce repeated SMTP resolution, and an opt-in **advanced sweep** (Inbox + subfolders).  
- **SMTP Resolution:** Extend to DLs; retain current fallback order; comment rationale in code for maintainability.  
- **Domain Parsing:** Support multiple delimiters (semicolon, comma, CR/LF); trim empties to avoid duplicates.  
- **User Experience:** Continue summarizing domain/email counts; additionally indicate whether a folder sweep was executed; consider simple progress feedback for large folders.  
- **Repo Hygiene:** Add LICENSE, CHANGELOG, and (optionally) screenshots/GIFs.  
- **Operational Guidance:** Emphasize macro signing (SelfCert) and safer Trust Center settings.

# Copy-Ready README Snippets
> **Compatibility**  
> This macro targets **classic Outlook for Windows** (VBA/COM). The **New Outlook for Windows** does **not** support VBA or COM add-ins; this project will not run there.

> **Security**  
> Sign your macro (e.g., with **SelfCert**) and avoid **“Enable all macros.”** Use **“Disable all macros with notification”** and trust only the signed macro you intend to run.

# Actionable Code Patterns (VBA Sketches)
## Optional full-folder sweep
```vb
Dim doFolderSweep As VbMsgBoxResult
doFolderSweep = MsgBox("Also sweep the entire current folder for matches now?", _
                       vbYesNo + vbQuestion, "Optional Folder Sweep")
If doFolderSweep = vbYes Then
    emailCount = emailCount + MoveEmailsToFolder(currentFolder.Items, deletedFolder, domainDict)
End If
````

## Extend SMTP resolution to Distribution Lists
```vb
If mail.SenderEmailType = "EX" Then
    Set addressEntry = mail.Sender
    If Not addressEntry Is Nothing Then
        Set exchangeUser = addressEntry.GetExchangeUser
        If Not exchangeUser Is Nothing Then
            resolvedAddress = exchangeUser.PrimarySmtpAddress
        Else
            Dim exDL As Outlook.ExchangeDistributionList
            Set exDL = addressEntry.GetExchangeDistributionList
            If Not exDL Is Nothing Then
                resolvedAddress = exDL.PrimarySmtpAddress
            End If
        End If
    End If
End If
' Fallbacks remain: PR_SMTP_ADDRESS via PropertyAccessor; then SenderEmailAddress.
```

## Normalize existing-domain delimiters
```vb
Dim raw As String, tokens() As String, t As Variant, d As String
raw = CStr(olCondition.Text)
raw = Replace(raw, vbCrLf, ";")
raw = Replace(raw, vbCr, ";")
raw = Replace(raw, vbLf, ";")
raw = Replace(raw, ",", ";")
Do While InStr(raw, ";;") > 0
    raw = Replace(raw, ";;", ";")
Loop
tokens = Split(raw, ";")
For Each t In tokens
    d = NormalizeDomain(CStr(t))
    If d <> "" Then domainDict(d) = True
Next t
```

## Guard rail for large domain lists
```vb
Const MAX_DOMAINS As Long = 400
If domainDict.Count > MAX_DOMAINS Then
    Dim choice As VbMsgBoxResult
    choice = MsgBox("Domain count (" & domainDict.Count & _
       ") exceeds " & MAX_DOMAINS & ". Proceed?", vbYesNo + vbExclamation, "Large Rule Warning")
    If choice = vbNo Then Exit Sub
End If
```

## Optional sender→domain caching during a run
```vb
Dim domainCache As Object: Set domainCache = CreateObject("Scripting.Dictionary")
Dim key As String, dom As String
key = olMail.EntryID
If domainCache.Exists(key) Then
    dom = domainCache(key)
Else
    dom = ExtractSenderDomain(olMail)
    domainCache(key) = dom
End If
```

## Optional export of merged domain list
```vb
Dim fso As Object, ts As Object, k As Variant
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.CreateTextFile(Environ$("USERPROFILE") & "\Documents\BlockedDomains.txt", True)
For Each k In domainDict.Keys
    ts.WriteLine CStr(k)
Next k
ts.Close
```

# Consolidated Checklist
- [ ] Add `LICENSE` (MIT suggested) and reference it in README
- [ ] Add **Compatibility** callout: Classic Outlook only; New Outlook unsupported
- [ ] Expand **Security**: signing with SelfCert; avoid “Enable all macros”
- [ ] Replace/remove visible metadata table; prefer YAML front matter
- [ ] Link Microsoft docs for `PropertyAccessor` and `PR_SMTP_ADDRESS`
- [ ] Make full-folder sweep optional (prompt)
- [ ] Extend SMTP resolution to **Exchange DLs**
- [ ] Parse multiple delimiters (semicolon, comma, CR/LF)
- [ ] Add soft guard for large domain lists / rule quotas
- [ ] Optional sender→domain caching
- [ ] Optional export of merged domain list
- [ ] Optional **advanced sweep** (Inbox + subfolders)
- [ ] Add screenshot/GIF; consider CHANGELOG and .gitignore
- [ ] Clarify substring matching: blocking `example.com` also blocks subdomains
- [ ] (Optional) Consider trimming trailing dots and IDN→Punycode handling

# Multiple-Choice Questions (from both reviews; choose one unless noted)
1. **Default behavior after updating the rule**
	A) Move only the **selected** messages now
	B) Prompt to **also sweep the current folder** (default **No**)
	C) Prompt to also sweep the current folder (default **Yes**)
	D) Always sweep the current folder **without prompting**

2. **Scope of sweeping (if chosen)** *(Select all that apply)*
	A) Current folder only
	B) Current folder + **subfolders**
	C) **Inbox** only
	D) Inbox + subfolders
	E) Custom folder picker at runtime

3. **SMTP resolution coverage**
	A) Current approach (User via `GetExchangeUser` → `PR_SMTP_ADDRESS` → `SenderEmailAddress`)
	B) **Add distribution lists** via `GetExchangeDistributionList`
	C) Add DLs and attempt header-based expansion if DL SMTP isn’t exposed
	D) Keep current approach; document DL limitation

4. **Existing-domain delimiter handling**
	A) Split on **semicolon** only
	B) Split on semicolon and **comma**
	C) Normalize **CR, LF, comma, semicolon** to a single delimiter and split
	D) Use **RegExp** for tokenization (adds reference)

5. **Large rule guardrail**
	A) Warn if > **200** domains
	B) Warn if > **400** domains
	C) Warn if > **600** domains
	D) No warning

6. **Performance optimization (caching)**
	A) **Enable caching** within a run (EntryID → domain)
	B) Skip caching (simplicity preferred)

7. **Audit/backup of blocked domains**
	A) **Write merged list** to a text file after updates
	B) Do not write to disk (privacy/IT policy concerns)
	C) Only write to disk **on user confirmation**

8. **Compatibility messaging in README**
	A) Prominent **top-level callout**
	B) Mentioned later in README
	C) Omit compatibility note

9. **Security guidance level**
	A) Basic tips (sign macros; avoid “Enable all macros”)
	B) Detailed steps (SelfCert walkthrough; Trust Center screenshots)
	C) No security guidance

10. **User-experience aids in README** *(Select all that apply)*
	A) **Screenshot/GIF** of running the macro
	B) Step-by-step **Ribbon/QAT button** setup
	C) **Troubleshooting** table (symptom → fix)
	D) “How SMTP is resolved” mini-diagram

11. **Project sequencing priority**
	A) Documentation/licensing first
	B) License/compat + code changes in parallel
	C) Code optimizations first
	D) Need stakeholder input

12. **Advanced sweeps (Inbox + subfolders)**
	A) Yes—high-priority addition
	B) Maybe—offer as opt-in
	C) No—scope creep / avoid
	D) Need to survey workflows

# Additional Enhancement Ideas
- Clarify in README that blocking `example.com` also blocks subdomains because the rule uses substring matching (echoing existing notes).
- Encourage periodic curation of the blocked list and mention potential **rule size quotas** in Outlook/Exchange.
- Consider a server-side alternative (transport rules) for environments that disallow VBA, noting behavior differences vs client rules.
- If IDN support becomes relevant, evaluate **Punycode** conversion during normalization.

# References
- GitHub repository: [https://github.com/shah0006/Add_to_Automatic_Delete_Email_Domain](https://github.com/shah0006/Add_to_Automatic_Delete_Email_Domain)
- Microsoft Q&A — New Outlook & VBA/COM add-ins: [https://learn.microsoft.com/en-us/answers/questions/4627044/new-outlook-vba](https://learn.microsoft.com/en-us/answers/questions/4627044/new-outlook-vba)
- Microsoft Learn — Get SMTP address of mail item sender: [https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-smtp-address-of-the-sender-of-a-mail-item](https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-smtp-address-of-the-sender-of-a-mail-item)
- Microsoft Learn — Obtain email address of a recipient (`PR_SMTP_ADDRESS`): [https://learn.microsoft.com/en-us/office/vba/outlook/concepts/address-book/obtain-the-e-mail-address-of-a-recipient](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/address-book/obtain-the-e-mail-address-of-a-recipient)
