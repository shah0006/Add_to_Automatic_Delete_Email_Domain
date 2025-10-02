Option Explicit

' Main procedure: Adds sender domains of selected emails to a blocked list rule
' and moves matching emails to the Deleted Items folder.
Sub AddSelectedEmailsDomainsToBlockedListAndMoveToDeleted()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olSelection As Selection
    Dim olMail As Outlook.MailItem
    Dim olRules As Outlook.Rules
    Dim olRule As Outlook.Rule
    Dim olCondition As Outlook.TextRuleCondition
    Dim senderDomain As String
    Dim currentFolder As Outlook.Folder
    Dim deletedFolder As Outlook.Folder
    Dim item As Object
    Dim domainDict As Object
    Dim tempDomain As Variant
    Dim userConfirmation As VbMsgBoxResult
    Dim ruleName As String
    Dim emailCount As Long

    On Error GoTo ErrorHandler

    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olSelection = olApp.ActiveExplorer.Selection
    Set domainDict = CreateObject("Scripting.Dictionary")
    domainDict.CompareMode = TextCompare

    ' Require at least one email to be selected
    If olSelection.Count = 0 Then
        MsgBox "Please select at least one email to add its domain to the blocked list.", vbExclamation
        Exit Sub
    End If

    Set olRules = olNamespace.DefaultStore.GetRules()

    ' Prompt user for a rule name
    ruleName = InputBox("Enter the name of the rule to update (or create):", _
                        "Rule Name", "Blocked Domains - Delete")
    If ruleName = "" Then Exit Sub

    ' Attempt to get or create the rule
    On Error Resume Next
    Set olRule = olRules.Item(ruleName)
    On Error GoTo ErrorHandler

    If olRule Is Nothing Then
        Set olRule = olRules.Create(ruleName, olRuleReceive)
        With olRule.Actions.MoveToFolder
            .Folder = olNamespace.GetDefaultFolder(olFolderDeletedItems)
            .Enabled = True
        End With
    End If

    ' Ensure Deleted Items folder is available
    Set deletedFolder = olNamespace.GetDefaultFolder(olFolderDeletedItems)
    If deletedFolder Is Nothing Then
        MsgBox "Unable to access the Deleted Items folder. Please try again.", vbExclamation
        Exit Sub
    End If

    ' Access the rule condition
    Set olCondition = olRule.Conditions.SenderAddress
    If olCondition Is Nothing Then
        MsgBox "Unable to access the Sender Address rule condition.", vbCritical
        Exit Sub
    End If

    ' Load existing blocked domains into dictionary
    If olCondition.Enabled Then
        If IsArray(olCondition.Text) Then
            For Each tempDomain In olCondition.Text
                senderDomain = NormalizeDomain(CStr(tempDomain))
                If senderDomain <> "" Then domainDict(senderDomain) = True
            Next tempDomain
        ElseIf TypeName(olCondition.Text) = "String" Then
            For Each tempDomain In Split(olCondition.Text, ";")
                senderDomain = NormalizeDomain(CStr(tempDomain))
                If senderDomain <> "" Then domainDict(senderDomain) = True
            Next tempDomain
        End If
    End If

    ' Extract domains from selected emails
    For Each item In olSelection
        If TypeName(item) = "MailItem" Then
            Set olMail = item
            senderDomain = ExtractSenderDomain(olMail)
            If senderDomain <> "" Then domainDict(senderDomain) = True
        End If
    Next item

    ' Exit if no domains were found
    If domainDict.Count = 0 Then
        MsgBox "No domains were extracted from the selected emails.", vbInformation
        Exit Sub
    End If

    ' Confirm with user before updating the blocked list
    tempDomain = domainDict.Keys
    userConfirmation = MsgBox("The following domains will be added/updated in the blocked list:" & vbNewLine & _
                              JoinStringArray(tempDomain, vbNewLine) & vbNewLine & vbNewLine & _
                              "Do you want to proceed?", vbYesNo + vbQuestion, "Confirm Domain Blocking")

    If userConfirmation = vbNo Then
        MsgBox "Operation cancelled by user.", vbInformation
        Exit Sub
    End If

    ' Update rule condition
    olCondition.Text = domainDict.Keys
    olCondition.Enabled = True

    With olRule.Actions.MoveToFolder
        .Folder = deletedFolder
        .Enabled = True
    End With

    olRules.Save

    ' Move matching emails
    emailCount = MoveEmailsToFolder(olSelection, deletedFolder, domainDict)

    Set currentFolder = olApp.ActiveExplorer.CurrentFolder
    If Not currentFolder Is Nothing Then
        emailCount = emailCount + MoveEmailsToFolder(currentFolder.Items, deletedFolder, domainDict)
    End If

    ' Summary message
    MsgBox "Operation completed successfully:" & vbNewLine & _
           "- " & domainDict.Count & " domain(s) added/updated in the blocked list." & vbNewLine & _
           "- " & emailCount & " email(s) moved to Deleted Items.", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & " (" & Err.Number & ")", vbCritical
End Sub

' Extracts and normalizes sender domain from a mail item
Function ExtractSenderDomain(mail As Outlook.MailItem) As String
    Dim emailAddress As String
    Dim domain As String

    emailAddress = ResolveSmtpAddress(mail)
    If emailAddress <> "" Then
        domain = GetDomainFromEmail(emailAddress)
        ExtractSenderDomain = NormalizeDomain(domain)
    Else
        ExtractSenderDomain = ""
    End If
End Function

' Resolves SMTP address, including Exchange addresses
Function ResolveSmtpAddress(mail As Outlook.MailItem) As String
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    Dim addressEntry As Outlook.AddressEntry
    Dim exchangeUser As Outlook.ExchangeUser
    Dim propertyAccessor As Outlook.PropertyAccessor
    Dim resolvedAddress As String

    resolvedAddress = ""

    On Error Resume Next
    If mail.SenderEmailType = "EX" Then
        Set addressEntry = mail.Sender
        If Not addressEntry Is Nothing Then
            Set exchangeUser = addressEntry.GetExchangeUser
            If Not exchangeUser Is Nothing Then
                resolvedAddress = exchangeUser.PrimarySmtpAddress
            End If
        End If
    End If
    On Error GoTo 0

    If resolvedAddress = "" Then
        On Error Resume Next
        Set propertyAccessor = mail.PropertyAccessor
        If Not propertyAccessor Is Nothing Then
            resolvedAddress = propertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        End If
        On Error GoTo 0
    End If

    If resolvedAddress = "" Then
        resolvedAddress = mail.SenderEmailAddress
    End If

    ResolveSmtpAddress = Trim$(resolvedAddress)
End Function

' Extracts domain from an email address
Function GetDomainFromEmail(emailAddress As String) As String
    Dim domainStartPos As Long

    domainStartPos = InStr(emailAddress, "@")
    If domainStartPos > 0 Then
        GetDomainFromEmail = Mid$(emailAddress, domainStartPos + 1)
    Else
        GetDomainFromEmail = ""
    End If
End Function

' Cleans domain string
Function NormalizeDomain(domain As String) As String
    Dim cleaned As String

    cleaned = LCase$(Trim$(domain))
    If Right$(cleaned, 1) = ">" Then cleaned = Left$(cleaned, Len(cleaned) - 1)
    If Left$(cleaned, 1) = "<" Then cleaned = Mid$(cleaned, 2)
    NormalizeDomain = cleaned
End Function

' Joins array of strings into one string
Function JoinStringArray(values As Variant, Optional delimiter As String = ";") As String
    If IsArray(values) Then
        JoinStringArray = Join(values, delimiter)
    Else
        JoinStringArray = CStr(values)
    End If
End Function

' Moves emails to target folder if sender domain is blocked
Function MoveEmailsToFolder(items As Object, targetFolder As Outlook.Folder, blockedDomains As Object) As Long
    Dim idx As Long
    Dim item As Object
    Dim olMail As Outlook.MailItem
    Dim senderDomain As String
    Dim movedCount As Long

    movedCount = 0

    If TypeName(items) = "Selection" Then
        For Each item In items
            If TypeName(item) = "MailItem" Then
                Set olMail = item
                senderDomain = ExtractSenderDomain(olMail)
                If senderDomain <> "" And blockedDomains.Exists(senderDomain) Then
                    olMail.Move targetFolder
                    movedCount = movedCount + 1
                End If
            End If
        Next item
    ElseIf TypeName(items) = "Items" Then
        For idx = items.Count To 1 Step -1
            Set item = items.Item(idx)
            If TypeName(item) = "MailItem" Then
                Set olMail = item
                senderDomain = ExtractSenderDomain(olMail)
                If senderDomain <> "" And blockedDomains.Exists(senderDomain) Then
                    olMail.Move targetFolder
                    movedCount = movedCount + 1
                End If
            End If
        Next idx
    End If

    MoveEmailsToFolder = movedCount
End Function
