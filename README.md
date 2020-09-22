<div align="center">

## Easy MAPI mail\!


</div>

### Description

This code will help understanding the use of the MAPI controls
 
### More Info
 
Recipient, CCRecipient, Subject, Message, Attachment

The great thing about using MAPI directly instead of using Outlook's Type library is that it is so much faster, and uses a lot less memory!!

No output

No side effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Conrad](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/conrad.md)
**Level**          |Advanced
**User Rating**    |3.9 (39 globes from 10 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/conrad-easy-mapi-mail__1-6377/archive/master.zip)





### Source Code

```
Public Function SendMAPIMail( _
MsgTo As String, _
Optional CC As String = "", _
Optional Subject As String = "", _
Optional Body As String = "", _
Optional Att As String = "") _
As Boolean
 'Code by Conrad
 'email cbrits@monotix.co.za
 '-----------------------------------------------
 '** PLEASE NOTE!! **
 'You need a form with both
 'controls (MapiMessages and MapiSession) on it
 '
 'Do the following:
 '-----------------
 '  1.Add a form, and name it frmMail.
 '  2.Go to Components...(Project menu) and find
 '   Microsoft MAPI Controls.
 '  3.Check it, and click OK. There will now
 '   be two
 '   new controls on your Control Tab.
 '  4.Add the two new controls to your form.
 '
 '-----------------------------------------------
 On Error GoTo ErrHndl
 Dim MAPISes As MAPISession
 Dim MAPIMsgs As MAPIMessages
 Screen.MousePointer = 11
 'set the objects to the controls of the form
 Set MAPISes = frmMail.MAPISession1
 Set MAPIMsgs = frmMail.MAPIMessages1
 'download new mail = false
 MAPISes.DownLoadMail = False
 'show the logon interface for the mail
 'account = true
 MAPISes.LogonUI = True
 'sign on to selected account
 MAPISes.SignOn
 DoEvents
 'check if logon was successful
 If MAPISes.SessionID = 0 Then
  SendMAPIMail = False
  MsgBox "Error on login to MAPI", _
      vbCritical, "MAPI"
  Exit Function
 End If
 'set the session IDs the same on both objects
 MAPIMsgs.SessionID = MAPISes.SessionID
 'Set the MSgIndex to -1, this needs to be
 'done for the Compose event to work
 MAPIMsgs.MsgIndex = -1
 'compose a new message
 MAPIMsgs.Compose
 'don't show the resolve address interface
 MAPIMsgs.AddressResolveUI = False
 'set the recipient
 MAPIMsgs.RecipIndex = 0
 MAPIMsgs.RecipType = mapToList
 MAPIMsgs.RecipAddress = MsgTo
 'resolve the recipient's email addresses
 MAPIMsgs.ResolveName
 'set the CC recipient
 MAPIMsgs.RecipIndex = 1
 MAPIMsgs.RecipType = mapCcList
 MAPIMsgs.RecipAddress = CC
 'resolve the recipient's email addresses
 MAPIMsgs.ResolveName
 'set the subject
 MAPIMsgs.MsgSubject = Subject
 'set the Message/Body/NoteText
 MAPIMsgs.MsgNoteText = Body
 If Att <> "" Then
  'set an attachment
  MAPIMsgs.AttachmentPathName = Att
 End If
 'send the message
 MAPIMsgs.Send
 'close the current session
 MAPISes.SignOff
 'clear objects
 Set MAPIMsgs = Nothing
 Set MAPISes = Nothing
 SendMAPIMail = True
 Screen.MousePointer = 0
 Exit Function
ErrHndl:
 Set MAPIMsgs = Nothing
 Set MAPISes = Nothing
 Screen.MousePointer = 0
 MsgBox "Error [" & Err & "] " & Error, vbCritical, "MAPI"
 Screen.MousePointer = 11
 On Error Resume Next
 frmMail.MAPISession1.SignOff
 SendMAPIMail = False
 Screen.MousePointer = 0
End Function
```

