Attribute VB_Name = "basMailClient"
Option Explicit

Public GWApp As Application
Public GWRootAccount As Account
Public gsRecipient As String
Public MyRecipients As Recipients

Public Function isgSendMail(sUserID As String, sPassword As String, sMessage As String, sSubject As String, Optional sFilePath As String = "")
On Error GoTo isgSendMail_Error

Set GWApp = New Application
Dim MyFolder As Folder
Set GWRootAccount = GWApp.Login(sUserID, , sPassword)
Dim MyMessage As Message
Dim sFID As Variant
Dim MyRecipient As Recipient

'Determine wich folder will create the message normally Mailbox
isgSendMail = True

For Each MyFolder In GWApp.RootAccount.AllFolders
    If MyFolder.Name = "Mailbox" Then
        'Return the folder ID
        sFID = MyFolder.FolderID
        Exit For
    End If
Next

'Set the Folder with the Mailbox folder ID
Set MyFolder = GWRootAccount.GetFolder(sFID)

'Start message
Set MyMessage = MyFolder.Messages.Add

'Message Info
MyMessage.Subject = sSubject
MyMessage.BodyText = sMessage

'Optional Settings
MyMessage.FromText = "GW-Mail Client"

'Loop through the objects to return the Email Addresses

For Each MyRecipient In MyRecipients
    'USE THIS LINE FOR INTERNAL NAMEING LIKE JOHN DOE instead of jdoe@domain.com
    
    'MyMessage.Recipients.Add MyRecipient.EmailAddress, "NGW", "egwTo"
    
    MyMessage.Recipients.Add MyRecipient.EmailAddress
Next

If sFilePath <> "" Then
    MyMessage.Attachments.Add sFilePath, egwFile, GetFileFromPath(sFilePath)
End If

MyMessage.Send
'MyMessage.Delete
Set MyMessage = Nothing

isgSendMail_Resume:
    Exit Function
isgSendMail_Error:
    isgSendMail = False
    MsgBox Error$, vbInformation
    Resume isgSendMail_Resume
End Function

Sub Main()
On Error GoTo Main_Error

Set MyRecipients = New Recipients

frmEmailer.Show

Main_Resume:
    Exit Sub
Main_Error:
    MsgBox Error$, vbInformation
    Resume Main_Resume
End Sub
