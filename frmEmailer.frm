VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEmailer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Groupwise Email Client by Tim Cook"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frmEmailer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2085
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Email Info:"
      Height          =   3270
      Left            =   1200
      TabIndex        =   13
      Top             =   690
      Width           =   6990
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   405
         Left            =   5565
         TabIndex        =   5
         Top             =   2805
         Width           =   1365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   405
         Left            =   5565
         TabIndex        =   4
         Top             =   2370
         Width           =   1365
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   60
         TabIndex        =   3
         Top             =   2385
         Width           =   5445
      End
      Begin VB.TextBox txtMessage 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1020
         Width           =   6840
      End
      Begin VB.TextBox txtSubject 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         TabIndex        =   1
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label Label7 
         Caption         =   "255 Char Limit"
         Height          =   210
         Left            =   1530
         TabIndex        =   20
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Recipients: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   75
         TabIndex        =   17
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label Label4 
         Caption         =   "Message: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   16
         Top             =   795
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Subject: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   105
         TabIndex        =   15
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   420
      Left            =   6675
      TabIndex        =   8
      Top             =   4860
      Width           =   1560
   End
   Begin VB.Frame Frame2 
      Caption         =   "Include File:"
      Height          =   810
      Left            =   1215
      TabIndex        =   12
      Top             =   3975
      Width           =   7020
      Begin VB.TextBox txtFilePath 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   4290
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   390
         Left            =   5565
         TabIndex        =   6
         Top             =   285
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Attatchment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   75
         TabIndex        =   19
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Login:"
      Height          =   630
      Left            =   1200
      TabIndex        =   9
      Top             =   45
      Width           =   6990
      Begin VB.TextBox txtPassword 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4455
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   255
         Width           =   2055
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Password: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3405
         TabIndex        =   11
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "User Name: *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "&Send Email"
      Height          =   420
      Left            =   5070
      TabIndex        =   7
      Top             =   4860
      Width           =   1560
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5265
      Left            =   45
      Picture         =   "frmEmailer.frx":030A
      Top             =   60
      Width           =   1140
   End
End
Attribute VB_Name = "frmEmailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'SAMPLE MAIL CLIENT

'With little modification you can use in Access and Where needed to
'work with groupwise as opposed to MAPI.

'Author: Tim Cook
'Do not change this example only work from it.


Private Sub cmdAdd_Click()
On Error GoTo cmdAdd_Click_Error

    frmAddRecipient.Show 1

cmdAdd_Click_Resume:
    Exit Sub
cmdAdd_Click_Error:
    MsgBox Error$, vbInformation
    Resume cmdAdd_Click_Resume
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo cmdBrowse_Click_Error

Dim sRetVal As String

Me.CommonDialog1.DialogTitle = "Select file to Send...."

Me.CommonDialog1.Filter = "ShowAll(*)"
Me.CommonDialog1.ShowOpen
sRetVal = Me.CommonDialog1.FileName

If sRetVal <> "" Then
    Me.txtFilePath = sRetVal
End If

cmdBrowse_Click_Resume:
    Exit Sub
cmdBrowse_Click_Error:
    MsgBox Error$, vbInformation
    Resume cmdBrowse_Click_Resume
End Sub

Private Sub cmdRemove_Click()
On Error GoTo cmdRemove_Click_Error

Dim MyRecipient As Recipient


If Me.List1.Text <> "" Then
    For Each MyRecipient In MyRecipients
        If MyRecipient.EmailAddress = Me.List1.Text Then
            MyRecipients.Remove MyRecipient.EmailAddress
            isgUpdateList
            Exit For
        End If
    Next
End If

cmdRemove_Click_Resume:
    Exit Sub
cmdRemove_Click_Error:
    MsgBox Error$, vbInformation
    Resume cmdRemove_Click_Resume
End Sub

Private Sub cmdSendMail_Click()
On Error GoTo cmdSendMail_Click_Error

frmEmailer.MousePointer = vbHourglass

If MyRecipients.Count Then
    If Me.txtMessage <> "" And Me.txtSubject <> "" And Me.txtUserName <> "" And Me.txtPassword <> "" Then
        If isgSendMail(Me.txtUserName, Me.txtPassword, Me.txtMessage, Me.txtFilePath, Me.txtFilePath) Then
            MsgBox "Message has been sent...", vbInformation
            Set MyRecipients = Nothing
            Set MyRecipients = New Recipients
            
            If MsgBox("Would you like to send another email?", vbInformation + vbYesNo, "Send Another?") = vbYes Then
                Me.txtFilePath = ""
                Me.txtSubject = ""
                Me.txtMessage = ""
                Me.List1.Clear
                Me.txtSubject.SetFocus
            Else
                Command2_Click
            End If
            
        End If
    Else
        MsgBox "Please enter the required fields!", vbInformation
    End If
Else
    MsgBox "Please enter User(s) to send the email to.", vbInformation
    Me.cmdAdd.SetFocus
End If

cmdSendMail_Click_Resume:
    frmEmailer.MousePointer = vbNormal
    Exit Sub
cmdSendMail_Click_Error:
    MsgBox Error$, vbInformation
    Resume cmdSendMail_Click_Resume
End Sub

Private Sub Command2_Click()
    Set MyRecipients = Nothing
    End
End Sub

Private Sub Form_Activate()
On Error GoTo Form_Activate_Error

isgUpdateList

Form_Activate_Resume:
    Exit Sub
Form_Activate_Error:
    MsgBox Error$, vbInformation
    Resume Form_Activate_Resume
End Sub

Sub isgUpdateList()
On Error GoTo isgUpdateList_Error

Dim MyRecipient As Recipient

Me.List1.Clear

For Each MyRecipient In MyRecipients
    Me.List1.AddItem MyRecipient.EmailAddress
Next

isgUpdateList_Resume:
    Exit Sub
isgUpdateList_Error:
    MsgBox Error$, vbInformation
    Resume isgUpdateList_Resume
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.txtUserName = SystemLogonName
End Sub
