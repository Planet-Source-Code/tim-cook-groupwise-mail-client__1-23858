VERSION 5.00
Begin VB.Form frmAddRecipient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Recipient to Send List:"
   ClientHeight    =   1470
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Add Recipient:"
      Height          =   960
      Left            =   600
      TabIndex        =   3
      Top             =   15
      Width           =   4890
      Begin VB.TextBox txtRecipient 
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
         TabIndex        =   0
         Top             =   465
         Width           =   4725
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Contact Name or Full Email address:"
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
         Left            =   105
         TabIndex        =   4
         Top             =   240
         Width           =   3930
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   4230
      TabIndex        =   2
      Top             =   1035
      Width           =   1245
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   2970
      TabIndex        =   1
      Top             =   1035
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "John Doe or johnDoe@domain.com"
      Height          =   240
      Left            =   210
      TabIndex        =   5
      Top             =   1110
      Width           =   2550
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   45
      Picture         =   "frmAddRecipient.frx":0000
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frmAddRecipient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub OKButton_Click()
    If Me.txtRecipient <> "" Then
        MyRecipients.Add Me.txtRecipient
        Unload Me
    Else
        MsgBox "Please enter an email address or press cancel to close!", vbInformation
    End If
End Sub
