VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recipients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Recipient"
Attribute VB_Ext_KEY = "Member0" ,"Recipient"
Option Explicit

Private mCol As Collection

Public Function Add(EmailAddress As Variant) As Recipient

    Dim oRecipient As Recipient
    Set oRecipient = New Recipient

    oRecipient.EmailAddress = EmailAddress
    
    mCol.Add oRecipient
    Set Add = oRecipient
    Set oRecipient = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As Recipient
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    Dim vItem As Recipient
    Dim l As Long
    
    l = 1
    For Each vItem In mCol
        
        If vItem.EmailAddress = vntIndexKey Then
            mCol.Remove l
        End If
        
        l = l + 1
    Next
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

