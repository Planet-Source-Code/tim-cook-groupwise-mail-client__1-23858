VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recipient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Recipients"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarEmailAddress As Variant

Public Property Let EmailAddress(vData As Variant)
    mvarEmailAddress = vData
End Property

Public Property Get EmailAddress() As Variant
    EmailAddress = mvarEmailAddress
End Property

