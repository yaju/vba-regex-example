VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colSubMatches As Collection

Friend Sub Add(ByVal val As String)
    colSubMatches.Add val
End Sub

Public Property Get Count() As Long
   Count = colSubMatches.Count
End Property

Public Property Get Item(ByVal pItemIndex As Variant) As String
Attribute Item.VB_UserMemId = 0
    Item = colSubMatches.Item(pItemIndex + 1)
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colSubMatches.[_NewEnum]
End Function

Private Sub Class_Terminate()
    Set colSubMatches = Nothing
End Sub

Private Sub Class_Initialize()
    Set colSubMatches = New Collection
End Sub
