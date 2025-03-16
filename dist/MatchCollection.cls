VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colMatches As Collection

Friend Sub Add(obj As Match)
    colMatches.Add obj
End Sub

Public Property Get Count() As Long
   Count = colMatches.Count
End Property

Public Property Get Item(ByVal pItemIndex As Variant) As Match
Attribute Item.VB_UserMemId = 0
    Set Item = colMatches.Item(pItemIndex + 1)
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colMatches.[_NewEnum]
End Function

Private Sub Class_Terminate()
    'RemoveAll
    Set colMatches = Nothing
End Sub

Private Sub Class_Initialize()
    Set colMatches = New Collection
End Sub
