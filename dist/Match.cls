VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oSubMatches As SubMatches
Private mValue As String
Private mFirstIndex As Long
Private mLength As Long

Public Property Get SubMatches() As SubMatches
    Set SubMatches = oSubMatches
End Property

Public Property Get Value() As String
    Value = mValue
End Property
Friend Property Let Value(ByVal pValue As String)
    mValue = pValue
End Property

Public Property Get FirstIndex() As Long
    FirstIndex = mFirstIndex
End Property
Friend Property Let FirstIndex(ByVal pFirstIndex As Long)
    mFirstIndex = pFirstIndex
End Property

Public Property Get Length() As Long
    Length = mLength
End Property
Friend Property Let Length(ByVal pLength As Long)
    mLength = pLength
End Property

Private Sub Class_Initialize()
    Set oSubMatches = New SubMatches
End Sub

Private Sub Class_Terminate()
    Set oSubMatches = Nothing
End Sub
