VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Wrapper for StaticRegex (https://github.com/sihlfall/vba-regex)
'designed as drop-in replacement (with extended functionality) for MS VBScript_RegExp_55.RegExp
'contributed by GCUser99 (https://github.com/GCuser99) on Sept 17, 2024
'contributed by yaju (https://github.com/yaju) on March 16, 2025
'TO DO: add named captures to SubMatches collection (using Key?)

Private mGlobal As Boolean
Private mMultiLine As Boolean
Private mIgnoreCase As Boolean
Private mPattern As String

Public Property Let globalFlag(ByVal pGlobal As Boolean)
    mGlobal = pGlobal
End Property
Public Property Get globalFlag() As Boolean
    globalFlag = mGlobal
End Property

Public Property Let multiLine(ByVal pMultiLine As Boolean)
    mMultiLine = pMultiLine
End Property
Public Property Get multiLine() As Boolean
    multiLine = mMultiLine
End Property

Public Property Let ignoreCase(ByVal pIgnoreCase As Boolean)
    mIgnoreCase = pIgnoreCase
End Property
Public Property Get ignoreCase() As Boolean
    ignoreCase = mIgnoreCase
End Property

Public Property Let pattern(ByVal pPattern As String)
    mPattern = pPattern
End Property
Public Property Get pattern() As String
    pattern = mPattern
End Property

Public Function Execute(ByVal sourceString As String) As MatchCollection

    Dim matcherState As StaticRegex.MatcherStateTy
    Dim regex As StaticRegex.RegexTy
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim i As Long
    Dim submatchString As String
    Dim matchStart As Long, matchLength As Long
    
    StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase

    StaticRegex.InitializeMatcherState matcherState, Not mGlobal, mMultiLine
    
    Set oMatches = New MatchCollection
    
    Do While StaticRegex.MatchNext(matcherState, regex, sourceString)
        With matcherState.captures
            Set oMatch = New Match
            matchStart = .entireMatch.start
            matchLength = .entireMatch.Length
            If matchStart > 0 Then
                oMatch.Value = Mid$(sourceString, matchStart, matchLength)
                oMatch.FirstIndex = matchStart - 1 'in VBScript_RegExp_55 this is zero based
                oMatch.Length = matchLength
            Else
                oMatch.Value = vbNullString
            End If

            'could add submatch via key associated with named group
            For i = 0 To .nNumberedCaptures - 1
                matchStart = .numberedCaptures(i).start
                matchLength = .numberedCaptures(i).Length
                If matchStart > 0 Then submatchString = Mid$(sourceString, matchStart, matchLength) Else submatchString = vbNullString
                oMatch.SubMatches.Add submatchString
            Next i
            'to do: need to add named captures to SubMatches collection (using Key?)
            'or alternatively, to another exposed additional collection such as "SubMatchesNamed"
            oMatches.Add oMatch
        End With
    Loop

    Set Execute = oMatches
End Function

Public Function Test(ByVal sourceString As String) As Boolean
    Dim regex As StaticRegex.RegexTy
    StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase
    Test = StaticRegex.Test(regex, sourceString, mMultiLine)
End Function

Public Function Replace(ByVal sourceString As String, ByVal replaceString As String) As String
    Dim regex As StaticRegex.RegexTy
    StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase
    Replace = StaticRegex.Replace(regex, replaceString, sourceString, Not mGlobal, mMultiLine)
End Function

Private Sub Class_Initialize()
    mGlobal = False 'StaticRegex defaults to True
    mMultiLine = False
    mIgnoreCase = False
End Sub

Public Function Split(ByVal sourceString As String) As Collection
    Dim regex As StaticRegex.RegexTy
    StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase
    Set Split = StaticRegex.Split(regex, sourceString, Not mGlobal, mMultiLine)
End Function
