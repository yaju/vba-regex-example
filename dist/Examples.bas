Attribute VB_Name = "Example"
Option Explicit

Public Sub main()

    Debug.Print "--- DefaultExample ---"
    
    DefaultExample
    
    Debug.Print "--- WrapperExample ---"
    
    WrapperExample

End Sub

Public Sub DefaultExample()
    Dim re As StaticRegex.RegexTy
    Dim matcherState As MatcherStateTy
    Dim sourceString As String
    Dim i As Long
    Dim ignoreCase As Boolean, globalFlag As Boolean, multiLine As Boolean
    Dim submatchString As String
    Dim matchStart As Long, matchLength As Long

    globalFlag = True
    multiLine = False
    ignoreCase = False
    StaticRegex.InitializeRegex re, "(€D+)(€d+)", ignoreCase
    StaticRegex.InitializeMatcherState matcherState, Not globalFlag, multiLine
    sourceString = "Abc123DEFGH4567ijkl890"

    Do While StaticRegex.MatchNext(matcherState, re, sourceString)
        With matcherState.captures
            For i = 0 To .nNumberedCaptures - 1
                matchStart = .numberedCaptures(i).start
                matchLength = .numberedCaptures(i).Length
                submatchString = vbNullString
                If matchStart > 0 Then submatchString = Mid$(sourceString, matchStart, matchLength)
                Debug.Print submatchString
            Next
        End With
   Loop

End Sub

Public Sub WrapperExample()

    Dim re As New RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim i As Long

    re.pattern = "(€D+)(€d+)"
    re.globalFlag = True
    Set mc = re.Execute("Abc123DEFGH4567ijkl890")
    For Each m In mc
        For i = 0 To m.SubMatches.Count - 1
            Debug.Print m.SubMatches(i)
        Next
    Next

End Sub
