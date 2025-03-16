Attribute VB_Name = "StaticRegex"
'
' VbaRegex
' by Sihlfall
' MIT license
'
' A regular expression engine written entirely in VBA.
'
' Standalone standard-module version.
'
' This file has been generated from the source files by a build script (make_aio.ps1).
' IF YOU NEED TO MAKE CHANGES, CONSIDER EDITING THE SOURCE FILES, NOT THIS FILE.
'
' Documentation, source code, tests, and build script are available at
'
' https://github.com/sihlfall/vba-regex
'
Option Private Module
Option Explicit

Private Enum StaticStringBuilderConstant
    ' Must be at least 2.
    STATIC_STRING_BUILDER_DEFAULT_MINIMUM_CAPACITY = 16
End Enum

Private Type StaticStringBuilder
    Active As Integer            ' index of the currently active buffer (0 or 1)
    Buffer(0 To 1) As String     ' .buffer(.active) is the currently active buffer
    Capacity As Long             ' current allocated capacity in characters
    Length As Long               ' current length of the string, in characters
    MinimumCapacity As Long      ' minimum capacity set (0 or >= 2)
End Type

Private Type ArrayBuffer
    Buffer() As Long
    Capacity As Long
    Length As Long
End Type

Private Enum ArrayBufferConstant
    ARRAY_BUFFER_MINIMUM_CAPACITY = 16
End Enum

Private Enum NumericConstantLong
    LONG_FIRST_BIT = &H80000000
    LONG_ALL_BUT_FIRST_BIT = Not LONG_FIRST_BIT
    LONG_MIN = &H80000000
    LONG_MAX = &H7FFFFFFF
    LONG_MAX_DIV_10 = LONG_MAX \ 10
End Enum

Private Enum RegexErrType
    REGEX_ERR = vbObjectError + 3000
    
    REGEX_ERR_INVALID_REGEXP_ESCAPE = REGEX_ERR + 1
    REGEX_ERR_INVALID_ESCAPE = REGEX_ERR + 2
    REGEX_ERR_INVALID_RANGE = REGEX_ERR + 3
    REGEX_ERR_UNTERMINATED_CHARCLASS = REGEX_ERR + 4
    REGEX_ERR_INVALID_REGEXP_GROUP = REGEX_ERR + 6
    REGEX_ERR_INVALID_REGEXP_CHARACTER = REGEX_ERR + 7
    REGEX_ERR_INVALID_QUANTIFIER = REGEX_ERR + 8
    REGEX_ERR_UNEXPECTED_END_OF_PATTERN = REGEX_ERR + 9
    REGEX_ERR_UNEXPECTED_REGEXP_TOKEN = REGEX_ERR + 10
    REGEX_ERR_UNEXPECTED_CLOSING_PAREN = REGEX_ERR + 11
    REGEX_ERR_INVALID_QUANTIFIER_NO_ATOM = REGEX_ERR + 12
    REGEX_ERR_INVALID_IDENTIFIER = REGEX_ERR + 13
    
    REGEX_ERR_INTERNAL_LOGIC_ERR = REGEX_ERR + 20
    
    REGEX_ERR_INVALID_FLAG_ERR = REGEX_ERR + 30
    REGEX_ERR_INVALID_REPLACEMENT_STRING = REGEX_ERR + 31
    
    REGEX_ERR_INVALID_INPUT = REGEX_ERR + 50
End Enum

Private Enum BytecodeDescriptionConstant
    BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT = 0
    BYTECODE_IDX_N_IDENTIFIERS = 1
    BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR = 2
    BYTECODE_IDENTIFIER_MAP_BEGIN = 3
    BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE = 3
    BYTECODE_IDENTIFIER_MAP_ENTRY_START_IN_PATTERN = 0
    BYTECODE_IDENTIFIER_MAP_ENTRY_LENGTH_IN_PATTERN = 1
    BYTECODE_IDENTIFIER_MAP_ENTRY_ID = 2
    
    ' Todo: Introduce special value or restrict max. explicit quantifier value to LONG_MAX - 1
    RE_QUANTIFIER_INFINITE = &H7FFFFFFF
End Enum

Private Enum ReOpType
    REOP_MATCH = 1
    REOP_CHAR = 2
    REOP_PERIOD = 3
    REOP_RANGES = 4 ' nranges [must be >= 1], chfrom, chto, chfrom, chto, ...
    REOP_INVRANGES = 5
    REOP_JUMP = 6
    REOP_SPLIT1 = 7 ' prefer direct
    REOP_SPLIT2 = 8 ' prefer jump
    REOP_SAVE = 11
    REOP_SET_NAMED = 12 ' id, capture num
    REOP_LOOKPOS = 13
    REOP_LOOKNEG = 14
    REOP_BACKREFERENCE = 15
    REOP_ASSERT_START = 16
    REOP_ASSERT_END = 17
    REOP_ASSERT_WORD_BOUNDARY = 18
    REOP_ASSERT_NOT_WORD_BOUNDARY = 19
    REOP_REPEAT_EXACTLY_INIT = 20 ' <none>
    REOP_REPEAT_EXACTLY_START = 21 ' quantity [must be >= 1], offset
    REOP_REPEAT_EXACTLY_END = 22 ' quantity [must be >= 1], offset
    REOP_REPEAT_MAX_HUMBLE_INIT = 23 ' <none>
    REOP_REPEAT_MAX_HUMBLE_START = 24 ' quantity, offset
    REOP_REPEAT_MAX_HUMBLE_END = 25 ' quantitiy, offset
    REOP_REPEAT_GREEDY_MAX_INIT = 26 ' <none>
    REOP_REPEAT_GREEDY_MAX_START = 27 ' quantity, offset
    REOP_REPEAT_GREEDY_MAX_END = 28 ' quantitiy, offset
    REOP_CHECK_LOOKAHEAD = 29 ' <none>
    REOP_CHECK_LOOKBEHIND = 30 ' <none>
    REOP_END_LOOKPOS = 31 ' <none>
    REOP_END_LOOKNEG = 32 ' <none>
    REOP_FAIL = 33
End Enum

Private Enum AstNodeType
    ' We guarantee: All AST_ constants are > 0.
    ' ! The parser relies on AST_ASSERT_LOOKAHEAD and LOOKBEHIND being > 0.
    MIN_AST_CODE = 0
    AST_EMPTY = 0
    AST_STRING = 1
    AST_DISJ = 2
    AST_CONCAT = 3
    AST_CHAR = 4
    AST_CAPTURE = 5
    AST_REPEAT_EXACTLY = 6
    AST_PERIOD = 7
    AST_ASSERT_START = 8
    AST_ASSERT_END = 9
    AST_ASSERT_WORD_BOUNDARY = 10
    AST_ASSERT_NOT_WORD_BOUNDARY = 11
    AST_MATCH = 12
    AST_ZEROONE_GREEDY = 13
    AST_ZEROONE_HUMBLE = 14
    AST_STAR_GREEDY = 15
    AST_STAR_HUMBLE = 16
    AST_REPEAT_MAX_GREEDY = 17
    AST_REPEAT_MAX_HUMBLE = 18
    AST_RANGES = 19
    AST_INVRANGES = 20
    AST_ASSERT_POS_LOOKAHEAD = 21
    AST_ASSERT_NEG_LOOKAHEAD = 22
    AST_ASSERT_POS_LOOKBEHIND = 23
    AST_ASSERT_NEG_LOOKBEHIND = 24
    AST_FAIL = 25
    AST_BACKREFERENCE = 26
    AST_NAMED = 27
    MAX_AST_CODE = 27
End Enum

Private Enum AstNodeDescriptionConstant
    NODE_TYPE = 0
    NODE_LCHILD = 1
    NODE_RCHILD = 2
End Enum

Private Enum AstTableDescriptionConstant
    AST_TABLE_OFFSET_NC = 0
    AST_TABLE_OFFSET_BLEN = 1
    AST_TABLE_OFFSET_ESFS = 2
    AST_TABLE_ENTRY_LENGTH = 3
    AST_TABLE_LENGTH = AST_TABLE_ENTRY_LENGTH * (MAX_AST_CODE + 1)
End Enum

Private astTableInitialized As Boolean ' default-initialized to False

Private UnicodeInitialized As Boolean ' auto-initialized to False
Private RangeTablesInitialized As Boolean ' auto-initialized to False

Private Enum StaticDataDescriptionConstant
    AST_TABLE_START = 0
    
    RANGE_TABLE_DIGIT_START = AST_TABLE_START + AST_TABLE_LENGTH
    RANGE_TABLE_DIGIT_LENGTH = 2
    
    RANGE_TABLE_WHITE_START = RANGE_TABLE_DIGIT_START + RANGE_TABLE_DIGIT_LENGTH
    RANGE_TABLE_WHITE_LENGTH = 22
    
    RANGE_TABLE_WORDCHAR_START = RANGE_TABLE_WHITE_START + RANGE_TABLE_WHITE_LENGTH
    RANGE_TABLE_WORDCHAR_LENGTH = 8
    
    RANGE_TABLE_NOTDIGIT_START = RANGE_TABLE_WORDCHAR_START + RANGE_TABLE_WORDCHAR_LENGTH
    RANGE_TABLE_NOTDIGIT_LENGTH = 4
    
    RANGE_TABLE_NOTWHITE_START = RANGE_TABLE_NOTDIGIT_START + RANGE_TABLE_NOTDIGIT_LENGTH
    RANGE_TABLE_NOTWHITE_LENGTH = 24
    
    RANGE_TABLE_NOTWORDCHAR_START = RANGE_TABLE_NOTWHITE_START + RANGE_TABLE_NOTWHITE_LENGTH
    RANGE_TABLE_NOTWORDCHAR_LENGTH = 10
    
    UNICODE_CANON_LOOKUP_TABLE_START = RANGE_TABLE_NOTWORDCHAR_START + RANGE_TABLE_NOTWORDCHAR_LENGTH
    UNICODE_CANON_LOOKUP_TABLE_LENGTH = 65536
    
    UNICODE_CANON_RUNS_TABLE_START = UNICODE_CANON_LOOKUP_TABLE_START + UNICODE_CANON_LOOKUP_TABLE_LENGTH
    UNICODE_CANON_RUNS_TABLE_LENGTH = 303
    
    STATIC_DATA_LENGTH = UNICODE_CANON_RUNS_TABLE_START + UNICODE_CANON_RUNS_TABLE_LENGTH
End Enum

Private StaticData(0 To STATIC_DATA_LENGTH - 1) As Long

Private Type StartLengthPair
    start As Long
    Length As Long
End Type

Private Type IdentifierTreeNode
    rbParent As Long
    rbChild(0 To 1) As Long ' 0 = left, 1 = right
    rbIsBlack As Boolean
    reference As StartLengthPair
    identifierId As Long
End Type

Private Type IdentifierTreeTy
    nEntries As Long
    root As Long
    bufferCapacity As Long
    Buffer() As IdentifierTreeNode
End Type

Private Type LexerContext
    iCurrent As Long
    iEnd As Long
    inputStr As String
    currentCharacter As Long
    identifierTree As IdentifierTreeTy
End Type

Private Type ReToken
    t As Long ' token type
    greedy As Boolean
    num As Long ' numeric value (character, count, id for named capture group, -1 for non-named capture group)
    qmin As Long
    qmax As Long
End Type

Private Enum TokenTypeIdType
    RETOK_EOF = 0
    RETOK_DISJUNCTION = 1
    RETOK_QUANTIFIER = 2
    RETOK_ASSERT_START = 3
    RETOK_ASSERT_END = 4
    RETOK_ASSERT_WORD_BOUNDARY = 5
    RETOK_ASSERT_NOT_WORD_BOUNDARY = 6
    RETOK_ASSERT_START_POS_LOOKAHEAD = 7
    RETOK_ASSERT_START_NEG_LOOKAHEAD = 8
    RETOK_ATOM_PERIOD = 9
    RETOK_ATOM_CHAR = 10
    RETOK_ATOM_DIGIT = 11                   ' assumptions in regexp compiler
    RETOK_ATOM_NOT_DIGIT = 12               ' -""-
    RETOK_ATOM_WHITE = 13                   ' -""-
    RETOK_ATOM_NOT_WHITE = 14               ' -""-
    RETOK_ATOM_WORD_CHAR = 15               ' -""-
    RETOK_ATOM_NOT_WORD_CHAR = 16           ' -""-
    RETOK_ATOM_BACKREFERENCE = 17
    RETOK_ATOM_START_CAPTURE_GROUP = 18
    RETOK_ATOM_START_NONCAPTURE_GROUP = 19
    RETOK_ATOM_START_CHARCLASS = 20
    RETOK_ATOM_START_CHARCLASS_INVERTED = 21
    RETOK_ASSERT_START_POS_LOOKBEHIND = 22
    RETOK_ASSERT_START_NEG_LOOKBEHIND = 23
    RETOK_ATOM_END = 24 ' closing parenthesis (ends (POS|NEG)_LOOK(AHEAD|BEHIND), CAPTURE_GROUP, NONCAPTURE_GROUP)
End Enum

Private Const LEXER_ENDOFINPUT As Long = -1

Private Enum LexerUnicodeCodepointConstant
    UNICODE_EXCLAMATION = 33  ' !
    UNICODE_DOLLAR = 36  ' $
    UNICODE_LPAREN = 40  ' (
    UNICODE_RPAREN = 41  ' )
    UNICODE_STAR = 42  ' *
    UNICODE_PLUS = 43  ' +
    UNICODE_COMMA = 44  ' ,
    UNICODE_MINUS = 45  ' -
    UNICODE_PERIOD = 46  ' .
    UNICODE_0 = 48  ' 0
    UNICODE_1 = 49  ' 1
    UNICODE_7 = 55  ' 7
    UNICODE_9 = 57  ' 9
    UNICODE_COLON = 58  ' :
    UNICODE_LT = 60  ' <
    UNICODE_EQUALS = 61  ' =
    UNICODE_GT = 62  ' >
    UNICODE_QUESTION = 63  ' ?
    UNICODE_UC_A = 65  ' A
    UNICODE_UC_B = 66  ' B
    UNICODE_UC_D = 68  ' D
    UNICODE_UC_F = 70  ' F
    UNICODE_UC_S = 83  ' S
    UNICODE_UC_W = 87  ' W
    UNICODE_UC_Z = 90  ' Z
    UNICODE_LBRACKET = 91  ' [
    UNICODE_BACKSLASH = 92  ' \
    UNICODE_RBRACKET = 93  ' ]
    UNICODE_CARET = 94  ' ^
    UNICODE_LC_A = 97  ' a
    UNICODE_LC_B = 98  ' b
    UNICODE_LC_C = 99  ' c
    UNICODE_LC_D = 100  ' d
    UNICODE_LC_F = 102  ' f
    UNICODE_LC_N = 110  ' n
    UNICODE_LC_R = 114  ' r
    UNICODE_LC_S = 115  ' s
    UNICODE_LC_T = 116  ' t
    UNICODE_LC_U = 117  ' u
    UNICODE_LC_V = 118  ' v
    UNICODE_LC_W = 119  ' w
    UNICODE_LC_X = 120  ' x
    UNICODE_LC_Z = 122  ' z
    UNICODE_LCURLY = 123  ' {
    UNICODE_PIPE = 124  ' |
    UNICODE_RCURLY = 125  ' }
    UNICODE_CP_ZWNJ = &H200C& ' zero-width non-joiner
    UNICODE_CP_ZWJ = &H200D&  ' zero-width joiner
End Enum

Private Enum DfsMatcherSharedConstant
    DEFAULT_STEPS_LIMIT = 10000
End Enum

Private Enum DfsMatcherPrivateConstant
    DEFAULT_MINIMUM_THREADSTACK_CAPACITY = 16
    Q_NONE = -2
    DFS_MATCHER_STACK_MINIMUM_CAPACITY = 16
    DFS_ENDOFINPUT = -1
End Enum

Private Type CapturesTy
    nNumberedCaptures As Long
    nNamedCaptures As Long
    entireMatch As StartLengthPair
    numberedCaptures() As StartLengthPair
    namedCaptures() As Long
End Type

Private Type DfsMatcherStackFrame
    master As Long
    capturesStackState As Long
    qStackLength As Long
    pc As Long
    sp As Long
    pcLandmark As Long
    spDelta As Long
    q As Long
    qTop As Long
End Type

Private Type DfsMatcherStack
    Buffer() As DfsMatcherStackFrame
    Capacity As Long
    Length As Long
End Type

Private Type DfsMatcherContext
    matcherStack As DfsMatcherStack
    capturesStack As ArrayBuffer
    qstack As ArrayBuffer
    nProperCapturePoints As Long
    nCapturePoints As Long ' capture points including slots for named captures
    
    master As Long
    
    capturesRequireCoW As Boolean
    qTop As Long
End Type

Private Enum ReplType
    REPL_END = 0
    REPL_DOLLAR = 1
    REPL_SUBSTR = 2
    REPL_PREFIX = 3
    REPL_SUFFIX = 4
    REPL_ACTUAL = 5
    REPL_NUMBERED = 6
    REPL_NAMED = 7
End Enum

Public Type RegexTy
    pattern As String
    bytecode() As Long
    isCaseInsensitive As Long
    stepsLimit As Long
End Type

Public Type MatcherStateTy
    multiLine As Boolean
    localMatch As Boolean
    current As Long
    captures As CapturesTy
    context As DfsMatcherContext
End Type

Private Sub AppendStr(ByRef sb As StaticStringBuilder, ByRef s As String)
    Dim Length As Long, nRequired As Long
    With sb
        Length = Len(s)
        If Length = 0 Then Exit Sub
        nRequired = .Length + Length
        If nRequired > .Capacity Then SwitchToLargerBuffer sb, nRequired
        Mid$(.Buffer(.Active), .Length + 1, Length) = s
        .Length = nRequired
    End With
End Sub

Private Sub Clear(ByRef sb As StaticStringBuilder)
    With sb
        .Active = 0
        .Buffer(0) = vbNullString
        .Buffer(1) = vbNullString
        .Capacity = 0
        .Length = 0
    End With
End Sub

Private Function GetLength(ByRef sb As StaticStringBuilder) As Long
    GetLength = sb.Length
End Function

Private Function GetStr(ByRef sb As StaticStringBuilder) As String
    With sb
        GetStr = Left$(.Buffer(.Active), .Length)
    End With
End Function

Private Function GetSubstr(ByRef sb As StaticStringBuilder, ByVal start As Long, ByVal Length As Long) As String
    Dim n As Long
    With sb
        n = .Length - start + 1
        If n <= 0 Then
            GetSubstr = vbNullString
            Exit Function
        End If
        If Length <= n Then n = Length
        GetSubstr = Mid$(.Buffer(.Active), start, n)
    End With
End Function

Private Sub SetMinimumCapacity(ByRef sb As StaticStringBuilder, ByVal MinimumCapacity As Long)
    If MinimumCapacity >= 2 Then sb.MinimumCapacity = MinimumCapacity Else sb.MinimumCapacity = 2
End Sub

Private Sub SwitchToLargerBuffer(ByRef sb As StaticStringBuilder, ByVal nRequired As Long)
    ' Allocate buffer that is able to hold nRequired characters.
    ' The new buffer size is calculated by repeatedly growing the current size by 50%.
    ' Copy string over to the new buffer.
    ' Deallocate the old buffer.
    With sb
        If .MinimumCapacity <= 1 Then .MinimumCapacity = STATIC_STRING_BUILDER_DEFAULT_MINIMUM_CAPACITY
        If .Capacity < .MinimumCapacity Then .Capacity = .MinimumCapacity
        Do
            If .Capacity >= nRequired Then Exit Do
            .Capacity = .Capacity + .Capacity \ 2
        Loop
        .Buffer(1 - .Active) = String(.Capacity, 0)
        Mid$(.Buffer(1 - .Active), 1, .Length) = .Buffer(.Active)
        .Buffer(.Active) = vbNullString
        .Active = 1 - .Active
    End With
End Sub

Private Sub AppendLong(ByRef lab As ArrayBuffer, ByVal v As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 1
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendTwo(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 2
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendThree(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 3
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendFour(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 4
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Buffer(.Length + 3) = v4
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendFive(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 5
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Buffer(.Length + 3) = v4
        .Buffer(.Length + 4) = v5
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendEight(ByRef lab As ArrayBuffer, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long, v6 As Long, v7 As Long, v8 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 8
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Buffer(.Length + 3) = v4
        .Buffer(.Length + 4) = v5
        .Buffer(.Length + 5) = v6
        .Buffer(.Length + 6) = v7
        .Buffer(.Length + 7) = v8
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendFill(ByRef lab As ArrayBuffer, ByVal cnt As Long, ByVal v As Long)
    Dim requiredCapacity As Long, i As Long
    With lab
        requiredCapacity = .Length + cnt
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        i = .Length
        Do While i < requiredCapacity: .Buffer(i) = v: i = i + 1: Loop
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendSlice(ByRef lab As ArrayBuffer, ByVal offset As Long, ByVal Length As Long)
    Dim requiredCapacity As Long, i As Long, j As Long, upper As Long
    With lab
        requiredCapacity = .Length + Length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        upper = offset + Length: i = offset: j = .Length
        Do While i < upper
            .Buffer(j) = .Buffer(i)
            i = i + 1: j = j + 1
        Loop
        .Length = requiredCapacity
    End With
End Sub

Private Sub AppendUnspecified(ByRef lab As ArrayBuffer, ByVal n As Long)
    Dim requiredCapacity As Long
    With lab
        .Length = .Length + n
        requiredCapacity = .Length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
    End With
End Sub

Private Sub AppendPrefixedPairsArray(ByRef lab As ArrayBuffer, ByVal prefix As Long, ByRef ary() As Long, ByVal aryStart As Long, ByVal aryLength As Long)
    ' prefix, number of pairs, pairs
    Dim requiredCapacity As Long, ub As Long, i As Long, j As Long
    With lab
        i = .Length
        .Length = .Length + aryLength + 2
        requiredCapacity = .Length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        
        .Buffer(i) = prefix: i = i + 1
        .Buffer(i) = aryLength \ 2: i = i + 1
        ub = aryStart + aryLength - 1
        For j = aryStart To ub
            .Buffer(i) = ary(j)
            i = i + 1
        Next
    End With
End Sub

Private Sub IncreaseCapacity(ByRef lab As ArrayBuffer, requiredCapacity As Long)
    Dim cap As Long
    With lab
        cap = .Capacity
        If cap <= ArrayBufferConstant.ARRAY_BUFFER_MINIMUM_CAPACITY Then cap = ArrayBufferConstant.ARRAY_BUFFER_MINIMUM_CAPACITY
        Do Until cap >= requiredCapacity
            cap = cap + cap \ 2
        Loop
        ReDim Preserve .Buffer(0 To cap - 1) As Long
        .Capacity = cap
    End With
End Sub

Private Function isCaseInsensitive(ByRef bytecode() As Long) As Boolean
    isCaseInsensitive = bytecode(BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR) <> 0
End Function

Private Function GetIdentifierId( _
    ByRef bytecode() As Long, _
    ByRef lake As String, _
    ByRef identifier As String _
) As Long
    Dim aa As Long, bb As Long, mm As Long, compare As Long, identifierLength As Long
    
    identifierLength = Len(identifier)
    
    aa = BYTECODE_IDENTIFIER_MAP_BEGIN
    bb = BYTECODE_IDENTIFIER_MAP_BEGIN + BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE * bytecode(BYTECODE_IDX_N_IDENTIFIERS)
    
    ' Find numeric id for identifier name.
    ' We are doing a binary search here.
    ' Invariant: Value we are looking for, if it exists, is always contained in the interval [aa;bb).
    Do
        If aa >= bb Then GetIdentifierId = -1: Exit Function ' identifier not found
        
        mm = aa + 3 * ((bb - aa) \ 6)
        If identifierLength < bytecode(mm + 1) Then
            bb = mm
        ElseIf identifierLength > bytecode(mm + 1) Then
            aa = mm + 3
        Else
            compare = StrComp( _
                identifier, _
                Mid$(lake, bytecode(mm), bytecode(mm + 1)), _
                vbBinaryCompare _
            )
            If compare < 0 Then
                bb = mm
            ElseIf compare > 0 Then
                aa = mm + 3
            Else
                ' found
                GetIdentifierId = bytecode(mm + 2)
                Exit Function
            End If
        End If
    Loop

End Function

Private Sub AstTableInitialize()
    InitializeAstTable StaticData
End Sub

Private Sub InitializeAstTable(ByRef t() As Long)
    Const b As Long = AST_TABLE_START
    Const nc As Long = b + AST_TABLE_OFFSET_NC
    Const blen As Long = b + AST_TABLE_OFFSET_BLEN
    Const esfs As Long = b + AST_TABLE_OFFSET_ESFS
    Const e As Long = AST_TABLE_ENTRY_LENGTH
    
    ' nc: number of children; negative values have special meaning
    '   -2: is AST_STRING
    '   -1: is AST_RANGES or AST_INVRANGES
    ' blen: length of bytecode generated for this node (meaningful only if .nc >= 0)
    ' esfs: extra stack space required when generating bytecode for this node
    '   Only nodes with children are permitted to require extra stack space.
    '   Hence .esfs > 0 must imply .nc >= 1.
    
    ' ! When adding new entries, make sure to adjust BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY below, if necessary!
    ' ! See comment on BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY below.
    
    t(nc + e * AST_EMPTY) = 0:                    t(blen + e * AST_EMPTY) = 0:                        t(esfs + e * AST_EMPTY) = 0
    t(nc + e * AST_STRING) = -2:                  t(blen + e * AST_STRING) = 2:                       t(esfs + e * AST_STRING) = 0
    t(nc + e * AST_DISJ) = 2:                     t(blen + e * AST_DISJ) = 4:                         t(esfs + e * AST_DISJ) = 1
    t(nc + e * AST_CONCAT) = 2:                   t(blen + e * AST_CONCAT) = 0:                       t(esfs + e * AST_CONCAT) = 0
    t(nc + e * AST_CHAR) = 0:                     t(blen + e * AST_CHAR) = 2:                         t(esfs + e * AST_CHAR) = 0
    t(nc + e * AST_CAPTURE) = 1:                  t(blen + e * AST_CAPTURE) = 4:                      t(esfs + e * AST_CAPTURE) = 0
    t(nc + e * AST_REPEAT_EXACTLY) = 1:           t(blen + e * AST_REPEAT_EXACTLY) = 7:               t(esfs + e * AST_REPEAT_EXACTLY) = 1
    t(nc + e * AST_PERIOD) = 0:                   t(blen + e * AST_PERIOD) = 1:                       t(esfs + e * AST_PERIOD) = 0
    t(nc + e * AST_ASSERT_START) = 0:             t(blen + e * AST_ASSERT_START) = 1:                 t(esfs + e * AST_ASSERT_START) = 0
    t(nc + e * AST_ASSERT_END) = 0:               t(blen + e * AST_ASSERT_END) = 1:                   t(esfs + e * AST_ASSERT_END) = 0
    t(nc + e * AST_ASSERT_WORD_BOUNDARY) = 0:     t(blen + e * AST_ASSERT_WORD_BOUNDARY) = 1:         t(esfs + e * AST_ASSERT_WORD_BOUNDARY) = 0
    t(nc + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 0: t(blen + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 1:     t(esfs + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 0
    t(nc + e * AST_MATCH) = 0:                    t(blen + e * AST_MATCH) = 1:                        t(esfs + e * AST_MATCH) = 0
    t(nc + e * AST_ZEROONE_GREEDY) = 1:           t(blen + e * AST_ZEROONE_GREEDY) = 2:               t(esfs + e * AST_ZEROONE_GREEDY) = 1
    t(nc + e * AST_ZEROONE_HUMBLE) = 1:           t(blen + e * AST_ZEROONE_HUMBLE) = 2:               t(esfs + e * AST_ZEROONE_HUMBLE) = 1
    t(nc + e * AST_STAR_GREEDY) = 1:              t(blen + e * AST_STAR_GREEDY) = 4:                  t(esfs + e * AST_STAR_GREEDY) = 1
    t(nc + e * AST_STAR_HUMBLE) = 1:              t(blen + e * AST_STAR_HUMBLE) = 4:                  t(esfs + e * AST_STAR_HUMBLE) = 1
    t(nc + e * AST_REPEAT_MAX_GREEDY) = 1:        t(blen + e * AST_REPEAT_MAX_GREEDY) = 7:            t(esfs + e * AST_REPEAT_MAX_GREEDY) = 1
    t(nc + e * AST_REPEAT_MAX_HUMBLE) = 1:        t(blen + e * AST_REPEAT_MAX_HUMBLE) = 7:            t(esfs + e * AST_REPEAT_MAX_HUMBLE) = 1
    t(nc + e * AST_RANGES) = -1:                  t(blen + e * AST_RANGES) = 2:                       t(esfs + e * AST_RANGES) = 0
    t(nc + e * AST_INVRANGES) = -1:               t(blen + e * AST_INVRANGES) = 2:                    t(esfs + e * AST_INVRANGES) = 0
    t(nc + e * AST_ASSERT_POS_LOOKAHEAD) = 1:     t(blen + e * AST_ASSERT_POS_LOOKAHEAD) = 4:         t(esfs + e * AST_ASSERT_POS_LOOKAHEAD) = 2
    t(nc + e * AST_ASSERT_NEG_LOOKAHEAD) = 1:     t(blen + e * AST_ASSERT_NEG_LOOKAHEAD) = 4:         t(esfs + e * AST_ASSERT_NEG_LOOKAHEAD) = 2
    t(nc + e * AST_ASSERT_POS_LOOKBEHIND) = 1:    t(blen + e * AST_ASSERT_POS_LOOKBEHIND) = 4:        t(esfs + e * AST_ASSERT_POS_LOOKBEHIND) = 2
    t(nc + e * AST_ASSERT_NEG_LOOKBEHIND) = 1:    t(blen + e * AST_ASSERT_NEG_LOOKBEHIND) = 4:        t(esfs + e * AST_ASSERT_NEG_LOOKBEHIND) = 2
    t(nc + e * AST_FAIL) = 0:                     t(blen + e * AST_FAIL) = 1:                         t(esfs + e * AST_FAIL) = 0
    t(nc + e * AST_BACKREFERENCE) = 0:            t(blen + e * AST_BACKREFERENCE) = 2:                t(esfs + e * AST_BACKREFERENCE) = 0
    t(nc + e * AST_NAMED) = 1:                    t(blen + e * AST_NAMED) = 3:                        t(esfs + e * AST_NAMED) = 0
End Sub

Private Sub AstToBytecode(ByRef ast() As Long, ByRef identifierTree As IdentifierTreeTy, ByVal caseInsensitive As Boolean, ByRef bytecode() As Long)
    Dim bytecodePtr As Long
    Dim curNode As Long, prevNode As Long
    Dim stack() As Long, sp As Long
    Dim direction As Long ' 0 = left before right, -1 = right before left
    Dim returningFromFirstChild As Long ' 0 = no, LONG_FIRST_BIT = yes
    
    ' temporaries, do not survive over more than one iteration
    Dim opcode1 As Long, opcode2 As Long, opcode3 As Long, tmp As Long, tmpCnt As Long, _
        e As Long, j As Long, patchPos As Long, maxSave As Long
    
    If Not astTableInitialized Then AstTableInitialize
    
    PrepareStackAndBytecodeBuffer ast, identifierTree, caseInsensitive, stack, bytecode
    
    sp = 0
    
    prevNode = -1
    curNode = ast(0) ' first word contains index of root
    bytecodePtr = 3 + 3 * bytecode(1)
    maxSave = -1
    direction = 0
    returningFromFirstChild = 0

ContinueLoop:
        Select Case ast(curNode + NODE_TYPE)
        Case AST_STRING
            tmpCnt = ast(curNode + 1) ' assert(tmpCnt >= 1)
            j = curNode + 2 + ((tmpCnt - 1) And direction)
            e = curNode + 1 + tmpCnt - ((tmpCnt - 1) And direction)
            tmp = 1 + 2 * direction
            Do
                bytecode(bytecodePtr) = REOP_CHAR: bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(j): bytecodePtr = bytecodePtr + 1
                If j = e Then Exit Do
                j = j + tmp
            Loop
            GoTo TurnToParent
        Case AST_RANGES
            opcode1 = REOP_RANGES
            GoTo HandleRanges
        Case AST_INVRANGES
            opcode1 = REOP_INVRANGES
            GoTo HandleRanges
        Case AST_CHAR
            bytecode(bytecodePtr) = REOP_CHAR: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = ast(curNode + 1): bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_PERIOD
            bytecode(bytecodePtr) = REOP_PERIOD: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_MATCH
            bytecode(bytecodePtr) = REOP_MATCH: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_START
            bytecode(bytecodePtr) = REOP_ASSERT_START: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_END
            bytecode(bytecodePtr) = REOP_ASSERT_END: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_WORD_BOUNDARY
            bytecode(bytecodePtr) = REOP_ASSERT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_NOT_WORD_BOUNDARY
            bytecode(bytecodePtr) = REOP_ASSERT_NOT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_DISJ
            If returningFromFirstChild Then ' previous was left child
                sp = sp - 1: patchPos = stack(sp)
                bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
                stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
                bytecode(patchPos) = bytecodePtr - patchPos - 1
                
                GoTo TurnToRightChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD) Then ' previous was right child
                sp = sp - 1: patchPos = stack(sp)
                bytecode(patchPos) = bytecodePtr - patchPos - 1
            
                GoTo TurnToParent
            Else ' previous was parent
                bytecode(bytecodePtr) = REOP_SPLIT1: bytecodePtr = bytecodePtr + 1
                stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        Case AST_CONCAT
            If returningFromFirstChild Then ' previous was first child
                GoTo TurnToSecondChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD + direction) Then ' previous was second child
                GoTo TurnToParent
            Else ' previous was parent
                GoTo TurnToFirstChild
            End If
        Case AST_CAPTURE
            If returningFromFirstChild Then
                bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
                tmp = ast(curNode + 2) * 2 + 1
                If tmp > maxSave Then maxSave = tmp
                bytecode(bytecodePtr) = tmp + direction: bytecodePtr = bytecodePtr + 1
                GoTo TurnToParent
            Else
                bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(curNode + 2) * 2 - direction: bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        Case AST_REPEAT_EXACTLY
            opcode1 = REOP_REPEAT_EXACTLY_INIT: opcode2 = REOP_REPEAT_EXACTLY_START: opcode3 = REOP_REPEAT_EXACTLY_END
            GoTo HandleRepeatQuantified
        Case AST_REPEAT_MAX_GREEDY
            opcode1 = REOP_REPEAT_GREEDY_MAX_INIT: opcode2 = REOP_REPEAT_GREEDY_MAX_START: opcode3 = REOP_REPEAT_GREEDY_MAX_END
            GoTo HandleRepeatQuantified
        Case AST_REPEAT_MAX_HUMBLE
            opcode1 = REOP_REPEAT_MAX_HUMBLE_INIT: opcode2 = REOP_REPEAT_MAX_HUMBLE_START: opcode3 = REOP_REPEAT_MAX_HUMBLE_END
            GoTo HandleRepeatQuantified
        Case AST_ZEROONE_GREEDY
            opcode1 = REOP_SPLIT1
            GoTo HandleZeroone
        Case AST_ZEROONE_HUMBLE
            opcode1 = REOP_SPLIT2
            GoTo HandleZeroone
        Case AST_STAR_GREEDY
            opcode1 = REOP_SPLIT1
            GoTo HandleStar
        Case AST_STAR_HUMBLE
            opcode1 = REOP_SPLIT2
            GoTo HandleStar
        Case AST_ASSERT_POS_LOOKAHEAD
            opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
            GoTo HandleLookahead
        Case AST_ASSERT_NEG_LOOKAHEAD
            opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
            GoTo HandleLookahead
        Case AST_ASSERT_POS_LOOKBEHIND
            opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
            GoTo HandleLookbehind
        Case AST_ASSERT_NEG_LOOKBEHIND
            opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
            GoTo HandleLookbehind
        Case AST_EMPTY
            GoTo TurnToParent
        Case AST_FAIL
            bytecode(bytecodePtr) = REOP_FAIL: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_BACKREFERENCE
            bytecode(bytecodePtr) = REOP_BACKREFERENCE: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = ast(curNode + 1): bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_NAMED
            If returningFromFirstChild Then
                ' nothing to be done
                GoTo TurnToParent
            Else
                bytecode(bytecodePtr) = REOP_SET_NAMED: bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(curNode + 2): bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(curNode + 3): bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        
        End Select
        
        Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable
        
HandleRanges: ' requires: opcode1
        tmpCnt = ast(curNode + 1)
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        j = curNode
        e = curNode + 1 + 2 * tmpCnt
        Do ' copy everything, including first word, which is the length
            j = j + 1
            bytecode(bytecodePtr) = ast(j): bytecodePtr = bytecodePtr + 1
        Loop Until j = e
        GoTo TurnToParent

HandleRepeatQuantified: ' requires: opcode1, opcode2, opcode 3
        tmpCnt = ast(curNode + 2)
        If returningFromFirstChild Then
            sp = sp - 1: patchPos = stack(sp)
            bytecode(bytecodePtr) = opcode3: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
            tmp = bytecodePtr - patchPos
            bytecode(bytecodePtr) = tmp: bytecodePtr = bytecodePtr + 1
            bytecode(patchPos) = tmp
            GoTo TurnToParent
        Else
            bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
            stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
            GoTo TurnToLeftChild
        End If

HandleZeroone: ' requires: opcode1
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleStar:
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        tmp = bytecodePtr - patchPos + 1
        bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = -(tmp + 2): bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = tmp
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleLookahead: ' requires opcode1, opcode2
    If returningFromFirstChild Then
        sp = sp - 1: direction = stack(sp)
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = REOP_CHECK_LOOKAHEAD: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        stack(sp) = direction: sp = sp + 1
        direction = 0
        GoTo TurnToLeftChild
    End If

HandleLookbehind: ' requires opcode1, opcode2
    If returningFromFirstChild Then
        sp = sp - 1: direction = stack(sp)
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = REOP_CHECK_LOOKBEHIND: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        stack(sp) = direction: sp = sp + 1
        direction = -1
        GoTo TurnToLeftChild
    End If

TurnToParent:
    prevNode = curNode
    If sp = 0 Then GoTo BreakLoop
    sp = sp - 1: tmp = stack(sp)
    curNode = tmp And LONG_ALL_BUT_FIRST_BIT
    returningFromFirstChild = tmp And LONG_FIRST_BIT
    GoTo ContinueLoop
TurnToLeftChild:
    prevNode = curNode
    stack(sp) = curNode Or LONG_FIRST_BIT: sp = sp + 1
    curNode = ast(curNode + NODE_LCHILD)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast(curNode + NODE_RCHILD)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToFirstChild:
    prevNode = curNode
    stack(sp) = curNode Or LONG_FIRST_BIT: sp = sp + 1
    curNode = ast(curNode + NODE_LCHILD - direction)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToSecondChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast(curNode + NODE_RCHILD + direction)
    returningFromFirstChild = 0
    GoTo ContinueLoop
    
BreakLoop:
    bytecode(0) = maxSave
    bytecode(bytecodePtr) = REOP_MATCH
End Sub

Private Sub PrepareStackAndBytecodeBuffer(ByRef ast() As Long, ByRef identifierTree As IdentifierTreeTy, ByVal caseInsensitive As Boolean, ByRef stack() As Long, ByRef bytecode() As Long)
    Dim sp As Long, prevNode As Long, curNode As Long, esfs As Long, stackCapacity As Long
    Dim tmp As Long, astTableIdx As Long
    Dim bytecodeLength As Long
    Dim returningFromFirstChild As Long ' 0 = no, LONG_FIRST_BIT = yes
    
    ' The initial stack capacity must be >= 2 * max([1 + entry.esfs for entry in AST_TABLE]),
    '   since when increasing the stack capacity, we increase by (current size \ 2) and
    '   we assume that this will be sufficient for the next stack frame.
    Const BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY As Long = 8
    
    stackCapacity = BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY
    ReDim stack(0 To BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY - 1) As Long

    sp = 0
    
    prevNode = -1
    curNode = ast(0) ' first word contains index of root
    returningFromFirstChild = 0

    bytecodeLength = 0
    
ContinueLoop:
        astTableIdx = ast(curNode + NODE_TYPE) * AST_TABLE_ENTRY_LENGTH
        esfs = StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_ESFS)
        
        Select Case StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_NC)
        Case -2
            bytecodeLength = bytecodeLength + 2 * ast(curNode + 1)
            GoTo TurnToParent
        Case -1
            bytecodeLength = bytecodeLength + 2 + 2 * ast(curNode + 1)
            GoTo TurnToParent
        Case 0
            bytecodeLength = bytecodeLength + StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
            GoTo TurnToParent
        Case 1
            If returningFromFirstChild Then
                GoTo TurnToParent
            Else
                bytecodeLength = bytecodeLength + StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
                GoTo TurnToLeftChild
            End If
        Case 2
            If returningFromFirstChild Then ' previous was left child
                GoTo TurnToRightChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD) Then ' previous was right child
                GoTo TurnToParent
            Else ' previous was parent
                bytecodeLength = bytecodeLength + StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
                GoTo TurnToLeftChild
            End If
        End Select
        
        Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable

TurnToParent:
    sp = sp - esfs
    If sp = 0 Then GoTo BreakLoop
    prevNode = curNode
    sp = sp - 1: tmp = stack(sp)
    returningFromFirstChild = tmp And LONG_FIRST_BIT: curNode = tmp And LONG_ALL_BUT_FIRST_BIT
    GoTo ContinueLoop
TurnToLeftChild:
    If sp >= stackCapacity - esfs Then
        stackCapacity = stackCapacity + stackCapacity \ 2
        ReDim Preserve stack(0 To stackCapacity - 1) As Long
    End If
    prevNode = curNode
    sp = sp + esfs: stack(sp) = curNode Or LONG_FIRST_BIT: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast(curNode + NODE_LCHILD)
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast(curNode + NODE_RCHILD)
    GoTo ContinueLoop
    
BreakLoop:
    ' Actual bytecode length is bytecodeLength + 4 + 3*identifierTree(N_NODES) due to intial nCaptures and final REOP_MATCH.
    ReDim bytecode(0 To bytecodeLength + 3 + 3 * identifierTree.nEntries) As Long
    bytecode(BYTECODE_IDX_N_IDENTIFIERS) = identifierTree.nEntries
    bytecode(BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR) = -caseInsensitive
    RedBlackDumpTree bytecode, BYTECODE_IDENTIFIER_MAP_BEGIN, identifierTree
End Sub

Private Sub UnicodeInitialize()
    InitializeUnicodeCanonLookupTable StaticData
    InitializeUnicodeCanonRunsTable StaticData
    UnicodeInitialized = True
End Sub

Private Sub RangeTablesInitialize()
    InitializeRangeTables StaticData
    RangeTablesInitialized = True
End Sub

Private Sub InitializeRangeTables(ByRef t() As Long)
    t(RANGE_TABLE_DIGIT_START + 0) = &H30&
    t(RANGE_TABLE_DIGIT_START + 1) = &H39&
    
    '---------------------------------
    
    t(RANGE_TABLE_WHITE_START + 0) = &H9&
    t(RANGE_TABLE_WHITE_START + 1) = &HD&
    t(RANGE_TABLE_WHITE_START + 2) = &H20&
    t(RANGE_TABLE_WHITE_START + 3) = &H20&
    t(RANGE_TABLE_WHITE_START + 4) = &HA0&
    t(RANGE_TABLE_WHITE_START + 5) = &HA0&
    t(RANGE_TABLE_WHITE_START + 6) = &H1680&
    t(RANGE_TABLE_WHITE_START + 7) = &H1680&
    t(RANGE_TABLE_WHITE_START + 8) = &H180E&
    t(RANGE_TABLE_WHITE_START + 9) = &H180E&
    t(RANGE_TABLE_WHITE_START + 10) = &H2000&
    t(RANGE_TABLE_WHITE_START + 11) = &H200A
    t(RANGE_TABLE_WHITE_START + 12) = &H2028&
    t(RANGE_TABLE_WHITE_START + 13) = &H2029&
    t(RANGE_TABLE_WHITE_START + 14) = &H202F
    t(RANGE_TABLE_WHITE_START + 15) = &H202F
    t(RANGE_TABLE_WHITE_START + 16) = &H205F&
    t(RANGE_TABLE_WHITE_START + 17) = &H205F&
    t(RANGE_TABLE_WHITE_START + 18) = &H3000&
    t(RANGE_TABLE_WHITE_START + 19) = &H3000&
    t(RANGE_TABLE_WHITE_START + 20) = &HFEFF&
    t(RANGE_TABLE_WHITE_START + 21) = &HFEFF&
    
    '---------------------------------
    
    t(RANGE_TABLE_WORDCHAR_START + 0) = &H30&
    t(RANGE_TABLE_WORDCHAR_START + 1) = &H39&
    t(RANGE_TABLE_WORDCHAR_START + 2) = &H41&
    t(RANGE_TABLE_WORDCHAR_START + 3) = &H5A&
    t(RANGE_TABLE_WORDCHAR_START + 4) = &H5F&
    t(RANGE_TABLE_WORDCHAR_START + 5) = &H5F&
    t(RANGE_TABLE_WORDCHAR_START + 6) = &H61&
    t(RANGE_TABLE_WORDCHAR_START + 7) = &H7A&
    
    '---------------------------------
    
    t(RANGE_TABLE_NOTDIGIT_START + 0) = LONG_MIN
    t(RANGE_TABLE_NOTDIGIT_START + 1) = &H2F&
    t(RANGE_TABLE_NOTDIGIT_START + 2) = &H3A&
    t(RANGE_TABLE_NOTDIGIT_START + 3) = LONG_MAX
    
    '---------------------------------
    
    t(RANGE_TABLE_NOTWHITE_START + 0) = LONG_MIN
    t(RANGE_TABLE_NOTWHITE_START + 1) = &H8&
    t(RANGE_TABLE_NOTWHITE_START + 2) = &HE&
    t(RANGE_TABLE_NOTWHITE_START + 3) = &H1F&
    t(RANGE_TABLE_NOTWHITE_START + 4) = &H21&
    t(RANGE_TABLE_NOTWHITE_START + 5) = &H9F&
    t(RANGE_TABLE_NOTWHITE_START + 6) = &HA1&
    t(RANGE_TABLE_NOTWHITE_START + 7) = &H167F&
    t(RANGE_TABLE_NOTWHITE_START + 8) = &H1681&
    t(RANGE_TABLE_NOTWHITE_START + 9) = &H180D&
    t(RANGE_TABLE_NOTWHITE_START + 10) = &H180F&
    t(RANGE_TABLE_NOTWHITE_START + 11) = &H1FFF&
    t(RANGE_TABLE_NOTWHITE_START + 12) = &H200B&
    t(RANGE_TABLE_NOTWHITE_START + 13) = &H2027&
    t(RANGE_TABLE_NOTWHITE_START + 14) = &H202A&
    t(RANGE_TABLE_NOTWHITE_START + 15) = &H202E&
    t(RANGE_TABLE_NOTWHITE_START + 16) = &H2030&
    t(RANGE_TABLE_NOTWHITE_START + 17) = &H205E&
    t(RANGE_TABLE_NOTWHITE_START + 18) = &H2060&
    t(RANGE_TABLE_NOTWHITE_START + 19) = &H2FFF&
    t(RANGE_TABLE_NOTWHITE_START + 20) = &H3001&
    t(RANGE_TABLE_NOTWHITE_START + 21) = &HFEFE&
    t(RANGE_TABLE_NOTWHITE_START + 22) = &HFF00&
    t(RANGE_TABLE_NOTWHITE_START + 23) = LONG_MAX
    
    '---------------------------------
    
    t(RANGE_TABLE_NOTWORDCHAR_START + 0) = LONG_MIN
    t(RANGE_TABLE_NOTWORDCHAR_START + 1) = &H2F&
    t(RANGE_TABLE_NOTWORDCHAR_START + 2) = &H3A&
    t(RANGE_TABLE_NOTWORDCHAR_START + 3) = &H40&
    t(RANGE_TABLE_NOTWORDCHAR_START + 4) = &H5B&
    t(RANGE_TABLE_NOTWORDCHAR_START + 5) = &H5E&
    t(RANGE_TABLE_NOTWORDCHAR_START + 6) = &H60&
    t(RANGE_TABLE_NOTWORDCHAR_START + 7) = &H60&
    t(RANGE_TABLE_NOTWORDCHAR_START + 8) = &H7B&
    t(RANGE_TABLE_NOTWORDCHAR_START + 9) = LONG_MAX
    
    '---------------------------------
End Sub

Private Function ReCanonicalizeChar(ByVal codepoint As Long) As Long
    ' ! This function must not alter codepoint if codepoint is negative, as ENDOFINPUT is represented by -1.
    If codepoint And &HFFFF0000 Then ' Codepoint not in [0;&HFFFF&]
        ReCanonicalizeChar = codepoint
    Else
        ReCanonicalizeChar = (codepoint + StaticData(UNICODE_CANON_LOOKUP_TABLE_START + codepoint)) And &HFFFF&
    End If
End Function

Private Function UnicodeIsLineTerminator(ByVal codepoint As Long) As Boolean
    ' ! This function must return False for negative values of codepoint, as ENDOFINPUT is represented by -1.
    If codepoint = &HA& Then
        UnicodeIsLineTerminator = True
    ElseIf codepoint = &HD& Then
        UnicodeIsLineTerminator = True
    ElseIf codepoint < &H2028& Then
        UnicodeIsLineTerminator = False
    ElseIf codepoint > &H2029& Then
        UnicodeIsLineTerminator = False
    Else
        UnicodeIsLineTerminator = True
    End If
End Function

Private Sub InitializeUnicodeCanonLookupTable(ByRef t() As Long)
    ' Array of integers would be sufficient
    Const b As Long = UNICODE_CANON_LOOKUP_TABLE_START
    
    Dim i As Long
    
    For i = b + 97 To b + 122: t(i) = -32: Next i
    t(b + 181) = 743
    For i = b + 224 To b + 246: t(i) = -32: Next i
    For i = b + 248 To b + 254: t(i) = -32: Next i
    t(b + 255) = 121
    For i = b + 257 To b + 303 Step 2: t(i) = -1: Next i
    t(b + 307) = -1
    t(b + 309) = -1
    t(b + 311) = -1
    For i = b + 314 To b + 328 Step 2: t(i) = -1: Next i
    For i = b + 331 To b + 375 Step 2: t(i) = -1: Next i
    t(b + 378) = -1
    t(b + 380) = -1
    t(b + 382) = -1
    t(b + 384) = 195
    t(b + 387) = -1
    t(b + 389) = -1
    t(b + 392) = -1
    t(b + 396) = -1
    t(b + 402) = -1
    t(b + 405) = 97
    t(b + 409) = -1
    t(b + 410) = 163
    t(b + 414) = 130
    t(b + 417) = -1
    t(b + 419) = -1
    t(b + 421) = -1
    t(b + 424) = -1
    t(b + 429) = -1
    t(b + 432) = -1
    t(b + 436) = -1
    t(b + 438) = -1
    t(b + 441) = -1
    t(b + 445) = -1
    t(b + 447) = 56
    t(b + 453) = -1
    t(b + 454) = -2
    t(b + 456) = -1
    t(b + 457) = -2
    t(b + 459) = -1
    t(b + 460) = -2
    For i = b + 462 To b + 476 Step 2: t(i) = -1: Next i
    t(b + 477) = -79
    For i = b + 479 To b + 495 Step 2: t(i) = -1: Next i
    t(b + 498) = -1
    t(b + 499) = -2
    t(b + 501) = -1
    For i = b + 505 To b + 543 Step 2: t(i) = -1: Next i
    For i = b + 547 To b + 563 Step 2: t(i) = -1: Next i
    t(b + 572) = -1
    t(b + 575) = 10815
    t(b + 576) = 10815
    t(b + 578) = -1
    For i = b + 583 To b + 591 Step 2: t(i) = -1: Next i
    t(b + 592) = 10783
    t(b + 593) = 10780
    t(b + 594) = 10782
    t(b + 595) = -210
    t(b + 596) = -206
    t(b + 598) = -205
    t(b + 599) = -205
    t(b + 601) = -202
    t(b + 603) = -203
    t(b + 604) = -23217
    t(b + 608) = -205
    t(b + 609) = -23221
    t(b + 611) = -207
    t(b + 613) = -23256
    t(b + 614) = -23228
    t(b + 616) = -209
    t(b + 617) = -211
    t(b + 618) = -23228
    t(b + 619) = 10743
    t(b + 620) = -23231
    t(b + 623) = -211
    t(b + 625) = 10749
    t(b + 626) = -213
    t(b + 629) = -214
    t(b + 637) = 10727
    t(b + 640) = -218
    t(b + 642) = -23229
    t(b + 643) = -218
    t(b + 647) = -23254
    t(b + 648) = -218
    t(b + 649) = -69
    t(b + 650) = -217
    t(b + 651) = -217
    t(b + 652) = -71
    t(b + 658) = -219
    t(b + 669) = -23275
    t(b + 670) = -23278
    t(b + 837) = 84
    t(b + 881) = -1
    t(b + 883) = -1
    t(b + 887) = -1
    t(b + 891) = 130
    t(b + 892) = 130
    t(b + 893) = 130
    t(b + 940) = -38
    t(b + 941) = -37
    t(b + 942) = -37
    t(b + 943) = -37
    For i = b + 945 To b + 961: t(i) = -32: Next i
    t(b + 962) = -31
    For i = b + 963 To b + 971: t(i) = -32: Next i
    t(b + 972) = -64
    t(b + 973) = -63
    t(b + 974) = -63
    t(b + 976) = -62
    t(b + 977) = -57
    t(b + 981) = -47
    t(b + 982) = -54
    t(b + 983) = -8
    For i = b + 985 To b + 1007 Step 2: t(i) = -1: Next i
    t(b + 1008) = -86
    t(b + 1009) = -80
    t(b + 1010) = 7
    t(b + 1011) = -116
    t(b + 1013) = -96
    t(b + 1016) = -1
    t(b + 1019) = -1
    For i = b + 1072 To b + 1103: t(i) = -32: Next i
    For i = b + 1104 To b + 1119: t(i) = -80: Next i
    For i = b + 1121 To b + 1153 Step 2: t(i) = -1: Next i
    For i = b + 1163 To b + 1215 Step 2: t(i) = -1: Next i
    For i = b + 1218 To b + 1230 Step 2: t(i) = -1: Next i
    t(b + 1231) = -15
    For i = b + 1233 To b + 1327 Step 2: t(i) = -1: Next i
    For i = b + 1377 To b + 1414: t(i) = -48: Next i
    For i = b + 4304 To b + 4346: t(i) = 3008: Next i
    t(b + 4349) = 3008
    t(b + 4350) = 3008
    t(b + 4351) = 3008
    For i = b + 5112 To b + 5117: t(i) = -8: Next i
    t(b + 7296) = -6254
    t(b + 7297) = -6253
    t(b + 7298) = -6244
    t(b + 7299) = -6242
    t(b + 7300) = -6242
    t(b + 7301) = -6243
    t(b + 7302) = -6236
    t(b + 7303) = -6181
    t(b + 7304) = -30270
    t(b + 7545) = -30204
    t(b + 7549) = 3814
    t(b + 7566) = -30152
    For i = b + 7681 To b + 7829 Step 2: t(i) = -1: Next i
    t(b + 7835) = -59
    For i = b + 7841 To b + 7935 Step 2: t(i) = -1: Next i
    For i = b + 7936 To b + 7943: t(i) = 8: Next i
    For i = b + 7952 To b + 7957: t(i) = 8: Next i
    For i = b + 7968 To b + 7975: t(i) = 8: Next i
    For i = b + 7984 To b + 7991: t(i) = 8: Next i
    For i = b + 8000 To b + 8005: t(i) = 8: Next i
    t(b + 8017) = 8
    t(b + 8019) = 8
    t(b + 8021) = 8
    t(b + 8023) = 8
    For i = b + 8032 To b + 8039: t(i) = 8: Next i
    t(b + 8048) = 74
    t(b + 8049) = 74
    t(b + 8050) = 86
    t(b + 8051) = 86
    t(b + 8052) = 86
    t(b + 8053) = 86
    t(b + 8054) = 100
    t(b + 8055) = 100
    t(b + 8056) = 128
    t(b + 8057) = 128
    t(b + 8058) = 112
    t(b + 8059) = 112
    t(b + 8060) = 126
    t(b + 8061) = 126
    t(b + 8112) = 8
    t(b + 8113) = 8
    t(b + 8126) = -7205
    t(b + 8144) = 8
    t(b + 8145) = 8
    t(b + 8160) = 8
    t(b + 8161) = 8
    t(b + 8165) = 7
    t(b + 8526) = -28
    For i = b + 8560 To b + 8575: t(i) = -16: Next i
    t(b + 8580) = -1
    For i = b + 9424 To b + 9449: t(i) = -26: Next i
    For i = b + 11312 To b + 11358: t(i) = -48: Next i
    t(b + 11361) = -1
    t(b + 11365) = -10795
    t(b + 11366) = -10792
    t(b + 11368) = -1
    t(b + 11370) = -1
    t(b + 11372) = -1
    t(b + 11379) = -1
    t(b + 11382) = -1
    For i = b + 11393 To b + 11491 Step 2: t(i) = -1: Next i
    t(b + 11500) = -1
    t(b + 11502) = -1
    t(b + 11507) = -1
    For i = b + 11520 To b + 11557: t(i) = -7264: Next i
    t(b + 11559) = -7264
    t(b + 11565) = -7264
    For i = b + 42561 To b + 42605 Step 2: t(i) = -1: Next i
    For i = b + 42625 To b + 42651 Step 2: t(i) = -1: Next i
    For i = b + 42787 To b + 42799 Step 2: t(i) = -1: Next i
    For i = b + 42803 To b + 42863 Step 2: t(i) = -1: Next i
    t(b + 42874) = -1
    t(b + 42876) = -1
    For i = b + 42879 To b + 42887 Step 2: t(i) = -1: Next i
    t(b + 42892) = -1
    t(b + 42897) = -1
    t(b + 42899) = -1
    t(b + 42900) = 48
    For i = b + 42903 To b + 42921 Step 2: t(i) = -1: Next i
    For i = b + 42933 To b + 42943 Step 2: t(i) = -1: Next i
    t(b + 42947) = -1
    t(b + 43859) = -928
    For i = b + 43888 To b + 43967: t(i) = 26672: Next i
    For i = b + 65345 To b + 65370: t(i) = -32: Next i
End Sub

Private Sub InitializeUnicodeCanonRunsTable(ByRef t() As Long)
    Const b As Long = UNICODE_CANON_RUNS_TABLE_START
    t(b + 0) = 97
    t(b + 1) = 123
    t(b + 2) = 181
    t(b + 3) = 182
    t(b + 4) = 224
    t(b + 5) = 247
    t(b + 6) = 248
    t(b + 7) = 255
    t(b + 8) = 256
    t(b + 9) = 257
    t(b + 10) = 304
    t(b + 11) = 307
    t(b + 12) = 312
    t(b + 13) = 314
    t(b + 14) = 329
    t(b + 15) = 331
    t(b + 16) = 376
    t(b + 17) = 378
    t(b + 18) = 383
    t(b + 19) = 384
    t(b + 20) = 385
    t(b + 21) = 387
    t(b + 22) = 390
    t(b + 23) = 392
    t(b + 24) = 393
    t(b + 25) = 396
    t(b + 26) = 397
    t(b + 27) = 402
    t(b + 28) = 403
    t(b + 29) = 405
    t(b + 30) = 406
    t(b + 31) = 409
    t(b + 32) = 410
    t(b + 33) = 411
    t(b + 34) = 414
    t(b + 35) = 415
    t(b + 36) = 417
    t(b + 37) = 422
    t(b + 38) = 424
    t(b + 39) = 425
    t(b + 40) = 429
    t(b + 41) = 430
    t(b + 42) = 432
    t(b + 43) = 433
    t(b + 44) = 436
    t(b + 45) = 439
    t(b + 46) = 441
    t(b + 47) = 442
    t(b + 48) = 445
    t(b + 49) = 446
    t(b + 50) = 447
    t(b + 51) = 448
    t(b + 52) = 453
    t(b + 53) = 454
    t(b + 54) = 455
    t(b + 55) = 456
    t(b + 56) = 457
    t(b + 57) = 458
    t(b + 58) = 459
    t(b + 59) = 460
    t(b + 60) = 461
    t(b + 61) = 462
    t(b + 62) = 477
    t(b + 63) = 478
    t(b + 64) = 479
    t(b + 65) = 496
    t(b + 66) = 498
    t(b + 67) = 499
    t(b + 68) = 500
    t(b + 69) = 501
    t(b + 70) = 502
    t(b + 71) = 505
    t(b + 72) = 544
    t(b + 73) = 547
    t(b + 74) = 564
    t(b + 75) = 572
    t(b + 76) = 573
    t(b + 77) = 575
    t(b + 78) = 577
    t(b + 79) = 578
    t(b + 80) = 579
    t(b + 81) = 583
    t(b + 82) = 592
    t(b + 83) = 593
    t(b + 84) = 594
    t(b + 85) = 595
    t(b + 86) = 596
    t(b + 87) = 597
    t(b + 88) = 598
    t(b + 89) = 600
    t(b + 90) = 601
    t(b + 91) = 602
    t(b + 92) = 603
    t(b + 93) = 604
    t(b + 94) = 605
    t(b + 95) = 608
    t(b + 96) = 609
    t(b + 97) = 610
    t(b + 98) = 611
    t(b + 99) = 612
    t(b + 100) = 613
    t(b + 101) = 614
    t(b + 102) = 615
    t(b + 103) = 616
    t(b + 104) = 617
    t(b + 105) = 618
    t(b + 106) = 619
    t(b + 107) = 620
    t(b + 108) = 621
    t(b + 109) = 623
    t(b + 110) = 624
    t(b + 111) = 625
    t(b + 112) = 626
    t(b + 113) = 627
    t(b + 114) = 629
    t(b + 115) = 630
    t(b + 116) = 637
    t(b + 117) = 638
    t(b + 118) = 640
    t(b + 119) = 641
    t(b + 120) = 642
    t(b + 121) = 643
    t(b + 122) = 644
    t(b + 123) = 647
    t(b + 124) = 648
    t(b + 125) = 649
    t(b + 126) = 650
    t(b + 127) = 652
    t(b + 128) = 653
    t(b + 129) = 658
    t(b + 130) = 659
    t(b + 131) = 669
    t(b + 132) = 670
    t(b + 133) = 671
    t(b + 134) = 837
    t(b + 135) = 838
    t(b + 136) = 881
    t(b + 137) = 884
    t(b + 138) = 887
    t(b + 139) = 888
    t(b + 140) = 891
    t(b + 141) = 894
    t(b + 142) = 940
    t(b + 143) = 941
    t(b + 144) = 944
    t(b + 145) = 945
    t(b + 146) = 962
    t(b + 147) = 963
    t(b + 148) = 972
    t(b + 149) = 973
    t(b + 150) = 975
    t(b + 151) = 976
    t(b + 152) = 977
    t(b + 153) = 978
    t(b + 154) = 981
    t(b + 155) = 982
    t(b + 156) = 983
    t(b + 157) = 984
    t(b + 158) = 985
    t(b + 159) = 1008
    t(b + 160) = 1009
    t(b + 161) = 1010
    t(b + 162) = 1011
    t(b + 163) = 1012
    t(b + 164) = 1013
    t(b + 165) = 1014
    t(b + 166) = 1016
    t(b + 167) = 1017
    t(b + 168) = 1019
    t(b + 169) = 1020
    t(b + 170) = 1072
    t(b + 171) = 1104
    t(b + 172) = 1120
    t(b + 173) = 1121
    t(b + 174) = 1154
    t(b + 175) = 1163
    t(b + 176) = 1216
    t(b + 177) = 1218
    t(b + 178) = 1231
    t(b + 179) = 1232
    t(b + 180) = 1233
    t(b + 181) = 1328
    t(b + 182) = 1377
    t(b + 183) = 1415
    t(b + 184) = 4304
    t(b + 185) = 4347
    t(b + 186) = 4349
    t(b + 187) = 4352
    t(b + 188) = 5112
    t(b + 189) = 5118
    t(b + 190) = 7296
    t(b + 191) = 7297
    t(b + 192) = 7298
    t(b + 193) = 7299
    t(b + 194) = 7301
    t(b + 195) = 7302
    t(b + 196) = 7303
    t(b + 197) = 7304
    t(b + 198) = 7305
    t(b + 199) = 7545
    t(b + 200) = 7546
    t(b + 201) = 7549
    t(b + 202) = 7550
    t(b + 203) = 7566
    t(b + 204) = 7567
    t(b + 205) = 7681
    t(b + 206) = 7830
    t(b + 207) = 7835
    t(b + 208) = 7836
    t(b + 209) = 7841
    t(b + 210) = 7936
    t(b + 211) = 7944
    t(b + 212) = 7952
    t(b + 213) = 7958
    t(b + 214) = 7968
    t(b + 215) = 7976
    t(b + 216) = 7984
    t(b + 217) = 7992
    t(b + 218) = 8000
    t(b + 219) = 8006
    t(b + 220) = 8017
    t(b + 221) = 8024
    t(b + 222) = 8032
    t(b + 223) = 8040
    t(b + 224) = 8048
    t(b + 225) = 8050
    t(b + 226) = 8054
    t(b + 227) = 8056
    t(b + 228) = 8058
    t(b + 229) = 8060
    t(b + 230) = 8062
    t(b + 231) = 8112
    t(b + 232) = 8114
    t(b + 233) = 8126
    t(b + 234) = 8127
    t(b + 235) = 8144
    t(b + 236) = 8146
    t(b + 237) = 8160
    t(b + 238) = 8162
    t(b + 239) = 8165
    t(b + 240) = 8166
    t(b + 241) = 8526
    t(b + 242) = 8527
    t(b + 243) = 8560
    t(b + 244) = 8576
    t(b + 245) = 8580
    t(b + 246) = 8581
    t(b + 247) = 9424
    t(b + 248) = 9450
    t(b + 249) = 11312
    t(b + 250) = 11359
    t(b + 251) = 11361
    t(b + 252) = 11362
    t(b + 253) = 11365
    t(b + 254) = 11366
    t(b + 255) = 11367
    t(b + 256) = 11368
    t(b + 257) = 11373
    t(b + 258) = 11379
    t(b + 259) = 11380
    t(b + 260) = 11382
    t(b + 261) = 11383
    t(b + 262) = 11393
    t(b + 263) = 11492
    t(b + 264) = 11500
    t(b + 265) = 11503
    t(b + 266) = 11507
    t(b + 267) = 11508
    t(b + 268) = 11520
    t(b + 269) = 11558
    t(b + 270) = 11559
    t(b + 271) = 11560
    t(b + 272) = 11565
    t(b + 273) = 11566
    t(b + 274) = 42561
    t(b + 275) = 42606
    t(b + 276) = 42625
    t(b + 277) = 42652
    t(b + 278) = 42787
    t(b + 279) = 42800
    t(b + 280) = 42803
    t(b + 281) = 42864
    t(b + 282) = 42874
    t(b + 283) = 42877
    t(b + 284) = 42879
    t(b + 285) = 42888
    t(b + 286) = 42892
    t(b + 287) = 42893
    t(b + 288) = 42897
    t(b + 289) = 42900
    t(b + 290) = 42901
    t(b + 291) = 42903
    t(b + 292) = 42922
    t(b + 293) = 42933
    t(b + 294) = 42944
    t(b + 295) = 42947
    t(b + 296) = 42948
    t(b + 297) = 43859
    t(b + 298) = 43860
    t(b + 299) = 43888
    t(b + 300) = 43968
    t(b + 301) = 65345
    t(b + 302) = 65371
End Sub

Private Sub EmitPredefinedRange( _
    ByRef outBuffer As ArrayBuffer, ByRef source() As Long, ByVal sourceStart As Long, ByVal sourceLength As Long _
)
    Dim i As Long, j As Long
    
    With outBuffer
        i = .Length
        AppendUnspecified outBuffer, sourceLength
        j = sourceStart + sourceLength - 2
        Do While j >= sourceStart
            .Buffer(i) = source(j): i = i + 1
            .Buffer(i) = source(j + 1): i = i + 1
            j = j - 2
        Loop
    End With
End Sub

Private Function UnicodeIsIdentifierPart(x As Long) As Boolean
    UnicodeIsIdentifierPart = False
End Function

Private Sub RegexpGenerateRanges(ByRef outBuffer As ArrayBuffer, _
    ByVal caseInsensitive As Boolean, ByVal r1 As Long, ByVal r2 As Long _
)
    Dim a As Long, b As Long, m As Long, d As Long, ub As Long, lastDelta As Long
    Dim r As Long, rc As Long

    If Not caseInsensitive Then
        AppendLong outBuffer, r1
        AppendLong outBuffer, r2
        Exit Sub
    End If
        
    rc = ReCanonicalizeChar(r1)
    lastDelta = rc - r1
    AppendLong outBuffer, rc

    a = UNICODE_CANON_RUNS_TABLE_START - 1
    ub = a + UNICODE_CANON_RUNS_TABLE_LENGTH
    
    If StaticData(ub) > r1 Then
        ' Find the index of the first element larger than r1.
        ' The index is guaranteed to be in the interval (a;b].
        b = ub
        Do
            d = b - a
            If d = 1 Then Exit Do
            m = a + d \ 2
            If StaticData(m) > r1 Then b = m Else a = m
        Loop
        
        ' Now b is the index of the first element larger than r1.
        Do
            r = StaticData(b)
            If r > r2 Then Exit Do
            AppendLong outBuffer, r - 1 + lastDelta
            
            rc = ReCanonicalizeChar(r)
            AppendLong outBuffer, rc
            lastDelta = rc - r

            If b = ub Then Exit Do
            b = b + 1
        Loop
    End If
    
    AppendLong outBuffer, r2 + lastDelta
End Sub

Private Sub RedBlackDumpTree(ByRef target() As Long, ByVal targetStartIdx As Long, ByRef tree As IdentifierTreeTy)
    Dim targetIdx As Long, currentNode As Long, nextNode As Long
    
    If tree.nEntries = 0 Then Exit Sub
    
    targetIdx = targetStartIdx
    
    currentNode = tree.root
    Do
        nextNode = tree.Buffer(currentNode).rbChild(0)
        If nextNode = -1 Then Exit Do
        currentNode = nextNode
    Loop
    
    Do
        ' handle current node
        target(targetIdx) = tree.Buffer(currentNode).reference.start: targetIdx = targetIdx + 1
        target(targetIdx) = tree.Buffer(currentNode).reference.Length: targetIdx = targetIdx + 1
        target(targetIdx) = tree.Buffer(currentNode).identifierId: targetIdx = targetIdx + 1
        
        nextNode = tree.Buffer(currentNode).rbChild(1)
        If nextNode <> -1 Then
            ' right-hand child exists
            ' -> go to leftmost descendant of right-hand child
            currentNode = nextNode
            Do
                nextNode = tree.Buffer(currentNode).rbChild(0)
                If nextNode = -1 Then Exit Do
                currentNode = nextNode
            Loop
        Else
            ' right-hand child does not exist
            ' -> go to first ancestor for which our subtree is the left-hand subtree
            Do
                nextNode = tree.Buffer(currentNode).rbParent
                If nextNode = -1 Then Exit Sub
                If currentNode = tree.Buffer(nextNode).rbChild(0) Then Exit Do
                currentNode = nextNode
            Loop
            currentNode = nextNode
        End If
    Loop
End Sub

Private Function RedBlackFindOrInsert(ByRef lake As String, ByRef tree As IdentifierTreeTy, ByRef vReference As StartLengthPair) As Long
    Dim parent As Long, found As Long
    Dim asRightHandChild As Boolean
    
    found = RedBlackFindPosition(parent, asRightHandChild, lake, tree, vReference)
    If found = -1 Then
        If tree.nEntries = tree.bufferCapacity Then
            tree.bufferCapacity = tree.bufferCapacity + (tree.bufferCapacity + 16) \ 2
            ReDim Preserve tree.Buffer(0 To tree.bufferCapacity - 1) As IdentifierTreeNode
        End If
        With tree.Buffer(tree.nEntries)
            .reference = vReference
            .identifierId = tree.nEntries
        End With
        RedBlackInsert tree, tree.nEntries, parent, asRightHandChild
        RedBlackFindOrInsert = tree.nEntries
        tree.nEntries = tree.nEntries + 1
    Else
        RedBlackFindOrInsert = tree.Buffer(found).identifierId
    End If
End Function

Private Function RedBlackComparator(ByRef lake As String, ByRef v1 As StartLengthPair, ByRef v2 As StartLengthPair) As Long
    If v1.Length < v2.Length Then RedBlackComparator = -1: Exit Function
    If v1.Length > v2.Length Then RedBlackComparator = 1: Exit Function
    RedBlackComparator = StrComp( _
        Mid$(lake, v1.start, v1.Length), _
        Mid$(lake, v2.start, v2.Length), _
        vbBinaryCompare _
    )
End Function

Private Function RedBlackFindPosition( _
    ByRef outParent As Long, ByRef outAsRightHandChild As Boolean, _
    ByRef lake As String, ByRef tree As IdentifierTreeTy, _
    ByRef vReference As StartLengthPair _
) As Long
    Dim cmp As Long, cur As Long, p As Long, rhc As Boolean
    
    cur = tree.root: p = -1: rhc = False
    Do Until cur = -1
        cmp = RedBlackComparator(lake, vReference, tree.Buffer(cur).reference)
        If cmp < 0 Then
            p = cur: rhc = False
            cur = tree.Buffer(cur).rbChild(0)
        ElseIf cmp = 0 Then
            RedBlackFindPosition = cur
            Exit Function
        Else
            p = cur: rhc = True
            cur = tree.Buffer(cur).rbChild(1)
        End If
    Loop
    outParent = p: outAsRightHandChild = rhc: RedBlackFindPosition = -1
End Function

Private Sub RedBlackInsert( _
    ByRef tree As IdentifierTreeTy, _
    ByVal newNode As Long, ByVal parent As Long, ByVal asRightHandChild As Boolean _
)
    Dim g As Long, u As Long, p As Long, n As Long, pIsRhc As Boolean
    Dim gg As Long, b As Long, c As Long, x As Long, y As Long, z As Long, nIsRhc As Boolean
    
    With tree.Buffer(newNode)
        .rbIsBlack = False
        .rbChild(0) = -1
        .rbChild(1) = -1
        .rbParent = parent
    End With
    If parent = -1 Then
        tree.root = newNode
        Exit Sub
    End If
    
    tree.Buffer(parent).rbChild(-asRightHandChild) = newNode
    
    n = newNode: p = parent
    Do
        If tree.Buffer(p).rbIsBlack Then Exit Sub
        ' p red
        g = tree.Buffer(p).rbParent
        If g = -1 Then ' p red and root
            tree.Buffer(p).rbIsBlack = True
            Exit Sub
        End If
        ' p red and not root (g exists)
        ' u is supposed to refer to the brother of p
        pIsRhc = tree.Buffer(g).rbChild(1) = p
        u = tree.Buffer(g).rbChild(1 + pIsRhc)
        If u = -1 Then GoTo ExitWithRotation
        If tree.Buffer(u).rbIsBlack Then GoTo ExitWithRotation

        ' p and u red, g exists
        tree.Buffer(p).rbIsBlack = True
        tree.Buffer(u).rbIsBlack = True
        tree.Buffer(g).rbIsBlack = False
        n = g
        p = tree.Buffer(n).rbParent
    Loop Until p = -1
    Exit Sub

ExitWithRotation: ' p red and u black (or does not exist), g exists
    ' For an explanation of the following, see
    '   https://en.wikibooks.org/w/index.php?title=F_Sharp_Programming/Advanced_Data_Structures&oldid=4052491 ,
    '   Section 3.1 ("Red Black Trees"), second diagram (following the sentence "The center tree is the balanced version.").
    nIsRhc = tree.Buffer(p).rbChild(1) = n
    If pIsRhc = nIsRhc Then ' outer child
        y = p
        If pIsRhc Then
            b = tree.Buffer(p).rbChild(0): c = tree.Buffer(n).rbChild(0): x = g: z = n
        Else
            b = tree.Buffer(n).rbChild(1): c = tree.Buffer(p).rbChild(1): x = n: z = g
        End If
    Else ' inner child
        y = n: With tree.Buffer(n): b = .rbChild(0): c = .rbChild(1): End With
        If pIsRhc Then
            x = g: z = p
        Else
            x = p: z = g
        End If
    End If
    
    gg = tree.Buffer(g).rbParent
    
    With tree.Buffer(x): .rbIsBlack = False: .rbParent = y: .rbChild(1) = b: End With
    With tree.Buffer(y): .rbIsBlack = True: .rbParent = gg: .rbChild(0) = x: .rbChild(1) = z: End With
    With tree.Buffer(z): .rbIsBlack = False: .rbParent = y: .rbChild(0) = c: End With
    
    If b <> -1 Then tree.Buffer(b).rbParent = x
    If c <> -1 Then tree.Buffer(c).rbParent = z
    
    If gg = -1 Then
        tree.root = y
    Else
        With tree.Buffer(gg): .rbChild(-(.rbChild(1) = g)) = y: End With
    End If
End Sub

Private Sub Initialize(ByRef lexCtx As LexerContext, ByRef inputStr As String)
    With lexCtx
        .inputStr = inputStr
        .iEnd = Len(.inputStr)
        .iCurrent = 0
        .currentCharacter = Not LEXER_ENDOFINPUT ' value does not matter, as long as it does not equal LEXER_ENDOFINPUT
        
        .identifierTree.nEntries = 0
        .identifierTree.root = -1
    End With
    Advance lexCtx
End Sub

Private Sub ParseReToken(ByRef lexCtx As LexerContext, ByRef outToken As ReToken)
    Dim x As Long
    
    ' used only locally
    Dim i As Long, val1 As Long, val2 As Long, digits As Long, tmp As Long
    
    ' Todo: remove
    Dim slp As StartLengthPair

    Dim emptyReToken As ReToken ' effectively a constant -- zeroed out by default
    outToken = emptyReToken

    x = Advance(lexCtx)
    Select Case x
    Case UNICODE_PIPE
        outToken.t = RETOK_DISJUNCTION
    Case UNICODE_CARET
        outToken.t = RETOK_ASSERT_START
    Case UNICODE_DOLLAR
        outToken.t = RETOK_ASSERT_END
    Case UNICODE_QUESTION
        With outToken
            .qmin = 0
            .qmax = 1
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_STAR
        With outToken
            .qmin = 0
            .qmax = RE_QUANTIFIER_INFINITE
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_PLUS
        With outToken
            .qmin = 1
            .qmax = RE_QUANTIFIER_INFINITE
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_LCURLY
        ' Production allows 'DecimalDigits', including leading zeroes
        val1 = 0
        val2 = RE_QUANTIFIER_INFINITE
        
        digits = 0

        Do
            x = Advance(lexCtx)
            If (x >= UNICODE_0) And (x <= UNICODE_9) Then
                digits = digits + 1
                ' Be careful to prevent overflow
                If val1 > LONG_MAX_DIV_10 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                val1 = val1 * 10
                tmp = x - UNICODE_0
                If LONG_MAX - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                val1 = val1 + tmp
            ElseIf x = UNICODE_COMMA Then
                If val2 <> RE_QUANTIFIER_INFINITE Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                If lexCtx.currentCharacter = UNICODE_RCURLY Then
                    ' form: { DecimalDigits , }, val1 = min count
                    If digits = 0 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                    outToken.qmin = val1
                    outToken.qmax = RE_QUANTIFIER_INFINITE
                    Advance lexCtx
                    Exit Do
                End If
                val2 = val1
                val1 = 0
                digits = 0 ' not strictly necessary because of lookahead '}' above
            ElseIf x = UNICODE_RCURLY Then
                If digits = 0 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                If val2 <> RE_QUANTIFIER_INFINITE Then
                    ' val2 = min count, val1 = max count
                    outToken.qmin = val2
                    outToken.qmax = val1
                Else
                    ' val1 = count
                    outToken.qmin = val1
                    outToken.qmax = val1
                End If
                Exit Do
            Else
                Err.Raise REGEX_ERR_INVALID_QUANTIFIER
            End If
        Loop
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            outToken.greedy = False
            Advance lexCtx
        Else
            outToken.greedy = True
        End If
        outToken.t = RETOK_QUANTIFIER
    Case UNICODE_PERIOD
        outToken.t = RETOK_ATOM_PERIOD
    Case UNICODE_BACKSLASH
        ' The E5.1 specification does not seem to allow IdentifierPart characters
        ' to be used as identity escapes.  Unfortunately this includes '$', which
        ' cannot be escaped as '\$'; it needs to be escaped e.g. as '\u0024'.
        ' Many other implementations (including V8 and Rhino, for instance) do
        ' accept '\$' as a valid identity escape, which is quite pragmatic, and
        ' ES2015 Annex B relaxes the rules to allow these (and other) real world forms.
        x = Advance(lexCtx)
        Select Case x
        Case UNICODE_LC_B
            outToken.t = RETOK_ASSERT_WORD_BOUNDARY
        Case UNICODE_UC_B
            outToken.t = RETOK_ASSERT_NOT_WORD_BOUNDARY
        Case UNICODE_LC_F
            outToken.num = &HC&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_N
            outToken.num = &HA&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_T
            outToken.num = &H9&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_R
            outToken.num = &HD&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_V
            outToken.num = &HB&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_C
            x = Advance(lexCtx)
            If (x >= UNICODE_LC_A And x <= UNICODE_LC_Z) Or (x >= UNICODE_UC_A And x <= UNICODE_UC_Z) Then
                outToken.num = x \ 32
                outToken.t = RETOK_ATOM_CHAR
            Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            End If
        Case UNICODE_LC_X
            outToken.num = LexerParseEscapeX(lexCtx)
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_U
            ' Todo: What does the following mean?
            ' The token value is the Unicode codepoint without
            ' it being decode into surrogate pair characters
            ' here.  The \u{H+} is only allowed in Unicode mode
            ' which we don't support yet.
            outToken.num = LexerParseEscapeU(lexCtx)
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_D
            outToken.t = RETOK_ATOM_DIGIT
        Case UNICODE_UC_D
            outToken.t = RETOK_ATOM_NOT_DIGIT
        Case UNICODE_LC_S
            outToken.t = RETOK_ATOM_WHITE
        Case UNICODE_UC_S
            outToken.t = RETOK_ATOM_NOT_WHITE
        Case UNICODE_LC_W
            outToken.t = RETOK_ATOM_WORD_CHAR
        Case UNICODE_UC_W
            outToken.t = RETOK_ATOM_NOT_WORD_CHAR
        Case UNICODE_0
            x = Advance(lexCtx)
            
            ' E5 Section 15.10.2.11
            If x >= UNICODE_0 And x <= UNICODE_9 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            outToken.num = 0
            outToken.t = RETOK_ATOM_CHAR
        Case Else
            If x >= UNICODE_1 And x <= UNICODE_9 Then
                val1 = 0
                i = 0
                Do
                    ' We have to be careful here to make sure there will be no overflow.
                    ' 2^31 - 1 backreferences is a bit ridiculous, though.
                    If val1 > LONG_MAX_DIV_10 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
                    val1 = val1 * 10
                    tmp = x - UNICODE_0
                    If LONG_MAX - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
                    val1 = val1 + tmp
                    x = lexCtx.currentCharacter
                    If x < UNICODE_0 Or x > UNICODE_9 Then Exit Do
                    Advance lexCtx
                    i = i + 1
                Loop
                outToken.t = RETOK_ATOM_BACKREFERENCE
                outToken.num = val1
            ElseIf (x >= 0 And Not UnicodeIsIdentifierPart(0)) Or x = UNICODE_CP_ZWNJ Or x = UNICODE_CP_ZWJ Then
                ' For ES5.1 identity escapes are not allowed for identifier
                ' parts.  This conflicts with a lot of real world code as this
                ' doesn't e.g. allow escaping a dollar sign as /\$/.
                outToken.num = x
                outToken.t = RETOK_ATOM_CHAR
            Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            End If
        End Select
    Case UNICODE_LPAREN
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            Advance lexCtx
            x = Advance(lexCtx)
            Select Case x
            Case UNICODE_EQUALS
                ' (?=
                outToken.t = RETOK_ASSERT_START_POS_LOOKAHEAD
            Case UNICODE_EXCLAMATION
                ' (?!
                outToken.t = RETOK_ASSERT_START_NEG_LOOKAHEAD
            Case UNICODE_COLON
                ' (?:
                outToken.t = RETOK_ATOM_START_NONCAPTURE_GROUP
            Case UNICODE_LT
                x = Advance(lexCtx)
                If x = UNICODE_EQUALS Then
                    outToken.t = RETOK_ASSERT_START_POS_LOOKBEHIND
                ElseIf x = UNICODE_EXCLAMATION Then
                    outToken.t = RETOK_ASSERT_START_NEG_LOOKBEHIND
                ElseIf IsIdentifierChar(x) Then
                    With lexCtx
                        val1 = .identifierTree.nEntries
                        val2 = .iCurrent - 1
                        Do
                            x = Advance(lexCtx)
                            If x = UNICODE_GT Then Exit Do
                            ' Todo: Allow unicode escape sequences
                            If Not IsIdentifierChar(x) Then Err.Raise REGEX_ERR_INVALID_IDENTIFIER
                        Loop
                        outToken.t = RETOK_ATOM_START_CAPTURE_GROUP
                        With slp
                            .start = val2: .Length = lexCtx.iCurrent - 1 - val2
                        End With
                        outToken.num = RedBlackFindOrInsert( _
                            lexCtx.inputStr, _
                            lexCtx.identifierTree, _
                            slp _
                        )
                    End With
                Else
                    Err.Raise REGEX_ERR_INVALID_REGEXP_GROUP
                End If
            Case Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_GROUP
            End Select
        Else
            ' (
            outToken.t = RETOK_ATOM_START_CAPTURE_GROUP
            outToken.num = -1
        End If
    Case UNICODE_RPAREN
        outToken.t = RETOK_ATOM_END
    Case UNICODE_LBRACKET
        ' To avoid creating a heavy intermediate value for the list of ranges,
        ' only the start token ('[' or '[^') is parsed here.  The regexp
        ' compiler parses the ranges itself.
        If lexCtx.currentCharacter = UNICODE_CARET Then
            Advance lexCtx
            outToken.t = RETOK_ATOM_START_CHARCLASS_INVERTED
        Else
            outToken.t = RETOK_ATOM_START_CHARCLASS
        End If
    Case UNICODE_RCURLY, UNICODE_RBRACKET
        ' Although these could be parsed as PatternCharacters unambiguously (here),
        ' * E5 Section 15.10.1 grammar explicitly forbids these as PatternCharacters.
        Err.Raise REGEX_ERR_INVALID_REGEXP_CHARACTER
    Case LEXER_ENDOFINPUT
        ' EOF
        outToken.t = RETOK_EOF
    Case Else
        ' PatternCharacter, all excluded characters are matched by cases above
        outToken.t = RETOK_ATOM_CHAR
        outToken.num = x
    End Select
End Sub

Private Sub ParseReRanges(lexCtx As LexerContext, ByRef outBuffer As ArrayBuffer, ByRef nranges As Long, ByVal ignoreCase As Boolean)
    Dim start As Long, ch As Long, x As Long, dash As Boolean, y As Long, bufferStart As Long
    
    bufferStart = outBuffer.Length
    
    ' start is -2 at the very beginning of the range expression,
    '   -1 when we have not seen a possible "start" character,
    '   and it equals the possible start character if we have seen one
    start = -2
    dash = False
    
    Do
ContinueLoop:
        x = Advance(lexCtx)

        If x < 0 Then GoTo FailUntermCharclass
        
        Select Case x
        Case UNICODE_RBRACKET
            If start >= 0 Then
                RegexpGenerateRanges outBuffer, ignoreCase, start, start
                Exit Do
            ElseIf start = -1 Then
                Exit Do
            Else ' start = -2
                ' ] at the very beginning of a range expression is interpreted literally,
                '   since empty ranges are not permitted.
                '   This corresponds to what RE2 does.
                ch = x
            End If
        Case UNICODE_MINUS
            If start >= 0 Then
                If Not dash Then
                    If lexCtx.currentCharacter <> UNICODE_RBRACKET Then
                        ' '-' as a range indicator
                        dash = True
                        GoTo ContinueLoop
                    End If
                End If
            End If
            ' '-' verbatim
            ch = x
        Case UNICODE_BACKSLASH
            '
            '  The escapes are same as outside a character class, except that \b has a
            '  different meaning, and \B and backreferences are prohibited (see E5
            '  Section 15.10.2.19).  However, it's difficult to share code because we
            '  handle e.g. "\n" very differently: here we generate a single character
            '  range for it.
            '

            ' XXX: ES2015 surrogate pair handling.

            x = Advance(lexCtx)

            Select Case x
            Case UNICODE_LC_B
                ' Note: '\b' in char class is different than outside (assertion),
                ' '\B' is not allowed and is caught by the duk_unicode_is_identifier_part()
                ' check below.
                '
                ch = &H8&
            Case x = UNICODE_LC_F
                ch = &HC&
            Case UNICODE_LC_N
                ch = &HA&
            Case UNICODE_LC_T
                ch = &H9&
            Case UNICODE_LC_R
                ch = &HD&
            Case UNICODE_LC_V
                ch = &HB&
            Case UNICODE_LC_C
                x = Advance(lexCtx)
                If ((x >= UNICODE_LC_A And x <= UNICODE_LC_Z) Or (x >= UNICODE_UC_A And x <= UNICODE_UC_Z)) Then
                    ch = x Mod 32
                Else
                    GoTo FailEscape
                End If
            Case UNICODE_LC_X
                ch = LexerParseEscapeX(lexCtx)
            Case UNICODE_LC_U
                ch = LexerParseEscapeU(lexCtx)
            Case UNICODE_LC_D
                EmitPredefinedRange outBuffer, StaticData, RANGE_TABLE_DIGIT_START, RANGE_TABLE_DIGIT_LENGTH
                ch = -1
            Case UNICODE_UC_D
                EmitPredefinedRange outBuffer, StaticData, RANGE_TABLE_NOTDIGIT_START, RANGE_TABLE_NOTDIGIT_LENGTH
                ch = -1
            Case UNICODE_LC_S
                EmitPredefinedRange outBuffer, StaticData, RANGE_TABLE_WHITE_START, RANGE_TABLE_WHITE_LENGTH
                ch = -1
            Case UNICODE_UC_S
                EmitPredefinedRange outBuffer, StaticData, RANGE_TABLE_NOTWHITE_START, RANGE_TABLE_NOTWHITE_LENGTH
                ch = -1
            Case UNICODE_LC_W
                EmitPredefinedRange outBuffer, StaticData, RANGE_TABLE_WORDCHAR_START, RANGE_TABLE_WORDCHAR_LENGTH
                ch = -1
            Case UNICODE_UC_W
                EmitPredefinedRange outBuffer, StaticData, RANGE_TABLE_NOTWORDCHAR_START, RANGE_TABLE_NOTWORDCHAR_LENGTH
                ch = -1
            Case Else
                If x < 0 Then GoTo FailEscape
                If x <= UNICODE_7 Then
                    If x >= UNICODE_0 Then
                        ' \0 or octal escape from \0 up to \377
                        ch = LexerParseLegacyOctal(lexCtx, x)
                    Else
                        ' IdentityEscape: ES2015 Annex B allows almost all
                        ' source characters here.  Match anything except
                        ' EOF here.
                        ch = x
                    End If
                Else
                    ' IdentityEscape: ES2015 Annex B allows almost all
                    ' source characters here.  Match anything except
                    ' EOF here.
                    ch = x
                End If
            End Select
        Case Else
            ' character represents itself
            ch = x
        End Select

        ' ch is a literal character here or -1 if parsed entity was
        ' an escape such as "\s".
        '

        If ch < 0 Then
            ' multi-character sets not allowed as part of ranges, see
            ' E5 Section 15.10.2.15, abstract operation CharacterRange.
            '
            If start >= 0 Then
                If dash Then
                    GoTo FailRange
                Else
                    RegexpGenerateRanges outBuffer, ignoreCase, start, start
                End If
            End If
            start = -1
            ' dash is already 0
        Else
            If start >= 0 Then
                If dash Then
                    If start > ch Then GoTo FailRange
                    RegexpGenerateRanges outBuffer, ignoreCase, start, ch
                    start = -1
                    dash = 0
                Else
                    RegexpGenerateRanges outBuffer, ignoreCase, start, start
                    start = ch
                    ' dash is already 0
                End If
            Else
                start = ch
            End If
        End If
    Loop

    If outBuffer.Length - 2 > bufferStart Then
        ' We have at least 2 intervals.
        HeapsortPairs outBuffer.Buffer, bufferStart, outBuffer.Length - 2
        outBuffer.Length = 2 + Unionize(outBuffer.Buffer, bufferStart, outBuffer.Length - 2)
    End If
    
    nranges = (outBuffer.Length - bufferStart) \ 2
    
    Exit Sub

FailEscape:
    Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE

FailRange:
    Err.Raise REGEX_ERR_INVALID_RANGE

FailUntermCharclass:
    Err.Raise REGEX_ERR_UNTERMINATED_CHARCLASS
End Sub

Private Sub HeapsortPairs(ByRef ary() As Long, ByVal b As Long, ByVal t As Long)
    Dim bb As Long
    Dim parent As Long, child As Long
    Dim smallestValueX As Long, smallestValueY As Long, tmpX As Long, tmpY As Long
    
    ' build heap
    ' bb marks the next element to be added to the heap
    bb = t - 2
    Do Until bb < b
        child = bb
        Do Until child = t
            parent = child + 2 + 2 * ((t - child) \ 4)
            If ary(parent) <= ary(child) Then Exit Do
            tmpX = ary(parent): tmpY = ary(parent + 1)
            ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
            ary(child) = tmpX: ary(child + 1) = tmpY
            child = parent
        Loop
        bb = bb - 2
    Loop

    ' demount heap
    ' bb marks the lower end of the remaining heap
    bb = b
    Do While bb < t
        smallestValueX = ary(t): smallestValueY = ary(t + 1)
        
        parent = t
        Do
            child = parent - t + parent - 2
            
            ' if there are no children, we are finished
            If child < bb Then Exit Do
            
            ' if there are two children, prefer the one with the smaller value
            If child > bb Then child = child + 2 * (ary(child - 2) <= ary(child))
            
            ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
            parent = child
        Loop
        
        ' now position parent is free
        
        ' if parent <> bb, free bb rather than parent
        ' by swapping the values in parent and bb and repairing the heap bottom-up
        If parent > bb Then
            ary(parent) = ary(bb): ary(parent + 1) = ary(bb + 1)
            child = parent
            Do Until child = t
                parent = child + 2 + 2 * ((t - child) \ 4)
                If ary(parent) <= ary(child) Then Exit Do
                tmpX = ary(parent): tmpY = ary(parent + 1)
                ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
                ary(child) = tmpX: ary(child + 1) = tmpY
                child = parent
            Loop
        End If
        
        ' now position bb is free
        
        ary(bb) = smallestValueX: ary(bb + 1) = smallestValueY
        bb = bb + 2
    Loop
End Sub

Private Function Unionize(ByRef ary() As Long, ByVal b As Long, ByVal t As Long)
    Dim i As Long, j As Long, lower As Long, upper As Long, nextLower As Long, nextUpper As Long
    
    lower = ary(b): upper = ary(b + 1)
    j = b
    For i = b + 2 To t Step 2
        nextLower = ary(i): nextUpper = ary(i + 1)
        If nextLower <= upper + 1 Then
            If nextUpper > upper Then upper = nextUpper
        Else
            ary(j) = lower: j = j + 1: ary(j) = upper: j = j + 1
            lower = nextLower: upper = nextUpper
        End If
    Next
    ary(j) = lower: ary(j + 1) = upper
    Unionize = j
End Function

Private Function LexerParseEscapeX(ByRef lexCtx As LexerContext) As Long
    Dim dig As Long, escval As Long, x As Long
    
    x = Advance(lexCtx)
    dig = HexvalValidate(x)
    If dig < 0 Then GoTo FailEscape
    escval = dig
    
    x = Advance(lexCtx)
    dig = HexvalValidate(x)
    If dig < 0 Then GoTo FailEscape
    escval = escval * 16 + dig
    
    LexerParseEscapeX = escval
    Exit Function
    
FailEscape:
    Err.Raise REGEX_ERR_INVALID_ESCAPE
End Function

Private Function LexerParseEscapeU(ByRef lexCtx As LexerContext) As Long
    Dim dig As Long, escval As Long, x As Long
    
    If lexCtx.currentCharacter = UNICODE_LCURLY Then
        Advance lexCtx
        
        escval = 0
        x = Advance(lexCtx)
        If x = UNICODE_RCURLY Then GoTo FailEscape ' Empty escape \u{}
        Do
            dig = HexvalValidate(x)
            If dig < 0 Then GoTo FailEscape
            If escval > &H10FFF Then GoTo FailEscape
            escval = escval * 16 + dig
            
            x = Advance(lexCtx)
        Loop Until x = UNICODE_RCURLY
        LexerParseEscapeU = escval
    Else
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        LexerParseEscapeU = escval
    End If
    Exit Function
    
FailEscape:
    Err.Raise REGEX_ERR_INVALID_ESCAPE
End Function

Private Function HexvalValidate(ByVal ch As Long) As Long
    Const HEX_DELTA_L As Long = UNICODE_LC_A - 10
    Const HEX_DELTA_U As Long = UNICODE_UC_A - 10

    HexvalValidate = -1
    If ch <= UNICODE_UC_F Then
        If ch <= UNICODE_9 Then
            If ch >= UNICODE_0 Then HexvalValidate = ch - UNICODE_0
        Else
            If ch >= UNICODE_UC_A Then HexvalValidate = ch - HEX_DELTA_U
        End If
    Else
        If ch <= UNICODE_LC_F Then
            If ch >= UNICODE_LC_A Then HexvalValidate = ch - HEX_DELTA_L
        End If
    End If
End Function

Private Function LexerParseLegacyOctal(ByRef lexCtx As LexerContext, ByVal x As Long)
    Dim cp As Long, tmp As Long, i As Long

    cp = x - UNICODE_0

    tmp = lexCtx.currentCharacter
    If tmp < UNICODE_0 Then GoTo ExitFunction
    If tmp > UNICODE_7 Then GoTo ExitFunction

    cp = cp * 8 + (tmp - UNICODE_0)
    Advance lexCtx

    If cp > 31 Then GoTo ExitFunction
    
    tmp = lexCtx.currentCharacter
    If tmp < UNICODE_0 Then GoTo ExitFunction
    If tmp > UNICODE_7 Then GoTo ExitFunction

    cp = cp * 8 + (tmp - UNICODE_0)
    Advance lexCtx

ExitFunction:
    LexerParseLegacyOctal = cp
End Function

Private Function IsIdentifierChar(ByVal c As Long) As Boolean
    ' Todo: Temporary Hack.
    IsIdentifierChar = ((c >= AscW("A")) And (c <= AscW("Z"))) Or ((c >= AscW("a")) And (c <= AscW("z")))
End Function

Private Function Advance(ByRef lexCtx As LexerContext) As Long
    Dim lower As Long, upper As Long

    With lexCtx
        Advance = .currentCharacter
        If .currentCharacter = LEXER_ENDOFINPUT Then Exit Function
        If .iCurrent = .iEnd Then
            .currentCharacter = LEXER_ENDOFINPUT
        Else
            .iCurrent = .iCurrent + 1
            .currentCharacter = AscW(Mid$(.inputStr, .iCurrent, 1)) And &HFFFF&
        End If
    End With
End Function

Private Sub Compile(ByRef outBytecode() As Long, ByRef s As String, Optional ByVal caseInsensitive As Boolean = False)
    Dim lex As LexerContext
    Dim ast As ArrayBuffer
    
    If Not UnicodeInitialized Then UnicodeInitialize
    If Not RangeTablesInitialized Then RangeTablesInitialize
    
    Initialize lex, s
    Parse lex, caseInsensitive, ast
    AstToBytecode ast.Buffer, lex.identifierTree, caseInsensitive, outBytecode
End Sub

Private Sub PerformPotentialConcat(ByRef ast As ArrayBuffer, ByRef potentialConcat2 As Long, ByRef potentialConcat1 As Long)
    Dim tmp As Long
    If potentialConcat2 <> -1 Then
        tmp = ast.Length
        AppendThree ast, AST_CONCAT, potentialConcat2, potentialConcat1
        potentialConcat2 = -1
        potentialConcat1 = tmp
    End If
End Sub

Private Sub Parse(ByRef lex As LexerContext, ByVal caseInsensitive As Boolean, ByRef ast As ArrayBuffer)
    Dim currToken As ReToken
    Dim currentAstNode As Long
    Dim potentialConcat2 As Long, potentialConcat1 As Long, pendingDisjunction As Long, currentDisjunction As Long
    Dim nCaptures As Long
    Dim tmp As Long, i As Long, qmin As Long, qmax As Long, n1 As Long, n2 As Long
    Dim parseStack As ArrayBuffer
    
    nCaptures = 0
    
    pendingDisjunction = -1
    potentialConcat2 = -1
    potentialConcat1 = -1
    currentDisjunction = -1
    
    AppendLong ast, 0 ' first word will be index of the root node, to be patched in the end

    Do
ContinueLoop:
        ParseReToken lex, currToken
        
        Select Case currToken.t
        Case RETOK_DISJUNCTION
            If potentialConcat1 = -1 Then
                currentAstNode = ast.Length
                AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            currentAstNode = ast.Length
            AppendThree ast, AST_DISJ, potentialConcat1, -1
            potentialConcat1 = -1

            If pendingDisjunction <> -1 Then
                ast.Buffer(pendingDisjunction + 2) = currentAstNode
            Else
                currentDisjunction = currentAstNode
            End If
        
            pendingDisjunction = currentAstNode
        
        Case RETOK_QUANTIFIER
            If potentialConcat1 = -1 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER_NO_ATOM
            
            qmin = currToken.qmin
            qmax = currToken.qmax
            
            If qmin > qmax Then
                currentAstNode = ast.Length
                AppendLong ast, AST_FAIL
                potentialConcat1 = currentAstNode
                GoTo ContinueLoop
            End If
            
            If qmin = 0 Then
                n1 = -1
            ElseIf qmin = 1 Then
                n1 = potentialConcat1
            ElseIf qmin > 1 Then
                currentAstNode = ast.Length
                AppendThree ast, AST_REPEAT_EXACTLY, potentialConcat1, qmin
                n1 = currentAstNode
            Else
                Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR
            End If
            
            If qmax = RE_QUANTIFIER_INFINITE Then
                If currToken.greedy Then tmp = AST_STAR_GREEDY Else tmp = AST_STAR_HUMBLE
                currentAstNode = ast.Length
                AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin = 1 Then
                If currToken.greedy Then tmp = AST_ZEROONE_GREEDY Else tmp = AST_ZEROONE_HUMBLE
                currentAstNode = ast.Length
                AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin > 1 Then
                If currToken.greedy Then tmp = AST_REPEAT_MAX_GREEDY Else tmp = AST_REPEAT_MAX_HUMBLE
                currentAstNode = ast.Length
                AppendThree ast, tmp, potentialConcat1, qmax - qmin
                n2 = currentAstNode
            ElseIf qmax = qmin Then
                n2 = -1
            Else
                Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR
            End If
            
            If n1 = -1 Then
                If n2 = -1 Then
                    currentAstNode = ast.Length
                    AppendLong ast, AST_EMPTY
                    potentialConcat1 = currentAstNode
                Else
                    potentialConcat1 = n2
                End If
            Else
                If n2 = -1 Then
                    potentialConcat1 = n1
                Else
                    currentAstNode = ast.Length
                    AppendThree ast, AST_CONCAT, n1, n2
                    potentialConcat1 = currentAstNode
                End If
            End If
            
        Case RETOK_ATOM_START_CAPTURE_GROUP
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            AppendFive parseStack, _
                nCaptures, currToken.num, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
                
        Case RETOK_ATOM_START_NONCAPTURE_GROUP
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            AppendFive parseStack, _
                -1, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
            
         Case RETOK_ASSERT_START_POS_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            AppendFive parseStack, _
                -(AST_ASSERT_POS_LOOKAHEAD - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            AppendFive parseStack, _
                -(AST_ASSERT_NEG_LOOKAHEAD - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
         Case RETOK_ASSERT_START_POS_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            AppendFive parseStack, _
                -(AST_ASSERT_POS_LOOKBEHIND - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            AppendFive parseStack, _
                -(AST_ASSERT_NEG_LOOKBEHIND - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        
        Case RETOK_ATOM_END
            If parseStack.Length = 0 Then Err.Raise REGEX_ERR_UNEXPECTED_CLOSING_PAREN
            
            ' Close disjunction
            If potentialConcat1 = -1 Then
                currentAstNode = ast.Length
                AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            If pendingDisjunction = -1 Then
                currentDisjunction = potentialConcat1
            Else
                ast.Buffer(pendingDisjunction + 2) = potentialConcat1
            End If

            potentialConcat1 = currentDisjunction
            
            ' Restore variables
            With parseStack
                .Length = .Length - 5
                tmp = .Buffer(.Length)
                n1 = .Buffer(.Length + 1)
                pendingDisjunction = .Buffer(.Length + 2)
                currentDisjunction = .Buffer(.Length + 3)
                potentialConcat2 = .Buffer(.Length + 4) ' This is correct, potentialConcat1 is the new node!
            End With
            
            If tmp > 0 Then ' capture group
                If n1 <> -1 Then
                    currentAstNode = ast.Length
                    AppendFour ast, AST_NAMED, potentialConcat1, n1, tmp
                    potentialConcat1 = currentAstNode
                End If
                currentAstNode = ast.Length
                AppendThree ast, AST_CAPTURE, potentialConcat1, tmp
                potentialConcat1 = currentAstNode
            ElseIf tmp = -1 Then ' non-capture group
                ' don't do anything
            Else ' lookahead or lookbehind
                currentAstNode = ast.Length
                AppendTwo ast, -(tmp + 2) + MIN_AST_CODE, potentialConcat1
                potentialConcat1 = currentAstNode
            End If

        Case RETOK_ATOM_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            
            currentAstNode = ast.Length
            tmp = currToken.num
            If caseInsensitive Then tmp = ReCanonicalizeChar(tmp)
            AppendTwo ast, AST_CHAR, tmp
                
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_PERIOD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendLong ast, AST_PERIOD
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_BACKREFERENCE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendTwo ast, AST_BACKREFERENCE, currToken.num
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ASSERT_START
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendLong ast, AST_ASSERT_START
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_END
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendLong ast, AST_ASSERT_END
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_START_CHARCLASS, RETOK_ATOM_START_CHARCLASS_INVERTED
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            If currToken.t = RETOK_ATOM_START_CHARCLASS Then
                AppendTwo ast, AST_RANGES, 0
            Else
                AppendTwo ast, AST_INVRANGES, 0
            End If
            
            tmp = 0 ' unnecessary, tmp is an output parameter indicating the number of ranges
            ' Todo: Remove that parameter from ParseReRanges -- we can calculate it by comparing
            '   old and new buffer length.
            ParseReRanges lex, ast, tmp, caseInsensitive
            
            ' patch range count
            ast.Buffer(currentAstNode + 1) = tmp
            
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_WORD_BOUNDARY
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendLong ast, AST_ASSERT_WORD_BOUNDARY
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_NOT_WORD_BOUNDARY
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendLong ast, AST_ASSERT_NOT_WORD_BOUNDARY
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_DIGIT
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendPrefixedPairsArray ast, AST_RANGES, StaticData, _
                RANGE_TABLE_DIGIT_START, RANGE_TABLE_DIGIT_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_NOT_DIGIT
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendPrefixedPairsArray ast, AST_RANGES, StaticData, _
                RANGE_TABLE_NOTDIGIT_START, RANGE_TABLE_NOTDIGIT_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_WHITE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendPrefixedPairsArray ast, AST_RANGES, StaticData, _
                RANGE_TABLE_WHITE_START, RANGE_TABLE_WHITE_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_NOT_WHITE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendPrefixedPairsArray ast, AST_RANGES, StaticData, _
                RANGE_TABLE_NOTWHITE_START, RANGE_TABLE_NOTWHITE_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ATOM_WORD_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendPrefixedPairsArray ast, AST_RANGES, StaticData, _
                RANGE_TABLE_WORDCHAR_START, RANGE_TABLE_WORDCHAR_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ATOM_NOT_WORD_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            AppendPrefixedPairsArray ast, AST_RANGES, StaticData, _
                RANGE_TABLE_NOTWORDCHAR_START, RANGE_TABLE_NOTWORDCHAR_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_EOF
            ' Todo: If Not expectEof Then Err.Raise REGEX_ERR_UNEXPECTED_END_OF_PATTERN
            
            ' Close disjunction
            If potentialConcat1 = -1 Then
                currentAstNode = ast.Length
                AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            If pendingDisjunction = -1 Then
                currentDisjunction = potentialConcat1
            Else
                ast.Buffer(pendingDisjunction + 2) = potentialConcat1
            End If
            
            currentAstNode = ast.Length
            AppendThree ast, AST_CAPTURE, currentDisjunction, 0
            ast.Buffer(0) = currentAstNode ' patch index of root node into the first word
            
            Exit Do
        Case Else
            Err.Raise REGEX_ERR_UNEXPECTED_REGEXP_TOKEN
        End Select
    Loop
End Sub

Private Function DfsMatch( _
    ByRef outCaptures As CapturesTy, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    Optional ByVal stepsLimit = DEFAULT_STEPS_LIMIT, _
    Optional ByVal multiLine As Boolean = False _
) As Long
    Dim context As DfsMatcherContext
    DfsMatch = DfsMatchFrom(context, outCaptures, bytecode, inputStr, 0, stepsLimit, multiLine:=multiLine)
End Function

Private Function DfsMatchFrom( _
    ByRef context As DfsMatcherContext, _
    ByRef outCaptures As CapturesTy, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    Optional ByVal stepsLimit As Long = DEFAULT_STEPS_LIMIT, _
    Optional ByVal multiLine As Boolean = False _
) As Long
    Dim nNamedCaptures As Long, nProperCapturePoints As Long, res As Long
    
    nProperCapturePoints = bytecode(0) + 1
    nNamedCaptures = bytecode(1)
    ' Todo: can we postpone this until we know that we will definitely need to fill outCaptures?
    With outCaptures
        .nNumberedCaptures = nProperCapturePoints \ 2 - 1
        .nNamedCaptures = bytecode(1)
        If .nNumberedCaptures > 0 Then ReDim .numberedCaptures(0 To .nNumberedCaptures - 1) As StartLengthPair
        If .nNamedCaptures > 0 Then ReDim .namedCaptures(0 To .nNamedCaptures - 1) As Long
    End With
    
    Do While sp <= Len(inputStr)
        InitializeMatcherContext context, nProperCapturePoints, nProperCapturePoints + nNamedCaptures
        res = DfsRunThreads(outCaptures, context, bytecode, inputStr, sp, stepsLimit, multiLine)
        If res <> -1 Then
            DfsMatchFrom = res
            Exit Function
        End If
        sp = sp + 1
    Loop

    DfsMatchFrom = -1
End Function

Private Function GetBc(ByRef bytecode() As Long, ByRef pc As Long) As Long
    If pc > UBound(bytecode) Then
        GetBc = 0 ' Todo: ??????
        Exit Function
    End If
    GetBc = bytecode(pc)
    pc = pc + 1
End Function

Private Function GetInputCharCode(ByRef inputStr As String, ByRef sp As Long, ByVal spDelta As Long) As Long
    If sp >= Len(inputStr) Then
        GetInputCharCode = DFS_ENDOFINPUT
    ElseIf sp < 0 Then
        GetInputCharCode = DFS_ENDOFINPUT
    Else
        GetInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) And &HFFFF& ' sp is 0-based and Mid$ is 1-based
        sp = sp + spDelta
    End If
End Function

Private Function PeekInputCharCode(ByRef inputStr As String, ByRef sp As Long) As Long
    If sp >= Len(inputStr) Then
        PeekInputCharCode = DFS_ENDOFINPUT
    ElseIf sp < 0 Then
        PeekInputCharCode = DFS_ENDOFINPUT
    Else
        PeekInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) And &HFFFF& ' sp is 0-based and Mid$ is 1-based
    End If
End Function

Private Function UnicodeReIsWordchar(c As Long)
    'TODO: Temporary hack
    UnicodeReIsWordchar = ((c >= AscW("A")) And (c <= AscW("Z"))) Or ((c >= AscW("a") And (c <= AscW("z"))))
End Function

Private Sub InitializeMatcherContext(ByRef context As DfsMatcherContext, ByVal nProperCapturePoints As Long, ByVal nCapturePoints As Long)
    With context
        ' Clear stacks
        .matcherStack.Length = 0
        .capturesStack.Length = 0
        .qstack.Length = 0
        
        .nProperCapturePoints = nProperCapturePoints
        .nCapturePoints = nCapturePoints
        AppendFill .capturesStack, nCapturePoints, -1
        
        .master = -1
        .capturesRequireCoW = False
        .qTop = 0
    End With
End Sub

Private Sub PushMatcherStackFrame( _
    ByRef context As DfsMatcherContext, _
    ByVal pc As Long, _
    ByVal sp As Long, _
    ByVal pcLandmark As Long, _
    ByVal spDelta As Long, _
    ByVal q As Long _
)
    With context.matcherStack
        If .Length = .Capacity Then
            ' Increase capacity
            If .Capacity < DFS_MATCHER_STACK_MINIMUM_CAPACITY Then .Capacity = DFS_MATCHER_STACK_MINIMUM_CAPACITY Else .Capacity = .Capacity + .Capacity \ 2
            ReDim Preserve .Buffer(0 To .Capacity - 1) As DfsMatcherStackFrame
        End If
        With .Buffer(.Length)
           .master = context.master
           .capturesStackState = context.capturesStack.Length Or (context.capturesRequireCoW And LONG_FIRST_BIT)
           .qStackLength = context.qstack.Length
           .pc = pc
           .sp = sp
           .pcLandmark = pcLandmark
           .spDelta = spDelta
           .q = q
           .qTop = context.qTop
        End With
        .Length = .Length + 1
    End With

    context.capturesRequireCoW = True
End Sub

Private Function PopMatcherStackFrame(ByRef context As DfsMatcherContext, ByRef pc As Long, ByRef sp As Long, ByRef pcLandmark As Long, ByRef spDelta As Long, ByRef q As Long) As Boolean
    With context.matcherStack
        If .Length = 0 Then
            PopMatcherStackFrame = False
            Exit Function
        End If
    
        .Length = .Length - 1
        With .Buffer(.Length)
            context.master = .master
            context.capturesStack.Length = .capturesStackState And LONG_ALL_BUT_FIRST_BIT
            context.qstack.Length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And LONG_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    
        PopMatcherStackFrame = True
    End With
End Function

Private Sub ReturnToMasterDiscardCaptures(ByRef context As DfsMatcherContext, ByRef pc As Long, ByRef sp As Long, ByRef pcLandmark As Long, ByRef spDelta As Long, ByRef q As Long)
    With context.matcherStack
        .Length = context.master
        With .Buffer(.Length)
            context.master = .master
            context.capturesStack.Length = .capturesStackState And LONG_ALL_BUT_FIRST_BIT
            context.qstack.Length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And LONG_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    End With
End Sub

Private Sub ReturnToMasterPreserveCaptures(ByRef context As DfsMatcherContext, ByRef pc As Long, ByRef sp As Long, ByRef pcLandmark As Long, ByRef spDelta As Long, ByRef q As Long)
    Dim masterCapturesStackLength As Long, i As Long
    
    With context.matcherStack
        .Length = context.master
        With .Buffer(.Length)
            context.master = .master
            masterCapturesStackLength = .capturesStackState And LONG_ALL_BUT_FIRST_BIT
            context.qstack.Length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And LONG_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    End With
    
    With context.capturesStack
        If .Length = masterCapturesStackLength Then Exit Sub

        If context.capturesRequireCoW Then
            masterCapturesStackLength = masterCapturesStackLength + context.nCapturePoints
            context.capturesRequireCoW = False
            If .Length = masterCapturesStackLength Then Exit Sub
        End If
        
        For i = 1 To context.nCapturePoints
            .Buffer(masterCapturesStackLength - i) = .Buffer(.Length - i)
        Next
        .Length = masterCapturesStackLength
    End With
End Sub

Private Sub CopyCaptures(ByRef context As DfsMatcherContext, ByRef captures As CapturesTy)
    Dim i As Long, baseIdx As Long, pt1 As Long, pt2 As Long
    
    With context
        baseIdx = .capturesStack.Length - .nCapturePoints
        pt1 = .capturesStack.Buffer(baseIdx)
        pt2 = .capturesStack.Buffer(baseIdx + 1)
        If pt1 = -1 Then
            With captures.entireMatch: .start = 0: .Length = 0: End With
        ElseIf pt2 < pt1 Then
            With captures.entireMatch: .start = 0: .Length = 0: End With
        Else
            With captures.entireMatch: .start = pt1 + 1: .Length = pt2 - pt1: End With
        End If
            
        For i = 1 To captures.nNumberedCaptures
            pt1 = .capturesStack.Buffer(baseIdx + 2 * i)
            pt2 = .capturesStack.Buffer(baseIdx + 2 * i + 1)
            If pt1 = -1 Then
                With captures.numberedCaptures(i - 1): .start = 0: .Length = 0: End With
            ElseIf pt2 < pt1 Then
                With captures.numberedCaptures(i - 1): .start = 0: .Length = 0: End With
            Else
                With captures.numberedCaptures(i - 1): .start = pt1 + 1: .Length = pt2 - pt1: End With
            End If
        Next
        
        baseIdx = baseIdx + 2 + 2 * captures.nNumberedCaptures
        For i = 0 To captures.nNamedCaptures - 1
            captures.namedCaptures(i) = .capturesStack.Buffer(baseIdx + i)
        Next
    End With
End Sub

Private Sub SetCapturePoint(ByRef context As DfsMatcherContext, ByVal idx As Long, ByVal v As Long)
    With context
        If .capturesRequireCoW Then
            AppendSlice .capturesStack, .capturesStack.Length - .nCapturePoints, .nCapturePoints
            .capturesRequireCoW = False
        End If
        .capturesStack.Buffer(.capturesStack.Length - .nCapturePoints + idx) = v
    End With
End Sub

Private Function GetCapturePoint(ByRef context As DfsMatcherContext, ByVal idx As Long) As Long
    With context
        GetCapturePoint = .capturesStack.Buffer( _
            .capturesStack.Length - .nCapturePoints + idx _
        )
    End With
End Function

Private Sub PushQCounter(ByRef context As DfsMatcherContext, ByVal q As Long)
    With context
        AppendThree .qstack, .qTop, .matcherStack.Length, q
        .qTop = .qstack.Length - 1
    End With
End Sub

Private Function PopQCounter(ByRef context As DfsMatcherContext) As Long
    With context
        If .qTop = 0 Then
            PopQCounter = Q_NONE
            Exit Function
        End If
        
        PopQCounter = .qstack.Buffer(.qTop)
        If .qstack.Buffer(.qTop - 1) = .matcherStack.Length Then .qstack.Length = .qstack.Length - 3
        .qTop = .qstack.Buffer(.qTop - 2)
    End With
End Function

Private Function DfsRunThreads( _
    ByRef outCaptures As CapturesTy, _
    ByRef context As DfsMatcherContext, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    ByVal stepsLimit As Long, _
    ByVal multiLine As Boolean _
) As Long
    Dim caseInsensitive As Boolean
    Dim pc As Long
    
    ' To avoid infinite loops
    Dim pcLandmark As Long
    
    Dim op As Long
    Dim c1 As Long, c2 As Long
    Dim t As Long
    Dim n As Long
    Dim successfulMatch As Boolean
    Dim r1 As Long, r2 As Long
    Dim b1 As Boolean, b2 As Boolean
    Dim aa As Long, bb As Long, mm As Long
    Dim idx As Long, off As Long
    Dim q As Long, qmin As Long, qmax As Long, qq As Long
    Dim qexact As Long
    Dim stepsCount As Long
    Dim spDelta As Long ' 1 when we walk forwards and -1 when we walk backwards
    
        caseInsensitive = bytecode(BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR) <> 0
        pc = 3 + 3 * bytecode(BYTECODE_IDX_N_IDENTIFIERS)
        pcLandmark = -1
        stepsCount = 0
        spDelta = 1
        q = Q_NONE
       
        GoTo ContinueLoopSuccess

        ' BEGIN LOOP
ContinueLoopFail:
            If Not PopMatcherStackFrame(context, pc, sp, pcLandmark, spDelta, q) Then
                DfsRunThreads = -1
                Exit Function
            End If

ContinueLoopSuccess:

            ' TODO: HACK to prevent infinite loop!!!!
            stepsCount = stepsCount + 1
            If stepsCount >= stepsLimit Then
                DfsRunThreads = -1
                Exit Function
            End If
            
            op = GetBc(bytecode, pc)
    
            ' #if defined(DUK_USE_DEBUG_LEVEL) && (DUK_USE_DEBUG_LEVEL >= 2)
            ' duk__regexp_dump_state(re_ctx);
            ' #End If
            ' DUK_DDD(DUK_DDDPRINT("match: rec=%ld, steps=%ld, pc (after op)=%ld, sp=%ld, op=%ld",
            '                     (long) re_ctx->recursion_depth,
            '                     (long) re_ctx->steps_count,
            '                     (long) (pc - re_ctx->bytecode),
            '                     (long) sp,
            '                     (long) op));
    
            Select Case op
            Case REOP_MATCH
                GoTo Match
            Case REOP_END_LOOKPOS
                ' Reached if pattern inside a positive lookahead matched.
                ReturnToMasterPreserveCaptures context, pc, sp, pcLandmark, spDelta, q
                ' Now we are at the REOP_LOOKPOS opcode.
                pc = pc + 1
                n = GetBc(bytecode, pc)
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_END_LOOKNEG
                ' Reached if pattern inside a positive lookahead matched.
                ReturnToMasterDiscardCaptures context, pc, sp, pcLandmark, spDelta, q
                GoTo ContinueLoopFail
            Case REOP_CHAR
                '
                '  Byte-based matching would be possible for case-sensitive
                '  matching but not for case-insensitive matching.  So, we
                '  match by decoding the input and bytecode character normally.
                '
                '  Bytecode characters are assumed to be already canonicalized.
                '  Input characters are canonicalized automatically by
                '  duk__inp_get_cp() if necessary.
                '
                '  There is no opcode for matching multiple characters.  The
                '  regexp compiler has trouble joining strings efficiently
                '  during compilation.  See doc/regexp.rst for more discussion.
    
                pcLandmark = pc - 1
                c1 = GetBc(bytecode, pc)
                c2 = GetInputCharCode(inputStr, sp, spDelta)
                If caseInsensitive Then c2 = ReCanonicalizeChar(c2)
                
                ' DUK_ASSERT(c1 >= 0);
    
                ' DUK_DDD(DUK_DDDPRINT("char match, c1=%ld, c2=%ld", (long) c1, (long) c2));
                If c1 <> c2 Then GoTo ContinueLoopFail
                GoTo ContinueLoopSuccess
            Case REOP_PERIOD
                pcLandmark = pc - 1
                c1 = GetInputCharCode(inputStr, sp, spDelta)
                If c1 < 0 Then
                    GoTo ContinueLoopFail
                ElseIf UnicodeIsLineTerminator(c1) Then
                    GoTo ContinueLoopFail
                End If
                GoTo ContinueLoopSuccess
            Case REOP_RANGES, REOP_INVRANGES
                pcLandmark = pc - 1
                n = GetBc(bytecode, pc) ' assert: >= 1
                c1 = GetInputCharCode(inputStr, sp, spDelta)
                If c1 < 0 Then GoTo ContinueLoopFail
                If caseInsensitive Then c1 = ReCanonicalizeChar(c1)
                
                aa = pc - 1
                pc = pc + 2 * n
                bb = pc + 1
                
                ' We are doing a binary search here.
                Do
                    mm = aa + 2 * ((bb - aa) \ 4)
                    If bytecode(mm) >= c1 Then bb = mm Else aa = mm
                    
                    If bb - aa = 2 Then
                        ' bb is the first upper bound index s.t. ary(bb)>=v
                        If bb >= pc Then successfulMatch = False Else successfulMatch = bytecode(bb - 1) <= c1
                        Exit Do
                    End If
                Loop
    
                If (op = REOP_RANGES) <> successfulMatch Then GoTo ContinueLoopFail

                GoTo ContinueLoopSuccess
            Case REOP_ASSERT_START
                If sp <= 0 Then GoTo ContinueLoopSuccess
                If Not multiLine Then GoTo ContinueLoopFail
                c1 = PeekInputCharCode(inputStr, sp - (spDelta + 1) \ 2)
                ' E5 Sections 15.10.2.8, 7.3
                If UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopSuccess
                GoTo ContinueLoopFail
            Case REOP_ASSERT_END
                c1 = PeekInputCharCode(inputStr, sp - (spDelta - 1) \ 2)
                If c1 = DFS_ENDOFINPUT Then GoTo ContinueLoopSuccess
                If Not multiLine Then GoTo ContinueLoopFail
                If UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopSuccess
                GoTo ContinueLoopFail
            Case REOP_ASSERT_WORD_BOUNDARY, REOP_ASSERT_NOT_WORD_BOUNDARY
                '
                '  E5 Section 15.10.2.6.  The previous and current character
                '  should -not- be canonicalized as they are now.  However,
                '  canonicalization does not affect the result of IsWordChar()
                '  (which depends on Unicode characters never canonicalizing
                '  into ASCII characters) so this does not matter.
                If sp <= 0 Then
                    b1 = False  ' not a wordchar
                Else
                    c1 = PeekInputCharCode(inputStr, sp - spDelta)
                    b1 = UnicodeReIsWordchar(c1)
                End If
                If sp > Len(inputStr) Then
                    b2 = False ' not a wordchar
                Else
                    c1 = PeekInputCharCode(inputStr, sp)
                    b2 = UnicodeReIsWordchar(c1)
                End If
    
                If (op = REOP_ASSERT_WORD_BOUNDARY) = (b1 = b2) Then GoTo ContinueLoopFail

                GoTo ContinueLoopSuccess
            Case REOP_JUMP
                n = GetBc(bytecode, pc)
                If n > 0 Then
                    ' forward jump (disjunction)
                    pc = pc + n
                    GoTo ContinueLoopSuccess
                Else
                    ' backward jump (end of loop)
                    t = pc + n
                    If pcLandmark <= t Then GoTo ContinueLoopSuccess ' empty match
                                    
                    pc = t: pcLandmark = t
                    GoTo ContinueLoopSuccess
                End If
            Case REOP_SPLIT1
                ' split1: prefer direct execution (no jump)
                n = GetBc(bytecode, pc)
                PushMatcherStackFrame context, pc + n, sp, pcLandmark, spDelta, q
                '.tsStack(.tsLastIndex - 1).pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_SPLIT2
                ' split2: prefer jump execution (not direct)
                n = GetBc(bytecode, pc)
                PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_EXACTLY_INIT
                PushQCounter context, q
                q = 0
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_EXACTLY_START
                pc = pc + 2 ' skip arguments
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_EXACTLY_END
                qexact = GetBc(bytecode, pc) ' quantity
                n = GetBc(bytecode, pc) ' offset
                q = q + 1
                If q < qexact Then
                    t = pc - n - 3
                    If pcLandmark > t Then pcLandmark = t
                    pc = t
                Else
                    q = PopQCounter(context)
                End If
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_MAX_HUMBLE_INIT, REOP_REPEAT_GREEDY_MAX_INIT
                PushQCounter context, q
                q = -1
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_MAX_HUMBLE_START
                qmax = GetBc(bytecode, pc)
                n = GetBc(bytecode, pc)
                
                q = q + 1
                If q < qmax Then PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                
                q = PopQCounter(context)
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_MAX_HUMBLE_END
                pc = pc + 1  ' skip first argument: quantity
                n = GetBc(bytecode, pc) ' offset
                t = pc - n - 3
                If pcLandmark <= t Then GoTo ContinueLoopFail ' empty match
                
                pc = t: pcLandmark = t
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_GREEDY_MAX_START
                qmax = GetBc(bytecode, pc)
                n = GetBc(bytecode, pc)
                
                q = q + 1
                If q < qmax Then
                    qq = PopQCounter(context)
                    PushMatcherStackFrame context, pc + n, sp, pcLandmark, spDelta, qq
                    PushQCounter context, qq
                Else
                    pc = pc + n
                    q = PopQCounter(context)
                End If
                
                GoTo ContinueLoopSuccess
            Case REOP_REPEAT_GREEDY_MAX_END
                pc = pc + 1 ' Skip first argument: quantity
                n = GetBc(bytecode, pc) ' offset
                t = pc - n - 3
                If pcLandmark <= t Then GoTo ContinueLoopSuccess ' empty match
                
                pc = t: pcLandmark = t
                GoTo ContinueLoopSuccess
            Case REOP_SAVE
                idx = GetBc(bytecode, pc)
                If idx >= context.nCapturePoints Then GoTo InternalError
                    ' idx is unsigned, < 0 check is not necessary
                    ' DUK_D(DUK_DPRINT("internal error, regexp save index insane: idx=%ld", (long) idx));
                '.tsStack(.tsLastIndex).saved(idx) = sp
                SetCapturePoint context, idx, sp
                GoTo ContinueLoopSuccess
            Case REOP_SET_NAMED
                r1 = GetBc(bytecode, pc)
                If r1 >= context.nCapturePoints Then GoTo InternalError
                r2 = GetBc(bytecode, pc)
                SetCapturePoint context, context.nProperCapturePoints + r1, r2
                GoTo ContinueLoopSuccess
            Case REOP_CHECK_LOOKAHEAD
                PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                context.master = context.matcherStack.Length - 1
                pc = pc + 2 ' jump over following REOP_LOOKPOS or REOP_LOOKNEG
                ' When we're moving forward, we are at the correct position. When we're moving backward, we have to step one towards the end.
                sp = sp + (1 - spDelta) \ 2
                spDelta = 1
                ' We could set pcLandmark to -1 again here, but we can be sure that pcLandmark < beginning of lookahead, so we can skip that
                GoTo ContinueLoopSuccess
            Case REOP_CHECK_LOOKBEHIND
                PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
                context.master = context.matcherStack.Length - 1
                pc = pc + 2 ' jump over following REOP_LOOKPOS or REOP_LOOKNEG
                ' When we're moving backward, we are at the correct position. When we're moving forward, we have to step one towards the beginning.
                sp = sp - (spDelta + 1) \ 2
                spDelta = -1
                ' We could set pcLandmark to -1 again here, but we can be sure that pcLandmark < beginning of lookahead, so we can skip that
                GoTo ContinueLoopSuccess
            Case REOP_LOOKPOS
                ' This point will only be reached if the pattern inside a negative lookahead/back did not match.
                n = GetBc(bytecode, pc)
                pc = pc + n
                GoTo ContinueLoopFail
            Case REOP_LOOKNEG
                ' This point will only be reached if the pattern inside a negative lookahead/back did not match.
                n = GetBc(bytecode, pc)
                pc = pc + n
                GoTo ContinueLoopSuccess
            Case REOP_BACKREFERENCE
                '
                '  Byte matching for back-references would be OK in case-
                '  sensitive matching.  In case-insensitive matching we need
                '  to canonicalize characters, so back-reference matching needs
                '  to be done with codepoints instead.  So, we just decode
                '  everything normally here, too.
                '
                '  Note: back-reference index which is 0 or higher than
                '  NCapturingParens (= number of capturing parens in the
                '  -entire- regexp) is a compile time error.  However, a
                '  backreference referring to a valid capture which has
                '  not matched anything always succeeds!  See E5 Section
                '  15.10.2.9, step 5, sub-step 3.
    
                pcLandmark = pc - 1
                idx = 2 * GetBc(bytecode, pc) ' backref n -> saved indices [n*2, n*2+1]
                If idx < 2 Then GoTo InternalError
                If idx + 1 >= context.nCapturePoints Then GoTo InternalError
                aa = GetCapturePoint(context, idx)
                bb = GetCapturePoint(context, idx + 1)
                If (aa >= 0) And (bb >= 0) Then
                    If spDelta = 1 Then
                        off = aa
                        Do While off < bb
                            c1 = GetInputCharCode(inputStr, off, 1)
                            c2 = GetInputCharCode(inputStr, sp, 1)
                            ' No need for an explicit c2 < 0 check: because c1 >= 0,
                            ' the comparison will always fail if c2 < 0.
                            If c1 <> c2 Then
                                If Not caseInsensitive Then GoTo ContinueLoopFail
                                If ReCanonicalizeChar(c1) <> ReCanonicalizeChar(c2) Then GoTo ContinueLoopFail
                            End If
                        Loop
                    Else
                        off = bb - 1
                        Do While off >= aa
                            c1 = GetInputCharCode(inputStr, off, -1)
                            c2 = GetInputCharCode(inputStr, sp, -1)
                            ' No need for an explicit c2 < 0 check: because c1 >= 0,
                            ' the comparison will always fail if c2 < 0.
                            If c1 <> c2 Then
                                If Not caseInsensitive Then GoTo ContinueLoopFail
                                If ReCanonicalizeChar(c1) <> ReCanonicalizeChar(c2) Then GoTo ContinueLoopFail
                            End If
                        Loop
                    End If
                Else
                    ' capture is 'undefined', always matches!
                End If
                GoTo ContinueLoopSuccess
            Case Else
                'DUK_D(DUK_DPRINT("internal error, regexp opcode error: %ld", (long) op));
                GoTo InternalError
            End Select
        ' END LOOP
    
Match:
        CopyCaptures context, outCaptures
        DfsRunThreads = sp
        Exit Function
        
InternalError:
        ' TODO: Raise correct exception
        Err.Raise 3000
        ' DUK_ERROR_INTERNAL(re_ctx->thr);
        ' DUK_WO_NORETURN(return -1;);
End Function

Private Sub ParseFormatString(ByRef parsedFormat As ArrayBuffer, ByRef formatString As String, ByRef bytecode() As Long, ByRef pattern As String)
    Dim curPos As Long, lastPos As Long, c As Long, formatStringLen As Long, num As Long, substrLen As Long, identifierId As Long
    
    Const UNICODE_DOLLAR As Long = 36
    Const UNICODE_AMP As Long = 38
    Const UNICODE_SQUOTE As Long = 39
    Const UNICODE_DIGIT_0 As Long = 48
    Const UNICODE_DIGIT_9 As Long = 57
    Const UNICODE_LT As Long = 60
    Const UNICODE_BACKTICK As Long = 96
    Const UNICODE_TILDE As Long = 126
    
    
    formatStringLen = Len(formatString)
    curPos = 1
    lastPos = 1
    Do
        curPos = InStr(curPos, formatString, "$", vbBinaryCompare)
        If curPos = 0 Then Exit Do
        If curPos = formatStringLen Then GoTo InvalidReplacementString
        curPos = curPos + 1
        c = AscW(Mid$(formatString, curPos, 1))
        If c = UNICODE_DOLLAR Then
            If curPos - lastPos = 1 Then
                AppendLong parsedFormat, REPL_DOLLAR
            Else
                AppendThree parsedFormat, REPL_SUBSTR, lastPos, curPos - lastPos
            End If
            curPos = curPos + 1
        Else
            substrLen = curPos - lastPos - 1
            If substrLen > 0 Then AppendThree parsedFormat, REPL_SUBSTR, lastPos, substrLen
            Select Case c
            Case UNICODE_AMP
                AppendLong parsedFormat, REPL_ACTUAL
                curPos = curPos + 1
            Case UNICODE_SQUOTE
                AppendLong parsedFormat, REPL_SUFFIX
                curPos = curPos + 1
            Case UNICODE_LT
                If curPos = formatStringLen Then GoTo InvalidReplacementString
                lastPos = curPos + 1
                curPos = InStr(lastPos, formatString, ">", vbBinaryCompare)
                If curPos = lastPos Then GoTo InvalidReplacementString ' empty identifier
                
                identifierId = GetIdentifierId(bytecode, pattern, Mid$(formatString, lastPos, curPos - lastPos))
                If identifierId >= 0 Then AppendTwo parsedFormat, REPL_NAMED, identifierId

                curPos = curPos + 1
            Case UNICODE_BACKTICK
                AppendLong parsedFormat, REPL_PREFIX
                curPos = curPos + 1
            Case UNICODE_TILDE
                ' ignore
                curPos = curPos + 1
            Case Else
                ' Todo: Check whether we can merge this with parsing a number within a regex
                If c < UNICODE_DIGIT_0 Then GoTo InvalidReplacementString
                If c > UNICODE_DIGIT_9 Then GoTo InvalidReplacementString
                num = 0
                Do
                    If num > LONG_MAX_DIV_10 Then GoTo InvalidReplacementString
                    
                    num = 10 * num
                    c = c - UNICODE_DIGIT_0
                    
                    If num > LONG_MAX - c Then GoTo InvalidReplacementString
                    
                    num = num + c
                    
                    curPos = curPos + 1
                    If curPos > formatStringLen Then Exit Do
                    c = AscW(Mid$(formatString, curPos, 1))
                    If c < UNICODE_DIGIT_0 Then Exit Do
                    If c > UNICODE_DIGIT_9 Then Exit Do
                Loop
                AppendTwo parsedFormat, REPL_NUMBERED, num
            End Select
        End If
        lastPos = curPos
    Loop
    
    substrLen = formatStringLen + 1 - lastPos
    If substrLen > 0 Then AppendThree parsedFormat, REPL_SUBSTR, lastPos, substrLen
    AppendLong parsedFormat, REPL_END
    
    Exit Sub
InvalidReplacementString:
    Err.Raise REGEX_ERR_INVALID_REPLACEMENT_STRING
End Sub

Private Sub AppendFormatted( _
    ByRef sb As StaticStringBuilder, _
    ByRef sHaystack As String, _
    ByRef captures As CapturesTy, _
    ByRef formatString As String, _
    ByRef parsed() As Long, _
    Optional ByVal parsedStartPos As Long = 0 _
)
    Dim j As Long, num As Long

    j = parsedStartPos
    Do
        Select Case parsed(j)
        Case REPL_END
            Exit Do
        Case REPL_DOLLAR
            AppendStr sb, "$"
            j = j + 1
        Case REPL_SUBSTR
            AppendStr sb, Mid$(formatString, parsed(j + 1), parsed(j + 2))
            j = j + 3
        Case REPL_PREFIX
            AppendStr sb, Left$(sHaystack, captures.entireMatch.start - 1)
            j = j + 1
        Case REPL_SUFFIX
            AppendStr sb, Mid$(sHaystack, captures.entireMatch.start + captures.entireMatch.Length)
            j = j + 1
        Case REPL_ACTUAL
            AppendStr sb, Mid$(sHaystack, captures.entireMatch.start, captures.entireMatch.Length)
            j = j + 1
        Case REPL_NUMBERED
            num = parsed(j + 1)
            If num <= captures.nNumberedCaptures Then
                With captures.numberedCaptures(num - 1)
                    If .Length > 0 Then AppendStr sb, Mid$(sHaystack, .start, .Length)
                End With
            End If
            j = j + 2
        Case REPL_NAMED
            num = captures.namedCaptures(parsed(j + 1))
            If num >= 0 Then
                If num <= captures.nNumberedCaptures Then
                    With captures.numberedCaptures(num - 1)
                        If .Length > 0 Then AppendStr sb, Mid$(sHaystack, .start, .Length)
                    End With
                End If
            End If
            j = j + 2
        Case Else
            Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR
        End Select
    Loop

End Sub

Public Function TryInitializeRegex( _
    ByRef regex As RegexTy, _
    ByRef pattern As String, _
    Optional ByVal caseInsensitive As Boolean = False _
)
    ' Todo:
    '   Actually, this is not what we want to have.
    '   We should change the compiler so that it reports syntax errors in the regex via a channel
    '     different to throwing.
    '   Then InitializeRegex should make use of TryInitializeRegex, not the other way round.
    On Error GoTo Fail
    InitializeRegex regex, pattern, caseInsensitive
    TryInitializeRegex = True
    Exit Function

Fail:
    TryInitializeRegex = False
End Function

Public Sub InitializeRegex( _
    ByRef regex As RegexTy, _
    ByRef pattern As String, _
    Optional ByVal caseInsensitive As Boolean = False _
)
    regex.pattern = pattern
    regex.isCaseInsensitive = caseInsensitive
    regex.stepsLimit = DEFAULT_STEPS_LIMIT
    Compile regex.bytecode, pattern, caseInsensitive:=caseInsensitive
End Sub

Public Function Test(ByRef regex As RegexTy, ByRef str As String, Optional ByVal multiLine As Boolean = False) As Boolean
    Dim captures As CapturesTy
    
    Test = DfsMatch( _
        captures, regex.bytecode, str, stepsLimit:=regex.stepsLimit, _
        multiLine:=multiLine _
    ) <> -1
End Function

Public Function Match( _
    ByRef matcherState As MatcherStateTy, ByRef regex As RegexTy, ByRef haystack As String, _
    Optional ByVal multiLine As Boolean = False, _
    Optional ByVal matchFrom As Long = 1 _
) As Boolean
    Match = DfsMatchFrom( _
        matcherState.context, matcherState.captures, regex.bytecode, haystack, matchFrom - 1, _
        stepsLimit:=regex.stepsLimit, _
        multiLine:=multiLine _
    ) <> -1
    matcherState.current = -1
End Function

Public Function GetCapture(ByRef matcherState As MatcherStateTy, ByRef haystack As String, Optional ByVal num As Long = 0) As String
    If num = 0 Then
        With matcherState.captures.entireMatch
            If .Length > 0 Then GetCapture = Mid$(haystack, .start, .Length) Else GetCapture = vbNullString
        End With
    ElseIf num <= matcherState.captures.nNumberedCaptures Then
        With matcherState.captures.numberedCaptures(num - 1)
            If .Length > 0 Then GetCapture = Mid$(haystack, .start, .Length) Else GetCapture = vbNullString
        End With
    Else
        GetCapture = vbNullString
    End If
End Function

Public Function GetCaptureByName( _
    ByRef matcherState As MatcherStateTy, _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    ByRef name As String _
) As String
    Dim identifierId As Long
    
    identifierId = GetIdentifierId(regex.bytecode, regex.pattern, name)
    If identifierId < 0 Then GetCaptureByName = vbNullString: Exit Function
    GetCaptureByName = GetCapture(matcherState, haystack, matcherState.captures.namedCaptures(identifierId))
End Function

Public Function MatchNext(ByRef matcherState As MatcherStateTy, ByRef regex As RegexTy, ByRef haystack As String) As Boolean
    Dim r As Long
    
    If matcherState.current = -1 Then Exit Function ' end of string reached, return False
    
    r = DfsMatchFrom( _
        matcherState.context, matcherState.captures, regex.bytecode, haystack, matcherState.current, _
        stepsLimit:=regex.stepsLimit, _
        multiLine:=matcherState.multiLine _
    )
    
    matcherState.current = r Or matcherState.localMatch
    MatchNext = r <> -1
End Function

Public Function Replace( _
    ByRef regex As RegexTy, _
    ByRef replacer As String, _
    ByRef haystack As String, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiLine As Boolean = False _
) As String
    Dim parsedFormat As ArrayBuffer, matcherState As MatcherStateTy, lastEndPos As Long, resultBuilder As StaticStringBuilder

    ParseFormatString parsedFormat, replacer, regex.bytecode, regex.pattern

    lastEndPos = 1
    matcherState.localMatch = localMatch
    matcherState.multiLine = multiLine
    Do While MatchNext(matcherState, regex, haystack)
        AppendStr resultBuilder, Mid$(haystack, lastEndPos, matcherState.captures.entireMatch.start - lastEndPos)
        AppendFormatted resultBuilder, haystack, matcherState.captures, replacer, parsedFormat.Buffer
        
        lastEndPos = matcherState.captures.entireMatch.start + matcherState.captures.entireMatch.Length
    Loop
    
    AppendStr resultBuilder, Mid$(haystack, lastEndPos)
    
    Replace = GetStr(resultBuilder)
End Function

Public Function MatchThenJoin( _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    Optional ByRef format As String = "$&", _
    Optional ByRef delimiter As String = vbNullString, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiLine As Boolean = False _
) As String
    Dim parsedFormat As ArrayBuffer, resultBuilder As StaticStringBuilder, matcherState As MatcherStateTy
    
    ParseFormatString parsedFormat, format, regex.bytecode, regex.pattern
    
    matcherState.localMatch = localMatch
    matcherState.multiLine = multiLine
    If MatchNext(matcherState, regex, haystack) Then
        AppendFormatted resultBuilder, haystack, matcherState.captures, format, parsedFormat.Buffer
        Do While MatchNext(matcherState, regex, haystack)
            AppendStr resultBuilder, delimiter
            AppendFormatted resultBuilder, haystack, matcherState.captures, format, parsedFormat.Buffer
        Loop
    End If
    
    MatchThenJoin = GetStr(resultBuilder)
End Function

Public Sub MatchThenList( _
    ByRef results() As String, _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    ByRef formatStrings() As String, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiLine As Boolean = False _
)
    Dim cola As Long, colb As Long, j As Long, k As Long, m As Long, mm As Long, nMatches As Long
    Dim parsedFormats As ArrayBuffer
    Dim matcherState As MatcherStateTy
    Dim resultBuilder As StaticStringBuilder

    cola = LBound(formatStrings)
    colb = UBound(formatStrings)
    
    For j = cola To colb
        k = parsedFormats.Length
        AppendLong parsedFormats, 0
        ParseFormatString parsedFormats, formatStrings(j), regex.bytecode, regex.pattern
        parsedFormats.Buffer(k) = parsedFormats.Length - k
    Next
    
    nMatches = 0
    
    matcherState.localMatch = localMatch
    matcherState.multiLine = multiLine
    Do While MatchNext(matcherState, regex, haystack)
        k = 0
        For j = cola To colb
            m = resultBuilder.Length
            AppendFormatted resultBuilder, haystack, matcherState.captures, formatStrings(j), parsedFormats.Buffer, k + 1
            AppendStr resultBuilder, ChrW$(m)
            k = k + parsedFormats.Buffer(k)
        Next
        
        nMatches = nMatches + 1
    Loop
    
    If nMatches = 0 Then
        ' hack to create a zero-length array
        results = Split(vbNullString)
    Else
        ReDim results(0 To nMatches - 1, cola To colb) As String
        m = resultBuilder.Length
        For j = nMatches - 1 To 0 Step -1
            For k = colb To cola Step -1
                mm = AscW(GetSubstr(resultBuilder, m, 1))
                results(j, k) = GetSubstr(resultBuilder, mm + 1, m - mm - 1)
                m = mm
            Next
        Next
    End If
End Sub

Public Sub InitializeMatcherState( _
    ByRef matcherState As MatcherStateTy, Optional ByVal localMatch = False, Optional ByVal multiLine = False _
)
    matcherState.current = 0
    matcherState.localMatch = localMatch
    matcherState.multiLine = multiLine
End Sub

Public Sub ResetMatcherState(ByRef matcherState As MatcherStateTy)
    matcherState.current = 0
End Sub

