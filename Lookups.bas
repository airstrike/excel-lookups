Attribute VB_Name = "Lookups"
Const MAX_RESULTS_SIZE As Long = 100

Public Function BaseLookup(ByVal Field As String, ParamArray Lookups() As Variant) As Variant

    BaseLookup = FlexLookup_( _
        Caller:=Application.Caller, ShtName:="BASE", _
        Field:=Field, Grouped:=True, Sorted:=True, _
        RowLookup:=False, _
        ProtoLookups:=Lookups)

End Function

Public Function RowLookup(ByVal ShtName As String, _
                            ByVal Grouped As Boolean, ByVal Sorted As Boolean, _
                            ParamArray Lookups() As Variant) As Variant
    RowLookup = FlexLookup_( _
        Caller:=Application.Caller, ShtName:=ShtName, _
        Field:="", Grouped:=Grouped, Sorted:=Sorted, _
        RowLookup:=True, _
        ProtoLookups:=Lookups)

End Function

Public Function MultiLookup(ByVal ShtName As String, ByVal Field As String, _
                            ByVal Grouped As Boolean, ByVal Sorted As Boolean, _
                            ParamArray Lookups() As Variant) As Variant

    MultiLookup = FlexLookup_( _
        Caller:=Application.Caller, ShtName:=ShtName, _
        Field:=Field, Grouped:=Grouped, Sorted:=Sorted, _
        ProtoLookups:=Lookups)

End Function

Public Function TableLookup(ByVal ShtName As String, ByVal Fields As Variant, _
                            ParamArray Lookups() As Variant)

    TableLookup = FlexLookup_( _
        Caller:=Application.Caller, ShtName:=ShtName, Sorted:=False, _
        Fields:=Fields, ProtoLookups:=Lookups)

End Function

Private Function FlexLookup_( _
    ByRef Caller As Variant, _
    ByVal ShtName As String, _
    ByVal ProtoLookups As Variant, _
    Optional ByVal Field As String = "", Optional ByVal Fields As Variant, _
    Optional ByVal Grouped As Boolean = True, Optional ByVal Sorted As Boolean = True, _
    Optional ByVal RowLookup As Boolean = False) As Variant
    ' ------------------------------------------------------------------------------------
    ' As a general rule of thumb, this function should not be accessed directly,
    ' but rather from one of its wrapper functions
    ' ------------------------------------------------------------------------------------
    
    Dim L As Long, x As Long, ProtoSize As Long, _
        FieldPos As Long, NonBlankFilters As Long, ResultSize As Long
    Dim Results() As Variant, Matches() As Variant, MatchesPos As Variant, _
        LookupFields() As Variant, LookupValues() As Variant, _
        LookupPos() As Variant
    
    Dim IsTableLookup As Boolean
    Dim RowIndexLookup As Boolean
    
    Dim Append As Boolean
    Dim LastValue As Variant
    Dim FieldCount As Long
    ReDim Results(0 To 0) As Variant
    ReDim Matches(0 To 0) As Variant
    
    Dim PreviousResults As New Scripting.Dictionary
    Dim ThisResult As Variant
    
    ReDim LookupFields(0 To 0) As Variant
    ReDim LookupValues(0 To 0) As Variant
    ReDim LookupPos(0 To 0) As Variant
    
    RowIndexLookup = False
    On Error Resume Next
    If IsNumeric(Field) Then RowIndexLookup = True
    On Error GoTo 0
    
    IsTableLookup = False
    If Field = "" Then IsTableLookup = True
    
    On Error Resume Next
    If UBound(Fields) - LBound(Fields) = 0 Then Fields = ¨(Fields)
    On Error GoTo 0
    FieldPos = GetFieldPos(ShtName, Field)
    
    ' If we're trying to retrieve a single field and can't find it, then raise an error
    If Field <> "" And FieldPos = 0 And Not RowIndexLookup Then GoTo RaiseNoMatchForField
    
    NumberOfFields = 1
    If Arrays.IsArrayAllocated(Fields) Then NumberOfFields = UBound(Fields)
    
    If Field = "" Then
        'NumberOfFields = 0
        ReDim ThisResult(0 To NumberOfFields - 1)
        ReDim Matches(0 To NumberOfFields - 1, 0 To 0) As Variant
        ReDim MatchesPos(0 To NumberOfFields - 1) As Variant
        For x = 1 To NumberOfFields
            ' We must cast the value to string to account for formulas in the
            ' lookup filters
            MatchesPos(x - 1) = GetFieldPos(ShtName, CStr(Fields(x)))
        Next
        
    End If
    
    If (RowIndexLookup = False And FieldPos = 0 And Field <> "") Or ShtName = "" Or Field = "" Then FlexLookup_ = Results
    
    ' If this is being called from a worksheet, there must be a selection range where the
    ' results will go, so we can stop iterating once we reach that number
    ' If there's no Caller.Rows/Caller.Columns, though, then just use the default 100
    On Error Resume Next
    CallerRows = 0
    CallerCols = 0
    CallerRows = Caller.Rows.Count
    CallerCols = Caller.Columns.Count
    MaxResultsSize = CallerRows * CallerCols
    On Error GoTo 0
    If MaxResultsSize = 0 Then MaxResultsSize = MAX_RESULTS_SIZE
    
    ' Make Lookups() from the ProtoLookups() sent from the wrapper functions
    If Not IsMissing(ProtoLookups) Then
        ReDim Lookups(0 To UBound(ProtoLookups))
        For x = LBound(ProtoLookups) To UBound(ProtoLookups)
            ' We must cast the value to string to account for formulas in the
            ' lookup filters
            Lookups(x) = CStr(ProtoLookups(x))
        Next
        
        ' Set some defaults before starting the loop
        L = UBound(Lookups) - LBound(Lookups) + 1
        NonBlankFilters = L
    End If
    
    FieldCount = 0
    ResultsSize = 0
    x = 0
    
    If L > 0 Then
        For i = LBound(Lookups) To UBound(Lookups) Step 2
            If ((Lookups(i) <> "") And (Lookups(i + 1) <> "")) Then
                FieldCount = FieldCount + 1
            Else
                NonBlankFilters = NonBlankFilters - 2
            End If
        Next
        
        If NonBlankFilters > 0 Then FieldCount = FieldCount - 1

        ReDim LookupFields(0 To FieldCount)
        ReDim LookupValues(0 To FieldCount)
        ReDim LookupPos(0 To FieldCount)
        
        If NonBlankFilters <= 0 Then GoTo StartReturn
        
        x = 0
        For i = LBound(Lookups) To UBound(Lookups) Step 2
            If ((Lookups(i) <> "") And (Lookups(i + 1) <> "")) Then
                LookupFields(x) = CStr(Lookups(i))
                LookupValues(x) = CStr(Lookups(i + 1))
                LookupPos(x) = GetFieldPos(ShtName, Lookups(i))
                x = x + 1
            End If

        Next i
        x = 0
        
    Else
        FieldCount = 0
    
    End If
    
StartReturn:
    With ActiveWorkbook.Sheets(ShtName)
        lastrow = .UsedRange.Rows.Count
        For xRow = 2 To lastrow Step 1
            Append = True

            If L = 0 Or NonBlankFilters = 0 Then
                If RowIndexLookup = True Then
                    InsertedValue = xRow
                ElseIf Field = "" Then
                    For x = LBound(MatchesPos) To UBound(MatchesPos)
                        s = .Cells(xRow, MatchesPos(x)).Value
                        ThisResult(x) = s
                    Next
                Else
                    InsertedValue = .Cells(xRow, FieldPos).Value
                    InsertedValueLength = Len(InsertedValue)
                End If
            Else
                For xField = LBound(LookupFields) To UBound(LookupFields)
                    Rowvalue = .Cells(xRow, LookupPos(xField)).Value
                    If LookupValues(xField) <> "" And CStr(Rowvalue) <> LookupValues(xField) Then
                        Append = False
                        GoTo SkipAppending
                    End If
                Next xField
                If RowIndexLookup Then
                    InsertedValue = xRow
                ElseIf Field = "" Then
                    For x = LBound(MatchesPos) To UBound(MatchesPos)
                        If MatchesPos(x) = 0 Then 'Skip empty columns
                            s = ""
                        Else
                            s = .Cells(xRow, MatchesPos(x)).Value
                        End If
                        ThisResult(x) = s
                    Next
                    InsertedValue = SHA1HASH(Join(ThisResult, ""))
                    InsertedValueLength = 1
                    
                Else
                    InsertedValue = .Cells(xRow, FieldPos).Value
                    InsertedValueLength = Len(InsertedValue)
                End If
            
            End If
            
            'Prevent errors in cells from propagating through the Function
            If IsError(InsertedValue) Then
                Append = False
                InsertedValue = ""
                InsertedValueLength = 0
            End If
            
            If Append = True And LastValue <> InsertedValue And InsertedValueLength > 0 Then
                Inserted = False
                If Field = "" Then
                    LastValue = SHA1HASH(Join(ThisResult, ""))
                    Inserted = AppendToArrayUniquely(Matches, ThisResult, PreviousResults, InsertedValue)
                    On Error Resume Next
                    PreviousResults.Add InsertedValue, 1
                    On Error GoTo 0
                    Z = 1
                Else
                    LastValue = InsertedValue
                    Inserted = AppendToArrayUniquely(Matches, InsertedValue)
                End If
                
                If Inserted = True Then
                    'If Field = "" Then
                        ResultsSize = ResultsSize + 1 * NumberOfFields
                    'Else
                    '    ResultsSize = ResultsSize + 1
                    'End If
                    If ResultsSize >= MaxResultsSize Then GoTo ReturnResults
                End If

            End If

SkipAppending:
        Next xRow
        
    End With
    
ReturnResults:
    If Sorted And Field <> "" Then
        Call QSortInPlace(Matches)
    End If
    
    On Error GoTo SimpleReturn
    If Field <> "" Then
        FlexLookup_ = ReturnArray(Matches, Caller)
    Else
        ReDim Matches2( _
            LBound(Matches, 1) To UBound(Matches, 2), _
            LBound(Matches, 2) To UBound(Matches, 1))
        Call TransposeArray(Matches, Matches2)
        Call QuickSortArray(Matches2, , , 0, vbTextCompare)
        Call TransposeArray(Matches2, Matches)
        Erase Matches2
        FlexLookup_ = ReturnTable(Matches, Caller)
    End If

    If Field = "" Then
        Z = 1
    End If
    GoTo ExitCleanly
    
SimpleReturn:
    FlexLookup_ = Matches
    If Matches(0) = vbEmpty Then
        FlexLookup_ = Nothing
    End If
    GoTo ExitCleanly
    
ErrHandler:
    FlexLookup_ = 0
    GoTo ExitCleanly

RaiseNoMatchForField:
    FlexLookup_ = "Error 1: No match for field '" & Field & "'."
    GoTo ExitCleanly
    
ExitCleanly:
    If UBound(Matches) - LBound(Matches) = 0 And VarType(Matches(0)) = vbEmpty Then
        Set FlexLookup_ = Nothing
    End If
    Set PreviousResults = Nothing
    Set ThisResult = Nothing
    Exit Function

End Function

Public Function COUNTB(ByVal V As Variant) As Long
    COUNTB = 0
    
    If IsError(V) Then Exit Function
    If Not IsArrayEmpty(V) Then
        If UBound(V) - LBound(V) = 0 Then
            If V(UBound(V)) = vbEmpty Then Exit Function
        End If
        
        COUNTB = WorksheetFunction.CountA(V)
    End If
    
End Function

Public Function UniqueLookup(Field As String, Optional Sorted As Boolean = False) As Variant
    Dim FieldPos As Long, xRow As Long, ReturnRows As Long
    Dim Results() As Variant
    ReDim Results(0 To 0) As Variant
    
    FieldPos = GetFieldPos(Field)
    
    ResultsSize = 0
    MaxResultsSize = Application.Caller.Rows.Count * Application.Caller.Columns.Count
    
    With ActiveWorkbook.Sheets("Base")
        lastrow = .UsedRange.Rows.Count
        
        For xRow = 2 To lastrow Step 1
            InsertedValue = .Cells(xRow, FieldPos).Value
            If LastValue <> InsertedValue Then
                LastValue = InsertedValue
                Inserted = False
                Inserted = AppendToArrayUniquely(Results, InsertedValue)
                
                If Inserted = True Then
                    ResultsSize = ResultsSize + 1
                    If ResultsSize >= MaxResultsSize Then GoTo ReturnResults
                End If
            End If
                        
        Next xRow
    
    End With
    
ReturnResults:
    
    If Sorted Then
        Call QSortInPlace(Results)
    End If
    
    On Error GoTo SimpleReturn
    UniqueLookup = ReturnArray(Results, Application.Caller)
    Exit Function
    
SimpleReturn:
    UniqueLookup = Results
    Exit Function
                                                  
End Function

Public Function GetFieldPos(ByVal ShtName As String, ByVal Field As String)
    On Error GoTo ErrHandler
    With Application.WorksheetFunction
        GetFieldPos = .Match(Field, ActiveWorkbook.Sheets(ShtName).Range("1:1"), 0)
        Exit Function
    End With
    
ErrHandler:
    GetFieldPos = 0
    On Error GoTo 0
    Exit Function
    
End Function

Public Function GetPivotTerra(ByVal DataFieldName As String, ByRef PTRange As Range, ParamArray OpArgs() As Variant) As Variant
    
    If IsMissing(OpArgs) Then
        GetPivotTerra = 0
        GoTo Ex
    End If
    
    With Application
        SU = .ScreenUpdating
        CU = .Calculation
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    Dim PT As PivotTable
    Dim DataField As PivotField
    Set PT = PTRange.PivotTable
    
    Dim ParsedArgs As Variant
    ReDim ParsedArgs(0 To 0)
    
    Dim x As Long
    x = 0
    For i = LBound(OpArgs) To UBound(OpArgs) Step 2
        If OpArgs(i + 1) = "" Then GoTo SkipEmpty
        If x = 0 Then
            ParsedArgs(x) = OpArgs(i)
        Else
            Call InsertElementIntoArray(ParsedArgs, x, OpArgs(i))
        End If
        'Select Case UCase(OpArgs(i))
        
        'Case "MES"
        '    thisValue = Int(OpArgs(i + 1))
        
        'Case Else
            On Error GoTo IsStr
            thisValue = Int(OpArgs(i + 1))
            GoTo IsInt
IsStr:
            thisValue = CStr(OpArgs(i + 1))
            Resume Next
IsInt:
        
        'End Select
        
        'Debug.Print x + 1 & " " & ThisValue
        Call InsertElementIntoArray(ParsedArgs, x + 1, thisValue)
        
        x = x + 2
SkipEmpty:
    Next
    
    Select Case UBound(ParsedArgs) - LBound(ParsedArgs) + 1
        Case 2
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1))
        
        Case 4
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3))
        
        Case 6
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5))
        
        Case 8
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7))
        
        Case 10
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9))
            
        Case 12
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9), ParsedArgs(10), ParsedArgs(11))
            
        Case 14
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9), ParsedArgs(10), ParsedArgs(11), _
            ParsedArgs(12), ParsedArgs(13))
                        
        Case 16
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9), ParsedArgs(10), ParsedArgs(11), _
            ParsedArgs(12), ParsedArgs(13), ParsedArgs(14), ParsedArgs(15))
                        
        Case 18
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9), ParsedArgs(10), ParsedArgs(11), _
            ParsedArgs(12), ParsedArgs(13), ParsedArgs(14), ParsedArgs(15), _
            ParsedArgs(16), ParsedArgs(17))
                        
        Case 20
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9), ParsedArgs(10), ParsedArgs(11), _
            ParsedArgs(12), ParsedArgs(13), ParsedArgs(14), ParsedArgs(15), _
            ParsedArgs(16), ParsedArgs(17), ParsedArgs(18), ParsedArgs(19))

        Case 22
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9), ParsedArgs(10), ParsedArgs(11), _
            ParsedArgs(12), ParsedArgs(13), ParsedArgs(14), ParsedArgs(15), _
            ParsedArgs(16), ParsedArgs(17), ParsedArgs(18), ParsedArgs(19), _
            ParsedArgs(20), ParsedArgs(21))

        Case 24
            GetPivotTerra = PT.GetPivotData(DataFieldName, _
            ParsedArgs(0), ParsedArgs(1), ParsedArgs(2), ParsedArgs(3), _
            ParsedArgs(4), ParsedArgs(5), ParsedArgs(6), ParsedArgs(7), _
            ParsedArgs(8), ParsedArgs(9), ParsedArgs(10), ParsedArgs(11), _
            ParsedArgs(12), ParsedArgs(13), ParsedArgs(14), ParsedArgs(15), _
            ParsedArgs(16), ParsedArgs(17), ParsedArgs(18), ParsedArgs(19), _
            ParsedArgs(20), ParsedArgs(21), ParsedArgs(22), ParsedArgs(23))
            
        Case Else
            GetPivotTerra = 0
            GoTo Ex
            
    End Select
        
    
Ex:
    'Restore original application status
    With Application
        .ScreenUpdating = SU
        .Calculation = CU
    End With
    
End Function

Function ¨(ParamArray Args() As Variant) As Variant
    ¨ = Args()
    
End Function

Function GetRow(ByVal Arr As Variant, Optional ByVal Index As Long = 0) As Variant
    Dim Results As Variant
    ReDim Results(LBound(Arr) To UBound(Arr))
    
    For i = LBound(Arr) To UBound(Arr) - 1
        Results(i) = Arr(i)(Index)
    Next
    
    Z = 1
    GetRow = Results
    
End Function
