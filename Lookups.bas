Attribute VB_Name = "Lookups"
Const MAX_RESULTS_SIZE As Long = 2500
Const DEFAULT_RETURN_VALUE As Variant = vbEmpty

Public Function RowLookup(ByVal Location As Variant, _
                            ByVal Unique As Long, ByVal Sorted As Variant, _
                            ParamArray Lookups() As Variant) As Variant
    RowLookup = FlexLookup_( _
        Caller:=Application.Caller, Location:=Location, _
        Field:="", Unique:=Unique, Sorted:=Sorted, _
        RowLookup:=True, _
        ProtoLookups:=Lookups)

End Function

Public Function MultiLookup(ByVal Location As Variant, ByVal Field As String, _
                            ByVal Unique As Long, ByVal Sorted As Variant, _
                            ParamArray Lookups() As Variant) As Variant

    MultiLookup = FlexLookup_( _
        Caller:=Application.Caller, Location:=Location, _
        Field:=Field, Unique:=Unique, Sorted:=Sorted, _
        ProtoLookups:=Lookups)

End Function

Public Function TableLookup(ByVal Location As Variant, ByVal Fields As Variant, _
                            ByVal Unique As Long, _
                            ParamArray Lookups() As Variant)

    TableLookup = FlexLookup_( _
        Caller:=Application.Caller, Location:=Location, _
        Fields:=Fields, Unique:=Unique, Sorted:=False, _
        ProtoLookups:=Lookups)

End Function

Private Function FlexLookup_( _
    ByRef Caller As Variant, _
    ByVal Location As Variant, _
    ByVal ProtoLookups As Variant, _
    Optional ByVal Field As String = "", Optional ByVal Fields As Variant, _
    Optional ByVal Unique As Long = 1, Optional ByVal Sorted As Variant = 1, _
    Optional ByVal RowLookup As Boolean = False) As Variant
    ' ------------------------------------------------------------------------------------
    ' As a general rule, this function should not be accessed directly. Instead, use one
    ' of its wrapper functions (TableLookup, MultiLookup, RowLookup) instead
    ' ------------------------------------------------------------------------------------
    Dim LocationRange As Range
    Set LocationRange = GetLocationRange(Location)

    
    Dim l As Long, x As Long, ProtoSize As Long, _
        FieldPos As Long, NonBlankFilters As Long, ResultSize As Long
    Dim Results() As Variant, Matches() As Variant, MatchesPos As Variant, _
        LookupFields() As Variant, LookupValues() As Variant, _
        LookupPos() As Variant
    
    Dim IsTableLookup As Boolean
    Dim RowIndexLookup As Boolean
    
    Dim Append As Boolean
    Dim LastValue As Variant
    Dim FieldCount As Long, FieldsRowsCount As Long, FieldsColumnsCount As Long
    Dim ReturnAsArray As Boolean
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
    FieldsBase = 0
    If TypeName(Fields) = "Range" Then FieldsBase = 1
    If UBound(Fields) - LBound(Fields) = 0 Then
        Fields = ¨(Fields)
    End If
    On Error GoTo 0
    FieldPos = GetFieldPos(Location, Field)
    
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
    
    ' If this is wrapped around a bigger formula, then return every possible result
    If TypeName(Caller) = "Range" Then
        If Caller.HasFormula Then
            If InStr(Caller.FormulaArray, "=MultiLookup(") = 0 And _
               InStr(Caller.FormulaArray, "=TableLookup(") = 0 Then
                ReturnAsArray = True
                MaxResultsSize = MAX_RESULTS_SIZE
            End If
        End If
    End If
    
    ' If we're trying to retrieve a single field and can't find it, then raise an error
    If Not IsTableLookup And FieldPos = 0 And Not RowIndexLookup Then GoTo RaiseNoMatchForField
    
    NumberOfFields = 1
    If Arrays.IsArrayAllocated(Fields) And UBound(Fields) > 0 Then NumberOfFields = UBound(Fields)
    
    On Error Resume Next
    Set Fields = Fields(0)
    On Error GoTo 0
    
    If IsTableLookup Then
        If RowLookup Then
            NumberOfFields = LocationRange.Columns.Count
        Else
            On Error Resume Next
            'If Fields for TableLookup() are provided directly in-cell e.g. {"foo","bar"}, we can't
            'access the .Columns / .Rows properties
            FieldsColumnsCount = UBound(Fields)
            FieldsRowsCount = UBound(Fields)
            FieldsColumnsCount = Fields.Columns.Count
            FieldsRowsCount = Fields.Rows.Count
            On Error GoTo 0
            'If ReturnAsArray Then
            '    NumberOfFields = Number
            'Else
                NumberOfFields = WorksheetFunction.Max( _
                    WorksheetFunction.Min(Caller.Columns.Count, FieldsColumnsCount), _
                    FieldsRowsCount)
            'End If
        End If
        
        'If RowLookup And UBound(Fields) = 0 Then Fields = LocationRange.Rows(0)
        ReDim ThisResult(0 To WorksheetFunction.Max(0, NumberOfFields - 1))
        ReDim Matches(0 To WorksheetFunction.Max(0, NumberOfFields - 1), 0 To 0) As Variant
        ReDim MatchesPos(0 To WorksheetFunction.Max(0, NumberOfFields - 1)) As Variant
        For x = FieldsBase To NumberOfFields 'FIXME? x = 1
            ' We must cast the value to string to account for formulas in the
            ' lookup filters
            On Error GoTo TrySingleElementArray
            MatchesPos(x - FieldsBase) = GetFieldPos(Location, CStr(Fields(x)))
            GoTo NextNumberOfFields
            
TrySingleElementArray:
            On Error GoTo 0
            MatchesPos(x - FieldsBase) = GetFieldPos(Location, CStr(Fields(x)(1)))
            
NextNumberOfFields:
            On Error GoTo 0
        Next
        
    End If
    
    If (RowIndexLookup = False And FieldPos = 0 And Not IsTableLookup) Or IsTableLookup Then FlexLookup_ = Results

    ' Make Lookups() from the ProtoLookups() sent from the wrapper functions
    If Not IsMissing(ProtoLookups) Then
        ReDim Lookups(0 To UBound(ProtoLookups))
        For x = LBound(ProtoLookups) To UBound(ProtoLookups)
            ' We must cast the value to string to account for formulas in the
            ' lookup filters
            Lookups(x) = CStr(ProtoLookups(x))
        Next
        
        ' Set some defaults before starting the loop
        l = UBound(Lookups) - LBound(Lookups) + 1
        NonBlankFilters = l
    End If
    
    FieldCount = 0
    ResultsSize = 0
    x = 0
    
    If l > 0 Then
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
                LookupPos(x) = GetFieldPos(Location, Lookups(i))
                x = x + 1
            End If

        Next i
        x = 0
        
    Else
        FieldCount = 0
    
    End If
    
StartReturn:
    With LocationRange
        lastrow = .Rows.Count
        For xRow = 2 To lastrow Step 1
            Append = True

            If l = 0 Or NonBlankFilters = 0 Then
                If RowIndexLookup = True Then
                    InsertedValue = xRow
                ElseIf IsTableLookup Then
                    For x = LBound(MatchesPos) To UBound(MatchesPos)
                        xMatch = MatchesPos(x)
                        If xMatch <> 0 Then
                            s = .Cells(xRow, MatchesPos(x)).Value
                        Else
                            s = vbEmpty
                        End If
                        ThisResult(x) = s
                    Next
                    InsertedValue = SHA1HASH(Join(ThisResult, ""))
                    InsertedValueLength = 1
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
                ElseIf IsTableLookup Then
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
            
            If Append = True And (Unique = False Or LastValue <> InsertedValue) And InsertedValueLength > 0 Then
                Inserted = False
                If IsTableLookup Then
                    LastValue = SHA1HASH(Join(ThisResult, ""))
                    Inserted = AppendToArray(Matches, ThisResult, PreviousResults, InsertedValue, Uniquely:=Unique)
                    On Error Resume Next
                    PreviousResults.Add InsertedValue, 1
                Else
                    LastValue = InsertedValue
                    Inserted = AppendToArray(Matches, InsertedValue, Uniquely:=Unique)
                End If
                
                If Inserted = True Then
                    ResultsSize = ResultsSize + 1 * NumberOfFields
                    If ResultsSize >= MaxResultsSize Then GoTo ReturnResults
                End If
            End If

SkipAppending:
        Next xRow
        
    End With
    
ReturnResults:
    If Sorted <> 0 And Not IsTableLookup Then
        Call QSortInPlace(Matches)
        If Sorted = -1 Then x = QSort.ReverseArrayInPlace(Matches)
    End If
    
    'If UBound(Matches) = 0 And FirstElementInArray(Matches) <> vbEmpty Then GoTo SimpleReturn
    
    'On Error GoTo TryToReturnAnything
    If IsTableLookup Then
        ReDim Matches2( _
            LBound(Matches, 1) To UBound(Matches, 2), _
            LBound(Matches, 2) To UBound(Matches, 1))
        Call TransposeArray(Matches, Matches2)
        If Sorted Then Call QuickSortArray(Matches2, , , 0, vbTextCompare)
        'If Sorted = -1 Then x = QSort.ReverseArrayInPlace(Matches2, True)
        Call TransposeArray(Matches2, Matches)
        Erase Matches2
        FlexLookup_ = ReturnTable(Matches, Caller)
    Else
        If UBound(Matches) = 0 And FirstElementInArray(Matches) = vbEmpty Then Matches(0) = DEFAULT_RETURN_VALUE
        FlexLookup_ = ReturnArray(Matches, Caller)
    End If
    
    'On Error GoTo 0
    'GoTo Returns

    
'TryToReturnAnything:
'    FlexLookup_ = Matches
    
'Returns:
    If Matches(0) = vbEmpty Then
        FlexLookup_ = Nothing
        Set PreviousResults = Nothing
        Set ThisResults = Nothing
        Exit Function
    End If
    GoTo ExitCleanly
    
ErrHandler:
    FlexLookup_ = DEFAULT_RETURN_VALUE
    GoTo ExitCleanly

RaiseNoMatchForField:
    FlexLookup_ = "Error 1: No match for field '" & Field & "'."
    GoTo ExitCleanly
    
ExitCleanly:
    On Error GoTo 0
    If UBound(Matches) - LBound(Matches) = 0 And _
        VarType(FirstElementInArray(Matches)) = vbEmpty Then 'And _
        'Matches = vbEmpty Then
        Set FlexLookup_ = DEFAULT_RETURN_VALUE
    End If
    Set PreviousResults = Nothing
    Set ThisResult = Nothing
    Exit Function

End Function

Public Function GetFieldPos(ByVal Location As Variant, ByVal Field As String)
    On Error GoTo ErrHandler
    Dim LocationRange As Range
    Set LocationRange = GetLocationRange(Location)
    
    With Application.WorksheetFunction
        GetFieldPos = .Match(Field, LocationRange.Rows(1), 0)
        Exit Function
    End With
    
ErrHandler:
    GetFieldPos = 0
    On Error GoTo 0
    Exit Function
    
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

    GetRow = Results
    
End Function

Function GetLocationRange(ByRef Location As Variant) As Range
    Dim r As Range
    On Error GoTo AsString
    Set r = Location
    GoTo EndFunction
    
AsString:
    Set r = ActiveWorkbook.Worksheets(Location).UsedRange
    GoTo EndFunction
    
EndFunction:
    Set GetLocationRange = r
End Function




