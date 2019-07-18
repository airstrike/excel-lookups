Attribute VB_Name = "Lookups"
Const MAX_RESULTS_SIZE As Long = 100
Const DEFAULT_RETURN_VALUE As Variant = ""

Public Function RowLookup(ByVal Location As Variant, _
                            ByVal Unique As Boolean, ByVal Sorted As Boolean, _
                            ParamArray Lookups() As Variant) As Variant
    RowLookup = FlexLookup_( _
        Caller:=Application.Caller, Location:=Location, _
        Field:="", Unique:=Unique, Sorted:=Sorted, _
        RowLookup:=True, _
        ProtoLookups:=Lookups)

End Function

Public Function MultiLookup(ByVal Location As Variant, ByVal Field As String, _
                            ByVal Unique As Boolean, ByVal Sorted As Boolean, _
                            ParamArray Lookups() As Variant) As Variant

    MultiLookup = FlexLookup_( _
        Caller:=Application.Caller, Location:=Location, _
        Field:=Field, Unique:=Unique, Sorted:=Sorted, _
        ProtoLookups:=Lookups)

End Function

Public Function TableLookup(ByVal Location As Variant, ByVal Fields As Variant, _
                            ByVal Unique As Boolean, _
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
    Optional ByVal Unique As Boolean = True, Optional ByVal Sorted As Boolean = True, _
    Optional ByVal RowLookup As Boolean = False) As Variant
    ' ------------------------------------------------------------------------------------
    ' As a general rule of thumb, this function should not be accessed directly,
    ' but rather from one of its wrapper functions
    ' ------------------------------------------------------------------------------------
    Dim LocationRange As Range
    Set LocationRange = GetLocationRange(Location)

    
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
    
    ' If we're trying to retrieve a single field and can't find it, then raise an error
    If Field <> "" And FieldPos = 0 And Not RowIndexLookup Then GoTo RaiseNoMatchForField
    
    NumberOfFields = 1
    If Arrays.IsArrayAllocated(Fields) And UBound(Fields) > 0 Then NumberOfFields = UBound(Fields)
    
    On Error Resume Next
    Set Fields = Fields(0)
    On Error GoTo 0
    
    If Field = "" Then
        If RowLookup Then
            NumberOfFields = LocationRange.Columns.Count
        Else
            NumberOfFields = WorksheetFunction.Max( _
                WorksheetFunction.Min(Caller.Columns.Count, Fields.Columns.Count), _
                Fields.Rows.Count)
        End If
        
        'If RowLookup And UBound(Fields) = 0 Then Fields = LocationRange.Rows(0)
        ReDim ThisResult(0 To NumberOfFields - 1)
        ReDim Matches(0 To NumberOfFields - 1, 0 To 0) As Variant
        ReDim MatchesPos(0 To NumberOfFields - 1) As Variant
        For x = 1 To NumberOfFields
            ' We must cast the value to string to account for formulas in the
            ' lookup filters
            MatchesPos(x - 1) = GetFieldPos(Location, CStr(Fields(x)))
        Next
        
    End If
    
    If (RowIndexLookup = False And FieldPos = 0 And Field <> "") Or Field = "" Then FlexLookup_ = Results

    
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

            If L = 0 Or NonBlankFilters = 0 Then
                If RowIndexLookup = True Then
                    InsertedValue = xRow
                ElseIf Field = "" Then
                    For x = LBound(MatchesPos) To UBound(MatchesPos)
                        s = .Cells(xRow, MatchesPos(x)).Value
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
            
            If Append = True And (Unique = False Or LastValue <> InsertedValue) And InsertedValueLength > 0 Then
                Inserted = False
                If Field = "" Then
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
    If Sorted And Field <> "" Then
        Call QSortInPlace(Matches)
    End If
    
    'If UBound(Matches) = 0 And FirstElementInArray(Matches) <> vbEmpty Then GoTo SimpleReturn
    
    On Error GoTo SimpleReturn
    If Field <> "" Then
        If UBound(Matches) = 0 And FirstElementInArray(Matches) = vbEmpty Then Matches(0) = DEFAULT_RETURN_VALUE
        FlexLookup_ = ReturnArray(Matches, Caller)
        GoTo ExitCleanly
    Else
        ReDim Matches2( _
            LBound(Matches, 1) To UBound(Matches, 2), _
            LBound(Matches, 2) To UBound(Matches, 1))
        Call TransposeArray(Matches, Matches2)
        If Sorted Then Call QuickSortArray(Matches2, , , 0, vbTextCompare)
        Call TransposeArray(Matches2, Matches)
        Erase Matches2
        FlexLookup_ = ReturnTable(Matches, Caller)
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
