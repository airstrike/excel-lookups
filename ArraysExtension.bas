Attribute VB_Name = "ArraysExtension"
'This module was previously named Arrays2.
Public Function ReturnArray(Arr, Optional ByRef Caller As Variant) As Variant
    Dim RowX As Long, ColX As Long, n As Long
    n = 0
    If IsMissing(Caller) Then
        CallerRows = UBound(Arr) + 1
    ElseIf TypeName(Caller) = "Range" Then
        If Caller.HasFormula Then
            If InStr(Caller.FormulaArray, "=MultiLookup(") = 0 And _
               InStr(Caller.FormulaArray, "=TableLookup(") = 0 Then
                CallerRows = UBound(Arr) + 1
            Else
                CallerRows = Caller.Rows.Count
                CallerCols = Caller.Columns.Count
            End If
        End If
    Else
        CallerRows = Caller.Rows.Count
        CallerCols = Caller.Columns.Count
    End If
    If CallerRows = 1 And CallerCols = 1 Then
        Dim SingleResult(1 To 1)
        SingleResult(1) = Left(GetItem(Arr, n), 254)
        ReturnArray = SingleResult
        Exit Function
    Else
        ReDim Results(1 To CallerRows, 0 To 0)
        If CallerRows = 1 And CallerCols = 1 Then
            'If we return just one cell, excel will repeat it for every cell in the worksheet,
            'so we need to pad the remaining cells with #N/A for consistency
            ReDim Results(1 To 2, 0 To 0)
            Results(1, 0) = Left(GetItem(Arr, n), 254)
            Results(2, 0) = CVErr(xlErrNA)
        Else
            If CallerCols = 1 Or CallerCols = vbEmpty Then
                ReDim Results(1 To CallerRows, 0 To 0)
                For RowX = 1 To CallerRows
                    Results(RowX, 0) = Left(GetItem(Arr, n), 254)
                    n = n + 1
                Next RowX
            Else
                ReDim Results(0 To 0, 1 To CallerCols)
                For ColX = 1 To CallerCols
                    Results(0, ColX) = Left(GetItem(Arr, n), 254)
                    n = n + 1
                Next ColX
            End If
        End If
    End If
    
Exiting:
    ReturnArray = Results

End Function

Public Function ReturnTable(Arr, Optional ByRef Caller As Variant) As Variant
    If IsMissing(Caller) Then
        CallerRows = UBound(Arr, 1)
        CallerCols = UBound(Arr, 2)
    ElseIf TypeName(Caller) = "Range" Then
        If Caller.HasFormula Then
            If InStr(Caller.FormulaArray, "=MultiLookup(") = 0 And _
               InStr(Caller.FormulaArray, "=TableLookup(") = 0 Then
                CallerRows = UBound(Arr, 1)
                CallerCols = UBound(Arr, 2)
            Else
                CallerRows = Caller.Rows.Count
                CallerCols = Caller.Columns.Count
            End If
        End If
    Else
        CallerRows = Caller.Rows.Count
        CallerCols = Caller.Columns.Count
    End If

    ReDim ReturnTableResults(1 To CallerRows, 1 To CallerCols) As Variant
    For RowNdx = 1 To CallerRows
        For ColNdx = 1 To CallerCols
            n = n + 1
            ReturnTableResults(RowNdx, ColNdx) = GetItem2Dim(Arr, RowNdx - 1, ColNdx - 1)
        Next ColNdx
     Next RowNdx
    
    ReturnTable = ReturnTableResults

End Function


Function ExcludeEmpty(ByRef Arr As Variant)
    Dim Results As Variant
    ReDim Results(1 To 1)
    
    For x = LBound(Arr) To UBound(Arr)
        If Arr(x) <> vbEmpty Then Results(UBound(Results)) = Arr(x)
    Next
    
    ExcludeEmpty = Results
    
End Function

Function FindInArray(InputArray, Value) As Variant
    For i = LBound(InputArray) To UBound(InputArray)
        If InputArray(i) = Value Then
            FindInArray = i
            Exit Function
        End If
    Next i
    
    FindInArray = Null
End Function

Function PresentInArray(InputArray, Value) As Boolean
    For i = LBound(InputArray) To UBound(InputArray)
        If InputArray(i) = Value Then
            PresentInArray = True
            Exit Function
        End If
    Next i
    
    PresentInArray = False
End Function

Function AppendToArray(InputArray, Value, _
    Optional ByRef Reference As Object, Optional ByVal ReferenceItem As String, _
    Optional ByVal Index As Long = -1, Optional ByVal Uniquely As Boolean = False) As Boolean
    Dim bool_ As Boolean
    bool_ = False
    Insert = False
    
    If Uniquely = False Then
        Insert = True
    Else
        If Reference Is Nothing Then
            If Not PresentInArray(InputArray, Value) Then Insert = True
        Else
            If Not Reference.Exists(ReferenceItem) Then Insert = True
        End If
    End If
    
    If Insert = True Then
        If Index = -1 Then Index = UBound(InputArray)
        If UBound(InputArray) = 0 And FirstElementInArray(InputArray) = vbEmpty Then
            On Error Resume Next
            If IsArray(Value) Then 'TableLookup
                If UBound(Value) = 0 Then 'TableLookup with a single column
                    InputArray(0) = Value(0) 'Avoids creating Array(0,0)(0) construct
                    InputArray(0, 0) = Value(0)
                End If
            Else
                InputArray(0) = Value
                InputArray(0, 0) = Value
            End If
            On Error GoTo 0
            bool_ = True
        Else
            If Arrays.NumberOfArrayDimensions(InputArray) = 1 Then
                ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray))
                bool_ = InsertElementIntoArray(InputArray, Index + 1, Value)
            Else
                bool_ = False
                If InputArray(0, 0) <> vbEmpty Then
                    ReDim Preserve InputArray( _
                        LBound(InputArray, 1) To UBound(InputArray, 1), _
                        LBound(InputArray, 2) To UBound(InputArray, 2) + 1)
                    f = UBound(InputArray, 2)
                Else
                    f = 0
                End If
                For x = LBound(Value) To UBound(Value)
                    InputArray(x, f) = Value(x)
                    bool_ = True
                Next
                
            End If
        End If
    End If

    AppendToArray = bool_
        
End Function

Function GetItem(Arr, Index, Optional Default As String = "")
    On Error GoTo ErrHandler
    GetItem = Arr(Index)
    Exit Function
    
ErrHandler:
    GetItem = Default

End Function

Function GetItem2Dim(ByVal Arr As Variant, IndexA, IndexB, Optional Default As String = "") As Variant
    On Error GoTo ErrHandler
    GetItem2Dim = Arr(IndexB, IndexA)
    If GetItem2Dim = vbEmpty And VarType(GetItem2Dim) <> vbDouble Then GetItem2Dim = Default
    Exit Function
    
ErrHandler:
    GetItem2Dim = Default

End Function

Public Function SortRange(ByVal Rng As Range, _
    Optional lngMin As Long = -1, _
    Optional lngMax As Long = -1, _
    Optional lngColumn As Long = 0, _
    Optional CompareMode As VbCompareMethod)

    Dim Arr() As Variant
    Dim d As String
    d = ""
    Arr = Rng.Value
    Dim Results() As Variant
    ReDim Results( _
        LBound(Arr, 1) - 1 To UBound(Arr, 2) - 1, _
        LBound(Arr, 2) - 1 To UBound(Arr, 1) - 1)
        
    
    d = ""
    Call TransposeArray(Arr, Results)
    Call QuickSortArray(Results, , , 0, vbTextCompare)
    Call TransposeArray(Results, Arr)
    'SortTable = Arr
    'Erase Results
    SortTable = ReturnTable(Arr, Rng)

End Function

Public Function FirstElementInArray(Arr As Variant)
    If Arrays.NumberOfArrayDimensions(Arr) = 1 Then
        FirstElementInArray = Arr(LBound(Arr, 1))
    Else
        FirstElementInArray = Arr(LBound(Arr, 1), LBound(Arr, 2))
    End If
End Function

Public Function IsArrayReallyEmpty(Arr As Variant) As Boolean
    IsArrayReallyEmpty = False
    If FirstElementInArray(Arr) <> vbEmpty And IsArrayEmpty(Arr) <> True Then IsArrayReallyEmpty = True
End Function

