Attribute VB_Name = "ArraysExtension"
'This module was previously named Arrays2.
Public Function Foobar() As Variant
    Dim Results(0 To 1)
    Results(0) = "foo"
    Results(1) = "bar"
    
    Foobar = Results

End Function

Public Function ReturnArray(Arr, Optional ByRef Application_Caller As Variant) As Variant
    Dim RowX As Long, ColX As Long, n As Long
    n = 0
    'If IsMissing(Application_Caller) Then
        CallerRows = UBound(Arr) + 1
    'Else
    '    CallerRows = Application_Caller.Rows.Count
    '    CallerCols = Application_Caller.Columns.Count
    'End If
    ReDim Results(1 To CallerRows, 0 To 0)
    If CallerRows = 1 Then
        'If we return just one cell, excel will repeat it for every cell in the worksheet,
        'so we need to pad the remaining cells with #N/A for consistency
        ReDim Results(1 To 2, 0 To 0)
        Results(1, 0) = Left(GetItem(Arr, n), 254)
        Results(2, 0) = CVErr(xlErrNA)
    Else
        ReDim Results(1 To CallerRows, 0 To 0)
        For RowX = 1 To CallerRows
            Results(RowX, 0) = Left(GetItem(Arr, n), 254)
            n = n + 1
        Next RowX
    End If
    
Exiting:
    ReturnArray = Results

End Function

Public Function ReturnTable(Arr, Optional ByRef Application_Caller As Variant)
    If IsMissing(Application_Caller) Then
        CallerRows = UBound(Arr)
        CallerCols = 1
    Else
        CallerRows = Application_Caller.Rows.Count
        CallerCols = Application_Caller.Columns.Count
    End If

    ReDim Results(1 To CallerRows, 1 To CallerCols)
    For RowNdx = 1 To CallerRows
        For ColNdx = 1 To CallerCols
            n = n + 1
            Results(RowNdx, ColNdx) = GetItem2Dim(Arr, RowNdx - 1, ColNdx - 1)
        Next ColNdx
     Next RowNdx
    
    ReturnTable = Results

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

Function AppendToArrayUniquely(InputArray, Value, _
    Optional ByRef Reference As Object, Optional ByVal ReferenceItem As String, _
    Optional ByVal Index As Long) As Boolean
    Dim bool_ As Boolean
    bool_ = False
    Insert = False
    
    If Reference Is Nothing Then
        If Not PresentInArray(InputArray, Value) Then Insert = True
    Else
        If Not Reference.Exists(ReferenceItem) Then Insert = True
    End If
    
    If Insert = True Then
        If Index = -1 Then Index = UBound(InputArray) + 1
        If UBound(InputArray) = 0 And IsEmpty(InputArray(0)) Then
            InputArray(0) = Value
            bool_ = True
        Else
            If Arrays.NumberOfArrayDimensions(InputArray) = 1 Then
                ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray))
                bool_ = InsertElementIntoArray(InputArray, Index + 1, Value)
            Else
                If InputArray(0, 0) <> vbEmpty Then
                    'ReDim Preserve InputArray(0 To UBound(InputArray) + 1, _
                        LBound(InputArray, 2) To UBound(InputArray, 2))
                    ReDim Preserve InputArray( _
                        LBound(InputArray, 1) To UBound(InputArray, 1), _
                        LBound(InputArray, 2) To UBound(InputArray, 2) + 1)
                    f = UBound(InputArray, 2)
                Else
                    f = 0
                End If
                For x = LBound(Value) To UBound(Value)
                    InputArray(x, f) = Value(x)
                Next
            End If
        End If
    End If

    AppendToArrayUniquely = bool_
        
End Function

Function GetItem(Arr, Index, Optional Default As String = "")
    On Error GoTo ErrHandler
    GetItem = Arr(Index)
    Exit Function
    
ErrHandler:
    GetItem = Default

End Function

Function GetItem2Dim(Arr, IndexA, IndexB, Optional Default As String = "")
    On Error GoTo ErrHandler
    GetItem2Dim = Arr(IndexB, IndexA)
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
