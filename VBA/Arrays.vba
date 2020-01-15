Function IsInArray(SearchString As String, arr As Variant) As Boolean
    ' Returns true if the given array <arr> contains the <SearchString>
    ' For Microsoft Excel by Max Schmeling

    Dim Item As Variant
    For Each Item In arr
        If Item = SearchString Then
            IsInArray = True
            Exit Function
        End if
    Next Item
    IsInArray = False
End Function


Function Length(arr As Variant) As Long
    ' Returns the amount of items in the array <arr>
    ' For Microsoft Excel by Max Schmeling

    If IsEmpty(arr) Then
       Length = 0
    Else
       Length = UBound(arr) - LBound(arr) + 1
    End If
End Function


Function CountDistinct(arr As Variant, Optional CompareMode As Integer = 1) As Long
    ' Returns the amount of distinct items in the given array <arr>
    ' <CompareMode> can be one of:
    ' - vbBinaryCompare = 0 (case-sensitive, faster)
    ' - vbTextCompare = 1 (case-insensitive, slower, default)
    '
    ' Written by Max Schmeling
    
    Dim Dict As Object
    Dim Item As Variant
    
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = CompareMode
    
    For Each Item In arr
        If Not Dict.exists(Item) Then
            Dict.Add Item, Nothing
        End If
    Next Item
    
    CountDistinct = Dict.Count
End Function


Function CountDistinctConditional(arr As Variant, _
                                Optional TargetColumn As Integer = -1, _
                                Optional CriteriaColumn As Integer = -1, _
                                Optional Criteria As Variant = "", _
                                Optional CompareMode As Integer = 1) As Long
    ' Returns the amount of distinct items in the given array <arr>.
    ' If <arr> is at least 2-Dimensional a seperate <CriteriaColumn>
    ' and a <Criteria> can be specified. The function will then return
    ' the amount of distinct items in <TargetColumn> that ocurre in the
    ' rows where <Criteria> occurres in <Criteria Column>. If <arr> is
    ' 1-Dimensional there is no need to specify <TargetColumn>.
    '
    ' <CompareMode> can be one of:
    ' - vbBinaryCompare = 0 (case-sensitive, faster)
    ' - vbTextCompare = 1 (case-insensitive, slower, default)
    '
    ' Written by Max Schmeling
    
    Dim Dict As Object
    Dim Item As Variant
    Dim CrItem As Variant
    Dim r As Long
    Dim c As Long
    Dim SkipMulti As Boolean
    
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = CompareMode
    
    TargetColumn = IIf(TargetColumn = -1, LBound(arr, 2), TargetColumn)
    CriteriaColumn = IIf(CriteriaColumn = -1, LBound(arr, 2), CriteriaColumn)
    SkipMulti = IIf(Criteria = vbNullString, True, False)
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        Item = arr(r, TargetColumn)
        CrItem = arr(r, CriteriaColumn)
        If Criteria = CrItem Or SkipMulti Then
            If Not Dict.exists(Item) Then
                Dict.Add Item, Nothing
            End If
        End If
    Next r
    
    CountDistinctConditional = Dict.Count
End Function
