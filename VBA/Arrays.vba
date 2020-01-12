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