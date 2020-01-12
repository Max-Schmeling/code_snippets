Function GetSelectedIndex(ListBoxObj As MSForms.ListBox) As Long
    ' Returns the index (0-based) of the selected item if an item is selected.
    ' If no item is selected -1 will be returned.
    ' For Microsoft Office Forms by Max Schmeling

    If ListBoxObj.ListIndex > -1 Then
        If ListBoxObj.Selected(ListBoxObj.ListIndex) Then
            GetSelectedIndex = ListBoxObj.ListIndex
        Else
             GetSelectedIndex = -1
        End If
    Else
        GetSelectedIndex = -1
    End If
End Function


Function GetLabelByIndex(ListBoxObj As MSForms.ListBox, Index As Long, Optional Column As Long = 0) As Variant
    ' Returns the label/text of the item within <ListBoxObj> at row <Index> in column <Column>
    ' For Microsoft Office Forms by Max Schmeling

    If Index >= 0 And Index < ListBoxObj.ListCount Then
        GetLabelByIndex = ListBoxObj.List(Index, Column)
    Else
        Set GetLabelByIndex = "" ' Beware: This could lead to ambiguity errors with labels that are = ""
    End If
End Function


Sub SortListbox(ListBoxObj As MSForms.ListBox, Mode As Integer, Optional Sortkey As Integer = 0, Optional Ascending As Boolean = True)
    ' Bubble sort algorithm for MSForms ListBoxes.
    ' ListBoxObj = reference to listbox object
    ' Mode = (1: numerical sorting, 2: alphabetical sorting)
    ' Sortkey = the column index to sort by 0-based
    ' Ascending = sort ascendingly or descendingly
    ' For Microsoft Office Forms by Max Schmeling

    Dim i As Long
    Dim j As Long
    Dim c As Integer
    Dim temp As Variant
    Dim LbList As Variant
    
    If ListBoxObj.ListCount >= 2 Then
    
        'Store the list in an array for sorting
        LbList = ListBoxObj.List
    
        'Bubble sort the array
        For i = LBound(LbList, 1) To UBound(LbList, 1) - 1
            For j = i + 1 To UBound(LbList, 1)
            
                If Ascending Then ' Sort ascending
                    If Mode = 1 Then ' Sort numerically
                        If CDbl(LbList(i, Sortkey)) > CDbl(LbList(j, Sortkey)) Then
                            ' swap the values
                            For c = LBound(LbList, 2) To UBound(LbList, 2) - 1
                                temp = LbList(i, c)
                                LbList(i, c) = LbList(j, c)
                                LbList(j, c) = temp
                            Next
                        End If
                    Else ' Sort alphabetically
                        If CStr(LbList(i, Sortkey)) > CStr(LbList(j, Sortkey)) Then
                            ' swap the values
                            For c = LBound(LbList, 2) To UBound(LbList, 2) - 1
                                temp = LbList(i, c)
                                LbList(i, c) = LbList(j, c)
                                LbList(j, c) = temp
                            Next
                        End If
                    End If
                    
                Else ' Sort descending
    
                    If Mode = 1 Then ' Sort numerically
                        If CDbl(LbList(i, Sortkey)) < CDbl(LbList(j, Sortkey)) Then
                            ' swap the values
                            For c = LBound(LbList, 2) To UBound(LbList, 2) - 1
                                temp = LbList(i, c)
                                LbList(i, c) = LbList(j, c)
                                LbList(j, c) = temp
                            Next
                        End If
                    Else ' Sort alphabetically
                        If CStr(LbList(i, Sortkey)) < CStr(LbList(j, Sortkey)) Then
                            ' swap the values
                            For c = LBound(LbList, 2) To UBound(LbList, 2) - 1
                                temp = LbList(i, c)
                                LbList(i, c) = LbList(j, c)
                                LbList(j, c) = temp
                            Next
                        End If
                    End If
    
                End If
            Next j
        Next i
        
        'Remove the contents of the listbox
        ListBoxObj.Clear
        
        'Repopulate with the sorted list
        ListBoxObj.List = LbList
    
    End If
End Sub