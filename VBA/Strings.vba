Function Startswith(Str As String, SubStr As String, Optional CaseSensitive As Boolean = True) As Boolean
    ' Returns True if the String <Str> starts with the string <SubStr>.
    ' Upper and lower case will be ignored if <CaseSensitive> = False
    ' For Microsoft Excel by Max Schmeling

    If Not CaseSensitive Then
        Str = LCase(Str)
        SubStr = LCase(SubStr)
    End If
    
    If Left$(Str, Len(SubStr)) = SubStr Then
        Startswith = True
    Else
        Startswith = False
    End If
End Function


Function Endswith(Str As String, SubStr As String, Optional CaseSensitive As Boolean = True) As Boolean
    ' Returns True if the String <Str> ends with the string <SubStr>.
    ' Upper and lower case will be ignored if <CaseSensitive> = False
    ' For Microsoft Excel by Max Schmeling

    If Not CaseSensitive Then
        Str = LCase(Str)
        SubStr = LCase(SubStr)
    End If
    
    If Right$(Str, Len(SubStr)) = SubStr Then
        Endswith= True
    Else
        Endswith= False
    End If
End Function


Function RemoveLeadingZeros(Str As String) As String
    ' Returns leading zeros from a number string if 
    ' <Str> startswith at least one '0'.
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    RemoveLeadingZeros = Str
    
    For i = 1 To Len(Str)
        If Mid$(Str, i, 1) = "0" Then
            RemoveLeadingZeros = Right$(Str, Len(Str) - i)
        Else
            Exit Function
        End If
    Next i
End Function


Function RemoveLetters(strInput As String) As String
    ' Removes all ascii lettsers
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String

    For i = 1 To Len(strInput)
        Select Case (Asc(Mid(strInput, i, 1)))
            Case 65 To 90, 97 To 122, 192 To 255: ' include 32 if you want to remove spaces too
            Case Else
                strOutput = strOutput & Mid(strInput, i, 1)
        End Select
    Next

    RemoveLetters = strOutput
End Function


Function RemoveNumeric(strInput As String) As String
    ' Removes all digits 0 - 9
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String

    For i = 1 To Len(strInput)
        Select Case (Asc(Mid(strInput, i, 1)))
            Case 48 To 57:
            Case Else
                strOutput = strOutput & Mid(strInput, i, 1)
        End Select
    Next

    RemoveNumeric= strOutput
End Function


Function RemoveAlphaNumeric(strInput As String) As String
    ' Removes all ascii lettsers and digits 0 - 9
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String

    For i = 1 To Len(strInput)
        Select Case (Asc(Mid(strInput, i, 1)))
            Case 48 To 57, 65 To 90, 97 To 122, 192 To 255:
            Case Else
                strOutput = strOutput & Mid(strInput, i, 1)
        End Select
    Next
    RemoveAlphaNumeric= strOutput
End Function


Function RemoveInvisibles(strInput As String) As String
    ' Removes invisible ascii characters
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String

    For i = 1 To Len(strInput)
        Select Case (Asc(Mid(strInput, i, 1)))
            Case 0 To 31, 127: ' include 32 if you want to remove spaces too
            Case Else
                strOutput = strOutput & Mid(strInput, i, 1)
        End Select
    Next
    RemoveInvisibles= strOutput
End Function


Function Remove(strInput As String, charString As String) As String
    ' Removes all characters from <strInput> that ocurre in <charString>
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String
    strOutput = strInput

    For i = 1 To Len(charString)
        strOutput = Replace(strOutput, Mid$(charString, i, 1), "")
    Next i
    Remove = strOutput
End Function


Function OnlyLetters(strInput As String) As String
    ' Removes all characters that are not letters
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String

    For i = 1 To Len(strInput)
        Select Case Asc(Mid(strInput, i, 1))
            Case 65 To 90, 97 To 122, 192 To 255: 'include 32 if you want to include space
                strOutput = strOutput & Mid(strInput, i, 1)
        End Select
    Next
    OnlyLetters = strOutput
End Function


Function OnlyNumeric(strInput As String) As String
    ' Removes all characters that are not digits 0-9
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String

    For i = 1 To Len(strInput)
        Select Case Asc(Mid(strInput, i, 1))
            Case 48 To 57:
                strOutput = strOutput & Mid(strInput, i, 1)
        End Select
    Next
    OnlyNumeric = strOutput
End Function


Function OnlyAlphaNumeric(strInput As String) As String
    ' Removes all characters that are not letters or digits 0-9
    ' For Microsoft Excel by Max Schmeling

    Dim i As Integer
    Dim strOutput As String

    For i = 1 To Len(strInput)
        Select Case Asc(Mid(strInput, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122, 192 To 255: 'include 32 if you want to include space
                strOutput = strOutput & Mid(strInput, i, 1)
        End Select
    Next
    OnlyAlphaNumeric= strOutput
End Function