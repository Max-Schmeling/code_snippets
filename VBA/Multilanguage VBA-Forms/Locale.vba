Option Private Module
Option Explicit

Public Language As Integer ' 0 to x ' where zero is the first language (eng) in Sheet_Locale headerrow and x ist the last


Public Function LocaleText(TextID, Optional Variables As Variant) As String
    ' This function needs to be called in all places where a text with variable language should be positioned.
    ' This means, instead of hardcoding your desired text string, you call this function and supply the <TextID>
    ' in column A in Sheet_Locale. Then the function will return the text in the line of the given TextID
    ' in the language determined by the public variable <Language>.
    ' If the text needs to contain substrings - such as numbers or whatever - the corresponding substrings
    ' can be replaced with '{x}' in the text template of each respective language in the Sheet_Locale. In order
    ' to supply a value for the variable, the second argument <Variables> is an array which should contain the
    ' contents of the respective occurrences of '{x}' in order from left to right.
    Dim LocaleArr As Variant
    Dim r As Long
    
    LocaleArr = Sheet_Locale.Range(Sheet_Locale.Range("D2"), Sheet_Locale.Cells(Sheet_Locale.Rows.Count, 1).End(xlUp))
    
    For r = LBound(LocaleArr, 1) To UBound(LocaleArr, 1)
        'Debug.Print LocaleArr(r, 1), TextID, CStr(LocaleArr(r, 1)) = TextID
        If CStr(LocaleArr(r, 1)) = TextID Then
        
            LocaleText = LocaleArr(r, 3 + Language) ' Text in current <Language>
            If LocaleText Like "*{#}*" And IsArray(Variables) Then ' Format text if it contains format codes. Ie. {#}
                LocaleText = ResolveLocaleText(LocaleText, Variables)
            End If
            
            Exit Function
        End If
    Next r
    
    LocaleText = "NA" ' this may never ocurre
End Function


Private Function ResolveLocaleText(FormatText As String, Variables As Variant) As String
    ' Replaces '{x}' in <FormatText> with the variabels in <Variables>
    ' x in '{x}' is an index starting at 0.
    ' This function is called by the Function LocaleText(). No need to call directly.
    
    Dim i As Integer
    Dim Item As Variant
    Dim Text As String
    
    Text = FormatText
    i = 0
    
    For Each Item In Variables
        Text = Replace(Text, "{" & i & "}", CStr(Variables(i)))
        i = i + 1
    Next Item
    
    ResolveLocaleText = Text
End Function