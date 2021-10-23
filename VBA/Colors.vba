Type RGB ' value range: 0-1
    Red As Double
    Green As Double
    Blue As Double
End Type


Type HSV ' value range: 0-1
    Hue As Double
    Saturation As Double
    Value As Double
End Type


Function Decimal_To_RGB(DecimalColor As Long) As RGB
    Dim cRGB As RGB
    cRGB.Red = (DecimalColor And 255) / 255
    cRGB.Green = (DecimalColor \ 256 And 255) / 255
    cRGB.Blue = (DecimalColor \ 256 ^ 2 And 255) / 255
    Decimal_To_RGB = cRGB
End Function


Function RGB_To_Decimal(cRGB As RGB) As Long
    RGB_To_Decimal = RGB(cRGB.Red * 255, cRGB.Green * 255, cRGB.Blue * 255)
End Function


Public Function RGB_To_HSV(cRGB As RGB) As HSV
    ' Taken from: https://axonflux.com/handy-rgb-to-hsl-and-rgb-to-hsv-color-model-c
    ' Translated from JS into VB(A)
    
    Dim cmin As Double
    Dim cmax As Double
    Dim d As Double
    Dim cHSV As HSV
    
    cmax = WorksheetFunction.max(cRGB.Red, cRGB.Green, cRGB.Blue)
    cmin = WorksheetFunction.min(cRGB.Red, cRGB.Green, cRGB.Blue)
    cHSV.Value = cmax
    d = cmax - cmin
    cHSV.Saturation = IIf(cmax = 0, 0, d / cmax)
    
    If cmax = cmin Then
        cHSV.Hue = 0
    Else
        Select Case cmax
            Case cRGB.Red
                cHSV.Hue = (cRGB.Green - cRGB.Blue) / d + IIf(cRGB.Green < cRGB.Blue, 6, 0)
            Case cRGB.Green
                cHSV.Hue = (cRGB.Blue - cRGB.Red) / d + 2
            Case cRGB.Blue
                cHSV.Hue = (cRGB.Red - cRGB.Green) / d + 4
        End Select
        cHSV.Hue = cHSV.Hue / 6
    End If
    
    RGB_To_HSV = cHSV
End Function


Function HSV_to_RGB(cHSV As HSV) As RGB
    ' Function taken from Python's stdlib and translated into (V)BA
    Dim i As Double
    Dim f As Double
    Dim p As Double
    Dim q As Double
    Dim t As Double
    Dim cRGB As RGB

    If cHSV.Saturation = 0# Then
        cRGB.Red = cHSV.Value
        cRGB.Green = cHSV.Value
        cRGB.Blue = cHSV.Value
    Else
        i = Int(cHSV.Hue * 6#) ' assumes that Int(...) truncates
        f = (cHSV.Hue * 6#) - i
        p = cHSV.Value * (1# - cHSV.Saturation)
        q = cHSV.Value * (1# - cHSV.Saturation * f)
        t = cHSV.Value * (1# - cHSV.Saturation * (1# - f))
        i = i Mod 6
        If i = 0# Then
            cRGB.Red = cHSV.Value
            cRGB.Green = t
            cRGB.Blue = p
        ElseIf i = 1# Then
            cRGB.Red = q
            cRGB.Green = cHSV.Value
            cRGB.Blue = p
        ElseIf i = 2# Then
            cRGB.Red = p
            cRGB.Green = cHSV.Value
            cRGB.Blue = t
        ElseIf i = 3# Then
            cRGB.Red = p
            cRGB.Green = q
            cRGB.Blue = cHSV.Value
        ElseIf i = 4# Then
            cRGB.Red = t
            cRGB.Green = p
            cRGB.Blue = cHSV.Value
        ElseIf i = 5# Then
            cRGB.Red = cHSV.Value
            cRGB.Green = p
            cRGB.Blue = q
        End If
    End If
    
    HSV_to_RGB = cRGB
End Function


Function Decimal_To_Hex(Color As Long, Optional Prefix As String = "") As String
    Dim RevHex As String
    Dim i As Integer
    
    RevHex = WorksheetFunction.Dec2Hex(Color, 6)
    ' Reverse byte order to represent RGB as web-/hexcode
    Decimal_To_Hex = Prefix
    For i = Len(RevHex) To 1 Step -1
        If i Mod 2 = 1 Then
            Decimal_To_Hex = Decimal_To_Hex & Mid$(RevHex, i, 2)
        End If
    Next
End Function

Function GetCellColor(Rng As Range, Mode As String, ColorFunction As String) As Variant
    ' Returns the color of the given <Rng> in the given <ColorFunction>
    ' @param <Mode>             can be one of: "fill", "font"
    ' @param <ColorFunction>    can be one of: "rgb", "hsv", "hex", decimal"
    
    Dim cRGB As RGB
    Dim cHSV As HSV
    Dim RevHex As String
    Dim i As Integer
    
    If Rng.Cells.CountLarge = 1 Then
        If Rng.Interior.ColorIndex <> xlNone Then
            If Mode = "fill" Then
                Select Case LCase(ColorFunction)
                    Case Is = "rgb"
                        cRGB = Decimal_To_RGB(Rng.Interior.Color)
                        GetCellColor = "RGB(" & CInt(cRGB.Red * 100) & "%, " & CInt(cRGB.Green * 100) & "%, " & CInt(cRGB.Blue * 100) & "%)"
                    Case Is = "hsv"
                        cHSV = RGB_To_HSV(Decimal_To_RGB(Rng.Interior.Color))
                        GetCellColor = "HSV(" & CInt(cHSV.Hue * 360) & "°, " & CInt(cHSV.Saturation * 100) & "%, " & CInt(cHSV.Value * 100) & "%)"
                    Case Is = "hex"
                        GetCellColor = Decimal_To_Hex(Rng.Interior.Color, "#")
                    Case Is = "decimal"
                        GetCellColor = Rng.Interior.Color
                End Select
            ElseIf Mode = "font" Then
                Select Case LCase(ColorFunction)
                    Case Is = "rgb"
                        cRGB = Decimal_To_RGB(Rng.Font.Color)
                        GetCellColor = "RGB(" & CInt(cRGB.Red * 100) & "%, " & CInt(cRGB.Green * 100) & "%, " & CInt(cRGB.Blue * 100) & "%)"
                    Case Is = "hsv"
                        cHSV = RGB_To_HSV(Decimal_To_RGB(Rng.Font.Color))
                        GetCellColor = "HSV(" & CInt(cHSV.Hue * 360) & "°, " & CInt(cHSV.Saturation * 100) & "%, " & CInt(cHSV.Value * 100) & "%)"
                    Case Is = "hex"
                        GetCellColor = Decimal_To_Hex(Rng.Font.Color, "#")
                    Case Is = "decimal"
                        GetCellColor = Rng.Font.Color
                End Select
            End If
        Else
            GetCellColor = CVErr(2000) ' Return #NULL Error
        End If
    Else
        GetCellColor = CVErr(2023) ' Return #REF/#BEZUG Error
    End If
End Function