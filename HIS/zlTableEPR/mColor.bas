Attribute VB_Name = "mColor"
'#########################################################################
'##模 块 名：mColor.bas
'##创 建 人：吴庆伟
'##日    期：2005年4月24日
'##修 改 人：
'##日    期：
'##描    述：颜色转换函数库
'##版    本：
'#########################################################################

Option Base 0
Option Explicit
'----------------------------------------------------------------------------------
'注释：
'a) HSL values ranges from 0 to 240
'   In practice it can be anything except that Hue describes 6 colors,
'   so a multiple of 6 is advantageous.
'   MS Win uses 240
'b) RGB are by definition integer values ranging from 0 to 255.
'c) Hue is undefined if Saturation (S) = 0 (total grey)
'----------------------------------------------------------------------------------
DefLng A-N, P-Z
DefBool O

Public Const MaxHSL As Integer = 240
Public Const MaxRGB As Integer = 256   '(Actually 0.0 to 259.999..)

'HSL DataType
Public Type HSLColor
    Hue As Long
    Sat As Long
    Lum As Long
End Type

Public Function HSLtoRGB(HueLumSat As HSLColor) As Long
    '-------------------------------------------------------
    'Convert HSL to RGB color
    '-------------------------------------------------------
    Dim R, G, b
    Dim H, l, s
    Dim Tint1, Tint2
    
    H = HueLumSat.Hue
    l = HueLumSat.Lum
    s = HueLumSat.Sat
    
    If s = 0 Then             'Achromatic, no color, greyscale -> R=G=B
        R = l * MaxRGB / MaxHSL 'Luminescence, converted to the proper range
        G = R                   'All RGB values same in greyscale
        b = R
    Else
        'Get the Tint Component Factors, which when applied to the Hue, separates
        'the Hue into 2 distinctive colors
        If l <= MaxHSL / 2 Then
          Tint2 = l * (MaxHSL + s) / MaxHSL
        Else
          Tint2 = l + s - (l * s / MaxHSL)
        End If
        Tint1 = 2 * l - Tint2
        'Get the RGB colors, in MaxHSL units and convert to MaxRGB units
        R = HueToRGB(Tint1, Tint2, H + MaxHSL / 3) * MaxRGB / MaxHSL
        G = HueToRGB(Tint1, Tint2, H) * MaxRGB / MaxHSL
        b = HueToRGB(Tint1, Tint2, H - MaxHSL / 3) * MaxRGB / MaxHSL
    End If
    'Validate out of bounds
    If R >= MaxRGB - 1 Then R = MaxRGB - 1
    If G >= MaxRGB - 1 Then G = MaxRGB - 1
    If b >= MaxRGB - 1 Then b = MaxRGB - 1
    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If b < 0 Then b = 0
    HSLtoRGB = RGB(CInt(R), CInt(G), CInt(b))
End Function

Private Function HueToRGB(ByVal Tint1 As Long, ByVal Tint2 As Long, ByVal Hue As Long) As Long
    '---------------------------------------------------------------------
    'Utility function to convert color tints 1 & 2 + hue to a single value
    '---------------------------------------------------------------------
    'Do a range check on Hue as the value passed was changed to outside
    'the normal range
    If Hue < 0 Then Hue = Hue + MaxHSL
    If Hue > MaxHSL Then Hue = Hue - MaxHSL
    
    
    If Hue < MaxHSL / 6 Then
        HueToRGB = Tint1 + (Tint2 - Tint1) * Hue / MaxHSL * 6
    ElseIf Hue < MaxHSL / 2 Then
        HueToRGB = Tint2
    ElseIf Hue < MaxHSL * 2 / 3 Then
        HueToRGB = Tint1 + (Tint2 - Tint1) * (MaxHSL * 2 / 3 - Hue) / MaxHSL * 6
    Else
        HueToRGB = Tint1
    End If
End Function

Private Function Max(t1 As Variant, ParamArray T() As Variant) As Variant
    '----------------------------------------------------
    'Determine the maximum of all values
    'Any number can be given (minimum 2), in any datatype
    '----------------------------------------------------
    Dim x As Variant, i As Long
    
    x = t1
    For i = 0 To UBound(T)
        If T(i) > x Then
            x = T(i)
        End If
    Next
    Max = x
End Function
Private Function Min(t1 As Variant, ParamArray T() As Variant) As Variant
    '----------------------------------------------------
    'Determine the minimum of all values
    'Any number, any type
    '----------------------------------------------------
    Dim x As Variant, i As Long
    
    x = t1
    For i = 0 To UBound(T)
        If T(i) < x Then
            x = T(i)
        End If
    Next
    Min = x
End Function

Public Function RGBRed(RGBColoror As Long) As Long
    '-----------------------------------------------------------
    'Return the Red component from an RGB Color
    '-----------------------------------------------------------
    RGBRed = RGBColoror And &HFF
End Function

Public Function RGBGreen(RGBColor As Long) As Long
    '------------------------------------------------------------
    'Return the Green component from an RGB Color
    '------------------------------------------------------------
    RGBGreen = ((RGBColor And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBColor As Long) As Long
    '------------------------------------------------------------
    'Return the Blue component from an RGB Color
    '------------------------------------------------------------
    RGBBlue = (RGBColor And &HFF0000) / &H10000
End Function

Public Function RGBtoHSL(RGBColor As Long) As HSLColor
    '----------------------------------------------------------------
    'Returns an HSLColor datatype containing Hue, Luminescence and
    'Saturation given an RGB Color value
    'Default variables are LONG
    '----------------------------------------------------------------
    Dim R, G, b                     'RGB Unit Colors
    Dim H, s, l                     'HSL Unit Colors
    Dim cMax, cMin, cMinus, cPlus   'Color separation
    Dim RDelta As Double            'Unit color separation, as % of cMax
    Dim GDelta As Double
    Dim BDelta As Double
      
    'Get the component colors, 0 to 255
    R = RGBRed(RGBColor)
    G = RGBGreen(RGBColor)
    b = RGBBlue(RGBColor)
    
    'Get the highest & lowest color values
    cMax = Max(R, G, b) 'Highest
    cMin = Min(R, G, b) 'Lowest
      
    'cMinus & cPlus are used to simplify the calculations
    cPlus = cMax + cMin
    cMinus = cMax - cMin
      
    'Calculate luminescence (lightness)
    'L = ((cPlus * MaxHSL) + MaxRGB) / (2 * MaxRGB)
    l = cPlus / (2 * MaxRGB) * MaxHSL
    
    If cMax = cMin Then
        'Achromatic (r=g=b, -> greyscale)
        'Saturation is 0 for greyscale and Hue is undefined (so use 0 - pure red)
        s = 0
        H = 0
    Else
        'Calculate color saturation
        If l <= (MaxHSL / 2) Then
            s = cMinus / cPlus * MaxHSL
        Else
            s = (cMinus * MaxHSL) / (2 * MaxRGB - cPlus)
        End If
            
        'Calculate hue
        'Deltas range from 0 to 59.99999
        RDelta = (cMax - R) / cMinus * MaxHSL / 6
        GDelta = (cMax - G) / cMinus * MaxHSL / 6
        BDelta = (cMax - b) / cMinus * MaxHSL / 6
            
        Select Case cMax
        Case R          '-60 to +60
            H = BDelta - GDelta
        Case G          '+60 to +180
            H = RDelta - BDelta + MaxHSL / 3
        Case b          '+180 to 300
            H = GDelta - RDelta + MaxHSL * 2 / 3
        End Select
        
        If H < 0 Then H = H + MaxHSL    'Convert Hue to 0 to 359 units
    End If
    
    RGBtoHSL.Hue = H
    RGBtoHSL.Lum = l
    RGBtoHSL.Sat = s
End Function



