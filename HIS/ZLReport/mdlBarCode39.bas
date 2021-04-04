Attribute VB_Name = "mdlBarCode39"
'/****************************************************************************
' * Summary   : 条形码生成程序
' * Version   : 1.00
' * Start Date: 2004-6-07
' * My home   : http://www.mndsoft.com
' * E-Mail    : Mnd@Mndsoft.Com
' ****************************************************************************/
Public Const STR_CODE_39 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*"
Private Code_B() As Variant

Private BarH As Long
Private zBarText As String
Private xObj As Object

Private xPos As Long, xTop As Long, zHasCaption As Boolean
Private xStart As Integer, posCtr As Integer, xTotal As Long, chkSum As Long, WithCheckSum As Boolean
Private Const ChkChar = 43

Public Function DrawBarCode39(zObj As Object, ByVal zBarH As Integer, ByVal BarText As String, Optional ByVal zWithCheckSum As Boolean, Optional ByVal HasCaption As Boolean) As StdPicture
'参数：zObj=用于绘制条码的临时Picture对象
'      zBarH=条码图形高度级别，1-N，一般为3
'返回：已经绘制好的条码图形，自动宽度
    If BarText = "" Then BarText = "0123456789"
    
    Set xObj = zObj
    WithCheckSum = zWithCheckSum
    Init_Table
    zBarText = BarText
    zHasCaption = HasCaption
    
    'If Not CheckCode Then Exit Function 'byZT容错显示
    
    BarH = zBarH * 10
    xTop = 10
    
    xObj.AutoRedraw = True
    Set xObj.Picture = Nothing
    xObj.BackColor = vbWhite
    xObj.ScaleMode = 3
    
    If HasCaption Then
        xObj.Height = (xObj.TextHeight(zBarText) + BarH + 10) * Screen.TwipsPerPixelY
    Else
        xObj.Height = (BarH + 20) * Screen.TwipsPerPixelY
    End If
    'xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
'    xObj.Width = (((Len(zBarText) + IIf(WithCheckSum, 3, 2)) * 12 + 15)) * 16
    
    xObj.Cls
    Paint_Bar zBarText
    xObj.Width = (xPos + 12) * Screen.TwipsPerPixelX
    Paint_Bar zBarText
   
    Set DrawBarCode39 = xObj.Image
    xObj.ScaleMode = 1
End Function

Function CheckCode() As Boolean
    Dim ii As Integer
    
    zBarText = UCase(Replace(zBarText, "*", ""))
    
    For ii = 1 To Len(zBarText)
        If InStr(STR_CODE_39, Mid(zBarText, ii, 1)) = 0 Then
           GoTo Err_Found
        End If
    Next
    CheckCode = True
    Exit Function
Err_Found:
'    Err.Raise vbObjectError + 513, "Bar 39", _
'      "An Invalid Character Found in Bar Text"
    CheckCode = False
End Function

Private Sub Paint_Bar(xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer
 
    xTotal = 0
    xPos = 1
    posCtr = 0
    
    Draw_Bar CStr(Code_B(ChkChar))
    
    For ii = 1 To Len(xstr)
        posCtr = InStr(STR_CODE_39, Mid(xstr, ii, 1)) - 1
        If posCtr = -1 Then posCtr = 0 'byZT容错显示
        
        xTotal = xTotal + posCtr
        
        Draw_Bar CStr(Code_B(posCtr))
        
    Next
    chkSum = xTotal Mod 43
    
    If WithCheckSum Then Draw_Bar CStr(Code_B(chkSum))
    
    Draw_Bar CStr(Code_B(ChkChar))
    
    If zHasCaption Then
        xObj.CurrentX = ((xPos + 5) / 2) - xObj.TextWidth(xstr) / 2 '水平坐标
        xObj.CurrentY = 5 + BarH    ' 垂直坐标
        xObj.Print xstr   ' 大印信息
    End If
End Sub

Private Sub Draw_Bar(Encoding As String)
    Dim ii As Integer
    For ii = 1 To Len(Encoding)
        xPos = xPos + 1
        xObj.Line (xPos + 2, xTop - 7)-(xPos + 2, xTop + BarH - 7), IIF(Mid(Encoding, ii, 1), vbBlack, vbWhite)
    Next
    xPos = xPos + 1
    xObj.Line (xPos + 2, xTop - 7)-(xPos + 2, xTop + BarH - 7), vbWhite
End Sub

Private Sub Init_Table()
    Code_B = Array( _
             "101001101101", "110100101011", "101100101011", "110110010101", "101001101011", "110100110101", _
             "101100110101", "101001011011", "110100101101", "101100101101", "110101001011", "101101001011", _
             "110110100101", "101011001011", "110101100101", "101101100101", "101010011011", "110101001101", _
             "101101001101", "101011001101", "110101010011", "101101010011", "110110101001", "101011010011", _
             "110101101001", "101101101001", "101010110011", "110101011001", "101101011001", "101011011001", _
             "110010101011", "100110101011", "110011010101", "100101101011", "110010110101", "100110110101", _
             "100101011011", "110010101101", "100110101101", "100100100101", "100100101001", "100101001001", _
             "101001001001", "100101101101" _
             )
End Sub
