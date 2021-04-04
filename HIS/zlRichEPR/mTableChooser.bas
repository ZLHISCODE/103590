Attribute VB_Name = "mTableChooser"
'#########################################################################
'##模 块 名：mTableChooser.bas
'##创 建 人：吴庆伟
'##日    期：2005年3月1日
'##修 改 人：
'##日    期：
'##描    述：表格选择器相关模块
'#########################################################################

Option Explicit
Private m_lCurrentX As Long
Private m_lCurrentY As Long

'################################################################################################################
'## 功能：  设置图片尺寸为默认尺寸
'##
'## 参数：  picDropDown     :图片控件
'##
'## 说明：  默认为 6×6 表格
'################################################################################################################
Public Sub DefaultSize(ByRef picDropDown As PictureBox)
   picDropDown.Width = (6 + 24 * 5) * Screen.TwipsPerPixelX
   picDropDown.Height = (6 + 24 * 4 + 24) * Screen.TwipsPerPixelY
   picDropDown.Cls
End Sub

'################################################################################################################
'## 功能：  绘制表格下拉选择器
'################################################################################################################
Public Sub DrawTableChooser( _
      ByRef cDW As cDropDownToolWindow, _
      ByVal xPixels As Long, _
      ByVal yPixels As Long, _
      ByVal Button As MouseButtonConstants, _
      Optional ByRef bIn As Boolean, _
      Optional ByRef xCellHit As Long, _
      Optional ByRef yCellHit As Long)
      
    Dim tR As RECT
    Dim tWR As RECT
    Dim tJunk As POINTAPI
    Dim lhdc As Long
    Dim hPen As Long, hPenOld As Long
    Dim bInX As Boolean, bInY As Boolean
    Dim hBrIn As Long, hBrOut As Long, hBr As Long
    Dim x As Long, y As Long, tBoxR As RECT
    Dim xMax As Long, yMax As Long, sStatus As String
    Dim bResize As Boolean
   
   ' This example uses API methods to draw on the
   ' PictureBox. But you don't have to use API
   ' methods, VB drawing methods work just as well
   ' and seem to be very quick these days.
   
   
   ' Get size of pic
   GetClientRect cDW.DropDownObject.hWnd, tR
      
   ' Cache HDC for speed:
   lhdc = cDW.DropDownObject.hdc
            
   ' The client area:
   If cDW.InRect(xPixels * Screen.TwipsPerPixelX, yPixels * Screen.TwipsPerPixelY) Then
      ' in.  Caption is cancel and boxes up to xPixels,yPixels are highlighted:
      bIn = True
   Else
      ' not in.  Caption is cancel and all boxes are blank
      ' either we exceed in x or y and thus should increase size..
      If (Button = vbLeftButton) Then
         If xPixels > tR.Right Then
            tR.Right = tR.Right + 24
            bResize = True
         End If
         If yPixels > tR.Bottom Then
            tR.Bottom = tR.Bottom + 24
            bResize = True
         End If
         If bResize Then
            cDW.Resize tR.Right - tR.Left, tR.Bottom - tR.Top
            ' HDC has changed:
            lhdc = cDW.DropDownObject.hdc
            ' Clear the area:
            LSet tWR = tR
            tWR.Right = tWR.Right + 3
            tWR.Bottom = tWR.Bottom + 3
            hBr = GetSysColorBrush(COLOR_BTNFACE)
            FillRect lhdc, tWR, hBr
            DeleteObject hBr
         End If
         bIn = bResize
      Else
         bIn = False
      End If
   End If
        
   ' Draw the border to the drop down window:
   hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
   hPenOld = SelectObject(lhdc, hPen)
   Rectangle lhdc, tR.Left, tR.Top, tR.Right, tR.Bottom
   SelectObject lhdc, hPenOld
   DeleteObject hPen
   MoveToEx lhdc, 1, tR.Bottom - 2, tJunk
   hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT))
   hPenOld = SelectObject(lhdc, hPen)
   LineTo lhdc, 1, tR.Top + 1
   LineTo lhdc, tR.Right - 1, tR.Top + 1
   SelectObject lhdc, hPenOld
   DeleteObject hPen
   
   ' Modify the client rectangle to exclude the
   ' border and the "status bar"
   InflateRect tR, -6, -27
   OffsetRect tR, -3, -24
   
   ' Set up a selected and non-selected brush:
   hBrIn = GetSysColorBrush(COLOR_HIGHLIGHT)
   hBrOut = GetSysColorBrush(COLOR_WINDOW)
   
   ' Draw the table cells to pick from, highlighting
   ' the ones the mouse has "selected":
   For x = tR.Left To tR.Right Step 24
      tBoxR.Left = x + 1
      tBoxR.Right = x + 23
      If (xPixels >= tBoxR.Left) Then
         bInX = True
      Else
         bInX = False
      End If
      For y = tR.Top To tR.Bottom Step 24
         tBoxR.Top = y + 1
         tBoxR.Bottom = y + 23
         If (yPixels > tBoxR.Top) Then
            bInY = True
         Else
            bInY = False
         End If
         If (bIn And bInX And bInY) Then
            hBr = hBrIn
            If (x > xMax) Then
               xMax = x
            End If
            If (y > yMax) Then
               yMax = y
            End If
         Else
            hBr = hBrOut
         End If
         FillRect lhdc, tBoxR, hBr
      Next y
   Next x
   DeleteObject hBrIn
   DeleteObject hBrOut
   
   ' Draw the "status bar"
   tR.Left = tR.Left + 1
   tR.Right = tR.Right + 6
   tR.Top = tR.Bottom + 27
   tR.Bottom = tR.Top + 20
   If (bIn) Then
      xCellHit = xMax \ 24 + 1
      yCellHit = yMax \ 24 + 1
      m_lCurrentX = xCellHit
      m_lCurrentY = yCellHit
      sStatus = yCellHit & " × " & xCellHit & " 表格"
   Else
      m_lCurrentX = 0
      m_lCurrentY = 0
      sStatus = "取消"
   End If
   DrawStatusText lhdc, tR, sStatus, 0
   
   ' Show the changes on screen:
   cDW.DropDownObject.Refresh
   
End Sub

Public Sub KeyEffect(ByVal iKey As KeyCodeConstants, ByRef bDoIt As Boolean, ByRef x As Single, ByRef y As Single, ByRef eButton As MouseButtonConstants, ByRef bCancel As Boolean, ByRef bSelect As Boolean)
   bDoIt = False
   x = (m_lCurrentX - 1) * 24 + 3
   y = (m_lCurrentY - 1) * 24 + 3
   eButton = vbLeftButton
   bCancel = False
   bSelect = False
   Select Case iKey
   Case vbKeyUp
      bDoIt = True
      y = y - 24
      If (y < 0) Then y = 0
   Case vbKeyDown
      bDoIt = True
      y = y + 24
   Case vbKeyRight
      bDoIt = True
      x = x + 24
   Case vbKeyLeft
      bDoIt = True
      x = x - 24
      If (x < 0) Then x = 0
   Case vbKeyReturn, vbKeySpace
      bDoIt = True
      bSelect = True
   Case vbKeyEscape
      bDoIt = True
      bCancel = True
   End Select
End Sub




