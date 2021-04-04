Attribute VB_Name = "mSetFormIcon"
'#########################################################################
'##模 块 名：mSetFormIcon.bas
'##创 建 人：吴庆伟
'##日    期：2005年4月1日
'##修 改 人：
'##日    期：
'##描    述：显示窗体真彩色图标
'#########################################################################

Option Explicit

'XP界面效果
Public Declare Function InitCommonControls Lib "COMCTL32.DLL" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long
   
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

'################################################################################################################
'## 功能：  设置窗体图标为资源文件中存储的真彩图标
'##
'## 参数：  hWnd            :窗体句柄 hWnd
'##         sIconResName    :图标资源的标识名称
'##         bSetAsAppIcon   :是否设置为程序图标，默认为 True
'################################################################################################################
Public Sub SetIcon( _
      ByVal hWnd As Long, _
      ByVal sIconResName As String, _
      Optional ByVal bSetAsAppIcon As Boolean = True)
Dim lhWndTop As Long
Dim lHWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lHWnd = hWnd
      lhWndTop = lHWnd
      Do While Not (lHWnd = 0)
         lHWnd = GetWindow(lHWnd, GW_OWNER)
         If Not (lHWnd = 0) Then
            lhWndTop = lHWnd
         End If
      Loop
   End If
   
   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cx = GetSystemMetrics(SM_CXSMICON)
   cy = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub
