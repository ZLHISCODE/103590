VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTemplateDate 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "对话框标题"
   ClientHeight    =   2670
   ClientLeft      =   2715
   ClientTop       =   3405
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Index           =   0
      Left            =   15
      ScaleHeight     =   2640
      ScaleWidth      =   3570
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   15
      Width           =   3570
      Begin VB.PictureBox picCmd 
         Height          =   315
         Left            =   2370
         ScaleHeight     =   255
         ScaleWidth      =   345
         TabIndex        =   11
         Top             =   2265
         Width           =   405
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   165
         Index           =   0
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2430
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   291
         _Version        =   393216
         BuddyControl    =   "txt(0)"
         BuddyDispid     =   196612
         BuddyIndex      =   0
         OrigLeft        =   105
         OrigTop         =   2460
         OrigRight       =   465
         OrigBottom      =   2670
         Max             =   23
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   165
         Index           =   1
         Left            =   480
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2430
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   291
         _Version        =   393216
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196612
         BuddyIndex      =   1
         OrigLeft        =   735
         OrigTop         =   2460
         OrigRight       =   1095
         OrigBottom      =   2670
         Max             =   59
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   165
         Index           =   2
         Left            =   900
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2430
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   291
         _Version        =   393216
         BuddyControl    =   "txt(2)"
         BuddyDispid     =   196612
         BuddyIndex      =   2
         OrigLeft        =   1350
         OrigTop         =   2460
         OrigRight       =   1710
         OrigBottom      =   2670
         Max             =   59
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   975
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   2235
         Width           =   210
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   540
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "00"
         Top             =   2235
         Width           =   210
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   105
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "00"
         Top             =   2235
         Width           =   210
      End
      Begin MSComCtl2.MonthView MonthView 
         CausesValidation=   0   'False
         Height          =   2160
         Left            =   45
         TabIndex        =   0
         Top             =   45
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483635
         BorderStyle     =   1
         Appearance      =   0
         MonthBackColor  =   12648447
         ShowToday       =   0   'False
         StartOfWeek     =   96403458
         TitleBackColor  =   -2147483635
         TrailingForeColor=   -2147483635
         CurrentDate     =   39874
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "："
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   2220
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "："
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   9
         Top             =   2220
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   735
      TabIndex        =   7
      Top             =   930
      Width           =   1100
   End
End
Attribute VB_Name = "frmTemplateDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'（１）窗体级变量定义
'######################################################################################################################

Private Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum

Private mblnStartUp As Boolean
Private mstrStatePath As String
Private mlngX As Long
Private mlngY As Long
Private mstrDate As String

Private msglTxtH As Single



Private mblnOK As Boolean
Private mlngTXTProc As Long


Private mfrmMain As Object
Private mstrFilterControl As String
Private mobjStateInfo As CommandBarControl

'######################################################################################################################
'

Private Function GetTrayHeight() As Long
    '******************************************************************************************************************
    '功能:获取任务栏的高度
    '******************************************************************************************************************
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function


Private Sub RestoreFormState()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error Resume Next
          
    '检查是否超过屏幕高和宽度
    Dim lngTrayH As Long
    Dim lngH0 As Long
    Dim lngH1 As Long
    
    lngTrayH = GetTrayHeight
    
    If Me.Left + Me.Width > Screen.Width Then
        If (Screen.Width - Me.Width) >= 0 Then
            Me.Left = Screen.Width - Me.Width
        Else
            Me.Left = 0
            Me.Width = Screen.Width
        End If
    End If
    
    If Me.Top + Me.Height > (Screen.Height - lngTrayH) Then
        
        If (Me.Top - Me.Height - msglTxtH) >= 0 Then
            '放在输入框的上面
            Me.Top = Me.Top - Me.Height - msglTxtH
        Else
            
            '分别计算放置上面和放置下面的高度,取最大高度
            lngH0 = Me.Top - msglTxtH
            lngH1 = Screen.Height - lngTrayH - Me.Top
            
            If lngH0 > lngH1 Then
            
                '上面高
                Me.Top = 0
                Me.Height = lngH0
            Else
                Me.Height = Screen.Height - lngTrayH - Me.Top
            End If
        End If
    End If
    
End Sub

Public Function ShowDialog(ByVal x As Single, _
                            ByVal y As Single, _
                            ByRef StrDate As String, _
                            Optional ByVal CtlHeight As Single = 300) As Boolean
    '******************************************************************************************************************
    '功能：显示查询选择器
    '参数：
    '******************************************************************************************************************
    Dim strTmp As String
    
    On Error GoTo errHand
    
    mstrDate = StrDate
    mblnStartUp = True
    mblnOK = False
            
    If IsDate(StrDate) = False And StrDate <> "" Then Exit Function
    
    Me.Left = x + 90
    Me.Top = y + 90
        
    If StrDate = "" Then
        strTmp = Format(CurrentDate, "yyyy-MM-dd HH:mm:ss")
    Else
        strTmp = Format(StrDate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    MonthView.value = CDate(Format(strTmp, "yyyy-MM-dd"))
    txt(0).Text = Mid(strTmp, 12, 2)
    txt(1).Text = Mid(strTmp, 15, 2)
    txt(2).Text = Mid(strTmp, 18, 2)
    
    Call RestoreFormState
    
    Me.Show 1
        
    If mblnOK Then StrDate = mstrDate
        
    ShowDialog = mblnOK
    
    Exit Function
    
errHand:
    
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

'######################################################################################################################

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    If MonthView.Visible Then MonthView.SetFocus
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    Me.Width = MonthView.Width + 90
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 30
    
End Sub

Private Sub MonthView_DateDblClick(ByVal DateDblClicked As Date)
    Call picCmd_Click
End Sub

Private Sub MonthView_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call picCmd_Click
    End If
End Sub

Private Sub picCmd_Click()
    mstrDate = Format(MonthView.value, "yyyy-MM-dd") & " " & Format(Val(txt(0).Text), "00") & ":" & Format(Val(txt(1).Text), "00") & ":" & Format(Val(txt(2).Text), "00")
    mblnOK = True
    Unload Me
End Sub

Private Sub picCmd_Paint()
    PicShowFlat picCmd, 1, "OK", taCenterAlign
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    TxtSelAll txt(Index)

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0

        If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0

    End If
    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetNewWindowLong(txt(Index).hwnd, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call RestoreWindowLong(txt(Index).hwnd)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub
End Sub


Private Sub PicShowFlat(objPic As Object, Optional IntStyle As Integer = -1, Optional strName As String = "", Optional intAlign As mTextAlign)
'功能：将PictureBox模拟成凹下或凸起的面板
'参数：intStyle:0=平面,-1=凹下,1=凸起
'      intAlign=如果要显示文本,则指定对齐方式
    
    Dim vRect As RECT, lngTmp As Long
    
    With objPic
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If IntStyle <> 0 Then
            vRect.Left = .ScaleLeft
            vRect.Top = .ScaleTop
            vRect.Right = .ScaleWidth
            vRect.Bottom = .ScaleHeight
            DrawEdge .hDC, vRect, CLng(IIf(IntStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If intAlign = taCenterAlign Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2 '中间对齐
            ElseIf intAlign = taRightAlign Then
                .CurrentX = .ScaleWidth - .TextWidth(strName) - 2 '右边对齐
            Else
                .CurrentX = 2 '左边对齐
            End If
            objPic.Print strName
        End If
    End With
End Sub


Private Sub TxtSelAll(objTxt As Object)
'功能：将编辑框的的文本全部选中
'参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hwnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

Private Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Private Function SetNewWindowLong(ByVal lngHwnd As Long, ByVal dwNewLong As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngTXTProc = GetWindowLong(lngHwnd, GWL_WNDPROC)
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, dwNewLong)
        
    SetNewWindowLong = True
    
End Function

Private Function RestoreWindowLong(ByVal lngHwnd As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, mlngTXTProc)
    
    RestoreWindowLong = True
End Function
