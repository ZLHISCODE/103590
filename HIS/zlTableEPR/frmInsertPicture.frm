VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInsertPicture 
   Caption         =   "插入图片"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   Icon            =   "frmInsertPicture.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   10050
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picOrig 
      Height          =   5655
      Left            =   3495
      ScaleHeight     =   5595
      ScaleWidth      =   5640
      TabIndex        =   4
      Top             =   330
      Width           =   5700
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   720
      ScaleHeight     =   5910
      ScaleWidth      =   2265
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   315
      Width           =   2265
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   225
         TabIndex        =   1
         Top             =   135
         Width           =   1725
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   4755
         Left            =   180
         TabIndex        =   2
         Top             =   705
         Width           =   2055
         _cx             =   3625
         _cy             =   8387
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInsertPicture.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Shape shpSearch 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   90
         Top             =   45
         Width           =   330
      End
      Begin VB.Shape shpThis 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   0
         Top             =   810
         Width           =   330
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6315
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10874
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   45
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmInsertPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ID_Marked = 301                       '标记图
Private Const ID_Out = 302                          '外部图
Private Const ID_InsAndExit = 303                   '插入并退出
Private Const ID_CancelAndExit = 304                '取消并退出
Private mlDesPicWidth As Long, mlDesPicHeight As Long, mPicReturn As StdPicture, mstrFileType As String
Public Function ShowMe(ByVal frmParent As Object, ByVal lDesPicWidth As Long, ByVal lDesPicHeight As Long, picReturn As StdPicture) As Boolean
    mlDesPicWidth = lDesPicWidth: mlDesPicHeight = lDesPicHeight: mstrFileType = "Marked"
    Set mPicReturn = New StdPicture
    stbThis.Panels(2).Text = "目标位置:宽度" & CInt(Me.ScaleX(mlDesPicWidth, vbTwips, vbPixels)) & " × 高度 " & CInt(Me.ScaleY(mlDesPicHeight, vbTwips, vbPixels))

    Me.Show vbModal, frmParent
    Set picReturn = mPicReturn
    If picReturn.Handle <> 0 Then ShowMe = True
End Function

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ID_CancelAndExit
            Set mPicReturn = New StdPicture
            Unload Me
        Case ID_InsAndExit
            Unload Me
        Case ID_Marked
            If picLeft.Visible = False Then
                picLeft.Visible = True
                mstrFileType = "Marked"
            Else
                picLeft.Visible = False
                mstrFileType = "Out"
            End If
            CommandBars_Resize
        Case ID_Out
            Call InsertLocalPic
    End Select
End Sub

Private Sub CommandBars_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
Bottom = Me.stbThis.Height
End Sub

Private Sub CommandBars_Resize()
Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    On Error Resume Next
    CommandBars.GetClientRect Left, Top, Right, Bottom
    If mstrFileType = "Marked" Then
        stbThis.Panels(1).Text = "【标记图】"
    Else
        stbThis.Panels(1).Text = "【本地图】"
    End If
    
    Dim lX As Long, lY As Long
    lX = Screen.TwipsPerPixelX
    lY = Screen.TwipsPerPixelY
    If Right >= Left And Bottom >= Top Then
        If picLeft.Visible Then
            picLeft.Move Left + lX * 2, Top + lY * 2, picLeft.Width, (Bottom - Top) - lY * 4
            picOrig.Move picLeft.Left + picLeft.Width + lX * 2, picLeft.Top, (Right - Left) - picLeft.Width - lX * 4, picLeft.Height
        Else
            picOrig.Move Left, Top, (Right - Left), (Bottom - Top)
        End If
    End If
End Sub

Private Sub Form_Load()
Dim objControl As CommandBarControl                 '工具栏控件
Dim BarTool As CommandBar
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBars.Icons = zlCommFun.GetPubIcons
    CommandBars.ActiveMenuBar.Visible = False
    CommandBars.EnableCustomization (False)
    CommandBars.Options.UseDisabledIcons = True
    CommandBars.Options.AlwaysShowFullMenus = True
    
    Set BarTool = CommandBars.Add("常用", xtpBarTop)
    With BarTool.Controls
        Set objControl = .Add(xtpControlButton, ID_Marked, "标记图(&M)"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 3552
        Set objControl = .Add(xtpControlButton, ID_Out, "本地图(&W)"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, ID_InsAndExit, "插入(&S)"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 3091
        Set objControl = .Add(xtpControlButton, ID_CancelAndExit, "关闭(&Q)"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 191
    End With

    '热键绑定
    CommandBars.KeyBindings.Add FCONTROL, vbKeyM, ID_Marked
    CommandBars.KeyBindings.Add FCONTROL, vbKeyW, ID_Out
    CommandBars.KeyBindings.Add FCONTROL, vbKeyS, ID_InsAndExit
    CommandBars.KeyBindings.Add FCONTROL, vbKeyQ, ID_CancelAndExit
    CommandBars.KeyBindings.Add FCONTROL, vbKeyReturn, ID_InsAndExit
    CommandBars.KeyBindings.Add 0, VK_ESCAPE, ID_CancelAndExit
    
    Call FillGrid
    Call RestoreWinState(Me, App.ProductName)
End Sub

'################################################################################################################
'## 功能：  填充标记图图片列表
'################################################################################################################
Private Sub FillGrid()
Dim rsTemp As ADODB.Recordset
    gstrSQL = "select 编码,名称,简码 from 病历标记图形"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "")
    vfgThis.Clear
    If Not rsTemp.EOF Then vfgThis.Rows = rsTemp.RecordCount + 1
    Dim i As Long
    i = 0
    vfgThis.Cell(flexcpText, 0, 0) = "编码"
    vfgThis.Cell(flexcpText, 0, 1) = "简码"
    vfgThis.Cell(flexcpText, 0, 2) = "名称"
    vfgThis.ColAlignment(1) = flexAlignLeftCenter
    Do While Not rsTemp.EOF
        i = i + 1
        vfgThis.Cell(flexcpText, i, 0) = Nvl(rsTemp("编码"))
        vfgThis.Cell(flexcpText, i, 1) = Nvl(rsTemp("简码"))
        vfgThis.Cell(flexcpText, i, 2) = Nvl(rsTemp("名称"))
        rsTemp.MoveNext
    Loop
    rsTemp.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Set mPicReturn = New StdPicture
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picLeft_Resize()
    Dim lX As Long, lY As Long
    lX = Screen.TwipsPerPixelX
    lY = Screen.TwipsPerPixelY
    With picLeft
        txtSearch.Move lX, lY, .ScaleWidth - lX * 2
        shpSearch.Move txtSearch.Left - lX, txtSearch.Top - lY, txtSearch.Width + lX * 2, txtSearch.Height + lY * 2
        vfgThis.Move txtSearch.Left, shpSearch.Top + shpSearch.Height + lY, txtSearch.Width, .ScaleHeight - shpSearch.Height - lY * 2
        shpThis.Move vfgThis.Left - lX, vfgThis.Top - lY, vfgThis.Width + 2 * lX, vfgThis.Height + 2 * lY
    End With
End Sub

Private Sub picOrig_DblClick()
    Unload Me
End Sub

Private Sub picOrig_Resize()
    Call DrawCenterPicture
End Sub

Private Sub txtSearch_GotFocus()
    zlControl.TxtSelAll txtSearch
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim i As Long
        For i = 1 To vfgThis.Rows - 1
            If UCase(vfgThis.Cell(flexcpText, i, 1)) Like UCase(Trim(txtSearch)) & "*" Or UCase(vfgThis.Cell(flexcpText, i, 2)) Like UCase(Trim(txtSearch)) & "*" Then
                vfgThis.Row = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub vfgThis_DblClick()
    If vfgThis.Row = 0 Then Exit Sub
    Unload Me
End Sub

Private Sub vfgThis_RowColChange()
    If vfgThis.Row = 0 Then Exit Sub
    ShowPicture vfgThis.Cell(flexcpText, vfgThis.Row, 0)
End Sub
Private Sub ShowPicture(ByVal strKey As String)
Dim strTemp As String
    On Error GoTo errHand:
    Screen.MousePointer = vbHourglass
    strTemp = zlBlobRead(0, strKey)
    If gobjFSO.FileExists(strTemp) Then
        Set mPicReturn = LoadPicture(strTemp)
        gobjFSO.DeleteFile strTemp
        Call DrawCenterPicture
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub DrawCenterPicture()
    Set picOrig.Picture = New StdPicture
    If mPicReturn.Handle = 0 Then Exit Sub
    picOrig.AutoRedraw = True
    If mlDesPicWidth >= picOrig.Width Or mlDesPicHeight >= picOrig.Height Then
        Call picOrig.PaintPicture(mPicReturn, 0, 0)
    Else
        Call picOrig.PaintPicture(mPicReturn, (picOrig.Width - mlDesPicWidth) / 2, (picOrig.Height - mlDesPicHeight) / 2, mlDesPicWidth, mlDesPicHeight)
    End If
    stbThis.Panels(2).Text = "目标位置:宽度 " & CInt(Me.ScaleX(mlDesPicWidth, vbTwips, vbPixels)) & " × 高度 " & CInt(Me.ScaleY(mlDesPicHeight, vbTwips, vbPixels))
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & " 源图大小:宽度 " & CInt(Me.ScaleX(mPicReturn.Width, vbTwips, vbPixels)) & " × 高度 " & CInt(Me.ScaleY(mPicReturn.Height, vbTwips, vbPixels))
End Sub
Private Sub InsertLocalPic()
Dim strTemp As String
    On Error GoTo errHand
    strTemp = GetOpenFile(Me.hWnd, "*.jpg", "所有图像文件" & Chr(0) & "*.jpg;*.bmp;*.gif;*.png;*.tif" & Chr(0), "插入本地图片")
    If strTemp <> "" Then
        Set mPicReturn = LoadPicture(strTemp)
        Call DrawCenterPicture
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
