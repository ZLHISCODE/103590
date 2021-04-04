VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDicSelect 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "字码选择器"
   ClientHeight    =   4335
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5BE9E&
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   0
      MousePointer    =   5  'Size
      ScaleHeight     =   105
      ScaleWidth      =   4875
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   4875
      Begin VB.Image imgTitle 
         Height          =   45
         Left            =   1350
         MousePointer    =   5  'Size
         Picture         =   "frmDicSelect.frx":0000
         Top             =   30
         Width           =   2250
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Height          =   345
      Left            =   45
      TabIndex        =   2
      Top             =   3945
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1588
            MinWidth        =   882
            Text            =   "编码种类"
            TextSave        =   "编码种类"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6615
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
   Begin VB.TextBox txtFind 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   1755
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   2445
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   2865
      _cx             =   5054
      _cy             =   4313
      Appearance      =   2
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDicSelect.frx":0082
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
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
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.Toolbar cbrThis 
      Height          =   330
      Left            =   1890
      TabIndex        =   4
      Top             =   135
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ils16"
      HotImageList    =   "ils16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "搜索"
            Key             =   "搜索"
            Object.ToolTipText     =   "搜索"
            Object.Tag             =   "搜索"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "选中"
            Key             =   "选中"
            Object.ToolTipText     =   "选中"
            Object.Tag             =   "选中"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "取消"
            Object.ToolTipText     =   "取消"
            Object.Tag             =   "取消"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1800
      Top             =   3285
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDicSelect.frx":0157
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDicSelect.frx":69B9
            Key             =   "find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDicSelect.frx":6D53
            Key             =   "cancel"
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpBorder3 
      BorderColor     =   &H00E09060&
      Height          =   375
      Left            =   900
      Top             =   3555
      Width           =   330
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00E09060&
      Height          =   375
      Left            =   90
      Top             =   3555
      Width           =   330
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00E09060&
      Height          =   375
      Left            =   495
      Top             =   3555
      Width           =   330
   End
End
Attribute VB_Name = "frmDicSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngCount As Long

Public Event pOK(strR As String)
Public Event pCancel()

'## 局部临时变量
Private lngX As Long, lngY As Long
Private mblnModel As Long

'################################################################################################################
'## 功能：  显示字典项目选择器
'##
'## 参数：  strDicName  :查找字符串
'##         (X,Y)       :显示位置（屏幕坐标）
'##         blnModel    :窗体显示模态（0-vbModeless  1-vbModal）
'##         ofrmParent  :父窗体
'################################################################################################################
Public Sub ShowMe(ByVal strDicName As String, ByVal X As Long, ByVal Y As Long, _
    Optional ByVal blnModel As Long = vbModeless, _
    Optional ByRef ofrmParent As Object, _
    Optional ByVal strFind As String)

    mblnModel = blnModel
    Me.stbThis.Panels(1).Text = strDicName
    
    Err = 0: On Error GoTo errHand
    Me.Left = X
    Me.Top = Y
    Me.Width = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmDicSelect", "MainWidth", 4000)
    txtFind.Text = ""
    vgdList.Rows = 1
    Me.txtFind = Trim(strFind)
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000: If Me.txtFind.Visible And Me.txtFind.Enabled Then Me.txtFind.SetFocus
    On Error Resume Next
    Call Form_Resize
    Me.Show blnModel, ofrmParent
    Call DoFind
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DoFind()
Dim lngMatch As Long
Dim rsTemp As New ADODB.Recordset

    lngMatch = Val(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0))
    If Trim(Me.txtFind.Text) = "" Then Exit Sub
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 编码, 名称, 简码 From The (Select Cast(Zl_Dic_Search([1], [2], " & lngMatch & ") As " & gstrDbOwner & ".t_Dic_Rowset) From Dual)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmDicSelect", Me.stbThis.Panels(1).Text, Trim(Me.txtFind.Text))
    Set Me.vgdList.DataSource = rsTemp
    With Me.vgdList
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
    End With
    If rsTemp.RecordCount = 0 Then
        cbrThis.Buttons(2).Enabled = False
        Me.stbThis.Panels(2).Text = "没有匹配的项目"
        Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000: If Me.txtFind.Visible And Me.txtFind.Enabled Then Me.txtFind.SetFocus
    Else
        cbrThis.Buttons(2).Enabled = True
        Me.stbThis.Panels(2).Text = "请选择希望的项目"
        If Me.vgdList.Visible And Me.vgdList.Enabled Then Me.vgdList.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call DoFind
    Case 2
        vgdList_DblClick
    Case 3
        Form_Deactivate
    End Select
End Sub

Private Sub Form_Activate()
    If Me.txtFind.Visible And Me.txtFind.Enabled Then Me.txtFind.SetFocus
End Sub

Private Sub Form_Deactivate()
    RaiseEvent pCancel
    If mblnModel = vbModal Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Dim lX As Long, lY As Long
    lX = Screen.TwipsPerPixelX
    lY = Screen.TwipsPerPixelY
    picTitle.Move 60, 60, ScaleWidth - 120
    stbThis.Move lX, ScaleHeight - stbThis.Height - lY, ScaleWidth - lX * 2
    Me.cbrThis.Move ScaleWidth - 80 - Me.cbrThis.Width, picTitle.Height + 120
    Me.txtFind.Move 80, Me.cbrThis.Top, Me.cbrThis.Left - Me.txtFind.Left - lX, Me.cbrThis.Height
    Me.vgdList.Move 80, Me.txtFind.Top + Me.txtFind.Height + lX * 4, ScaleWidth - 160, Me.ScaleHeight - Me.stbThis.Height - Me.stbThis.Height - picTitle.Height - 250
    shpBorder1.Move 0, 0, ScaleWidth, ScaleHeight
    shpBorder2.Move vgdList.Left - lX, vgdList.Top - lY, vgdList.Width + lX * 3, vgdList.Height + lY * 2
    shpBorder3.Move shpBorder2.Left, txtFind.Top - lY, shpBorder2.Width, txtFind.Height + lY * 2
    If Me.Top + Me.Height > Screen.Height - 800 Then Me.Top = Me.Top - Me.Height - 200
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Me.Left - Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ils16.ListImages.Clear
    ImageList_Destroy ils16.hImageList
    Set imgTitle.Picture = Nothing
    Set picTitle.Picture = Nothing
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
    Me.stbThis.Panels(2).Text = "输入希望查找项目的编码/名称/简码"
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call DoFind
        Exit Sub
    ElseIf KeyAscii = vbKeyEscape Then
        Form_Deactivate
    End If
End Sub

Private Sub vgdList_DblClick()
    If vgdList.ROW <= 0 Then Exit Sub
    Dim i As Long, strReturn As String
    
    If vgdList.ROW < 1 Then Exit Sub
    strReturn = ""
    For i = 0 To vgdList.Cols - 1
        strReturn = strReturn & ";" & vgdList.TextMatrix(vgdList.ROW, i)
    Next
    If Len(strReturn) > 0 Then strReturn = Mid(strReturn, 2)
    RaiseEvent pOK(strReturn)
    If mblnModel = vbModal Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub vgdList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        vgdList_DblClick
    ElseIf KeyAscii = vbKeyEscape Then
        Form_Deactivate
    End If
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTitle.Tag = "Down"
    lngX = X
    lngY = Y
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picTitle.Tag = "Down" Then
        Me.Move Me.Left + X - lngX, Me.Top + Y - lngY
    Else
        If X > 0 And X < picTitle.ScaleWidth And Y > 0 And Y < picTitle.ScaleHeight Then
            SetCapture picTitle.hwnd
            picTitle.Cls
            picTitle.BackColor = &HC2EEFF
            picTitle.Line (0, 0)-(picTitle.ScaleWidth - Screen.TwipsPerPixelX, picTitle.ScaleHeight - Screen.TwipsPerPixelY), &H800000, B
            Me.stbThis.Panels(2).Text = "按下鼠标拖拽可以移动编辑器"
            If picTitle.Tag = "Down" Then
                Me.Move Me.Left + X - lngX, Me.Top + Y - lngY
            End If
        Else
            ReleaseCapture
            picTitle.Cls
            picTitle.BackColor = &HF5BE9E
            Me.stbThis.Panels(2).Text = "输入希望查找项目的编码/名称/简码"
        End If
    End If
End Sub

Private Sub picTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTitle.Tag = ""
    If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
End Sub

Private Sub picTitle_Resize()
    imgTitle.Move (picTitle.ScaleWidth - imgTitle.Width) / 2, 30
End Sub

