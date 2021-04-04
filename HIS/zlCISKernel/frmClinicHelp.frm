VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmClinicHelp 
   BackColor       =   &H8000000C&
   Caption         =   "诊疗操作与药物参考查阅"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "frmClinicHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab stbCatalog 
      Height          =   6690
      Left            =   90
      TabIndex        =   7
      Top             =   585
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   11800
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "目录(&C)"
      TabPicture(0)   =   "frmClinicHelp.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "搜索(&S)"
      TabPicture(1)   =   "frmClinicHelp.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblNote"
      Tab(1).Control(1)=   "lvwFind"
      Tab(1).Control(2)=   "cboFind"
      Tab(1).Control(3)=   "cmdFind"
      Tab(1).Control(4)=   "cmdShow(0)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "书签(&M)"
      TabPicture(2)   =   "frmClinicHelp.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwMark"
      Tab(2).Control(1)=   "cmdMark"
      Tab(2).Control(2)=   "cmdDel"
      Tab(2).Control(3)=   "cmdShow(1)"
      Tab(2).ControlCount=   4
      Begin VB.PictureBox picList 
         Height          =   5880
         Left            =   195
         ScaleHeight     =   5820
         ScaleWidth      =   3000
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "100"
         Top             =   465
         Width           =   3060
         Begin VB.CommandButton cmdKind 
            Caption         =   "诊疗操作参考  (&5)"
            Height          =   350
            Index           =   4
            Left            =   0
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   1335
            Width           =   2295
         End
         Begin VB.CommandButton cmdKind 
            Caption         =   "中药配方参考  (&4)"
            Height          =   350
            Index           =   3
            Left            =   0
            TabIndex        =   19
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   1005
            Width           =   2295
         End
         Begin VB.CommandButton cmdKind 
            Caption         =   "中草药应用参考(&3)"
            Height          =   350
            Index           =   2
            Left            =   0
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   675
            Width           =   2295
         End
         Begin VB.CommandButton cmdKind 
            Caption         =   "中成药应用参考(&2)"
            Height          =   350
            Index           =   1
            Left            =   0
            TabIndex        =   21
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   345
            Width           =   2295
         End
         Begin VB.CommandButton cmdKind 
            Caption         =   "西药应用参考  (&1)"
            Height          =   350
            Index           =   0
            Left            =   0
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   15
            Width           =   2295
         End
         Begin MSComctlLib.TreeView tvwList 
            Height          =   4005
            Left            =   30
            TabIndex        =   23
            Tag             =   "1000"
            Top             =   1740
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   7064
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imgList"
            Appearance      =   0
         End
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "显示(&D)"
         Enabled         =   0   'False
         Height          =   350
         Index           =   0
         Left            =   -72630
         TabIndex        =   14
         Top             =   5640
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "列出搜索项目(&L)"
         Height          =   350
         Left            =   -73230
         TabIndex        =   13
         Top             =   1305
         Width           =   1530
      End
      Begin VB.ComboBox cboFind 
         Height          =   300
         Left            =   -74850
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   810
         Width           =   3165
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "显示(&D)"
         Height          =   350
         Index           =   1
         Left            =   -72660
         TabIndex        =   10
         Top             =   5685
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&E)"
         Height          =   350
         Left            =   -72765
         TabIndex        =   9
         Top             =   450
         Width           =   1110
      End
      Begin VB.CommandButton cmdMark 
         Caption         =   "当前项标记为书签(&M)"
         Height          =   350
         Left            =   -74820
         TabIndex        =   8
         Top             =   435
         Width           =   2010
      End
      Begin MSComctlLib.ListView lvwMark 
         Height          =   4605
         Left            =   -74835
         TabIndex        =   11
         Top             =   870
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   8123
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwFind 
         Height          =   3660
         Left            =   -74865
         TabIndex        =   15
         Top             =   1815
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   6456
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "键入要查找的编码、名称、别名或简码(&W):"
         Height          =   180
         Left            =   -74835
         TabIndex        =   16
         Top             =   450
         Width           =   3495
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7890
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClinicHelp.frx":03DE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14949
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
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   4155
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox picItem 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4605
      Picture         =   "frmClinicHelp.frx":0C70
      ScaleHeight     =   675
      ScaleWidth      =   7245
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   645
      Width           =   7245
      Begin VB.Label lblAlias 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "甲氰咪胍泰胃美"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3495
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "西咪替丁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   5940
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3210
      Top             =   7380
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
            Picture         =   "frmClinicHelp.frx":128C
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicHelp.frx":1826
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicHelp.frx":1DC0
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   9330
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRefer 
      Height          =   5295
      Left            =   4665
      TabIndex        =   3
      Top             =   1860
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   -2147483628
      Rows            =   5
      Cols            =   4
      FixedRows       =   0
      BackColorFixed  =   -2147483628
      ForeColorFixed  =   32768
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   225
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "比例尺寸"
      Height          =   180
      Left            =   7920
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmClinicHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSteps As String    '后退和前进项目记录步骤

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String, aryTemp() As String
Dim mstrSQL As String
Dim mstrLike As String

Private WithEvents objParentForm As Form
Attribute objParentForm.VB_VarHelpID = -1

Public Sub ShowMe(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long, Optional ByVal blnShowInTaskBar As Boolean)
    '---------------------------------------------
    '功能：根据上级程序要求，以模态或非模态显示诊疗参考
    '入参：frmParent-父窗体；
    '      blnModal-是否模态显示（通常和上级窗体一致）；
    '      lngItemId-要显示的诊疗项目ID，不为0时，缺省不显示目录区；
    '---------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, blnShowInTaskBar)
    
    If lngItemID = 0 Then
        Me.stbCatalog.Visible = True: Me.picVBar.Visible = True
        Call cmdKind_Click(0)
        Call stbCatalog_Click(0)
    Else
        mstrSQL = "Select nvl(max(参考目录ID),0) from 诊疗项目目录 where ID=[1]"
        err = 0: On Error GoTo Errhand
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngItemID)
        lngItemID = rsTemp.Fields(0).Value
        If lngItemID = 0 Then
            If MsgBox("项目不存在对应参考，是否自行查找参考？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Unload Me: Exit Sub
            Me.stbCatalog.Visible = True: Me.picVBar.Visible = True
            Call cmdKind_Click(0)
            Call stbCatalog_Click(0)
        Else
            Me.stbCatalog.Visible = False: Me.picVBar.Visible = False
            Call zlShowRef(lngItemID)
        End If
    End If
    
    err = 0: On Error Resume Next
    Set objParentForm = frmParent
    Me.Show bytModal, frmParent
    Exit Sub

Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboFind_Click()
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cboFind_GotFocus()
    Me.cboFind.SelStart = 0: Me.cboFind.SelLength = 100
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Hide
        Me.stbCatalog.Visible = False: Me.picVBar.Visible = False
        If Me.WindowState <> 2 Then
            Me.Left = Me.Left + (Me.stbCatalog.Width + Me.picVBar.Width)
            Me.Width = Me.Width - (Me.stbCatalog.Width + Me.picVBar.Width)
        End If
        Call cbsThis_Resize
    Case conMenu_View_Show
        Me.stbCatalog.Visible = True: Me.picVBar.Visible = True
        If Me.WindowState <> 2 Then
            If Me.Left - (Me.stbCatalog.Width + Me.picVBar.Width) >= 0 Then
                Me.Left = Me.Left - (Me.stbCatalog.Width + Me.picVBar.Width)
            Else
                Me.Left = 0
            End If
            Me.Width = Me.Width + (Me.stbCatalog.Width + Me.picVBar.Width)
        End If
        
        '定位到当前显示项目所在的目录
        On Error Resume Next
        Call cmdKind_Click(Me.picItem.Tag - 1)
        Call stbCatalog_Click(0)
        Call cbsThis_Resize
    Case conMenu_View_Backward
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        If intCount > LBound(aryTemp) Then Call zlShowRef(Val(aryTemp(intCount - 1)))
    Case conMenu_View_Forward
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        If intCount < UBound(aryTemp) Then Call zlShowRef(Val(aryTemp(intCount + 1)))
    Case conMenu_File_Print
        Dim objPrint As New zlPrint1Grd
        Dim objRow As New zlTabAppRow
        Dim bytMode As Byte
        If Me.hgdRefer.TextMatrix(Me.hgdRefer.FixedRows, 2) = "" Then Exit Sub
        Set objPrint.Body = Me.hgdRefer
        With objPrint.Title
            .Text = Me.lblItem.Caption & ".应用参考"
            .Font.Size = 11
        End With
        objRow.Add Me.lblAlias.Caption
        objPrint.UnderAppRows.Add objRow
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Case conMenu_Help_Help
        ShowHelp App.ProductName, Me.hWnd, Me.Name ', Int((glngSys) / 100)
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    err = 0: On Error Resume Next
    If Me.picVBar.Visible = True Then
        With Me.picVBar
            .Top = lngScaleTop: .Height = lngScaleBottom
            If .Left < 2000 Then .Left = 2000
            If .Left > lngScaleRight - 4000 Then .Left = lngScaleRight - 4000
        End With
        With Me.stbCatalog
            .Left = lngScaleLeft: .Width = Me.picVBar.Left - .Left + 15
            .Top = lngScaleTop: .Height = lngScaleBottom - .Top + 15
        End With
        With Me.picList
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = 390: .Height = Me.stbCatalog.Height - .Top - 90
        End With
        With Me.lblNote
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = 450
        End With
        With Me.cboFind
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = Me.lblNote.Top + Me.lblNote.Height + 45
        End With
        With Me.cmdFind
            .Left = Me.stbCatalog.Width - .Width - 90
            .Top = Me.cboFind.Top + Me.cboFind.Height + 90
        End With
        With Me.cmdShow(0)
            .Left = Me.stbCatalog.Width - .Width - 90
            .Top = Me.stbCatalog.Height - 180 - .Height
        End With
        With Me.lvwFind
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = Me.cmdFind.Top + Me.cmdFind.Height + 90
            .Height = Me.cmdShow(0).Top - .Top - 90
        End With
        
        With Me.cmdMark
            .Left = 90: .Top = 450
        End With
        With Me.cmdDel
            .Left = Me.cmdMark.Left + Me.cmdMark.Width + 45: .Top = 450
        End With
        With Me.cmdShow(1)
            .Left = Me.stbCatalog.Width - .Width - 90
            .Top = Me.stbCatalog.Height - 180 - .Height
        End With
        With Me.lvwMark
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = Me.cmdMark.Top + Me.cmdMark.Height + 90
            .Height = Me.cmdShow(1).Top - .Top - 90
        End With
        
        With Me.picItem
            .Left = Me.picVBar.Left + Me.picVBar.Width: .Width = lngScaleRight - .Left
            .Top = lngScaleTop
        End With
    Else
        With Me.picItem
            .Left = lngScaleLeft: .Width = lngScaleRight - .Left
            .Top = lngScaleTop
        End With
    End If
    
    With Me.lblItem
        .Left = 45: .Width = Me.picItem.ScaleWidth
    End With
    With Me.lblAlias
        .Left = 45: .Width = Me.picItem.ScaleWidth - 90
    End With
    Me.picItem.Height = Me.lblAlias.Top + Me.lblAlias.Height + 90
    
    With Me.hgdRefer
        .Redraw = False
        .Left = Me.picItem.Left: .Width = lngScaleRight - .Left
        .Top = Me.picItem.Top + Me.picItem.Height: .Height = lngScaleBottom - .Top
        .ColWidth(0) = 250
        .ColWidth(1) = Me.TextWidth("空格")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
         Call zlGrdRowHeight
        .Redraw = True
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Hide: Control.Enabled = Me.picVBar.Visible
    Case conMenu_View_Show: Control.Enabled = Not Me.picVBar.Visible
    Case conMenu_View_Backward
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        If intCount > LBound(aryTemp) Then
            Control.Enabled = True
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_Forward
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        If intCount < UBound(aryTemp) Then
            Control.Enabled = True
        Else
            Control.Enabled = False
        End If
    End Select
End Sub

Private Sub cmdDel_Click()
    If Me.lvwMark.SelectedItem Is Nothing Then Exit Sub
    Me.lvwMark.ListItems.Remove (Me.lvwMark.SelectedItem.Key)
    If Me.lvwMark.ListItems.Count > 0 Then
        Me.cmdDel.Enabled = True: Me.cmdShow(1).Enabled = True
    Else
        Me.cmdDel.Enabled = False: Me.cmdShow(1).Enabled = False
    End If
    
    strTemp = ""
    For Each objItem In Me.lvwMark.ListItems
        strTemp = strTemp & "," & Mid(objItem.Key, 2)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    Call zlDatabase.SetPara("诊疗参考书签", strTemp, glngSys, p药品诊疗参考)
End Sub

Private Sub cmdFind_Click()
    If Trim(Me.cboFind.Text) = "" Then
        MsgBox "请输入查找的内容", vbExclamation, gstrSysName
        Me.cboFind.SetFocus: Exit Sub
    End If
    strTemp = ""
    For intCount = 0 To Me.cboFind.ListCount
        strTemp = strTemp & ";" & Me.cboFind.List(intCount)
    Next
    If InStr(1, strTemp, ";" & Trim(Me.cboFind.Text)) = 0 Then
        Me.cboFind.AddItem Trim(Me.cboFind.Text), 0
    End If
    
    Me.lvwFind.ListItems.Clear
    mstrSQL = "select distinct I.分类ID,I.ID,I.编码,I.名称" & _
            " from 诊疗参考目录 I,诊疗参考别名 N" & _
            " where I.ID=N.参考目录ID" & _
            "       and (I.编码 like [1] or N.名称 like [2] or N.简码 like [2])"
    err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Trim(Me.cboFind.Text) & "%", mstrLike & Trim(Me.cboFind.Text) & "%")
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwFind.ListItems.Add(, "_" & !ID, !名称, "item", "item")
            objItem.SubItems(Me.lvwFind.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
    End With
    If Me.lvwFind.ListItems.Count > 0 Then
        Me.lvwFind.ListItems(1).Selected = True
        Me.lvwFind.SelectedItem.EnsureVisible
        Me.cmdShow(0).Enabled = True
    Else
        Me.cmdShow(0).Enabled = False
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdKind_Click(Index As Integer)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    If Me.stbCatalog.Visible And Me.tvwList.Visible Then Me.tvwList.SetFocus
    If Val(Me.tvwList.Tag) <> Index Then
        Call picList_Resize
        Me.tvwList.Tag = Index
        Call zlRefList
    End If
End Sub

Private Sub cmdMark_Click()
    If Val(Me.lblItem.Tag) = 0 Then Exit Sub
    mstrSQL = "select I.分类ID,I.ID,I.编码,I.名称" & _
            " from 诊疗参考目录 I" & _
            " where I.ID=[1]"
    err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, CLng(Me.lblItem.Tag))
    err = 0: On Error Resume Next
    With rsTemp
        If Not .EOF Then
            strTemp = ""
            For Each objItem In Me.lvwMark.ListItems
                strTemp = strTemp & "," & Mid(objItem.Key, 2)
            Next
            If Len(Mid(strTemp & "," & !ID, 2)) > 1000 Then
                MsgBox "你的书签项目太多，已不能再增加。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            Set objItem = Me.lvwMark.ListItems.Add(, "_" & !ID, !名称, "item", "item")
            objItem.SubItems(Me.lvwMark.ColumnHeaders("编码").Index - 1) = !编码
            Me.lvwMark.ListItems("_" & Me.lblItem.Tag).Selected = True
            Me.lvwMark.SelectedItem.EnsureVisible
            
            strTemp = strTemp & "," & Mid(objItem.Key, 2)
        Else
            Exit Sub
        End If
    End With
    
    If Me.lvwMark.ListItems.Count > 0 Then
        Me.cmdDel.Enabled = True: Me.cmdShow(1).Enabled = True
    Else
        Me.cmdDel.Enabled = False: Me.cmdShow(1).Enabled = False
    End If
    
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    Call zlDatabase.SetPara("诊疗参考书签", strTemp, glngSys, p药品诊疗参考)
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdShow_Click(Index As Integer)
    If Index = 0 Then
        If Me.lvwFind.SelectedItem Is Nothing Then Exit Sub
        Call zlShowRef(Mid(Me.lvwFind.SelectedItem.Key, 2))
    Else
        If Me.lvwMark.SelectedItem Is Nothing Then Exit Sub
        Call zlShowRef(Mid(Me.lvwMark.SelectedItem.Key, 2))
    End If
End Sub

Private Sub Form_Load()
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    Me.cbsThis.ActiveMenuBar.Visible = False
    
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Hide, "隐藏")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "显示")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "后退"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "前进")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("H"), conMenu_View_Hide
        .Add FCONTROL, Asc("S"), conMenu_View_Show
        .Add FCONTROL, Asc("B"), conMenu_View_Backward
        .Add FCONTROL, Asc("F"), conMenu_View_Forward
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    
    '截面元素形态设置
    Me.lvwFind.ListItems.Clear
    With Me.lvwFind.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 3200
        .Add , "编码", "编码", 1000
    End With
    With Me.lvwFind
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1: .SortOrder = lvwAscending
    End With
    Me.lvwMark.ListItems.Clear
    With Me.lvwMark.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 3200
        .Add , "编码", "编码", 1000
    End With
    With Me.lvwMark
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1: .SortOrder = lvwAscending
    End With
    
    With Me.hgdRefer
        .ColWidth(0) = 250
        .ColWidth(1) = Me.TextWidth("空格")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
    End With
        
    '提取书签显示
    strTemp = zlDatabase.GetPara("诊疗参考书签", glngSys, p药品诊疗参考)
    If strTemp = "" Then Exit Sub
        
    mstrSQL = "select I.分类ID,I.ID,I.编码,I.名称" & _
            " from 诊疗参考目录 I" & _
            " where I.ID in (" & strTemp & ")"
    err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwMark.ListItems.Add(, "_" & !ID, !名称, "item", "item")
            objItem.SubItems(Me.lvwMark.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
    End With
    If Me.lvwMark.ListItems.Count > 0 Then
        Me.lvwMark.ListItems(1).Selected = True
        Me.lvwMark.SelectedItem.EnsureVisible
        Me.cmdDel.Enabled = True: Me.cmdShow(1).Enabled = True
    Else
        Me.cmdDel.Enabled = False: Me.cmdShow(1).Enabled = False
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub hgdRefer_DblClick()
    With Me.hgdRefer
        If .TextMatrix(.Row, 0) <> "Δ" And .TextMatrix(.Row, 0) <> "▼" Then Exit Sub
        .Redraw = False
        For intCount = .Row + 1 To .Rows - 1
            If .TextMatrix(intCount, 0) = "Δ" Or .TextMatrix(intCount, 0) = "▼" Then Exit For
            If .TextMatrix(.Row, 0) = "▼" Then
                .RowHeight(intCount) = 255
            Else
                .RowHeight(intCount) = 0
            End If
        Next
        If .TextMatrix(.Row, 0) = "▼" Then
            .TextMatrix(.Row, 0) = "Δ"
        Else
            .TextMatrix(.Row, 0) = "▼"
        End If
        Call zlGrdRowHeight
        .Redraw = True
    End With
End Sub

Private Sub hgdRefer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call hgdRefer_DblClick
End Sub

Private Sub lvwFind_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwFind.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwFind.SortOrder = IIF(Me.lvwFind.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwFind.SortKey = ColumnHeader.Index - 1
        Me.lvwFind.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwFind_DblClick()
    If Me.lvwFind.SelectedItem Is Nothing Then Exit Sub
    Call zlShowRef(Mid(Me.lvwFind.SelectedItem.Key, 2))
End Sub

Private Sub lvwMark_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwMark.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwMark.SortOrder = IIF(Me.lvwMark.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwMark.SortKey = ColumnHeader.Index - 1
        Me.lvwMark.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMark_DblClick()
    If Me.lvwMark.SelectedItem Is Nothing Then Exit Sub
    Call zlShowRef(Mid(Me.lvwMark.SelectedItem.Key, 2))
End Sub

Private Sub objParentForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub picList_Resize()
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picList.ScaleLeft + 0
        Me.cmdKind(intCount).Width = Me.picList.ScaleWidth + 15
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picList.ScaleTop + 285 * intCount
            Me.tvwList.Top = Me.picList.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picList.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.tvwList.Left = Me.picList.ScaleLeft + 15
    Me.tvwList.Width = Me.picList.ScaleWidth
    Me.tvwList.Height = Me.picList.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + X
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call cbsThis_Resize
End Sub

Private Sub stbCatalog_Click(PreviousTab As Integer)
    Select Case Me.stbCatalog.Tab
    Case 0
        Me.picList.Visible = True
        Me.lblNote.Visible = False
        Me.cboFind.Visible = False
        Me.cmdFind.Visible = False
        Me.lvwFind.Visible = False
        Me.cmdShow(0).Visible = False
        Me.cmdMark.Visible = False
        Me.cmdDel.Visible = False
        Me.lvwMark.Visible = False
        Me.cmdShow(1).Visible = False
    Case 1
        Me.picList.Visible = False
        Me.lblNote.Visible = True
        Me.cboFind.Visible = True
        Me.cmdFind.Visible = True
        Me.lvwFind.Visible = True
        Me.cmdShow(0).Visible = True
        Me.cmdMark.Visible = False
        Me.cmdDel.Visible = False
        Me.lvwMark.Visible = False
        Me.cmdShow(1).Visible = False
    Case 2
        Me.picList.Visible = False
        Me.lblNote.Visible = False
        Me.cboFind.Visible = False
        Me.cmdFind.Visible = False
        Me.lvwFind.Visible = False
        Me.cmdShow(0).Visible = False
        Me.cmdMark.Visible = True
        Me.cmdDel.Visible = True
        Me.lvwMark.Visible = True
        Me.cmdShow(1).Visible = True
    End Select
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    '明细项目，直接调用显示参考
    If Left(Node.Key, 1) = "_" Then Call zlShowRef(Mid(Node.Key, 2)): Exit Sub
        
    '分类项目，检查是否装入项目进行操作
    If Left(Node.Key, 1) = "C" And Node.Children > 0 Then Exit Sub
    
    mstrSQL = "select I.分类ID,I.ID,I.编码,I.名称" & _
            " from 诊疗参考目录 I" & _
            " where 分类ID=[1]" & _
            " order by I.编码"
    err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, CLng(Mid(Node.Key, 2)))
    With rsTemp
        Do While Not .EOF
            Set objNode = Me.tvwList.Nodes.Add("C" & !分类ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "item")
            .MoveNext
        Loop
        
        err = 0: On Error Resume Next
        If Val(Me.lblItem.Tag) <> 0 Then
            Me.tvwList.Nodes("_" & Val(Me.lblItem.Tag)).Selected = True
        End If
        If Node.Key <> Me.tvwList.SelectedItem.Key Then
            Node.Expanded = True
            Me.tvwList.SelectedItem.EnsureVisible
            Call zlShowRef(Mid(Me.tvwList.SelectedItem.Key, 2))
        End If
    End With
    Exit Sub

Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRefList()
    '---------------------------------------------
    '填写诊疗参考分类与目录
    '---------------------------------------------
    Me.tvwList.Visible = False
    Me.tvwList.Nodes.Clear
    
    '填写分类
    strTemp = Val(Me.tvwList.Tag) + 1
    mstrSQL = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗参考分类" & _
            " Where 类型 = [1]" & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, strTemp)
    With rsTemp
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwList.Nodes.Add(, , "C" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwList.Nodes.Add("C" & !上级ID, tvwChild, "C" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        
        err = 0: On Error Resume Next
        If Val(Me.picVBar.Tag) <> 0 Then
            Me.tvwList.Nodes("C" & Val(Me.picVBar.Tag)).Selected = True
        End If
        If Not (Me.tvwList.SelectedItem Is Nothing) Then
            Me.tvwList.SelectedItem.EnsureVisible
            Call tvwList_NodeClick(Me.tvwList.SelectedItem)
        End If
    End With
    
    Me.tvwList.Visible = True
    Exit Sub

Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlShowRef(lngItemID As Long)
    '------------------------------------------------
    '功能：显示指定项目的诊疗参考内容
    '------------------------------------------------
    '清理详细信息显示区
    With Me.hgdRefer
        .Rows = 1: .ColAlignment(0) = 1: .ColAlignment(1) = 1: .ColAlignment(2) = 1
        For intCount = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCount) = ""
        Next
    End With
    
    '------------------------------------------------
    Me.hgdRefer.Redraw = False
    mstrSQL = "select distinct I.分类ID,I.ID,I.编码,N.性质,N.名称,I.类型" & _
            " from 诊疗参考目录 I,诊疗参考别名 N" & _
            " where I.ID=N.参考目录ID and I.ID=[1]"
    err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngItemID)
    With rsTemp
        Me.lblAlias.Caption = ""
        Do While Not .EOF
            If !性质 = 1 Then
                Me.picVBar.Tag = !分类ID: Me.picItem.Tag = !类型: Me.lblItem.Tag = !ID: Me.lblItem.Caption = IIF(IsNull(!名称), "", !名称)
            Else
                Me.lblAlias.Caption = Me.lblAlias.Caption & "、" & IIF(IsNull(!名称), "", !名称)
            End If
            .MoveNext
        Loop
        If Me.lblAlias.Caption <> "" Then Me.lblAlias.Caption = Mid(Me.lblAlias.Caption, 2)
    End With
    
    mstrSQL = "select 项目层次,项目序号,内容行号,decode(nvl(内容行号,0),0,参考项目,内容文本) as 内容" & _
            " from 诊疗参考内容" & _
            " where 参考目录id=[1] and (nvl(内容行号,0)=0 or length(ltrim(nvl(内容文本,'')))<>0)" & _
            " order by 项目序号,内容行号"
    err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngItemID)
    With rsTemp
        If .EOF Or .BOF Then
            Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + 1
        Else
            Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdRefer.RowHeight(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = 255
            Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 0) = ""
            If !内容行号 = 0 Then
                If !项目层次 = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 0) = "Δ"
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = "【" & !内容 & "】"
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "【" & !内容 & "】"
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 0) = ""
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "［" & !内容 & "］"
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = False
                End If
            Else
                If !项目层次 = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = Space(4) & !内容
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !内容
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !内容
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = False
                End If
            End If
            .MoveNext
        Loop
    End With
    Call cbsThis_Resize
    Me.hgdRefer.Redraw = True
    
    If mstrSteps = "" Then
        mstrSteps = Val(Me.lblItem.Tag)
    Else
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        If intCount > UBound(aryTemp) Then mstrSteps = mstrSteps & ";" & Val(Me.lblItem.Tag)
    End If
    Exit Sub

Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '根据调整内容调整内容网格的行高度，以保证内容的正常显示
    '---------------------------------------------
    Dim lngColWidth As Long
    With Me.hgdRefer
        For intRow = .FixedRows To .Rows - 1
            If .RowHeight(intRow) <> 0 Then
                If .TextMatrix(intRow, 1) = "" Then
                    lngColWidth = .ColWidth(2)
                Else
                    lngColWidth = .ColWidth(1) + .ColWidth(2)
                End If
                Me.lblScale.Width = lngColWidth - 90
                Me.lblScale.Caption = .TextMatrix(intRow, 2)
                .RowHeight(intRow) = Me.lblScale.Height + 75
            End If
        Next
    End With
End Sub


