VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrTendEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   8565
   Begin VB.PictureBox picPati 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6615
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   1875
      Begin VB.ComboBox cbo病人 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   1845
      End
   End
   Begin VB.PictureBox picSignCheck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   3540
      ScaleHeight     =   2835
      ScaleWidth      =   4725
      TabIndex        =   13
      Top             =   1170
      Visible         =   0   'False
      Width           =   4755
      Begin VB.CommandButton cmdSignAll 
         Caption         =   "全部"
         Height          =   350
         Left            =   270
         TabIndex        =   18
         ToolTipText     =   "确认"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdSignCur 
         Caption         =   "验证"
         Height          =   350
         Left            =   2790
         TabIndex        =   16
         ToolTipText     =   "确认"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         Height          =   350
         Left            =   3690
         TabIndex        =   17
         ToolTipText     =   "取消"
         Top             =   2370
         Width           =   840
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSignData 
         Height          =   1635
         Left            =   -30
         TabIndex        =   15
         Top             =   630
         Width           =   4755
         _cx             =   8387
         _cy             =   2884
         Appearance      =   2
         BorderStyle     =   1
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendEditor.ctx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "以下是签名历史记录，可选择单行验证，也可进行全部验证。"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   810
         TabIndex        =   14
         Top             =   150
         Width           =   3720
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   120
         Picture         =   "usrTendEditor.ctx":0062
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.PictureBox pic护理等级 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1965
      Begin VB.ComboBox cbo护理等级 
         Height          =   300
         Left            =   420
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lbl护理等级 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "模板"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.PictureBox picNothing 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1620
      ScaleHeight     =   405
      ScaleWidth      =   1725
      TabIndex        =   9
      Top             =   60
      Width           =   1725
      Begin VB.Label lblNothing 
         BackStyle       =   0  'Transparent
         Caption         =   "请先选择病人！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   10
         Top             =   90
         Width           =   1875
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   8385
      TabIndex        =   2
      Top             =   510
      Width           =   8385
      Begin MSComctlLib.ListView lvwMultiSel 
         Height          =   1725
         Left            =   2310
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   3043
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   945
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   945
         Begin VB.CommandButton cmd未记说明 
            Caption         =   "‥"
            Height          =   225
            Left            =   630
            TabIndex        =   4
            Top             =   30
            Width           =   255
         End
         Begin VB.ComboBox cbo部位 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txt数据 
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   945
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Vsf 
         Height          =   3975
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8385
         _cx             =   14790
         _cy             =   7011
         Appearance      =   0
         BorderStyle     =   1
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   600
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendEditor.ctx":0CA4
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
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
         Begin VB.PictureBox picSign 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   3600
            ScaleHeight     =   195
            ScaleWidth      =   945
            TabIndex        =   19
            Tag             =   "225"
            Top             =   390
            Visible         =   0   'False
            Width           =   975
            Begin VB.Label lbl验证签名 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "验证签名"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   210
               TabIndex        =   20
               Top             =   0
               Width           =   720
            End
            Begin VB.Image imgSign 
               Height          =   240
               Left            =   -30
               Picture         =   "usrTendEditor.ctx":0D06
               Tag             =   "240"
               Top             =   -30
               Width           =   240
            End
         End
      End
   End
   Begin VB.TextBox txt显示天数 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   5970
      MaxLength       =   2
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   90
      Width           =   645
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   4020
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendEditor.ctx":7558
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "usrTendEditor.ctx":DDBA
      Left            =   690
      Top             =   150
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "usrTendEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public mblnEditable As Boolean

Private objESign As Object
Private mfrmParent As Object
Private mblnInit As Boolean
Private mstrSel As String                   '复制行:1;复制某单元格:1.1
Private mblnShow As Boolean                 '是否显示录入框
Private mblnChange As Boolean               '是否修改数据
Private mintPreDays As Long
Private mstrMaxDate As String
Private mstrSelItems As String              '保存用户本次增加的列，以免刷新后重新设置
Private mblnCheckVersion As Boolean

Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mbyt护理等级 As Byte
Private mint婴儿 As Integer
Private mbln心率 As Boolean                 '是否需要录入心率
Private mstrPrivs As String

Private mlngOper As Long                    '手术列号
Private mlngSigner As Long                  '签名人
Private mlngSignTime As Long                '签名时间
Private mlngRecord As Long                  '记录ID
Private mlngGroup As Long                   '组号
Private mlngCert As Long                    '证书ID
Private mlngCertLevel As Long               '护士/护士长签名
Public mstrPigeonhole As String             '归档人

Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
Private mrsSelItems As New ADODB.Recordset          '当前录入的护理记录项目清单

Private Enum 选择
    列
    行
End Enum

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Public Event AfterDataChanged()
Public Event AfterArchiveChanged()
Public Event AfterRefresh()
Public Event AfterSelChange(ByVal lngCert As Long, ByVal strCertLevel As String)
Public Event DbClick(ByVal strData As String)
Public Event AfterRowColChange(ByVal strInfo As String)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'记录上次选择行,顶行,以便刷新后重新定位
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long
Private mbytFontSize As Byte '字体大小 9、12

Public Sub ReSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Object
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objExtendedBar As CommandBar
    Dim lngCol As Long, lngReDraw As Long
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))

    UserControl.FontSize = mbytFontSize
    UserControl.FontName = "宋体"
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = mbytFontSize
    Set cbsThis.Options.Font = CtlFont
    
    Set CtlFont = dkpMain.PaintManager.CaptionFont
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = mbytFontSize
    Set dkpMain.PaintManager.CaptionFont = CtlFont
    
        '显示天数工具栏
    '------------------------------------------------------------------------------------------------------------------
    lbl护理等级.FontSize = mbytFontSize
    cbo护理等级.FontSize = mbytFontSize
    cbo护理等级.Left = lbl护理等级.Left + lbl护理等级.Width + 30
    lbl护理等级.Top = cbo护理等级.Top + (cbo护理等级.Height - lbl护理等级.Height) \ 2
    cbo护理等级.Width = 1575 + IIf(mbytFontSize = 12, 360, 0)
    pic护理等级.Width = cbo护理等级.Width + cbo护理等级.Left
    pic护理等级.Height = cbo护理等级.Top * 2 + cbo护理等级.Height
    txt显示天数.FontSize = mbytFontSize
    cbo病人.FontSize = mbytFontSize
    picPati.Width = cbo病人.Width + cbo病人.Left
    picPati.Height = cbo病人.Height + cbo病人.Top
    
    If Not cbsThis Is Nothing Then
        Set objExtendedBar = cbsThis.Add("查看", xtpBarTop)
        objExtendedBar.ContextMenuPresent = False
        objExtendedBar.ShowTextBelowIcons = False
        objExtendedBar.EnableDocking xtpFlagHideWrap
        With objExtendedBar.Controls
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Visible = mblnEditable
            pic护理等级.Visible = mblnEditable
            cbrCustom.Handle = pic护理等级.hWnd
            cbrCustom.ToolTipText = "护理等级"
            
            Set cbrControl = .Add(xtpControlLabel, 0, "显示天数")
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Handle = txt显示天数.hWnd
            cbrCustom.ToolTipText = "显示几天以内的数据"
            Set cbrControl = .Add(xtpControlLabel, 0, "病人")
            If Not mblnEditable Then cbrControl.Visible = False
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Visible = mblnEditable
            picPati.Visible = mblnEditable
            cbrCustom.Handle = picPati.hWnd
            cbrCustom.ToolTipText = "病人列表"
        End With
        
        For Each objCtrl In cbsThis.Item(cbsThis.Count - 1).Controls
            objCtrl.Delete
        Next
        If Not cbsThis.Item(cbsThis.Count - 1) Is Nothing Then cbsThis.Item(cbsThis.Count - 1).Delete
        cbsThis.Item(2).Visible = mblnEditable
    End If
    
    '开始进行表格处理
    With vsf
        lngReDraw = .Redraw
        .Redraw = flexRDNone
        .FontSize = mbytFontSize
        .FontName = "宋体"
        .RowHeightMin = BlowUp(IIf(mblnEditable, 600, 300))
        .RowHeightMax = BlowUp(2000)
        .ColWidth(0) = BlowUp(300)
        .ColWidth(1) = BlowUp(1000)
        .ColWidth(2) = BlowUp(800)
        For lngCol = 3 To .Cols - 1
            .ColWidth(lngCol) = BlowUp(900)
        Next lngCol
        Call vsf.AutoSize(0, vsf.Cols - 1)
        .Redraw = lngReDraw
        .Refresh
    End With
    
    cbsThis.RecalcLayout
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange
    If mbytFontSize <> 12 Then Exit Function
    BlowUp = CInt(dblChange + (dblChange * 1 / 3))
End Function

Public Function GetCopyData() As String
    Dim intCol As Integer
    Dim lngOrder As Long
    Dim blnCopy As Boolean, blnDo As Boolean
    Dim strOrder As String, strData As String
    On Error GoTo errHand
    '只复制有效项目区域(含手术)
    
    If vsf.Row <> vsf.RowSel Then
        MsgBox "不支持多行复制！", vbInformation, gstrSysName
        Exit Function
    End If
    
    For intCol = vsf.Col To vsf.ColSel
        mrsSelItems.Filter = "列=" & intCol
        If mrsSelItems.RecordCount <> 0 Then
            lngOrder = mrsSelItems!项目序号
            blnCopy = True
        ElseIf vsf.Col = mlngOper Then
            lngOrder = mlngOper
            blnCopy = True
        Else
            blnCopy = False
        End If
        
        If blnCopy Then
            strOrder = strOrder & IIf(Not blnDo, "", ",") & lngOrder
            strData = strData & IIf(Not blnDo, "", ",") & vsf.TextMatrix(vsf.Row, intCol)
            blnDo = True
        End If
    Next
    mrsSelItems.Filter = 0
    
    GetCopyData = strOrder & "|" & strData
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsSelItems.Filter = 0
End Function

Public Function IsPigeonhole() As Boolean
    IsPigeonhole = (mstrPigeonhole <> "")
End Function

Private Sub cbo病人_Click()
    If mblnInit = False Then Exit Sub
    If cbo病人.Tag = cbo病人.ListIndex Then Exit Sub
    
    cbo病人.Tag = cbo病人.ListIndex
    Call ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng病区ID, cbo病人.ItemData(cbo病人.ListIndex), mbyt护理等级, mstrPrivs, False, mblnEditable)
End Sub

Private Sub cbo部位_Click()
    If txt数据.Enabled = False Or Val(cbo部位.Tag) = 1 Then txt数据.Text = cbo部位.Text
End Sub

Private Sub cbo护理等级_Click()
    If mblnInit = False Then Exit Sub
    Call ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng病区ID, cbo病人.ItemData(cbo病人.ListIndex), mbyt护理等级, mstrPrivs, False, mblnEditable)
End Sub

'----------------------------------------------------------------
'录入相关的控制说明：
'固定/表示录入脉搏短拙与物理降温
'在输入内容后按下键则弹出部位或方式
'用*或小键盘上其它字符代替下键
'按Del键清除当前列的内容
'-----------------
'列格式说明:日期,时间,(体温,脉搏...,)手术,(大便次数...,)签名人,签名时间
'日期,时间是固定的
'其后是模板定义项目,然后是手术列,也是固定的
'其查询天数内,如是存在当前表格外的项目,自动将项目添加到表格中
'最后是签名人,签名时间,记录ID,组号,记录人
'-----------------
'表格属性说明
'RowData:0-未修改;1-新增或修改
'CellData:0-未修改;1-新增或修改
'-----------------
'只有记录ID为空的行,才允许删除整行;否则,只能清除除日期,时间外的数据
'----------------------------------------------------------------


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Dim blnClear As Boolean     '清除吗?
    Dim blnData As Boolean      '有数据则为真
    Dim strSymbol As String
    Dim strSelItems As String
    Dim strDelItem As String
    Dim lngOrder As Long
    Dim intRow As Integer, intCol As Integer, intRowSel As Integer, intColSel As Integer
    Dim intRow_ As Integer, intCol_ As Integer
    
    Select Case Control.ID
    Case conMenu_Edit_Copy
        mstrSel = vsf.Row & "," & vsf.Col & "," & vsf.RowSel & "," & vsf.ColSel
    Case conMenu_Edit_PASTE
        Call PasteData
    Case conMenu_Edit_Clear
        
        '依次将有数据的列找出来,将其colData设置为1,然后将所选单元格的内容清空
        blnEnable = picInput.Visible
        intRow = vsf.Row
        intCol = vsf.Col
        intRowSel = vsf.RowSel
        intColSel = vsf.ColSel
        
        If vsf.Row > vsf.RowSel Then intRow = vsf.RowSel: intRowSel = vsf.Row
        If vsf.Col > vsf.ColSel Then intCol = vsf.ColSel: intColSel = vsf.Col
        If intColSel >= mlngSigner Then intColSel = mlngSigner - 1
        
        For intRow_ = intRow To intRowSel
            For intCol_ = intCol To intColSel
                If vsf.TextMatrix(intRow_, intCol_) <> "" Then
                    '只有记录ID为空的行,才允许删除整行;否则,只能清除除日期,时间外的数据
                    If Not (Val(vsf.TextMatrix(intRow_, mlngRecord)) <> 0 And intCol_ <= 2) Then
                        blnClear = CheckVersion(intRow_, intCol_)
                        
                        If blnClear Then
                            vsf.Cell(flexcpData, intRow_, intCol_) = 1
                            vsf.Cell(flexcpText, intRow_, intCol_) = ""
                            vsf.RowData(intRow_) = 1
                            mblnChange = True
                        End If
                    End If
                End If
            Next
        Next
        
        '对于记录ID为空,且整行无数据的无效行,删除掉
        intRowSel = vsf.Rows - 1        '最后一行永远不删,当做新增空白行,留给用户录入
        intColSel = mlngSigner - 1
        For intRow = intRowSel To 1 Step -1
            blnData = False
            For intCol = IIf(Val(vsf.RowData(intRow)) = 0, 1, 3) To intColSel
                If vsf.TextMatrix(intRow, intCol) <> "" Then
                    blnData = True
                    Exit For
                End If
            Next
            If Not blnData Then
                If Val(vsf.TextMatrix(intRow, mlngRecord)) <> 0 Then   '历史数据隐藏
                    vsf.RowHidden(intRow) = True
                Else
                    If intRow <> vsf.Rows - 1 Then
                        vsf.RemoveItem intRow               '新记录删除
                    End If
                End If
            End If
        Next
        
        mblnShow = False
        picInput.Visible = False
        
        '清除选择区域
        vsf.RowSel = vsf.Row
        vsf.ColSel = vsf.Col
        vsf.SetFocus
        If blnEnable Then Call Vsf_EnterCell
        If mblnChange Then RaiseEvent AfterDataChanged
    Case conMenu_Edit_SPECIALCHAR
        strSymbol = frmInsSymbol.ShowMe(False, 0)
        txt数据.Text = txt数据.Text & strSymbol
    Case conMenu_Edit_Append
        '手术列与签名人之间的列,都是临时添加的项目,这部分项目是按项目序号大小顺序添加的,因此,在手工添加时,也应该保证此顺序,避免刷新后列顺序发生变化
        With mrsSelItems
            '得到已选择项目的序号清单
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                strSelItems = strSelItems & "," & !项目序号
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
        End With
        strSelItems = strSelItems & ","
        
        strSelItems = frmTendItemChoose.ShowSelect(strSelItems, cbo护理等级.ListIndex, cbo病人.ItemData(cbo病人.ListIndex), mlng科室ID)
        If strSelItems = "" Then Exit Sub
        mstrSelItems = mstrSelItems & IIf(mstrSelItems = "", "", vbCrLf) & strSelItems
        
        Call InsertColumn(strSelItems)
    Case conMenu_Edit_Delete
        '如果查询列表中有数据则不允许删除
        intCol = vsf.Col
        intRowSel = vsf.Rows - 1
        For intRow = vsf.Row To intRowSel
            If vsf.TextMatrix(intRow, intCol) <> "" Or vsf.Cell(flexcpData, intRow, intCol) <> 0 Then
                MsgBox "当前项目有数据，不允许删除！", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
        
        Call DeleteColumn(intCol)
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '审签后的数据不允许做任何处理
    
    If mblnInit = False Then Exit Sub
    Select Case Control.ID
    Case conMenu_Edit_PASTE
        Control.Visible = mblnEditable
        Control.Enabled = (mstrSel <> "") And Not IsPigeonhole And mblnEditable And vsf.TextMatrix(vsf.Row, mlngCertLevel) <> "护士长"
    Case conMenu_Edit_Copy, conMenu_Edit_SPECIALCHAR, conMenu_Edit_Append
        Control.Visible = mblnEditable
        Control.Enabled = Not IsPigeonhole And mblnEditable And (InStr(1, mstrPrivs, "护理记录登记") <> 0)
    Case conMenu_Edit_Clear '签名的数据不允许清除
        Control.Visible = mblnEditable
        Control.Enabled = Not IsPigeonhole And mblnEditable And (InStr(1, mstrPrivs, "护理记录登记") <> 0) And mblnCheckVersion And vsf.TextMatrix(vsf.Row, mlngCertLevel) <> "护士长"
        'If Control.Enabled Then Control.Enabled = (Vsf.TextMatrix(Vsf.Row, mlngSigner) = "")
        
        '如果是多选,则允许清除
        If vsf.RowSel <> vsf.Row Or vsf.ColSel <> vsf.Col Then Control.Enabled = True
    Case conMenu_Edit_Delete
        Dim blnDel As Boolean
        If mrsSelItems.State = 1 Then
            mrsSelItems.Filter = "列=" & vsf.Col
            If mrsSelItems.RecordCount <> 0 Then
                blnDel = (mrsSelItems!固定 = 0)
            End If
            mrsSelItems.Filter = 0
        End If
        Control.Visible = mblnEditable
        Control.Enabled = Not IsPigeonhole And mblnEditable And blnDel And (InStr(1, mstrPrivs, "护理记录登记") <> 0) And vsf.TextMatrix(vsf.Row, mlngCertLevel) <> "护士长"
    End Select
End Sub

Private Sub cmd未记说明_Click()
    If cbo部位.Visible Then
        If Val(cbo部位.Tag) = 0 Then
            Call txt数据_KeyDown(vbKeyDown, vbShiftMask)
        Else
            Call txt数据_KeyDown(vbKeyDown, 0)
            txt数据.Text = ""
            txt数据.SetFocus
        End If
    Else
        Call txt数据_KeyDown(vbKeyW, vbCtrlMask)
    End If
End Sub

Private Sub InitEnv()
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("数据补录时限", glngSys))
    
    '打开现存在的所有护理记录项目
    gstrSQL = " Select 项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
              " From 护理记录项目 B" & _
              " Where B.应用方式<>0 " & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitBill()
    Dim intCol As Integer, intCols As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    '初始化内存记录集
    strFields = "列," & adDouble & ",18|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|固定," & adDouble & ",2"
    Call Record_Init(mrsSelItems, strFields)
    strFields = "列|项目序号|项目名称|固定"
    
    '先添加模板设定的项目
    strSQL = " Select B.项目序号,B.项目名称,B.项目单位,B.项目类型,1 AS 固定" & _
             " From 护理项目模板 A,护理记录项目 B" & _
             " Where a.项目序号 = b.项目序号 And B.应用方式<>0 And A.科室ID=[3] And A.护理等级 = [1] And B.适用病人 IN (0,[2])" & _
             " And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[3])))" & _
             " Order by A.排列序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "先添加模板设定的项目", cbo护理等级.ListIndex, IIf(cbo病人.ItemData(cbo病人.ListIndex) = 0, 1, 2), mlng科室ID)
    If rsTemp.RecordCount = 0 Then
        '按以前的规则提取项目清单供录入
        strSQL = " Select B.项目序号,B.项目名称,B.项目单位,B.项目类型,0 AS 固定" & _
                 " From 护理记录项目 B" & _
                 " Where B.应用方式<>0 And B.护理等级>=[1] And B.适用病人 IN (0,[2])" & _
                 " And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[3])))" & _
                 " Order by B.项目序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "按以前的规则提取项目清单供录入", cbo护理等级.ListIndex, IIf(cbo病人.ItemData(cbo病人.ListIndex) = 0, 1, 2), mlng科室ID)
    End If
    
    With vsf
        intCols = .Cols - 1
        For intCol = 1 To intCols
            .ColHidden(intCol) = False
        Next
        
        .Clear
        .Rows = 2
        .FixedCols = 1
        .Cols = rsTemp.RecordCount + .FixedCols + 3     '加上日期时间列,再加上固定的手术列
        .RowHeightMin = IIf(mblnEditable, 600, 300)
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .WordWrap = True
        
        .TextMatrix(0, 1) = "日期"
        .TextMatrix(0, 2) = "时间"
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 600
        .ColWidth(2) = 800
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        
        intCol = 3
        Do While Not rsTemp.EOF
            If rsTemp!项目名称 Like "舒张压*" And .TextMatrix(0, intCol - 1) Like "收缩压*" Then
                .TextMatrix(0, intCol - 1) = "血压" & IIf(NVL(rsTemp!项目单位) = "", "", vbCrLf & "(" & rsTemp!项目单位 & ")")
                .Cols = .Cols - 1
                intCol = intCol - 1
            Else
                .TextMatrix(0, intCol) = rsTemp!项目名称 & IIf(NVL(rsTemp!项目单位) = "", "", vbCrLf & "(" & rsTemp!项目单位 & ")")
            End If
            .ColWidth(intCol) = 900
            .ColAlignment(intCol) = IIf(rsTemp!项目类型 = 0, flexAlignCenterCenter, flexAlignLeftTop)       '数字则居中显示,非数字以用户录入的数据显示
            
            '将目前已选择的项目加入内存记录集中
            strValues = intCol & "|" & rsTemp!项目序号 & "|" & rsTemp!项目名称 & "|" & rsTemp!固定
            Call Record_Add(mrsSelItems, strFields, strValues)
            
            intCol = intCol + 1
            rsTemp.MoveNext
        Loop
        '.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .MergeCells = flexMergeFree
        .WordWrap = True
        
        '将目前已选择的项目加入内存记录集中
        strValues = .Cols - 1 & "|0|手术|1"
        Call Record_Add(mrsSelItems, strFields, strValues)
        
        mlngOper = .Cols - 1
        .TextMatrix(0, .Cols - 1) = "手术"
        .TextMatrix(1, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        .TextMatrix(1, 2) = Format(zlDatabase.Currentdate, "HH:mm")
    End With
    
    '检查是否需要录入心率
    mrsSelItems.Filter = "项目序号=-1"
    mbln心率 = (mrsSelItems.RecordCount <> 0)
    mrsSelItems.Filter = 0
End Sub

Private Sub ReadData()
    Dim arrColumn
    Dim intStart As Integer, intEnd As Integer
    
    Dim int心率应用 As Integer
    Dim strStart As String, strEnd As String
    Dim rsColumns As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '读取近期多少天的数据
    
    mrsItems.Filter = "项目序号=-1"
    If mrsItems.RecordCount <> 0 Then
        int心率应用 = mrsItems!应用方式
    End If
    mrsItems.Filter = 0
    strStart = Format(DateAdd("d", -1 * Val(txt显示天数.Text), zlDatabase.Currentdate), "yyyy-MM-dd") & " 00:00:00"
    strEnd = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd") & " 23:59:59"
    
    '检查是否归档
    gstrSQL = " Select 归档人 From 病人护理记录 Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And Rownum<2"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否归档", mlng病人ID, mlng主页ID, cbo病人.ItemData(cbo病人.ListIndex))
    If rsTemp.RecordCount <> 0 Then mstrPigeonhole = NVL(rsTemp!归档人)
    
    '1、先提取出查询时间范围内自己添加的项目,依次加到表格中
    gstrSQL = " Select Distinct Y.项目序号,Y.项目名称 From (" & _
                    " Select A.项目序号 " & _
                    " From 病人护理内容 A,病人护理记录 C" & _
                    " Where C.ID = A.记录id AND A.记录类型 =1 AND C.病人来源 = 2 AND ((NVL(A.记录标记,0) <> 1 And a.项目序号>0) or a.项目序号=-1 ) " & _
                         " AND C.发生时间 Between [1] And [2] And C.病人ID=[3] And C.主页ID=[4]" & _
                    "       " & _
                    "      ) X,护理记录项目 Y " & _
              " Where Y.项目序号 = X.项目序号 AND nvl(Y.护理等级,3) >=[6] And Nvl(y.应用方式,0)=1 And Nvl(y.适用病人,0) In (0,[7]) And (Y.适用科室=1 Or (Y.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=Y.项目序号 And D.科室id=[5])))  " & _
              " Order By Y.项目序号"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsColumns = zlDatabase.OpenSQLRecord(gstrSQL, "先提取出查询时间范围内自己添加的项目,依次加到表格中", CDate(strStart), CDate(strEnd), _
                mlng病人ID, mlng主页ID, mlng科室ID, cbo护理等级.ListIndex, IIf(cbo病人.ItemData(cbo病人.ListIndex) = 0, 1, 2))
    Call AddColumns(rsColumns)
    
    '将用户选择的列添加进去
    If mstrSelItems <> "" Then
        arrColumn = Split(mstrSelItems, vbCrLf)
        intEnd = UBound(arrColumn)
        For intStart = 0 To intEnd
            Call InsertColumn(arrColumn(intStart))
        Next
    End If
    vsf.Cell(flexcpAlignment, 0, 0, 0, vsf.Cols - 1) = flexAlignCenterCenter
    
    '2、提取数据
    gstrSQL = " Select X.* From ("
    If int心率应用 = 2 Then
        gstrSQL = gstrSQL & _
                    "Select A.项目序号,DECODE(A.记录类型,4,A.项目名称, A.记录内容) As 记录结果, " & _
                        "D.项目ID AS 证书ID,Nvl(A.终止版本,A.开始版本) AS 实际版本,D.记录人 AS 签名人,NVL(D.项目名称,to_char(D.修改时间,'yyyy-MM-dd hh24:mi:ss')) As 签名时间,NVL(D.记录内容,'护士') AS 签名级别," & _
                        "Decode(a.记录内容,Null,'',A.体温部位) As 部位,b.记录内容 As 标记,b.记录标记," & _
                        "C.发生时间 As 完成日期,A.记录id,A.记录组号,a.未记说明,a.记录人 " & _
                    " From 病人护理内容 A, 病人护理内容 B,病人护理记录 C,病人护理内容 D " & _
                    " Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                         " AND A.记录类型 =1 AND C.病人来源 = 2 AND NVL(A.记录标记,0) <> 1 " & _
                         " And D.记录类型(+)=5 And D.记录ID(+)=C.ID And D.终止版本(+) Is NULL" & _
                         " AND C.发生时间 Between [1] And [2] And C.病人ID=[3] And C.主页ID=[4] and C.婴儿=[8]"
    Else
        gstrSQL = gstrSQL & _
                    "Select A.项目序号,DECODE(A.记录类型,4,A.项目名称, A.记录内容) As 记录结果, " & _
                        "D.项目ID AS 证书ID,Nvl(A.终止版本,A.开始版本) AS 实际版本,D.记录人 AS 签名人,NVL(D.项目名称,to_char(D.修改时间,'yyyy-MM-dd hh24:mi:ss')) As 签名时间,NVL(D.记录内容,'护士') AS 签名级别," & _
                        "Decode(a.记录内容,Null,'',A.体温部位) As 部位,Decode(a.项目序号,2,'',-1,'',b.记录内容) As 标记,Decode(a.项目序号,2,0,-1,0,b.记录标记) As 记录标记," & _
                        "C.发生时间 As 完成日期,A.记录id,A.记录组号,a.未记说明,a.记录人 " & _
                    " From 病人护理内容 A, 病人护理内容 B,病人护理记录 C,病人护理内容 D " & _
                    " Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                         " AND A.记录类型 =1 AND C.病人来源 = 2 AND ((NVL(A.记录标记,0) <> 1 And a.项目序号>0) or a.项目序号=-1 or (a.项目序号=0 and a.记录类型=4)) " & _
                         " And D.记录类型(+)=5 And D.记录ID(+)=C.ID And D.终止版本(+) Is NULL" & _
                         " AND C.发生时间 Between [1] And [2] And C.病人ID=[3] And C.主页ID=[4] and C.婴儿=[8]"
    End If
    gstrSQL = gstrSQL & _
                "       And a.终止版本 Is Null And b.终止版本 Is Null " & _
                "       And Decode(a.项目序号,2,-1,a.项目序号)=b.项目序号(+)) X,护理记录项目 Y " & _
                "Where Y.项目序号 = X.项目序号 AND nvl(Y.护理等级,3) >=[6] And Nvl(y.应用方式,0)=1 And Nvl(y.适用病人,0) In (0,[7]) And (Y.适用科室=1 Or (Y.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=Y.项目序号 And D.科室id=[5])))  "
    
    '加上手术项目
    gstrSQL = gstrSQL & _
                " UNION " & _
                " Select A.项目序号,DECODE(A.记录类型,4,A.项目名称, A.记录内容) As 记录结果, " & _
                    "D.项目ID AS 证书ID,Nvl(A.终止版本,A.开始版本) AS 实际版本,D.记录人 AS 签名人,NVL(D.项目名称,to_char(D.修改时间,'yyyy-MM-dd hh24:mi:ss')) As 签名时间,NVL(D.记录内容,'护士') AS 签名级别," & _
                    "Decode(a.记录内容,Null,'',A.体温部位) As 部位,Decode(a.项目序号,2,'',-1,'',b.记录内容) As 标记,Decode(a.项目序号,2,0,-1,0,b.记录标记) As 记录标记," & _
                    "C.发生时间 As 完成日期,A.记录id,A.记录组号,a.未记说明,a.记录人 " & _
                " From 病人护理内容 A, 病人护理内容 B,病人护理记录 C,病人护理内容 D " & _
                " Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                     " AND A.记录类型 =4 AND C.病人来源 = 2 And a.终止版本 Is Null And b.终止版本 Is Null And D.终止版本(+) Is NULL" & _
                     " And D.记录类型(+)=5 And D.记录ID(+)=C.ID" & _
                     " AND C.发生时间 Between [1] And [2] And C.病人ID=[3] And C.主页ID=[4] And C.婴儿=[8]"
    
    gstrSQL = " Select * From (" & gstrSQL & ") Order By 完成日期,记录ID,记录组号,DECODE(项目序号,0,999,项目序号)"

    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取数据", CDate(strStart), CDate(strEnd), _
                mlng病人ID, mlng主页ID, mlng科室ID, cbo护理等级.ListIndex, IIf(cbo病人.ItemData(cbo病人.ListIndex) = 0, 1, 2), cbo病人.ItemData(cbo病人.ListIndex))

    '准备添加数据(遇到没有的项目,直接在表格中增加该列,同时处理内部记录集
    Call ShowData(rsData)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DeleteColumn(ByVal intCol As Integer)
    Dim lngOrder As Long
    Dim strName As String
    Dim arrColumn
    Dim intStart As Integer, intEnd As Integer
    '删除指定的列
    
    mrsSelItems.Filter = "列=" & intCol
    lngOrder = mrsSelItems!项目序号
    strName = mrsSelItems!项目名称
    mrsSelItems.Filter = 0
    
    '删除列
    vsf.ColPosition(intCol) = vsf.Cols - 1
    vsf.Cols = vsf.Cols - 1
    '处理内部记录集
    With mrsSelItems
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !列 > intCol Then
                !列 = !列 - 1
                .Update
            ElseIf !列 = intCol Then
                .Delete
            Else
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    '相关模块变量的更新
    If mlngOper > intCol Then mlngOper = mlngOper - 1
    mlngSigner = mlngSigner - 1
    mlngSignTime = mlngSignTime - 1
    mlngRecord = mlngRecord - 1
    mlngGroup = mlngGroup - 1
    mlngCert = mlngCert - 1
    mlngCertLevel = mlngCertLevel - 1
    
    arrColumn = Split(mstrSelItems, vbCrLf)
    intEnd = UBound(arrColumn)
    mstrSelItems = ""
    For intStart = 0 To intEnd
        If Val(Split(arrColumn(intStart), "|")(0)) <> lngOrder Then
            mstrSelItems = mstrSelItems & IIf(mstrSelItems = "", "", vbCrLf) & arrColumn(intStart)
        End If
    Next
End Sub

Private Sub InsertColumn(ByVal strSelItems As String)
    Dim lngOrder As Long
    
    '如果已存在该列则退出
    mrsSelItems.Filter = "项目序号=" & Val(Split(strSelItems, "|")(0))
    If mrsSelItems.RecordCount <> 0 Then
        mrsSelItems.Filter = 0
        Exit Sub
    End If
    
    '将用户选择的项目添加到表格中
    mrsItems.Filter = "项目序号=" & Val(Split(strSelItems, "|")(0))
    vsf.Cols = vsf.Cols + 1
    vsf.TextMatrix(0, vsf.Cols - 1) = Split(strSelItems, "|")(1) & IIf(NVL(mrsItems!项目单位) = "", "", vbCrLf & "(" & mrsItems!项目单位 & ")")
    vsf.ColAlignment(vsf.Cols - 1) = IIf(mrsItems!项目类型 = 0, flexAlignCenterCenter, flexAlignLeftTop)       '数字则居中显示,非数字以用户录入的数据显示
    mrsItems.Filter = 0
    'Vsf.Cell(flexcpAlignment, 0, Vsf.Cols - 1, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter  '整列进行设置
        
    '取手术后那列的项目序号
    With mrsSelItems
        .Filter = "列>" & mlngOper
        .Sort = "列"
        Do While Not .EOF
            If !项目序号 > Val(Split(strSelItems, "|")(0)) Then
                lngOrder = !列
                Exit Do
            End If
            .MoveNext
        Loop
        If lngOrder = 0 Then lngOrder = mlngSigner  '没找着,说明没得添加项目,取签名列
        
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    vsf.ColPosition(vsf.Cols - 1) = lngOrder      '签名人列开始往后移
    '处理内部记录集
    With mrsSelItems
        Do While Not .EOF
            If !列 >= lngOrder Then
                !列 = !列 + 1
                .Update
            End If
            .MoveNext
        Loop
    End With
    strValues = lngOrder & "|" & Split(strSelItems, "|")(0) & "|" & Split(strSelItems, "|")(1) & "|0"
    Call Record_Add(mrsSelItems, strFields, strValues)
    '相关模块变量的更新
    mlngSigner = mlngSigner + 1
    mlngSignTime = mlngSignTime + 1
    mlngRecord = mlngRecord + 1
    mlngGroup = mlngGroup + 1
    mlngCert = mlngCert + 1
    mlngCertLevel = mlngCertLevel + 1
End Sub

Private Sub AddColumns(ByVal rsColumns As ADODB.Recordset)
    '将历史数据中存在的多余列添加到表格中
    With rsColumns
        Do While Not .EOF
            mrsSelItems.Filter = "项目序号=" & !项目序号
            If mrsSelItems.RecordCount = 0 Then
                mrsItems.Filter = "项目序号=" & !项目序号
                vsf.Cols = vsf.Cols + 1
                vsf.TextMatrix(0, vsf.Cols - 1) = .Fields("项目名称").Value & IIf(NVL(mrsItems!项目单位) = "", "", vbCrLf & "(" & mrsItems!项目单位 & ")")
                vsf.ColAlignment(vsf.Cols - 1) = IIf(mrsItems.Fields("项目类型").Value = 0, flexAlignCenterCenter, flexAlignLeftTop)
                mrsItems.Filter = 0
                
                strValues = vsf.Cols - 1 & "|" & !项目序号 & "|" & !项目名称 & "|0"
                Call Record_Add(mrsSelItems, strFields, strValues)
            End If
            .MoveNext
        Loop
    End With
    
    '固定加入签名人,签名时间列
    With vsf
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "签名人"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        mlngSigner = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "签名时间"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        mlngSignTime = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "证书ID"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngCert = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "记录ID"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngRecord = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "签名级别"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngCertLevel = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "组号"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngGroup = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "记录人"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    mrsSelItems.Filter = 0
End Sub

Private Sub ShowData(ByVal rsData As ADODB.Recordset)
    On Error GoTo errHand
    Dim lngRow As Long
    Dim lngRecord As Long   '记录ID
    Dim lngGroup As Long    '组号
    Dim strData As String
    Dim strTime As String
    Dim lng终止版本 As Long, bln上色 As Boolean
    Dim rsTemp As New ADODB.Recordset   '提取当前记录最大的终止版本
    
    '再循环写数据
    lngRow = 1
    With rsData
        Do While Not .EOF
            If lngRecord <> !记录ID Or lngGroup <> !记录组号 Then
                '提取当前记录最大的终止版本
                gstrSQL = " Select max(开始版本),Max(终止版本) From 病人护理内容 Where 记录ID=[1]"
                If mblnMoved_HL Then
                    gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
                    gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
                End If
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前记录最大的终止版本", CLng(!记录ID))
                lng终止版本 = NVL(rsTemp.Fields(0).Value, 1)
                If lng终止版本 < NVL(rsTemp.Fields(1).Value, 1) Then lng终止版本 = NVL(rsTemp.Fields(1).Value, 1)
                
                '新的记录
                If lngRecord <> 0 Then
                    '增加行
                    lngRow = lngRow + 1
                    If lngRow > vsf.Rows - 1 Then vsf.Rows = vsf.Rows + 1
                    vsf.TextMatrix(lngRow, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                    vsf.TextMatrix(lngRow, 2) = Format(zlDatabase.Currentdate, "HH:mm")
                    'Vsf.Cell(flexcpAlignment, lngRow, 0, lngRow, Vsf.Cols - 1) = flexAlignCenterCenter
                Else
                    '第一条记录
                End If
                strTime = Format(!完成日期, "yyyy-MM-dd HH:mm")
                
                '先写入签名人及签名时间
                lngRecord = !记录ID
                lngGroup = !记录组号
                bln上色 = True
                If Not IsNull(!签名人) Then
                    bln上色 = False
                    vsf.Cell(flexcpPicture, lngRow, 0) = imgRow.ListImages(1).Picture
                End If
                vsf.Cell(flexcpPictureAlignment, lngRow, 0) = flexAlignCenterCenter
                vsf.TextMatrix(lngRow, 1) = Split(strTime, " ")(0)
                vsf.TextMatrix(lngRow, 2) = Split(strTime, " ")(1)
                vsf.TextMatrix(lngRow, mlngCert) = Val(NVL(.Fields("证书ID").Value, 0))
                vsf.TextMatrix(lngRow, mlngCertLevel) = NVL(.Fields("签名级别").Value)
                vsf.TextMatrix(lngRow, mlngSigner) = NVL(.Fields("签名人").Value)
                vsf.TextMatrix(lngRow, mlngSignTime) = Format(.Fields("签名时间").Value, "yyyy-MM-dd HH:mm:ss")
                vsf.TextMatrix(lngRow, mlngRecord) = CLng(.Fields("记录ID").Value)
                vsf.TextMatrix(lngRow, mlngGroup) = CLng(.Fields("记录组号").Value)
                vsf.TextMatrix(lngRow, vsf.Cols - 1) = NVL(.Fields("记录人").Value)
                vsf.RowData(lngRow) = 0
                
                If bln上色 Then '签名人为空,且终止版本大于1,才说明需要上色;排开初步产生的数据不需要上色的情况
                    bln上色 = (lng终止版本 > 1)
                End If
            End If
            
            '先写入普通的护理项目
            If !项目序号 <> 0 Then
                '如果未记说明不为空,显示未记说明
                If Not IsNull(.Fields("未记说明").Value) Then
                    strData = .Fields("未记说明").Value
                Else
                    strData = NVL(.Fields("记录结果").Value)
                    If Not IsNull(.Fields("标记").Value) Then
                        strData = strData & "/" & .Fields("标记").Value
                    End If
                    If Not IsNull(.Fields("部位").Value) Then
                        strData = .Fields("部位").Value & ":" & strData
                    ElseIf !项目序号 = 1 Then
                        strData = "腋温:" & strData
                    End If
                End If
                
                mrsSelItems.Filter = "项目序号=" & !项目序号
                If mrsSelItems.RecordCount <> 0 Then
                    If !项目序号 = 5 Then   '收缩压,如果对应单元格有内容,则说明已填入舒张压,以/组合显示
                        If vsf.TextMatrix(lngRow, mrsSelItems!列) <> "" Then
                            vsf.TextMatrix(lngRow, mrsSelItems!列) = vsf.TextMatrix(lngRow, mrsSelItems!列) & "/" & strData
                        Else
                            vsf.TextMatrix(lngRow, mrsSelItems!列) = strData
                        End If
                    Else
                        vsf.TextMatrix(lngRow, mrsSelItems!列) = strData
                    End If
                End If
            Else
                '再写入手术
                strData = NVL(.Fields("记录结果").Value)
                mrsSelItems.Filter = "项目序号=0"
                If mrsSelItems.RecordCount <> 0 Then
                    vsf.TextMatrix(lngRow, mrsSelItems!列) = strData
                End If
            End If
            
            '上色(手术除外)
            If !实际版本 = lng终止版本 And bln上色 Then
                vsf.Cell(flexcpForeColor, lngRow, mrsSelItems!列) = &HFF&
            End If
            
            .MoveNext
        Loop
    End With
    mrsSelItems.Filter = 0
    
    '增加空白行
    If Val(vsf.TextMatrix(vsf.Rows - 1, mlngRecord)) <> 0 Then
        lngRow = lngRow + 1
        If lngRow > vsf.Rows - 1 Then vsf.Rows = vsf.Rows + 1
        vsf.TextMatrix(lngRow, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        vsf.TextMatrix(lngRow, 2) = Format(zlDatabase.Currentdate, "HH:mm")
        'Vsf.Cell(flexcpAlignment, lngRow, 0, lngRow, Vsf.Cols - 1) = flexAlignCenterCenter
    End If
    
    '使用CellData来保存修改标志
    vsf.Cell(flexcpData, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = 0
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsSelItems.Filter = 0
End Sub

Private Sub PasteData()
    Dim arrSel, arrData, arrRow
    Dim strSource As String
    Dim intRow As Integer, intCol As Integer
    Dim intSourceRow As Integer, intSourceCol As Integer, intSourceRowSel As Integer, intSourceColSel As Integer
    '检查待粘贴的区域，是否与拷贝的内容存在重叠，算法上有区别
    
    arrSel = Split(mstrSel, ",")
    intSourceRow = arrSel(0)
    intSourceCol = arrSel(1)
    intSourceRowSel = arrSel(2)
    intSourceColSel = arrSel(3)
    '如果反起选择的,则需要调整一下起始行,列,终止行,列
    If intSourceRow > intSourceRowSel Then intRow = intSourceRow: intSourceRow = intSourceRowSel: intSourceRowSel = intRow
    If intSourceCol > intSourceColSel Then intCol = intSourceCol: intSourceCol = intSourceColSel: intSourceColSel = intCol
    '签名人,签名时间,记录ID,组号这四列不复制
    'If intSourceColSel > Vsf.Cols - 5 Then intSourceColSel = Vsf.Cols - 5
    If intSourceColSel >= mlngSigner Then intSourceColSel = mlngSigner - 1
    
    '需粘贴的起始列必须与拷贝的起始列相同,才允许执行粘贴操作
    If vsf.Col <> intSourceCol Then
        MsgBox "待粘贴的列必须与复制的起始列相同！", vbInformation, gstrSysName
        Exit Sub
    End If
    If vsf.Row = intSourceRow Then Exit Sub
    
    '得到待粘贴区域
    If vsf.Row > intSourceRow And vsf.Row <= intSourceRowSel Then
        If MsgBox("你所选择的粘贴区域与复制区域重合了，你确定要进行粘贴操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    '将复制区域的数据临时写入变量中
    For intRow = intSourceRow To intSourceRowSel
        strSource = strSource & IIf(intRow = intSourceRow, "", "|优优|")
        For intCol = intSourceCol To intSourceColSel
            strSource = strSource & IIf(intCol = intSourceCol, "", "|小宝|") & vsf.TextMatrix(intRow, intCol)
        Next
    Next
    
    '进行复制(手术列不复制)
    If strSource = "" Then Exit Sub
    arrData = Split(strSource, "|优优|")
    intSourceRowSel = vsf.Row + (intSourceRowSel - intSourceRow)
    For intRow = vsf.Row To intSourceRowSel
        arrRow = Split(arrData(intRow - vsf.Row), "|小宝|")
        If intRow > vsf.Rows - 1 Then Exit For
        For intCol = intSourceCol To intSourceColSel
            '原来有值,或者复制单元格有值,才填写修改标志
            If intCol <> mlngOper Then
                If vsf.TextMatrix(intRow, intCol) <> "" Or arrRow(intCol - vsf.Col) <> "" Then
                    vsf.TextMatrix(intRow, intCol) = arrRow(intCol - vsf.Col)
                    vsf.Cell(flexcpData, intRow, intCol) = 1
                    vsf.RowData(intRow) = 1
                    mblnChange = True
                End If
            End If
        Next
    Next
    If mblnChange Then RaiseEvent AfterDataChanged
End Sub

Private Sub InitPanelMain()
    Dim objPane As Pane
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    
    dkpMain.SetCommandBars cbsThis
    
    Set objPane = dkpMain.CreatePane(1, 100, 200, DockTopOf, Nothing): objPane.Title = "编辑": objPane.Options = PaneNoCaption
    objPane.Handle = picMain.hWnd
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    '63174:刘鹏飞,2013-07-03,将根据mblnEditable判断是否加载菜单取消，在菜单的Update事件中进行控制菜单是否可见.
    '因为窗体加载时没有选择病人mblnEditable=False,部分菜单没有加载，在选择病人时mblnEditable=ture但菜单不会在加载。
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 16, 16
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '快键绑定
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制"): cbrControl.ToolTipText = "复制(Ctrl+C)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "粘贴"):  cbrControl.ToolTipText = "粘贴(Ctrl+V)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"):   cbrControl.ToolTipText = "清除"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "特殊符号"):  cbrControl.ToolTipText = "插入特殊符号(Ctrl+D)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "添加"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "添加项目(Alt+A)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"):  cbrControl.ToolTipText = "删除项目(Alt+D)"
    End With
    
    '显示天数工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objExtendedBar = cbsThis.Add("查看", xtpBarTop)
    objExtendedBar.ContextMenuPresent = False
    objExtendedBar.ShowTextBelowIcons = False
    objExtendedBar.EnableDocking xtpFlagHideWrap
    With objExtendedBar.Controls
        
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.Visible = mblnEditable
        pic护理等级.Visible = mblnEditable
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = pic护理等级.hWnd
        cbrCustom.ToolTipText = "护理等级"
        
        Set cbrControl = .Add(xtpControlLabel, 0, "显示天数")
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = txt显示天数.hWnd
        cbrCustom.ToolTipText = "显示几天以内的数据"
        Set cbrControl = .Add(xtpControlLabel, 0, "病人")
        If Not mblnEditable Then cbrControl.Visible = False
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.Visible = mblnEditable
        picPati.Visible = mblnEditable
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = picPati.hWnd
        cbrCustom.ToolTipText = "病人列表"
    End With
    
    'Call SetDockRight(objExtendedBar, cbrToolBar)
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next

     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
        .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
        .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
        .Add FALT, Asc("A"), conMenu_Edit_Append
        .Add FALT, Asc("D"), conMenu_Edit_Delete
    End With
    
    InitMenuBar = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function WriteIntoVsf(Optional ByRef strInfo As String) As Boolean
    Dim blnAllow As Boolean
    Dim StrText As String
    Dim strMsg As String
    Dim lngRecord As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    
    Dim intType As Integer, lngOrder As Long, lngClass As Long, strName As String, lngLength As Long, str值域 As String
    
    If picInput.Visible Then
        lngRow = Split(txt数据.Tag, "|")(0)
        lngCol = Split(txt数据.Tag, "|")(1)
        If txt数据.Enabled Then
            '检查数据合法性
            If Val(cbo部位.Tag) = 0 Then
                If txt数据.Text <> "" Then
                    StrText = IIf(cbo部位.Visible And Trim(cbo部位.Text) <> "", cbo部位.Text & ":", "") & Trim(txt数据.Text)
                End If
            Else
                StrText = IIf(Trim(txt数据.Text) <> "", Trim(txt数据.Text), cbo部位.Text)
            End If
            If lngCol <= 2 Then
                If Trim(StrText) <> "" Then
                    strMsg = "Msgbox"
                    blnAllow = CheckDate2(lngRow, lngCol, StrText, strMsg)
                    strInfo = strMsg
                End If
            Else
                '定位列对应的护理记录进行检查
                mrsSelItems.Filter = "列=" & lngCol
                mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
                
                intType = mrsItems!项目类型     '0-数值；1-文字
                lngClass = mrsItems!项目性质
                lngOrder = mrsItems!项目序号
                strName = mrsItems!项目名称
                lngLength = mrsItems!项目长度 + IIf(NVL(mrsItems!项目小数, 0) = 0, 0, NVL(mrsItems!项目小数, 0) + 1)
                If intType = 0 Then
                    str值域 = NVL(mrsItems!项目值域)
                Else
                    str值域 = ""
                    StrText = txt数据.Text      '非数字型项目,以用户原始录入为准
                End If
                
                '如果是大段文本则不检查数据合法性
                If intType = 1 And lngLength > 100 Then
                    '不做任何处理
                    blnAllow = True
                Else
                    strMsg = "Msgbox"       '传入非空值,表示如果出错,则在该变量中返回错误信息
                    blnAllow = CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, str值域, strMsg)
                    strInfo = strMsg
                End If
                
                mrsItems.Filter = 0
                mrsSelItems.Filter = 0
            End If
            
            If blnAllow Then vsf.TextMatrix(lngRow, lngCol) = StrText
        Else
            blnAllow = True
            vsf.TextMatrix(lngRow, lngCol) = txt数据.Text
        End If
    Else
        blnAllow = True
        lngRow = Split(lvwMultiSel.Tag, "|")(0)
        lngCol = Split(lvwMultiSel.Tag, "|")(1)
        vsf.TextMatrix(lngRow, lngCol) = strInfo
    End If
    txt数据.Tag = ""
    cbo部位.Visible = False
    txt数据.Height = picInput.Height
    picInput.Visible = False
    lvwMultiSel.Visible = False
    
    '更新修改标志
    If blnAllow Then
        If picInput.Tag <> vsf.TextMatrix(lngRow, lngCol) Then
            '如果是修改的时间,需要把记录ID相同的所有记录的时间全部修改了
            If lngCol <= 2 And Val(vsf.TextMatrix(lngRow, mlngRecord)) <> 0 Then
                lngRows = vsf.Rows - 1
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                For lngRow = 1 To lngRows
                    If Val(vsf.TextMatrix(lngRow, mlngRecord)) = lngRecord Then
                        vsf.TextMatrix(lngRow, lngCol) = StrText
                        '修改标志
                        vsf.RowData(lngRow) = 1
                        vsf.Cell(flexcpData, lngRow, lngCol) = 1
                    End If
                Next
            Else
                '修改标志
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                vsf.RowData(lngRow) = 1
                vsf.Cell(flexcpData, lngRow, lngCol) = 1
            End If
            mblnChange = True
        End If
        
        WriteIntoVsf = True
        If mblnChange Then RaiseEvent AfterDataChanged
    End If
End Function

Private Sub lvwMultiSel_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strData As String
    Dim intCol As Integer, intMax As Integer
    Dim blnAllow As Boolean
    
    If KeyCode = vbKeyReturn Then
        intMax = lvwMultiSel.ListItems.Count
        For intCol = 1 To intMax
            If lvwMultiSel.ListItems(intCol).Checked Then
                strData = strData & IIf(strData = "", "", ",") & lvwMultiSel.ListItems(intCol).Text
            End If
        Next
        blnAllow = WriteIntoVsf(strData)
        Call vsf_KeyDown(vbKeyReturn, Shift)
'    ElseIf KeyCode = vbKeyLeft Then
'        Call vsf_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    vsf.Width = picMain.Width
    vsf.Height = picMain.Height - vsf.Top
End Sub

Private Sub cbo部位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call txt数据_KeyDown(vbKeyReturn, 0): Exit Sub
End Sub

Private Sub txt数据_GotFocus()
    Call zlControl.TxtSelAll(txt数据)
End Sub

Private Sub txt数据_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrText As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode = vbKeyDown And InStr(1, "体温脉搏呼吸手术", Mid(vsf.TextMatrix(0, vsf.Col), 1, 2)) <> 0 Then
        If Shift = 0 Then
            cbo部位.Tag = 0
            cbo部位.Clear
            If Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "体温" Then
                cbo部位.AddItem "腋温"
                cbo部位.AddItem "口温"
                cbo部位.AddItem "肛温"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "脉搏" Then
                cbo部位.AddItem ""
                cbo部位.AddItem "起搏器"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "呼吸" Then
                cbo部位.AddItem "自主呼吸"
                cbo部位.AddItem "呼吸机"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "手术" Then
                cbo部位.AddItem "手术"
                cbo部位.AddItem "分娩"
                cbo部位.AddItem "手术分娩"
            End If
            If cbo部位.ListCount <> 0 Then cbo部位.ListIndex = 0
            cmd未记说明.ToolTipText = IIf(Val(cbo部位.Tag) = 0, "切换到未记说明", "切换到部位")
        ElseIf Shift = vbShiftMask Then
            gstrSQL = " Select 名称 From 常用体温说明 Order by 编码"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "提取未记说明")
            With rsTemp
                cbo部位.Clear
                Do While Not .EOF
                    cbo部位.AddItem !名称
                    .MoveNext
                Loop
                cbo部位.ListIndex = 0
                cbo部位.Tag = 1
            End With
        End If
        
        With cbo部位
            .Top = picInput.Height - .Height
            .Width = picInput.Width
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
        txt数据.Height = picInput.Height - cbo部位.Height
        If cbo部位.Tag = 1 Then txt数据.Text = cbo部位.Text
        cmd未记说明.ToolTipText = IIf(Val(cbo部位.Tag) = 0, "切换到未记说明", "切换到部位")
    ElseIf KeyCode = vbKeyReturn Then
        Dim strData As String
        Dim lngCol As Long
        Dim blnAllow As Boolean
        
        blnAllow = True
        If Shift = vbCtrlMask Then Exit Sub
        If picInput.Visible And txt数据.Tag <> "" Then
            lngCol = Split(txt数据.Tag, "|")(1)
            If InStr(1, "体温脉搏呼吸", Mid(vsf.TextMatrix(0, lngCol), 1, 2)) <> 0 Then
                '检查数据合法性
                If cbo部位.Tag = 0 Then
                    If txt数据.Text <> "" Then
                        strData = IIf(cbo部位.Visible And Trim(cbo部位.Text) <> "", cbo部位.Text & ":", "") & Trim(txt数据.Text)
                    End If
                Else
                    strData = IIf(Trim(txt数据.Text) <> "", Trim(txt数据.Text), cbo部位.Text)
                End If
            Else
'                mrsSelItems.Filter = "列=" & lngCol
'                If mrsSelItems.RecordCount <> 0 Then
'                    mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
'                    If mrsItems.RecordCount <> 0 Then
'                        If mrsItems!项目类型 = 1 Then
'                            strData = txt数据.Text
'                        Else
                            strData = Trim(txt数据.Text)
'                        End If
'                    End If
'                End If
'                mrsSelItems.Filter = 0
'                mrsItems.Filter = 0
            End If
            If strData <> picInput.Tag Then blnAllow = WriteIntoVsf(strData)
        End If
        
        If blnAllow Then
            Call vsf_KeyDown(vbKeyReturn, Shift)
        Else
            Call Vsf_EnterCell
            RaiseEvent AfterRowColChange(strData)
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If txt数据.SelStart = 0 Then Call vsf_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyW And Shift = vbCtrlMask Then
        If Not (cmd未记说明.Visible And cbo部位.Visible = False) Then Exit Sub
        StrText = frmWordsEditor.ShowMe(Me, mlng病人ID, mlng主页ID, txt数据.Text)
        If StrText = "" Then Exit Sub
        txt数据.Text = StrText

        DoEvents
        txt数据.SetFocus
        Call txt数据_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub txt显示天数_GotFocus()
    Call zlControl.TxtSelAll(txt显示天数)
End Sub

Private Sub txt显示天数_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If KeyCode = vbKeyReturn Then Call txt显示天数_Validate(blnCancel)
End Sub

Private Sub txt显示天数_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt显示天数_Validate(Cancel As Boolean)
    If Val(txt显示天数.Text) = Val(txt显示天数.Tag) Then Exit Sub
    txt显示天数.Tag = txt显示天数.Text
    Call ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng病区ID, cbo病人.ItemData(cbo病人.ListIndex), cbo护理等级.ListIndex, mstrPrivs, False, mblnEditable)
End Sub

Private Sub UserControl_GotFocus()
    Call Vsf_EnterCell
End Sub

Private Sub UserControl_Initialize()
    mstrSel = ""
    mstrSelItems = ""
    mblnShow = False
    mblnChange = False
    mblnInit = False
    txt显示天数.Tag = 1
    
    With cbo护理等级
        .Clear
        .AddItem "特级护理模板"
        .AddItem "一级护理模板"
        .AddItem "二级护理模板"
        .AddItem "三级护理模板"
    End With
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If InStr(1, "TXT数据,CBO部位", UCase(ActiveControl.Name)) <> 0 Then
            mblnShow = False
            picInput.Visible = False
            vsf.SetFocus
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    If mlng病人ID = 0 Then
        picNothing.Visible = True
        picNothing.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        lblNothing.Move UserControl.ScaleWidth / 2 - lblNothing.Width / 2, UserControl.ScaleHeight / 2 - lblNothing.Height
    Else
        picNothing.Visible = False
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    If mblnInit = False Then Exit Sub
    If mblnEditable = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    
    '显示当前项目的相关信息
    mrsSelItems.Filter = "列=" & NewCol
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!项目值域) <> "" Then
                If mrsItems!项目类型 = 0 Then
                    strInfo = "有效范围:" & Split(mrsItems!项目值域, ";")(0) & "～" & Split(mrsItems!项目值域, ";")(1)
                Else
                    strInfo = "有效范围:" & mrsItems!项目值域
                End If
            Else
                strInfo = ""
            End If
            
            If mrsSelItems!项目序号 = 1 Then
                strInfo = strInfo & Space(5) & "物理降温表示法:39/37.5"
            ElseIf mrsSelItems!项目序号 = 3 Then
                If mbln心率 = False Then strInfo = strInfo & Space(5) & "脉搏短拙表示法:130/120"
            ElseIf vsf.TextMatrix(0, NewCol) Like "血压*" Then
                strInfo = strInfo & Space(5) & "录入规则:收缩压/舒张压"
            End If
            
            If mrsSelItems!项目序号 >= 1 And mrsSelItems!项目序号 <= 3 Then
                strInfo = strInfo & Space(5) & "按↓进行部位选择;按SHIFT+↓进行未记说明的选择"
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    RaiseEvent AfterRowColChange(strInfo)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    blnScroll = True
    Call Vsf_EnterCell
    blnScroll = False
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call Vsf_EnterCell
End Sub

Private Sub vsf_DblClick()
    Dim blnDo As Boolean
    Dim lngOrder As Long
    
    If mblnEditable Then
        mblnShow = True
        Call Vsf_EnterCell
    Else
        If vsf.Row = 0 Then Exit Sub
        
        mrsSelItems.Filter = "列=" & vsf.Col
        blnDo = (mrsSelItems.RecordCount <> 0)
        If blnDo Then lngOrder = mrsSelItems!项目序号
        mrsSelItems.Filter = 0
        If blnDo Then RaiseEvent DbClick(lngOrder & "|" & vsf.TextMatrix(vsf.Row, vsf.Col))
    End If
End Sub

Private Sub Vsf_EnterCell()
    Dim arrData
    Dim strData As String
    Dim intIndex As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnAllow As Boolean, blnWords As Boolean
    Dim intCol As Integer, intMax As Integer
    
    If mblnInit = False Then Exit Sub
    Call ShowSignMarker
    
    '如果已录入数据则保存
    blnAllow = True
    If picInput.Visible And txt数据.Tag <> "" Then
        lngRow = Split(txt数据.Tag, "|")(0)
        lngCol = Split(txt数据.Tag, "|")(1)
        If InStr(1, "体温脉搏呼吸", Mid(vsf.TextMatrix(0, lngCol), 1, 2)) <> 0 Then
            '检查数据合法性
            If cbo部位.Tag = 0 Then
                If txt数据.Text <> "" Then
                    strData = IIf(cbo部位.Visible And Trim(cbo部位.Text) <> "", cbo部位.Text & ":", "") & txt数据.Text
                End If
            Else
                strData = IIf(Trim(txt数据.Text) <> "", Trim(txt数据.Text), cbo部位.Text)
            End If
        Else
'            mrsSelItems.Filter = "列=" & lngCol
'            If mrsSelItems.RecordCount <> 0 Then
'                mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
'                If mrsItems.RecordCount <> 0 Then
'                    If mrsItems!项目类型 = 1 Then
'                        strData = txt数据.Text
'                    Else
                        strData = Trim(txt数据.Text)
'                    End If
'                End If
'            End If
'            mrsSelItems.Filter = 0
'            mrsItems.Filter = 0
        End If
        If strData <> picInput.Tag Then blnAllow = WriteIntoVsf(strData)
    ElseIf lvwMultiSel.Visible Then
        intMax = lvwMultiSel.ListItems.Count
        For intCol = 1 To intMax
            If lvwMultiSel.ListItems(intCol).Checked Then
                strData = strData & IIf(strData = "", "", ",") & lvwMultiSel.ListItems(intCol).Text
            End If
        Next
        blnAllow = WriteIntoVsf(strData)
    End If
    Call vsf.AutoSize(0, vsf.Cols - 1)
    picInput.Visible = False
    lvwMultiSel.Visible = False
    If blnAllow = False Then
        If vsf.Row <> lngRow Then vsf.Row = lngRow
        If vsf.Col <> lngCol Then vsf.Col = lngCol
        RaiseEvent AfterRowColChange(strData)
        Exit Sub
    End If
    
    RaiseEvent AfterSelChange(IIf(Trim(vsf.TextMatrix(vsf.Row, mlngSigner)) <> "", 1, 0), vsf.TextMatrix(vsf.Row, mlngCertLevel))
    
    mblnCheckVersion = CheckVersion
    If InStr(1, mstrPrivs, "护理记录登记") = 0 Then Exit Sub
    If mblnShow = False Or IsPigeonhole Or Not mblnEditable Then Exit Sub
    If vsf.Col = 0 Or vsf.Row = 0 Then Exit Sub
    If vsf.Col = mlngOper And mblnCheckVersion = False Then Exit Sub
    If vsf.Col >= mlngSigner Then Exit Sub          '签名人,签名时间以及组号不允许编辑,组号隐藏
    If vsf.RowIsVisible(vsf.Row) = False Then Exit Sub
    If Not blnScroll And vsf.Visible Then vsf.SetFocus
    
    '准备显示
    With picInput
        .Tag = vsf.TextMatrix(vsf.Row, vsf.Col)             '保存编辑前的数据
        .Left = vsf.ColPos(vsf.Col) + vsf.Left
        .Top = vsf.RowPos(vsf.Row) + vsf.Top
        .Width = vsf.ColWidth(vsf.Col)
        .FontName = vsf.FontName
        .FontSize = vsf.FontSize
        If vsf.Row = vsf.Rows - 1 Then
            .Height = vsf.ROWHEIGHT(vsf.Row)    '取其行高
        Else
            .Height = vsf.RowPos(vsf.Row + 1) - vsf.RowPos(vsf.Row)
        End If
        If .Height > vsf.RowHeightMax Then .Height = vsf.RowHeightMax
        If .Height < vsf.RowHeightMin Then .Height = vsf.RowHeightMin
        .ZOrder 0
        .Visible = True
    End With
    With cbo部位
        .FontName = vsf.FontName
        .FontSize = vsf.FontSize
        .Visible = False
        .Clear
        .Tag = 0
        blnAllow = True
        If Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "体温" Then
            .AddItem "腋温"
            .AddItem "口温"
            .AddItem "肛温"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "脉搏" Then
            .AddItem ""
            .AddItem "起搏器"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "呼吸" Then
            .AddItem "自主呼吸"
            .AddItem "呼吸机"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "手术" Then
            .AddItem "手术"
            .AddItem "分娩"
            .AddItem "手术分娩"
            .Visible = True
            blnAllow = False
        Else
            '定位列,如果是单选,则将值域加入下拉框
            mrsSelItems.Filter = "列=" & vsf.Col
            If mrsSelItems.RecordCount <> 0 Then
                mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!项目表示 = 2 Then
                        '单选
                        .AddItem " "
                        arrData = Split(NVL(mrsItems!项目值域), ";")
                        intMax = UBound(arrData)
                        For intCol = 0 To intMax
                            If Mid(arrData(intCol), 1, 1) = "√" Then intIndex = intCol
                            .AddItem Replace(arrData(intCol), "√", "")
                        Next
                        blnAllow = False
                    ElseIf mrsItems!项目表示 = 3 Then
                        '多选
                        picInput.Visible = False
                        lvwMultiSel.Font.Name = vsf.FontName
                        lvwMultiSel.Font.Size = vsf.FontSize
                        lvwMultiSel.Left = picInput.Left + picInput.Width - lvwMultiSel.Width
                        lvwMultiSel.Top = picInput.Top + picInput.Height
                        lvwMultiSel.Visible = True
                        If lvwMultiSel.Top + lvwMultiSel.Height > picMain.Height Then lvwMultiSel.Top = picInput.Top - lvwMultiSel.Height
                        
                        '加入数据
                        lvwMultiSel.ListItems.Clear
                        arrData = Split(NVL(mrsItems!项目值域), ";")
                        intMax = UBound(arrData)
                        For intCol = 0 To intMax
                            strData = Replace(arrData(intCol), "√", "")
                            lvwMultiSel.ListItems.Add , "K" & intCol, strData
                            If Mid(arrData(intCol), 1, 1) = "√" Then lvwMultiSel.ListItems(intCol + 1).Selected = True
                            If InStr(1, "," & vsf.TextMatrix(vsf.Row, vsf.Col) & ",", "," & strData & ",") <> 0 Then lvwMultiSel.ListItems(intCol + 1).Checked = True
                        Next
                        lvwMultiSel.Tag = vsf.Row & "|" & vsf.Col
                        lvwMultiSel.SetFocus
                    ElseIf mrsItems!项目类型 = 1 And mrsItems!项目长度 >= 200 Then
                        blnWords = True
                    End If
                End If
            End If
            mrsSelItems.Filter = 0
            mrsItems.Filter = 0
        End If
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    With txt数据
        .Enabled = blnAllow          '如果当前列是手术列,则不允许录入
        .Text = vsf.TextMatrix(vsf.Row, vsf.Col)
        If .Enabled Then
            If InStr(1, .Text, ":") <> 0 And cbo部位.ListCount Then
                With cbo部位
                    If InStr(1, txt数据.Text, ":") <> 0 Then
                        .Text = Split(txt数据.Text, ":")(0)
                    End If
                    '.Top = picInput.Height - .Height
                    .Width = picInput.Width
                    .Visible = True
                    .FontName = vsf.FontName
                    .FontSize = vsf.FontSize
                    .ZOrder 0
                End With
                .Text = Split(.Text, ":")(1)
            End If
        Else
            If .Text <> "" Then cbo部位.Text = .Text
            'If .Text = "" Then .Text = cbo部位.Text
            With cbo部位
                '.Top = picInput.Height - .Height
                .Width = picInput.Width
                .Visible = True
                .FontName = vsf.FontName
                .FontSize = vsf.FontSize
                .ZOrder 0
            End With
        End If
        .FontName = vsf.FontName
        .FontSize = vsf.FontSize
        .Width = picInput.Width
        .Height = picInput.Height - IIf(cbo部位.Visible, cbo部位.Height, 0)
        .Tag = vsf.Row & "|" & vsf.Col
    End With
    If cbo部位.Enabled Then
        cbo部位.Top = picInput.Height - cbo部位.Height
        cbo部位.Width = txt数据.Width
    End If
    
    cmd未记说明.Visible = (InStr(1, "体温脉搏呼吸", Mid(vsf.TextMatrix(0, vsf.Col), 1, 2)) <> 0) Or blnWords
    If cmd未记说明.Visible Then
        cmd未记说明.FontName = vsf.FontName
        cmd未记说明.FontSize = vsf.FontSize
        '如果是体温曲线项目,如果录入的数据不是数值型,则将标志改为1
        If InStr(1, txt数据.Text, "/") = 0 Then
            If Trim(Split(txt数据.Text & "|", "|")(0)) <> "" And Trim(Split(txt数据.Text & "|", "|")(0)) <> "不升" Then
                If Not IsNumeric(Split(txt数据.Text & "|", "|")(0)) Then
                    strData = Split(txt数据.Text & "|", "|")(0)
                    Call txt数据_KeyDown(vbKeyDown, vbShiftMask)
                    txt数据.Text = strData
                End If
            End If
        End If
        If blnWords Then
            cmd未记说明.ToolTipText = "可以按Ctrl+W调出词句编辑器"
        Else
            cmd未记说明.ToolTipText = IIf(Val(cbo部位.Tag) = 0, "切换到未记说明", "切换到部位")
        End If
        cmd未记说明.Left = txt数据.Width - cmd未记说明.Width
    End If
    
    On Error Resume Next
    If txt数据.Enabled Then
        txt数据.SetFocus
    Else
        cbo部位.SetFocus
    End If
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intStep As Integer
    
    '如果是上下左右,吃掉
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 _
        Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then    'Or KeyCode = vbKeyLeft
        Exit Sub
    End If
    If KeyCode = vbKeyLeft And (picInput.Visible = False And lvwMultiSel.Visible = False) Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        '清除当前单元格的内容
        vsf.TextMatrix(vsf.Row, vsf.Col) = ""
        cbo部位.Visible = False
        txt数据.Text = ""
        txt数据.Height = picInput.Height
    End If
    
    If KeyCode = vbKeyReturn Then
        '跳到下一个有效单元格
toNextCol:
        If vsf.Col < mlngSigner Then
            vsf.Col = vsf.Col + 1
            If vsf.Col = mlngSigner Then GoTo toNextCol
            If vsf.ColHidden(vsf.Col) Then GoTo toNextCol
        Else
toNextRow:
            If vsf.Row = vsf.Rows - 1 Then
                vsf.Rows = vsf.Rows + 1
                vsf.TextMatrix(vsf.Rows - 1, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                vsf.TextMatrix(vsf.Rows - 1, 2) = Format(zlDatabase.Currentdate, "HH:mm")
                'Vsf.Cell(flexcpAlignment, Vsf.Rows - 1, 0, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter
            End If
            vsf.Row = vsf.Row + 1
            If vsf.RowHidden(vsf.Row) Then GoTo toNextRow
            vsf.Col = 1
        End If
        If vsf.ColIsVisible(vsf.Col) = False Then
            vsf.LeftCol = vsf.Col
        End If
        If vsf.RowIsVisible(vsf.Row) = False Then
            vsf.TopRow = vsf.Row
        End If
        Exit Sub
    End If
    
    If KeyCode = vbKeyLeft Then
        '跳到上一个有效单元格
toPreCol:
        If vsf.Col > 1 Then
            vsf.Col = vsf.Col - 1
            If vsf.Col >= mlngSigner Then GoTo toPreCol
            If vsf.Col = mlngOper Then GoTo toPreCol
            If vsf.ColHidden(vsf.Col) Then GoTo toPreCol
        Else
toPreRow:
            If vsf.Row > 1 Then
                vsf.Row = vsf.Row - 1
                vsf.Col = vsf.Cols - 1
                GoTo toPreCol
            Else
                vsf.Row = 1
            End If
            If vsf.RowHidden(vsf.Row) Then GoTo toPreRow
            vsf.Col = 1
        End If
        If vsf.ColIsVisible(vsf.Col) = False Then
            vsf.LeftCol = vsf.Col
        End If
        If vsf.RowIsVisible(vsf.Row) = False Then
            vsf.TopRow = vsf.Row
        End If
        Exit Sub
    End If
    
    mblnShow = True
    Call Vsf_EnterCell
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer, intCol As Integer
    If Button <> 1 Then Exit Sub
    
    intCol = vsf.MouseCol
    intRow = vsf.MouseRow
    If intRow = 0 And intCol = 0 Then
        Call vsf.Select(0, 0, vsf.Rows - 1, vsf.Cols - 1)
    ElseIf intCol = 0 Then
        Call vsf.Select(intRow, 0, intRow, vsf.Cols - 1)
    ElseIf intRow = 0 Then
        Call vsf.Select(0, intCol, vsf.Rows - 1, intCol)
    End If
End Sub

Public Sub ArchiveMe()
    On Error GoTo errHand
    
    If mlng病人ID = 0 Or mblnMoved_HL Then Exit Sub
    If MsgBox("需要将该病人本次住院所有护理记录归档吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
        Dim strNow As String

        strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        gstrSQL = "Zl_电子护理记录_Archive(" & mlng病人ID & "," & mlng主页ID & "," & cbo病人.ItemData(cbo病人.ListIndex) & ",'" & gstrUserName & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "归档")

        mstrPigeonhole = gstrUserName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnArchiveMe()
    On Error GoTo errHand
    
    If mlng病人ID = 0 Or mblnMoved_HL Then Exit Sub
    If mstrPigeonhole <> "" Then
        If MsgBox("需要撤销该病人本次住院所有已归档护理记录吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

            gstrSQL = "Zl_电子护理记录_UnArchive(" & mlng病人ID & "," & mlng主页ID & "," & cbo病人.ItemData(cbo病人.ListIndex) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "撤销归档")
            mstrPigeonhole = ""
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SignMe()
    Dim blnSign As Boolean          '是否签名成功
    Dim strTime As String
    Dim strSignTime As String       '保证所有签名的签名时间一致,便于取消签名时按签名时间统一取消
    Dim str状态 As String           '保存签名选项,避免循环签名时不停的弹出签名窗口
    Dim intRow As Integer, intRows As Integer
    On Error GoTo errHand
    '按发生时间循环进行签名
    
    If mlng病人ID = 0 Or mblnMoved_HL Then Exit Sub
    
    intRows = vsf.Rows - 1
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    For intRow = 1 To intRows
        If vsf.TextMatrix(intRow, mlngSigner) = "" And vsf.TextMatrix(intRow, vsf.Cols - 1) = gstrUserName Then
            If strTime <> vsf.TextMatrix(intRow, 1) & " " & vsf.TextMatrix(intRow, 2) & ":00" And Val(vsf.TextMatrix(intRow, mlngRecord)) <> 0 Then
                strTime = vsf.TextMatrix(intRow, 1) & " " & vsf.TextMatrix(intRow, 2) & ":00"
                If SignName(strTime, strSignTime, str状态) = False Then Exit For
                blnSign = True
            End If
        End If
    Next
    If blnSign Then Call ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng病区ID, cbo病人.ItemData(cbo病人.ListIndex), cbo护理等级.ListIndex, mstrPrivs, False, mblnEditable)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnSignMe()
    Dim blnUnSign As Boolean
    Dim strTime As String               '记录时间
    Dim strSignTime As String           '签名时间
    Dim intRow As Integer, intRows As Integer
    Dim lng终止版本 As Long             '最大版本
    Dim blnClear As Boolean             '取消签名时是否清除该版本的数据回退到上次签名后的状态
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '不批量取消签名,只取消当前选择的记录
    
    If mlng病人ID = 0 Or mblnMoved_HL Then Exit Sub
    
    If vsf.TextMatrix(vsf.Row, mlngSigner) <> gstrUserName Then
        MsgBox "不允许取消其他操作员的签名！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    blnClear = (MsgBox("取消签名时是否该版本的数据回退到上次签名后的状态？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
'    '把同一签名时间的数据提取出来,依次取消签名
'    strSignTime = Vsf.TextMatrix(Vsf.Row, mlngSignTime)
'    gstrSQL = " Select A.发生时间 From 病人护理记录 A,病人护理内容 B" & _
'              " Where A.ID=B.记录ID And B.记录类型=5 And B.项目名称=[4]" & _
'              " And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3] And A.病人来源=2"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取消签名", mlng病人id, mlng主页id, cbo病人.itemdata(cbo病人.listindex), strSignTime)
'    With rsTemp
'        Do While Not .EOF
'            If UnSignName(Format(!发生时间, "yyyy-MM-dd HH:mm:ss"), blnClear) = False Then Exit Sub
'            blnUnSign = True
'            .MoveNext
'        Loop
'    End With
    
    If UnSignName(vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00", blnClear) = False Then Exit Sub
    Call ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng病区ID, cbo病人.ItemData(cbo病人.ListIndex), cbo护理等级.ListIndex, mstrPrivs, False, mblnEditable)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal strStart As String, ByVal strSignTime As String, str状态 As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim oSign As cEPRSign
    Dim strSource As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '检查当前是否已经签名了
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From 病人护理内容 a,病人护理记录 b Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.记录类型=5 And Nvl(a.开始版本,1)=Nvl(b.最后版本,1)"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "检查当前是否已经签名了", mlng病人ID, mlng主页ID, CDate(strStart), cbo病人.ItemData(cbo病人.ListIndex))
    If rs.BOF = False Then
        MsgBox "当前没有需要签名的信息！", vbInformation, gstrSysName
        Exit Function
    End If
        
    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.修改时间" & vbNewLine & _
             " From 病人护理内容 a,病人护理记录 b " & vbNewLine & _
             " Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.终止版本 Is Null" & vbNewLine & _
             " Order by A.项目序号"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "获取要签名的内容", mlng病人ID, mlng主页ID, CDate(strStart), cbo病人.ItemData(cbo病人.ListIndex))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    Debug.Print "签名：" & Now & vbCrLf & strSource
    If strSource = "" Then
        MsgBox "当前没有需要签名的信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '76223:刘鹏飞,2014-08-05,电子签名添加时间戳信息
    '------------------------------------------------------------------------------------------------------------------
    Set oSign = frmCaseTendSign.ShowMe(Me, mstrPrivs, strSource, mlng病人ID, mlng主页ID, mlng病区ID, str状态)
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_电子护理记录_SignName("
        gstrSQL = gstrSQL & mlng病人ID & "," & mlng主页ID & "," & cbo病人.ItemData(cbo病人.ListIndex) & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "'" & oSign.姓名 & "',"
        gstrSQL = gstrSQL & "'" & oSign.签名信息 & "',"
        gstrSQL = gstrSQL & oSign.证书ID & ","
        gstrSQL = gstrSQL & oSign.签名方式 & ",'" & oSign.时间戳 & "','" & oSign.时间戳信息 & "')"

        Call zlDatabase.ExecuteProcedure(gstrSQL, "执行签名")
        SignName = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UnSignName(ByVal strStart As String, ByVal blnClear As Boolean) As Boolean
    '******************************************************************************************************************
    '功能:
    '
    '
    '******************************************************************************************************************
    Dim strSource As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    '检查当前是否已经签名了
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select A.项目ID AS 证书ID From 病人护理内容 a,病人护理记录 b Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.记录类型=5"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "检查当前是否已经签名了", mlng病人ID, mlng主页ID, CDate(strStart), cbo病人.ItemData(cbo病人.ListIndex))
    If rs.BOF Then
        MsgBox "当前没有需要取消的签名！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '如果是电子签名,则需要验证
    '------------------------------------------------------------------------------------------------------------------
    If Val(NVL(rs!证书ID, 0)) > 0 Then
        '数字签名验证
        Err.Clear
        If gobjTendESign Is Nothing Then
            On Error Resume Next
            Set gobjTendESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err.Clear
            On Error GoTo 0
            If Not gobjTendESign Is Nothing Then Call gobjTendESign.Initialize(gcnOracle, glngSys)
        End If
        If Not gobjTendESign Is Nothing Then
            If Not gobjTendESign.CheckCertificate(gstrDBUser) Then Exit Function
        Else
            MsgBox "电子签名部件未能正确安装，回退操作不能继续！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Zl_电子护理记录_Unsignname("
    gstrSQL = gstrSQL & mlng病人ID & ","
    gstrSQL = gstrSQL & mlng主页ID & ","
    gstrSQL = gstrSQL & cbo病人.ItemData(cbo病人.ListIndex) & ","
    gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & _
                      IIf(blnClear, "1", "0") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "执行取消签名")
    
    UnSignName = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    Call ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng病区ID, cbo病人.ItemData(cbo病人.ListIndex), mbyt护理等级, mstrPrivs, True, mblnEditable)
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    mblnShow = False
    picInput.Visible = False
    
    SaveME = True
    
    Call ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng病区ID, cbo病人.ItemData(cbo病人.ListIndex), mbyt护理等级, mstrPrivs, False, mblnEditable)
End Function

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngPatiID As Long, ByVal lngPageId As Long, lngDeptId As Long, _
    Optional ByVal intBaby As Integer = 0, Optional ByVal byt护理级别 As Byte = 3, Optional ByVal strPrivs As String, _
    Optional ByVal blnCancel As Boolean = False, Optional ByVal blnEditable As Boolean = True)
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       lngDeptID           要显示护理记录的科室
    '       intBaby             婴儿标志
    '       blnEditable         如果为假,说明是做为查询子窗体在使用,取消与编辑相关的功能
    '返回： 无
    '******************************************************************************************************************
'    Dim bln护理级别 As Boolean
    
    Err = 0
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    mblnInit = False
    mblnEditable = blnEditable And Not mblnMoved_HL
    
    lngLastRow = vsf.Row
    lngLastTopRow = vsf.TopRow
    lngLastPatientID = mlng病人ID
    If lngLastRow < 1 Then lngLastRow = 1
    If lngLastTopRow < 1 Then lngLastTopRow = 1
    
    If mblnChange And Not blnCancel Then
        If MsgBox("当前病人的数据还未保存，点“是”进行保存，点“否”将放弃本次修改！", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call Vsf_EnterCell
            Call SaveData
        End If
    End If
    mblnShow = False
    picInput.Visible = False
    
    mlng病人ID = lngPatiID
    mlng主页ID = lngPageId
    mlng病区ID = lngDeptId
    mint婴儿 = intBaby
    mbyt护理等级 = byt护理级别
    mstrPrivs = strPrivs
    Set mfrmParent = frmParent
    
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd")
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitPanelMain
        Call InitEnv            '初始化环境
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    
    '提取该病人所属科室,以便后面提取模板
    Call UserControl_Resize
    If mlng病人ID = 0 Then Exit Sub
    gstrSQL = " Select 出院科室ID From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人所属科室", mlng病人ID, mlng主页ID)
    mlng科室ID = rsTemp!出院科室ID
    
    '当病人发生变化时,清除以下变量
    If lngLastPatientID <> mlng病人ID Then
        mstrSel = ""
        mstrSelItems = ""
        cbo护理等级.ListIndex = mbyt护理等级
        
        '提取病人的婴儿
        gstrSQL = " Select NVL(A.婴儿姓名,NVL(C.姓名,B.姓名) ||'之子'||A.序号) AS 姓名,A.序号" & _
                  " From 病人信息 B,病案主页 C,病人新生儿记录 A" & _
                  " Where B.病人ID=C.病人ID And A.病人ID=C.病人ID And A.主页ID=C.主页ID And C.病人ID=[1] And C.主页ID=[2]" & _
                  " Order By A.序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人的婴儿", mlng病人ID, mlng主页ID)
        With cbo病人
            .Clear
            .AddItem "病人本人"
            
            Do While Not rsTemp.EOF
                .AddItem rsTemp!姓名
                .ItemData(.NewIndex) = rsTemp!序号
                rsTemp.MoveNext
            Loop
        End With
    End If
    cbo病人.ListIndex = mint婴儿
    
    Call InitBill
    Call ReadData
    mblnInit = True
    
    '恢复定位
    If lngLastPatientID <> mlng病人ID Then
        lngLastRow = 1
        lngLastTopRow = 1
    End If
    
    cbo病人.Tag = cbo病人.ListIndex
    If vsf.Rows - 1 > lngLastRow Then vsf.Row = lngLastRow
    If vsf.RowIsVisible(vsf.Row) Then vsf.TopRow = lngLastTopRow
    Call Vsf_EnterCell
    Call ReSetFontSize(mbytFontSize)
    mblnChange = False
    RaiseEvent AfterRefresh
    
    'Call OutputRsData(mrsSelItems)
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckData() As Boolean
    Dim StrText As String
    Dim strMaxDate As String, str值域 As String
    Dim lngRow As Long, lngRows As Long, lngCol As Long
    Dim intType As Integer, lngOrder As Long, lngClass As Long, strName As String, lngLength As Long
    On Error GoTo errHand
    '检查数据录入合法性
    
    lngRows = vsf.Rows - 1
    '先检查日期是否合法
    For lngRow = 1 To lngRows
        If Val(vsf.RowData(lngRow)) = 1 Then
            If Not CheckDate1(lngRow) Then
                vsf.Row = lngRow
                vsf.Col = 1
                If vsf.RowIsVisible(vsf.Row) Then vsf.TopRow = vsf.Row
                Exit Function
            End If
        End If
    Next
    
    '依次检查各个项目的录入合法性
    With mrsSelItems
        .MoveFirst
        Do While Not .EOF
            mrsItems.Filter = "项目序号=" & !项目序号
            If mrsItems.RecordCount <> 0 Then
                lngCol = !列
                intType = mrsItems!项目类型     '0-数值；1-文字
                lngClass = mrsItems!项目性质
                lngOrder = mrsItems!项目序号
                strName = mrsItems!项目名称
                lngLength = mrsItems!项目长度 + IIf(NVL(mrsItems!项目小数, 0) = 0, 0, NVL(mrsItems!项目小数, 0) + 1)
                If intType = 0 Then
                    str值域 = NVL(mrsItems!项目值域)
                Else
                    str值域 = ""
                End If
                '数值项目:只有体温,呼吸与脉搏,以及血压才存在/录入
                '文本项目:只检查是否超长
                If Not (intType = 1 And lngLength > 100) Then
                    For lngRow = 1 To lngRows
                        If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                            StrText = vsf.TextMatrix(lngRow, lngCol)
                            If Trim(StrText) <> "" Then
                                If Not CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, str值域) Then
                                    vsf.Row = lngRow
                                    If vsf.RowIsVisible(vsf.Row) Then vsf.TopRow = vsf.Row
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    mrsItems.Filter = 0
    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsItems.Filter = 0
End Function

Private Function CheckDate1(ByVal lngRow As Long) As Boolean
    If Not IsDate(vsf.TextMatrix(lngRow, 1)) Then
        MsgBox "日期格式错误：yyyy-MM-dd", vbInformation, gstrSysName
        Exit Function
    End If
    If vsf.TextMatrix(lngRow, 1) > mstrMaxDate Then
        MsgBox "第" & lngRow & "行的日期大于了参数[超期录入天数：" & mintPreDays & "天]所指定的范围！", vbInformation, gstrSysName
        Exit Function
    End If
    If Trim(vsf.TextMatrix(lngRow, 2)) = "" Then
        MsgBox "第" & lngRow & "行的时间不能为空！", vbInformation, gstrSysName
        Exit Function
    End If
    If Len(vsf.TextMatrix(lngRow, 2)) = 2 Then vsf.TextMatrix(lngRow, 2) = vsf.TextMatrix(lngRow, 2) & ":00"
    
    CheckDate1 = True
End Function

Private Function CheckDate2(ByVal lngRow As Long, ByVal lngCol As Long, StrText As String, Optional ByRef strInfo As String = "") As Boolean
    Dim strMsg As String
    Dim strDate As String
    Dim rsTemp As New ADODB.Recordset
    
    '不能小于入院日期,时间不足补位时,要考虑是否合法
    If lngCol = 1 Then
        gstrSQL = " Select 入院日期 From 病案主页 Where 病人ID=" & mlng病人ID & " And 主页ID=" & mlng主页ID
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "取入院日期")
        strDate = Format(rsTemp!入院日期, "yyyy-MM-dd")
        
        If Not IsDate(StrText) Then
            strMsg = "日期格式错误：yyyy-MM-dd"
            GoTo errHand
        End If
        If StrText < strDate Then
            strMsg = "第" & lngRow & "行的日期小于了入院日期！"
            GoTo errHand
        End If
        If StrText > mstrMaxDate Then
            strMsg = "第" & lngRow & "行的日期大于了参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
            GoTo errHand
        End If
    End If
    
    If lngCol = 2 Then
        If Trim(StrText) = "" Then
            strMsg = "第" & lngRow & "行的时间不能为空！"
            GoTo errHand
        End If
        If Len(StrText) <= 2 Then StrText = String(2 - Len(StrText), "0") & StrText
        If Val(Mid(StrText, 1, 2)) < 0 Or Val(Mid(StrText, 1, 2)) > 23 Then
            strMsg = "第" & lngRow & "行的时间数据非法！"
            GoTo errHand
        End If
        If Len(StrText) = 2 Then StrText = StrText & ":00"
        If Len(StrText) < 5 And InStr(1, StrText, ":") > 0 Then StrText = String(5 - Len(StrText), "0") & StrText
        If Mid(StrText, 3, 1) <> ":" Then
            strMsg = "第" & lngRow & "行的时间数据格式非法[09:00]！"
            GoTo errHand
        End If
        If Len(StrText) < 5 Then StrText = StrText & String(5 - Len(StrText), "0")
        If Not (Val(Mid(StrText, 4, 2)) >= 0 And Val(Mid(StrText, 4, 2)) <= 59) Then
            strMsg = "第" & lngRow & "行的时间数据格式非法[09:00]！"
            GoTo errHand
        End If
        vsf.TextMatrix(lngRow, 2) = StrText
    
        '数据发生时间不能在当前操作员所属科室的有效时间以前
        If Not CheckTime(lngRow, mlng病人ID, mlng主页ID, vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), strMsg) Then
            GoTo errHand
        End If
    End If
    CheckDate2 = True
    Exit Function
errHand:
    If strInfo = "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        strInfo = strMsg
    End If
End Function

Private Function CheckValid(ByRef StrText As String, ByVal lngOrder As Long, ByVal lngType As Long, ByVal strCap As String, _
    ByVal lngLength As Long, ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal str值域 As String, _
    Optional ByRef strInfo As String = "") As Boolean
    Dim arrData
    Dim strMsg As String
    Dim intDo As Integer, intCount As Integer
    Dim strPart As String, strValue1 As String, strValue2 As String, strTextClone As String
    
    If StrText = "" Then
        CheckValid = True
        Exit Function
    End If
    
    '先取出部位,上/下数据
    strTextClone = StrText
    If InStr(1, strTextClone, ":") <> 0 Then
        strPart = Split(strTextClone, ":")(0)
        strTextClone = Split(strTextClone, ":")(1)
    End If
    If InStr(1, strTextClone, "/") <> 0 Then
        strValue1 = Split(strTextClone, "/")(0)
        strValue2 = Split(strTextClone, "/")(1)
    Else
        strValue1 = strTextClone
    End If
    
    If lngType = 2 Then '如果是活动项目则可能存在部位,把部位提出来,只检查录入的数据是否超过限制
        If InStr(1, StrText, ":") <> 0 Then
            StrText = Split(StrText, ":")(1)
        End If
    End If
    
'    If str值域 = "" Then  '普通项目
'        If Not (lngOrder = 9 Or lngOrder = 10) Then '大便次数与排出量不进行有效范围检查
'            If LenB(StrConv(strText, vbFromUnicode)) > lngLength Then
'                strMsg = "第" & lngRow & "行的" & strCap & "超长，请检查！"
'                GoTo errHand
'            End If
'        End If
'    Else                    '体温脉搏呼吸以及血压
        '没有心率的时候，才允许录入脉搏
        If lngOrder = 2 And mbln心率 Then
            If InStr(1, StrText, "/") <> 0 Then
                strMsg = "请将测得的心率数据录入单独的心率单元格中！"
                GoTo errHand
            End If
        End If
        If lngOrder = 3 Then
            If InStr(1, StrText, "/") <> 0 Then
                strMsg = "呼吸数据录入错误！"
                GoTo errHand
            End If
        End If
        If lngOrder = 4 Or lngOrder = 5 Then
            '血压值必须含/
            If vsf.TextMatrix(0, lngCol) Like "血压*" Then
                If InStr(1, StrText, "/") = 0 Then
                    strMsg = "血压数据的格式错误：收缩压/舒张压！"
                    GoTo errHand
                End If
                If Trim(Split(StrText, "/")(0)) = "" Or Trim(Split(StrText, "/")(1)) = "" Then
                    strMsg = "血压数据错误：收缩压/舒张压！"
                    GoTo errHand
                End If
            End If
        End If
        If UBound(Split(StrText, "/")) > 1 Then
            strMsg = "第" & lngRow & "行的" & strCap & "数据录入错误，请检查！"
            GoTo errHand
        End If
        
        arrData = Split(StrText, "/")
        intCount = UBound(arrData)
        For intDo = 0 To intCount
            StrText = arrData(intDo)
            If InStr(1, StrText, ":") <> 0 Then StrText = Split(StrText, ":")(1)
            '曲线项目不检查是否超长
            If lngOrder > 3 Then
                If LenB(StrConv(StrText, vbFromUnicode)) > lngLength Then
                    strMsg = "第" & lngRow & "行的" & strCap & "超长，请检查！"
                    vsf.TopRow = lngRow
                    GoTo errHand
                End If
            End If
            If IsNumeric(StrText) Then    '有效范围与当前录入值都是数值型才检查,否则当成是未记说明
                If Not (lngOrder = 9 Or lngOrder = 10) Then '大便次数与排出量不进行有效范围检查
                    If str值域 <> "" Then
                        If IsNumeric(Split(str值域, ";")(0)) Then
                            If Not (Val(StrText) >= Split(str值域, ";")(0) And Val(StrText) <= Split(str值域, ";")(1)) Then
                                strMsg = "第" & lngRow & "行的" & strCap & "超出有效范围（" & Split(str值域, ";")(0) & "-" & Split(str值域, ";")(1) & "），请检查！"
                                GoTo errHand
                            End If
                        End If
                    End If
                    If mrsItems!项目类型 = 0 Then
                        If NVL(mrsItems!项目小数, 0) <> 0 Then
                            If intDo = 0 Then
                                strValue1 = Format(StrText, "#0." & String(mrsItems!项目小数, "0"))
                            Else
                                strValue2 = Format(StrText, "#0." & String(mrsItems!项目小数, "0"))
                            End If
                        Else
                            If intDo = 0 Then
                                strValue1 = Format(StrText, "#0")
                            Else
                                strValue2 = Format(StrText, "#0")
                            End If
                        End If
                    End If
                End If
            End If
        Next
'    End If
    
    '拼装输入串
    StrText = IIf(strPart <> "", strPart & ":", "") & strValue1 & IIf(strValue2 <> "", "/" & strValue2, "")
    
    CheckValid = True
    Exit Function
errHand:
    If strInfo = "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        strInfo = strMsg    '输出错误信息,由调用程序处理
    End If
End Function

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean, blnOper As Boolean         '指定某个时间段里是否出现手术
    Dim lngOrder As Long
    Dim strTime As String, strTmp As String, strSQLtmp As String, strMsg As String
    Dim intAllow As Integer, intType As Integer, lngClass As Long
    Dim str内容 As String, str标记 As String, str部位 As String, str未记说明 As String 'str标记:只保存特殊降温或脉搏短拙
    Dim lngRecord As Long, lngGroup As Long, lngMAX As Long
    Dim lngRow As Long, lngRows As Long, lngCol As Long, lngCols As Long
    Dim strDate As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim intPos As Integer, intMax As Integer
    Dim strSQL() As String
    On Error GoTo errHand
    '同一个时间里(同一条记录ID),不允许出现多组手术,也就是只允许一个组号里有手术的存在
    '如果记录ID=0,新增的记录,其时间点已存在历史记录的,则产生最大的组号保存
    
    If mblnMoved_HL Then Exit Function
    
    ReDim Preserve strSQL(1 To 1)
    lngRows = vsf.Rows - 1
    lngCols = mlngSigner - 1         '后面的签名人,签名时间,记录ID,组号不处理
    intAllow = IIf(InStr(mstrPrivs, "他人护理记录") > 0, 1, 0)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '1、如果时间有变化，先处理时间
    For lngRow = 1 To lngRows
        '数据发生时间不能在当前操作员所属科室的有效时间以前
        strMsg = "msgbox"
        If Val(vsf.RowData(lngRow)) = 1 Then
            If Not CheckTime(lngRow, mlng病人ID, mlng主页ID, vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2), Mid(strDate, 1, 16), strMsg) Then
                Exit Function
            End If
        End If
        
        If Val(vsf.TextMatrix(lngRow, mlngRecord)) <> 0 And (vsf.Cell(flexcpData, lngRow, 1) = 1 Or vsf.Cell(flexcpData, lngRow, 2) = 1) Then
            If lngRecord <> Val(vsf.TextMatrix(lngRow, mlngRecord)) Then
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                strStart = vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2) & ":00"
                gstrSQL = "Zl_病人护理记录_UpdateReplace(" & lngRecord & ",0," & cbo病人.ItemData(cbo病人.ListIndex) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
                strSQL(ReDimArray(strSQL)) = gstrSQL
            End If
        End If
    Next
    
    '2、再依次处理编辑过的元素
    For lngRow = 1 To lngRows
        '先定位修改过的行,再在列中循环找到修改过的列
        If vsf.TextMatrix(lngRow, 1) <> "" And vsf.TextMatrix(lngRow, 2) <> "" Then
            If strTime <> vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2) Then
                strTime = vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2)
                blnOper = False
            End If
            
            strDate = strTime
            strStart = strDate & ":00"
            strEnd = Format(DateAdd("n", 1, CDate(strDate)), "yyyy-MM-dd HH:mm") & ":00"
            
            If Val(vsf.RowData(lngRow)) = 1 Then
                '有组号则取组号，无组号，则取当前最大组号
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                lngGroup = Val(vsf.TextMatrix(lngRow, mlngGroup))
                '有可能原来的数据中的组号不是按顺序增加的,因此此段进行校正
                If lngGroup = 0 Then
                    '取最大的组号
                    gstrSQL = " select max(记录组号) AS 组号 " & _
                              " From 病人护理内容" & _
                              " where 记录ID=(" & _
                              "     select ID from 病人护理记录" & _
                              "     where 病人ID=[1] and 主页ID=[2] and 婴儿=[3] and 科室ID=[4] and 发生时间=[5])"
                    If mblnMoved_HL Then
                        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
                        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取最大的组号", mlng病人ID, mlng主页ID, cbo病人.ItemData(cbo病人.ListIndex), mlng科室ID, CDate(strStart))
                    lngGroup = NVL(rsTemp!组号, 0) + 1
                End If
                
                '一个元素一个元素的处理
                For lngCol = 3 To lngCols
                    If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                        '此数据进行了新增或修改操作
                        gstrSQL = "Zl_病人护理记录_UpdateRecord("
                        gstrSQL = gstrSQL & mlng病人ID & "," & mlng主页ID & "," & cbo病人.ItemData(cbo病人.ListIndex) & ","
                        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & IIf(lngCol <> mlngOper, 1, 4) & ","
                        
                        lngOrder = 0
                        If lngCol <> mlngOper Then
                            mrsSelItems.Filter = "列=" & lngCol
                            mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
                            intType = mrsItems!项目类型
                            lngClass = mrsItems!项目性质
                            lngOrder = mrsItems!项目序号
                        End If
                        strSQLtmp = gstrSQL     '为血压单独处理
                        gstrSQL = gstrSQL & lngOrder & ","
                        
                        str部位 = "": str标记 = "": str未记说明 = ""
                        str内容 = vsf.TextMatrix(lngRow, lngCol)
                        If (lngOrder = 1 Or lngOrder = 2 Or lngOrder = 3) Or lngClass = 2 Then
                            If InStr(1, str内容, ":") <> 0 Then
                                str部位 = Trim(Split(str内容, ":")(0))
                                str内容 = Trim(Split(str内容, ":")(1))
                            End If
                            If InStr(1, str内容, "/") <> 0 Then
                                str标记 = Trim(Split(str内容, "/")(1))
                                str内容 = Trim(Split(str内容, "/")(0))
                            End If
                        ElseIf lngOrder = 4 Then        '因为是按列循环,所以只会处理一次,如果是合并录入收缩压与舒张压,则在保存后再处理下
                            If InStr(1, str内容, "/") <> 0 Then
                                str内容 = Split(str内容, "/")(lngOrder - 4)
                            End If
                        End If
                        '只有曲线项目才存在未记说明的概念
                        If lngOrder <= 3 And Not IsNumeric(str内容) And lngCol <> mlngOper Then
                            If (lngOrder = 1 And str内容 <> "不升") Or lngOrder <> 1 Then
                                str未记说明 = str内容
                                str内容 = ""
                            End If
                        End If
                        
                        '体温脉搏项目,如果有/填1
                        If lngOrder = -1 Then
                            gstrSQL = gstrSQL & "1,"
                        Else
                            gstrSQL = gstrSQL & "0,"
                        End If
                        
                        If lngCol <> mlngOper Or blnOper = False Then
                            gstrSQL = gstrSQL & "'" & str内容 & "','" & str部位 & "'," & intAllow & "," & IIf(IsNumeric(str内容), 0, 1) & "," & lngGroup & ",'" & str未记说明 & "')"
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                            
                            '如果是血压
                            If lngOrder = 4 And vsf.TextMatrix(0, lngCol) Like "血压*" Then
                                If str内容 <> "" Then str内容 = Split(vsf.TextMatrix(lngRow, lngCol), "/")(1)       '不为空时进行赋值,为空则说明现在是清除数据
                                strSQLtmp = strSQLtmp & "5,0,"
                                gstrSQL = strSQLtmp & "'" & str内容 & "','" & str部位 & "'," & intAllow & "," & IIf(IsNumeric(str内容), 0, 1) & "," & lngGroup & ",'" & str未记说明 & "')"
                                strSQL(ReDimArray(strSQL)) = gstrSQL
                            End If
                            
                            If lngCol = mlngOper Then blnOper = True
                        End If
                        
                        '----------------------------------------------------------------------------
                        '没有选择心率,就允许他在脉搏处同时录入(如果都为空,完成标记部分数据清除的功能)
                        If (lngOrder = 1 Or lngOrder = 2 And mbln心率 = False) Then
            
                            gstrSQL = "Zl_病人护理记录_UpdateRecord("
                            gstrSQL = gstrSQL & mlng病人ID & "," & mlng主页ID & "," & cbo病人.ItemData(cbo病人.ListIndex) & ","
                            gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                            gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            gstrSQL = gstrSQL & "1,"
                            gstrSQL = gstrSQL & IIf(lngOrder = 2, -1, lngOrder) & ","
                            gstrSQL = gstrSQL & "1,"
                                                            
                            If str标记 <> "" And str内容 <> "" Then
                                Select Case intType
                                Case 0          '数值
                                    strTmp = Val(str标记)
                                Case 1          '文本
                                    strTmp = str标记
                                End Select
                                gstrSQL = gstrSQL & "'" & strTmp & "','" & str部位 & "'," & intAllow & "," & IIf(IsNumeric(strTmp), 0, 1) & "," & lngGroup & ",Null)"
                            Else
                                gstrSQL = gstrSQL & "NULL,'" & str部位 & "'," & intAllow & ",0," & lngGroup & ",Null)"
                            End If
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    '循环执行SQL保存数据
    gcnOracle.BeginTrans
    blnTrans = True
    intMax = UBound(strSQL)
    For intPos = 1 To intMax
        If strSQL(intPos) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(intPos), "保存体温数据")
    Next
    SaveData = True
    gcnOracle.CommitTrans
    blnTrans = False
    
    mblnChange = False
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
    
    RaiseEvent AfterDataChanged
    RaiseEvent AfterRefresh
    Exit Function
    
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
End Function


'---------------------------------------------------------------------------------
'以下是基础函数或过程
'---------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名,值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    With rsObj
        Do While Not .EOF
            Debug.Print !列 & "," & !项目序号 & "," & !项目名称
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function CheckVersion(Optional ByVal lngRow As Long = 0, Optional ByVal lngCol As Long = 0) As Boolean
    Dim lng项目序号 As Long
    Dim lng当前版本 As Long
    Dim lng最高版本 As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '因手术只产生一条记录,只有当签名记录的最大版本小于手术数据的开始版本时,才允许进行编辑(含清除功能)
    '如果要清除一行,该行存在手术记录,如果不允许对手术列进行编辑,则取消该操作
    
    If lngRow = 0 Then lngRow = vsf.Row
    If lngCol = 0 Then lngCol = vsf.Col
    If Val(vsf.TextMatrix(lngRow, mlngRecord)) = 0 Then CheckVersion = True: Exit Function      '新记录直接退出
    If vsf.Cell(flexcpData, lngRow, lngCol) <> 0 Then CheckVersion = True: Exit Function                              '本次新增的数据允许清除
    
    '取当前单元格的项目序号
    mrsSelItems.Filter = "列=" & lngCol
    If mrsSelItems.RecordCount <> 0 Then
        lng项目序号 = mrsSelItems!项目序号
    Else
        mrsSelItems.Filter = 0
        Exit Function
    End If
    mrsSelItems.Filter = 0
    
    '取当前记录+组号的最大版本
    gstrSQL = " Select Max(开始版本) AS 最高版本 From 病人护理内容 Where 记录ID=[1] And 记录类型=5"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前记录+组号的最大版本", Val(vsf.TextMatrix(lngRow, mlngRecord)), Val(vsf.TextMatrix(lngRow, mlngGroup)))
    lng最高版本 = NVL(rsTemp!最高版本, 0)
    
    '取当前项目的当前版本
    gstrSQL = " Select MAX(开始版本) AS 当前版本 From 病人护理内容 Where 记录ID=[1] And 记录组号=[2]" & IIf(lngCol = mlngOper, " And 记录类型=4", " And 项目序号=[3]")
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前记录+组号的最大版本", Val(vsf.TextMatrix(lngRow, mlngRecord)), Val(vsf.TextMatrix(lngRow, mlngGroup)), lng项目序号)
    lng当前版本 = NVL(rsTemp!当前版本, 1)
    
    '只有当前版本大于最高版本,才允许清除(签名的数据也不允许清除)
    '同时如果最高版本=1,且签名人为空,也允许清除
    CheckVersion = ((lng当前版本 > lng最高版本) Or (lng最高版本 = 1 And vsf.Cell(flexcpForeColor, lngRow, lngCol) = &HFF&))
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '数据发生时间必须在当前科室的有效时间范围内
    
    blnMsg = (strMsg <> "")
    gstrSQL = " Select 开始原因,病区ID,to_char(开始时间,'yyyy-MM-dd hh24:mi') AS 开始时间,to_char(NVL(终止时间,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS 终止时间 " & _
              " From 病人变动记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]" & _
              " Order by 开始时间,开始原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前科室有效时间范围", lng病人ID, lng主页ID)
    With rsTemp
        .Filter = "病区ID=" & mlng病区ID
        Do While Not .EOF
            If strTime >= !开始时间 And strTime <= !终止时间 Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '找到了就退出
        If blnExist Then
            If Not IsAllowInput(lng病人ID, lng主页ID, strTime, strCurTime) Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        '没找到,就整理原因进行准确性提示
        .Filter = "开始原因=1"
        If .RecordCount <> 0 Then
            If !开始原因 = 1 And strTime < !开始时间 Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入院时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=2"
        If .RecordCount <> 0 Then
            If !开始原因 = 2 And strTime < !开始时间 Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入科时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=10"
        If .RecordCount <> 0 Then
            If !开始原因 = 10 And strTime > !终止时间 Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & !终止时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '其他情况说明
        strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[不在当前病区的有效时间范围内]"
        GoTo exitHand
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Sub imgSign_Click()
    Call picSign_Click
End Sub

Private Sub lbl验证签名_Click()
    Call picSign_Click
End Sub

Private Sub picSign_Click()
    '加载签名历史记录
    Dim str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    vsfSignData.Clear
    str发生时间 = vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00"
    gstrSQL = "" & _
        " SELECT A.记录人 AS 签名人,NVL(to_char(A.修改时间,'yyyy-MM-dd hh24:mi:ss'),A.项目名称) AS 签名时间,A.记录内容 AS 签名信息,A.记录标记 AS 签名规则,A.ID,DECODE(A.项目ID,NULL,'有效','未验证') AS 有效性,A.开始版本,A.项目序号 AS 签名规则版本" & vbNewLine & _
        " FROM 病人护理内容 A,病人护理记录 B" & vbNewLine & _
        " WHERE A.记录ID=B.ID AND A.记录类型=5" & vbNewLine & _
        " AND B.病人ID=[1] AND B.主页ID=[2] AND B.婴儿=[3] AND B.发生时间=[4] " & vbNewLine & _
        " Order by A.项目名称 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名历史记录", mlng病人ID, mlng主页ID, mint婴儿, CDate(str发生时间))
    
    Set vsfSignData.DataSource = rsTemp
    With vsfSignData
        .ColWidth(0) = 1000
        .ColWidth(1) = 1800
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .Row = 1
        .Col = 5
    End With
    
    picSign.Visible = False
    With picSignCheck
        .Left = vsf.Left + (vsf.Width - .Width) / 2
        .Top = vsf.Top + (vsf.Height - .Height) / 2
        .Visible = True
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    picSignCheck.Visible = False
End Sub

Private Sub cmdSignCur_Click()
    '单行验证
    Dim lngLoop As Long
    Dim int版本 As Integer
    Dim strSource As String, str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If (Val(vsfSignData.TextMatrix(vsfSignData.Row, 4)) = 0) Then Exit Sub
    If (Val(vsfSignData.TextMatrix(vsfSignData.Row, 7)) < 2) Then
        MsgBox "由于签名规则变化，老版签名数据暂不支持签名校验功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
    int版本 = vsfSignData.TextMatrix(vsfSignData.Row, 6)
    str发生时间 = vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00"
    Set rsTemp = GetSignData(str发生时间, int版本)
    Do While Not rsTemp.EOF
        For lngLoop = 0 To rsTemp.Fields.Count - 1
            strSource = strSource & CStr(zlCommFun.NVL(rsTemp.Fields(lngLoop).Value, ""))
        Next
        rsTemp.MoveNext
    Loop
    Debug.Print "验证签名：" & Now & vbCrLf & strSource
    
    '数字签名
    Err.Clear
    If gobjTendESign Is Nothing Then
        On Error Resume Next
        Set gobjTendESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        On Error GoTo 0
        If Not gobjTendESign Is Nothing Then
            Call gobjTendESign.Initialize(gcnOracle, glngSys)
        End If
    End If
    If gobjTendESign Is Nothing Then
        MsgBox "电子签名部件未能正确安装，验证操作不能继续！", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjTendESign.VerifySignature(strSource, Val(vsfSignData.TextMatrix(vsfSignData.Row, 4)), 5) Then
        vsfSignData.TextMatrix(vsfSignData.Row, 5) = "有效"
        Call vsfSignData_EnterCell
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSignAll_Click()
    Dim lngSel As Long
    Dim lngRow As Long, lngRows As Long
    '全部验证
    
    lngSel = vsfSignData.Row
    vsfSignData.Redraw = flexRDNone
    lngRows = vsfSignData.Rows - 1
    For lngRow = 1 To lngRows
        vsfSignData.Row = lngRow
        Call cmdSignCur_Click
    Next
    vsfSignData.Row = lngSel
    vsfSignData.Redraw = flexRDDirect
End Sub

Private Function ShowSignMarker(Optional ByVal bln外部 As Boolean = False) As Boolean
    Dim str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    '显示历史签名标记
    
    picSign.Visible = False
    picSignCheck.Visible = False
    If Not bln外部 Then
        If vsf.Col <> mlngSigner Then Exit Function
    End If
    If vsf.TextMatrix(vsf.Row, mlngSigner) = "" Then Exit Function
    
    str发生时间 = vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00"
    gstrSQL = "" & _
        " SELECT A.记录人 AS 签名人,NVL(to_char(A.修改时间,'yyyy-MM-dd hh24:mi:ss'),A.项目名称) AS 签名时间,A.记录内容 AS 签名信息,A.记录标记 AS 签名规则,A.ID" & vbNewLine & _
        " FROM 病人护理内容 A,病人护理记录 B" & vbNewLine & _
        " WHERE A.记录ID=B.ID AND A.记录类型=5" & vbNewLine & _
        " AND B.病人ID=[1] AND B.主页ID=[2] AND B.婴儿=[3] AND B.发生时间=[4] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名历史记录", mlng病人ID, mlng主页ID, mint婴儿, CDate(str发生时间))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    With picSign
        .Top = vsf.Top + vsf.CellTop + vsf.CellHeight - .Height
        .Left = vsf.Left + vsf.CellLeft + 500
        .Visible = True
    End With
    ShowSignMarker = True
End Function

Private Sub vsfSignData_EnterCell()
    cmdSignCur.Enabled = (vsfSignData.TextMatrix(vsfSignData.Row, 5) <> "有效")
End Sub

Private Function GetSignData(ByVal str发生时间 As String, ByVal int版本 As Integer) As ADODB.Recordset
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    If int版本 = 1 Then
        gstrSQL = "" & _
            "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.修改时间" & vbNewLine & _
            "  From 病人护理内容 a, 病人护理记录 b" & vbNewLine & _
            " Where b.病人id = [1] And b.主页id = [2] And B.婴儿=[3] And b.发生时间 =[4]" & vbNewLine & _
            "   And a.记录id = b.ID and A.记录类型 <>5 and A.开始版本=1" & vbNewLine & _
            " ORDER BY 项目序号"
    Else
        gstrSQL = "" & _
            "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.修改时间" & vbNewLine & _
            "  From 病人护理内容 a, 病人护理记录 b" & vbNewLine & _
            " Where b.病人id = [1] And b.主页id = [2] And B.婴儿=[3] And b.发生时间 =[4]" & vbNewLine & _
            "   And a.记录id = b.ID and A.记录类型 <>5" & vbNewLine & _
            "   and (A.开始版本=[5] or (A.开始版本 <[5] and A.终止版本 IS NULL) or (A.开始版本<[5] and A.终止版本>[5]))" & vbNewLine & _
            " ORDER BY 项目序号"
    End If
    Set GetSignData = zlDatabase.OpenSQLRecord(gstrSQL, "提取指定版本的数据", mlng病人ID, mlng主页ID, mint婴儿, CDate(str发生时间), int版本)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SignMarker()
    '供外部主程序调用
    If Not ShowSignMarker(True) Then Exit Sub
    Call picSign_Click
End Sub
