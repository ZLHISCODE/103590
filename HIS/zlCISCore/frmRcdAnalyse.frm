VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form frmRcdAnalyse 
   Caption         =   "病人病史分析"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8100
   Icon            =   "frmRcdAnalyse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8100
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   -15
      TabIndex        =   20
      Top             =   1335
      Width           =   6450
   End
   Begin VB.TextBox txtPati 
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   900
      MaxLength       =   11
      TabIndex        =   1
      ToolTipText     =   "请按""-病人ID""、""+住院号""、""*门诊号""形式输入或直接输入姓名查找"
      Top             =   150
      Width           =   900
   End
   Begin VB.OptionButton optChart 
      Appearance      =   0  'Flat
      Caption         =   "图形对比(&G)(仅数值项目可进行图形对比)"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3960
      TabIndex        =   19
      Top             =   3885
      Width           =   3705
   End
   Begin VB.OptionButton optChart 
      Appearance      =   0  'Flat
      Caption         =   "数据表格(&T)"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   2580
      TabIndex        =   18
      Top             =   3885
      Value           =   -1  'True
      Width           =   1290
   End
   Begin MSChart20Lib.MSChart chtItem 
      Height          =   1260
      Left            =   2850
      OleObjectBlob   =   "frmRcdAnalyse.frx":08CA
      TabIndex        =   17
      Top             =   4230
      Width           =   5490
   End
   Begin VSFlex8Ctl.VSFlexGrid hgdDiag 
      Height          =   1305
      Left            =   165
      TabIndex        =   15
      Top             =   2280
      Width           =   5580
      _cx             =   9842
      _cy             =   2302
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   18
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRcdAnalyse.frx":2C17
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
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3255
      Left            =   1620
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4695
      Visible         =   0   'False
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6105
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRcdAnalyse.frx":2C40
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9234
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.CommandButton cmdShow 
      Caption         =   "开始分析(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   180
      TabIndex        =   6
      Top             =   945
      Width           =   1605
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6645
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   6645
      TabIndex        =   11
      Top             =   60
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -15
      Top             =   4815
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
            Picture         =   "frmRcdAnalyse.frx":34D2
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   1185
      Left            =   210
      TabIndex        =   9
      Top             =   4005
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   2090
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
   Begin MSComctlLib.ListView lvwText 
      Height          =   1200
      Left            =   180
      TabIndex        =   8
      Top             =   2775
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   2117
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
   Begin VSFlex8Ctl.VSFlexGrid hgdText 
      Height          =   1305
      Left            =   2610
      TabIndex        =   14
      Top             =   2520
      Width           =   5580
      _cx             =   9842
      _cy             =   2302
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   18
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRcdAnalyse.frx":3A6C
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
   End
   Begin VSFlex8Ctl.VSFlexGrid hgdItem 
      Height          =   1305
      Left            =   2535
      TabIndex        =   16
      Top             =   4170
      Width           =   5580
      _cx             =   9842
      _cy             =   2302
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   18
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRcdAnalyse.frx":3A95
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   300
      Left            =   915
      TabIndex        =   4
      Top             =   525
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   60817411
      CurrentDate     =   37922
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   300
      Left            =   3330
      TabIndex        =   5
      Top             =   525
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   60817411
      CurrentDate     =   37922
   End
   Begin MSComctlLib.TabStrip tabTopic 
      Height          =   1020
      Left            =   0
      TabIndex        =   7
      Top             =   1530
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   1799
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "疾病诊断对照(&1)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "病历文本对比(&2)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "诊治所见分析(&3)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   7035
      Picture         =   "frmRcdAnalyse.frx":3ABE
      Top             =   1005
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "姓名：高学娅    性别：女    年龄：65"
      Height          =   180
      Left            =   1920
      TabIndex        =   2
      Top             =   195
      Width           =   3240
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "日期(&D)                        至"
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   585
      Width           =   2970
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuPreview 
         Caption         =   "预览(V)"
      End
      Begin VB.Menu mnuPopuPrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPopuExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuPopuSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopuCopy 
         Caption         =   "表复制(C)"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmRcdAnalyse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String, aryTemp() As String

Private WithEvents objParentForm As Form
Attribute objParentForm.VB_VarHelpID = -1

Public Sub ShowMe(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngPatiId As Long)
    '---------------------------------------------
    '功能：根据上级程序要求，以模态或非模态显示病人病史分析
    '入参：frmParent-父窗体；
    '      blnModal-是否模态显示（通常和上级窗体一致）；
    '      lngPatiId-要显示的病人ID，不传递或传递时，用户可改变；
    '---------------------------------------------
    If lngPatiId <> 0 Then
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                    " where 病人id=" & lngPatiId
        Err = 0: On Error GoTo ErrHand
        With rsTemp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
            If .RecordCount = 1 Then
                Me.txtPati.Tag = !病人ID: Me.txtPati.Text = Me.txtPati.Tag
                Me.lblInfo.Caption = "姓名：" & Trim(!姓名) & _
                        Space(3) & "性别：" & IIf(IsNull(!性别), "", !性别) & _
                        Space(3) & "年龄：" & IIf(IsNull(!年龄), "", !年龄)
                Me.lblInfo.Tag = Trim(!姓名)
                Call zlClearTopic
                Me.cmdShow.Enabled = True
                Call cmdShow_Click
            End If
        End With
    End If
    
    On Error Resume Next
    Set objParentForm = frmParent
    Me.Show bytModal, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdShow_Click()
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        '提取诊断记录
        gstrSql = "select L.记录日期,L.诊断类型,L.诊断描述||decode(L.是否疑诊,1,' ？','') as 诊断描述,L.记录来源,L.记录人" & _
                " from 病人诊断记录 L" & _
                " where L.病人ID=" & Val(Me.txtPati.Text) & " and 取消人 is null" & _
                "       and L.记录日期 between to_date('" & Format(Me.dtpFrom.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')" & _
                "       and to_date('" & Format(Me.dtpTo.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')+1-1/24/60/60" & _
                " order by L.记录日期"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Do While Not .EOF
            If Me.hgdDiag.Rows - 1 < .AbsolutePosition Then Me.hgdDiag.Rows = Me.hgdDiag.Rows + 1
            Me.hgdDiag.TextMatrix(.AbsolutePosition, 0) = Format(!记录日期, "YYYY-MM-DD")
            Select Case !诊断类型
            Case 1
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "西医门诊诊断"
            Case 2
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "西医入院诊断"
            Case 3
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "西医出院诊断"
            Case 5
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "院内感染"
            Case 6
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "病理诊断"
            Case 7
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "损伤中毒码"
            Case 8
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "术前诊断"
            Case 9
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "术后诊断"
            Case 11
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "中医门诊诊断"
            Case 12
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "中医入院诊断"
            Case 13
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "中医出院诊断"
            Case Else
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 1) = "其他诊断"
            End Select
            Me.hgdDiag.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!诊断描述), "", !诊断描述)
            Select Case !记录来源
            Case 1
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 3) = "病历"
            Case 2
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 3) = "入院登记"
            Case 3
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 3) = "首页"
            Case Else
                Me.hgdDiag.TextMatrix(.AbsolutePosition, 3) = "未知"
            End Select
            Me.hgdDiag.TextMatrix(.AbsolutePosition, 4) = IIf(IsNull(!记录人), "", !记录人)
            .MoveNext
        Loop
        Call Me.hgdDiag.AutoSize(2)
        
        '提取病历文本项目
        gstrSql = "select I.ID,I.编码,I.名称" & _
                " from (select distinct C.元素编码" & _
                "       from 病人病历记录 L,病人病历内容 C" & _
                "       where L.ID=C.病历记录ID and L.作废人 is null and C.元素类型=0" & _
                "           and L.病人ID=" & Val(Me.txtPati.Text) & _
                "           and L.书写日期 between to_date('" & Format(Me.dtpFrom.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')" & _
                "           and to_date('" & Format(Me.dtpTo.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')+1-1/24/60/60) D," & _
                "      病历元素目录 I" & _
                " where D.元素编码=I.编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwText.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "item": objItem.SmallIcon = "item"
            objItem.SubItems(Me.lvwText.ColumnHeaders("编码").Index - 1) = !编码
            .MoveNext
        Loop
        If Me.lvwText.ListItems.Count > 0 Then
            Me.lvwText.ListItems(1).Selected = True
            Me.lvwText.SelectedItem.EnsureVisible
            Call lvwText_ItemClick(Me.lvwText.SelectedItem)
        End If
        
        '提取诊治所见项目
        gstrSql = "select I.ID,I.中文名,I.英文名,I.单位,I.类型" & _
                " from (select distinct S.所见项ID" & _
                "       from 病人病历记录 L,病人病历内容 C,病人病历所见单 S" & _
                "       where L.ID=C.病历记录ID and C.ID=S.病历ID and L.作废人 is null" & _
                "           and L.病人ID=" & Val(Me.txtPati.Text) & _
                "           and L.书写日期 between to_date('" & Format(Me.dtpFrom.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')" & _
                "           and to_date('" & Format(Me.dtpTo.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')+1-1/24/60/60) S," & _
                "      诊治所见项目 I,诊治所见分类 K" & _
                " where S.所见项ID=I.ID and I.分类ID=K.ID and K.性质<>1"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !中文名)
            objItem.Icon = "item": objItem.SmallIcon = "item"
            objItem.SubItems(Me.lvwItem.ColumnHeaders("英文名").Index - 1) = IIf(IsNull(!英文名), "", !英文名)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            objItem.Tag = IIf(IsNull(!类型), 0, !类型)
            .MoveNext
        Loop
        If Me.lvwItem.ListItems.Count > 0 Then
            Me.lvwItem.ListItems(1).Selected = True
            Me.lvwItem.SelectedItem.EnsureVisible
            Call lvwItem_ItemClick(Me.lvwItem.SelectedItem)
        End If
        
    End With
    Me.cmdShow.Enabled = False
    Me.cmdShow.Caption = "重新分析(&A)"
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpFrom_Change()
    If Me.dtpFrom.Value > Me.dtpTo.Value Then Me.dtpTo.Value = Me.dtpFrom.Value
    If Format(Me.dtpFrom.Tag, "YYYY-MM-DD") <> Format(Me.dtpFrom.Value, "YYYY-MM-DD") Then
        Me.dtpFrom.Tag = Format(Me.dtpFrom.Value, "YYYY-MM-DD")
        Call zlClearTopic
        Me.cmdShow.Enabled = True
    End If
End Sub

Private Sub dtpFrom_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpTo_Change()
    If Me.dtpFrom.Value > Me.dtpTo.Value Then Me.dtpFrom.Value = Me.dtpTo.Value
    If Format(Me.dtpTo.Tag, "YYYY-MM-DD") <> Format(Me.dtpTo.Value, "YYYY-MM-DD") Then
        Me.dtpTo.Tag = Format(Me.dtpTo.Value, "YYYY-MM-DD")
        Call zlClearTopic
        Me.cmdShow.Enabled = True
    End If
End Sub

Private Sub dtpTo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    
    '界面元素形态设置
    Me.dtpFrom.MaxDate = Date: Me.dtpFrom.Value = DateAdd("m", -1, Date)
    Me.dtpTo.MaxDate = Date: Me.dtpTo.Value = Date
    
    With Me.hgdDiag
        .Rows = .FixedRows + 1: .Cols = 5
        .TextMatrix(0, 0) = "日期": .TextMatrix(0, 1) = "类型": .TextMatrix(0, 2) = "诊断": .TextMatrix(0, 3) = "来源": .TextMatrix(0, 4) = "记录人"
        .ColWidth(0) = 1000: .ColWidth(1) = 1200: .ColWidth(2) = 4300: .ColWidth(3) = 1000: .ColWidth(4) = 800
        For intCol = .FixedCols To .Cols - 1
            .FixedAlignment(intCol) = 4: .ColAlignment(intCol) = 0
        Next
    End With
    
    Me.lvwText.ListItems.Clear
    With Me.lvwText.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2000
        .Add , "编码", "编码", 750
    End With
    With Me.lvwText
        .Width = 2800
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1: .SortOrder = lvwAscending
    End With
    
    With Me.hgdText
        .Rows = .FixedRows + 1: .Cols = 4
        .TextMatrix(0, 0) = "日期": .TextMatrix(0, 1) = "内容": .TextMatrix(0, 2) = "位置": .TextMatrix(0, 3) = "书写人"
        .ColWidth(0) = 1000: .ColWidth(1) = 3500: .ColWidth(2) = 1200: .ColWidth(3) = 800
        For intCol = .FixedCols To .Cols - 1
            .FixedAlignment(intCol) = 4: .ColAlignment(intCol) = 0
        Next
    End With
    
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "中文名", "中文名", 1400
        .Add , "英文名", "英文名", 850
        .Add , "单位", "单位", 800
    End With
    With Me.lvwItem
        .Width = 3100
        .SortKey = .ColumnHeaders("中文名").Index - 1: .SortOrder = lvwAscending
    End With
    With Me.hgdItem
        .Rows = .FixedRows + 1: .Cols = 4
        .TextMatrix(0, 0) = "日期": .TextMatrix(0, 1) = "数值(或内容)": .TextMatrix(0, 2) = "位置": .TextMatrix(0, 3) = "书写人"
        .ColWidth(0) = 1600: .ColWidth(1) = 2500: .ColWidth(2) = 1500: .ColWidth(3) = 800
        For intCol = .FixedCols To .Cols - 1
            .FixedAlignment(intCol) = 4: .ColAlignment(intCol) = 0
        Next
    End With
    ReDim aryTemp(1, 2)
    Me.chtItem.ChartData = aryTemp
    
    With Me.lvwPati.ColumnHeaders
        .Clear
        .Add , "病人ID", "病人ID", 800
        .Add , "门诊号", "门诊号", 800
        .Add , "住院号", "住院号", 800
        .Add , "姓名", "姓名", 900
        .Add , "性别", "性别", 600
        .Add , "年龄", "年龄", 600
    End With
    With Me.lvwPati
        .SortKey = .ColumnHeaders("病人ID").Index - 1: .SortOrder = lvwAscending
    End With

    Call tabTopic_Click
End Sub

Private Sub Form_Resize()
    Dim lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Err = 0: On Error Resume Next
    
    Me.cmdClose.Left = Me.ScaleWidth - Me.cmdClose.Width - 180
    Me.cmdHelp.Left = Me.cmdClose.Left
    Me.fraLine.Width = Me.cmdClose.Left - 180
    Me.imgLogo.Left = Me.cmdClose.Left + (Me.cmdClose.Width - Me.imgLogo.Width) / 2
    
    With Me.tabTopic
        .Left = 0: .Width = Me.ScaleWidth - .Left + 15
        .Height = Me.ScaleHeight - lngStatus - .Top + 15
    End With
    
    With Me.hgdDiag
        .Left = Me.tabTopic.Left + 90: .Width = Me.tabTopic.Width - .Left - 90
        .Top = Me.tabTopic.Top + 375: .Height = Me.tabTopic.Height - (.Top - Me.tabTopic.Top) - 90
    End With

    With Me.lvwText
        .Left = Me.tabTopic.Left + 90
        .Top = Me.tabTopic.Top + 375: .Height = Me.tabTopic.Height - (.Top - Me.tabTopic.Top) - 90
    End With
    With Me.hgdText
        .Left = Me.lvwText.Left + Me.lvwText.Width + 60: .Width = Me.tabTopic.Width - .Left - 90
        .Top = Me.lvwText.Top + 15: .Height = Me.tabTopic.Height - (.Top - Me.tabTopic.Top) - 90
    End With
    
    With Me.lvwItem
        .Left = Me.tabTopic.Left + 90
        .Top = Me.tabTopic.Top + 375: .Height = Me.tabTopic.Height - (.Top - Me.tabTopic.Top) - 90
    End With
    With Me.optChart(0)
        .Top = Me.lvwItem.Top: .Left = Me.lvwItem.Left + Me.lvwItem.Width + 60
    End With
    With Me.optChart(1)
        .Top = Me.lvwItem.Top: .Left = Me.optChart(0).Left + Me.optChart(0).Width + 60
    End With
    With Me.hgdItem
        .Left = Me.lvwItem.Left + Me.lvwItem.Width + 60: .Width = Me.tabTopic.Width - .Left - 90
        .Top = Me.optChart(0).Top + 300: .Height = Me.tabTopic.Height - (.Top - Me.tabTopic.Top) - 90
    End With
    With Me.chtItem
        .Left = Me.hgdItem.Left: .Width = Me.hgdItem.Width
        .Top = Me.hgdItem.Top: .Height = Me.hgdItem.Height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub hgdDiag_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call Me.hgdDiag.AutoSize(2)
End Sub

Private Sub hgdDiag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Call PopupMenu(Me.mnuPopu, 2)
End Sub

Private Sub hgdItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call Me.hgdItem.AutoSize(1)
End Sub

Private Sub hgdItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Call PopupMenu(Me.mnuPopu, 2)
End Sub

Private Sub hgdText_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call Me.hgdText.AutoSize(1)
End Sub

Private Sub hgdText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Call PopupMenu(Me.mnuPopu, 2)
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItem.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItem.SortOrder = IIf(Me.lvwItem.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItem.SortKey = ColumnHeader.Index - 1
        Me.lvwItem.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Me.hgdItem
        .Rows = .FixedRows + 1
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        '提取诊断记录
        gstrSql = "select decode(C.元素类型,-2,to_date(C.标题文本,'YYYY-MM-DD HH24:MI:SS'),L.书写日期) as 日期,S.所见内容 as 内容,L.病历名称,L.书写人" & _
                " from 病人病历记录 L,病人病历内容 C,病人病历所见单 S" & _
                " where L.ID=C.病历记录ID and C.ID=S.病历ID and L.作废人 is null" & _
                "       and S.所见项ID=" & Mid(Item.Key, 2) & _
                "       and L.病人ID=" & Val(Me.txtPati.Text) & _
                "       and decode(C.元素类型,-2,to_date(C.标题文本,'YYYY-MM-DD HH24:MI:SS'),L.书写日期) between" & _
                "       to_date('" & Format(Me.dtpFrom.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')" & _
                "       and to_date('" & Format(Me.dtpTo.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')+1-1/24/60/60" & _
                " order by decode(C.元素类型,-2,to_date(C.标题文本,'YYYY-MM-DD HH24:MI:SS'),L.书写日期)"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        If .EOF Then Me.hgdItem.Rows = 1: Exit Sub
        Do While Not .EOF
            If Me.hgdItem.Rows - 1 < .AbsolutePosition Then Me.hgdItem.Rows = Me.hgdItem.Rows + 1
            Me.hgdItem.TextMatrix(.AbsolutePosition, 0) = Format(!日期, "YYYY-MM-DD HH:MM")
            Me.hgdItem.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!内容), "", !内容)
            Me.hgdItem.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!病历名称), "", !病历名称)
            Me.hgdItem.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!书写人), "", !书写人)
            .MoveNext
        Loop
        Call Me.hgdItem.AutoSize(1)
    End With
    
    Err = 0: On Error GoTo 0
    If Val(Item.Tag) <> 0 Then
        Me.optChart(0).Value = True: Me.optChart(1).Value = False
        Me.optChart(0).Enabled = False: Me.optChart(1).Enabled = False
        Exit Sub
    Else
        Me.optChart(0).Enabled = True: Me.optChart(1).Enabled = True
    End If
    
    ReDim aryTemp(1 To Me.hgdItem.Rows - 1, 2)
    For intRow = 1 To Me.hgdItem.Rows - 1
        aryTemp(intRow, 1) = Format(CDate(Me.hgdItem.TextMatrix(intRow, 0)), "M月D日")
        aryTemp(intRow, 2) = Val(Me.hgdItem.TextMatrix(intRow, 1))
    Next
    
    With Me.chtItem
        .AllowDynamicRotation = False: .AllowDithering = False
        .Legend.Location.Visible = False
        .chartType = VtChChartType2dLine
        .ColumnCount = 1: .ColumnLabelCount = 1
        .RowCount = Me.hgdItem.Rows - 1
        .ChartData = aryTemp
        .Plot.SeriesCollection(1).Pen.VtColor.Set 45, 6, 198
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwPati.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwPati.SortOrder = IIf(Me.lvwPati.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwPati.SortKey = ColumnHeader.Index - 1
        Me.lvwPati.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwPati_DblClick()
    If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwPati
        If Val(Me.txtPati.Tag) <> Val(.SelectedItem.Text) Then
            Me.txtPati.Tag = .SelectedItem.Text
            Me.txtPati.Text = Me.txtPati.Tag
            Me.lblInfo.Caption = "姓名：" & .SelectedItem.SubItems(.ColumnHeaders("姓名").Index - 1) & _
                    Space(3) & "性别：" & .SelectedItem.SubItems(.ColumnHeaders("性别").Index - 1) & _
                    Space(3) & "年龄：" & .SelectedItem.SubItems(.ColumnHeaders("年龄").Index - 1)
            Me.lblInfo.Tag = .SelectedItem.SubItems(.ColumnHeaders("姓名").Index - 1)
            Call zlClearTopic
            Me.cmdShow.Enabled = True
        End If
        Me.txtPati.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
        Call lvwPati_DblClick
    End Select
End Sub

Private Sub lvwPati_LostFocus()
    Me.lvwPati.Visible = False
End Sub

Private Sub lvwText_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwText.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwText.SortOrder = IIf(Me.lvwText.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwText.SortKey = ColumnHeader.Index - 1
        Me.lvwText.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwText_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Me.hgdText
        .Rows = .FixedRows + 1
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        '提取诊断记录
        gstrSql = "select L.书写日期,T.内容,L.病历名称,L.书写人" & _
                " from 病人病历记录 L,病人病历内容 C,病人病历文本段 T" & _
                " where L.ID=C.病历记录ID and C.ID=T.病历ID and L.作废人 is null" & _
                "       and C.元素类型=0 and C.元素编码='" & Item.SubItems(Me.lvwText.ColumnHeaders("编码").Index - 1) & "'" & _
                "       and L.病人ID=" & Val(Me.txtPati.Text) & _
                "       and L.书写日期 between to_date('" & Format(Me.dtpFrom.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')" & _
                "       and to_date('" & Format(Me.dtpTo.Value, "YYYY-MM-DD") & "','YYYY-MM-DD')+1-1/24/60/60" & _
                " order by L.书写日期"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Do While Not .EOF
            If Me.hgdText.Rows - 1 < .AbsolutePosition Then Me.hgdText.Rows = Me.hgdText.Rows + 1
            Me.hgdText.TextMatrix(.AbsolutePosition, 0) = Format(!书写日期, "YYYY-MM-DD")
            Me.hgdText.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!内容), "", !内容)
            Me.hgdText.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!病历名称), "", !病历名称)
            Me.hgdText.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!书写人), "", !书写人)
            .MoveNext
        Loop
        Call Me.hgdText.AutoSize(1)
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuPopuCopy_Click()
    Dim objTab As Object
    If Me.tabTopic.Tabs(1).Selected Then
        Set objTab = Me.hgdDiag
    ElseIf Me.tabTopic.Tabs(2).Selected Then
        Set objTab = Me.hgdText
    ElseIf Me.tabTopic.Tabs(3).Selected Then
        Set objTab = Me.hgdItem
    End If
    strTemp = ""
    With objTab
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                If intCol = .Cols - 1 Then
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & vbCrLf
                Else
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & vbTab
                End If
            Next
        Next
    End With
    VB.Clipboard.Clear
    VB.Clipboard.SetText strTemp
End Sub

Private Sub mnuPopuExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuPopuPreview_Click()
    Call zlRptPrint(2)
End Sub

Private Sub mnuPopuPrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub objParentForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub optChart_Click(Index As Integer)
    If Me.optChart(0).Value Then
        Me.hgdItem.Visible = True: Me.chtItem.Visible = False
    Else
        Me.hgdItem.Visible = False: Me.chtItem.Visible = True
    End If
End Sub

Private Sub tabTopic_Click()
    If Me.tabTopic.Tabs(1).Selected Then
        Me.hgdDiag.Visible = True
        Me.lvwText.Visible = False: Me.hgdText.Visible = False
        Me.lvwItem.Visible = False
        Me.optChart(0).Visible = False: Me.optChart(1).Visible = False
        Me.hgdItem.Visible = False: Me.chtItem.Visible = False
        Me.stbThis.Panels(2).Text = "病人各种诊断的对比情况"
    ElseIf Me.tabTopic.Tabs(2).Selected Then
        Me.hgdDiag.Visible = False
        Me.lvwText.Visible = True: Me.hgdText.Visible = True
        Me.optChart(0).Visible = False: Me.optChart(1).Visible = False
        Me.hgdItem.Visible = False: Me.chtItem.Visible = False
        Me.stbThis.Panels(2).Text = "选择要对比的病历元素项目，可查看该元素在各病历中的记录变化"
    ElseIf Me.tabTopic.Tabs(3).Selected Then
        Me.hgdDiag.Visible = False
        Me.lvwText.Visible = False: Me.hgdText.Visible = False
        Me.lvwItem.Visible = True
        Me.optChart(0).Visible = True: Me.optChart(1).Visible = True
        If Me.optChart(0).Value = True Then
            Me.hgdItem.Visible = True: Me.chtItem.Visible = False
        Else
            Me.hgdItem.Visible = False: Me.chtItem.Visible = True
        End If
        Me.stbThis.Panels(2).Text = "选择要对比的诊治项目，查看病人该项数值的记录反映病人病情变化"
    End If
End Sub

Private Sub txtPati_GotFocus()
    Me.txtPati.SelStart = 0: Me.txtPati.SelLength = 100
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If InStr("~!@#$^&()|=`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Me.txtPati.Text = Trim(Me.txtPati.Text)
    If Me.txtPati.Text = "" Then Me.txtPati.Text = Me.txtPati.Tag: Exit Sub
    
    Select Case Left(Me.txtPati.Text, 1)
    Case "-", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" '病人ID
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 病人id=" & Abs(Val(Me.txtPati.Text))
    Case "+"        '住院号
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 住院号=" & Val(Me.txtPati.Text)
    Case "*"        '门诊号
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 门诊号=" & Val(Mid(Me.txtPati.Text, 2))
    Case Else       '病人姓名
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 姓名 like '" & Me.txtPati.Text & "%'"
    End Select
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        If .BOF Or .EOF = 1 Then
            MsgBox "未找到指定病人", vbExclamation, gstrSysName
            Me.txtPati.Text = "": Me.txtPati.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Val(Me.txtPati.Tag) <> !病人ID Then
                Me.txtPati.Tag = !病人ID: Me.txtPati.Text = Me.txtPati.Tag
                Me.lblInfo.Caption = "姓名：" & Trim(!姓名) & _
                        Space(3) & "性别：" & IIf(IsNull(!性别), "", !性别) & _
                        Space(3) & "年龄：" & IIf(IsNull(!年龄), "", !年龄)
                Me.lblInfo.Tag = !姓名
                Call zlClearTopic
                Me.cmdShow.Enabled = True
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwPati.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwPati.ListItems.Add(, "_" & !病人ID, !病人ID)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("门诊号").Index - 1) = IIf(IsNull(!门诊号), "", !门诊号)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("住院号").Index - 1) = IIf(IsNull(!住院号), "", !住院号)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("姓名").Index - 1) = IIf(IsNull(!姓名), "", !姓名)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("性别").Index - 1) = IIf(IsNull(!性别), "", !性别)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("年龄").Index - 1) = IIf(IsNull(!年龄), "", !年龄)
            .MoveNext
        Loop
        Me.lvwPati.ListItems(1).Selected = True
    End With
    With Me.lvwPati
        .Left = Me.txtPati.Left
        .Top = Me.txtPati.Top + Me.txtPati.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPati_LostFocus()
    Me.txtPati.Text = Me.txtPati.Tag
End Sub

Private Sub zlClearTopic()
    '---------------------------------------------
    '根清除分析内容：在分析要求改变时调用
    '---------------------------------------------
    With Me.hgdDiag
        .Rows = .FixedRows + 1
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    Me.lvwText.ListItems.Clear
    With Me.hgdText
        .Rows = .FixedRows + 1
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    Me.lvwItem.ListItems.Clear
    With Me.hgdItem
        .Rows = .FixedRows + 1
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    ReDim aryTemp(1, 2)
    Me.chtItem.ChartData = aryTemp
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    On Error Resume Next
    If Me.tabTopic.Tabs(1).Selected Then
        objPrint.Title.Text = "“" & Me.lblInfo.Tag & "”疾病诊断对照"
        Set objPrint.Body = Me.hgdDiag
    ElseIf Me.tabTopic.Tabs(2).Selected Then
        objPrint.Title.Text = "“" & Me.lblInfo.Tag & "”" & Me.lvwText.SelectedItem.Text & "对比"
        Set objPrint.Body = Me.hgdText
    ElseIf Me.tabTopic.Tabs(3).Selected Then
        objPrint.Title.Text = "“" & Me.lblInfo.Tag & "”" & Me.lvwItem.SelectedItem.Text & "对比"
        Set objPrint.Body = Me.hgdItem
    End If
    objPrint.Title.Font.Size = 11
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

