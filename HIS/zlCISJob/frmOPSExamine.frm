VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOPSExamine 
   Caption         =   "手术审核管理"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15165
   Icon            =   "frmOPSExamine.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraDoctor 
      Caption         =   "医生"
      ForeColor       =   &H000040C0&
      Height          =   8775
      Left            =   120
      TabIndex        =   30
      Top             =   600
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptDoc 
         Height          =   7140
         Left            =   105
         TabIndex        =   31
         Top             =   1500
         Width           =   3330
         _Version        =   589884
         _ExtentX        =   5874
         _ExtentY        =   12594
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   33
         Top             =   788
         Width           =   1905
      End
      Begin VB.CheckBox chkIsShowAll 
         Caption         =   "只显示有申请的医生"
         Height          =   180
         Left            =   1080
         TabIndex        =   32
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室(&D)"
         Height          =   180
         Left            =   315
         TabIndex        =   36
         Top             =   420
         Width           =   630
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找(&F)"
         Height          =   180
         Left            =   315
         TabIndex        =   35
         Top             =   840
         Width           =   630
      End
   End
   Begin VB.PictureBox picUnAudited 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4440
      ScaleHeight     =   5895
      ScaleWidth      =   9735
      TabIndex        =   20
      Top             =   2640
      Width           =   9735
      Begin VB.PictureBox picDateY 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   9375
         TabIndex        =   38
         Top             =   0
         Width           =   9375
         Begin VB.CommandButton cmdFindY 
            Caption         =   "查找(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.ComboBox cboDateY 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   30
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTimeY 
            Height          =   300
            Index           =   1
            Left            =   4515
            TabIndex        =   41
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106496003
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTimeY 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   42
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106496003
            CurrentDate     =   40256
         End
         Begin VB.Label lblDateY 
            Caption         =   "开嘱时间"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   75
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从                 至"
            Height          =   180
            Left            =   2460
            TabIndex        =   43
            Top             =   90
            Width           =   1890
         End
      End
      Begin VB.PictureBox picDate 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9375
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   9375
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "查找(&F)"
            Height          =   350
            Left            =   5910
            TabIndex        =   22
            Top             =   0
            Visible         =   0   'False
            Width           =   1100
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   4515
            TabIndex        =   24
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106496003
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   25
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106496003
            CurrentDate     =   40256
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从                 至"
            Height          =   180
            Left            =   2460
            TabIndex        =   27
            Top             =   90
            Width           =   1890
         End
         Begin VB.Label lblDate 
            Caption         =   "审核时间"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   75
            Width           =   735
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAudit 
         Height          =   4860
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   8505
         _cx             =   15002
         _cy             =   8572
         Appearance      =   1
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOPSExamine.frx":6852
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
   End
   Begin VB.Frame fraPati 
      Caption         =   "病人信息"
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   11295
      Begin VB.PictureBox picInShow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   8445
         TabIndex        =   2
         Top             =   360
         Width           =   8450
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   7800
            TabIndex        =   3
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   4
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   5
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblInformation 
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   7
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblCaption 
            Caption         =   "体重："
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "病况："
            Height          =   255
            Index           =   5
            Left            =   7200
            TabIndex        =   11
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "护理等级："
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   10
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblCaption 
            Caption         =   "床号："
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   9
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblCaption 
            Caption         =   "入院时间："
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   8
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.ComboBox cbo过敏 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   697
         Width           =   4815
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   15
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblInformation 
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblCaption 
         Caption         =   "性别："
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "年龄："
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "过敏药物："
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         Caption         =   "诊断："
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   7335
      Left            =   3720
      TabIndex        =   29
      Top             =   1800
      Width           =   11355
      _Version        =   589884
      _ExtentX        =   20029
      _ExtentY        =   12938
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   37
      Top             =   10575
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   635
      SimpleText      =   $"frmOPSExamine.frx":68ED
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOPSExamine.frx":6934
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21669
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
   Begin MSComctlLib.ImageList img16 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSExamine.frx":71C8
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSExamine.frx":DA2A
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSExamine.frx":1428C
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSExamine.frx":14826
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgAdvice 
      Left            =   1200
      Top             =   0
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
            Picture         =   "frmOPSExamine.frx":14DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSExamine.frx":1535A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSExamine.frx":158F4
            Key             =   "签名"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmOPSExamine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstrPrivs As String
Private mlngModul As Long
Private mlngCodeType As Long         '0-拼音,1-五笔
Private mobjBar As CommandBar
Private mlngLevel As Long
Private mblnIsUpdate As Boolean

Private mlngFindNum As Long
Private mstrChangeRows As String   '记录修改的行
Private mstr签名IDs As String      '取消审核的时候记录一个病人处理过的签名ID
'手术审核暂时不启用签名功能，所以判断加了 And 1 = 0
Private mblnTmp As Boolean
Private mrsDefine As ADODB.Recordset
Private Enum Enum_Dor
    COL_人员ID = 0
    col_姓名 = 1
    COL_手术等级 = 2
    COL_拼音简码 = 3
    COL_五笔简码 = 4
    COL_所属部门 = 5
    COL_所属部门ID = 6
End Enum

Private Enum Enum_Advice
    col_选择 = 0
    col_取消选择 = 1
    col_审核说明 = 2
    col_审核时间 = 3
    col_病人姓名 = 4
    COL_医嘱内容 = 5
    col_期效 = 6
'用简洁模式，所以总量和单量隐藏起来，和医嘱内容合并
    COL_单量 = 7
    COL_主刀医生 = 8
    col_开始时间 = 9
'隐藏列
    col_医嘱ID = 10
    col_相关ID = 11
    col_性别 = 12
    col_年龄 = 13
    col_体重 = 14
    col_入院时间 = 15
    col_床号 = 16
    col_病况 = 17
    col_护理等级 = 18
    col_病人Id = 19
    col_主页ID = 20
    col_组ID = 21
    COL_诊疗类别 = 22
    col_病人来源 = 23
    col_签名id = 24
    col_医嘱状态 = 25
End Enum

Private Enum enum_Info
    info_入院时间 = 0
    info_性别 = 1
    info_年龄 = 2
    info_床号 = 3
    info_护理等级 = 4
    info_病况 = 5
    info_诊断 = 6
    info_体重 = 7
End Enum

Private Sub cboDept_Click()
    Call LoadDoc
End Sub

Private Sub LoadDoc()
'加载权限比操作员低的医生
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    If cboDept.ListIndex = -1 Then Exit Sub
    
    strSQL = "Select DISTINCT a.Id, A.性别" & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", ",b.部门ID,e.名称 as 所属部门") & ",a.姓名,a.手术等级, Upper(zlSpellCode(a.姓名)) As 拼音简码, Upper(Zlwbcode(a.姓名)) As 五笔简码" & vbNewLine & _
            "From 人员表 A, 部门人员 B, 人员性质说明 D,部门表 E" & IIf(chkIsShowAll.Value, " ,(Select Distinct 开嘱医生 From 病人医嘱记录 F Where f.审核状态 = 1 And F.开嘱时间 Between [3] And [4] and f.医嘱状态=1 And f.诊疗类别 ='F') F", "") & vbNewLine & _
            "Where a.Id = b.人员id And e.ID=b.部门ID And d.人员id = a.Id  And d.人员性质 = '医生' And " & vbNewLine & _
            "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)  " & vbNewLine & _
            "   " & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", "And b.部门id=[2]") & _
            IIf(chkIsShowAll.Value, " And  f.开嘱医生 = a.姓名 ", "")
            
    On Error GoTo errH
    
    rptDoc.Records.DeleteAll
    vsAudit.Rows = 1: vsAudit.AddItem ""
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngLevel, Val(cboDept.ItemData(cboDept.ListIndex)), CDate(dtpTimeY(0).Value), CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60))
    
    With rptDoc
        Do While Not rsTmp.EOF
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                objItem.Icon = img16.ListImages.Item(IIf(rsTmp!性别 & "" = "女", "feMale", "Male")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!手术等级 & "")
            Set objItem = objRecord.AddItem(rsTmp!拼音简码 & "")
            Set objItem = objRecord.AddItem(rsTmp!五笔简码 & "")
            If Val(cboDept.ItemData(cboDept.ListIndex)) <> -1 Then
                Set objItem = objRecord.AddItem(rsTmp!所属部门 & "")
                Set objItem = objRecord.AddItem(rsTmp!部门ID & "")
            End If
            rsTmp.MoveNext
        Loop
        .Populate
    End With
    mlngFindNum = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    If rptDoc.Visible = False Then Exit Sub
    If rptDoc.Records.Count > 0 Then
        If rptDoc.SelectedRows.Count = 0 Then Exit Sub
        strSubhead = rptDoc.SelectedRows(0).Record(col_姓名).Value & "手术审核清单"
    Else
        Exit Sub
    End If
    
    '调用打印部件处理
    Set objPrint.Body = Me.vsAudit
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub Cancle()
'功能：取消保存
    Dim i As Long
    With vsAudit
        If MsgBox("本次修改的内容未保存，是否继续？", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            If tbcSub.Selected.Tag = "已审核" Then
                Call LoadAdvice(True)
            Else
                Call LoadAdvice
            End If
            mblnIsUpdate = False
            mstrChangeRows = ""
        End If
    End With
End Sub

Private Sub SaveAudit()
'功能：保存审核信息
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strDate As String
    
    With vsAudit
        If .EditText <> "" Then .TextMatrix(.Row, .Col) = .EditText
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        If tbcSub.Selected.Tag = "待审核" Then
            For i = 1 To .Rows - 1
                '一个病人调用一次
                If RowIn同一病人(i, lngBegin, lngEnd, vsAudit) Then
                    Call SaveAuditOnePati(lngBegin, lngEnd, strDate)
                    i = lngEnd
                Else
                    Call SaveAuditOnePati(i, i, strDate)
                End If
            Next
            Call LoadAdvice
        Else
            Call SaveAuditUpdate
            Call LoadAdvice(True)
        End If
        mstrChangeRows = ""
        mblnIsUpdate = False
    End With
End Sub

Private Sub SaveAuditUpdate()
'功能：修改已审核未通过的审核说明
    Dim i As Long
    Dim strSQL As String
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strDate As String
    Dim varArr As Variant
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    If mstrChangeRows <> "" Then
        varArr = Split(mstrChangeRows, ",")
        With vsAudit
            If .EditText <> "" Then .TextMatrix(.Row, .Col) = .EditText
            For i = 0 To UBound(varArr)
                If .TextMatrix(Val(varArr(i)), col_医嘱ID) <> "" And Val(varArr(i)) <> 0 Then
                    strSQL = "Zl_医嘱审核管理_Update(" & Val(.TextMatrix(Val(varArr(i)), col_医嘱ID)) & "," & strDate & ",'" & .TextMatrix(Val(varArr(i)), col_审核说明) & "',1)"
                    colSQL.Add strSQL, "C" & colSQL.Count + 1
                End If
            Next
        End With
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.Count
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveAuditOnePati(ByVal lngBegin As Long, ByVal lngEnd As Long, ByVal strDate As String)
'功能：保存审核信息
'参数：从第几行开始，到第几行结束（同一个病人）
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, j As Long
    Dim strIDs As String
    Dim strSource As String, strSign As String
    Dim lng签名ID As Long, lng证书ID As Long
    Dim intRule As Integer, strTimeStamp As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    Dim strSignSQL As String
    
    With vsAudit
        For i = lngBegin To lngEnd
            If .TextMatrix(i, col_医嘱ID) = "" Then Exit Sub
            If Val(.Cell(flexcpData, i, col_选择) & "") <> "0" Then
                strSQL = Val(.TextMatrix(i, col_医嘱ID)) & "|" & "Zl_医嘱审核管理_Audit(" & Val(.TextMatrix(i, col_医嘱ID)) & "," & Val(.Cell(flexcpData, i, col_选择) & "") & "," & _
                        "'" & UserInfo.姓名 & "'," & strDate & ",'" & .TextMatrix(i, col_审核说明) & "'"
                colSQL.Add strSQL, "C" & colSQL.Count + 1
                If Val(.Cell(flexcpData, i, col_选择) & "") = 1 Then
                    strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_医嘱ID)
                End If
            End If
        Next
      
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.Count
            strSQL = Mid(colSQL("C" & i), InStr(colSQL("C" & i), "|") + 1) & ",1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Next
        If strSignSQL <> "" Then
            Call zlDatabase.ExecuteProcedure(strSignSQL, Me.Caption)
        End If
    gcnOracle.CommitTrans: blnTrans = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpTime(0).MaxDate = curDate
    dtpTime(1).MaxDate = curDate
    cmdFind.Visible = False
    
    Select Case cboTime.ListIndex
    Case 0 '今日
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 1 '最近二天
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 2 '最近三天
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 3 '最近一周
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 4 '最近一月
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 5 '指  定
        If Me.Visible Then
            dtpTime(0).SetFocus
            cmdFind.Visible = True
        End If
    End Select
    
    If cboTime.ListIndex <> cboTime.ListCount - 1 And Me.Visible Then
        Call LoadAdvice(True)
    End If
End Sub

Private Sub CancleAudit()
'取消审核
    Dim i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnIsCheck As Boolean
    
    With vsAudit
        '判断是否有勾选的，有勾选就以勾选为准
        If MsgBox("取消审核的医嘱可在待审核中重新审核，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, col_取消选择))) = 1 Then Exit For
        Next
        blnIsCheck = i < .Rows
        
        If blnIsCheck Then
            For i = i To .Rows - 1
                If Abs(Val(.TextMatrix(i, col_取消选择))) = 1 Then
                    If RowIn同一病人(i, lngBegin, lngEnd, vsAudit) Then
                        Call CancleAuditOnePati(lngBegin, lngEnd)
                        i = lngEnd
                    Else
                        Call CancleAuditOnePati(i, i)
                    End If
                End If
            Next
        Else
            If .Row = 0 Then Exit Sub
            If gintCA > 0 And 1 = 0 Then
                If RowIn同一病人(.Row, lngBegin, lngEnd, vsAudit) And Val(.TextMatrix(.Row, col_签名id)) <> 0 Then
                    '如果是选择的情况，则不用递归，直接处理和选中行签名ID一样的医嘱
                    Call CancleAuditOnePati(lngBegin, lngEnd, Not blnIsCheck, Val(.TextMatrix(.Row, col_签名id)), False)
                Else
                    Call CancleAuditOnePati(.Row, .Row, Not blnIsCheck)
                End If
            Else
                Call CancleAuditOnePati(.Row, .Row, Not blnIsCheck)
            End If
        End If
        Call LoadAdvice(True)
    End With
End Sub

Private Sub CancleAuditOnePati(ByVal lngBegin As Long, ByVal lngEnd As Long, Optional ByVal blnIsNoCheck As Boolean, _
        Optional ByVal lng签名ID_IN As Long, Optional ByVal blnIsRecursive As Boolean = True)
'功能：取消审核
'参数：lngBegin从第几行开始，lngEnd到第几行结束（同一个病人）
'     blnIsNoCheck=没有勾选则已选中行为准取消审核
'     lng签名ID_IN：用于递归调用，如果第一次循环中发现有签名ID<>0，则递归调用本函数，并把这里的签名ID传入，
'    加入到字符串mstr签名IDs里，第二次进来则处理签名ID的医嘱,如果再发现与传入的签名ID不一样，并且又不在字符串mstr签名IDs中，则为新的，则再递归调用。
'    blnIsRecursive:是否递归，默认为要递归
    Dim strSQL As String, i As Long, j As Long
    Dim strIDs As String, blnTrans As Boolean
    Dim strSource As String, strSign As String
    Dim lng证书ID As Long, lng签名ID As Long
    Dim intRule As Integer, strTimeStamp As String
    Dim lngGroupBegin As Long, lngGroupEnd As Long
    
    With vsAudit
        For i = lngBegin To lngEnd
            If Abs(Val(.TextMatrix(i, col_取消选择))) = 1 Or blnIsNoCheck Then
                strIDs = strIDs & IIf(strIDs = "", "", ",") & .TextMatrix(i, col_医嘱ID)
            End If
        Next
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    '取消签名
    If gintCA > 0 And 1 = 0 And strSign <> "" Then
        Call zlDatabase.ExecuteProcedure(strSign, Me.Caption)
    End If
    '取消审核
    If strIDs <> "" Then
        strSQL = "Zl_医嘱审核管理_Cancel('" & strIDs & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '执行完后，判断是否发现有签名ID的医嘱，然后递归调用
    If blnIsRecursive Then
        If lng签名ID <> 0 Then
            mstr签名IDs = mstr签名IDs & "," & lng签名ID
            Call CancleAuditOnePati(lngBegin, lngEnd, blnIsNoCheck, lng签名ID)
        End If
    End If
    mstr签名IDs = "0"
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim objPopup As CommandBarPopup
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
    Case conMenu_Edit_Untread     '取消
        Call Cancle
    Case conMenu_Edit_Save        '保存
        Call SaveAudit
    Case conMenu_Edit_AdviceUnAudit '取消审核
        Call CancleAudit
    Case conMenu_Tool_Archive '电子病案查阅
        If vsAudit.Row = 0 Or vsAudit.TextMatrix(1, col_医嘱ID) = "" Then Exit Sub
        Call frmArchiveView.ShowArchive(Me, Val(vsAudit.TextMatrix(vsAudit.Row, col_病人Id)), Val(vsAudit.TextMatrix(vsAudit.Row, col_主页ID)))
    Case conMenu_View_Find '查找
        txtFind.SetFocus '有时需要定位一下
        If txtFind.Text <> "" Then
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_FindNext '查找下一个
        If txtFind.Text = "" Then
            txtFind.SetFocus
        Else
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        cbsMain_Resize
    Case conMenu_View_Refresh '刷新
        If tbcSub.Selected.Tag = "待审核" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If rptDoc.SelectedRows.Count = 0 Or vsAudit.Row <= 0 Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "执行科室=" & cboDept.ItemData(cboDept.ListIndex))
            Else
                With vsAudit
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "执行科室=" & cboDept.ItemData(cboDept.ListIndex), "审核人=" & rptDoc.SelectedRows(0).Record(col_姓名).Value, _
                        "病人ID=" & .TextMatrix(.Row, col_病人Id), "主页ID=" & .TextMatrix(.Row, col_主页ID), "医嘱ID=" & .TextMatrix(.Row, col_医嘱ID))
                End With
            End If
        End If
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With fraDoctor
        .Top = lngTop
        .Left = lngLeft + 100
        .Height = lngBottom - lngTop - stbThis.Height
    End With
    rptDoc.Height = fraDoctor.Height - 1600
    
    With fraPati
        .Top = fraDoctor.Top
        .Left = fraDoctor.Left + fraDoctor.Width + 45
        .Width = lngRight - fraDoctor.Width - 200
    End With
    
    With tbcSub
        .Top = fraPati.Top + fraPati.Height + 45
        .Left = fraPati.Left
        .Height = fraDoctor.Height - fraPati.Height - 45
        .Width = fraPati.Width + 50
    End With
    
    Me.Refresh
End Sub

Private Sub SetControlVisible(ByRef Control As XtremeCommandBars.ICommandBarControl)
    '根据权限设置按钮可见状态
    
    Select Case Control.ID
        Case conMenu_Edit_AdviceUnAudit
            If tbcSub.Selected.Tag <> "已审核" Then Control.Visible = False: Exit Sub
        Case conMenu_Tool_Archive '电子病案查阅
            If GetInsidePrivs(p电子病案查阅) = "" Then
                Control.Visible = False
                Exit Sub
            End If
    End Select
    Control.Visible = True
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim rptRecord As ReportRecord
        
'    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    Select Case Control.ID
    
        Case conMenu_Edit_Untread, conMenu_Edit_Save   '保存,取消
            Control.Enabled = mblnIsUpdate
        Case conMenu_View_Refresh, conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '刷新,打印
            Control.Enabled = Not mblnIsUpdate
            If mblnIsUpdate Then
                cboDept.Enabled = False
                txtFind.Enabled = False
                fraDoctor.Enabled = False
                cboDept.BackColor = &H8000000F
                txtFind.BackColor = &H8000000F
                cmdFind.Enabled = True
                cboTime.Enabled = False
                tbcSub.Item(IIf(tbcSub.Selected.Index = 0, 1, 0)).Enabled = False
            Else
                cboDept.Enabled = True
                txtFind.Enabled = True
                fraDoctor.Enabled = True
                cboTime.Enabled = True
                cmdFind.Enabled = True
                cboDept.BackColor = &H80000005
                txtFind.BackColor = &H80000005
                tbcSub.Item(IIf(tbcSub.Selected.Index = 0, 1, 0)).Enabled = True
            End If
        Case conMenu_Edit_AdviceUnAudit '取消审核
            Control.Enabled = vsAudit.Row <> 0 And Not mblnIsUpdate And vsAudit.TextMatrix(vsAudit.Row, col_医嘱状态) = "1"
        Case conMenu_Tool_Archive '电子病案查阅
            Control.Enabled = vsAudit.Row <> 0 And vsAudit.TextMatrix(1, col_医嘱ID) <> ""
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext '查找下一个
            Control.Visible = False
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkIsShowAll_Click()
    If mblnTmp Then Exit Sub
    
    Call LoadDoc
End Sub

Private Sub cmdFind_Click()
    Call LoadAdvice(True)
End Sub

Private Sub GetLocalSetting()
'获取本地参数
    cboTime.ListIndex = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "时间范围", 0)
    mblnTmp = True
    chkIsShowAll.Value = Val(zlDatabase.GetPara("只显示有申请的医生", glngSys, mlngModul, 0) & "")
    mblnTmp = False
End Sub

Private Sub cmdFindY_Click()
    Call LoadAdvice(False)
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim strHead As String
    
    mstrPrivs = GetInsidePrivs(p手术审核管理)
    mlngModul = p手术审核管理
    mlngCodeType = zlDatabase.GetPara("简码方式")
    mblnIsUpdate = False
    mstrChangeRows = ""
    mstr签名IDs = "0"
    
    '---cboTime
    cboTime.AddItem "今    日"
    cboTime.AddItem "最近二天"
    cboTime.AddItem "最近三天"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "[指  定]"
    cboTime.ListIndex = 0
    
    '---cboDateY
    cboDateY.AddItem "今    日"
    cboDateY.AddItem "最近二天"
    cboDateY.AddItem "最近三天"
    cboDateY.AddItem "最近一周"
    cboDateY.AddItem "最近一月"
    cboDateY.AddItem "[指  定]"
    cboDateY.ListIndex = 3
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "  待审核  ", picUnAudited.hwnd, 0).Tag = "待审核"
        .InsertItem(1, "  已审核  ", picUnAudited.hwnd, 0).Tag = "已审核"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlcommfun.GetPubIcons
    Call MainDefCommandBar
    
    'vsFlexGrid
    '-----------------------------------------------------
    strHead = ",450,1; ;审核说明,2560,1;审核时间;病人姓名,1000,1;医嘱内容,3500,1;期效,500,1;单量;主刀医生,1350,1;开始时间,2000,1;医嘱ID;相关ID ; 性别;年龄;体重;入院时间;床号; 组号; 病况;护理等级;病人ID; 主页ID; 组ID;诊疗类别 ;病人来源;签名id;医嘱状态"
    Call InitTable(vsAudit, strHead)
    vsAudit.ExtendLastCol = True
    vsAudit.Editable = flexEDKbdMouse
    vsAudit.Cell(flexcpPicture, 0, col_选择) = img16.ListImages("unCheck").Picture
    vsAudit.Cell(flexcpPictureAlignment, 0, col_选择) = flexPicAlignCenterCenter
    vsAudit.ColDataType(col_取消选择) = flexDTBoolean
    vsAudit.Cell(flexcpPicture, 0, col_取消选择) = img16.ListImages("unCheck").Picture
    vsAudit.Cell(flexcpPictureAlignment, 0, col_取消选择) = flexPicAlignCenterCenter
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    Call RestoreWinState(Me, App.ProductName)
    
    Set mrsDefine = InitAdviceDefine
    Call GetLocalSetting '本地参数
    
    Call LoadDept
End Sub

Private Sub LoadDept()
'加载操作员所属科室
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long
    
    strSQL = "Select B.ID,B.编码,B.名称 " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", ",A.缺省") & vbNewLine & _
            "From " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", "部门人员 A, ") & _
            " 部门表 B, 部门性质说明 C" & vbNewLine & _
            " Where B.Id = C.部门id " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", " And a.部门id = B.Id And A.人员ID = [1] ") & vbNewLine & _
            "  And C.工作性质 = '临床' And C.服务对象 <> 0  And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) Order By B.编码"

    On Error GoTo errH
    cboDept.Clear
    '所有部门
    If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 Then
        cboDept.AddItem "所有部门"
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        '所属缺省
        If InStr(";" & mstrPrivs & ";", ";所有部门;") = 0 Then
            If rsTmp!缺省 = 1 Then
                Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    Call LoadDoc
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptDoc
        
        Set objCol = .Columns.Add(COL_人员ID, "人员ID", 0, False)
        Set objCol = .Columns.Add(col_姓名, "姓名", 70, True)
        Set objCol = .Columns.Add(COL_手术等级, "手术等级", 80, True)
        Set objCol = .Columns.Add(COL_拼音简码, "拼音简码", 0, False)
        Set objCol = .Columns.Add(COL_五笔简码, "五笔简码", 0, False)
        Set objCol = .Columns.Add(COL_所属部门, "所属部门", 0, False)
        Set objCol = .Columns.Add(COL_所属部门ID, "所属部门ID", 0, False)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的医生..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
        If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 Then .GroupsOrder.Add .Columns(COL_所属部门)
    End With
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "取消审核(&U)")
        objControl.BeginGroup = True
        objControl.IconId = 21905
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("工具栏", xtpBarTop)
    With mobjBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdviceUnAudit, "取消审核(&U)")
            objControl.BeginGroup = True
            objControl.IconId = 21905
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnIsUpdate = True Then
        If MsgBox("当前输入的内容未保存，是否要退出？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmParent Is Nothing Then Set mfrmParent = Nothing
    mlngFindNum = 0
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "时间范围", cboTime.ListIndex
    zlDatabase.SetPara "只显示有申请的医生", chkIsShowAll.Value & "", glngSys, mlngModul
End Sub

Private Sub cboDateY_Click()
    Dim curDate As Date
    
    dtpTimeY(0).Enabled = cboDateY.ListIndex = cboDateY.ListCount - 1
    dtpTimeY(1).Enabled = cboDateY.ListIndex = cboDateY.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpTimeY(0).MaxDate = curDate
    dtpTimeY(1).MaxDate = curDate
    cmdFindY.Visible = False
    
    Select Case cboDateY.ListIndex
    Case 0 '今日
        dtpTimeY(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 1 '最近二天
        dtpTimeY(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 2 '最近三天
        dtpTimeY(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 3 '最近一周
        dtpTimeY(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 4 '最近一月
        dtpTimeY(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTimeY(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 5 '指  定
        If Me.Visible Then
            dtpTimeY(0).SetFocus
            cmdFindY.Visible = True
        End If
    End Select
    
    If cboDateY.ListIndex <> cboDateY.ListCount - 1 Then
        If Me.Visible Then Call LoadAdvice(True)
    End If
End Sub

Private Sub picUnAudited_Resize()
    On Error Resume Next
    picDate.Move 0, 0, picUnAudited.Width
    picDateY.Move 0, 0, picUnAudited.Width
    vsAudit.Move 0, picDate.Top + picDate.Height, picUnAudited.Width, picUnAudited.Height - picDate.Top + picDate.Height
End Sub

Private Sub rptDoc_SelectionChanged()
    If mlngFindNum <> 0 Then mlngFindNum = rptDoc.SelectedRows(0).Index + 1
    
    '加载医嘱列表
    If tbcSub.Selected.Tag = "待审核" Then
        If Me.Visible Then Call LoadAdvice
    Else
        If Me.Visible Then Call LoadAdvice(True)
    End If
End Sub

Private Sub rptDoc_SortOrderChanged()
    mlngFindNum = 0
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Tag = "已审核" Then
        picDate.Visible = True
        picDateY.Visible = False
        Call picUnAudited_Resize
        vsAudit.ColWidth(col_取消选择) = 250
        vsAudit.ColHidden(col_取消选择) = False
        vsAudit.ColWidth(col_审核时间) = 1800
        vsAudit.ColHidden(col_审核时间) = False
        Set vsAudit.Cell(flexcpPicture, 0, col_选择) = Nothing
        vsAudit.TextMatrix(0, col_选择) = "状态"
        If Me.Visible Then Call LoadAdvice(True)
    Else
        picDate.Visible = False
        picDateY.Visible = True
        Call picUnAudited_Resize
        vsAudit.Cell(flexcpPicture, 0, col_选择) = img16.ListImages("unCheck").Picture
        vsAudit.TextMatrix(0, col_选择) = ""
        vsAudit.ColWidth(col_取消选择) = 0
        vsAudit.ColHidden(col_取消选择) = True
        vsAudit.ColWidth(col_审核时间) = 0
        vsAudit.ColHidden(col_审核时间) = True
        If Me.Visible Then Call LoadAdvice
    End If
End Sub

Private Sub txtFind_Change()
    mlngFindNum = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub LoadAdvice(Optional ByVal blnIsAudited As Boolean)
'加载待审核和已审核的医嘱
'参数：是否加载已审核医嘱,为空为加载待审核医嘱
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngID As Long       '用于定位
    Dim strFormat As String
    Dim strTmp As String
    Dim blnDo As Boolean
    
    strSQL = "Select a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, a.诊疗类别, Null As 选择, Null As 输入, A.姓名, p.当前床号 As 床号," & vbNewLine & _
            "       Decode(Nvl(a.医嘱期效, 0), 0, '长嘱', '临嘱') As 期效, To_Char(a.开始执行时间, 'YYYY-MM-DD HH24:MI') As 开始时间, a.医嘱内容," & vbNewLine & _
            "       Decode(a.单次用量, Null, Null, a.单次用量 || b.计算单位) As 单量, Max(decode(g.中文名,'主刀医生', d.内容,'')) As 主刀医生, a.执行时间方案 As 执行时间方案, a.病人id," & vbNewLine & _
            "       a.主页id, a.诊疗项目id, a.频率次数, a.频率间隔, a.间隔单位, b.计算单位 As 单量单位, e.体重,e.入院日期,e.入院病况,A.年龄,A.性别,f.名称 as 护理等级,a.病人来源" & vbNewLine & _
            IIf(blnIsAudited, ", c.操作类型, c.操作说明, c.签名id ,a.医嘱状态,c.操作时间 as 审核时间", "") & _
            " From 病人医嘱记录 A, 病人信息 P, 诊疗项目目录 B, 病案主页 E,收费项目目录 F,诊治所见项目 G,病人医嘱附件 D" & vbNewLine & _
            IIf(blnIsAudited, ", (Select 医嘱id,操作时间,操作说明,操作类型,签名ID" & vbNewLine & _
                            "From (Select C.医嘱id,C.操作时间,C.操作说明,C.操作类型,C.签名ID, Row_Number() Over(Partition By C.医嘱id Order By C.操作时间 Desc) Top" & vbNewLine & _
                            "       From 病人医嘱状态 C" & vbNewLine & _
                            "       Where c.操作时间 Between [3] And [4] " & vbNewLine & _
                            "       and C.操作类型 in(11,12) And C.操作人员 =[2])" & vbNewLine & _
                            "Where Top = 1)  C", "") & _
            " Where a.病人id = p.病人id And a.诊疗项目id = b.Id  And f.id(+)=e.护理等级id  And" & vbNewLine & _
            "      e.病人id(+) = a.病人id And e.主页id(+) = a.主页id And d.医嘱ID(+)=a.id And d.要素id=g.id(+) And g.中文名(+)='主刀医生' " & _
            IIf(blnIsAudited, " And c.医嘱id = a.Id ", _
            " And A.开嘱时间 between [5] and [6]  And a.医嘱状态 = 1 And a.审核状态 = 1 ") & vbNewLine & _
            "    And a.开嘱医生=[1] And a.诊疗类别 ='F'  And a.相关ID is null " & _
            " Group By a.Id, a.相关id, a.诊疗类别,  A.姓名, p.当前床号,a.医嘱期效, a.开始执行时间, a.医嘱内容,a.单次用量, a.执行时间方案 , a.病人id, a.主页id, a.诊疗项目id, a.频率次数, a.频率间隔, a.间隔单位, b.计算单位, e.体重,e.入院日期,e.入院病况,A.年龄,A.性别,f.名称,a.病人来源" & _
            IIf(blnIsAudited, ",c.操作类型, c.操作说明, c.签名id ,a.医嘱状态,c.操作时间 ", "") & _
            " Order By A.姓名,To_Char(a.开始执行时间, 'YYYY-MM-DD HH24:MI'),Nvl(a.相关id, a.Id),a.id"
            
    On Error GoTo errH
    If rptDoc.SelectedRows.Count = 0 Then Exit Sub
    If rptDoc.SelectedRows(0).GroupRow Then vsAudit.Rows = 1: vsAudit.AddItem "": Exit Sub
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rptDoc.SelectedRows(0).Record(col_姓名).Value, UserInfo.姓名, CDate(dtpTime(0).Value), CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60), CDate(dtpTimeY(0).Value), CDate(dtpTimeY(1).Value + 1 - 1 / 24 / 60 / 60))
    
    With vsAudit
        If Val(.TextMatrix(.Row, col_医嘱ID)) <> 0 Then lngID = Val(.TextMatrix(.Row, col_医嘱ID))
        If Not blnIsAudited Then .Cell(flexcpPicture, 0, col_选择) = img16.ListImages("unCheck").Picture
        .Cell(flexcpPicture, 0, col_取消选择) = img16.ListImages("unCheck").Picture
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(i, col_病人姓名) = rsTmp!姓名 & ""
                .TextMatrix(i, col_期效) = rsTmp!期效 & ""
                .TextMatrix(i, COL_单量) = rsTmp!单量 & ""
                .TextMatrix(i, COL_主刀医生) = rsTmp!主刀医生 & ""
                .TextMatrix(i, col_开始时间) = rsTmp!开始时间 & ""
                .TextMatrix(i, col_医嘱ID) = rsTmp!ID & ""
                If Val(rsTmp!ID & "") = lngID And lngID <> 0 Then
                    .Row = i
                End If
                .TextMatrix(i, col_相关ID) = rsTmp!相关ID & ""
                .TextMatrix(i, col_性别) = rsTmp!性别 & ""
                .TextMatrix(i, col_年龄) = rsTmp!年龄 & ""
                .TextMatrix(i, col_体重) = rsTmp!体重 & ""
                .TextMatrix(i, col_入院时间) = rsTmp!入院日期 & ""
                .TextMatrix(i, col_床号) = rsTmp!床号 & ""
                .TextMatrix(i, col_护理等级) = rsTmp!护理等级 & ""
                .TextMatrix(i, col_病人Id) = rsTmp!病人ID & ""
                .TextMatrix(i, col_主页ID) = rsTmp!主页ID & ""
                .TextMatrix(i, col_组ID) = rsTmp!组ID & ""
                .TextMatrix(i, COL_诊疗类别) = rsTmp!诊疗类别 & ""
                .TextMatrix(i, col_病人来源) = rsTmp!病人来源 & ""
                .TextMatrix(i, col_病况) = rsTmp!入院病况 & ""
                '显示简洁模式下的医嘱内容
                strFormat = rsTmp!医嘱内容
'                If .TextMatrix(i, COL_频率) <> "一次性" Then
'                    blnDo = True
'                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[总量]") = 0
'                    If blnDo Then
'                        strTmp = .TextMatrix(i, COL_总量)
'                        If strTmp <> "" Then strFormat = strFormat & ",共" & strTmp
'                    End If
                    
'                    blnDo = True
'                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[单量]") = 0
'                    If blnDo Then
'                        strTmp = .TextMatrix(i, COL_单量)
'                        If strTmp <> "" Then strFormat = strFormat & ",每次" & strTmp
'                    End If
'                End If
                .TextMatrix(i, COL_医嘱内容) = strFormat
                If blnIsAudited Then
                    .TextMatrix(i, col_签名id) = rsTmp!签名id & ""
                    .TextMatrix(i, col_医嘱状态) = rsTmp!医嘱状态 & ""
                    .Cell(flexcpData, i, col_选择) = Val(rsTmp!操作类型 & "") - 10
                    .Cell(flexcpPicture, i, col_选择) = imgAdvice.ListImages(Val(.Cell(flexcpData, i, col_选择))).Picture
                    .Cell(flexcpPictureAlignment, i, col_选择) = flexPicAlignCenterCenter
                    .TextMatrix(i, col_审核说明) = rsTmp!操作说明 & ""
                    .TextMatrix(i, col_审核时间) = Format(rsTmp!审核时间 & "", "yyyy-MM-dd HH:mm:ss")
                    '如果医嘱不再是新开状态，则改变字体颜色
                    If Val(rsTmp!医嘱状态 & "") <> 1 Then
                        .Cell(flexcpForeColor, i, col_选择, i, col_签名id) = &HC00000
                    End If
                End If
                
                rsTmp.MoveNext
                i = i + 1
            Loop
            vsAudit.Cell(flexcpBackColor, 1, IIf(blnIsAudited, 1, 0), i - 1, col_审核说明) = &HFAEADA
            If blnIsAudited Then
                For j = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, j, col_选择)) = 1 Or (.TextMatrix(j, col_医嘱状态) & "" <> "1" And .TextMatrix(j, col_医嘱状态) & "" <> "") Then
                        .Cell(flexcpBackColor, j, col_审核说明) = &H80000005
                        If .TextMatrix(j, col_医嘱状态) & "" <> "1" And .TextMatrix(j, col_医嘱状态) & "" <> "" Then
                            '已校对的医嘱不允许修改或回退
                            .Cell(flexcpBackColor, j, col_取消选择) = &H80000005
                        End If
                    End If
                Next
            End If
        Else
            .AddItem ""
            vsAudit.Cell(flexcpBackColor, 1, IIf(blnIsAudited, 1, 0), 1, col_审核说明) = &HFAEADA
        End If
        If blnIsAudited Then
            stbThis.Panels(2).Text = cboTime.Text & "时间内【" & rptDoc.SelectedRows(0).Record(col_姓名).Value & "】有" & IIf(.TextMatrix(1, col_医嘱ID) = "", 0, .Rows - 1) & "条医嘱进行了审核。"
        Else
            stbThis.Panels(2).Text = "【" & rptDoc.SelectedRows(0).Record(col_姓名).Value & "】有" & IIf(.TextMatrix(1, col_医嘱ID) = "", 0, .Rows - 1) & "条医嘱需要审核。"
        End If
        '自动调整行高
        .AutoSize COL_医嘱内容
        .Redraw = flexRDDirect
        If .Row > 0 Then Call vsAudit_AfterRowColChange(1, 1, .Row, col_审核说明)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDoc
        strMsg = UCase(Trim(txtFind.Text))
        If zlcommfun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindNum To rptDoc.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(col_姓名).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Or _
                            .Rows(i).Record(IIf(mlngCodeType = 0, COL_拼音简码, COL_五笔简码)).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Then
                        '该行选中且显示在可见区域,并引发SelectionChanged事件
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                Else
                    If .Rows(i).Record(col_姓名).Value Like IIf(gstrLike = "", "", "*") & strMsg & "*" Then
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                End If
            End If
        Next
        If mlngFindNum = 0 Then
            MsgBox "当前部门没有找到您查找的医生。", vbInformation, Me.Caption
        ElseIf mlngFindNum <> 0 And blnIsFind = False Then
            MsgBox "已经是最后一个医生了。", vbInformation, Me.Caption
            mlngFindNum = 0
        End If
    End With
End Sub

Private Sub vsAudit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    With vsAudit
        If NewCol = col_审核说明 Or NewCol = col_选择 Or NewCol = col_取消选择 Then
            If (Val(vsAudit.Cell(flexcpData, NewRow, col_选择) & "") = "1" And NewCol = col_审核说明) Or _
                    (vsAudit.TextMatrix(NewRow, col_医嘱状态) & "" <> "1" And vsAudit.TextMatrix(NewRow, col_医嘱状态) & "" <> "" And NewCol = col_审核说明) _
                    Or (tbcSub.Selected.Tag = "已审核" And NewCol = col_选择) Then
                vsAudit.FocusRect = flexFocusNone
            Else
                If .TextMatrix(NewRow, col_医嘱状态) & "" <> "1" And .TextMatrix(NewRow, col_医嘱状态) & "" <> "" Then
                    vsAudit.FocusRect = flexFocusNone
                Else
                    vsAudit.FocusRect = flexFocusHeavy
                End If
            End If
        Else
            vsAudit.FocusRect = flexFocusNone
        End If
        
        '颜色
        .ForeColorSel = .Cell(flexcpForeColor, NewRow, NewCol)

        If vsAudit.TextMatrix(NewRow, col_医嘱ID) <> "" And NewRow <> 0 Then
            lblInformation(info_入院时间).Caption = Format(.TextMatrix(NewRow, col_入院时间), "yyyy-MM-dd")
            lblInformation(info_性别).Caption = .TextMatrix(NewRow, col_性别)
            lblInformation(info_年龄).Caption = .TextMatrix(NewRow, col_年龄)
            lblInformation(info_病况).Caption = .TextMatrix(NewRow, col_病况)
            lblInformation(info_床号).Caption = .TextMatrix(NewRow, col_床号)
            lblInformation(info_护理等级).Caption = .TextMatrix(NewRow, col_护理等级)
            lblInformation(info_体重).Caption = IIf(Val(.TextMatrix(NewRow, col_体重) & "") = 0, "", .TextMatrix(NewRow, col_体重) & "Kg")
            
            '过敏记录
            Call LoadPatiAllergy(Val(.TextMatrix(NewRow, col_病人Id) & ""), cbo过敏)
            
            '诊断
            lblInformation(info_诊断).Caption = GetPatiDiagnose(Val(.TextMatrix(NewRow, col_病人Id) & ""), _
            Val(.TextMatrix(NewRow, col_主页ID) & ""), _
            Val(.TextMatrix(NewRow, col_病人来源)))
            '住院信息显示
            picInShow.Visible = Not .TextMatrix(NewRow, col_病人来源) = "1"
        Else
            lblInformation(info_入院时间).Caption = ""
            lblInformation(info_性别).Caption = ""
            lblInformation(info_年龄).Caption = ""
            lblInformation(info_病况).Caption = ""
            lblInformation(info_床号).Caption = ""
            lblInformation(info_护理等级).Caption = ""
            lblInformation(info_体重).Caption = ""
            
            '过敏记录
            cbo过敏.Clear
            
            '诊断
            lblInformation(info_诊断).Caption = ""
            
            picInShow.Visible = True
        End If
    End With
End Sub

Private Sub vsAudit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col_审核说明) Then
        Cancel = True
    Else
        If Val(vsAudit.Cell(flexcpData, Row, col_选择) & "") = "1" Or vsAudit.TextMatrix(1, col_医嘱ID) & "" = "" Or _
                (vsAudit.TextMatrix(Row, col_医嘱状态) & "" <> "1" And vsAudit.TextMatrix(Row, col_医嘱状态) & "" <> "") Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAudit_Click()
    Dim i As Long
    
    With vsAudit
        If tbcSub.Selected.Tag = "已审核" Then
            If .MouseCol = col_取消选择 And .MouseRow = .FixedRows - 1 Then
                If .TextMatrix(1, col_医嘱ID) = "" Then Exit Sub
                If .ColData(col_取消选择) = "Check" Then
                    .Cell(flexcpPicture, 0, col_取消选择) = img16.ListImages("unCheck").Picture
                    .ColData(col_取消选择) = ""
                Else
                    .Cell(flexcpPicture, 0, col_取消选择) = img16.ListImages("Check").Picture
                    .ColData(col_取消选择) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, col_医嘱ID) = "" Then Exit For
                    If .ColData(col_取消选择) = "Check" Then
                        If Not (.TextMatrix(i, col_医嘱状态) & "" <> "1" And .TextMatrix(i, col_医嘱状态) & "" <> "") Then
                            .TextMatrix(i, col_取消选择) = -1
                        End If
                    Else
                        .TextMatrix(i, col_取消选择) = 0
                    End If
                    
                Next
            ElseIf .MouseCol = col_取消选择 And .MouseRow > .FixedRows - 1 And .MouseRow < .Rows Then
                 Call vsAudit_KeyPress(vbKeySpace)
            End If
        Else
            If .MouseCol = col_选择 And .MouseRow = .FixedRows - 1 Then
                If .TextMatrix(1, col_医嘱ID) = "" Then Exit Sub
                For i = 1 To .Rows - 1
                    If .ColData(col_选择) = "" Then
                        If .TextMatrix(i, col_审核说明) <> "" Then
                            If MsgBox("您已经填写了审核说明，修改为通过将删除说明，是否继续？", vbQuestion + vbDefaultButton1 + vbYesNo, Me.Caption) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                If .ColData(col_选择) = "Check" Then
                    .Cell(flexcpPicture, 0, col_选择) = img16.ListImages("unCheck").Picture
                    .ColData(col_选择) = ""
                Else
                    .Cell(flexcpPicture, 0, col_选择) = img16.ListImages("Check").Picture
                    .ColData(col_选择) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, col_医嘱ID) = "" Then Exit For
                    If .ColData(col_选择) = "Check" Then
                        .Cell(flexcpPicture, i, col_选择) = imgAdvice.ListImages(1).Picture
                        .Cell(flexcpData, i, col_选择) = 1
                        .Cell(flexcpPictureAlignment, i, col_选择) = flexPicAlignCenterCenter
                        vsAudit.Cell(flexcpBackColor, i, col_审核说明) = &H80000005
                        .TextMatrix(i, col_审核说明) = ""
                    Else
                        Set .Cell(flexcpPicture, i, col_选择) = Nothing
                        .Cell(flexcpData, i, col_选择) = 0
                        vsAudit.Cell(flexcpBackColor, i, col_审核说明) = &HFAEADA
                    End If
                    
                Next
                mblnIsUpdate = True
            End If
        End If
    End With
End Sub

Private Sub vsAudit_DblClick()
    With vsAudit
        If .MouseCol = col_选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAudit_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Function RowIn同一病人(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'功能：判断指定病人是否有其他医嘱
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If lngRow = 0 Then Exit Function
        If .TextMatrix(lngRow - 1, col_病人姓名) = .TextMatrix(lngRow, col_病人姓名) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If .TextMatrix(lngRow + 1, col_病人姓名) = .TextMatrix(lngRow, col_病人姓名) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If .TextMatrix(i, col_病人姓名) = .TextMatrix(lngRow, col_病人姓名) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If .TextMatrix(i, col_病人姓名) = .TextMatrix(lngRow, col_病人姓名) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn同一病人 = blnTmp
    End With
End Function

Private Sub vsAudit_KeyPress(KeyAscii As Integer)
    With vsAudit
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call UnAuditEnterNextCell
        ElseIf .Col = col_审核说明 Then
            .ComboList = "" '使按钮状态进入输入状态
        ElseIf .Col = col_选择 And KeyAscii = vbKeySpace Then
            Call AuditStateCheck
        ElseIf .Col = col_取消选择 And KeyAscii = vbKeySpace Then
            Call AuditCancleCheck
        End If
    End With
End Sub

Private Sub AuditCancleCheck()
'功能：已审核取消选择的同步选择一组药品
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim lngCheck As Long
    Dim blnIsAudit As Boolean   '判断医嘱是新开状态
    
    With vsAudit
        If tbcSub.Selected.Tag = "待审核" Then Exit Sub
        If .TextMatrix(.Row, col_医嘱ID) = "" Or (.TextMatrix(.Row, col_医嘱状态) & "" <> "1" And .TextMatrix(.Row, col_医嘱状态) & "" <> "") Then Exit Sub
        '如果是启用了签名参数，则检查是否有一起签名的，一起勾选。
        If gintCA = 0 Or (.TextMatrix(.Row, col_签名id) = "" And gintCA > 0 And 1 = 0) Then
            lngBegin = .Row: lngEnd = .Row
        Else
            If Not RowIn同一病人(.Row, lngBegin, lngEnd, vsAudit) Then
                lngBegin = .Row: lngEnd = .Row
            End If
        End If
        lngCheck = Val(.TextMatrix(lngBegin, col_取消选择))
        For i = lngBegin To lngEnd
            If gintCA = 0 Or (.TextMatrix(.Row, col_签名id) = "" And gintCA > 0 And 1 = 0) Then
                .TextMatrix(i, col_取消选择) = IIf(lngCheck = 0, -1, 0)
            Else
                If .TextMatrix(i, col_签名id) <> "" And .TextMatrix(i, col_签名id) = .TextMatrix(.Row, col_签名id) Then
                    If Val(.TextMatrix(i, col_医嘱状态)) = 1 Then
                        .TextMatrix(i, col_取消选择) = IIf(lngCheck = 0, -1, 0)
                    Else
                        blnIsAudit = True
                        Exit For
                    End If
                End If
                If i = lngEnd Then stbThis.Panels(2).Text = "一同勾选/取消的医嘱为批量签名审核的。"
            End If
        Next
        '如果含有不是新增的医嘱则取消选择，并提示
        If blnIsAudit Then
            For i = lngBegin To lngEnd
                If .TextMatrix(i, col_签名id) <> "" And .TextMatrix(i, col_签名id) = .TextMatrix(.Row, col_签名id) Then
                    .TextMatrix(i, col_取消选择) = 0
                End If
            Next
            MsgBox "发现有批量审核签名的医嘱已经校对，不能取消审核。", vbInformation, Me.Caption
        End If
    End With
End Sub

Private Sub AuditStateCheck(Optional ByVal lngState As Long)
'同步选择一组药品
'参数：lngState=0或者null 为进入下一个状态，1=√ ，2=？，3=待审核
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    
    With vsAudit
        If tbcSub.Selected.Tag = "已审核" Then Exit Sub
        If .TextMatrix(.Row, col_医嘱ID) = "" Or (.TextMatrix(.Row, col_医嘱状态) & "" <> "1" And .TextMatrix(.Row, col_医嘱状态) & "" <> "") Then Exit Sub
        lngBegin = .Row: lngEnd = .Row
        
        For i = lngBegin To lngEnd
            If lngState = 1 Or Val(.Cell(flexcpData, i, col_选择) & "") = 0 Then
                If .TextMatrix(i, col_审核说明) <> "" Then
                    If MsgBox("您已经填写了审核说明，修改为通过将删除说明，是否继续？", vbQuestion + vbDefaultButton1 + vbYesNo, Me.Caption) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        For i = lngBegin To lngEnd
            If lngState = 1 Or Val(.Cell(flexcpData, i, col_选择) & "") = 0 Then
                .TextMatrix(i, col_审核说明) = ""
            End If
            .Cell(flexcpData, i, col_选择) = IIf(lngState = 0, Val(.Cell(flexcpData, i, col_选择) & "") + IIf(Val(.Cell(flexcpData, i, col_选择) & "") = 2, -2, 1), IIf(lngState = 3, 0, lngState))
            If Val(.Cell(flexcpData, i, col_选择) & "") = 0 Then
                Set .Cell(flexcpPicture, i, col_选择) = Nothing
            Else
                .Cell(flexcpPicture, i, col_选择) = imgAdvice.ListImages(Val(.Cell(flexcpData, i, col_选择) & "")).Picture
            End If
            .Cell(flexcpPictureAlignment, i, col_选择) = flexPicAlignCenterCenter
            vsAudit.Cell(flexcpBackColor, i, col_审核说明) = IIf(Val(.Cell(flexcpData, i, col_选择) & "") = 1, &H80000005, &HFAEADA)
        Next
        mblnIsUpdate = True
    End With
End Sub


Private Sub vsAudit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = col_审核说明 Then
        If zlcommfun.ActualLen(vsAudit.Editable) - zlcommfun.ActualLen(vsAudit.EditSelText) >= 100 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
            If KeyAscii = vbKeyReturn Then
                Call UnAuditEnterNextCell
                Exit Sub
            End If
            KeyAscii = 0
        ElseIf Chr(KeyAscii) = "'" Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub UnAuditEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAudit
        If .Col = col_审核说明 Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlcommfun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAudit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If vsAudit.Rows <= 1 Then Exit Sub
    If vsAudit.TextMatrix(1, col_医嘱ID) <> "" And (vsAudit.MouseCol = col_选择 Or vsAudit.MouseCol = col_审核说明) And vsAudit.MouseRow = 0 And tbcSub.Selected.Tag = "待审核" Then
        strTip = "选中第一列的单元格按空格或双击可改变审核结果。" & vbCrLf & "？为不通过，√为通过。"
        zlcommfun.ShowTipInfo vsAudit.hwnd, strTip, True
    Else
        strTip = ""
        zlcommfun.ShowTipInfo 0, strTip, True
    End If
End Sub

Private Sub vsAudit_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = col_审核说明 Then
        vsAudit.EditSelStart = 0
        vsAudit.EditSelLength = Len(vsAudit.EditText)
    End If
End Sub

Private Sub vsAudit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col_审核说明 Then
        If vsAudit.EditText <> vsAudit.TextMatrix(Row, Col) Then
            If Val(vsAudit.Cell(flexcpData, Row, col_选择) & "") = "0" And tbcSub.Selected.Tag = "待审核" Then
                Call AuditStateCheck(2)
            End If
            mblnIsUpdate = True
            If tbcSub.Selected.Tag = "已审核" Then
                mstrChangeRows = mstrChangeRows & IIf(mstrChangeRows = "", "", ",") & Row
            End If
        End If
    End If
End Sub


