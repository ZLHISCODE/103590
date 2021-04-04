VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmStyle_CommonCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "窗体样式配置"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   Icon            =   "frmStyle_CommonCfg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkCallTarget 
      Caption         =   "显示就诊目的地"
      Height          =   255
      Left            =   2160
      TabIndex        =   41
      Top             =   5950
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd显示设备设置 
      Caption         =   "设备设置(&S)"
      Height          =   350
      Left            =   4040
      TabIndex        =   38
      Top             =   5880
      Width           =   1200
   End
   Begin VB.CheckBox chkScrollDisplay 
      Caption         =   "滚动显示已过号内容"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5950
      Width           =   1935
   End
   Begin TabDlg.SSTab sstFormSetup 
      Height          =   5640
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9948
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "液晶屏位置"
      TabPicture(0)   =   "frmStyle_CommonCfg.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "显示队列配置"
      TabPicture(1)   =   "frmStyle_CommonCfg.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfQueueSetup"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "窗体皮肤设置"
      TabPicture(2)   =   "frmStyle_CommonCfg.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(1)=   "cboStyleType"
      Tab(2).Control(2)=   "fraRemarkInfo"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "医生\科室信息"
      TabPicture(3)   =   "frmStyle_CommonCfg.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "fraDeptInfo"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame3 
         Height          =   700
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   7140
         Begin VB.ComboBox cboCurRoom 
            Height          =   300
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   240
            Width           =   2205
         End
         Begin VB.ComboBox cboCurDept 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "当前诊室/执行间"
            Height          =   180
            Left            =   3360
            TabIndex        =   46
            Top             =   285
            Width           =   720
         End
         Begin VB.Label lblCurDept 
            AutoSize        =   -1  'True
            Caption         =   "当前科室"
            Height          =   180
            Left            =   120
            TabIndex        =   45
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   -74880
         TabIndex        =   31
         Top             =   3600
         Width           =   7140
         Begin VB.CheckBox chkConvertQueueName 
            Caption         =   "转换成老版队列名称"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "将队列名称的存储格式转换为老版本的格式"
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkShowDeptName 
            Caption         =   "显示科室名"
            Height          =   255
            Left            =   3950
            TabIndex        =   37
            ToolTipText     =   "在样式窗口中显示诊室标题时，是否需要显示对应的科室名"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtQueueRows 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   6400
            TabIndex        =   33
            Top             =   220
            Width           =   375
         End
         Begin VB.CheckBox chkFontAutoSizeToList 
            Caption         =   "列表字体自适应"
            Height          =   255
            Left            =   2240
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblQueue 
            AutoSize        =   -1  'True
            Caption         =   "排队列表显示     行"
            Height          =   180
            Left            =   5300
            TabIndex        =   34
            Top             =   270
            Width           =   1710
         End
      End
      Begin VB.Frame fraRemarkInfo 
         Caption         =   "底端文本"
         Height          =   800
         Left            =   -74880
         TabIndex        =   29
         Top             =   4720
         Width           =   7140
         Begin VB.TextBox txtRemarkInfo 
            Appearance      =   0  'Flat
            Height          =   460
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   240
            Width           =   6945
         End
      End
      Begin VB.ComboBox cboStyleType 
         Height          =   300
         Left            =   -73990
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   400
         Width           =   2535
      End
      Begin VB.Frame Frame9 
         Caption         =   "皮肤设置"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   27
         Top             =   440
         Width           =   7140
         Begin VB.PictureBox picStyleView 
            BackColor       =   &H80000008&
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3675
            ScaleWidth      =   6840
            TabIndex        =   40
            Top             =   360
            Width           =   6900
            Begin VB.Image imgStyleView 
               Height          =   3780
               Left            =   0
               Picture         =   "frmStyle_CommonCfg.frx":007C
               Stretch         =   -1  'True
               Top             =   0
               Width           =   6900
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "显示数据过滤条件设置"
         Height          =   1140
         Left            =   -74880
         TabIndex        =   24
         Top             =   4320
         Width           =   7140
         Begin VB.TextBox txtFilter 
            Appearance      =   0  'Flat
            Height          =   705
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   300
            Width           =   6945
         End
      End
      Begin VB.Frame fraDeptInfo 
         Caption         =   "科室简介"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   22
         Top             =   3975
         Width           =   7140
         Begin VB.TextBox txtDeptInfo 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   300
            Width           =   6825
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "医生相关信息"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   18
         Top             =   1200
         Width           =   7140
         Begin VB.TextBox txtIntroduction 
            Appearance      =   0  'Flat
            Height          =   2025
            Index           =   0
            Left            =   4080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   500
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.CommandButton cmdSetDocPhoto 
            Caption         =   "照片设置(&S)"
            Height          =   350
            Left            =   2360
            TabIndex        =   20
            Top             =   2160
            Width           =   1605
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   2360
            ScaleHeight     =   1785
            ScaleWidth      =   1575
            TabIndex        =   19
            Top             =   285
            Width           =   1605
            Begin VB.Image imgDoctorPhoto 
               Height          =   1815
               Index           =   0
               Left            =   0
               Picture         =   "frmStyle_CommonCfg.frx":30D77
               Stretch         =   -1  'True
               Top             =   0
               Visible         =   0   'False
               Width           =   1605
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfDoctorInfo 
            Height          =   2235
            Left            =   120
            TabIndex        =   21
            Top             =   285
            Width           =   2115
            _cx             =   3731
            _cy             =   3942
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483638
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   2
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   272
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
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
         Begin VB.Label lblDoctorIntro 
            AutoSize        =   -1  'True
            Caption         =   "医生简介："
            Height          =   180
            Left            =   4080
            TabIndex        =   36
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5100
         Left            =   120
         TabIndex        =   3
         Top             =   400
         Width           =   7140
         Begin VB.OptionButton optFullScreen 
            Caption         =   "全屏"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   760
            Width           =   855
         End
         Begin VB.OptionButton optCustom 
            Caption         =   "自定义"
            Height          =   255
            Left            =   2160
            TabIndex        =   14
            Top             =   760
            Width           =   900
         End
         Begin VB.Frame frmCustom 
            Caption         =   "自定义位置(分辨率为单位)"
            Height          =   1245
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   6900
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   1
               Left            =   840
               TabIndex        =   9
               Top             =   360
               Width           =   2535
            End
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   2
               Left            =   840
               TabIndex        =   8
               Top             =   830
               Width           =   2535
            End
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   3
               Left            =   4440
               TabIndex        =   7
               Top             =   360
               Width           =   2295
            End
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   4
               Left            =   4440
               TabIndex        =   6
               Top             =   840
               Width           =   2295
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "左"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   13
               Top             =   405
               Width           =   180
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "顶"
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   12
               Top             =   888
               Width           =   180
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "宽度"
               Height          =   180
               Index           =   3
               Left            =   3840
               TabIndex        =   11
               Top             =   405
               Width           =   360
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "高度"
               Height          =   180
               Index           =   4
               Left            =   3840
               TabIndex        =   10
               Top             =   885
               Width           =   360
            End
         End
         Begin VB.ComboBox cboLCDNum 
            Height          =   300
            ItemData        =   "frmStyle_CommonCfg.frx":32128
            Left            =   1200
            List            =   "frmStyle_CommonCfg.frx":3212F
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   250
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "显示器编号"
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   900
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfQueueSetup 
         Height          =   3105
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   7140
         _cx             =   12594
         _cy             =   5477
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483638
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6540
      TabIndex        =   1
      Top             =   5895
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5400
      TabIndex        =   0
      Top             =   5895
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgDoctorPhoto 
      Left            =   4800
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmStyle_CommonCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'这里的配置窗口由主配置窗口frmMain调用
'在该通用样式配置窗口下，配置的内容包含如下：
'1.窗口显示位置，包含显示器编号，显示区域(left,top,right,bottom),当为全屏显示时，则不需要对显示区域进行配置
'2.当前排队叫号显示所在科室及诊室配置
'3.显示队列配置，如果业务类型为0表示属于门诊排队叫号的显示业务，业务类型定义可参考变量glngBusinessType，
'                不同的业务类型获取队列名称的方式也有所区别
'                   pacs业务队列名称规则为：科室名称 + "-" + 执行间 如“放射科-DR1室”
'4.自定义排队叫号显示数据过滤条件设置
'5.广告显示配置

Private mlngWindowNo As Long
Private mlngStyleType As Long
Private mobjParent As Object
Private mrsRecord As ADODB.Recordset
Private mstrDoctorPhoto() As String         '医生照片的十六进制串
Private mstrStyleTylePath As String
Private mintSelectedQueueNum As Integer     '已选择的队列数
Private mlngQueueListMaxRows As Long        '列表可以显示数据的总行数
Private mblnShowListHeader As Boolean       '显示列表标题
Private mrsDept As ADODB.Recordset
Private mstrCurDiagnoseRoom As String       '本机执行间
Private mstrPreSelectDept As String         '上一次选择的科室名

Public Function OpenShowConfig(ByVal lngWindowNo As Long, ByVal lngStyleType As TShowStyle, objOwner As Object) As Boolean
'显示配置窗口
    OpenShowConfig = False

    mlngWindowNo = lngWindowNo
    mlngStyleType = lngStyleType
    mintSelectedQueueNum = 0
    Set mobjParent = objOwner
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical, TBusinessType.btPacs, TBusinessType.btPeis
            Call InitFace
            
            Call InitQueueSetup
            
            Call InitDoctorInfo
            
            Call InitLocalPars
            
            Call Me.Show(vbModal, objOwner)
    End Select
    
    OpenShowConfig = True
End Function

Private Sub InitFace()
'根据显示样式初始化所包含内容
    Dim i As Integer
    
    If mlngStyleType = TShowStyle.ssSingleMan Then         '单病人
        cmd显示设备设置.Visible = False
        chkScrollDisplay.Visible = False
        lblQueue.Visible = False
        txtQueueRows.Visible = False
        fraRemarkInfo.Enabled = False
        
        lblDoctorIntro.Enabled = False
        For i = 0 To txtIntroduction.Count - 1
            txtIntroduction(i).Enabled = False
        Next
        fraDeptInfo.Enabled = False
    
    ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then   '单队列
        cmd显示设备设置.Visible = False
        chkCallTarget.Visible = True
    
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then   '多队列
        lblQueue.Visible = False
        txtQueueRows.Visible = False
        chkShowDeptName.Visible = False
        cmd显示设备设置.Visible = False
        sstFormSetup.TabVisible(3) = False
        
    ElseIf mlngStyleType = TShowStyle.ssOld Then   '老版本
        chkFontAutoSizeToList.Visible = False
        lblQueue.Visible = False
        txtQueueRows.Visible = False
        chkShowDeptName.Visible = False
        chkScrollDisplay.Visible = False
        
        sstFormSetup.TabVisible(0) = False
        sstFormSetup.TabVisible(2) = False
        sstFormSetup.TabVisible(3) = False
    End If
End Sub

Private Sub InitDoctorInfo()
'初始化医生相关信息配置
    Dim i As Integer
    Dim strDoctorInfo As String
    Dim strDoctorPhoto As String
    Dim strIntroduction As String
    Dim strWorkingTime As String
    
On Error GoTo ErrorHand

    If mlngStyleType <> TShowStyle.ssSingleMan And mlngStyleType <> TShowStyle.ssSingleQueue Then Exit Sub
    
    With vsfDoctorInfo
        .Cols = 3
        .Rows = 20
        ReDim mstrDoctorPhoto(.Rows - 2) As String
        
        For i = 1 To .Rows - 1
            Load imgDoctorPhoto(i)
            Load txtIntroduction(i)
            imgDoctorPhoto(i).Visible = True
            txtIntroduction(i).Visible = True
        Next
        
        .TextMatrix(0, 0) = "医生姓名"
        .TextMatrix(0, 1) = "值班时间"
        
        .Editable = flexEDKbdMouse
        
        .ColHidden(2) = True
        .ColComboList(1) = " |星期日|星期一|星期二|星期三|星期四|星期五|星期六"
        
        strDoctorInfo = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生信息")    '姓名和职位
        strDoctorPhoto = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生照片")    '
        strWorkingTime = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "值班时间")   '
        strIntroduction = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生简介")   '
        
        For i = 1 To .Rows - 1
            If strDoctorInfo <> "" Then
                vsfDoctorInfo.TextMatrix(i, 0) = Split(Split(Mid(strDoctorInfo, 2), "|")(i - 1), "-")(1)
                vsfDoctorInfo.TextMatrix(i, 2) = Split(Split(Mid(strDoctorInfo, 2), "|")(i - 1), "-")(0)
            End If

            If strDoctorPhoto <> "" Then
                Call LoadPictureInfo(imgDoctorPhoto(i), Split(Mid(strDoctorPhoto, 2), "|")(i - 1))
                mstrDoctorPhoto(i - 1) = Split(Mid(strDoctorPhoto, 2), "|")(i - 1)
            End If
            If strWorkingTime <> "" Then vsfDoctorInfo.TextMatrix(i, 1) = Split(Mid(strWorkingTime, 2), "|")(i - 1)
            If strIntroduction <> "" Then txtIntroduction(i).Text = Split(Mid(strIntroduction, 2), "|")(i - 1)
            
            vsfDoctorInfo.AutoSize 0, vsfDoctorInfo.Cols - 1
        Next
        
        If .Rows > 1 Then .RowSel = 1
        .AutoSize 0, .Cols - 1
        
        .Editable = flexEDNone
    End With

Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub LoadDoctorInfo(ByVal lngCurDeptID As Long)
'单队列时，根据现实队列配置，加载对应科室的医生名字等信息
    Dim i As Integer
    Dim strSql As String
    Dim strDoctorNames As String
    Dim str工作性质 As String
    
    vsfDoctorInfo.ColComboList(0) = ""
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            str工作性质 = "临床"
        Case TBusinessType.btPacs
            str工作性质 = "检查"
        Case TBusinessType.btPeis
            str工作性质 = "体检"
        'case
        '
        '
    End Select
    
    strSql = "select A.ID as 部门ID,A.名称,C.姓名,C.ID as 人员ID from 部门表 A,部门人员 B,人员表 C,部门性质说明 D " & _
             "Where A.ID=[1] And A.ID = B.部门ID And B.人员ID = C.ID And D.部门ID = A.ID " & _
             "And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
             "And D.工作性质 IN('" & str工作性质 & "') Order by A.编码"
    
    Set mrsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngCurDeptID)
    
    If mrsRecord.RecordCount <= 0 Then Exit Sub
    
    Do While Not mrsRecord.EOF
        strDoctorNames = strDoctorNames & "|" & Nvl(mrsRecord!姓名)
        mrsRecord.MoveNext
    Loop
    
   vsfDoctorInfo.Editable = flexEDKbdMouse
   vsfDoctorInfo.ColComboList(0) = " " & strDoctorNames
End Sub

Private Sub InitLocalPars()
'初始化样式配置参数
    Dim i As Integer, j As Integer
    Dim lngLCDNum As Long
    Dim lngCurLCDNo As Long
    Dim strCallingQueues As String
    Dim strLCDLocation As String
    Dim objForder As Folder
    Dim objFile As File

On Error GoTo ErrorHand
    '过滤条件
    txtFilter.Text = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "过滤条件", "")
    chkConvertQueueName.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "转换队列名称", 0))
    
    '排队列表显示的队列名
    strCallingQueues = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示队列")
    
    For i = 1 To vsfQueueSetup.Rows - 1
        For j = 0 To vsfQueueSetup.Cols - 1
            If InStr(strCallingQueues, vsfQueueSetup.TextMatrix(0, j) & "|" & vsfQueueSetup.TextMatrix(vsfQueueSetup.Rows - 1, j) & ":" & vsfQueueSetup.TextMatrix(i, j)) > 0 Then
                If vsfQueueSetup.TextMatrix(i, j) <> "" Then
                    vsfQueueSetup.Cell(flexcpChecked, i, j) = 1
                    
                    mintSelectedQueueNum = mintSelectedQueueNum + 1
                End If
            End If
        Next
    Next
    
    If mlngStyleType = TShowStyle.ssOld Then Exit Sub
    
    '加载所有科室
    mstrCurDiagnoseRoom = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "本机执行间", "")
    Call InitRoom
    
    If mlngStyleType = TShowStyle.ssMultiQueue Then
        vsfQueueSetup.ToolTipText = "您已选择了" & mintSelectedQueueNum & "个队列！" & vbCrLf & "最多允许选择" & mlngQueueListMaxRows & "个队列！"
    End If

    '显示模式,0-全屏；1-自定义
    If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示模式", 0)) = 0 Then
        optFullScreen.value = True
    Else
        optCustom.value = True
    End If
    
    '自定义显示位置
    If optCustom.value Then
        strLCDLocation = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "自定义位置")
        
        For i = 0 To UBound(Split(strLCDLocation, "|"))
            txtRect(i + 1).Text = Mid(Split(strLCDLocation, "|")(i), 3)
        Next
    End If
    
    '加载显示器编号
    Call InitMonitor
    
    lngLCDNum = UBound(gmonitors) - 1
    lngCurLCDNo = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示器编号")) - 1
    
    cboLCDNum.Clear
    
    For i = 1 To lngLCDNum
        cboLCDNum.AddItem i
        If i - 1 = lngCurLCDNo Then cboLCDNum.ListIndex = i - 1
    Next
    
    If cboLCDNum.ListIndex < 0 Then cboLCDNum.ListIndex = 0
    
    '加载队列样式
    cboStyleType.Clear
    
    If mlngStyleType = TShowStyle.ssSingleMan Then
        If gobjFile.FolderExists(App.Path & "\Skin\单病人样式") Then
            Set objForder = gobjFile.GetFolder(App.Path & "\Skin\单病人样式")
        Else
            Set objForder = gobjFile.GetFolder(App.Path & "\zlQueueShow\Skin\单病人样式")
        End If
    ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\单队列样式") Then
            Set objForder = gobjFile.GetFolder(App.Path & "\Skin\单队列样式")
        Else
            Set objForder = gobjFile.GetFolder(App.Path & "\zlQueueShow\Skin\单队列样式")
        End If
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\多队列样式") Then
            Set objForder = gobjFile.GetFolder(App.Path & "\Skin\多队列样式")
        Else
            Set objForder = gobjFile.GetFolder(App.Path & "\zlQueueShow\Skin\多队列样式")
        End If
    End If
    
    For Each objFile In objForder.Files
        If Mid(objFile.Name, Len(objFile.Name) - 2) = "jpg" Then
            cboStyleType.AddItem Mid(objFile.Name, 1, Len(objFile.Name) - 4)
            
            If objFile.Path = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式") & ".jpg" Then
                cboStyleType.Text = Mid(objFile.Name, 1, Len(objFile.Name) - 4)
            End If
        End If
    Next
    
    If cboStyleType.ListCount > 0 And cboStyleType.ListIndex < 0 Then cboStyleType.ListIndex = 0
    If cboStyleType.ListIndex >= 0 Then
        If mlngStyleType = TShowStyle.ssSingleMan Then
            '''''''''''''
            
        ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then
            If gobjFile.FolderExists(App.Path & "\Skin\单队列样式") Then
                Call SetIniFile(App.Path & "\Skin\单队列样式\" & cboStyleType.Text & ".ini")
            Else
                Call SetIniFile(App.Path & "\zlQueueShow\Skin\单队列样式\" & cboStyleType.Text & ".ini")
            End If
            
            mblnShowListHeader = CBool(ReadValue("排队列表区域", "是否显示列表标题"))
            
            If mblnShowListHeader Then
                mlngQueueListMaxRows = Val(ReadValue("排队列表区域", "总行数")) - 1
            Else
                mlngQueueListMaxRows = Val(ReadValue("排队列表区域", "总行数"))
            End If
        Else
            If gobjFile.FolderExists(App.Path & "\Skin\多队列样式") Then
                Call SetIniFile(App.Path & "\Skin\多队列样式\" & cboStyleType.Text & ".ini")
            Else
                Call SetIniFile(App.Path & "\zlQueueShow\Skin\多队列样式\" & cboStyleType.Text & ".ini")
            End If
            
            mlngQueueListMaxRows = Val(ReadValue("准备就诊列表区域", "总行数")) - 1
        End If
    End If
    
    chkFontAutoSizeToList.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "列表字体自适应", 1))
    chkShowDeptName.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "诊室标题是否显示科室名", 0))
    
    If mlngStyleType = TShowStyle.ssSingleMan Then Exit Sub
    
    '是否滚动显示呼叫信息
    chkScrollDisplay.value = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "滚动显示", 1)
    
    If mlngStyleType = TShowStyle.ssSingleQueue Then
        chkCallTarget.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示就诊目的地", 0))
        txtQueueRows.Text = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "排队列表显示行", mlngQueueListMaxRows))
        txtQueueRows.ToolTipText = "排队列表数据显示行数，最多显示" & mlngQueueListMaxRows & "行"
        txtDeptInfo.Text = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "科室简介")
    End If
    
    '底端文本
    txtRemarkInfo.Text = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "底端文本", "请未叫到号的患者耐心等待!")
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub InitRoom()
'功能：加载所有执行科室
    Dim i As Integer
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
    cboCurDept.Clear
    
    If mrsDept.RecordCount <= 0 Then Exit Sub
    mrsDept.MoveFirst
    
    For i = 0 To mrsDept.RecordCount - 1
        cboCurDept.AddItem Nvl(mrsDept!名称)
        cboCurDept.ItemData(i) = Nvl(mrsDept!ID)
        
        If mstrCurDiagnoseRoom <> "" Then
            If Nvl(mrsDept!名称) = Split(mstrCurDiagnoseRoom, "-")(0) Then
                mstrPreSelectDept = Nvl(mrsDept!名称)
                cboCurDept.ListIndex = i
            End If
        End If
        
        mrsDept.MoveNext
    Next
    
    If cboCurDept.ListCount > 0 And cboCurDept.ListIndex < 0 Then
        strSql = "select d.名称 from 上机人员表 A,人员表 B,部门人员 C,部门表 D " & _
                 "where A.人员ID=B.ID And b.id=c.人员id and c.部门id=d.id and c.缺省=1 and A.用户名=[1]"
        
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))
        
        If rsRecord.RecordCount > 0 Then
            For i = 0 To cboCurDept.ListCount - 1
                If cboCurDept.List(i) = Nvl(rsRecord!名称) Then
                    mstrPreSelectDept = Nvl(rsRecord!名称)
                    cboCurDept.ListIndex = i
                End If
            Next
        End If
    End If
    
    If cboCurDept.ListCount > 0 And cboCurDept.ListIndex < 0 Then cboCurDept.ListIndex = 0
End Sub

Private Sub InitQueueSetup()
'加载队列信息
    Dim i As Integer, j As Integer
    Dim intQueueNum As Integer  '每个科室对应的队列数
    Dim str来源 As String, str工作性质 As String
    Dim strSql As String
    Dim rsRoom As ADODB.Recordset
    Dim lngRoomMaxNum As Long

On Error GoTo ErrorHand
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            str工作性质 = "临床"
            str来源 = ",1,3,"
        Case TBusinessType.btPacs
            str工作性质 = "检查"
            str来源 = ",1,2,3,"
        Case TBusinessType.btPeis
            str工作性质 = "体检"
            str来源 = ",1,2,3,"
        'case
        '
        '
    End Select
    
    strSql = "Select Distinct A.ID,A.编码,A.名称 From 部门表 A,部门性质说明 B " & _
             "Where B.部门ID = A.ID  And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
             "And B.工作性质 IN('" & str工作性质 & "') And instr('" & str来源 & "',','||B.服务对象||',')> 0 Order by A.编码"
    
    Set mrsDept = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "显示队列配置")
    
    If mrsDept.RecordCount <= 0 Then Exit Sub
    
    '根据不同业务查找对应队列配置
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "select 科室id,房间," & _
                           "case " & _
                                "when instr(房间, '一') > 0 then replace(房间, '一', '1') " & _
                                "when instr(房间, '二') > 0 then replace(房间, '二', '2') " & _
                                "when instr(房间, '三') > 0 then replace(房间, '三', '3') " & _
                                "when instr(房间, '四') > 0 then replace(房间, '四', '4') " & _
                                "when instr(房间, '五') > 0 then replace(房间, '五', '5') " & _
                                "when instr(房间, '六') > 0 then replace(房间, '六', '6') " & _
                                "when instr(房间, '七') > 0 then replace(房间, '七', '7') " & _
                                "when instr(房间, '八') > 0 then replace(房间, '八', '8') " & _
                                "else replace(房间, '九', '9') " & _
                           "end As ord " & _
                    "from ( Select Distinct P.科室ID, '执行间-'||R.名称 as 房间 From 门诊诊室 R, 挂号安排诊室 S, 挂号安排 P " & _
                    "Where R.名称 = S.门诊诊室 And S.号表id = P.ID ) a order by 科室id,ord "
            
        Case TBusinessType.btPacs
            If gstrCompareVersion < "010.034.000" Then
                strSql = "select 科室id,房间," & _
                              "case " & _
                                  "when instr(房间, '一') > 0 then replace(房间, '一', '1') " & _
                                  "when instr(房间, '二') > 0 then replace(房间, '二', '2') " & _
                                  "when instr(房间, '三') > 0 then replace(房间, '三', '3') " & _
                                  "when instr(房间, '四') > 0 then replace(房间, '四', '4') " & _
                                  "when instr(房间, '五') > 0 then replace(房间, '五', '5') " & _
                                  "when instr(房间, '六') > 0 then replace(房间, '六', '6') " & _
                                  "when instr(房间, '七') > 0 then replace(房间, '七', '7') " & _
                                  "when instr(房间, '八') > 0 then replace(房间, '八', '8') " & _
                                  "else replace(房间, '九', '9') end As ord " & _
                              "from (select A.科室id, '执行间-'||执行间 as 房间 from 医技执行房间 A,部门表 B,部门性质说明 C " & _
                                  "Where A.科室ID=B.id And B.ID=C.部门ID And C.工作性质='" & str工作性质 & "' And 科室id not in " & _
                                  "(select 科室id from 影像流程参数 where 参数名='排队叫号方式' and 参数值=1 ) " & _
                                  "union select 科室id,'执行间-科室队列' as 房间 from 影像流程参数 where 参数名='排队叫号方式' and 参数值=1) " & _
                              "order by 科室id,ord"
            Else
                strSql = "select 科室id,房间,ord from " & _
                         "(select 科室id,房间, " & _
                            "case " & _
                                "when instr(房间, '一') > 0 then replace(房间, '一', '1') " & _
                                "when instr(房间, '二') > 0 then replace(房间, '二', '2') " & _
                                "when instr(房间, '三') > 0 then replace(房间, '三', '3') " & _
                                "when instr(房间, '四') > 0 then replace(房间, '四', '4') " & _
                                "when instr(房间, '五') > 0 then replace(房间, '五', '5') " & _
                                "when instr(房间, '六') > 0 then replace(房间, '六', '6') " & _
                                "when instr(房间, '七') > 0 then replace(房间, '七', '7') " & _
                                "when instr(房间, '八') > 0 then replace(房间, '八', '8') " & _
                                "else replace(房间, '九', '9') " & _
                            "end As ord " & _
                        "from (select 科室id, '执行间-'||执行间 as 房间 from 医技执行房间 A,部门表 B,部门性质说明 C " & _
                              "Where A.科室ID=B.id And B.ID=C.部门ID And C.工作性质='" & str工作性质 & "') " & _
                        "union select b.id,'执行分组-'||a.组名 房间,'周' as ord from 影像执行分组 a,部门表 b where a.科室id=b.id) " & _
                        "order by 科室id,ord"

            End If
            
        Case TBusinessType.btPeis
            strSql = "select 科室id,房间," & _
                           "case " & _
                                "when instr(房间, '一') > 0 then replace(房间, '一', '1') " & _
                                "when instr(房间, '二') > 0 then replace(房间, '二', '2') " & _
                                "when instr(房间, '三') > 0 then replace(房间, '三', '3') " & _
                                "when instr(房间, '四') > 0 then replace(房间, '四', '4') " & _
                                "when instr(房间, '五') > 0 then replace(房间, '五', '5') " & _
                                "when instr(房间, '六') > 0 then replace(房间, '六', '6') " & _
                                "when instr(房间, '七') > 0 then replace(房间, '七', '7') " & _
                                "when instr(房间, '八') > 0 then replace(房间, '八', '8') " & _
                                "else replace(房间, '九', '9') " & _
                          "end As ord " & _
                    "from (select 科室id, '执行间-'||执行间 as 房间 from 医技执行房间 A,部门表 B,部门性质说明 C  " & _
                     "Where A.科室ID=B.id And B.ID=C.部门ID And C.工作性质='" & str工作性质 & "')" & _
                    "order by 科室id,ord"""

        'case
        '.
        '.
    End Select
    
    Set rsRoom = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "显示队列配置")
    
    With vsfQueueSetup
        .Editable = flexEDKbdMouse
        .Cols = mrsDept.RecordCount
        
        If rsRoom.RecordCount > 0 Then
            .Rows = rsRoom.RecordCount + 3
        Else
            .Rows = 2
        End If
        
        If glngBusinessType = TBusinessType.btClinical And (mlngStyleType = TShowStyle.ssSingleQueue Or mlngStyleType = TShowStyle.ssMultiQueue) Then
            .Rows = .Rows + 1
        End If
        
        mrsDept.MoveFirst

        For i = 0 To .Cols - 1
            '加载列头，即科室名称
            .TextMatrix(0, i) = Nvl(mrsDept!名称)
            
            intQueueNum = 0
            
            If rsRoom.RecordCount > 0 Then
                '加载对应科室执行间
                rsRoom.MoveFirst
                For j = 0 To rsRoom.RecordCount - 1
                    If Nvl(mrsDept!ID) = Nvl(rsRoom!科室id) Then
                        intQueueNum = intQueueNum + 1
                        
                        .TextMatrix(intQueueNum, i) = Split(Nvl(rsRoom!房间), "-")(1)
                        
                        '对分组加以颜色区分
                        Select Case Split(Nvl(rsRoom!房间), "-")(0)
                            Case "执行间"
                                '''''
                            Case "执行分组"
                                .Cell(flexcpForeColor, intQueueNum, i) = vbRed
                        End Select
                        
                        .Cell(flexcpChecked, intQueueNum, i) = 2
                    End If
                    
                    rsRoom.MoveNext
                Next
            End If

            If glngBusinessType = TBusinessType.btPacs And gstrCompareVersion >= "010.034.000" Then 'PACS业务增加"未分配队列"
                intQueueNum = intQueueNum + 1
                .TextMatrix(intQueueNum, i) = "未分配队列"
                .Cell(flexcpChecked, intQueueNum, i) = 2
                .Cell(flexcpForeColor, intQueueNum, i) = vbRed
            End If
            
            '临川业务，单队和多队列时增加科室队列，用于显示整个科室的排队信息
            If glngBusinessType = TBusinessType.btClinical And (mlngStyleType = TShowStyle.ssSingleQueue Or mlngStyleType = TShowStyle.ssMultiQueue) Then
                intQueueNum = intQueueNum + 1
                .TextMatrix(intQueueNum, i) = "科室队列"
                .Cell(flexcpChecked, intQueueNum, i) = 2
                .Cell(flexcpForeColor, intQueueNum, i) = vbRed
            End If
            
            If intQueueNum = 0 Then .ColHidden(i) = True     '没有执行间时，样式为单队列时无需显示科室
            If intQueueNum > lngRoomMaxNum Then lngRoomMaxNum = intQueueNum
            
            '存储对应列科室的ID和编码，此行不显示
            .TextMatrix(.Rows - 1, i) = Nvl(mrsDept!ID) & "_" & Nvl(mrsDept!编码)
        
            mrsDept.MoveNext
        Next
        
        For i = lngRoomMaxNum + 1 To .Rows - 1
            .RowHidden(i) = True
        Next
        
        '自动列宽
        .AutoSize 0, .Cols - 1
        
        '最后一列自动填充满列表
        .ExtendLastCol = True
    End With
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub cboCurDept_Click()
'加载当前科室下的执行房间/诊室
    Dim strSql As String
    Dim rsRoom As ADODB.Recordset
On Error GoTo ErrorHand
    If mstrPreSelectDept <> "" Then
        If mstrPreSelectDept <> cboCurDept.Text Then
            If MsgBox("此操作会清除所有医生的相关信息配置，是否继续？", vbYesNo + vbDefaultButton2) = vbNo Then
                cboCurDept.Text = mstrPreSelectDept
                Exit Sub
            Else
                '清除医生相关信息配置
                Call ClearDoctorInfo
            End If
        End If
    End If
    
    mstrPreSelectDept = cboCurDept.Text
    
    '加载医生相关信息
    Call LoadDoctorInfo(Val(cboCurDept.ItemData(cboCurDept.ListIndex)))
    
    '加载诊室
    cboCurRoom.Clear
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "Select Distinct P.科室ID, '执行间-'||R.名称 as 房间 From 门诊诊室 R, 挂号安排诊室 S, 挂号安排 P " & _
                     "Where R.名称 = S.门诊诊室 And S.号表id = P.ID and p.科室id=[1]  And R.缺省标志 <> 1 "
                     
        Case TBusinessType.btPacs
            strSql = "select 科室id, '执行间-'||执行间 as 房间 from 医技执行房间 A,部门表 B,部门性质说明 C " & _
                     "Where A.科室ID=B.id And B.ID=C.部门ID and a.科室id=[1] And C.工作性质='检查'"
                     
        Case TBusinessType.btPeis
            strSql = "select 科室id, '执行间-'||执行间 as 房间 from 医技执行房间 A,部门表 B,部门性质说明 C  " & _
                   "Where A.科室ID=B.id And B.ID=C.部门ID and a.科室id=[1]  And C.工作性质='体检'"
                   
    End Select
    
    Set rsRoom = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "查询科室执行间", Val(cboCurDept.ItemData(cboCurDept.ListIndex)))
    
    If rsRoom.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsRoom.EOF
        cboCurRoom.AddItem Split(Nvl(rsRoom!房间), "-")(1)
        
        If mstrCurDiagnoseRoom <> "" Then
            If Split(Nvl(rsRoom!房间), "-")(1) = Split(mstrCurDiagnoseRoom, "-")(1) Then cboCurRoom.Text = Split(Nvl(rsRoom!房间), "-")(1)
        End If
        
        rsRoom.MoveNext
    Loop
    
    If cboCurRoom.ListCount > 0 And cboCurRoom.ListIndex < 0 Then cboCurRoom.ListIndex = 0
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cboStyleType_Click()
    If mlngStyleType = TShowStyle.ssSingleMan Then
        If gobjFile.FolderExists(App.Path & "\Skin\单病人样式") Then
            mstrStyleTylePath = App.Path & "\Skin\单病人样式\" & cboStyleType.Text & ".jpg"
        Else
            mstrStyleTylePath = App.Path & "\zlQueueShow\Skin\单病人样式\" & cboStyleType.Text & ".jpg"
        End If
    ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\单队列样式") Then
            mstrStyleTylePath = App.Path & "\Skin\单队列样式\" & cboStyleType.Text & ".jpg"
        Else
            mstrStyleTylePath = App.Path & "\zlQueueShow\Skin\单队列样式\" & cboStyleType.Text & ".jpg"
        End If
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\多队列样式") Then
            mstrStyleTylePath = App.Path & "\Skin\多队列样式\" & cboStyleType.Text & ".jpg"
        Else
            mstrStyleTylePath = App.Path & "\zlQueueShow\Skin\多队列样式\" & cboStyleType.Text & ".jpg"
        End If
    End If
    
    If gobjFile.FileExists(mstrStyleTylePath) Then
        imgStyleView.Picture = LoadPicture(mstrStyleTylePath)
    
        Call ResizeImg(imgStyleView, 0, 0, picStyleView.Width, picStyleView.Height)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'保存配置
On Error GoTo ErrorHand
    Call SaveLocalPars
    
    Unload Me
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub SaveLocalPars()
'保存配置
    Dim i As Integer, j As Integer
    Dim strCallingQueues As String
    Dim strDoctorInfo As String     '保存格式："医生1的姓名和职位|医生2的姓名和职位|。。。。"
    Dim strDoctorPhoto As String    '保存格式："医生1的照片|医生2的照片|。。。。"
    Dim strIntroduction As String   '保存格式："医生1的简介|医生2的简介|。。。。"
    Dim strWorkingTime As String    '保存格式："医生1的值班时间|医生2的值班时间|。。。。"
    
    For j = 0 To vsfQueueSetup.Cols - 1
        For i = 1 To vsfQueueSetup.Rows - 1
            If vsfQueueSetup.Cell(flexcpChecked, i, j) = 1 Then
                strCallingQueues = strCallingQueues & "," & vsfQueueSetup.TextMatrix(0, j) & "|" & vsfQueueSetup.TextMatrix(vsfQueueSetup.Rows - 1, j) & ":" & vsfQueueSetup.TextMatrix(i, j)
            End If
        Next
    Next
    '保存设置的队列，格式："科室名称|科室ID:执行间"，如："放射科|64:CT一室"
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示队列", Mid(strCallingQueues, 2)
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "过滤条件", txtFilter.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "转换队列名称", chkConvertQueueName.value
    
    If mlngStyleType = TShowStyle.ssOld Then Exit Sub
    
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示模式", IIf(optFullScreen.value = True, 0, 1)
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "自定义位置", "左:" & Val(txtRect(1).Text) & "|顶:" & Val(txtRect(2).Text) & "|宽:" & Val(txtRect(3).Text) & "|高:" & Val(txtRect(4).Text)
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示器编号", cboLCDNum.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "列表字体自适应", chkFontAutoSizeToList.value
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "诊室标题是否显示科室名", chkShowDeptName.value
    
    If cboCurDept.Text <> "" Then
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "本机执行间", cboCurDept.Text & "-" & cboCurRoom.Text
    Else
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "本机执行间", ""
    End If

    If mlngStyleType = TShowStyle.ssSingleMan Or mlngStyleType = TShowStyle.ssSingleQueue Then
        For i = 1 To vsfDoctorInfo.Rows - 1
            strDoctorInfo = strDoctorInfo & "|" & Nvl(vsfDoctorInfo.TextMatrix(i, 2)) & "-" & Nvl(vsfDoctorInfo.TextMatrix(i, 0))
            strDoctorPhoto = strDoctorPhoto & "|" & mstrDoctorPhoto(i - 1)
            strWorkingTime = strWorkingTime & "|" & Nvl(vsfDoctorInfo.TextMatrix(i, 1))
            strIntroduction = strIntroduction & "|" & txtIntroduction(i)
        Next
        
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生信息", strDoctorInfo     '姓名和职位
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生照片", strDoctorPhoto    '
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "值班时间", strWorkingTime    '
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "医生简介", strIntroduction   '
        
        If mlngStyleType = TShowStyle.ssSingleMan Then
            If gobjFile.FolderExists(App.Path & "\Skin\单病人样式") Then
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\Skin\单病人样式\" & cboStyleType.Text
            Else
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\zlQueueShow\Skin\单病人样式\" & cboStyleType.Text
            End If
            
            Exit Sub
        Else
            SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "显示就诊目的地", chkCallTarget.value
            
            If gobjFile.FolderExists(App.Path & "\Skin\单队列样式") Then
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\Skin\单队列样式\" & cboStyleType.Text
            Else
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\zlQueueShow\Skin\单队列样式\" & cboStyleType.Text
            End If
        End If
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\多队列样式") Then
            SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\Skin\多队列样式\" & cboStyleType.Text
        Else
            SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "皮肤样式", App.Path & "\zlQueueShow\Skin\多队列样式\" & cboStyleType.Text
        End If
    End If
    
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "底端文本", txtRemarkInfo.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "滚动显示", chkScrollDisplay.value
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "排队列表显示行", txtQueueRows.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "科室简介", txtDeptInfo.Text
End Sub

Private Sub cmdSetDocPhoto_Click()
    Dim strFileName As String
    Dim arrByte() As Byte
    Dim arrPic() As String
    Dim lngCount As Long, lngFileSize As Long
On Error GoTo ErrorHand
    dlgDoctorPhoto.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|(*.bmp)|*.bmp|(*.*)|*.*"
    dlgDoctorPhoto.ShowOpen

    strFileName = dlgDoctorPhoto.FileName

    If strFileName = "" Then Exit Sub

    '读取文件长度
    lngFileSize = FileLen(strFileName)

    ReDim arrByte(0 To lngFileSize - 1) '定义数值长度
    ReDim arrPic(0 To lngFileSize - 1) '定义数值长度

    Open strFileName For Binary As #1
    Get #1, , arrByte
    Close #1

    '将字节转换为16进制
    For lngCount = LBound(arrByte) To UBound(arrByte)
        arrPic(lngCount) = Hex(arrByte(lngCount))
        If Len(arrPic(lngCount)) = 1 Then arrPic(lngCount) = "0" & arrPic(lngCount)
    Next

    imgDoctorPhoto(vsfDoctorInfo.RowSel).Picture = LoadPicture(strFileName)
    
    mstrDoctorPhoto(vsfDoctorInfo.RowSel - 1) = Join(arrPic, "")
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmd显示设备设置_Click()
    Call InitOldLCDShow
    
    Call gobjQueueShow.zlSetup(Me)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If KeyAscii = vbKeyEscape Then Unload Me
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub optCustom_Click()
    Dim i As Integer

On Error GoTo ErrorHand
    frmCustom.Enabled = True
    
    For i = 1 To txtRect.Count
        lblRect(i).Enabled = True
        txtRect(i).Enabled = True
    Next
    
    txtRect(1).Text = 0
    txtRect(2).Text = 0
    txtRect(3).Text = Screen.Width / Screen.TwipsPerPixelX
    txtRect(4).Text = Screen.Height / Screen.TwipsPerPixelY
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub optFullScreen_Click()
    Dim i As Integer
On Error GoTo ErrorHand
    frmCustom.Enabled = False
    
    For i = 1 To txtRect.Count
        lblRect(i).Enabled = False
        txtRect(i).Enabled = False
        txtRect(i).Text = ""
    Next
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtQueueRows_Change()
On Error GoTo ErrorHand

    If txtQueueRows.Text <= 0 Then txtQueueRows.Text = 1
    If txtQueueRows.Text > mlngQueueListMaxRows Then txtQueueRows.Text = mlngQueueListMaxRows
    txtQueueRows.Text = Val(txtQueueRows.Text)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtQueueRows_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If InStr("01234567890." & Chr(8), Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtRect_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrorHand
    If InStr("01234567890." & Chr(8), Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHand
    With vsfDoctorInfo
        If .ColSel <> 0 Then Exit Sub
        If vsfDoctorInfo.TextMatrix(.RowSel, .ColSel) = "" Then Exit Sub

        mrsRecord.Filter = ""
        mrsRecord.Filter = "姓名='" & Trim(vsfDoctorInfo.TextMatrix(.RowSel, .ColSel)) & "'"

        If mrsRecord.RecordCount > 0 Then vsfDoctorInfo.TextMatrix(.RowSel, 2) = Nvl(mrsRecord!人员id)
        
        .AutoSize 0, vsfDoctorInfo.Cols - 1
    End With
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrorHand
    FinishEdit = True
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_KeyDown(KeyCode As Integer, Shift As Integer)
'在配置列表中按下“delete”键时，提示是否删除医生相关信息配置
    
On Error GoTo ErrorHand
    If KeyCode = vbKeyDelete Then  '清除医生相关信息配置
        If MsgBox("此操作会清除所有医生的相关信息配置，是否继续？", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        
        Call ClearDoctorInfo
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub ClearDoctorInfo()
'清除医生相关信息配置
    Dim i As Integer, j As Integer
    
    For i = 1 To vsfDoctorInfo.Rows - 1
        For j = 0 To vsfDoctorInfo.Cols - 1
            vsfDoctorInfo.TextMatrix(i, j) = ""
        Next
        
        imgDoctorPhoto(i).Picture = imgDoctorPhoto(0).Picture
    Next
    
    For i = 0 To UBound(mstrDoctorPhoto)
        mstrDoctorPhoto(i) = ""
    Next
    
    For i = 0 To txtIntroduction.Count - 1
        txtIntroduction(i).Text = ""
    Next
End Sub

Private Sub vsfQueueSetup_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrorHand
    Cancel = True
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfQueueSetup_Click()
    Dim i As Integer, j As Integer
    Dim lngRowSel As Long, lngColSel As Long

On Error GoTo ErrorHand

    lngRowSel = vsfQueueSetup.RowSel
    lngColSel = vsfQueueSetup.ColSel
    
    If lngRowSel <= 0 Then Exit Sub
    If vsfQueueSetup.TextMatrix(lngRowSel, lngColSel) = "" Then Exit Sub

    Select Case mlngStyleType
        Case TShowStyle.ssSingleMan
            For i = 1 To vsfQueueSetup.Rows - 1
                For j = 0 To vsfQueueSetup.Cols - 1
                    '在PACS排队业务下，单病人时能选择同一个科室下的一个或多个队列
                    If glngBusinessType = TBusinessType.btPacs Then
                        If vsfQueueSetup.TextMatrix(i, j) <> "" And j <> lngColSel Then
                            vsfQueueSetup.Cell(flexcpChecked, i, j) = 2
                        End If
                    Else    '其它业务下最多只能选择一个队列
                        If vsfQueueSetup.TextMatrix(i, j) <> "" And (i <> lngRowSel Or j <> lngColSel) Then
                            vsfQueueSetup.Cell(flexcpChecked, i, j) = 2
                        End If
                    End If
                Next
            Next
            
        Case TShowStyle.ssSingleQueue  '单队列时只能选择一个执行间
            For i = 1 To vsfQueueSetup.Rows - 1
                For j = 0 To vsfQueueSetup.Cols - 1
                    If vsfQueueSetup.TextMatrix(i, j) <> "" And (i <> lngRowSel Or j <> lngColSel) Then
                        vsfQueueSetup.Cell(flexcpChecked, i, j) = 2
                    End If
                Next
            Next

        Case TShowStyle.ssMultiQueue  '多队列时不限制选择
        
        Case TShowStyle.ssOld '老版本时不限制选择
        ''''''''''''''
    End Select
    
    If vsfQueueSetup.Cell(flexcpChecked, lngRowSel, lngColSel) = 1 Then
        If mlngStyleType = TShowStyle.ssMultiQueue Then mintSelectedQueueNum = mintSelectedQueueNum - 1
        
        vsfQueueSetup.Cell(flexcpChecked, lngRowSel, lngColSel) = 2
        vsfDoctorInfo.Editable = flexEDNone         '没有选择执行间时不允许设置医生值班信息
        
        If mlngStyleType = TShowStyle.ssSingleMan Or mlngStyleType = TShowStyle.ssSingleQueue Then vsfDoctorInfo.ColComboList(0) = ""
    Else
        If mlngStyleType = TShowStyle.ssMultiQueue Then
            If mintSelectedQueueNum >= mlngQueueListMaxRows Then
                MsgBox "您已选择了" & mintSelectedQueueNum & "个队列！" & vbCrLf & "最多允许选择" & mlngQueueListMaxRows & "个队列！", vbExclamation, gstrSysName
                Exit Sub
            End If
        
            mintSelectedQueueNum = mintSelectedQueueNum + 1
        End If
        
        vsfQueueSetup.Cell(flexcpChecked, lngRowSel, lngColSel) = 1
    End If
    
    If mlngStyleType = TShowStyle.ssMultiQueue Then
        vsfQueueSetup.ToolTipText = "您已选择了" & mintSelectedQueueNum & "个队列！" & vbCrLf & "最多允许选择" & mlngQueueListMaxRows & "个队列！"
    End If
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_SelChange()
On Error GoTo ErrorHand
    Call ShowImageAndIntroduction(vsfDoctorInfo.RowSel)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub ShowImageAndIntroduction(ByVal intIndex As Integer)
    Dim i As Integer
On Error GoTo ErrorHand
    For i = 0 To imgDoctorPhoto.Count - 1
        imgDoctorPhoto(i).Visible = False
        txtIntroduction(i).Visible = False
    Next
    
    imgDoctorPhoto(intIndex).Visible = True
    txtIntroduction(intIndex).Visible = True
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub


