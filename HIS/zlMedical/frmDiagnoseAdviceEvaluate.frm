VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDiagnoseAdviceEvaluate 
   Caption         =   "体检诊断评估"
   ClientHeight    =   6270
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   10500
   Icon            =   "frmDiagnoseAdviceEvaluate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10500
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   5040
      Left            =   135
      ScaleHeight     =   5040
      ScaleWidth      =   9420
      TabIndex        =   17
      Top             =   795
      Width           =   9420
      Begin VB.CheckBox chk 
         Caption         =   "评估分组(&G)"
         Height          =   225
         Left            =   5835
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1305
      End
      Begin VB.Frame fra 
         Height          =   4560
         Left            =   5850
         TabIndex        =   18
         Top             =   405
         Width           =   3450
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   4
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   570
            Width           =   2190
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   4
            Left            =   855
            TabIndex        =   24
            Top             =   915
            Width           =   480
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   5
            Left            =   1575
            TabIndex        =   23
            Top             =   915
            Width           =   480
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   3
            Left            =   2115
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   915
            Width           =   915
         End
         Begin VB.CommandButton cmd 
            Caption         =   "移出规则(&R) >>"
            Height          =   350
            Index           =   1
            Left            =   75
            TabIndex        =   13
            Top             =   3090
            Width           =   1440
         End
         Begin VB.CommandButton cmdOpen 
            Height          =   300
            Left            =   3060
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":076A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1260
            Width           =   300
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            Left            =   855
            TabIndex        =   11
            Top             =   2340
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<< 添加规则(&A)"
            Height          =   350
            Index           =   0
            Left            =   75
            TabIndex        =   12
            Top             =   2685
            Width           =   1440
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   855
            TabIndex        =   5
            Top             =   1260
            Width           =   2190
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1620
            Width           =   2190
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   855
            TabIndex        =   10
            Top             =   1980
            Width           =   2190
         End
         Begin VB.TextBox txt 
            BackColor       =   &H8000000A&
            Height          =   300
            Index           =   1
            Left            =   855
            TabIndex        =   3
            Top             =   225
            Width           =   2190
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&2.性  别"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   75
            TabIndex        =   28
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&3.年  龄"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   75
            TabIndex        =   27
            Top             =   975
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   1380
            TabIndex        =   26
            Top             =   960
            Width           =   180
         End
         Begin VB.Label lbl 
            Caption         =   "评估分组时，只有任意一组条件成立都认为此评估成立。"
            Height          =   435
            Index           =   6
            Left            =   540
            TabIndex        =   21
            Top             =   3615
            Width           =   2745
            WordWrap        =   -1  'True
         End
         Begin VB.Image img 
            Height          =   240
            Left            =   180
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":0CF4
            Top             =   3585
            Width           =   240
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "单位:s"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   5
            Left            =   1590
            TabIndex        =   20
            Top             =   3165
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "类型:数值型"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   4
            Left            =   1575
            TabIndex        =   19
            Top             =   2805
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&4.项  目"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   4
            Top             =   1335
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&5.条  件"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   7
            Top             =   1695
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&6.项目值"
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   9
            Top             =   2010
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&1.组  名"
            ForeColor       =   &H8000000C&
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   2
            Top             =   300
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   4185
         Left            =   30
         TabIndex        =   0
         Top             =   135
         Width           =   5790
         _cx             =   10213
         _cy             =   7382
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         Begin VB.Line lnX 
            Index           =   1
            Visible         =   0   'False
            X1              =   1140
            X2              =   2925
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line lnY 
            Index           =   1
            Visible         =   0   'False
            X1              =   1965
            X2              =   1965
            Y1              =   1590
            Y2              =   2805
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   5910
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":127E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13441
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8115
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":1B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":1D2C
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":1F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2166
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2386
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7515
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":25A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":27C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":29DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2D2C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2F4C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10500
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   1270
         ButtonWidth     =   1402
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存(Alt+S)"
               Object.Tag             =   "&S.保存"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.重填"
               Key             =   "重填"
               Object.ToolTipText     =   "重填(Alt+R)"
               Object.Tag             =   "&R.重填"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "重填(&R)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmDiagnoseAdviceEvaluate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private Type Items
    项目 As String
End Type

Private usrSaveItem As Items

'（２）自定义过程或函数************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    mnuFileSave.Enabled = True
    mnuFileRestore.Enabled = True

    If vData = False Then
        mnuFileSave.Enabled = False
        mnuFileRestore.Enabled = False

    End If

    tbrThis.Buttons("保存").Enabled = mnuFileSave.Enabled
    tbrThis.Buttons("重填").Enabled = mnuFileRestore.Enabled
        
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next

    Call ResetVsf(vsf)
    Call AppendRows(vsf, lnX, lnY)
    
    On Error GoTo 0
    
    Call InitData
    
    EditChanged = True
    
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, _
                            ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant
    Dim lngLoop As Long
    
    mblnStartUp = True
    mblnOK = False
                    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    If ReadData(mlngKey) = False Then Exit Function
    
    Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    
    EditChanged = False
    
    stbThis.Panels(2).Text = "填写体检诊断的评估规则。"
                
    Me.Show 1, frmMain
        
    ShowEdit = mblnOK
    
End Function

Private Function InitSysFlag() As Boolean

    cbo(1).Clear
    cbo(2).Clear
    
    Select Case Mid(lbl(4).Caption, 4)
    Case "数字型"
        cbo(1).AddItem "[最低值]"
        cbo(1).AddItem "[最高值]"
        cbo(2).AddItem "[最低值]"
        cbo(2).AddItem "[最高值]"
    End Select

    Select Case cbo(0).Text
    Case "等于"
        cbo(1).AddItem "[偏高]"
        cbo(1).AddItem "[偏低]"
        cbo(1).AddItem "[异常]"
        cbo(2).AddItem "[偏高]"
        cbo(2).AddItem "[偏低]"
        cbo(2).AddItem "[异常]"
    End Select
        
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT B.ID,A.分组名 AS 组名,B.中文名 AS 项目,A.关系式 AS 条件,A.条件值 AS 项目值,A.性别,A.开始年龄,A.结束年龄 from 体检诊断评估 A,诊治所见项目 B WHERE A.项目id=B.ID AND A.诊断序号=[1] ORDER BY 分组名"
           
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        
        If zlCommFun.NVL(rs("组名").Value) <> "" Then chk.Value = 1
                
        Call LoadGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
        
    On Error GoTo errHand
    
    Dim strVsf As String
   
    strVsf = "组名,900,1,1,1,;性别,600,1,1,1,;开始年龄,900,1,1,1,;结束年龄,900,1,1,1,;项目,2100,1,1,1,;条件,900,1,1,1,;项目值,900,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)

    lbl(4).Caption = ""
    lbl(5).Caption = ""
'    vsf.ColHidden(0) = True
    vsf.MergeCol(0) = True
    
    With cbo(4)
        .Clear
        .AddItem ""
        .AddItem "1-男"
        .AddItem "2-女"
        .ListIndex = 0
    End With
    
    With cbo(3)
        .Clear
        .AddItem "1-岁"
        .AddItem "2-月"
        .AddItem "3-天"
        .ListIndex = 0
    End With
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

                                                                
    ValidEdit = True
    
End Function

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_体检诊断评估_DELETE(" & mlngKey & ")"
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            strSQL(ReDimArray(strSQL)) = "ZL_体检诊断评估_INSERT(" & mlngKey & ",'" & _
                                        vsf.TextMatrix(lngLoop, 0) & "'," & _
                                        Val(vsf.RowData(lngLoop)) & ",'" & _
                                        vsf.TextMatrix(lngLoop, 5) & "','" & _
                                        vsf.TextMatrix(lngLoop, 6) & "','" & _
                                        vsf.TextMatrix(lngLoop, 1) & "','" & _
                                        vsf.TextMatrix(lngLoop, 2) & "','" & _
                                        vsf.TextMatrix(lngLoop, 3) & "')"
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Sub FillOperate(ByVal bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim strText As String
    
    strText = cbo(0).Text
    
    cbo(0).Clear
    cbo(1).Clear
    Select Case bytMode
    Case 0  '数字型
        cbo(0).AddItem "等于"
        cbo(0).AddItem "大于"
        cbo(0).AddItem "小于"
        cbo(0).AddItem "大于等于"
        cbo(0).AddItem "小于等于"
        cbo(0).AddItem "不等于"
        cbo(0).AddItem "在范围内"
    Case 1, 2 '文字型
        cbo(0).AddItem "等于"
        cbo(0).AddItem "大于"
        cbo(0).AddItem "小于"
        cbo(0).AddItem "大于等于"
        cbo(0).AddItem "小于等于"
        cbo(0).AddItem "不等于"
        cbo(0).AddItem "包含"
    Case 3  '阴阳型(罗辑型)
        cbo(0).AddItem "等于"
        cbo(0).AddItem "不等于"
        cbo(0).AddItem "包含"
'
'        cbo(1).AddItem "阴性"
'        cbo(1).AddItem "阳性"
'        cbo(1).ListIndex = 0
    End Select
    
    On Error Resume Next
    
    cbo(0).Text = strText
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
End Sub

Private Sub cbo_Click(Index As Integer)
    Select Case Index
    Case 0
        cbo(2).Visible = (cbo(Index).List(cbo(Index).ListIndex) = "在范围内")
        
        Call InitSysFlag
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Index > 0 Then
            
            If Chr(KeyAscii) = "'" Then KeyAscii = 0
            
            Select Case Val(cmdOpen.Tag)
            Case 0
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.")
            Case 1
                
            Case 2
                
            End Select
        End If
    End If
    
End Sub

Private Sub chk_Click()
    
    If chk.Value = 1 Then
        txt(1).Enabled = True
        txt(1).BackColor = &H80000005
        lbl(3).ForeColor = &H80000012
        
        ResetVsf vsf
        Call AppendRows(vsf, lnX, lnY)
'        vsf.ColHidden(0) = False
    Else
        txt(1).Enabled = False
        txt(1).BackColor = &H8000000A
        lbl(3).ForeColor = &H8000000C
        
        ResetVsf vsf
        Call AppendRows(vsf, lnX, lnY)
'        vsf.ColHidden(0) = True
    End If
    
    EditChanged = True
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Dim intRow As Long
    
    Select Case Index
    Case 0
        If Val(vsf.RowData(vsf.Rows - 1)) = 0 Then
            intRow = vsf.Rows - 1
        Else
            vsf.Rows = vsf.Rows + 1
            intRow = vsf.Rows - 1
        End If
        
        If chk.Value = 1 Then
            If Trim(txt(1).Text) = "" Then
                ShowSimpleMsg "必须输入评估的组名！"
                Exit Sub
            End If
        End If
        
        If Val(fra.Tag) = 0 Then Exit Sub
        
        vsf.RowData(intRow) = fra.Tag
        vsf.TextMatrix(intRow, 0) = txt(1).Text
        
        vsf.TextMatrix(intRow, 1) = zlCommFun.GetNeedName(cbo(4).Text)
        
        If Trim(txt(4).Text) <> "" Then
            vsf.TextMatrix(intRow, 2) = Trim(txt(4).Text) & zlCommFun.GetNeedName(cbo(3).Text)
        Else
            vsf.TextMatrix(intRow, 2) = ""
        End If
        
        If Trim(txt(5).Text) <> "" Then
            vsf.TextMatrix(intRow, 3) = Trim(txt(5).Text) & zlCommFun.GetNeedName(cbo(3).Text)
        Else
            vsf.TextMatrix(intRow, 3) = ""
        End If
        
        vsf.TextMatrix(intRow, 4) = txt(0).Text
        vsf.TextMatrix(intRow, 5) = cbo(0).Text
        
        Select Case cbo(1).Text
        Case "[最低值]", "[最高值]", "[偏高]", "[偏低]", "[异常]"
            
        Case Else
            If Val(cmdOpen.Tag) = 0 Then
                cbo(1).Text = Val(cbo(1).Text)
            End If
        End Select
        
        Select Case cbo(2).Text
        Case "[最低值]", "[最高值]", "[偏高]", "[偏低]", "[异常]"
            
        Case Else
            If Val(cmdOpen.Tag) = 0 Then
                cbo(2).Text = Val(cbo(2).Text)
            End If
        End Select
        
        If cbo(2).Visible Then
            vsf.TextMatrix(intRow, 6) = cbo(1).Text & " 至 " & cbo(2).Text
        Else
            vsf.TextMatrix(intRow, 6) = cbo(1).Text
        End If
        
        
        vsf.Col = 0
        vsf.Sort = flexSortGenericAscending
        
        EditChanged = True
        
        Call AppendRows(vsf, lnX, lnY)
        
        LocationObj txt(0)
        
    Case 1
        
        If vsf.Rows <> 2 Then
            vsf.RemoveItem vsf.Row
        Else
            Call ResetVsf(vsf)
        End If
        Call AppendRows(vsf, lnX, lnY)
        
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        
        EditChanged = True
        
    End Select
End Sub

Private Sub cmdOpen_Click()
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
        
    gstrSQL = GetPublicSQL(SQL.诊治项目选择)
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt(0), "编码,1200,0,1;名称,1800,0,0;临床意义,1800,0,0", Me.Name & "\诊治项目选择", "请选择一个诊治项目。", rsData, rs, 8790, 5100) Then
        
        txt(0).Text = zlCommFun.NVL(rs("名称").Value)
        fra.Tag = zlCommFun.NVL(rs("ID").Value)
        cmdOpen.Tag = zlCommFun.NVL(rs("类型").Value, 0)
        txt(0).Tag = ""
        
        Select Case Val(cmdOpen.Tag)
        Case 0
            lbl(4).Caption = "类型:数字型"
        Case 1
            lbl(4).Caption = "类型:文字型"
        Case 2
            lbl(4).Caption = "类型:罗辑型"
        Case Else
            lbl(4).Caption = "类型:"
        End Select
        
        lbl(5).Caption = "单位:" & zlCommFun.NVL(rs("单位").Value)
        
        usrSaveItem.项目 = txt(0).Text
                                
        Call FillOperate(Val(cmdOpen.Tag))
        
        Call InitSysFlag
        
        cbo(1).Text = ""
        cbo(2).Text = ""
        
    End If

    txt(0).SetFocus
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyS
            If tbrThis.Buttons("保存").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("保存"))
        Case vbKeyR
            If tbrThis.Buttons("重填").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("重填"))
        Case vbKeyH
            If tbrThis.Buttons("帮助").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("帮助"))
        Case vbKeyX
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End If
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With picBack
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        
    End With
    
    With vsf
        .Left = 0
        .Top = 0
        .Width = picBack.Width - .Left - fra.Width - 30
        .Height = picBack.Height - .Top
    End With
    
    With chk
        .Left = vsf.Left + vsf.Width + 30
        .Top = 30
    End With
    
    With fra
        .Left = chk.Left
        .Top = chk.Top + chk.Height + 30 - 90
        .Height = picBack.Height - .Top
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
        
    If MsgBox("确实要恢复以前所选项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    Call ReadData(mlngKey)
    
    Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    Dim lngKey As Long
            
    If ValidEdit = False Then Exit Sub
    If SaveEdit(lngKey) Then
        mblnOK = True
        EditChanged = False
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub


Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "保存"
        Call mnuFileSave_Click
    Case "重填"
        Call mnuFileRestore_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    
    If Index = 0 Then
        txt(Index).Tag = "Changed"
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)

    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt(Index)
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsData As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        
        If Index = 0 And txt(Index).Tag <> "" Then
            
            strText = UCase(txt(Index).Text) & "%"
            
            gstrSQL = GetPublicSQL(SQL.诊治项目过滤选择)
            
            If ParamInfo.项目输入匹配方式 = 0 Then strTmp = "%" & strText
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
            
            If ShowTxtFilter(Me, txt(Index), "编码,900,0,1;名称,2400,0,0;英文名,1200,0,0;临床意义,900,0,0", Me.Name & "\诊治项目过滤选择", "请从下表中选择一个项目", rsData, rs) Then
                
                txt(0).Text = zlCommFun.NVL(rs("名称").Value)
                fra.Tag = zlCommFun.NVL(rs("ID").Value)
                cmdOpen.Tag = zlCommFun.NVL(rs("类型").Value, 0)
                txt(0).Tag = ""
                usrSaveItem.项目 = txt(0).Text
                
                Call FillOperate(Val(cmdOpen.Tag))
                
                cbo(1).Text = ""
                cbo(2).Text = ""
                
                Select Case Val(cmdOpen.Tag)
                Case 0
                    lbl(4).Caption = "类型:数字型"
                Case 1
                    lbl(4).Caption = "类型:文字型"
                Case 2
                    lbl(4).Caption = "类型:罗辑型"
                Case Else
                    lbl(4).Caption = "类型:"
                End Select
                lbl(5).Caption = "单位:" & zlCommFun.NVL(rs("单位").Value)
                
                Call InitSysFlag
                
            Else
                txt(0).Text = usrSaveItem.项目
                Exit Sub
            End If
        End If
                                
        zlCommFun.PressKey vbKeyTab
        If Index = 0 Then zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
        Case 0
            
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            
            If Chr(KeyAscii) = "*" Then
                KeyAscii = 0
                Call cmdOpen_Click
            End If
            
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If (txt(Index).Tag = "Changed") And Index = 0 Then
        txt(Index).Text = usrSaveItem.项目
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    
    Call SelectRow(vsf, OldRow, NewRow)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.焦点
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.非焦点
    Call SelectRow(vsf, 1, vsf.Row)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

