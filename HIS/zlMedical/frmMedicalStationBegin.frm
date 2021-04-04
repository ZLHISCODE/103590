VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedicalStationBegin 
   Caption         =   "开始体检"
   ClientHeight    =   5880
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9750
   Icon            =   "frmMedicalStationBegin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9750
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5520
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationBegin.frx":076A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12118
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9750
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
         TabIndex        =   9
         Top             =   30
         Width           =   9630
         _ExtentX        =   16986
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
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.全选"
               Key             =   "全选"
               Object.ToolTipText     =   "全选(Alt+A)"
               Object.Tag             =   "&A.全选"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.全清"
               Key             =   "全清"
               Object.ToolTipText     =   "全清(Alt+C)"
               Object.Tag             =   "&C.全清"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&B.接受"
               Key             =   "接受"
               Object.ToolTipText     =   "接受(Alt+B)"
               Object.Tag             =   "&B.接受"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&B.报到"
               Key             =   "报到"
               Object.ToolTipText     =   "报到(Alt+B)"
               Object.Tag             =   "&B.报到"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8145
      Top             =   4740
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
            Picture         =   "frmMedicalStationBegin.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":1778
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":1EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":266C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":288C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7485
      Top             =   4740
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
            Picture         =   "frmMedicalStationBegin.frx":2AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":3226
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":39A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":411A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":433A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4875
      Left            =   6300
      TabIndex        =   10
      Top             =   630
      Width           =   3405
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   45
         Picture         =   "frmMedicalStationBegin.frx":455A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
         Width           =   285
      End
      Begin VB.CheckBox chk 
         Caption         =   "找到后处理为不报到(&5)"
         Height          =   240
         Index           =   1
         Left            =   855
         TabIndex        =   14
         Top             =   2250
         Width           =   2370
      End
      Begin VB.CheckBox chk 
         Caption         =   "找到后处理为报到(&4)"
         Height          =   240
         Index           =   0
         Left            =   855
         TabIndex        =   13
         Top             =   1935
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1305
         TabIndex        =   12
         Text            =   "cbo"
         Top             =   195
         Width           =   1995
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   930
         Width           =   1995
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   1125
         TabIndex        =   5
         Top             =   1365
         Width           =   1470
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&3.按姓名定位"
         Height          =   240
         Left            =   4710
         TabIndex        =   11
         Top             =   210
         Width           =   1425
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   1995
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.门 诊 号"
         Height          =   180
         Index           =   3
         Left            =   375
         TabIndex        =   3
         Tag             =   "门诊号"
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.组    别"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   1
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.工作单位"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   0
         Top             =   255
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   4740
      Left            =   60
      TabIndex        =   6
      Top             =   735
      Width           =   6210
      _cx             =   10954
      _cy             =   8361
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -30
         X2              =   1755
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "全清(&C)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "接受(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileCome 
         Caption         =   "报到(&B)"
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
Attribute VB_Name = "frmMedicalStationBegin"
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
Private mblnStarted As Boolean
Private mlngFindRow As Long
Private mintSort As Integer

Private Enum mCol
    报到 = 0
    姓名
    门诊号
    健康号
    就诊卡号
    身份证号
    性别
    工作单位
    组别
End Enum

Public WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1

'（２）自定义过程或函数************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long

    mnuFileSave.Enabled = True
        
    If vData = False Then
        mnuFileSave.Enabled = False
    End If

    If mblnStarted = False Then
        tbrThis.Buttons("接受").Enabled = mnuFileSave.Enabled
    Else
        tbrThis.Buttons("报到").Enabled = mnuFileSave.Enabled
    End If
    
End Property

Private Sub RefreshState()
    
    Dim lngLoop As Long
    Dim intCount As Integer
    
    intCount = 0
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, 0))) = 1 Then
            intCount = intCount + 1
        End If
    Next
    
    stbThis.Panels(2).Text = "当前选中 " & intCount & " 人"
End Sub

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next



    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng病人id As Long = 0, Optional blnStarted As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    mblnStarted = blnStarted
    mlngKey = lngKey
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadData(mlngKey) = False Then Exit Function
    
    '如果是单个人,直接处理,不弹出界面
    If lng病人id > 0 Then
        vsf.TextMatrix(1, mCol.报到) = 1
        
        If MsgBox("真的要接受体检吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveEdit() Then ShowEdit = True
        End If
        
        Exit Function
    End If
    
    EditChanged = (Val(vsf.RowData(1)) > 0)
    
    Call RefreshState
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand

    gstrSQL = "SELECT 0 AS 报到,A.病人id AS ID,B.姓名,B.性别,B.工作单位,B.门诊号,b.健康号,b.就诊卡号,b.身份证号,A.组别名称 AS 组别,'' AS 未到原因 " & _
                "FROM 体检人员档案 A,病人信息 B " & _
                "WHERE A.体检状态 IN (1,4) AND A.体检报到=0 AND A.病人id=B.病人id and A.登记id=[1]"
                
    gstrSQL = gstrSQL & " Order By 门诊号"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    gstrSQL = "SELECT Distinct B.工作单位 " & _
                "FROM 体检人员档案 A,病人信息 B " & _
                "WHERE B.工作单位 Is Not Null And A.体检状态 IN (1,4) AND A.体检报到 In ([2],[3]) AND A.病人id=B.病人id and A.登记id=[1]"
                    
    cbo(1).Clear
    cbo(1).AddItem ""
    cbo(1).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, 0, 0)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(1).AddItem rs("工作单位").Value
            rs.MoveNext
        Loop
    End If
    
    gstrSQL = "SELECT Distinct A.组别名称 " & _
                "FROM 体检人员档案 A " & _
                "WHERE A.组别名称 Is Not Null And A.体检状态 IN (1,4) AND A.体检报到=0 AND A.登记id=[1]"
                    
    cbo(0).Clear
    cbo(0).AddItem ""
    cbo(0).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(0).AddItem rs("组别名称").Value
            rs.MoveNext
        Loop
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
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mlngFindRow = 1
    mintSort = 0
    
    strVsf = "报到,450,1,1,1,;姓名,900,1,1,1,;门诊号,900,7,1,1,;健康号,900,7,1,1,;就诊卡号,900,1,1,1,;身份证号,1200,1,1,0,;性别,810,1,1,1,;工作单位,2280,1,1,1,;组别,1200,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
    vsf.Editable = True
    
    Call AppendRows(vsf, lnX, lnY)
    
    tbrThis.Buttons("接受").Visible = True
    tbrThis.Buttons("报到").Visible = True
    
    If mblnStarted = False Then
        Me.Caption = "接受体检"
        tbrThis.Buttons("报到").Visible = False
        mnuFileCome.Visible = False
    Else
        Me.Caption = "人员报到"
        
        tbrThis.Buttons("接受").Visible = False
        mnuFileSave.Visible = False
    End If
    
    cbo(0).Clear
    cbo(0).AddItem ""
    cbo(0).ListIndex = 0
    
    cbo(1).Clear
    cbo(1).AddItem ""
    cbo(1).ListIndex = 0
    
    
    
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


Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL  As String
    Dim lngCount As Long
    Dim lngDept As Long
    Dim lngSendNo As Long
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim str采集No As String
    Dim strNO As String
    Dim lngTotal As Long
    Dim strTmp As String
    Dim strSample As String
    Dim strDeptID As String
    Dim blnVerfiy As Boolean
    Dim blnCheck As Boolean
    Dim varTmp As Variant
    
    On Error GoTo errHand

    Me.Enabled = False

    Call frmWait.OpenWait(Me, IIf(tbrThis.Buttons("接受").Visible, "接受体检", "体检报到"))
    frmWait.WaitInfo = "正在进行报到处理..."

    '变量初始化处理
    
    
    '读取自动打印申请单相关参数
    strTmp = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\体检申请单条件", "条件1", "")
    
    If Trim(strTmp) <> "" And InStr(strTmp, "'") > 0 Then
        
        strTmp = Mid(strTmp, InStr(strTmp, "'") + 1)

        varTmp = Split(strTmp, "'")
        
        On Error Resume Next
        
        blnCheck = (Val(varTmp(0)) = 1)
        blnVerfiy = (Val(varTmp(1)) = 1)
        
        strSample = ""
        For lngLoop = 2 To UBound(varTmp) - 1
            strSample = strSample & "''" & varTmp(lngLoop)
        Next
        If strSample <> "" Then strSample = strSample & "''"
        
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\体检申请单条件", "打印执行科室", "")
    If strTmp <> "" Then
        strTmp = "'" & strTmp & "'"
        varTmp = Split(strTmp, "'")
        
        strDeptID = ""
        For lngLoop = 0 To UBound(varTmp)
            strDeptID = strDeptID & "," & Val(varTmp(lngLoop))
        Next
        If strDeptID <> "" Then strDeptID = Mid(strDeptID, 2)
        
    End If
    
    
    lngDept = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)

    
    lngTotal = vsf.Rows - 1
    lngSendNo = GetNextNo(10)
    
    frmWait.WaitInfo = "正在进行报到接受..."
    
    strSQL = "Select a.病人id,a.清单id,b.结算途径,b.采集方式id From 体检项目医嘱 a,体检项目清单 b Where a.清单id=b.id and b.登记id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
    If rs.BOF Then GoTo errHand
    
    For lngLoop = 1 To lngTotal
        
        If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, mCol.报到))) = 1 Then
            
            Call SQLRecord(rsSQL)

            frmWait.WaitInfo = "正在进行报到接受“" & vsf.TextMatrix(lngLoop, mCol.姓名) & " ”..." & Format(100 * lngLoop / lngTotal, "0.00") & "%"
            
            '产生费用单据号
            rs.Filter = ""
            rs.Filter = "病人id=" & Val(vsf.RowData(lngLoop))
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    
                    str采集No = ""
                    strNO = ""
                    
                    If zlCommFun.NVL(rs("结算途径").Value, 1) = 1 Then
                        '记帐
                        strNO = GetNextNo(14)
                    Else
                        strNO = GetNextNo(13)
                    End If
                    
                    If zlCommFun.NVL(rs("采集方式id").Value, 0) > 0 Then
                        '采集
                        If zlCommFun.NVL(rs("结算途径").Value, 1) = 1 Then
                            '记帐
                            str采集No = GetNextNo(14)
                        Else
                            str采集No = GetNextNo(13)
                        End If
                    End If
                    
                    strSQL = "ZL_体检项目医嘱_NO(" & zlCommFun.NVL(rs("清单id").Value, 0) & "," & Val(vsf.RowData(lngLoop)) & ",'" & strNO & "','" & str采集No & "')"
                    Call SQLRecordAdd(rsSQL, strSQL)
                    
                    rs.MoveNext
                Loop
            End If
            
            '开始执行
            blnTran = True
            gcnOracle.BeginTrans
            If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
            For lngCount = 1 To rsSQL.RecordCount
                Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
                rsSQL.MoveNext
            Next
            
            '接受或报到开始
            strSQL = "zl_体检人员档案_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(lngLoop)) & "," & lngDept & ",NULL,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '产生相关费用
            If MakeMedicalCharge(rsSQL, mlngKey) = False Then
                frmWait.CloseWait
                Me.Enabled = True
                gcnOracle.RollbackTrans
                blnTran = False
                Exit Function
            End If
            
            '接受或报到结束
            strSQL = "zl_体检人员档案_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(lngLoop)) & "," & lngDept & ",NULL,2)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            gcnOracle.CommitTrans
            blnTran = False
            
            If Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印指引单", 0)) = 1 Or Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印申请单", 0)) = 1 Then
                
                Call frmWait.HideWait
                
            
                '自动打印指引单
                If Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印指引单", 0)) = 1 Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861", Me, "登记id=" & mlngKey, "病人id=" & Val(vsf.RowData(lngLoop)), 2)
                End If
                
                '自动打印申请单
                If Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印申请单", 0)) = 1 Then
                    Call OutPutQuestBill(Me, mlngKey, Val(vsf.RowData(lngLoop)), strDeptID, strSample, blnVerfiy, blnCheck, 2)
                End If
                
                Call frmWait.ShowWait
            End If
            
        End If
    Next


    frmWait.CloseWait
    Me.Enabled = True
        
    SaveEdit = True

    Exit Function

errHand:

    frmWait.CloseWait
    Me.Enabled = True

    If ErrCenter = 1 Then
        Resume
    End If

    If blnTran Then
        gcnOracle.RollbackTrans
        ShowSimpleMsg "未完全报到成功或部份接受成功！"
    End If

End Function


Private Sub cbo_Click(Index As Integer)
    mlngFindRow = 0
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Function FindData() As Boolean
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnFind1 As Boolean
    Dim blnFind2 As Boolean
    Dim blnFind3 As Boolean
    Dim blnFind4 As Boolean
    Dim strCol As String
    
    FindData = True
    
    If mlngFindRow >= vsf.Rows - 1 Then mlngFindRow = 0

    strCol = Mid(lbl(3).Caption, 4)
    lngCol = GetCol(vsf, strCol)
    
    For lngRow = mlngFindRow + 1 To vsf.Rows - 1
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(1).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.工作单位)), UCase(cbo(1).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            
            Select Case strCol
            Case "门 诊 号"
                If UCase(vsf.TextMatrix(lngRow, mCol.门诊号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "健 康 号"
                If UCase(vsf.TextMatrix(lngRow, mCol.健康号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "就诊卡号"
                If UCase(vsf.TextMatrix(lngRow, mCol.就诊卡号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "身份证号"
                If UCase(vsf.TextMatrix(lngRow, mCol.身份证号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "姓    名"
                If UCase(vsf.TextMatrix(lngRow, mCol.姓名)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "姓名拼音"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.姓名))) = UCase(txt(1).Text) Then blnFind2 = True
            Case "姓名五笔"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.姓名)), 1) = UCase(txt(1).Text) Then blnFind2 = True
            End Select

        End If
        
        
        If cbo(0).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.组别)), UCase(cbo(0).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
            mlngFindRow = lngRow
            
            vsf.Row = mlngFindRow
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
    Next
    
    For lngRow = 1 To mlngFindRow
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(1).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.工作单位)), UCase(cbo(1).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.门诊号)), UCase(txt(1).Text)) > 0 Then
                blnFind2 = True
            End If
            
            Select Case strCol
            Case "门 诊 号"
                If UCase(vsf.TextMatrix(lngRow, mCol.门诊号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "健 康 号"
                If UCase(vsf.TextMatrix(lngRow, mCol.健康号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "就诊卡号"
                If UCase(vsf.TextMatrix(lngRow, mCol.就诊卡号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "身份证号"
                If UCase(vsf.TextMatrix(lngRow, mCol.身份证号)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "姓    名"
                If UCase(vsf.TextMatrix(lngRow, mCol.姓名)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "姓名拼音"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.姓名))) = UCase(txt(1).Text) Then blnFind2 = True
            Case "姓名五笔"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.姓名)), 1) = UCase(txt(1).Text) Then blnFind2 = True
            End Select
            
        End If
        
        If cbo(0).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.组别)), UCase(cbo(0).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
            mlngFindRow = lngRow
            
            vsf.Row = mlngFindRow
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
        

    Next
    FindData = False
    
End Function

Private Sub chk_Click(Index As Integer)
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 300 * 3)
    
    txt(1).Text = ""
    LocationObj txt(1)
End Sub

Private Sub cmdSelect_Click()
    
    If FindData Then
    
        If chk(0).Value = 1 Then
            vsf.TextMatrix(vsf.Row, mCol.报到) = 1
            EditChanged = True
            Call RefreshState
        End If
        
        If chk(1).Value = 1 Then
            vsf.TextMatrix(vsf.Row, mCol.报到) = 0
            EditChanged = True
            Call RefreshState
        End If
    End If
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("全选").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全选"))
        Case vbKeyC
            If tbrThis.Buttons("全清").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全清"))
        Case vbKeyB
            If tbrThis.Buttons("开始").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("开始"))
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

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Load()

    glngFormW = 9870
    glngFormH = 6570
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    
    lbl(3).Caption = "&3." & (GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", "门 诊 号"))
    lbl(3).Tag = Mid(lbl(3).Caption, 4)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With vsf
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left - fra.Width - 15
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra
        .Left = vsf.Left + vsf.Width + 15
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = vsf.Height + 90
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", lbl(3).Tag)
End Sub


Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.报到) = 0
        End If
    Next
    
    EditChanged = False
End Sub

Private Sub mnuFileCome_Click()
    Call mnuFileSave_Click
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileSave_Click()
    
    If ValidEdit = False Then Exit Sub

    If SaveEdit() Then
        EditChanged = False
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.报到) = 1
            EditChanged = True
        End If
    Next
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

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    mobjPopMenu.Add 1, "&1.门 诊 号", , , True, , (lbl(3).Tag = "门 诊 号")
    mobjPopMenu.Add 2, "&2.健 康 号", , , True, , (lbl(3).Tag = "健 康 号")
    mobjPopMenu.Add 3, "&3.就诊卡号", , , True, , (lbl(3).Tag = "就诊卡号")
    mobjPopMenu.Add 4, "&4.姓    名", , , True, , (lbl(3).Tag = "姓    名")
    mobjPopMenu.Add 5, "&5.姓名拼音", , , True, , (lbl(3).Tag = "姓名拼音")
    mobjPopMenu.Add 6, "&6.姓名五笔", , , True, , (lbl(3).Tag = "姓名五笔")
    mobjPopMenu.Add 7, "&7.身份证号", , , True, , (lbl(3).Tag = "身份证号")
    
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    
    Caption = Mid(Caption, 4)
    
    lbl(3).Caption = "&3." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
    lbl(3).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "全选"
        Call mnuFileSelectAll_Click
    Case "全清"
        Call mnuFileClearAll_Click
    Case "报到"
        Call mnuFileSave_Click
    Case "接受"
        Call mnuFileSave_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_Change(Index As Integer)
    mlngFindRow = 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    cmdSelect.Default = True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    Dim blnCard As Boolean
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    strCol = Mid(lbl(3).Caption, 4)
    lngCol = GetCol(vsf, strCol)
            
    If strCol = "就诊卡号" And KeyAscii <> vbKeyReturn Then
        '就诊卡号，自动识别

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.就诊卡号码长度 - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = 0
            Call cmdSelect_Click
        End If

    End If
    
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        zlCommFun.OpenIme False
    End Select
    
    cmdSelect.Default = False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.报到))) = 1 Then
        EditChanged = True
        Call RefreshState
        Exit Sub
    End If
        
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.报到))) = 1 Then
            EditChanged = True
            Call RefreshState
            Exit Sub
        End If
    Next
    
    If lngLoop = vsf.Rows Then EditChanged = False
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace And vsf.Col <> mCol.报到 Then
    
        If Abs(Val(vsf.TextMatrix(vsf.Row, mCol.报到))) = 1 Then
            vsf.TextMatrix(vsf.Row, mCol.报到) = 0
        Else
            vsf.TextMatrix(vsf.Row, mCol.报到) = 1
        End If
        
        EditChanged = True
        
        Call RefreshState
            
    End If
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If vsf.MouseRow = 0 Then
        If mintSort = flexSortGenericAscending Then
            mintSort = flexSortGenericDescending
        Else
            mintSort = flexSortGenericAscending
        End If
        
        vsf.Sort = mintSort
    End If
    
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.报到 Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

