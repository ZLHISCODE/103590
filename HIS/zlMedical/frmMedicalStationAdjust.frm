VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMedicalStationAdjust 
   Caption         =   "检查结果批量调整"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "frmMedicalStationAdjust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10350
   Begin VB.PictureBox picResult 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   3435
      ScaleHeight     =   435
      ScaleWidth      =   3435
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4710
      Width           =   3435
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1185
         TabIndex        =   13
         Top             =   90
         Width           =   2880
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.调整为新值"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   12
         Top             =   150
         Width           =   1080
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10350
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   10230
         _ExtentX        =   18045
         _ExtentY        =   1138
         ButtonWidth     =   1296
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
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
               Caption         =   "&D.调整"
               Key             =   "调整"
               Object.ToolTipText     =   "调整(Alt+D)"
               Object.Tag             =   "&D.调整"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Left            =   4470
      Top             =   750
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
            Picture         =   "frmMedicalStationAdjust.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":28E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":2B00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   5460
      Top             =   750
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
            Picture         =   "frmMedicalStationAdjust.frx":2D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":349A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":3C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":438E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":45AE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   5865
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationAdjust.frx":47CE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13203
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
   Begin VB.Frame fra 
      Height          =   4770
      Left            =   135
      TabIndex        =   19
      Top             =   840
      Width           =   3255
      Begin VB.CommandButton cmdCalc 
         Caption         =   "按新值填写(&J)"
         Height          =   350
         Left            =   255
         TabIndex        =   15
         Top             =   4335
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&P"
         Height          =   300
         Left            =   2835
         TabIndex        =   10
         Top             =   2325
         Width           =   300
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   2880
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "立即搜索(&S)"
         Height          =   350
         Left            =   270
         TabIndex        =   11
         Top             =   2715
         Width           =   1185
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   2325
         Width           =   2535
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   1095
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   435
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   100990979
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1860
         TabIndex        =   3
         Top             =   435
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   100990979
         CurrentDate     =   38229
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.体检科室"
         Height          =   180
         Index           =   6
         Left            =   90
         TabIndex        =   6
         Top             =   1470
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   4
         Left            =   1620
         TabIndex        =   2
         Top             =   495
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.体检项目"
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   8
         Top             =   2100
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.体检单号"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   855
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.体检时间"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1980
      Left            =   3795
      TabIndex        =   14
      Top             =   1560
      Width           =   3795
      _cx             =   6694
      _cy             =   3492
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8475
      Top             =   2250
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
            Picture         =   "frmMedicalStationAdjust.frx":5062
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":53FC
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationAdjust.frx":5796
            Key             =   "单据"
         EndProperty
      EndProperty
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
         Caption         =   "调整(&D)"
         Shortcut        =   ^D
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
Attribute VB_Name = "frmMedicalStationAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mstrSQL As String
Private mblnChangeEdit As Boolean
Private mstrPrivs As String


Private mlngKey As Long

Private Enum mCol
    选择 = 0
    门诊号
    姓名
    上次结果
    调整结果
    团体
    体检时间
    病人id
    病历id
'    状态
End Enum

Private Type Items
    诊治项目 As String
End Type

Private usrSaveGroup As Items

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '功能:根据修改状态设置按钮、菜单等的可用状态
    '-----------------------------------------------------------------------------------------
    
    mnuFileSave.Enabled = True
        
    If mblnChangeEdit = False Then mnuFileSave.Enabled = False
        
    tbrThis.Buttons("调整").Enabled = mnuFileSave.Enabled
        
End Sub

Private Sub RefreshStatus()
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    If vsf.Rows = 2 And Trim(vsf.TextMatrix(1, 1)) = "" Then
        stbThis.Panels(2).Text = "没有信息。"
    Else
        stbThis.Panels(2).Text = "共找到 " & vsf.Rows - 1 & " 个信息。"
    End If
    
End Sub

Public Function ShowEdit(ByVal frmMain As Form, Optional ByVal strPrivs As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示本编辑窗体
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
                    
    mblnChangeEdit = False
    Call AdjustEnableState
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strVsf As String
    
    On Error GoTo errHand
    
    strVsf = "选择,540,4,1,1,;门诊号,810,7,1,1,;姓名,900,1,1,1,;上次结果,1200,1,1,1,;调整结果,1200,1,1,1,;团体,1500,1,1,1,;体检时间,1670,1,1,1,;病人id,0,1,1,0,;病历id,0,1,1,0,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(mCol.选择) = flexDTBoolean
    'Set vsf.Cell(flexcpPicture, 0, mCol.状态) = ils13.ListImages("状态").Picture
    vsf.Editable = flexEDKbdMouse
    
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strWhere As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Call ResetVsf(vsf)
    Call AppendRows(vsf, lnX, lnY)
    
    strWhere = " AND b.体检时间 BETWEEN TO_DATE('" & Format(dtp(0).Value, dtp(0).CustomFormat) & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(dtp(1).Value, dtp(1).CustomFormat) & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss') "
    If cbo(1).ListIndex > 0 Then strWhere = strWhere & " AND a.体检部门id+0=" & cbo(1).ItemData(cbo(1).ListIndex)
    If Trim(txt(0).Text) <> "" Then
        
        varTmp2 = Split(Trim(txt(0).Text), ",")
        strTmp = " 1=2 "
        For lngLoop = 0 To UBound(varTmp2)
            If InStr(varTmp2(lngLoop), "-") = 0 Then
                strTmp = strTmp & "  OR a.体检号='" & varTmp2(lngLoop) & "'"
            Else
                strTmp = strTmp & "  OR a.体检号 BETWEEN '" & Mid(varTmp2(lngLoop), 1, InStr(varTmp2(lngLoop), "-") - 1) & "' AND '" & Mid(varTmp2(lngLoop), InStr(varTmp2(lngLoop), "-") + 1) & "'"
            End If
        Next
        If strTmp <> " 1=2 " Then strWhere = strWhere & " AND (" & strTmp & ")"
        
    End If
    
    mstrSQL = "" & _
     "Select b.id,0 as 选择,i.门诊号,i.姓名,e.挂号单,e.病人id,h.病历id,h.所见内容 As 上次结果,h.所见内容 As 调整结果,j.名称 As 团体,b.体检时间,'报告' As 状态 " & _
        "From  体检登记记录 a, " & _
              "体检人员档案 b, " & _
              "体检项目清单 c, " & _
              "体检项目医嘱 d, " & _
              "病人医嘱记录 e, " & _
              "病人医嘱发送 f, " & _
              "病人病历内容 g, " & _
              "病人病历所见单 h, " & _
              "病人信息 i, " & _
              "合约单位 j " & _
        "where a.ID=b.登记id " & _
              "and a.id=c.登记id " & _
              "and d.清单id=c.id " & _
              "and e.id=d.医嘱id " & _
              "and f.医嘱id=e.id " & _
              "and g.病历记录id=f.报告id " & _
              "and g.id=h.病历id " & _
              "and i.病人id=b.病人id " & _
              "and i.病人id=e.病人id " & _
              "and d.病人id=b.病人id " & _
              "and a.合约单位id=j.id(+) " & _
              "and h.所见项ID+0=" & mlngKey & strWhere
    
    mstrSQL = "Select * From (" & mstrSQL & ") Order By 门诊号"
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If rs.BOF = False Then
        Call LoadGrid(vsf, rs, Array("", "", "", "", "", "", "yyyy-MM-dd"), , ils13)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    If InStr(mstrPrivs, "未收费体检") = 0 Then
        For mlngLoop = vsf.Rows - 1 To 1 Step -1
            If Val(vsf.RowData(mlngLoop)) > 0 Then
                
                gstrSQL = GetPublicSQL(SQL.个人费用概况)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(mlngLoop)))
                If CalcCharge(rsData, rs) Then
                    If rs("未收金额").Value > 0 Then vsf.RemoveItem mlngLoop
                End If
                
            End If
        Next
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strError As String
    Dim rs As New ADODB.Recordset
    
    '检验输入的新值是否正确,主要是检验公式
    
    On Error GoTo errHand
    
    
            
    ValidData = True
    
    Exit Function
errHand:
    LocationObj txt(1)
    strError = "调整结果值或公式不合法！"
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strSQL() As String
        
    On Error GoTo errHand
    ReDim strSQL(1 To 1)

    For mlngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.选择))) = 1 And Val(vsf.RowData(mlngLoop)) > 0 Then

            strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_单项填写(" & Val(vsf.TextMatrix(mlngLoop, mCol.病历id)) & "," & mlngKey & ",'" & vsf.TextMatrix(mlngLoop, mCol.调整结果) & "')"
            
        End If
    Next
    
    blnTran = True
    
    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    SaveData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Function AdjustResult() As Boolean
    Dim lngLoop As Long
    
    For mlngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.选择))) = 1 And Val(vsf.RowData(mlngLoop)) > 0 Then
            
            vsf.TextMatrix(mlngLoop, mCol.调整结果) = txt(1).Text

        End If
    Next
    
    AdjustResult = True
    
End Function

Private Sub cmdCalc_Click()
    Call AdjustResult
End Sub

Private Sub cmdOpen_Click()
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    gstrSQL = GetPublicSQL(SQL.检查诊治项目选择)
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt(2), "编码,1200,0,1;名称,1800,0,0;临床意义,1800,0,0", Me.Name & "\诊治项目选择", "请选择一个诊治项目。", rsData, rs, 8790, 5100) Then
        
        txt(2).Text = zlCommFun.NVL(rs("名称").Value)
        mlngKey = zlCommFun.NVL(rs("ID").Value)
        cmdOpen.Tag = zlCommFun.NVL(rs("类型").Value, 0)
        txt(2).Tag = ""
        
        usrSaveGroup.诊治项目 = txt(2).Text
                                
        txt(1).Text = ""
        
    End If
    
    txt(2).SetFocus
End Sub

Private Sub cmdRefresh_Click()
    
    Call ReadData
            
    mblnChangeEdit = False
    Call AdjustEnableState
    Call RefreshStatus
    
    vsf.Col = 1
    vsf.SetFocus
    vsf.Col = 0
End Sub


Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    '检验部门
    cbo(1).Clear
    For mlngLoop = 0 To mfrmMain.cboDept.ListCount - 1
        cbo(1).AddItem mfrmMain.cboDept.List(mlngLoop)
        cbo(1).ItemData(cbo(1).NewIndex) = mfrmMain.cboDept.ItemData(mlngLoop)
    Next
    cbo(1).ListIndex = mfrmMain.cboDept.ListIndex
            
    dtp(0).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtp(1).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
'    cbo(0).ListIndex = 0
    
    If cbo(1).ListIndex = -1 Then
        zlControl.CboLocate cbo(1), UserInfo.部门ID, True
        If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    End If
    
    txt(0).Text = ""
    txt(2).Text = ""
    dtp(0).SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("全选").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全选"))
        Case vbKeyC
            If tbrThis.Buttons("全清").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全清"))
        Case vbKeyB
            If tbrThis.Buttons("调整").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("调整"))
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
    
    With fra
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With vsf
        .Left = fra.Left + fra.Width
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - picResult.Height
    End With
    
    With picResult
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height
        .Width = vsf.Width
    End With
    
    txt(1).Width = picResult.Width - txt(1).Left - 60
   
    Call AppendRows(vsf, lnX, lnY)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChangeEdit Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuFileClearAll_Click()
    For mlngLoop = 1 To vsf.Rows - 1
        vsf.TextMatrix(mlngLoop, mCol.选择) = 0
    Next
    mblnChangeEdit = False
    Call AdjustEnableState
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    If mblnChangeEdit Then
        
        If MsgBox("真的要将选中的调整为新值吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        If AdjustResult = False Then Exit Sub
        If SaveData() = False Then Exit Sub
        
        mblnOK = True
        
        mblnChangeEdit = False
        Call AdjustEnableState

        ShowSimpleMsg "体检结果批量调整成功！"
        
        Call ResetVsf(vsf)
        txt(1).Text = ""
        
        Exit Sub
    End If
End Sub

Private Sub mnuFileSelectAll_Click()
    For mlngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) > 0 Then
            vsf.TextMatrix(mlngLoop, mCol.选择) = 1
        End If
    Next
    
    mblnChangeEdit = True
    Call AdjustEnableState
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
    Case "全选"
        Call mnuFileSelectAll_Click
    Case "全清"
        Call mnuFileClearAll_Click
    Case "调整"
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
    If Index = 2 Then txt(2).Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    If KeyAscii = vbKeyReturn Then
        If Index = 2 Then
            If txt(2).Tag <> "" Then
                txt(2).Tag = ""
                
                strText = UCase(txt(Index).Text) & "%"
                If ParamInfo.项目输入匹配方式 = 0 Then strTmp = " %" & strText
                
                gstrSQL = GetPublicSQL(SQL.检查诊治项目过滤选择)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
                
                If ShowTxtFilter(Me, txt(Index), "编码,900,0,1;名称,2400,0,0;英文名,1200,0,0;临床意义,900,0,0", Me.Name & "\诊治项目过滤选择", "请从下表中选择一个项目", rsData, rs) Then
                    
                    txt(2).Text = zlCommFun.NVL(rs("名称").Value)
                    mlngKey = zlCommFun.NVL(rs("ID").Value)
                    cmdOpen.Tag = zlCommFun.NVL(rs("类型").Value, 0)
                    txt(2).Tag = ""
                    usrSaveGroup.诊治项目 = txt(2).Text
                    txt(1).Text = ""
                    
                Else
                    txt(2).Text = usrSaveGroup.诊治项目
                    Exit Sub
                End If
            Else
                zlCommFun.PressKey vbKeyTab
                zlCommFun.PressKey vbKeyTab
            End If
            txt(2).Tag = ""
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case Index
        Case 0
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "ABCDEFGHIJKLMNOPQRSTWXYZUV01234567890,-")
        Case 1
            If Val(cmdOpen.Tag) = 0 Then
                '数值型
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.-")
            End If
        End Select
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Index = 2 Then
        If (txt(2).Tag = "Changed") Then txt(2).Text = usrSaveGroup.诊治项目
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.选择))) = 1 Then
        mblnChangeEdit = True
        Call AdjustEnableState
        Exit Sub
    End If
        
    For mlngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.选择))) = 1 Then
            mblnChangeEdit = True
            Call AdjustEnableState
            Exit Sub
        End If
    Next
    
    If mlngLoop = vsf.Rows Then
        mblnChangeEdit = False
        Call AdjustEnableState
    End If
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf.RowData(Row)) = 0 Then Cancel = True
    If Col <> 0 Then Cancel = True
    
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

