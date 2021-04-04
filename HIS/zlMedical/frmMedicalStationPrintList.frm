VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMedicalStationPrintList 
   Caption         =   "体检报告单"
   ClientHeight    =   5865
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9510
   Icon            =   "frmMedicalStationPrintList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9510
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2880
      Left            =   300
      TabIndex        =   0
      Top             =   915
      Width           =   6825
      _cx             =   12039
      _cy             =   5080
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
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
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
   Begin VB.Frame fraInfo 
      Height          =   630
      Left            =   450
      TabIndex        =   6
      Top             =   4305
      Width           =   3165
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   675
         Picture         =   "frmMedicalStationPrintList.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1980
         TabIndex        =   2
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.姓名"
         Height          =   180
         Index           =   1
         Left            =   1020
         TabIndex        =   1
         Tag             =   "姓名"
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Tag             =   "姓名"
         Top             =   285
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5505
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationPrintList.frx":09F0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11721
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   9510
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9390
         _ExtentX        =   16563
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
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.个人"
               Key             =   "个人"
               Object.ToolTipText     =   "个人(Alt+S)"
               Object.Tag             =   "&S.个人"
               ImageIndex      =   5
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&G.团体"
               Key             =   "团体"
               Object.ToolTipText     =   "团体(Alt+G)"
               Object.Tag             =   "&G.团体"
               ImageIndex      =   6
               Style           =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.全选"
               Key             =   "全选"
               Object.ToolTipText     =   "全选(Alt+A)"
               Object.Tag             =   "&A.全选"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.全清"
               Key             =   "全清"
               Object.ToolTipText     =   "全清(Alt+C)"
               Object.Tag             =   "&C.全清"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8010
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":1284
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":19FE
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":2178
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":2398
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":25B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":2D32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8685
      Top             =   1185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":34AC
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":3C26
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":43A0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":45C0
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":47E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintList.frx":4EDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileRptSingle 
         Caption         =   "个人报告单(&S)"
         Begin VB.Menu mnuFileRptSinglePrintView 
            Caption         =   "预览(&1)"
         End
         Begin VB.Menu mnuFileRptSinglePrint 
            Caption         =   "打印(&2)"
         End
         Begin VB.Menu mnuFileRptSingleOutExcel 
            Caption         =   "输出到Excel(&3)"
         End
      End
      Begin VB.Menu mnuFileRptGroup 
         Caption         =   "团体报告单(&G)"
         Begin VB.Menu mnuFileRptGroupPrintView 
            Caption         =   "预览(&1)"
         End
         Begin VB.Menu mnuFileRptGroupPrint 
            Caption         =   "打印(&2)"
         End
         Begin VB.Menu mnuFileRptGroupOutExcel 
            Caption         =   "输入到Excel(&3)"
         End
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "全清(&C)"
         Shortcut        =   ^C
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
Attribute VB_Name = "frmMedicalStationPrintList"
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

Private Enum mCol
    选择 = 0
    姓名
    体检编号
    性别
    出生日期
    婚姻状况
    组别
End Enum

Private WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

'（２）自定义过程或函数************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    
    
    mnuFileRptSinglePrint.Enabled = True
    mnuFileRptSinglePrintView.Enabled = True
    mnuFileRptSingleOutExcel.Enabled = True
        
    If vData = False Then
        mnuFileRptSinglePrint.Enabled = False
        mnuFileRptSinglePrintView.Enabled = False
        mnuFileRptSingleOutExcel.Enabled = False
    End If
      
    
End Property

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

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False

    mlngKey = lngKey
    
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadData(mlngKey, lng病人id) = False Then Exit Function
    
    '如果是单个人,直接处理,不弹出界面
    If lng病人id > 0 Then
        mnuFileRptGroup.Visible = False
        tbrThis.Buttons("团体").Visible = False
    End If
    
    EditChanged = (Val(vsf.RowData(1)) > 0)

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long, ByVal lng病人id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
    
    gstrSQL = "SELECT 1 AS 选择,A.病人id AS ID,B.姓名,B.性别,A.体检编号,B.婚姻状况,TO_CHAR(B.出生日期,'yyyy-mm-dd') AS 出生日期,A.组别名称 AS 组别,'' AS 未到原因 " & _
                "FROM 体检人员档案 A,病人信息 B " & _
                "WHERE A.体检报到=1 AND A.体检状态 IN (4,5) AND A.病人id=B.病人id and A.登记id=" & lngKey
    If lng病人id > 0 Then gstrSQL = gstrSQL & " AND B.病人id=" & lng病人id
    
    gstrSQL = gstrSQL & " Order By A.体检编号"
    
    Call OpenRecord(rs, gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
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
    Dim strVsf As String
    
    On Error GoTo errHand
    
    strVsf = "选择,450,1,1,1,;姓名,1500,1,1,1,;体检编号,810,7,1,1,;性别,810,1,1,1,;出生日期,1080,1,1,1,;婚姻状况,1200,1,1,1,;组别,1200,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
    vsf.Editable = True
    
    Call AppendRows(vsf, lnX, lnY)
    
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

Private Function PrintData(ByVal bytMode As Byte, Optional ByVal blnGroup As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngLoop As Long

    On Error GoTo errHand
    
    If blnGroup Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_3", Me, "登记id=" & mlngKey, bytMode)
    Else
        
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, mCol.选择))) = 1 Then
                
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_2", Me, "登记id=" & mlngKey, "病人id=" & Val(vsf.RowData(lngLoop)), bytMode)
                
                '如果是预览，只一次预览
                If bytMode = 1 Then Exit For
                
            End If
        Next
    End If
    
    PrintData = True

    Exit Function

errHand:

    If ErrCenter = 1 Then Resume

End Function


Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.x * Screen.TwipsPerPixelX, objPoint.y * Screen.TwipsPerPixelY - 650)
    
    txt(1).Text = ""
    LocationObj txt(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("全选").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全选"))
        Case vbKeyC
            If tbrThis.Buttons("全清").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全清"))
        Case vbKeyV
            If tbrThis.Buttons("预览").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("预览"))
        Case vbKeyP
            If tbrThis.Buttons("打印").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("打印"))
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

    Call RestoreWinState(Me, App.ProductName)
       
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0")) = 1 Then
        '使用个性化设置
      
        lbl(1).Caption = "&6." & (GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", "姓名"))
        lbl(1).Tag = Mid(lbl(1).Caption, 4)

    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With vsf
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fraInfo.Height + 90
    End With
    
    With fraInfo
        .Left = 0
        .Top = vsf.Top + vsf.Height - 75
        .Width = vsf.Width
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", lbl(1).Tag)
End Sub


Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.选择) = 0
        End If
    Next
    
    EditChanged = False
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileRptGroupOutExcel_Click()
    Call PrintData(3, True)
End Sub

Private Sub mnuFileRptGroupPrint_Click()
    Call PrintData(2, True)
End Sub

Private Sub mnuFileRptGroupPrintView_Click()
    Call PrintData(1, True)
End Sub

Private Sub mnuFileRptSingleOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFileRptSinglePrint_Click()
    
    If PrintData(2) Then
        ShowSimpleMsg " 已经打印完成！ "
    End If
    
End Sub

Private Sub mnuFileRptSinglePrintView_Click()
    
    Call PrintData(1)
    
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.选择) = 1
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
    
    Select Case mbytPopMenu

    Case 3
        
        mobjPopMenu.Add 1, "&1.姓名", , , True, , (lbl(1).Tag = "姓名")
        mobjPopMenu.Add 2, "&2.体检编号", , , True, , (lbl(1).Tag = "体检编号")
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu

    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(1).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub
Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "全选"
        Call mnuFileSelectAll_Click
    Case "全清"
        Call mnuFileClearAll_Click
    Case "个人"
        Me.PopupMenu mnuFileRptSingle, , Button.Left + 90, Button.Top + Button.Height + 45
    Case "团体"
        Me.PopupMenu mnuFileRptGroup, , Button.Left + 90, Button.Top + Button.Height + 45
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngLoop As Long
    Dim strCol As String
    Dim lngCol As Long
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    Dim lngRow As Long
    
    If KeyAscii = vbKeyReturn Then
        
        If Index = 1 And Trim(txt(Index).Text) <> "" Then
            
            strCol = Mid(lbl(1).Caption, 4)
            lngCol = GetCol(vsf, strCol)
            If lngCol < 0 Then Exit Sub
            
            lngRow = 0
            If vsf.Row + 1 <= vsf.Rows - 1 Then
                For lngLoop = vsf.Row + 1 To vsf.Rows - 1
                    If InStr(vsf.TextMatrix(lngLoop, lngCol), txt(Index).Text) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf.Row
                    If InStr(vsf.TextMatrix(lngLoop, lngCol), txt(Index).Text) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "没有找到符合要求的信息！"
                txt(Index).Text = ""
            Else
                vsf.ShowCell lngRow, vsf.Col
                vsf.Row = lngRow
            End If
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.选择))) = 1 Then
        EditChanged = True
        Exit Sub
    End If
        
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.选择))) = 1 Then
            EditChanged = True
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

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.选择 Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

