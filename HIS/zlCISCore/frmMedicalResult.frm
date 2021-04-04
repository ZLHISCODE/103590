VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMedicalResult 
   Caption         =   "体检结果评估"
   ClientHeight    =   5700
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9180
   Icon            =   "frmMedicalResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9180
   Begin VB.PictureBox picTitle 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9165
      TabIndex        =   1
      Top             =   735
      Width           =   9165
      Begin VB.CommandButton cmdSearch 
         Caption         =   "开始评估(&F)"
         Height          =   350
         Left            =   7785
         TabIndex        =   2
         Top             =   90
         Width           =   1320
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "姓名:张三 性别:男 年龄:60 婚姻状况:已婚"
         Height          =   180
         Left            =   60
         TabIndex        =   3
         Top             =   195
         Width           =   3510
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalResult.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11113
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
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2205
      Left            =   345
      TabIndex        =   4
      Top             =   1605
      Width           =   4920
      _cx             =   8678
      _cy             =   3889
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
      HighLight       =   0
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
         Index           =   1
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   1
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9180
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
         TabIndex        =   6
         Top             =   30
         Width           =   9060
         _ExtentX        =   15981
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
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存(Alt+S)"
               Object.Tag             =   "&S.保存"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Left            =   7950
      Top             =   1050
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
            Picture         =   "frmMedicalResult.frx":0E1E
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":1598
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":1D12
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":1F2C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":214C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1050
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
            Picture         =   "frmMedicalResult.frx":236C
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":2AE6
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":3260
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":347A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalResult.frx":369A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSelAll 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu mnuFileClsAll 
         Caption         =   "全清(&C)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_2 
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
Attribute VB_Name = "frmMedicalResult"
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

Private mlng病人id As Long
Private mlng医嘱id As Long
Private mstr挂号单 As String

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
    
    tbrThis.Buttons("保存").Enabled = mnuFileSave.Enabled
            
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    Call AppendRows(vsf, lnX, lnY)
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
                
    '病人id,医嘱id,挂号单
    
    mlng病人id = Val(Split(strParam, "'")(0))
    mlng医嘱id = Val(Split(strParam, "'")(1))
    mstr挂号单 = Split(strParam, "'")(2)
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If mlng病人id > 0 Then Call ReadData(mlng病人id)
    
    EditChanged = False
    
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
    Dim strSQL As String
    
    On Error GoTo errHand
    
    '读取病人信息
    strSQL = "SELECT * FROM 病人信息 WHERE 病人id=" & lngKey
    Call OpenRecord(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
        lblInfo.Caption = "姓名:" & zlCommFun.NVL(rs("姓名")) & " 性别:" & zlCommFun.NVL(rs("性别")) & " 年龄:" & zlCommFun.NVL(rs("年龄")) & " 婚姻状况:" & zlCommFun.NVL(rs("婚姻状况"))
    End If
                                
    Call AppendRows(vsf, lnX, lnY)
    
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
    
    vsf.Cols = 3
    vsf.ColWidth(0) = 255
    vsf.ExtendLastCol = True
    
    vsf.TextMatrix(0, 0) = ""
    vsf.TextMatrix(0, 1) = "评估结果"
    vsf.TextMatrix(0, 2) = "异常结果"
    
    vsf.ColWidth(1) = 2400
    vsf.Editable = flexEDKbdMouse
    vsf.ColDataType(0) = flexDTBoolean
    
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

Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngLoop As Long
    
    
    On Error GoTo errHand
    
    
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function

Private Sub cmdSearch_Click()
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim blnFound As Boolean
    Dim bytSave As Byte
    Dim strResult As String
    Dim str性别 As String
    Dim str年龄 As String
    Dim strValue As String
    Dim strWorn As String
    Dim strRefence As String
    
    '开始评估体检结论
    
    cmdSearch.Enabled = False
    
    bytSave = vsf.Redraw
    vsf.Redraw = flexRDNone
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        
    stbThis.Panels(2).Text = "正在确定筛查范围..."
    DoEvents
    
    '提取有条件的诊断项目
    strSQL = _
            "SELECT A.名称,A.序号,B.分组名,A.是否疾病 " & _
            "FROM   体检诊断建议 A, " & _
                    "体检诊断评估 B " & _
            "Where A.序号 = B.诊断序号 " & _
            "GROUP BY A.名称,A.序号,B.分组名,A.是否疾病"
            
    Call OpenRecord(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
    
        strSQL = "Select 性别,年龄 From 病人信息 Where 病人id=[1]"
        Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人id)
        If rs2.BOF Then Exit Sub
        
        str性别 = zlCommFun.NVL(rs2("性别").Value)
        str年龄 = zlCommFun.NVL(rs2("年龄").Value)
'        If str年龄 = "" Then str年龄 = zlCommFun.NVL(rs2("实际年龄").Value)
        
        
        Do While Not rs.EOF
            strResult = ""
            
            stbThis.Panels(2).Text = "正在进行“" & zlCommFun.NVL(rs("名称").Value) & "”评估..."
            DoEvents
            
            '提取判断条件
            strSQL = _
                    "SELECT B.ID,B.类型,B.替换域,B.中文名,A.关系式,A.条件值,A.性别,A.开始年龄,A.结束年龄 " & _
                    "FROM 体检诊断评估 A, " & _
                         "诊治所见项目 B " & _
                    "Where A.项目ID = B.ID " & _
                          "AND A.诊断序号=" & rs("序号").Value & " " & _
                          IIf(zlCommFun.NVL(rs("分组名")) = "", "AND A.分组名 IS NULL", "AND A.分组名='" & zlCommFun.NVL(rs("分组名")) & "'")
                          
            Call OpenRecord(rs2, strSQL, Me.Caption)
            If rs2.BOF = False Then
                
                '
                blnFound = True
                                
                Do While Not rs2.EOF
                    
                    strTmp = ""
                    
                    If zlCommFun.NVL(rs2("替换域").Value, 0) = 0 Then
                        '不是替换域
                        
                        '读取内容
                        strSQL = "SELECT S.标题,S.所见内容,Y.部件 " & _
                                " FROM 病人病历所见单 S,病人病历内容 X,病历元素目录 Y " & _
                                " WHERE X.ID=S.病历id And Y.编码(+)=X.元素编码 And S.所见项ID+0=" & rs2("ID").Value & _
                                "       AND S.病历ID=(" & _
                                "           SELECT MAX(S.病历ID)" & _
                                "           from (SELECT S.病历ID FROM 病人病历所见单 S WHERE S.所见项ID+0=" & rs2("ID").Value & ") S," & _
                                "                (SELECT C.ID,C.病历记录ID " & _
                                "                 FROM 病人病历记录 L,病人病历内容 C" & _
                                "                 WHERE L.ID=C.病历记录ID" & _
                                "                       AND L.病人ID=" & mlng病人id
                                
                        
                        If mlng医嘱id > 0 Then
                            strSQL = strSQL & "                       AND L.ID IN (SELECT DISTINCT 报告id FROM 病人医嘱发送 WHERE 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE 病人来源=4 AND ID=" & mlng医嘱id & " UNION ALL SELECT ID FROM 病人医嘱记录 WHERE 病人来源=4 AND 相关id=" & mlng医嘱id & "))"
                        Else
                            strSQL = strSQL & "                       AND L.ID IN (SELECT DISTINCT 报告id FROM 病人医嘱发送 WHERE 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE 病人来源=4 AND 挂号单='" & mstr挂号单 & "' and 病人id=" & mlng病人id & "))"
                        End If
                        
                        'strSQL = strSQL & "                       AND L.ID IN (SELECT DISTINCT 报告id FROM 病人医嘱发送 WHERE 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE 病人来源=4 AND 挂号单='" & mstr挂号单 & "' and 病人id=" & mlng病人id & "))"
                        
                        strSQL = strSQL & "       ) C" & _
                            "           WHERE C.ID=S.病历ID)"
                            
                        Call OpenRecord(rs3, strSQL, Me.Caption)
                        If rs3.BOF = False Then
                            strTmp = zlCommFun.NVL(rs3("所见内容").Value, "")
                        Else
                            blnFound = False
                            Exit Do
                        End If
                        
                    Else
                        '是替换域
                        
                        strTmp = GetSpecValue(rs2("中文名").Value, CStr(mlng病人id), "0", 0)
                        
                    End If
                    
                    
                    '调用病人条件判断
                    If zlCommFun.NVL(rs2("性别").Value) <> "" Then
                                                
                        If InStr(str性别, zlCommFun.NVL(rs2("性别").Value)) = 0 Then
                            '不成立
                            blnFound = False
                            Exit Do
                        End If
                        
                    End If
                    
                    If zlCommFun.NVL(rs2("开始年龄").Value) <> "" Or zlCommFun.NVL(rs2("结束年龄").Value) <> "" Then
                        '
                        If zlVerifyAge(str年龄, zlCommFun.NVL(rs2("开始年龄").Value), zlCommFun.NVL(rs2("开始年龄").Value)) = False Then
                            blnFound = False
                            Exit Do
                        End If
                        
                    End If
                        
                    strValue = strTmp
                    strWorn = ""
                    strRefence = ""
                    
                    If UCase(zlCommFun.NVL(rs3("部件").Value, "")) = "ZL9CISCORE.USRVERIFYREPORT" Then
                        If strTmp <> "" Then
                            
                            strValue = Split(strTmp, "'")(0)
                            strWorn = Split(strTmp, "'")(1)
                            strRefence = Split(strTmp, "'")(2)
                            
                            strTmp = Split(strTmp, "'")(0) & "(" & Split(strTmp, "'")(2) & ")"
                        End If
                    End If
                    
                    '调用满足判断
                    If Not zlVerifyValue(strValue, zlCommFun.NVL(rs2("类型"), 0), zlCommFun.NVL(rs2("关系式")), zlCommFun.NVL(rs2("条件值")), strWorn, strRefence) Then
                        blnFound = False
                        Exit Do
                    End If
                    
                    If strTmp <> "" Then
                        If strResult <> zlCommFun.NVL(rs3("标题").Value, "") & ":" & strTmp Then
                            strResult = strResult & zlCommFun.NVL(rs3("标题").Value, "") & ":" & strTmp
                        End If
                    End If
                    
                    
                    rs2.MoveNext
                Loop
                
                If blnFound Then
                    
                    '此项条件成立
                            
                    '判断是否已经有了
                    For lngLoop = 1 To vsf.Rows - 1
                        If Val(vsf.RowData(lngLoop)) = rs("序号").Value Then
                            Exit For
                        End If
                    Next
                                                
                    If lngLoop >= vsf.Rows Then
                        
                        '没有就增加上
                        
                        If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                            vsf.Rows = vsf.Rows + 1
                        End If
                        
                        vsf.RowData(vsf.Rows - 1) = rs("序号").Value
                        vsf.TextMatrix(vsf.Rows - 1, 0) = "1"
                        vsf.TextMatrix(vsf.Rows - 1, 1) = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(vsf.Rows - 1, 2) = strResult
                        vsf.Cell(flexcpData, vsf.Rows - 1, 1) = zlCommFun.NVL(rs("是否疾病").Value, 0)
                                                
                    End If
                    
                End If
                
            End If
            
            rs.MoveNext
        Loop
    End If
    
    If Val(vsf.RowData(1)) = 0 Then
        stbThis.Panels(2).Text = "没有找到评估结果。"
        
        EditChanged = False
    Else
        stbThis.Panels(2).Text = "共找到 " & vsf.Rows - 1 & " 条评估结果。"
        
        EditChanged = True
    End If
    
    vsf.Redraw = bytSave
    
    cmdSearch.Enabled = True
    AppendRows vsf, lnX, lnY
    
End Sub

Private Function zlVerifyAge(ByVal str年龄 As String, ByVal str开始年龄 As String, ByVal str结束年龄 As String) As Boolean
    
    Dim strAgeNumber As String
    Dim strAgeNumberBegin As String
    Dim strAgeNumberEnd As String
    Dim strAgeUnit As String
    
    On Error GoTo errHand
    
    If str开始年龄 = "" And str结束年龄 = "" Then
        zlVerifyAge = True
        Exit Function
    End If
    
    If str开始年龄 = "" And str结束年龄 <> "" Then str开始年龄 = str结束年龄
    If str开始年龄 <> "" And str结束年龄 = "" Then str结束年龄 = str开始年龄
        
    Call AnalyseAge(str开始年龄, strAgeNumberBegin, strAgeUnit)
    Select Case strAgeUnit
    Case "月"
        strAgeNumberBegin = Val(strAgeNumberBegin) * 30
    Case "年"
        strAgeNumberBegin = Val(strAgeNumberBegin) * 365
    End Select
    
    Call AnalyseAge(str结束年龄, strAgeNumberEnd, strAgeUnit)
    Select Case strAgeUnit
    Case "月"
        strAgeNumberEnd = Val(strAgeNumberEnd) * 30
    Case "年"
        strAgeNumberEnd = Val(strAgeNumberEnd) * 365
    End Select
    
    Call AnalyseAge(str年龄, strAgeNumber, strAgeUnit)
    Select Case strAgeUnit
    Case "月"
        strAgeNumber = Val(strAgeNumber) * 30
    Case "年"
        strAgeNumber = Val(strAgeNumber) * 365
    End Select
        
    If Val(strAgeNumber) >= Val(strAgeNumberBegin) And Val(strAgeNumber) <= Val(strAgeNumberEnd) Then
        zlVerifyAge = True
        Exit Function
    End If
    
    zlVerifyAge = False
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AnalyseAge(strOld As String, ByRef strAgeNumber As String, ByRef strAgeUnit As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '功能:将数据库中保存的年龄按估计的格式加载到界面
    
    Dim strTmp As Long
    
    If strOld = "岁" Then Exit Function
    
    If InStr(strOld, "岁") > 0 Then
        If InStr(strOld, "岁") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "岁") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "岁"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "月") > 0 Then
        If InStr(strOld, "月") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "月") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "月"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "天") > 0 Then
        If InStr(strOld, "天") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "天") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "天"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf IsNumeric(strOld) Then
        strAgeNumber = strOld
        strAgeUnit = "岁"
    Else
        strAgeNumber = strOld
        strAgeUnit = ""
    End If
    
    AnalyseAge = True
    
End Function

Private Function zlVerifyValue(strVerify As String, bytType As Byte, ByVal strFormula As String, ByVal strAskValue As String, ByVal strWorn As String, ByVal str正常结果 As String) As Boolean
    '-------------------------------------------------
    '功能：判断当前数据是否满足条件表达式
    '入参： strVerify-需判断的数值
    '       bytType-数值类型
    '       strFormula-关系式（文字说明）
    '       strAskValue-要求的数值或范围域
    '出参：正确返回true，否则返回false
    '-------------------------------------------------
    Dim aryTemp() As String
    Dim varTmp As Variant
    
    zlVerifyValue = False
    
    
    Select Case strAskValue
    Case "[最低值]"
    
        If InStr(str正常结果, "～") > 0 Then
            varTmp = Split(str正常结果, "～")
            strAskValue = Val(varTmp(0))
        End If
        
    Case "[最高值]"
        
        If InStr(str正常结果, "～") > 0 Then
            varTmp = Split(str正常结果, "～")
            strAskValue = Val(varTmp(1))
        End If
        
    Case "[偏低]"
        
        If Trim(strWorn) = "偏低" Then zlVerifyValue = True
        Exit Function
        
    Case "[偏高]"
        
        If Trim(strWorn) = "偏高" Then zlVerifyValue = True
        Exit Function
        
    Case "[异常]"
        
        If Trim(strWorn) = "异常" Then zlVerifyValue = True
        Exit Function
        
    End Select
    
    Select Case Val(bytType)
    Case 0  '数值
        Select Case Trim(strFormula)
        Case "等于"
            If Val(strVerify) = Val(strAskValue) Then zlVerifyValue = True
        Case "不等于"
            If Val(strVerify) <> Val(strAskValue) Then zlVerifyValue = True
        Case "大于"
            If Val(strVerify) > Val(strAskValue) Then zlVerifyValue = True
        Case "小于"
            If Val(strVerify) < Val(strAskValue) Then zlVerifyValue = True
        Case "小于等于"
            If Val(strVerify) <= Val(strAskValue) Then zlVerifyValue = True
        Case "大于等于"
            If Val(strVerify) >= Val(strAskValue) Then zlVerifyValue = True
        Case "介于", "在范围内"
            aryTemp = Split(strAskValue, "至")
            If UBound(aryTemp) = 1 Then
                aryTemp(0) = Trim(aryTemp(0))
                aryTemp(1) = Trim(aryTemp(1))
                
                If Val(strVerify) >= Val(aryTemp(0)) And Val(strVerify) <= Val(aryTemp(1)) Then zlVerifyValue = True
                If Val(strVerify) >= Val(aryTemp(1)) And Val(strVerify) <= Val(aryTemp(0)) Then zlVerifyValue = True
            End If
'        Case "存在"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Val(strVerify) & ",") > 0 Then zlVerifyValue = True
'        Case "不存在"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Val(strVerify) & ",") = 0 Then zlVerifyValue = True
        End Select
    Case 1  '文字
        Select Case Trim(strFormula)
        Case "等于"
            If Trim(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
        
        Case "大于"
            
            If Trim(strVerify) > Trim(strAskValue) Then zlVerifyValue = True
            
        Case "小于"
            
            If Trim(strVerify) < Trim(strAskValue) Then zlVerifyValue = True
            
        Case "大于等于"
            
            If Trim(strVerify) >= Trim(strAskValue) Then zlVerifyValue = True
            
        Case "小于等于"
            
            If Trim(strVerify) <= Trim(strAskValue) Then zlVerifyValue = True
            
        Case "不等于"
            If Trim(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
            
        Case "包含"
            If InStr(1, Trim(strVerify), Trim(strAskValue)) > 0 Then zlVerifyValue = True
            
        Case "不包含"
            If InStr(1, Trim(strVerify), Trim(strAskValue)) = 0 Then zlVerifyValue = True
'        Case "存在"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Trim(strVerify) & ",") > 0 Then zlVerifyValue = True
'        Case "不存在"
'            strAskValue = Replace(strAskValue, Space(1), "")
'            If InStr(1, "," & strAskValue & ",", "," & Trim(strVerify) & ",") = 0 Then zlVerifyValue = True
        End Select
'    Case 2  '日期
'        strVerify = Format(strVerify, "YYYY-MM-DD")
'        Select Case Trim(strFormula)
'        Case "等于"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
'        Case "不等于"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
'        Case "晚于"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) > Trim(strAskValue) Then zlVerifyValue = True
'        Case "早于"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) < Trim(strAskValue) Then zlVerifyValue = True
'        Case "不晚于"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) <= Trim(strAskValue) Then zlVerifyValue = True
'        Case "不早于"
'            strAskValue = Format(strAskValue, "YYYY-MM-DD")
'            If Trim(strVerify) >= Trim(strAskValue) Then zlVerifyValue = True
'        Case "介于", "在范围内"
'            aryTemp = Split(strAskValue, "至")
'            If UBound(aryTemp) = 1 Then
'                aryTemp(0) = Format(Trim(aryTemp(0)), "YYYY-MM-DD")
'                aryTemp(1) = Format(Trim(aryTemp(1)), "YYYY-MM-DD")
'                If Trim(strVerify) >= Trim(aryTemp(0)) And Trim(strVerify) <= Trim(aryTemp(1)) Then zlVerifyValue = True
'                If Trim(strVerify) >= Trim(aryTemp(1)) And Trim(strVerify) <= Trim(aryTemp(0)) Then zlVerifyValue = True
'            End If
'        End Select
    Case 2  '逻辑
        Select Case Trim(strFormula)
        Case "等于"
            If Val(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
        Case "不等于"
            If Val(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
        End Select
    Case Else
    End Select
End Function


'（３）窗体及其控件的事件处理******************************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("全选").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全选"))
        Case vbKeyC
            If tbrThis.Buttons("全清").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全清"))
        Case vbKeyS
            If tbrThis.Buttons("保存").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("保存"))
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
        
    With picTitle
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth
    End With
    
    With vsf
        .Left = 0
        .Top = picTitle.Top + picTitle.Height + 30
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
        
    cmdSearch.Left = picTitle.Width - cmdSearch.Width - 60
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuFileClear_Click()
    If MsgBox("确实要清除所设置的项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    EditChanged = True
    
End Sub

Private Sub mnuFileClsAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If vsf.TextMatrix(lngLoop, 0) <> "0" Then
                vsf.TextMatrix(lngLoop, 0) = "0"
                EditChanged = True
            End If
        End If
    Next
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
    
    If MsgBox("确实要恢复以前所选项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    Call ReadData(mlngKey)
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit Then
        
        On Error Resume Next
        
        Call mfrmMain.EditRefresh(vsf)
        
        On Error GoTo 0
        
        EditChanged = False
        
        Unload Me
        
    End If
    
End Sub

Private Sub mnuFileSelAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If vsf.TextMatrix(lngLoop, 0) = "0" Then
                vsf.TextMatrix(lngLoop, 0) = "1"
                EditChanged = True
            End If
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
    cbrThis.Bands(1).MINHEIGHT = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "保存"
        Call mnuFileSave_Click
    Case "全选"
        Call mnuFileSelAll_Click
    Case "全清"
        Call mnuFileClsAll_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    vsf.TextMatrix(Row, Col) = Abs(vsf.Value)
    EditChanged = True
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = 0)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col <> 0 Then Cancel = True
    If Val(vsf.RowData(Row)) = 0 Then Cancel = True
    
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

