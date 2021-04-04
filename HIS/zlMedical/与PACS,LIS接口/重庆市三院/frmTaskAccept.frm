VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmTaskAccept 
   Caption         =   "接受检验结果"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   8775
   Icon            =   "frmTaskAccept.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6615
      Top             =   2025
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3060
      Left            =   360
      TabIndex        =   0
      Top             =   1110
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   5398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "体检单"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "单位名称"
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3900
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":6852
            Key             =   "package"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":D0B4
            Key             =   "package_ok"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7995
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":13916
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":13B36
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":13D56
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskAccept.frx":144D0
            Key             =   "Send"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8775
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   1138
         ButtonWidth     =   1455
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  启动  "
               Key             =   "启动"
               Object.ToolTipText     =   "启动"
               Object.Tag             =   "  启动 "
               ImageKey        =   "Send"
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "指定"
               Key             =   "指定"
               Object.ToolTipText     =   "指定"
               Object.Tag             =   "指定"
               ImageKey        =   "Search"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1320
      Left            =   3195
      TabIndex        =   3
      Top             =   810
      Width           =   3135
      _cx             =   5530
      _cy             =   2328
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
      Cols            =   6
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   4845
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTaskAccept.frx":14C4A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10398
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
   Begin VB.Image imgY 
      Height          =   4680
      Left            =   3075
      MousePointer    =   9  'Size W E
      Top             =   660
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileParam 
         Caption         =   "参数设置(&P)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "接受(&E)"
      Begin VB.Menu mnuEditAutoAccept 
         Caption         =   "启动自动接受服务(&R)"
      End
      Begin VB.Menu mnuEditAcceptPerson 
         Caption         =   "接受指定人员数据(&A)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditResetPerson 
         Caption         =   "清除个人已接数据(&C)"
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
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frmTaskAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcnnSQLServer As New ADODB.Connection
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mstrKey As String

Private Enum mCol
    体检单号
    姓名
    性别
    身份证号
    门诊号
    登记id
    工作单位
    项目 = 0
    结果
    标志
    参考
End Enum

Private mlngCount As Long
Private mlngTotal As Long

Private Function ReadUnit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    lvw.ListItems.Clear
    
    Set objItem = lvw.ListItems.Add(, "_0", "--------", 1, 1)
    objItem.SubItems(1) = "[个人体检清单]"
    
    mstrSQL = "Select A.ID,A.体检号,B.名称 From 体检登记记录 A,合约单位 B Where A.是否团体=1 AND A.体检状态=4 AND B.ID=A.合约单位id"
    Set rs = OpenRecord(rs, mstrSQL, gstrSysName)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            Set objItem = lvw.ListItems.Add(, "_" & rs("ID").Value, rs("体检号").Value, 2, 2)
            objItem.SubItems(1) = rs("名称").Value
    
            rs.MoveNext
        Loop
    End If
        
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
End Function

Private Function ReadItems(ByVal lng登记id As Long, ByVal lng病人id As Long) As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
'    vsfItem.Rows = 2
'    vsfItem.RowData(1) = 0
'    vsfItem.Cell(flexcpText, 1, 0, 1, vsfItem.Cols - 1) = ""

    mstrSQL = ""

'    rs.Open mstrSQL, gcnOracle
'    If rs.BOF = False Then
'
'        Call LoadGrid(vsf, rs)
'
'    End If
    
    ReadItems = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
    
End Function

Private Function ReadPerson(ByVal lng登记id As Long) As Boolean
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    If lng登记id = 0 Then
        '个人
        
        mstrSQL = "Select B.病人id As ID,C.体检号 As 体检单号,A.登记id,B.姓名,B.性别,B.身份证号,B.门诊号,B.工作单位 From 体检人员档案 A,病人信息 B,体检登记记录 C Where A.病人id=B.病人id AND C.ID=A.登记id AND C.体检状态=4 AND a.体检报到=1 AND C.合约单位id Is Null"
        
    Else
        
        '团体
        mstrSQL = "Select B.病人id As ID,C.体检号 As 体检单号,A.登记id,B.姓名,B.性别,B.身份证号,B.门诊号,B.工作单位 From 体检人员档案 A,病人信息 B,体检登记记录 C Where a.体检报到=1 and A.病人id=B.病人id AND C.ID=A.登记id AND C.ID=" & lng登记id
        
    End If
    rs.Open mstrSQL, gcnOracle
    If rs.BOF = False Then
         
        Call LoadGrid(vsf, rs)
        
    End If
    
    ReadPerson = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
    
End Function

Private Function ConnectSQLServer(ByVal strSvr As String, ByVal strDb As String, ByVal strUser As String, ByVal strPsw As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    If mcnnSQLServer.State = adStateOpen Then mcnnSQLServer.Close
    mcnnSQLServer.Open "Provider=SQLOLEDB.1;Password=" & strPsw & ";Persist Security Info=True;User ID=" & strUser & ";Initial Catalog=" & strDb & ";Data Source=" & strSvr
    If mcnnSQLServer.State <> adStateOpen Then
        
        ShowSimpleMsg "连接到LIS服务器失败！"
        
        Exit Function
    End If
    
    ConnectSQLServer = True
    Exit Function
errHand:
    ShowSimpleMsg Err.Description
End Function

Private Function AcceptResult() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngTotal As Long
    Dim lngCount As Long
    
    On Error GoTo errHand
    
    frmWait.OpenWait Me, "接受检验数据", True
    
    mstrSQL = "Select A.登记id,A.病人id,B.姓名 From 体检人员档案 A,病人信息 B Where A.病人id=B.病人id AND A.体检状态=4"
    rs.Open mstrSQL, gcnOracle
    If rs.BOF = False Then
        
        lngTotal = rs.RecordCount
        For lngCount = 1 To lngTotal
            
            frmWait.WaitInfo = "正在接受“" & NVL(rs("姓名")) & "”的检验数据..."
            frmWait.WaitProgress = Format(100 * lngCount / lngTotal, "0.00")
            
            Call AcceptOneResult(NVL(rs("登记id"), 0), NVL(rs("病人id"), 0))
            rs.MoveNext
        Next
        
        frmWait.CloseWait
    End If
    
    frmWait.CloseWait
    AcceptResult = True
    
    Exit Function
    
errHand:
    frmWait.CloseWait
    ShowSimpleMsg Err.Description
End Function

Private Function ClearOneResult(ByVal lng登记id As Long, ByVal lng病人id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim str体检单号 As String
    Dim lng门诊号 As Long
    
    '检查是否已经接受，只接受未接受的，通过标本号判断
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    mstrSQL = "Select C.体检号,B.门诊号 From 体检人员档案 A,病人信息 B,体检登记记录 C Where A.病人id=B.病人id AND C.ID=A.登记id"
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.Open mstrSQL, gcnOracle
    If rsTmp.BOF Then Exit Function
    
    str体检单号 = UCase(NVL(rsTmp("体检号")))
    lng门诊号 = NVL(rsTmp("门诊号"), 0)
            
    '清除原有的检验结果
    mstrSQL = "ZL_ZLLIS_清除结果('" & str体检单号 & "'," & lng病人id & ")"
    gcnOracle.Execute mstrSQL, , adCmdStoredProc

    gcnOracle.CommitTrans
    
    ClearOneResult = True
    
    ShowSimpleMsg "已经清除了此人已接受的检验数据！"
    Exit Function
    
errHand:
    
    ShowSimpleMsg Err.Description
    gcnOracle.RollbackTrans
    
End Function

Private Function AcceptOneResult(ByVal lng登记id As Long, ByVal lng病人id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim str体检单号 As String
    Dim lng门诊号 As Long
    Dim strCode As String
    Dim lng申请id As Long
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim lng组合项目id As Long
    
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    mstrSQL = "Select C.体检号,B.门诊号 " & _
            "From 体检人员档案 A,病人信息 B,体检登记记录 C Where A.病人id=B.病人id AND C.ID=A.登记id AND C.ID=" & lng登记id & " AND A.病人id=" & lng病人id
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.Open mstrSQL, gcnOracle
    If rsTmp.BOF Then Exit Function
    
    str体检单号 = UCase(NVL(rsTmp("体检号")))
    lng门诊号 = NVL(rsTmp("门诊号"), 0)
    
    '填写新的检验结果
    mstrSQL = "Select 检验项目代码,项目名称,单位,数值性结果,字符性结果,参与范围,异常情况,检验者 " & _
                "From Lis_Value Where 体检单号='" & str体检单号 & "' AND 病人id='" & lng门诊号 & "'"
                
    rs.Open mstrSQL, mcnnSQLServer
    If rs.BOF = False Then
        Do While Not rs.EOF
            
'            If NVL(rs("检验项目代码")) <> "" Then
'                strCode = "''"
'                varAry = Split(NVL(rs("检验项目代码")), ",")
'                For lngLoop = 0 To UBound(varAry)
'                    strCode = strCode & ",'" & varAry(lngLoop) & "'"
'                Next
'            End If
'            strCode = "'" & NVL(rs("检验项目代码")) & "'"

            lng申请id = 0
            
            '1.找申请id
'            mstrSQL = "Select LIS组合编码 From 诊疗项目目录_LIS Where LIS编码='" & NVL(rs("检验项目代码")) & "'"
'            If rsTmp.State = adStateOpen Then rsTmp.Close
'            rsTmp.Open mstrSQL, gcnOracle
'            If rsTmp.BOF = False Then strCode = strCode & ",'" & NVL(rsTmp("LIS组合编码")) & "'"

            lng组合项目id = 0
                        
            '找组合项目id
            mstrSQL = "Select e.id " & _
                        "From   诊疗项目目录_LIS a," & _
                                "诊疗项目目录 b," & _
                                "检验报告项目 c," & _
                                "检验报告项目 d," & _
                                "诊疗项目目录 e " & _
                        "where  a.诊疗项目id=b.id " & _
                                "and Nvl(b.组合项目,0)=0 " & _
                                "and c.诊疗项目id=b.id " & _
                                "and c.报告项目id=d.报告项目id " & _
                                "and d.诊疗项目id=e.id " & _
                                "and e.组合项目=1 and instr(','||a.LIS编码||',','," & NVL(rs("检验项目代码")) & ",')>0"
                        
            If rsTmp.State = adStateOpen Then rsTmp.Close
            rsTmp.Open mstrSQL, gcnOracle
            If rsTmp.BOF = False Then
                
                lng组合项目id = NVL(rsTmp("ID"))
                
                mstrSQL = "Select   C.医嘱id As 申请id " & _
                            "From   体检项目清单 B," & _
                                    "体检项目医嘱 C, " & _
                                    "检验报告项目 D " & _
                            "Where  B.诊疗项目id= " & lng组合项目id & _
                                    " AND C.清单ID=B.ID  " & _
                                    " AND C.病人id=" & lng病人id & " " & _
                                    " AND B.登记id=" & lng登记id
                                    
    '            mstrSQL = "Select C.医嘱id As 申请id " & _
    '                        "From   (Select 诊疗项目id From 诊疗项目目录_LIS Where LIS编码 In (" & strCode & ")) A, " & _
    '                                "体检项目清单 B," & _
    '                                "体检项目医嘱 C " & _
    '                        "Where  A.诊疗项目id=B.诊疗项目id " & _
    '                                "AND C.清单ID=B.ID " & _
    '                                "AND C.病人id=" & lng病人id & " " & _
    '                                "AND B.登记id=" & lng登记id
                
                If rsTmp.State = adStateOpen Then rsTmp.Close
                rsTmp.Open mstrSQL, gcnOracle
                If rsTmp.BOF = False Then lng申请id = NVL(rsTmp("申请id"), 0)
                
                If lng申请id > 0 Then
                    
                    '2.找对应的项目
                    strCode = "'" & NVL(rs("检验项目代码")) & "'"
                    
                    mstrSQL = "Select C.ID,C.中文名 " & _
                                "From   诊疗项目目录_LIS A," & _
                                        "检验报告项目 B," & _
                                        "诊治所见项目 C  " & _
                                "Where C.ID=B.报告项目id AND A.诊疗项目id=B.诊疗项目id AND A.LIS编码=" & strCode
                    
                    If rsTmp.State = adStateOpen Then rsTmp.Close
                    rsTmp.Open mstrSQL, gcnOracle
                    
                    If rsTmp.BOF = False Then
                        mstrSQL = "ZL_ZLLIS_填写结果("
                        
                        mstrSQL = mstrSQL & lng申请id & ","
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rsTmp("中文名"))) & "',"
                        mstrSQL = mstrSQL & NVL(rsTmp("ID"), 0) & ","
                        
                        If IsNull(rs("数值性结果").Value) = False Then
                            mstrSQL = mstrSQL & "'" & Trim(NVL(rs("数值性结果"))) & "',"
                            mstrSQL = mstrSQL & "1,"
                        Else
                            mstrSQL = mstrSQL & "'" & Trim(NVL(rs("字符性结果"))) & "',"
                            mstrSQL = mstrSQL & "0,"
                        End If
                        
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rs("单位").Value)) & "',"
                        
                        Select Case Trim(NVL(rs("异常情况")))
                        Case "L", "l"
                            mstrSQL = mstrSQL & "'偏低',"
                        Case "H", "h"
                            mstrSQL = mstrSQL & "'偏高',"
                        Case Else
                            mstrSQL = mstrSQL & "'正常',"
                        End Select
                        
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rs("参与范围"))) & "',"
                        mstrSQL = mstrSQL & "'" & Trim(NVL(rs("检验者"))) & "'"
                        
                        mstrSQL = mstrSQL & ")"
                        
                        gcnOracle.Execute mstrSQL, , adCmdStoredProc
                        
                    End If
                End If
            End If
            rs.MoveNext
        Loop
    End If
    
    gcnOracle.CommitTrans
    
    AcceptOneResult = True
    
    Exit Function
    
errHand:
    
    ShowSimpleMsg Err.Description
    gcnOracle.RollbackTrans
    
End Function

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    Dim strUser As String
    Dim strPsw As String
    Dim strSvr As String
    
    strUser = GetSetting("ZLSOFT", "注册信息\登陆信息_LIS", "USER", "HISJQ")
    strSvr = GetSetting("ZLSOFT", "注册信息\登陆信息_LIS", "SERVER", "")
    
    If frmLisLogin.ShowLogin(Me, strUser, strPsw, strSvr) Then
        
        If ConnectSQLServer(strSvr, "CliMis", strUser, strPsw) = False Then
            Unload Me
            Exit Sub
        End If
    Else
        Unload Me
        Exit Sub
    End If
    
    SaveSetting "ZLSOFT", "注册信息\登陆信息_LIS", "USER", strUser
    SaveSetting "ZLSOFT", "注册信息\登陆信息_LIS", "SERVER", strSvr
    
    Call ReadUnit
    
    If Not (lvw.SelectedItem Is Nothing) Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
    
End Sub

Private Sub Form_Load()
    Dim strVsf As String
    
    mblnStartUp = True
    
    strVsf = "体检单号,900,1,1,1,;姓名,900,1,1,1,;性别,600,1,1,1,;身份证号,1800,1,1,1,;门诊号,1200,1,1,1,;登记id,0,1,1,1,;工作单位,1200,1,1,1,"
    Call CreateVsf(vsf, strVsf)
'
'    With vsfItem
'        .Cols = 4
'
'        .TextMatrix(0, mCol.项目) = "项目"
'        .TextMatrix(0, mCol.结果) = "结果"
'        .TextMatrix(0, mCol.标志) = "标志"
'        .TextMatrix(0, mCol.参考) = "参考"
'
'        .ColWidth(mCol.项目) = 1500
'        .ColWidth(mCol.结果) = 1200
'        .ColWidth(mCol.标志) = 600
'        .ColWidth(mCol.参考) = 1500
'    End With
    
    mstrKey = ""
    mlngTotal = Val(GetSetting("ZLSOFT", "公共全局\检验接口", "接受间隔", "10"))
    If mlngTotal < 5 Then mlngTotal = 5
    If mlngTotal > 30 Then mlngTotal = 30
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With lvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
       
    With imgY
        .Top = lvw.Top
        .Height = lvw.Height
    End With
    
    With vsf
        .Left = imgY.Left + imgY.Width
        .Top = lvw.Top + 30
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
'    With vsfItem
'        .Left = vsf.Left
'        .Top = vsf.Top + vsf.Height + 45
'        .Width = vsf.Width
'    End With
    
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If mstrKey <> Item.Key Then
        
        mstrKey = Item.Key
        Call ReadPerson(Val(Mid(Item.Key, 2)))
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        
    End If
    
    
End Sub

Private Sub mnuEditAcceptPerson_Click()
    Dim blnSvr As Boolean
    
    If Val(vsf.RowData(vsf.Row)) <= 0 Then Exit Sub
    
    If MsgBox("确实需要接受此人的检验数据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    blnSvr = tmr.Enabled
    tmr.Enabled = False
    
    frmWait.OpenWait Me, "接受检验数据", True
    frmWait.WaitInfo = "正在接受“" & vsf.TextMatrix(vsf.Row, mCol.姓名) & "”检验数据"
    
    If AcceptOneResult(Val(vsf.TextMatrix(vsf.Row, mCol.登记id)), Val(vsf.RowData(vsf.Row))) Then
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    End If
    
    frmWait.CloseWait
    tmr.Enabled = blnSvr
    
End Sub

Private Sub mnuEditAutoAccept_Click()
    '
    mnuEditAutoAccept.Checked = Not mnuEditAutoAccept.Checked
    
    If mnuEditAutoAccept.Checked Then
        tbrThis.Buttons("启动").Value = tbrPressed
        tmr.Enabled = True
    Else
        tbrThis.Buttons("启动").Value = tbrUnpressed
        tmr.Enabled = False
    End If
    
End Sub

Private Sub mnuEditResetPerson_Click()
    
    Dim rs As New ADODB.Recordset
    
    If Val(vsf.RowData(vsf.Row)) = 0 Then Exit Sub
    
    On Error GoTo errHand
    
    If MsgBox("确实需要清除此人已接受的数据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If ClearOneResult(Val(vsf.TextMatrix(vsf.Row, mCol.登记id)), Val(vsf.RowData(vsf.Row))) Then
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    End If
    
    Exit Sub
    
errHand:
    ShowSimpleMsg Err.Description
    
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileParam_Click()
    
    If frmTaskAcceptParam.ShowParam(Me) Then
        mlngTotal = Val(GetSetting("ZLSOFT", "公共全局\检验接口", "接受间隔", "10"))
        If mlngTotal < 5 Then mlngTotal = 5
        If mlngTotal > 30 Then mlngTotal = 30
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuHelpTopic_Click()
    Call ShowHelp(Me.hWnd, Me.Name)
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
    Case "启动"
        Call mnuEditAutoAccept_Click
    Case "指定"
        Call mnuEditAcceptPerson_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tmr_Timer()
    
    mlngCount = mlngCount + 1
    
    If mlngCount >= mlngTotal Then
        mlngCount = 0
        tmr.Enabled = False
        If AcceptResult Then
            Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        End If
        tmr.Enabled = True
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If OldRow <> NewRow Then
        Call ReadItems(Val(vsf.TextMatrix(NewRow, mCol.登记id)), Val(vsf.RowData(NewRow)))
    End If
    
End Sub

