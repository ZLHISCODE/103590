VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIllImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "疾病编码导入"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10815
   Icon            =   "frmIllImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10815
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10815
      TabIndex        =   15
      Top             =   7710
      Width           =   10815
      Begin VB.CommandButton cmdCancel 
         Caption         =   "退出(&C)"
         Height          =   350
         Left            =   9600
         TabIndex        =   17
         Tag             =   "分类"
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "导入(&O)"
         Height          =   350
         Left            =   8400
         TabIndex        =   16
         Tag             =   "分类"
         Top             =   240
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar prg 
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   90
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin TabDlg.SSTab sstType 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "EXCEL"
      TabPicture(0)   =   "frmIllImport.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "脚本"
      TabPicture(1)   =   "frmIllImport.frx":686E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   6
         Top             =   420
         Width           =   10335
         Begin VB.CommandButton cmd 
            Caption         =   "文件(&F)"
            Height          =   350
            Index           =   0
            Left            =   9000
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   615
            Width           =   1100
         End
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   9
            ToolTipText     =   "疾病编码分类EXCEL表格路径"
            Top             =   600
            Width           =   7935
         End
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   2295
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   1440
            Width           =   10095
            _cx             =   17806
            _cy             =   4048
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   500
            ColWidthMax     =   10000
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
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   2535
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   4170
            Width           =   10095
            _cx             =   17806
            _cy             =   4471
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   300
            ColWidthMin     =   500
            ColWidthMax     =   10000
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "说明:EXCEL文件包含【疾病编码分类】和【疾病编码目录】两个表单，表单样式见下方表格示例。"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   7740
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "疾病编码目录表格示例"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   3930
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "文件位置"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "疾病编码分类表格示例"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   10335
         Begin VB.CommandButton cmd 
            Caption         =   "文件(&F)"
            Height          =   350
            Index           =   1
            Left            =   9000
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   750
            Width           =   1100
         End
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   2
            ToolTipText     =   "疾病编码分类EXCEL表格路径"
            Top             =   750
            Width           =   7935
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "说明:请选择需要执行的脚本文件。文件内容必须是由【疾病编码管理】导出的脚本样式。"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   7110
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "文件位置"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmIllImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mstrTYPE As String = "章,编码范围,名称"
Private Const mstrCONTENT As String = "编码,附码,名称"

Private mconn As ADODB.Connection
Private mRsType As ADODB.Recordset
Private mrsContent As ADODB.Recordset
Private mrs类别 As ADODB.Recordset
Private mbytModel  As Byte   '0-EXCEL方式,不导入分类;1-Excel方式:导入分类;2-脚步导入,不导入分类;3-脚步导入,导入分类

Private Enum E_ITEM
    E_分类 = 0
    E_目录 = 1
End Enum

Private Enum E_PAGE
    E_EXCEL = 0
    E_SCRIPT = 1
End Enum

Private Sub InitVsItem()
'功能:初始化示例表格
    Dim strHead As String
    Dim strRowContent As String 'strRowContent=表格的预定义行内容,格式为：列1,内容1,列2,内容2:行1;列1,内容1,列2,内容2:行2;
    
    strHead = "章,2000,4;编码范围,2000,1;名称,6000,1"
    
    strRowContent = "0,第一章,1,A00-B99,2,某些传染病和寄生虫病;0,,1,A00-A09,2,肠道传染病;0,,1,A15-Al9,2,结核病;0,,1,A20-A28,2,某些动物源性细菌性疾病;" & vbCrLf & _
                    "0,,1,B99-B99,2,其他传染病;" & vbCrLf & _
                    "0,第二章,1,C00-D48,2,肿瘤;0,,1,C00-C14,2,唇、口腔和咽恶性肿瘤;0,,1,C15-C26,2,消化器官恶性肿瘤;0,,1,C30-C39,2,呼吸和胸腔内器官恶性肿瘤"
    Grid.Init vsItem(E_分类), strHead, strRowContent, 0, 1
    strHead = "编码,2000,1;附码,2000,1;名称,6000,1"

    strRowContent = "0,A00.000,1,,2,古典生物型霍乱;0,A00.100,1,,2,埃尔托型霍乱;" & vbCrLf & _
                    "0,A01.001+,1,K77.0*,2,伤寒性肝炎;0,A01.100,1,,2,副伤寒甲"
    Grid.Init vsItem(E_目录), strHead, strRowContent, 0, 1
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = E_EXCEL Then
        OpenFile "EXCEL Files(*.xls,*.xlsx)|*.xls;*.xlsx", Index
    Else
        OpenFile "SQL Files(*.sql)|*.SQL", Index
    End If
    Set mconn = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strFile As String, StrInfo As String
    Dim objFile As New FileSystemObject
    Dim i As Byte
    Dim objExcel As Object      'Excel.Application '定义Excel类
    Dim objBook As Object        'Excel.Workbook '定义工件簿类
    Dim objsheet As Object      'Excel.Worksheet '定义工作表类
    Dim arrTmp As Variant
    
    On Error GoTo errH
    Me.MousePointer = 11
    Debug.Print Now & vbCrLf
    
    StrInfo = "执行此操作之前请对【疾病编码分类】和【疾病编码目录】的数据进行数据备份。这是一件非常严肃的事情,请先备份再执行本操作。" & vbCrLf & vbCrLf & _
             " 是否继续?"
    If MsgBox(StrInfo, vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
        GoTo errHandle
    End If
        
    gstrSQL = "Select 编码,是否分类 From 疾病编码类别 where 编码 IN ('D','Y','M','S') order by 优先级"
    Set mrs类别 = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
    If sstType.Tab = E_EXCEL Then
    'EXCEL
    '检查文件是否选择
        If Trim(txt(E_EXCEL).Text) = "" Then
            MsgBox "请选择包含表【疾病编码分类】和【疾病编码目录】的EXCEL文件。", vbOKOnly + vbInformation, gstrSysName
            cmd(E_EXCEL).SetFocus
            GoTo errHandle
        End If
        
        If Not objFile.FileExists(Trim(txt(E_EXCEL).Text)) Then
            MsgBox "该文件不存在。文件位置:" & Trim(txt(E_EXCEL).Text), vbOKOnly + vbInformation, gstrSysName
            txt(E_EXCEL).SetFocus
            GoTo errHandle
        End If
        On Error Resume Next
        Set objExcel = CreateObject("Excel.Application") '创建Excel应用类
        If Err.Number <> 0 Then
            MsgBox "请检查本机是否正确安装EXCEL。", vbInformation + vbOKOnly, gstrSysName
            GoTo errHandle
        End If
        Err.Clear: On Error GoTo 0
        
        On Error GoTo errH
        
        arrTmp = Split("0,0", ",")
        objExcel.Visible = False '设置Excel可见
        Set objBook = objExcel.Workbooks.Open(Trim(txt(E_EXCEL).Text)) '打开Excel工作簿
        For i = 1 To objBook.Worksheets.Count
            Set objsheet = objBook.Worksheets(i)  '打开Excel工作表
            If objsheet.Name = "疾病编码目录" Then
                arrTmp(1) = 1
            ElseIf objsheet.Name = "疾病编码分类" Then
                arrTmp(0) = 1
            End If
        Next
        objBook.Close
        objExcel.Quit
        Set objExcel = Nothing
        
        If arrTmp(0) = 1 And arrTmp(1) = 1 Then
            mbytModel = 1
        ElseIf arrTmp(0) = 0 And arrTmp(1) = 1 Then
            If MsgBox("该文件表单名称仅有【疾病编码目录】,没有【疾病编码分类】。继续操作将只导入【疾病编码目录】的数据。" & vbCrLf & vbCrLf & _
                        "是否继续？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                GoTo errHandle
            End If
            mbytModel = 0
        Else
            MsgBox "请检查Exele文件" & Trim(txt(E_EXCEL).Text) & vbCrLf & _
                "该文件表单名称不包含【疾病编码目录】,无法进行下一步操作。", vbInformation + vbOKOnly, gstrSysName
            GoTo errHandle
        End If
        
        If Not InitOLEConn(Trim(txt(E_EXCEL).Text)) Then GoTo errHandle
        '打开记录集
        Set mRsType = New ADODB.Recordset
        Set mrsContent = New ADODB.Recordset
        On Error Resume Next
        If mbytModel = 1 Then
            mRsType.Open "Select [章],[编码范围],[名称] FROM [疾病编码分类$]", mconn, adOpenStatic, adLockOptimistic
            If Err.Number <> 0 Then
               MsgBox "错误号:" & Err.Number & vbCrLf & "错误信息:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrSysName
               Err.Clear
               GoTo errHandle
            End If
        End If
        mrsContent.Open "Select [编码],[附码],[名称] FROM [疾病编码目录$]", mconn, adOpenStatic, adLockOptimistic
        If Err.Number <> 0 Then
           MsgBox "错误号:" & Err.Number & vbCrLf & "错误信息:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrSysName
           Err.Clear
           GoTo errHandle
        End If
        Err.Clear: On Error GoTo 0
        
        On Error GoTo errH
        If Not CheckRS() Then GoTo errHandle
        If Not FuncUpdateRS() Then GoTo errHandle
        Call SaveData(E_EXCEL)
    ElseIf sstType.Tab = E_SCRIPT Then
        '脚步
        strFile = Trim(txt(E_SCRIPT).Text)
        If strFile = "" Then
            MsgBox "请选择导入所需的脚本文件。", vbInformation, gstrSysName
            cmd(E_SCRIPT).SetFocus
            GoTo errHandle
        End If
        If strFile <> "" Then
            If Not FuncCreateRSBySQL(strFile) Then GoTo errHandle
            Call SaveData(E_SCRIPT)
        End If
    End If
    
    Debug.Print Now
errHandle:
    prg.Visible = False
    lblInfo.Caption = ""
    Me.MousePointer = 0
    Exit Sub
    
errH:
    lblInfo.Caption = ""
    prg.Visible = False
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    prg.Visible = False
    Call InitVsItem
End Sub

Public Sub ShowMe(ByVal frmParent As Form)

    Me.Show 1, frmParent
End Sub

Private Sub OpenFile(ByVal strFilter As String, ByVal intIndex As Integer)
    dlgOpenFile.Filter = strFilter
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        txt(intIndex).Text = dlgOpenFile.FileName
    End If
    
End Sub

Public Function CheckRS() As Boolean
    Dim lngRow As Long, lngCol As Long
    Dim reg As New RegExp
    
    '数据格式检查
    On Error GoTo errH
    prg.Visible = True
    If mbytModel = 1 Then
        With mRsType
            lblInfo.Caption = "【疾病编码分类】表单格式检查..."
            reg.Pattern = "^[A-Z]{1}[0-9]{2}-[A-Z]{1}[0-9]{2}$"
            reg.IgnoreCase = False
            '编码范围,名称,类别不能为空
            For lngRow = 1 To .RecordCount
                If Trim(!编码范围 & "") = "" Then
                    MsgBox "EXCEL文件【疾病编码分类】,【编码范围】列值不能为空！" & "所在行:" & lngRow + 1, vbInformation, gstrSysName
                    CheckRS = False
                    Exit Function
                End If
    
                If Trim(!名称 & "") = "" Then
                    MsgBox "EXCEL文件【疾病编码分类】,【名称】列值不能为空！" & "所在行:" & lngRow + 1, vbInformation, gstrSysName
                    CheckRS = False
                    Exit Function
                End If
                '编码范围输入格式有误
                If Not reg.Test(!编码范围 & "") Then
                    MsgBox "EXCEL文件【疾病编码分类】,【编码范围】列、第" & lngRow + 1 & "行:" & vbCrLf & _
                            "编码范围是由疾病编码的前三位（1位大写字母加2位数字）加分隔符""-""组成。示例:A00-B99", vbInformation, gstrSysName
                        CheckRS = False
                    Exit Function
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
    End If
    With mrsContent
        '列名和列顺序检查
        lblInfo.Caption = "【疾病编码目录】表单格式检查..."
        reg.Pattern = "^([A-Z]{1}[0-9]{2}.(([Xx0-9]{0,2}\+?)|([0-9]{0,3}\+?)))|(M[8-9]{1}[0-9]{4}/[01236])|([0-9]{2}.([0-9]{4}|[0-9]{5}))$"
        reg.IgnoreCase = False
        '编码,名称,类别不能为空
        For lngRow = 1 To .RecordCount
            If Trim(!编码 & "") = "" Then
                MsgBox "EXCEL文件【疾病编码目录】,【编码】列值不能为空！" & "所在行:" & lngRow + 1, vbInformation, gstrSysName
                CheckRS = False
                Exit Function
            End If

            If Trim(!名称 & "") = "" Then
                MsgBox "EXCEL文件【疾病编码目录】,【名称】列值不能为空！" & "所在行:" & lngRow + 1, vbInformation, gstrSysName
                CheckRS = False
                Exit Function
            End If
            '编码范围输入格式有误
            If Not reg.Test(!编码 & "") Then
                MsgBox "EXCEL文件【疾病编码目录】,【编码】列、第" & lngRow + 1 & "行:" & vbCrLf & _
                        "编码格式不对请检查。", vbInformation, gstrSysName
                    CheckRS = False
                Exit Function
            End If
            
            prg.value = Int((lngRow / .RecordCount) * 100)
            .MoveNext
        Next
    End With
    prg.Visible = False
    CheckRS = True
    Exit Function
errH:
    prg.Visible = False
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function InitRS(Optional ByVal bytFunc As Byte = 0) As ADODB.Recordset
'功能:构造医嘱记录
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '字段名称|字段类型|字段长度 缺省字段类型 为adVarChar
    
    Select Case bytFunc
    
    Case 0
        strFields = "ID|adBigInt|18,上级ID|adBigInt|18,类别||1,内容||4000,标记|adInteger|1"    '标记 0-Insert Into 行 ;1-Select 行; 2-Select 结尾行
    Case 1
        strFields = "ID|adBigInt|18,分类ID|adBigInt|18,类别||1,编码||100,内容||4000,标记|adInteger|1,是否分类|adInteger|1"  '是否分类 0-不显示分类;1-显示分类
    End Select
    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            If UCase(arrSubFeld(1) & "") = UCase("adVarChar") Then
                FieldType = adVarChar
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adBigInt") Then
                FieldType = adBigInt
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adInteger") Then
                FieldType = adInteger
            Else
                FieldType = adVarChar
            End If
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitRS = rs
End Function

Private Function FuncUpdateRS() As Boolean
    Dim lngRow As Long, lngLevel As Long
    Dim i As Long, lngPos As Long, j As Long, k As Long
    Dim rsTmp As ADODB.Recordset
    
    Dim lngNum As Long
    Dim strType As String, strTypes As String
    Dim strCode As String, strCodeA As String, strcodeB As String
    Dim strGroup As String
    Dim lngGroupId As Long
    Dim arrTmp As Variant
    Dim arrList As Variant
    Dim arrLevel As Variant
    
    On Error GoTo errH
    lblInfo.Caption = "正在组织【疾病编码分类】的数据..."
    prg.Visible = True
    If mbytModel = 0 Or mbytModel = 2 Then
        gstrSQL = "Select a.Id, a.上级id, 序号, Level, a.名称, a.编码范围, a.类别, a.是否病人, 0 As 操作" & vbNewLine & _
                "From 疾病编码分类 A, 疾病编码类别 B" & vbNewLine & _
                "Where a.类别 = b.编码 And (a.撤档时间 Is Null Or Trunc(a.撤档时间) = To_Date('3000-01-01', 'YYYY-MM-dd')) And" & vbNewLine & _
                "      a.类别 In ('D', 'Y', 'M', 'S') And b.是否分类 = 1" & vbNewLine & _
                "Start With a.上级id Is Null" & vbNewLine & _
                "Connect By Prior ID = a.上级id"
        Call zldatabase.OpenRecordset(mRsType, gstrSQL, Me.Caption, adOpenStatic, adLockOptimistic)
        Set mRsType = zldatabase.CopyNewRec(mRsType) '复制便于后面更新字段值
    End If
    
    If mbytModel = 1 Then
        Set mRsType = zldatabase.CopyNewRec(mRsType, , , Array("ID", adBigInt, 18, Empty, "上级ID", adBigInt, 18, Empty, "序号", adInteger, 6, Empty, "编码A", adVarChar, 60, Empty, _
                        "编码B", adVarChar, 60, Empty, "类别", adVarChar, 1, Empty, "是否病人", adInteger, 1, Empty, "操作", adInteger, 1, Empty, "Level", adInteger, 1, Empty))
        '追加字段
        strTypes = ""
        With mRsType
            .Filter = ""
            For lngRow = 1 To .RecordCount
                '名称编码
                !ID = lngRow
                !名称 = FuncGetStr(!名称)
                !编码范围 = UCase(FuncGetStr(!编码范围))
                
                !编码A = UCase(Split(!编码范围 & "", "-")(0))
                !编码B = UCase(Split(!编码范围 & "", "-")(1))
                !是否病人 = 1 '默认为1; 0-疾病疗效只能是其他
                !操作 = 1
                !Level = 1
                !类别 = FuncCheckType(!编码A & "")
                If InStr("," & strTypes & ",", "," & !类别 & ",") = 0 Then
                    strTypes = strTypes & "," & !类别
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
            strTypes = Mid(strTypes, 2)
            
            If strTypes <> "" Then
                arrTmp = Split(strTypes, ",")
                For i = LBound(arrTmp) To UBound(arrTmp)
                    mRsType.Filter = "类别='" & arrTmp(i) & "'"
                    lngNum = 0
                    For j = 1 To mRsType.RecordCount
                        lngNum = lngNum + 1
                        mRsType!序号 = lngNum
                        mRsType.MoveNext
                    Next
                Next
            End If
            '根据编码范围查找上级ID
            .Filter = ""
            For lngRow = 1 To .RecordCount
                lngPos = .AbsolutePosition
                If Trim(!章 & "") <> "" And strGroup <> Trim(!章 & "") Then
                    strGroup = Trim(!章 & "")
                    !上级id = 0
                    !Level = 1
                Else
                    strCodeA = !编码A
                    strcodeB = !编码B
                    Do While Not .BOF
                        .MovePrevious
                        If !上级id = 0 Then
                            lngGroupId = !ID
                            lngLevel = 1
                            .AbsolutePosition = lngPos
                            !上级id = lngGroupId
                            !Level = (lngLevel + 1)
                            Exit Do
                        End If
                        If strCodeA >= !编码A And strcodeB <= !编码B Then
                            lngGroupId = !ID
                            lngLevel = !Level
                            .AbsolutePosition = lngPos
                            !上级id = lngGroupId
                            !Level = (lngLevel + 1)
                            Exit Do
                        End If
                    Loop
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
    End If
    
    If mbytModel = 0 Or mbytModel = 1 Then
        lblInfo.Caption = "正在组织【疾病编码目录】的数据..."
        strCode = ""
        Set mrsContent = zldatabase.CopyNewRec(mrsContent, , , Array("ID", adBigInt, 18, Empty, "序号", adInteger, 10, Empty, "分类ID", adBigInt, 18, Empty, "是否分类", adInteger, 3, _
                        Empty, "类别", adVarChar, 1, Empty))
        strTypes = ""
        With mrsContent
            If .RecordCount > 0 Then .MoveFirst
            For lngRow = 1 To .RecordCount
                !ID = lngRow
                !序号 = 1
                !分类id = 0
                !名称 = FuncGetStr(!名称)
                !编码 = UCase(FuncGetStr(!编码))
                !类别 = FuncCheckType(!编码 & "")
                If strCode = !类别 & "_" & !编码 Then
                    !序号 = lngNum + 1
                End If
                '记录上一个编码及序号
                strCode = !类别 & "_" & !编码
                lngNum = Val(!序号 & "")
                If InStr("," & strTypes & ",", "," & !类别 & ",") = 0 Then
                    strTypes = strTypes & "," & !类别
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
        strTypes = strTypes & ","
    End If
    If mbytModel = 2 Then
        With mrsContent
            .Filter = "标记>0"
            For i = 1 To .RecordCount
                !分类id = 0
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
    End If
    lblInfo.Caption = "正在更新【疾病编码目录】的【分类ID】..."
    mrs类别.Filter = ""
    Do While Not mrs类别.EOF
        If Val(mrs类别!是否分类 & "") = 0 Then
            gstrSQL = "Select ID" & vbNewLine & _
                    "From 疾病编码分类 A" & vbNewLine & _
                    "Where a.类别 = [1] And (a.撤档时间 Is Null Or Trunc(a.撤档时间) = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mrs类别!编码)
            If Not rsTmp.EOF Then lngNum = rsTmp!ID
            mrsContent.Filter = "类别 ='" & mrs类别!编码 & "'"
            For j = 1 To mrsContent.RecordCount
                mrsContent!分类id = lngNum
                mrsContent!是否分类 = 1   '不新增分类,沿用之前分类
                mrsContent.MoveNext
            Next
        Else
            mRsType.Filter = "类别 = '" & mrs类别!编码 & "'"
            mRsType.Sort = "Level desc"
            For i = 1 To mRsType.RecordCount
                arrList = Split(Trim(mRsType!编码范围 & ""), ",")
                'A15.0-A15.3,A16.0-A16.2
                For j = LBound(arrList) To UBound(arrList)
                    arrTmp = Split(Trim(arrList(j)), "-")
                    If UBound(arrTmp) = 1 Then
                        mrsContent.Filter = "编码 >= '" & arrTmp(0) & "' And 编码 <= '" & arrTmp(1) & "' And 分类ID = 0 And  类别 ='" & mRsType!类别 & "'"
                        For k = 1 To mrsContent.RecordCount
                            mrsContent!分类id = mRsType!ID
                            mrsContent.MoveNext
                        Next
                        
                        mrsContent.Filter = "类别 ='" & mRsType!类别 & "' And 分类ID = 0 And 编码 like '" & arrTmp(1) & "%'"
                        For k = 1 To mrsContent.RecordCount
                            mrsContent!分类id = mRsType!ID
                            mrsContent.MoveNext
                        Next
                        prg.value = Int((i / mRsType.RecordCount) * 100)
                    ElseIf UBound(arrTmp) = 0 Then
                        mrsContent.Filter = "类别 ='" & mRsType!类别 & "' And 分类ID = 0 And 编码 like '" & arrTmp(0) & "%'"
                        For k = 1 To mrsContent.RecordCount
                            mrsContent!分类id = mRsType!ID
                            mrsContent.MoveNext
                        Next
                    End If
                Next
                mRsType.MoveNext
            Next
        End If
        mrs类别.MoveNext
    Loop
    If mbytModel = 0 Or mbytModel = 1 Then
        mrsContent.Filter = "分类ID = 0"
    ElseIf mbytModel = 2 Then
        mrsContent.Filter = "标记>0 And 分类ID = 0"
    End If
    If mrsContent.RecordCount > 0 Then
        strCode = ""
        For i = 1 To mrsContent.RecordCount
            strCode = strCode & "," & mrsContent!编码
            If i > 10 Then strCode = strCode & "...": Exit For
            mrsContent.MoveNext
        Next
        strCode = Mid(strCode, 2)
        If MsgBox("存在" & mrsContent.RecordCount & "行【疾病编码目录】的分类ID无法确认,对应【编码】如下:" & strCode & vbCrLf & _
                "是否继续？" & vbCrLf & _
                "选择【是】会将无法确认【分类ID】的项目添加到指定分类【其他】中并继续下一步操作。" & vbCrLf & _
                "选择【否】会终止本次操作。请检查【疾病编码分类】的【编码范围】列的值能否将上述【疾病编码目录】的编码值包含。", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
            mrsContent.MoveFirst

            For i = 1 To mrsContent.RecordCount
                mRsType.Filter = "类别='" & mrsContent!类别 & "' And 名称='其他' And 上级ID = 0"
                If mRsType.RecordCount > 0 Then
                    lngGroupId = mRsType!ID
                Else
                    If mbytModel = 0 Or mbytModel = 2 Then
                        mRsType.Filter = "类别='" & mrsContent!类别 & "'"
                        mRsType.Sort = "序号 desc"
                        If Not mRsType.EOF Then
                            lngNum = Val(mRsType!序号 & "") + 1
                        Else
                            lngNum = (mRsType.RecordCount + 1)
                        End If
                    Else
                        mRsType.Filter = "类别='" & mrsContent!类别 & "'"
                        lngNum = (mRsType.RecordCount + 1)
                    End If
                    mRsType.Filter = ""
                    lngGroupId = mRsType.RecordCount + 1
                    mRsType.AddNew
                    mRsType!ID = lngGroupId
                    mRsType!上级id = 0
                    mRsType!序号 = lngNum
                    mRsType!名称 = "其他"
                    mRsType!类别 = mrsContent!类别
                    mRsType!是否病人 = 1
                    mRsType!操作 = 1 '1-新增
                    mRsType.Update
                End If
                mrsContent!分类id = lngGroupId
                mrsContent.MoveNext
            Next
        Else
            prg.Visible = False
            Me.MousePointer = 0
            Exit Function
        End If
    End If
    mRsType.Filter = ""
    mRsType.Sort = ""
    prg.Visible = False
    FuncUpdateRS = True
    Exit Function
errH:
    Resume
    prg.Visible = False
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Private Function SaveData(ByVal bytFunc As Byte) As Boolean
'功能:
'参数:bytFunc=0 Excel导入;=1 脚步文件导入
    Dim colType As New Collection
    Dim colSQL As New Collection
    Dim lngId As Long, i As Long
    Dim strValue As String
    Dim strTitle As String
    Dim strTemp As String
    Dim blnOver As Boolean
    Dim datCurr As Date
    Dim arrTmp As Variant
    Dim strDate As String, strType As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lngMin As Long, lngMax As Long
    
    On Error GoTo errH
    'GET ID
    prg.Visible = True
    lblInfo.Caption = "正在生成【疾病编码分类】的【ID】、【上级ID】..."
    If mbytModel < 3 Then
        mRsType.Filter = "操作 =1 "
    Else
        mRsType.Filter = "标记>0"
    End If
    '疾病编码分类未新增的情况下
    lngId = FuncGetNo("疾病编码分类", mRsType.RecordCount)  '提前获取ID
    For i = 1 To mRsType.RecordCount
        'Debug.Print mRsType!ID & "_" & mRsType!上级ID & "_" & mRsType!名称
        colType.Add lngId, "_" & mRsType!ID: lngId = lngId + 1
        If mbytModel = 1 Or mbytModel = 3 Then
            If Not InStr("," & strType & ",", "," & UCase(mRsType!类别 & "") & ",") > 0 Then
                strType = strType & "," & UCase(mRsType!类别 & "")
            End If
        End If
        prg.value = Int((i / mRsType.RecordCount) * 100)
        mRsType.MoveNext
    Next
    
    If mbytModel = 0 Or mbytModel = 2 Then
    '沿用旧的分类ID，不生成新的ID
        mRsType.Filter = "操作 = 0 "
        For i = 1 To mRsType.RecordCount
            colType.Add Val(mRsType!ID & ""), "_" & mRsType!ID
            prg.value = Int((i / mRsType.RecordCount) * 100)
            mRsType.MoveNext
        Next
    End If

    If mbytModel < 3 Then
        mRsType.Filter = "操作 =1 "
    Else
        mRsType.Filter = "标记>0"
    End If

    For i = 1 To mRsType.RecordCount
        mRsType!ID = colType("_" & mRsType!ID)
        If Val(mRsType!上级id & "") <> 0 Then
            mRsType!上级id = colType("_" & mRsType!上级id)
        End If
        prg.value = Int((i / mRsType.RecordCount) * 100)
        mRsType.MoveNext
    Next
    
    datCurr = zldatabase.Currentdate
    strDate = Format(datCurr, "yyyy-MM-dd HH:mm:ss")
    strType = Mid(strType, 2)
    If strType <> "" Then
        arrTmp = Split(strType, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            gstrSQL = "Update 疾病编码分类 " & vbNewLine & _
                        " Set 撤档时间 = " & "To_Date('" & DateAdd("n", -1, datCurr) & "','YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
                        " Where 类别 = '" & arrTmp(i) & "' And (撤档时间 Is Null Or Trunc(撤档时间) = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            colSQL.Add gstrSQL
        Next
    End If
    
    strType = "": lblInfo.Caption = "正在生成【疾病编码目录】的【ID】..."
    If bytFunc = 0 Then
        mrsContent.Filter = ""
    Else
        mrsContent.Filter = "标记>0"
    End If
    lngId = FuncGetNo("疾病编码目录", mrsContent.RecordCount)  '提前获取ID
    For i = 1 To mrsContent.RecordCount
        mrsContent!ID = lngId: lngId = lngId + 1
        '没有分类的项目已经提前处理:M-肿瘤形态学编码
        If Val(mrsContent!是否分类 & "") = 0 Then
            mrsContent!分类id = colType("_" & mrsContent!分类id)
        End If
        If Not InStr("," & strType & ",", "," & UCase(mrsContent!类别 & "") & ",") > 0 Then
            strType = strType & "," & UCase(mrsContent!类别 & "")
        End If
        prg.value = Int((i / mrsContent.RecordCount) * 100)
        mrsContent.MoveNext
    Next
    strType = Mid(strType, 2)
    If strType <> "" Then
        arrTmp = Split(strType, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            gstrSQL = "Update 疾病编码目录 " & vbNewLine & _
                      "Set 撤档时间 = " & "To_Date('" & DateAdd("n", -1, datCurr) & "','YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
                      "Where 类别 = '" & arrTmp(i) & "' And (撤档时间 Is Null Or Trunc(撤档时间) = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            colSQL.Add gstrSQL
        Next
    End If
    
    strDate = "To_Date('" & datCurr & "','YYYY-MM-DD HH24:MI:SS')"
    If mbytModel < 3 Then
        '构造疾病编码分类的SQL
        strTitle = "Insert Into 疾病编码分类(ID, 上级id, 序号, 名称, 简码, 类别, 编码范围, 是否病人, 建档时间) " & vbCrLf
        mRsType.Filter = "操作 =1"
        lblInfo.Caption = "正在生成【疾病编码分类】的SQL语句..."
        strTemp = "": strValue = ""
        With mRsType
            For i = 1 To .RecordCount
                strTemp = "Select " & !ID & "," & IIF(Val(!上级id & "") = 0, "Null", !上级id) & "," & !序号 & _
                    ",'" & Trim(Replace(!名称, "'", "''")) & "','" & _
                    Mid(zlcommfun.SpellCode(Trim(Replace(!名称, "'", "''")) & "※0"), 1, 20) & "','" & Trim(!类别) & "','" & Trim(Replace(!编码范围 & "", "'", "''")) & "'," & _
                    Val(!是否病人 & "") & "," & strDate & " From Dual UNION ALL" & vbCrLf
                If Len(strTitle & strValue & strTemp) > 100000 Then
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                    colSQL.Add strTitle & strValue
                    strValue = strTemp
                    blnOver = True
                Else
                    blnOver = False
                    strValue = strValue & strTemp
                End If
                .MoveNext
                If .EOF Then
                    If Not blnOver Then
                        strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                        colSQL.Add strTitle & strValue
                        Exit For
                    End If
                End If
                prg.value = Int((i / .RecordCount) * 100)
            Next
        End With
    Else
        '构造疾病编码分类的SQL
        mRsType.Filter = "标记=0"
        strTitle = mRsType!内容 & vbCrLf
        mRsType.Filter = "标记>0"
        lblInfo.Caption = "正在生成【疾病编码分类】的SQL语句..."
        strValue = "": lngMin = 0: lngMax = 0
        With mRsType
            For i = 1 To .RecordCount
                If i = 1 Then lngMin = Val(!ID & "")
                If i = .RecordCount Then lngMax = Val(!ID & "")
                strValue = strValue & "Select " & !ID & "," & IIF(Val(!上级id & "") = 0, "Null", !上级id) & "," & !内容 & vbCrLf
                If !标记 = 2 Then
                    colSQL.Add strTitle & strValue
                    strValue = ""
                End If
                .MoveNext
                prg.value = Int((i / .RecordCount) * 100)
            Next
        End With
        If lngMin <> lngMax Then
            colSQL.Add "Update 疾病编码分类 Set 建档时间 = " & strDate & " Where ID Between " & lngMin & " And " & lngMax
        End If
    End If
    
    If bytFunc = 0 Then
        'Update 疾病编码目录 Set 五笔码 = ZLTOOLS.zlWbCode(名称, 20);
        '构造疾病编码目录的SQL
        strTitle = "Insert Into 疾病编码目录 (ID, 分类id, 编码, 序号, 附码, 名称, 简码, 五笔码, 类别, 建档时间)" & vbCrLf
        
        lblInfo.Caption = "正在生成【疾病编码目录】的SQL语句..."
        strTemp = "": strValue = ""
        With mrsContent
            .Filter = ""
            For i = 1 To .RecordCount
                strTemp = "Select " & !ID & "," & !分类id & ",'" & !编码 & "'," & !序号 & ",'" & !附码 & "','" & !名称 & "','" & Mid(zlcommfun.SpellCode(Trim(Replace(!名称, "'", "''")) & "※0"), 1, 20) & "','" & _
                            Mid(zlcommfun.SpellCode(Trim(Replace(!名称, "'", "''")) & "※1"), 1, 20) & "','" & !类别 & "'," & strDate & " From Dual UNION ALL" & vbCrLf
                If Len(strTitle & strValue & strTemp) > 100000 Then
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                    colSQL.Add strTitle & strValue
                    strValue = strTemp
                    blnOver = True
                Else
                    blnOver = False
                    strValue = strValue & strTemp
                End If
                .MoveNext
                If .EOF Then
                    If Not blnOver Then
                        strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                        colSQL.Add strTitle & strValue
                        Exit For
                    End If
                End If
                prg.value = Int((i / .RecordCount) * 100)
            Next
         End With
    Else
        'Update 疾病编码目录 Set 五笔码 = ZLTOOLS.zlWbCode(名称, 20);
        '构造疾病编码目录的SQL
        mrsContent.Filter = "标记=0"
        strTitle = mrsContent!内容 & vbCrLf
        
        lblInfo.Caption = "正在生成【疾病编码目录】的SQL语句..."
        strValue = "": lngMin = 0: lngMax = 0
        With mrsContent
            .Filter = "标记>0"
            For i = 1 To .RecordCount
                If i = 1 Then lngMin = Val(!ID & "")
                If i = .RecordCount Then lngMax = Val(!ID & "")
                strValue = strValue & "Select " & !ID & "," & !分类id & "," & !内容 & vbCrLf
                If !标记 = 2 Then
                    colSQL.Add strTitle & strValue
                    strValue = ""
                End If
                 .MoveNext
                prg.value = Int((i / .RecordCount) * 100)
            Next
            If lngMin <> lngMax Then
                colSQL.Add "Update 疾病编码目录 Set 建档时间 = " & strDate & " Where ID Between " & lngMin & " And " & lngMax
            End If
         End With
    End If

    '疾病编码目录 编码和名称 都相同的情况下不插入新数据,保留原数据。如果原数据的分类已经停用则修改分类ID
    gstrSQL = "Zl_疾病编码目录_Redo(" & strDate & ",To_Date('" & DateAdd("n", -1, datCurr) & "','YYYY-MM-DD HH24:MI:SS'))"
    lblInfo.Caption = "正在提交【疾病编码分类】、【疾病编码目录】的数据..."
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colSQL.Count
        Call zldatabase.OpenRecordset(rsTmp, CStr(colSQL(i)), Me.Caption)
        WriteLog vbCrLf & colSQL(i)
        prg.value = (Int(i / colSQL.Count) * 100)
    Next
    
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    WriteLog vbCrLf & gstrSQL
    gcnOracle.CommitTrans: blnTrans = False

        
    MsgBox "导入成功!", vbInformation, Me.Caption
    
    SaveData = True
    lblInfo.Caption = ""
    Me.MousePointer = 0
    prg.Visible = False
    Exit Function
errH:
    prg.Visible = False
    lblInfo.Caption = ""
    If blnTrans Then gcnOracle.RollbackTrans
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function InitOLEConn(ByVal strFilePath As String) As Boolean
    Dim strConnect As String
    Dim objFile As New FileSystemObject
    
    On Error GoTo errH
    
    If mconn Is Nothing Then
        Set mconn = New ADODB.Connection
        If UCase(objFile.GetExtensionName(strFilePath)) = "XLS" Then
            strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath & ";Extended Properties=""Excel 12.0;HDR=YES"""
        Else
            strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath & ";Extended Properties=""Excel 12.0;HDR=YES"""
        End If
        mconn.ConnectionString = strConnect
    End If
    If mconn.State = adStateClosed Then mconn.Open
    InitOLEConn = True
    Exit Function
errH:
    If Err.Number = -2147467259 Then
        MsgBox "请检查EXCEL文件是否已经打开", vbInformation, gstrSysName
    Else
        MsgBox Err.Description, vbInformation, gstrSysName
    End If
    Set mconn = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mconn Is Nothing Then
        If mconn.State = adStateOpen Then mconn.Close
        Set mconn = Nothing
    End If
    Set mrsContent = Nothing
    Set mRsType = Nothing
End Sub


Private Function FuncGetNo(ByVal strTable As String, ByVal lngCount As Long) As Long
'功能:获取序列,并修正序列
    Dim rsTemp As New ADODB.Recordset
    Dim lngMax As Long
    Dim lngCurr As Long
    Dim strOwner As String
    
    On Error GoTo errH
    gstrSQL = "Select Sequence_Owner From All_Sequences Where Sequence_Name = Upper('" & strTable & "_ID')"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then strOwner = rsTemp!Sequence_Owner & ""

    '序列修正
    gstrSQL = "Select Max(ID) as ID From " & strTable
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then lngMax = rsTemp!ID
    gstrSQL = "Select " & strOwner & "." & strTable & "_ID.Nextval As CurrID From Dual"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then lngCurr = rsTemp!CurrID
    If Abs(lngMax - lngCurr) > 1 Then
        '--修正成反向增量
        gstrSQL = "Alter Sequence " & strOwner & "." & strTable & "_ID Increment By " & (lngMax - lngCurr)
        gcnOracle.Execute gstrSQL
        ' --移动一次增量
        gstrSQL = "Select " & strOwner & "." & strTable & "_ID.Nextval From Dual"
        gcnOracle.Execute gstrSQL
        '--恢复原始增量
        gstrSQL = "Alter Sequence " & strOwner & "." & strTable & "_ID Increment By 1"
        gcnOracle.Execute gstrSQL
    End If
    gstrSQL = "Select " & strOwner & "." & strTable & "_ID.Nextval AS NO From Dual Connect By Rownum <= [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngCount)
    If Not rsTemp.EOF Then FuncGetNo = rsTemp!NO
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncCheckType(ByVal strCode As String) As String
'功能:根据编码确定类别
'
    Dim strTest As String
    Dim re As New RegExp
    
    '肿瘤形态学编码
    On Error Resume Next
    re.Pattern = "^(M[8-9]{1}[0-9]{4}/[01236])$"
    If re.Test(strCode) Then
        FuncCheckType = "M": Exit Function
    End If
    Err.Clear: On Error GoTo 0
    
    'Y-损伤中毒的外部原因(V01～Y98)
    strTest = Left(strCode, 3)
    If (strTest >= "V01" And strTest <= "Y98") Then
        FuncCheckType = "Y": Exit Function
    End If
    
    'ICD-10排除损伤中毒的外部原因
    If (strTest >= "A00" And strTest <= "Z99") And Not (strTest >= "V01" And strTest <= "Y98") Then
        FuncCheckType = "D": Exit Function
    End If

    'ICD-9-CM3手术编码
    strTest = Left(strCode, 2)
    If strTest >= "00" And strTest <= 99 Then
        FuncCheckType = "S": Exit Function
    End If
End Function

Private Function FuncCreateRSBySQL(ByVal strFile As String) As Boolean

    Dim StrInfo As String, strTXT As String
    Dim objFile As New FileSystemObject
    Dim objStream As TextStream
    Dim bytType As Byte '0-疾病编码分类;1-疾病编码目录
    Dim arrItem As Variant
    Dim lngNum As Long
    Dim j As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    
    Set mRsType = InitRS(0)
    Set mrsContent = InitRS(1)
    If Not objFile.FileExists(strFile) Then
        MsgBox "当前文件:" & strFile & "不存在。", vbInformation, gstrSysName
        Exit Function
    End If
            
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    Do While Not objStream.AtEndOfStream
        strTXT = Trim(objStream.ReadLine)
        If InStr(UCase(strTXT), UCase("Insert Into")) > 0 And InStr(strTXT, "疾病编码目录") > 0 Then
            bytType = 1
            mrsContent.Filter = "标记=0"
            If mrsContent.RecordCount = 0 Then
                mrsContent.AddNew
                mrsContent!标记 = 0
                mrsContent!内容 = Trim(strTXT)
            End If
        ElseIf InStr(UCase(strTXT), UCase("Insert Into")) > 0 And InStr(strTXT, "疾病编码分类") > 0 Then
            bytType = 0
            mRsType.Filter = "标记=0"
            If mRsType.RecordCount = 0 Then
                mRsType.AddNew
                mRsType!标记 = 0
                mRsType!内容 = Trim(strTXT)
            End If
        ElseIf InStr(UCase(strTXT), UCase("Select")) > 0 And InStr(UCase(strTXT), UCase("From Dual UNION ALL")) > 0 Then
            arrItem = Split(strTXT, ",")
            If bytType = 0 Then
                mRsType.AddNew
                mRsType!ID = Val(Replace(UCase(arrItem(0)), UCase("Select "), ""))
                mRsType!上级id = IIF(UCase(arrItem(1)) = "NULL", 0, Val(arrItem(1)))
                mRsType!类别 = Replace(arrItem(2), "'", "")
                mRsType!内容 = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mRsType!标记 = 1
            Else
                mrsContent.AddNew
                mrsContent!ID = Val(Replace(UCase(arrItem(0)), UCase("Select"), ""))
                mrsContent!分类id = Val(arrItem(1))
                mrsContent!类别 = Replace(arrItem(2), "'", "")
                mrsContent!编码 = Replace(arrItem(3), "'", "")
                mrsContent!内容 = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mrsContent!标记 = 1
            End If
        ElseIf InStr(UCase(strTXT), UCase("Select")) > 0 And InStr(UCase(strTXT), UCase("From Dual")) > 0 And InStr(UCase(strTXT), UCase("UNION ALL")) = 0 Then
            If Right(strTXT, 1) = ";" Then strTXT = Left(strTXT, Len(strTXT) - 1)
            arrItem = Split(strTXT, ",")
            If bytType = 0 Then
                mRsType.AddNew
                mRsType!ID = Val(Replace(UCase(arrItem(0)), UCase("Select"), ""))
                mRsType!上级id = IIF(UCase(arrItem(1)) = "NULL", 0, Val(arrItem(1)))
                mRsType!类别 = Replace(arrItem(2), "'", "")
                mRsType!内容 = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mRsType!标记 = 2
                mRsType.UpdateBatch
            Else
                mrsContent.AddNew
                mrsContent!ID = Val(Replace(UCase(arrItem(0)), UCase("Select"), ""))
                mrsContent!分类id = Val(arrItem(1))
                mrsContent!类别 = Replace(arrItem(2), "'", "")
                mrsContent!编码 = Replace(arrItem(3), "'", "")
                mrsContent!内容 = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mrsContent!标记 = 2
                mrsContent.UpdateBatch
            End If
        End If
    Loop
    mRsType.Filter = ""
    If mRsType.RecordCount = 0 Then
        mbytModel = 2
        If Not FuncUpdateRS() Then Exit Function
    Else
        mbytModel = 3
    End If
    
    FuncCreateRSBySQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

