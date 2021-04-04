VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHomePage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "主页设置"
   ClientHeight    =   4005
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5880
   Icon            =   "frmHomePage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3240
      TabIndex        =   15
      Top             =   3585
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4470
      TabIndex        =   16
      Top             =   3585
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   75
      TabIndex        =   17
      Top             =   3570
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   3510
      Left            =   30
      TabIndex        =   18
      Top             =   15
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   6191
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "主页信息"
      TabPicture(0)   =   "frmHomePage.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdTest"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cbo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "主页背景"
      TabPicture(1)   =   "frmHomePage.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSize(1)"
      Tab(1).Control(1)=   "UsrPicture(1)"
      Tab(1).Control(2)=   "cmdOpen(1)"
      Tab(1).Control(3)=   "cmdClear(1)"
      Tab(1).Control(4)=   "cmdPos(1)"
      Tab(1).Control(5)=   "cmbMode"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "宣传标语"
      TabPicture(2)   =   "frmHomePage.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPos(2)"
      Tab(2).Control(1)=   "cmdClear(2)"
      Tab(2).Control(2)=   "cmdOpen(2)"
      Tab(2).Control(3)=   "UsrPicture(2)"
      Tab(2).Control(4)=   "lblSize(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "医院标志"
      TabPicture(3)   =   "frmHomePage.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdPos(3)"
      Tab(3).Control(1)=   "cmdClear(3)"
      Tab(3).Control(2)=   "cmdOpen(3)"
      Tab(3).Control(3)=   "UsrPicture(3)"
      Tab(3).Control(4)=   "lblSize(3)"
      Tab(3).ControlCount=   5
      Begin VB.ComboBox cmbMode 
         Height          =   300
         Left            =   -70980
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2010
         Width           =   1110
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3015
         Width           =   2610
      End
      Begin VB.CommandButton cmdTest 
         Cancel          =   -1  'True
         Caption         =   "试听(&H)"
         Height          =   350
         Left            =   3960
         TabIndex        =   5
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CommandButton cmdPos 
         Caption         =   "位置(&P)"
         Height          =   350
         Index           =   3
         Left            =   -70980
         TabIndex        =   14
         Top             =   1605
         Width           =   1100
      End
      Begin VB.CommandButton cmdPos 
         Caption         =   "位置(&P)"
         Height          =   350
         Index           =   2
         Left            =   -70980
         TabIndex        =   11
         Top             =   1590
         Width           =   1100
      End
      Begin VB.CommandButton cmdPos 
         Caption         =   "位置(&P)"
         Height          =   350
         Index           =   1
         Left            =   -70980
         TabIndex        =   8
         Top             =   1515
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除(&L)"
         Height          =   350
         Index           =   3
         Left            =   -70980
         TabIndex        =   13
         Top             =   1020
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除(&L)"
         Height          =   350
         Index           =   2
         Left            =   -70980
         TabIndex        =   10
         Top             =   1005
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除(&L)"
         Height          =   350
         Index           =   1
         Left            =   -70980
         TabIndex        =   7
         Top             =   1020
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "图片(&B)"
         Height          =   350
         Index           =   3
         Left            =   -70980
         TabIndex        =   12
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "图片(&B)"
         Height          =   350
         Index           =   2
         Left            =   -70980
         TabIndex        =   9
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "图片(&B)"
         Height          =   350
         Index           =   1
         Left            =   -70980
         TabIndex        =   6
         Top             =   600
         Width           =   1100
      End
      Begin VB.Frame Frame2 
         Caption         =   "主页图片"
         Height          =   2535
         Left            =   150
         TabIndex        =   19
         Top             =   405
         Width           =   5085
         Begin VB.CommandButton cmdPos 
            Caption         =   "位置(&P)"
            Height          =   350
            Index           =   0
            Left            =   3750
            TabIndex        =   2
            Top             =   1290
            Width           =   1100
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "清除(&L)"
            Height          =   350
            Index           =   0
            Left            =   3750
            TabIndex        =   1
            Top             =   630
            Width           =   1100
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "图片(&A)"
            Height          =   350
            Index           =   0
            Left            =   3750
            TabIndex        =   0
            Top             =   225
            Width           =   1100
         End
         Begin zl9NewQuery.ctlPicture UsrPicture 
            Height          =   2205
            Index           =   0
            Left            =   105
            TabIndex        =   20
            Top             =   240
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   3889
         End
         Begin VB.Label lblSize 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "800 X 600"
            Height          =   180
            Index           =   0
            Left            =   3675
            TabIndex        =   21
            Top             =   2175
            Width           =   1260
         End
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   2565
         Index           =   1
         Left            =   -74805
         TabIndex        =   22
         Top             =   600
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   4524
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   2565
         Index           =   2
         Left            =   -74805
         TabIndex        =   24
         Top             =   600
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   4524
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   2565
         Index           =   3
         Left            =   -74805
         TabIndex        =   26
         Top             =   600
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   4524
      End
      Begin VB.Label Label3 
         Caption         =   "背景音乐(&M)"
         Height          =   270
         Left            =   135
         TabIndex        =   3
         Top             =   3075
         Width           =   1185
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   3
         Left            =   -71175
         TabIndex        =   27
         Top             =   2580
         Width           =   1500
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   2
         Left            =   -71175
         TabIndex        =   25
         Top             =   2625
         Width           =   1470
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   1
         Left            =   -71235
         TabIndex        =   23
         Top             =   2745
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmHomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarSvrPicRange As String           '保存增加图片的范围
Private mvarSvrPicType As String            '保存增加图片的类型
Private mstrHomeCode As String

Private Sub cbo_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click(Index As Integer)
    UsrPicture(Index).Tag = ""
    UsrPicture(Index).Cls
    cmdOK.Tag = "1"
End Sub

Private Sub cmdOK_Click()
    Dim strSQL(1 To 4) As String
    Dim i As Long
    
    If cmdOK.Tag = "1" Then
        On Error GoTo errHand
        gcnOracle.BeginTrans
        strSQL(1) = "zl_咨询页面目录_delete(0)"
        strSQL(2) = "zl_咨询页面目录_insert(0,'主页',1,0," & IIf(Val(UsrPicture(2).Tag) = 0, "NULL", Val(UsrPicture(2).Tag)) & "," & IIf(Val(UsrPicture(1).Tag) = 0, "NULL", Val(UsrPicture(1).Tag)) & "," & cbo.ItemData(cbo.ListIndex) & ",NULL,1,'" & mstrHomeCode & "','ZY')"
        strSQL(3) = "zl_咨询段落目录_insert(0,1,'',NULL,0,0,'宋体;12;0;0;0',0,NULL,0," & IIf(Val(UsrPicture(0).Tag) = 0, "NULL", Val(UsrPicture(0).Tag)) & ",0)"
        strSQL(4) = "zl_咨询段落目录_insert(0,2,'',NULL,0,0,'宋体;12;0;0;0',0,NULL,0," & IIf(Val(UsrPicture(3).Tag) = 0, "NULL", Val(UsrPicture(3).Tag)) & ",0)"
        For i = 1 To 4
            'gcnOracle.Execute strSQL(i), , adCmdStoredProc
            Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
        Next
        gcnOracle.CommitTrans
    End If
    Unload Me
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub cmdOpen_Click(Index As Integer)
    Dim lngKey As Long
    Dim strFilter As String
    Dim strTitle As String
    
    Select Case Index
    Case 0
        strFilter = "9;0;1;2;3;4"
        strTitle = "添加主页内容"
    Case 1
        strFilter = "4;0;1;2;3;9"
        strTitle = "添加主页背景"
    Case 2
        strFilter = "1;0;2;3;4;9"
        strTitle = "添加主页宣传标语"
    Case 3
        strFilter = "0;1;2;3;4;9"
        strTitle = "添加医院标志图片"
    End Select
    If frmPicSelect.OpenPictureBox(Me, strTitle, strFilter, lngKey, mvarSvrPicRange, mvarSvrPicType) Then
        '更新图片显示
        UsrPicture(Index).Tag = lngKey
        Call ShowPicture(lngKey, Index)
        cmdOK.Tag = "1"
    End If
End Sub

Private Sub cmdPos_Click(Index As Integer)
    Select Case Index
    Case 0
        Call frmPosSample.ShowPageSample("主页图片")
    Case 1
        Call frmPosSample.ShowPageSample("主页背景")
    Case 2
        Call frmPosSample.ShowPageSample("宣传标语")
    Case 3
        Call frmPosSample.ShowPageSample("标志图片")
    End Select
End Sub

Private Sub cmdTest_Click()
    Dim vFileData As New FileSystemObject
    Dim strFile As String
    
    Call MusicClose
    
    
    If cbo.ListIndex < 0 Then Exit Sub
    If cbo.ItemData(cbo.ListIndex) <= 0 Then Exit Sub
    
    '1.检查图形目录是否存在
    On Error Resume Next
    vFileData.CreateFolder App.Path & "\图形"
    
    '2.检查本系统中可能使用到的图片
    gstrSQL = "select 序号,类型,名称,修改日期 from 咨询图片元素 where 序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(cbo.ItemData(cbo.ListIndex)))
    If gRs.BOF Then Exit Sub
    
    strFile = IIf(IsNull(gRs!名称), "", gRs!名称)
    If strFile <> "" Then Call CheckFileNew(strFile, IIf(IsNull(gRs!类型), 0, gRs!类型), gRs!序号, gRs!修改日期, vFileData)
            
    Call MusicPlay(strFile)

End Sub

Private Sub Command1_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    mvarSvrPicRange = ""
    mvarSvrPicType = ""
    
    For i = 0 To lblSize.UBound
        lblSize(i).Caption = ""
    Next
    
    
    cmbMode.AddItem "平铺"
    cmbMode.AddItem "拉伸"
    cmbMode.AddItem "居中"
    Select Case GetPara("背景显示模式", "平铺")
        Case "拉伸"
            cmbMode.ListIndex = 1
        Case "居中"
            cmbMode.ListIndex = 2
        Case Else
            cmbMode.ListIndex = 0
    End Select
    
    cbo.AddItem "[无]"
    gstrSQL = "select 序号,名称 from 咨询图片元素 where 类型=3"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cbo.AddItem IIf(IsNull(gRs!名称), "", gRs!名称)
            cbo.ItemData(cbo.NewIndex) = IIf(IsNull(gRs!序号), 0, gRs!序号)
            gRs.MoveNext
        Wend
    End If
    cbo.ListIndex = 0
    
    '读取主页信息
    On Error GoTo errHand
    gstrSQL = "select 宣传标语,页面背景,背景音乐,编码 from 咨询页面目录 where 页面序号=0"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        mstrHomeCode = IIf(IsNull(gRs!编码), "", gRs!编码)
        UsrPicture(1).Tag = IIf(IsNull(gRs!页面背景), 0, gRs!页面背景)
        UsrPicture(2).Tag = IIf(IsNull(gRs!宣传标语), 0, gRs!宣传标语)
        cbo.ListIndex = FindCboIndex(cbo, IIf(IsNull(gRs!背景音乐), 0, gRs!背景音乐))
        
        gstrSQL = "select A.插图序号 from 咨询段落目录 A where A.页面序号=0 and A.段落序号=1"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then UsrPicture(0).Tag = IIf(IsNull(gRs!插图序号), 0, gRs!插图序号)
        
        gstrSQL = "select A.插图序号 from 咨询段落目录 A where A.页面序号=0 and A.段落序号=2"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then UsrPicture(3).Tag = IIf(IsNull(gRs!插图序号), 0, gRs!插图序号)
        
        Call ShowPicture(Val(UsrPicture(0).Tag), 0)
        Call ShowPicture(Val(UsrPicture(1).Tag), 1)
        Call ShowPicture(Val(UsrPicture(2).Tag), 2)
        Call ShowPicture(Val(UsrPicture(3).Tag), 3)
    End If
            
    cmdOK.Tag = ""
    Exit Sub
errHand:
    cmdOK.Tag = ""
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowPicture(ByVal PicNo As Long, ByVal Index As Long)
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "select 序号,宽度,高度,类型 from 咨询图片元素 where 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PicNo)
    If rs.BOF = False Then
        Call UsrPicture(Index).ShowPictureByFieldNew(rs!序号, rs!宽度 * Screen.TwipsPerPixelX, rs!高度 * Screen.TwipsPerPixelY, IIf(IsNull(rs!类型), 0, rs!类型))
        If Index = 0 Then lblSize(Index).Caption = "宽度:" & Format(rs!宽度 * Screen.TwipsPerPixelX / 567, "0.0(厘米)") & vbCrLf & "高度:" & Format(rs!高度 * Screen.TwipsPerPixelY / 567, "0.0(厘米)")
    End If
    CloseRecord rs
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case cmbMode.ListIndex
        Case 1
            SetPara "背景显示模式", "拉伸"
        Case 2
            SetPara "背景显示模式", "居中"
        Case Else
            SetPara "背景显示模式", "平铺"
    End Select
End Sub

