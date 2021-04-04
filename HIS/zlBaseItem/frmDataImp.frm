VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataImp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医价数据导入"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmDataImp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6555
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ProgressBar pbImport 
      Height          =   200
      Left            =   120
      TabIndex        =   11
      Top             =   3680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdModi 
      Caption         =   "修改(&B)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      ToolTipText     =   "修改数据源"
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   5160
      TabIndex        =   6
      ToolTipText     =   "删除数据源"
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
      Height          =   350
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "增加数据源"
      Top             =   1560
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   6555
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "导入(&I)"
      Height          =   350
      Left            =   3975
      TabIndex        =   2
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   270
      Picture         =   "frmDataImp.frx":058A
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   3
      Top             =   4110
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   3960
      Width           =   6555
   End
   Begin MSComctlLib.ListView lvwSource 
      Height          =   2160
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   3810
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "医价数据源(&S)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDataImp.frx":06D4
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   120
      Picture         =   "frmDataImp.frx":075C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmDataImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim strSourceName As String
    strSourceName = ""
    
    If frmSrcEdit.EditSource(Me, strSourceName) Then ListSource strSourceName
    lvwSource.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim curIndex As Long
    
    If lvwSource.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("是否删除：" + lvwSource.SelectedItem.Text + "？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
    Call DeleteSetting("ZLSOFT", "医价数据", UCase(lvwSource.SelectedItem.Text))
    Call DeleteSetting("ZLSOFT", "医价数据\" & UCase(lvwSource.SelectedItem.Text))
    
    curIndex = lvwSource.SelectedItem.Index
    Call ListSource
    
    On Error Resume Next
    If curIndex > lvwSource.ListItems.Count - 1 Then curIndex = curIndex - 1
    If curIndex > -1 Then lvwSource.SelectedItem = lvwSource.ListItems(curIndex)
    lvwSource.SetFocus
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdModi_Click()
    Dim strSourceName As String
    strSourceName = Me.lvwSource.SelectedItem.Text
    
    If frmSrcEdit.EditSource(Me, strSourceName) Then ListSource strSourceName
    lvwSource.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim SourceName As String
    Dim strDSN As String, strSourceSQL As String, strDestFields As String, ifDeleteData As Boolean
    Dim DataEngine As New DAO.DBEngine, DBWork As DAO.Workspace
    Dim strConnect As String, objDbase As DAO.Database, rsSource As DAO.Recordset
    
    If Me.lvwSource.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo DBError
    SourceName = Me.lvwSource.SelectedItem.Text
    strDSN = GetSetting("ZLSOFT", "医价数据\" & SourceName, "ODBC", "")
    strSourceSQL = GetSetting("ZLSOFT", "医价数据\" & SourceName, "数据源", "")
    strDestFields = GetSetting("ZLSOFT", "医价数据\" & SourceName, "字段", "")
    ifDeleteData = GetSetting("ZLSOFT", "医价数据\" & SourceName, "清除数据", "false")
    
    '未定义数据导入方式
    If Len(strSourceSQL) = 0 Then
        If MsgBox("未设置数据导入方式，是否继续？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        Call frmDataSet.ShowMe(Me, strDSN, strSourceSQL, strDestFields, ifDeleteData)
        If Len(strSourceSQL) = 0 Then Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    With Me.pbImport
        .Min = 1: .Max = 100: .Value = 1: .Visible = True
    End With
    strConnect = SetConnect(Me.hwnd, "", "DSN=" + strDSN)
    Set DBWork = DataEngine.CreateWorkspace("JetWork", "Admin", "", dbUseJet)
    Set objDbase = getDatabase(DBWork, strConnect)
    Me.pbImport.Value = 5
    '读取源数据
    Set rsSource = objDbase.OpenRecordset(strSourceSQL, dbOpenSnapshot, dbDenyWrite + dbReadOnly)
    Me.pbImport.Value = 10
    
    '导入数据
    ImportData rsSource, strDestFields, ifDeleteData, pbImport
    
    objDbase.Close
    Me.MousePointer = vbDefault
    Me.pbImport.Visible = False
    Exit Sub
    
DBError:
    If ErrCenter() = 1 Then Resume
    Me.MousePointer = vbDefault
    Me.pbImport.Visible = False
    Call SaveErrLog
End Sub

Private Sub ImportData(rsSource As DAO.Recordset, DestFields As String, ifDeleteData As Boolean, Optional pbImport As Object)
    Dim rsTmp As New ADODB.Recordset
    Dim strFldSQL As String, strValueSQl As String, strTmp As String, tmpDate As String
    Dim lngKeyFldIndex As Long
    Dim aFields() As String, i As Long
    Dim dblProgBarInit As Double, lngRecords As Long
    
    If rsSource Is Nothing Then Exit Sub
    If Len(DestFields) = 0 Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo DBError
    If ifDeleteData Then Call zlDatabase.OpenRecordset(rsTmp, "Delete From 标准医价规范", Me.Caption)
    
    aFields = Split(DestFields, ","): strFldSQL = "Insert Into 标准医价规范("
    For i = 0 To UBound(aFields)
        strTmp = strTmp & "," & Trim(aFields(i))
        If aFields(i) = "项目编码" Then lngKeyFldIndex = i
    Next
    If Len(strTmp) > 0 Then strFldSQL = strFldSQL & Mid(strTmp, 2) & ") Values("
    
    rsSource.MoveFirst: rsSource.MoveLast: lngRecords = rsSource.RecordCount
    rsSource.MoveFirst: dblProgBarInit = pbImport.Value
    Do While Not rsSource.EOF
        '删除编号相同的原记录
        If Not ifDeleteData Then
            Call zlDatabase.OpenRecordset(rsTmp, "Delete From 标准医价规范 Where 项目编码='" & rsSource(lngKeyFldIndex) & "'", Me.Caption)
        End If
        
        strValueSQl = ""
        For i = 0 To UBound(aFields)
            If aFields(i) Like "*日期*" Then
                tmpDate = rsSource(i)
                If Not IsDate(rsSource(i)) And InStr(rsSource(i), ".") > 0 Then
                    tmpDate = Mid(rsSource(i), 1, InStr(rsSource(i), ".") - 1)
                End If
                strValueSQl = strValueSQl & ",To_Date('" & Format(tmpDate, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strValueSQl = strValueSQl & ",'" & rsSource(i) & "'"
            End If
        Next
        If Len(strValueSQl) > 0 Then strTmp = strFldSQL & Mid(strValueSQl, 2) & ")"
        Call zlDatabase.OpenRecordset(rsTmp, strTmp, Me.Caption)
        
        On Error Resume Next
        pbImport.Value = pbImport.Value + (1 / lngRecords) * (pbImport.Max - dblProgBarInit)
        On Error GoTo DBError
    
        rsSource.MoveNext
    Loop
    
    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err.Raise Err.Number, "导入医价数据"
End Sub
Private Sub Form_Load()
    With Me.lvwSource.ColumnHeaders
        .Clear
        .Add , "_名称", "名称", 2500
        .Add , "_说明", "说明", 2000
    End With
    
    ListSource
End Sub

Private Sub ListSource(Optional ByVal DefaultItem As String = "")
    Dim aSourceList As Variant, tmpItem As ListItem
    Dim i As Integer
    aSourceList = GetAllSettings("ZLSOFT", "医价数据")
    If Not IsEmpty(aSourceList) Then
        Me.lvwSource.ListItems.Clear
        For i = 0 To UBound(aSourceList, 1)
            With Me.lvwSource
                Set tmpItem = .ListItems.Add(, "_" & i, aSourceList(i, 0))
                tmpItem.SubItems(1) = GetSetting("ZLSOFT", "医价数据\" & UCase(tmpItem.Text), "说明", "")
                If DefaultItem = aSourceList(i, 0) Then tmpItem.Selected = True
            End With
        Next
        If Len(DefaultItem) = 0 Then Me.lvwSource.ListItems(1).Selected = True
    End If
End Sub

Private Sub lvwSource_DblClick()
    Call cmdModi_Click
End Sub
