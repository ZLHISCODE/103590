VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDataSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据提取设置"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmDataSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5925
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frmDataSet.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4710
      TabIndex        =   6
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3585
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "(&D)导入时删除系统中医价数据"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   6075
   End
   Begin VB.ComboBox cmbTable 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
   Begin ZL9BillEdit.BillEdit msfFields 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "数据源表(&S)："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "可选字段(&F)："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const col_Check As Integer = 0
Private Const col_Name As Integer = 1
Private Const col_Type As Integer = 2
Private Const col_Length As Integer = 3
Private Const col_DestFld As Integer = 4

Private Const Dest_Table As String = "标准医价规范"
Private Const UQ_Field As String = "项目编码"

'返回参数
Private strSourceSQL As String, strDestFields As String, ifClear As Boolean

Private objDbase As DAO.Database
Private DataEngine As New DAO.DBEngine, DBWork As DAO.Workspace

Public Sub ShowMe(objParent As Object, ByVal DataSource As String, _
    ByRef SourceSQL As String, ByRef DestFields As String, ByRef ifDeleteData As Boolean)
    
    Dim strTmp As String, strTables() As Variant
    Dim i As Long
    strDestFields = "": strSourceSQL = "": ifClear = False
    
    strTmp = SetConnect(objParent.hwnd, "", "DSN=" + DataSource)
    If Len(Trim(strTmp)) > 0 Then
        objParent.MousePointer = vbHourglass
        Set DBWork = DataEngine.CreateWorkspace("JetWork", "Admin", "", dbUseJet)
        Set objDbase = getDatabase(DBWork, strTmp)
        strTables = getTables(strTmp)
        objParent.MousePointer = vbDefault
        If UBound(strTables, 1) < 0 Then Exit Sub
        
        With Me.cmbTable
            .Clear
            For i = 0 To UBound(strTables, 1)
                .AddItem strTables(i)
            Next
            .ListIndex = 0
        End With
        
        SourceSQL = "": DestFields = "": ifDeleteData = False
        Me.Show vbModal, objParent
        If Len(Trim(strSourceSQL)) > 0 Then
            SourceSQL = strSourceSQL: DestFields = strDestFields
            ifDeleteData = ifClear
        End If
    End If
End Sub

Private Sub cmbTable_Click()
    Dim objTable As DAO.TableDef
    Dim i As Long, tmpItem As ListItem
    
    On Error GoTo DBError
    Set objTable = objDbase.TableDefs(cmbTable.Text)
'    With lvwField
'        .ListItems.Clear
'        .SortKey = 0: .Sorted = True
'    End With
'
'    With objTable
'        For i = 0 To .Fields.Count - 1
'            Set tmpItem = lvwField.ListItems.Add(, "_" & i, .Fields(i).Name)
'            tmpItem.SubItems(1) = GetFieldTypeName(.Fields(i).Type)
'            tmpItem.SubItems(2) = Format(.Fields(i).Size, "@@@@@@")
'        Next
'    End With
    
    With msfFields
        .ClearBill
        If objTable.Fields.Count > 0 Then
            .Rows = objTable.Fields.Count + 1: .Active = True
        Else
            .Rows = 2: .Active = False
        End If
        For i = 0 To objTable.Fields.Count - 1
            .TextMatrix(i + 1, col_Check) = ""
            .TextMatrix(i + 1, col_Name) = objTable.Fields(i).Name
            .TextMatrix(i + 1, col_Type) = GetFieldTypeName(objTable.Fields(i).Type)
            .TextMatrix(i + 1, col_Length) = objTable.Fields(i).Size
            .TextMatrix(i + 1, col_DestFld) = ""
        Next
    End With
    Exit Sub
    
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    strDestFields = "": strSourceSQL = "": ifClear = False
    With msfFields
        If .Active Then
            For i = 1 To .Rows - 1
                If Len(Trim(.TextMatrix(i, col_DestFld))) > 0 Then
                    If InStr(strDestFields, "," & .TextMatrix(i, col_DestFld) & " ") > 0 Then
                        MsgBox "目标字段：" & .TextMatrix(i, col_DestFld) & " 重复，请修改！", vbInformation, gstrSysName
                        .Row = i: .Col = col_DestFld: Exit Sub
                    Else
                        strDestFields = strDestFields + "," + .TextMatrix(i, col_DestFld) + " "
                        strSourceSQL = strSourceSQL + "," + .TextMatrix(i, col_Name) + " "
                    End If
                End If
            Next
            If InStr(strDestFields, "," & UQ_Field & " ") = 0 Then
                MsgBox "必须指定目标字段：" & UQ_Field & "！", vbInformation, gstrSysName
                .Row = 1: .Col = col_DestFld: Exit Sub
            End If
            
            If Len(strDestFields) > 0 Then strDestFields = Mid(strDestFields, 2)
            If Len(strSourceSQL) > 0 Then strSourceSQL = "Select " & _
                Mid(strSourceSQL, 2) & " From [" & Me.cmbTable.Text & "]"
            ifClear = Me.chkClear
        End If
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    With Me.msfFields
        .Active = False: .AllowAddRow = False
        .MsfObj.FixedCols = 0: .Rows = 2: .Cols = 5
        .TextMatrix(0, col_Check) = "选择":     .ColWidth(col_Check) = 500:  .ColAlignment(col_Check) = 1: .ColData(col_Check) = -1
        .TextMatrix(0, col_Name) = "字段名称":       .ColWidth(col_Name) = 1500:   .ColAlignment(col_Name) = 1: .ColData(col_Name) = 5
        .TextMatrix(0, col_Type) = "类型":   .ColWidth(col_Type) = 800:  .ColAlignment(col_Type) = 1: .ColData(col_Type) = 5
        .TextMatrix(0, col_Length) = "长度":       .ColWidth(col_Length) = 1000:   .ColAlignment(col_Length) = -1: .ColData(col_Length) = 5
        .TextMatrix(0, col_DestFld) = "目标字段":     .ColWidth(col_DestFld) = 1500:   .ColAlignment(col_DestFld) = 1: .ColData(col_DestFld) = 1
        .PrimaryCol = 0: .LocateCol = 0
        
        .Row = 1: .Col = 0
    End With
'    With Me.lvwField
'        .ColumnHeaders.Add , "_FldName", "字段名", 1500
'        .ColumnHeaders.Add , "_FldType", "类型", 1000
'        .ColumnHeaders.Add , "_FldLen", "长度", 1000, 1
'    End With
End Sub
'
'Private Sub lvwField_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    With lvwField
'        .SortKey = ColumnHeader.Index - 1: .SortOrder = (.SortOrder + 1) Mod 2: .Sorted = True
'    End With
'End Sub

Private Sub msfFields_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub msfFields_CellCheck(Row As Long, Col As Long)
    With msfFields
        If .TextMatrix(.Row, col_Check) <> .CheckChar Then
            .TextMatrix(.Row, col_DestFld) = ""
        ElseIf .TextMatrix(.Row, col_DestFld) = "" Then
            .TextMatrix(.Row, col_Check) = ""
        End If
    End With
End Sub

Private Sub msfFields_CommandClick()
    With Me.msfFields
        Select Case .Col
            Case col_DestFld
                Call GetDestField(False)
        End Select
    End With
End Sub

Private Sub GetDestField(ByVal ifSearch As Boolean)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    With Me.msfFields
        strSQL = "Select Column_ID as ID, Column_Name As 字段名," & _
            "Decode(Data_Type,'NUMBER','数字','DATE','日期','VARCHAR2','字符串','VARCHAR','字符串','CHAR','字符串','其他') as 类型,Data_Length as 长度 " & _
            "From User_Tab_Columns Where Table_Name='" & Dest_Table & "'" & IIf(Not ifSearch, "", " And Column_Name Like '%" & .Text & "%'")

        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "目标字段", , , , , , True, Me.Left + .Left + 3000, Me.Top + .Top + .CellTop, .RowHeight(.Row), True, , True)
        If Not rsTmp Is Nothing Then
            .Text = rsTmp(1)
            .TextMatrix(.Row, .Col) = rsTmp(1)
            
            If .TextMatrix(.Row, col_Check) <> .CheckChar Then .TextMatrix(.Row, col_Check) = .CheckChar
        Else
            If ifSearch Then .Text = "": .TextMatrix(.Row, .Col) = "": .TextMatrix(.Row, col_Check) = ""
        End If
    End With
End Sub

Private Sub msfFields_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    With Me.msfFields
        Select Case .Col
            Case col_DestFld
                If KeyCode <> vbKeyReturn Then Exit Sub
                Call GetDestField(True)
        End Select
    End With
End Sub
