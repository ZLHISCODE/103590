VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReportImport 
   Caption         =   "报表导入"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmReportImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8445
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCopyTypeSet 
      Caption         =   "设置"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   4950
      Width           =   615
   End
   Begin VB.ComboBox cboCopyType 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdImportTypeSet 
      Caption         =   "设置"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   4590
      Width           =   615
   End
   Begin VB.ComboBox cboImportType 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _cx             =   14843
      _cy             =   7858
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "批量设置覆盖方式"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   4980
      Width           =   1440
   End
   Begin VB.Label lblImportType 
      AutoSize        =   -1  'True
      Caption         =   "批量设置导入方式"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   4620
      Width           =   1440
   End
End
Attribute VB_Name = "frmReportImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsReports As ADODB.Recordset
Private mrsFiles As ADODB.Recordset
Private mlngSys As Long
Private mlngGroup As Long
Private mblnAllImp As Boolean
Private mblnOK As Boolean

Public Function ShowMe(ByVal lngSys As Long, ByVal lngGroup As Long, ByVal blnAllImp As Boolean, ByVal rsReports As ADODB.Recordset, ByRef rsFiles As ADODB.Recordset, ByVal objParent As Object) As Boolean
    Set mrsReports = rsReports
    Set mrsFiles = rsFiles
    mlngSys = lngSys
    mlngGroup = lngGroup
    mblnAllImp = blnAllImp
    Call InitData
    Me.Show 1, objParent
    Set rsFiles = mrsFiles
    ShowMe = mblnOK
End Function


Public Function InitData() As Boolean
    Dim arrTmp As Variant, strInfo As String
    Dim strFilter As String
    Dim intErrType As Integer, intImpType As Integer, lngImpGroup As Long, lngRPTID As Long
    Dim strMsg As String, strOption As String, strReturn As String
    Dim i As Long, lngCount As Long
    Dim cllCover As New Collection '被覆盖的报表ID,用来排除一次导入一个报表多次覆盖
    Dim blnSingle  As Boolean, strFileName As String
    Dim strFlag As String
    
    With cboImportType
        .Clear
        .AddItem "新增导入"
        .AddItem "覆盖导入"
        .ListIndex = 0
    End With
    
    With cboCopyType
        .Clear
        .AddItem "整体覆盖"
        .AddItem "数据源覆盖"
        .ListIndex = 0
    End With
    '初始化表格
    With vsf
        .Cols = 7
        .Rows = 1
        .TextMatrix(0, 0) = "报表编号"
        .ColKey(0) = "报表编号"
        .ColDataType(0) = flexDTString
        .ColWidth(0) = 1200
        
        .TextMatrix(0, 1) = "报表名称"
        .ColKey(1) = "报表名称"
        .ColDataType(1) = flexDTString
        .ColWidth(0) = 1200
        
        .TextMatrix(0, 2) = "导入类型"
        .ColKey(2) = "导入类型"
        .Editable = flexEDKbdMouse
        .ColComboList(2) = "新增导入|覆盖导入"
        .ColWidth(2) = 1200
        
        .TextMatrix(0, 3) = "覆盖类型"
        .ColKey(3) = "覆盖类型"
        .Editable = flexEDKbdMouse
        .ColComboList(3) = Space$(1) & "|整体覆盖|数据源覆盖"
        
        .TextMatrix(0, 4) = "说明"
        .ColKey(4) = "说明"
        .ColDataType(4) = flexDTString
        .ColWidth(4) = 1800
        
        .TextMatrix(0, 5) = "错误标识"
        .ColKey(5) = "错误标识"
        .ColDataType(5) = flexDTLong
        .ColWidth(5) = 0
        
        .TextMatrix(0, 6) = "文件名"
        .ColKey(6) = "文件名"
        .ColDataType(6) = flexDTString
        .ColWidth(6) = 0
    End With
    
    With mrsFiles
        .Filter = "": .Sort = "FilePath Desc"
        blnSingle = mrsFiles.RecordCount = 1 '是否单个报表导入
'        If blnSingle Then strFileName = mrsFiles!FileName
        Do While Not .EOF
            intErrType = 0: intImpType = 0: lngImpGroup = 0: lngRPTID = 0
            arrTmp = Split(GetReportInfo(!FilePath & ""), ";") '获取文件信息
            If UBound(arrTmp) <> 2 Then intErrType = 4 '文件检查
            If Val(arrTmp(2)) <> 9 And intErrType = 0 Then intErrType = 5  '版本检查
            strFileName = arrTmp(1)
            If intErrType = 0 Then
                If mlngSys = 0 Then '非系统报表要求分组的报表中不能存在相同报表
                    '非固定报表全部导入已经确定报表要导入的分组
                    mrsReports.Filter = "名称='" & arrTmp(1) & "' And 编号='" & arrTmp(0) & "'" & IIF(mlngSys = 0 And mblnAllImp, " And 组ID=" & !组ID, "")
                    If mrsReports.EOF Then mrsReports.Filter = "名称='" & arrTmp(1) & "'" & IIF(mlngSys = 0 And mblnAllImp, " And 组ID=" & !组ID, "")
                Else '系统报表通过编号直接查找
                    mrsReports.Filter = "编号='" & arrTmp(0) & "'"
                End If
                '确定报表导入的分组，如果存在的同名的，优先查找没有分组的报表
                mrsReports.Sort = "ID Desc,组ID"
                If Not mrsReports.EOF Then
                    lngRPTID = mrsReports!id: lngImpGroup = mrsReports!组ID
                    If lngRPTID = 0 Then intErrType = 1 '该报表已经被标记新增
                    If intErrType = 0 Then
                        On Error Resume Next
                        cllCover.Add "1", "_" & lngRPTID
                        If Err.Number <> 0 Then Err.Clear: intErrType = 2 '该报表已经被标记覆盖
                        On Error GoTo errH
                    End If
                    If intErrType = 0 Then intImpType = 2
                    '编号名称不匹配
                    If intErrType = 0 And (CStr(arrTmp(0)) <> mrsReports!编号 & "" Or CStr(arrTmp(1)) <> mrsReports!名称) Then intErrType = 6
                Else
                    If mlngSys <> 0 Then intErrType = 3  '系统固定报表必须覆盖同名报表
                    If intErrType = 0 Then intImpType = 1  '非系统报表没有同名，则新增报表
                End If
                If mlngSys = 0 And mblnAllImp Then lngImpGroup = !组ID '非固定报表导入取原来的分组
                '该报表是新增报表，则加入缓存，防止多次增加
                If mrsReports.EOF And mlngSys = 0 Then mrsReports.AddNew Array("Id", "编号", "名称", "组iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), lngImpGroup)
            End If
            
            vsf.Rows = vsf.Rows + 1
            vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("报表编号")) = arrTmp(0)
            vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("报表名称")) = arrTmp(1)
            vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("文件名")) = !FileName
            Select Case intErrType
            Case 2
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("说明")) = "报表""" & strFileName & """存在相同内容的报表无法导入！"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("错误标识")) = 1
            Case 3
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("说明")) = "报表""" & strFileName & """由于没有可以覆盖的报表而无法导入！"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("错误标识")) = 1
            Case 4
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("说明")) = "报表""" & strFileName & """由于内容存在问题而无法导入！"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("错误标识")) = 1
            Case 5
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("说明")) = "报表""" & strFileName & """由于版本不对而无法导入！"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("错误标识")) = 1
            Case 6
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("说明")) = "报表""" & strFileName & """编号或名称与覆盖的报表不相符，请选择确认！"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("错误标识")) = 1
            Case Else
                Select Case intImpType
                Case 1
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("说明")) = "将新增导入报表" & strFileName
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("导入类型")) = "新增导入"
                Case 2
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("说明")) = "报表" & strFileName & "将会覆盖原有报表"
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("导入类型")) = "覆盖导入"
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("覆盖类型")) = "整体覆盖"
                End Select
            End Select

            .Update Array("组ID", "同名ID", "导入类型", "ErrType"), Array(lngImpGroup, lngRPTID, intImpType, intErrType)
            
            .MoveNext
        Loop
    End With
    mrsFiles.Filter = ""
errH:
End Function

Private Sub cmdCopyTypeSet_Click()
    Dim i As Integer
    Dim strCopyType As String
    Dim strFileName As String
    
    With vsf
        strCopyType = cboCopyType.Text
        mrsFiles.Filter = ""
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("导入类型")) = "覆盖导入" Then
                '判断是否因错误无法导入,判断是否组ID存在(若不是分组报表则可以新增导入)
                If Val(.TextMatrix(i, .ColIndex("错误标识"))) = 0 Then
                    .TextMatrix(i, .ColIndex("覆盖类型")) = strCopyType
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdImportTypeSet_Click()
    Dim i As Integer
    Dim strImportType As String
    Dim strFileName As String
    
    With vsf
        strImportType = cboImportType.Text
        mrsFiles.Filter = ""
        For i = 1 To .Rows - 1
            '判断是否分组报表,若是分组报表，则不允许将覆盖导入为新增导入
            mrsFiles.Filter = "FileName='" & .TextMatrix(i, .ColIndex("文件名")) & "'"
            Select Case strImportType
            Case "新增导入"
                '判断是否因错误无法导入,判断是否组ID存在(若不是分组报表则可以新增导入)
                If Val(.TextMatrix(i, .ColIndex("错误标识"))) = 0 And Val(mrsFiles("组ID")) = 0 Then
                    .TextMatrix(i, .ColIndex("导入类型")) = "新增导入"
                    .TextMatrix(i, .ColIndex("覆盖类型")) = Space$(1)
                    .TextMatrix(i, .ColIndex("说明")) = "将新增导入报表" & .TextMatrix(i, .ColIndex("报表名称"))
                End If
            Case "覆盖导入"
                '判断是否因错误无法导入,判断是否组ID存在(若不是分组报表则可以新增导入)
                If Val(.TextMatrix(i, .ColIndex("错误标识"))) = 0 And Val(mrsFiles("同名ID")) > 0 Then
                    .TextMatrix(i, .ColIndex("导入类型")) = "覆盖导入"
                    If .TextMatrix(i, .ColIndex("覆盖类型")) = Space$(1) Then
                        .TextMatrix(i, .ColIndex("覆盖类型")) = "整体覆盖"
                    End If
                    .TextMatrix(i, .ColIndex("说明")) = "报表" & .TextMatrix(i, .ColIndex("报表名称")) & "将会覆盖原有报表"
                End If
            End Select
        Next
    End With
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim intImpType As Integer
    Dim intCopyType As Integer
    Dim intSameID As Integer
    With mrsFiles
        .Filter = ""
        For i = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(i, vsf.ColIndex("错误标识"))) = 0 Then
                intImpType = IIF(vsf.TextMatrix(i, vsf.ColIndex("导入类型")) = "新增导入", 1, 2)
                
                intCopyType = IIF(vsf.TextMatrix(i, vsf.ColIndex("覆盖类型")) = "数据源覆盖", 1, 0)
                mrsFiles.Filter = "FileName='" & vsf.TextMatrix(i, vsf.ColIndex("文件名")) & "'"
                intSameID = Val(mrsFiles("同名ID").Value)
                If intImpType = 1 Then
                    intSameID = 0
                End If
                mrsFiles.Update Array("导入类型", "同名ID", "覆盖类型"), Array(intImpType, intSameID, intCopyType)
            Else
                mrsFiles.Filter = "FileName='" & vsf.TextMatrix(i, vsf.ColIndex("文件名")) & "'"
                mrsFiles.Delete (adAffectCurrent)
            End If
        Next
    End With
    mblnOK = True
    Unload Me
End Sub

Private Sub Command2_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = vsf.ColIndex("导入类型") Then
        If vsf.TextMatrix(Row, vsf.ColIndex("导入类型")) = "新增导入" Then
            vsf.TextMatrix(Row, vsf.ColIndex("覆盖类型")) = Space$(1)
            vsf.TextMatrix(Row, vsf.ColIndex("说明")) = "将新增导入报表" & vsf.TextMatrix(Row, vsf.ColIndex("报表名称"))
        End If
        If vsf.TextMatrix(Row, vsf.ColIndex("导入类型")) = "覆盖导入" Then
            If vsf.TextMatrix(Row, vsf.ColIndex("覆盖类型")) = Space$(1) Then
                vsf.TextMatrix(Row, vsf.ColIndex("覆盖类型")) = "整体覆盖"
                vsf.TextMatrix(Row, vsf.ColIndex("说明")) = "报表" & vsf.TextMatrix(Row, vsf.ColIndex("报表名称")) & "将会覆盖原有报表"
            End If
        End If
    End If
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = vsf.ColIndex("覆盖类型") Then
        '判断是否因错误无法导入
        If Val(vsf.TextMatrix(NewRow, vsf.ColIndex("错误标识"))) > 0 Then Cancel = True
        If vsf.TextMatrix(NewRow, vsf.ColIndex("导入类型")) = "新增导入" Then Cancel = True
    End If
    If NewCol = vsf.ColIndex("导入类型") Then
        '判断是否因错误无法导入
        If Val(vsf.TextMatrix(NewRow, vsf.ColIndex("错误标识"))) > 0 Then Cancel = True
        '判断判断该报表是否分组报表，若不是分组报表则允许新增报表，否则不允许新增导入，并禁止编辑该数据单元格
        mrsFiles.Filter = ""
        mrsFiles.Filter = "FileName='" & vsf.TextMatrix(NewRow, vsf.ColIndex("文件名")) & "'"
        If Val(mrsFiles("组ID").Value) > 0 Then
            Cancel = True
        End If
        If Val(mrsFiles("同名ID").Value) = 0 Then
            Cancel = True
        End If
    End If
End Sub

