VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPacsApplyWord 
   Caption         =   "影像申请常用词句"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "frmPacsApplyWord.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6255
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdEdit 
      Caption         =   "编辑(&E)"
      Height          =   360
      Left            =   3000
      TabIndex        =   6
      Top             =   5400
      Width           =   900
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfWord 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _cx             =   10610
      _cy             =   9128
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
      BackColorFixed  =   14811105
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
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
      ExplorerBar     =   1
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   5235
      TabIndex        =   5
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   360
      Left            =   1080
      TabIndex        =   2
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   900
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Left            =   4200
      Picture         =   "frmPacsApplyWord.frx":6852
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNoCheck 
      Height          =   255
      Left            =   3960
      Picture         =   "frmPacsApplyWord.frx":6BC4
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPacsApplyWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDept As Long    '科室ID
Private mlngNo As Long      '人员ID
Private mstrSort As String  '项目分类
Private mstrWord As String  '返回词句
Private mblnEdit As Boolean '是否发生编辑
Private mblnIsEdit As Boolean
Private mstrCurWord As String '用于词句是否发生改动的记录

Private Const M_STR_TITLE = "影像申请单常用词句"

Private Enum TColName
    colID = 0
    col序号 = 1
    col图标 = 2
    col是否通用 = 3
    col创建人 = 4
    col排序 = 5
    col词句内容 = 6
End Enum

Public Function ShowPacsApplyWord(lngDept As Long, lngNo As Long, strSort As String, ower As Object) As String
    mlngDept = lngDept
    mlngNo = lngNo
    mstrSort = strSort
    mstrWord = ""
    mblnEdit = False
    mblnIsEdit = False
    
    Me.Show 1, ower
    
    ShowPacsApplyWord = mstrWord
End Function

Private Sub cmdAdd_Click()
    On Error GoTo errHandle
        
    cmdSave.Enabled = True
    
    Call AddRow(0, 0, "", "")
    
    vsfWord.Select vsfWord.Row, TColName.col词句内容
    vsfWord.EditCell
    mblnEdit = True
    If vsfWord.Row > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    Unload Me
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub cmdDelete_Click()
'删除词句时有两种情况：一是通过勾选进行多条删除，而是只删除当前选中词句
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    Dim blnSelect As Boolean
    Dim i As Long
    
    On Error GoTo errHandle
    
    For i = 1 To vsfWord.Rows - 1
        If vsfWord.Cell(flexcpData, i, TColName.col图标) = 1 Then
            If Val(vsfWord.RowData(i)) < 0 Then
                MsgBox "您不是这条词句的创建者，无法执行该操作！", vbInformation, M_STR_TITLE
                vsfWord.Select i, 1
                vsfWord.ShowCell i, 1
                Exit Sub
            End If
            blnSelect = True
        End If
    Next
    
    If blnSelect Then
    '删除多条词句
        If MsgBox("是否删除所选词句？", vbYesNo, M_STR_TITLE) = vbYes Then
            i = 1
            While i <= vsfWord.Rows - 1
                If vsfWord.Cell(flexcpData, i, TColName.col图标) = 1 Then
                    strSQL = "Zl_影像申请常用词句_Delete(" & Val(GetValue(i, TColName.colID)) & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, "删除词句")
                    vsfWord.RemoveItem i
                    vsfWord.Refresh
                Else
                    i = i + 1
                End If
                
            Wend
        Else
            Exit Sub
        End If
    Else
    '删除当前一条词句
        If vsfWord.Row < 1 Then
            MsgBox "请先选择需要操作的词句。", vbInformation, M_STR_TITLE
            Exit Sub
        End If
        
        If Val(vsfWord.RowData(vsfWord.Row)) < 0 Then
            MsgBox "您不是这条词句的创建者，无法执行该操作！", vbInformation, M_STR_TITLE
            Exit Sub
        End If
        
        If MsgBox("是否删除词句【" & Trim(GetValue(vsfWord.Row, TColName.col词句内容)) & "】？", vbYesNo, M_STR_TITLE) = vbYes Then
            strSQL = "Zl_影像申请常用词句_Delete(" & Val(GetValue(vsfWord.Row, TColName.colID)) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "删除词句")
            vsfWord.RemoveItem vsfWord.Row
            RefreshNum vsfWord.Row
        Else
            Exit Sub
        End If
    End If

    If vsfWord.Row > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    
    If vsfWord.Rows <= 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub RefreshNum(lngRow As Long)
    Dim i As Long
    Dim lngCount As Long
    
    If lngRow <= 0 Then Exit Sub
    lngCount = 0
    With vsfWord
        For i = lngRow To .Rows - 1
            .TextMatrix(i, TColName.col序号) = .Row + lngCount
            lngCount = lngCount + 1
        Next
    End With
End Sub

Private Function IsCreator(lngID As Long) As Boolean
'判断当前操作人员是否为词句的创建者
    
    If lngID <> mlngNo Then
        IsCreator = False
    Else
        IsCreator = True
    End If
End Function

Private Sub cmdEdit_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Val(cmdEdit.Tag) = 0 Then
        cmdEdit.Tag = 1
        cmdEdit.Caption = "退出(&E)"
        cmdAdd.Visible = True
        cmdDelete.Visible = True
        cmdSave.Visible = True
        cmdSave.Enabled = False
        cmdSure.Visible = False
        cmdCancel.Visible = False
        vsfWord.ColHidden(TColName.col是否通用) = False
        vsfWord.ColHidden(TColName.col创建人) = False
                  
        If vsfWord.Rows > 1 Then
            vsfWord.Cell(flexcpSort, 1, TColName.col排序, vsfWord.Rows - 1, TColName.col词句内容) = flexSortStringNoCaseAscending
        End If
        
        For i = 1 To vsfWord.Rows - 1
            If vsfWord.RowHidden(i) = True Then
                vsfWord.RowHidden(i) = False
            End If
            If vsfWord.RowData(i) < 0 Then
                vsfWord.Cell(flexcpBackColor, i, TColName.col是否通用, i, TColName.col词句内容) = &HC0FFFF
            End If
        Next
    Else
        If mblnEdit Then
            If MsgBox("编辑内容是否保存？", vbYesNo, M_STR_TITLE) = vbYes Then
                If Not SaveData Then
                    Exit Sub
                End If
            Else
                Call InitData
            End If
        End If
        mblnEdit = False
        
        cmdEdit.Tag = 0
        cmdEdit.Caption = "编辑(&E)"
        cmdAdd.Visible = False
        cmdDelete.Visible = False
        cmdSave.Visible = False
        cmdSure.Visible = True
        cmdCancel.Visible = True
        
        cmdSave.Enabled = False
        If vsfWord.Row > 0 Then
            cmdDelete.Enabled = True
        Else
            cmdDelete.Enabled = False
        End If
        
        vsfWord.ColHidden(TColName.col是否通用) = True
        vsfWord.ColHidden(TColName.col创建人) = True

        If vsfWord.Rows > 1 Then
            vsfWord.Cell(flexcpSort, 1, TColName.col词句内容, vsfWord.Rows - 1, TColName.col词句内容) = flexSortStringNoCaseAscending
        End If

        For i = 1 To vsfWord.Rows - 1
            If IsRepeted(i - 1, Trim(GetValue(i, TColName.col词句内容))) Then
                vsfWord.RowHidden(i) = True
            End If
            
            If vsfWord.RowData(i) < 0 Then
                vsfWord.Cell(flexcpBackColor, i, TColName.col是否通用, i, TColName.col创建人) = &H80000005
            End If
        Next
    End If
    
    Call Form_Resize
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If Not SaveData Then Exit Sub
    
    If vsfWord.Row > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    
    cmdSave.Enabled = False
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    Dim blnTag As Boolean
    Dim lngID As Long
    Dim i As Long
    
    For i = 1 To vsfWord.Rows - 1
        '新增和修改词句时，要判断词句是否已存在
        If Val(GetValue(i, TColName.colID)) = 0 Or Val(vsfWord.RowData(i)) = 1 Then
            If CheckRepeted(i) Then
                MsgBox IIF(vsfWord.Cell(flexcpData, i, TColName.col是否通用) = 0, "您已添加了该词句。", "该词句在科室通用词句中已存在。"), vbInformation, M_STR_TITLE
                vsfWord.Select i, TColName.col词句内容
                vsfWord.EditCell
                Exit Function
            End If
        End If
        
        If Len(GetValue(i, TColName.col词句内容)) > 200 Then
            MsgBox "词句内容过长，不能超过200个字。", vbInformation, M_STR_TITLE
            vsfWord.Select i, TColName.col词句内容
            vsfWord.EditCell
            Exit Function
        End If
    Next
    
    i = 1
    While i <= vsfWord.Rows - 1
        blnTag = False
        '新增
        If Val(GetValue(i, TColName.colID)) = 0 Then
            If Len(Trim(GetValue(i, TColName.col词句内容))) > 0 Then
                
                strSQL = "select Zl_影像申请常用词句_Insert([1],[2],[3],[4],[5]) as 返回值 from dual"
                Set rsResult = zlDatabase.OpenSQLRecord(strSQL, "新增数据", mstrSort, Replace(Trim(GetValue(i, TColName.col词句内容)), "'", "''"), vsfWord.Cell(flexcpData, i, TColName.col是否通用), mlngDept, mlngNo)
                
                If rsResult.RecordCount > 0 Then
                    vsfWord.TextMatrix(i, TColName.colID) = Val(Nvl(rsResult.Fields!返回值))
                    vsfWord.TextMatrix(i, TColName.col创建人) = UserInfo.姓名
                    lngID = Val(Nvl(rsResult.Fields!返回值))
                End If
            Else
                vsfWord.RemoveItem i
                blnTag = True
            End If
        End If
        
        '修改
        If Not blnTag Then
            If Val(vsfWord.RowData(i)) = 1 Then
                strSQL = "Zl_影像申请常用词句_Update(" & Val(GetValue(i, TColName.colID)) & ",'" & Replace(Trim(GetValue(i, TColName.col词句内容)), "'", "''") & "'," & vsfWord.Cell(flexcpData, i, TColName.col是否通用) & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "修改数据")
                
                vsfWord.RowData(i) = 0
                lngID = Val(GetValue(i, TColName.colID))
            End If
            
            i = i + 1
        End If
    Wend
    
    If vsfWord.Rows > 1 Then
        vsfWord.Cell(flexcpSort, 1, TColName.col排序, vsfWord.Rows - 1, TColName.col词句内容) = flexSortStringNoCaseAscending
    End If
    
    If vsfWord.Rows > 1 And lngID > 0 Then
        For i = 1 To vsfWord.Rows - 1
            If lngID = Val(GetValue(i, TColName.colID)) Then
                vsfWord.Select i, 1
                vsfWord.ShowCell i, 1
            End If
        Next
    End If
    mblnEdit = False
    SaveData = True
End Function

Private Function GetValue(lngRow As Long, lngCol As Long) As String
    GetValue = vsfWord.TextMatrix(lngRow, lngCol)
End Function

Private Function CheckRepeted(lngRow As Long) As Boolean
'编辑词句判断是否同级别重复
    Dim i As Long
    
    CheckRepeted = False
    
    For i = 1 To vsfWord.Rows - 1
        If Trim(GetValue(lngRow, TColName.col词句内容)) = Trim(GetValue(i, TColName.col词句内容)) And vsfWord.Cell(flexcpData, i, TColName.col是否通用) = vsfWord.Cell(flexcpData, lngRow, TColName.col是否通用) And Len(Trim(GetValue(lngRow, TColName.col词句内容))) > 0 And i <> lngRow Then
            CheckRepeted = True
            Exit Function
        End If
    Next
End Function

Private Sub cmdSure_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
'    For i = 1 To vsfWord.Rows - 1
'        If vsfWord.Cell(flexcpData, i, TColName.col图标) = 1 Then
'             mstrWord = mstrWord & IIF(Len(mstrWord) = 0, "", "；") & Trim(vsfWord.TextMatrix(i, TColName.col词句内容))
'        End If
'    Next
    mstrWord = Trim(GetValue(vsfWord.Row, TColName.col词句内容))
    
    Unload Me
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Call InitGrid
    Call InitFace
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Val(cmdEdit.Tag) = 1 And mblnIsEdit Then
        Call vsfWord_AfterEdit(vsfWord.Row, vsfWord.Col)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.ScaleHeight > 3000 And Me.ScaleWidth > 6000 Then
        vsfWord.Left = 120
        vsfWord.Top = 120
        
        vsfWord.Height = Me.ScaleHeight - cmdAdd.Height - 360
        vsfWord.Width = Me.ScaleWidth - 240
        
        cmdAdd.Left = vsfWord.Left
        cmdAdd.Top = vsfWord.Top + vsfWord.Height + 120
        
        cmdDelete.Left = cmdAdd.Left + cmdAdd.Width + 60
        cmdDelete.Top = cmdAdd.Top
        
        cmdSave.Left = cmdDelete.Left + cmdDelete.Width + 60
        cmdSave.Top = cmdAdd.Top
        
        cmdEdit.Left = IIF(Val(cmdEdit.Tag) = 1, Me.ScaleWidth - cmdCancel.Width - 120, vsfWord.Left)
        cmdEdit.Top = cmdAdd.Top
        
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
        cmdCancel.Top = cmdAdd.Top
        
        cmdSure.Left = cmdCancel.Left - cmdSure.Width - 60
        cmdSure.Top = cmdAdd.Top
        
'        vsfWord.ColWidth(TColName.col词句内容) = vsfWord.Width - 340 - IIF(vsfWord.ColHidden(TColName.col创建人), 0, vsfWord.ColWidth(TColName.col创建人)) - IIF(vsfWord.ColHidden(TColName.col是否通用), 0, vsfWord.ColWidth(TColName.col是否通用))
    End If
End Sub

Private Sub InitFace()
'初始化界面显示

    Me.Caption = mstrSort & IIF(InStr(mstrSort, "项目") = 0, "项目", "") & "选择"
    
    Call InitData
    
    cmdEdit.Tag = 0
    
    cmdAdd.Visible = False
    cmdDelete.Visible = False
    cmdSave.Visible = False
End Sub

Private Sub InitGrid()
    
    With vsfWord
        
        .Cols = 7
        .ColHidden(TColName.colID) = True
        .FixedRows = 1
        .FixedCols = 0
        .ColWidth(TColName.colID) = 0
        .ColWidth(TColName.col序号) = 480
        .ColWidth(TColName.col图标) = 600
        .ColWidth(TColName.col是否通用) = 480
        .ColWidth(TColName.col排序) = 0
        .ColWidth(TColName.col创建人) = 1000
        .RowHeightMin = 350
        .RowHeightMax = 350
        .ExtendLastCol = True
        .ScrollTrack = True
        .ColHidden(TColName.col图标) = True
        .ColHidden(TColName.col是否通用) = True
        .ColHidden(TColName.col序号) = True
        .ColHidden(TColName.col排序) = True
        .ColHidden(TColName.col创建人) = True
        
        .TextMatrix(0, TColName.col序号) = "序号"
        .TextMatrix(0, TColName.colID) = "ID"
        .TextMatrix(0, TColName.col图标) = "选择"
        .TextMatrix(0, TColName.col是否通用) = "通用"
        .TextMatrix(0, TColName.col排序) = "排序"
        .TextMatrix(0, TColName.col词句内容) = "词句内容"
        .TextMatrix(0, TColName.col创建人) = "创建人"
        
    End With
End Sub

Private Sub InitData()
'初始化界面数据
    Dim strSQL As String
    Dim rsResult As New ADODB.Recordset
    Dim blnOwer As Boolean
    Dim blnHidden As Boolean
    
    vsfWord.Rows = 1
    strSQL = "Select a.Id, a.词句内容, a.是否通用, a.创建人员id,b.姓名 as 创建人" & vbNewLine & _
                "From (Select Id, 项目分类, 词句内容, 是否通用, 创建人员id" & vbNewLine & _
                "       From 影像申请常用词句" & vbNewLine & _
                "       Where 创建人员id = [1] And 是否通用 = [2]" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Id, 项目分类, 词句内容, 是否通用, 创建人员id From 影像申请常用词句 Where 科室id = [3] And 是否通用 = [4]) a,人员表 b" & vbNewLine & _
                "Where a.创建人员id = b.id and a.项目分类 = [5]" & vbNewLine & _
                "Order By a.词句内容"

    Set rsResult = zlDatabase.OpenSQLRecord(strSQL, "获取词句", mlngNo, 0, mlngDept, 1, mstrSort)
    
    While Not rsResult.EOF
        blnOwer = IsCreator(Val(Nvl(rsResult.Fields!创建人员ID)))
        blnHidden = IsRepeted(vsfWord.Rows - 1, Nvl(rsResult.Fields!词句内容))
        AddRow Val(Nvl(rsResult.Fields!ID)), Val(Nvl(rsResult.Fields!是否通用)), Nvl(rsResult.Fields!词句内容), Nvl(rsResult.Fields!创建人), blnOwer, blnHidden
        rsResult.MoveNext
    Wend
    
    If vsfWord.Rows > 1 Then
        vsfWord.Select 1, 1
        vsfWord.ShowCell 1, 1
    End If
End Sub

Private Function IsRepeted(lngRow As Long, strValue As String) As Boolean
'非编辑界面重复词句判断
    Dim i As Long
    
    If lngRow < 1 Then Exit Function
    
    IsRepeted = False
    For i = 1 To lngRow
        If Trim(GetValue(i, TColName.col词句内容)) = Trim(strValue) Then
            IsRepeted = True
            Exit Function
        End If
    Next
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandle
    
    If mblnEdit Then
        If MsgBox("编辑内容是否保存？", vbYesNo, M_STR_TITLE) = vbYes Then
            If Not SaveData Then
                Cancel = 1
                Exit Sub
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clea
End Sub

Private Sub vsfWord_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    mblnIsEdit = False
    
    If Col = TColName.col词句内容 And Row > 0 And Val(cmdEdit.Tag) = 1 Then
        '修改词句时不能空
        If Len(Trim(GetValue(Row, TColName.col词句内容))) = 0 And Val(GetValue(Row, TColName.colID)) > 0 Then
            MsgBox "词句内容不能为空。", vbInformation, M_STR_TITLE
            vsfWord.TextMatrix(Row, TColName.col词句内容) = mstrCurWord
            Exit Sub
        End If
        
        '判断哪些词句进行过修改
        If Val(GetValue(Row, TColName.colID)) > 0 Then
            vsfWord.RowData(Row) = 1
            mblnEdit = True
            cmdSave.Enabled = True
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub vsfWord_Click()
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If vsfWord.Row < 1 Then Exit Sub
    If Val(cmdEdit.Tag) = 1 Then
        lngRow = vsfWord.Row
        If vsfWord.Col = TColName.col图标 And lngRow > 0 Then
            If vsfWord.Cell(flexcpData, lngRow, TColName.col图标) = 0 Then
                vsfWord.Cell(flexcpData, lngRow, TColName.col图标) = 1
                vsfWord.Cell(flexcpPicture, lngRow, TColName.col图标) = imgCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.col图标) = flexPicAlignCenterCenter
            Else
                vsfWord.Cell(flexcpData, lngRow, TColName.col图标) = 0
                vsfWord.Cell(flexcpPicture, lngRow, TColName.col图标) = imgNoCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.col图标) = flexPicAlignCenterCenter
            End If
        End If
    
        If vsfWord.Col = TColName.col是否通用 And lngRow > 0 And Val(vsfWord.RowData(lngRow)) >= 0 Then
            mblnEdit = True
            vsfWord.RowData(lngRow) = 1
            cmdSave.Enabled = True
            If vsfWord.Cell(flexcpData, lngRow, TColName.col是否通用) = 0 Then
                vsfWord.Cell(flexcpData, lngRow, TColName.col是否通用) = 1
                vsfWord.TextMatrix(lngRow, TColName.col排序) = 1
                vsfWord.Cell(flexcpPicture, lngRow, TColName.col是否通用) = imgCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.col是否通用) = flexPicAlignCenterCenter
            Else
                vsfWord.Cell(flexcpData, lngRow, TColName.col是否通用) = 0
                vsfWord.TextMatrix(lngRow, TColName.col排序) = 0
                vsfWord.Cell(flexcpPicture, lngRow, TColName.col是否通用) = imgNoCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.col是否通用) = flexPicAlignCenterCenter
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub vsfWord_DblClick()
    On Error GoTo errHandle
    
    If vsfWord.Row > 0 Then
        If Val(cmdEdit.Tag) = 0 Then
            If vsfWord.Row <= 0 Then Exit Sub
            
            If Val(cmdEdit.Tag) = 0 Then
                mstrWord = Trim(GetValue(vsfWord.Row, TColName.col词句内容))
                Unload Me
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub vsfWord_RowColChange()
    On Error GoTo errHandle
    
    If vsfWord.Row <= 0 Then Exit Sub
    If Val(cmdEdit.Tag) = 1 Then
        If Val(vsfWord.RowData(vsfWord.Row)) >= 0 And (vsfWord.Col = TColName.col词句内容) Then
            vsfWord.Editable = flexEDKbdMouse
        Else
            vsfWord.Editable = flexEDNone
        End If
    Else
        vsfWord.Editable = flexEDNone
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub AddRow(lngID As Long, lngGeneral As Long, strWord As String, strOwer As String, Optional blnOwer As Boolean = True, Optional blnHidden As Boolean = False)
    With vsfWord
        .Rows = vsfWord.Rows + 1
        .ShowCell .Rows - 1, TColName.col词句内容
        .Select .Rows - 1, TColName.col词句内容
        .TextMatrix(vsfWord.Rows - 1, TColName.colID) = lngID
        .TextMatrix(vsfWord.Rows - 1, TColName.col序号) = vsfWord.Rows - 1
        .Cell(flexcpPicture, vsfWord.Rows - 1, TColName.col图标) = imgNoCheck.Picture
        .Cell(flexcpData, vsfWord.Rows - 1, TColName.col图标) = 0
        .Cell(flexcpPictureAlignment, vsfWord.Rows - 1, TColName.col图标) = flexPicAlignCenterCenter
        
        .Cell(flexcpPicture, vsfWord.Rows - 1, TColName.col是否通用) = IIF(lngGeneral = 1, imgCheck.Picture, imgNoCheck.Picture)
        .Cell(flexcpData, vsfWord.Rows - 1, TColName.col是否通用) = lngGeneral
        
        '是否通用列的排序列
        .TextMatrix(vsfWord.Rows - 1, TColName.col排序) = lngGeneral
        .Cell(flexcpPictureAlignment, vsfWord.Rows - 1, TColName.col是否通用) = flexPicAlignCenterCenter
        
        .TextMatrix(vsfWord.Rows - 1, TColName.col词句内容) = strWord
        .Cell(flexcpAlignment, vsfWord.Rows - 1, TColName.col词句内容) = flexAlignLeftCenter
        
        .TextMatrix(vsfWord.Rows - 1, TColName.col创建人) = strOwer
        .Cell(flexcpAlignment, vsfWord.Rows - 1, TColName.col创建人) = flexAlignLeftCenter
        
        If Not blnOwer Then
            .RowData(vsfWord.Rows - 1) = -1
        Else
            .RowData(vsfWord.Rows - 1) = 0
        End If
        
        If blnHidden Then
            .RowHidden(vsfWord.Rows - 1) = True
        End If
    End With
End Sub

Private Sub vsfWord_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑前的词句
    On Error GoTo errHandle
    
    mstrCurWord = GetValue(Row, Col)

    cmdSave.Enabled = True
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub


