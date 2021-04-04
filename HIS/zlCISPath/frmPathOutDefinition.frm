VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathOutDefinition 
   Caption         =   "病人出径登记表定义"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12015
   Icon            =   "frmPathOutDefinition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   12015
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12015
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8295
      Width           =   12015
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   6120
         TabIndex        =   10
         ToolTipText     =   "按""Delete""键快捷删除，选中列表列可删除列表内容，选择其他行则删除整行数据。"
         Top             =   200
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9480
         TabIndex        =   9
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "上移(&U)"
         Height          =   350
         Index           =   0
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "下移(&L)"
         Height          =   350
         Index           =   1
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   200
         Width           =   1100
      End
      Begin VB.CheckBox chk必须填写 
         BackColor       =   &H00F0F4E4&
         Caption         =   "结束或退出路径时必须填写本表"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   188
         Width           =   3015
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10680
         TabIndex        =   4
         Top             =   200
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   12000
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   240
         Picture         =   "frmPathOutDefinition.frx":6852
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   12000
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPathOutDefinition.frx":70DA
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   8175
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   7410
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   11895
      _cx             =   20981
      _cy             =   13070
      Appearance      =   0
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPathOutDefinition.frx":7193
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox picAddRow 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00808080&
         Height          =   200
         Left            =   11480
         Picture         =   "frmPathOutDefinition.frx":73B9
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   280
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmPathOutDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Item_Type
    T0数字 = 0
    T1字符 = 1
    T2日期 = 2
    T3布尔型 = 3
    T4单选项 = 4
    T5多选项 = 5
    T6标题 = 6
End Enum

Private Enum CNAME
    col_项目序号 = 0    '顺序号
    col_项目名称 = 1
    Col_页数 = 2
    col_类型 = 3
    col_必填 = 4
    col_通用 = 5
    col_列表 = 6
    col_插入 = 7
    '隐藏列
    col_状态 = 8    '0-新增，1-原始，2-修改
    COL_行号 = 9    '行ID
    col_多选序号 = 10
End Enum
Private Const color_Unmodify = &H8000000F
Private Const mstrComboList = "0-数字|1-字符|2-日期|3-布尔|4-单选项|5-多选项|6-标题"
Private mstrDelItem As String '删除了的项目序号串
Private mstrCaption As String   '窗体名
Private mlng路径ID As Long
Private mintType As Integer  '0-住院；1-门诊

Public Function ShowMe(frmMain As Object, ByVal lng路径ID As Long, ByVal strCaption As String, Optional ByVal intType As Integer) As Boolean
    mlng路径ID = lng路径ID
    mstrCaption = strCaption
    mintType = intType
    Me.Show 1, frmMain
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Call vsItem_KeyDown(vbKeyDelete, 0)
End Sub

Private Sub cmdMove_Click(Index As Integer)
    With vsItem
        If Index = 0 And .Row > .FixedRows Then
            .RowPosition(.Row) = .Row - 1
            .Row = .Row - 1
        ElseIf Index = 1 And .Row < .Rows - 1 Then
            .RowPosition(.Row) = .Row + 1
            .Row = .Row + 1
        End If
        Call FuncNoASC
    End With
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long, arrtmp As Variant
    Dim blnOneNullRow As Boolean
    Dim lngMaxPage As Long
    Dim strPage As String
    Dim strMsg As String
    Dim MaxPage As Long
    
    With vsItem
        '如果删除了所有行，则只有一行空行
        blnOneNullRow = (.Rows = .FixedRows + 1 And .TextMatrix(.FixedRows, col_项目名称) = "")
        
        If Not blnOneNullRow Then
            '检查名称
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, col_项目名称) = "" And i >= 1 Then
                    '多选首行不能为空
                    If Val(.TextMatrix(i - 1, col_类型)) <> T5多选项 Or Val(.TextMatrix(i, col_类型)) <> T5多选项 Then
                        Exit For
                    End If
                End If
                If Val(.TextMatrix(i, col_类型)) = T5多选项 And .TextMatrix(i, col_列表) = "" Then
                    .Select i, col_列表
                    MsgBox "第" & i & "行的列表值为空，请输入项目列表值。", vbInformation, gstrSysName
                    .SetFocus
                    Exit Sub
                End If
                '页数
                strPage = strPage & "," & IIf(Val(.TextMatrix(i, Col_页数)) = 0, 1, Val(.TextMatrix(i, Col_页数)))
                If IIf(Val(.TextMatrix(i, Col_页数)) = 0, 1, Val(.TextMatrix(i, Col_页数))) > MaxPage Then
                    MaxPage = IIf(Val(.TextMatrix(i, Col_页数)) = 0, 1, Val(.TextMatrix(i, Col_页数)))
                End If
            Next
            strPage = Mid(strPage, 2)
            If i <= .Rows - 1 Then
                MsgBox "第" & i & "行项目名称为空，请输入项目名称。", vbInformation, gstrSysName
                .Select i, col_项目名称
                .SetFocus
                Exit Sub
            End If
            
            '检查选项列表
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, col_类型)) = T4单选项 Then
                    If Trim(.TextMatrix(i, col_列表)) = "" Then
                        Exit For
                    ElseIf InStr(.TextMatrix(i, col_列表), ",") = 0 And Mid(.TextMatrix(i, col_列表), 1, 1) <> "[" Then
                        Exit For
                    Else
                        If ZLCommFun.ActualLen(.TextMatrix(i, col_列表)) >= 100 Then
                            Exit For
                        End If
                        arrtmp = Split(.TextMatrix(i, col_列表), ",")
                        For j = 0 To UBound(arrtmp)
                            If Trim(arrtmp(j)) = "" Then
                                Exit For
                            End If
                        Next
                        If j <= UBound(arrtmp) Then
                            Exit For
                        End If
                    End If
                End If
            Next
            If i <= .Rows - 1 Then
                MsgBox "第" & i & "行选项列表格式不符合要求，每个选项不能为空，多个选项请以逗号分隔。", vbInformation, gstrSysName
                .Select i, col_列表
                .SetFocus
                Exit Sub
            End If
            
        ElseIf chk必须填写.Value = 1 Then
            MsgBox "没有定义填写项目时，不能设置为必须填写。", vbInformation, gstrSysName
            .Select .FixedRows, col_列表
            .SetFocus
            Exit Sub
        End If
        '检查页号是否连续
        If Not (.Rows = 2 And .TextMatrix(1, col_项目名称) = "") Then
            If Not CheckPageNum(strPage, MaxPage, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                .SetFocus
                Exit Sub
            End If
        End If
        
    End With
    
    
    If SaveData(blnOneNullRow) Then
        Unload Me
    End If
End Sub

Private Function CheckPageNum(ByVal strPage As String, ByVal MaxPage As Long, strMsg As String) As Boolean
'功能：判断页数是否连续
    Dim strSql As String, rsTmp As Recordset
    
    strSql = "Select Rownum From Dual Connect By Rownum < [1]+1 Minus Select Column_Value From Table(f_Num2list([2]))"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, MaxPage, strPage)
    If rsTmp.RecordCount > 0 Then
        strMsg = "缺少第" & rsTmp!Rownum & "页的内容，请检查页数是否连续。"
    Else
       CheckPageNum = True
    End If
End Function

Private Function SaveData(blnOneNullRow As Boolean) As Boolean
'功能：保存数据
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSql As String, strTmp As String
    Dim lng类型 As Long, i As Long, arrtmp As Variant
    Dim lngMaxNO As Long, lngPage As Long, j As Long
    Dim blnGrant As Boolean
    
    If Not blnOneNullRow Then
        With vsItem
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, col_状态)) = 0 Then
                    If lngMaxNO = 0 Then lngMaxNO = GetMaxItemNO
                    lngMaxNO = lngMaxNO + 1
                    
                    strSql = "0,2," & lngMaxNO
                ElseIf Val(.TextMatrix(i, col_状态)) = 2 Then
                    strSql = "1,2," & .TextMatrix(i, COL_行号)
                Else
                    strSql = ""
                End If
                If strSql <> "" Then
                    lng类型 = Val("" & .TextMatrix(i, col_类型))
                    If lng类型 = T0数字 Or lng类型 = T1字符 Or lng类型 = T2日期 Then
                        strTmp = lng类型 & "|" & IIf(.Cell(flexcpChecked, i, col_必填) = 1, 1, 0)
                    ElseIf lng类型 = T4单选项 Then
                        strTmp = lng类型 & "|" & .TextMatrix(i, col_列表)
                        If Mid(.TextMatrix(i, col_列表), 1, 1) = "[" And Mid(.TextMatrix(i, col_列表), Len(.TextMatrix(i, col_列表))) = "]" Then
                            blnGrant = True
                        End If
                    ElseIf lng类型 = T5多选项 Then
                        strTmp = lng类型 & "|" & IIf(.Cell(flexcpChecked, i, col_必填) = 1, 1, 0) & "|" & .TextMatrix(i, col_列表)
                    Else
                        strTmp = lng类型
                    End If
                    lngPage = Val(.TextMatrix(i, Col_页数))
                    If lngPage = 0 Then
                        For j = i - 1 To 1 Step -1
                            If Val(.TextMatrix(j, Col_页数)) <> 0 Then lngPage = Val(.TextMatrix(j, Col_页数)): Exit For
                        Next
                    End If
                    strSql = "Zl_路径报表定义_Update(" & strSql & "," & _
                            ZVal(Val(.TextMatrix(i, col_项目序号))) & ",'" & .TextMatrix(i, col_项目名称) & "','" & strTmp & "',Null," & ZVal(lngPage) & "," _
                            & IIf(.Cell(flexcpChecked, i, col_通用) = 1, 1, 0) & "," & ZVal(mlng路径ID) & "," & ZVal(.TextMatrix(i, col_多选序号)) & "," & glngSys & ")"
                    colSQL.Add strSql, "C" & colSQL.count + 1
                End If
            Next
        End With
    End If
    
    If mstrDelItem <> "" Then
        arrtmp = Split(mstrDelItem, ",")
        For i = 0 To UBound(arrtmp)
            strSql = "Zl_路径报表定义_Update(2,2," & arrtmp(i) & ",Null,Null,Null,Null,NULL,NULL,NULL,NULL," & glngSys & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        Next
    End If
    
    On Error GoTo errH
    If colSQL.count > 0 Then
        gcnOracle.BeginTrans: blnTrans = True
            For i = 1 To colSQL.count
                Call zlDatabase.ExecuteProcedure(IIf(mintType = 1, Replace(colSQL("C" & i), "路径报表定义", "门诊路径报表定义"), colSQL("C" & i)), Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    '保存参数设置
    Call zlDatabase.SetPara("必须填写出径登记表", chk必须填写.Value, glngSys, IIf(mintType = 1, P门诊路径应用, P临床路径应用))
    
    SaveData = True
    If blnGrant Then
        MsgBox "您选择了字典表作为单选项的数据源，请到管理工具中对""临床路径应用""模块进行重新授权。", vbInformation, Me.Caption
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '不允许输入分隔符及单引号
    End If
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, arrtmp As Variant, arrType As Variant
 
    arrType = Split(mstrComboList, "|")
    
    strSql = "Select a.行号, a.页数, Nvl(b.序号, a.项目序号) As 项目序号, a.项目文本1, a.项目文本2, a.路径id, a.多选序号" & vbNewLine & _
                "From 路径报表结构 A, 路径报表序号 B" & vbNewLine & _
                "Where a.报表id = b.报表id(+) And a.行号 = b.行号(+) And a.报表id = 2 And" & vbNewLine & _
                "      (Nvl(a.路径id, b.路径id) = [1] And (Exists (Select 1 From 路径报表序号 Where 报表id = 2 And 路径id = [1]) Or Not Exists (Select 1 From 路径报表结构 Where 报表id = 2 And a.路径id Is Null)))" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.行号, a.页数, a.项目序号, a.项目文本1, a.项目文本2, a.路径id, a.多选序号" & vbNewLine & _
                "From 路径报表结构 A" & vbNewLine & _
                "Where a.报表id = 2 And a.路径id Is Null And Not Exists (Select 1 From 路径报表序号 Where 报表id = 2 And 路径id = [1])" & vbNewLine & _
                "Order By 项目序号, 多选序号"


    On Error GoTo errH
    If mintType = 1 Then strSql = Replace(strSql, "路径报表结构", "门诊路径报表结构"): strSql = Replace(strSql, "路径报表序号", "门诊路径报表序号")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
    With vsItem
        .Rows = .FixedRows
        If rsTmp.RecordCount = 0 Then
            Call AddNewRow
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col_项目序号) = rsTmp!项目序号 & ""
                .TextMatrix(i, Col_页数) = rsTmp!页数 & ""
                .TextMatrix(i, col_项目名称) = rsTmp!项目文本1 & ""
                .Cell(flexcpChecked, i, col_通用) = IIf(rsTmp!路径ID & "" = "", 1, 0)
                
                arrtmp = Split(rsTmp!项目文本2, "|")    '=类型|是否必填或选项1,选项2,...
                .TextMatrix(i, col_类型) = arrType(Val(arrtmp(0)))
                
                If Val(arrtmp(0)) = T3布尔型 Or Val(arrtmp(0)) = T4单选项 Then
                    .Cell(flexcpChecked, i, col_必填) = 1
                    .Cell(flexcpBackColor, i, col_必填) = color_Unmodify
                ElseIf Val(arrtmp(0)) = T6标题 Then
                    .Cell(flexcpChecked, i, col_必填) = 0
                    .Cell(flexcpBackColor, i, col_必填) = color_Unmodify
                ElseIf Val(arrtmp(0)) = T5多选项 Then
                    If Val(arrtmp(1)) = 1 Then
                        .Cell(flexcpChecked, i, col_必填) = 1
                    End If
                    .TextMatrix(i, col_列表) = arrtmp(2)
                ElseIf UBound(arrtmp) > 0 Then  '数字，字符，日期
                    If Val(arrtmp(1)) = 1 Then
                        .Cell(flexcpChecked, i, col_必填) = 1
                    End If
                End If
                
                If Val(arrtmp(0)) = T4单选项 Then
                    If UBound(arrtmp) > 0 Then
                        .TextMatrix(i, col_列表) = arrtmp(1)
                    End If
                End If
                
                .TextMatrix(i, col_状态) = 1
                .TextMatrix(i, COL_行号) = rsTmp!行号
                
                rsTmp.MoveNext
            Next
        End If
    End With
    Call FuncNoASC
    vsItem.Row = 1
    chk必须填写.Value = zlDatabase.GetPara("必须填写出径登记表", glngSys, IIf(mintType = 1, P门诊路径应用, P临床路径应用), 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)

    mstrDelItem = ""
    Me.Caption = mstrCaption
    vsItem.ColComboList(col_类型) = mstrComboList
    vsItem.ColComboList(Col_页数) = "1|2|3|4|5|6|7|8|9|10"
    vsItem.ColDataType(col_必填) = flexDTBoolean
    vsItem.Editable = flexEDKbdMouse
    picAddRow.Visible = False
    Call LoadData
End Sub

Private Sub Form_Resize()
    Dim lngWidth As Long, i As Long
    On Error Resume Next
    Line1(0).X2 = Me.Width
    Line1(1).X2 = Me.Width
    Line1(2).X2 = Me.Width
    Line1(3).X2 = Me.Width
    vsItem.Width = Me.Width - 320
    vsItem.Height = Me.Height - vsItem.Top - picBottom.Height - 590
    picAddRow.Left = vsItem.Left + vsItem.Width - vsItem.ColWidth(col_插入) - 30
    For i = 0 To vsItem.Cols - 1
        If Not vsItem.ColHidden(i) And i <> col_列表 And i <> col_插入 Then lngWidth = lngWidth + vsItem.ColWidth(i)
    Next
    vsItem.ColWidth(col_列表) = vsItem.Width - lngWidth - 470
    picBottom.Top = Me.Height - vsItem.Height - vsItem.Top
    picBottom.Width = Me.Width
    cmdOK.Left = Me.Width - cmdOK.Width - 1800
    cmdCancel.Left = Me.Width - cmdCancel.Width - 500
    If vsItem.Width < 7888 Then
        picAddRow.Visible = False
    Else
        picAddRow.Visible = True
    End If
    If Me.Width < 9900 Then Me.Width = 9900
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picAddRow_Click()
    Dim i As Long
    
    If vsItem.Row = vsItem.Rows - 1 Then
        Call AddNewRow
    Else
        Call AddNewRow(vsItem.Row)
    End If
    
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngType As Long, i As Long
    '根据当前列数据设置关联列
    With vsItem
        If Col = col_类型 Then
            lngType = Val(.TextMatrix(Row, col_类型))
            
            If lngType = T3布尔型 Or lngType = T4单选项 Then
                .Cell(flexcpChecked, Row, col_必填) = 1
                .Cell(flexcpBackColor, Row, col_必填) = color_Unmodify
            ElseIf lngType = T6标题 Then
                .Cell(flexcpChecked, Row, col_必填) = 0
                .Cell(flexcpBackColor, Row, col_必填) = color_Unmodify
            Else
                .Cell(flexcpBackColor, Row, col_必填) = vbWhite
            End If
            
            '清除列表值
            If lngType <> T4单选项 And lngType <> T5多选项 Then
                If .TextMatrix(Row, col_列表) <> "" Then
                    .TextMatrix(Row, col_列表) = ""
                End If
            End If
            '多选清空项目序号
            Call FuncNoASC
        ElseIf Col = col_项目名称 Then
            .TextMatrix(Row, col_项目名称) = .TextMatrix(Row, col_项目名称)
            '多选清空项目序号
            Call FuncNoASC
        ElseIf Col = Col_页数 Or Col = col_必填 Or Col = col_通用 Then
            For i = Row + 1 To .Rows - 1
                If (.TextMatrix(i, col_项目名称) = "" Or .TextMatrix(i, col_项目名称) = .TextMatrix(Row, col_项目名称)) And Val(.TextMatrix(i, col_类型)) = T5多选项 And Val(.TextMatrix(Row, col_类型)) = T5多选项 Then
                    .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                    If Val(.TextMatrix(i, col_状态)) = 1 Then '为了简化处理，只要修改了都标记为已改，不管实际是否改变
                        .TextMatrix(i, col_状态) = 2
                    End If
                Else
                    Exit For
                End If
            Next
            If (Col = col_必填 Or Col = col_通用) And .Cell(flexcpData, Row, col_项目序号) <> "" Then
                For i = Row - 1 To 1 Step -1
                    If (.TextMatrix(i, col_项目名称) = "" Or .TextMatrix(i, col_项目名称) = .TextMatrix(Row, col_项目名称)) And Val(.TextMatrix(i, col_类型)) = T5多选项 And Val(.TextMatrix(Row, col_类型)) = T5多选项 Then
                        .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                        If Val(.TextMatrix(i, col_状态)) = 1 Then '为了简化处理，只要修改了都标记为已改，不管实际是否改变
                            .TextMatrix(i, col_状态) = 2
                        End If
                    Else
                        If Val(.TextMatrix(i, col_类型)) = T5多选项 Then
                            .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                            If Val(.TextMatrix(i, col_状态)) = 1 Then '为了简化处理，只要修改了都标记为已改，不管实际是否改变
                                .TextMatrix(i, col_状态) = 2
                            End If
                        End If
                        Exit For
                    End If
                Next
            End If
        ElseIf Col = col_列表 Then
            If Val(.TextMatrix(Row, col_类型)) = T4单选项 Then .ColComboList(col_列表) = "..."
        End If
        
        If Val(.TextMatrix(Row, col_状态)) = 1 Then '为了简化处理，只要修改了都标记为已改，不管实际是否改变
            .TextMatrix(Row, col_状态) = 2
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And NewRow > vsItem.FixedRows - 1 Then
        If Me.Visible Then
            If picAddRow.Visible = False Then picAddRow.Visible = True
        End If
        picAddRow.Top = vsItem.Cell(flexcpTop, NewRow, col_插入) + 30
        picAddRow.Left = vsItem.Left + vsItem.Cell(flexcpLeft, NewRow, col_插入) + 50
    End If
    If NewCol = col_列表 And Val(vsItem.TextMatrix(NewRow, col_类型)) = T4单选项 Then
        vsItem.ColComboList(col_列表) = "..."
    Else
        vsItem.ColComboList(col_列表) = ""
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col_必填 Then
        If vsItem.Cell(flexcpBackColor, Row, col_必填) = color_Unmodify Then
            Cancel = True
        End If
    ElseIf Col = col_列表 Then
        If Val(vsItem.TextMatrix(Row, col_类型)) <> T4单选项 And Val(vsItem.TextMatrix(Row, col_类型)) <> T5多选项 Then
            Cancel = True
        End If
    ElseIf Col = Col_页数 Then
        If vsItem.Cell(flexcpData, Row, col_项目序号) <> "" Then
            Cancel = True
        End If
    ElseIf Col = col_插入 Then
        Cancel = True
    End If
End Sub

Private Sub vsItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As Recordset, strSql As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsItem
        If Col = col_列表 And Val(.TextMatrix(Row, col_类型)) = T4单选项 Then
            '检查是否存在出径记录
            If Val(.TextMatrix(Row, col_状态)) <> 0 Then
                If CheckItemNO(Val(.TextMatrix(Row, COL_行号))) Then
                    If MsgBox("在“病人出径记录”中已存在当前行的相关数据，修改后可能会丢失数据，你确定要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            '如果有逗号则不查找字典表
            strSql = "Select Rownum As ID, 系统, 表名 From zlBaseCode"
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "公共字典表", _
                False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
            '判断是否有数据
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, Col) = "[" & rsTmp!表名 & "]"
            End If
        End If
    End With
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, lngRowNO As Long
    '删除行
    If KeyCode = vbKeyDelete Then
        With vsItem
            If .Col = col_列表 Then
                .TextMatrix(.Row, .Col) = ""
            Else
                If .Row = .FixedRows - 1 Then .Row = .FixedRows
                If .TextMatrix(.Row, col_项目名称) = "" And .Rows = 2 Then
                    MsgBox "没有可删除的项目了。", vbInformation, gstrSysName
                    Exit Sub
                End If
                If MsgBox("你确定要删除当前行吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    lngRowNO = Val(.TextMatrix(.Row, COL_行号))

                    If CheckItemNO(lngRowNO) Then
                        If MsgBox("在" & IIf(mintType = 1, "“病人门诊出径记录”", "“病人出径记录”") & "中已存在当前行的相关数据，删除后将破坏数据之间的关联，你确定要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        End If
                    End If
                    
                    If lngRowNO <> 0 Then
                        If mstrDelItem = "" Then
                            mstrDelItem = lngRowNO
                        Else
                            mstrDelItem = mstrDelItem & "," & lngRowNO
                        End If
                    End If
                    For i = .Row + 1 To .Rows - 1
                        .TextMatrix(i, col_项目序号) = Val(.TextMatrix(i, col_项目序号)) - 1
                    Next
                    .RemoveItem .Row
                    If .Rows = .FixedRows Then Call AddNewRow
                End If
            End If
        End With
    ElseIf KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsItem_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    '新增行(最后一行按回车)
    Dim i As Long
    With vsItem
        If KeyAscii = vbKeyReturn Then
            If .Row = .Rows - 1 And .Col = col_列表 Then
                Call AddNewRow
                .Select .Rows - 1, col_项目名称
            ElseIf .Col = col_列表 Then
                KeyAscii = 0
                .Select .Row + 1, col_项目名称
            Else
                KeyAscii = 0
                .Col = .Col + 1
            End If
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsItem_CellButtonClick(.Row, .Col)
            Else
                .ColComboList(col_列表) = "" '使按钮状态进入输入状态
            End If
            
        End If
    End With
End Sub

Private Function FuncNoASC() As Long
'功能：为表格从新排序,并且一组的必填和页面修改为相同
    Dim i As Long, lngNo As Long, lng序号 As Long
    
    With vsItem
        lngNo = 0
        .TextMatrix(1, col_项目序号) = "1"
        If Val(.TextMatrix(1, col_状态)) = 1 Then
            .TextMatrix(1, col_状态) = 2
        End If
        For i = 2 To .Rows - 1
            If Val(.TextMatrix(i, col_类型)) = T5多选项 Then
                If Val(.TextMatrix(i - 1, col_类型)) = T5多选项 And (.TextMatrix(i, col_项目名称) = .TextMatrix(i - 1, col_项目名称) Or .TextMatrix(i, col_项目名称) = "") Then
                    .TextMatrix(i, col_项目序号) = .TextMatrix(i - 1, col_项目序号)
                    .Cell(flexcpData, i, col_项目序号) = "1"
                    .TextMatrix(i, Col_页数) = .TextMatrix(i - 1, Col_页数)
                    .Cell(flexcpChecked, i, col_必填) = .Cell(flexcpChecked, i - 1, col_必填)
                    .TextMatrix(i, col_多选序号) = Val(.TextMatrix(i - 1, col_多选序号)) + 1
                    lngNo = lngNo + 1
                Else
                    .TextMatrix(i, col_项目序号) = i - lngNo
                    .Cell(flexcpData, i, col_项目序号) = ""
                    lng序号 = 1
                    If Val(.TextMatrix(i, col_类型)) = T5多选项 Then .TextMatrix(i, col_多选序号) = lng序号
                End If
            Else
                .TextMatrix(i, col_项目序号) = i - lngNo
                .Cell(flexcpData, i, col_项目序号) = ""
                lng序号 = 1
                If Val(.TextMatrix(i, col_类型)) = T5多选项 Then .TextMatrix(i, col_多选序号) = lng序号
            End If
            If Val(.TextMatrix(i, col_状态)) = 1 Then '为了简化处理，只要修改了都标记为已改，不管实际是否改变
                .TextMatrix(i, col_状态) = 2
            End If
        Next
    End With
End Function


Private Sub AddNewRow(Optional ByVal lngRow As Long)
'功能：新增一空白行
'参数：lngRow-0，最后一行新增，否则为插入
    With vsItem
        If lngRow = 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
        Else
            vsItem.AddItem "", lngRow
        End If
        If lngRow <> 1 Then
            .TextMatrix(lngRow, col_类型) = Split(mstrComboList, "|")(Val(.TextMatrix(lngRow - 1, col_类型)))
            .TextMatrix(lngRow, Col_页数) = .TextMatrix(lngRow - 1, Col_页数)
            .Cell(flexcpChecked, lngRow, col_必填) = .Cell(flexcpChecked, lngRow - 1, col_必填)
            .Cell(flexcpChecked, lngRow, col_通用) = .Cell(flexcpChecked, lngRow - 1, col_通用)
        Else
            .TextMatrix(lngRow, col_类型) = Split(mstrComboList, "|")(0)
        End If
        Call vsItem_AfterEdit(lngRow, col_类型)
        .TextMatrix(lngRow, col_状态) = 0
        Call FuncNoASC
    End With
End Sub

Private Function GetMaxItemNO() As Long
'功能：获取项目列表的当前最大行号
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Nvl(Max(行号),0) as 最大行号 From 路径报表结构 Where 报表ID = 2"
    On Error GoTo errH
    If mintType = 1 Then strSql = Replace(strSql, "路径报表结构", "门诊路径报表结构"): strSql = Replace(strSql, "路径报表序号", "门诊路径报表序号")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        GetMaxItemNO = rsTmp!最大行号
    Else
        GetMaxItemNO = 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckItemNO(ByVal lngRowNO As Long) As Boolean
'功能：检查当前行是否已存在病人相关数据
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    If mintType = 1 Then
        strSql = "Select 1 From 病人门诊出径记录 Where 行号 = [1] And Rownum=1"
    Else
        strSql = "Select 1 From 病人出径记录 Where 行号 = [1] And Rownum=1"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngRowNO)
    CheckItemNO = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As Recordset, strSql As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsItem
        If Col = col_列表 And Val(.TextMatrix(Row, col_类型)) = T4单选项 Or Col = col_类型 And .EditText <> .TextMatrix(Row, col_类型) Or Col = col_通用 And Val(.TextMatrix(Row, col_通用)) <> 0 Then
            '检查是否存在出径记录
            If Val(.TextMatrix(Row, col_状态)) <> 0 Then
                If CheckItemNO(Val(.TextMatrix(Row, COL_行号))) Then
                    If MsgBox("在" & IIf(mintType = 1, "“病人门诊出径记录”", "“病人出径记录”") & "中已存在当前行的相关数据，修改后可能会丢失数据，你确定要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True
                        If Col = col_类型 Then .EditText = .TextMatrix(Row, col_类型)
                        Exit Sub
                    End If
                End If
            End If
        End If
        If Col = col_列表 And Val(.TextMatrix(Row, col_类型)) = T4单选项 Then
            '如果有逗号则不查找字典表
            If .EditText = "" Then Exit Sub
            If InStr(.EditText, ",") = 0 Then
                strSql = "Select Rownum As ID, 系统, 表名 From zlBaseCode Where 表名 Like [1]"
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                On Error GoTo errH
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "公共字典表", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    gstrLike & .EditText & "%")
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    '判断是否有数据
                    If Not rsTmp Is Nothing Then
                        .EditText = "[" & rsTmp!表名 & "]"
                    End If
                End If
            End If
            If Mid(.EditText, 1, 1) = "[" And Mid(.EditText, Len(.EditText)) = "]" Then
                strSql = "Select Count(1) as 存在 From zlBaseCode Where 表名=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Mid(.EditText, 2, Len(.EditText) - 2))
                If Val(rsTmp!存在 & "") = 0 Then
                    MsgBox "没有这个找到这个字典表：" & Mid(.EditText, 2, Len(.EditText) - 2)
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
