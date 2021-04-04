VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmPathOutLog 
   Caption         =   "病人出径登记"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12015
   Icon            =   "frmPathOutLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   12015
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   12000
      TabIndex        =   6
      Top             =   840
      Width           =   12000
      Begin VSFlex8Ctl.VSFlexGrid vsItem 
         Height          =   7410
         Left            =   0
         TabIndex        =   7
         Top             =   0
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
         Rows            =   7
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathOutLog.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
         OwnerDraw       =   1
         Editable        =   2
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
      End
   End
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
      Begin VB.CommandButton cmdPrintToEXCEL 
         Caption         =   "输出到EXCEL"
         Height          =   350
         Left            =   7800
         TabIndex        =   8
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10800
         TabIndex        =   4
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9600
         TabIndex        =   3
         Top             =   200
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11880
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   12000
         Y1              =   30
         Y2              =   30
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
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "请填写以下要求登记的信息。浅绿色背景的单元格为必填项，日期请按YYYY-MM-DD格式录入。提交病案审查后将不允许再修改。"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   7455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   12000
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathOutLog.frx":6990
         Top             =   45
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   8895
      _Version        =   589884
      _ExtentX        =   15690
      _ExtentY        =   8916
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPathOutLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun As Long '0-新增，1-查看，2-修改
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng路径ID As Long
Private mlng病人路径ID As Long
Private mcolSQL As New Collection

Private mblnOK As Boolean
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
    col_项目值 = 2
    col_备注 = 3
    
    '隐藏列
    col_类型 = 4
    col_必填 = 5
    col_状态 = 6    '1-原始，2-修改
    col_行号 = 7    '行ID
    Col_页数 = 8
End Enum
Private Const Color_MustAddBack = &HE1FFE1

Public Function ShowMe(frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngFun As Long, ByRef colSQL As Collection, _
                    ByVal lng路径ID As Long, Optional ByVal lng病人路径Id As Long) As Boolean
'参数： lngFun=0-新增，1-查看，2-修改
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlngFun = lngFun
    mlng路径ID = lng路径ID
    mlng病人路径ID = lng病人路径Id
    
    If mlngFun = 1 Then
        If CheckPatiPathOutLog(lng病人ID, lng主页ID) = False Then
            MsgBox "未登记该病人的出径信息。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Set mcolSQL = Nothing
    
    Me.Show 1, frmMain
    
    Set colSQL = mcolSQL
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strMsg As String, intType As Integer
    Dim strValue As String
    Dim blnIsCheck As Boolean   '判断多选必须的情况
    Dim j As Long
    
    With vsItem
        For i = .FixedRows To .Rows - 1
            intType = Val(.TextMatrix(i, col_类型))
            strValue = Trim(.TextMatrix(i, col_项目值))
            If intType = T0数字 Then
                If strValue <> "" And Not IsNumeric(strValue) Then
                    strMsg = "要求输入的内容必须是数字型。"
                    Exit For
                End If
            ElseIf intType = T2日期 Then
                If strValue <> "" And Not IsDate(strValue) Then
                    strMsg = "要求输入的内容必须是日期型。"
                    Exit For
                End If
            ElseIf intType = T1字符 Then
                If strValue <> "" Then
                    If zlCommFun.ActualLen(strValue) >= 100 Then
                        strMsg = "项目值最多允许输入100个字符或50个汉字。"
                        Exit For
                    End If
                End If
            End If
            If intType = T0数字 Or intType = T1字符 Or intType = T2日期 Or intType = T4单选项 Then
                If Val(.TextMatrix(i, col_必填)) = 1 Then
                    If strValue = "" Then
                        strMsg = "要求必须填写内容。"
                        Exit For
                    End If
                End If
            End If
            
            If intType = T5多选项 And Val(.TextMatrix(i, col_必填)) = 1 Then
                For j = i To .Rows - 1
                    If .TextMatrix(j, col_项目序号) <> .TextMatrix(i, col_项目序号) Then Exit For
                    If .Cell(flexcpChecked, j, col_项目值) = 1 Then blnIsCheck = True
                Next
                If Not blnIsCheck Then
                    strMsg = "要求必须填写内容。"
                    Exit For
                End If
                blnIsCheck = False
                i = j - 1
            End If
            
            If .TextMatrix(i, col_备注) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, col_备注)) >= 1000 Then
                    strMsg = "备注内容最多允许输入1000个字符或500个汉字。"
                    Exit For
                End If
            End If
        Next
        If i <= .Rows - 1 Then
            tbcSub.Item(Val(.TextMatrix(i, Col_页数)) - 1).Selected = True
            MsgBox "第" & .TextMatrix(i, col_项目序号) & "号项目，" & strMsg, vbInformation, gstrSysName
            .Select i, col_项目值
            .SetFocus
            Exit Sub
        End If
        
        If SaveData = False Then
            Exit Sub
        End If
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Function SaveData() As Boolean
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSql As String, i As Long, intType As Integer
    Dim strDate As String, str数字值 As String, str字符值 As String, str日期值 As String, strValue As String

    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"

    With vsItem
        For i = .FixedRows To .Rows - 1
            str数字值 = "Null"
            str字符值 = "Null"
            str日期值 = "Null"
            strSql = ""

            strValue = Trim(.TextMatrix(i, col_项目值))
            intType = Val(.TextMatrix(i, col_类型))
            Select Case intType
            Case T0数字
                str数字值 = strValue
            Case T3布尔型, T5多选项
                str数字值 = IIf(.Cell(flexcpChecked, i, col_项目值) = 1, 1, 0)
            Case T1字符, T4单选项
                str字符值 = "'" & strValue & "'"
            Case T2日期
                If strValue <> "" Then
                    str日期值 = "To_Date('" & Format(strValue, "yyyy-MM-DD") & "','yyyy-mm-dd')"
                End If
            End Select

            If mlngFun = 0 Then
                '新增(未填写的行不保存)
                If intType = T3布尔型 Or strValue <> "" Or Trim(.TextMatrix(i, col_备注)) <> "" Then
                    strSql = "0"
                End If
            Else
                '修改
                If Val(.TextMatrix(i, col_状态)) = 2 Then
                    strSql = "1"
                End If
            End If

            If strSql <> "" Then
                strSql = "Zl_病人出径记录_Update(" & strSql & "," & mlng病人ID & "," & mlng主页ID & "," & .TextMatrix(i, col_行号) & _
                         "," & str数字值 & "," & str字符值 & "," & str日期值 & ",'" & Trim(.TextMatrix(i, col_备注)) & "','" & UserInfo.姓名 & "'," & strDate & "," & mlng病人路径ID & ")"
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        Next
    End With

    Set mcolSQL = colSQL

    SaveData = True
End Function

Private Sub cmdPrintToEXCEL_Click()
'功能:输出到EXCEL
    Dim objReport As ReportControl
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim rsTmp As Recordset, strSql As String
    
    On Error GoTo errH
    
    strSql = "Select NVL(B.姓名,A.姓名) 姓名 From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID)
    vsItem.ColWidth(col_备注) = 3950
    Set objPrint.Body = Me.vsItem
    vsItem.AutoSize vsItem.FixedCols, vsItem.Cols - 1, , 45 '高度自适应
    objPrint.Title.Text = "出径登记表"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("病人：" & rsTmp!姓名)
    strSql = "Select 名称 From 临床路径目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
    Call objAppRow.Add("路径：" & rsTmp!名称)
    Call objPrint.UnderAppRows.Add(objAppRow)
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    zlPrintOrView1Grd objPrint, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '不允许输入分隔符及单引号
    End If
End Sub

Private Sub Form_Load()
    Dim lngPage As Long, i As Long
    
    Call RestoreWinState(Me, App.ProductName, mlngFun)
    
    mblnOK = False
    For i = 0 To vsItem.Cols - 1
        If vsItem.ColHidden(i) Then vsItem.ColWidth(i) = 0
    Next
    If mlngFun = 1 Then
        vsItem.Editable = flexEDNone
        cmdOK.Visible = False
        cmdCancel.Caption = "退出(&X)"
    End If
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .ColorSet.HeaderFaceLight = &HF0F4E4
            .ColorSet.HeaderFaceDark = &HF0F4E4
        End With
        '如果只有一页则隐藏选项卡头
        lngPage = CheckPage
        If lngPage > 0 Then
            For i = 0 To lngPage - 1
                .InsertItem(i, "第" & i + 1 & "页", picItem.Hwnd, 0).Tag = i + 1
            Next
            .Item(0).Selected = True
            Call tbcSub_SelectedChanged(.Item(0))
        End If
        If lngPage <= 1 Then
            .PaintManager.HeaderMargin.Top = -20   '不显示Tab标题区域
        End If
    End With
    
    '绑定后在加载，否则会出现显示不出的问题
    Call LoadData
End Sub

Private Function CheckPage() As Long
'返回：页数
    Dim strSql As String
    Dim rsTmp As Recordset
    
    strSql = "Select Count(Distinct NVL(页数,1)) as 页数 From 路径报表结构 Where 报表id = 2 And 路径id Is Null Or 路径id = [1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
    CheckPage = Val(rsTmp!页数 & "")
    
End Function

Private Sub LoadData()
    Dim i As Long, arrtmp As Variant, intType As Integer
    Dim strSql As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    If mlngFun = 1 Or mlngFun = 2 Then
        strSql = "Select a.行号, NVL(a.页数,1)  as 页数, a.项目序号, a.项目文本1, a.项目文本2, a.路径id, a.多选序号, b.数字值, b.字符值, b.日期值, b.备注" & vbNewLine & _
                "From (Select a.行号, a.页数, Nvl(b.序号, a.项目序号) As 项目序号, a.项目文本1, a.项目文本2, a.路径id, a.多选序号" & vbNewLine & _
                "       From 路径报表结构 A, 路径报表序号 B" & vbNewLine & _
                "       Where a.报表id = b.报表id(+) And a.行号 = b.行号(+) And a.报表id = 2 And" & vbNewLine & _
                "             (Nvl(a.路径id, b.路径id) = [3] And (Exists (Select 1 From 路径报表序号 Where 报表ID = 2  And 路径id = [3]) Or Not Exists (Select 1 From 路径报表结构 Where 报表id = 2 And a.路径id Is Null)))" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.行号, NVL(a.页数,1)  as 页数, a.项目序号, a.项目文本1, a.项目文本2, a.路径id, a.多选序号" & vbNewLine & _
                "From 路径报表结构 A" & vbNewLine & _
                "Where a.报表id = 2 And a.路径id Is Null And Not Exists (Select 1 From 路径报表序号 Where 报表id = 2 And 路径id = [3])) A, 病人出径记录 B" & vbNewLine & _
                "Where a.行号 = b.行号(+) And b.病人id(+) = [1] And b.主页id(+) = [2] And B.路径记录ID(+)=[4]" & vbNewLine & _
                "Order By 项目序号, 多选序号"
 
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, mlng路径ID, mlng病人路径ID)
    Else
        strSql = "Select a.行号, NVL(a.页数,1) as 页数, Nvl(b.序号, a.项目序号) As 项目序号, a.项目文本1, a.项目文本2, a.路径id, a.多选序号" & vbNewLine & _
                "From 路径报表结构 A, 路径报表序号 B" & vbNewLine & _
                "Where a.报表id = b.报表id(+) And a.行号 = b.行号(+) And a.报表id = 2 And" & vbNewLine & _
                "      (Nvl(a.路径id, b.路径id) = [1] And (Exists (Select 1 From 路径报表序号 Where 报表id = 2 And 路径id = [1]) Or Not Exists (Select 1 From 路径报表结构 Where 报表id = 2 And a.路径id Is Null)))" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.行号, NVL(a.页数,1)  as 页数, a.项目序号, a.项目文本1, a.项目文本2, a.路径id, a.多选序号" & vbNewLine & _
                "From 路径报表结构 A" & vbNewLine & _
                "Where a.报表id = 2 And a.路径id Is Null And Not Exists (Select 1 From 路径报表序号 Where 报表id = 2 And 路径id = [1])" & vbNewLine & _
                "Order By 项目序号, 多选序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
    End If
    With vsItem
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, col_项目序号) = rsTmp!项目序号 & ""
            .TextMatrix(i, col_项目名称) = rsTmp!项目文本1 & ""
            .TextMatrix(i, col_状态) = "1"
            .TextMatrix(i, col_行号) = rsTmp!行号 & ""
            .TextMatrix(i, Col_页数) = rsTmp!页数 & ""
            
            arrtmp = Split(rsTmp!项目文本2, "|")    '=类型|是否必填或选项1,选项2,...
            intType = arrtmp(0)
            .TextMatrix(i, col_类型) = intType
            
            If intType = T3布尔型 Or intType = T4单选项 Then
                .TextMatrix(i, col_必填) = 1
                .Cell(flexcpPictureAlignment, i, col_项目值) = flexPicAlignCenterCenter
                .Cell(flexcpAlignment, i, col_项目值) = flexAlignCenterCenter
                .Cell(flexcpBackColor, i, col_项目值) = Color_MustAddBack
            ElseIf UBound(arrtmp) > 0 Then  '数字，字符，日期
                .TextMatrix(i, col_必填) = arrtmp(1)
                If Val(arrtmp(1)) = 1 Then
                    .Cell(flexcpBackColor, i, col_项目值) = Color_MustAddBack
                End If
                If intType = T5多选项 Then
                    .TextMatrix(i, col_备注) = arrtmp(2)
                    .Cell(flexcpPictureAlignment, i, col_项目值) = flexPicAlignCenterCenter
                    .Cell(flexcpAlignment, i, col_项目值) = flexAlignCenterCenter
                End If
            Else
                .TextMatrix(i, col_必填) = 0
            End If
            
            If intType = T4单选项 Then
                If UBound(arrtmp) > 0 Then
                    .RowData(i) = CStr(Replace(arrtmp(1), ",", "|"))   '作为ColComboList的值
                    .TextMatrix(i, col_项目值) = Split(arrtmp(1), ",")(0)   '第一项作为缺省值
                    If .TextMatrix(i, col_项目值) <> "" Then
                        If Mid(.TextMatrix(i, col_项目值), 1, 1) = "[" And Mid(.TextMatrix(i, col_项目值), Len(.TextMatrix(i, col_项目值))) = "]" Then
                            .Cell(flexcpData, i, col_项目值) = Mid(.TextMatrix(i, col_项目值), 2, InStr(.TextMatrix(i, col_项目值), "]") - 2)
                            .TextMatrix(i, col_项目值) = ""
                            .RowData(i) = "" '当类型为“T4单选项”时，将会作为下拉方式或者绑定方式的判断依据
                        End If
                    End If
                End If
            ElseIf intType = T3布尔型 Or intType = T5多选项 Then
                .Cell(flexcpChecked, i, col_项目值) = 2
            ElseIf intType = T6标题 Then
                .TextMatrix(i, col_项目值) = .TextMatrix(i, col_项目名称)
                .TextMatrix(i, col_备注) = .TextMatrix(i, col_项目名称)
                .MergeRow(i) = True
            End If
            
            
            If (mlngFun = 1 Or mlngFun = 2) And intType <> T6标题 Then
                Select Case intType
                Case T0数字
                    .TextMatrix(i, col_项目值) = "" & rsTmp!数字值
                Case T3布尔型, T5多选项
                    .Cell(flexcpChecked, i, col_项目值) = IIf(Val("" & rsTmp!数字值) = 1, 1, 2)
                Case T1字符, T4单选项
                    .TextMatrix(i, col_项目值) = "" & rsTmp!字符值
                Case T2日期
                    If Not IsNull(rsTmp!日期值) Then
                        .TextMatrix(i, col_项目值) = Format(rsTmp!日期值 & "", "yyyy-MM-DD")
                    End If
                End Select
                
                If intType <> T5多选项 Then .TextMatrix(i, col_备注) = "" & rsTmp!备注
                
                '保存原值，用于判断是否修改
                If mlngFun = 2 Then
                    .Cell(flexcpData, i, col_备注) = "" & .TextMatrix(i, col_备注)
                    If .Cell(flexcpData, i, col_项目值) = "" Then .Cell(flexcpData, i, col_项目值) = "" & .TextMatrix(i, col_项目值)
                End If
            End If
            
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Line1(0).X2 = Me.Width
    Line1(1).X2 = Me.Width
    Line1(2).X2 = Me.Width
    Line1(3).X2 = Me.Width
    tbcSub.Top = picInfo.Height
    tbcSub.Left = 20
    picBottom.Top = Me.Height - tbcSub.Height - tbcSub.Top
    tbcSub.Width = Me.Width - 270
    tbcSub.Height = Me.Height - tbcSub.Top - picBottom.Height - 590
    picBottom.Width = Me.Width
    cmdOK.Left = Me.Width - cmdOK.Width - 1800
    cmdCancel.Left = Me.Width - cmdCancel.Width - 500
    cmdPrintToEXCEL.Left = Me.Width - cmdPrintToEXCEL.Width - 3000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mlngFun)
End Sub


Private Sub picItem_Resize()
    On Error Resume Next
    vsItem.Move 0, 0, picItem.Width, picItem.Height
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        If Val(vsItem.TextMatrix(i, Col_页数)) = Val(Item.Tag & "") Then
            vsItem.RowHidden(i) = False
        Else
            vsItem.RowHidden(i) = True
        End If
    Next
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col_备注 Or Col = col_项目值 Then
        With vsItem
            If CStr(.Cell(flexcpData, Row, Col)) <> CStr(.TextMatrix(Row, Col)) Or .TextMatrix(Row, col_类型) = T3布尔型 Or .TextMatrix(Row, col_类型) = T5多选项 Then
                .TextMatrix(Row, col_状态) = 2
            Else
                .TextMatrix(Row, col_状态) = 1
            End If
            If .TextMatrix(Row, col_类型) = T4单选项 And Col = col_项目值 And .RowData(Row) = "" Then
                .ColComboList(col_项目值) = "..."
            End If
        End With
    End If
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And NewRow >= vsItem.FixedRows Then
        If vsItem.TextMatrix(NewRow, col_类型) = T4单选项 Then
            If vsItem.RowData(NewRow) = "" Then   '绑定数据源方式
                vsItem.ColComboList(col_项目值) = "..."
            Else '下拉方式
                vsItem.ColComboList(col_项目值) = vsItem.RowData(NewRow)
            End If
        Else
            vsItem.ColComboList(col_项目值) = ""
        End If
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col_项目值 Or Col = col_备注) Then
        Cancel = True
    Else
        If vsItem.TextMatrix(Row, col_类型) = "6" Then
            Cancel = True
        End If
        If Col = col_备注 And vsItem.TextMatrix(Row, col_类型) = "5" Then
            Cancel = True
        End If
            
    End If
End Sub

Private Sub vsItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strTable As String, strSql As String
    Dim rsTmp As Recordset
    Dim vPoint As POINTAPI, blnCancel As Boolean

    With vsItem
        If Col = col_项目值 Then
            strTable = .Cell(flexcpData, Row, Col)
            If strTable <> "" Then
                strSql = "Select Rownum as ID,名称 From " & strTable
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                On Error GoTo errH
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTable, True, "", "", True, True, True, _
                                                     vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有找到 " & strTable & " 的数据。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = "-"
                    End If
                    Exit Sub
                Else
                    .TextMatrix(Row, Col) = rsTmp!名称 & ""
                    If CStr(.Cell(flexcpData, Row, Col)) <> CStr(.TextMatrix(Row, Col)) Then
                        .TextMatrix(Row, col_状态) = 2  '直接双击选择时不会触发vsItem_AfterEdit
                    End If
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

Private Sub vsItem_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsItem
        If Val(.TextMatrix(Row, col_类型)) = T6标题 And Col > 0 And Col < col_备注 Then
            vRect.Left = Right - 2
            vRect.Right = Right
            vRect.Top = Top
            vRect.Bottom = Bottom - 1
        Else
            lngLeft = col_项目序号: lngRight = col_项目名称
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            If Not RowIn一组项目(Row, lngBegin, lngEnd) Then Exit Sub
    
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
        End If
      

        If Between(Row, .Row, .RowSel) Then
            'SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowIn一组项目(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    
    With vsItem
        If .TextMatrix(lngRow, col_项目序号) = "" Then Exit Function
        If lngRow = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col_项目序号)) = Val(.TextMatrix(lngRow, col_项目序号)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col_项目序号)) = Val(.TextMatrix(lngRow, col_项目序号)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col_项目序号)) = Val(.TextMatrix(lngRow, col_项目序号)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col_项目序号)) = Val(.TextMatrix(lngRow, col_项目序号)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一组项目 = blnTmp
    End With
End Function

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsItem_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsItem
            If .Row = .Rows - 1 And .Col = col_备注 Then
                Call zlCommFun.PressKey(vbKeyTab)
            ElseIf .Col = col_备注 Then
                KeyAscii = 0
                .Select .Row + 1, col_项目值
            Else
                KeyAscii = 0
                .Col = .Col + 1
            End If
        End With
    Else
        If KeyAscii = Asc("*") Then
            KeyAscii = 0
            Call vsItem_CellButtonClick(vsItem.Row, vsItem.Col)
        Else
            vsItem.ColComboList(col_项目值) = "" '使按钮状态进入输入状态
        End If
    End If
End Sub

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：根据行要求的类型检查输入数据的有效性
    Dim intType As Integer, blnValidate As Boolean, strMsg As String
    Dim strValue As String
    Dim strTable As String, strSql As String
    Dim rsTmp As Recordset
    Dim vPoint As POINTAPI, blnCancel As Boolean

    With vsItem
        If Col = col_项目值 Then
            blnValidate = True
            intType = Val(vsItem.TextMatrix(Row, col_类型))
            strValue = vsItem.EditText

            If strValue <> "" Then
                Select Case intType
                Case T0数字
                    blnValidate = IsNumeric(strValue)
                    strMsg = "要求输入的内容必须是数字型。"
                Case T2日期
                    blnValidate = IsDate(strValue)
                    strMsg = "要求输入的内容必须是日期型。"
                Case T4单选项
                    If .RowData(Row) = "" Then
                        strTable = .Cell(flexcpData, Row, Col)
                        If strTable <> "" Then
                            strSql = "Select Rownum as ID,名称 From " & strTable & " Where 名称 Like [1] Or Upper(zlspellcode(名称)) like [2]"
                            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                            On Error GoTo errH
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTable, True, "", "", True, True, True, _
                                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, gstrLike & strValue & "%", UCase(strValue) & "%")
                            If rsTmp Is Nothing Then
                                If Not blnCancel Then
                                    strMsg = "没有查找到指定的数据。"
                                    blnValidate = False
                                Else
                                    Cancel = True
                                End If
                            Else
                                .EditText = rsTmp!名称 & ""
                                .TextMatrix(Row, Col) = rsTmp!名称 & ""
                            End If
                        End If
                    End If
                End Select
            End If
            If blnValidate = False Then
                MsgBox strMsg, vbInformation, gstrSysName
                Cancel = True
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



