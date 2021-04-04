VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl usrInSpecResult 
   BackColor       =   &H80000005&
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LockControls    =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   8130
   Begin VB.PictureBox PicItem 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   8130
      TabIndex        =   7
      Top             =   0
      Width           =   8130
      Begin VB.CommandButton cmdP1 
         Caption         =   "&P"
         Height          =   300
         Left            =   4230
         TabIndex        =   6
         ToolTipText     =   "选择标本"
         Top             =   -15
         Width           =   315
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   1290
         TabIndex        =   5
         Top             =   0
         Width           =   2955
      End
      Begin VB.CommandButton cmdP 
         Caption         =   "…"
         Height          =   285
         Left            =   6825
         TabIndex        =   10
         ToolTipText     =   "选择标本"
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lblBBCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "标本(&B)"
         Height          =   180
         Left            =   4725
         TabIndex        =   8
         Top             =   45
         Width           =   630
      End
      Begin VB.Label lblBB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标本名称"
         Height          =   180
         Left            =   5445
         TabIndex        =   9
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检验项目(&C)"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   45
         Width           =   990
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   8130
         Y1              =   315
         Y2              =   315
      End
   End
   Begin VB.ListBox listCell 
      Height          =   1110
      Left            =   765
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   570
      Width           =   3075
   End
   Begin VB.ComboBox CmbCell 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   300
      Width           =   810
   End
   Begin VB.TextBox txtCell 
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   345
      Width           =   810
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfMain 
      Height          =   1530
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   2699
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   11
      FixedCols       =   0
      BackColorSel    =   -2147483639
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483631
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
End
Attribute VB_Name = "usrInSpecResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const LAWLChar = "';`|,"""
Private i As Long, j As Long
Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private Enum EnmCTLType
    CTLTxt = 0  '-文本
    CTLUpDown = 1 '-上下,
    CTLDownList = 2 '-下拉,
    CTLCheck = 3 '-复选,
    CTLOption = 4   '-单选
End Enum
Private Enum EnmValType
    ValNumber = 0   '数值型
    ValText = 1     '文本型
    ValDate = 2     '日期型
End Enum

Private Enum EnmGridCol
    ItemID = 0
    Item类型 = 1
    Item表示法 = 2
    Item数值域 = 3
    Item初始值 = 4
    Item长度 = 5
    Item小数长 = 6
    Item行号 = 7
    Item指标名 = 8
    Item英文 = 9
    Item正常值 = 10
    Item所见内容 = 11
    Item单位 = 12
    Item中文 = 13
End Enum

Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
Private mID诊疗项目 As Long
Private mblnCancel As Boolean
Private mlng病历id As Long
Private mShowItem As Boolean

Private mblnLawless As Boolean
Private mblnFirst As Boolean

Private mItemIndex As Long

Private Function zlGetSymbol(StrInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & StrInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & StrInput & "') from dual"
    End If
    On Error GoTo ErrHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Property Get ShowItem() As Boolean
'显示检验项目,此是用户可以自己改变检验项目
    ShowItem = mShowItem
End Property

Public Property Let ShowItem(ByVal New_ShowItem As Boolean)
    mShowItem = New_ShowItem
    If mShowItem = True Then
        PicItem.Height = 510
        PicItem.Visible = True
    Else
        PicItem.Height = 0
        PicItem.Visible = False
    End If
    UserControl_Resize
    PropertyChanged "ShowItem"
End Property

Public Property Get ID诊疗项目() As Long
    ID诊疗项目 = mID诊疗项目
End Property

Public Property Let ID诊疗项目(ByVal New_ID诊疗项目 As Long)
'设置诊疗项目
On Error GoTo ErrHandle
Dim lngWidth As Long
Dim lngWidth单位 As Long
Dim lngWidth行号 As Long

Dim rs诊疗项目 As New ADODB.Recordset

    txtCell.Visible = False
    CmbCell.Visible = False
    listCell.Visible = False
    lngWidth = 800
    
    '初始化控件
    mblnCancel = True
    InitMe
    mblnCancel = False
    mID诊疗项目 = New_ID诊疗项目
    PropertyChanged "ID诊疗项目"
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Property
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Property
        
    strSQL = _
        "SELECT C.ID," & vbCrLf & _
        "       A.排列序号 序号," & vbCrLf & _
        "       A.检验标本," & vbCrLf & _
        "       C.类型," & vbCrLf & _
        "       C.表示法," & vbCrLf & _
        "       C.数值域," & vbCrLf & _
        "       C.初始值," & vbCrLf & _
        "       C.长度," & vbCrLf & _
        "       C.小数," & vbCrLf & _
        "       C.中文名 指标名," & vbCrLf & _
        "       C.英文名 英文名," & vbCrLf & _
        "       C.单位" & vbCrLf & _
        "  FROM 诊疗项目目录 B, 检验报告项目 A,诊治所见项目 C" & vbCrLf & _
        " WHERE B.ID IN (SELECT DISTINCT 诊疗项目ID FROM 检验报告项目) AND  A.报告项目id=C.Id AND " & vbCrLf & _
        "      B.标本部位 = A.检验标本 AND B.ID = A.诊疗项目ID  AND A.诊疗项目ID =" & New_ID诊疗项目 & vbCrLf & _
        " ORDER BY A.排列序号"
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "检验结果报告")
    '读出检验项目的检验指标
    If rsTmp.RecordCount > 0 Then
        mblnCancel = True
        rsTmp.MoveFirst
        '读出
        msfMain.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            msfMain.TextMatrix(i, EnmGridCol.ItemID) = zlCommFun.Nvl(rsTmp!ID, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item类型) = zlCommFun.Nvl(rsTmp!类型, 1)
            msfMain.TextMatrix(i, EnmGridCol.Item表示法) = zlCommFun.Nvl(rsTmp!表示法, 0)
            If zlCommFun.Nvl(rsTmp!类型, 1) = 0 Then
                If Trim(zlCommFun.Nvl(rsTmp!数值域)) <> ";" Then
                    msfMain.TextMatrix(i, EnmGridCol.Item正常值) = Replace(Trim(zlCommFun.Nvl(rsTmp!数值域)), ";", " ～ ")
                End If
            End If
            msfMain.TextMatrix(i, EnmGridCol.Item数值域) = Trim(zlCommFun.Nvl(rsTmp!数值域))
            msfMain.TextMatrix(i, EnmGridCol.Item初始值) = Trim(zlCommFun.Nvl(rsTmp!初始值))
            msfMain.TextMatrix(i, EnmGridCol.Item长度) = zlCommFun.Nvl(rsTmp!长度, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item小数长) = zlCommFun.Nvl(rsTmp!小数, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item指标名) = Trim(zlCommFun.Nvl(rsTmp!指标名)) & IIf(Trim(zlCommFun.Nvl(rsTmp!英文名)) = "", "", "［" & Trim(zlCommFun.Nvl(rsTmp!英文名)) & "］")
            msfMain.TextMatrix(i, EnmGridCol.Item英文) = Trim(zlCommFun.Nvl(rsTmp!英文名))
            msfMain.TextMatrix(i, EnmGridCol.Item所见内容) = Trim(zlCommFun.Nvl(rsTmp!初始值))
            msfMain.TextMatrix(i, EnmGridCol.Item单位) = Trim(zlCommFun.Nvl(rsTmp!单位))
            If i = 1 Then lblBB.Caption = zlCommFun.Nvl(rsTmp!检验标本)
            rsTmp.MoveNext
        Next
        strSQL = "select * from 诊疗项目目录  where id=" & mID诊疗项目
        Call zlDatabase.OpenRecordset(rs诊疗项目, strSQL, "检验结果报告")
        If rs诊疗项目.RecordCount > 0 Then
            txtItem.Text = zlCommFun.Nvl(rs诊疗项目!名称)
            txtItem.SelStart = Len(txtItem.Text)
            txtItem.Tag = zlCommFun.Nvl(rs诊疗项目!名称)
            cmdP1.Tag = rs诊疗项目!ID
        Else
            txtItem.Text = ""
            txtItem.Tag = ""
            cmdP1.Tag = 0
        End If
        ReSetRowCode msfMain
        PicItem_Resize
        If rs诊疗项目.RecordCount > 0 Then
            On Error Resume Next
            If msfMain.Enabled And msfMain.Visible Then
                msfMain.SetFocus
            End If
        End If
        mblnCancel = False
    Else
        lblBB.Caption = ""
        txtItem.Text = ""
        txtItem.Tag = ""
    End If
    UserControl_Resize
    mblnCancel = False
    Exit Property
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Property

Private Sub ReadData(lng病历ID As Long)
'功能:根据指定ID读取指标
On Error GoTo ErrHandle
Dim rs诊疗项目 As New ADODB.Recordset
Dim rs诊疗项目1 As New ADODB.Recordset
Dim strItemName As String   '用来保存诊疗项目名称
Dim lngWidth As Long

    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Sub
    
    txtCell.Visible = False
    CmbCell.Visible = False
    listCell.Visible = False
    '先根据病历ID读出数据如果没有数据再根据项目ID读出指标.
    strSQL = _
        "SELECT A.所见项ID ID,A.控件号 序号, " & vbCrLf & _
        "           B.类型, " & vbCrLf & _
        "           B.表示法, " & vbCrLf & _
        "           B.数值域, " & vbCrLf & _
        "           B.初始值, " & vbCrLf & _
        "           B.长度, " & vbCrLf & _
        "           B.小数, " & vbCrLf & _
        "           B.中文名 指标名," & vbCrLf & _
        "           B.英文名 英文名," & vbCrLf & _
        "           A.所见内容,A.合并号," & vbCrLf & _
        "           A.计量单位 单位 " & vbCrLf & _
        " FROM 病人病历所见单 a,诊治所见项目 b,检验报告项目 c   " & vbCrLf & _
        " WHERE a.所见项ID(+)=B.ID AND c.报告项目id=b.id " & vbCrLf & _
        "   AND nvl(a.合并号,0)=c.诊疗项目id " & vbCrLf & _
        "   AND a.病历ID=" & lng病历ID & vbCrLf & _
        " ORDER BY c.排列序号"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "检验结果报告")
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        '如果有数据就读出数据,
        msfMain.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            msfMain.TextMatrix(i, EnmGridCol.ItemID) = zlCommFun.Nvl(rsTmp!ID, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item类型) = zlCommFun.Nvl(rsTmp!类型, 1)
            msfMain.TextMatrix(i, EnmGridCol.Item表示法) = zlCommFun.Nvl(rsTmp!表示法, 0)
            If zlCommFun.Nvl(rsTmp!类型, 1) = 0 Then
                If Trim(zlCommFun.Nvl(rsTmp!数值域)) <> ";" Then
                    msfMain.TextMatrix(i, EnmGridCol.Item正常值) = Replace(Trim(zlCommFun.Nvl(rsTmp!数值域)), ";", " ～ ")
                End If
            End If
            msfMain.TextMatrix(i, EnmGridCol.Item数值域) = Trim(zlCommFun.Nvl(rsTmp!数值域))
            msfMain.TextMatrix(i, EnmGridCol.Item初始值) = Trim(zlCommFun.Nvl(rsTmp!初始值))
            msfMain.TextMatrix(i, EnmGridCol.Item长度) = zlCommFun.Nvl(rsTmp!长度, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item小数长) = zlCommFun.Nvl(rsTmp!小数, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item指标名) = Trim(zlCommFun.Nvl(rsTmp!指标名)) & IIf(Trim(zlCommFun.Nvl(rsTmp!英文名)) = "", "", "［" & Trim(zlCommFun.Nvl(rsTmp!英文名)) & "］")
            msfMain.TextMatrix(i, EnmGridCol.Item英文) = Trim(zlCommFun.Nvl(rsTmp!英文名))
            msfMain.TextMatrix(i, EnmGridCol.Item所见内容) = zlCommFun.Nvl(rsTmp!所见内容)
            msfMain.TextMatrix(i, EnmGridCol.Item单位) = Trim(zlCommFun.Nvl(rsTmp!单位))
            If i = 1 Then
                mID诊疗项目 = zlCommFun.Nvl(rsTmp!合并号, 0)
                cmdP1.Tag = mID诊疗项目
            End If
            '重新调整指标名的宽度
            If UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item指标名)) > lngWidth Then
                lngWidth = UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item指标名))
            End If
            rsTmp.MoveNext
        Next
        '重新调整指标名的宽度
        If msfMain.ColWidth(EnmGridCol.Item指标名) < lngWidth Then
            msfMain.ColWidth(EnmGridCol.Item指标名) = lngWidth
        End If
        i = msfMain.Width - (msfMain.ColWidth(EnmGridCol.Item行号) + msfMain.ColWidth(EnmGridCol.Item指标名) + msfMain.ColWidth(EnmGridCol.Item单位)) - Screen.TwipsPerPixelX * 6
        msfMain.ColWidth(EnmGridCol.Item所见内容) = IIf(i < 200, 200, i)
        
        '读出诊疗项目以前的名称
        strSQL = "select * from 病人病历所见单 where 控件号 in (-2,-1) and 合并号=" & mID诊疗项目 & " and 病历ID=" & lng病历ID
        Call zlDatabase.OpenRecordset(rs诊疗项目, strSQL, "检验结果报告")
        rs诊疗项目.Filter = "控件号=-2"
        If rs诊疗项目.RecordCount > 0 Then
            txtItem.Tag = zlCommFun.Nvl(rs诊疗项目!标题)
            '再次检查是否存在该项目并读名称显示
            strSQL = "select * from 诊疗项目目录  where id=" & mID诊疗项目
            If rs诊疗项目1.State = adStateOpen Then rs诊疗项目1.Close
            Set rs诊疗项目1 = Nothing
            Call zlDatabase.OpenRecordset(rs诊疗项目1, strSQL, "检验结果报告")
            If rs诊疗项目1.RecordCount > 0 Then
                txtItem.Text = zlCommFun.Nvl(rs诊疗项目1!名称)
            Else
                '如果没有就初始化
                InitMe
                txtItem.Text = ""
                txtItem.Tag = ""
            End If
        Else
            '如没有病历中的项目就检查有没有那个项目
            strSQL = "select * from 诊疗项目目录  where id=" & mID诊疗项目
            Call zlDatabase.OpenRecordset(rs诊疗项目1, strSQL, "检验结果报告")
            ID诊疗项目 = mID诊疗项目
            If rs诊疗项目1.RecordCount < 1 Then
                '如果没有就退出
                Exit Sub
            End If
        End If
        '得到标本
        rs诊疗项目.Filter = "控件号=-1"
        If rs诊疗项目.RecordCount > 0 Then
            lblBB.Caption = zlCommFun.Nvl(rs诊疗项目!标题)
        Else
            lblBB.Caption = ""
        End If
        
        mblnCancel = True
        ReSetRowCode msfMain
        PicItem_Resize
        mblnCancel = False
        ReSetRowCode msfMain
    Else
        '否则没有数据就初始化检验结果项目的表格
        ID诊疗项目 = mID诊疗项目
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function SaveData(lng病人ID As Long, lng主页ID As Long, lng病历ID As Long, strReturnSQL As String, strError As String) As Boolean
'保存检验结果项目数据
Dim strTmp As String, strName As String
Dim lngRow As Long
Dim lngCol As Long

On Error GoTo ErrHandle

    If msfMain.Rows < 3 And Trim(msfMain.TextMatrix(1, EnmGridCol.ItemID)) = "" Then
        strReturnSQL = ""
        strError = "检验结果不能为空"
        Exit Function
    End If
    If mID诊疗项目 < 1 Then
        strReturnSQL = ""
        strError = "检验项目不确定，请重新选择检验项目"
        Exit Function
    End If
    '得到项目名称
    strName = IIf(txtItem.Tag = "", Trim(txtItem.Text), txtItem.Tag)
    '检验项目
    strTmp = mID诊疗项目 & "''"
    strTmp = strTmp & strName & "''"
    strTmp = strTmp & " ''"
    strTmp = strTmp & "-2''-2'' ''"
    '检验标本
    strName = Trim(lblBB.Caption)
    strTmp = strTmp & mID诊疗项目 & "''"
    strTmp = strTmp & strName & "''"
    strTmp = strTmp & " ''"
    strTmp = strTmp & "-1''-1'' ''"
    For lngRow = 1 To msfMain.Rows - 1
        '如果有英文名就取英文名,否则就取中文名，同时替换用来保存数据库的分隔符
        strName = IIf(Replace(Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item英文)), "'", "’") = "", Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item指标名)), Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item英文)))
        '存储过程参数格式:检验项目ID'标题文本'单位'所见项目ID'控件号'所见内容'检验项目ID1'标题文本1'单位1'所见项目ID1'控件号1'所见内容1'
        strTmp = strTmp & mID诊疗项目 & "''"
        strTmp = strTmp & strName & "''"
        strTmp = strTmp & Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item单位)) & "''"
        strTmp = strTmp & Trim(msfMain.TextMatrix(lngRow, EnmGridCol.ItemID)) & "''"
        
        For lngCol = 1 To Len("'`|,""")
            If InStr(Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item所见内容)), Mid("'`|,""", lngCol, 1)) > 0 Then
                msfMain.Row = lngRow: msfMain.Col = EnmGridCol.Item所见内容
                msfMain_EnterCell
                strError = "第" & lngRow & "行存在非法字符！"
                SetErr 0, "第" & lngRow & "行存在非法字符！"
                Exit Function
            End If
        Next
        strTmp = strTmp & CStr(lngRow - 1) & "''" & Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item所见内容)) & "''"
    Next
    strReturnSQL = "ZL_检验结果记录_INSERT(" & lng病历ID & ",'" & strTmp & "')"
    SaveData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitMe()
Dim rsNewTmp As New ADODB.Recordset
    '初始化控件
    msfMain.Clear
    msfMain.Rows = 2
    msfMain.FixedRows = 1
    msfMain.Cols = 13
    '第一行设置为空
    msfMain.RowHeight(0) = 0
    '表头
    msfMain.TextMatrix(0, EnmGridCol.ItemID) = "ID"
    msfMain.TextMatrix(0, EnmGridCol.Item类型) = "类型"
    msfMain.TextMatrix(0, EnmGridCol.Item表示法) = "表示法"
    msfMain.TextMatrix(0, EnmGridCol.Item数值域) = "数值域"
    msfMain.TextMatrix(0, EnmGridCol.Item初始值) = "初始值"
    msfMain.TextMatrix(0, EnmGridCol.Item长度) = "长度"
    msfMain.TextMatrix(0, EnmGridCol.Item小数长) = "小数长"
    msfMain.TextMatrix(0, EnmGridCol.Item行号) = "行号"
    msfMain.TextMatrix(0, EnmGridCol.Item指标名) = "指标名"
    msfMain.TextMatrix(0, EnmGridCol.Item英文) = "英文名"
    msfMain.TextMatrix(0, EnmGridCol.Item正常值) = "正常值"
    msfMain.TextMatrix(0, EnmGridCol.Item所见内容) = "所见内容"
    msfMain.TextMatrix(0, EnmGridCol.Item单位) = "单位"
    '设置各列的宽
    msfMain.ColWidth(EnmGridCol.ItemID) = 0
    msfMain.ColWidth(EnmGridCol.Item类型) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item表示法) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item数值域) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item初始值) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item长度) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item小数长) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item行号) = 300
    msfMain.ColWidth(EnmGridCol.Item指标名) = 2400
    msfMain.ColWidth(EnmGridCol.Item正常值) = 2400
    msfMain.ColWidth(EnmGridCol.Item英文) = 0
    
    '单位列的宽
    msfMain.ColWidth(EnmGridCol.Item单位) = 1000
    '自动计算出所见内容的宽
    i = msfMain.Width - (msfMain.ColWidth(EnmGridCol.Item行号) + msfMain.ColWidth(EnmGridCol.Item指标名) + msfMain.ColWidth(EnmGridCol.Item正常值) + msfMain.ColWidth(EnmGridCol.Item单位)) - Screen.TwipsPerPixelX * 6
    msfMain.ColWidth(EnmGridCol.Item所见内容) = IIf(i < 200, 200, i)
    '设置列对齐
    msfMain.ColAlignment(EnmGridCol.Item行号) = AlignmentSettings.flexAlignLeftCenter
    msfMain.ColAlignment(EnmGridCol.Item指标名) = AlignmentSettings.flexAlignLeftCenter
    msfMain.ColAlignment(EnmGridCol.Item正常值) = AlignmentSettings.flexAlignCenterCenter
    msfMain.ColAlignment(EnmGridCol.Item英文) = AlignmentSettings.flexAlignLeftCenter
    msfMain.ColAlignment(EnmGridCol.Item所见内容) = msfMain.ColAlignment(EnmGridCol.Item指标名)
    msfMain.ColAlignment(EnmGridCol.Item单位) = msfMain.ColAlignment(EnmGridCol.Item指标名)
    UserControl_Resize
    txtItem.Tag = ""
    txtItem.Text = ""
End Sub

Private Sub ReSetRowCode(objMSH As MSHFlexGrid)
'对行号进行重新设置
Dim lngWidth行号 As Long

    For i = 1 To objMSH.Rows - 1
'        objMSH.RowHeight(i) = CmbCell.Height
        objMSH.TextMatrix(i, EnmGridCol.Item行号) = CStr(i) & "、"
        If UserControl.TextWidth(objMSH.TextMatrix(i, EnmGridCol.Item行号)) > lngWidth行号 Then lngWidth行号 = UserControl.TextWidth(objMSH.TextMatrix(i, EnmGridCol.Item行号))
    Next
    objMSH.ColWidth(EnmGridCol.Item行号) = lngWidth行号
End Sub

Private Function InDesign() As Boolean
'功能：判断当前运行程序是否在VB的工程环境中
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Private Sub SetErr(lngErrNum As Long, strErr As String)
'设置错误描述及错误号
'如果lngErrNum=-1 表示 控件自己定义的错误
mReturnErrnumber = lngErrNum
mReturnErrDescription = strErr
End Sub

Public Property Get ID病人病历() As Long
'返回病人病历ID
    ID病人病历 = mlng病历id
End Property

Public Property Let ID病人病历(ByVal New_ID病人病历 As Long)
'设置病人病历ID,并检查该病历是不是存在
    mlng病历id = New_ID病人病历
    ReadData mlng病历id
End Property

Public Property Get ReturnErrNumber() As Long
'返回最后一次的错误号
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
'返回最后一次错误描述字符串
    ReturnErrDescription = mReturnErrDescription
End Property

Public Property Get DispMode() As Boolean
'是否为显示模式
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    msfMain_EnterCell
    PropertyChanged "DispMode"
End Property

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub CmbCell_Click()
If mblnCancel = False Then
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) = CmbCell.Text
Else
    mblnCancel = False
End If
End Sub

Private Sub CmbCell_DblClick()
If mblnCancel = False Then
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) = CmbCell.Text
End If
End Sub

Private Sub CmbCell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown) And Shift = 2) Then
        KeyCode = 0
        If msfMain.Row < msfMain.Rows - 1 Then
            mblnCancel = True
            msfMain.Row = msfMain.Row + 1
            msfMain_EnterCell
            mblnCancel = True
            Exit Sub
        End If
    ElseIf (KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp) And Shift = 2 Then
        KeyCode = 0
        If msfMain.Row > 1 Then
            mblnCancel = True
            msfMain.Row = msfMain.Row - 1
            msfMain_EnterCell
            mblnCancel = True
            Exit Sub
        End If
    End If
    '
    If mblnCancel = False Then
        msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) = CmbCell.Text
    End If
End Sub

Private Sub cmdP_Click()
On Error GoTo ErrHandle
Dim strSQL As String
Dim strReturn As String
Dim CurPoint As POINTAPI
Dim rsNewTmp As New ADODB.Recordset
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Sub
    strSQL = "SELECT DISTINCT 检验标本 标本名称 FROM 检验报告项目 where 诊疗项目ID=" & mID诊疗项目
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "检验结果报告")
    If rsTmp.RecordCount = 1 Then
        lblBB.Caption = zlCommFun.Nvl(rsTmp!标本名称)
        PicItem_Resize
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        '定位选择器
        CurPoint.X = (cmdP.Left) / Screen.TwipsPerPixelX
        CurPoint.Y = (cmdP.Top + cmdP.Height + Screen.TwipsPerPixelY * 2) / Screen.TwipsPerPixelY
        ClientToScreen PicItem.hwnd, CurPoint
        CurPoint.X = CurPoint.X * Screen.TwipsPerPixelX
        CurPoint.Y = CurPoint.Y * Screen.TwipsPerPixelY
        If CurPoint.X + 3000 > Screen.Width Then CurPoint.X = Screen.Width - 3200
        If CurPoint.X < 0 Then CurPoint.X = 0
        If CurPoint.Y + 2400 > Screen.Height Then CurPoint.Y = Screen.Height - 2800
        If CurPoint.Y < 0 Then CurPoint.Y = 0
        strReturn = frmSelectChild.ShowSelectChild(Me, CurPoint.X, CurPoint.Y, 3200, 2400, rsTmp, "2800")
        If Trim(strReturn) = "" Or Trim(strReturn) = ";" Then Exit Sub
        lblBB.Caption = Split(strReturn, ";")(0)
        PicItem_Resize
    End If
    '设置指定标本的指标项目
    strSQL = _
        "SELECT C.ID," & vbCrLf & _
        "       A.排列序号 序号," & vbCrLf & _
        "       A.检验标本," & vbCrLf & _
        "       C.类型," & vbCrLf & _
        "       C.表示法," & vbCrLf & _
        "       C.数值域," & vbCrLf & _
        "       C.初始值," & vbCrLf & _
        "       C.长度," & vbCrLf & _
        "       C.小数," & vbCrLf & _
        "       C.中文名 指标名," & vbCrLf & _
        "       C.英文名 英文名," & vbCrLf & _
        "       C.单位" & vbCrLf & _
        "  FROM 诊疗项目目录 B, 检验报告项目 A,诊治所见项目 C" & vbCrLf & _
        " WHERE B.ID IN (SELECT DISTINCT 诊疗项目ID FROM 检验报告项目) AND  A.报告项目id=C.Id AND " & vbCrLf & _
        "      A.检验标本='" & lblBB.Caption & "' AND B.ID = A.诊疗项目ID  AND A.诊疗项目ID =" & ID诊疗项目 & vbCrLf & _
        " ORDER BY A.排列序号"
    Call zlDatabase.OpenRecordset(rsNewTmp, strSQL, "检验结果报告")
    '读出检验项目的检验指标
    If rsNewTmp.RecordCount > 0 Then
        mblnCancel = True
        rsNewTmp.MoveFirst
        '读出
        msfMain.Rows = rsNewTmp.RecordCount + 1
        For i = 1 To rsNewTmp.RecordCount
            msfMain.TextMatrix(i, EnmGridCol.ItemID) = zlCommFun.Nvl(rsNewTmp!ID, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item类型) = zlCommFun.Nvl(rsNewTmp!类型, 1)
            msfMain.TextMatrix(i, EnmGridCol.Item表示法) = zlCommFun.Nvl(rsNewTmp!表示法, 0)
            If zlCommFun.Nvl(rsNewTmp!类型, 1) = 0 Then
                If Trim(zlCommFun.Nvl(rsNewTmp!数值域)) <> ";" Then
                    msfMain.TextMatrix(i, EnmGridCol.Item正常值) = Replace(Trim(zlCommFun.Nvl(rsNewTmp!数值域)), ";", " ～ ")
                End If
            End If
            msfMain.TextMatrix(i, EnmGridCol.Item数值域) = Trim(zlCommFun.Nvl(rsNewTmp!数值域))
            msfMain.TextMatrix(i, EnmGridCol.Item初始值) = Trim(zlCommFun.Nvl(rsNewTmp!初始值))
            msfMain.TextMatrix(i, EnmGridCol.Item长度) = zlCommFun.Nvl(rsNewTmp!长度, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item小数长) = zlCommFun.Nvl(rsNewTmp!小数, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item指标名) = Trim(zlCommFun.Nvl(rsNewTmp!指标名)) & IIf(Trim(zlCommFun.Nvl(rsNewTmp!英文名)) = "", "", "［" & Trim(zlCommFun.Nvl(rsNewTmp!英文名)) & "］")
            msfMain.TextMatrix(i, EnmGridCol.Item英文) = Trim(zlCommFun.Nvl(rsNewTmp!英文名))
            msfMain.TextMatrix(i, EnmGridCol.Item所见内容) = Trim(zlCommFun.Nvl(rsNewTmp!初始值))
            msfMain.TextMatrix(i, EnmGridCol.Item单位) = Trim(zlCommFun.Nvl(rsNewTmp!单位))
            rsNewTmp.MoveNext
        Next
        ReSetRowCode msfMain
        PicItem_Resize
    End If
    UserControl_Resize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub listCell_ItemCheck(Item As Integer)
Dim strTmp As String
    If mblnCancel = True Then Exit Sub
    For i = 0 To listCell.ListCount - 1
        If listCell.Selected(i) = True Then
            strTmp = strTmp & listCell.List(i) & ";"
        End If
    Next
    If Right(strTmp, 1) = ";" Then strTmp = Left(strTmp, Len(strTmp) - 1)
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) = strTmp
End Sub

Private Sub listCell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If msfMain.Row < msfMain.Rows - 1 Then
            listCell.Visible = False
            msfMain.Row = msfMain.Row + 1
            msfMain_EnterCell
        End If
    End If
End Sub

Private Sub SetSelColor(objMsf As MSHFlexGrid, ByVal lngRow As Long, Optional ByVal oleForeColor As OLE_COLOR = 0, Optional ByVal oleBackColor As OLE_COLOR = &HFFFFFF)
'设置选择行的颜色
Dim lngSelCol As Long, lngSelRow As Long

    objMsf.Redraw = False
    lngSelCol = objMsf.Col
    lngSelRow = objMsf.Row
    
    For i = 1 To objMsf.Rows - 1
        objMsf.Row = i
        If i = lngRow Then
            For j = 0 To objMsf.Cols - 1
                objMsf.Col = j
                objMsf.CellFontBold = True
                objMsf.CellForeColor = oleForeColor
                objMsf.CellBackColor = oleBackColor
            Next
        Else
            For j = 0 To objMsf.Cols - 1
                objMsf.Col = j
                objMsf.CellFontBold = False
                objMsf.CellForeColor = 0
                objMsf.CellBackColor = RGB(255, 255, 255)
            Next
        End If
    Next
    objMsf.Col = lngSelCol
    objMsf.Row = lngSelRow
    objMsf.Refresh
    objMsf.Redraw = True
End Sub

Private Sub msfMain_EnterCell()
On Error Resume Next
mblnCancel = True
txtCell.Visible = False
listCell.Visible = False
CmbCell.Visible = False

SetSelColor msfMain, msfMain.Row
If msfMain.Row > 0 Then
    '给那行的高重新赋值
    msfMain.RowHeight(msfMain.Row) = 255
    txtCell.Height = 255
    If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item表示法)) = CStr(EnmCTLType.CTLTxt) Or Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item表示法)) = CStr(EnmCTLType.CTLUpDown) Then
    '如果为文本和上下时
        '求左边
        txtCell.Left = msfMain.ColPos(EnmGridCol.Item所见内容) + Screen.TwipsPerPixelX * 2
        '求宽
        i = msfMain.ColWidth(EnmGridCol.Item所见内容) - Screen.TwipsPerPixelX * 4
        txtCell.Width = IIf(i < Screen.TwipsPerPixelX * 4, Screen.TwipsPerPixelX * 4, i)
        '求顶
        txtCell.Top = msfMain.Top + msfMain.RowPos(msfMain.Row) + Screen.TwipsPerPixelY * 2
        '求高
        i = msfMain.CellHeight - Screen.TwipsPerPixelY * 4
        txtCell.Height = IIf(i < Screen.TwipsPerPixelY * 4, Screen.TwipsPerPixelY * 4, i)
        '得到当前内容
        mblnCancel = True
        txtCell.Text = msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)
        mblnCancel = False
        If txtCell.Enabled And UserControl.Enabled And mDispMode = False Then
            txtCell.Visible = True
            txtCell.ZOrder
            txtCell.SetFocus
        End If
    ElseIf Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item表示法)) = CStr(EnmCTLType.CTLDownList) Or Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item表示法)) = CStr(EnmCTLType.CTLOption) Then
        '如果为下拉和单选时
        '求左边
        CmbCell.Left = msfMain.ColPos(EnmGridCol.Item所见内容) + Screen.TwipsPerPixelX * 2
        '求宽
        i = msfMain.ColWidth(EnmGridCol.Item所见内容) - Screen.TwipsPerPixelX * 4
        CmbCell.Width = IIf(i < Screen.TwipsPerPixelX * 4, Screen.TwipsPerPixelX * 4, i)
        '求顶
        CmbCell.Top = msfMain.Top + msfMain.RowPos(msfMain.Row) + Screen.TwipsPerPixelY * 2
        msfMain.RowHeight(msfMain.Row) = CmbCell.Height
        '得到当前内容
        mblnCancel = True
        CmbCell.Clear
        If InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item数值域)), ";") > 0 Then
            '选初始化
            For i = 0 To UBound(Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item数值域)), ";"))
                CmbCell.AddItem Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item数值域)), ";")(i)
            Next
            CmbCell.ListIndex = 0
            '设置值
            If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)) = "" Then
                msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) = CmbCell.Text
            Else
                For i = 0 To CmbCell.ListCount - 1
                    If CmbCell.List(i) = msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) Then
                        CmbCell.ListIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
        If CmbCell.Enabled And UserControl.Enabled And mDispMode = False Then
            CmbCell.Visible = True
            CmbCell.ZOrder
            CmbCell.SetFocus
        End If
        mblnCancel = False
    ElseIf Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item表示法)) = CStr(EnmCTLType.CTLCheck) Then
        '如果为复选项时
        '求左边
        listCell.Left = msfMain.ColPos(EnmGridCol.Item所见内容) + Screen.TwipsPerPixelX * 2
        '求宽
        i = msfMain.ColWidth(EnmGridCol.Item所见内容) - Screen.TwipsPerPixelX * 4
        listCell.Width = IIf(i < Screen.TwipsPerPixelX * 4, Screen.TwipsPerPixelX * 4, i)
        '求顶
        listCell.Top = msfMain.Top + msfMain.RowPos(msfMain.Row) + msfMain.CellHeight + Screen.TwipsPerPixelY * 2
        '重新设置高
        listCell.Height = 1200
        If listCell.Top + listCell.Height > UserControl.Height Then
            listCell.Top = listCell.Top - msfMain.CellHeight - Screen.TwipsPerPixelY * 2 - listCell.Height
        End If
        listCell.Clear
        mblnCancel = True
        If InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item数值域)), ";") > 0 Then
            '选初始化
            For i = 0 To UBound(Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item数值域)), ";"))
                listCell.AddItem Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item数值域)), ";")(i)
            Next
            '再设置值
            For i = 0 To listCell.ListCount - 1
                For j = 0 To UBound(Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)), ";"))
                    If listCell.List(i) = Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)), ";")(j) Then
                        listCell.Selected(i) = True
                    End If
                Next
            Next
        End If
        If listCell.Enabled And UserControl.Enabled And mDispMode = False Then
            listCell.Visible = True
            listCell.ZOrder
            listCell.SetFocus
        End If
        mblnCancel = False
    End If
End If
End Sub

Private Sub msfMain_Scroll()
    txtCell.Visible = False
    listCell.Visible = False
    CmbCell.Visible = False
End Sub

Private Sub PicItem_Resize()
    Line1.X2 = PicItem.ScaleWidth - Line1.X1
    cmdP.Left = Line1.X2 - cmdP.Width
    lblBB.Left = cmdP.Left - lblBB.Width - Screen.TwipsPerPixelX * 10
    lblBBCaption.Left = lblBB.Left - lblBBCaption.Width - Screen.TwipsPerPixelX * 4
End Sub

Private Sub txtCell_Change()
If mblnCancel = False Then
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) = txtCell.Text
End If
End Sub

Private Sub txtCell_GotFocus()
    If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item类型)) Then
        If Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item类型)) = 0 And IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item长度)) Then
            If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item长度)) Then
                If Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item长度)) > 0 Then
                    txtCell.MaxLength = Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item长度))
                End If
            End If
        End If
    End If
    zlControl.TxtSelAll txtCell
    zlCommFun.OpenIme True
End Sub

Private Sub txtCell_KeyDown(KeyCode As Integer, Shift As Integer)
Dim blnCancel As Boolean
'先检查是不是输入了非法字符
    If InStr(LAWLChar, Chr(KeyCode)) > 0 Then
        KeyCode = 0
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown, vbKeyPageDown
            KeyCode = 0
            If msfMain.Row < msfMain.Rows - 1 Then
                txtCell_Validate blnCancel
                If mblnLawless = True Then mblnLawless = False:   Exit Sub
                msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容) = txtCell.Text
                txtCell.Visible = False
                msfMain.Row = msfMain.Row + 1
                msfMain_EnterCell
                Exit Sub
            End If
        Case vbKeyUp, vbKeyPageUp
            KeyCode = 0
            If msfMain.Row > 1 Then
                txtCell_Validate blnCancel
                If mblnLawless = True Then mblnLawless = False:   Exit Sub
                txtCell.Visible = False
                msfMain.Row = msfMain.Row - 1
                msfMain_EnterCell
                Exit Sub
            End If
    End Select
End Sub

Private Sub txtCell_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item类型)) = "0" Then
        If IsNumeric(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item小数长))) Then
            If Format(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item小数长))) > 0 Then
                '为小数时
                If InStr("0123456789.", Chr(KeyAscii)) < 1 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            Else
                '为整数
                If InStr("0123456789", Chr(KeyAscii)) < 1 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Else
            '为整数
            If InStr("0123456789", Chr(KeyAscii)) < 1 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtCell_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txtCell_Validate(Cancel As Boolean)
'长度检查
If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item长度)) Then
    If zlCommFun.ActualLen(txtCell.Text) > 1000 And Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item长度)) > 0 And _
        zlCommFun.ActualLen(txtCell.Text) > Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item长度)) Then
        MsgBox "输入超长,请重新输入！", vbInformation, gstrSysName
        mblnLawless = True
        Cancel = True
        Exit Sub
    End If
    If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item类型)) = "0" Then
        If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)) = False And Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)) <> "" Then
            MsgBox "只能输入数值,请重新输入！", vbInformation, gstrSysName
            mblnLawless = True
            Cancel = True
            Exit Sub
        End If
        If Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item小数长)) > 0 And InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)), ".") > 0 Then
            i = InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容)), ".")
            i = Len(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item所见内容))) - i
            If i > Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item小数长)) Then
               MsgBox "输入小数部分超长,请重新输入！", vbInformation, gstrSysName
               mblnLawless = True
               Cancel = True
               Exit Sub
            End If
        End If
    Else
        If zlCommFun.ActualLen(txtCell.Text) > 1000 Then
            MsgBox "输入超长,请重新输入！", vbInformation, gstrSysName
            mblnLawless = True
            Cancel = True
            Exit Sub
        End If
    End If
Else
    If zlCommFun.ActualLen(txtCell.Text) > 1000 Then
        MsgBox "输入超长,请重新输入！", vbInformation, gstrSysName
        mblnLawless = True
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub txtItem_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtItem
End Sub

Public Property Get Cur当前标本() As String
    '
    Cur当前标本 = lblBB.Caption
End Property

Public Property Let Cur当前标本(ByVal New_Cur As String)
    lblBB.Caption = New_Cur
End Property

Private Sub cmdP1_Click()
On Error GoTo ErrHandle
Dim strWidth As String
Dim CurPoint As POINTAPI

    strSQL = _
        "SELECT a.* FROM (SELECT DISTINCT A.ID, A.编码, A.名称, B.名称 别名, A.计算单位" & vbCrLf & _
        "  FROM 诊疗项目目录 A, 诊疗项目别名 B" & vbCrLf & _
        " WHERE B.诊疗项目ID = A.ID(+) AND" & vbCrLf & _
        "      A.ID IN (SELECT DISTINCT 诊疗项目ID FROM 检验报告项目)) A" & vbCrLf & _
        " order by a.编码 "
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "检验结果报告")
    If rsTmp.RecordCount = 1 Then
        ID诊疗项目 = rsTmp!ID
    ElseIf rsTmp.RecordCount > 0 Then
        '定位选择器
        CurPoint.X = (txtItem.Left) / Screen.TwipsPerPixelX
        CurPoint.Y = (txtItem.Top + txtItem.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
        ClientToScreen UserControl.hwnd, CurPoint
        CurPoint.X = CurPoint.X * Screen.TwipsPerPixelX
        CurPoint.Y = CurPoint.Y * Screen.TwipsPerPixelY
        If CurPoint.X + 4800 > Screen.Width Then CurPoint.X = Screen.Width - 5180
        If CurPoint.X < 0 Then CurPoint.X = 0
        If CurPoint.Y + Screen.TwipsPerPixelY * 200 > Screen.Height Then CurPoint.Y = CurPoint.Y - txtItem.Height - Screen.TwipsPerPixelY * 200 - Screen.TwipsPerPixelY * 2
        If CurPoint.Y < 0 Then CurPoint.Y = 0
        
        '初始选择器
        strWidth = "0;800;1500;1500;1000"
        strWidth = frmSelectChild.ShowSelectChild(Nothing, CurPoint.X, CurPoint.Y, 4800 + 380, Screen.TwipsPerPixelY * 200, rsTmp, strWidth)
        If Trim(strWidth) = "" Or Trim(strWidth) = ";;" Then
            Exit Sub
        End If
        ID诊疗项目 = CLng(Split(strWidth, ";")(0))
    Else
        txtItem.Text = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtItem_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
Dim blnMatching As Boolean
Dim strWidth As String
Dim CurPoint As POINTAPI

    If KeyCode = vbKeyReturn Then
        blnMatching = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", "0") = "0", True, False)
        KeyCode = 0
        strSQL = _
            "SELECT a.* FROM (SELECT DISTINCT A.ID, A.编码,A.名称, B.名称 别名, A.计算单位" & vbCrLf & _
            "  FROM 诊疗项目目录 A, 诊疗项目别名 B" & vbCrLf & _
            " WHERE B.诊疗项目ID = A.ID(+) AND " & vbCrLf & _
            "      (Upper(Nvl(a.编码,'')) like '" & UCase(txtItem.Text) & "%' Or  " & vbCrLf & _
            "       Upper(Nvl(a.名称,'')) like '" & IIf(blnMatching = True, "%", "") & UCase(txtItem.Text) & "%' Or  " & vbCrLf & _
            "       Upper(Nvl(b.名称,'')) like '" & IIf(blnMatching = True, "%", "") & UCase(txtItem.Text) & "%' Or  " & vbCrLf & _
            "       Upper(Nvl(b.简码,'')) like '" & IIf(blnMatching = True, "%", "") & UCase(txtItem.Text) & "%' )  and  " & vbCrLf & _
            "      A.ID IN (SELECT DISTINCT 诊疗项目ID FROM 检验报告项目)) A " & vbCrLf & _
            " order by a.编码 "

        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "检验结果报告")
        If rsTmp.RecordCount = 1 Then
            ID诊疗项目 = rsTmp!ID
        ElseIf rsTmp.RecordCount > 0 Then
            '定位选择器
            CurPoint.X = (txtItem.Left) / Screen.TwipsPerPixelX
            CurPoint.Y = (txtItem.Top + txtItem.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen UserControl.hwnd, CurPoint
            CurPoint.X = CurPoint.X * Screen.TwipsPerPixelX
            CurPoint.Y = CurPoint.Y * Screen.TwipsPerPixelY
            If CurPoint.X + 4800 > Screen.Width Then CurPoint.X = Screen.Width - 5180
            If CurPoint.X < 0 Then CurPoint.X = 0
            If CurPoint.Y + Screen.TwipsPerPixelY * 200 > Screen.Height Then CurPoint.Y = CurPoint.Y - txtItem.Height - Screen.TwipsPerPixelY * 200 - Screen.TwipsPerPixelY * 2
            If CurPoint.Y < 0 Then CurPoint.Y = 0
            
            '初始选择器
            strWidth = "0;800;1500;1500;1000"
            strWidth = frmSelectChild.ShowSelectChild(Nothing, CurPoint.X, CurPoint.Y, 4800 + 380, Screen.TwipsPerPixelY * 200, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;" Then
                Exit Sub
            End If
            ID诊疗项目 = CLng(Split(strWidth, ";")(0))
        Else
            ID诊疗项目 = 0
            txtItem.Text = ""
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtItem_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub UserControl_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub UserControl_Initialize()
    UserControl.Font.Name = "宋体"
    UserControl.Font.Size = 9
    UserControl.Font.Bold = True
    ID诊疗项目 = mID诊疗项目
    mItemIndex = -1
    mblnCancel = False
End Sub

Private Sub UserControl_InitProperties()
    mDispMode = False
    mShowItem = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", False)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    ShowItem = PropBag.ReadProperty("ShowItem", True)
    mblnCancel = True
    ID诊疗项目 = PropBag.ReadProperty("ID诊疗项目", 0)
    mblnCancel = False
    txtCell.Visible = False
    CmbCell.Visible = False
    listCell.Visible = False
End Sub

Public Property Get Text() As String
'为每一个控件加上文本转储属性
Dim i As Long
Dim strTmp As String, strName As String

'通过用户输入的内容得到转储文本
    If msfMain.Rows < 2 Then Exit Property
    If msfMain.Rows = 2 And msfMain.TextMatrix(1, EnmGridCol.Item指标名) = "" Then Exit Property
    '得到项目名称
    strName = IIf(txtItem.Tag = "", Trim(txtItem.Text), txtItem.Tag)
    '检验项目
    strTmp = strName & "（" & lblBB.Caption & "）："
    For i = 1 To msfMain.Rows - 1
        strTmp = strTmp & IIf(Trim(msfMain.TextMatrix(i, EnmGridCol.Item英文)) <> "", Trim(msfMain.TextMatrix(i, EnmGridCol.Item英文)), msfMain.TextMatrix(i, EnmGridCol.Item指标名)) & " " & msfMain.TextMatrix(i, EnmGridCol.Item所见内容) & msfMain.TextMatrix(i, EnmGridCol.Item单位) & IIf(i = msfMain.Rows - 1, "", "，")
    Next
    Text = strTmp
End Property

Private Sub UserControl_Resize()
    Dim lngWidth As Long
    Dim lngWidth单位 As Long
    
    msfMain.Left = 0
    msfMain.Top = PicItem.Top + PicItem.Height
    msfMain.Width = ScaleWidth
    i = ScaleHeight - (PicItem.Top + PicItem.Height)
    msfMain.Height = IIf(i > Screen.TwipsPerPixelY, i, Screen.TwipsPerPixelY)
    msfMain_EnterCell
    
    msfMain.ColWidth(EnmGridCol.Item指标名) = 2400
    msfMain.ColWidth(EnmGridCol.Item正常值) = 1200
    For i = 1 To msfMain.Rows - 1
        If UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item指标名)) > lngWidth Then
            lngWidth = UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item指标名)) / 2
        End If
        If UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item单位)) > lngWidth单位 Then
            lngWidth单位 = UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item单位)) / 2
        End If
    Next
    If msfMain.ColWidth(EnmGridCol.Item指标名) < lngWidth Then
        msfMain.ColWidth(EnmGridCol.Item指标名) = lngWidth
    End If
    If msfMain.ColWidth(EnmGridCol.Item单位) < lngWidth单位 Then
        msfMain.ColWidth(EnmGridCol.Item单位) = lngWidth单位
    End If
    '自动计算出所见内容的宽
    i = msfMain.Width - (msfMain.ColWidth(EnmGridCol.Item行号) + msfMain.ColWidth(EnmGridCol.Item指标名) + msfMain.ColWidth(EnmGridCol.Item正常值) + msfMain.ColWidth(EnmGridCol.Item英文) + msfMain.ColWidth(EnmGridCol.Item单位)) - Screen.TwipsPerPixelX * 20
    msfMain.ColWidth(EnmGridCol.Item所见内容) = IIf(i < 200, 200, i)
    msfMain_EnterCell
    mblnCancel = False
End Sub

Private Sub UserControl_Show()
    mblnCancel = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ShowItem", mShowItem, True)
    Call PropBag.WriteProperty("DispMode", mDispMode, False)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ID诊疗项目", mID诊疗项目, 0)
End Sub
 
Private Sub UserControl_EnterFocus()
    On Error Resume Next
    UserControl.Parent.CallBack_GotFocus
End Sub

