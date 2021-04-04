VERSION 5.00
Begin VB.Form frmVerifyEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "组合项目编辑"
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8505
   Icon            =   "frmVerifyEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Height          =   3105
      Left            =   45
      TabIndex        =   33
      Top             =   -45
      Width           =   7110
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   5655
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   165
         Width           =   1395
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   270
         Index           =   2
         Left            =   4545
         TabIndex        =   2
         Top             =   180
         Width           =   285
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   8
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   165
         Width           =   3390
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1620
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1125
         TabIndex        =   15
         Top             =   1635
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   3360
         TabIndex        =   23
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   1125
         TabIndex        =   21
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1125
         TabIndex        =   19
         Top             =   2010
         Width           =   3390
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   3360
         TabIndex        =   12
         Top             =   1275
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1125
         TabIndex        =   10
         Top             =   1275
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1125
         TabIndex        =   8
         Top             =   915
         Width           =   3390
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1125
         TabIndex        =   6
         Top             =   570
         Width           =   1155
      End
      Begin VB.CheckBox chk 
         Caption         =   "需要执行安排(&E)"
         Height          =   210
         Index           =   1
         Left            =   1140
         TabIndex        =   25
         Top             =   2805
         Width           =   1680
      End
      Begin VB.CheckBox chk 
         Caption         =   "住院(&2)"
         Height          =   210
         Index           =   3
         Left            =   5760
         TabIndex        =   27
         Top             =   1320
         Width           =   1005
      End
      Begin VB.CheckBox chk 
         Caption         =   "门诊(&1)"
         Height          =   210
         Index           =   2
         Left            =   5760
         TabIndex        =   26
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   0
         TabIndex        =   38
         Top             =   480
         Width           =   7110
      End
      Begin VB.Frame Frame2 
         Height          =   2685
         Left            =   5310
         TabIndex        =   34
         Top             =   420
         Width           =   30
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "服务范围"
         Height          =   180
         Index           =   15
         Left            =   5775
         TabIndex        =   39
         Top             =   660
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5445
         Picture         =   "frmVerifyEdit.frx":000C
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "类型(&T)"
         Height          =   180
         Index           =   1
         Left            =   4995
         TabIndex        =   3
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "上级(&D)"
         Height          =   180
         Index           =   12
         Left            =   465
         TabIndex        =   0
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lbl 
         Caption         =   "说明是否允许该项目作为门诊病人或住院病人的诊疗措施应用，或者不能直接应用于病人。"
         Height          =   1155
         Index           =   26
         Left            =   5745
         TabIndex        =   36
         Top             =   1635
         Width           =   1275
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(拼音)"
         Height          =   180
         Index           =   10
         Left            =   2280
         TabIndex        =   22
         Top             =   2445
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(拼音)"
         Height          =   180
         Index           =   8
         Left            =   2280
         TabIndex        =   11
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(五笔)"
         Height          =   180
         Index           =   11
         Left            =   4545
         TabIndex        =   24
         Top             =   2445
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(五笔)"
         Height          =   180
         Index           =   9
         Left            =   4545
         TabIndex        =   13
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "计算单位(&U)"
         Height          =   180
         Index           =   7
         Left            =   105
         TabIndex        =   14
         Top             =   1710
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "适用性别(&X)"
         Height          =   180
         Index           =   6
         Left            =   2340
         TabIndex        =   16
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "别名简码(&F)"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   20
         Top             =   2430
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "别名(&A)"
         Height          =   180
         Index           =   4
         Left            =   465
         TabIndex        =   18
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   3
         Left            =   465
         TabIndex        =   9
         Top             =   1365
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   465
         TabIndex        =   7
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   5
         Top             =   645
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7275
      TabIndex        =   30
      Top             =   75
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7275
      TabIndex        =   31
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7275
      TabIndex        =   32
      Top             =   1410
      Width           =   1100
   End
   Begin VB.Frame Frame4 
      Height          =   2715
      Left            =   45
      TabIndex        =   35
      Top             =   2985
      Width           =   7110
      Begin VB.PictureBox vsf 
         Height          =   2505
         Left            =   1440
         ScaleHeight     =   2445
         ScaleWidth      =   5550
         TabIndex        =   29
         Top             =   150
         Width           =   5610
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   105
         Picture         =   "frmVerifyEdit.frx":1D06
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "组成项目(&M)"
         Height          =   180
         Index           =   14
         Left            =   435
         TabIndex        =   28
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         Caption         =   "在开本项目的检验申请单时，将同时检验设置的所有组成项目。"
         Height          =   1815
         Index           =   13
         Left            =   435
         TabIndex        =   37
         Top             =   600
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmVerifyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mlngUpKey As Long
Private mlngKey As Long
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSQL As String
            
Private Function ShowOpenList(Optional strText As String) As Byte
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    
    On Error GoTo errHand
    
    strLvw = "编码,1200,0,1;检验项目,2700,0,0;英文缩写,900,0,0;标本类型,900,0,0"

    ShowOpenList = 2
    
    strTmp = Trim(vsf.TextMatrix(1, 4))
    For mlngLoop = 2 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) > 0 And vsf.TextMatrix(mlngLoop, 4) <> "" And mlngLoop <> vsf.Row Then
            strTmp = GetCommon(strTmp, Split(vsf.TextMatrix(mlngLoop, 4), ","))
            If strTmp = "" Then
                ShowSimpleMsg "设置的检验项目没有共同的标本类型！"
                Exit Function
            End If
        End If
    Next
    
    strText = UCase(strText)
    
    strSQL = "SELECT C.ID,C.编码,C.中文名 AS 检验项目,C.英文名 AS 英文名称,D.缩写 AS 英文缩写,zlGetSample(C.ID) AS 标本类型,D.计算公式 " & _
                "FROM 诊疗项目目录 A,检验报告项目 B,诊治所见项目 C,检验项目 D " & _
                "WHERE A.ID=B.诊疗项目ID AND  D.项目类别 IN (1,3) AND NVL(A.组合项目,0)=0 " & _
                    IIf(strTmp = "", "", "AND C.ID IN (SELECT 项目id FROM 检验项目参考 WHERE INSTR('," & strTmp & ",',','||标本类型||',')>0)") & _
                    "AND B.报告项目ID=C.ID AND C.ID=D.诊治项目id AND A.类别='C'"
                    
    strSQL = strSQL & " AND (UPPER(A.编码) Like '%" & strText & "%' OR UPPER(D.缩写) LIKE '%" & strText & "%' OR A.名称 Like '%" & strText & "%' OR A.ID IN (SELECT 诊疗项目id FROM 诊疗项目别名 WHERE (名称 Like '%" & strText & "%' OR UPPER(简码) Like '%" & strText & "%')))"
        
    Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
    
    If rs.BOF Then
        
        ShowOpenList = 0
        
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then
    
        If CheckHave(zlCommFun.Nvl(rs("ID").value)) Then
            MsgBox "选择的项目“" & zlCommFun.Nvl(rs("检验项目").value) & "”以前已经选择！", vbInformation, gstrSysName
            Exit Function
        End If
        
        vsf.EditText = zlCommFun.Nvl(rs("检验项目").value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("检验项目").value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("检验项目").value)
        vsf.TextMatrix(vsf.Row, 2) = zlCommFun.Nvl(rs("英文名称").value)
        vsf.TextMatrix(vsf.Row, 3) = zlCommFun.Nvl(rs("英文缩写").value)
        vsf.TextMatrix(vsf.Row, 4) = zlCommFun.Nvl(rs("标本类型").value)
        vsf.TextMatrix(vsf.Row, 5) = zlCommFun.Nvl(rs("计算公式").value)
        vsf.RowData(vsf.Row) = zlCommFun.Nvl(rs("ID").value)
        
        ShowOpenList = 1
        Exit Function
    End If
    
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY + 30, 4200, 2100, Me.Name & "\检验项目选择", "请从下表中选择一个项目") Then
        
        If CheckHave(zlCommFun.Nvl(rs("ID").value)) Then
            MsgBox "选择的项目“" & zlCommFun.Nvl(rs("检验项目").value) & "”以前已经选择！", vbInformation, gstrSysName
            Exit Function
        End If
        
        vsf.EditText = zlCommFun.Nvl(rs("检验项目").value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("检验项目").value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("检验项目").value)
        vsf.TextMatrix(vsf.Row, 2) = zlCommFun.Nvl(rs("英文名称").value)
        vsf.TextMatrix(vsf.Row, 3) = zlCommFun.Nvl(rs("英文缩写").value)
        vsf.TextMatrix(vsf.Row, 4) = zlCommFun.Nvl(rs("标本类型").value)
        vsf.TextMatrix(vsf.Row, 5) = zlCommFun.Nvl(rs("计算公式").value)
        vsf.RowData(vsf.Row) = zlCommFun.Nvl(rs("ID").value)
        
        ShowOpenList = 1
        
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
            
            
Private Function ShowOpenTree() As Byte
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    
    On Error GoTo errHand
    
    strLvw = "编码,1200,0,1;名称,2700,0,0;英文缩写,900,0,0;标本类型,900,0,0"

    ShowOpenTree = 2
    
    strTmp = Trim(vsf.TextMatrix(1, 4))
    For mlngLoop = 2 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) > 0 And vsf.TextMatrix(mlngLoop, 4) <> "" And mlngLoop <> vsf.Row Then
            strTmp = GetCommon(strTmp, Split(vsf.TextMatrix(mlngLoop, 4), ","))
            If strTmp = "" Then
                ShowSimpleMsg "设置的检验项目没有共同的标本类型！"
                Exit Function
            End If
        End If
    Next
    
    
    strSQL = "select * " & _
             "from (Select DISTINCT ID,上级ID,0 as 末级,编码,名称 ,'' as 英文名称,'' as 英文缩写,'' AS 标本类型,'' AS 计算公式, " & _
                                   "DECODE(上级ID, Null, ID * POWER(10, 20), 上级ID * POWER(10, 20) + ID) As 排序 " & _
                     "From 诊疗分类目录 " & _
                    "Where 类型 = 5 " & _
                    "Start With ID IN (SELECT DISTINCT 分类id FROM 诊疗项目目录 WHERE 类别 = 'C') " & _
                   "Connect by Prior 上级ID = ID " & _
                   "Union All " & _
                     "Select C.ID,A.分类id AS 上级ID,1 as 末级, A.编码,A.名称,C.英文名 AS 英文名称,D.缩写 AS 英文缩写,zlGetSample(C.ID) AS 标本类型,D.计算公式, " & _
                            "1 AS 排序 " & _
                       "FROM 诊疗项目目录 A,检验报告项目 B,诊治所见项目 C,检验项目 D " & _
                      "Where A.ID=B.诊疗项目id AND B.报告项目id=C.ID AND C.ID=D.诊治项目id AND D.项目类别 IN (1,3) AND NVL(A.组合项目,0)=0 AND A.类别 = 'C' AND (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) " & _
                            IIf(strTmp = "", "", "AND C.ID IN (SELECT 项目id FROM 检验项目参考 WHERE INSTR('," & strTmp & ",',','||标本类型||',')>0)") & _
                   ") A " & _
            "ORDER BY A.末级, A.排序, A.编码"
                        
    Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
    
    If rs.BOF Then Exit Function
    
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectExplorer.ShowSelect(Me, _
                            rs, _
                            sglX, _
                            sglY, _
                            5400, _
                            2400, _
                            vsf.CellHeight, _
                            "检验项目树型选择", _
                            strLvw, _
                            "请选择一个检验项目") Then
                            
        If CheckHave(zlCommFun.Nvl(rs("ID").value)) Then
            MsgBox "选择的项目“" & zlCommFun.Nvl(rs("名称").value) & "”以前已经选择！", vbInformation, gstrSysName
            Exit Function
        End If
        
        vsf.EditText = zlCommFun.Nvl(rs("名称").value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("名称").value)
        vsf.TextMatrix(vsf.Row, 2) = zlCommFun.Nvl(rs("英文名称").value)
        vsf.TextMatrix(vsf.Row, 3) = zlCommFun.Nvl(rs("英文缩写").value)
        vsf.TextMatrix(vsf.Row, 4) = zlCommFun.Nvl(rs("标本类型").value)
        vsf.TextMatrix(vsf.Row, 5) = zlCommFun.Nvl(rs("计算公式").value)
        vsf.RowData(vsf.Row) = zlCommFun.Nvl(rs("ID").value)
        
        ShowOpenTree = 1
        
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    For mlngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) = lngKey And vsf.Row <> mlngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.Nvl(rsData("ID")))
        
        On Error GoTo errHand
        For lngLoop = 0 To objMsf.Cols - 1
        
            On Error Resume Next
            strMask = ""
            strMask = MaskArray(lngLoop)
                                    
            On Error GoTo errHand
            If strMask <> "" Then
                objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
            Else
                objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop)))
            End If
                        
        Next
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowEdit(ByVal frmMain As Form, ByVal lngUpKey As Long, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    mblnStartUp = True
    mblnOK = False
    
    mlngUpKey = lngUpKey
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then
        cmdOK.Tag = ""
        Exit Function
    End If
    
    If ReadData = False Then
        cmdOK.Tag = ""
        Exit Function
    End If
    
    If mlngKey = 0 Then
        
        If GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\诊疗项目增加\", "编码", 0) = 0 Then
        
            mstrSQL = "SELECT NVL(MAX(编码),'0000000') AS 编码 FROM 诊疗项目目录 WHERE 类别 >= 'A'"
            zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
            If mRs.BOF = False Then txt(0).Text = Right(String(10, "0") & Val(mRs("编码")) + 1, Len(mRs("编码")))
            
        Else
            strTmp = Mid(txt(8).Text, 2, InStr(1, txt(8).Text, "]") - 2)
            
            mstrSQL = "SELECT NVL(MAX(编码),'0000000') AS 编码 FROM 诊疗项目目录 WHERE 类别 >= 'A' and 编码 like '" & strTmp & "%'"
            zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
            If mRs.BOF = False Then txt(0).Text = strTmp & Right(String(10, "0") & Val(mRs("编码")) + 1, Len(mRs("编码")) - Len(strTmp))
            
        End If
    End If
    
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '1.最大输入长度
    txt(1).MaxLength = GetMaxLength("诊疗项目目录", "名称")
    txt(0).MaxLength = GetMaxLength("诊疗项目目录", "编码")
    txt(2).MaxLength = GetMaxLength("诊疗项目别名", "简码")
    txt(3).MaxLength = GetMaxLength("诊疗项目别名", "简码")
    txt(7).MaxLength = GetMaxLength("诊疗项目目录", "计算单位")
    txt(4).MaxLength = GetMaxLength("诊疗项目别名", "名称")
    txt(5).MaxLength = GetMaxLength("诊疗项目别名", "简码")
    txt(6).MaxLength = GetMaxLength("诊疗项目别名", "简码")
            
    '2.检验类型
    mstrSQL = "SELECT 编码||'-'||名称,0 FROM 诊疗检验类型"
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    Call AddComboData(cbo(0), rs)
    If cbo(0).ListCount > 0 Then cbo(0).ListIndex = 0
    
    '3.适用性别
    cbo(1).AddItem "1-所有"
    cbo(1).ItemData(cbo(1).NewIndex) = 0
    
    cbo(1).AddItem "2-男性"
    cbo(1).ItemData(cbo(1).NewIndex) = 1
    
    cbo(1).AddItem "3-女性"
    cbo(1).ItemData(cbo(1).NewIndex) = 2
    
    cbo(1).ListIndex = 0
    
    '2.项目分类
    mstrSQL = "SELECT '['||编码||']'||名称 AS 名称 FROM 诊疗分类目录 WHERE ID=" & mlngUpKey
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        txt(8).Text = zlCommFun.Nvl(rs("名称").value)
        txt(8).Tag = mlngUpKey
    End If
    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "检验项目", 3300, 1, "...", 1
        .NewColumn "英文名称", 1500, 1
        .NewColumn "英文缩写", 810, 1
        .NewColumn "标本类型", 900, 1
        .NewColumn "计算公式", 0, 1
        .FixedCols = 1
        .ColHidden(2) = True
    End With
        
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    mstrSQL = "SELECT A.*,B.名称 AS 上级名称 " & _
                "FROM 诊疗项目目录 A,诊疗分类目录 B " & _
                "WHERE A.分类id=B.ID(+) AND A.ID=" & mlngKey
                
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        txt(0).Text = zlCommFun.Nvl(rs("编码").value)
        txt(1).Text = zlCommFun.Nvl(rs("名称").value)
        txt(7).Text = zlCommFun.Nvl(rs("计算单位").value)
        txt(8).Text = zlCommFun.Nvl(rs("上级名称").value)
        txt(8).Tag = zlCommFun.Nvl(rs("分类id").value)
        
        On Error Resume Next
        cbo(1).ListIndex = zlCommFun.Nvl(rs("适用性别").value, 0)
        On Error GoTo errHand
                
        chk(1).value = zlCommFun.Nvl(rs("执行安排").value, 0)
        
        zlControl.CboLocate cbo(0), zlCommFun.Nvl(rs("操作类型").value)
        
        Select Case zlCommFun.Nvl(rs("服务对象").value, 1)
        Case 1
            chk(2).value = 1
        Case 2
            chk(3).value = 1
        Case 3
            chk(2).value = 1
            chk(3).value = 1
        End Select
                
    End If
    
    mstrSQL = "SELECT 名称,简码,码类,性质 " & _
                "FROM 诊疗项目别名 A " & _
                "WHERE A.性质 IN (1,9) AND A.诊疗项目id=" & mlngKey
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            If rs("性质") = 1 Then
                If rs("码类").value = 1 Then
                    txt(2).Text = zlCommFun.Nvl(rs("简码").value)
                Else
                    txt(3).Text = zlCommFun.Nvl(rs("简码").value)
                End If
            Else
                txt(4).Text = zlCommFun.Nvl(rs("名称").value)
                If rs("码类").value = 1 Then
                    txt(5).Text = zlCommFun.Nvl(rs("简码").value)
                Else
                    txt(6).Text = zlCommFun.Nvl(rs("简码").value)
                End If
            End If
            rs.MoveNext
        Loop
    End If
                
    mstrSQL = "SELECT '' AS 序号," & _
                      "A.报告项目id AS ID," & _
                      "C.名称 AS 检验项目," & _
                      "D.英文名 AS 英文名称," & _
                      "E.缩写 AS 英文缩写,zlGetSample(D.ID) AS 标本类型,E.计算公式 " & _
                 "FROM 检验报告项目 A," & _
                      "(SELECT 报告项目id FROM 检验报告项目 WHERE 诊疗项目id = " & mlngKey & ") B," & _
                      "诊疗项目目录 C,诊治所见项目 D,检验项目 E,检验报告项目 F " & _
                "WHERE A.报告项目id = B.报告项目id AND A.诊疗项目id <> " & mlngKey & " AND " & _
                      "nvl(C.组合项目,0) = 0 AND A.诊疗项目id = C.ID AND C.ID=F.诊疗项目id AND F.报告项目id=D.ID AND D.ID=E.诊治项目id"
    
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        vsf.TextMatrix(0, 0) = "序号"
        Call FillGrid(vsf, rs)
        vsf.TextMatrix(0, 0) = ""
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varTmp As Variant
    Dim lngLeftPos As Long
    Dim lngRightPos As Long
    Dim strTmpID As String
    Dim strID As String
    Dim rs As New ADODB.Recordset
                
    If Trim(txt(0).Text) = "" Then
        ShowSimpleMsg "项目编码不能为空值！"
        LocationObj txt(0)
        Exit Function
    End If
    
    strTmp = CheckNumeric(txt(0).Text, txt(0).MaxLength, 0, 1)
    If strTmp <> "" Then
        ShowSimpleMsg "编码" & strTmp
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        ShowSimpleMsg "项目名称不能为空值！"
        LocationObj txt(1)
        Exit Function
    End If
    
    '检验计算公式是否正确
    For mlngLoop = 1 To vsf.Rows - 1
        strTmpID = strTmpID & "," & Trim(vsf.RowData(mlngLoop))
    Next
    strTmpID = strTmpID & ","
    
    For mlngLoop = 1 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(mlngLoop, 5)) <> "" Then
            
            '是计算项目,有计算公式
            strTmp = Trim(vsf.TextMatrix(mlngLoop, 5))
            
            lngLeftPos = InStr(strTmp, "[")
            lngRightPos = InStr(strTmp, "]")
            Do While lngLeftPos > 0 And lngRightPos > 0
                
                strID = Trim(Mid(strTmp, lngLeftPos + 1, lngRightPos - lngLeftPos - 1))
                If strID <> "" Then
                    '找到此项目
                    If InStr(strTmpID, "," & strID & ",") = 0 Then
                        '没有在其中
                        
                        Call zlDatabase.OpenRecordset(rs, "SELECT 中文名 FROM 诊治所见项目 WHERE ID=" & Val(strID), Me.Caption)
                        If rs.BOF = False Then
                            
                            ShowSimpleMsg "计算公式中“" & zlCommFun.Nvl(rs("中文名")) & "”项目未被包含在当前组合检验项目中！"
                            
                            vsf.Row = mlngLoop
                            vsf.Col = 1
                            vsf.ShowCell vsf.Row, vsf.Col
                                
                            Exit Function
                        End If
                        
                        
                        ShowSimpleMsg "“" & Trim(vsf.TextMatrix(mlngLoop, 1)) & "”项目的计算公式有误！"
                            
                        vsf.Row = mlngLoop
                        vsf.Col = 1
                        vsf.ShowCell vsf.Row, vsf.Col
                            
                        Exit Function
                        
                    End If
                End If
                
                strTmp = Mid(strTmp, lngRightPos + 1)
                lngLeftPos = InStr(strTmp, "[")
                lngRightPos = InStr(strTmp, "]")
            Loop
            
        End If
    Next
            
    '检验标本类型是否相同
    
    If Val(vsf.RowData(1)) > 0 Then
        strTmp = vsf.TextMatrix(1, 4)
        
        For mlngLoop = 2 To vsf.Rows - 1
            If Val(vsf.RowData(mlngLoop)) > 0 And vsf.TextMatrix(mlngLoop, 4) <> "" Then
                
                strTmp = GetCommon(strTmp, Split(vsf.TextMatrix(mlngLoop, 4), ","))
                
                If strTmp = "" Then
                    
                    ShowSimpleMsg "设置的检验项目没有共同的标本类型！"
                    vsf.Row = mlngLoop
                    vsf.Col = 4
                    vsf.ShowCell vsf.Row, vsf.Col
                    
                    Exit Function
                End If
                
            End If
        Next
    End If
    
    
    ValidData = True
    
End Function

Private Function GetCommon(ByVal str标准 As String, ByVal var检查 As Variant) As String
                    
    Dim lngLoop As Long
        
    GetCommon = ""
    
    For lngLoop = 0 To UBound(var检查)
        If InStr("," & str标准 & ",", "," & var检查(lngLoop) & ",") > 0 Then
            GetCommon = GetCommon & "," & var检查(lngLoop)
        End If
    Next
    
    If GetCommon <> "" Then GetCommon = Mid(GetCommon, 2)
    
End Function


Private Function SaveData(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim blnTran As Boolean
    Dim strSQL(1 To 2) As String
    
    On Error GoTo errHand
    
    If mlngKey = 0 Then
        '新增
        lngKey = zlDatabase.GetNextId("诊疗项目目录")
        
        strSQL(1) = "ZL_诊疗项目_INSERT('C'," & Val(txt(8).Tag) & "," & _
                                        lngKey & ",'" & _
                                        txt(0).Text & "','" & _
                                        txt(1).Text & "','" & _
                                        txt(2).Text & "','" & _
                                        txt(3).Text & "','" & _
                                        txt(4).Text & "','" & _
                                        txt(5).Text & "','" & _
                                        txt(6).Text & "','" & _
                                        zlCommFun.GetNeedName(cbo(0).Text) & "',1,1,3,'" & _
                                        txt(7).Text & "'," & _
                                        cbo(1).ItemData(cbo(1).ListIndex) & "," & _
                                        chk(1).value & "," & _
                                        IIf(chk(2).value And chk(3).value, 3, IIf(chk(2).value, 1, IIf(chk(3).value, 2, 0))) & "," & _
                                        "1,NULL,NULL,1,NULL,NULL,'',NULL)"
    Else
        '修改
        lngKey = mlngKey

        strSQL(1) = "ZL_诊疗项目_UPDATE('C'," & Val(txt(8).Tag) & "," & _
                                        lngKey & ",'" & _
                                        txt(0).Text & "','" & _
                                        txt(1).Text & "','" & _
                                        txt(2).Text & "','" & _
                                        txt(3).Text & "','" & _
                                        txt(4).Text & "','" & _
                                        txt(5).Text & "','" & _
                                        txt(6).Text & "','" & _
                                        zlCommFun.GetNeedName(cbo(0).Text) & "',1,1,3,'" & _
                                        txt(7).Text & "'," & _
                                        cbo(1).ItemData(cbo(1).ListIndex) & "," & _
                                        chk(1).value & "," & _
                                        IIf(chk(2).value And chk(3).value, 3, IIf(chk(2).value, 1, IIf(chk(3).value, 2, 0))) & "," & _
                                        "1,NULL,NULL,1,NULL,NULL,'',NULL,1)"
    End If
    
    For mlngLoop = 1 To vsf.Rows - 1
        If vsf.RowData(mlngLoop) > 0 Then
            strValue = strValue & "|null^" & Val(vsf.RowData(mlngLoop))
        End If
    Next
    If strValue <> "" Then strValue = Mid(strValue, 2)
    strSQL(2) = "ZL_检验报告项目_UPDATE(" & lngKey & ",'" & strValue & "')"
    
    blnTran = True
    
    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    SaveData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Sub cbo_Click(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim objPoint As POINTAPI
    Dim rs As New ADODB.Recordset
    
    mstrSQL = "Select ID," & _
                    "上级ID," & _
                    "0 as 末级," & _
                    "'['||编码||']'||名称 AS 名称 " & _
              "From 诊疗分类目录 where 类型=5 " & _
              "Start With 上级ID IS NULL " & _
                "Connect by Prior ID=上级ID "
                
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF Then Exit Sub
    
    Call ClientToScreen(txt(8).hwnd, objPoint)
    If frmSelectTree.ShowSelect(Me, rs, objPoint.x * 15 - 30, objPoint.y * 15 + txt(8).Height - 30, txt(8).Width, 3000, txt(8).Height, txt(8).Tag, "检验分类选择", "请选择一个分类位置") Then
        txt(8).Text = zlCommFun.Nvl(rs("名称").value)
        txt(8).Tag = zlCommFun.Nvl(rs("ID").value)
    End If
    
    txt(8).SetFocus
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub CmdOK_Click()
    Dim lngKey As Long
    Dim strTmp As String
    
    If cmdOK.Tag <> "" Then
        
        If ValidData = False Then Exit Sub
        
        If SaveData(lngKey) = False Then Exit Sub
        
        mblnOK = True
        
        '刷新主界面的数据显示
        Call mfrmMain.EditRefresh(2, lngKey)
        
        If mlngKey = 0 Then
            
            '清除控件内容
            txt(0).Text = ""
            txt(1).Text = ""
            txt(2).Text = ""
            txt(3).Text = ""
            txt(4).Text = ""
            txt(5).Text = ""
            txt(6).Text = ""
            txt(7).Text = ""
            
            '清除上一项目的子项目
            Call ClearGrid(vsf)
            
            '产生缺省的项目编码
            If GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\诊疗项目增加\", "编码", 0) = 0 Then
            
                mstrSQL = "SELECT NVL(MAX(编码),'0000000') AS 编码 FROM 诊疗项目目录 WHERE 类别 >= 'A'"
                zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
                If mRs.BOF = False Then txt(0).Text = Right(String(10, "0") & Val(mRs("编码")) + 1, Len(mRs("编码")))
                
            Else
                strTmp = Mid(txt(8).Text, 2, InStr(1, txt(8).Text, "]") - 2)
                
                mstrSQL = "SELECT NVL(MAX(编码),'0000000') AS 编码 FROM 诊疗项目目录 WHERE 类别 >= 'A' and 编码 like '" & strTmp & "%'"
                zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
                If mRs.BOF = False Then txt(0).Text = strTmp & Right(String(10, "0") & Val(mRs("编码")) + 1, Len(mRs("编码")) - Len(strTmp))
                
            End If
            
            '定位并选中项目编码
            LocationObj txt(0)
            
            cmdOK.Tag = ""
            Exit Sub
        End If
        
    End If
    
    cmdOK.Tag = ""
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
    Case 1, 4, 7
        Call zlCommFun.OpenIme(True)
    End Select
    
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        If Index = 8 Then zlCommFun.PressKey vbKeyTab
    Else
        Select Case Index
        Case 2, 3, 5, 6
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 0
            KeyAscii = FilterKeyAscii(KeyAscii, 1)
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 4, 7
        Call zlCommFun.OpenIme(False)
    End Select
    
    If Index = 1 Then
        If InStr(txt(Index).Text, "'") = 0 Then
            txt(2).Text = zlGetSymbol(txt(Index).Text, 0)
            txt(3).Text = zlGetSymbol(txt(Index).Text, 1)
        End If
    End If
    
    If Index = 4 Then
        If InStr(txt(Index).Text, "'") = 0 Then
            txt(5).Text = zlGetSymbol(txt(Index).Text, 0)
            txt(6).Text = zlGetSymbol(txt(Index).Text, 1)
        End If
    End If
    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    cmdOK.Tag = "Changed"
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    If Val(vsf.RowData(Row)) <= 0 Then
        Col = 1
        Cancel = True
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Select Case ShowOpenTree
    Case 0
        '没有匹配的项目
        MsgBox "没有找到相匹配的结果或者标本类型不一致！", vbInformation, gstrSysName
        
    Case 1
        '选取了一个项目
        cmdOK.Tag = "Changed"
    Case 2
        '取消了本次选择
        
    End Select
    
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                Cancel = True
                Exit Sub
            End If
                        
            Select Case ShowOpenList(vsf.EditText)
            Case 0
                '没有匹配的项目
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
                MsgBox "没有找到相匹配的结果或者标本类型不一致！", vbInformation, gstrSysName
                
            Case 1
                '选取了一个项目
                cmdOK.Tag = "Changed"
            Case 2
                '取消了本次选择
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
            End Select
        End If
    Else
        cmdOK.Tag = "Changed"
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn Then
        If Col = 1 And Trim(vsf.TextMatrix(Row, Col)) = "" Then
            zlCommFun.PressKey vbKeyTab
            Cancel = True
            KeyAscii = 0
        End If
    End If
End Sub


